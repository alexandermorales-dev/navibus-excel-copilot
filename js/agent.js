/* ============================================
   agent.js — Iterative tool-calling agent loop
   Drives Gemini + Tools + Journal to fulfill a user request:
     call → if functionCalls, dispatch tools, push functionResponse(s)
     back → repeat until the model emits a final text answer (or budget
     exhausted, or user clicks Stop).

   Conversation history (Gemini contents format) is owned by App and
   passed in; the agent appends model turns (with functionCall parts)
   and user turns (with functionResponse parts) as it goes.
   ============================================ */

const Agent = {
  MAX_ROUNDS: 20,           // hard cap on tool-call rounds per request
  MAX_CONSECUTIVE_ERRORS: 3,// bail if the same tool keeps failing

  /**
   * Run the agent loop for one user message.
   *
   * @param {object} opts
   * @param {string} opts.userText
   * @param {Array}  opts.conversation  — App.conversation (mutated in place)
   * @param {AbortSignal} [opts.signal] — for Stop button
   * @param {function} opts.onThinking  — (chunk, full) live reasoning
   * @param {function} opts.onText      — (chunk, full) live final text
   * @param {function} opts.onToolStart — (callId, name, args) before dispatch
   * @param {function} opts.onToolEnd   — (callId, name, result) after dispatch
   * @param {function} opts.onRound     — (round, maxRounds) at start of each round
   *
   * Returns: { ok, finalText, rounds, toolCalls, sealed }
   *   On error: { ok:false, error, errorType, partial? }
   */
  async run(opts) {
    const {
      userText, conversation, signal,
      onThinking, onText, onToolStart, onToolEnd, onRound
    } = opts;

    Journal.beginRequest();

    const systemPrompt = Prompt.build();
    const tools = Tools.declarations();
    let toolCallCount = 0;
    let consecutiveErrors = 0;
    let lastErrorTool = null;
    let finalText = '';
    let aborted = false;

    for (let round = 1; round <= this.MAX_ROUNDS; round++) {
      if (signal && signal.aborted) { aborted = true; break; }
      if (onRound) onRound(round, this.MAX_ROUNDS);

      const result = await Gemini.generateWithFallback({
        systemPrompt,
        contents: conversation,
        tools,
        toolConfig: { functionCallingConfig: { mode: 'AUTO' } },
        onThinking,
        onText: (chunk, full) => {
          finalText = full;
          if (onText) onText(chunk, full);
        },
        signal
      });

      if (!result.ok) {
        // Seal journal so undo is available for whatever did get written.
        const sealed = Journal.sealRequest();
        return {
          ok: false,
          error: result.error,
          errorType: result.errorType,
          rounds: round - 1,
          toolCalls: toolCallCount,
          partial: toolCallCount > 0,
          sealed
        };
      }

      // No function calls and no text → model ended empty-handed.
      if (result.functionCalls.length === 0) {
        finalText = result.text || (I18n.lang === 'es'
          ? 'Listo.'
          : 'Done.');
        break;
      }

      // Append the model turn (with functionCall parts + any text) to history.
      // CRITICAL: preserve thought_signature on functionCall parts — Gemini
      // requires it to be echoed back when thinking is enabled, or the next
      // request returns HTTP 400 "missing thought_signature".
      const modelParts = [];
      if (result.text) modelParts.push({ text: result.text });
      for (const fc of result.functionCalls) {
        const fcPart = { functionCall: { name: fc.name, args: fc.args } };
        if (fc.thought_signature) {
          fcPart.functionCall.thought_signature = fc.thought_signature;
        }
        if (fc.thought) {
          fcPart.thought = true;
        }
        modelParts.push(fcPart);
      }
      conversation.push({ role: 'model', parts: modelParts });

      // Dispatch each function call and collect responses.
      const responseParts = [];
      for (let i = 0; i < result.functionCalls.length; i++) {
        if (signal && signal.aborted) { aborted = true; break; }

        const fc = result.functionCalls[i];
        const callId = `r${round}c${i}`;
        toolCallCount++;

        if (onToolStart) onToolStart(callId, fc.name, fc.args);

        const toolResult = await Tools.dispatch(fc.name, fc.args);

        if (onToolEnd) onToolEnd(callId, fc.name, toolResult);

        // Track consecutive errors per tool to bail on stuck loops.
        if (!toolResult.ok) {
          if (lastErrorTool === fc.name) consecutiveErrors++;
          else { consecutiveErrors = 1; lastErrorTool = fc.name; }
        } else {
          consecutiveErrors = 0;
          lastErrorTool = null;
        }

        // Gemini expects functionResponse.name to match the call name, and
        // response to be a JSON object. Echo thought_signature if present.
        const respPart = {
          functionResponse: {
            name: fc.name,
            response: toolResult.ok
              ? { ok: true, ...toolResult.result }
              : { ok: false, error: toolResult.error }
          }
        };
        if (fc.thought_signature) {
          respPart.thought_signature = fc.thought_signature;
        }
        responseParts.push(respPart);

        if (consecutiveErrors >= this.MAX_CONSECUTIVE_ERRORS) {
          // Tell the model to stop retrying this tool and explain.
          responseParts.push({
            text: I18n.lang === 'es'
              ? `SYSTEM: La herramienta "${fc.name}" ha fallado ${consecutiveErrors} veces consecutivas con error: ${toolResult.error}. Deja de intentar esta herramienta y responde al usuario en texto explicando el problema y qué información necesitas para continuar.`
              : `SYSTEM: Tool "${fc.name}" has failed ${consecutiveErrors} consecutive times with error: ${toolResult.error}. Stop trying this tool and respond to the user in text explaining the problem and what information you need to continue.`
          });
          consecutiveErrors = 0;
          lastErrorTool = null;
        }
      }

      if (aborted) break;

      // Push the function responses as a user turn (Gemini convention).
      conversation.push({ role: 'user', parts: responseParts });

      // Trim large tool results in older history to control context growth.
      this.trimHistory(conversation);

      // If the model also emitted final text alongside the calls, treat that
      // as a mid-stream narration, not the final answer — keep looping so it
      // can act on its own narration. The loop ends when a round returns text
      // WITHOUT function calls.
    }

    const sealed = Journal.sealRequest();

    if (aborted) {
      return {
        ok: true,
        finalText: finalText || (I18n.lang === 'es' ? '(Detenido por el usuario)' : '(Stopped by user)'),
        rounds: this.MAX_ROUNDS,
        toolCalls: toolCallCount,
        sealed,
        aborted: true
      };
    }

    // If we hit the round budget without a final text, synthesize a note.
    if (!finalText) {
      finalText = I18n.lang === 'es'
        ? `Alcancé el límite de ${this.MAX_ROUNDS} rondas de herramientas. Se completaron ${toolCallCount} llamadas a herramientas. Puedes pedirme que continúe o que verifique el resultado.`
        : `I hit the limit of ${this.MAX_ROUNDS} tool rounds. ${toolCallCount} tool calls completed. You can ask me to continue or to verify the result.`;
    }

    // Append the final text answer to conversation history as a model turn.
    conversation.push({ role: 'model', parts: [{ text: finalText }] });

    return {
      ok: true,
      finalText,
      rounds: this.MAX_ROUNDS,
      toolCalls: toolCallCount,
      sealed
    };
  },

  /**
   * Trim large data payloads from old tool responses to keep context bounded.
   * Keeps the most recent 6 turns intact; older functionResponse.result.data
   * arrays are replaced with a summary string.
   */
  trimHistory(conversation) {
    const TURNS_KEPT_FULL = 6;
    // Work from oldest; skip the last TURNS_KEPT_FULL turns.
    for (let i = 0; i < conversation.length - TURNS_KEPT_FULL; i++) {
      const turn = conversation[i];
      if (!turn || !turn.parts) continue;
      for (const part of turn.parts) {
        if (part.functionResponse && part.functionResponse.response) {
          const r = part.functionResponse.response;
          // Replace large arrays with a size note.
          if (Array.isArray(r.data)) {
            const rows = r.data.length;
            const cols = rows > 0 && Array.isArray(r.data[0]) ? r.data[0].length : 0;
            r.data = `[${rows}x${cols} array omitted to save context]`;
          }
          if (Array.isArray(r.sample)) {
            r.sample = `[${r.sample.length}x sample omitted]`;
          }
          if (Array.isArray(r.matches)) {
            r.matches = `[${r.matches.length} matches omitted]`;
          }
        }
      }
    }
  }
};
