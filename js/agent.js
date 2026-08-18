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
  MAX_CONTINUATIONS: 3,     // auto-continue attempts when MAX_TOKENS truncates a response

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
   * @param {function} opts.onToolError — (name, error) when a tool fails, for prominent UI notification
   * @param {function} opts.onRound     — (round, maxRounds) at start of each round
   *
   * Returns: { ok, finalText, rounds, toolCalls, sealed }
   *   On error: { ok:false, error, errorType, partial? }
   */
  async run(opts) {
    const {
      userText, conversation, signal,
      onThinking, onText, onToolStart, onToolEnd, onToolError, onRound
    } = opts;

    Journal.beginRequest();
    Config.resetRotation(); // start each user request at key pool index 0

    const systemPrompt = Prompt.build();
    const tools = Tools.declarations();
    let toolCallCount = 0;
    let consecutiveErrors = 0;
    let lastErrorTool = null;
    let finalText = '';
    let aborted = false;
    const toolErrors = []; // collect all tool errors to surface to the user

    // Auto-continue state: when the model hits MAX_TOKENS, the response is
    // truncated mid-sentence. We save the partial text, push it to history,
    // and ask the model to continue — concatenating the pieces seamlessly.
    let continuationBase = '';
    let continuationsUsed = 0;

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
          // Prepend any text from previous truncated segments so the UI
          // shows the complete answer as it streams in.
          finalText = continuationBase + full;
          if (onText) onText(chunk, finalText);
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
          sealed,
          toolErrors
        };
      }

      // No function calls → the model is done (or was truncated).
      if (result.functionCalls.length === 0) {
        // MAX_TOKENS: the response was cut off mid-generation. Auto-continue
        // by asking the model to resume from where it stopped, up to a cap.
        if (result.finishReason === 'MAX_TOKENS' && continuationsUsed < this.MAX_CONTINUATIONS && result.text) {
          continuationBase = finalText; // save accumulated text so far
          conversation.push({ role: 'model', parts: result.rawParts });
          conversation.push({ role: 'user', parts: [{ text: I18n.lang === 'es'
            ? 'Continúa tu respuesta anterior exactamente desde donde se detuvo. No repitas lo que ya dijiste.'
            : 'Continue your previous response exactly from where it stopped. Do not repeat what you already said.'
          }] });
          continuationsUsed++;
          continue; // next round calls generateWithFallback again
        }

        // If this followed a continuation, finalText was already set by the
        // onText callback to continuationBase + full — don't overwrite it
        // with just result.text (the continuation piece alone).
        if (!continuationBase) {
          finalText = result.text || (I18n.lang === 'es'
            ? 'Listo.'
            : 'Done.');
        }
        // If continuationBase is set but result.text is empty, finalText
        // already holds the partial text from the onText callback.
        break;
      }

      // The model returned function calls — we're in tool-calling mode, so
      // reset the continuation base (any partial text is already in history).
      continuationBase = '';

      // Append the model turn to history using the RAW parts from the API
      // response. This preserves thoughtSignature at the Part level, which
      // Gemini 3 requires — reconstructing parts manually drops the signature
      // and causes HTTP 400 on the next turn.
      conversation.push({ role: 'model', parts: result.rawParts });

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
          // Collect error for surfacing to user, and fire prominent UI callback.
          toolErrors.push({ tool: fc.name, error: toolResult.error });
          if (onToolError) onToolError(fc.name, toolResult.error);
        } else {
          consecutiveErrors = 0;
          lastErrorTool = null;
        }

        // Gemini expects functionResponse.name to match the call name, and
        // response to be a JSON object.
        responseParts.push({
          functionResponse: {
            name: fc.name,
            response: toolResult.ok
              ? { ok: true, ...toolResult.result }
              : { ok: false, error: toolResult.error }
          }
        });

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
        aborted: true,
        toolErrors
      };
    }

    // If we hit the round budget without a final text, synthesize a note.
    if (!finalText) {
      finalText = I18n.lang === 'es'
        ? `Alcancé el límite de ${this.MAX_ROUNDS} rondas de herramientas. Se completaron ${toolCallCount} llamadas a herramientas. Puedes pedirme que continúe o que verifique el resultado.`
        : `I hit the limit of ${this.MAX_ROUNDS} tool rounds. ${toolCallCount} tool calls completed. You can ask me to continue or to verify the result.`;
    }

    // If there were tool errors that the model didn't address in its final
    // answer, append them so the user sees what went wrong.
    if (toolErrors.length > 0 && finalText) {
      const errorSummary = toolErrors.map(e => `${e.tool}: ${e.error}`).join('; ');
      const hasErrorMention = finalText.toLowerCase().includes('error') || finalText.toLowerCase().includes('no se pudo') || finalText.toLowerCase().includes('could not') || finalText.toLowerCase().includes('failed');
      if (!hasErrorMention) {
        finalText += I18n.lang === 'es'
          ? `\n\n⚠ Errores de herramientas: ${errorSummary}`
          : `\n\n⚠ Tool errors: ${errorSummary}`;
      }
    }

    // Append the final text answer to conversation history as a model turn.
    conversation.push({ role: 'model', parts: [{ text: finalText }] });

    return {
      ok: true,
      finalText,
      rounds: this.MAX_ROUNDS,
      toolCalls: toolCallCount,
      sealed,
      toolErrors
    };
  },

  /**
   * Trim large data payloads from old tool responses to keep context bounded.
   * Keeps the most recent 6 turns intact; older functionResponse.result.data
   * arrays are replaced with a summary string.
   */
  trimHistory(conversation) {
    // Aggressive trimming: keep only the last 3 turns full. Older turns
    // get their large payloads replaced with compact summaries to save
    // input tokens (and thus quota) on subsequent API calls.
    const TURNS_KEPT_FULL = 3;
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
          // Trim the overview string — it's the single largest payload
          // (all sheets, headers, stats, sample rows). Replace with a
          // compact note so the model knows it saw the overview but
          // doesn't resend thousands of tokens every round.
          if (typeof r.overview === 'string' && r.overview.length > 200) {
            const sheetCount = r.sheets ? r.sheets.length : '?';
            r.overview = `[workbook overview omitted to save context - ${sheetCount} sheets. Call get_workbook_overview again if you need details.]`;
          }
          // Trim the sheets array from old overview responses.
          if (Array.isArray(r.sheets) && r.sheets.length > 0) {
            r.sheets = `[${r.sheets.length} sheet names omitted]`;
          }
        }
      }
    }
  }
};
