/* ============================================
   agent.js — Iterative tool-calling agent loop
   Drives Groq + Tools + Journal to fulfill a user request:
     call → if tool_calls, dispatch tools, push tool response messages
     back → repeat until the model emits a final text answer (or budget
     exhausted, or user clicks Stop).

   Smart Routing: when Config.model === 'smart', the agent classifies
   the user's request and routes to a fast model (8B) for simple tasks
   or a capable model (70B) for complex builds. It can also upgrade
   mid-request if the fast model starts calling complex tools.

   Conversation history (OpenAI messages format) is owned by App and
   passed in; the agent appends assistant turns (with tool_calls) and
   tool response messages as it goes.
   ============================================ */

// Tools that signal a complex workflow requiring the capable model.
const COMPLEX_TOOLS = new Set([
  'write_range', 'format_range', 'add_sheet', 'delete_sheet',
  'create_table', 'create_pivot', 'create_chart', 'add_slicer',
  'conditional_format', 'insert_rows_cols', 'delete_rows_cols', 'sort_range'
]);

// Intent keywords that signal a complex build/analysis request.
const COMPLEX_KEYWORDS = [
  // Build/create actions
  'create', 'build', 'make', 'generate', 'construct', 'set up', 'setup',
  'dashboard', 'report', 'panel', 'informe', 'crear', 'construir', 'generar',
  'tabla', 'gráfico', 'grafico', 'pivot', 'dinámica', 'dinamica',
  // Analysis actions
  'analyze', 'analysis', 'analizar', 'análisis', 'compare', 'comparar',
  'troubleshoot', 'debug', 'fix', 'repair', 'corregir', 'arreglar',
  'formula', 'fórmula', 'calculate', 'calcular', 'summary', 'resumen',
  // Formatting/layout
  'format', 'formatear', 'layout', 'design', 'diseño', 'diseñar',
  'conditional', 'condicional', 'slicer', 'segmentador',
  // Complex operations
  'merge', 'combinar', 'consolidate', 'consolidar', 'reconcile', 'conciliar',
  'forecast', 'pronóstico', 'pronostico', 'trend', 'tendencia',
  'variance', 'varianza', 'budget', 'presupuesto'
];

const Router = {
  /**
   * Classify the user's request as simple or complex.
   * Returns the model name to use for the first round.
   */
  classify(userText) {
    // If user picked a specific model, always use it.
    if (Config.model !== 'smart') return Config.model;

    const lower = userText.toLowerCase();

    // Check for complex keywords
    for (const kw of COMPLEX_KEYWORDS) {
      if (lower.includes(kw)) return Config.capableModel;
    }

    // Multi-sentence requests are likely complex
    const sentences = userText.split(/[.!?;]+/).filter(s => s.trim().length > 0);
    if (sentences.length > 2) return Config.capableModel;

    // Default: fast model for simple questions and lookups
    return Config.fastModel;
  },

  /**
   * Decide whether to upgrade from fast to capable model based on
   * the tools called in the current round.
   * Returns the model to use for the next round, or null to keep current.
   */
  shouldUpgrade(currentModel, functionCalls) {
    if (currentModel === Config.capableModel) return null;
    if (Config.model !== 'smart') return null;

    for (const fc of functionCalls) {
      if (COMPLEX_TOOLS.has(fc.name)) {
        return Config.capableModel;
      }
    }
    return null;
  }
};

const Agent = {
  MAX_ROUNDS: 20,           // hard cap on tool-call rounds per request
  MAX_CONSECUTIVE_ERRORS: 3,// bail if the same tool keeps failing
  MAX_CONTINUATIONS: 3,     // auto-continue attempts when MAX_TOKENS truncates a response

  /**
   * Run the agent loop for one user message.
   *
   * @param {object} opts
   * @param {string} opts.userText
   * @param {Array}  opts.conversation  — App.conversation (mutated in place, OpenAI messages format)
   * @param {AbortSignal} [opts.signal] — for Stop button
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
      onText, onThinking, onToolStart, onToolEnd, onToolError, onRound
    } = opts;

    Journal.beginRequest();

    const systemPrompt = Prompt.build();
    const tools = Tools.declarations();
    let toolCallCount = 0;
    let consecutiveErrors = 0;
    let lastErrorTool = null;
    let finalText = '';
    let aborted = false;
    const toolErrors = []; // collect all tool errors to surface to the user

    // Smart routing: pick the initial model based on request complexity.
    let currentModel = Router.classify(userText);
    console.log(`Router: using ${currentModel} for this request`);

    // Auto-continue state: when the model hits MAX_TOKENS, the response is
    // truncated mid-sentence. We save the partial text, push it to history,
    // and ask the model to continue — concatenating the pieces seamlessly.
    let continuationBase = '';
    let continuationsUsed = 0;

    for (let round = 1; round <= this.MAX_ROUNDS; round++) {
      if (signal && signal.aborted) { aborted = true; break; }
      if (onRound) onRound(round, this.MAX_ROUNDS);

      const result = await Groq.generateWithFallback({
        systemPrompt,
        messages: conversation,
        tools,
        model: currentModel,
        onText: (chunk, full) => {
          // Prepend any text from previous truncated segments so the UI
          // shows the complete answer as it streams in.
          finalText = continuationBase + full;
          if (onText) onText(chunk, finalText);
        },
        onThinking: (text) => {
          // Fired when the model emits text alongside tool calls (reasoning).
          if (onThinking) onThinking(text);
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
          conversation.push(result.message);
          conversation.push({ role: 'user', content: I18n.lang === 'es'
            ? 'Continúa tu respuesta anterior exactamente desde donde se detuvo. No repitas lo que ya dijiste.'
            : 'Continue your previous response exactly from where it stopped. Do not repeat what you already said.'
          });
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

      // If the model emitted text alongside tool calls, that's reasoning —
      // fire onThinking so the UI can show it in the thinking block.
      if (result.text && onThinking) onThinking(result.text);

      // Smart routing: check if we should upgrade to the capable model
      // for the next round based on the tools being called.
      const upgrade = Router.shouldUpgrade(currentModel, result.functionCalls);
      if (upgrade) {
        console.log(`Router: upgrading from ${currentModel} to ${upgrade} (complex tools detected)`);
        currentModel = upgrade;
      }

      // Append the assistant message (with tool_calls) to history.
      conversation.push(result.message);

      // Dispatch each function call and collect responses.
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

        // OpenAI format: tool response is a message with role "tool",
        // tool_call_id matching the call, and content as a JSON string.
        const toolResponseContent = toolResult.ok
          ? JSON.stringify({ ok: true, ...toolResult.result })
          : JSON.stringify({ ok: false, error: toolResult.error });

        conversation.push({
          role: 'tool',
          tool_call_id: fc.id,
          content: toolResponseContent
        });

        if (consecutiveErrors >= this.MAX_CONSECUTIVE_ERRORS) {
          // Tell the model to stop retrying this tool and explain.
          conversation.push({
            role: 'user',
            content: I18n.lang === 'es'
              ? `SYSTEM: La herramienta "${fc.name}" ha fallado ${consecutiveErrors} veces consecutivas con error: ${toolResult.error}. Deja de intentar esta herramienta y responde al usuario en texto explicando el problema y qué información necesitas para continuar.`
              : `SYSTEM: Tool "${fc.name}" has failed ${consecutiveErrors} consecutive times with error: ${toolResult.error}. Stop trying this tool and respond to the user in text explaining the problem and what information you need to continue.`
          });
          consecutiveErrors = 0;
          lastErrorTool = null;
        }
      }

      if (aborted) break;

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

    // Append the final text answer to conversation history as an assistant message.
    conversation.push({ role: 'assistant', content: finalText });

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
   * Trim large data payloads from old tool response messages to keep context
   * bounded. Keeps the most recent 3 turns intact; older tool messages get
   * their large content payloads replaced with compact summaries.
   */
  trimHistory(conversation) {
    const TURNS_KEPT_FULL = 6;
    for (let i = 0; i < conversation.length - TURNS_KEPT_FULL; i++) {
      const msg = conversation[i];
      if (!msg || msg.role !== 'tool') continue;
      try {
        const r = JSON.parse(msg.content);
        let modified = false;
        // Replace large arrays with a size note.
        if (Array.isArray(r.data)) {
          const rows = r.data.length;
          const cols = rows > 0 && Array.isArray(r.data[0]) ? r.data[0].length : 0;
          r.data = `[${rows}x${cols} array omitted to save context]`;
          modified = true;
        }
        if (Array.isArray(r.sample)) {
          r.sample = `[${r.sample.length}x sample omitted]`;
          modified = true;
        }
        if (Array.isArray(r.matches)) {
          r.matches = `[${r.matches.length} matches omitted]`;
          modified = true;
        }
        // Trim the overview string — it's the single largest payload.
        if (typeof r.overview === 'string' && r.overview.length > 200) {
          const sheetCount = r.sheets ? r.sheets.length : '?';
          r.overview = `[workbook overview omitted to save context - ${sheetCount} sheets. Call get_workbook_overview again if you need details.]`;
          modified = true;
        }
        // Trim the sheets array from old overview responses.
        if (Array.isArray(r.sheets) && r.sheets.length > 0) {
          r.sheets = `[${r.sheets.length} sheet names omitted]`;
          modified = true;
        }
        if (modified) {
          msg.content = JSON.stringify(r);
        }
      } catch (e) {
        // Not JSON or already trimmed — skip.
      }
    }
  }
};
