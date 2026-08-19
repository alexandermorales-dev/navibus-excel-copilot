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
  MAX_ROUNDS: 15,           // hard cap on tool-call rounds per request
  MAX_CONSECUTIVE_ERRORS: 3,// bail if the same tool keeps failing
  MAX_CONTINUATIONS: 3,     // auto-continue attempts when MAX_TOKENS truncates a response
  STALE_ROUNDS: 3,          // if no successful write after this many rounds, nudge the model
  REMINDER_ROUNDS: 4,       // re-inject original request to keep the model focused

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

    const systemPrompt = Prompt.build(userText);
    const tools = Tools.declarations();
    let toolCallCount = 0;
    let consecutiveErrors = 0;
    let lastErrorTool = null;
    let finalText = '';
    let aborted = false;
    const toolErrors = []; // collect all tool errors to surface to the user
    let writeSuccessCount = 0;
    let roundsSinceWrite = 0;
    let lastRound = 0;  // track final round for post-loop summary
    const writeOps = []; // track what was actually written for summary

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
      lastRound = round;

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
          finalText = result.text || '';
        }
        // If continuationBase is set but result.text is empty, finalText
        // already holds the partial text from the onText callback.
        break;
      }

      // The model returned function calls — we're in tool-calling mode, so
      // reset the continuation base (any partial text is already in history).
      continuationBase = '';

      // If the model emitted reasoning alongside tool calls that wasn't
      // already streamed (e.g. non-streaming fallback or missed chunks),
      // fire onThinking so the UI can show it in the thinking block.
      if (result.reasoning && onThinking) onThinking(result.reasoning);

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

        // Track successful write operations for stale detection and summary.
        if (toolResult.ok && COMPLEX_TOOLS.has(fc.name)) {
          writeSuccessCount++;
          roundsSinceWrite = 0;
          writeOps.push({ name: fc.name, args: fc.args, round });
        }

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
          // Don't bail — tell the model to try a DIFFERENT approach.
          conversation.push({
            role: 'user',
            content: I18n.lang === 'es'
              ? `SYSTEM: La herramienta "${fc.name}" ha fallado ${consecutiveErrors} veces con error: ${toolResult.error}. NO insistas con la misma herramienta y los mismos argumentos. Prueba un enfoque diferente: usa find_in_workbook para localizar la hoja/rango correcto, o usa read_range con un rango más pequeño, o crea una hoja nueva y procede con lo que ya sabes. Tienes datos leídos en tu historial — úsalos. El usuario quiere un resultado, no una explicación de por qué falló.`
              : `SYSTEM: Tool "${fc.name}" has failed ${consecutiveErrors} times with error: ${toolResult.error}. DO NOT retry the same tool with the same arguments. Try a different approach: use find_in_workbook to locate the correct sheet/range, or use read_range with a smaller range, or create a new sheet and proceed with what you already know. You have data in your history — use it. The user wants a result, not an explanation of why it failed.`
          });
          consecutiveErrors = 0;
          lastErrorTool = null;
        }
      }

      if (aborted) break;

      // Trim large tool results in older history to control context growth.
      this.trimHistory(conversation);

      // Stale detection: if the model has been reading for too many rounds
      // without writing anything, nudge it to start building — and remind
      // it what the user actually asked for so it doesn't lose the goal.
      roundsSinceWrite++;
      if (writeSuccessCount === 0 && roundsSinceWrite >= this.STALE_ROUNDS && round < this.MAX_ROUNDS) {
        conversation.push({
          role: 'user',
          content: I18n.lang === 'es'
            ? `SYSTEM: Has pasado ${roundsSinceWrite} rondas leyendo datos sin escribir nada. Ya tienes suficiente información. DEJA de leer y EMPIEZA a construir AHORA. El usuario pidió: "${userText}". Crea la hoja, escribe los KPIs con fórmulas, formatea, y crea los gráficos con los datos que ya leíste. Si falta algún dato, procede con lo que tienes y pide el resto al final.`
            : `SYSTEM: You have spent ${roundsSinceWrite} rounds reading data without writing anything. You have enough information. STOP reading and START building NOW. The user asked: "${userText}". Create the sheet, write KPIs with formulas, format, and create charts with the data you already read. If some data is missing, proceed with what you have and ask for the rest in your final answer.`
        });
        roundsSinceWrite = 0; // reset to avoid repeated nudges
      }

      // Periodic reminder: re-inject the original request so the model
      // stays focused on the user's actual goal instead of wandering.
      if (round > 0 && round % this.REMINDER_ROUNDS === 0 && round < this.MAX_ROUNDS) {
        conversation.push({
          role: 'user',
          content: I18n.lang === 'es'
            ? `SYSTEM: Recordatorio — la solicitud original del usuario era: "${userText}". Asegúrate de estar avanzando hacia ese objetivo, no explorando datos sin propósito.`
            : `SYSTEM: Reminder — the user's original request was: "${userText}". Make sure you are progressing toward that goal, not exploring data aimlessly.`
        });
      }

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

    // If the model returned no text (or just "Done"), construct a meaningful
    // summary based on what was actually done during the run.
    if (!finalText || finalText.trim().length < 5) {
      if (writeOps.length === 0) {
        finalText = I18n.lang === 'es'
          ? 'No pude completar la operación. ' + (toolErrors.length > 0
            ? `Encontré ${toolErrors.length} errores al intentar usar las herramientas. Intenta reformular tu solicitud o verifica que los datos existan.`
            : 'No se realizaron cambios en el libro. Intenta reformular tu solicitud.')
          : 'I could not complete the operation. ' + (toolErrors.length > 0
            ? `I encountered ${toolErrors.length} errors while trying to use tools. Try rephrasing your request or verify the data exists.`
            : 'No changes were made to the workbook. Try rephrasing your request.');
      } else {
        // Construct a summary of what was built
        const sheets = writeOps.filter(o => o.name === 'add_sheet').map(o => o.args?.name).filter(Boolean);
        const writes = writeOps.filter(o => o.name === 'write_range').length;
        const formats = writeOps.filter(o => o.name === 'format_range').length;
        const charts = writeOps.filter(o => o.name === 'create_chart').length;
        const tables = writeOps.filter(o => o.name === 'create_table').length;
        const pivots = writeOps.filter(o => o.name === 'create_pivot').length;

        const parts = [];
        if (sheets.length) parts.push(sheets.length === 1 ? `hoja "${sheets[0]}"` : `hojas: ${sheets.join(', ')}`);
        if (writes) parts.push(`${writes} escrituras`);
        if (formats) parts.push(`${formats} formatos`);
        if (charts) parts.push(`${charts} gráficos`);
        if (tables) parts.push(`${tables} tablas`);
        if (pivots) parts.push(`${pivots} tablas dinámicas`);

        const es = I18n.lang === 'es';
        if (parts.length > 0) {
          finalText = es
            ? `Completé la operación: ${parts.join(', ')}. `
            : `Completed: ${parts.map(p => p.replace('escrituras','writes').replace('formatos','formats').replace('gráficos','charts').replace('tablas','tables').replace('tablas dinámicas','pivot tables')).join(', ')}. `;
        } else {
          finalText = es ? 'Operación completada.' : 'Operation completed.';
        }

        // If we hit the round limit, note it
        if (lastRound >= this.MAX_ROUNDS && !aborted) {
          finalText += es
            ? ' \u26A0 Se alcanzó el límite de rondas — el resultado puede estar incompleto. Puedes pedirme que continúe.'
            : ' \u26A0 Hit the round limit — the result may be incomplete. You can ask me to continue.';
        }

        // If there were errors, mention count
        if (toolErrors.length > 0) {
          finalText += es
            ? ` \u26A0 ${toolErrors.length} herramientas fallaron durante la construcción.`
            : ` \u26A0 ${toolErrors.length} tools failed during the build.`;
        }
      }
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
