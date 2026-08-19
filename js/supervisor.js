/* ============================================
   supervisor.js — Gemini-powered supervisor
   Periodically reviews the agent's progress and
   provides guidance to keep it on track.

   Uses gemini-3.1-flash-lite (free tier) via the
   OpenAI-compatible endpoint to save tokens: the
   supervisor only sees a COMPRESSED summary of
   what the agent has done, not raw tool outputs.
   ============================================ */

const Supervisor = {
  BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
  MODEL: 'gemini-3.1-flash-lite',

  /**
   * Ask Gemini to review the agent's progress and provide guidance.
   *
   * @param {object} opts
   * @param {string} opts.userRequest  — the original user message
   * @param {Array}  opts.actions      — compact list of {tool, detail, ok} done so far
   * @param {number} opts.round        — current round number
   * @param {number} opts.maxRounds    — max rounds allowed
   * @param {string} opts.lang         — 'es' or 'en'
   * @returns {Promise<string>}        — guidance text, or '' if unavailable
   */
  async review(opts) {
    const { userRequest, actions, round, maxRounds, lang } = opts;

    const key = Config.geminiApiKey;
    if (!key) return ''; // no Gemini key → skip supervision silently

    // Build a compact action log — this is the token-saving trick.
    // We send ~1 line per action, not the full tool output arrays.
    const actionLog = actions.map((a, i) =>
      `${i + 1}. ${a.tool}(${a.detail}) → ${a.ok ? 'ok' : 'error: ' + (a.error || 'unknown')}`
    ).join('\n');

    const es = lang === 'es';
    const systemPrompt = es
      ? `Eres un supervisor experto que guía a un agente AI que construye paneles en Excel. Revisas el progreso y das orientación concisa. Respondes en español. Sé breve: máximo 3-4 líneas. Identifica: (1) si va por buen camino, (2) qué falta por hacer según la solicitud original, (3) si se está desviando o atascando. NO des instrucciones genéricas — sé específico sobre qué hojas crear, qué datos faltan, qué gráficos o tablas aún se necesitan.`
      : `You are an expert supervisor guiding an AI agent that builds Excel dashboards. You review progress and give concise guidance. Be brief: max 3-4 lines. Identify: (1) whether it's on track, (2) what's still missing per the original request, (3) whether it's wandering or stuck. Do NOT give generic instructions — be specific about which sheets to create, what data is missing, what charts or tables are still needed.`;

    const userPrompt = es
      ? `Solicitud original del usuario: "${userRequest}"

Ronda actual: ${round} de ${maxRounds}

Acciones realizadas hasta ahora:
${actionLog || '(ninguna aún)'}

¿Va por buen camino? ¿Qué falta? Responde en máximo 4 líneas, sé específico.`
      : `Original user request: "${userRequest}"

Current round: ${round} of ${maxRounds}

Actions taken so far:
${actionLog || '(none yet)'}

Is it on track? What's missing? Answer in max 4 lines, be specific.`;

    try {
      const resp = await fetch(this.BASE_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${key}`
        },
        body: JSON.stringify({
          model: this.MODEL,
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userPrompt }
          ],
          temperature: 0.3,
          max_tokens: 300
        })
      });

      if (!resp.ok) {
        console.warn(`Supervisor: Gemini returned ${resp.status}`);
        return '';
      }

      const data = await resp.json();
      const text = data.choices?.[0]?.message?.content || '';
      return text.trim();
    } catch (e) {
      console.warn('Supervisor: review failed:', e.message);
      return '';
    }
  }
};
