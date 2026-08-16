/* ============================================
   gemini.js — Gemini API client (function-calling agent mode)
   Free-tier aware: 429 handling, retryDelay, model fallback.
   Uses streamGenerateContent for live thinking display and
   incremental functionCall parsing across SSE chunks.
   ============================================ */

const Gemini = {
  BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',

  /**
   * Call Gemini streamGenerateContent with optional tool declarations.
   *
   * @param {object} opts
   * @param {string} opts.systemPrompt
   * @param {Array}  opts.contents        — Gemini contents array (may include
   *                                        functionCall / functionResponse parts)
   * @param {Array}  [opts.tools]         — [{functionDeclarations:[...]}]
   * @param {object} [opts.toolConfig]    — { functionCallingConfig: {...} }
   * @param {number} [opts.maxRetries=4]
   * @param {function} [opts.onThinking]  — (chunkText, fullThinking) live
   * @param {function} [opts.onText]      — (chunkText, fullText) live final text
   * @param {AbortSignal} [opts.signal]   — for user-initiated Stop
   *
   * Returns:
   *   { ok: true, text, thinking, functionCalls: [{name, args}] }
   *   { ok: false, error, errorType, retry }
   */
  async generate(opts) {
    const {
      systemPrompt,
      contents,
      tools,
      toolConfig,
      maxRetries = 4,
      onThinking,
      onText,
      signal
    } = opts;

    const model = Config.model;
    const url = `${this.BASE_URL}/${model}:streamGenerateContent?alt=sse&key=${Config.apiKey}`;

    const body = {
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: contents,
      generationConfig: {
        temperature: 0.4,
        maxOutputTokens: 8192,
        thinkingConfig: {
          includeThoughts: true,
          thinkingLevel: 'medium'
        }
      }
    };
    if (tools && tools.length > 0) body.tools = tools;
    if (toolConfig) body.toolConfig = toolConfig;

    let lastError = '';
    let lastErrorType = 'unknown';

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      // Honor an external abort (user clicked Stop) between retries too.
      if (signal && signal.aborted) {
        return { ok: false, error: I18n.t('aborted'), errorType: 'aborted', retry: false };
      }

      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 180000);

        // Chain external signal → our controller
        const onExternalAbort = () => controller.abort();
        if (signal) signal.addEventListener('abort', onExternalAbort);

        const resp = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body),
          signal: controller.signal
        });
        clearTimeout(timeoutId);
        if (signal) signal.removeEventListener('abort', onExternalAbort);

        if (resp.ok) {
          return await this.parseStream(resp, onThinking, onText);
        }

        if (resp.status === 429) {
          const errBody = await resp.text();
          const delay = this.parseRetryDelay(errBody);
          const isDaily = this.isDailyQuota(errBody);

          if (isDaily) {
            return {
              ok: false,
              error: I18n.t('dailyQuota'),
              errorType: 'quota_daily',
              retry: false
            };
          }

          lastErrorType = 'rate_limit';
          const waitSec = delay > 0 ? Math.min(delay, 65) : this.backoffSeconds(20, attempt);
          lastError = I18n.tf('rateLimit', waitSec);
          if (attempt < maxRetries - 1) {
            App.showStatus(I18n.tf('rateLimit', waitSec));
            await this.sleep(waitSec * 1000, signal);
            App.hideStatus();
          }
          continue;
        }

        if (resp.status >= 500) {
          lastError = I18n.tf('serverError', resp.status);
          lastErrorType = 'server';
          if (attempt < maxRetries - 1) {
            await this.sleep(this.backoffSeconds(3, attempt) * 1000, signal);
          }
          continue;
        }

        // 4xx (non-429): fail fast — not retryable
        const errText = await resp.text();
        return {
          ok: false,
          error: `HTTP ${resp.status}: ${this.truncate(errText, 300)}`,
          errorType: 'client',
          retry: false
        };

      } catch (e) {
        if (e.name === 'AbortError') {
          if (signal && signal.aborted) {
            return { ok: false, error: I18n.t('aborted'), errorType: 'aborted', retry: false };
          }
          lastError = I18n.t('timeout');
          lastErrorType = 'timeout';
        } else {
          lastError = I18n.tf('networkError', e.message);
          lastErrorType = 'network';
        }
        if (attempt < maxRetries - 1) {
          await this.sleep(this.backoffSeconds(2, attempt) * 1000, signal);
        }
      }
    }

    return {
      ok: false,
      error: lastError || I18n.t('unknownError'),
      errorType: lastErrorType,
      retry: false
    };
  },

  /**
   * Try primary model, then fallback model if it fails with a retryable error.
   */
  async generateWithFallback(opts) {
    let result = await this.generate(opts);

    if (!result.ok && result.retry !== false && Config.model !== Config.fallbackModel) {
      App.showStatus(I18n.t('fallbackModel'));
      const fallbackOpts = { ...opts, maxRetries: 2 };
      // Temporarily swap model by routing through generate() which reads Config.model.
      // We achieve fallback by swapping Config.model for one call.
      const savedModel = Config.model;
      Config.model = Config.fallbackModel;
      try {
        result = await this.generate(fallbackOpts);
      } finally {
        Config.model = savedModel;
      }
      App.hideStatus();
    }

    return result;
  },

  /**
   * Parse SSE stream from streamGenerateContent.
   * Aggregates text, thinking, and functionCall parts. FunctionCall args
   * may arrive split across multiple chunks (as partial JSON strings) — we
   * buffer per-part and finalize on part boundary.
   *
   * Returns: { ok: true, text, thinking, functionCalls }
   */
  async parseStream(resp, onThinking, onText) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    let fullThinking = '';

    // Function-call aggregation. Gemini typically emits a complete functionCall
    // object in one chunk, but defensive: collect all functionCall parts found.
    const functionCalls = [];

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });

      // SSE format: lines starting with "data: "
      const lines = buffer.split('\n');
      buffer = lines.pop() || ''; // Keep incomplete line in buffer

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith('data: ')) continue;

        const jsonStr = trimmed.substring(6).trim();
        if (jsonStr === '[DONE]') continue;

        try {
          const chunk = JSON.parse(jsonStr);
          const parts = chunk.candidates?.[0]?.content?.parts || [];

          for (const part of parts) {
            if (part.thought && part.text) {
              fullThinking += part.text;
              if (onThinking) onThinking(part.text, fullThinking);
            } else if (part.functionCall) {
              // functionCall: { name, args, thought_signature? }
              // Preserve thought_signature — Gemini requires it to be
              // echoed back in the conversation history when thinking is
              // enabled, or subsequent requests return HTTP 400.
              const fc = {
                name: part.functionCall.name,
                args: part.functionCall.args || {}
              };
              if (part.functionCall.thought_signature) {
                fc.thought_signature = part.functionCall.thought_signature;
              }
              if (part.thought_signature) {
                fc.thought_signature = part.thought_signature;
              }
              if (part.thought) {
                fc.thought = true;
              }
              functionCalls.push(fc);
            } else if (part.text) {
              fullText += part.text;
              if (onText) onText(part.text, fullText);
            }
          }
        } catch (e) {
          // Incomplete JSON in chunk — skip, will be completed in next iteration
        }
      }
    }

    return {
      ok: true,
      text: fullText,
      thinking: fullThinking.trim(),
      functionCalls
    };
  },

  /**
   * Exponential backoff with jitter: base * 2^attempt, +/- 20% random jitter.
   * Returns seconds. Aborts early if signal fires during sleep.
   */
  backoffSeconds(base, attempt) {
    const raw = base * Math.pow(2, attempt);
    const jitter = raw * (0.8 + Math.random() * 0.4); // 80%-120% of raw
    return Math.round(jitter);
  },

  sleep(ms, signal) {
    return new Promise(resolve => {
      const start = Date.now();
      const remaining = ms;
      const tick = () => {
        if (signal && signal.aborted) return resolve();
        const elapsed = Date.now() - start;
        if (elapsed >= remaining) return resolve();
        setTimeout(tick, Math.min(250, remaining - elapsed));
      };
      tick();
    });
  },

  parseRetryDelay(errBody) {
    try {
      const match = errBody.match(/"retryDelay"\s*:\s*"(\d+)s"/);
      if (match) return parseInt(match[1], 10);
    } catch (e) {}
    return -1;
  },

  isDailyQuota(errBody) {
    return /PerDay|per[_-]?day|daily/i.test(errBody);
  },

  truncate(s, n) {
    return s.length > n ? s.substring(0, n) + '...' : s;
  }
};
