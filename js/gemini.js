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

    const body = {
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: contents,
      generationConfig: {
        temperature: 0.4,
        maxOutputTokens: 16384,
        thinkingConfig: {
          includeThoughts: true,
          thinkingLevel: 'low'
        }
      }
    };
    if (tools && tools.length > 0) body.tools = tools;
    if (toolConfig) body.toolConfig = toolConfig;

    let lastError = '';
    let lastErrorType = 'unknown';

    // Key selection strategy:
    //   1. If the user has their own API key, try it FIRST on every request.
    //   2. On 429 from the user's key → mark it exhausted for this request,
    //      fall back to round-robin among the HARDCODED keys only.
    //   3. On 400/403 from the user's key (invalid/disabled) → auto-clear it
    //      from Settings + localStorage, notify once, then use hardcoded keys.
    //   4. Hardcoded keys round-robin among themselves on 429s.
    const userKey = (Config.apiKey && Config.apiKey.length > 0) ? Config.apiKey : '';
    let userKeyExhausted = false; // 429'd during this request — skip for rest of request
    const hardcodedPool = Config.hardcodedKeys.filter(k => k);
    const effectiveMaxRetries = Math.max(maxRetries, hardcodedPool.length + (userKey ? 3 : 2));

    for (let attempt = 0; attempt < effectiveMaxRetries; attempt++) {
      // Honor an external abort (user clicked Stop) between retries too.
      if (signal && signal.aborted) {
        return { ok: false, error: I18n.t('aborted'), errorType: 'aborted', retry: false };
      }

      // Pick the key: user's key first (if available and not exhausted this
      // request), otherwise round-robin among the hardcoded keys.
      let apiKey = '';
      let isUserKey = false;
      if (userKey && !userKeyExhausted) {
        apiKey = userKey;
        isUserKey = true;
      } else {
        apiKey = Config.nextHardcodedKey();
      }
      if (!apiKey) {
        return { ok: false, error: 'No API key available.', errorType: 'client', retry: false };
      }
      const url = `${this.BASE_URL}/${model}:streamGenerateContent?alt=sse&key=${apiKey}`;

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

          console.log(`429: ${isUserKey ? 'user key' : 'hardcoded key'}, daily=${isDaily}, delay=${delay}s, body=${this.truncate(errBody, 200)}`);

          if (isUserKey) {
            // User's key hit rate limit — don't try it again this request;
            // fall back to the hardcoded pool immediately.
            userKeyExhausted = true;
            continue;
          }

          if (attempt < effectiveMaxRetries - 1) {
            if (hardcodedPool.length > 1) {
              // More hardcoded keys to try — rotate immediately without waiting.
              continue;
            }
            // Only one key — must wait for the rate limit to clear.
            const waitSec = delay > 0 ? Math.min(delay, 65) : this.backoffSeconds(20, attempt);
            lastError = I18n.tf('rateLimit', waitSec);
            lastErrorType = isDaily ? 'quota_daily' : 'rate_limit';
            App.showStatus(I18n.tf('rateLimit', waitSec));
            await this.sleep(waitSec * 1000, signal);
            App.hideStatus();
            continue;
          }

          // Exhausted all retries across all keys.
          lastError = isDaily ? I18n.t('dailyQuota') : I18n.tf('rateLimit', delay > 0 ? delay : 20);
          lastErrorType = isDaily ? 'quota_daily' : 'rate_limit';
          return {
            ok: false,
            error: lastError,
            errorType: lastErrorType,
            retry: false
          };
        }

        // 400/403: invalid or disabled key.
        if (resp.status === 400 || resp.status === 403) {
          const errText = await resp.text();
          if (isUserKey) {
            // User's key is invalid/disabled — clear it, notify, fall back.
            console.log(`${resp.status} on user key — clearing and falling back to hardcoded keys. body=${this.truncate(errText, 200)}`);
            Config.clearUserKey();
            userKeyExhausted = true; // prevents re-reading the now-empty Config.apiKey
            App.showStatus(I18n.lang === 'es'
              ? 'Tu API key es inválida — se eliminó de Configuración. Usando claves integradas.'
              : 'Your API key is invalid — cleared from Settings. Using built-in keys.');
            setTimeout(() => App.hideStatus(), 5000);
            continue;
          }
          // Hardcoded key 400/403 — fail fast (bad request or disabled key).
          return {
            ok: false,
            error: `HTTP ${resp.status}: ${this.truncate(errText, 300)}`,
            errorType: 'client',
            retry: false
          };
        }

        if (resp.status >= 500) {
          lastError = I18n.tf('serverError', resp.status);
          lastErrorType = 'server';
          if (attempt < effectiveMaxRetries - 1) {
            await this.sleep(this.backoffSeconds(3, attempt) * 1000, signal);
          }
          continue;
        }

        // Other 4xx: fail fast — not retryable
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
        if (attempt < effectiveMaxRetries - 1) {
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
   * Aggregates text, thinking, and functionCall parts.
   *
   * CRITICAL: preserves thoughtSignature on parts — Gemini 3 requires it
   * to be echoed back in conversation history on the next turn, or the
   * API returns HTTP 400. The signature is at the Part level (alongside
   * functionCall/text), NOT inside the functionCall object.
   *
   * Returns: { ok: true, text, thinking, functionCalls, rawParts, finishReason }
   *   rawParts = the original Part objects from the model response, to be
   *   pushed into conversation history verbatim (preserving thoughtSignature).
   *   finishReason = the model's stop reason (e.g. 'STOP', 'MAX_TOKENS').
   *   Used by the agent to detect truncated responses and auto-continue.
   */
  async parseStream(resp, onThinking, onText) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    let fullThinking = '';
    let finishReason = null;

    const functionCalls = [];
    const rawParts = []; // preserve original parts for conversation history

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
          const candidate = chunk.candidates?.[0];
          // Capture finishReason — the last chunk carries it. 'MAX_TOKENS'
          // means the response was truncated; the agent uses this to
          // auto-continue instead of treating partial text as final.
          if (candidate?.finishReason) {
            finishReason = candidate.finishReason;
          }
          const parts = candidate?.content?.parts || [];

          for (const part of parts) {
            if (part.thought && part.text) {
              // Thinking part — stream it live, but DON'T add to rawParts
              // (thinking parts should not be echoed back in history).
              fullThinking += part.text;
              if (onThinking) onThinking(part.text, fullThinking);
            } else if (part.functionCall) {
              // functionCall part — preserve the ENTIRE original part object
              // so thoughtSignature (at the part level) is kept intact.
              rawParts.push(part);
              functionCalls.push({
                name: part.functionCall.name,
                args: part.functionCall.args || {}
              });
            } else if (part.text) {
              // Regular text part — accumulate and add to rawParts.
              fullText += part.text;
              if (onText) onText(part.text, fullText);
              rawParts.push({ text: part.text });
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
      functionCalls,
      rawParts,
      finishReason
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
    // Only match explicit daily quota markers in Google's error format.
    // Avoid matching the word "daily" in generic error messages.
    return /"quotaType"\s*:\s*"PerDay"|perDay|RESOURCE_EXHAUSTED.*daily/i.test(errBody);
  },

  truncate(s, n) {
    return s.length > n ? s.substring(0, n) + '...' : s;
  }
};
