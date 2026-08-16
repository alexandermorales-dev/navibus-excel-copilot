/* ============================================
   gemini.js — Gemini API client
   Free-tier aware: 429 handling, retryDelay, model fallback
   Uses streamGenerateContent for live thinking display
   ============================================ */

const Gemini = {
  BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',

  /**
   * Call Gemini streamGenerateContent endpoint.
   * onThinking(text) is called live as thinking chunks arrive.
   * Returns: { ok: true, text, thinking } | { ok: false, error, retry }
   */
  async generate(model, systemPrompt, contents, maxRetries = 4, onThinking) {
    const url = `${this.BASE_URL}/${model}:streamGenerateContent?alt=sse&key=${Config.apiKey}`;
    const body = {
      systemInstruction: { parts: [{ text: systemPrompt }] },
      contents: contents,
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 8192,
        responseMimeType: 'application/json',
        thinkingConfig: {
          includeThoughts: true,
          thinkingLevel: 'medium'
        }
      }
    };

    let lastError = '';
    let lastErrorType = 'unknown';

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 180000);

        const resp = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body),
          signal: controller.signal
        });
        clearTimeout(timeoutId);

        if (resp.ok) {
          // Parse SSE stream
          const result = await this.parseStream(resp, onThinking);
          return result;
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
            await this.sleep(waitSec * 1000);
            App.hideStatus();
          }
          continue;
        }

        if (resp.status >= 500) {
          lastError = I18n.tf('serverError', resp.status);
          lastErrorType = 'server';
          if (attempt < maxRetries - 1) {
            await this.sleep(this.backoffSeconds(3, attempt) * 1000);
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
          lastError = I18n.t('timeout');
          lastErrorType = 'timeout';
        } else {
          lastError = I18n.tf('networkError', e.message);
          lastErrorType = 'network';
        }
        if (attempt < maxRetries - 1) {
          await this.sleep(this.backoffSeconds(2, attempt) * 1000);
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
   * Exponential backoff with jitter: base * 2^attempt, +/- 20% random jitter.
   * Returns seconds.
   */
  backoffSeconds(base, attempt) {
    const raw = base * Math.pow(2, attempt);
    const jitter = raw * (0.8 + Math.random() * 0.4); // 80%-120% of raw
    return Math.round(jitter);
  },

  /**
   * Parse SSE stream from streamGenerateContent.
   * Calls onThinking(chunk) as thinking parts arrive.
   * Returns: { ok: true, text, thinking }
   */
  async parseStream(resp, onThinking) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    let fullThinking = '';

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
              // This is a thinking chunk — send it live
              fullThinking += part.text;
              if (onThinking) onThinking(part.text, fullThinking);
            } else if (part.text) {
              // This is regular output
              fullText += part.text;
            }
          }
        } catch (e) {
          // Incomplete JSON in chunk — skip, will be completed in next iteration
        }
      }
    }

    return { ok: true, text: fullText, thinking: fullThinking.trim() };
  },

  /**
   * Try primary model, then fallback model if it fails with a non-quota error.
   */
  async generateWithFallback(systemPrompt, contents, onThinking) {
    let result = await this.generate(Config.model, systemPrompt, contents, 4, onThinking);

    if (!result.ok && result.retry !== false && Config.model !== Config.fallbackModel) {
      App.showStatus(I18n.t('fallbackModel'));
      result = await this.generate(Config.fallbackModel, systemPrompt, contents, 2, onThinking);
      App.hideStatus();
    }

    return result;
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
  },

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
};
