/* ============================================
   llm.js — OpenAI-compatible chat client

   One client for every provider: Gemini, Groq and OpenRouter all expose
   the same /chat/completions shape, so only the base URL, auth header
   and model id vary. Routing and failover are delegated to Quota.

   On a retryable failure the call is retried against a DIFFERENT
   provider rather than sleeping, which is strictly better than waiting
   whenever a second key is configured.
   ============================================ */

const LLM = {
  REQUEST_TIMEOUT_MS: 180000,
  MAX_WAIT_MS: 65000,        // longest we'll sleep for a rate-limit window
  MAX_ATTEMPTS: 4,           // total provider attempts per logical call

  /**
   * @param {object} opts
   * @param {string}  opts.role           — 'plan' | 'answer' | 'repair'
   * @param {string}  opts.systemPrompt
   * @param {Array}   opts.messages       — OpenAI messages array
   * @param {boolean} [opts.json]         — request a strict JSON object back
   * @param {number}  [opts.maxTokens]
   * @param {number}  [opts.temperature]
   * @param {Array}   [opts.tools]        — optional tool declarations
   * @param {AbortSignal} [opts.signal]
   * @param {function} [opts.onText]      — (chunk, full)
   * @param {function} [opts.onThinking]  — (fullReasoning)
   * @param {function} [opts.onProvider]  — (providerLabel, model, isFailover)
   *
   * Returns { ok, text, reasoning, functionCalls, usage, providerId, model }
   *      or { ok:false, error, errorType }
   */
  async chat(opts) {
    const {
      role = 'plan', systemPrompt, messages, json = false,
      maxTokens = 8192, temperature = 0.3, tools,
      signal, onText, onThinking, onProvider
    } = opts;

    const estTokens = Quota.estimateTokens(systemPrompt) +
                      Quota.estimateTokens(messages) +
                      Math.min(maxTokens, 2048);

    const tried = [];
    let lastError = 'No provider available';
    let lastErrorType = 'quota';

    for (let attempt = 0; attempt < this.MAX_ATTEMPTS; attempt++) {
      if (signal && signal.aborted) {
        return { ok: false, error: I18n.t('aborted'), errorType: 'aborted' };
      }

      const route = Quota.pick(role, estTokens);
      if (!route) {
        return {
          ok: false,
          errorType: tried.length > 0 ? lastErrorType : 'quota',
          error: tried.length > 0 ? lastError : I18n.t('allQuotaSpent')
        };
      }

      if (route.waitMs > 0) {
        // pick() only returns a wait when no provider is immediately free.
        if (route.waitMs > this.MAX_WAIT_MS) {
          return { ok: false, error: I18n.t('allQuotaSpent'), errorType: 'quota' };
        }
        const secs = Math.ceil(route.waitMs / 1000);
        this._status(I18n.tf('rateLimit', secs));
        await this.sleep(route.waitMs, signal);
        this._hideStatus();
        if (signal && signal.aborted) {
          return { ok: false, error: I18n.t('aborted'), errorType: 'aborted' };
        }
      }

      if (onProvider) onProvider(route.provider.label, route.model, tried.length > 0);
      tried.push(route.providerId);

      const result = await this._callOnce({
        route, systemPrompt, messages, json, maxTokens, temperature, tools,
        signal, onText, onThinking
      });

      if (result.ok) {
        Quota.reward(route.providerId);
        Quota.record(route.providerId, { tokens: result.usage?.total_tokens || estTokens });
        return { ...result, providerId: route.providerId, model: route.model };
      }

      // A failed request still consumes free-tier allowance, except when
      // it never reached the provider (abort / local network failure).
      if (result.errorType !== 'aborted' && result.errorType !== 'network') {
        Quota.record(route.providerId, { tokens: result.usage?.total_tokens || 0 });
      }

      if (result.errorType === 'aborted') return result;

      Quota.penalize(route.providerId, {
        status: result.status,
        retryAfterSec: result.retryAfterSec
      });

      lastError = result.error;
      lastErrorType = result.errorType;
      console.warn(`LLM: ${route.providerId} failed (${result.errorType}): ${result.error}`);
    }

    return { ok: false, error: lastError, errorType: lastErrorType };
  },

  /**
   * A single request against one provider.
   */
  async _callOnce({ route, systemPrompt, messages, json, maxTokens, temperature, tools, signal, onText, onThinking }) {
    const { provider, model } = route;
    const key = Config.keyFor(provider.id);
    if (!key) return { ok: false, error: 'No key configured', errorType: 'client' };

    const body = {
      model,
      messages: [{ role: 'system', content: systemPrompt }, ...messages],
      temperature,
      max_tokens: maxTokens,
      stream: true,
      stream_options: { include_usage: true }
    };
    if (json && provider.jsonMode) body.response_format = { type: 'json_object' };
    if (tools && tools.length > 0) {
      body.tools = tools;
      body.tool_choice = 'auto';
    }

    let attemptBody = body;

    for (let pass = 0; pass < 2; pass++) {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.REQUEST_TIMEOUT_MS);
      const onExternalAbort = () => controller.abort();
      if (signal) signal.addEventListener('abort', onExternalAbort);

      try {
        const resp = await fetch(provider.chatUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${key}`,
            ...(provider.extraHeaders || {})
          },
          body: JSON.stringify(attemptBody),
          signal: controller.signal
        });

        if (resp.ok) {
          const parsed = await this.parseStream(resp, onText, onThinking, signal);
          return parsed;
        }

        const errText = await resp.text().catch(() => '');

        // Some providers/models reject response_format or stream_options.
        // Strip the optional fields and retry once before giving up.
        if (resp.status === 400 && pass === 0 && (attemptBody.response_format || attemptBody.stream_options)) {
          const retry = { ...attemptBody };
          delete retry.response_format;
          delete retry.stream_options;
          attemptBody = retry;
          console.warn(`LLM: ${provider.id} rejected optional fields, retrying without them`);
          continue;
        }

        return {
          ok: false,
          status: resp.status,
          retryAfterSec: this._retryAfter(resp),
          error: this._describeHttp(resp.status, errText, provider),
          errorType: this._classify(resp.status)
        };
      } catch (e) {
        if (e.name === 'AbortError') {
          if (signal && signal.aborted) {
            return { ok: false, error: I18n.t('aborted'), errorType: 'aborted' };
          }
          return { ok: false, error: I18n.t('timeout'), errorType: 'timeout' };
        }
        return { ok: false, error: I18n.tf('networkError', e.message), errorType: 'network' };
      } finally {
        clearTimeout(timeoutId);
        if (signal) signal.removeEventListener('abort', onExternalAbort);
      }
    }

    return { ok: false, error: I18n.t('unknownError'), errorType: 'unknown' };
  },

  _retryAfter(resp) {
    const h = resp.headers.get('Retry-After') || resp.headers.get('retry-after');
    if (!h) return null;
    const n = parseInt(h, 10);
    return isNaN(n) ? null : n;
  },

  _classify(status) {
    if (status === 429) return 'rate_limit';
    if (status === 402) return 'quota';
    if (status === 401 || status === 403) return 'auth';
    if (status >= 500) return 'server';
    return 'client';
  },

  _describeHttp(status, body, provider) {
    const snippet = String(body || '').slice(0, 200);
    if (status === 401 || status === 403) return I18n.tf('authFailed', provider.label);
    if (status === 402) return I18n.tf('providerQuota', provider.label);
    if (status === 429) return I18n.tf('providerRateLimited', provider.label);
    if (status >= 500) return I18n.tf('serverError', `${provider.label} ${status}`);
    return `${provider.label} HTTP ${status}${snippet ? ': ' + snippet : ''}`;
  },

  /**
   * Parse an SSE stream, accumulating text, reasoning, tool calls and usage.
   */
  async parseStream(resp, onText, onThinking, signal) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    let fullReasoning = '';
    let finishReason = null;
    let usage = null;
    const toolCallMap = {};

    while (true) {
      if (signal && signal.aborted) break;
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;
        const jsonStr = trimmed.slice(5).trim();
        if (!jsonStr || jsonStr === '[DONE]') continue;

        let chunk;
        try {
          chunk = JSON.parse(jsonStr);
        } catch (e) {
          continue; // partial JSON — completed on a later read
        }

        if (chunk.usage) usage = chunk.usage;

        const choice = chunk.choices && chunk.choices[0];
        if (!choice) continue;
        if (choice.finish_reason) finishReason = choice.finish_reason;

        const delta = choice.delta;
        if (!delta) continue;

        if (delta.content) {
          fullText += delta.content;
          if (onText) onText(delta.content, fullText);
        }

        const reasoningChunk = delta.reasoning || delta.reasoning_content;
        if (reasoningChunk) {
          fullReasoning += reasoningChunk;
          if (onThinking) onThinking(fullReasoning);
        }

        if (delta.tool_calls) {
          for (const tc of delta.tool_calls) {
            const idx = tc.index ?? 0;
            if (!toolCallMap[idx]) {
              toolCallMap[idx] = { id: tc.id || '', type: 'function', function: { name: '', arguments: '' } };
            }
            if (tc.id) toolCallMap[idx].id = tc.id;
            if (tc.function && tc.function.name) toolCallMap[idx].function.name += tc.function.name;
            if (tc.function && tc.function.arguments) toolCallMap[idx].function.arguments += tc.function.arguments;
          }
        }
      }
    }

    const functionCalls = Object.keys(toolCallMap)
      .sort((a, b) => parseInt(a, 10) - parseInt(b, 10))
      .map(idx => {
        const tc = toolCallMap[idx];
        let args = {};
        try {
          args = tc.function.arguments ? JSON.parse(tc.function.arguments) : {};
        } catch (e) {
          console.warn('LLM: unparseable tool arguments', tc.function.arguments);
        }
        return { id: tc.id, name: tc.function.name, args };
      });

    return {
      ok: true,
      text: fullText,
      reasoning: fullReasoning,
      functionCalls,
      usage,
      truncated: finishReason === 'length'
    };
  },

  /**
   * Extract a JSON object from a model response.
   *
   * Free-tier models frequently wrap JSON in prose or fenced code blocks,
   * emit trailing commas, or use smart quotes. Rather than burning a retry
   * (and a request) on formatting, recover locally where possible.
   *
   * Returns { ok, value } or { ok:false, error }
   */
  extractJSON(text) {
    if (!text || typeof text !== 'string') {
      return { ok: false, error: 'Empty response' };
    }

    const candidates = [];

    // 1. Fenced code block (```json ... ``` or ``` ... ```)
    const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) candidates.push(fence[1]);

    // 2. The raw text as-is.
    candidates.push(text);

    // 3. Widest brace-balanced span, for JSON embedded in prose.
    const span = this._braceSpan(text);
    if (span) candidates.push(span);

    for (const raw of candidates) {
      for (const attempt of [raw, this._repairJSON(raw)]) {
        const trimmed = (attempt || '').trim();
        if (!trimmed) continue;
        try {
          const value = JSON.parse(trimmed);
          if (value && typeof value === 'object') return { ok: true, value };
        } catch (e) { /* try next candidate */ }
      }
    }

    return { ok: false, error: 'No valid JSON object found in response' };
  },

  /**
   * Widest {...} span with balanced braces, ignoring braces inside strings.
   */
  _braceSpan(text) {
    const start = text.indexOf('{');
    if (start === -1) return null;
    let depth = 0;
    let inString = false;
    let escaped = false;
    for (let i = start; i < text.length; i++) {
      const ch = text[i];
      if (escaped) { escaped = false; continue; }
      if (ch === '\\') { escaped = true; continue; }
      if (ch === '"') { inString = !inString; continue; }
      if (inString) continue;
      if (ch === '{') depth++;
      else if (ch === '}') {
        depth--;
        if (depth === 0) return text.slice(start, i + 1);
      }
    }
    // Unterminated — return the remainder so _repairJSON can try to close it.
    return text.slice(start);
  },

  /**
   * Best-effort repair of the malformations small models actually produce.
   */
  _repairJSON(raw) {
    if (!raw) return raw;
    let s = String(raw).trim();

    // Strip a stray leading fence marker.
    s = s.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();

    // Normalize smart quotes outside of intended content.
    s = s.replace(/[\u201c\u201d]/g, '"').replace(/[\u2018\u2019]/g, "'");

    // Remove trailing commas before a closing brace/bracket.
    s = s.replace(/,(\s*[}\]])/g, '$1');

    // Close unbalanced braces/brackets (truncated output).
    const counts = this._countDelims(s);
    if (counts.curly > 0 || counts.square > 0) {
      // Drop a dangling incomplete key/value fragment before closing.
      s = s.replace(/,\s*"[^"]*"\s*:?\s*$/, '').replace(/,\s*$/, '');
      s += ']'.repeat(Math.max(0, counts.square)) + '}'.repeat(Math.max(0, counts.curly));
    }

    return s;
  },

  _countDelims(s) {
    let curly = 0, square = 0, inString = false, escaped = false;
    for (const ch of s) {
      if (escaped) { escaped = false; continue; }
      if (ch === '\\') { escaped = true; continue; }
      if (ch === '"') { inString = !inString; continue; }
      if (inString) continue;
      if (ch === '{') curly++;
      else if (ch === '}') curly--;
      else if (ch === '[') square++;
      else if (ch === ']') square--;
    }
    return { curly, square };
  },

  sleep(ms, signal) {
    return new Promise(resolve => {
      const start = Date.now();
      const tick = () => {
        if (signal && signal.aborted) return resolve();
        const elapsed = Date.now() - start;
        if (elapsed >= ms) return resolve();
        setTimeout(tick, Math.min(250, ms - elapsed));
      };
      tick();
    });
  },

  _status(msg) {
    if (typeof App !== 'undefined' && App.showStatus) App.showStatus(msg);
  },

  _hideStatus() {
    if (typeof App !== 'undefined' && App.hideStatus) App.hideStatus();
  }
};
