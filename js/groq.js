/* ============================================
   groq.js — Groq API client via OpenRouter
   Uses OpenAI-compatible chat/completions with streaming
   for live text display and incremental tool_call parsing.
   ============================================ */

const Groq = {
  BASE_URL: 'https://openrouter.ai/api/v1/chat/completions',

  /**
   * Call Groq chat/completions with optional tool declarations.
   *
   * @param {object} opts
   * @param {string} opts.systemPrompt
   * @param {Array}  opts.messages       — OpenAI messages array
   * @param {Array}  [opts.tools]        — [{type:'function', function:{name,description,parameters}}]
   * @param {string} [opts.model]        — override model (defaults to Config.model or Config.capableModel)
   * @param {number} [opts.maxRetries=4]
   * @param {function} [opts.onText]      — (chunkText, fullText) live final text
   * @param {function} [opts.onThinking]  — (fullReasoning) live reasoning/thinking text
   * @param {AbortSignal} [opts.signal]   — for user-initiated Stop
   *
   * Returns:
   *   { ok: true, text, reasoning, functionCalls: [{id, name, args}], message, finishReason }
   *   { ok: false, error, errorType, retry }
   */
  async generate(opts) {
    const {
      systemPrompt,
      messages,
      tools,
      model,
      maxRetries = 4,
      onText,
      onThinking,
      signal
    } = opts;

    const useModel = model || Config.capableModel;

    // 8B model has a smaller output budget; 70B can use the full 16k.
    const maxTokens = useModel === Config.fastModel ? 8192 : 16384;

    const body = {
      model: useModel,
      messages: [
        { role: 'system', content: systemPrompt },
        ...messages
      ],
      temperature: 0.4,
      max_tokens: maxTokens,
      stream: true,
      // Request reasoning/thinking tokens from models that support them.
      // Ignored by non-reasoning models, so it's safe to always send.
      reasoning: { enabled: true, exclude: false }
    };
    if (tools && tools.length > 0) {
      body.tools = tools;
      body.tool_choice = 'auto';
    }

    let lastError = '';
    let lastErrorType = 'unknown';

    for (let attempt = 0; attempt < maxRetries; attempt++) {
      if (signal && signal.aborted) {
        return { ok: false, error: I18n.t('aborted'), errorType: 'aborted', retry: false };
      }

      const apiKey = Config.activeKey;
      console.log(`Groq: activeKey length=${apiKey ? apiKey.length : 0}, apiKey starts with=${apiKey ? apiKey.substring(0, 8) : 'none'}`);
      if (!apiKey) {
        return { ok: false, error: 'No API key configured. Open Settings and enter your Groq API key.', errorType: 'client', retry: false };
      }

      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 180000);

        const onExternalAbort = () => controller.abort();
        if (signal) signal.addEventListener('abort', onExternalAbort);

        const resp = await fetch(this.BASE_URL, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`,
            'HTTP-Referer': 'https://alexandermorales-dev.github.io/navibus-excel-copilot/',
            'X-Title': 'Excel AI Copilot'
          },
          body: JSON.stringify(body),
          signal: controller.signal
        });
        clearTimeout(timeoutId);
        if (signal) signal.removeEventListener('abort', onExternalAbort);

        if (resp.ok) {
          return await this.parseStream(resp, onText, onThinking, signal);
        }

        if (resp.status === 429) {
          const retryAfter = resp.headers.get('Retry-After');
          const delay = retryAfter ? parseInt(retryAfter, 10) : this.backoffSeconds(20, attempt);
          console.log(`429 rate limited, waiting ${delay}s`);
          lastError = I18n.tf('rateLimit', delay);
          lastErrorType = 'rate_limit';
          App.showStatus(I18n.tf('rateLimit', delay));
          if (attempt < maxRetries - 1) {
            await this.sleep(delay * 1000, signal);
            App.hideStatus();
            continue;
          }
          App.hideStatus();
          return { ok: false, error: lastError, errorType: lastErrorType, retry: false };
        }

        if (resp.status === 401 || resp.status === 403) {
          const errText = await resp.text();
          console.log(`${resp.status} auth error: ${this.truncate(errText, 300)}`);
          return {
            ok: false,
            error: `Authentication failed (${resp.status}). Check your API key in Settings. ${this.truncate(errText, 200)}`,
            errorType: 'client',
            retry: false
          };
        }

        if (resp.status === 402) {
          // Quota/credits exhausted — don't retry, tell the user clearly.
          const errText = await resp.text();
          console.log(`402 quota exceeded: ${this.truncate(errText, 300)}`);
          return {
            ok: false,
            error: I18n.t('quotaExceeded'),
            errorType: 'quota',
            retry: false
          };
        }

        if (resp.status >= 500) {
          lastError = I18n.tf('serverError', resp.status);
          lastErrorType = 'server';
          if (attempt < maxRetries - 1) {
            await this.sleep(this.backoffSeconds(3, attempt) * 1000, signal);
          }
          continue;
        }

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
   * Single-model — no fallback, just calls generate().
   * Kept for interface compatibility with the agent.
   */
  async generateWithFallback(opts) {
    return await this.generate(opts);
  },

  /**
   * Parse SSE stream from chat/completions.
   * Accumulates text content, reasoning/thinking, and tool_calls across chunks.
   *
   * Returns: { ok: true, text, reasoning, functionCalls, message, finishReason }
   *   message = the full assistant message object for conversation history
   *   finishReason = 'STOP' or 'MAX_TOKENS' (mapped from OpenAI format)
   */
  async parseStream(resp, onText, onThinking, signal) {
    const reader = resp.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullText = '';
    let fullReasoning = '';
    let finishReason = null;

    // Tool calls are accumulated by index across chunks
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
        if (!trimmed || !trimmed.startsWith('data: ')) continue;

        const jsonStr = trimmed.substring(6).trim();
        if (jsonStr === '[DONE]') continue;

        try {
          const chunk = JSON.parse(jsonStr);
          const choice = chunk.choices?.[0];
          if (!choice) continue;

          if (choice.finish_reason) {
            finishReason = choice.finish_reason;
          }

          const delta = choice.delta;
          if (!delta) continue;

          // Accumulate text content
          if (delta.content) {
            fullText += delta.content;
            if (onText) onText(delta.content, fullText);
          }

          // Accumulate reasoning/thinking tokens. OpenRouter emits these
          // under `delta.reasoning`; some upstreams use `reasoning_content`
          // as an alias, so we check both for robustness.
          const reasoningChunk = delta.reasoning || delta.reasoning_content;
          if (reasoningChunk) {
            fullReasoning += reasoningChunk;
            if (onThinking) onThinking(fullReasoning);
          }

          // Accumulate tool calls by index
          if (delta.tool_calls) {
            for (const tc of delta.tool_calls) {
              const idx = tc.index ?? 0;
              if (!toolCallMap[idx]) {
                toolCallMap[idx] = {
                  id: tc.id || '',
                  type: 'function',
                  function: { name: '', arguments: '' }
                };
              }
              if (tc.id) toolCallMap[idx].id = tc.id;
              if (tc.function?.name) toolCallMap[idx].function.name += tc.function.name;
              if (tc.function?.arguments) toolCallMap[idx].function.arguments += tc.function.arguments;
            }
          }
        } catch (e) {
          // Incomplete JSON — will be completed in next iteration
        }
      }
    }

    // Build function calls array from accumulated map
    const functionCalls = [];
    const toolCallsArr = [];
    const indices = Object.keys(toolCallMap).sort((a, b) => parseInt(a) - parseInt(b));
    for (const idx of indices) {
      const tc = toolCallMap[idx];
      let args = {};
      try {
        args = tc.function.arguments ? JSON.parse(tc.function.arguments) : {};
      } catch (e) {
        console.warn('Failed to parse tool call arguments:', tc.function.arguments);
      }
      functionCalls.push({
        id: tc.id,
        name: tc.function.name,
        args: args
      });
      toolCallsArr.push(tc);
    }

    // Build the assistant message for conversation history
    const message = {
      role: 'assistant',
      content: fullText || null,
      // Preserve reasoning so it can be replayed on follow-up turns.
      reasoning: fullReasoning || undefined,
      tool_calls: toolCallsArr.length > 0 ? toolCallsArr : undefined
    };

    // Map OpenAI finish_reason to internal format
    let mappedFinishReason = 'STOP';
    if (finishReason === 'length') mappedFinishReason = 'MAX_TOKENS';
    else if (finishReason === 'tool_calls') mappedFinishReason = 'STOP';

    return {
      ok: true,
      text: fullText,
      reasoning: fullReasoning,
      functionCalls,
      message,
      finishReason: mappedFinishReason
    };
  },

  backoffSeconds(base, attempt) {
    const raw = base * Math.pow(2, attempt);
    const jitter = raw * (0.8 + Math.random() * 0.4);
    return Math.round(jitter);
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

  truncate(s, n) {
    return s.length > n ? s.substring(0, n) + '...' : s;
  }
};
