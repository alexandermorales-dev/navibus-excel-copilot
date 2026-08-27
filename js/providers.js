/* ============================================
   providers.js — Provider registry & model resolution

   All three supported providers expose an OpenAI-compatible
   /chat/completions endpoint, so a single client (llm.js) serves
   them all — only the base URL, auth header and model id differ.

   Model ids are NOT hardcoded. Free-tier model slugs get deprecated
   and renamed constantly, which repeatedly broke this add-in. Instead
   each provider declares preference-ordered patterns per role, and we
   resolve them against the list actually returned by GET /models.
   A static fallback list is used when discovery is unavailable.
   ============================================ */

const Providers = {
  /**
   * Free-tier limits below are the published values used by the quota
   * governor to route requests. They are deliberately in one place so
   * they can be corrected without touching any other file.
   */
  registry: {
    gemini: {
      id: 'gemini',
      label: 'Google Gemini',
      keyUrl: 'https://aistudio.google.com/apikey',
      keyPlaceholder: 'AIza...',
      chatUrl: 'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
      modelsUrl: 'https://generativelanguage.googleapis.com/v1beta/openai/models',
      // Free tier limits as of Jul 2026 (per-model, not per-key):
      //   gemini-3.6-flash:       15 RPM, 1500 RPD (best reasoning, free)
      //   gemini-3.5-flash:       15 RPM, 1500 RPD (free, similar to 3.6)
      //   gemini-3.5-flash-lite:  15 RPM, 1000 RPD (free, fast, high volume)
      //   gemini-3.1-flash-lite:  15 RPM, 1000 RPD (free, workhorse)
      //   gemini-*-pro-*:         NOT free tier — never select
      // We use the flash numbers as the pooled estimate since the
      // alternate-model fallback can tap its separate quota.
      limits: { rpd: 1500, rpm: 15, tpm: 250000 },
      jsonMode: true,
      // Highest priority: most headroom per request.
      priority: 1,
      // Gemini's OpenAI-compatible endpoint needs this to return thought
      // summaries in the stream. Without it, reasoning deltas are never sent.
      requestExtras: {
        extra_body: {
          google: {
            thinking_config: { include_thoughts: true }
          }
        }
      },
      // Only free-tier models. Pro and preview models are excluded —
      // they either have no free tier or lower limits that cause 402/403.
      // Filter: drop Pro models (no free tier), TTS/image/embed variants.
      modelFilter: (id) => !/\bpro\b/i.test(id) && !/tts|image|embed|aqa|guard|veo/i.test(id),
      models: {
        // Plan needs reasoning — prefer full flash over lite.
        // Order: specific versions → "latest" non-lite → any flash (not lite) → lite.
        plan:   [/^models\/gemini-3\.6-flash$/i, /^models\/gemini-3\.5-flash$/i,
                 /^gemini-3\.6-flash$/i, /^gemini-3\.5-flash$/i,
                 /^gemini-flash-latest$/i,
                 /flash(?!-lite)/i, /flash-lite/i],
        // Answer/repair need speed, not deep reasoning — prefer lite.
        answer: [/flash-lite/i, /flash(?!-lite)/i],
        repair: [/flash-lite/i, /flash(?!-lite)/i]
      },
      fallbackModels: {
        plan:   ['gemini-3.6-flash', 'gemini-flash-latest', 'gemini-3.5-flash', 'gemini-3.5-flash-lite'],
        answer: ['gemini-3.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-3.6-flash'],
        repair: ['gemini-3.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-3.6-flash']
      }
    },

    groq: {
      id: 'groq',
      label: 'Groq',
      keyUrl: 'https://console.groq.com/keys',
      keyPlaceholder: 'gsk_...',
      chatUrl: 'https://api.groq.com/openai/v1/chat/completions',
      modelsUrl: 'https://api.groq.com/openai/v1/models',
      // 1,000 req/day is plenty, but the token ceiling is tight — this is
      // why the compact prompt matters. Cached prompt tokens do not count
      // toward Groq rate limits, so we keep the system prefix stable.
      limits: { rpd: 1000, rpm: 30, tpm: 8000 },
      jsonMode: true,
      priority: 2,
      models: {
        plan:   [/llama-3\.3-70b/i, /gpt-oss-120b/i, /(70|120)b/i, /qwen3-32b/i, /maverick|scout/i],
        answer: [/llama-3\.1-8b-instant/i, /8b/i, /gpt-oss-20b/i, /scout/i],
        repair: [/llama-3\.3-70b/i, /gpt-oss-120b/i, /(70|120)b/i]
      },
      fallbackModels: {
        plan:   ['llama-3.3-70b-versatile', 'openai/gpt-oss-120b'],
        answer: ['llama-3.1-8b-instant'],
        repair: ['llama-3.3-70b-versatile']
      }
    },

    openrouter: {
      id: 'openrouter',
      label: 'OpenRouter',
      keyUrl: 'https://openrouter.ai/keys',
      keyPlaceholder: 'sk-or-v1-...',
      chatUrl: 'https://openrouter.ai/api/v1/chat/completions',
      modelsUrl: 'https://openrouter.ai/api/v1/models',
      // Only 50 requests/day on the free tier, and failed requests still
      // count. Kept strictly as an emergency reserve.
      limits: { rpd: 50, rpm: 20, tpm: 200000 },
      jsonMode: true,
      priority: 3,
      // Restrict to :free variants so a stray key with credits can't be billed.
      modelFilter: (id) => /:free$/.test(id),
      models: {
        plan:   [/llama-3\.3-70b/i, /gpt-oss-120b/i, /nemotron.*(super|120b)/i, /(70|120|235)b/i],
        answer: [/(8|9|12)b/i, /flash|lite|mini|small|nano/i],
        repair: [/llama-3\.3-70b/i, /gpt-oss-120b/i, /(70|120)b/i]
      },
      fallbackModels: {
        plan:   ['meta-llama/llama-3.3-70b-instruct:free', 'openai/gpt-oss-120b:free'],
        answer: ['meta-llama/llama-3.3-70b-instruct:free'],
        repair: ['meta-llama/llama-3.3-70b-instruct:free']
      },
      // OpenRouter attributes traffic via these headers.
      extraHeaders: {
        'HTTP-Referer': 'https://alexandermorales-dev.github.io/navibus-excel-copilot/',
        'X-Title': 'Excel AI Copilot'
      }
    }
  },

  ROLES: ['plan', 'answer', 'repair'],

  // Discovered model ids per provider: { gemini: ['gemini-3.6-flash', ...] }
  discovered: {},

  ids() {
    return Object.keys(this.registry);
  },

  get(id) {
    return this.registry[id] || null;
  },

  /**
   * Providers that have a key configured, ordered by priority.
   */
  configured() {
    return this.ids()
      .filter(id => Config.keyFor(id))
      .map(id => this.registry[id])
      .sort((a, b) => a.priority - b.priority);
  },

  /**
   * Fetch the model list for a provider and cache it. Also serves as the
   * "Test connection" probe in Settings: a 200 here means the key works.
   *
   * Returns { ok, models, error }
   */
  async discover(id) {
    const p = this.get(id);
    if (!p) return { ok: false, error: `Unknown provider: ${id}` };
    const key = Config.keyFor(id);
    if (!key) return { ok: false, error: 'No key configured' };

    try {
      const resp = await fetch(p.modelsUrl, {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${key}`, ...(p.extraHeaders || {}) }
      });
      if (!resp.ok) {
        const body = await resp.text().catch(() => '');
        return { ok: false, error: this._httpError(resp.status, body) };
      }
      const json = await resp.json();
      // OpenAI shape: { data: [{ id }] }. Gemini's compat layer prefixes
      // ids with "models/", which the chat endpoint also accepts.
      const raw = Array.isArray(json.data) ? json.data : [];
      let models = raw
        .map(m => String(m.id || m.name || ''))
        .filter(Boolean)
        .map(m => m.replace(/^models\//, ''));
      if (p.modelFilter) models = models.filter(p.modelFilter);
      // Drop non-text endpoints that would fail a chat call.
      models = models.filter(m => !/embed|whisper|tts|imagen|veo|image|vision-only|guard|aqa/i.test(m));

      if (models.length === 0) {
        return { ok: false, error: 'No usable chat models returned' };
      }
      this.discovered[id] = models;
      Config.saveDiscovered(id, models);
      return { ok: true, models };
    } catch (e) {
      return { ok: false, error: e.message || String(e) };
    }
  },

  /**
   * Models known to be deprecated/removed. These are filtered out of
   * discovered lists so stale localStorage caches don't resurrect them.
   * Updated when Google sunsets a model (404 on first call).
   */
  DEPRECATED: [
    /^gemini-2\.[05]-/i,       // gemini-2.5-flash, gemini-2.0-flash (deprecated Jul 2026)
    /^gemini-1\.[05]-/i,       // gemini-1.5-* (older)
  ],

  isDeprecated(model) {
    return this.DEPRECATED.some(re => re.test(model));
  },

  /**
   * Resolve the model id to use for a provider/role.
   * Prefers discovered models matched against the provider's preference
   * patterns; falls back to the static list when discovery hasn't run.
   * Deprecated models are filtered out of discovered lists.
   */
  resolveModel(id, role) {
    const p = this.get(id);
    if (!p) return null;

    // Explicit user override wins — but clear it if the model is deprecated.
    const override = Config.modelOverride(id);
    if (override) {
      if (this.isDeprecated(override)) {
        console.warn(`Providers: overriding deprecated model ${override}, clearing override`);
        Config.setModelOverride(id, '');
      } else {
        return override;
      }
    }

    const raw = this.discovered[id] || Config.loadDiscovered(id) || [];
    // Filter out deprecated models from stale cached lists.
    const available = raw.filter(m => !this.isDeprecated(m));
    if (available.length !== raw.length) {
      // Update the cache so we don't re-filter every call.
      this.discovered[id] = available;
      Config.saveDiscovered(id, available);
    }

    const patterns = p.models[role] || p.models.plan || [];

    for (const re of patterns) {
      const hit = available.find(m => re.test(m));
      if (hit) return hit;
    }

    const fallback = (p.fallbackModels[role] || p.fallbackModels.plan || [])[0];
    // If we have a discovered list but nothing matched, take the first
    // entry rather than a stale hardcoded id that may be deprecated.
    if (available.length > 0) return available[0];
    return fallback || null;
  },

  /**
   * Alternate models for the same provider/role, excluding the primary
   * model. Used for in-provider model fallback before switching to a
   * different provider — e.g. try gemini-3.6-flash, then gemini-3.1-flash-lite,
   * before jumping to Groq.
   *
   * Returns an array of model ids (may be empty).
   */
  altModels(id, role, primaryModel) {
    const p = this.get(id);
    if (!p) return [];

    // If the user set an explicit override, don't try alternates.
    if (Config.modelOverride(id)) return [];

    const available = this.discovered[id] || Config.loadDiscovered(id) || [];
    const patterns = p.models[role] || p.models.plan || [];
    const fallbacks = p.fallbackModels[role] || p.fallbackModels.plan || [];
    const result = [];
    const seen = new Set([primaryModel?.toLowerCase()]);

    // Other pattern matches from discovered list.
    for (const re of patterns) {
      for (const m of available) {
        if (re.test(m) && !seen.has(m.toLowerCase())) {
          seen.add(m.toLowerCase());
          result.push(m);
        }
      }
    }

    // Static fallbacks.
    for (const m of fallbacks) {
      if (!seen.has(m.toLowerCase())) {
        seen.add(m.toLowerCase());
        result.push(m);
      }
    }

    return result.slice(0, 2);   // cap at 2 alternates
  },

  _httpError(status, body) {
    const snippet = String(body || '').slice(0, 200);
    if (status === 401 || status === 403) return `Invalid or unauthorized key (${status})`;
    if (status === 429) return 'Rate limited (429) — key is valid but throttled';
    if (status === 402) return 'Quota or credits exhausted (402)';
    return `HTTP ${status}${snippet ? ': ' + snippet : ''}`;
  }
};
