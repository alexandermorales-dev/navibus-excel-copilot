/* ============================================
   quota.js — Free-tier quota governor

   Tracks, per provider and per day: requests made, tokens consumed,
   rolling per-minute windows, and 429 cooldowns. Routes each call to
   the best provider that still has headroom, and fails over instead of
   sleeping whenever a second key is available.

   Why this exists: every provider's free tier binds on a different
   axis (OpenRouter on requests/day, Groq on tokens/minute, Gemini on
   requests/day with generous tokens). Pooling three keys behind a
   governor that understands each ceiling turns ~50 usable requests a
   day into ~2,000.

   Daily counters roll over on UTC date change. Providers reset on
   their own schedules (Groq at midnight UTC, Gemini on Pacific time),
   so this is an approximation — if we roll over early, the provider
   answers 429 and we mark it exhausted anyway.
   ============================================ */

const Quota = {
  STORAGE_KEY: 'quota_ledger_v1',

  // Persisted per-day counters: { [providerId]: {requests, tokens, exhausted, keyInvalid} }
  ledger: {},
  day: '',

  // Transient rolling windows: { [providerId]: [{ts, tokens}] }
  _window: {},
  // Transient cooldowns and 429 streaks.
  _cooldownUntil: {},
  _streak: {},

  MAX_429_STREAK: 3,   // consecutive 429s before we treat the day as spent

  load() {
    const saved = Store.getJSON(this.STORAGE_KEY, null);
    this.day = this._utcDay();
    if (saved && saved.day === this.day && saved.ledger) {
      this.ledger = saved.ledger;
    } else {
      this.ledger = {};
    }
    for (const id of Providers.ids()) this._entry(id);
  },

  save() {
    Store.setJSON(this.STORAGE_KEY, { day: this.day, ledger: this.ledger });
  },

  _utcDay() {
    return new Date().toISOString().slice(0, 10);
  },

  /**
   * Reset counters when the UTC date changes mid-session.
   */
  _rollover() {
    const today = this._utcDay();
    if (today !== this.day) {
      this.day = today;
      this.ledger = {};
      this._streak = {};
      this._cooldownUntil = {};
      for (const id of Providers.ids()) this._entry(id);
      this.save();
    }
  },

  _entry(id) {
    if (!this.ledger[id]) {
      this.ledger[id] = { requests: 0, tokens: 0, exhausted: false, keyInvalid: false };
    }
    return this.ledger[id];
  },

  _pruneWindow(id) {
    const cutoff = Date.now() - 60000;
    const w = (this._window[id] || []).filter(e => e.ts > cutoff);
    this._window[id] = w;
    return w;
  },

  usedRpm(id) {
    return this._pruneWindow(id).length;
  },

  usedTpm(id) {
    return this._pruneWindow(id).reduce((sum, e) => sum + e.tokens, 0);
  },

  /**
   * Rough token estimate. Cheap and provider-agnostic; actual usage from
   * the response replaces it when the provider reports it.
   */
  estimateTokens(text) {
    if (!text) return 0;
    const s = typeof text === 'string' ? text : JSON.stringify(text);
    return Math.ceil(s.length / 4);
  },

  /**
   * Can this provider take a request of roughly estTokens right now?
   * Returns { ok, waitMs, reason }. waitMs > 0 means "available after a wait".
   */
  availability(id, estTokens = 2000) {
    this._rollover();
    const p = Providers.get(id);
    if (!p) return { ok: false, waitMs: Infinity, reason: 'unknown' };
    if (!Config.keyFor(id)) return { ok: false, waitMs: Infinity, reason: 'nokey' };

    const e = this._entry(id);
    if (e.keyInvalid) return { ok: false, waitMs: Infinity, reason: 'invalidkey' };
    if (e.exhausted) return { ok: false, waitMs: Infinity, reason: 'exhausted' };

    const lim = p.limits;
    if (lim.rpd && e.requests >= lim.rpd) {
      return { ok: false, waitMs: Infinity, reason: 'rpd' };
    }

    const now = Date.now();
    const cooldown = this._cooldownUntil[id] || 0;
    if (cooldown > now) {
      return { ok: false, waitMs: cooldown - now, reason: 'cooldown' };
    }

    const w = this._pruneWindow(id);
    const oldest = w.length > 0 ? w[0].ts : now;
    const ageOutMs = Math.max(0, 60000 - (now - oldest)) + 250;

    if (lim.rpm && w.length >= lim.rpm) {
      return { ok: false, waitMs: ageOutMs, reason: 'rpm' };
    }
    if (lim.tpm) {
      const used = w.reduce((s, x) => s + x.tokens, 0);
      if (used + estTokens > lim.tpm) {
        // A single call larger than the whole per-minute ceiling can never
        // fit — treat it as unusable rather than waiting forever.
        if (estTokens > lim.tpm) return { ok: false, waitMs: Infinity, reason: 'toolarge' };
        return { ok: false, waitMs: ageOutMs, reason: 'tpm' };
      }
    }

    return { ok: true, waitMs: 0, reason: 'ok' };
  },

  /**
   * Choose the provider+model for a role.
   *
   * Providers are tried in registry priority order, which is set so the
   * scarcest quota (OpenRouter, 50/day) is spent last. If none is free
   * right now, the one with the shortest wait is returned with waitMs so
   * the caller can sleep — but only after every alternative was ruled out.
   *
   * Returns { providerId, provider, model, waitMs } or null.
   */
  pick(role, estTokens = 2000) {
    const candidates = Providers.configured();
    if (candidates.length === 0) return null;

    let best = null;
    for (const p of candidates) {
      const av = this.availability(p.id, estTokens);
      const model = Providers.resolveModel(p.id, role);
      if (!model) continue;
      if (av.ok) return { providerId: p.id, provider: p, model, waitMs: 0 };
      if (av.waitMs !== Infinity && (!best || av.waitMs < best.waitMs)) {
        best = { providerId: p.id, provider: p, model, waitMs: av.waitMs };
      }
    }
    return best;
  },

  /**
   * Record a completed attempt. Called for failures too: on free tiers a
   * failed request still counts against the daily allowance.
   */
  record(id, { tokens = 0, requests = 1 } = {}) {
    this._rollover();
    const e = this._entry(id);
    e.requests += requests;
    e.tokens += tokens;
    if (!this._window[id]) this._window[id] = [];
    this._window[id].push({ ts: Date.now(), tokens });
    this.save();
  },

  /**
   * Register a failure so routing can react to it.
   */
  penalize(id, { status, retryAfterSec } = {}) {
    const e = this._entry(id);

    if (status === 401 || status === 403) {
      e.keyInvalid = true;
      this.save();
      return;
    }
    if (status === 402) {
      e.exhausted = true;
      this.save();
      return;
    }
    if (status === 429) {
      this._streak[id] = (this._streak[id] || 0) + 1;
      // Providers do not distinguish "too fast" from "out for the day" in
      // the status code, so a persistent 429 streak is treated as spent.
      if (this._streak[id] >= this.MAX_429_STREAK) {
        e.exhausted = true;
        this.save();
        return;
      }
      const wait = retryAfterSec ? retryAfterSec * 1000 : Math.min(60000, 5000 * Math.pow(2, this._streak[id] - 1));
      this._cooldownUntil[id] = Date.now() + wait;
      this.save();
      return;
    }
    // Server/network errors: brief cooldown, no day-level penalty.
    this._cooldownUntil[id] = Date.now() + 3000;
  },

  /**
   * Clear the 429 streak after a success.
   */
  reward(id) {
    this._streak[id] = 0;
  },

  remaining(id) {
    const p = Providers.get(id);
    if (!p) return 0;
    const e = this._entry(id);
    if (e.exhausted || e.keyInvalid) return 0;
    return Math.max(0, (p.limits.rpd || 0) - e.requests);
  },

  /**
   * Snapshot for the UI quota bar.
   */
  summary() {
    this._rollover();
    return Providers.ids()
      .filter(id => Config.keyFor(id))
      .map(id => {
        const p = Providers.get(id);
        const e = this._entry(id);
        let state = 'ok';
        if (e.keyInvalid) state = 'invalid';
        else if (e.exhausted || e.requests >= (p.limits.rpd || Infinity)) state = 'exhausted';
        else if ((this._cooldownUntil[id] || 0) > Date.now()) state = 'cooldown';
        else if (this.remaining(id) < (p.limits.rpd || 0) * 0.15) state = 'low';
        return {
          id,
          label: p.label,
          used: e.requests,
          limit: p.limits.rpd || 0,
          remaining: this.remaining(id),
          state
        };
      });
  },

  /**
   * Total pooled requests left today across all configured providers.
   */
  pooledRemaining() {
    return Providers.ids()
      .filter(id => Config.keyFor(id))
      .reduce((sum, id) => sum + this.remaining(id), 0);
  },

  /**
   * Reset all counters — exposed in Settings for when a provider's window
   * rolls over on a different schedule than our UTC assumption.
   */
  reset() {
    this.ledger = {};
    this._window = {};
    this._streak = {};
    this._cooldownUntil = {};
    for (const id of Providers.ids()) this._entry(id);
    this.save();
  }
};
