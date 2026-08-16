/* ============================================
   config.js — Settings & state management
   Includes hardcoded API key pool with round-robin rotation.
   ============================================ */

const Config = {
  apiKey: '',
  model: 'gemini-3.6-flash',
  fallbackModel: 'gemini-3.5-flash-lite',

  // Hardcoded key pool — rotated round-robin on every Gemini request.
  // If the user enters their own key in Settings, it's prepended to the
  // pool so it's tried first, then we cycle through these.
  // Stored base64-encoded to avoid plaintext exposure in the repo.
  _encodedKeys: [
    'QVEuQWI4Uk42SU8yWTdSYU83a3l4VDVxWkM0ZkppWVZqQzhSekZFUDJDUGl0Y21wNG81Wmc=',
    'QVEuQWI4Uk42STVZZmJRbzZDY19fUll1T05hb2FiOHctWjFOOUttQnVLQnotSEQ2cmJzc1E='
  ],

  get hardcodedKeys() {
    return this._encodedKeys.map(k => {
      try { return atob(k); } catch (e) { return ''; }
    });
  },

  // Round-robin index into the key pool.
  _keyIndex: 0,

  // Models that have been deprecated — auto-migrate to current names
  deprecatedModels: {
    'gemini-2.5-flash': 'gemini-3.6-flash',
    'gemini-2.5-flash-lite': 'gemini-3.5-flash-lite',
    'gemini-2.0-flash': 'gemini-3.6-flash',
    'gemini-2.0-flash-lite': 'gemini-3.5-flash-lite',
  },

  load() {
    try {
      this.apiKey = localStorage.getItem('gemini_api_key') || '';
      let savedModel = localStorage.getItem('gemini_model');
      // Auto-migrate deprecated model names
      if (savedModel && this.deprecatedModels[savedModel]) {
        console.log('Migrating deprecated model:', savedModel, '->', this.deprecatedModels[savedModel]);
        savedModel = this.deprecatedModels[savedModel];
        localStorage.setItem('gemini_model', savedModel);
      }
      this.model = savedModel || 'gemini-3.6-flash';
    } catch (e) {
      console.warn('Config load failed:', e);
    }
  },

  save() {
    try {
      localStorage.setItem('gemini_api_key', this.apiKey);
      localStorage.setItem('gemini_model', this.model);
    } catch (e) {
      console.warn('Config save failed:', e);
    }
  },

  hasApiKey() {
    // We always have keys available (hardcoded pool), so this is always true.
    // Kept for compatibility with the UI gating logic.
    return true;
  },

  /**
   * Build the key pool: user's key first (if set), then hardcoded keys.
   * @returns {string[]}
   */
  keyPool() {
    const pool = [];
    if (this.apiKey && this.apiKey.length > 0) {
      pool.push(this.apiKey);
    }
    for (const k of this.hardcodedKeys) {
      if (k && !pool.includes(k)) pool.push(k);
    }
    return pool;
  },

  /**
   * Get the next API key (round-robin). Advances the index so the next
   * call uses a different key, distributing load across the pool.
   * @returns {string}
   */
  nextKey() {
    const pool = this.keyPool();
    if (pool.length === 0) return '';
    const key = pool[this._keyIndex % pool.length];
    this._keyIndex = (this._keyIndex + 1) % pool.length;
    return key;
  },

  /**
   * Reset the rotation index (e.g. at the start of a new user request).
   */
  resetRotation() {
    this._keyIndex = 0;
  }
};
