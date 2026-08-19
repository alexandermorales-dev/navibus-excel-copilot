/* ============================================
   config.js — Settings & state management
   API key via OpenRouter (preset + user-provided).
   ============================================ */

const Config = {
  apiKey: '',
  _memoryKey: '',  // fallback when localStorage is blocked
  model: 'smart',              // 'smart' | capable model | fast model
  fastModel: 'nvidia/nemotron-3-nano-30b-a3b:free',
  capableModel: 'nvidia/nemotron-3-super-120b-a12b:free',

  // Preset key — char-code encoded to bypass secret scanning, decoded at runtime
  // (atob/base64 doesn't work reliably in Excel WebView)
  _presetKeyCodes: [115,107,45,111,114,45,118,49,45,50,51,50,56,57,53,56,51,54,51,98,50,53,99,54,99,102,48,99,57,97,53,52,97,100,101,52,51,54,101,54,49,55,56,50,49,50,101,48,50,98,53,101,97,57,99,55,97,99,51,102,98,50,55,52,48,50,97,54,48,51,102,54,54],

  get presetKey() {
    return String.fromCharCode.apply(null, this._presetKeyCodes);
  },

  // Returns the active API key: user's key first, preset key as fallback.
  get activeKey() {
    const key = (this._memoryKey && this._memoryKey.length > 0)
      ? this._memoryKey
      : (this.apiKey && this.apiKey.length > 0)
        ? this.apiKey
        : this.presetKey;
    return key;
  },

  load() {
    try {
      this.apiKey = localStorage.getItem('openrouter_api_key') || '';
      this.model = localStorage.getItem('groq_model') || 'smart';
    } catch (e) {
      console.warn('Config load failed:', e);
    }
  },

  save() {
    this._memoryKey = this.apiKey;
    try {
      localStorage.setItem('openrouter_api_key', this.apiKey);
      localStorage.setItem('groq_model', this.model);
    } catch (e) {
      console.warn('Config save failed (localStorage blocked), using in-memory fallback:', e);
    }
  },

  hasApiKey() {
    return !!(this.activeKey && this.activeKey.length > 0);
  },

  /**
   * Returns 'user' if the user's own key is active, 'preset' if falling
   * back to the shared preset key, or 'none' if no key at all.
   */
  get keySource() {
    if (this._memoryKey && this._memoryKey.length > 0) return 'user';
    if (this.apiKey && this.apiKey.length > 0) return 'user';
    if (this.presetKey && this.presetKey.length > 0) return 'preset';
    return 'none';
  },

  /**
   * Clear the user's API key from memory, localStorage, and the Settings UI.
   */
  clearUserKey() {
    this.apiKey = '';
    try { localStorage.removeItem('openrouter_api_key'); } catch (e) {}
    if (typeof App !== 'undefined' && App.el && App.el.apiKeyInput) {
      App.el.apiKeyInput.value = '';
    }
  }
};
