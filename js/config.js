/* ============================================
   config.js — Settings & state management
   API key via OpenRouter (preset + user-provided).
   ============================================ */

const Config = {
  apiKey: '',
  model: 'smart',              // 'smart' | 'groq/llama-3.3-70b-versatile' | 'groq/llama-3.1-8b-instant'
  fastModel: 'groq/llama-3.1-8b-instant',
  capableModel: 'groq/llama-3.3-70b-versatile',

  // Preset key — base64-encoded, loaded from gitignored preset-key.js
  _presetKeyEncoded: (typeof PRESET_KEY !== 'undefined') ? PRESET_KEY : '',

  get presetKey() {
    try { return atob(this._presetKeyEncoded); } catch (e) { return ''; }
  },

  // Returns the active API key: user's key first, preset key as fallback.
  get activeKey() {
    return (this.apiKey && this.apiKey.length > 0) ? this.apiKey : this.presetKey;
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
    try {
      localStorage.setItem('openrouter_api_key', this.apiKey);
      localStorage.setItem('groq_model', this.model);
    } catch (e) {
      console.warn('Config save failed:', e);
    }
  },

  hasApiKey() {
    return !!(this.activeKey && this.activeKey.length > 0);
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
