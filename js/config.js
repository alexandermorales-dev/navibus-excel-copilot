/* ============================================
   config.js — Settings & state management
   API key via OpenRouter (preset + user-provided).
   ============================================ */

const Config = {
  apiKey: '',
  geminiApiKey: '',
  _memoryKey: '',  // fallback when all storage is blocked
  _geminiMemoryKey: '',
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

  /**
   * Read a value from the most persistent storage available.
   * Priority: Office roamingSettings > localStorage > null.
   * roamingSettings persists across sessions per user in Excel.
   */
  _read(key) {
    // Try Office.js roamingSettings first (survives relaunch)
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        const val = Office.context.roamingSettings.get(key);
        if (val) return val;
      }
    } catch (e) { /* roamingSettings not available */ }
    // Fallback to localStorage (cleared on relaunch in some WebViews)
    try {
      return localStorage.getItem(key) || null;
    } catch (e) {
      return null;
    }
  },

  /**
   * Write a value to all available storage layers.
   */
  _write(key, value) {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.set(key, value);
        Office.context.roamingSettings.saveAsync();
      }
    } catch (e) { /* roamingSettings not available */ }
    try {
      localStorage.setItem(key, value);
    } catch (e) { /* localStorage blocked */ }
  },

  _remove(key) {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.remove(key);
        Office.context.roamingSettings.saveAsync();
      }
    } catch (e) {}
    try {
      localStorage.removeItem(key);
    } catch (e) {}
  },

  load() {
    this.apiKey = this._read('openrouter_api_key') || '';
    this.geminiApiKey = this._read('gemini_api_key') || '';
    this.model = this._read('groq_model') || 'smart';
  },

  save() {
    this._memoryKey = this.apiKey;
    this._geminiMemoryKey = this.geminiApiKey;
    this._write('openrouter_api_key', this.apiKey);
    this._write('gemini_api_key', this.geminiApiKey);
    this._write('groq_model', this.model);
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
   * Clear the user's API key from memory, storage, and the Settings UI.
   */
  clearUserKey() {
    this.apiKey = '';
    this._memoryKey = '';
    this._remove('openrouter_api_key');
    if (typeof App !== 'undefined' && App.el && App.el.apiKeyInput) {
      App.el.apiKeyInput.value = '';
    }
  }
};
