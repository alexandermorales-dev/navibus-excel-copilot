/* ============================================
   config.js — Settings & state management
   ============================================ */

const Config = {
  apiKey: '',
  model: 'gemini-3.6-flash',
  fallbackModel: 'gemini-3.5-flash-lite',

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
    return this.apiKey && this.apiKey.length > 0;
  }
};
