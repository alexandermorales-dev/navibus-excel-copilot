/* ============================================
   config.js — Persistent storage + settings

   Store: a three-layer persistence helper. Office roamingSettings is
   the only layer that reliably survives an Excel relaunch; localStorage
   is a fallback; an in-memory map covers WebViews that block both.

   Config: per-provider API keys, model overrides, and the cached model
   lists from provider discovery.

   NOTE: this file previously shipped a shared OpenRouter key baked into
   a public repository. It has been removed — every user brings their own
   key, so nobody shares (or can drain) anyone else's free-tier quota.
   ============================================ */

const Store = {
  _memory: {},

  get(key) {
    if (Object.prototype.hasOwnProperty.call(this._memory, key)) {
      return this._memory[key];
    }
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        const val = Office.context.roamingSettings.get(key);
        if (val !== null && val !== undefined) {
          this._memory[key] = val;
          return val;
        }
      }
    } catch (e) { /* roamingSettings unavailable */ }
    try {
      const val = localStorage.getItem(key);
      if (val !== null) {
        this._memory[key] = val;
        return val;
      }
    } catch (e) { /* localStorage blocked */ }
    return null;
  },

  set(key, value) {
    this._memory[key] = value;
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.set(key, value);
        Office.context.roamingSettings.saveAsync();
      }
    } catch (e) { /* roamingSettings unavailable */ }
    try {
      localStorage.setItem(key, value);
    } catch (e) { /* localStorage blocked */ }
  },

  remove(key) {
    delete this._memory[key];
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.roamingSettings) {
        Office.context.roamingSettings.remove(key);
        Office.context.roamingSettings.saveAsync();
      }
    } catch (e) { /* roamingSettings unavailable */ }
    try {
      localStorage.removeItem(key);
    } catch (e) { /* localStorage blocked */ }
  },

  getJSON(key, fallback) {
    const raw = this.get(key);
    if (!raw) return fallback;
    try {
      return JSON.parse(raw);
    } catch (e) {
      return fallback;
    }
  },

  setJSON(key, value) {
    try {
      this.set(key, JSON.stringify(value));
    } catch (e) { /* unserializable — ignore */ }
  }
};

const Config = {
  // Per-provider API keys, keyed by provider id.
  keys: {},
  // Optional per-provider model override chosen by the user in Settings.
  overrides: {},

  load() {
    for (const id of Providers.ids()) {
      this.keys[id] = Store.get(`key_${id}`) || '';
      this.overrides[id] = Store.get(`model_${id}`) || '';
      const models = Store.getJSON(`models_${id}`, null);
      if (Array.isArray(models) && models.length > 0) {
        Providers.discovered[id] = models;
      }
    }
  },

  keyFor(providerId) {
    return this.keys[providerId] || '';
  },

  setKey(providerId, value) {
    const key = (value || '').trim();
    this.keys[providerId] = key;
    if (key) Store.set(`key_${providerId}`, key);
    else Store.remove(`key_${providerId}`);
  },

  modelOverride(providerId) {
    return this.overrides[providerId] || '';
  },

  setModelOverride(providerId, model) {
    this.overrides[providerId] = model || '';
    if (model) Store.set(`model_${providerId}`, model);
    else Store.remove(`model_${providerId}`);
  },

  saveDiscovered(providerId, models) {
    Store.setJSON(`models_${providerId}`, models);
  },

  loadDiscovered(providerId) {
    return Store.getJSON(`models_${providerId}`, null);
  },

  /**
   * True when at least one provider has a key — the app is usable with
   * a single key configured.
   */
  hasApiKey() {
    return Providers.ids().some(id => !!this.keys[id]);
  },

  configuredCount() {
    return Providers.ids().filter(id => !!this.keys[id]).length;
  }
};
