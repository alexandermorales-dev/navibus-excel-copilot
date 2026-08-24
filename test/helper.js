/* ============================================
   helper.js — test harness for browser-global modules

   The add-in has no build step: every js/ file declares a top-level
   const and is loaded via <script>. These tests load the same files
   into a vm context in the same order, so what is tested is exactly
   what ships — no module wrappers, no bundler.
   ============================================ */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const JS_DIR = path.join(__dirname, '..', 'js');

/**
 * Minimal localStorage stand-in so persistence paths are exercised.
 */
function fakeLocalStorage() {
  const map = new Map();
  return {
    getItem: (k) => (map.has(k) ? map.get(k) : null),
    setItem: (k, v) => { map.set(k, String(v)); },
    removeItem: (k) => { map.delete(k); },
    clear: () => map.clear(),
    _map: map
  };
}

/**
 * Top-level `const Name = ...` declarations, which is how every module in
 * js/ exposes itself. In a vm context these live in the script's lexical
 * scope rather than on globalThis, so the files are concatenated into a
 * single script and the declared names are returned explicitly.
 */
function declaredNames(code) {
  const names = [];
  const re = /^const\s+([A-Za-z_$][\w$]*)\s*=/gm;
  let m;
  while ((m = re.exec(code)) !== null) names.push(m[1]);
  return names;
}

/**
 * Load the given js/ files into a fresh sandbox and return the modules
 * they declare, alongside the sandbox globals.
 *
 * @param {string[]} files   — file names relative to js/, in load order
 * @param {object}  [extras] — extra globals injected before loading
 */
function loadApp(files, extras = {}) {
  const sandbox = {
    console,
    TextEncoder,
    TextDecoder,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    // Host JSON so parsed results carry host prototypes and can be compared
    // with deepStrictEqual across the vm realm boundary.
    JSON,
    localStorage: fakeLocalStorage(),
    fetch: async () => { throw new Error('fetch not stubbed in this test'); },
    ...extras
  };
  sandbox.globalThis = sandbox;
  sandbox.window = sandbox;
  vm.createContext(sandbox);

  const sources = files.map(f => fs.readFileSync(path.join(JS_DIR, f), 'utf8'));
  const names = [...new Set(sources.flatMap(declaredNames))];
  const script = sources.join('\n;\n') + `\n;({ ${names.join(', ')} });`;

  const modules = vm.runInContext(script, sandbox, { filename: files.join('+') });
  return Object.assign(sandbox, modules);
}

/**
 * Sandbox with the full non-UI stack loaded and a usable I18n.
 * Office.js dependent modules (schema/tools/journal) are excluded unless
 * requested, since they require a live Excel host.
 */
function loadCore(extraFiles = [], extras = {}) {
  return loadApp([
    'config.js', 'i18n.js', 'providers.js', 'quota.js', 'llm.js', ...extraFiles
  ], extras);
}

module.exports = { loadApp, loadCore, fakeLocalStorage, JS_DIR };
