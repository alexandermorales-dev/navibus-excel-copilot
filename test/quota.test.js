/* Tests for the quota governor: routing, ceilings, failover, persistence.
   Run with: node --test "test/*.test.js" */

const test = require('node:test');
const assert = require('node:assert');
const { loadCore } = require('./helper.js');

/**
 * Fresh sandbox with the given provider keys configured.
 */
function setup(keys = { gemini: 'k1', groq: 'k2', openrouter: 'k3' }) {
  const app = loadCore();
  const { Config, Quota, Providers } = app;
  Config.load();
  for (const [id, key] of Object.entries(keys)) Config.setKey(id, key);
  // Pin model resolution so tests don't depend on discovery.
  for (const id of Providers.ids()) {
    Providers.discovered[id] = Providers.get(id).fallbackModels.plan.slice();
  }
  Quota.load();
  return app;
}

test('pick: prefers Gemini, the highest-priority provider', () => {
  const { Quota } = setup();
  const route = Quota.pick('plan', 1000);
  assert.strictEqual(route.providerId, 'gemini');
  assert.strictEqual(route.waitMs, 0);
});

test('pick: falls through to Groq when Gemini is exhausted', () => {
  const { Quota } = setup();
  Quota.ledger.gemini.exhausted = true;
  assert.strictEqual(Quota.pick('plan', 1000).providerId, 'groq');
});

test('pick: falls through to OpenRouter when the first two are spent', () => {
  const { Quota } = setup();
  Quota.ledger.gemini.exhausted = true;
  Quota.ledger.groq.exhausted = true;
  assert.strictEqual(Quota.pick('plan', 1000).providerId, 'openrouter');
});

test('pick: returns null when nothing is configured', () => {
  const { Quota } = setup({});
  assert.strictEqual(Quota.pick('plan', 1000), null);
});

test('pick: skips a provider whose per-day request cap is reached', () => {
  const { Quota, Providers } = setup();
  Quota.ledger.gemini.requests = Providers.get('gemini').limits.rpd;
  assert.strictEqual(Quota.pick('plan', 1000).providerId, 'groq');
});

test('pick: skips a provider that cannot fit the request in its TPM ceiling', () => {
  const { Quota } = setup();
  // Groq's 8k TPM cannot hold a 50k-token call; Gemini's 250k can.
  const route = Quota.pick('plan', 50000);
  assert.strictEqual(route.providerId, 'gemini');

  Quota.ledger.gemini.exhausted = true;
  const next = Quota.pick('plan', 50000);
  // Groq is structurally too small, so it must land on OpenRouter.
  assert.strictEqual(next.providerId, 'openrouter');
});

test('availability: toolarge is permanent, not a wait', () => {
  const { Quota } = setup();
  const av = Quota.availability('groq', 50000);
  assert.strictEqual(av.ok, false);
  assert.strictEqual(av.reason, 'toolarge');
  assert.strictEqual(av.waitMs, Infinity);
});

test('availability: RPM ceiling produces a bounded wait, not exhaustion', () => {
  const { Quota, Providers } = setup({ gemini: 'k1' });
  const rpm = Providers.get('gemini').limits.rpm;
  for (let i = 0; i < rpm; i++) Quota.record('gemini', { tokens: 10 });
  const av = Quota.availability('gemini', 100);
  assert.strictEqual(av.ok, false);
  assert.strictEqual(av.reason, 'rpm');
  assert.ok(av.waitMs > 0 && av.waitMs <= 61000, `waitMs=${av.waitMs}`);
});

test('pick: prefers an immediately-free provider over waiting on a faster one', () => {
  const { Quota, Providers } = setup();
  const rpm = Providers.get('gemini').limits.rpm;
  for (let i = 0; i < rpm; i++) Quota.record('gemini', { tokens: 10 });
  const route = Quota.pick('plan', 100);
  assert.strictEqual(route.providerId, 'groq');
  assert.strictEqual(route.waitMs, 0);
});

test('pick: returns the shortest wait when every provider is throttled', () => {
  const { Quota, Providers } = setup();
  for (const id of Providers.ids()) {
    const rpm = Providers.get(id).limits.rpm;
    for (let i = 0; i < rpm; i++) Quota.record(id, { tokens: 10 });
  }
  const route = Quota.pick('plan', 100);
  assert.ok(route, 'expected a throttled route rather than null');
  assert.ok(route.waitMs > 0);
});

test('penalize: 401 marks the key invalid permanently', () => {
  const { Quota } = setup();
  Quota.penalize('gemini', { status: 401 });
  assert.strictEqual(Quota.availability('gemini', 100).reason, 'invalidkey');
  assert.strictEqual(Quota.pick('plan', 100).providerId, 'groq');
});

test('penalize: 402 marks the provider exhausted for the day', () => {
  const { Quota } = setup();
  Quota.penalize('gemini', { status: 402 });
  assert.strictEqual(Quota.availability('gemini', 100).reason, 'exhausted');
});

test('penalize: 429 sets a cooldown honouring Retry-After', () => {
  const { Quota } = setup();
  Quota.penalize('gemini', { status: 429, retryAfterSec: 30 });
  const av = Quota.availability('gemini', 100);
  assert.strictEqual(av.reason, 'cooldown');
  assert.ok(av.waitMs > 25000 && av.waitMs <= 30000);
});

test('penalize: 429 without Retry-After uses at least 20s cooldown (RPM window)', () => {
  const { Quota } = setup();
  Quota.penalize('gemini', { status: 429 });
  const av = Quota.availability('gemini', 100);
  assert.strictEqual(av.reason, 'cooldown');
  assert.ok(av.waitMs >= 15000, `expected >=15s cooldown, got ${av.waitMs}ms`);
});

test('penalize: rapid 429s do NOT cause day exhaustion (RPM protection)', () => {
  const { Quota } = setup();
  // Three 429s in rapid succession should only set cooldowns, not mark
  // the provider as exhausted for the day. This is the RPM scenario:
  // the provider just needs a minute to reset its rolling window.
  for (let i = 0; i < Quota.MAX_429_STREAK; i++) Quota.penalize('gemini', { status: 429 });
  assert.strictEqual(Quota.availability('gemini', 100).reason, 'cooldown');
});

test('penalize: 429 after cooldown expired escalates to day exhaustion', () => {
  const { Quota } = setup();
  // Simulate: 429 → cooldown → wait for cooldown → 429 → cooldown → wait → 429
  for (let i = 0; i < Quota.MAX_429_STREAK; i++) {
    Quota.penalize('gemini', { status: 429 });
    // Simulate the cooldown expiring before the next 429.
    Quota._cooldownUntil.gemini = Date.now() - 1000;
  }
  assert.strictEqual(Quota.availability('gemini', 100).reason, 'exhausted');
});

test('reward: clears the 429 streak so one bad minute is not fatal', () => {
  const { Quota } = setup();
  Quota.penalize('gemini', { status: 429 });
  Quota.penalize('gemini', { status: 429 });
  Quota.reward('gemini');
  Quota.penalize('gemini', { status: 429 });
  assert.notStrictEqual(Quota.availability('gemini', 100).reason, 'exhausted');
});

test('pick: prefers waiting for high-priority provider on short cooldown', () => {
  const { Quota, Providers } = setup();
  // Put Gemini (priority 1) in a short cooldown.
  Quota.penalize('gemini', { status: 429 });
  const av = Quota.availability('gemini', 100);
  assert.strictEqual(av.reason, 'cooldown');
  // pick() should prefer waiting for Gemini over OpenRouter (priority 3).
  const route = Quota.pick('plan', 2000);
  assert.ok(route, 'pick should return a route');
  assert.strictEqual(route.providerId, 'gemini');
  assert.ok(route.waitMs > 0, 'should have a wait time');
  assert.ok(route.waitMs <= Quota.PREFER_WAIT_MS, 'wait should be within prefer threshold');
});

test('pick: fails over to lower-priority provider when cooldown is long', () => {
  const { Quota, Providers } = setup();
  // Put Gemini in a long cooldown (simulated by setting a far-future cooldown).
  Quota._cooldownUntil.gemini = Date.now() + 120000;  // 2 minutes
  // pick() should fail over to the next available provider.
  const route = Quota.pick('plan', 2000);
  assert.ok(route, 'pick should return a route');
  assert.notStrictEqual(route.providerId, 'gemini');
});

/* ---------- alternate model fallback ---------- */

test('altModels: returns alternate models for the same provider', () => {
  const { Providers } = setup();
  const primary = Providers.resolveModel('gemini', 'plan');
  const alts = Providers.altModels('gemini', 'plan', primary);
  assert.ok(alts.length > 0, 'should have at least one alternate model');
  assert.ok(!alts.includes(primary), 'alternates should exclude the primary model');
});

test('altModels: returns empty when user has explicit override', () => {
  const { Providers, Config } = setup();
  Config.setModelOverride('gemini', 'gemini-3.6-flash');
  const primary = Providers.resolveModel('gemini', 'plan');
  const alts = Providers.altModels('gemini', 'plan', primary);
  assert.strictEqual(alts.length, 0);
});

test('altModels: caps at 2 alternates', () => {
  const { Providers } = setup();
  // Populate discovered with many flash variants.
  Providers.discovered.gemini = [
    'gemini-3.6-flash', 'gemini-3.5-flash', 'gemini-3.1-flash-lite',
    'gemini-3.5-flash-lite', 'gemini-flash-latest'
  ];
  const primary = Providers.resolveModel('gemini', 'plan');
  const alts = Providers.altModels('gemini', 'plan', primary);
  assert.ok(alts.length <= 2);
});

test('resolveModel: filters deprecated models from discovered list', () => {
  const { Providers } = setup();
  // Simulate a stale cache that still includes gemini-2.5-flash.
  Providers.discovered.gemini = [
    'gemini-2.5-flash', 'gemini-3.6-flash', 'gemini-3.1-flash-lite'
  ];
  const model = Providers.resolveModel('gemini', 'plan');
  // Should NOT return the deprecated gemini-2.5-flash.
  assert.ok(!/gemini-2\.5/.test(model), `expected non-deprecated model, got ${model}`);
  assert.ok(/gemini-3/.test(model), `expected gemini-3.x, got ${model}`);
});

test('resolveModel: clears deprecated user override', () => {
  const { Providers, Config } = setup();
  Config.setModelOverride('gemini', 'gemini-2.5-flash');
  const model = Providers.resolveModel('gemini', 'plan');
  assert.ok(!/gemini-2\.5/.test(model), `expected non-deprecated model, got ${model}`);
  // Override should have been cleared.
  assert.strictEqual(Config.modelOverride('gemini'), '');
});

test('record: counts requests and tokens', () => {
  const { Quota } = setup();
  Quota.record('gemini', { tokens: 500 });
  Quota.record('gemini', { tokens: 250 });
  assert.strictEqual(Quota.ledger.gemini.requests, 2);
  assert.strictEqual(Quota.ledger.gemini.tokens, 750);
});

test('remaining and pooledRemaining reflect all configured providers', () => {
  const { Quota, Providers } = setup();
  const expected = Providers.ids().reduce((s, id) => s + Providers.get(id).limits.rpd, 0);
  assert.strictEqual(Quota.pooledRemaining(), expected);
  Quota.record('gemini', { tokens: 1 });
  assert.strictEqual(Quota.pooledRemaining(), expected - 1);
});

test('pooledRemaining ignores providers without a key', () => {
  const { Quota, Providers } = setup({ gemini: 'k1' });
  assert.strictEqual(Quota.pooledRemaining(), Providers.get('gemini').limits.rpd);
});

test('ledger persists across a reload within the same day', () => {
  const app = setup();
  app.Quota.record('gemini', { tokens: 400 });
  const saved = app.Store.getJSON(app.Quota.STORAGE_KEY, null);
  assert.strictEqual(saved.ledger.gemini.requests, 1);

  // Simulate a relaunch: same storage contents, fresh Quota state.
  app.Quota.ledger = {};
  app.Quota.load();
  assert.strictEqual(app.Quota.ledger.gemini.requests, 1);
});

test('day rollover resets counters', () => {
  const { Quota, Store } = setup();
  Quota.record('gemini', { tokens: 400 });
  // Backdate the stored ledger, then force a rollover check.
  Store.setJSON(Quota.STORAGE_KEY, { day: '2000-01-01', ledger: Quota.ledger });
  Quota.day = '2000-01-01';
  Quota._rollover();
  assert.strictEqual(Quota.ledger.gemini.requests, 0);
});

test('summary reports state per configured provider', () => {
  const { Quota } = setup({ gemini: 'k1', groq: 'k2' });
  Quota.penalize('groq', { status: 402 });
  const rows = Quota.summary();
  assert.strictEqual(rows.length, 2);
  assert.strictEqual(rows.find(r => r.id === 'gemini').state, 'ok');
  assert.strictEqual(rows.find(r => r.id === 'groq').state, 'exhausted');
});

test('reset clears counters and penalties', () => {
  const { Quota } = setup();
  Quota.record('gemini', { tokens: 10 });
  Quota.penalize('gemini', { status: 402 });
  Quota.reset();
  assert.strictEqual(Quota.ledger.gemini.requests, 0);
  assert.strictEqual(Quota.availability('gemini', 100).ok, true);
});

/* ---------- model resolution ---------- */

test('resolveModel: matches discovered models by preference order', () => {
  const { Providers } = setup();
  Providers.discovered.groq = ['llama-3.1-8b-instant', 'llama-3.3-70b-versatile', 'whisper-large-v3'];
  assert.strictEqual(Providers.resolveModel('groq', 'plan'), 'llama-3.3-70b-versatile');
  assert.strictEqual(Providers.resolveModel('groq', 'answer'), 'llama-3.1-8b-instant');
});

test('resolveModel: prefers non-lite flash for planning on Gemini', () => {
  const { Providers } = setup();
  Providers.discovered.gemini = ['gemini-9.9-flash-lite', 'gemini-9.9-flash'];
  assert.strictEqual(Providers.resolveModel('gemini', 'plan'), 'gemini-9.9-flash');
  assert.strictEqual(Providers.resolveModel('gemini', 'answer'), 'gemini-9.9-flash-lite');
});

test('resolveModel: unseen future model names still resolve via patterns', () => {
  const { Providers } = setup();
  // A model naming scheme that did not exist when this code was written.
  Providers.discovered.gemini = ['gemini-7.0-flash-lite', 'gemini-7.0-flash'];
  assert.strictEqual(Providers.resolveModel('gemini', 'plan'), 'gemini-7.0-flash');
});

test('resolveModel: falls back to the first discovered model when no pattern matches', () => {
  const { Providers } = setup();
  Providers.discovered.groq = ['some-brand-new-model-v9'];
  assert.strictEqual(Providers.resolveModel('groq', 'plan'), 'some-brand-new-model-v9');
});

test('resolveModel: user override wins over discovery', () => {
  const { Providers, Config } = setup();
  Providers.discovered.groq = ['llama-3.3-70b-versatile'];
  Config.setModelOverride('groq', 'my-pinned-model');
  assert.strictEqual(Providers.resolveModel('groq', 'plan'), 'my-pinned-model');
});

test('configured: returns only keyed providers in priority order', () => {
  const { Providers } = setup({ openrouter: 'k3', gemini: 'k1' });
  assert.strictEqual(Providers.configured().map(p => p.id).join(','), 'gemini,openrouter');
});

test('config: hasApiKey is true with a single provider configured', () => {
  const { Config } = setup({ groq: 'only-one' });
  assert.strictEqual(Config.hasApiKey(), true);
  assert.strictEqual(Config.configuredCount(), 1);
});

test('config: no preset or fallback key exists', () => {
  const { Config } = setup({});
  assert.strictEqual(Config.hasApiKey(), false);
  assert.strictEqual(Config.keyFor('openrouter'), '');
  // Guard against reintroducing a shared baked-in key.
  assert.strictEqual(Config.presetKey, undefined);
  assert.strictEqual(Config.activeKey, undefined);
});
