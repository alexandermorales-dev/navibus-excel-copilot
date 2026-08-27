/* Tests for context.js sheet scope detection and agent.js response validation.
   Run with: node --test test/ */

const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./helper.js');

// Load the non-Office.js modules needed for these tests.
// Schema is stubbed because it requires a live Excel host (Office.js).
const ctx = loadApp([
  'config.js', 'i18n.js', 'providers.js', 'quota.js', 'llm.js',
  'context.js', 'prompt.js', 'intent.js', 'agent.js'
], {
  // Minimal Schema stub: toText returns a compact description so
  // Context.render() can produce output without Office.js.
  Schema: {
    toText(snap) {
      if (!snap || snap.sheetCount === 0) return 'The workbook is empty.';
      return snap.sheets.map(s => `  - "${s.name}": ${s.rows} rows x ${s.cols} cols`).join('\n');
    }
  }
});

/* ---------- Sheet scope detection ---------- */

function fakeSnap(sheets) {
  return {
    sheetCount: sheets.length,
    sheets: sheets.map(s => ({
      name: s,
      empty: false,
      error: false,
      hasHeaders: true,
      headers: ['Col1', 'Col2'],
      columnTypes: ['text', 'number'],
      columnStats: [null, { sum: 100, avg: 50, min: 10, max: 90, count: 2 }],
      sampleRows: [['a', 1], ['b', 2]],
      address: `A1:B3`,
      rows: 3,
      cols: 2
    }))
  };
}

test('detectSheetScope: matches exact sheet name', () => {
  const snap = fakeSnap(['Ventas', 'Compras']);
  assert.strictEqual(ctx.Context.detectSheetScope('analiza la hoja Ventas', snap), 'Ventas');
});

test('detectSheetScope: matches case-insensitively', () => {
  const snap = fakeSnap(['Ventas']);
  assert.strictEqual(ctx.Context.detectSheetScope('analiza la hoja ventas', snap), 'Ventas');
});

test('detectSheetScope: matches accent-insensitively', () => {
  const snap = fakeSnap(['Análisis']);
  assert.strictEqual(ctx.Context.detectSheetScope('mira la hoja analisis', snap), 'Análisis');
});

test('detectSheetScope: returns null when no sheet is mentioned', () => {
  const snap = fakeSnap(['Ventas', 'Compras']);
  assert.strictEqual(ctx.Context.detectSheetScope('crea un dashboard', snap), null);
});

test('detectSheetScope: prefers longer sheet name on partial overlap', () => {
  const snap = fakeSnap(['Sales', 'Sales Data']);
  assert.strictEqual(ctx.Context.detectSheetScope('analiza Sales Data', snap), 'Sales Data');
});

test('detectSheetScope: does not match substring inside another word', () => {
  const snap = fakeSnap(['In']);
  // "In" should not match "Investment" in the text
  assert.strictEqual(ctx.Context.detectSheetScope('Show me the Investment data', snap), null);
});

test('detectSheetScope: ignores empty and error sheets', () => {
  const snap = {
    sheetCount: 2,
    sheets: [
      { name: 'Empty', empty: true },
      { name: 'Errored', error: true },
      { name: 'Real', empty: false, error: false }
    ]
  };
  assert.strictEqual(ctx.Context.detectSheetScope('look at Real', snap), 'Real');
  assert.strictEqual(ctx.Context.detectSheetScope('look at Empty', snap), null);
});

test('detectSheetScope: returns null for null/empty input', () => {
  const snap = fakeSnap(['Ventas']);
  assert.strictEqual(ctx.Context.detectSheetScope(null, snap), null);
  assert.strictEqual(ctx.Context.detectSheetScope('', snap), null);
  assert.strictEqual(ctx.Context.detectSheetScope('test', null), null);
});

/* ---------- buildPlanMessages with scope directive ---------- */

test('buildPlanMessages: injects SCOPE directive when sheet is mentioned', () => {
  const snap = fakeSnap(['Ventas', 'Compras']);
  const built = ctx.Context.buildPlanMessages({
    userText: 'analiza la hoja Ventas',
    snap
  });
  assert.ok(built.messages[0].content.includes('## SCOPE'));
  assert.ok(built.messages[0].content.includes('"Ventas"'));
  assert.strictEqual(built.scopeSheet, 'Ventas');
});

test('buildPlanMessages: no SCOPE directive when no sheet mentioned', () => {
  const snap = fakeSnap(['Ventas', 'Compras']);
  const built = ctx.Context.buildPlanMessages({
    userText: 'crea un dashboard',
    snap
  });
  assert.ok(!built.messages[0].content.includes('## SCOPE'));
  assert.strictEqual(built.scopeSheet, null);
});

/* ---------- Agent response validation ---------- */

test('_validateResponse: empty answer and empty ops fails', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: '', ops: [] }, snap, 'build');
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.errorType, 'empty');
});

test('_validateResponse: schema echoed in answer fails', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: '{"intent":"build","ops":[]}', ops: [] }, snap, 'build');
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.errorType, 'empty');
});

test('validateResponse: qa with no answer fails', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: '', ops: [] }, snap, 'qa');
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.errorType, 'empty');
});

test('_validateResponse: build with ops but no answer passes', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: '', ops: [{ op: 'add_sheet', name: 'X' }] }, snap, 'build');
  assert.strictEqual(r.ok, true);
});

test('_validateResponse: build with answer but no ops passes', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: 'I cannot do that.', ops: [] }, snap, 'build');
  assert.strictEqual(r.ok, true);
});

test('_validateResponse: qa with answer passes', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: 'The total is 100.', ops: [] }, snap, 'qa');
  assert.strictEqual(r.ok, true);
});

test('_validateResponse: non-object plan fails', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse(null, snap, 'build');
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.errorType, 'plan');
});

test('_validateResponse: whitespace-only answer with no ops fails', () => {
  const snap = fakeSnap(['Data']);
  const r = ctx.Agent._validateResponse({ answer: '   \n  ', ops: [] }, snap, 'build');
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.errorType, 'empty');
});
