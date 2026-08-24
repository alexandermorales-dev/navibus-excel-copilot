/* Tests for op validation: the stage that catches bad plans locally
   instead of spending an API round-trip to discover them.
   Run with: node --test "test/*.test.js" */

const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./helper.js');

/**
 * Snapshot fixture resembling what Schema.snapshot() produces.
 */
function snapshot() {
  return {
    sheetCount: 2,
    sheets: [
      {
        name: 'Datos', empty: false, rows: 501, cols: 6,
        address: 'Datos!A1:F501', hasHeaders: true,
        headers: ['Fecha', 'Region', 'Producto', 'Cantidad', 'Importe', 'Margen'],
        columnTypes: ['date', 'text', 'text', 'number', 'currency', 'percent'],
        columnStats: [null, null, null,
          { sum: 5000, avg: 10, min: 1, max: 99, count: 500, nonNumericCells: 0 },
          { sum: 250000, avg: 500, min: 12, max: 9800, count: 500, nonNumericCells: 0 },
          { sum: 150, avg: 0.3, min: 0.1, max: 0.6, count: 500, nonNumericCells: 0 }],
        sampleRows: [['2026-01-01', 'Norte', 'A', 5, 1200, 0.31]],
        truncated: true, statsPartial: false, statsRowCount: 500
      },
      { name: 'Notas', empty: true, rows: 0, cols: 0 }
    ]
  };
}

function setup() {
  const app = loadApp([
    'config.js', 'i18n.js', 'providers.js', 'quota.js', 'llm.js',
    'schema.js', 'tools.js', 'recipes.js', 'ops.js', 'context.js'
  ]);
  return app;
}

/* ---------- basic structure ---------- */

test('validate: rejects a non-object plan', () => {
  const { Ops } = setup();
  assert.strictEqual(Ops.validate(null, snapshot()).ok, false);
});

test('validate: an empty ops array is valid (question answered in text)', () => {
  const { Ops } = setup();
  const r = Ops.validate({ intent: 'qa', answer: 'The total is 250000.', ops: [] }, snapshot());
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.ops.length, 0);
});

test('validate: drops an op with no "op" field', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ sheet: 'Datos', range: 'A1' }] }, snapshot());
  assert.strictEqual(r.ops.length, 0);
  assert.match(r.dropped[0].reason, /Missing "op"/);
});

test('validate: drops an unknown op name', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'make_magic', sheet: 'Datos' }] }, snapshot());
  assert.strictEqual(r.ops.length, 0);
  assert.match(r.dropped[0].reason, /Unknown op/);
});

test('validate: drops an op missing required fields', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos' }] }, snapshot());
  assert.match(r.dropped[0].reason, /Missing required field/);
});

/* ---------- sheet resolution ---------- */

test('validate: corrects sheet-name case instead of failing', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'datos', range: 'H1', values: [[1]] }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
  assert.strictEqual(r.ops[0].sheet, 'Datos');
  assert.match(r.warnings[0], /corrected/);
});

test('validate: corrects accents and punctuation in sheet names', () => {
  const { Ops } = setup();
  const snap = snapshot();
  snap.sheets.push({ name: 'Región Sur', empty: false, rows: 10, cols: 2, address: 'A1:B10', hasHeaders: true, headers: ['a', 'b'], columnTypes: ['text', 'number'], columnStats: [], sampleRows: [] });
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'region sur', range: 'D1', values: [[1]] }] }, snap);
  assert.strictEqual(r.ops[0].sheet, 'Región Sur');
});

test('validate: drops a write to a sheet that does not exist at all', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Ventas2030', range: 'A1', values: [[1]] }] }, snapshot());
  assert.strictEqual(r.ops.length, 0);
  assert.match(r.dropped[0].reason, /does not exist/);
});

test('validate: allows writing to a sheet created earlier in the same batch', () => {
  const { Ops } = setup();
  const r = Ops.validate({
    ops: [
      { op: 'add_sheet', name: 'Panel' },
      { op: 'write_range', sheet: 'Panel', range: 'A1', values: [['Hi']] }
    ]
  }, snapshot());
  assert.strictEqual(r.ops.length, 2);
  assert.strictEqual(r.dropped.length, 0);
});

test('validate: renames a colliding new sheet and rewrites later references', () => {
  const { Ops } = setup();
  const r = Ops.validate({
    ops: [
      { op: 'add_sheet', name: 'Notas' },
      { op: 'write_range', sheet: 'Notas', range: 'A1', values: [['x']] }
    ]
  }, snapshot());
  assert.strictEqual(r.ops[0].name, 'Notas 2');
  // The dependent write must follow the sheet to its new name.
  assert.strictEqual(r.ops[1].sheet, 'Notas 2');
  assert.match(r.warnings[0], /already exists/);
});

/* ---------- ranges ---------- */

test('validate: rejects a malformed range', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'not-a-range', values: [[1]] }] }, snapshot());
  assert.match(r.dropped[0].reason, /Invalid range/);
});

test('validate: accepts whole-column band ranges', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'format_range', sheet: 'Datos', range: 'A:A', columnWidth: 100 }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
});

test('validate: accepts a sheet-qualified range', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'create_chart', sheet: 'Datos', type: 'pie', sourceRange: 'Datos!A1:B10' }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
});

test('validate: uppercases lowercase range letters', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'h1:i2', values: [[1, 2], [3, 4]] }] }, snapshot());
  assert.strictEqual(r.ops[0].range, 'H1:I2');
});

/* ---------- values and formulas ---------- */

test('validate: wraps a flat values array into a single row', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'H1', values: ['a', 'b'] }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
  assert.strictEqual(r.ops[0].values.length, 1);
  assert.strictEqual(r.ops[0].values[0].length, 2);
});

test('validate: rejects unbalanced formula parentheses', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'H1', values: [['=SUM(A1:A5']] }] }, snapshot());
  assert.match(r.dropped[0].reason, /unbalanced parentheses/);
});

test('validate: rejects an unmatched closing parenthesis', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'H1', values: [['=SUM(A1:A5))']] }] }, snapshot());
  assert.match(r.dropped[0].reason, /unmatched/);
});

test('validate: rejects an unterminated string in a formula', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'H1', values: [['=IF(A1="x,1,2)']] }] }, snapshot());
  assert.match(r.dropped[0].reason, /unterminated string/);
});

test('validate: allows parentheses inside formula strings', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'write_range', sheet: 'Datos', range: 'H1', values: [['=CONCAT("a(b)",A1)']] }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
});

/* ---------- op-specific behaviour ---------- */

test('validate: strips an invalid color but keeps the format op', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'format_range', sheet: 'Datos', range: 'A1', bold: true, fillColor: 'reddish' }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
  assert.strictEqual(r.ops[0].fillColor, undefined);
  assert.strictEqual(r.ops[0].bold, true);
  assert.match(r.warnings[0], /invalid color/);
});

test('validate: accepts colors without a leading hash', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'format_range', sheet: 'Datos', range: 'A1', fillColor: '1a237e' }] }, snapshot());
  assert.strictEqual(r.ops[0].fillColor, '1a237e');
});

test('validate: chart without any data source is rejected', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'create_chart', sheet: 'Datos', type: 'pie' }] }, snapshot());
  assert.match(r.dropped[0].reason, /requires sourceRange/);
});

test('validate: table name spaces are replaced rather than failing', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'create_table', sheet: 'Datos', range: 'A1:C5', name: 'My Table' }] }, snapshot());
  assert.strictEqual(r.ops[0].name, 'My_Table');
});

test('validate: refuses to delete a sheet the user did not ask about', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'delete_sheet', name: 'Datos' }] }, snapshot());
  assert.strictEqual(r.ops.length, 0);
  assert.match(r.dropped[0].reason, /not explicitly requested/);
});

test('validate: allows an explicitly requested sheet deletion', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'delete_sheet', name: 'Datos', userRequested: true }] }, snapshot());
  assert.strictEqual(r.ops.length, 1);
});

/* ---------- argument mapping ---------- */

test('SPEC: chart map accepts dataRange/anchor aliases', () => {
  const { Ops } = setup();
  const args = Ops.SPEC.create_chart.map({
    sheet: 'P', type: 'pie', dataRange: 'A1:B5', anchor: 'E5', title: 'T'
  });
  assert.strictEqual(args.sourceRange, 'A1:B5');
  assert.strictEqual(args.dest, 'E5');
  assert.strictEqual(args.type, 'Pie');
});

test('SPEC: chart map normalizes loose type names', () => {
  const { Ops } = setup();
  assert.strictEqual(Ops.SPEC.create_chart.map({ sheet: 'P', type: 'column_clustered' }).type, 'ColumnClustered');
  assert.strictEqual(Ops.SPEC.create_chart.map({ sheet: 'P', type: 'bar' }).type, 'BarClustered');
  assert.strictEqual(Ops.SPEC.create_chart.map({ sheet: 'P', type: 'nonsense' }).type, 'ColumnClustered');
});

test('SPEC: pivot map converts field/aggregation to col/agg', () => {
  const { Ops } = setup();
  const args = Ops.SPEC.create_pivot.map({
    name: 'p', sheet: 'P', sourceSheet: 'Datos', sourceRange: 'A1:F9',
    rows: ['Region'], values: [{ field: 'Importe', aggregation: 'sum' }]
  });
  assert.strictEqual(args.source, 'Datos!A1:F9');
  assert.strictEqual(args.values[0].col, 'Importe');
  assert.strictEqual(args.values[0].agg, 'sum');
  assert.strictEqual(args.destSheet, 'P');
});

test('SPEC: pivot map accepts plain string values', () => {
  const { Ops } = setup();
  const args = Ops.SPEC.create_pivot.map({ name: 'p', sheet: 'P', values: ['Importe'] });
  assert.strictEqual(args.values[0].col, 'Importe');
  assert.strictEqual(args.values[0].agg, 'sum');
});

test('SPEC: sort map converts column/order to keyColumn/ascending', () => {
  const { Ops } = setup();
  const args = Ops.SPEC.sort_range.map({ sheet: 'P', range: 'A1:C9', column: 2, order: 'desc' });
  assert.strictEqual(args.keyColumn, 2);
  assert.strictEqual(args.ascending, false);
});

test('SPEC: insert_rows_cols accepts a column letter for "at"', () => {
  const { Ops } = setup();
  assert.strictEqual(Ops.SPEC.insert_rows_cols.map({ sheet: 'D', kind: 'columns', at: 'G' }).at, 7);
  assert.strictEqual(Ops.SPEC.insert_rows_cols.map({ sheet: 'D', kind: 'rows', at: 3 }).at, 3);
});

test('SPEC: format map accepts "align" as an alias', () => {
  const { Ops } = setup();
  assert.strictEqual(Ops.SPEC.format_range.map({ sheet: 'P', range: 'A1', align: 'Center' }).horizontalAlignment, 'Center');
});

/* ---------- recipe validation ---------- */

test('validate: recipe without a source is rejected', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.dashboard', sheet: 'Panel' }] }, snapshot());
  assert.match(r.dropped[0].reason, /requires source.sheet/);
});

test('validate: recipe with an unknown source sheet is rejected', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.dashboard', sheet: 'Panel', source: { sheet: 'Nope' }, groupBy: 'Region', valueColumn: 'Importe' }] }, snapshot());
  assert.match(r.dropped[0].reason, /does not exist/);
});

test('validate: recipe source sheet case is corrected', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.dashboard', sheet: 'Panel', source: { sheet: 'DATOS' }, groupBy: 'Region', valueColumn: 'Importe' }] }, snapshot());
  assert.strictEqual(r.ops[0].source.sheet, 'Datos');
});

test('validate: dashboard needs kpis or a breakdown', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.dashboard', sheet: 'Panel', source: { sheet: 'Datos' } }] }, snapshot());
  assert.match(r.dropped[0].reason, /requires kpis or groupBy/);
});

test('validate: summary_table needs groupBy and valueColumn', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.summary_table', sheet: 'S', source: { sheet: 'Datos' }, groupBy: 'Region' }] }, snapshot());
  assert.match(r.dropped[0].reason, /requires groupBy and valueColumn/);
});

test('validate: recipe target sheet collision is renamed', () => {
  const { Ops } = setup();
  const r = Ops.validate({ ops: [{ op: 'recipe.dashboard', sheet: 'Notas', source: { sheet: 'Datos' }, groupBy: 'Region', valueColumn: 'Importe' }] }, snapshot());
  assert.strictEqual(r.ops[0].sheet, 'Notas 2');
});

/* ---------- problem description for the repair call ---------- */

test('describeProblems: renders each category compactly', () => {
  const { Ops } = setup();
  const text = Ops.describeProblems({
    dropped: [{ op: 'write_range', reason: 'Sheet "X" does not exist' }],
    failed: [{ op: 'create_chart', args: { sheet: 'P', range: 'A1' }, error: 'boom' }],
    problems: [{ kind: 'formula_error', detail: '#REF! in P!B4' }]
  });
  assert.match(text, /REJECTED write_range/);
  assert.match(text, /FAILED create_chart/);
  assert.match(text, /FORMULA_ERROR/);
});
