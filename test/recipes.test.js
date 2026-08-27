/* Tests for the deterministic layout engine.

   This is the arithmetic that used to be delegated to the model via prose
   ("leave 2 spacer rows", "charts need width 300-450"), which is why
   dashboards came out with overlapping charts and unformatted cells.
   Run with: node --test "test/*.test.js" */

const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./helper.js');

function snapshot() {
  return {
    sheetCount: 1,
    sheets: [{
      name: 'Datos', empty: false, rows: 501, cols: 6,
      address: 'Datos!A1:F501', hasHeaders: true,
      headers: ['Fecha', 'Region', 'Producto', 'Cantidad', 'Importe', 'Margen'],
      columnTypes: ['date', 'text', 'text', 'number', 'currency', 'percent'],
      columnStats: [], sampleRows: []
    }]
  };
}

function setup() {
  return loadApp([
    'config.js', 'i18n.js', 'providers.js', 'quota.js', 'llm.js',
    'schema.js', 'tools.js', 'recipes.js', 'ops.js', 'context.js'
  ]);
}

/**
 * Context stub: readRange returns the group-by column values so the
 * recipe can compute its distinct categories without Excel. When the
 * aggregate method reads the value column, the stub returns numeric
 * values matching the category count so the numeric-data check passes.
 */
function ctxWith(categories) {
  const values = categories.map((_, i) => [100 + i * 10]);
  let callCount = 0;
  return {
    snap: snapshot(),
    readRange: async () => {
      callCount++;
      // First call: group-by column (categories). Second call: value column (numbers).
      if (callCount === 1) {
        return { ok: true, data: categories.map(c => [c]), truncated: false };
      }
      return { ok: true, data: values, truncated: false };
    }
  };
}

const REGIONS = ['Norte', 'Sur', 'Este', 'Norte', 'Sur', 'Oeste'];

/* ---------- parsing / helpers ---------- */

test('chartType: normalizes the names models actually emit', () => {
  const { Recipes } = setup();
  assert.strictEqual(Recipes.chartType('columnClustered'), 'ColumnClustered');
  assert.strictEqual(Recipes.chartType('column_clustered'), 'ColumnClustered');
  assert.strictEqual(Recipes.chartType('Column Clustered'), 'ColumnClustered');
  assert.strictEqual(Recipes.chartType('pie'), 'Pie');
  assert.strictEqual(Recipes.chartType('LINE'), 'Line');
  assert.strictEqual(Recipes.chartType(undefined), 'ColumnClustered');
});

test('isCircular: pie and doughnut have no axes', () => {
  const { Recipes } = setup();
  assert.strictEqual(Recipes.isCircular('pie'), true);
  assert.strictEqual(Recipes.isCircular('doughnut'), true);
  assert.strictEqual(Recipes.isCircular('columnClustered'), false);
});

test('formatFor: derives the number format from the column type', () => {
  const { Recipes } = setup();
  const desc = snapshot().sheets[0];
  assert.strictEqual(Recipes.formatFor(desc, 4, 'SUM'), '$#,##0.00');   // currency
  assert.strictEqual(Recipes.formatFor(desc, 5, 'SUM'), '0.0%');        // percent
  assert.strictEqual(Recipes.formatFor(desc, 3, 'SUM'), '#,##0');       // number
  assert.strictEqual(Recipes.formatFor(desc, 4, 'COUNT'), '#,##0');     // counts are integers
});

test('resolveSource: derives data row bounds from the real used range', () => {
  const { Recipes } = setup();
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, { snap: snapshot() });
  assert.strictEqual(src.ok, true);
  assert.strictEqual(src.headerRow, 1);
  assert.strictEqual(src.firstData, 2);
  assert.strictEqual(src.lastData, 501);
});

test('resolveSource: rejects a sheet with no header row', () => {
  const { Recipes } = setup();
  const snap = snapshot();
  snap.sheets[0].hasHeaders = false;
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, { snap });
  assert.strictEqual(src.ok, false);
  assert.match(src.error, /no header row/);
});

test('colRange: builds an absolute reference for a named column', () => {
  const { Recipes } = setup();
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, { snap: snapshot() });
  const col = Recipes.colRange(src, 'Importe');
  assert.strictEqual(col.letter, 'E');
  assert.strictEqual(col.ref, "'Datos'!$E$2:$E$501");
  assert.strictEqual(Recipes.colRange(src, 'Nope'), null);
});

/* ---------- aggregation ---------- */

test('aggregate: computes distinct categories from a local read', async () => {
  const { Recipes } = setup();
  const ctx = ctxWith(REGIONS);
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, ctx);
  const agg = await Recipes.aggregate({ groupBy: 'Region', valueColumn: 'Importe', agg: 'sum' }, ctx, src);
  assert.strictEqual(agg.ok, true);
  assert.strictEqual(agg.categories.length, 4);
  assert.strictEqual(agg.categories.join(','), 'Este,Norte,Oeste,Sur'); // sorted
});

test('aggregate: caps high-cardinality columns', async () => {
  const { Recipes } = setup();
  const many = Array.from({ length: 400 }, (_, i) => `id-${i}`);
  const ctx = ctxWith(many);
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, ctx);
  const agg = await Recipes.aggregate({ groupBy: 'Region', valueColumn: 'Importe' }, ctx, src);
  assert.strictEqual(agg.categories.length, 25);
  assert.strictEqual(agg.truncated, true);
});

test('aggregate: ignores blank cells', async () => {
  const { Recipes } = setup();
  const ctx = ctxWith(['A', '', null, 'B', undefined]);
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, ctx);
  const agg = await Recipes.aggregate({ groupBy: 'Region', valueColumn: 'Importe' }, ctx, src);
  assert.strictEqual(agg.categories.join(','), 'A,B');
});

test('aggregate: fails clearly on an unknown column', async () => {
  const { Recipes } = setup();
  const ctx = ctxWith(REGIONS);
  const src = Recipes.resolveSource({ source: { sheet: 'Datos' } }, ctx);
  const agg = await Recipes.aggregate({ groupBy: 'Nope', valueColumn: 'Importe' }, ctx, src);
  assert.strictEqual(agg.ok, false);
  assert.match(agg.error, /not found/);
});

/* ---------- summary block ---------- */

async function summaryOps(overrides = {}) {
  const app = setup();
  const { Recipes } = app;
  const ctx = ctxWith(REGIONS);
  const op = { source: { sheet: 'Datos' }, groupBy: 'Region', valueColumn: 'Importe', agg: 'sum', ...overrides };
  const src = Recipes.resolveSource(op, ctx);
  const breakdown = await Recipes.aggregate(op, ctx, src);
  const built = Recipes.buildSummaryBlock({
    sheet: 'Panel', startRow: 5, breakdown, src, op, tableName: 'tblX'
  });
  return { app, built, op };
}

test('summary block: total row sits directly below the category rows', async () => {
  const { built } = await summaryOps();
  assert.strictEqual(built.headerRow, 5);
  assert.strictEqual(built.firstRow, 6);
  assert.strictEqual(built.lastRow, 9);   // 4 categories
  assert.strictEqual(built.totalRow, 10);
});

test('summary block: uses live SUMIF formulas, not pasted values', async () => {
  const { built } = await summaryOps();
  const write = built.ops.find(o => o.op === 'write_range' && o.range === 'A6');
  assert.ok(write, 'expected a category write at A6');
  const [label, value, pct] = write.values[0];
  assert.strictEqual(label, 'Este');
  assert.strictEqual(value, "=SUMIF('Datos'!$B$2:$B$501,A6,'Datos'!$E$2:$E$501)");
  assert.strictEqual(pct, '=IFERROR(B6/$B$10,0)');
});

test('summary block: each category row references its own label cell', async () => {
  const { built } = await summaryOps();
  const write = built.ops.find(o => o.op === 'write_range' && o.range === 'A6');
  write.values.forEach((row, i) => {
    const rowNum = 6 + i;
    assert.ok(row[1].includes(`,A${rowNum},`), `row ${rowNum} should key off A${rowNum}, got ${row[1]}`);
    assert.ok(row[2].includes(`B${rowNum}/`), `row ${rowNum} pct should divide B${rowNum}`);
  });
});

test('summary block: count aggregation uses COUNTIF with no value range', async () => {
  const { built } = await summaryOps({ agg: 'count' });
  const write = built.ops.find(o => o.op === 'write_range' && o.range === 'A6');
  assert.strictEqual(write.values[0][1], "=COUNTIF('Datos'!$B$2:$B$501,A6)");
});

test('summary block: average total aggregates the source, not the rows above', async () => {
  const { built } = await summaryOps({ agg: 'average' });
  const totalWrite = built.ops.find(o => o.op === 'write_range' && o.range === 'A10');
  // Averaging the per-category averages would be arithmetically wrong.
  assert.match(totalWrite.values[0][1], /^=IFERROR\(AVERAGE\(/);
});

test('summary block: sum total adds the category rows', async () => {
  const { built } = await summaryOps({ agg: 'sum' });
  const totalWrite = built.ops.find(o => o.op === 'write_range' && o.range === 'A10');
  assert.strictEqual(totalWrite.values[0][1], '=SUM(B6:B9)');
});

test('summary block: the Excel table excludes the total row', async () => {
  const { built } = await summaryOps();
  const table = built.ops.find(o => o.op === 'create_table');
  assert.strictEqual(table.range, 'A5:C9');   // header..lastRow, not totalRow
});

test('summary block: chart source excludes the total row and includes the header', async () => {
  const { built } = await summaryOps();
  assert.strictEqual(built.chartDataRange, 'A5:B9');
});

test('summary block: number format follows the value column type', async () => {
  const { built } = await summaryOps();
  assert.strictEqual(built.numberFormat, '$#,##0.00'); // Importe is currency
});

/* ---------- full dashboard layout ---------- */

async function dashboard(overrides = {}) {
  const app = setup();
  const ctx = ctxWith(REGIONS);
  const op = {
    op: 'recipe.dashboard',
    sheet: 'Panel',
    title: 'PANEL',
    source: { sheet: 'Datos' },
    kpis: [
      { label: 'Total', column: 'Importe', agg: 'sum' },
      { label: 'Promedio', column: 'Importe', agg: 'average' },
      { label: 'Operaciones', column: 'Importe', agg: 'count' }
    ],
    groupBy: 'Region',
    valueColumn: 'Importe',
    agg: 'sum',
    charts: [{ type: 'columnClustered', title: 'A' }, { type: 'pie', title: 'B' }],
    ...overrides
  };
  const res = await app.Recipes.dashboard(op, ctx);
  return { app, res };
}

test('dashboard: creates its own sheet first', async () => {
  const { res } = await dashboard();
  assert.strictEqual(res.ok, true);
  assert.strictEqual(res.ops[0].op, 'add_sheet');
  assert.strictEqual(res.ops[0].name, 'Panel');
  assert.strictEqual(res.createdSheet, 'Panel');
});

test('dashboard: every op targets the dashboard sheet', async () => {
  const { res } = await dashboard();
  for (const op of res.ops) {
    if (op.op === 'add_sheet') continue;
    assert.strictEqual(op.sheet, 'Panel', `op ${op.op} targeted ${op.sheet}`);
  }
});

test('dashboard: expands into a full layout without the model emitting format ops', async () => {
  const { res } = await dashboard();
  const kinds = res.ops.reduce((m, o) => { m[o.op] = (m[o.op] || 0) + 1; return m; }, {});
  assert.ok(kinds.write_range >= 6, `expected several writes, got ${kinds.write_range}`);
  assert.ok(kinds.format_range >= 10, `expected many formats, got ${kinds.format_range}`);
  assert.strictEqual(kinds.create_chart, 2);
  assert.strictEqual(kinds.create_table, 1);
  assert.strictEqual(kinds.conditional_format, 1);
});

test('dashboard: KPI value cells are formulas over the source column', async () => {
  const { res } = await dashboard();
  const formulas = res.ops
    .filter(o => o.op === 'write_range')
    .flatMap(o => o.values.flat())
    .filter(v => typeof v === 'string' && v.startsWith('='));
  assert.ok(formulas.some(f => f === "=SUM('Datos'!$E$2:$E$501)"));
  assert.ok(formulas.some(f => f === "=AVERAGE('Datos'!$E$2:$E$501)"));
  assert.ok(formulas.some(f => f === "=COUNT('Datos'!$E$2:$E$501)"));
});

test('dashboard: KPI cards do not overlap horizontally', async () => {
  const { app, res } = await dashboard();
  const { Recipes } = app;
  // Each card spans KPI_SPAN columns with KPI_GAP between; check the
  // merged label ranges are disjoint.
  const spans = res.ops
    .filter(o => o.op === 'format_range' && o.merge && /^[A-Z]+3:[A-Z]+3$/.test(o.range))
    .map(o => {
      const [a, b] = o.range.split(':');
      return [app.Tools.letterToNum(a.replace(/\d/g, '')), app.Tools.letterToNum(b.replace(/\d/g, ''))];
    })
    .sort((x, y) => x[0] - y[0]);
  assert.strictEqual(spans.length, 3);
  for (let i = 1; i < spans.length; i++) {
    assert.ok(spans[i][0] > spans[i - 1][1], `card ${i} starts at ${spans[i][0]} but previous ended at ${spans[i - 1][1]}`);
  }
  assert.strictEqual(Recipes.KPI_GAP >= 1, true);
});

test('dashboard: skips a KPI whose column does not exist, keeping the rest', async () => {
  const { res } = await dashboard({
    kpis: [
      { label: 'Good', column: 'Importe', agg: 'sum' },
      { label: 'Bad', column: 'NoSuchColumn', agg: 'sum' }
    ]
  });
  assert.strictEqual(res.ok, true);
  assert.ok(res.warnings.some(w => /NoSuchColumn/.test(w)));
  const formulas = res.ops.filter(o => o.op === 'write_range').flatMap(o => o.values.flat());
  assert.ok(formulas.some(v => v === "=SUM('Datos'!$E$2:$E$501)"));
  assert.ok(!formulas.some(v => typeof v === 'string' && v.includes('NoSuchColumn')));
});

test('dashboard: caps KPIs at four', async () => {
  const { res } = await dashboard({
    kpis: Array.from({ length: 8 }, (_, i) => ({ label: `K${i}`, column: 'Importe', agg: 'sum' }))
  });
  // Match only KPI aggregates over the SOURCE column; the summary block's
  // own total row is also a =SUM() but over the dashboard's own cells.
  const kpiCells = res.ops.filter(o => o.op === 'write_range' &&
    o.values.flat().some(v => typeof v === 'string' && v === "=SUM('Datos'!$E$2:$E$501)"));
  assert.strictEqual(kpiCells.length, 4, `expected exactly 4 KPI value writes, got ${kpiCells.length}`);
});

test('dashboard: stacked charts are vertically separated', async () => {
  const { app, res } = await dashboard();
  const charts = res.ops.filter(o => o.op === 'create_chart');
  const rows = charts.map(c => parseInt(String(c.dest).replace(/[A-Z]/g, ''), 10));
  assert.strictEqual(rows.length, 2);
  const gap = Math.abs(rows[1] - rows[0]);
  assert.ok(gap >= app.Recipes.chartRowSpan(), `charts ${gap} rows apart, need >= ${app.Recipes.chartRowSpan()}`);
});

test('dashboard: charts are anchored clear of the summary table columns', async () => {
  const { app, res } = await dashboard();
  const charts = res.ops.filter(o => o.op === 'create_chart');
  for (const c of charts) {
    const col = app.Tools.letterToNum(String(c.dest).replace(/\d/g, ''));
    // The table occupies A:C, so charts must start at D or later.
    assert.ok(col > 3, `chart anchored at column ${col}, which overlaps the table`);
  }
});

test('dashboard: pie charts get no axis titles', async () => {
  const { res } = await dashboard();
  const pie = res.ops.find(o => o.op === 'create_chart' && o.type === 'Pie');
  assert.strictEqual(pie.xAxisTitle, undefined);
  assert.strictEqual(pie.yAxisTitle, undefined);
  const col = res.ops.find(o => o.op === 'create_chart' && o.type === 'ColumnClustered');
  assert.strictEqual(col.xAxisTitle, 'Region');
  assert.strictEqual(col.yAxisTitle, 'Importe');
});

test('dashboard: charts are skipped when there is no summary table to chart', async () => {
  const { res } = await dashboard({ groupBy: undefined, valueColumn: undefined });
  assert.strictEqual(res.ops.filter(o => o.op === 'create_chart').length, 0);
  assert.ok(res.warnings.some(w => /Charts skipped/.test(w)));
});

test('dashboard: insights are formulas so they stay true after edits', async () => {
  const { res } = await dashboard();
  const insight = res.ops.find(o => o.op === 'write_range' &&
    o.values.some(r => typeof r[0] === 'string' && r[0].includes('Top ')));
  assert.ok(insight, 'expected an insights write op');
  for (const row of insight.values) {
    assert.ok(row[0].startsWith('='), `insight should be a formula: ${row[0]}`);
  }
});

test('dashboard: no two write ops target the same anchor cell', async () => {
  const { res } = await dashboard();
  const anchors = res.ops
    .filter(o => o.op === 'write_range')
    .map(o => `${o.sheet}!${o.range}`);
  assert.strictEqual(new Set(anchors).size, anchors.length,
    `duplicate write anchors: ${anchors.filter((a, i) => anchors.indexOf(a) !== i)}`);
});

test('dashboard: fails cleanly when the source sheet is missing', async () => {
  const app = setup();
  const res = await app.Recipes.dashboard(
    { sheet: 'P', source: { sheet: 'Ghost' }, kpis: [] },
    { snap: snapshot(), readRange: async () => ({ ok: false, error: 'x' }) }
  );
  assert.strictEqual(res.ok, false);
  assert.match(res.error, /does not exist/);
});

test('dashboard: never autofits its own sheet', async () => {
  const { res } = await dashboard();
  // Recipes set deliberate widths; autofit would collapse the merged cards.
  assert.strictEqual(res.ops.filter(o => o.op === 'autofit').length, 0);
});

/* ---------- summary_table recipe ---------- */

test('summary_table: produces a titled, formatted table with no charts', async () => {
  const app = setup();
  const ctx = ctxWith(REGIONS);
  const res = await app.Recipes.summaryTable({
    op: 'recipe.summary_table', sheet: 'Resumen',
    source: { sheet: 'Datos' }, groupBy: 'Region', valueColumn: 'Importe', agg: 'sum'
  }, ctx);
  assert.strictEqual(res.ok, true);
  assert.strictEqual(res.ops[0].op, 'add_sheet');
  assert.strictEqual(res.ops.filter(o => o.op === 'create_chart').length, 0);
  assert.strictEqual(res.ops.filter(o => o.op === 'create_table').length, 1);
});

test('tableName: is unique and contains no spaces', () => {
  const { Recipes } = setup();
  const a = Recipes.tableName('My Dashboard');
  assert.ok(!/\s/.test(a));
  assert.match(a, /^tblMyDashboard\d+$/);
});
