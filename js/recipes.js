/* ============================================
   recipes.js — Deterministic layout generators

   The dashboard structure used to live in the system prompt as ~60 lines
   of prose specifying exact hex colors, font sizes, row heights and merge
   instructions. The model then had to translate that into ~40 individual
   format ops on every single request. It cost thousands of tokens per
   call, and small models got the arithmetic wrong, producing overlapping
   charts and unformatted cells.

   Here it is code. The model picks WHAT to show (which columns, which
   aggregation, which chart types); this file decides WHERE everything
   goes and HOW it looks. Output is consistent every run, and the layout
   arithmetic is unit-testable without Excel.

   A recipe expands into a flat list of ordinary ops, so the executor,
   journal and undo path stay unchanged.
   ============================================ */

const Recipes = {
  /* ---------- design tokens ---------- */
  T: {
    navy:      '#1a237e',
    navyMid:   '#3949ab',
    accent:    '#1a73e8',
    white:     '#FFFFFF',
    cardBg:    '#f5f7fa',
    cardBorder:'#d7dce5',
    label:     '#5f6b7a',
    subtext:   '#98a2b3',
    ink:       '#1f2937',

    bannerHeight: 38,
    spacerHeight: 8,
    kpiLabelHeight: 18,
    kpiValueHeight: 30,
    kpiSubHeight: 16,
    chartWidth: 360,
    chartHeight: 230,

    labelColWidth: 150,
    valueColWidth: 105,
    kpiColWidth: 88
  },

  // Columns each KPI card spans, and the gap between cards.
  KPI_SPAN: 2,
  KPI_GAP: 1,
  // First column of the layout (A).
  START_COL: 1,

  AGGS: { sum: 'SUM', average: 'AVERAGE', count: 'COUNT', min: 'MIN', max: 'MAX' },

  /**
   * Expand a recipe op into plain ops.
   *
   * @param {object} op   — the recipe op from the plan
   * @param {object} ctx  — { snap, readRange(sheet, range) }
   * Returns { ok, ops, warnings } or { ok:false, error }
   */
  async expand(op, ctx) {
    switch (op.op) {
      case 'recipe.dashboard':     return await this.dashboard(op, ctx);
      case 'recipe.summary_table': return await this.summaryTable(op, ctx);
      default: return { ok: false, error: `Unknown recipe: ${op.op}` };
    }
  },

  /* ----------------------------------------------------------
     Shared source resolution
     ---------------------------------------------------------- */

  /**
   * Resolve the source sheet, data range and column positions from header
   * names, using the snapshot. Everything downstream works in absolute
   * addresses so no formula depends on the model's arithmetic.
   */
  resolveSource(op, ctx) {
    const src = op.source || {};
    const sheetName = Context.resolveSheetName(ctx.snap, src.sheet);
    if (!sheetName) {
      return { ok: false, error: `Source sheet "${src.sheet}" does not exist` };
    }
    const desc = Context.findSheet(ctx.snap, sheetName);
    if (!desc || desc.empty) {
      return { ok: false, error: `Source sheet "${sheetName}" is empty` };
    }
    if (!desc.hasHeaders) {
      return { ok: false, error: `Source sheet "${sheetName}" has no header row, so columns cannot be resolved by name` };
    }

    // Trust the sheet's real used range over any range the model supplied:
    // it is authoritative and avoids off-by-N row errors.
    const bounds = Schema.parseAddress(desc.address) || { startRow: 1, endRow: desc.rows };
    const headerRow = bounds.startRow || 1;
    const firstData = headerRow + 1;
    const lastData = bounds.endRow || desc.rows;

    return {
      ok: true,
      sheet: sheetName,
      desc,
      headerRow,
      firstData,
      lastData,
      rowCount: Math.max(0, lastData - firstData + 1)
    };
  },

  /**
   * Absolute A1 range for a named column's data rows, e.g. 'Data'!$D$2:$D$500
   */
  colRange(src, header) {
    const idx = Context.columnIndex(src.desc, header);
    if (idx < 0) return null;
    const letter = Tools.colToLetter(idx + 1);
    return {
      letter,
      idx,
      ref: `'${src.sheet}'!$${letter}$${src.firstData}:$${letter}$${src.lastData}`
    };
  },

  /* ----------------------------------------------------------
     recipe.dashboard
     ---------------------------------------------------------- */
  async dashboard(op, ctx) {
    const src = this.resolveSource(op, ctx);
    if (!src.ok) return { ok: false, error: src.error };

    const warnings = [];
    const ops = [];
    const sheet = op.sheet || 'Dashboard';

    // Validate the KPI columns up front; drop the ones we cannot resolve
    // rather than emitting formulas that would produce #REF!.
    const kpis = [];
    for (const k of (op.kpis || [])) {
      const col = this.colRange(src, k.column);
      if (!col) {
        warnings.push(`KPI "${k.label || k.column}" skipped: column "${k.column}" not found on "${src.sheet}"`);
        continue;
      }
      const fn = this.AGGS[String(k.agg || 'sum').toLowerCase()] || 'SUM';
      kpis.push({
        label: k.label || `${k.agg || 'sum'} of ${k.column}`,
        formula: `=${fn}(${col.ref})`,
        format: k.format || this.formatFor(src.desc, col.idx, fn),
        subtext: k.subtext || `${fn} · ${k.column}`
      });
      if (kpis.length >= 4) break;
    }

    // Aggregate the category breakdown. Reading is free via Office.js, so
    // the distinct categories are computed locally instead of being
    // guessed by the model.
    let breakdown = null;
    if (op.groupBy && op.valueColumn) {
      breakdown = await this.aggregate(op, ctx, src);
      if (!breakdown.ok) {
        warnings.push(breakdown.error);
        breakdown = null;
      }
    }

    ops.push({ op: 'add_sheet', name: sheet, tabColor: this.T.navy });

    const layout = this.planDashboardLayout({ kpis, breakdown, charts: op.charts || [] });
    let row = 1;

    /* --- title banner --- */
    const title = op.title || `DASHBOARD — ${src.sheet}`;
    const bannerEnd = Tools.colToLetter(layout.width);
    ops.push(
      { op: 'write_range', sheet, range: 'A1', values: [[title]] },
      { op: 'format_range', sheet, range: `A1:${bannerEnd}1`,
        merge: true, bold: true, fontSize: 16, fontColor: this.T.white,
        fillColor: this.T.navy, horizontalAlignment: 'Center',
        verticalAlignment: 'Center', rowHeight: this.T.bannerHeight }
    );
    row = 2;
    ops.push({ op: 'format_range', sheet, range: `A${row}:${bannerEnd}${row}`, rowHeight: this.T.spacerHeight });
    row++;

    /* --- KPI cards --- */
    if (kpis.length > 0) {
      const labelRow = row, valueRow = row + 1, subRow = row + 2;
      kpis.forEach((k, i) => {
        const startCol = this.START_COL + i * (this.KPI_SPAN + this.KPI_GAP);
        const a = Tools.colToLetter(startCol);
        const b = Tools.colToLetter(startCol + this.KPI_SPAN - 1);

        ops.push(
          { op: 'write_range', sheet, range: `${a}${labelRow}`, values: [[k.label.toUpperCase()]] },
          { op: 'write_range', sheet, range: `${a}${valueRow}`, values: [[k.formula]], numberFormat: k.format },
          { op: 'write_range', sheet, range: `${a}${subRow}`, values: [[k.subtext]] },
          { op: 'format_range', sheet, range: `${a}${labelRow}:${b}${labelRow}`,
            merge: true, bold: true, fontSize: 9, fontColor: this.T.label,
            fillColor: this.T.cardBg, horizontalAlignment: 'Center',
            rowHeight: this.T.kpiLabelHeight,
            borders: { EdgeTop: 'Thin', EdgeLeft: 'Thin', EdgeRight: 'Thin' } },
          { op: 'format_range', sheet, range: `${a}${valueRow}:${b}${valueRow}`,
            merge: true, bold: true, fontSize: 18, fontColor: this.T.accent,
            fillColor: this.T.cardBg, horizontalAlignment: 'Center',
            verticalAlignment: 'Center', numberFormat: k.format,
            rowHeight: this.T.kpiValueHeight,
            borders: { EdgeLeft: 'Thin', EdgeRight: 'Thin' } },
          { op: 'format_range', sheet, range: `${a}${subRow}:${b}${subRow}`,
            merge: true, fontSize: 8, fontColor: this.T.subtext,
            fillColor: this.T.cardBg, horizontalAlignment: 'Center',
            rowHeight: this.T.kpiSubHeight,
            borders: { EdgeBottom: 'Thin', EdgeLeft: 'Thin', EdgeRight: 'Thin' } }
        );
      });
      row = subRow + 1;
      ops.push({ op: 'format_range', sheet, range: `A${row}:${bannerEnd}${row}`, rowHeight: this.T.spacerHeight });
      row++;
    }

    /* --- summary table (+ charts anchored beside it) --- */
    let tableInfo = null;
    if (breakdown) {
      const built = this.buildSummaryBlock({
        sheet, startRow: row, breakdown, src, op,
        tableName: this.tableName(sheet)
      });
      ops.push(...built.ops);
      tableInfo = built;
      row = built.endRow + 1;
    }

    /* --- charts --- */
    const charts = (op.charts || []).slice(0, 3);
    if (charts.length > 0 && tableInfo) {
      const anchorCol = Tools.colToLetter(tableInfo.chartAnchorCol);
      charts.forEach((c, i) => {
        const anchorRow = tableInfo.startRow + i * this.chartRowSpan();
        ops.push({
          op: 'create_chart',
          sheet,
          type: this.chartType(c.type),
          sourceRange: tableInfo.chartDataRange,
          title: c.title || `${op.valueColumn} by ${op.groupBy}`,
          dest: `${anchorCol}${anchorRow}`,
          width: this.T.chartWidth,
          height: this.T.chartHeight,
          xAxisTitle: this.isCircular(c.type) ? undefined : op.groupBy,
          yAxisTitle: this.isCircular(c.type) ? undefined : op.valueColumn
        });
      });
      // Keep the insights block below whichever is taller: table or charts.
      const chartsEndRow = tableInfo.startRow + charts.length * this.chartRowSpan();
      row = Math.max(row, chartsEndRow) + 1;
    } else if (charts.length > 0) {
      warnings.push('Charts skipped: a summary table is required as the chart source (provide groupBy and valueColumn)');
    }

    /* --- insights, as live formulas over the summary table --- */
    if (tableInfo) {
      ops.push(...this.insightOps({ sheet, startRow: row, tableInfo, op, bannerEnd }));
      row += 5;
    }

    /* --- column widths --- */
    ops.push(...this.columnWidthOps(sheet, layout, kpis.length));

    return { ok: true, ops, warnings, createdSheet: sheet };
  },

  /* ----------------------------------------------------------
     recipe.summary_table — the graceful-degradation path
     ---------------------------------------------------------- */
  async summaryTable(op, ctx) {
    const src = this.resolveSource(op, ctx);
    if (!src.ok) return { ok: false, error: src.error };

    const breakdown = await this.aggregate(op, ctx, src);
    if (!breakdown.ok) return { ok: false, error: breakdown.error };

    const sheet = op.sheet || 'Summary';
    const ops = [{ op: 'add_sheet', name: sheet, tabColor: this.T.navyMid }];

    const title = op.title || `${op.valueColumn} BY ${op.groupBy}`.toUpperCase();
    ops.push(
      { op: 'write_range', sheet, range: 'A1', values: [[title]] },
      { op: 'format_range', sheet, range: 'A1:C1', merge: true, bold: true,
        fontSize: 14, fontColor: this.T.white, fillColor: this.T.navy,
        horizontalAlignment: 'Center', rowHeight: this.T.bannerHeight }
    );

    const built = this.buildSummaryBlock({
      sheet, startRow: 3, breakdown, src, op, tableName: this.tableName(sheet)
    });
    ops.push(...built.ops);
    ops.push(
      { op: 'format_range', sheet, range: 'A:A', columnWidth: this.T.labelColWidth },
      { op: 'format_range', sheet, range: 'B:C', columnWidth: this.T.valueColWidth }
    );

    return { ok: true, ops, warnings: [], createdSheet: sheet };
  },

  /* ----------------------------------------------------------
     Aggregation — done locally over Office.js reads (no API cost)
     ---------------------------------------------------------- */

  /**
   * Read the group-by and value columns and compute the distinct
   * categories. Only the category list is needed: the written cells use
   * live SUMIF/AVERAGEIF formulas so the table stays correct when the
   * source data changes.
   */
  async aggregate(op, ctx, src) {
    const groupCol = this.colRange(src, op.groupBy);
    if (!groupCol) return { ok: false, error: `Group column "${op.groupBy}" not found on "${src.sheet}"` };
    const valueCol = this.colRange(src, op.valueColumn);
    if (!valueCol) return { ok: false, error: `Value column "${op.valueColumn}" not found on "${src.sheet}"` };

    const read = await ctx.readRange(
      src.sheet,
      `${groupCol.letter}${src.firstData}:${groupCol.letter}${src.lastData}`
    );
    if (!read.ok) return { ok: false, error: `Could not read "${op.groupBy}": ${read.error}` };

    const seen = new Map();
    for (const row of (read.data || [])) {
      const v = Array.isArray(row) ? row[0] : row;
      if (v === null || v === undefined || v === '') continue;
      const key = String(v);
      if (!seen.has(key)) seen.set(key, 0);
      seen.set(key, seen.get(key) + 1);
    }

    if (seen.size === 0) {
      return { ok: false, error: `Column "${op.groupBy}" has no non-empty values to group by` };
    }

    // Cap categories so a high-cardinality column (an id, a timestamp)
    // cannot produce a thousand-row "summary".
    const MAX_CATEGORIES = 25;
    let categories = [...seen.keys()];
    let truncated = false;
    if (categories.length > MAX_CATEGORIES) {
      categories = categories
        .sort((a, b) => seen.get(b) - seen.get(a))
        .slice(0, MAX_CATEGORIES);
      truncated = true;
    } else {
      categories.sort((a, b) => a.localeCompare(b));
    }

    const agg = String(op.agg || 'sum').toLowerCase();
    return {
      ok: true,
      categories,
      truncated,
      partial: !!read.truncated,
      groupCol,
      valueCol,
      agg: this.AGGS[agg] ? agg : 'sum'
    };
  },

  /**
   * The summary table block: header row, one formula row per category, a
   * total row, an Excel Table, and data bars on the value column.
   */
  buildSummaryBlock({ sheet, startRow, breakdown, src, op, tableName }) {
    const ops = [];
    const { categories, groupCol, valueCol, agg } = breakdown;

    const headerRow = startRow;
    const firstRow = headerRow + 1;
    const lastRow = headerRow + categories.length;
    const totalRow = lastRow + 1;

    const ifFn = { sum: 'SUMIF', average: 'AVERAGEIF', count: 'COUNTIF', min: 'SUMIF', max: 'SUMIF' }[agg] || 'SUMIF';
    const groupRef = `'${src.sheet}'!$${groupCol.letter}$${src.firstData}:$${groupCol.letter}$${src.lastData}`;
    const valueRef = `'${src.sheet}'!$${valueCol.letter}$${src.firstData}:$${valueCol.letter}$${src.lastData}`;

    // Header
    const valueHeader = `${agg.charAt(0).toUpperCase() + agg.slice(1)} of ${op.valueColumn}`;
    ops.push(
      { op: 'write_range', sheet, range: `A${headerRow}`,
        values: [[op.groupBy, valueHeader, '% of total']] },
      { op: 'format_range', sheet, range: `A${headerRow}:C${headerRow}`,
        bold: true, fontColor: this.T.white, fillColor: this.T.navyMid,
        horizontalAlignment: 'Center', rowHeight: 20,
        borders: { EdgeBottom: 'Thin' } }
    );

    // One row per category, all live formulas keyed off the label cell so
    // editing a label re-points its own aggregate.
    const rows = categories.map((cat, i) => {
      const rowNum = firstRow + i;
      const catCell = `A${rowNum}`;
      const valueFormula = ifFn === 'COUNTIF'
        ? `=COUNTIF(${groupRef},${catCell})`
        : `=${ifFn}(${groupRef},${catCell},${valueRef})`;
      return [String(cat), valueFormula, `=IFERROR(B${rowNum}/$B$${totalRow},0)`];
    });
    ops.push({ op: 'write_range', sheet, range: `A${firstRow}`, values: rows });

    const numberFormat = op.valueFormat || this.formatFor(src.desc, valueCol.idx, agg);
    ops.push(
      { op: 'format_range', sheet, range: `B${firstRow}:B${lastRow}`, numberFormat },
      { op: 'format_range', sheet, range: `C${firstRow}:C${lastRow}`, numberFormat: '0.0%' },
      { op: 'format_range', sheet, range: `A${firstRow}:C${lastRow}`,
        borders: { EdgeLeft: 'Thin', EdgeRight: 'Thin', EdgeBottom: 'Thin' } }
    );

    // Total row. For an average, totalling the column would be wrong, so
    // aggregate the source directly instead of summing the rows above.
    const totalFormula = agg === 'average'
      ? `=IFERROR(AVERAGE(${valueRef}),0)`
      : `=SUM(B${firstRow}:B${lastRow})`;
    ops.push(
      { op: 'write_range', sheet, range: `A${totalRow}`,
        values: [['TOTAL', totalFormula, `=IFERROR(SUM(C${firstRow}:C${lastRow}),0)`]] },
      { op: 'format_range', sheet, range: `A${totalRow}:C${totalRow}`,
        bold: true, fillColor: this.T.cardBg, borders: { EdgeTop: 'Thin', EdgeBottom: 'Thin' } },
      { op: 'format_range', sheet, range: `B${totalRow}`, numberFormat },
      { op: 'format_range', sheet, range: `C${totalRow}`, numberFormat: '0.0%' }
    );

    // Excel Table over the category rows only (not the total).
    ops.push({ op: 'create_table', sheet, range: `A${headerRow}:C${lastRow}`,
               name: tableName, style: 'TableStyleMedium2' });
    ops.push({ op: 'conditional_format', sheet, range: `B${firstRow}:B${lastRow}`, type: 'dataBar' });

    return {
      ops,
      startRow: headerRow,
      headerRow,
      firstRow,
      lastRow,
      totalRow,
      endRow: totalRow,
      // Charts read the category + value columns, header included so the
      // series picks up its name.
      chartDataRange: `A${headerRow}:B${lastRow}`,
      chartAnchorCol: 5, // column E — clear of the 3-column table
      valueColumnLetter: 'B',
      numberFormat
    };
  },

  /**
   * Insights as formulas rather than baked-in text, so they stay true
   * after the source data changes.
   */
  insightOps({ sheet, startRow, tableInfo, op, bannerEnd }) {
    const { firstRow, lastRow, totalRow } = tableInfo;
    const valRange = `$B$${firstRow}:$B$${lastRow}`;
    const catRange = `$A$${firstRow}:$A$${lastRow}`;
    const fmt = tableInfo.numberFormat;

    const topName = `INDEX(${catRange},MATCH(MAX(${valRange}),${valRange},0))`;
    const lowName = `INDEX(${catRange},MATCH(MIN(${valRange}),${valRange},0))`;

    const lines = [
      [`=IFERROR("Top ${op.groupBy}: "&${topName}&" with "&TEXT(MAX(${valRange}),"${fmt}")&" ("&TEXT(IFERROR(MAX(${valRange})/$B$${totalRow},0),"0.0%")&" of total)","")`],
      [`=IFERROR("Lowest ${op.groupBy}: "&${lowName}&" with "&TEXT(MIN(${valRange}),"${fmt}"),"")`],
      [`=IFERROR("Spread: "&COUNT(${valRange})&" categories, average "&TEXT(AVERAGE(${valRange}),"${fmt}")&", top ${op.groupBy} is "&TEXT(IFERROR(MAX(${valRange})/AVERAGE(${valRange}),0),"0.0")&"x the average","")`]
    ];

    return [
      { op: 'write_range', sheet, range: `A${startRow}`, values: [['KEY OBSERVATIONS']] },
      { op: 'format_range', sheet, range: `A${startRow}:${bannerEnd}${startRow}`,
        merge: true, bold: true, fontSize: 11, fontColor: this.T.navy, rowHeight: 20 },
      { op: 'write_range', sheet, range: `A${startRow + 1}`, values: lines },
      { op: 'format_range', sheet, range: `A${startRow + 1}:A${startRow + 3}`,
        fontSize: 10, fontColor: this.T.ink }
    ];
  },

  /* ----------------------------------------------------------
     Layout arithmetic
     ---------------------------------------------------------- */

  /**
   * Overall width of the dashboard in columns, wide enough for the KPI
   * cards, the summary table and the charts beside it.
   */
  planDashboardLayout({ kpis, breakdown, charts }) {
    const kpiWidth = kpis.length > 0
      ? kpis.length * this.KPI_SPAN + (kpis.length - 1) * this.KPI_GAP
      : 0;
    // Table occupies A:C; charts start at E and span ~5 columns.
    const contentWidth = (breakdown ? 3 : 0);
    const chartWidth = (charts && charts.length > 0 && breakdown) ? 4 + 5 : 0;
    return {
      width: Math.max(kpiWidth, contentWidth, chartWidth, 6),
      kpiWidth,
      kpiCount: kpis.length
    };
  },

  /**
   * Rows a chart occupies, so stacked charts never overlap. Excel rows
   * default to ~15 points.
   */
  chartRowSpan() {
    return Math.ceil(this.T.chartHeight / 15) + 1;
  },

  columnWidthOps(sheet, layout, kpiCount) {
    const ops = [];
    if (kpiCount > 0) {
      // KPI cards are merged blocks: give every spanned column equal width.
      const lastKpiCol = Tools.colToLetter(Math.max(1, layout.kpiWidth));
      ops.push({ op: 'format_range', sheet, range: `A:${lastKpiCol}`, columnWidth: this.T.kpiColWidth });
    }
    // The label column carries category names, so make it wider.
    ops.push({ op: 'format_range', sheet, range: 'A:A', columnWidth: this.T.labelColWidth });
    return ops;
  },

  /* ----------------------------------------------------------
     Helpers
     ---------------------------------------------------------- */

  /**
   * Pick a number format from the source column's detected type, so
   * currency stays currency and counts stay integers.
   */
  formatFor(desc, colIdx, agg) {
    const fn = String(agg || '').toUpperCase();
    if (fn === 'COUNT') return '#,##0';
    const type = desc && desc.columnTypes ? desc.columnTypes[colIdx] : null;
    switch (type) {
      case 'currency': return '$#,##0.00';
      case 'percent':  return '0.0%';
      case 'date':     return 'yyyy-mm-dd';
      case 'number':   return fn === 'AVERAGE' || fn === 'average' ? '#,##0.00' : '#,##0';
      default:         return '#,##0.00';
    }
  },

  CHART_TYPES: {
    columnclustered: 'ColumnClustered',
    column: 'ColumnClustered',
    bar: 'BarClustered',
    barclustered: 'BarClustered',
    line: 'Line',
    linemarkers: 'LineMarkers',
    pie: 'Pie',
    doughnut: 'Doughnut',
    area: 'Area',
    scatter: 'XYScatter'
  },

  chartType(t) {
    const key = String(t || 'columnClustered').toLowerCase().replace(/[^a-z]/g, '');
    return this.CHART_TYPES[key] || 'ColumnClustered';
  },

  isCircular(t) {
    const norm = this.chartType(t);
    return norm === 'Pie' || norm === 'Doughnut';
  },

  /**
   * Table names must be unique in the workbook and cannot contain spaces.
   */
  tableName(sheet) {
    const base = 'tbl' + String(sheet).replace(/[^A-Za-z0-9]/g, '');
    return `${base}${Date.now().toString().slice(-5)}`;
  }
};
