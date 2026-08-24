/* ============================================
   tools.js — Tool registry for the agentic loop
   - handlers: async (args) -> { ok, result } | { ok:false, error }
   Every mutating tool records pre-images with the Journal so the user
   can undo the whole request. Every tool returns verified post-state so
   the verifier can detect #REF! errors after a write.
   ============================================ */

const Tools = {
  // Cap to keep tool results small enough for the model's context.
  READ_CELL_CAP: 5000,
  READ_ROW_CAP: 200,
  READ_COL_CAP: 50,

  // Session-level cache for get_workbook_overview. The overview is expensive
  // to compute (reads ALL sheets) and rarely changes within a session.
  // Invalidated when sheets are added/deleted or the user clears chat.
  _overviewCache: null,

  invalidateOverviewCache() {
    this._overviewCache = null;
  },

  /* ----------------------------------------------------------
     DISPATCH
     ---------------------------------------------------------- */
  async dispatch(name, args) {
    const handler = this.handlers[name];
    if (!handler) return { ok: false, error: `Unknown tool: ${name}` };
    try {
      return await handler.call(this, args || {});
    } catch (e) {
      return { ok: false, error: e.message || String(e) };
    }
  },

  // Names of mutating tools (for journal pre-image decisions).
  MUTATING: new Set([
    'write_range', 'format_range', 'clear_range', 'add_sheet', 'delete_sheet',
    'create_table', 'create_pivot', 'create_chart', 'add_slicer',
    'conditional_format', 'insert_rows_cols', 'delete_rows_cols', 'sort_range'
  ]),

  isMutating(name) {
    return this.MUTATING.has(name);
  },

  /* ----------------------------------------------------------
     HANDLERS
     Each returns { ok, result } or { ok:false, error }.
     ---------------------------------------------------------- */
  handlers: {
    // ---------- READ ----------
    async get_workbook_overview() {
      // Return cached overview if available (saves reading all sheets again).
      if (this._overviewCache) {
        return { ok: true, result: this._overviewCache, cached: true };
      }
      const snap = await Schema.snapshot();
      const result = { overview: Schema.toText(snap), sheets: snap.sheets.map(s => ({ name: s.name, rows: s.rows, cols: s.cols, address: s.address })) };
      this._overviewCache = result;
      return { ok: true, result };
    },

    async read_range({ sheet, range, what = 'values' }) {
      // Bound the read to keep payload small.
      const bounded = await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const r = s.getRange(range);
        r.load('address, rowCount, columnCount');
        await ctx.sync();

        let rowCount = r.rowCount;
        let colCount = r.columnCount;
        let target = r;
        let truncated = false;

        if (rowCount * colCount > Tools.READ_CELL_CAP) {
          // Truncate rows first, then cols.
          const startCell = range.split(':')[0];
          rowCount = Math.min(rowCount, Tools.READ_ROW_CAP);
          colCount = Math.min(colCount, Tools.READ_COL_CAP);
          const endCell = Tools.offsetRange(startCell, rowCount - 1, colCount - 1);
          target = s.getRange(`${startCell}:${endCell}`);
          truncated = true;
          target.load('values, formulas, numberFormats');
          await ctx.sync();
        } else {
          target.load('values, formulas, numberFormats');
          await ctx.sync();
        }

        let data;
        if (what === 'formulas') data = target.formulas;
        else if (what === 'formats') data = target.numberFormats;
        else data = target.values;

        return { address: r.address, rowCount, colCount, truncated, data };
      });
      return { ok: true, result: bounded };
    },

    async find_in_workbook({ query, sheet }) {
      const matches = [];
      const sheetsToSearch = sheet ? [sheet] : await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        ws.load('items/name, items/visibility');
        await ctx.sync();
        return (ws.items || []).filter(s => s.visibility === 'Visible' || s.visibility === undefined).map(s => s.name);
      });

      const qLower = String(query).toLowerCase();
      for (const sh of sheetsToSearch) {
        try {
          await Excel.run(async (ctx) => {
            const s = ctx.workbook.worksheets.getItem(sh);
            const used = s.getUsedRangeOrNullObject();
            used.load('values, address, rowCount, columnCount');
            await ctx.sync();
            if (used.isNullObject) return;
            // Cap scan to avoid huge sheets.
            const maxRows = Math.min(used.rowCount, 5000);
            const maxCols = Math.min(used.columnCount, 50);
            for (let r = 0; r < maxRows; r++) {
              for (let c = 0; c < maxCols; c++) {
                const v = used.values[r] ? used.values[r][c] : null;
                if (v === null || v === '') continue;
                if (String(v).toLowerCase().includes(qLower)) {
                  const addr = Tools.addressFromIndex(used.address, r, c);
                  matches.push({ sheet: sh, address: addr, value: v });
                  if (matches.length >= 20) return;
                }
              }
            }
          });
          if (matches.length >= 20) break;
        } catch (e) { /* skip unreadable sheet */ }
      }
      return { ok: true, result: { matches, truncated: matches.length >= 20 } };
    },

    async get_objects({ sheet }) {
      const result = await Excel.run(async (ctx) => {
        const out = { tables: [], pivots: [], charts: [], slicers: [], names: [] };
        const wantSheet = sheet ? String(sheet).toLowerCase() : null;

        // Workbook-scoped collections — load all properties in one batch.
        const tables = ctx.workbook.tables;
        tables.load('items/name, items/worksheet/name, items/range/address');
        const pivots = ctx.workbook.pivotTables;
        pivots.load('items/name, items/worksheet/name, items/layout/range');
        const slicers = ctx.workbook.slicers;
        slicers.load('items/name, items/worksheet/name, items/caption');
        const names = ctx.workbook.names;
        names.load('items/name, items/value');
        await ctx.sync();

        for (const t of tables.items) {
          let tSheet = '';
          try { tSheet = t.worksheet.name || ''; } catch (e) {}
          if (wantSheet && tSheet.toLowerCase() !== wantSheet) continue;
          out.tables.push({ name: t.name, sheet: tSheet, range: t.range.address });
        }
        for (const p of pivots.items) {
          let pSheet = '';
          let pRange = '';
          try { pSheet = p.worksheet.name || ''; } catch (e) {}
          try { pRange = p.layout.range || ''; } catch (e) {}
          if (wantSheet && pSheet.toLowerCase() !== wantSheet) continue;
          out.pivots.push({ name: p.name, sheet: pSheet, range: pRange });
        }
        for (const sl of slicers.items) {
          let slSheet = '';
          try { slSheet = sl.worksheet.name || ''; } catch (e) {}
          if (wantSheet && slSheet.toLowerCase() !== wantSheet) continue;
          out.slicers.push({ name: sl.name, sheet: slSheet, caption: sl.caption });
        }
        for (const n of names.items) {
          out.names.push({ name: n.name, value: n.value });
        }

        // Charts are sheet-scoped.
        if (sheet) {
          const s = ctx.workbook.worksheets.getItem(sheet);
          const charts = s.charts;
          charts.load('items/name, items/chartType');
          await ctx.sync();
          for (const c of charts.items) {
            out.charts.push({ name: c.name, sheet, type: c.chartType });
          }
        } else {
          const ws = ctx.workbook.worksheets;
          ws.load('items/name, items/visibility');
          await ctx.sync();
          for (const w of ws.items) {
            if (w.visibility !== 'Visible' && w.visibility !== undefined) continue;
            try {
              const charts = w.charts;
              charts.load('items/name, items/chartType');
              await ctx.sync();
              for (const c of charts.items) {
                out.charts.push({ name: c.name, sheet: w.name, type: c.chartType });
              }
            } catch (e) {}
          }
        }

        return out;
      });
      return { ok: true, result };
    },

    // ---------- WRITE / EDIT ----------
    async write_range({ sheet, range, values, numberFormat }) {
      if (!Array.isArray(values) || !values.every(r => Array.isArray(r))) {
        return { ok: false, error: 'values must be a 2D array.' };
      }
      if (values.length === 0 || !values[0] || values[0].length === 0) {
        return { ok: false, error: 'values array is empty — nothing to write.' };
      }

      // Normalize ragged arrays: pad short rows with null so every row has
      // the same column count. Office.js requires a rectangular matrix.
      const maxCols = Math.max(...values.map(r => r.length));
      const normalized = values.map(r => {
        if (r.length === maxCols) return r;
        const padded = r.slice();
        while (padded.length < maxCols) padded.push(null);
        return padded;
      });

      const startCell = range.split(':')[0];
      const valRows = normalized.length;
      const valCols = maxCols;
      const endCell = Tools.offsetRange(startCell, valRows - 1, valCols - 1);
      const fullRange = `${startCell}:${endCell}`;

      // Pre-image for undo.
      await Journal.recordRangePreImage(sheet, fullRange);

      const hasFormulas = normalized.some(row => Array.isArray(row) && row.some(v => typeof v === 'string' && v.startsWith('=')));

      // Office.js requires formulas to be a 2D array of strings only.
      // Convert null/number/boolean to strings; formula cells pass through.
      const formulaMatrix = hasFormulas
        ? normalized.map(row => row.map(v => {
            if (v === null || v === undefined) return '';
            if (typeof v === 'string') return v;
            return String(v);
          }))
        : null;

      const post = await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const r = s.getRange(fullRange);
        if (hasFormulas) r.formulas = formulaMatrix;
        else r.values = normalized;
        if (numberFormat) r.numberFormat = numberFormat;
        // Read back a sample to surface formula errors.
        r.load('values, rowCount, columnCount');
        await ctx.sync();
        const sampleRows = Math.min(r.rowCount, 5);
        const sampleCols = Math.min(r.columnCount, 5);
        const sample = [];
        for (let i = 0; i < sampleRows; i++) {
          sample.push(r.values[i].slice(0, sampleCols));
        }
        return { address: fullRange, rowCount: r.rowCount, columnCount: r.columnCount, sample };
      });
      return { ok: true, result: post };
    },

    async format_range(action) {
      await Journal.recordRangePreImage(action.sheet, action.range);
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(action.sheet);
        const range = s.getRange(action.range);
        const fmt = range.format;

        const safeSet = (setter) => { try { setter(); } catch (e) { /* skip invalid arg */ } };
        if (action.bold !== undefined) safeSet(() => fmt.font.bold = action.bold);
        if (action.italic !== undefined) safeSet(() => fmt.font.italic = action.italic);
        if (action.fontSize !== undefined && typeof action.fontSize === 'number') safeSet(() => fmt.font.size = action.fontSize);
        if (action.fontName !== undefined && typeof action.fontName === 'string') safeSet(() => fmt.font.name = action.fontName);
        if (action.fontColor !== undefined) {
          const sc = Tools.sanitizeColor(action.fontColor);
          if (sc) safeSet(() => fmt.font.color = sc);
        }
        if (action.fillColor !== undefined) {
          const sc = Tools.sanitizeColor(action.fillColor);
          if (sc) safeSet(() => fmt.fill.color = sc);
        }
        if (action.horizontalAlignment !== undefined) safeSet(() => fmt.horizontalAlignment = action.horizontalAlignment);
        if (action.verticalAlignment !== undefined) safeSet(() => fmt.verticalAlignment = action.verticalAlignment);
        if (action.wrapText !== undefined) safeSet(() => fmt.wrapText = action.wrapText);
        if (action.numberFormat !== undefined && typeof action.numberFormat === 'string') safeSet(() => range.numberFormat = action.numberFormat);
        if (action.columnWidth !== undefined && typeof action.columnWidth === 'number' && action.columnWidth > 0) safeSet(() => fmt.columnWidth = action.columnWidth);
        if (action.rowHeight !== undefined && typeof action.rowHeight === 'number' && action.rowHeight > 0) safeSet(() => fmt.rowHeight = action.rowHeight);
        if (action.merge) safeSet(() => range.merge(true));
        if (action.borders) safeSet(() => Tools.applyBorders(ctx, range, action.borders));
        await ctx.sync();
      });
      return { ok: true, result: { sheet: action.sheet, range: action.range } };
    },

    async clear_range({ sheet, range }) {
      await Journal.recordRangePreImage(sheet, range);
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        s.getRange(range).clear();
        await ctx.sync();
      });
      return { ok: true, result: { sheet, range } };
    },

    async add_sheet({ name, tabColor }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.add(name);
        s.activate();
        s.tabColor = tabColor || '#4a9eff';
        await ctx.sync();
      });
      Journal.recordCreatedObject('sheet', name, name);
      this.invalidateOverviewCache(); // structure changed
      return { ok: true, result: { name } };
    },

    async delete_sheet({ name, userRequested }) {
      // Safety: allow if user explicitly asked, or if created this request.
      const createdThisRequest = Journal.current &&
        Journal.current.entries.some(e => e.type === 'created' && e.kind === 'sheet' && e.name === name);
      if (!userRequested && !createdThisRequest) {
        return { ok: false, error: `Refused to delete sheet "${name}" — it was not created in this request and the user did not explicitly ask to delete it. Ask the user to confirm before deleting existing sheets.` };
      }
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(name);
        s.delete();
        await ctx.sync();
      });
      // Best-effort undo: we can't restore data, but record a marker.
      if (Journal.current) Journal.current.entries.push({ type: 'modified', kind: 'sheet', name, sheet: name });
      this.invalidateOverviewCache(); // structure changed
      return { ok: true, result: { name, note: 'Sheet deleted. Data on it is not recoverable via undo.' } };
    },

    async create_table({ sheet, range, name, style }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const r = s.getRange(range);
        const t = ctx.workbook.tables.add(r, true);
        t.name = name;
        if (style) t.style = style;
        await ctx.sync();
      });
      Journal.recordCreatedObject('table', name, sheet);
      return { ok: true, result: { name, sheet, range } };
    },

    async create_pivot({ destSheet, dest, source, name, rows, cols, values, filters }) {
      await Excel.run(async (ctx) => {
        const ds = ctx.workbook.worksheets.getItem(destSheet);
        const pivot = ctx.workbook.pivotTables.add(name, source, ds.getRange(dest));
        if (Array.isArray(rows)) for (const f of rows) pivot.rowFields.add(f);
        if (Array.isArray(cols)) for (const f of cols) pivot.columnFields.add(f);
        if (Array.isArray(values)) {
          for (const v of values) {
            const vf = pivot.dataFields.add(v.col);
            if (v.agg) vf.summarizeBy = Tools.mapAggregation(v.agg);
          }
        }
        if (Array.isArray(filters)) for (const f of filters) pivot.filterFields.add(f);
        await ctx.sync();
      });
      Journal.recordCreatedObject('pivot', name, destSheet);
      return { ok: true, result: { name, destSheet, dest } };
    },

    async create_chart(action) {
      const chartInfo = await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(action.sheet);
        let chart;
        if (action.sourceRange) {
          const r = s.getRange(action.sourceRange);
          chart = s.charts.add(action.type, r, action.seriesBy || 'Auto');
        } else if (action.sourcePivot) {
          const pivot = ctx.workbook.pivotTables.getItem(action.sourcePivot);
          pivot.load('layout/range');
          await ctx.sync();
          const rangeStr = pivot.layout.range;
          const sourceSheet = ctx.workbook.worksheets.getItem(action.sourceSheet || action.sheet);
          const r = sourceSheet.getRange(rangeStr);
          chart = s.charts.add(action.type, r, action.seriesBy || 'Auto');
        } else {
          throw new Error('create_chart requires sourceRange or sourcePivot');
        }
        if (action.title) chart.title.text = action.title;
        if (action.dest) chart.setPosition(action.dest);
        chart.width = action.width || 380;
        chart.height = action.height || 240;
        // Axis titles are best-effort: not every chart type exposes both
        // axes, and pie/doughnut expose neither.
        const safeSet = (setter) => { try { setter(); } catch (e) { /* axis unavailable */ } };
        if (action.xAxisTitle) {
          safeSet(() => {
            chart.axes.categoryAxis.title.text = action.xAxisTitle;
            chart.axes.categoryAxis.title.visible = true;
          });
        }
        if (action.yAxisTitle) {
          safeSet(() => {
            chart.axes.valueAxis.title.text = action.yAxisTitle;
            chart.axes.valueAxis.title.visible = true;
          });
        }
        if (action.showLegend === false) safeSet(() => chart.legend.visible = false);
        chart.load('name');
        await ctx.sync();
        return { name: chart.name || 'chart' };
      });
      Journal.recordCreatedObject('chart', chartInfo.name, action.sheet);
      return { ok: true, result: { name: chartInfo.name, sheet: action.sheet, dest: action.dest } };
    },

    async add_slicer({ sheet, sourcePivot, sourceTable, field, dest, name }) {
      const slicerInfo = await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const source = sourcePivot || sourceTable;
        if (!source) throw new Error('add_slicer requires sourcePivot or sourceTable');
        const slicer = ctx.workbook.slicers.add(source, field, s.getRange(dest));
        if (name) slicer.name = name;
        slicer.load('name');
        await ctx.sync();
        return { name: slicer.name || 'slicer' };
      });
      Journal.recordCreatedObject('slicer', slicerInfo.name, sheet);
      return { ok: true, result: slicerInfo };
    },

    async conditional_format({ sheet, range, type, rules }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const r = s.getRange(range);
        if (type === 'colorScale') {
          r.conditionalFormats.add('ColorScale', {
            threeColorScale: {
              minimum: { color: '#63BE7B', type: 'Min' },
              midpoint: { color: '#FFEB84', type: 'Percentile', percentile: 50 },
              maximum: { color: '#F8696B', type: 'Max' }
            }
          });
        } else if (type === 'dataBar') {
          r.conditionalFormats.add('DataBar', {
            dataBar: { barDirection: 'LeftToRight', color: '#4a9eff' }
          });
        } else if (type === 'cellValue' && Array.isArray(rules)) {
          for (const rule of rules) {
            r.conditionalFormats.add('CellValue', {
              cellValue: { operator: rule.operator || 'GreaterThan', formula: [rule.value] },
              format: { fill: { color: rule.color || '#FFC7CE' } }
            });
          }
        } else {
          throw new Error(`Unsupported conditional_format type: ${type}`);
        }
        await ctx.sync();
      });
      return { ok: true, result: { sheet, range, type } };
    },

    async autofit({ sheet, cols, rows }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const used = s.getUsedRangeOrNullObject();
        used.load('address, rowCount, columnCount');
        await ctx.sync();
        if (used.isNullObject && !cols && !rows) return;

        const colLetters = cols || (() => {
          const out = [];
          for (let c = 0; c < used.columnCount; c++) out.push(Tools.colToLetter(c + 1));
          return out;
        })();
        const rowCount = rows ? Math.max(...rows) : used.rowCount;

        for (const col of colLetters) {
          const colRange = s.getRange(`${col}1:${col}${rowCount}`);
          colRange.format.autofitColumns();
        }
        if (rows) {
          for (const rn of rows) {
            const rowRange = s.getRange(`A${rn}:${Tools.colToLetter(used.columnCount || 1)}${rn}`);
            rowRange.format.autofitRows();
          }
        } else if (!used.isNullObject) {
          for (let r = 0; r < used.rowCount; r++) {
            const rn = r + 1;
            const rowRange = s.getRange(`A${rn}:${Tools.colToLetter(used.columnCount)}${rn}`);
            rowRange.format.autofitRows();
          }
        }
        await ctx.sync();
      });
      return { ok: true, result: { sheet } };
    },

    async insert_rows_cols({ sheet, kind, at, count = 1 }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        if (kind === 'rows') {
          const r = s.getRange(`${at}:${at + count - 1}`);
          r.insert('Down');
        } else {
          const startCol = Tools.colToLetter(at);
          const endCol = Tools.colToLetter(at + count - 1);
          const r = s.getRange(`${startCol}:${endCol}`);
          r.insert('Right');
        }
        await ctx.sync();
      });
      if (Journal.current) Journal.recordStructuralChange(sheet, 'insert', { kind, at, count });
      return { ok: true, result: { sheet, kind, at, count } };
    },

    async delete_rows_cols({ sheet, kind, at, count = 1 }) {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        if (kind === 'rows') {
          const r = s.getRange(`${at}:${at + count - 1}`);
          r.delete('Up');
        } else {
          const startCol = Tools.colToLetter(at);
          const endCol = Tools.colToLetter(at + count - 1);
          const r = s.getRange(`${startCol}:${endCol}`);
          r.delete('Left');
        }
        await ctx.sync();
      });
      if (Journal.current) Journal.recordStructuralChange(sheet, 'delete', { kind, at, count });
      return { ok: true, result: { sheet, kind, at, count, note: 'Structural delete — undo is partial.' } };
    },

    async sort_range({ sheet, range, keyColumn, ascending = true, hasHeader = true }) {
      await Journal.recordRangePreImage(sheet, range);
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet);
        const r = s.getRange(range);
        const sortFields = [{
          key: keyColumn - 1, // 0-based within range
          sortOn: 'Value',
          ascending: ascending,
          color: null,
          icon: null
        }];
        r.sort.apply(sortFields, hasHeader, 'TopToBottom', null, null, null);
        await ctx.sync();
      });
      return { ok: true, result: { sheet, range, keyColumn, ascending } };
    }
  },

  /* ----------------------------------------------------------
     HELPERS
     ---------------------------------------------------------- */
  sanitizeColor(c) {
    if (!c || typeof c !== 'string') return null;
    if (/^#[0-9a-fA-F]{6}$/.test(c)) return c;
    if (/^[0-9a-fA-F]{6}$/.test(c)) return '#' + c;
    if (/^#[0-9a-fA-F]{3}$/.test(c)) {
      return '#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
    }
    if (/^[0-9a-fA-F]{3}$/.test(c)) {
      return '#' + c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
    }
    return null;
  },

  mapAggregation(agg) {
    const map = {
      sum: 'Sum', count: 'Count', average: 'Average',
      max: 'Max', min: 'Min'
    };
    return map[agg] || 'Sum';
  },

  applyBorders(ctx, range, borders) {
    if (typeof borders === 'string') {
      // Apply same style to all edges.
      const lineStyle = this.borderLineStyle(borders);
      range.format.borders.getItem('EdgeTop').style = lineStyle;
      range.format.borders.getItem('EdgeBottom').style = lineStyle;
      range.format.borders.getItem('EdgeLeft').style = lineStyle;
      range.format.borders.getItem('EdgeRight').style = lineStyle;
    } else if (borders && typeof borders === 'object') {
      for (const key of Object.keys(borders)) {
        try {
          range.format.borders.getItem(key).style = this.borderLineStyle(borders[key]);
        } catch (e) { /* ignore invalid edge key */ }
      }
    }
  },

  borderLineStyle(s) {
    const map = { Thin: 'Thin', Medium: 'Medium', Thick: 'Thick', Dashed: 'Dash' };
    return map[s] || 'Thin';
  },

  colToLetter(n) {
    let r = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      r = String.fromCharCode(65 + rem) + r;
      n = Math.floor((n - 1) / 26);
    }
    return r;
  },

  letterToNum(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  },

  /**
   * Given a start cell like "B3", return the cell offset by dRow rows and
   * dCol columns. Used to compute the end of a write range from values dims.
   */
  offsetRange(startCell, dRow, dCol) {
    const m = startCell.match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return startCell;
    const startColNum = this.letterToNum(m[1].toUpperCase());
    const startRow = parseInt(m[2], 10);
    const endColNum = startColNum + dCol;
    const endRow = startRow + dRow;
    return `${this.colToLetter(endColNum)}${endRow}`;
  },

  /**
   * Given a base range address like "Sheet!A1:D100" and 0-based row/col
   * offsets, return a single-cell address within that range.
   */
  addressFromIndex(rangeAddr, row, col) {
    const rangePart = rangeAddr.includes('!') ? rangeAddr.split('!')[1] : rangeAddr;
    const m = rangePart.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
    if (!m) return rangeAddr;
    const startColNum = this.letterToNum(m[1].toUpperCase());
    const startRow = parseInt(m[2], 10);
    return `${this.colToLetter(startColNum + col)}${startRow + row}`;
  }
};
