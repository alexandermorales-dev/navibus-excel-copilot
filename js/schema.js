/* ============================================
   schema.js — Workbook snapshot via Office.js
   Reads sheet names, used ranges, headers, column types, sample rows
   ============================================ */

const Schema = {
  // Number of sample data rows to include per sheet in the snapshot text.
  SAMPLE_ROWS: 20,

  /**
   * Build a compact snapshot of the workbook for the LLM.
   * Returns a string description of the workbook structure.
   */
  async snapshot() {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name, items/visibility');
      await ctx.sync();

      const allSheets = sheets.items || [];
      const visibleSheets = allSheets.filter(s =>
        s.visibility === 'Visible' || s.visibility === undefined
      );
      const sheetDescs = [];

      for (const sheet of visibleSheets) {
        try {
          const desc = await this.describeSheet(ctx, sheet);
          sheetDescs.push(desc);
        } catch (e) {
          sheetDescs.push({
            name: sheet.name || 'unknown',
            empty: true,
            rows: 0,
            cols: 0,
            error: e.message
          });
        }
      }

      return {
        sheetCount: visibleSheets.length,
        sheets: sheetDescs
      };
    });
  },

  async describeSheet(ctx, sheet) {
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load('address, rowCount, columnCount, values, numberFormats');
    await ctx.sync();

    if (usedRange.isNullObject) {
      return {
        name: sheet.name,
        empty: true,
        rows: 0,
        cols: 0
      };
    }

    const rowCount = usedRange.rowCount;
    const colCount = usedRange.columnCount;
    const values = usedRange.values;
    const formats = usedRange.numberFormats;

    // Defensive: values should be a 2D array. If not, treat as empty.
    if (!Array.isArray(values) || values.length === 0 || !Array.isArray(values[0])) {
      return {
        name: sheet.name,
        empty: true,
        rows: rowCount || 0,
        cols: colCount || 0
      };
    }

    const safeFormats = Array.isArray(formats) ? formats : [];

    // Detect if first row is a header row (all text, no duplicates)
    const firstRow = values[0] || [];
    const isHeader = this.looksLikeHeader(firstRow);

    let headers = [];
    let sampleRows = [];
    let columnTypes = [];

    let columnStats = [];

    if (isHeader) {
      headers = firstRow.map(v => String(v));
      // Sample data rows so the LLM can see deeper data
      const sampleCount = Math.min(this.SAMPLE_ROWS, rowCount - 1);
      for (let i = 1; i <= sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map(v => this.formatValue(v)));
        }
      }
      // Infer column types from formats + sample data
      columnTypes = this.inferColumnTypes(headers, values, safeFormats);
      // Compute full-range numeric stats (not limited to the sample) so the
      // LLM can answer aggregate questions (totals/averages) accurately.
      columnStats = this.computeColumnStats(headers, values, columnTypes);
    } else {
      // No clear header — show first N rows as data
      const sampleCount = Math.min(this.SAMPLE_ROWS, rowCount);
      for (let i = 0; i < sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map(v => this.formatValue(v)));
        }
      }
    }

    return {
      name: sheet.name,
      empty: false,
      rows: rowCount,
      cols: colCount,
      hasHeaders: isHeader,
      headers: headers,
      columnTypes: columnTypes,
      columnStats: columnStats,
      sampleRows: sampleRows,
      address: usedRange.address,
      truncated: isHeader ? (rowCount - 1) > sampleRows.length : rowCount > sampleRows.length
    };
  },

  /**
   * Compute min/max/sum/average/count for numeric-like columns using the
   * FULL used range values (already loaded in memory), not just the sample.
   * Returns an array parallel to headers; null entries for non-numeric cols.
   */
  computeColumnStats(headers, values, columnTypes) {
    const stats = [];
    for (let c = 0; c < headers.length; c++) {
      const type = columnTypes[c];
      if (type !== 'number' && type !== 'currency' && type !== 'percent') {
        stats.push(null);
        continue;
      }
      let sum = 0, count = 0, min = null, max = null;
      for (let r = 1; r < values.length; r++) {
        const v = values[r] ? values[r][c] : undefined;
        if (typeof v === 'number' && !isNaN(v)) {
          sum += v;
          count++;
          if (min === null || v < min) min = v;
          if (max === null || v > max) max = v;
        }
      }
      if (count === 0) {
        stats.push(null);
        continue;
      }
      stats.push({
        sum: this.roundStat(sum),
        avg: this.roundStat(sum / count),
        min: this.roundStat(min),
        max: this.roundStat(max),
        count
      });
    }
    return stats;
  },

  roundStat(n) {
    return Math.round(n * 100) / 100;
  },

  looksLikeHeader(row) {
    if (!row || row.length === 0) return false;
    // All non-empty, all strings (or convertible to string), no nulls
    const nonEmpty = row.filter(v => v !== null && v !== '' && v !== undefined);
    if (nonEmpty.length < 2) return false;
    // At least 80% should be text (not pure numbers)
    const textCount = nonEmpty.filter(v => typeof v === 'string' || (typeof v === 'number' && isNaN(Number(v)))).length;
    return textCount / nonEmpty.length >= 0.5;
  },

  inferColumnTypes(headers, values, formats) {
    const types = [];
    for (let c = 0; c < headers.length; c++) {
      const fmt = (Array.isArray(formats) && formats[0]) ? formats[0][c] : 'General';
      // Check a few data rows to infer type
      let sampleVals = [];
      for (let r = 1; r < Math.min(values.length, 6); r++) {
        if (values[r] && values[r][c] !== null && values[r][c] !== '') {
          sampleVals.push(values[r][c]);
        }
      }
      types.push(this.guessType(sampleVals, fmt));
    }
    return types;
  },

  guessType(samples, format) {
    if (samples.length === 0) return 'unknown';
    // Date format detection
    if (/d|m|y|h|s/i.test(format) && format !== 'General') return 'date';
    // Currency
    if (/\$|€|£|¥/.test(format)) return 'currency';
    // Percentage
    if (/%/.test(format)) return 'percent';
    // Check sample values
    const allNumbers = samples.every(v => typeof v === 'number');
    if (allNumbers) return 'number';
    const allStrings = samples.every(v => typeof v === 'string');
    if (allStrings) return 'text';
    return 'mixed';
  },

  formatValue(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return v;
    return String(v);
  },

  /**
   * Format the snapshot as a compact text description for the LLM prompt.
   */
  toText(snap) {
    if (snap.sheetCount === 0) return 'The workbook is empty (no sheets).';
    let lines = [`Workbook has ${snap.sheetCount} sheet(s):\n`];
    for (const s of snap.sheets) {
      if (s.empty) {
        lines.push(`  - "${s.name}": empty sheet`);
        continue;
      }
      lines.push(`  - "${s.name}": ${s.rows} rows × ${s.cols} cols (used range: ${s.address})`);
      if (s.hasHeaders) {
        lines.push(`    Headers: [${s.headers.map(h => `"${h}"`).join(', ')}]`);
        if (s.columnTypes.length > 0) {
          lines.push(`    Types: [${s.columnTypes.join(', ')}]`);
        }
        if (Array.isArray(s.columnStats) && s.columnStats.some(st => st)) {
          lines.push(`    Column stats (computed over ALL ${s.rows - 1} data rows, exact — safe to quote directly):`);
          for (let c = 0; c < s.headers.length; c++) {
            const st = s.columnStats[c];
            if (!st) continue;
            lines.push(`      "${s.headers[c]}": sum=${st.sum}, avg=${st.avg}, min=${st.min}, max=${st.max}, count=${st.count}`);
          }
        }
        if (s.sampleRows.length > 0) {
          lines.push(`    Sample rows${s.truncated ? ` (showing ${s.sampleRows.length} of ${s.rows - 1} data rows — truncated)` : ''}:`);
          for (const row of s.sampleRows) {
            lines.push(`      [${row.join(' | ')}]`);
          }
        }
      } else {
        lines.push(`    (no clear header row)`);
        if (s.sampleRows.length > 0) {
          lines.push(`    First rows:`);
          for (const row of s.sampleRows) {
            lines.push(`      [${row.join(' | ')}]`);
          }
        }
      }
      lines.push('');
    }
    return lines.join('\n');
  }
};
