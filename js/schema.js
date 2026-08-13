/* ============================================
   schema.js — Workbook snapshot via Office.js
   Reads sheet names, used ranges, headers, column types, sample rows
   ============================================ */

const Schema = {
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

    if (isHeader) {
      headers = firstRow.map(v => String(v));
      // Sample up to 3 data rows
      const sampleCount = Math.min(3, rowCount - 1);
      for (let i = 1; i <= sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map(v => this.formatValue(v)));
        }
      }
      // Infer column types from formats + sample data
      columnTypes = this.inferColumnTypes(headers, values, safeFormats);
    } else {
      // No clear header — show first 3 rows as data
      const sampleCount = Math.min(3, rowCount);
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
      sampleRows: sampleRows,
      address: usedRange.address
    };
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
      lines.push(`  - "${s.name}": ${s.rows} rows × ${s.cols} cols`);
      if (s.hasHeaders) {
        lines.push(`    Headers: [${s.headers.map(h => `"${h}"`).join(', ')}]`);
        if (s.columnTypes.length > 0) {
          lines.push(`    Types: [${s.columnTypes.join(', ')}]`);
        }
        if (s.sampleRows.length > 0) {
          lines.push(`    Sample rows:`);
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
