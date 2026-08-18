/* ============================================
   schema.js — Workbook snapshot via Office.js
   Reads sheet names, used ranges, headers, column types, sample rows
   ============================================ */

const Schema = {
  // Number of sample data rows to include per sheet in the snapshot text.
  SAMPLE_ROWS: 10,
  // Max cells to load in a single Office.js range call. Sheets exceeding this
  // are read via bounded sub-ranges to avoid payload limits and slow loads.
  CELL_CAP: 200000,
  // Max rows for stats computation when a sheet is too large to load fully.
  STAT_ROW_CAP: 5000,

  /**
   * Build a compact snapshot of the workbook for the LLM.
   * @param {string[]} [copilotSheets] — names of sheets created by the copilot
   *   in a prior run, so they can be tagged as output sheets in the snapshot.
   * Returns a string description of the workbook structure.
   */
  async snapshot(copilotSheets) {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name, items/visibility');
      await ctx.sync();

      const allSheets = sheets.items || [];
      const visibleSheets = allSheets.filter(s =>
        s.visibility === 'Visible' || s.visibility === undefined
      );
      const sheetDescs = [];
      const copilotSet = new Set(copilotSheets || []);

      for (const sheet of visibleSheets) {
        try {
          const desc = await this.describeSheet(ctx, sheet);
          desc.isCopilotSheet = copilotSet.has(sheet.name);
          sheetDescs.push(desc);
        } catch (e) {
          sheetDescs.push({
            name: sheet.name || 'unknown',
            error: true,
            errorMessage: e.message || String(e),
            rows: 0,
            cols: 0
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
    // Load metadata only first — avoids payload explosion on large sheets
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load('address, rowCount, columnCount');
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
    const totalCells = rowCount * colCount;
    const isLarge = totalCells > this.CELL_CAP;

    let values, formats;
    let statsRowCount;

    if (isLarge) {
      // Load a bounded sub-range to avoid Office.js payload limits.
      const parsed = this.parseAddress(usedRange.address);
      const startCol = parsed ? parsed.startCol : 'A';
      const endCol = parsed ? parsed.endCol : this.colToLetter(colCount);
      const statCap = Math.min(this.STAT_ROW_CAP, Math.floor(this.CELL_CAP / colCount));
      const boundedEndRow = Math.min(statCap, rowCount);
      const boundedRange = sheet.getRange(`${startCol}1:${endCol}${boundedEndRow}`);
      boundedRange.load('values, numberFormats');
      await ctx.sync();
      values = boundedRange.values;
      formats = boundedRange.numberFormats;
      statsRowCount = boundedEndRow - 1;
    } else {
      usedRange.load('values, numberFormats');
      await ctx.sync();
      values = usedRange.values;
      formats = usedRange.numberFormats;
      statsRowCount = rowCount - 1;
    }

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

    // Detect if first row is a header row
    const firstRow = values[0] || [];
    const isHeader = this.looksLikeHeader(firstRow);

    let headers = [];
    let sampleRows = [];
    let columnTypes = [];
    let columnStats = [];

    if (isHeader) {
      headers = firstRow.map(v => String(v));
      // Infer types BEFORE building sample rows so we can convert dates
      columnTypes = this.inferColumnTypes(headers.length, values, safeFormats, 1);
      // Sample data rows so the LLM can see deeper data
      const sampleCount = Math.min(this.SAMPLE_ROWS, values.length - 1);
      for (let i = 1; i <= sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map((v, c) => this.formatValue(v, columnTypes[c])));
        }
      }
      // Compute numeric stats over all loaded data rows
      columnStats = this.computeColumnStats(headers, values, columnTypes, 1);
    } else {
      // No clear header — infer types starting from row 0
      columnTypes = this.inferColumnTypes(colCount, values, safeFormats, 0);
      const sampleCount = Math.min(this.SAMPLE_ROWS, values.length);
      for (let i = 0; i < sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map((v, c) => this.formatValue(v, columnTypes[c])));
        }
      }
    }

    const sampleRowCount = isHeader
      ? Math.min(this.SAMPLE_ROWS, rowCount - 1)
      : Math.min(this.SAMPLE_ROWS, rowCount);

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
      truncated: isHeader ? (rowCount - 1) > sampleRowCount : rowCount > sampleRowCount,
      statsPartial: isLarge,
      statsRowCount: isHeader ? statsRowCount : 0
    };
  },

  /**
   * Compute min/max/sum/average/count for numeric-like columns using the
   * FULL used range values (already loaded in memory), not just the sample.
   * Returns an array parallel to headers; null entries for non-numeric cols.
   */
  computeColumnStats(headers, values, columnTypes, startRow = 1) {
    const stats = [];
    for (let c = 0; c < headers.length; c++) {
      const type = columnTypes[c];
      if (type !== 'number' && type !== 'currency' && type !== 'percent') {
        stats.push(null);
        continue;
      }
      let sum = 0, count = 0, min = null, max = null, nonNumeric = 0;
      for (let r = startRow; r < values.length; r++) {
        const v = values[r] ? values[r][c] : undefined;
        if (typeof v === 'number' && !isNaN(v)) {
          sum += v;
          count++;
          if (min === null || v < min) min = v;
          if (max === null || v > max) max = v;
        } else if (v !== null && v !== undefined && v !== '') {
          nonNumeric++;
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
        count,
        nonNumericCells: nonNumeric
      });
    }
    return stats;
  },

  /**
   * Round to 6 significant figures — preserves small values (percent,
   * unit prices) that 2-decimal rounding would destroy.
   */
  roundStat(n) {
    if (n === 0 || !isFinite(n)) return n;
    const abs = Math.abs(n);
    const magnitude = Math.floor(Math.log10(abs));
    const factor = Math.pow(10, 6 - 1 - magnitude);
    return Math.round(n * factor) / factor;
  },

  looksLikeHeader(row) {
    if (!row || row.length === 0) return false;
    const nonEmpty = row.filter(v => v !== null && v !== '' && v !== undefined);
    if (nonEmpty.length < 2) return false;
    // At least 80% of non-empty cells should be text (not pure numbers)
    const textCount = nonEmpty.filter(v => typeof v === 'string').length;
    return textCount / nonEmpty.length >= 0.8;
  },

  inferColumnTypes(colCount, values, formats, startRow = 1) {
    const types = [];
    for (let c = 0; c < colCount; c++) {
      const fmt = (Array.isArray(formats) && formats[startRow]) ? formats[startRow][c]
                 : (Array.isArray(formats) && formats[0]) ? formats[0][c]
                 : 'General';
      let sampleVals = [];
      for (let r = startRow; r < Math.min(values.length, startRow + 5); r++) {
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

  formatValue(v, colType) {
    if (v === null || v === undefined) return '';
    if (colType === 'date' && typeof v === 'number') {
      return this.excelSerialToDate(v);
    }
    if (typeof v === 'number') return v;
    return String(v);
  },

  /**
   * Convert an Excel serial date number to an ISO date string.
   * Excel's epoch is Dec 30 1899 (accounting for the 1900 leap-year bug).
   */
  excelSerialToDate(serial) {
    const epochMs = Date.UTC(1899, 11, 30);
    const date = new Date(epochMs + serial * 86400000);
    return date.toISOString().split('T')[0];
  },

  /**
   * Parse an Office.js range address like "Sheet1!A1:M500" or "A1:M500".
   * Returns { startCol, startRow, endCol, endRow } or null.
   */
  parseAddress(addr) {
    const rangePart = addr.includes('!') ? addr.split('!')[1] : addr;
    const match = rangePart.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!match) return null;
    return {
      startCol: match[1],
      startRow: parseInt(match[2], 10),
      endCol: match[3],
      endRow: parseInt(match[4], 10)
    };
  },

  /**
   * Convert a 1-based column index to Excel column letter(s).
   */
  colToLetter(n) {
    let result = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  },

  /**
   * Format the snapshot as a compact text description for the LLM prompt.
   */
  toText(snap) {
    if (snap.sheetCount === 0) return 'The workbook is empty (no sheets).';
    let lines = [`Workbook has ${snap.sheetCount} sheet(s):\n`];
    for (const s of snap.sheets) {
      if (s.error) {
        lines.push(`  - "${s.name}": ⚠ COULD NOT READ THIS SHEET (error: ${s.errorMessage}). Do NOT assume it is empty — tell the user you were unable to read it.`);
        lines.push('');
        continue;
      }
      if (s.empty) {
        lines.push(`  - "${s.name}": empty sheet`);
        lines.push('');
        continue;
      }
      const copilotTag = s.isCopilotSheet ? ' [copilot output sheet — not source data]' : '';
      lines.push(`  - "${s.name}": ${s.rows} rows × ${s.cols} cols (used range: ${s.address})${copilotTag}`);
      if (s.hasHeaders) {
        lines.push(`    Headers: [${s.headers.map(h => `"${h}"`).join(', ')}]`);
        if (s.columnTypes.length > 0) {
          lines.push(`    Types: [${s.columnTypes.join(', ')}]`);
        }
        if (Array.isArray(s.columnStats) && s.columnStats.some(st => st)) {
          if (s.statsPartial) {
            lines.push(`    Column stats (computed over first ${s.statsRowCount} of ${s.rows - 1} data rows — PARTIAL, may not reflect all data):`);
          } else {
            lines.push(`    Column stats (computed over ALL ${s.rows - 1} data rows, exact — safe to quote directly):`);
          }
          for (let c = 0; c < s.headers.length; c++) {
            const st = s.columnStats[c];
            if (!st) continue;
            let statLine = `      "${s.headers[c]}": sum=${st.sum}, avg=${st.avg}, min=${st.min}, max=${st.max}, count=${st.count}`;
            if (st.nonNumericCells > 0) {
              statLine += ` (${st.nonNumericCells} non-numeric cells excluded)`;
            }
            lines.push(statLine);
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
          lines.push(`    First rows${s.truncated ? ` (showing ${s.sampleRows.length} of ${s.rows} rows — truncated)` : ''}:`);
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
