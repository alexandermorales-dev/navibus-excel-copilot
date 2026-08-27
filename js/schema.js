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

    // Detect non-tabular layouts: multi-block columns, merged header rows,
    // or data starting several rows down. This helps the model choose the
    // right approach (recipe vs raw ops vs "normalize first").
    const layout = this.detectLayout(values, isHeader, rowCount, colCount);
    const layoutType = layout.type;
    const headerRowIndex = layout.headerRow || 0;

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
    } else if (layoutType === 'titled' && headerRowIndex > 0) {
      // Titled sheet: headers are at headerRowIndex (1-based), data starts
      // below. Read the header row and treat it like a tabular sheet.
      const hdrIdx = headerRowIndex - 1;  // convert to 0-based array index
      const headerRow = values[hdrIdx] || [];
      headers = headerRow.map(v => String(v));
      // Infer types from data rows (below header)
      columnTypes = this.inferColumnTypes(headers.length, values, safeFormats, headerRowIndex);
      // Sample data rows below the header
      const sampleCount = Math.min(this.SAMPLE_ROWS, values.length - headerRowIndex);
      for (let i = headerRowIndex; i < headerRowIndex + sampleCount; i++) {
        if (values[i]) {
          sampleRows.push(values[i].map((v, c) => this.formatValue(v, columnTypes[c])));
        }
      }
      // Compute stats from data rows
      columnStats = this.computeColumnStats(headers, values, columnTypes, headerRowIndex);
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

    // For titled sheets, hasHeaders is false but we populated headers.
    // Set a flag so toText knows to show them.
    const effectiveHasHeaders = isHeader || (layoutType === 'titled' && headers.length > 0);

    const sampleRowCount = effectiveHasHeaders
      ? Math.min(this.SAMPLE_ROWS, rowCount - (headerRowIndex || 1))
      : Math.min(this.SAMPLE_ROWS, rowCount);

    return {
      name: sheet.name,
      empty: false,
      rows: rowCount,
      cols: colCount,
      hasHeaders: effectiveHasHeaders,
      layoutType: layoutType,
      headerRowIndex: headerRowIndex,
      headers: headers,
      columnTypes: columnTypes,
      columnStats: columnStats,
      sampleRows: sampleRows,
      address: usedRange.address,
      truncated: effectiveHasHeaders ? (rowCount - (headerRowIndex || 1)) > sampleRowCount : rowCount > sampleRowCount,
      statsPartial: isLarge,
      statsRowCount: effectiveHasHeaders ? statsRowCount : 0
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

  /**
   * Detect the overall layout of a sheet. This helps the model understand
   * what it's looking at and choose the right approach:
   *
   *   'tabular'    — clean flat table, headers in row 1, one record per row
   *   'titled'     — title/info rows at top, tabular data starts a few rows down
   *   'multiblock' — data laid out in side-by-side column blocks (e.g. one
   *                  block per ship/department), not a flat table
   *   'unknown'    — can't determine, treat with caution
   */
  detectLayout(values, isHeader, rowCount, colCount) {
    if (!values || values.length === 0) return { type: 'unknown', headerRow: 0 };

    // First, check for titled layout: first few rows are sparse, then
    // a header row appears. This takes priority over multiblock because
    // a titled sheet may have empty columns in the title rows that would
    // falsely trigger multiblock detection.
    if (!isHeader && rowCount > 5) {
      for (let r = 1; r < Math.min(8, values.length); r++) {
        if (this.looksLikeHeader(values[r])) {
          // Found a header row below row 1. Now check if the data below
          // it is clean tabular (no empty separator columns in data rows).
          const dataStart = r + 1;
          if (dataStart < values.length) {
            // Check column density in data rows — if most columns have
            // data, it's a titled tabular sheet, not multiblock.
            let dataCols = 0;
            for (let c = 0; c < colCount; c++) {
              for (let dr = dataStart; dr < Math.min(values.length, dataStart + 10); dr++) {
                if (values[dr] && values[dr][c] !== null && values[dr][c] !== '' && values[dr][c] !== undefined) {
                  dataCols++;
                  break;
                }
              }
            }
            // If >70% of columns have data, it's titled tabular.
            // headerRow is 1-based (r is 0-based index into values).
            if (dataCols / colCount > 0.7) return { type: 'titled', headerRow: r + 1 };
          }
          // Header found but data is sparse — could be multiblock with
          // a title above. Fall through to multiblock check.
          break;
        }
      }
    }

    // Check for multi-block: large column count relative to data density,
    // with empty column gaps (columns that are entirely empty in a way
    // that separates blocks). Only check data rows (skip title rows).
    if (colCount > 10) {
      // Find the first row with substantial data (3+ non-empty cells).
      let dataStartRow = 0;
      for (let r = 0; r < Math.min(values.length, 10); r++) {
        const nonEmpty = (values[r] || []).filter(v => v !== null && v !== '' && v !== undefined).length;
        if (nonEmpty >= 3) { dataStartRow = r; break; }
      }

      // Find columns that are entirely empty across data rows — these
      // act as separators between blocks.
      const emptyCols = [];
      for (let c = 0; c < colCount; c++) {
        let hasData = false;
        for (let r = dataStartRow; r < Math.min(values.length, dataStartRow + 20); r++) {
          if (values[r] && values[r][c] !== null && values[r][c] !== '' && values[r][c] !== undefined) {
            hasData = true;
            break;
          }
        }
        if (!hasData) emptyCols.push(c);
      }
      // If there are 2+ empty separator columns with data blocks between
      // them, it's a multi-block layout.
      if (emptyCols.length >= 2) {
        let blocks = 1;
        for (let i = 1; i < emptyCols.length; i++) {
          if (emptyCols[i] - emptyCols[i - 1] > 1) blocks++;
        }
        if (blocks >= 2) return { type: 'multiblock', headerRow: 0 };
      }
    }

    // Check for very wide sheets (100+ cols) that aren't real tables —
    // often formatting artifacts or pivot cache sheets.
    if (colCount > 100 && rowCount < 200) return { type: 'unknown', headerRow: 0 };

    if (isHeader) return { type: 'tabular', headerRow: 1 };
    return { type: 'unknown', headerRow: 0 };
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
   * Convert Excel column letter(s) to a 1-based column index.
   */
  letterToCol(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  },

  /**
   * Detail levels for snapshot rendering. The snapshot is injected into
   * every planning prompt, so its size directly drives token cost — and on
   * providers whose free tier binds on tokens-per-minute, size decides
   * whether a call fits at all. Callers pick a level to fit their budget.
   *
   *   compact — names, dimensions, headers. Enough to pick a sheet.
   *   normal  — + column types, exact column stats, a few sample rows.
   *   full    — + more sample rows and column letters.
   */
  DETAIL: {
    compact: { stats: false, types: false, sampleRows: 0 },
    normal:  { stats: true,  types: true,  sampleRows: 3 },
    full:    { stats: true,  types: true,  sampleRows: 10 }
  },

  /**
   * Format the snapshot as a compact text description for the LLM prompt.
   * @param {object} snap
   * @param {string} [level] — 'compact' | 'normal' | 'full'
   */
  toText(snap, level = 'normal') {
    const cfg = this.DETAIL[level] || this.DETAIL.normal;
    if (!snap || snap.sheetCount === 0) return 'The workbook is empty (no sheets).';

    const lines = [`Workbook has ${snap.sheetCount} sheet(s):`, ''];
    for (const s of snap.sheets) {
      if (s.error) {
        lines.push(`  - "${s.name}": COULD NOT READ THIS SHEET (${s.errorMessage}). Do not assume it is empty.`);
        continue;
      }
      if (s.empty) {
        lines.push(`  - "${s.name}": empty sheet`);
        continue;
      }

      const copilotTag = s.isCopilotSheet ? ' [previous copilot output — not source data]' : '';
      const layout = s.layoutType && s.layoutType !== 'tabular'
        ? ` [LAYOUT: ${s.layoutType}${s.layoutType === 'multiblock' ? ' — data is in side-by-side column blocks, NOT a flat table. Do NOT guess cell references. To use this sheet\'s data, request ranges via "needs" and the system will read the actual values for you. Do NOT use recipes on this sheet.' : s.layoutType === 'titled' ? ` — title rows at top, headers in row ${s.headerRowIndex}, data starts row ${s.headerRowIndex + 1}` : ' — non-standard layout, inspect sample rows carefully'}]`
        : '';
      lines.push(`  - "${s.name}": ${s.rows} rows x ${s.cols} cols, used range ${s.address}${copilotTag}${layout}`);

      if (!s.hasHeaders) {
        lines.push('    (no clear header row)');
        this._pushSamples(lines, s, cfg.sampleRows, false);
        continue;
      }

      // Pair each header with its column letter so the model can build
      // ranges without guessing which letter a column lives in.
      // Account for the used range's start column (e.g. B6:Q2879 starts at B).
      const bounds = this.parseAddress(s.address);
      const startColNum = bounds ? this.letterToCol(bounds.startCol) : 1;
      const cols = s.headers.map((h, i) => `${this.colToLetter(startColNum + i)}="${h}"`);
      lines.push(`    Columns: ${cols.join(', ')}`);

      if (cfg.types && s.columnTypes.length > 0) {
        lines.push(`    Types: ${s.columnTypes.join(', ')}`);
      }

      if (cfg.stats && Array.isArray(s.columnStats) && s.columnStats.some(st => st)) {
        const scope = s.statsPartial
          ? `first ${s.statsRowCount} of ${s.rows - 1} data rows, PARTIAL`
          : `all ${s.rows - 1} data rows, exact`;
        lines.push(`    Numeric stats (${scope}):`);
        for (let c = 0; c < s.headers.length; c++) {
          const st = s.columnStats[c];
          if (!st) continue;
          let line = `      "${s.headers[c]}": sum=${st.sum} avg=${st.avg} min=${st.min} max=${st.max} n=${st.count}`;
          if (st.nonNumericCells > 0) line += ` (${st.nonNumericCells} non-numeric skipped)`;
          lines.push(line);
        }
      }

      this._pushSamples(lines, s, cfg.sampleRows, true);
    }

    return lines.join('\n');
  },

  _pushSamples(lines, s, maxRows, hasHeaders) {
    if (!maxRows || !s.sampleRows || s.sampleRows.length === 0) return;
    const rows = s.sampleRows.slice(0, maxRows);
    const total = hasHeaders ? s.rows - 1 : s.rows;
    const note = total > rows.length ? ` (${rows.length} of ${total})` : '';
    lines.push(`    Sample rows${note}:`);
    for (const row of rows) lines.push(`      [${row.join(' | ')}]`);
  }
};
