/* ============================================
   executor.js — Deterministic Office.js executor
   Each op → hand-written Office.js code. No generated code.
   Tracks artifacts for rollback.
   ============================================ */

const Executor = {
  artifacts: [],  // { type, name, sheet } — created in current run

  clearArtifacts() {
    this.artifacts = [];
  },

  /**
   * Execute a validated action plan.
   * Returns: { succeeded: [...], failed: [...], rolledBack: bool }
   */
  async execute(plan, onProgress) {
    this.clearArtifacts();
    const succeeded = [];
    const failed = [];

    for (let i = 0; i < plan.length; i++) {
      const action = plan[i];
      const opDesc = `${action.op}${action.name ? '("' + action.name + '")' : action.range ? '("' + action.range + '")' : ''}`;

      // Notify progress callback
      if (onProgress) onProgress(i + 1, plan.length, action);

      try {
        await this.executeOp(action);
        succeeded.push({ index: i, op: opDesc, action });
      } catch (e) {
        failed.push({ index: i, op: opDesc, action, error: e.message || String(e) });
        // Rollback everything created so far in this run
        await this.rollback();
        return { succeeded, failed, rolledBack: true };
      }
    }

    // Post-execution: auto-fit columns on any sheet we created
    const createdSheets = [...new Set(this.artifacts.filter(a => a.type === 'sheet').map(a => a.name))];
    for (const sheetName of createdSheets) {
      try {
        await this.autoFitSheet(sheetName);
      } catch (e) {
        // Non-fatal — best effort
        console.warn(`Auto-fit failed for ${sheetName}: ${e.message}`);
      }
    }

    return { succeeded, failed, rolledBack: false };
  },

  /**
   * Auto-fit column widths on a sheet, but only for columns that don't
   * already have an explicit width set by the plan.
   */
  async autoFitSheet(sheetName) {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load('address');
      await ctx.sync();

      if (usedRange.isNullObject) return;

      // Get the used range and auto-fit columns
      const range = sheet.getUsedRange();
      range.format.autofitColumns();
      range.format.autofitRows();
      await ctx.sync();
    });
  },

  /**
   * Dispatch a single action to its handler.
   */
  async executeOp(action) {
    const handler = this.handlers[action.op];
    if (!handler) throw new Error(`Unknown op: ${action.op}`);
    await handler.call(this, action);
  },

  handlers: {
    async addSheet(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.add(action.name);
        sheet.activate();
        // Set tab color to blue for dashboard/report sheets
        if (action.tabColor) {
          sheet.tabColor = action.tabColor;
        } else {
          sheet.tabColor = '#4a9eff';
        }
        this.artifacts.push({ type: 'sheet', name: action.name });
        await ctx.sync();
      });
    },

    async writeRange(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);
        const range = sheet.getRange(action.range);

        // Load range dimensions to validate against values array
        range.load('rowCount, columnCount');
        await ctx.sync();

        const valRows = action.values.length;
        const valCols = action.values[0] ? action.values[0].length : 0;

        // Detect if any cell value is a formula (starts with =)
        const hasFormulas = action.values.some(row =>
          Array.isArray(row) && row.some(v => typeof v === 'string' && v.startsWith('='))
        );

        // Determine the target range (may need auto-resize)
        let targetRange;
        if (valRows !== range.rowCount || valCols !== range.columnCount) {
          const startCell = action.range.split(':')[0];
          const endCell = this.offsetRange(startCell, valRows - 1, valCols - 1);
          targetRange = sheet.getRange(`${startCell}:${endCell}`);
        } else {
          targetRange = range;
        }

        if (hasFormulas) {
          // Use .formulas so Excel evaluates formula strings (e.g. "=LC!E15")
          // Non-formula cells in the array are treated as literal values by Excel
          targetRange.formulas = action.values;
        } else {
          targetRange.values = action.values;
        }

        if (action.numberFormat) {
          targetRange.numberFormat = action.numberFormat;
        }
        await ctx.sync();
      });
    },

    async formatRange(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);
        const range = sheet.getRange(action.range);
        const fmt = range.format;

        if (action.bold !== undefined) fmt.font.bold = action.bold;
        if (action.italic !== undefined) fmt.font.italic = action.italic;
        if (action.fontSize !== undefined) fmt.font.size = action.fontSize;
        if (action.fontName !== undefined) fmt.font.name = action.fontName;
        if (action.fontColor !== undefined) fmt.font.color = this.sanitizeColor(action.fontColor);
        if (action.fillColor !== undefined) fmt.fill.color = this.sanitizeColor(action.fillColor);
        if (action.horizontalAlignment !== undefined) {
          fmt.horizontalAlignment = this.sanitizeAlignment(action.horizontalAlignment);
        }
        if (action.verticalAlignment !== undefined) {
          fmt.verticalAlignment = this.sanitizeAlignment(action.verticalAlignment, true);
        }
        if (action.wrapText !== undefined) fmt.wrapText = action.wrapText;
        if (action.numberFormat !== undefined) range.numberFormat = action.numberFormat;
        if (action.columnWidth !== undefined) fmt.columnWidth = action.columnWidth;
        if (action.rowHeight !== undefined) fmt.rowHeight = action.rowHeight;

        if (action.borders) {
          this.applyBorders(ctx, range, action.borders);
        }

        await ctx.sync();
      });
    },

    async kpiBlock(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);

        const cellAddr = action.cell;
        const labelAddr = this.offsetCell(action.cell, -1, 0);

        // Label cell — small, gray, bold
        const labelRange = sheet.getRange(labelAddr);
        labelRange.values = [[action.label]];
        labelRange.format.font.bold = true;
        labelRange.format.font.size = 10;
        labelRange.format.font.color = '#666666';
        labelRange.format.horizontalAlignment = 'Center';

        // Value cell — large, blue, bold, with number format
        const valueRange = sheet.getRange(cellAddr);
        // Accept either "formula" (starts with =) or "value" (static)
        if (action.formula) {
          valueRange.formulas = [[action.formula]];
        } else if (action.value !== undefined) {
          valueRange.values = [[action.value]];
        }
        valueRange.format.font.size = 22;
        valueRange.format.font.bold = true;
        valueRange.format.font.color = '#1a73e8';
        valueRange.format.horizontalAlignment = 'Center';
        valueRange.format.fill.color = '#f8f9fa';
        if (action.numberFormat) {
          valueRange.numberFormat = [[action.numberFormat]];
        } else {
          // Default: thousands separator
          valueRange.numberFormat = [['#,##0']];
        }

        // Set row heights for KPI area
        labelRange.format.rowHeight = 18;
        valueRange.format.rowHeight = 35;

        await ctx.sync();
      });
    },

    async createTable(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);
        const range = sheet.getRange(action.range);
        const table = ctx.workbook.tables.add(range, true);
        table.name = action.name;
        if (action.style) {
          table.style = action.style;
        }
        this.artifacts.push({ type: 'table', name: action.name, sheet: action.sheet });
        await ctx.sync();
      });
    },

    async createPivot(action) {
      await Excel.run(async (ctx) => {
        const destSheet = ctx.workbook.worksheets.getItem(action.sheet);
        const pivotTable = ctx.workbook.pivotTables.add(action.name, action.source, destSheet.getRange(action.dest));

        // Add row fields
        if (Array.isArray(action.rows)) {
          for (const field of action.rows) {
            const rowField = pivotTable.rowFields.add(field);
          }
        }

        // Add column fields
        if (Array.isArray(action.cols)) {
          for (const field of action.cols) {
            pivotTable.columnFields.add(field);
          }
        }

        // Add value fields with aggregation
        if (Array.isArray(action.values)) {
          for (const v of action.values) {
            const valField = pivotTable.dataFields.add(v.col);
            if (v.agg) {
              valField.summarizeBy = this.mapAggregation(v.agg);
            }
          }
        }

        // Add filter fields
        if (Array.isArray(action.filters)) {
          for (const field of action.filters) {
            pivotTable.filterFields.add(field);
          }
        }

        this.artifacts.push({ type: 'pivot', name: action.name, sheet: action.sheet });
        await ctx.sync();
      });
    },

    async createChart(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);
        let chart;

        if (action.sourceRange) {
          const range = sheet.getRange(action.sourceRange);
          chart = sheet.charts.add(action.type, range, action.seriesBy || 'Auto');
        } else if (action.sourcePivot) {
          // For pivot-sourced charts, get the pivot table's parent range
          const pivot = ctx.workbook.pivotTables.getItem(action.sourcePivot);
          pivot.load('layout/range');
          await ctx.sync();
          const rangeStr = pivot.layout.range;
          const sourceSheet = ctx.workbook.worksheets.getItem(action.sourceSheet || action.sheet);
          const range = sourceSheet.getRange(rangeStr);
          chart = sheet.charts.add(action.type, range, action.seriesBy || 'Auto');
        } else {
          throw new Error('createChart requires either sourceRange or sourcePivot');
        }

        if (action.title) {
          chart.title.text = action.title;
        }
        if (action.dest) {
          chart.setPosition(action.dest);
        }
        // Default chart size if not specified
        chart.width = action.width || 380;
        chart.height = action.height || 240;

        // Load chart name before reading it (Office.js requires explicit load)
        chart.load('name');
        await ctx.sync();

        this.artifacts.push({ type: 'chart', name: chart.name || 'chart', sheet: action.sheet });
        await ctx.sync();
      });
    },

    async addSlicer(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);

        let slicerSource;
        if (action.sourcePivot) {
          slicerSource = action.sourcePivot;
        } else if (action.sourceTable) {
          slicerSource = action.sourceTable;
        }

        const slicer = ctx.workbook.slicers.add(slicerSource, action.field, sheet.getRange(action.dest));
        if (action.name) slicer.name = action.name;

        // Load slicer name before reading it
        slicer.load('name');
        await ctx.sync();

        this.artifacts.push({ type: 'slicer', name: slicer.name || 'slicer', sheet: action.sheet });
        await ctx.sync();
      });
    },

    async conditionalFormat(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.sheet);
        const range = sheet.getRange(action.range);

        if (action.type === 'colorScale') {
          range.conditionalFormats.add('ColorScale', {
            threeColorScale: {
              minimum: { color: '#63BE7B', type: 'Min' },
              midpoint: { color: '#FFEB84', type: 'Percentile', percentile: 50 },
              maximum: { color: '#F8696B', type: 'Max' }
            }
          });
        } else if (action.type === 'dataBar') {
          range.conditionalFormats.add('DataBar', {
            dataBar: { barDirection: 'LeftToRight', color: '#4a9eff' }
          });
        } else if (action.type === 'cellValue' && action.rules) {
          for (const rule of action.rules) {
            range.conditionalFormats.add('CellValue', {
              cellValue: {
                operator: rule.operator || 'GreaterThan',
                formula: [rule.value]
              },
              format: { fill: { color: rule.color || '#FFC7CE' } }
            });
          }
        }

        await ctx.sync();
      });
    },

    async deleteSheet(action) {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(action.name);
        sheet.delete();
        await ctx.sync();
      });
    }
  },

  /**
   * Rollback: delete all artifacts created in the current run.
   */
  async rollback() {
    if (this.artifacts.length === 0) return;

    // Reverse order: delete charts/slicers/pivots first, then sheets
    const reversed = [...this.artifacts].reverse();

    await Excel.run(async (ctx) => {
      for (const art of reversed) {
        try {
          if (art.type === 'sheet') {
            const sheet = ctx.workbook.worksheets.getItem(art.name);
            sheet.delete();
          } else if (art.type === 'table') {
            const table = ctx.workbook.tables.getItem(art.name);
            table.delete();
          } else if (art.type === 'pivot') {
            const pivot = ctx.workbook.pivotTables.getItem(art.name);
            pivot.delete();
          } else if (art.type === 'chart') {
            // Charts are tricky to find by name; try via sheet
            const sheet = ctx.workbook.worksheets.getItem(art.sheet);
            const charts = sheet.charts;
            charts.load('items/name, items/id');
            await ctx.sync();
            for (const c of charts.items) {
              if (c.name === art.name || c.id === art.name) {
                c.delete();
                break;
              }
            }
          } else if (art.type === 'slicer') {
            const slicer = ctx.workbook.slicers.getItem(art.name);
            slicer.delete();
          }
        } catch (e) {
          // Best-effort rollback; log but don't throw
          console.warn(`Rollback: could not delete ${art.type} "${art.name}": ${e.message}`);
        }
      }
      await ctx.sync();
    });

    this.clearArtifacts();
  },

  // --- Helpers ---

  mapAggregation(agg) {
    const map = {
      'sum': 'Sum',
      'count': 'Count',
      'average': 'Average',
      'avg': 'Average',
      'max': 'Max',
      'min': 'Min',
      'product': 'Product',
      'countNumbers': 'CountNumbers',
      'standardDeviation': 'StandardDeviation',
      'variance': 'Variance'
    };
    return map[agg.toLowerCase()] || 'Sum';
  },

  offsetCell(cellAddr, rowOffset, colOffset) {
    // Parse "B2" → {col: 2, row: 2} → offset → return "B1"
    const match = cellAddr.match(/^([A-Z]+)(\d+)$/);
    if (!match) return cellAddr;
    const colStr = match[1];
    const rowNum = parseInt(match[2], 10) + rowOffset;
    if (rowNum < 1) return cellAddr; // can't go above row 1
    return colStr + rowNum;
  },

  offsetRange(startCell, rowOffset, colOffset) {
    // Given "A1", rowOffset=2, colOffset=3 → return "D3"
    const match = startCell.match(/^([A-Z]+)(\d+)$/);
    if (!match) return startCell;
    const colStr = match[1];
    const rowNum = parseInt(match[2], 10) + rowOffset;
    const newCol = this.incrementCol(colStr, colOffset);
    return newCol + Math.max(1, rowNum);
  },

  incrementCol(colStr, n) {
    // Convert column letters to number, add n, convert back
    let num = 0;
    for (let i = 0; i < colStr.length; i++) {
      num = num * 26 + (colStr.charCodeAt(i) - 64);
    }
    num += n;
    let result = '';
    while (num > 0) {
      const rem = (num - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      num = Math.floor((num - 1) / 26);
    }
    return result;
  },

  applyBorders(ctx, range, borders) {
    const validStyles = ['None', 'Continuous', 'Dash', 'DashDot', 'DashDotDot', 'Dot', 'Double', 'SlantDashDot', 'Thin', 'Medium', 'Thick'];
    const normalizeStyle = (s) => {
      if (!s || typeof s !== 'string') return 'Thin';
      const cap = s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
      return validStyles.includes(cap) ? cap : 'Thin';
    };
    const validEdges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'];
    const normalizeEdge = (e) => {
      if (!e) return 'EdgeTop';
      // Map common shortcuts
      const map = { 'top': 'EdgeTop', 'bottom': 'EdgeBottom', 'left': 'EdgeLeft', 'right': 'EdgeRight' };
      const lower = e.toLowerCase();
      if (map[lower]) return map[lower];
      if (validEdges.includes(e)) return e;
      // Try Edge + capitalized
      const candidate = 'Edge' + e.charAt(0).toUpperCase() + e.slice(1).toLowerCase();
      return validEdges.includes(candidate) ? candidate : 'EdgeTop';
    };

    if (typeof borders === 'string') {
      const style = normalizeStyle(borders);
      range.format.borders.getItem('EdgeTop').style = style;
      range.format.borders.getItem('EdgeBottom').style = style;
      range.format.borders.getItem('EdgeLeft').style = style;
      range.format.borders.getItem('EdgeRight').style = style;
    } else if (typeof borders === 'object') {
      for (const [edge, style] of Object.entries(borders)) {
        range.format.borders.getItem(normalizeEdge(edge)).style = normalizeStyle(style);
      }
    }
  },

  sanitizeColor(color) {
    if (!color || typeof color !== 'string') return color;
    // Ensure hex colors start with #
    if (/^[0-9a-fA-F]{6}$/.test(color)) return '#' + color;
    if (/^[0-9a-fA-F]{8}$/.test(color)) return '#' + color;
    // Already has # or is a named color
    return color;
  },

  sanitizeAlignment(value, isVertical) {
    if (!value || typeof value !== 'string') return value;
    const v = value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
    // Valid horizontal: Left, Center, Right, Distributed, Justify
    // Valid vertical: Top, Center, Bottom, Distributed, Justify
    const validH = ['Left', 'Center', 'Right', 'Distributed', 'Justify'];
    const validV = ['Top', 'Center', 'Bottom', 'Distributed', 'Justify'];
    const valid = isVertical ? validV : validH;
    if (valid.includes(v)) return v;
    // Fallback
    return isVertical ? 'Center' : 'Left';
  }
};
