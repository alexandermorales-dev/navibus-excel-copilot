/* ============================================
   actions.js — Action DSL: schemas + validation
   Each op is validated before execution.
   ============================================ */

const Actions = {
  /**
   * Validate a full action plan (array of action objects).
   * Returns: { valid: true } | { valid: false, errors: [...] }
   */
  validatePlan(plan) {
    const errors = [];

    if (!Array.isArray(plan)) {
      return { valid: false, errors: ['Action plan must be a JSON array.'] };
    }

    if (plan.length === 0) {
      return { valid: false, errors: ['Action plan is empty.'] };
    }

    for (let i = 0; i < plan.length; i++) {
      const action = plan[i];
      const prefix = `Action ${i + 1}`;

      if (!action || typeof action !== 'object') {
        errors.push(`${prefix}: must be an object.`);
        continue;
      }

      const op = action.op;
      if (!op) {
        errors.push(`${prefix}: missing "op" field.`);
        continue;
      }

      const validator = this.validators[op];
      if (!validator) {
        errors.push(`${prefix}: unknown op "${op}". Valid ops: ${Object.keys(this.validators).join(', ')}`);
        continue;
      }

      const result = validator(action, prefix);
      if (!result.valid) {
        errors.push(...result.errors);
      }
    }

    return errors.length === 0 ? { valid: true } : { valid: false, errors };
  },

  validators: {
    addSheet(action, prefix) {
      const errors = [];
      if (!action.name || typeof action.name !== 'string') {
        errors.push(`${prefix}: "name" (string) is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    writeRange(action, prefix) {
      const errors = [];
      if (!action.sheet || typeof action.sheet !== 'string') {
        errors.push(`${prefix}: "sheet" (string) is required.`);
      }
      if (!action.range || typeof action.range !== 'string') {
        errors.push(`${prefix}: "range" (string, e.g. "A1:B5") is required.`);
      }
      if (!Array.isArray(action.values)) {
        errors.push(`${prefix}: "values" (2D array) is required.`);
      } else {
        // Check it's a 2D array
        if (!action.values.every(row => Array.isArray(row))) {
          errors.push(`${prefix}: "values" must be a 2D array (array of arrays).`);
        }
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    formatRange(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" is required.`);
      if (!action.range) errors.push(`${prefix}: "range" is required.`);
      // At least one format property must be present
      const fmtProps = ['bold', 'italic', 'fontSize', 'fontName', 'fillColor',
                        'fontColor', 'horizontalAlignment', 'verticalAlignment',
                        'numberFormat', 'wrapText', 'columnWidth', 'rowHeight',
                        'borders', 'merge'];
      const hasAny = fmtProps.some(p => action[p] !== undefined);
      if (!hasAny) {
        errors.push(`${prefix}: at least one format property is required (${fmtProps.join(', ')}).`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    kpiBlock(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" is required.`);
      if (!action.cell || typeof action.cell !== 'string') {
        errors.push(`${prefix}: "cell" (string, e.g. "B2") is required.`);
      }
      if (!action.label || typeof action.label !== 'string') {
        errors.push(`${prefix}: "label" (string) is required.`);
      }
      // Accept either "formula" (starts with =) or "value" (a static number/string)
      if (!action.formula && action.value === undefined) {
        errors.push(`${prefix}: "formula" (string, starts with "=") or "value" is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    createTable(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" is required.`);
      if (!action.range) errors.push(`${prefix}: "range" is required.`);
      if (!action.name || typeof action.name !== 'string') {
        errors.push(`${prefix}: "name" (string) is required.`);
      }
      if (action.style && typeof action.style !== 'string') {
        errors.push(`${prefix}: "style" must be a string if provided.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    createPivot(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" (destination sheet) is required.`);
      if (!action.source || typeof action.source !== 'string') {
        errors.push(`${prefix}: "source" (range string, e.g. "Datos!A1:M500") is required.`);
      }
      if (!action.name || typeof action.name !== 'string') {
        errors.push(`${prefix}: "name" is required.`);
      }
      if (!action.dest || typeof action.dest !== 'string') {
        errors.push(`${prefix}: "dest" (destination cell, e.g. "A1") is required.`);
      }
      // At least one of rows/cols/values
      const hasRows = Array.isArray(action.rows) && action.rows.length > 0;
      const hasCols = Array.isArray(action.cols) && action.cols.length > 0;
      const hasValues = Array.isArray(action.values) && action.values.length > 0;
      if (!hasRows && !hasCols && !hasValues) {
        errors.push(`${prefix}: at least one of "rows", "cols", or "values" is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    createChart(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" is required.`);
      if (!action.type || typeof action.type !== 'string') {
        errors.push(`${prefix}: "type" (e.g. "columnClustered", "line", "pie") is required.`);
      }
      // Either sourceRange or sourcePivot
      if (!action.sourceRange && !action.sourcePivot) {
        errors.push(`${prefix}: either "sourceRange" or "sourcePivot" is required.`);
      }
      if (!action.dest || typeof action.dest !== 'string') {
        errors.push(`${prefix}: "dest" (destination cell, e.g. "E2") is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    addSlicer(action, prefix) {
      const errors = [];
      if (!action.sourcePivot && !action.sourceTable) {
        errors.push(`${prefix}: either "sourcePivot" or "sourceTable" is required.`);
      }
      if (!action.field || typeof action.field !== 'string') {
        errors.push(`${prefix}: "field" (column/slicer field name) is required.`);
      }
      if (!action.sheet) errors.push(`${prefix}: "sheet" (where to place slicer) is required.`);
      if (!action.dest) errors.push(`${prefix}: "dest" (cell position) is required.`);
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    conditionalFormat(action, prefix) {
      const errors = [];
      if (!action.sheet) errors.push(`${prefix}: "sheet" is required.`);
      if (!action.range) errors.push(`${prefix}: "range" is required.`);
      if (!action.type || typeof action.type !== 'string') {
        errors.push(`${prefix}: "type" (e.g. "colorScale", "dataBar", "cellValue") is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    },

    deleteSheet(action, prefix) {
      const errors = [];
      if (!action.name || typeof action.name !== 'string') {
        errors.push(`${prefix}: "name" (sheet to delete) is required.`);
      }
      return errors.length ? { valid: false, errors } : { valid: true };
    }
  },

  /**
   * List of all valid op names.
   */
  validOps() {
    return Object.keys(this.validators);
  }
};
