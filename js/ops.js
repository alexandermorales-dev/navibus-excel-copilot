/* ============================================
   ops.js — Op schema, validation, execution, verification

   The model returns one JSON op batch instead of calling tools one at a
   time over many rounds. That batch then passes through three stages that
   all run locally over Office.js and therefore cost no API quota:

     validate — catch and repair bad ops before touching the workbook.
                Previously every one of these mistakes cost a full
                round-trip to discover.
     execute  — run ops in order through the existing Tools handlers, so
                all Office.js code and the undo journal work unchanged.
     verify   — read the result back and detect formula errors, empty
                KPIs and missing objects. This used to be done by the
                model with read_range calls, which was pure waste:
                reading cells is free.

   Op names match the Tools handler names so there is no hidden mapping,
   plus recipe.* composites expanded by recipes.js. Where a handler's
   argument names are awkward for a model to remember, the spec maps
   model-friendly aliases onto them here rather than in the prompt.
   ============================================ */

const Ops = {
  /**
   * Per-op contract:
   *   tool      — Tools handler to dispatch to
   *   required  — argument names that must be present
   *   sheetArg  — which field names a sheet name, for existence checks
   *   creates   — 'sheet' when the op brings a sheet into existence
   *   map       — normalize model output to handler arguments
   *   verify    — describes what to read back afterwards
   */
  SPEC: {
    add_sheet: {
      tool: 'add_sheet', required: ['name'], creates: 'sheet',
      map: (a) => ({ name: a.name, tabColor: a.tabColor })
    },
    write_range: {
      tool: 'write_range', required: ['sheet', 'range', 'values'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, range: a.range, values: a.values, numberFormat: a.numberFormat }),
      verify: 'values'
    },
    format_range: {
      tool: 'format_range', required: ['sheet', 'range'], sheetArg: 'sheet',
      map: (a) => {
        const out = { sheet: a.sheet, range: a.range };
        for (const k of ['bold', 'italic', 'fontSize', 'fontName', 'fontColor', 'fillColor',
                         'horizontalAlignment', 'verticalAlignment', 'wrapText',
                         'numberFormat', 'columnWidth', 'rowHeight', 'merge', 'borders']) {
          if (a[k] !== undefined) out[k] = a[k];
        }
        // Accept "align" as a friendlier alias.
        if (a.align && !out.horizontalAlignment) out.horizontalAlignment = a.align;
        return out;
      }
    },
    clear_range: {
      tool: 'clear_range', required: ['sheet', 'range'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, range: a.range })
    },
    delete_sheet: {
      tool: 'delete_sheet', required: ['name'],
      map: (a) => ({ name: a.name, userRequested: a.userRequested })
    },
    create_table: {
      tool: 'create_table', required: ['sheet', 'range', 'name'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, range: a.range, name: a.name, style: a.style || 'TableStyleMedium2' })
    },
    create_chart: {
      tool: 'create_chart', required: ['sheet', 'type'], sheetArg: 'sheet',
      // The handler wants sourceRange/dest; models reliably produce
      // dataRange/anchor, so both are accepted.
      map: (a) => ({
        sheet: a.sheet,
        type: Recipes.chartType(a.type),
        sourceRange: a.sourceRange || a.dataRange,
        sourcePivot: a.sourcePivot,
        sourceSheet: a.sourceSheet || a.dataSheet,
        title: a.title,
        dest: a.dest || a.anchor,
        width: a.width || Recipes.T.chartWidth,
        height: a.height || Recipes.T.chartHeight,
        xAxisTitle: a.xAxisTitle,
        yAxisTitle: a.yAxisTitle,
        seriesBy: a.seriesBy
      }),
      verify: 'chart'
    },
    create_pivot: {
      tool: 'create_pivot', required: ['name'],
      map: (a) => ({
        name: a.name,
        destSheet: a.destSheet || a.sheet,
        dest: a.dest || a.anchor || 'A1',
        source: a.source || (a.sourceSheet && a.sourceRange ? `${a.sourceSheet}!${a.sourceRange}` : undefined),
        rows: a.rows,
        cols: a.cols || a.columns,
        // Handler expects {col, agg}; models write {field, aggregation}.
        values: Array.isArray(a.values)
          ? a.values.map(v => (typeof v === 'string'
              ? { col: v, agg: 'sum' }
              : { col: v.col || v.field, agg: v.agg || v.aggregation || 'sum' }))
          : a.values,
        filters: a.filters
      })
    },
    add_slicer: {
      tool: 'add_slicer', required: ['sheet', 'field'], sheetArg: 'sheet',
      map: (a) => ({
        sheet: a.sheet, field: a.field, name: a.name,
        sourcePivot: a.sourcePivot, sourceTable: a.sourceTable,
        dest: a.dest || a.anchor || 'A1'
      })
    },
    conditional_format: {
      tool: 'conditional_format', required: ['sheet', 'range'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, range: a.range, type: a.type || a.kind || 'dataBar', rules: a.rules })
    },
    autofit: {
      tool: 'autofit', required: ['sheet'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, cols: a.cols, rows: a.rows })
    },
    sort_range: {
      tool: 'sort_range', required: ['sheet', 'range'], sheetArg: 'sheet',
      map: (a) => ({
        sheet: a.sheet, range: a.range,
        // Handler wants a 1-based keyColumn within the range.
        keyColumn: a.keyColumn || a.column || 1,
        ascending: a.ascending !== undefined ? a.ascending : (a.order !== 'desc'),
        hasHeader: a.hasHeader !== undefined ? a.hasHeader : true
      })
    },
    insert_rows_cols: {
      tool: 'insert_rows_cols', required: ['sheet', 'kind', 'at'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, kind: a.kind, at: Ops._toIndex(a.at), count: a.count || 1 })
    },
    delete_rows_cols: {
      tool: 'delete_rows_cols', required: ['sheet', 'kind', 'at'], sheetArg: 'sheet',
      map: (a) => ({ sheet: a.sheet, kind: a.kind, at: Ops._toIndex(a.at), count: a.count || 1 })
    },
    read_range: {
      tool: 'read_range', required: ['sheet', 'range'], sheetArg: 'sheet', readOnly: true,
      map: (a) => ({ sheet: a.sheet, range: a.range, what: a.what || 'values' })
    },
    // Alias: models frequently invent "merge_range" — map it to format_range
    // with merge:true rather than dropping it as an unknown op.
    merge_range: {
      tool: 'format_range', required: ['sheet', 'range'], sheetArg: 'sheet',
      map: (a) => {
        const out = { sheet: a.sheet, range: a.range, merge: true };
        for (const k of ['bold', 'fontSize', 'fontColor', 'fillColor',
                         'horizontalAlignment', 'verticalAlignment', 'rowHeight']) {
          if (a[k] !== undefined) out[k] = a[k];
        }
        return out;
      }
    }
  },

  RECIPES: new Set(['recipe.dashboard', 'recipe.summary_table']),

  MUTATING: new Set([
    'add_sheet', 'write_range', 'format_range', 'clear_range', 'delete_sheet',
    'create_table', 'create_chart', 'create_pivot', 'add_slicer',
    'conditional_format', 'autofit', 'sort_range', 'insert_rows_cols', 'delete_rows_cols'
  ]),

  RANGE_RE: /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/,
  // Whole-column / whole-row forms like "A:A" or "1:3".
  BAND_RE: /^(\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\d{1,7}:\d{1,7})$/,
  ERROR_VALUES: ['#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NULL!', '#NUM!', '#SPILL!', '#CALC!'],

  _toIndex(at) {
    // Handler indexes columns numerically; accept a letter for convenience.
    if (typeof at === 'number') return at;
    const s = String(at || '').trim();
    if (/^\d+$/.test(s)) return parseInt(s, 10);
    if (/^[A-Za-z]{1,3}$/.test(s)) return Tools.letterToNum(s.toUpperCase());
    return 1;
  },

  /* ----------------------------------------------------------
     VALIDATION — runs before anything touches the workbook
     ---------------------------------------------------------- */

  /**
   * Validate and normalize a plan.
   *
   * Anything fixable is fixed silently (sheet-name case and accents, range
   * casing, missing chart source, colliding sheet names). Anything not
   * fixable is dropped with a recorded reason, so one bad op cannot take
   * down the whole batch.
   *
   * @param {object} plan  — { intent, answer, ops, needs }
   * @param {object} snap  — workbook snapshot
   * Returns { ok, ops, dropped, warnings, plannedSheets, error }
   */
  validate(plan, snap) {
    if (!plan || typeof plan !== 'object') {
      return { ok: false, error: 'Plan is not an object', ops: [], dropped: [], warnings: [] };
    }

    const rawOps = Array.isArray(plan.ops) ? plan.ops : [];
    const ops = [];
    const dropped = [];
    const warnings = [];

    // Sheets that will exist by the time later ops run: existing sheets
    // plus any this batch creates. Without this, writes to a sheet created
    // earlier in the same batch would be wrongly rejected.
    const existing = new Set((snap && snap.sheets ? snap.sheets : []).map(s => String(s.name).toLowerCase()));
    // lowercase name -> exact name, so later ops can be resolved to the
    // spelling the sheet will actually have.
    const planned = new Map();
    const createdNames = [];

    for (let i = 0; i < rawOps.length; i++) {
      const raw = rawOps[i];
      if (!raw || typeof raw !== 'object') {
        dropped.push({ index: i, op: String(raw), reason: 'Not an object' });
        continue;
      }

      const name = raw.op || raw.name_of_op || raw.action;
      if (!name) {
        dropped.push({ index: i, op: '(missing)', reason: 'Missing "op" field' });
        continue;
      }

      if (this.RECIPES.has(name)) {
        const check = this.validateRecipe(raw, snap);
        if (!check.ok) {
          dropped.push({ index: i, op: name, reason: check.error });
          continue;
        }
        const target = check.op.sheet;
        const unique = this.uniqueSheetName(target, existing, planned);
        if (unique !== target) {
          warnings.push(`Sheet "${target}" already exists — using "${unique}"`);
          check.op.sheet = unique;
        }
        planned.set(unique.toLowerCase(), unique);
        createdNames.push(unique);
        ops.push(check.op);
        continue;
      }

      const spec = this.SPEC[name];
      if (!spec) {
        dropped.push({ index: i, op: name, reason: `Unknown op "${name}"` });
        continue;
      }

      const args = { ...raw };
      delete args.op;

      // Required fields
      const missing = spec.required.filter(k => args[k] === undefined || args[k] === null || args[k] === '');
      if (missing.length > 0) {
        dropped.push({ index: i, op: name, reason: `Missing required field(s): ${missing.join(', ')}` });
        continue;
      }

      // Sheet creation: rename on collision rather than failing.
      if (spec.creates === 'sheet') {
        const unique = this.uniqueSheetName(args.name, existing, planned);
        if (unique !== args.name) {
          warnings.push(`Sheet "${args.name}" already exists — using "${unique}"`);
          // Later ops in this batch referenced the original name, so
          // rewrite them too.
          this._renameInRemaining(rawOps, i + 1, args.name, unique);
          args.name = unique;
        }
        planned.set(String(args.name).toLowerCase(), String(args.name));
        createdNames.push(args.name);
      }

      // Sheet existence. Resolve to the workbook's exact spelling: Office.js
      // is inconsistent about case, and formulas built from a mis-cased name
      // are harder to debug later.
      if (spec.sheetArg && args[spec.sheetArg]) {
        const given = String(args[spec.sheetArg]);
        const resolved = Context.resolveSheetName(snap, given) || this._matchPlanned(planned, given);
        if (!resolved) {
          dropped.push({ index: i, op: name, reason: `Sheet "${given}" does not exist` });
          continue;
        }
        if (resolved !== given) {
          warnings.push(`Sheet "${given}" corrected to "${resolved}"`);
          args[spec.sheetArg] = resolved;
        }
      }

      // Range syntax
      const rangeErr = this.checkRanges(name, args);
      if (rangeErr) {
        dropped.push({ index: i, op: name, reason: rangeErr });
        continue;
      }

      // Op-specific checks
      const specific = this.checkSpecific(name, args, snap);
      if (specific.error) {
        dropped.push({ index: i, op: name, reason: specific.error });
        continue;
      }
      if (specific.warning) warnings.push(specific.warning);

      ops.push({ op: name, ...args });
    }

    return { ok: true, ops, dropped, warnings, createdSheets: createdNames };
  },

  validateRecipe(raw, snap) {
    const op = { ...raw };
    if (!op.source || !op.source.sheet) {
      return { ok: false, error: `${op.op} requires source.sheet` };
    }
    const resolved = Context.resolveSheetName(snap, op.source.sheet);
    if (!resolved) {
      return { ok: false, error: `Source sheet "${op.source.sheet}" does not exist` };
    }
    op.source = { ...op.source, sheet: resolved };
    if (!op.sheet) op.sheet = op.op === 'recipe.dashboard' ? 'Dashboard' : 'Summary';

    if (op.op === 'recipe.summary_table' && (!op.groupBy || !op.valueColumn)) {
      return { ok: false, error: 'recipe.summary_table requires groupBy and valueColumn' };
    }
    if (op.op === 'recipe.dashboard') {
      const hasKpis = Array.isArray(op.kpis) && op.kpis.length > 0;
      const hasBreakdown = !!(op.groupBy && op.valueColumn);
      if (!hasKpis && !hasBreakdown) {
        return { ok: false, error: 'recipe.dashboard requires kpis or groupBy+valueColumn' };
      }
    }
    return { ok: true, op };
  },

  /**
   * Range syntax check. A malformed range would otherwise surface as an
   * opaque Office.js exception one API round later.
   */
  checkRanges(name, args) {
    const fields = ['range', 'sourceRange', 'dataRange', 'dest', 'anchor'];
    for (const f of fields) {
      const v = args[f];
      if (v === undefined || v === null || v === '') continue;
      const val = String(v).trim();
      // Strip a sheet qualifier: "Data!A1:D5" is acceptable input.
      const bare = val.includes('!') ? val.split('!').pop() : val;
      if (!this.RANGE_RE.test(bare) && !this.BAND_RE.test(bare)) {
        return `Invalid range in "${f}": "${val}"`;
      }
      args[f] = val.toUpperCase() === val ? val : val.toUpperCase();
    }
    return null;
  },

  checkSpecific(name, args, snap) {
    switch (name) {
      case 'write_range': {
        if (!Array.isArray(args.values) || args.values.length === 0) {
          return { error: 'values must be a non-empty array' };
        }
        // Accept a flat array as a single row.
        if (!Array.isArray(args.values[0])) args.values = [args.values];
        if (!args.values.every(r => Array.isArray(r))) {
          return { error: 'values must be a 2D array (array of row arrays)' };
        }
        const bad = this.checkFormulas(args.values);
        if (bad) return { error: bad };
        return {};
      }
      case 'format_range': {
        for (const f of ['fontColor', 'fillColor']) {
          if (args[f] && !Tools.sanitizeColor(args[f])) {
            // Formatting is cosmetic — drop the bad color, keep the op.
            const bad = args[f];
            delete args[f];
            return { warning: `Ignored invalid color "${bad}" in format_range` };
          }
        }
        return {};
      }
      case 'create_chart': {
        if (!args.sourceRange && !args.dataRange && !args.sourcePivot) {
          return { error: 'create_chart requires sourceRange (or sourcePivot)' };
        }
        return {};
      }
      case 'create_table': {
        if (!this.RANGE_RE.test(String(args.range).replace(/\$/g, ''))) {
          return { error: 'create_table needs an explicit range like "A1:C10"' };
        }
        if (/\s/.test(String(args.name))) {
          args.name = String(args.name).replace(/\s+/g, '_');
          return { warning: `Table name cannot contain spaces — using "${args.name}"` };
        }
        return {};
      }
      case 'delete_sheet': {
        if (!args.userRequested) {
          return { error: `Refusing to delete "${args.name}": not explicitly requested by the user` };
        }
        return {};
      }
      default:
        return {};
    }
  },

  /**
   * Cheap structural checks on formulas. Catches the mistakes that produce
   * an immediate Excel error without needing to evaluate anything.
   */
  checkFormulas(values) {
    for (const row of values) {
      for (const v of row) {
        if (typeof v !== 'string' || !v.startsWith('=')) continue;
        let depth = 0;
        let inStr = false;
        for (const ch of v) {
          if (ch === '"') inStr = !inStr;
          else if (!inStr && ch === '(') depth++;
          else if (!inStr && ch === ')') depth--;
          if (depth < 0) return `Formula has an unmatched ")": ${v.slice(0, 60)}`;
        }
        // Check the string state first: an unterminated quote also swallows
        // closing parens, so reporting "unbalanced parentheses" would point
        // at the symptom rather than the cause.
        if (inStr) return `Formula has an unterminated string: ${v.slice(0, 60)}`;
        if (depth !== 0) return `Formula has unbalanced parentheses: ${v.slice(0, 60)}`;
      }
    }
    return null;
  },

  /**
   * Resolve a sheet name against sheets this batch is about to create,
   * which aren't in the snapshot yet.
   */
  _matchPlanned(planned, given) {
    return planned.get(String(given).toLowerCase()) || null;
  },

  uniqueSheetName(name, existing, planned) {
    const base = String(name || 'Sheet').slice(0, 28);
    const taken = (n) => existing.has(n.toLowerCase()) || planned.has(n.toLowerCase());
    if (!taken(base)) return base;
    for (let i = 2; i < 100; i++) {
      const candidate = `${base} ${i}`;
      if (!taken(candidate)) return candidate;
    }
    return `${base} ${Date.now().toString().slice(-4)}`;
  },

  _renameInRemaining(rawOps, from, oldName, newName) {
    for (let j = from; j < rawOps.length; j++) {
      const o = rawOps[j];
      if (!o || typeof o !== 'object') continue;
      for (const f of ['sheet', 'destSheet', 'sourceSheet', 'name']) {
        if (o[f] === oldName) o[f] = newName;
      }
    }
  },

  /* ----------------------------------------------------------
     EXECUTION
     ---------------------------------------------------------- */

  /**
   * Run ops in order via the existing Tools handlers.
   *
   * Dependency handling: if a sheet fails to be created, every later op
   * targeting that sheet is skipped rather than producing a cascade of
   * identical errors (which used to fill the activity feed and confuse
   * the model on the next round).
   *
   * @param {Array} ops
   * @param {object} opts — { snap, signal, onOpStart, onOpEnd, onRecipe }
   * Returns { results, failed, executed, createdSheets, writes }
   */
  async execute(ops, opts = {}) {
    const { snap, signal, onOpStart, onOpEnd } = opts;
    const results = [];
    const failed = [];
    const deadSheets = new Set();
    const createdSheets = [];
    const writes = [];
    let executed = 0;

    const readRange = async (sheet, range) => {
      const r = await Tools.dispatch('read_range', { sheet, range });
      return r.ok ? { ok: true, ...r.result } : { ok: false, error: r.error };
    };

    // Expand recipes first so the checklist shown to the user and the
    // dependency analysis both see the real op list.
    const flat = [];
    for (const op of ops) {
      if (!this.RECIPES.has(op.op)) {
        flat.push(op);
        continue;
      }
      const expansion = await Recipes.expand(op, { snap, readRange });
      if (!expansion.ok) {
        failed.push({ op: op.op, error: expansion.error });
        results.push({ op: op.op, ok: false, error: expansion.error, recipe: true });
        continue;
      }
      for (const w of (expansion.warnings || [])) {
        results.push({ op: op.op, ok: true, warning: w, recipe: true, skipped: true });
      }
      flat.push(...expansion.ops.map(o => ({ ...o, _from: op.op })));
    }

    for (let i = 0; i < flat.length; i++) {
      if (signal && signal.aborted) break;
      const op = flat[i];
      const spec = this.SPEC[op.op];
      if (!spec) {
        failed.push({ op: op.op, error: `Unknown op "${op.op}"` });
        continue;
      }

      const targetSheet = spec.sheetArg ? op[spec.sheetArg] : (spec.creates === 'sheet' ? op.name : null);
      if (targetSheet && deadSheets.has(String(targetSheet).toLowerCase())) {
        results.push({ op: op.op, ok: false, skipped: true, sheet: targetSheet,
                       error: `Skipped: sheet "${targetSheet}" could not be created` });
        continue;
      }

      const args = spec.map ? spec.map(op) : op;
      const callId = `op${i}`;
      if (onOpStart) onOpStart(callId, op.op, op);

      const res = await Tools.dispatch(spec.tool, args);

      if (onOpEnd) onOpEnd(callId, op.op, res);

      if (res.ok) {
        executed++;
        results.push({ op: op.op, ok: true, args, result: res.result });
        if (spec.creates === 'sheet') {
          createdSheets.push(op.name);
          Context.noteCopilotSheet(op.name);
        }
        if (spec.verify === 'values' && res.result && res.result.address) {
          writes.push({ sheet: args.sheet, range: res.result.address, values: args.values });
        }
      } else {
        failed.push({ op: op.op, args, error: res.error });
        results.push({ op: op.op, ok: false, args, error: res.error });
        if (spec.creates === 'sheet') {
          deadSheets.add(String(op.name).toLowerCase());
        }
      }
    }

    if (createdSheets.length > 0) {
      Tools.invalidateOverviewCache();
      Context.invalidate();
    }

    return { results, failed, executed, createdSheets, writes };
  },

  /* ----------------------------------------------------------
     VERIFICATION — free, because reading cells is an Office.js call
     ---------------------------------------------------------- */

  /**
   * Read back everything that was written and look for problems.
   *
   * @param {object} exec — the result of execute()
   * Returns { ok, problems, values }
   *   problems: [{ kind, sheet, address, detail }]
   *   values:   [{ sheet, address, label, value }] verified cell readings
   */
  async verify(exec, opts = {}) {
    const { signal } = opts;
    const problems = [];
    const values = [];

    for (const w of exec.writes) {
      if (signal && signal.aborted) break;
      const res = await Tools.dispatch('read_range', { sheet: w.sheet, range: w.range });
      if (!res.ok) {
        problems.push({ kind: 'unreadable', sheet: w.sheet, address: w.range, detail: res.error });
        continue;
      }
      const data = res.result.data || [];

      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < (data[r] || []).length; c++) {
          const cell = data[r][c];
          if (typeof cell !== 'string') continue;
          const hit = this.ERROR_VALUES.find(e => cell === e || cell.startsWith(e));
          if (hit) {
            const addr = Tools.addressFromIndex(w.range, r, c);
            const src = this._sourceFormula(w.values, r, c);
            problems.push({
              kind: 'formula_error',
              sheet: w.sheet,
              address: addr,
              detail: `${hit} in ${w.sheet}!${addr}${src ? ` (formula: ${src})` : ''}`
            });
          }
        }
      }

      // Collect single-cell formula results as verified values: these are
      // the KPI cells, and they let the final answer cite real numbers
      // without spending another API call.
      if (Array.isArray(w.values) && w.values.length === 1 && w.values[0].length === 1) {
        const src = w.values[0][0];
        if (typeof src === 'string' && src.startsWith('=')) {
          const value = data[0] && data[0][0];
          values.push({ sheet: w.sheet, address: w.range, formula: src, value });
          if (value === 0 || value === null || value === '') {
            problems.push({
              kind: 'suspicious_zero',
              sheet: w.sheet,
              address: w.range,
              detail: `${w.sheet}!${w.range} evaluates to ${value === '' ? 'blank' : value} (formula: ${src})`
            });
          }
        }
      }
    }

    return { ok: problems.length === 0, problems, values };
  },

  _sourceFormula(values, r, c) {
    if (!Array.isArray(values) || !values[r]) return null;
    const v = values[r][c];
    return typeof v === 'string' && v.startsWith('=') ? v.slice(0, 80) : null;
  },

  /**
   * Pair verified KPI readings with their labels so the answer can quote
   * them. Labels come from a write op that placed text immediately left of
   * the value cell, which is how the recipes lay out KPI cards.
   */
  labelValues(exec, verified) {
    const out = [];
    for (const v of verified) {
      const label = this._nearbyLabel(exec, v);
      out.push({ label, value: v.value, address: `${v.sheet}!${v.address}` });
    }
    return out;
  },

  _nearbyLabel(exec, v) {
    const cell = String(v.address).split(':')[0];
    const m = cell.match(/^\$?([A-Z]+)\$?(\d+)$/);
    if (!m) return null;
    const col = m[1];
    const row = parseInt(m[2], 10);
    // The recipes write the KPI label one row above the value.
    const labelCell = `${col}${row - 1}`;
    for (const r of exec.results) {
      if (!r.ok || r.op !== 'write_range' || !r.args) continue;
      if (r.args.sheet !== v.sheet) continue;
      const anchor = String(r.args.range).split(':')[0].replace(/\$/g, '');
      if (anchor !== labelCell) continue;
      const val = r.args.values && r.args.values[0] && r.args.values[0][0];
      if (typeof val === 'string' && !val.startsWith('=')) return val;
    }
    return null;
  },

  /**
   * Compact, token-cheap description of what went wrong, for the repair
   * call. Only failures are included — successful ops are already applied
   * and re-sending them would invite duplicate work.
   *
   * For formula errors (#VALUE!, #REF!, etc.), the actual values of the
   * referenced cells are read back so the model can understand WHY the
   * formula failed and fix it — e.g. seeing that DC!D20 contains text
   * explains the #VALUE! error.
   */
  async describeProblems({ dropped, failed, problems }) {
    const lines = [];
    for (const d of (dropped || [])) {
      lines.push(`REJECTED ${d.op}: ${d.reason}`);
    }
    for (const f of (failed || [])) {
      lines.push(`FAILED ${f.op}${f.args && f.args.range ? ` at ${f.args.sheet}!${f.args.range}` : ''}: ${f.error}`);
    }
    for (const p of (problems || [])) {
      let line = `${p.kind.toUpperCase()}: ${p.detail}`;
      // For formula errors, read the referenced cells so the repair model
      // can see what's actually there and derive a correct formula.
      if (p.kind === 'formula_error' && p.sheet && p.address) {
        const cellRefs = this._extractCellRefs(p.detail);
        if (cellRefs.length > 0) {
          const values = await this._readCellRefs(cellRefs);
          if (values.length > 0) {
            line += '\n  Referenced cell values:';
            for (const v of values.slice(0, 8)) {
              line += `\n    ${v.ref} = ${v.value}`;
            }
          }
        }
      }
      lines.push(line);
    }
    return lines.join('\n');
  },

  /**
   * Extract cell references (Sheet!A1 or A1) from a formula string.
   */
  _extractCellRefs(text) {
    const refs = [];
    // Match patterns like Sheet!A1, 'Sheet'!A1, Sheet!$A$1
    const re = /(?:'([^']+)'!|([A-Za-z_][A-Za-z0-9_]*)!)?\$?([A-Z]{1,3})\$?(\d{1,7})/g;
    let m;
    while ((m = re.exec(text)) !== null) {
      const sheet = m[1] || m[2] || null;
      const col = m[3];
      const row = m[4];
      refs.push({ sheet, cell: `${col}${row}` });
      if (refs.length >= 10) break;   // cap to keep the repair prompt small
    }
    return refs;
  },

  /**
   * Read the actual values of referenced cells via Office.js.
   */
  async _readCellRefs(refs) {
    // Group by sheet to minimize Office.js calls.
    const bySheet = new Map();
    for (const r of refs) {
      const key = r.sheet || '';
      if (!bySheet.has(key)) bySheet.set(key, []);
      bySheet.get(key).push(r);
    }

    const results = [];
    for (const [sheet, cells] of bySheet) {
      if (!sheet) continue;   // can't read without a sheet name
      // Read each cell individually — they're typically scattered.
      for (const c of cells) {
        try {
          const res = await Tools.dispatch('read_range', { sheet, range: c.cell });
          if (res.ok && res.result && res.result.data) {
            const val = res.result.data[0] && res.result.data[0][0];
            const display = val === null || val === undefined || val === ''
              ? '(empty)'
              : typeof val === 'string' ? `"${val}"` : String(val);
            results.push({ ref: `${sheet}!${c.cell}`, value: display });
          }
        } catch (e) { /* skip unreadable cells */ }
      }
    }
    return results;
  }
};
