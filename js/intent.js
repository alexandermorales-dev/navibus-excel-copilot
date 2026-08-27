/* ============================================
   intent.js — Local intent classification and fast paths

   Two jobs, both done without an API call:

   1. classify() picks which prompt module to send, so a question about
      the data never pays for the write-op reference (~350 tokens instead
      of ~1,200).

   2. fastPath() answers outright when the answer is already known.
      Schema computes exact sum/avg/min/max/count for every numeric
      column over every row, so "what is the total of Amount?" needs no
      model at all. These answers state exactly what was computed, so a
      misread question is visible to the user rather than hidden behind
      confident phrasing.

   The fast path is deliberately conservative: it fires only on an
   explicit aggregate keyword plus a confident column match. Anything
   ambiguous falls through to the planner.
   ============================================ */

const Intent = {
  BUILD_WORDS: [
    'create', 'build', 'make', 'generate', 'add a', 'set up', 'setup', 'draw',
    'dashboard', 'report', 'chart', 'graph', 'pivot', 'table', 'summary',
    'crea', 'crear', 'construye', 'construir', 'genera', 'generar', 'haz',
    'hazme', 'arma', 'armar', 'panel', 'informe', 'reporte', 'gráfico',
    'grafico', 'tabla', 'dinámica', 'dinamica', 'resumen', 'dibuja'
  ],

  EDIT_WORDS: [
    'sort', 'format', 'clean', 'rename', 'delete', 'remove', 'clear',
    'highlight', 'autofit', 'resize', 'insert', 'fill', 'replace', 'update',
    'ordena', 'ordenar', 'formatea', 'formatear', 'limpia', 'limpiar',
    'renombra', 'renombrar', 'elimina', 'eliminar', 'borra', 'borrar',
    'resalta', 'resaltar', 'ajusta', 'ajustar', 'inserta', 'insertar',
    'reemplaza', 'reemplazar', 'actualiza', 'actualizar'
  ],

  QUESTION_WORDS: [
    'what', 'which', 'how many', 'how much', 'who', 'when', 'where', 'why',
    'is there', 'are there', 'do i', 'does', 'can you tell', 'show me',
    'qué', 'que ', 'cuál', 'cual', 'cuánto', 'cuanto', 'cuántos', 'cuantos',
    'cuántas', 'cuantas', 'quién', 'quien', 'cuándo', 'cuando', 'dónde',
    'donde', 'por qué', 'porque', 'hay ', 'existe', 'dime', 'muéstrame'
  ],

  // Questions about Excel itself rather than this workbook.
  EXPLAIN_WORDS: [
    'how do i use', 'how does', 'what is the difference', 'explain',
    'what does the', 'syntax', 'shortcut', 'formula for',
    'cómo se usa', 'como se usa', 'qué diferencia', 'que diferencia',
    'explica', 'explícame', 'explicame', 'sintaxis', 'atajo'
  ],

  AGGREGATES: {
    sum:     ['total of', 'sum of', 'total', 'sum', 'suma de', 'suma', 'total de', 'sumatoria'],
    average: ['average of', 'mean of', 'average', 'mean', 'promedio de', 'promedio', 'media de'],
    max:     ['maximum of', 'max of', 'highest', 'largest', 'maximum', 'max',
              'máximo de', 'maximo de', 'máximo', 'maximo', 'mayor'],
    min:     ['minimum of', 'min of', 'lowest', 'smallest', 'minimum', 'min',
              'mínimo de', 'minimo de', 'mínimo', 'minimo', 'menor'],
    count:   ['how many rows', 'how many records', 'row count', 'number of rows',
              'cuántas filas', 'cuantas filas', 'cuántos registros', 'cuantos registros',
              'número de filas', 'numero de filas']
  },

  /**
   * Words that indicate the question is about a SUBSET of the data. Column
   * statistics are whole-column aggregates, so they cannot answer these —
   * the fast path must decline.
   */
  SUBSET_WORDS: [
    ' by ', ' per ', ' for ', ' in ', ' where ', ' filtered', ' only ',
    ' each ', ' top ', ' bottom ', ' between ', ' excluding ', ' without ',
    ' vs ', ' versus ', ' compared',
    ' por ', ' para ', ' en ', ' donde ', ' solo ', ' sólo ', ' cada ',
    ' entre ', ' excluyendo ', ' sin ', ' mejor', ' peor', ' comparad'
  ],

  /**
   * Pick the prompt module for a request.
   */
  classify(text) {
    const t = ' ' + String(text || '').toLowerCase().trim() + ' ';

    if (this._hasAny(t, this.EXPLAIN_WORDS) && !this._mentionsWorkbook(t)) {
      return 'explain';
    }
    if (this._hasAny(t, this.BUILD_WORDS)) return 'build';
    if (this._hasAny(t, this.EDIT_WORDS)) return 'edit';
    if (this._hasAny(t, this.QUESTION_WORDS) || t.includes('?')) return 'qa';

    // Default to build: an imperative with no question marker is usually
    // a request to make something.
    return 'build';
  },

  _hasAny(haystack, needles) {
    return needles.some(n => haystack.includes(n));
  },

  _mentionsWorkbook(t) {
    return ['my data', 'my sheet', 'my workbook', 'this sheet', 'this workbook',
            'mis datos', 'mi hoja', 'mi libro', 'esta hoja', 'este libro']
      .some(w => t.includes(w));
  },

  /**
   * Try to satisfy the request with no API call at all.
   *
   * Returns { kind:'answer', text, source } when answered locally,
   * or null to fall through to the planner.
   */
  fastPath(userText, snap) {
    if (!snap || Context.isEmpty(snap)) return null;
    const text = String(userText || '').toLowerCase();

    // Subset questions need real row-level filtering, which the whole-column
    // statistics cannot provide. Decline rather than answer approximately.
    const padded = ' ' + text + ' ';
    if (this._hasAny(padded, this.SUBSET_WORDS)) return null;

    // Detect a specific sheet mention so we can restrict the search.
    const scopeSheet = Context.detectSheetScope(userText, snap);

    // Row-count questions.
    const rowCount = this._rowCountAnswer(padded, snap, scopeSheet);
    if (rowCount) return rowCount;

    return this._aggregateAnswer(text, snap, scopeSheet);
  },

  _rowCountAnswer(padded, snap, scopeSheet) {
    if (!this._hasAny(padded, this.AGGREGATES.count)) return null;
    let sheets = snap.sheets.filter(s => !s.empty && !s.error && !s.isCopilotSheet);
    if (scopeSheet) {
      // Restrict to the named sheet when the user specified one.
      const norm = (s) => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      const target = norm(scopeSheet);
      sheets = sheets.filter(s => norm(s.name) === target);
    }
    if (sheets.length !== 1) return null;   // ambiguous which sheet is meant
    const s = sheets[0];
    const dataRows = s.hasHeaders ? s.rows - 1 : s.rows;
    return {
      kind: 'answer',
      source: 'local',
      text: I18n.lang === 'es'
        ? `La hoja **${s.name}** tiene **${this._fmt(dataRows)}** filas de datos (${s.cols} columnas, rango ${s.address}).`
        : `Sheet **${s.name}** has **${this._fmt(dataRows)}** data rows (${s.cols} columns, range ${s.address}).`
    };
  },

  /**
   * Match "<aggregate> of <column>" against a numeric column with exact
   * precomputed statistics. When scopeSheet is provided, only that sheet
   * is searched, so a column name shared across sheets doesn't produce a
   * false match on the wrong sheet.
   */
  _aggregateAnswer(text, snap, scopeSheet) {
    let agg = null;
    for (const [name, words] of Object.entries(this.AGGREGATES)) {
      if (name === 'count') continue;
      if (this._hasAny(text, words)) { agg = name; break; }
    }
    if (!agg) return null;

    // Find every numeric column whose header appears in the question.
    // A single unambiguous hit is required.
    const hits = [];
    const norm = (s) => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const target = scopeSheet ? norm(scopeSheet) : null;

    for (const s of snap.sheets) {
      if (s.empty || s.error || !s.hasHeaders || !Array.isArray(s.columnStats)) continue;
      if (target && norm(s.name) !== target) continue;   // restrict to scoped sheet
      for (let c = 0; c < s.headers.length; c++) {
        const st = s.columnStats[c];
        if (!st) continue;
        const header = String(s.headers[c]).toLowerCase().trim();
        if (header.length < 3) continue;   // too short to match confidently
        if (text.includes(header)) hits.push({ sheet: s, header: s.headers[c], stats: st, partial: s.statsPartial });
      }
    }

    if (hits.length !== 1) return null;    // zero or ambiguous → let the model decide
    const hit = hits[0];
    if (hit.partial) return null;          // stats don't cover all rows

    const value = hit.stats[agg === 'average' ? 'avg' : agg];
    if (value === undefined || value === null) return null;

    const label = {
      sum:     { es: 'Suma',    en: 'Sum' },
      average: { es: 'Promedio', en: 'Average' },
      max:     { es: 'Máximo',  en: 'Maximum' },
      min:     { es: 'Mínimo',  en: 'Minimum' }
    }[agg][I18n.lang === 'es' ? 'es' : 'en'];

    // State the computation explicitly so a misread question is obvious.
    return {
      kind: 'answer',
      source: 'local',
      text: I18n.lang === 'es'
        ? `**${label}** de **${hit.header}** en la hoja **${hit.sheet.name}**: **${this._fmt(value)}**\n\nCalculado sobre las ${this._fmt(hit.stats.count)} filas con valores numéricos.`
        : `**${label}** of **${hit.header}** on sheet **${hit.sheet.name}**: **${this._fmt(value)}**\n\nComputed over the ${this._fmt(hit.stats.count)} rows containing numeric values.`
    };
  },

  _fmt(n) {
    if (typeof n !== 'number') return String(n);
    if (!isFinite(n)) return String(n);
    const decimals = Number.isInteger(n) ? 0 : 2;
    return n.toLocaleString(I18n.lang === 'es' ? 'es-ES' : 'en-US', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  }
};
