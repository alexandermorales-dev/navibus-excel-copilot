/* ============================================
   context.js — Workbook context for the planner

   The previous design made the model discover the workbook by calling
   get_workbook_overview and then read_range several times, costing one
   API request per step before any work could start. Office.js reads are
   free and fast, so the snapshot is taken locally and handed to the model
   in its first message instead. That alone removes 3-5 requests per task.

   The snapshot is placed in the first USER message rather than the system
   prompt so the system prefix stays byte-identical between calls, which
   is what lets providers serve it from prompt cache (on Groq, cached
   tokens don't count against rate limits).
   ============================================ */

const Context = {
  // Token ceiling for the rendered snapshot. Kept well below the smallest
  // provider's per-minute budget so a planning call always fits.
  MAX_SNAPSHOT_TOKENS: 3000,

  // Sheets created by the copilot in this session, so they can be labelled
  // as output rather than mistaken for source data on follow-up requests.
  copilotSheets: [],

  _cache: null,

  noteCopilotSheet(name) {
    if (name && !this.copilotSheets.includes(name)) this.copilotSheets.push(name);
  },

  invalidate() {
    this._cache = null;
  },

  reset() {
    this._cache = null;
    this.copilotSheets = [];
  },

  /**
   * Take a fresh workbook snapshot. Cached per run; callers invalidate
   * after mutations.
   */
  async snapshot({ force = false } = {}) {
    if (this._cache && !force) return this._cache;
    const snap = await Schema.snapshot(this.copilotSheets);
    this._cache = snap;
    return snap;
  },

  /**
   * Render a snapshot at the most detailed level that fits the budget.
   * Large workbooks degrade to less detail rather than blowing the budget
   * or getting truncated mid-table.
   *
   * Returns { text, level, tokens }
   */
  render(snap, { maxTokens = this.MAX_SNAPSHOT_TOKENS } = {}) {
    for (const level of ['full', 'normal', 'compact']) {
      const text = Schema.toText(snap, level);
      const tokens = Quota.estimateTokens(text);
      if (tokens <= maxTokens || level === 'compact') {
        return { text: this._clip(text, maxTokens), level, tokens };
      }
    }
    // Unreachable, but keep the contract explicit.
    return { text: '', level: 'compact', tokens: 0 };
  },

  /**
   * Hard clip as a last resort, on a line boundary so a partial row never
   * looks like real data to the model.
   */
  _clip(text, maxTokens) {
    const maxChars = maxTokens * 4;
    if (text.length <= maxChars) return text;
    const cut = text.lastIndexOf('\n', maxChars);
    return text.slice(0, cut > 0 ? cut : maxChars) +
      '\n    [snapshot truncated — call read_data for any sheet not shown above]';
  },

  /**
   * True when there is nothing to work with, so we can answer without
   * spending a request.
   */
  isEmpty(snap) {
    if (!snap || snap.sheetCount === 0) return true;
    return snap.sheets.every(s => s.empty || s.error);
  },

  /**
   * Look up a sheet description by name, case-insensitively.
   */
  findSheet(snap, name) {
    if (!snap || !name) return null;
    const target = String(name).trim().toLowerCase();
    return snap.sheets.find(s => String(s.name).toLowerCase() === target) || null;
  },

  /**
   * Resolve a possibly-misspelled sheet name to a real one. Small models
   * routinely get case and accents wrong; fixing it locally is free,
   * whereas letting the op fail costs a repair request.
   */
  resolveSheetName(snap, name) {
    if (!snap || !name) return null;
    const exact = this.findSheet(snap, name);
    if (exact) return exact.name;

    const norm = (s) => String(s).toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');
    const target = norm(name);
    const hit = snap.sheets.find(s => norm(s.name) === target);
    return hit ? hit.name : null;
  },

  /**
   * Index of a column by header name on a sheet, or -1.
   */
  columnIndex(sheetDesc, header) {
    if (!sheetDesc || !sheetDesc.headers || !header) return -1;
    const norm = (s) => String(s).trim().toLowerCase();
    const target = norm(header);
    return sheetDesc.headers.findIndex(h => norm(h) === target);
  },

  /**
   * The primary data sheet: the largest non-empty sheet that the copilot
   * did not create. Used when the user doesn't name a sheet.
   */
  primarySheet(snap) {
    if (!snap) return null;
    const candidates = snap.sheets.filter(s => !s.empty && !s.error && !s.isCopilotSheet);
    if (candidates.length === 0) return null;
    return candidates.reduce((best, s) => (s.rows * s.cols > best.rows * best.cols ? s : best));
  },

  /**
   * Build the messages array for a planning call: snapshot plus request.
   * A single user message keeps the shape stable across providers.
   */
  buildPlanMessages({ userText, snap, maxTokens, extra }) {
    const rendered = this.render(snap, { maxTokens });
    const parts = [
      '## WORKBOOK',
      rendered.text,
      '',
      '## REQUEST',
      userText
    ];
    if (extra) {
      parts.push('', '## ADDITIONAL DATA', extra);
    }
    return {
      messages: [{ role: 'user', content: parts.join('\n') }],
      level: rendered.level,
      tokens: rendered.tokens
    };
  },

  /**
   * Render extra ranges the planner explicitly asked for, so the follow-up
   * call carries exactly the data it said it needed and nothing more.
   */
  renderNeeds(reads) {
    const out = [];
    for (const r of reads) {
      if (!r.ok) {
        out.push(`${r.sheet}!${r.range}: ERROR ${r.error}`);
        continue;
      }
      const rows = (r.data || []).map(row => `[${row.join(' | ')}]`);
      out.push(`${r.sheet}!${r.address || r.range}${r.truncated ? ' (truncated)' : ''}:`);
      out.push(...rows.map(l => '  ' + l));
    }
    return out.join('\n');
  }
};
