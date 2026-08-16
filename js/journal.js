/* ============================================
   journal.js — Mutation journal for "Undo last request"
   Records pre-images of every cell range mutated during a user request
   (values + number formats) and every object (sheet/table/pivot/chart/
   slicer) created or modified, so a single user request can be undone
   even though Office.js bypasses Excel's native undo stack.

   Scope: per-request. App calls beginRequest() when a user message
   starts and the journal is sealed when the agent loop ends. The final
   message gets an "Undo" button iff the sealed journal has entries.
   ============================================ */

const Journal = {
  // Stack of completed requests, newest last. Each entry:
  //   { id, startedAt, entries: [{type, ...}] }
  // We keep only the most recent few to bound memory.
  history: [],
  MAX_HISTORY: 10,

  // The currently-open request (or null).
  current: null,

  beginRequest() {
    this.current = { id: Date.now() + '-' + Math.random().toString(36).slice(2, 8),
                     startedAt: Date.now(),
                     entries: [] };
  },

  isActive() {
    return this.current !== null;
  },

  /**
   * Record a pre-image of a range about to be overwritten.
   * Captures values + numberFormats so undo restores both.
   * Silently no-ops if not in an active request (e.g. read-only tools).
   */
  async recordRangePreImage(sheetName, rangeAddr) {
    if (!this.current) return;
    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(rangeAddr);
        range.load('values, numberFormats, address, rowCount, columnCount');
        await ctx.sync();
        // Skip truly empty ranges to save memory.
        const hasValues = range.values.some(row => row.some(v => v !== null && v !== ''));
        this.current.entries.push({
          type: 'range',
          sheet: sheetName,
          range: range.address, // normalized full address
          values: hasValues ? range.values : null,
          formats: hasValues ? range.numberFormats : null,
          hadValues: hasValues
        });
      });
    } catch (e) {
      // Best-effort: if pre-image capture fails (e.g. range doesn't exist yet),
      // record a "clear on undo" entry.
      this.current.entries.push({
        type: 'range',
        sheet: sheetName,
        range: rangeAddr,
        values: null,
        formats: null,
        hadValues: false
      });
    }
  },

  /**
   * Record that an object was created during this request so undo can delete it.
   */
  recordCreatedObject(kind, name, sheet) {
    if (!this.current) return;
    this.current.entries.push({ type: 'created', kind, name, sheet });
  },

  /**
   * Record that an existing object was modified (e.g. a table's style changed)
   * — undo will not attempt to restore the prior state of these (too complex),
   * but we log it so the UI can warn "partial undo".
   */
  recordModifiedObject(kind, name, sheet) {
    if (!this.current) return;
    this.current.entries.push({ type: 'modified', kind, name, sheet });
  },

  /**
   * Record that rows/cols were inserted or deleted on an existing sheet.
   * Undo of structural changes is best-effort and flagged as "partial".
   */
  recordStructuralChange(sheetName, kind, details) {
    if (!this.current) return;
    this.current.entries.push({ type: 'structural', sheet: sheetName, kind, details });
  },

  /**
   * Seal the current request and push onto history. Returns the sealed
   * request (or null if no entries were recorded, meaning nothing to undo).
   */
  sealRequest() {
    if (!this.current) return null;
    const req = this.current;
    this.current = null;
    if (req.entries.length === 0) return null;
    this.history.push(req);
    while (this.history.length > this.MAX_HISTORY) this.history.shift();
    return req;
  },

  discardRequest() {
    this.current = null;
  },

  hasUndoable() {
    return this.history.length > 0;
  },

  lastRequest() {
    return this.history.length > 0 ? this.history[this.history.length - 1] : null;
  },

  /**
   * Undo the most recent request. Reverses entries in reverse order:
   *   - created objects → delete
   *   - range pre-images → restore values+formats (or clear if hadValues=false)
   *   - structural / modified → cannot fully restore; flagged in result
   *
   * Returns: { ok, undone, skipped, partial, error? }
   */
  async undoLast() {
    const req = this.history.pop();
    if (!req) return { ok: false, error: 'Nothing to undo.' };

    const undone = [];
    const skipped = [];
    let partial = false;

    // Reverse so deletions happen before range restores.
    const entries = [...req.entries].reverse();

    try {
      await Excel.run(async (ctx) => {
        for (const e of entries) {
          try {
            if (e.type === 'created') {
              if (e.kind === 'sheet') {
                const s = ctx.workbook.worksheets.getItem(e.name);
                s.delete();
              } else if (e.kind === 'table') {
                const t = ctx.workbook.tables.getItem(e.name);
                t.delete();
              } else if (e.kind === 'pivot') {
                const p = ctx.workbook.pivotTables.getItem(e.name);
                p.delete();
              } else if (e.kind === 'chart') {
                const s = ctx.workbook.worksheets.getItem(e.sheet);
                const charts = s.charts;
                charts.load('items/name, items/id');
                await ctx.sync();
                for (const c of charts.items) {
                  if (c.name === e.name || c.id === e.name) { c.delete(); break; }
                }
              } else if (e.kind === 'slicer') {
                const sl = ctx.workbook.slicers.getItem(e.name);
                sl.delete();
              }
              undone.push(`${e.kind} "${e.name}"`);
            } else if (e.type === 'range') {
              const sheet = ctx.workbook.worksheets.getItem(e.sheet);
              const range = sheet.getRange(e.range);
              if (e.hadValues) {
                range.values = e.values;
                range.numberFormats = e.formats;
              } else {
                range.clear();
              }
              undone.push(`range ${e.sheet}!${e.range}`);
            } else if (e.type === 'structural' || e.type === 'modified') {
              partial = true;
              skipped.push(`${e.type} on ${e.sheet || e.name}`);
            }
          } catch (err) {
            partial = true;
            skipped.push(`${e.type} (${err.message})`);
          }
        }
        await ctx.sync();
      });
    } catch (err) {
      return { ok: false, error: err.message, undone, skipped, partial: true };
    }

    return { ok: true, undone, skipped, partial };
  },

  /**
   * Clear all history (e.g. on chat clear).
   */
  clear() {
    this.history = [];
    this.current = null;
  }
};
