/* ============================================
   agent.js — Plan / execute / verify orchestrator

   Replaces the previous iterative ReAct loop, which made up to 15 API
   requests per user message: one to discover the workbook, several to
   read ranges, several to write, several to verify, plus supervisor
   calls. On a free tier capped at 50 requests a day that allowed about
   three tasks before running dry.

   The flow now is:

     snapshot        local, free
     fast path       0 calls when the answer is already computed
     plan            1 call, with the whole workbook already in context
     [data fetch]    at most 1 extra call, only if the planner asks
     validate        local, free — rejects bad ops before Excel sees them
     execute         local, via the existing Tools handlers
     verify          local, free — reads results back, finds #REF! etc.
     repair          only when verification found something, max 2

   Typical cost is 1 request; worst case 4. Everything that does not need
   a language model is done in code, where it is also deterministic and
   testable.
   ============================================ */

const Agent = {
  MAX_REPAIRS: 2,
  PLAN_MAX_TOKENS: 6000,
  REPAIR_MAX_TOKENS: 3000,
  ANSWER_MAX_TOKENS: 800,

  /**
   * @param {object} opts
   * @param {string} opts.userText
   * @param {Array}  opts.conversation — prior turns, for follow-up context
   * @param {AbortSignal} [opts.signal]
   * @param {function} [opts.onPhase]     — (phase, detail)
   * @param {function} [opts.onThinking]  — (reasoningText)
   * @param {function} [opts.onProvider]  — (label, model, isFailover)
   * @param {function} [opts.onPlan]      — (ops, planText)
   * @param {function} [opts.onOpStart]   — (id, opName, args)
   * @param {function} [opts.onOpEnd]     — (id, opName, result)
   *
   * Returns { ok, finalText, calls, sealed, aborted, warnings, error }
   */
  async run(opts) {
    const {
      userText, conversation = [], signal,
      onPhase, onThinking, onProvider, onPlan, onOpStart, onOpEnd
    } = opts;

    const phase = (p, d) => { if (onPhase) onPhase(p, d); };
    const lang = Prompt.detectLang(userText);
    const stats = { calls: 0, tokens: 0 };
    const warnings = [];

    Journal.beginRequest();

    try {
      /* ---------- 1. Snapshot (local, free) ---------- */
      phase('reading');
      let snap;
      try {
        snap = await Context.snapshot({ force: true });
      } catch (e) {
        return this._fail(I18n.t('workbookError') + ': ' + (e.message || e), stats);
      }

      if (Context.isEmpty(snap)) {
        return this._done(I18n.t('emptyWorkbook'), stats, warnings);
      }

      /* ---------- 2. Local fast path (0 calls) ---------- */
      const intent = Intent.classify(userText);
      if (intent === 'qa') {
        const fast = Intent.fastPath(userText, snap);
        if (fast) {
          console.log('Agent: answered locally, 0 API calls');
          return this._done(fast.text, stats, warnings);
        }
      }

      /* ---------- 3. Plan (1 call) ---------- */
      phase('planning');
      const planned = await this._plan({
        userText, snap, intent, lang, conversation, signal,
        onThinking, onProvider, stats
      });
      if (!planned.ok) return this._fail(planned.error, stats, planned.errorType);
      if (signal && signal.aborted) return this._aborted(stats);

      let plan = planned.plan;

      /* ---------- 3b. One optional data-fetch round ---------- */
      if (Array.isArray(plan.needs) && plan.needs.length > 0 && (!plan.ops || plan.ops.length === 0)) {
        phase('reading');
        const extra = await this._fetchNeeds(plan.needs);
        const second = await this._plan({
          userText, snap, intent, lang, conversation, signal,
          onThinking, onProvider, stats,
          extra,
          allowNeeds: false   // one fetch only; no further escalation
        });
        if (!second.ok) return this._fail(second.error, stats, second.errorType);
        plan = second.plan;
      }

      /* ---------- 4. Validate (local, free) ---------- */
      const validation = Ops.validate(plan, snap);
      if (!validation.ok) return this._fail(I18n.t('badPlan'), stats);
      warnings.push(...validation.warnings);

      let answer = typeof plan.answer === 'string' ? plan.answer.trim() : '';

      // Nothing to do: a question answered in text.
      if (validation.ops.length === 0) {
        if (validation.dropped.length > 0) {
          console.warn('Agent: all ops rejected', validation.dropped);
        }
        if (!answer) return this._fail(I18n.t('badPlan'), stats);
        return this._done(answer, stats, warnings, validation.dropped);
      }

      if (onPlan) onPlan(validation.ops, plan.plan || '');

      /* ---------- 5. Execute (local) ---------- */
      phase('writing');
      let exec = await Ops.execute(validation.ops, { snap, signal, onOpStart, onOpEnd });
      if (signal && signal.aborted) {
        return this._aborted(stats, exec);
      }

      /* ---------- 6. Verify (local, free) ---------- */
      phase('verifying');
      let verification = await Ops.verify(exec, { signal });

      /* ---------- 7. Repair (only if needed) ---------- */
      let repairs = 0;
      let problems = this._collectProblems(validation, exec, verification);

      while (problems.length > 0 && repairs < this.MAX_REPAIRS) {
        if (signal && signal.aborted) break;
        repairs++;
        phase('repairing', I18n.tf('repairing', problems.length));

        const repaired = await this._repair({
          userText, snap, lang, signal, stats, onThinking, onProvider,
          description: Ops.describeProblems({
            dropped: validation.dropped,
            failed: exec.failed,
            problems: verification.problems
          })
        });
        if (!repaired.ok || !repaired.plan) break;

        const rv = Ops.validate(repaired.plan, await Context.snapshot({ force: true }));
        if (rv.ops.length === 0) break;

        const rexec = await Ops.execute(rv.ops, { snap, signal, onOpStart, onOpEnd });
        const rverify = await Ops.verify(rexec, { signal });

        // Merge so the final answer reflects everything that happened.
        exec = this._mergeExec(exec, rexec);
        verification = rverify;
        validation.dropped = rv.dropped;
        if (repaired.plan.answer) answer = String(repaired.plan.answer).trim();

        const next = this._collectProblems(rv, rexec, rverify);
        if (next.length >= problems.length) break;  // not converging — stop
        problems = next;
      }

      /* ---------- 8. Final answer, with verified values appended ---------- */
      phase('done');
      const finalText = this._composeAnswer({
        answer, exec, verification, warnings, lang, problems
      });

      return this._done(finalText, stats, warnings, validation.dropped);

    } catch (e) {
      console.error('Agent: unexpected failure', e);
      return this._fail(e.message || String(e), stats);
    }
  },

  /* ----------------------------------------------------------
     PLANNING
     ---------------------------------------------------------- */
  async _plan({ userText, snap, intent, lang, conversation, signal, onThinking, onProvider, stats, extra, allowNeeds = true }) {
    const systemPrompt = Prompt.build({ intent, lang, allowNeeds });
    const built = Context.buildPlanMessages({ userText, snap, extra });

    // Carry a little prior context so follow-ups ("now add a chart") work,
    // without replaying the whole transcript.
    const history = this._recentHistory(conversation);
    const messages = [...history, ...built.messages];

    const res = await LLM.chat({
      role: 'plan',
      systemPrompt,
      messages,
      json: true,
      maxTokens: this.PLAN_MAX_TOKENS,
      temperature: 0.2,
      signal,
      onThinking,
      onProvider
    });

    stats.calls++;
    if (res.usage && res.usage.total_tokens) stats.tokens += res.usage.total_tokens;

    if (!res.ok) return { ok: false, error: res.error, errorType: res.errorType };

    const parsed = LLM.extractJSON(res.text);
    if (!parsed.ok) {
      console.warn('Agent: unparseable plan:', String(res.text || '').slice(0, 400));
      return { ok: false, error: I18n.t('badPlan'), errorType: 'plan' };
    }

    console.log(`Agent: plan from ${res.providerId}/${res.model} (snapshot detail: ${built.level}, ${built.tokens} tok)`);
    return { ok: true, plan: parsed.value };
  },

  async _repair({ userText, snap, lang, signal, stats, onThinking, onProvider, description }) {
    const rendered = Context.render(snap, { maxTokens: 1200 });
    const messages = [{
      role: 'user',
      content: [
        '## WORKBOOK', rendered.text, '',
        '## ORIGINAL REQUEST', userText, '',
        '## PROBLEMS TO FIX', description
      ].join('\n')
    }];

    const res = await LLM.chat({
      role: 'repair',
      systemPrompt: Prompt.repair(lang),
      messages,
      json: true,
      maxTokens: this.REPAIR_MAX_TOKENS,
      temperature: 0.1,
      signal,
      onThinking,
      onProvider
    });

    stats.calls++;
    if (res.usage && res.usage.total_tokens) stats.tokens += res.usage.total_tokens;
    if (!res.ok) return { ok: false, error: res.error };

    const parsed = LLM.extractJSON(res.text);
    if (!parsed.ok) return { ok: false, error: I18n.t('badPlan') };
    return { ok: true, plan: parsed.value };
  },

  /**
   * Read the ranges the planner asked for. Office.js reads are free, so
   * this costs nothing beyond the single follow-up planning call.
   */
  async _fetchNeeds(needs) {
    const reads = [];
    for (const n of needs.slice(0, 5)) {
      if (!n || !n.sheet || !n.range) continue;
      const res = await Tools.dispatch('read_range', { sheet: n.sheet, range: n.range });
      reads.push(res.ok
        ? { ok: true, sheet: n.sheet, range: n.range, ...res.result }
        : { ok: false, sheet: n.sheet, range: n.range, error: res.error });
    }
    return Context.renderNeeds(reads);
  },

  /**
   * Last user/assistant exchange only. Enough for "now sort it descending"
   * to make sense, without the unbounded context growth of the old loop.
   */
  _recentHistory(conversation) {
    const usable = conversation.filter(m => m.role === 'user' || m.role === 'assistant');
    const tail = usable.slice(-2);
    return tail.map(m => ({
      role: m.role,
      content: String(m.content || '').slice(0, 600)
    }));
  },

  /* ----------------------------------------------------------
     PROBLEM COLLECTION
     ---------------------------------------------------------- */

  /**
   * Only problems worth spending a request on. Cosmetic warnings and ops
   * the validator already repaired are excluded.
   */
  _collectProblems(validation, exec, verification) {
    const out = [];
    for (const d of validation.dropped) out.push({ kind: 'rejected', detail: d.reason });
    for (const f of exec.failed) out.push({ kind: 'failed', detail: `${f.op}: ${f.error}` });
    for (const p of verification.problems) {
      // A zero KPI is often legitimate (an empty category), so it is only
      // worth a repair when something else is also wrong.
      if (p.kind === 'suspicious_zero') continue;
      out.push({ kind: p.kind, detail: p.detail });
    }
    return out;
  },

  _mergeExec(a, b) {
    return {
      results: [...a.results, ...b.results],
      failed: b.failed,
      executed: a.executed + b.executed,
      createdSheets: [...a.createdSheets, ...b.createdSheets],
      writes: [...a.writes, ...b.writes]
    };
  },

  /* ----------------------------------------------------------
     ANSWER COMPOSITION
     ---------------------------------------------------------- */

  /**
   * Build the user-facing message. The planner supplies prose; the
   * verified numbers are appended from cells actually read back, so the
   * answer cites real values without spending another API call and
   * without the model being able to invent figures.
   */
  _composeAnswer({ answer, exec, verification, warnings, lang, problems }) {
    const es = lang === 'es';
    const parts = [];

    parts.push(answer || this._describeWork(exec, es));

    const labelled = Ops.labelValues(exec, verification.values)
      .filter(v => v.label && v.value !== null && v.value !== undefined && v.value !== '');

    if (labelled.length > 0) {
      parts.push('');
      parts.push(es ? '**Valores verificados:**' : '**Verified values:**');
      for (const v of labelled) {
        parts.push(`- ${v.label}: **${Intent._fmt(v.value)}**`);
      }
    }

    if (exec.createdSheets.length > 0) {
      parts.push('');
      parts.push(es
        ? `Hoja creada: ${exec.createdSheets.map(s => `**${s}**`).join(', ')}`
        : `Created sheet: ${exec.createdSheets.map(s => `**${s}**`).join(', ')}`);
    }

    // Be explicit about what did not work rather than reporting success.
    if (problems && problems.length > 0) {
      parts.push('');
      parts.push(es
        ? `No pude resolver ${problems.length} problema(s):`
        : `${problems.length} issue(s) could not be resolved:`);
      for (const p of problems.slice(0, 5)) parts.push(`- ${p.detail}`);
    }

    const realWarnings = (warnings || []).filter(Boolean);
    if (realWarnings.length > 0) {
      parts.push('');
      parts.push(es ? '_Notas:_' : '_Notes:_');
      for (const w of realWarnings.slice(0, 5)) parts.push(`- _${w}_`);
    }

    return parts.join('\n');
  },

  _describeWork(exec, es) {
    if (exec.executed === 0) {
      return es
        ? 'No se realizaron cambios en el libro.'
        : 'No changes were made to the workbook.';
    }
    const counts = {};
    for (const r of exec.results) {
      if (!r.ok || r.skipped) continue;
      counts[r.op] = (counts[r.op] || 0) + 1;
    }
    const bits = Object.entries(counts).map(([op, n]) => `${n}x ${I18n.toolLabel(op)}`);
    return (es ? 'Listo. ' : 'Done. ') + bits.join(', ');
  },

  /* ----------------------------------------------------------
     RESULT HELPERS
     ---------------------------------------------------------- */
  _done(finalText, stats, warnings, dropped) {
    const sealed = Journal.sealRequest();
    return {
      ok: true,
      finalText,
      calls: stats.calls,
      tokens: stats.tokens,
      sealed,
      warnings: warnings || [],
      dropped: dropped || []
    };
  },

  _fail(error, stats, errorType) {
    const sealed = Journal.sealRequest();
    return {
      ok: false,
      error,
      errorType: errorType || 'unknown',
      calls: stats.calls,
      tokens: stats.tokens,
      sealed
    };
  },

  _aborted(stats, exec) {
    const sealed = Journal.sealRequest();
    return {
      ok: true,
      aborted: true,
      finalText: I18n.t('aborted'),
      calls: stats.calls,
      tokens: stats.tokens,
      sealed,
      warnings: []
    };
  }
};
