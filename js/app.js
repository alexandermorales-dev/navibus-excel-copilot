/* ============================================
   app.js — UI controller for the agentic Excel Copilot
   Owns: chat DOM, settings, conversation history, and the
   Agent.run() lifecycle (start/stop/undo).
   No more plan/repair/summary logic — the agent handles all of that.
   ============================================ */

const App = {
  conversation: [],   // OpenAI messages format
  isRunning: false,
  lastUserText: '',   // for manual retry
  abortController: null, // for Stop button

  el: {},

  init() {
    this.el = {
      messageList: document.getElementById('messageList'),
      messageInput: document.getElementById('messageInput'),
      sendBtn: document.getElementById('sendBtn'),
      stopBtn: document.getElementById('stopBtn'),
      settingsBtn: document.getElementById('settingsBtn'),
      settingsPanel: document.getElementById('settingsPanel'),
      providerList: document.getElementById('providerList'),
      advancedList: document.getElementById('advancedList'),
      resetQuotaBtn: document.getElementById('resetQuotaBtn'),
      quotaBar: document.getElementById('quotaBar'),
      clearChatBtn: document.getElementById('clearChatBtn'),
      statusBar: document.getElementById('statusBar')
    };

    Config.load();
    Quota.load();

    this.el.sendBtn.addEventListener('click', () => this.sendMessage());
    this.el.stopBtn.addEventListener('click', () => this.stopRun());
    this.el.messageInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        this.sendMessage();
      }
    });
    this.el.messageInput.addEventListener('input', () => this.updateSendButton());

    this.el.settingsBtn.addEventListener('click', () => {
      this.el.settingsPanel.classList.toggle('hidden');
    });

    this.el.resetQuotaBtn.addEventListener('click', () => {
      Quota.reset();
      this.renderQuotaBar();
      this.addMessage('system', I18n.t('quotaReset'));
    });

    this.el.clearChatBtn.addEventListener('click', () => this.clearChat());

    const boot = () => {
      I18n.init();
      this.renderProviderSettings();
      this.localizeUI();
      this.renderQuotaBar();
      this.updateSendButton();
      this.updateStopButton();
      // No key yet → open Settings so onboarding is the first thing seen.
      if (!Config.hasApiKey()) this.el.settingsPanel.classList.remove('hidden');
    };

    if (typeof Office !== 'undefined' && Office.onReady) {
      Office.onReady((info) => {
        try {
          if (info && info.host === Office.HostType.Excel) {
            console.log('Excel AI Copilot ready');
          }
        } catch (e) { /* info or HostType may be undefined outside Excel */ }
        boot();
      });
      // Fallback: if Office.onReady doesn't fire within 3s (e.g. outside
      // Excel), initialize anyway so the UI isn't stuck in the default lang.
      setTimeout(() => { if (!I18n.initialized) boot(); }, 3000);
    } else {
      console.warn('Office.js not loaded — running in browser dev mode');
      boot();
    }
  },

  /* ----------------------------------------------------------
     PROVIDER SETTINGS / ONBOARDING
     Rows are generated from the Providers registry so adding a
     provider there needs no HTML changes.
     ---------------------------------------------------------- */
  renderProviderSettings() {
    const list = this.el.providerList;
    const advanced = this.el.advancedList;
    if (!list) return;
    list.innerHTML = '';
    if (advanced) advanced.innerHTML = '';

    for (const id of Providers.ids()) {
      const p = Providers.get(id);
      const row = document.createElement('div');
      row.className = 'provider-row';
      row.innerHTML = `
        <div class="provider-head">
          <label for="key_${id}">${this.escapeHtml(p.label)}</label>
          <a href="${p.keyUrl}" target="_blank" class="provider-link">${this.escapeHtml(I18n.t('getKey'))}</a>
        </div>
        <div class="api-key-wrapper">
          <input type="password" id="key_${id}" placeholder="${this.escapeHtml(p.keyPlaceholder)}" autocomplete="off" spellcheck="false" />
          <button type="button" class="api-key-toggle" data-toggle="${id}" title="Show/Hide" aria-label="Toggle visibility">
            <svg class="icon-eye" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
            </svg>
            <svg class="icon-eye-off hidden" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
              <line x1="1" y1="1" x2="23" y2="23"/>
            </svg>
          </button>
          <button type="button" class="btn-test" data-test="${id}">${this.escapeHtml(I18n.t('testKey'))}</button>
        </div>
        <p class="provider-status" data-status="${id}"></p>
      `;
      list.appendChild(row);

      const input = row.querySelector(`#key_${id}`);
      input.value = Config.keyFor(id);

      const commit = () => {
        Config.setKey(id, input.value);
        this.updateSendButton();
        this.renderQuotaBar();
      };
      input.addEventListener('change', commit);
      input.addEventListener('blur', commit);
      // Commit on input too: some Excel WebViews don't fire change reliably.
      input.addEventListener('input', commit);

      row.querySelector(`[data-toggle="${id}"]`).addEventListener('click', (e) => {
        const btn = e.currentTarget;
        const eye = btn.querySelector('.icon-eye');
        const eyeOff = btn.querySelector('.icon-eye-off');
        const show = input.type === 'password';
        input.type = show ? 'text' : 'password';
        eye.classList.toggle('hidden', show);
        eyeOff.classList.toggle('hidden', !show);
      });

      row.querySelector(`[data-test="${id}"]`).addEventListener('click', (e) => {
        this.testProvider(id, e.currentTarget);
      });

      this.showProviderStatus(id, Config.keyFor(id) ? '' : I18n.t('noKeySet'), 'muted');

      // Advanced: pin a specific model instead of automatic resolution.
      if (advanced) {
        const adv = document.createElement('div');
        adv.className = 'settings-row';
        adv.innerHTML = `
          <label for="model_${id}">${this.escapeHtml(p.label)} — ${this.escapeHtml(I18n.t('modelOverride'))}</label>
          <select id="model_${id}"></select>
        `;
        advanced.appendChild(adv);
        this._fillModelSelect(id, adv.querySelector(`#model_${id}`));
      }
    }
  },

  _fillModelSelect(id, select) {
    if (!select) return;
    const current = Config.modelOverride(id);
    const models = Providers.discovered[id] || Config.loadDiscovered(id) || [];
    select.innerHTML = `<option value="">${this.escapeHtml(I18n.t('modelAuto'))}</option>` +
      models.map(m => `<option value="${this.escapeHtml(m)}">${this.escapeHtml(m)}</option>`).join('');
    select.value = current;
    select.onchange = () => Config.setModelOverride(id, select.value);
  },

  /**
   * Probe a provider's /models endpoint. A 200 proves the key works and
   * gives us the real model list, so resolution never relies on a
   * hardcoded slug that may have been deprecated.
   */
  async testProvider(id, btn) {
    if (!Config.keyFor(id)) {
      this.showProviderStatus(id, I18n.t('noKeySet'), 'muted');
      return;
    }
    const original = btn.textContent;
    btn.disabled = true;
    btn.textContent = I18n.t('testing');
    this.showProviderStatus(id, I18n.t('testing'), 'muted');

    const res = await Providers.discover(id);

    btn.disabled = false;
    btn.textContent = original;

    if (res.ok) {
      const model = Providers.resolveModel(id, 'plan');
      this.showProviderStatus(id, I18n.tf('testOk', model), 'ok');
      const select = document.getElementById(`model_${id}`);
      if (select) this._fillModelSelect(id, select);
      this.renderQuotaBar();
    } else {
      this.showProviderStatus(id, I18n.tf('testFailed', res.error), 'error');
    }
  },

  showProviderStatus(id, text, kind) {
    const el = document.querySelector(`[data-status="${id}"]`);
    if (!el) return;
    el.textContent = text || '';
    el.className = `provider-status provider-status-${kind || 'muted'}`;
  },

  /* ----------------------------------------------------------
     QUOTA BAR
     ---------------------------------------------------------- */
  renderQuotaBar() {
    const bar = this.el.quotaBar;
    if (!bar) return;
    const rows = Quota.summary();
    if (rows.length === 0) {
      bar.classList.add('hidden');
      bar.innerHTML = '';
      return;
    }
    bar.classList.remove('hidden');
    const chips = rows.map(r => {
      let detail;
      if (r.state === 'invalid') detail = I18n.t('quotaInvalid');
      else if (r.state === 'exhausted') detail = I18n.t('quotaExhausted');
      else if (r.state === 'cooldown') detail = I18n.t('quotaCooldown');
      else detail = `${r.remaining}/${r.limit}`;
      return `<span class="quota-chip quota-${r.state}">${this.escapeHtml(r.label)} <b>${this.escapeHtml(detail)}</b></span>`;
    }).join('');
    bar.innerHTML = `<span class="quota-label">${this.escapeHtml(I18n.t('quotaTitle'))}</span>${chips}`;
    bar.title = I18n.t('quotaTitle');
  },

  localizeUI() {
    document.querySelectorAll('[data-i18n]').forEach(el => {
      el.textContent = I18n.t(el.getAttribute('data-i18n'));
    });
    document.querySelectorAll('[data-i18n-ph]').forEach(el => {
      el.placeholder = I18n.t(el.getAttribute('data-i18n-ph'));
    });
  },

  updateSendButton() {
    const hasText = this.el.messageInput.value.trim().length > 0;
    const hasKey = Config.hasApiKey();
    this.el.sendBtn.disabled = !hasText || !hasKey || this.isRunning;
  },

  updateStopButton() {
    this.el.stopBtn.disabled = !this.isRunning;
  },

  async sendMessage() {
    const text = this.el.messageInput.value.trim();
    if (!text || this.isRunning) return;

    if (!Config.hasApiKey()) {
      this.addMessage('system', I18n.t('needApiKey'));
      this.el.settingsPanel.classList.remove('hidden');
      return;
    }

    this.isRunning = true;
    this.updateSendButton();
    this.updateStopButton();
    this.el.messageInput.value = '';
    this.lastUserText = text;
    this._hasWrittenThisRun = false;

    this.addMessage('user', text);
    this.conversation.push({ role: 'user', content: text });

    this.abortController = new AbortController();

    // Create the live activity container for this run.
    const activityEl = this.addActivityContainer();

    try {
      const result = await Agent.run({
        userText: text,
        conversation: this.conversation,
        signal: this.abortController.signal,
        onText: (chunk, full) => this.updateLiveAnswer(activityEl, full),
        onThinking: (text) => this.updateLiveThinking(activityEl, text),
        onToolStart: (callId, name, args) => {
          this.addToolRow(activityEl, callId, name, args);
          if (this.isWriteTool(name)) this._hasWrittenThisRun = true;
          this.updateRunStatus(activityEl, this.toolPhase(name), this.toolPhaseDetail(name, args));
        },
        onToolEnd: (callId, name, toolResult) => this.finalizeToolRow(activityEl, callId, toolResult),
        onToolError: (name, error) => this.showToolError(activityEl, name, error),
        onRound: (round, maxRounds) => {
          if (round > 1) this.updateRunStatus(activityEl, 'thinking');
        }
      });

      this.finalizeLiveThinking(activityEl);

      if (!result.ok) {
        this.updateRunStatus(activityEl, 'error');
        this.finalizeLiveAnswer(activityEl);
        this.addErrorMessageWithRetry(`${I18n.t('aiError')}: ${result.error}`);
      } else {
        this.updateRunStatus(activityEl, result.aborted ? 'stopped' : 'done');
        this.collapseActivityFeed(activityEl);
        this.finalizeLiveAnswer(activityEl, result.sealed && !result.aborted);
        const ansEl = activityEl.querySelector('.live-answer');
        if ((!ansEl || ansEl.classList.contains('hidden')) && result.finalText) {
          this.addFinalMessage(result.finalText, result.sealed, result.aborted);
        }
      }
    } catch (e) {
      this.finalizeLiveThinking(activityEl);
      this.updateRunStatus(activityEl, 'error');
      this.collapseActivityFeed(activityEl);
      this.finalizeLiveAnswer(activityEl);
      this.addErrorMessageWithRetry(`${I18n.t('genericError')}: ${e.message || String(e)}`);
    }

    this.abortController = null;
    this.isRunning = false;
    this._hasWrittenThisRun = false;
    this.updateSendButton();
    this.updateStopButton();
  },

  stopRun() {
    if (this.abortController) this.abortController.abort();
  },

  async retryLastMessage() {
    if (this.isRunning || !this.lastUserText) return;
    this.isRunning = true;
    this._hasWrittenThisRun = false;
    this.updateSendButton();
    this.updateStopButton();
    this.abortController = new AbortController();

    const activityEl = this.addActivityContainer();
    this.updateLiveThinking(activityEl, I18n.t('retryingMessage'));

    try {
      const result = await Agent.run({
        userText: this.lastUserText,
        conversation: this.conversation,
        signal: this.abortController.signal,
        onText: (chunk, full) => this.updateLiveAnswer(activityEl, full),
        onThinking: (text) => this.updateLiveThinking(activityEl, text),
        onToolStart: (callId, name, args) => {
          this.addToolRow(activityEl, callId, name, args);
          if (this.isWriteTool(name)) this._hasWrittenThisRun = true;
          this.updateRunStatus(activityEl, this.toolPhase(name), this.toolPhaseDetail(name, args));
        },
        onToolEnd: (callId, name, toolResult) => this.finalizeToolRow(activityEl, callId, toolResult),
        onToolError: (name, error) => this.showToolError(activityEl, name, error),
        onRound: (round, maxRounds) => {
          if (round > 1) this.updateRunStatus(activityEl, 'thinking');
        }
      });
      this.finalizeLiveThinking(activityEl);
      if (!result.ok) {
        this.updateRunStatus(activityEl, 'error');
        this.finalizeLiveAnswer(activityEl);
        this.addErrorMessageWithRetry(`${I18n.t('aiError')}: ${result.error}`);
      } else {
        this.updateRunStatus(activityEl, result.aborted ? 'stopped' : 'done');
        this.collapseActivityFeed(activityEl);
        this.finalizeLiveAnswer(activityEl, result.sealed && !result.aborted);
        const ansEl = activityEl.querySelector('.live-answer');
        if ((!ansEl || ansEl.classList.contains('hidden')) && result.finalText) {
          this.addFinalMessage(result.finalText, result.sealed, result.aborted);
        }
      }
    } catch (e) {
      this.finalizeLiveThinking(activityEl);
      this.updateRunStatus(activityEl, 'error');
      this.collapseActivityFeed(activityEl);
      this.finalizeLiveAnswer(activityEl);
      this.addErrorMessageWithRetry(`${I18n.t('genericError')}: ${e.message || String(e)}`);
    }

    this.abortController = null;
    this.isRunning = false;
    this.updateSendButton();
    this.updateStopButton();
  },

  /* ----------------------------------------------------------
     ACTIVITY FEED UI
     ---------------------------------------------------------- */
  addActivityContainer() {
    const msg = document.createElement('div');
    msg.className = 'message assistant';
    msg.innerHTML = `
      <div class="message-avatar">AI</div>
      <div class="message-content">
        <div class="run-status run-status-active">
          <span class="run-status-spinner"></span>
          <span class="run-status-text">${this.escapeHtml(I18n.t('statusThinking'))}</span>
          <span class="run-status-count"></span>
        </div>
        <div class="thinking-block thinking-live">
          <div class="thinking-header">
            <span class="thinking-icon">\u{1F4AD}</span>
            <span class="thinking-label">${this.escapeHtml(I18n.t('reasoning'))}</span>
            <span class="thinking-dots"><span></span><span></span><span></span></span>
          </div>
          <div class="thinking-text thinking-stream"></div>
        </div>
        <div class="activity-feed"></div>
        <div class="live-answer hidden"></div>
      </div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  /**
   * Update the run status header to reflect the current phase.
   * Phases: thinking → reading → writing → verifying → done/error/stopped
   */
  updateRunStatus(activityEl, phase, detail) {
    if (!activityEl) return;
    const statusEl = activityEl.querySelector('.run-status');
    const textEl = activityEl.querySelector('.run-status-text');
    const countEl = activityEl.querySelector('.run-status-count');
    if (!statusEl || !textEl) return;

    statusEl.classList.remove('run-status-active', 'run-status-done', 'run-status-error', 'run-status-stopped');

    let label = '';
    switch (phase) {
      case 'thinking':  label = I18n.t('statusThinking');  statusEl.classList.add('run-status-active'); break;
      case 'reading':   label = I18n.t('statusReading');   statusEl.classList.add('run-status-active'); break;
      case 'writing':   label = I18n.t('statusWriting');   statusEl.classList.add('run-status-active'); break;
      case 'verifying': label = I18n.t('statusVerifying'); statusEl.classList.add('run-status-active'); break;
      case 'done':      label = I18n.t('statusDone');      statusEl.classList.add('run-status-done'); break;
      case 'error':     label = I18n.t('statusError');     statusEl.classList.add('run-status-error'); break;
      case 'stopped':   label = I18n.t('statusStopped');   statusEl.classList.add('run-status-stopped'); break;
      default: label = phase;
    }
    textEl.textContent = label;
    if (countEl) countEl.textContent = detail || '';
  },

  updateLiveThinking(activityEl, fullText) {
    if (!activityEl || !fullText) return;
    const streamEl = activityEl.querySelector('.thinking-stream');
    if (streamEl) streamEl.innerHTML = this.formatContent(this.escapeHtml(fullText));
    this.scrollToBottom();
  },

  finalizeLiveThinking(activityEl) {
    if (!activityEl) return;
    const header = activityEl.querySelector('.thinking-header');
    if (header) {
      header.innerHTML = `<span class="thinking-icon">\u{1F4AD}</span><span class="thinking-label">${this.escapeHtml(I18n.t('reasoningDone'))}</span>`;
    }
    const block = activityEl.querySelector('.thinking-block');
    if (block) {
      block.classList.remove('thinking-live');
      block.classList.add('thinking-done');
    }
    // Keep thinking visible by default — user can collapse if they want.
    const streamEl = activityEl.querySelector('.thinking-stream');
    if (streamEl && streamEl.textContent.trim().length > 0) {
      header.style.cursor = 'pointer';
      header.onclick = () => streamEl.classList.toggle('collapsed');
    }
  },

  updateLiveAnswer(activityEl, fullText) {
    if (!activityEl || !fullText) return;
    const ansEl = activityEl.querySelector('.live-answer');
    if (ansEl) {
      ansEl.classList.remove('hidden');
      ansEl._rawText = fullText;
      ansEl.innerHTML = this.formatContent(this.escapeHtml(fullText));
    }
    this.scrollToBottom();
  },

  finalizeLiveAnswer(activityEl, addUndo = false) {
    if (!activityEl) return;
    const ansEl = activityEl.querySelector('.live-answer');
    if (ansEl && ansEl.classList.contains('hidden')) {
      ansEl.remove();
    } else if (ansEl) {
      ansEl.classList.add('live-answer-done', 'final-answer');
      const rawText = ansEl._rawText || ansEl.textContent || '';
      ansEl.innerHTML = this.formatMarkdown(rawText);
      if (addUndo) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn-undo';
        btn.textContent = I18n.t('undo');
        btn.addEventListener('click', () => this.undoLastRequest(btn));
        ansEl.appendChild(btn);
      }
    }
  },

  addToolRow(activityEl, callId, name, args) {
    const feed = activityEl.querySelector('.activity-feed');
    if (!feed) return;
    const row = document.createElement('div');
    row.className = 'tool-row';
    row.dataset.callId = callId;
    row.innerHTML = `
      <span class="tool-spinner"></span>
      <span class="tool-label">${this.escapeHtml(I18n.toolLabel(name))}</span>
      <span class="tool-detail">${this.escapeHtml(this.toolDetail(name, args))}</span>
    `;
    feed.appendChild(row);
    this.scrollToBottom();
  },

  finalizeToolRow(activityEl, callId, toolResult) {
    const row = activityEl.querySelector(`.tool-row[data-call-id="${callId}"]`);
    if (!row) return;
    const spinner = row.querySelector('.tool-spinner');
    if (spinner) {
      spinner.classList.remove('tool-spinner');
      if (toolResult.ok) {
        spinner.classList.add('tool-ok');
        spinner.textContent = '\u2713';
      } else {
        spinner.classList.add('tool-warn');
        spinner.textContent = '\u26A0';
        row.classList.add('tool-row-warn');
      }
    }
    this.scrollToBottom();
  },

  showToolError(activityEl, toolName, error) {
    // Silently handled — the agent receives tool errors and can self-correct.
  },

  collapseActivityFeed(activityEl) {
    const feed = activityEl.querySelector('.activity-feed');
    if (!feed) return;
    const rows = feed.querySelectorAll('.tool-row');
    if (rows.length === 0) {
      feed.style.display = 'none';
      return;
    }
    feed.classList.add('activity-feed-collapsed');
    feed.onclick = () => feed.classList.toggle('activity-feed-collapsed');
    feed.title = I18n.lang === 'es' ? 'Clic para ver detalles' : 'Click to see details';
  },

  toolDetail(name, args) {
    // Short human-readable summary of a tool call's arguments.
    const a = args || {};
    switch (name) {
      case 'get_workbook_overview': return '';
      case 'read_range': return `${a.sheet || ''}!${a.range || ''}${a.what && a.what !== 'values' ? ' (' + a.what + ')' : ''}`;
      case 'find_in_workbook': return `"${a.query || ''}"${a.sheet ? ' en ' + a.sheet : ''}`;
      case 'get_objects': return a.sheet ? a.sheet : '';
      case 'write_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'format_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'clear_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'add_sheet': return a.name || '';
      case 'delete_sheet': return a.name || '';
      case 'create_table': return a.name || '';
      case 'create_pivot': return a.name || '';
      case 'create_chart': return a.title || a.type || '';
      case 'add_slicer': return a.field || '';
      case 'conditional_format': return `${a.sheet || ''}!${a.range || ''}`;
      case 'autofit': return a.sheet || '';
      case 'insert_rows_cols': return `${a.sheet || ''} ${a.kind || ''} @${a.at || ''}`;
      case 'delete_rows_cols': return `${a.sheet || ''} ${a.kind || ''} @${a.at || ''}`;
      case 'sort_range': return `${a.sheet || ''}!${a.range || ''}`;
      default: return '';
    }
  },

  /**
   * Map a tool name to a high-level phase for the status indicator.
   * Reading tools → 'reading', writing/creating tools → 'writing',
   * read_range after writes → 'verifying' (heuristic: any read_range
   * that follows a write in the same run is treated as verification).
   */
  toolPhase(name) {
    const READ_TOOLS = ['get_workbook_overview', 'read_range', 'find_in_workbook', 'get_objects'];
    const WRITE_TOOLS = ['write_range', 'format_range', 'clear_range', 'add_sheet', 'delete_sheet',
                         'create_table', 'create_pivot', 'create_chart', 'add_slicer',
                         'conditional_format', 'autofit', 'insert_rows_cols', 'delete_rows_cols', 'sort_range'];
    if (WRITE_TOOLS.includes(name)) return 'writing';
    if (READ_TOOLS.includes(name)) return this._hasWrittenThisRun ? 'verifying' : 'reading';
    return 'thinking';
  },

  isWriteTool(name) {
    const WRITE_TOOLS = ['write_range', 'format_range', 'clear_range', 'add_sheet', 'delete_sheet',
                         'create_table', 'create_pivot', 'create_chart', 'add_slicer',
                         'conditional_format', 'autofit', 'insert_rows_cols', 'delete_rows_cols', 'sort_range'];
    return WRITE_TOOLS.includes(name);
  },

  toolPhaseDetail(name, args) {
    const a = args || {};
    switch (name) {
      case 'get_workbook_overview': return '';
      case 'read_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'write_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'format_range': return `${a.sheet || ''}!${a.range || ''}`;
      case 'add_sheet': return a.name || '';
      case 'create_chart': return a.title || a.type || '';
      case 'create_pivot': return a.name || '';
      case 'create_table': return a.name || '';
      default: return this.toolDetail(name, args);
    }
  },

  /* ----------------------------------------------------------
     FINAL MESSAGE
     ---------------------------------------------------------- */
  addFinalMessage(text, sealed, aborted) {
    const msg = document.createElement('div');
    msg.className = 'message assistant';
    const undoBtn = (sealed && !aborted)
      ? `<button type="button" class="btn-undo" data-undo>${this.escapeHtml(I18n.t('undo'))}</button>`
      : '';
    msg.innerHTML = `
      <div class="message-avatar">AI</div>
      <div class="message-content">
        <div class="final-answer">${this.formatMarkdown(text)}</div>
        ${undoBtn}
      </div>
    `;
    const btn = msg.querySelector('[data-undo]');
    if (btn) btn.addEventListener('click', () => this.undoLastRequest(btn));
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
  },

  async undoLastRequest(btn) {
    if (btn) { btn.disabled = true; btn.textContent = I18n.t('undoing'); }
    const res = await Journal.undoLast();
    if (btn) { btn.disabled = false; btn.textContent = I18n.t('undo'); }
    if (res.ok) {
      const note = res.partial
        ? (I18n.lang === 'es'
            ? `Revertido (parcial: ${res.skipped.length} cambio(s) no reversibles).`
            : `Undone (partial: ${res.skipped.length} change(s) not reversible).`)
        : I18n.t('undone');
      this.addMessage('system', note);
    } else {
      this.addMessage('system', `${I18n.t('undoFailed')}: ${res.error || ''}`);
    }
  },

  /* ----------------------------------------------------------
     BASIC MESSAGES
     ---------------------------------------------------------- */
  addMessage(role, text) {
    const msg = document.createElement('div');
    msg.className = `message ${role}`;
    const avatar = role === 'user' ? 'You' : role === 'system' ? '!' : 'AI';
    msg.innerHTML = `
      <div class="message-avatar">${avatar}</div>
      <div class="message-content">${this.formatContent(this.escapeHtml(text))}</div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  addErrorMessageWithRetry(text) {
    const msg = document.createElement('div');
    msg.className = 'message system';
    msg.innerHTML = `
      <div class="message-avatar">!</div>
      <div class="message-content">
        <p>${this.formatContent(this.escapeHtml(text))}</p>
        <button type="button" class="btn-retry">${this.escapeHtml(I18n.t('retryButton'))}</button>
      </div>
    `;
    const btn = msg.querySelector('.btn-retry');
    if (btn) btn.addEventListener('click', () => { btn.disabled = true; this.retryLastMessage(); });
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  /* ----------------------------------------------------------
     STATUS / MISC
     ---------------------------------------------------------- */
  showStatus(text) {
    this.el.statusBar.textContent = text;
    this.el.statusBar.classList.remove('hidden');
  },

  hideStatus() {
    this.el.statusBar.classList.add('hidden');
  },

  clearChat() {
    this.conversation = [];
    Journal.clear();
    Tools.invalidateOverviewCache(); // fresh session — force re-read on next request
    this.el.messageList.innerHTML = '';
    this.addMessage('assistant', I18n.t('clearChatDone'));
  },

  scrollToBottom() {
    this.el.messageList.scrollTop = this.el.messageList.scrollHeight;
  },

  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text == null ? '' : String(text);
    return div.innerHTML;
  },

  formatContent(html) {
    return html.replace(/\n/g, '<br>');
  },

  /**
   * Minimal Markdown rendering for final answers: bold, italic, inline code,
   * bullet lists, numbered lists, and paragraphs. Keeps output safe (escapes
   * HTML first) and simple (no full parser).
   */
  formatMarkdown(text) {
    let s = this.escapeHtml(text);
    // Code spans
    s = s.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Bold
    s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    // Italic (avoid touching ** which we already handled)
    s = s.replace(/(^|[^*])\*([^*]+)\*(?!\*)/g, '$1<em>$2</em>');

    // Split into lines for list/paragraph handling.
    const lines = s.split('\n');
    const out = [];
    let inUl = false, inOl = false;
    const closeLists = () => {
      if (inUl) { out.push('</ul>'); inUl = false; }
      if (inOl) { out.push('</ol>'); inOl = false; }
    };

    for (let raw of lines) {
      const line = raw.trim();
      if (/^[-*]\s+/.test(line)) {
        if (!inUl) { closeLists(); out.push('<ul>'); inUl = true; }
        out.push(`<li>${line.replace(/^[-*]\s+/, '')}</li>`);
      } else if (/^\d+\.\s+/.test(line)) {
        if (!inOl) { closeLists(); out.push('<ol>'); inOl = true; }
        out.push(`<li>${line.replace(/^\d+\.\s+/, '')}</li>`);
      } else if (line === '') {
        closeLists();
        out.push('');
      } else {
        closeLists();
        out.push(`<p>${line}</p>`);
      }
    }
    closeLists();
    return out.join('\n');
  }
};

/* Helper on App for i18n tool labels (defined in i18n.js as toolLabel map). */

document.addEventListener('DOMContentLoaded', () => {
  App.init();
});

