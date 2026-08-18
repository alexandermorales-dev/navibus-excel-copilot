/* ============================================
   app.js — UI controller for the agentic Excel Copilot
   Owns: chat DOM, settings, conversation history, and the
   Agent.run() lifecycle (start/stop/undo).
   No more plan/repair/summary logic — the agent handles all of that.
   ============================================ */

const App = {
  conversation: [],   // Gemini contents format
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
      apiKeyInput: document.getElementById('apiKeyInput'),
      apiKeyToggle: document.getElementById('apiKeyToggle'),
      modelSelect: document.getElementById('modelSelect'),
      clearChatBtn: document.getElementById('clearChatBtn'),
      statusBar: document.getElementById('statusBar')
    };

    Config.load();
    this.el.apiKeyInput.value = Config.apiKey;
    this.el.modelSelect.value = Config.model;
    this.updateSendButton();
    this.updateStopButton();

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

    this.el.apiKeyInput.addEventListener('change', () => {
      Config.apiKey = this.el.apiKeyInput.value.trim();
      Config.save();
      this.updateSendButton();
    });

    this.el.apiKeyToggle.addEventListener('click', () => {
      const input = this.el.apiKeyInput;
      const eyeIcon = this.el.apiKeyToggle.querySelector('.icon-eye');
      const eyeOffIcon = this.el.apiKeyToggle.querySelector('.icon-eye-off');
      if (input.type === 'password') {
        input.type = 'text';
        eyeIcon.classList.add('hidden');
        eyeOffIcon.classList.remove('hidden');
      } else {
        input.type = 'password';
        eyeIcon.classList.remove('hidden');
        eyeOffIcon.classList.add('hidden');
      }
    });

    this.el.modelSelect.addEventListener('change', () => {
      Config.model = this.el.modelSelect.value;
      Config.save();
    });

    this.el.clearChatBtn.addEventListener('click', () => this.clearChat());

    if (typeof Office !== 'undefined' && Office.onReady) {
      Office.onReady((info) => {
        try {
          if (info && info.host === Office.HostType.Excel) {
            console.log('Excel AI Copilot ready');
          }
        } catch (e) { /* info or HostType may be undefined outside Excel */ }
        I18n.init();
        this.localizeUI();
      });
      // Fallback: if Office.onReady doesn't fire within 3s (e.g. outside
      // Excel), initialize i18n anyway so the UI isn't stuck in default lang.
      setTimeout(() => {
        if (!I18n.initialized) { I18n.init(); this.localizeUI(); }
      }, 3000);
    } else {
      console.warn('Office.js not loaded — running in browser dev mode');
      I18n.init();
      this.localizeUI();
    }
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

    this.addMessage('user', text);
    this.conversation.push({ role: 'user', parts: [{ text }] });

    this.abortController = new AbortController();

    // Create the live activity container for this run.
    const activityEl = this.addActivityContainer();

    try {
      const result = await Agent.run({
        userText: text,
        conversation: this.conversation,
        signal: this.abortController.signal,
        onThinking: (chunk, full) => this.updateLiveThinking(activityEl, full),
        onText: (chunk, full) => this.updateLiveAnswer(activityEl, full),
        onToolStart: (callId, name, args) => this.addToolRow(activityEl, callId, name, args),
        onToolEnd: (callId, name, toolResult) => this.finalizeToolRow(activityEl, callId, toolResult),
        onToolError: (name, error) => this.showToolError(activityEl, name, error)
      });

      this.finalizeLiveThinking(activityEl);
      this.finalizeLiveAnswer(activityEl);

      if (!result.ok) {
        this.addErrorMessageWithRetry(`${I18n.t('geminiError')}: ${result.error}`);
      } else {
        // Render the final answer as a proper message (with undo if sealed).
        this.addFinalMessage(result.finalText, result.sealed, result.aborted);
      }
    } catch (e) {
      this.finalizeLiveThinking(activityEl);
      this.addErrorMessageWithRetry(`${I18n.t('genericError')}: ${e.message || String(e)}`);
    }

    this.abortController = null;
    this.isRunning = false;
    this.updateSendButton();
    this.updateStopButton();
  },

  stopRun() {
    if (this.abortController) this.abortController.abort();
  },

  async retryLastMessage() {
    if (this.isRunning || !this.lastUserText) return;
    this.isRunning = true;
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
        onThinking: (chunk, full) => this.updateLiveThinking(activityEl, full),
        onText: (chunk, full) => this.updateLiveAnswer(activityEl, full),
        onToolStart: (callId, name, args) => this.addToolRow(activityEl, callId, name, args),
        onToolEnd: (callId, name, toolResult) => this.finalizeToolRow(activityEl, callId, toolResult),
        onToolError: (name, error) => this.showToolError(activityEl, name, error)
      });
      this.finalizeLiveThinking(activityEl);
      this.finalizeLiveAnswer(activityEl);
      if (!result.ok) {
        this.addErrorMessageWithRetry(`${I18n.t('geminiError')}: ${result.error}`);
      } else {
        this.addFinalMessage(result.finalText, result.sealed, result.aborted);
      }
    } catch (e) {
      this.finalizeLiveThinking(activityEl);
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
    // Collapse long thinking by default; user can expand.
    const streamEl = activityEl.querySelector('.thinking-stream');
    if (streamEl && streamEl.textContent.trim().length > 0) {
      streamEl.classList.add('collapsed');
      header.style.cursor = 'pointer';
      header.onclick = () => streamEl.classList.toggle('collapsed');
    }
  },

  updateLiveAnswer(activityEl, fullText) {
    if (!activityEl || !fullText) return;
    const ansEl = activityEl.querySelector('.live-answer');
    if (ansEl) {
      ansEl.classList.remove('hidden');
      ansEl.innerHTML = this.formatContent(this.escapeHtml(fullText));
    }
    this.scrollToBottom();
  },

  finalizeLiveAnswer(activityEl) {
    if (!activityEl) return;
    const ansEl = activityEl.querySelector('.live-answer');
    if (ansEl && ansEl.classList.contains('hidden')) {
      // No live answer was streamed (e.g. agent ended via tool round with no text).
      // The final message will be added separately; hide this placeholder.
      ansEl.remove();
    } else if (ansEl) {
      // Promote the live answer to a finalized styled block.
      ansEl.classList.add('live-answer-done');
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
      spinner.classList.add(toolResult.ok ? 'tool-ok' : 'tool-fail');
      spinner.textContent = toolResult.ok ? '\u2713' : '\u2717';
    }
    if (!toolResult.ok) {
      row.classList.add('tool-row-error');
      const errSpan = document.createElement('span');
      errSpan.className = 'tool-error';
      errSpan.textContent = toolResult.error;
      row.appendChild(errSpan);
    }
    this.scrollToBottom();
  },

  /**
   * Show a prominent error banner when a tool fails, so the user is
   * always aware that something went wrong — even if the model later
   * self-corrects or omits the error from its final answer.
   */
  showToolError(activityEl, toolName, error) {
    if (!activityEl) return;
    const feed = activityEl.querySelector('.activity-feed');
    if (!feed) return;
    // Avoid duplicate banners for the same tool+error.
    const existing = feed.querySelector('.tool-error-banner');
    if (existing && existing.dataset.error === error) return;
    const banner = document.createElement('div');
    banner.className = 'tool-error-banner';
    banner.dataset.error = error;
    const label = I18n.toolLabel(toolName);
    banner.innerHTML = `<span class="tool-error-icon">\u26A0</span><span><strong>${this.escapeHtml(label)}</strong>: ${this.escapeHtml(error)}</span>`;
    feed.appendChild(banner);
    this.scrollToBottom();
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

