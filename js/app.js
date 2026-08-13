/* ============================================
   app.js â€” Main UI controller + agent loop
   ============================================ */

const App = {
  conversation: [],   // { role, parts: [{text}] } â€” Gemini format
  isRunning: false,

  // DOM refs
  el: {},

  init() {
    // Cache DOM elements
    this.el = {
      messageList: document.getElementById('messageList'),
      messageInput: document.getElementById('messageInput'),
      sendBtn: document.getElementById('sendBtn'),
      settingsBtn: document.getElementById('settingsBtn'),
      settingsPanel: document.getElementById('settingsPanel'),
      apiKeyInput: document.getElementById('apiKeyInput'),
      modelSelect: document.getElementById('modelSelect'),
      clearChatBtn: document.getElementById('clearChatBtn'),
      statusBar: document.getElementById('statusBar')
    };

    // Load saved settings
    Config.load();
    this.el.apiKeyInput.value = Config.apiKey;
    this.el.modelSelect.value = Config.model;
    this.updateSendButton();

    // Event listeners
    this.el.sendBtn.addEventListener('click', () => this.sendMessage());
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

    this.el.modelSelect.addEventListener('change', () => {
      Config.model = this.el.modelSelect.value;
      Config.save();
    });

    this.el.clearChatBtn.addEventListener('click', () => {
      this.clearChat();
    });

    // Initialize Office.js
    if (typeof Office !== 'undefined') {
      Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
          console.log('Excel AI Copilot ready');
        }
      });
    } else {
      // Running outside Office (browser dev mode)
      console.warn('Office.js not loaded â€” running in browser dev mode');
    }
  },

  updateSendButton() {
    const hasText = this.el.messageInput.value.trim().length > 0;
    const hasKey = Config.hasApiKey();
    this.el.sendBtn.disabled = !hasText || !hasKey || this.isRunning;
  },

  async sendMessage() {
    const text = this.el.messageInput.value.trim();
    if (!text || this.isRunning) return;

    if (!Config.hasApiKey()) {
      this.addMessage('system', 'Por favor, ingresa tu API key de Gemini en ConfiguraciÃ³n primero.');
      this.el.settingsPanel.classList.remove('hidden');
      return;
    }

    this.isRunning = true;
    this.updateSendButton();
    this.el.messageInput.value = '';

    // Show user message
    this.addMessage('user', text);

    // Add to conversation history
    this.conversation.push({
      role: 'user',
      parts: [{ text: text }]
    });

    // Show typing indicator
    const typingEl = this.addTypingIndicator();

    try {
      await this.runAgentLoop(text, typingEl);
    } catch (e) {
      this.removeElement(typingEl);
      this.addMessage('system', `Error: ${e.message || String(e)}`);
    }

    this.isRunning = false;
    this.updateSendButton();
  },

  async runAgentLoop(userText, typingEl) {
    // 1. Read workbook schema
    this.updateTypingText(typingEl, 'Analizando libro...');
    let schemaSnap;
    try {
      schemaSnap = await Schema.snapshot();
    } catch (e) {
      this.removeElement(typingEl);
      this.addMessage('system', `No se pudo leer el libro: ${e.message}`);
      return;
    }

    // 2. Build system prompt
    const systemPrompt = Prompt.build(schemaSnap);

    // 3. Call Gemini with live thinking display
    this.updateTypingText(typingEl, 'Pensando...');

    // Replace typing indicator with a live thinking block
    this.removeElement(typingEl);
    const thinkingEl = this.addLiveThinkingMessage();

    const result = await Gemini.generateWithFallback(
      systemPrompt,
      this.conversation,
      (chunk, fullThinking) => {
        // Live update: append chunk to thinking display
        this.updateLiveThinking(thinkingEl, fullThinking);
      }
    );

    // Finalize the thinking block
    this.finalizeLiveThinking(thinkingEl);

    if (!result.ok) {
      this.addMessage('system', `Error de Gemini: ${result.error}`);
      return;
    }

    // 4. Parse the response
    const parsed = Repair.parsePlan(result.text);
    if (!parsed.ok) {
      // Not a JSON plan â€” treat as a text response
      this.addMessage('assistant', result.text);
      this.conversation.push({ role: 'model', parts: [{ text: result.text }] });
      return;
    }

    // 6. Validate the plan
    const validation = Actions.validatePlan(parsed.plan);
    if (!validation.valid) {
      this.addMessage('system', `ValidaciÃ³n fallida:\n${validation.errors.join('\n')}`);
      await this.attemptRepair(systemPrompt, validation.errors, schemaSnap);
      return;
    }

    // 7. Execute the plan with live progress
    this.conversation.push({ role: 'model', parts: [{ text: result.text }] });
    const progressEl = this.addProgressMessage(parsed.plan);

    const execResult = await Executor.execute(parsed.plan, (current, total, action) => {
      this.updateProgress(progressEl, current, total, action);
    });

    this.finalizeProgress(progressEl);

    // 8. Show results
    this.showRunSummary(execResult);

    if (execResult.failed.length > 0 && execResult.rolledBack) {
      this.addMessage('system', 'EjecuciÃ³n fallida. Intentando reparar...');
      await this.attemptRepair(systemPrompt, execResult.failed, schemaSnap, parsed.plan);
    } else if (execResult.failed.length === 0) {
      // 9. Generate a final summary + suggestions message
      await this.generateSummaryAndSuggestions(parsed.plan, execResult, schemaSnap);
    }
  },

  async generateSummaryAndSuggestions(plan, execResult, schemaSnap) {
    // Build a short prompt asking for a summary of what was done + suggestions
    const planSummary = plan.map(a => {
      switch (a.op) {
        case 'addSheet': return `CreÃ³ hoja "${a.name}"`;
        case 'kpiBlock': return `KPI "${a.label}"`;
        case 'createChart': return `GrÃ¡fico "${a.title || a.type}"`;
        case 'createPivot': return `Tabla dinÃ¡mica "${a.name}"`;
        case 'createTable': return `Tabla "${a.name}"`;
        case 'addSlicer': return `Segmentador "${a.field}"`;
        case 'conditionalFormat': return `Formato condicional en ${a.range}`;
        case 'writeRange': return `Datos en ${a.range}`;
        case 'formatRange': return `Formato en ${a.range}`;
        default: return a.op;
      }
    }).join(', ');

    const summaryPrompt = `Resumen de lo que se creÃ³: ${planSummary}

Datos del libro: ${Schema.toText(schemaSnap)}

Responde en ESPAÃ‘OL con un mensaje breve y profesional que incluya:
1. Un resumen de 1-2 lÃ­neas de lo que se creÃ³ (en lenguaje natural, no tÃ©cnico)
2. 2-3 sugerencias especÃ­ficas de anÃ¡lisis adicionales o mejoras que el usuario podrÃ­a pedir, basadas en los datos reales del libro

MantÃ©n el mensaje conciso (mÃ¡ximo 5-6 lÃ­neas). No devuelvas JSON. Solo texto plano en espaÃ±ol.`;

    const result = await Gemini.generateWithFallback(
      summaryPrompt,
      [{ role: 'user', parts: [{ text: summaryPrompt }] }],
      null  // no thinking callback for this short call
    );

    if (result.ok && result.text) {
      this.addMessage('assistant', result.text);
      this.conversation.push({ role: 'model', parts: [{ text: result.text }] });
    }
  },

  async attemptRepair(systemPrompt, failedActions, schemaSnap, originalPlan) {
    const repairResult = await Repair.repairActions(
      systemPrompt,
      this.conversation,
      failedActions,
      schemaSnap,
      originalPlan
    );

    if (!repairResult.ok) {
      this.addMessage('system', `ReparaciÃ³n fallida: ${repairResult.error}`);
      return;
    }

    // Execute repaired plan
    const execResult = await Executor.execute(repairResult.plan);
    this.showRunSummary(execResult);

    if (execResult.failed.length > 0) {
      this.addMessage('system', 'La reparaciÃ³n tambiÃ©n fallÃ³. No hay mÃ¡s intentos.');
    }
  },

  showRunSummary(result) {
    // Build a user-friendly summary in Spanish
    const summary = this.buildFriendlySummary(result);

    const card = document.createElement('div');
    card.className = 'message assistant';
    card.innerHTML = `
      <div class="message-avatar">N</div>
      <div class="message-content">
        <div class="run-card">
          <div class="run-card-title">
            ${summary.title}
            <span class="badge ${summary.badgeClass}">${summary.badgeText}</span>
          </div>
          <div class="run-card-body">${summary.body}</div>
          ${summary.errors.length > 0 ? `
            <div class="run-card-errors">
              <strong>Errores:</strong>
              <ul class="run-card-list">
                ${summary.errors.map(e => `<li><span class="icon-fail">âœ—</span> ${e}</li>`).join('')}
              </ul>
            </div>
          ` : ''}
        </div>
      </div>
    `;
    this.el.messageList.appendChild(card);
    this.scrollToBottom();
  },

  buildFriendlySummary(result) {
    const succeeded = result.succeeded;
    const failed = result.failed;

    // Aggregate what was created
    const created = {
      sheets: [],
      kpis: 0,
      tables: [],
      pivots: [],
      charts: [],
      slicers: [],
      ranges: 0,
      formats: 0,
      conditionalFormats: 0,
    };

    for (const s of succeeded) {
      const a = s.action;
      switch (a.op) {
        case 'addSheet': created.sheets.push(a.name); break;
        case 'kpiBlock': created.kpis++; break;
        case 'createTable': created.tables.push(a.name); break;
        case 'createPivot': created.pivots.push(a.name); break;
        case 'createChart': created.charts.push(a.title || a.type || 'grÃ¡fico'); break;
        case 'addSlicer': created.slicers.push(a.field || 'slicer'); break;
        case 'writeRange': created.ranges++; break;
        case 'formatRange': created.formats++; break;
        case 'conditionalFormat': created.conditionalFormats++; break;
      }
    }

    // Build friendly body text
    const parts = [];
    if (created.sheets.length > 0) {
      parts.push(`Nueva hoja: <strong>${created.sheets.join(', ')}</strong>`);
    }
    if (created.kpis > 0) {
      parts.push(`${created.kpis} KPI${created.kpis > 1 ? 's' : ''}`);
    }
    if (created.tables.length > 0) {
      parts.push(`Tabla${created.tables.length > 1 ? 's' : ''}: ${created.tables.join(', ')}`);
    }
    if (created.pivots.length > 0) {
      parts.push(`Tabla${created.pivots.length > 1 ? 's' : ''} dinÃ¡mica${created.pivots.length > 1 ? 's' : ''}: ${created.pivots.join(', ')}`);
    }
    if (created.charts.length > 0) {
      parts.push(`GrÃ¡fico${created.charts.length > 1 ? 's' : ''}: ${created.charts.join(', ')}`);
    }
    if (created.slicers.length > 0) {
      parts.push(`Segmentador${created.slicers.length > 1 ? 'es' : ''}: ${created.slicers.join(', ')}`);
    }
    if (created.conditionalFormats > 0) {
      parts.push(`${created.conditionalFormats} formato${created.conditionalFormats > 1 ? 's' : ''} condicional${created.conditionalFormats > 1 ? 'es' : ''}`);
    }
    if (created.ranges > 0 && created.formats > 0) {
      parts.push(`${created.ranges} bloque${created.ranges > 1 ? 's' : ''} de datos con ${created.formats} formato${created.formats > 1 ? 's' : ''}`);
    }

    const body = parts.length > 0
      ? `<p>${parts.join(' Â· ')}</p>`
      : '<p>AcciÃ³n completada.</p>';

    // Determine title and badge
    let title, badgeText, badgeClass;
    if (failed.length === 0) {
      title = 'Listo';
      badgeText = 'âœ“ Completado';
      badgeClass = 'success';
    } else if (result.rolledBack) {
      title = 'Revertido';
      badgeText = 'âœ— Revertido';
      badgeClass = 'error';
    } else {
      title = 'Parcial';
      badgeText = 'âš  Parcial';
      badgeClass = 'partial';
    }

    const errors = failed.map(f => `${this.opLabel(f.action?.op)}: ${f.error}`);

    return { title, badgeText, badgeClass, body, errors };
  },

  opLabel(op) {
    const labels = {
      addSheet: 'Crear hoja',
      writeRange: 'Escribir datos',
      formatRange: 'Formatear',
      kpiBlock: 'KPI',
      createTable: 'Crear tabla',
      createPivot: 'Tabla dinÃ¡mica',
      createChart: 'GrÃ¡fico',
      addSlicer: 'Segmentador',
      conditionalFormat: 'Formato condicional',
      deleteSheet: 'Eliminar hoja',
    };
    return labels[op] || op || 'OperaciÃ³n';
  },

  // --- UI helpers ---

  addMessage(role, text) {
    const msg = document.createElement('div');
    msg.className = `message ${role}`;

    const avatar = role === 'user' ? 'Fr' : role === 'system' ? '!' : 'N';
    const content = this.escapeHtml(text);

    msg.innerHTML = `
      <div class="message-avatar">${avatar}</div>
      <div class="message-content">${this.formatContent(content)}</div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  addTypingIndicator() {
    const msg = document.createElement('div');
    msg.className = 'message assistant';
    msg.innerHTML = `
      <div class="message-avatar">N</div>
      <div class="message-content">
        <div class="typing-status" id="typingText">Procesando...</div>
        <div class="typing-indicator">
          <span></span><span></span><span></span>
        </div>
      </div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  updateTypingText(typingEl, text) {
    if (!typingEl) return;
    const statusEl = typingEl.querySelector('.typing-status');
    if (statusEl) statusEl.textContent = text;
  },

  addLiveThinkingMessage() {
    const msg = document.createElement('div');
    msg.className = 'message assistant';
    msg.innerHTML = `
      <div class="message-avatar">N</div>
      <div class="message-content">
        <div class="thinking-block thinking-live">
          <div class="thinking-header">
            <span class="thinking-icon">ðŸ’­</span>
            <span class="thinking-label">Razonando...</span>
            <span class="thinking-dots">
              <span></span><span></span><span></span>
            </span>
          </div>
          <div class="thinking-text thinking-stream"></div>
        </div>
      </div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  updateLiveThinking(thinkingEl, fullText) {
    if (!thinkingEl) return;
    const streamEl = thinkingEl.querySelector('.thinking-stream');
    if (streamEl) {
      streamEl.innerHTML = this.formatContent(this.escapeHtml(fullText));
    }
    this.scrollToBottom();
  },

  finalizeLiveThinking(thinkingEl) {
    if (!thinkingEl) return;
    const header = thinkingEl.querySelector('.thinking-header');
    if (header) {
      header.innerHTML = '<span class="thinking-icon">ðŸ’­</span><span class="thinking-label">Razonamiento del modelo</span>';
    }
    // Make it collapsible after completion
    const block = thinkingEl.querySelector('.thinking-block');
    if (block) {
      block.classList.remove('thinking-live');
      block.classList.add('thinking-done');
    }
  },

  addProgressMessage(plan) {
    const msg = document.createElement('div');
    msg.className = 'message assistant';
    msg.innerHTML = `
      <div class="message-avatar">N</div>
      <div class="message-content">
        <div class="progress-container">
          <div class="progress-header">
            <span class="progress-label">Ejecutando plan...</span>
            <span class="progress-count">0 / ${plan.length}</span>
          </div>
          <div class="progress-bar-track">
            <div class="progress-bar-fill" style="width: 0%"></div>
          </div>
          <div class="progress-step"></div>
        </div>
      </div>
    `;
    this.el.messageList.appendChild(msg);
    this.scrollToBottom();
    return msg;
  },

  updateProgress(progressEl, current, total, action) {
    if (!progressEl) return;
    const pct = Math.round((current / total) * 100);
    const fill = progressEl.querySelector('.progress-bar-fill');
    const count = progressEl.querySelector('.progress-count');
    const step = progressEl.querySelector('.progress-step');
    const label = progressEl.querySelector('.progress-label');

    if (fill) fill.style.width = pct + '%';
    if (count) count.textContent = `${current} / ${total}`;
    if (label) label.textContent = 'Construyendo...';
    if (step) {
      const desc = this.opDescription(action);
      step.textContent = desc;
    }
    this.scrollToBottom();
  },

  finalizeProgress(progressEl) {
    if (!progressEl) return;
    const label = progressEl.querySelector('.progress-label');
    if (label) label.textContent = 'Completado';
    const fill = progressEl.querySelector('.progress-bar-fill');
    if (fill) fill.style.width = '100%';
  },

  opDescription(action) {
    const labels = {
      addSheet: `Creando hoja "${action.name}"`,
      writeRange: `Escribiendo datos en ${action.range}`,
      formatRange: `Formateando ${action.range}`,
      kpiBlock: `Creando KPI "${action.label}"`,
      createTable: `Creando tabla "${action.name}"`,
      createPivot: `Creando tabla dinÃ¡mica "${action.name}"`,
      createChart: `Creando grÃ¡fico${action.title ? ': ' + action.title : ''}`,
      addSlicer: `Creando segmentador "${action.field}"`,
      conditionalFormat: `Aplicando formato condicional a ${action.range}`,
      deleteSheet: `Eliminando hoja "${action.name}"`,
    };
    return labels[action.op] || `Ejecutando ${action.op}`;
  },

  removeElement(el) {
    if (el && el.parentNode) {
      el.parentNode.removeChild(el);
    }
  },

  showStatus(text) {
    this.el.statusBar.textContent = text;
    this.el.statusBar.classList.remove('hidden');
  },

  hideStatus() {
    this.el.statusBar.classList.add('hidden');
  },

  clearChat() {
    this.conversation = [];
    this.el.messageList.innerHTML = '';
    this.addMessage('assistant', 'ConversaciÃ³n borrada. Â¿QuÃ© te gustarÃ­a analizar, Francisco?');
  },

  scrollToBottom() {
    this.el.messageList.scrollTop = this.el.messageList.scrollHeight;
  },

  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  },

  formatContent(html) {
    // Convert newlines to <br>, preserve code blocks
    return html.replace(/\n/g, '<br>');
  }
};

// System prompt builder
const Prompt = {
  build(schemaSnap) {
    const schemaText = Schema.toText(schemaSnap);
    return `Eres Navibus AI, el asistente inteligente de Navibus, una empresa de ferrys/transporte marÃ­timo. EstÃ¡s integrado en Excel como panel lateral para ayudar a Francisco, el gerente, a crear informes, paneles de control, tablas, grÃ¡ficos y tablas dinÃ¡micas sobre la operaciÃ³n de la flota.

IDENTIDAD: Eres Navibus AI. Te diriges al usuario como Francisco. Conoces el negocio de transporte marÃ­timo: viajes, rutas, embarcaciones, costos operativos, combustible, mantenimiento, ocupaciÃ³n, horarios, tripulaciÃ³n. Tu tono es profesional pero cercano, como un analista de confianza que conoce el negocio.

IDIOMA: EspaÃ±ol. Eres un asistente de habla hispana. TODO tu output debe estar en espaÃ±ol â€” incluyendo tu razonamiento interno (thinking/reasoning). Piensa en espaÃ±ol, razona en espaÃ±ol, responde en espaÃ±ol. NUNCA uses inglÃ©s para nada. Si piensas en inglÃ©s, traduce tu razonamiento al espaÃ±ol antes de emitirlo.

TerminologÃ­a profesional del sector marÃ­timo en espaÃ±ol: "Panel de Control", "Informe Ejecutivo", "Flota", "EmbarcaciÃ³n", "Viajes", "Rutas", "OcupaciÃ³n", "Costos Operativos", "Combustible", "Mantenimiento", "TripulaciÃ³n", "Temporada Alta/Baja", "Tabla DinÃ¡mica", "Segmentador de Datos", "Formato Condicional".

ESTADO ACTUAL DEL LIBRO:
${schemaText}

## TU ROL: ASESOR PROACTIVO

Eres un asesor profesional, no solo un ejecutor de comandos. Analiza los datos del libro y actÃºa como un gerente financiero/tÃ©cnico:

- Cuando veas los datos, IDENTIFICA patrones, tendencias y oportunidades de anÃ¡lisis.
- Sugiere mÃ©tricas relevantes segÃºn el tipo de datos (ej: totales, promedios, tendencias mensuales, comparativas).
- Si el usuario pide algo genÃ©rico como "crea un dashboard", PROPÃ“N un diseÃ±o especÃ­fico basado en los datos reales que ves.
- DespuÃ©s de crear algo, SUGIERE 2-3 anÃ¡lisis adicionales o mejoras que podrÃ­an ser Ãºtiles.
- Usa un tono profesional pero cercano, como un consultor experto.

## CUÃNDO RESPONDER TEXTO vs JSON

- Si el usuario saluda, hace una pregunta, o NO pide crear algo: responde con TEXTO PLANO en espaÃ±ol. NO devuelvas JSON.
- Si el usuario pide CREAR, GENERAR, HACER algo (panel, informe, tabla, grÃ¡fico, KPI, etc.): responde con un ARRAY JSON de acciones.
- Antes de crear algo complejo, puedes hacer una pregunta aclaratoria breve en texto plano si es necesario.
- DespuÃ©s de ejecutar un plan, la prÃ³xima respuesta del usuario puede incluir seguimiento â€” mantÃ©n contexto de lo que se creÃ³.

## FORMATO DEL PLAN DE ACCIONES

Cuando devuelvas un plan, DEBE ser un array JSON que empieza con [ y termina con ]. Sin markdown, sin texto antes o despuÃ©s del JSON.

Cada acciÃ³n tiene un campo "op" y parÃ¡metros especÃ­ficos:

[{"op": "addSheet", "name": "Panel_Control"},
 {"op": "writeRange", "sheet": "Panel_Control", "range": "A1:H2", "values": [["TÃ­tulo",""],["SubtÃ­tulo",""]]},
 {"op": "formatRange", "sheet": "Panel_Control", "range": "A1:H1", "bold": true, "fontSize": 16, "fillColor": "#1a237e", "fontColor": "#FFFFFF", "horizontalAlignment": "Left"},
 {"op": "kpiBlock", "sheet": "Panel_Control", "cell": "B3", "label": "Total Viajes", "formula": "=COUNTA(Datos!A2:A500)"},
 {"op": "kpiBlock", "sheet": "Panel_Control", "cell": "E3", "label": "Costo Promedio", "formula": "=AVERAGE(Datos!D2:D500)", "numberFormat": "#,##0.00"},
 {"op": "createTable", "sheet": "Datos", "range": "A1:D100", "name": "TablaDatos", "style": "TableStyleMedium2"},
 {"op": "createPivot", "sheet": "Panel_Control", "source": "Datos!A1:M500", "name": "PivotVentas", "dest": "A10", "rows": ["Region"], "values": [{"col": "Importe", "agg": "sum"}]},
 {"op": "createChart", "sheet": "Panel_Control", "type": "columnClustered", "sourceRange": "A10:B20", "dest": "E10", "title": "Ventas por RegiÃ³n", "width": 400, "height": 250},
 {"op": "addSlicer", "sheet": "Panel_Control", "sourcePivot": "PivotVentas", "field": "Mes", "dest": "E2"},
 {"op": "conditionalFormat", "sheet": "Panel_Control", "range": "B3:B10", "type": "dataBar"},
 {"op": "deleteSheet", "name": "HojaVieja"}]

## PROPIEDADES DE FORMATO (para formatRange)
bold (bool), italic (bool), fontSize (number), fontName (string), fillColor (hex como "#1a237e"),
fontColor (hex), horizontalAlignment ("Left"|"Center"|"Right"), verticalAlignment ("Top"|"Center"|"Bottom"),
numberFormat (string como "#,##0" o "â‚¬#,##0.00" o "0.0%"), wrapText (bool),
columnWidth (number â€” ANCHO DE COLUMNA EN PIXELES, usar 15-20 para texto corto, 25-35 para texto largo, 12-15 para nÃºmeros),
rowHeight (number â€” ALTO DE FILA, usar 20-25 para normal, 30-40 para encabezados),
borders (string como "Thin" o objeto como {"EdgeTop":"Thin","EdgeBottom":"Thin"})

## TIPOS DE GRÃFICO
"columnClustered", "columnStacked", "barClustered", "barStacked", "line", "pie", "doughnut", "area"

## AGREGACIONES (para tablas dinÃ¡micas)
"sum", "count", "average", "max", "min"

## KPI BLOCKS
- Usa "formula" para valores calculados (ej: "=SUM(Datos!C2:C500)", "=COUNTA(Datos!A2:A500)").
- Usa "value" para valores estÃ¡ticos (ej: 42, "N/A").
- Opcionalmente incluye "numberFormat" (ej: "#,##0", "â‚¬#,##0.00", "0.0%").

## REGLAS DE LAYOUT PROFESIONAL (CRÃTICO)

1. ESTRUCTURA DEL PANEL (de arriba a abajo):
   - Fila 1: TÃ­tulo del panel (combinar A1:H1, fontSize 16, fondo azul oscuro, texto blanco)
   - Fila 2: Espacio en blanco (rowHeight 10)
   - Filas 3-5: KPIs en fila horizontal (un KPI cada 3 columnas: B3, E3, H3)
   - Fila 6-7: Espacio en blanco
   - Fila 8+: Tabla de datos o tabla dinÃ¡mica (columnas A-D)
   - Fila 8+: GrÃ¡ficos a la derecha (columnas F-H o mÃ¡s allÃ¡)
   - Fila 25+: SecciÃ³n de observaciones/insights (texto con anÃ¡lisis)

2. ANCHOS DE COLUMNA â€” SIEMPRE incluir formatRange con columnWidth:
   - Columna A: 20-25 (etiquetas/nombres)
   - Columnas B-D: 15-18 (datos numÃ©ricos)
   - Columnas E-H: 15-18 (datos adicionales o espacio para grÃ¡ficos)
   - NUNCA dejar columnas con ancho por defecto â€” siempre especificar columnWidth

3. ALTOS DE FILA â€” SIEMPRE incluir formatRange con rowHeight:
   - Fila de tÃ­tulo: 35-40
   - Filas de KPI: 25-30
   - Filas de encabezado de tabla: 25
   - Filas de datos: 20-22
   - Filas de espacio: 10-15

4. ESPACIADO â€” NUNCA superponer elementos:
   - Dejar al menos 2 filas en blanco entre secciones
   - Los grÃ¡ficos necesitan espacio: especificar width (300-450) y height (200-280)
   - Posicionar grÃ¡ficos con dest en celdas que tengan espacio libre debajo
   - KPIs: separar horizontalmente por al menos 2 columnas

5. COLORES PROFESIONALES:
   - TÃ­tulo: fondo "#1a237e" (azul oscuro), texto "#FFFFFF"
   - Encabezados de tabla: fondo "#3949ab" (azul medio), texto "#FFFFFF", bold
   - KPIs: etiqueta en gris "#666666", valor en azul "#1a73e8" o verde "#2e7d32"
   - Filas alternas: usar fillColor "#f5f5f5" para filas pares (efecto cebra)
   - Bordes de tabla: "Thin" color "#cccccc"

## REGLAS ESTRICTAS
1. NUNCA escribas o modifiques hojas de datos existentes. SIEMPRE crea una hoja nueva primero con addSheet.
2. Pon todos los elementos del panel en UNA hoja nueva.
3. Usa nombres reales de hojas y columnas del libro. NO inventes nombres de columnas.
4. Las fÃ³rmulas deben referenciar la hoja de datos real (ej: =SUM(Datos!C2:C500)).
5. Para tablas dinÃ¡micas, "source" debe ser un rango completo como "Hoja!A1:M500" incluyendo encabezados.
6. Para KPIs, la etiqueta va en la fila encima del valor (el sistema lo maneja automÃ¡ticamente).
7. MANTÃ‰N LOS PLANES CONCISOS: prefiere menos acciones bien estructuradas.
8. Devuelve SOLO el array JSON. Sin markdown, sin explicaciÃ³n fuera del JSON.
9. DespuÃ©s de crear algo, la prÃ³xima respuesta puede incluir sugerencias de anÃ¡lisis adicionales.`;
  }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});
