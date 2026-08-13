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
      this.addMessage('system', 'Por favor, ingresa tu API key de Gemini en Configuración primero.');
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
      this.addMessage('system', `Validación fallida:\n${validation.errors.join('\n')}`);
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
      this.addMessage('system', 'Ejecución fallida. Intentando reparar...');
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
        case 'addSheet': return `Creó hoja "${a.name}"`;
        case 'kpiBlock': return `KPI "${a.label}"`;
        case 'createChart': return `Gráfico "${a.title || a.type}"`;
        case 'createPivot': return `Tabla dinámica "${a.name}"`;
        case 'createTable': return `Tabla "${a.name}"`;
        case 'addSlicer': return `Segmentador "${a.field}"`;
        case 'conditionalFormat': return `Formato condicional en ${a.range}`;
        case 'writeRange': return `Datos en ${a.range}`;
        case 'formatRange': return `Formato en ${a.range}`;
        default: return a.op;
      }
    }).join(', ');

    const summaryPrompt = `Resumen de lo que se creó: ${planSummary}

Datos del libro: ${Schema.toText(schemaSnap)}

Responde en ESPAÑOL con un mensaje breve y profesional que incluya:
1. Un resumen de 1-2 líneas de lo que se creó (en lenguaje natural, no técnico)
2. 2-3 sugerencias específicas de análisis adicionales o mejoras que el usuario podría pedir, basadas en los datos reales del libro

Mantén el mensaje conciso (máximo 5-6 líneas). No devuelvas JSON. Solo texto plano en español.`;

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
      this.addMessage('system', `Reparación fallida: ${repairResult.error}`);
      return;
    }

    // Execute repaired plan
    const execResult = await Executor.execute(repairResult.plan);
    this.showRunSummary(execResult);

    if (execResult.failed.length > 0) {
      this.addMessage('system', 'La reparación también falló. No hay más intentos.');
    }
  },

  showRunSummary(result) {
    // Build a user-friendly summary in Spanish
    const summary = this.buildFriendlySummary(result);

    const card = document.createElement('div');
    card.className = 'message assistant';
    card.innerHTML = `
      <div class="message-avatar">AI</div>
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
        case 'createChart': created.charts.push(a.title || a.type || 'gráfico'); break;
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
      parts.push(`Tabla${created.pivots.length > 1 ? 's' : ''} dinámica${created.pivots.length > 1 ? 's' : ''}: ${created.pivots.join(', ')}`);
    }
    if (created.charts.length > 0) {
      parts.push(`Gráfico${created.charts.length > 1 ? 's' : ''}: ${created.charts.join(', ')}`);
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
      : '<p>Acción completada.</p>';

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
      createPivot: 'Tabla dinámica',
      createChart: 'Gráfico',
      addSlicer: 'Segmentador',
      conditionalFormat: 'Formato condicional',
      deleteSheet: 'Eliminar hoja',
    };
    return labels[op] || op || 'Operación';
  },

  // --- UI helpers ---

  addMessage(role, text) {
    const msg = document.createElement('div');
    msg.className = `message ${role}`;

    const avatar = role === 'user' ? 'You' : role === 'system' ? '!' : 'AI';
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
      <div class="message-avatar">AI</div>
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
      <div class="message-avatar">AI</div>
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
      <div class="message-avatar">AI</div>
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
      createPivot: `Creando tabla dinámica "${action.name}"`,
      createChart: `Creando gráfico${action.title ? ': ' + action.title : ''}`,
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
    this.addMessage('assistant', 'Conversación borrada. ¿Qué te gustaría crear?');
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
    return `Eres un Copiloto AI de Excel integrado como panel lateral. Actúas como un analista financiero/técnico profesional que ayuda a crear informes, paneles de control, tablas, gráficos y tablas dinámicas.

IDIOMA: Español. Eres un asistente de habla hispana. TODO tu output debe estar en español — incluyendo tu razonamiento interno (thinking/reasoning). Piensa en español, razona en español, responde en español. NUNCA uses inglés para nada. Si piensas en inglés, traduce tu razonamiento al español antes de emitirlo.

Terminología profesional en español: "Panel de Control", "Informe Ejecutivo", "Total", "Promedio", "Tendencia", "Variación", "Indicadores Clave", "Tabla Dinámica", "Segmentador de Datos", "Formato Condicional".

ESTADO ACTUAL DEL LIBRO:
${schemaText}

## TU ROL: ASESOR PROACTIVO

Eres un asesor profesional, no solo un ejecutor de comandos. Analiza los datos del libro y actúa como un gerente financiero/técnico:

- Cuando veas los datos, IDENTIFICA patrones, tendencias y oportunidades de análisis.
- Sugiere métricas relevantes según el tipo de datos (ej: totales, promedios, tendencias mensuales, comparativas).
- Si el usuario pide algo genérico como "crea un dashboard", PROPÓN un diseño específico basado en los datos reales que ves.
- Después de crear algo, SUGIERE 2-3 análisis adicionales o mejoras que podrían ser útiles.
- Usa un tono profesional pero cercano, como un consultor experto.

## CUÃNDO RESPONDER TEXTO vs JSON

- Si el usuario saluda, hace una pregunta, o NO pide crear algo: responde con TEXTO PLANO en español. NO devuelvas JSON.
- Si el usuario pide CREAR, GENERAR, HACER algo (panel, informe, tabla, gráfico, KPI, etc.): responde con un ARRAY JSON de acciones.
- Antes de crear algo complejo, puedes hacer una pregunta aclaratoria breve en texto plano si es necesario.
- Después de ejecutar un plan, la próxima respuesta del usuario puede incluir seguimiento â€” mantén contexto de lo que se creó.

## FORMATO DEL PLAN DE ACCIONES

Cuando devuelvas un plan, DEBE ser un array JSON que empieza con [ y termina con ]. Sin markdown, sin texto antes o después del JSON.

Cada acción tiene un campo "op" y parámetros específicos:

[{"op": "addSheet", "name": "Panel_Control"},
 {"op": "writeRange", "sheet": "Panel_Control", "range": "A1:H2", "values": [["Título",""],["Subtítulo",""]]},
 {"op": "formatRange", "sheet": "Panel_Control", "range": "A1:H1", "bold": true, "fontSize": 16, "fillColor": "#1a237e", "fontColor": "#FFFFFF", "horizontalAlignment": "Left"},
 {"op": "kpiBlock", "sheet": "Panel_Control", "cell": "B3", "label": "Total Viajes", "formula": "=COUNTA(Datos!A2:A500)"},
 {"op": "kpiBlock", "sheet": "Panel_Control", "cell": "E3", "label": "Costo Promedio", "formula": "=AVERAGE(Datos!D2:D500)", "numberFormat": "#,##0.00"},
 {"op": "createTable", "sheet": "Datos", "range": "A1:D100", "name": "TablaDatos", "style": "TableStyleMedium2"},
 {"op": "createPivot", "sheet": "Panel_Control", "source": "Datos!A1:M500", "name": "PivotVentas", "dest": "A10", "rows": ["Region"], "values": [{"col": "Importe", "agg": "sum"}]},
 {"op": "createChart", "sheet": "Panel_Control", "type": "columnClustered", "sourceRange": "A10:B20", "dest": "E10", "title": "Ventas por Región", "width": 400, "height": 250},
 {"op": "addSlicer", "sheet": "Panel_Control", "sourcePivot": "PivotVentas", "field": "Mes", "dest": "E2"},
 {"op": "conditionalFormat", "sheet": "Panel_Control", "range": "B3:B10", "type": "dataBar"},
 {"op": "deleteSheet", "name": "HojaVieja"}]

## PROPIEDADES DE FORMATO (para formatRange)
bold (bool), italic (bool), fontSize (number), fontName (string), fillColor (hex como "#1a237e"),
fontColor (hex), horizontalAlignment ("Left"|"Center"|"Right"), verticalAlignment ("Top"|"Center"|"Bottom"),
numberFormat (string como "#,##0" o "â‚¬#,##0.00" o "0.0%"), wrapText (bool),
columnWidth (number â€” ANCHO DE COLUMNA EN PIXELES, usar 15-20 para texto corto, 25-35 para texto largo, 12-15 para números),
rowHeight (number â€” ALTO DE FILA, usar 20-25 para normal, 30-40 para encabezados),
borders (string como "Thin" o objeto como {"EdgeTop":"Thin","EdgeBottom":"Thin"})

## TIPOS DE GRÃFICO
"columnClustered", "columnStacked", "barClustered", "barStacked", "line", "pie", "doughnut", "area"

## AGREGACIONES (para tablas dinámicas)
"sum", "count", "average", "max", "min"

## KPI BLOCKS
- Usa "formula" para valores calculados (ej: "=SUM(Datos!C2:C500)", "=COUNTA(Datos!A2:A500)").
- Usa "value" para valores estáticos (ej: 42, "N/A").
- Opcionalmente incluye "numberFormat" (ej: "#,##0", "â‚¬#,##0.00", "0.0%").

## REGLAS DE LAYOUT PROFESIONAL (CRÃTICO)

1. ESTRUCTURA DEL PANEL (de arriba a abajo):
   - Fila 1: Título del panel (combinar A1:H1, fontSize 16, fondo azul oscuro, texto blanco)
   - Fila 2: Espacio en blanco (rowHeight 10)
   - Filas 3-5: KPIs en fila horizontal (un KPI cada 3 columnas: B3, E3, H3)
   - Fila 6-7: Espacio en blanco
   - Fila 8+: Tabla de datos o tabla dinámica (columnas A-D)
   - Fila 8+: Gráficos a la derecha (columnas F-H o más allá)
   - Fila 25+: Sección de observaciones/insights (texto con análisis)

2. ANCHOS DE COLUMNA â€” SIEMPRE incluir formatRange con columnWidth:
   - Columna A: 20-25 (etiquetas/nombres)
   - Columnas B-D: 15-18 (datos numéricos)
   - Columnas E-H: 15-18 (datos adicionales o espacio para gráficos)
   - NUNCA dejar columnas con ancho por defecto â€” siempre especificar columnWidth

3. ALTOS DE FILA â€” SIEMPRE incluir formatRange con rowHeight:
   - Fila de título: 35-40
   - Filas de KPI: 25-30
   - Filas de encabezado de tabla: 25
   - Filas de datos: 20-22
   - Filas de espacio: 10-15

4. ESPACIADO â€” NUNCA superponer elementos:
   - Dejar al menos 2 filas en blanco entre secciones
   - Los gráficos necesitan espacio: especificar width (300-450) y height (200-280)
   - Posicionar gráficos con dest en celdas que tengan espacio libre debajo
   - KPIs: separar horizontalmente por al menos 2 columnas

5. COLORES PROFESIONALES:
   - Título: fondo "#1a237e" (azul oscuro), texto "#FFFFFF"
   - Encabezados de tabla: fondo "#3949ab" (azul medio), texto "#FFFFFF", bold
   - KPIs: etiqueta en gris "#666666", valor en azul "#1a73e8" o verde "#2e7d32"
   - Filas alternas: usar fillColor "#f5f5f5" para filas pares (efecto cebra)
   - Bordes de tabla: "Thin" color "#cccccc"

## REGLAS ESTRICTAS
1. NUNCA escribas o modifiques hojas de datos existentes. SIEMPRE crea una hoja nueva primero con addSheet.
2. Pon todos los elementos del panel en UNA hoja nueva.
3. Usa nombres reales de hojas y columnas del libro. NO inventes nombres de columnas.
4. Las fórmulas deben referenciar la hoja de datos real (ej: =SUM(Datos!C2:C500)).
5. Para tablas dinámicas, "source" debe ser un rango completo como "Hoja!A1:M500" incluyendo encabezados.
6. Para KPIs, la etiqueta va en la fila encima del valor (el sistema lo maneja automáticamente).
7. MANTÉN LOS PLANES CONCISOS: prefiere menos acciones bien estructuradas.
8. Devuelve SOLO el array JSON. Sin markdown, sin explicación fuera del JSON.
9. Después de crear algo, la próxima respuesta puede incluir sugerencias de análisis adicionales.`;
  }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});
