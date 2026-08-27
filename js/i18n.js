/* ============================================
   i18n.js — Bilingual support (English / Spanish)
   Detects Office display language and provides localized strings
   for UI, tool labels, and status messages.
   ============================================ */

const I18n = {
  lang: 'es', // default to Spanish
  initialized: false,

  init() {
    if (this.initialized) return; // guard against double-init
    this.initialized = true;
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.displayLang) {
        const displayLang = Office.context.displayLang.toLowerCase();
        this.lang = displayLang.startsWith('es') ? 'es' : 'en';
      }
    } catch (e) {
      this.lang = 'es';
    }
    console.log(`I18n: language detected as "${this.lang}"`);
  },

  isSpanish() { return this.lang === 'es'; },
  isEnglish() { return this.lang === 'en'; },

  strings: {
    // --- Welcome & chat ---
    welcomeTitle: {
      es: '¡Hola! Soy tu Excel AI Copilot. Puedo ayudarte a crear informes, paneles de control, tablas, gráficos y más — y responder preguntas sobre tus datos.',
      en: "Hi! I'm your Excel AI Copilot. I can build reports, dashboards, tables, charts and more — and answer questions about your data."
    },
    welcomeHint: {
      es: 'Para empezar, agrega al menos una API key gratuita en Configuración.',
      en: 'To get started, add at least one free API key in Settings.'
    },
    placeholder: {
      es: 'Pregúntame sobre tus datos o pídeme crear algo... (ej: crea un panel de control)',
      en: 'Ask about your data or ask me to build something... (e.g., create a dashboard)'
    },
    clearChat: {
      es: 'Borrar conversación',
      en: 'Clear Chat History'
    },
    clearChatDone: {
      es: 'Conversación borrada. ¿Qué te gustaría hacer?',
      en: 'Chat cleared. What would you like to do?'
    },
    needApiKey: {
      es: 'No hay API key configurada. Agrega al menos una key gratuita en Configuración.',
      en: 'No API key configured. Add at least one free key in Settings.'
    },

    // --- Onboarding / provider settings ---
    providersTitle: { es: 'Proveedores de IA', en: 'AI providers' },
    providersHint: {
      es: 'Agrega al menos una key. Con varias, el copiloto reparte las peticiones entre ellas y cambia de proveedor automáticamente al agotarse una cuota.',
      en: 'Add at least one key. With several, the copilot spreads requests across them and switches provider automatically when a quota runs out.'
    },
    getKey:        { es: 'Obtener key gratis', en: 'Get free key' },
    testKey:       { es: 'Probar',    en: 'Test' },
    testing:       { es: 'Probando...', en: 'Testing...' },
    testOk:        { es: (m) => `Conectado — modelo: ${m}`, en: (m) => `Connected — model: ${m}` },
    testFailed:    { es: (e) => `Falló: ${e}`, en: (e) => `Failed: ${e}` },
    noKeySet:      { es: 'Sin configurar', en: 'Not configured' },
    advancedTitle: { es: 'Avanzado', en: 'Advanced' },
    modelAuto:     { es: 'Automático (recomendado)', en: 'Automatic (recommended)' },
    modelOverride: { es: 'Modelo fijo', en: 'Pinned model' },
    resetQuota:    { es: 'Reiniciar contadores de cuota', en: 'Reset quota counters' },
    quotaReset:    { es: 'Contadores de cuota reiniciados.', en: 'Quota counters reset.' },

    // --- Quota bar ---
    quotaTitle:     { es: 'Peticiones restantes hoy', en: 'Requests left today' },
    quotaExhausted: { es: 'agotado', en: 'spent' },
    quotaInvalid:   { es: 'key inválida', en: 'invalid key' },
    quotaCooldown:  { es: 'en espera', en: 'cooling down' },
    allQuotaSpent: {
      es: 'Se agotó la cuota gratuita de todos los proveedores configurados. Agrega otra key en Configuración o vuelve a intentarlo mañana.',
      en: 'The free quota of every configured provider is spent. Add another key in Settings or try again tomorrow.'
    },
    switchedProvider: {
      es: (p) => `Cambiado a ${p} (el proveedor anterior estaba limitado)`,
      en: (p) => `Switched to ${p} (previous provider was throttled)`
    },
    usingProvider: { es: (p, m) => `${p} · ${m}`, en: (p, m) => `${p} · ${m}` },

    // --- Status / live activity ---
    analyzing: { es: 'Analizando libro...', en: 'Analyzing workbook...' },
    thinking:  { es: 'Pensando...',       en: 'Thinking...' },
    building:  { es: 'Construyendo...',   en: 'Building...' },
    completed: { es: 'Completado',        en: 'Completed' },
    reasoning: { es: 'Razonando...',      en: 'Reasoning...' },
    reasoningDone: { es: 'Razonamiento del modelo', en: 'Model reasoning' },
    retryingMessage: { es: 'Reintentando...', en: 'Retrying...' },
    // --- Run status header (phase indicator) ---
    statusThinking:  { es: 'Pensando...',           en: 'Thinking...' },
    statusReading:   { es: 'Leyendo el libro...',   en: 'Reading workbook...' },
    statusPlanning:  { es: 'Planificando...',       en: 'Planning...' },
    statusWriting:   { es: 'Escribiendo en Excel...', en: 'Writing to Excel...' },
    statusVerifying: { es: 'Verificando...',        en: 'Verifying...' },
    statusRepairing: { es: 'Corrigiendo...',        en: 'Repairing...' },
    statusDone:      { es: 'Completado',            en: 'Completed' },
    statusError:     { es: 'Error',                 en: 'Error' },
    statusStopped:   { es: 'Detenido',              en: 'Stopped' },

    // --- Plan checklist ---
    planLabel:   { es: 'Plan', en: 'Plan' },
    planSteps:   { es: (n) => `${n} paso(s)`, en: (n) => `${n} step(s)` },
    callsUsed:   { es: (n) => `${n} llamada(s) a la IA`, en: (n) => `${n} AI call(s)` },
    repairing:   { es: (n) => `Corrigiendo ${n} problema(s)...`, en: (n) => `Repairing ${n} issue(s)...` },
    verifiedOk:  { es: 'Verificado sin errores', en: 'Verified, no errors' },
    verifyIssues: { es: (n) => `${n} problema(s) tras la verificación`, en: (n) => `${n} issue(s) after verification` },

    // --- Stop / Undo ---
    stop:    { es: 'Detener',   en: 'Stop' },
    undo:    { es: 'Deshacer',  en: 'Undo' },
    undoing: { es: 'Deshaciendo...', en: 'Undoing...' },
    undone:  { es: 'Cambios revertidos.', en: 'Changes undone.' },
    undoFailed: { es: 'No se pudo deshacer', en: 'Could not undo' },
    aborted: { es: 'Detenido por el usuario.', en: 'Stopped by user.' },

    // --- Tool labels (shown in the activity feed) ---
    toolLabel: {
      get_workbook_overview: { es: 'Analizar libro',       en: 'Analyze workbook' },
      read_range:            { es: 'Leer rango',           en: 'Read range' },
      find_in_workbook:      { es: 'Buscar en libro',      en: 'Find in workbook' },
      get_objects:           { es: 'Listar objetos',       en: 'List objects' },
      write_range:           { es: 'Escribir datos',       en: 'Write data' },
      format_range:          { es: 'Formatear',            en: 'Format' },
      clear_range:           { es: 'Limpiar rango',        en: 'Clear range' },
      add_sheet:             { es: 'Crear hoja',           en: 'Create sheet' },
      delete_sheet:          { es: 'Eliminar hoja',        en: 'Delete sheet' },
      create_table:          { es: 'Crear tabla',          en: 'Create table' },
      create_pivot:          { es: 'Crear tabla dinámica', en: 'Create pivot table' },
      create_chart:          { es: 'Crear gráfico',        en: 'Create chart' },
      add_slicer:            { es: 'Crear segmentador',    en: 'Create slicer' },
      conditional_format:    { es: 'Formato condicional',  en: 'Conditional format' },
      autofit:               { es: 'Auto-ajustar',         en: 'Auto-fit' },
      insert_rows_cols:      { es: 'Insertar filas/cols',  en: 'Insert rows/cols' },
      delete_rows_cols:      { es: 'Eliminar filas/cols',  en: 'Delete rows/cols' },
      sort_range:            { es: 'Ordenar rango',        en: 'Sort range' },
      freeze_panes:          { es: 'Inmovilizar paneles',  en: 'Freeze panes' },
      'recipe.dashboard':    { es: 'Panel de control',     en: 'Dashboard' },
      'recipe.summary_table':{ es: 'Tabla resumen',        en: 'Summary table' },
      'recipe.kpi_row':      { es: 'Tarjetas de KPI',      en: 'KPI cards' },
      'recipe.title_banner': { es: 'Banner de título',     en: 'Title banner' }
    },

    // --- Errors ---
    aiError:       { es: 'Error de IA',     en: 'AI error' },
    workbookError: { es: 'No se pudo leer el libro', en: 'Could not read workbook' },
    genericError:  { es: 'Error',             en: 'Error' },
    retryButton:   { es: 'Reintentar',        en: 'Retry' },

    // --- Provider / transport status messages ---
    rateLimit: {
      es: (s) => `Límite de tasa — esperando ${s}s...`,
      en: (s) => `Rate limited — waiting ${s}s...`
    },
    serverError: {
      es: (s) => `Error del servidor ${s}`,
      en: (s) => `Server error ${s}`
    },
    authFailed: {
      es: (p) => `${p}: la API key es inválida o no está autorizada. Verifícala en Configuración.`,
      en: (p) => `${p}: API key is invalid or unauthorized. Check it in Settings.`
    },
    providerQuota: {
      es: (p) => `${p}: cuota agotada.`,
      en: (p) => `${p}: quota exhausted.`
    },
    providerRateLimited: {
      es: (p) => `${p}: límite de tasa alcanzado.`,
      en: (p) => `${p}: rate limit reached.`
    },
    timeout: { es: 'Tiempo de espera agotado (180s)', en: 'Request timed out (180s)' },
    networkError: {
      es: (s) => `Error de red: ${s}`,
      en: (s) => `Network error: ${s}`
    },
    unknownError: { es: 'Error desconocido tras reintentos', en: 'Unknown error after retries' },
    badPlan: {
      es: 'El modelo no devolvió un plan válido. Intenta reformular tu solicitud.',
      en: 'The model did not return a valid plan. Try rephrasing your request.'
    },
    emptyWorkbook: {
      es: 'El libro parece estar vacío. Agrega tus datos en Excel e inténtalo de nuevo.',
      en: 'The workbook looks empty. Add your data in Excel and try again.'
    },
    emptyResponse: {
      es: 'El modelo devolvió una respuesta vacía. Reintentando automáticamente...',
      en: 'The model returned an empty response. Retrying automatically...'
    },
    truncatedPlan: {
      es: 'La respuesta fue demasiado larga y se truncó. Usa recetas (recipe.dashboard, recipe.summary_table) en lugar de escribir datos celda por celda. Las recetas generan el layout automáticamente.',
      en: 'The response was too long and got truncated. Use recipes (recipe.dashboard, recipe.summary_table) instead of writing data cell by cell. Recipes generate the layout automatically.'
    },
    retryNudge1: {
      es: 'Tu respuesta anterior estaba vacía o no era JSON válido. Devuelve SOLO un objeto JSON con los campos "answer" y "ops". No incluyas texto adicional fuera del JSON. Si no puedes cumplir la solicitud, explica por qué en "answer" y deja "ops" vacío.',
      en: 'Your previous response was empty or not valid JSON. Return ONLY a JSON object with "answer" and "ops" fields. No text outside the JSON. If you cannot fulfill the request, explain why in "answer" and leave "ops" empty.'
    },
    retryNudge2: {
      es: 'Intento 2. Devuelve un objeto JSON válido: {"intent":"...","answer":"texto breve","ops":[...]}. El campo "answer" debe contener un mensaje real para el usuario. No repitas el esquema. Usa recipe.dashboard o recipe.summary_table para dashboards y tablas — NO escribas los datos celda por celda.',
      en: 'Attempt 2. Return a valid JSON object: {"intent":"...","answer":"short message","ops":[...]}. The "answer" field must contain a real message for the user. Do not repeat the schema. Use recipe.dashboard or recipe.summary_table for dashboards and tables — do NOT write data cell by cell.'
    },
    retryNudge3: {
      es: 'Último intento. Responde con un JSON mínimo: {"answer":"explica brevemente qué pasó","ops":[]}. Si no puedes hacer lo que pide el usuario, dilo en "answer".',
      en: 'Final attempt. Reply with minimal JSON: {"answer":"briefly explain what happened","ops":[]}. If you cannot do what the user asked, say so in "answer".'
    }
  },

  t(key) {
    const entry = this.strings[key];
    if (!entry) return key;
    if (typeof entry === 'function') return entry;
    return entry[this.lang] || entry.en || key;
  },

  tf(key, ...args) {
    const entry = this.strings[key];
    if (!entry) return key;
    const fn = entry[this.lang] || entry.en;
    if (typeof fn === 'function') return fn(...args);
    return fn;
  },

  // Look up a tool label by tool name (used in the activity feed).
  toolLabel(name) {
    const labels = this.strings.toolLabel || {};
    const entry = labels[name];
    if (!entry) return name;
    return entry[this.lang] || entry.en || name;
  }
};
