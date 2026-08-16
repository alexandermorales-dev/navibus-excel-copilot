/* ============================================
   i18n.js — Bilingual support (English / Spanish)
   Detects Office display language and provides localized strings
   for UI, tool labels, and status messages.
   ============================================ */

const I18n = {
  lang: 'es', // default to Spanish

  init() {
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
      es: 'Ingresa tu API key de Gemini en Configuración para empezar.',
      en: 'Enter your Gemini API key in Settings to get started.'
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
      es: 'Por favor, ingresa tu API key de Gemini en Configuración primero.',
      en: 'Please enter your Gemini API key in Settings first.'
    },

    // --- Status / live activity ---
    analyzing: { es: 'Analizando libro...', en: 'Analyzing workbook...' },
    thinking:  { es: 'Pensando...',       en: 'Thinking...' },
    building:  { es: 'Construyendo...',   en: 'Building...' },
    completed: { es: 'Completado',        en: 'Completed' },
    reasoning: { es: 'Razonando...',      en: 'Reasoning...' },
    reasoningDone: { es: 'Razonamiento del modelo', en: 'Model reasoning' },
    retryingMessage: { es: 'Reintentando...', en: 'Retrying...' },

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
      sort_range:            { es: 'Ordenar rango',        en: 'Sort range' }
    },

    // --- Errors ---
    geminiError:   { es: 'Error de Gemini',   en: 'Gemini error' },
    workbookError: { es: 'No se pudo leer el libro', en: 'Could not read workbook' },
    genericError:  { es: 'Error',             en: 'Error' },
    retryButton:   { es: 'Reintentar',        en: 'Retry' },

    // --- Gemini API status messages ---
    dailyQuota: {
      es: 'Cuota diaria de Gemini agotada. Se reinicia a medianoche (hora del Pacífico).',
      en: 'Daily Gemini quota exhausted. Resets at midnight Pacific time.'
    },
    rateLimit: {
      es: (s) => `Límite de tasa — esperando ${s}s...`,
      en: (s) => `Rate limited — waiting ${s}s...`
    },
    serverError: {
      es: (s) => `Error del servidor ${s}`,
      en: (s) => `Server error ${s}`
    },
    timeout: { es: 'Tiempo de espera agotado (180s)', en: 'Request timed out (180s)' },
    networkError: {
      es: (s) => `Error de red: ${s}`,
      en: (s) => `Network error: ${s}`
    },
    unknownError: { es: 'Error desconocido tras reintentos', en: 'Unknown error after retries' },
    fallbackModel: {
      es: 'Modelo principal falló, probando alternativa...',
      en: 'Primary model failed, trying fallback...'
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
