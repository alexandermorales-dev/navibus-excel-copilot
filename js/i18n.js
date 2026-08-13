/* ============================================
   i18n.js — Bilingual support (English / Spanish)
   Detects Office display language and provides
   localized strings for UI, prompts, and labels.
   ============================================ */

const I18n = {
  lang: 'es', // default to Spanish

  // Detect language from Office context
  init() {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.displayLang) {
        const displayLang = Office.context.displayLang.toLowerCase();
        if (displayLang.startsWith('es')) {
          this.lang = 'es';
        } else {
          this.lang = 'en';
        }
      }
    } catch (e) {
      // Default to Spanish if detection fails
      this.lang = 'es';
    }
    console.log(`I18n: language detected as "${this.lang}"`);
  },

  isSpanish() {
    return this.lang === 'es';
  },

  isEnglish() {
    return this.lang === 'en';
  },

  // All UI strings
  strings: {
    // --- Welcome & chat ---
    welcomeTitle: {
      es: '¡Hola! Soy tu Excel AI Copilot. Puedo ayudarte a crear informes, paneles de control, tablas, gráficos y más.',
      en: "Hi! I'm your Excel AI Copilot. I can help you create reports, dashboards, tables, charts, and more."
    },
    welcomeHint: {
      es: 'Ingresa tu API key de Gemini en Configuración para empezar.',
      en: 'Enter your Gemini API key in Settings to get started.'
    },
    placeholder: {
      es: 'Pídeme crear algo... (ej: crea un panel de control)',
      en: 'Ask me to build something... (e.g., create a dashboard)'
    },
    clearChat: {
      es: 'Borrar conversación',
      en: 'Clear Chat History'
    },
    clearChatDone: {
      es: 'Conversación borrada. ¿Qué te gustaría crear?',
      en: 'Chat cleared. What would you like to build?'
    },
    needApiKey: {
      es: 'Por favor, ingresa tu API key de Gemini en Configuración primero.',
      en: 'Please enter your Gemini API key in Settings first.'
    },

    // --- Status messages ---
    analyzing: {
      es: 'Analizando libro...',
      en: 'Analyzing workbook...'
    },
    thinking: {
      es: 'Pensando...',
      en: 'Thinking...'
    },
    building: {
      es: 'Construyendo...',
      en: 'Building...'
    },
    completed: {
      es: 'Completado',
      en: 'Completed'
    },
    reasoning: {
      es: 'Razonando...',
      en: 'Reasoning...'
    },
    reasoningDone: {
      es: 'Razonamiento del modelo',
      en: 'Model reasoning'
    },

    // --- Run summary ---
    summaryTitle: {
      es: 'Listo',
      en: 'Done'
    },
    summarySuccess: {
      es: '✓ Completado',
      en: '✓ Completed'
    },
    summaryRolledBack: {
      es: '✗ Revertido',
      en: '✗ Rolled back'
    },
    summaryPartial: {
      es: '⚠ Parcial',
      en: '⚠ Partial'
    },
    summaryRolledBackTitle: {
      es: 'Revertido',
      en: 'Rolled back'
    },
    summaryPartialTitle: {
      es: 'Parcial',
      en: 'Partial'
    },
    errorsLabel: {
      es: 'Errores',
      en: 'Errors'
    },
    actionCompleted: {
      es: 'Acción completada.',
      en: 'Action completed.'
    },
    newSheet: {
      es: 'Nueva hoja',
      en: 'New sheet'
    },
    kpis: {
      es: (n) => `${n} KPI${n > 1 ? 's' : ''}`,
      en: (n) => `${n} KPI${n > 1 ? 's' : ''}`
    },
    tables: {
      es: (n) => `Tabla${n > 1 ? 's' : ''}`,
      en: (n) => `Table${n > 1 ? 's' : ''}`
    },
    pivots: {
      es: (n) => `Tabla${n > 1 ? 's' : ''} dinámica${n > 1 ? 's' : ''}`,
      en: (n) => `Pivot table${n > 1 ? 's' : ''}`
    },
    charts: {
      es: (n) => `Gráfico${n > 1 ? 's' : ''}`,
      en: (n) => `Chart${n > 1 ? 's' : ''}`
    },
    slicers: {
      es: (n) => `Segmentador${n > 1 ? 'es' : ''}`,
      en: (n) => `Slicer${n > 1 ? 's' : ''}`
    },
    conditionalFormats: {
      es: (n) => `${n} formato${n > 1 ? 's' : ''} condicional${n > 1 ? 'es' : ''}`,
      en: (n) => `${n} conditional format${n > 1 ? 's' : ''}`
    },
    dataBlocks: {
      es: (r, f) => `${r} bloque${r > 1 ? 's' : ''} de datos con ${f} formato${f > 1 ? 's' : ''}`,
      en: (r, f) => `${r} data block${r > 1 ? 's' : ''} with ${f} format${f > 1 ? 's' : ''}`
    },

    // --- Progress step descriptions ---
    opDesc: {
      addSheet: {
        es: (a) => `Creando hoja "${a.name}"`,
        en: (a) => `Creating sheet "${a.name}"`
      },
      writeRange: {
        es: (a) => `Escribiendo datos en ${a.range}`,
        en: (a) => `Writing data to ${a.range}`
      },
      formatRange: {
        es: (a) => `Formateando ${a.range}`,
        en: (a) => `Formatting ${a.range}`
      },
      kpiBlock: {
        es: (a) => `Creando KPI "${a.label}"`,
        en: (a) => `Creating KPI "${a.label}"`
      },
      createTable: {
        es: (a) => `Creando tabla "${a.name}"`,
        en: (a) => `Creating table "${a.name}"`
      },
      createPivot: {
        es: (a) => `Creando tabla dinámica "${a.name}"`,
        en: (a) => `Creating pivot table "${a.name}"`
      },
      createChart: {
        es: (a) => `Creando gráfico${a.title ? ': ' + a.title : ''}`,
        en: (a) => `Creating chart${a.title ? ': ' + a.title : ''}`
      },
      addSlicer: {
        es: (a) => `Creando segmentador "${a.field}"`,
        en: (a) => `Creating slicer "${a.field}"`
      },
      conditionalFormat: {
        es: (a) => `Aplicando formato condicional a ${a.range}`,
        en: (a) => `Applying conditional format to ${a.range}`
      },
      deleteSheet: {
        es: (a) => `Eliminando hoja "${a.name}"`,
        en: (a) => `Deleting sheet "${a.name}"`
      }
    },

    // --- Op labels (for error display) ---
    opLabel: {
      addSheet: { es: 'Crear hoja', en: 'Create sheet' },
      writeRange: { es: 'Escribir datos', en: 'Write data' },
      formatRange: { es: 'Formatear', en: 'Format' },
      kpiBlock: { es: 'KPI', en: 'KPI' },
      createTable: { es: 'Crear tabla', en: 'Create table' },
      createPivot: { es: 'Tabla dinámica', en: 'Pivot table' },
      createChart: { es: 'Gráfico', en: 'Chart' },
      addSlicer: { es: 'Segmentador', en: 'Slicer' },
      conditionalFormat: { es: 'Formato condicional', en: 'Conditional format' },
      deleteSheet: { es: 'Eliminar hoja', en: 'Delete sheet' }
    },

    // --- Error & repair messages ---
    execFailed: {
      es: 'Ejecución fallida. Intentando reparar...',
      en: 'Execution failed. Attempting repair...'
    },
    repairFailed: {
      es: 'Reparación fallida',
      en: 'Repair failed'
    },
    repairAlsoFailed: {
      es: 'La reparación también falló. No hay más intentos.',
      en: 'Repair also failed. No further attempts.'
    },
    validationFailed: {
      es: 'Validación fallida',
      en: 'Validation failed'
    },
    geminiError: {
      es: 'Error de Gemini',
      en: 'Gemini error'
    },
    workbookError: {
      es: 'No se pudo leer el libro',
      en: 'Could not read workbook'
    },
    genericError: {
      es: 'Error',
      en: 'Error'
    },

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
    timeout: {
      es: 'Tiempo de espera agotado (120s)',
      en: 'Request timed out (120s)'
    },
    networkError: {
      es: (s) => `Error de red: ${s}`,
      en: (s) => `Network error: ${s}`
    },
    unknownError: {
      es: 'Error desconocido tras reintentos',
      en: 'Unknown error after retries'
    },
    fallbackModel: {
      es: 'Modelo principal falló, probando alternativa...',
      en: 'Primary model failed, trying fallback...'
    },

    // --- Summary & suggestions prompt ---
    summaryPrompt: {
      es: (planSummary, schemaText) => `Resumen de lo que se creó: ${planSummary}

Datos del libro: ${schemaText}

Responde en ESPAÑOL con un mensaje breve y profesional que incluya:
1. Un resumen de 1-2 líneas de lo que se creó (en lenguaje natural, no técnico)
2. 2-3 sugerencias específicas de análisis adicionales o mejoras que el usuario podría pedir, basadas en los datos reales del libro

Mantén el mensaje conciso (máximo 5-6 líneas). No devuelvas JSON. Solo texto plano en español.`,
      en: (planSummary, schemaText) => `Summary of what was created: ${planSummary}

Workbook data: ${schemaText}

Respond in ENGLISH with a brief, professional message that includes:
1. A 1-2 line summary of what was created (in natural language, not technical)
2. 2-3 specific suggestions for additional analysis or improvements the user could request, based on the actual workbook data

Keep the message concise (maximum 5-6 lines). Do not return JSON. Plain text only in English.`
    }
  },

  // Get a string by key (for simple strings)
  t(key) {
    const entry = this.strings[key];
    if (!entry) return key;
    if (typeof entry === 'function') return entry;
    return entry[this.lang] || entry.en || key;
  },

  // Get a string by key with arguments (for function strings)
  tf(key, ...args) {
    const entry = this.strings[key];
    if (!entry) return key;
    const fn = entry[this.lang] || entry.en;
    if (typeof fn === 'function') return fn(...args);
    return fn;
  },

  // Get an op-specific string
  tOp(opKey, subKey, action) {
    const entry = this.strings[opKey];
    if (!entry || !entry[subKey]) return opKey;
    const fn = entry[subKey][this.lang] || entry[subKey].en;
    if (typeof fn === 'function') return fn(action);
    return fn;
  }
};
