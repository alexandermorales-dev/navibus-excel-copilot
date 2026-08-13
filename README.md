# Excel AI Copilot

An intelligent Excel add-in that uses Google's Gemini API to create reports, dashboards, tables, charts, and pivot tables — without generating or injecting VBA code.

## How it works

1. You type a request in the sidebar (e.g., "crea un panel de control con KPIs").
2. The add-in reads the workbook structure (sheet names, headers, data types).
3. Gemini returns a **JSON action plan** (validated, not raw code).
4. A deterministic executor performs each action via Office.js — no compile errors, ever.
5. If an action fails, it rolls back all changes and attempts one automatic repair round.
6. On success, the assistant generates a summary in Spanish with suggestions for follow-up analysis.

## Features

- **Live reasoning stream**: Watch the model's thinking appear in real-time as it analyzes your data.
- **Progress bar**: Real-time tracking of each action as it executes.
- **Executive summary**: On completion, shows what was created and suggests next analyses.
- **Professional layout**: Automatic column widths, row heights, and element spacing — no overlapping data.
- **Auto-repair**: If an action fails, rolls back everything and retries with error context.
- **Dark theme**: Modern dark UI with blue accents.
- **Quota handling**: Detects rate limits (429) and daily quota exhaustion with clear messages.
- **Spanish-first**: All UI text, responses, reasoning, and suggestions are in Spanish.

## Requirements

- **Microsoft Excel 2021+ or Microsoft 365** (desktop, Windows). Requires WebView2 + Office.js pivot/slicer APIs.
- **A free Gemini API key** — get one at [aistudio.google.com/apikey](https://aistudio.google.com/apikey).
- For local development: Python 3.x (used as the HTTPS dev server).

## Local development setup

### 1. Start the dev server

```bash
python dev_server.py
```

The server runs on `https://localhost:3000` with self-signed SSL certificates (required because Office.js demands HTTPS).

> First time? Run `.certs/install-cert.ps1` to install the self-signed certificate into the Windows certificate store.

### 2. Sideload the manifest in Excel

1. Open Excel.
2. Go to **Insert** → **Add-ins** → **Manage My Add-ins** → **Upload My Add-in**.
3. Select `manifest-localhost.xml`.
4. The **AI Copilot** button appears on the Home tab.

### 3. Enter your API key

1. Click the **gear icon** in the sidebar header.
2. Paste your Gemini API key.
3. Select a model (default: `gemini-3.6-flash`).
4. Close settings and start chatting.

## Sharing with other users

Once hosted on GitHub Pages:

1. Push this project to a GitHub repository.
2. Enable **GitHub Pages** (Settings → Pages → Source: main branch, root).
3. Edit `manifest.xml` — replace `YOUR_GITHUB_USERNAME` with your GitHub username and `YOUR_REPO_NAME` with your repo name.
4. Share `manifest.xml` with the other person.
5. They open Excel → **Insert** → **Add-ins** → **Upload My Add-in** → select the file.
6. They enter their own Gemini API key in settings.

No installer, no Node.js, no build step on their end.

## Project structure

```
├── manifest.xml              # For GitHub Pages (sharing)
├── manifest-localhost.xml    # For local development
├── index.html                # Task pane entry point
├── css/
│   └── style.css             # Dark-themed UI
├── js/
│   ├── config.js             # Settings (API key, model) in localStorage
│   ├── gemini.js             # Gemini API client (streaming, 429, fallback)
│   ├── schema.js             # Workbook snapshot via Office.js
│   ├── actions.js            # Action DSL validation
│   ├── executor.js           # Deterministic Office.js executor + rollback + auto-fit
│   ├── repair.js             # Failed-action repair round-trip
│   └── app.js                # UI controller + agent loop + system prompt
├── assets/
│   └── icon-*.png            # Add-in icons
├── dev_server.py             # HTTPS development server
└── README.md
```

## Action DSL

The add-in uses a JSON action plan instead of generated code. Available operations:

| Op | Description |
|---|---|
| `addSheet` | Create a new worksheet |
| `writeRange` | Write values to a range |
| `formatRange` | Apply formatting (bold, fill, borders, column width, etc.) |
| `kpiBlock` | Create a labeled KPI cell with a formula or static value |
| `createTable` | Create an Excel Table from a range |
| `createPivot` | Create a PivotTable |
| `createChart` | Create a chart from a range or pivot |
| `addSlicer` | Add a slicer to a pivot or table |
| `conditionalFormat` | Add conditional formatting (data bars, color scale, etc.) |
| `deleteSheet` | Delete a worksheet |

Each action is validated before execution. If validation or execution fails, the add-in rolls back any changes made and sends the errors back to Gemini for one repair attempt.

## Gemini API notes

- Default model: `gemini-3.6-flash` (free tier).
- Fallback model: `gemini-3.5-flash-lite` (separate quota).
- Handles 429 rate limits by parsing Google's `retryDelay` and waiting with a visible countdown.
- Per-day quota exhaustion fails fast with a clear message.
- API key is stored in `localStorage` and only sent to Google's API.
- Model reasoning is streamed live via `streamGenerateContent` with `includeThoughts: true`.

## Limitations

- Requires Excel 2021+ (Office.js pivot/slicer APIs).
- Office.js bypasses Excel's undo stack (mitigated by artifact-log rollback).
- The DSL supports a defined set of operations — new ops can be added by extending `actions.js` and `executor.js`.
- Calculated pivot fields and classic pivot layout are not supported by the Office.js API.
- The model's internal reasoning language may occasionally be in English despite Spanish instructions (model limitation).

## Troubleshooting

**"Por favor, ingresa tu API key"** — Click the gear icon and paste your key from [aistudio.google.com/apikey](https://aistudio.google.com/apikey).

**Add-in doesn't appear** — Ensure you uploaded the correct manifest (`manifest-localhost.xml` for local dev). Check that the server is running on port 3000.

**Sidebar shows old code after updates** — Office's cache is aggressive. Close Excel and delete the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, then reopen Excel and reload the add-in.

**Pivot/slicer errors** — Confirm you're running Excel 2021+ or Microsoft 365. Older versions don't support the required Office.js APIs.

**Rate limited (429)** — The free tier has per-minute and per-day limits. Wait a minute for per-minute limits. For per-day limits, try again after midnight Pacific time.
