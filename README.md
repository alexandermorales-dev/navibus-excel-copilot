# Excel AI Copilot

An intelligent Excel add-in that uses Google's Gemini API to **answer questions about your workbook** and **build reports, dashboards, tables, charts, and pivots** — via an iterative, tool-calling agent that reads the real workbook, acts step-by-step, verifies its own results, and self-corrects.

## How it works

The add-in runs an **agentic loop** (no more single-shot "JSON plan"):

1. You type a request in the sidebar (e.g. "crea un panel de control con KPIs" or "¿cuál es el total de ventas por región?").
2. The agent calls Gemini with a set of **tools** (functions) it can invoke.
3. Gemini decides what to do: it calls tools like `get_workbook_overview`, `read_range`, `add_sheet`, `write_range`, `create_chart`, etc.
4. Each tool runs against the real workbook via Office.js and returns verified results (including formula errors like `#REF!`) back to Gemini.
5. Gemini sees the real results, self-corrects if something failed, and continues until the task is done.
6. It finishes with a concise summary in your language, plus suggested follow-up analyses.
7. Every mutation is journaled, so a **Undo** button on the final message reverts the whole request.

## Features

- **Agentic tool-calling loop** — Gemini reads, writes, and verifies step-by-step instead of guessing a plan up front.
- **Live reasoning stream** — watch the model's thinking appear in real time.
- **Live tool activity feed** — each tool call shows as a row with a spinner → ✓/✗ and a short description.
- **Self-correction** — when a tool returns an error, the model sees it and retries with a fix (no more full rollbacks).
- **Stop button** — abort a run mid-flight.
- **Undo last request** — reverts all mutations from the most recent request (per-request journal).
- **Edit existing sheets** — the agent can add formula columns, sort, clean, and format your data when asked (with undo).
- **Q&A about your data** — the agent reads exact cells with `read_range` to answer accurately, no more guessing from samples.
- **Markdown rendering** — final answers support bold, lists, and inline code.
- **Dark theme, ES/EN auto-detect** — UI and model output follow Office's display language.
- **Quota handling** — 429 rate limits show a visible countdown; daily quota fails fast with a clear message.

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
│   └── style.css             # Dark-themed UI + activity feed + markdown
├── js/
│   ├── config.js             # Settings (API key, model) in localStorage
│   ├── i18n.js               # ES/EN string table + Office locale detection
│   ├── gemini.js             # Gemini API client (function calling, streaming, 429, fallback)
│   ├── schema.js             # Workbook snapshot via Office.js (powers get_workbook_overview / read_range)
│   ├── journal.js            # Per-request mutation journal → Undo
│   ├── tools.js              # Tool declarations + Office.js implementations (read/write/edit)
│   ├── prompt.js             # Agentic system prompts (ES/EN)
│   ├── agent.js              # Iterative agent loop (call → tool → result → repeat)
│   └── app.js                # UI controller: chat, activity feed, Stop, Undo
├── assets/
│   └── icon-*.png            # Add-in icons
├── dev_server.py             # HTTPS development server
└── README.md
```

## Tools available to the agent

| Tool | Description |
|---|---|
| `get_workbook_overview` | Sheets, used ranges, headers, column types, numeric stats (sum/avg/min/max/count) |
| `read_range` | Read values / formulas / formats of any range |
| `find_in_workbook` | Locate a value or label across a sheet or the whole workbook |
| `get_objects` | List existing tables, pivots, charts, slicers, named ranges |
| `write_range` | Write a 2D array (formulas auto-detected); returns a sample to catch errors |
| `format_range` | Bold, fill, borders, column width, row height, merge, number format, alignment |
| `clear_range` | Clear values and formatting from a range |
| `add_sheet` / `delete_sheet` | Create / delete worksheets (delete guarded) |
| `create_table` | Create an Excel Table from a range |
| `create_pivot` | Create a PivotTable |
| `create_chart` | Create a chart from a range or pivot |
| `add_slicer` | Add a slicer to a pivot or table |
| `conditional_format` | Color scale, data bar, or cell-value rules |
| `autofit` | Auto-fit column widths / row heights |
| `insert_rows_cols` / `delete_rows_cols` | Structural edits (undo is partial) |
| `sort_range` | Sort a range by a key column |

Each mutating tool records a pre-image with the Journal so the user can undo the whole request.

## Gemini API notes

- Default model: `gemini-3.6-flash` (free tier).
- Fallback model: `gemini-3.5-flash-lite` (separate quota).
- Uses native **function calling** (`tools` + `toolConfig`).
- Handles 429 rate limits by parsing Google's `retryDelay` and waiting with a visible countdown.
- Per-day quota exhaustion fails fast with a clear message.
- API key is stored in `localStorage` and only sent to Google's API.
- Model reasoning is streamed live via `streamGenerateContent` with `includeThoughts: true`.
- The agent loop makes multiple Gemini calls per request (5–15+); this is expected and necessary for self-correction.

## Limitations

- Requires Excel 2021+ (Office.js pivot/slicer APIs).
- Office.js bypasses Excel's undo stack — mitigated by the per-request Journal and the **Undo** button.
- Free-tier rate limits may be hit faster than before because the agent makes multiple calls per request.
- Calculated pivot fields and classic pivot layout are not supported by the Office.js API.
- Structural edits (insert/delete rows/cols) and deletions of existing data sheets are only partially undoable — the agent will confirm with you before destructive operations.

## Troubleshooting

**"Por favor, ingresa tu API key"** — Click the gear icon and paste your key from [aistudio.google.com/apikey](https://aistudio.google.com/apikey).

**Add-in doesn't appear** — Ensure you uploaded the correct manifest (`manifest-localhost.xml` for local dev). Check that the server is running on port 3000.

**Sidebar shows old code after updates** — Office's cache is aggressive. Close Excel and delete the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, then reopen Excel and reload the add-in.

**Pivot/slicer errors** — Confirm you're running Excel 2021+ or Microsoft 365. Older versions don't support the required Office.js APIs.

**Rate limited (429)** — The free tier has per-minute and per-day limits. Wait a minute for per-minute limits. For per-day limits, try again after midnight Pacific time.

**Agent hit the round limit** — The agent caps at 20 tool-call rounds per request. If it stops mid-task, send a follow-up message like "continue" or "verify what you built so far".
