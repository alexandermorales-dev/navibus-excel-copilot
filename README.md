# Excel AI Copilot

An intelligent Excel add-in that answers questions about your workbook and builds reports, dashboards, tables, charts, and pivots — using a **plan → execute → verify** orchestrator that typically costs **1 API call per request** instead of the 5–15 calls of a traditional ReAct loop.

## How it works

1. You type a request (e.g. "crea un panel de control con KPIs" or "¿cuál es el total de ventas por región?").
2. A **local snapshot** of the workbook (sheets, headers, column types, sample rows, numeric stats) is built in-code — no API call.
3. **Fast-path check**: common requests (sum, average, count, sort, filter) are answered directly from the snapshot when possible — **0 API calls**.
4. **Plan**: one API call sends the snapshot + request to the model, which returns a JSON op list (write, format, create_chart, etc.).
5. **Validate**: ops are checked locally against the snapshot before Excel is touched — bad sheet names, missing fields, and unknown ops are dropped silently.
6. **Execute**: ops run against the real workbook via Office.js, with deterministic layout recipes handling formatting in-code (no prompt needed).
7. **Verify**: results are read back locally to catch `#REF!` and other formula errors.
8. **Repair** (only if verification found issues): up to 2 follow-up calls to fix problems.
9. A concise summary is rendered in your language, with suggested follow-ups.
10. Every mutation is journaled — an **Undo** button reverts the whole request.

Typical cost: **1 request**. Worst case: **4**. Everything that doesn't need a language model is done in code, where it's deterministic and testable.

## Multi-provider failover

The add-in pools three free-tier providers so a single exhausted quota doesn't block you:

| Provider | Free tier | Best for |
|---|---|---|
| **Google Gemini** | 1,000 req/day, 250k tokens/min | Primary — generous token ceiling for snapshot-bearing plan calls |
| **Groq** | 1,000 req/day, 8k tokens/min | Fast answers & repairs (cached prompt tokens don't count toward TPM) |
| **OpenRouter** | 50 req/day, 200k tokens/min | Emergency reserve only |

Requests are routed by priority (Gemini → Groq → OpenRouter). When a provider hits a rate limit, the quota governor automatically fails over to the next one — you see a badge showing which provider served the request.

**Bring your own keys.** There is no shared preset key. Each user enters their own keys in Settings; only providers with a key entered are used. Keys are stored in `localStorage` and sent only to the respective provider's API.

## Features

- **Plan → execute → verify** — 1 API call typically; no iterative tool-calling loop.
- **Local snapshot injection** — the model sees the whole workbook structure without spending a call to discover it.
- **Deterministic layout recipes** — formatting (column widths, KPI card layout, chart placement, color scales) is computed in code, not prompted.
- **Multi-provider failover** with a visible quota bar per provider.
- **Live reasoning stream** — watch the model's thinking appear in real time.
- **Plan checklist** — see the intended op sequence before it runs; rows tick off as ops execute.
- **Self-repair** — when verification finds `#REF!` or other errors, the model gets one chance to fix them.
- **Stop button** — abort a run mid-flight.
- **Undo last request** — reverts all mutations from the most recent request.
- **Fast-path answers** — simple aggregations and lookups answered locally from the snapshot.
- **Markdown rendering** — final answers support bold, lists, and inline code.
- **Dark theme, ES/EN auto-detect** — UI and model output follow Office's display language.

## Requirements

- **Microsoft Excel 2021+ or Microsoft 365** (desktop, Windows). Requires WebView2 + Office.js pivot/slicer APIs.
- **At least one free API key** from:
  - [Google Gemini](https://aistudio.google.com/apikey) (recommended — best free-tier limits)
  - [Groq](https://console.groq.com/keys)
  - [OpenRouter](https://openrouter.ai/keys)
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

### 3. Enter your API keys

1. Click the **gear icon** in the sidebar header.
2. Paste at least one API key (Gemini recommended). Each provider has a **Test** button to verify the key works.
3. The quota bar shows remaining requests/day for each configured provider.
4. Close settings and start chatting.

## Sharing with other users

Once hosted on GitHub Pages:

1. Push this project to a GitHub repository.
2. Enable **GitHub Pages** (Settings → Pages → Source: main branch, root).
3. Edit `manifest.xml` — replace `YOUR_GITHUB_USERNAME` with your GitHub username and `YOUR_REPO_NAME` with your repo name.
4. Share `manifest.xml` with the other person.
5. They open Excel → **Insert** → **Add-ins** → **Upload My Add-in** → select the file.
6. They enter their own API keys in settings.

No installer, no Node.js, no build step on their end.

## Project structure

```
├── manifest.xml              # For GitHub Pages (sharing)
├── manifest-localhost.xml    # For local development
├── index.html                # Task pane entry point
├── package.json              # Test runner (node --test)
├── css/
│   └── style.css             # Dark-themed UI + plan checklist + markdown
├── js/
│   ├── config.js             # Per-provider keys + settings in localStorage
│   ├── i18n.js               # ES/EN string table + Office locale detection
│   ├── providers.js          # Provider registry (Gemini/Groq/OpenRouter) + model resolution
│   ├── quota.js              # Quota governor — tracks RPD/RPM/TPM, routes & fails over
│   ├── llm.js                # Single OpenAI-compatible client for all providers
│   ├── schema.js             # Workbook snapshot via Office.js (detail levels)
│   ├── context.js            # Snapshot injection + prompt-cache-friendly system prefix
│   ├── prompt.js             # Compact plan/repair/answer prompts (ES/EN)
│   ├── intent.js             # Local fast-paths (sum/avg/count/sort/lookup from snapshot)
│   ├── ops.js                # Op schema, validator, executor, verifier
│   ├── recipes.js            # Deterministic layout recipes (KPI cards, dashboards, charts)
│   ├── journal.js            # Per-request mutation journal → Undo
│   ├── tools.js              # Office.js tool handlers (read/write/format/create)
│   ├── agent.js              # Plan → execute → verify orchestrator
│   └── app.js                # UI controller: chat, plan checklist, Stop, Undo
├── test/
│   ├── helper.js             # VM sandbox loader for browser-globals in node --test
│   ├── llm.test.js           # LLM client tests
│   ├── quota.test.js         # Quota governor tests
│   ├── ops.test.js           # Op validator + verifier tests
│   └── recipes.test.js       # Layout recipe tests
├── assets/
│   └── icon-*.png            # Add-in icons
├── dev_server.py             # HTTPS development server
└── README.md
```

## Ops available to the planner

The model returns a JSON array of ops. Each op is validated locally before execution.

| Op | Description |
|---|---|
| `read_range` | Read values / formulas / formats of any range (used for data fetch when snapshot is insufficient) |
| `write_range` | Write a 2D array (formulas auto-detected); verifier reads back to catch errors |
| `format_range` | Bold, fill, borders, column width, row height, merge, number format, alignment |
| `clear_range` | Clear values and formatting from a range |
| `add_sheet` / `delete_sheet` | Create / delete worksheets (delete guarded — existing sheets require confirmation) |
| `create_table` | Create an Excel Table from a range |
| `create_pivot` | Create a PivotTable |
| `create_chart` | Create a chart from a range or pivot |
| `add_slicer` | Add a slicer to a pivot or table |
| `conditional_format` | Color scale, data bar, or cell-value rules |
| `autofit` | Auto-fit column widths / row heights |
| `insert_rows_cols` / `delete_rows_cols` | Structural edits (undo is partial) |
| `sort_range` | Sort a range by a key column |

Each mutating op records a pre-image with the Journal so the user can undo the whole request.

## Layout recipes

Formatting that used to be prompted (and frequently produced broken layouts) is now computed deterministically in `recipes.js`:

- **KPI cards** — 3-column grid with header, value, and delta cells, auto-sized.
- **Dashboard layout** — KPIs on top, breakdown table + chart below, column widths pre-computed.
- **Summary blocks** — aggregation table with banner, headers, and totals row.
- **Chart placement** — row spans calculated from KPI count so charts don't overlap.
- **Number formats** — per-column based on aggregation type (currency, integer, percent).

The model just says "build a dashboard with these KPIs and this breakdown"; the recipe figures out the cells, widths, and formats.

## Tests

```bash
node --test test/
```

122 tests cover the quota governor, op validator/verifier, and layout recipes. Tests run in a VM sandbox that loads the source files the same way the browser does (no Node-specific shims).

## Limitations

- Requires Excel 2021+ (Office.js pivot/slicer APIs).
- Office.js bypasses Excel's undo stack — mitigated by the per-request Journal and the **Undo** button.
- Calculated pivot fields and classic pivot layout are not supported by the Office.js API.
- Structural edits (insert/delete rows/cols) and deletions of existing data sheets are only partially undoable — the agent will confirm with you before destructive operations.
- Free-tier model slugs change frequently; the provider registry resolves models dynamically from `/models` endpoints with static fallbacks.

## Troubleshooting

**"No API key configured"** — Click the gear icon and paste at least one key. Gemini is recommended for the best free-tier limits.

**Add-in doesn't appear** — Ensure you uploaded the correct manifest (`manifest-localhost.xml` for local dev). Check that the server is running on port 3000.

**Sidebar shows old code after updates** — Office's cache is aggressive. Close Excel and delete the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, then reopen Excel and reload the add-in.

**Pivot/slicer errors** — Confirm you're running Excel 2021+ or Microsoft 365. Older versions don't support the required Office.js APIs.

**Rate limited (429)** — The quota governor fails over to the next provider automatically. If all providers are exhausted, wait for the reset (Gemini: midnight Pacific; Groq: midnight UTC; OpenRouter: midnight UTC) or add a key for another provider.

**Repair failed** — The verifier found errors and the model couldn't fix them in 2 attempts. Check the plan checklist for warning rows, then send a follow-up message describing what went wrong.
