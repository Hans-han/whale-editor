# 🐋 Whale Editor

> AI-powered Office document editor — DeepSeek-driven OOXML patching with stable-prefix prompt caching.

Drag a `.docx` / `.pptx` (or legacy `.doc` / `.ppt` via libreoffice) into the browser, describe what you want changed (or attach a PDF / Markdown / Word format spec as reference), and Whale Editor's planner-executor agent will:

1. **Scan** the document into a structured manifest
2. **Plan** — a non-thinking model (`deepseek-chat` by default) breaks your intent into 1–8 atomic patch steps
3. **Execute** — each step is run by the executor model (e.g. `deepseek-v4-flash`) with structured patch operations
4. **Validate** every patch against OOXML schema before committing
5. **Stream** progress back live (NDJSON) — including a 🐋-on-♾️ animation, per-step before→after diffs, and live cache-hit metrics
6. **Reuse** the same source file up to 100 times within a 10-min sliding window so DeepSeek's KV cache stays warm — typical hit rate **80–95%+** on follow-up edits

Built on TypeScript / Node.js. Web UI at `apps/web/`, an experimental Office.js task pane in `apps/office-addin/`.

---

## Architecture

```
   Web UI (apps/web/)                        Optional: Word / PowerPoint addin
   ▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔                        ▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔▔
   - drag-drop file                          (sideload manifest.xml,
   - intent textarea                          requires HTTPS)
   - reference attachments
   - BYO API key (header)                              │
              │                                        │
              ▼                                        ▼
   ┌─────────────────────────────────────────────────────────────┐
   │  POST /api/agent/auto      (multipart + NDJSON streaming)   │
   │  POST /api/agent/run       (legacy single-task, file out)   │
   │  GET  /api/health          (status + libreoffice check)     │
   └─────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
                  ┌────────────────────────────┐
                  │       runAuto()            │
                  │  ─────────────────────     │
                  │  manifest → planner →      │
                  │  N × (runAgent step)       │
                  │  → stream events to client │
                  └────────────────────────────┘
                                  │
            ┌─────────────────────┼─────────────────────┐
            ▼                     ▼                     ▼
   ┌────────────────┐   ┌────────────────┐   ┌─────────────────┐
   │ Planner (LLM)  │   │ Executor agent │   │  OOXML patches  │
   │ deepseek-chat  │   │ deepseek-v4-…  │   │ Zod-validated   │
   │ submit_plan()  │   │ apply_patch()  │   │ ZIP/XML safe    │
   └────────────────┘   └────────────────┘   └─────────────────┘
```

Stable prompt prefix (system rules + tool spec + manifest + intent) is reused across reruns so 80–95%+ of input tokens hit DeepSeek's prompt cache.

---

## Quick Start

Requires Node.js ≥ 18 and a DeepSeek API key ([get one here](https://platform.deepseek.com/)).

```bash
git clone https://github.com/Hans-han/whale-editor.git
cd whale-editor
npm install
cp .env.example .env
# edit .env, paste your sk-... DeepSeek key
npm run build
npm start
```

Open **http://localhost:3000/** in your browser.

> Drop a `.docx` or `.pptx`, type what you want changed (or attach a reference doc), and click **生成方案并执行**. The 🐋 will start swimming. Every iteration shows a live cache-hit ratio.

### .env minimum

```env
DEEPSEEK_API_KEY=sk-paste-yours-here
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-v4-flash
# Optional: planner uses deepseek-chat by default (non-thinking, faster, more reliable for tool calls)
# DEEPSEEK_PLANNER_MODEL=deepseek-chat
```

You can also leave `DEEPSEEK_API_KEY` blank and let each user paste their own key in the UI (top-right `⚙ API Key`). The key is stored in `localStorage` and sent per-request via `X-DeepSeek-Key` header — never logged server-side.

---

## What you can upload

### Target document (the thing being edited)

| Format | Native | Via libreoffice* |
|---|---|---|
| `.docx` / `.pptx` | ✅ | — |
| `.doc` / `.ppt` (legacy binary) | — | ✅ auto-converted to OOXML |

\* If you want `.doc` / `.ppt` support, install libreoffice once: `brew install --cask libreoffice` (macOS) or your distro's equivalent. Whale Editor auto-detects `soffice` and converts on upload.

### Reference materials (planning context only — extracted to text)

`.pdf` · `.docx` · `.doc` · `.pptx` · `.ppt` · `.md` · `.txt` (up to 6 attachments per run, 50 MB each, 100 KB text per file).

---

## Optional: Word task-pane add-in (HTTPS required)

Office desktop add-ins need HTTPS. If you want to load Whale Editor inside Word/PowerPoint as a task pane:

```bash
# 1. Generate self-signed certs (one-time)
npm run generate-certs

# 2. Trust the cert for your platform's WebView
#    (macOS) sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/localhost.crt

# 3. Run in office-https mode
npm run start:office-https

# 4. Sideload apps/office-addin/manifest.xml in Word
#    (Insert → My Add-ins → Upload → manifest.xml)
```

The add-in talks to the same backend; the web UI remains the recommended primary interface.

---

## Configuration

### Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `DEEPSEEK_API_KEY` | (empty — UI fallback) | Server-side default key. If empty, every request must supply `X-DeepSeek-Key`. |
| `DEEPSEEK_BASE_URL` | `https://api.deepseek.com/v1` | OpenAI-compatible base URL |
| `DEEPSEEK_MODEL` | `deepseek-v4-flash` | Executor model |
| `DEEPSEEK_PLANNER_MODEL` | `deepseek-chat` | Planner model (kept non-thinking for tool-call reliability) |
| `PORT` | `3000` | HTTP/HTTPS port |
| `HOST` | `0.0.0.0` | Bind address |
| `MAX_ITERATIONS` | `4` | Per-step retry cap |
| `MAX_PATCH_OPERATIONS` | `20` | Operations per patch |
| `MAX_OUTPUT_TOKENS` | `4096` | LLM output budget per call (raise for big patches) |
| `LOG_LEVEL` | `info` | `debug` / `info` / `warn` / `error` |

### Using another OpenAI-compatible provider

`DEEPSEEK_BASE_URL` can point at any OpenAI-compatible endpoint (Together, OpenRouter, your own gateway, etc.). The provider abstraction lives in [`src/llm/`](src/llm/) — extend `BaseLLMProvider` if your provider needs custom headers or a non-OpenAI shape.

---

## API endpoints

### `POST /api/agent/auto` (primary)

Multipart upload with NDJSON streaming response.

| Field | Required | Description |
|---|---|---|
| `file` | first run only | `.docx` / `.pptx` / `.doc` / `.ppt` |
| `sessionId` | follow-up | UUID returned by previous run; bypasses re-upload, hits warm cache |
| `intent` | yes (text + refs ≥ 1) | Natural-language instructions |
| `references[]` | optional | Up to 6 reference files (pdf/doc/docx/ppt/pptx/md/txt) |
| `fileType` | optional | `docx` or `pptx`; inferred from extension if missing |
| `X-DeepSeek-Key` (header) | optional | BYO API key, overrides server default |

The response is `application/x-ndjson`; one JSON event per line:

```json
{"type":"converting","from":"doc","filename":"report.doc"}
{"type":"session","session":{"id":"uuid","reuseRemaining":99,"expiresInMs":600000,...}}
{"type":"manifest_summary","fileType":"docx","objectCount":1572}
{"type":"plan","plan":{"rationale":"...","steps":[...]},"usage":{...}}
{"type":"step_start","index":0,"total":3,"step":{"title":"Fix cover page","task":"..."}}
{"type":"step_done","index":0,"success":true,"iterations":1,"operations":[...],"usage":{...}}
{"type":"complete","success":true,"modifiedFileBase64":"...","totalUsage":{"cacheHitRatio":0.93,...}}
```

### `POST /api/agent/run` (legacy, single-task)

Multipart upload, returns the modified file as binary download (or JSON on error).

```bash
curl -X POST http://localhost:3000/api/agent/run \
  -F "file=@document.docx" \
  -F "task=Change the title to 'New Title'" \
  --output modified.docx
```

### `GET /api/health`

Returns server status + libreoffice availability:

```json
{"status":"ok","timestamp":"...","libreoffice":{"available":true,"path":"/opt/homebrew/bin/soffice"}}
```

---

## Supported patch operations

### DOCX

`replace_paragraph_text` · `replace_text_in_paragraph` · `update_table_cell_text` · `insert_paragraph_after` · `delete_paragraph` · `apply_paragraph_style` · `update_header_text` · `update_footer_text`

### PPTX

`replace_shape_text` · `replace_slide_title` · `update_speaker_notes` · `insert_slide_from_layout` · `delete_slide` · `move_shape` · `resize_shape` · `apply_text_style` · `replace_image` · `fit_text_to_shape`

Every operation goes through Zod schema validation before execution, then full ZIP+XML integrity checks afterwards. Failed patches retry with the validation error fed back into the prompt (max `MAX_ITERATIONS` per step).

---

## Why stable prompt prefix?

DeepSeek (and most modern LLM providers) bill **prompt-cache hits at a fraction of the cold price**. The cache works best when the leading bytes of the prompt are byte-identical across requests.

Whale Editor's `buildPrompt()` puts:

1. system rules
2. tool specs
3. patch schema + validation rules
4. document manifest (deterministic stable JSON)
5. user task (the only varying tail)

…in that order. Re-edits within the 10-min session window keep `(1, 2, 3, 4)` byte-identical, so the typically-massive manifest portion of the prompt is cached. Each LLM call logs `cacheHitTokens / cacheMissTokens / cacheHitRatio / stablePrefixHash` — verify in `/tmp/office-agent.log` (or wherever your logger ships).

---

## Project structure

```
whale-editor/
├── apps/
│   ├── web/              # Primary web UI (vanilla HTML/CSS/JS)
│   │   ├── index.html    # 🐋 + ♾️ progress, drop zones, BYO key
│   │   ├── app.js
│   │   └── app.css
│   └── office-addin/     # Optional Word/PowerPoint task pane
│       └── manifest.xml
├── src/
│   ├── api/              # Express server + multipart + NDJSON
│   ├── agent/            # autoLoop, planner, executor loop, prompt builder
│   ├── llm/              # DeepSeek provider (OpenAI-compatible)
│   ├── office/           # OOXML manifest generators (DOCX/PPTX)
│   ├── patches/          # Patch validator + executor
│   ├── utils/            # extractText (refs), convertOoxml (libreoffice),
│   │                     # sessionCache, logger, hashing
│   └── types/            # Shared TypeScript types
├── prompts/              # System prompt(s)
├── scripts/              # generate-certs.ts
├── schemas/              # JSON schemas
└── tests/                # Vitest suites + minimal docx/pptx fixtures
```

---

## Running tests

```bash
npm test          # compiled JS (run `npm run build` first)
npm run test:ts   # source TS via tsx, no build needed
```

39+ tests cover manifest generation, patch operations, validator, prompt builder determinism, and utility functions.

---

## Security

- **API keys** never appear in logs; user-supplied keys live in browser `localStorage` only and are sent per-request as a header.
- **No macros / external URLs**: agent never executes VBA or follows external links from documents.
- **Document content is data, not instructions**: the planner system prompt explicitly forbids treating document text as control commands.
- **Strict patch validation**: all operations are Zod-validated; the executed buffer is re-validated as a ZIP/OOXML package before being returned.
- **Self-signed dev certs** (`certs/`) are gitignored. For production, terminate TLS at a reverse proxy with a real certificate.

---

## Limitations

- Complex DOCX content (images, charts, SmartArt, tracked changes) is preserved but not yet directly editable via patches.
- PPTX animation/transition editing not supported.
- `.doc` / `.ppt` legacy binary editing requires libreoffice (auto-detected, friendly error if missing).
- Single document per run — no batch processing yet.
- Visual rendering verification (diff screenshots) not yet implemented.

---

## Roadmap

- Visual slide thumbnails (PPTX render + diff)
- Tables, images, charts in patch ops
- XLSX support
- Cumulative-edit mode (each rerun edits the previous result, not the original)
- Public deployment template (Cloudflare Tunnel + real domain)
- Optional human-in-the-loop approval before each step

---

## License

MIT — see [LICENSE](./LICENSE).
