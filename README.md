# Office Agent Plugin

A TypeScript/Node.js Office document editing agent that integrates with DeepSeek V4. Supports `.docx` and `.pptx` files with structured patch operations, validation loops, and cache-friendly prompt building.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                          API Server                             │
│  POST /api/agent/run (upload file + task → modified file)       │
└─────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                         Agent Loop                              │
│  1. Inspect file → generate manifest                           │
│  2. Build cache-friendly prompt (stable prefix first)          │
│  3. Ask LLM for patch plan (tool calling)                      │
│  4. Validate patch JSON (Zod schema)                           │
│  5. Execute patch (deterministic XML modification)             │
│  6. Validate output file (ZIP + XML integrity)                 │
│  7. If fails → retry with error context (max 4 iterations)    │
└─────────────────────────────────────────────────────────────────┘
                                │
            ┌───────────────────┼───────────────────┐
            ▼                   ▼                   ▼
    ┌──────────────┐   ┌──────────────┐   ┌──────────────┐
    │  LLM Provider │   │ Office Core  │   │ Patch System │
    │  (DeepSeek)   │   │ (DOCX/PPTX)  │   │ (Validator)  │
    └──────────────┘   └──────────────┘   └──────────────┘
```

## Quick Start

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Environment

Copy `.env.example` to `.env` and fill in your DeepSeek API credentials:

```bash
cp .env.example .env
```

Edit `.env`:
```env
DEEPSEEK_API_KEY=sk-your-api-key-here
DEEPSEEK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-v4
```

**Important**: Do NOT hardcode your API key in source code. Use environment variables.

### 3. Build and Run

```bash
npm run build
npm start
```

The server will start at `http://localhost:3000`.

### 4. Upload a Document

```bash
curl -X POST http://localhost:3000/api/agent/run \
  -F "file=@document.docx" \
  -F "task=Change the title to 'New Title'" \
  --output modified.docx
```

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `DEEPSEEK_API_KEY` | (required) | Your DeepSeek API key |
| `DEEPSEEK_BASE_URL` | `https://api.deepseek.com/v1` | API base URL |
| `DEEPSEEK_MODEL` | `deepseek-v4` | Model name |
| `PORT` | `3000` | Server port |
| `HOST` | `0.0.0.0` | Server host |
| `MAX_ITERATIONS` | `4` | Max agent loop iterations |
| `MAX_PATCH_OPERATIONS` | `20` | Max operations per patch |
| `MAX_OUTPUT_TOKENS` | `4096` | Max LLM output tokens |
| `LOG_LEVEL` | `info` | Logging level |

### Using Other Providers

The LLM provider is configurable. To use a different provider:

1. Implement the `LLMProvider` interface
2. Update the provider instantiation in `src/api/server.ts`

The interface is OpenAI-compatible, so any provider supporting the chat/completions API should work.

## Features

### Supported Operations

#### DOCX Operations
- `replace_paragraph_text` - Replace entire paragraph text
- `replace_text_in_paragraph` - Find and replace specific text
- `update_table_cell_text` - Update table cell content
- `insert_paragraph_after` - Insert new paragraph after target
- `delete_paragraph` - Remove a paragraph
- `apply_paragraph_style` - Change paragraph style
- `update_header_text` - Update header content
- `update_footer_text` - Update footer content

#### PPTX Operations
- `replace_shape_text` - Replace text in a shape
- `replace_slide_title` - Update slide title
- `update_speaker_notes` - Modify speaker notes
- `insert_slide_from_layout` - Add new slide from layout
- `delete_slide` - Remove a slide
- `move_shape` - Change shape position
- `resize_shape` - Change shape dimensions
- `apply_text_style` - Modify text formatting
- `replace_image` - Replace image content
- `fit_text_to_shape` - Auto-fit text to shape bounds

### Cache-Friendly Prompt Building

The prompt builder optimizes for DeepSeek V4's cache mechanism:

1. **Stable prefix first**: System rules, tool specs, patch schema, validation rules, manifest
2. **Dynamic suffix last**: Iteration number, tool results, validation errors

This ensures maximum cache hit rates, reducing costs and latency.

### Validation

Every patch execution includes:
- ZIP package integrity check
- Required OOXML parts existence
- XML parsing validation
- Relationship target verification
- Content type validation

## Running Tests

```bash
npm test
```

All 39 tests cover:
- DOCX/PPTX manifest generation
- DOCX/PPTX patch operations
- Patch validation (Zod schemas)
- Prompt builder determinism
- Utility functions (stable JSON, hashing)
- File validation

## Project Structure

```
office-agent-plugin/
├── src/
│   ├── api/              # Express server
│   ├── agent/            # Agent loop and prompt builder
│   ├── llm/              # LLM provider adapter
│   ├── office/           # OOXML package handling
│   │   ├── docx/         # DOCX-specific logic
│   │   └── pptx/         # PPTX-specific logic
│   ├── patches/          # Patch validation and execution
│   ├── types/            # TypeScript type definitions
│   └── utils/            # Utility functions
├── tests/                # Test files
│   └── fixtures/         # Test fixtures (minimal docx/pptx)
├── prompts/              # System prompts
└── schemas/              # JSON schemas
```

## Security Considerations

1. **Prompt Injection Defense**: Document content is treated as DATA, not instructions. Any text attempting to change agent behavior is ignored.

2. **No Macro Execution**: The agent never executes VBA macros or accesses external URLs.

3. **API Key Security**: API keys are loaded from environment variables, never hardcoded.

4. **Logging Safety**: Document content is never logged. Only metadata and token usage are recorded.

5. **Input Validation**: All patch operations are validated against Zod schemas before execution.

## Limitations

1. **DOCX Support**: Basic paragraph and table operations. Complex formatting (images, charts, SmartArt) not yet supported.

2. **PPTX Support**: Text replacement and basic shape operations. Complex animations and transitions not supported.

3. **No Office.js Integration**: Currently operates via file upload/download, not direct Office integration.

4. **No Visual Validation**: Cannot verify visual appearance of modified documents.

5. **Single Document**: Processes one document at a time. No batch processing.

## Why Stable Prompt Prefix?

DeepSeek V4 (and similar LLMs) use KV-cache to speed up inference. The cache works best when the prompt prefix is identical across requests.

By putting stable content first:
- System rules
- Tool specifications
- Patch schema
- Validation rules
- Document manifest
- User task

And dynamic content last:
- Current iteration
- Previous tool results
- Validation errors

We maximize cache hits, reducing:
- Latency (faster responses)
- Cost (cache hits are cheaper)
- Token usage (cached tokens don't count)

## Viewing Cache Metrics

Cache metrics are logged with each LLM call:

```json
{
  "inputTokens": 1500,
  "outputTokens": 200,
  "cacheHitTokens": 1200,
  "cacheMissTokens": 300,
  "cacheHitRatio": 0.8
}
```

The `stablePrefixHash` is also logged to verify prompt consistency.

## Roadmap

1. **Full Office.js Add-in**: Direct integration with Word/PowerPoint
2. **Richer DOCX Support**: Images, charts, styles, headers/footers
3. **Richer PPTX Layout QA**: Visual validation of slide layouts
4. **XLSX Support**: Spreadsheet editing capabilities
5. **Visual Slide Thumbnails**: Image-based validation
6. **Human-in-the-Loop Approval**: Review patches before execution

## License

MIT
