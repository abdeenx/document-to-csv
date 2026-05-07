# document-to-csv

A typesafe TypeScript CLI that converts document images and PDFs into structured **CSV**, **Excel**, or **Word** files using local LLMs served by [LM Studio](https://lm-studio.ai).

Four pipelines are available, each using a different combination of local LLMs:

| Flag | Models needed | Output |
|------|--------------|--------|
| *(none / CSV)* | DeepSeek-OCR + Gemma4 | `.csv` |
| `--excel` | DeepSeek-OCR + Gemma4 | `.xlsx` |
| `--word` | DeepSeek-OCR + dots.ocr + GLM-OCR + Chandra-OCR + Gemma4 | `.docx` (4-model corroboration) |
| `--qwen-word` | Qwen2.5-VL-7B-Instruct | `.docx` (single-model, fast) |

All models are served locally by LM Studio. No internet connection required after initial download.

---

## Prerequisites

### LM Studio

Download from [lm-studio.ai](https://lm-studio.ai), load both models, and start the local server (default `http://localhost:1234`).

### PDF renderer (PDF inputs only)

Required for the DeepSeek-OCR visual pass on each page and for image embedding in Excel output.

```bash
brew install poppler       # provides pdftoppm  ← recommended
brew install mupdf-tools   # provides mutool    ← fallback
```

If neither is installed the CLI falls back to text-layer-only extraction and prints a hint.

### Node.js & pnpm

```bash
node --version   # 22 or 24
pnpm --version   # 10+
```

---

## Installation

```bash
git clone https://github.com/abdeenx/document-to-csv.git
cd document-to-csv
pnpm install
```

---

## Usage

```bash
pnpm --filter @workspace/scripts run document-to-csv <file> [options]
```

### Arguments

| Argument | Description |
|----------|-------------|
| `<file>` | Path to the input file. Supported: `pdf`, `jpg`, `jpeg`, `png`, `gif`, `webp` |

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o, --output <path>` | `<basename>.csv` / `.xlsx` / `.docx` | Output file path |
| `--excel` | off | Write `.xlsx` instead of `.csv` (see [Excel output](#excel-output)) |
| `--word` | off | Write `.docx` via 4-model OCR + corroboration (see [Word output](#word-output)) |
| `--enhance` | off | Re-run OCR on weak pages in an existing `--word` progress file, then regenerate |
| `--qwen-word` | off | Write `.docx` via a single Qwen2.5-VL pass per page (see [Qwen Word output](#qwen-word-output-qwen-word)) |
| `--lm-studio-url <url>` | `http://localhost:1234/v1` | LM Studio server URL |
| `--ocr-model <id>` | `mlx-community/DeepSeek-OCR-8bit` | DeepSeek OCR model ID |
| `--structurer-model <id>` | `zecanard/...` | Corroboration / structuring model |
| `--qwen-model <id>` | `mlx-community/Qwen2.5-VL-7B-Instruct-8bit` | Qwen VLM model for `--qwen-word` |
| `-v, --verbose` | off | Step-by-step logging |
| `-h, --help` | — | Show help |

`--excel`, `--word`, and `--qwen-word` are mutually exclusive. Both `--word` and `--qwen-word` require a PDF input.

### Examples

```bash
# Image → CSV
pnpm --filter @workspace/scripts run document-to-csv ./invoice.png

# PDF → CSV with verbose logging
pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --verbose

# PDF → layout-faithful Excel file
pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel

# PDF → Word document via 4-model corroboration (Arabic + Latin, resumable)
pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word

# PDF → Word, verbose, custom output path
pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf \
  --word \
  --output ./exports/contract.docx \
  --verbose

# PDF → Word via Qwen2.5-VL (single model, fast, Arabic + Latin, resumable)
pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen-word

# Qwen Word with custom output path
pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf \
  --qwen-word \
  --output ./exports/arabic.docx \
  --verbose

# Custom output path, remote LM Studio
pnpm --filter @workspace/scripts run document-to-csv ./statement.pdf \
  --excel \
  --output ./exports/statement.xlsx \
  --lm-studio-url http://192.168.1.10:1234/v1
```

---

## How it works

### Mode 1 — Image → CSV / Excel

```
Image file
   │
   ├─ Step 1 — DeepSeek-OCR extracts text (base64 encoded, vision call)
   │
   └─ Step 2 — Gemma4 tool-use loop
                Receives: image + OCR text
                Calls write_csv tool → CSV / Excel
```

### Mode 2 — PDF → CSV / Excel

```
PDF file
   │
   ├─ Step 1 — Text extraction (parallel)
   │    ├─ pdfjs-dist: embedded text layer (structural, tab-delimited)
   │    └─ DeepSeek-OCR: visual pass on each rendered page (pdftoppm / mutool)
   │
   ├─ Step 2 — Gemma4 visual verification (non-tool call)
   │    Receives: rendered page images + both text layers
   │    Produces: single reconciled, tab-delimited text
   │
   └─ Step 3 — Gemma4 tool-use loop
                Receives: verified text + page images
                Calls write_csv tool → CSV / Excel
```

When `--excel` is used on a PDF, page layout data (text positions + image bounding boxes) is extracted in parallel with Step 1 at no extra cost.

### Mode 3 — PDF → Word (`--word`)

```
PDF file
   │
   ├─ Step 1 — pdfjs (page count only; text layer discarded)
   │
   └─ Per page (resumable — progress saved to .progress.json after each page):
        │
        ├─ Render page → JPEG (pdftoppm / mutool)
        │
        ├─ Passes 1–4 in parallel — four OCR models
        │    DeepSeek-OCR  — visual OCR
        │    dots.ocr      — visual OCR
        │    GLM-OCR       — visual OCR
        │    Chandra-OCR   — visual OCR
        │    All outputs: thinking blocks stripped
        │
        └─ Pass 5 — Gemma4 corroboration
             Receives: page image + all 4 OCR texts
             Produces: single most accurate version of the page
             Rules: trust content in 3+ sources; use page image as
                    visual ground truth; thinking blocks stripped

→ Word generator builds .docx from corroborated page texts
  Arabic paragraphs → bidirectional + right-aligned
  Latin paragraphs  → left-to-right
  Tables → tab-delimited rows reconstructed as Word tables
  Headings → Word Heading 2 style
  Each PDF page → new page in the Word document
```

### Mode 4 — PDF → Word via Qwen2.5-VL (`--qwen-word`)

```
PDF file
   │
   ├─ Step 1 — pdfjs (page count only)
   │
   └─ Per page (resumable — progress saved to .qwen-progress.json after each page):
        │
        ├─ Render page → JPEG (pdftoppm / mutool)
        │
        └─ Single pass — Qwen2.5-VL-7B-Instruct
             Receives: rendered page image
             Produces: extracted text (Arabic + Latin)
             Thinking blocks stripped before saving

→ Word generator builds .docx from Qwen-extracted page texts
  Same Arabic RTL, headings, tables, and page-break logic as --word
```

**When to use `--qwen-word` vs `--word`:**

| Consideration | `--qwen-word` | `--word` |
|---------------|--------------|---------|
| Models needed | 1 | 5 |
| Speed | Faster (1 call/page) | Slower (5 calls/page) |
| Accuracy strategy | Single strong VLM | Multi-model consensus |
| Best for | Quick conversions, lighter hardware | Maximum accuracy, longer documents |

---

## Excel output

When `--excel` is passed the workbook contains two sheets:

### `Data` sheet
A clean, styled table — bold header row with light blue fill, alternating row shading, frozen first row, auto-fitted columns. This is equivalent to the CSV but formatted for Excel.

### `Document` sheet (PDF inputs only)
A layout-faithful reconstruction of the original PDF page:

- **Text** — every text run from the PDF is placed in the cell that corresponds to its exact X/Y position. The page coordinate space (in PDF points) is mapped to a fine Excel grid (6 pt per column, 12 pt per row).
- **Images** — images embedded in the PDF (logos, charts, stamps, etc.) are detected by replaying the PDF's drawing operator list and tracking the current transformation matrix (CTM). Each detected image region is cropped from the rendered page raster using [sharp](https://sharp.pixelplumbing.com) and embedded in the Excel sheet anchored to the cell at the matching position.

---

## Word output

When `--word` is passed on a PDF, a `.docx` file is produced with:

- **Structure preserved** — headings, paragraphs, tables, and lists are reconstructed from the corroborated text.
- **Arabic RTL support** — paragraphs and table cells containing Arabic text are automatically marked bidirectional with right alignment so Word renders them correctly.
- **One Word page per PDF page** — each PDF page starts on a new page in the output document.
- **Resumable** — a `.progress.json` file is written next to the output file after every completed page. If the process is interrupted, re-running the same command picks up exactly where it left off. Delete the progress file to force a complete re-run.

```bash
# First run (processes pages 1–20, interrupted at page 12)
pnpm --filter @workspace/scripts run document-to-csv ./doc.pdf --word
# → doc.progress.json contains pages 1–12

# Second run (reads pages 1–12 from cache, resumes at page 13)
pnpm --filter @workspace/scripts run document-to-csv ./doc.pdf --word
```

---

## Architecture

| Component | File | Role |
|-----------|------|------|
| CLI entry | `scripts/src/document-to-csv.ts` | Argument parsing, orchestration, mode routing |
| Types | `scripts/src/types.ts` | Zod schemas + TypeScript interfaces for all data shapes |
| LM Studio client | `scripts/src/lm-studio-client.ts` | Typed OpenAI-compatible client factory |
| OCR | `scripts/src/ocr.ts` | `callOcrModel()` — base64 encode + DeepSeek-OCR vision call; two system prompts (CSV mode, Word mode) |
| PDF extraction | `scripts/src/pdf.ts` | pdfjs text layer + DeepSeek-OCR pass + layout extraction for Excel + raw per-page texts for Word |
| CSV generation | `scripts/src/csv-generator.ts` | `verifyDocumentWithGemma()` pre-pass + `generateCsvWithGemma()` agentic tool-use loop |
| Excel generation | `scripts/src/excel-generator.ts` | `generateExcel()` — styled Data sheet + layout-faithful Document sheet |
| PDF → Word pipeline | `scripts/src/pdf-to-word.ts` | 4-pass per-page extraction, progress tracking, corroboration |
| Word generation | `scripts/src/word-generator.ts` | `generateWordDoc()` — docx with Arabic RTL, headings, tables, page breaks |

### Key decisions

- **LM Studio via OpenAI SDK** — LM Studio exposes an OpenAI-compatible `/v1` endpoint, so the standard `openai` npm package is used directly. No custom HTTP layer needed.
- **Three-pass PDF→CSV pipeline** — (1) pdfjs text layer for structure, (2) DeepSeek-OCR for visual verification, (3) Gemma4 reconciliation before CSV generation. This gives better results than any single source alone.
- **Four-pass PDF→Word pipeline** — the Word mode adds a third independent extraction pass (Gemma4 vision) then a dedicated corroboration step. Content confirmed by 2+ sources is trusted; single-source content is verified against the page image.
- **Per-page progress persistence** — the Word pipeline writes a JSON progress file after every page. On restart, already-completed pages are loaded from the file instead of being re-processed. This makes long conversions safe to interrupt.
- **Agentic tool-use loop (CSV/Excel)** — Gemma4 calls the `write_csv` tool up to 8 iterations. If it stops without calling the tool it is re-prompted automatically.
- **Zod validation on every tool call** — `WriteCsvToolArgsSchema` validates Gemma4's `write_csv` arguments before accepting them.
- **RFC 4180 CSV sanitizer** — the LLM's CSV output is re-parsed and re-serialized programmatically to guarantee correct quoting and consistent column counts.
- **pdfjs-dist v5 legacy build** — the main ESM build requires DOM globals absent in Node.js. The legacy CJS build is loaded via `createRequire`. `GlobalWorkerOptions.workerSrc` must point to the actual worker file; pdfjs v5 removed the in-process fake worker.
- **CTM-tracked image detection** — PDF image positions are derived by replaying the operator list and accumulating the current transformation matrix through `save`/`restore`/`transform` ops. Regions are cropped from the rendered raster using sharp.
- **Arabic RTL in Word** — paragraphs and table cells containing Arabic Unicode characters are detected at render time and set `bidirectional: true` + `AlignmentType.RIGHT` + `rightToLeft: true` on text runs, producing correct RTL rendering in Microsoft Word without any manual markup.

---

## Supported document types

| Type | CSV/Excel | Word |
|------|-----------|------|
| Invoices | Tabular line items, totals, key-value fields | Full layout with headings and tables |
| Bank statements | Account summaries, transaction tables | Multi-page statements |
| Contracts / legal | — | Multi-page documents, Arabic + Latin mixed |
| Reports | Multi-section documents, mixed tables | Full narrative + tables |
| Forms | Key-value pairs, checkboxes | Form fields as key-value text |
| Spreadsheet screenshots | Grid data, merged cells | — |
| Scanned documents | Export as PNG/JPEG first | Export as PNG/JPEG first |

---

## Typecheck

```bash
pnpm --filter @workspace/scripts run typecheck
# or full workspace:
pnpm run typecheck
```
