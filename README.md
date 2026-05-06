# document-to-csv

A typesafe TypeScript CLI that converts document images and PDFs into structured **CSV** or **Excel** files using local LLMs served by [LM Studio](https://lm-studio.ai).

Two models collaborate:

| Role | Default model |
|------|---------------|
| **OCR / vision** | `mlx-community/DeepSeek-OCR-8bit` |
| **Structuring** | `zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine` |

Both must be loaded in LM Studio before running. No internet connection required after that — everything runs locally.

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
| `-o, --output <path>` | `<basename>.csv` / `.xlsx` | Output file path |
| `--excel` | off | Write `.xlsx` instead of `.csv` (see [Excel output](#excel-output)) |
| `--lm-studio-url <url>` | `http://localhost:1234/v1` | LM Studio server URL |
| `--ocr-model <id>` | `mlx-community/DeepSeek-OCR-8bit` | Vision/OCR model ID |
| `--structurer-model <id>` | `zecanard/...` | Structuring model ID |
| `-v, --verbose` | off | Step-by-step logging |
| `-h, --help` | — | Show help |

### Examples

```bash
# Image → CSV
pnpm --filter @workspace/scripts run document-to-csv ./invoice.png

# PDF → CSV with verbose logging
pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --verbose

# PDF → layout-faithful Excel file
pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel

# Custom output path, remote LM Studio
pnpm --filter @workspace/scripts run document-to-csv ./statement.pdf \
  --excel \
  --output ./exports/statement.xlsx \
  --lm-studio-url http://192.168.1.10:1234/v1
```

---

## How it works

### Image input

```
Image file
   │
   ├─ Step 1 — DeepSeek-OCR extracts text (base64 encoded, vision call)
   │
   └─ Step 2 — Gemma4 tool-use loop
                Receives: image + OCR text
                Calls write_csv tool → CSV / Excel
```

### PDF input

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
   │    (corrects column misalignment between text layer and OCR)
   │
   └─ Step 3 — Gemma4 tool-use loop
                Receives: verified text + page images
                Calls write_csv tool → CSV / Excel
```

When `--excel` is used on a PDF, page layout data (text positions + image bounding boxes) is extracted in parallel with Step 1 at no extra cost.

---

## Excel output

When `--excel` is passed the workbook contains two sheets:

### `Data` sheet
A clean, styled table — bold header row with light blue fill, alternating row shading, frozen first row, auto-fitted columns. This is equivalent to the CSV but formatted for Excel.

### `Document` sheet (PDF inputs only)
A layout-faithful reconstruction of the original PDF page:

- **Text** — every text run from the PDF is placed in the cell that corresponds to its exact X/Y position. The page coordinate space (in PDF points) is mapped to a fine Excel grid (6 pt per column, 12 pt per row).
- **Images** — images embedded in the PDF (logos, charts, stamps, etc.) are detected by replaying the PDF's drawing operator list and tracking the current transformation matrix (CTM). Each detected image region is cropped from the rendered page raster using [sharp](https://sharp.pixelplumbing.com) and embedded in the Excel sheet anchored to the cell at the matching position.

The result is an Excel file where you can see both the clean extracted table (Data sheet) and a view of the original document with images in their correct relative positions (Document sheet).

---

## Architecture

| Component | File | Role |
|-----------|------|------|
| CLI entry | `scripts/src/document-to-csv.ts` | Argument parsing, orchestration, step logging |
| Types | `scripts/src/types.ts` | Zod schemas + TypeScript interfaces for all data shapes |
| LM Studio client | `scripts/src/lm-studio-client.ts` | Typed OpenAI-compatible client factory |
| OCR | `scripts/src/ocr.ts` | `callOcrModel()` — base64 encode + DeepSeek-OCR vision call |
| PDF extraction | `scripts/src/pdf.ts` | pdfjs text layer + DeepSeek-OCR pass + layout extraction for Excel |
| CSV generation | `scripts/src/csv-generator.ts` | `verifyDocumentWithGemma()` pre-pass + `generateCsvWithGemma()` agentic loop |
| Excel generation | `scripts/src/excel-generator.ts` | `generateExcel()` — styled Data sheet + layout-faithful Document sheet |

### Key decisions

- **LM Studio via OpenAI SDK** — LM Studio exposes an OpenAI-compatible `/v1` endpoint, so the standard `openai` npm package is used directly. No custom HTTP layer needed.
- **Three-pass PDF pipeline** — (1) pdfjs text layer for structure, (2) DeepSeek-OCR for visual verification, (3) Gemma4 reconciliation before CSV generation. This gives better results than any single source alone.
- **Agentic tool-use loop** — Gemma4 calls the `write_csv` tool up to 8 iterations. If it stops without calling the tool it is re-prompted automatically.
- **Zod validation on every tool call** — `WriteCsvToolArgsSchema` validates Gemma4's `write_csv` arguments before accepting them.
- **RFC 4180 CSV sanitizer** — the LLM's CSV output is re-parsed and re-serialized programmatically to guarantee correct quoting and consistent column counts.
- **pdfjs-dist v5 legacy build** — the main ESM build requires DOM globals absent in Node.js. The legacy CJS build is loaded via `createRequire` (Node.js 22+ allows `require()` of `.mjs` files). `GlobalWorkerOptions.workerSrc` must point to the actual worker file; pdfjs v5 removed the in-process fake worker.
- **CTM-tracked image detection** — PDF image positions are derived by replaying the operator list and accumulating the current transformation matrix through `save`/`restore`/`transform` ops. No raw PDF image object extraction needed — regions are cropped from the existing rendered raster.

---

## Supported document types

| Type | Works best for |
|------|---------------|
| Invoices | Tabular line items, totals, key-value header fields |
| Bank statements | Account summaries, transaction tables |
| Reports | Multi-section documents, mixed text and tables |
| Forms | Key-value pairs, checkboxes (visual toggle detection) |
| Spreadsheet screenshots | Grid data, merged cells |
| Scanned documents | Export pages as PNG/JPEG first, then pass to the CLI |

---

## Typecheck

```bash
pnpm --filter @workspace/scripts run typecheck
# or full workspace:
pnpm run typecheck
```
