# document-to-csv

  A typesafe TypeScript CLI that converts document images and PDFs into structured **CSV**, **Excel**, or **Word** files using local LLMs served by [LM Studio](https://lm-studio.ai).

  Five models collaborate:

  | Role | Default model |
  |------|---------------|
  | **OCR / vision** | `mlx-community/DeepSeek-OCR-8bit` |
  | **OCR / vision** | `mlx-community/dots.ocr-bf16` |
  | **OCR / vision** | `mlx-community/GLM-OCR-bf16` |
  | **OCR / vision** | `jwindle47/chandra-ocr-2-8bit-mlx` |
  | **Corroboration & structuring** | `zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine` |

  All models must be loaded in LM Studio before running. No internet connection required after that — everything runs locally.

  ---

  ## Prerequisites

  ### LM Studio

  Download from [lm-studio.ai](https://lm-studio.ai), load the models, and start the local server (default `http://localhost:1234`).

  ### PDF renderer (PDF inputs only)

  Required for OCR visual passes on each page and for image embedding in Excel output.

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
  | `--word` | off | Write `.docx` from a PDF (see [Word output](#word-output)) |
  | `--enhance` | off | Re-run OCR on weak pages in an existing progress file (requires `--word`) |
  | `--lm-studio-url <url>` | `http://localhost:1234/v1` | LM Studio server URL |
  | `--ocr-model <id>` | `mlx-community/DeepSeek-OCR-8bit` | DeepSeek-OCR model ID |
  | `--dots-ocr-model <id>` | `mlx-community/dots.ocr-bf16` | dots.ocr model ID |
  | `--glm-ocr-model <id>` | `mlx-community/GLM-OCR-bf16` | GLM-OCR model ID |
  | `--chandra-ocr-model <id>` | `jwindle47/chandra-ocr-2-8bit-mlx` | Chandra-OCR model ID |
  | `--structurer-model <id>` | `zecanard/...` | Corroboration / structuring model ID |
  | `-v, --verbose` | off | Step-by-step logging |
  | `-h, --help` | — | Show help |

  `--excel` and `--word` are mutually exclusive. `--word` requires a PDF input. `--enhance` requires `--word`.

  ### Examples

  ```bash
  # Image → CSV
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png

  # PDF → CSV with verbose logging
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --verbose

  # PDF → layout-faithful Excel file
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel

  # PDF → Word document (Arabic + Latin, resumable)
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word

  # PDF → Word, verbose, custom output path
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf \
    --word \
    --output ./exports/contract.docx \
    --verbose

  # Enhance an existing conversion (retry pages with < 3 OCR sources, re-corroborate)
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word --enhance

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
     ├─ Step 1 — pdfjs text layer (all pages, structural)
     │
     └─ Per page (resumable — progress saved to .progress.json after each page):
          │
          ├─ Render page → JPEG (pdftoppm / mutool)
          │
          ├─ Steps 2–5 — Four OCR models in parallel
          │    ├─ DeepSeek-OCR   — visual OCR, structure-aware prompt
          │    ├─ dots.ocr       — visual OCR, structure-aware prompt
          │    ├─ GLM-OCR        — visual OCR, structure-aware prompt
          │    └─ Chandra-OCR    — visual OCR, structure-aware prompt
          │    All preserve Arabic (RTL) and Latin characters.
          │    Results appear in the terminal as each model finishes.
          │
          └─ Step 6 — Gemma4 corroboration
               Receives: page image + pdfjs text + all 4 OCR results (5 sources)
               Produces: single most accurate version of the page
               Rules:
                 - Content in 3+ sources → trusted, always included
                 - Content in 2 sources  → included unless image contradicts it
                 - Content in 1 source   → included only if image clearly confirms it
                 - Use page image as visual ground truth for disagreements

  → Word generator builds .docx from corroborated page texts
    Arabic paragraphs → bidirectional + right-aligned
    Latin paragraphs  → left-to-right
    Tables → tab-delimited rows reconstructed as Word tables
    Headings → Word Heading 2 style
    Each PDF page → new page in the Word document
  ```

  **Per-page terminal output:**

  ```
  [Word] Page 3/10:
         Rendering... done (91 KB)
         Extracting (4 OCR models in parallel):
           ✓ dots.ocr          143 chars  (1.9s)
           ✓ DeepSeek-OCR      141 chars  (2.2s)
           ✓ GLM-OCR           138 chars  (2.5s)
           ✓ Chandra-OCR       145 chars  (2.8s)
         Corroborating (5 sources)... done (1821 chars, 5.4s)
         Page 3 done in 10.1s
  ```

  ### Mode 3b — Enhance mode (`--word --enhance`)

  Runs against an **existing progress file** to improve pages where fewer than 3 OCR models produced substantial text (> 20 characters).

  ```
  Existing .progress.json
     │
     └─ For each completed page:
          │
          ├─ Count OCR sources with substantial text (> 20 chars)
          │
          ├─ If ≥ 3 sources → skip (already sufficient)
          │
          └─ If < 3 sources → enhance:
               ├─ Render page → JPEG
               ├─ Retry each weak model up to 3 times
               │    Stop retrying a model as soon as it returns substantial text.
               ├─ Re-corroborate with all available sources (best-case scenario)
               └─ Save updated progress immediately (resumable if interrupted)

  → Regenerate .docx from updated progress file
  ```

  **Sample enhance output:**

  ```
  [Enhance] Page 3/12: 1/4 OCR sources populated — enhancing...
           Rendering... done (88 KB)
           Retrying dots.ocr (attempt 1/3)...     ✓  134 chars
           Retrying GLM-OCR (attempt 1/3)...      ✗  8 chars (too short)
           Retrying GLM-OCR (attempt 2/3)...      ✓  119 chars
           Retrying Chandra-OCR (attempt 1/3)...  ✓  145 chars
           → 4/4 sources populated. Re-corroborating...
           ✓ Corroborated (1943 chars, 5.1s)

  [Enhance] Complete. 2 page(s) improved, 8 page(s) already had sufficient coverage.
  ```

  Enhance mode is safe to run multiple times — pages that already have sufficient coverage are skipped, and the progress file is updated incrementally so an interrupted enhance run can be resumed by re-running the same command.

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
  - **Enhance mode** — if some pages had low OCR coverage, run `--enhance` against the same command to retry weak models and re-corroborate without reprocessing pages that are already good.

  ```bash
  # First run (processes pages 1–20, interrupted at page 12)
  pnpm --filter @workspace/scripts run document-to-csv ./doc.pdf --word
  # → doc.progress.json contains pages 1–12

  # Second run (reads pages 1–12 from cache, resumes at page 13)
  pnpm --filter @workspace/scripts run document-to-csv ./doc.pdf --word

  # Enhance run (checks all completed pages, retries those with < 3 OCR sources)
  pnpm --filter @workspace/scripts run document-to-csv ./doc.pdf --word --enhance
  ```

  ---

  ## Architecture

  | Component | File | Role |
  |-----------|------|------|
  | CLI entry | `scripts/src/document-to-csv.ts` | Argument parsing, orchestration, mode routing (CSV / Excel / Word / Enhance) |
  | Types | `scripts/src/types.ts` | Zod schemas + TypeScript interfaces for all data shapes |
  | LM Studio client | `scripts/src/lm-studio-client.ts` | Typed OpenAI-compatible client factory |
  | OCR | `scripts/src/ocr.ts` | `callOcrModel()` — base64 encode + vision call; two system prompts (CSV mode, Word mode) |
  | PDF extraction | `scripts/src/pdf.ts` | pdfjs text layer + OCR pass + layout extraction for Excel + raw per-page texts for Word |
  | CSV generation | `scripts/src/csv-generator.ts` | `verifyDocumentWithGemma()` pre-pass + `generateCsvWithGemma()` agentic tool-use loop |
  | Excel generation | `scripts/src/excel-generator.ts` | `generateExcel()` — styled Data sheet + layout-faithful Document sheet |
  | PDF → Word pipeline | `scripts/src/pdf-to-word.ts` | 4 OCR models in parallel per page, Gemma4 corroboration, progress tracking, enhance mode |
  | Word generation | `scripts/src/word-generator.ts` | `generateWordDoc()` — docx with Arabic RTL, headings, tables, page breaks |

  ### Key decisions

  - **LM Studio via OpenAI SDK** — LM Studio exposes an OpenAI-compatible `/v1` endpoint, so the standard `openai` npm package is used directly. No custom HTTP layer needed.
  - **Three-pass PDF→CSV pipeline** — (1) pdfjs text layer for structure, (2) DeepSeek-OCR for visual verification, (3) Gemma4 reconciliation before CSV generation. This gives better results than any single source alone.
  - **4 OCR models + corroboration for PDF→Word** — DeepSeek-OCR, dots.ocr, GLM-OCR, and Chandra-OCR run in parallel on each rendered page image, then Gemma4 reconciles all four results against the pdfjs text layer (5 sources total). Running models in parallel keeps per-page latency close to the slowest single model rather than the sum of all. Content confirmed by 3+ sources is trusted outright; content from 2 sources is included unless the page image contradicts it; single-source content is checked against the image.
  - **Enhance mode** — an optional second pass over an existing progress file. Pages with fewer than 3 substantial OCR results are re-rendered and the weak models are retried up to 3 times each before re-corroborating. Progress is saved after every improved page, so enhance runs are themselves resumable.
  - **Per-page progress persistence** — the Word pipeline writes a JSON progress file after every page. On restart, already-completed pages are loaded from the file instead of being re-processed. This makes long conversions safe to interrupt. The schema uses `.default("")` on all optional fields so older progress files parse cleanly when new OCR model fields are added.
  - **Agentic tool-use loop (CSV/Excel)** — Gemma4 calls the `write_csv` tool up to 8 iterations. If it stops without calling the tool it is re-prompted automatically.
  - **Zod validation on every tool call** — `WriteCsvToolArgsSchema` validates Gemma4's `write_csv` arguments before accepting them.
  - **RFC 4180 CSV sanitizer** — the LLM's CSV output is re-parsed and re-serialized programmatically to guarantee correct quoting and consistent column counts.
  - **pdfjs-dist v5 legacy build** — the main ESM build requires DOM globals absent in Node.js. The legacy CJS build is loaded via `createRequire`. `GlobalWorkerOptions.workerSrc` must point to the actual worker file; pdfjs v5 removed the in-process fake worker.
  - **CTM-tracked image detection** — PDF image positions are derived by replaying the operator list and accumulating the current transformation matrix through `save`/`restore`/`transform` ops. Regions are cropped from the rendered raster using sharp.
  - **Arabic RTL in Word** — paragraphs and table cells containing Arabic Unicode characters are detected at render time and set `bidirectional: true` + `AlignmentType.RIGHT` + `rightToLeft: true` on text runs, producing correct RTL rendering in Microsoft Word without any manual markup.
  - **Thinking-block stripping** — Gemma4's chain-of-thought blocks (`<thinking>`, `<think>`, `<|channel>thought…<channel|>`) are stripped before corroborated text is saved to the progress file or written to the document.

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
  