# Document-to-CSV / Excel / Word Harness

A typesafe TypeScript CLI that uses DeepSeek-OCR (vision) + Gemma4 (tool use) via LM Studio to extract text from document images and PDFs, then produce structured CSV, Excel, or Word files. Supports Arabic (RTL) and Latin text.

## Run & Operate

- `pnpm --filter @workspace/scripts run document-to-csv <file> [options]` — main CLI
- `pnpm --filter @workspace/scripts run typecheck` — typecheck scripts package
- `pnpm run typecheck` — full typecheck across all packages

Key flags:
- `--excel` — write `.xlsx`; embeds rendered PDF pages as image sheets
- `--word` — write `.docx` from PDF; 4 OCR models in parallel + Gemma4 corroboration; Arabic RTL; resumable
- `--enhance` — re-run OCR on weak pages in an existing `--word` progress file, then regenerate docx
- `--qwen-word` — write `.docx` from PDF using a **single Qwen2.5-VL call per page**; no corroboration needed; Arabic RTL; resumable
- `--verbose` — step-by-step logging
- `--output <path>` — custom output path

PDF renderer (required for OCR + `--excel` image embedding + `--word`):
```
brew install poppler      # pdftoppm (recommended)
brew install mupdf-tools  # mutool (fallback)
```

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- LLM client: `openai` SDK (LM Studio OpenAI-compatible endpoint)
- PDF: `pdfjs-dist` v5 (legacy build via `createRequire`)
- Image resize: `sharp`
- Excel output: `exceljs`
- Word output: `docx` v9
- Validation: Zod (`zod/v4`)

## Pipelines at a glance

| Flag | Models required | Output |
|------|----------------|--------|
| *(none)* | DeepSeek-OCR + Gemma4 | CSV |
| `--excel` | DeepSeek-OCR + Gemma4 | XLSX |
| `--word` | DeepSeek-OCR + dots.ocr + GLM-OCR + Chandra-OCR + Gemma4 | DOCX (4-model corroboration) |
| `--qwen-word` | Qwen2.5-VL-7B-Instruct-8bit | DOCX (single-model, fast) |

## Where things live

- `scripts/src/types.ts` — all Zod schemas and TypeScript types (incl. `CliArgsSchema`, `WordProgressSchema`, `QwenProgressSchema`, `PageExtraction`)
- `scripts/src/lm-studio-client.ts` — typed OpenAI client factory
- `scripts/src/ocr.ts` — `callOcrModel()` with optional system prompt override; `stripThinking()` (exported); `OCR_SYSTEM_PROMPT_CSV` + `OCR_SYSTEM_PROMPT_WORD`
- `scripts/src/pdf.ts` — pdfjs extraction + OCR pass + `extractPdfjsPageTextsRaw()` (word) + `extractPdfPageLayouts()` (excel)
- `scripts/src/csv-generator.ts` — `verifyDocumentWithGemma()` + `generateCsvWithGemma()` tool-use loop + RFC 4180 sanitizer
- `scripts/src/excel-generator.ts` — `generateExcel()`: styled data sheet + layout-faithful document sheet
- `scripts/src/pdf-to-word.ts` — 4-OCR-model parallel pipeline + Gemma4 corroboration + progress tracking (`--word`)
- `scripts/src/qwen-pdf-to-word.ts` — single-model Qwen2.5-VL pipeline + progress tracking (`--qwen-word`)
- `scripts/src/word-generator.ts` — `generateWordDoc()`: docx with Arabic RTL, headings, tables, page breaks
- `scripts/src/document-to-csv.ts` — CLI entry; routes PDF/image and --excel/--word/--qwen-word/--csv

## Architecture decisions

- **LM Studio via OpenAI SDK**: OpenAI-compatible `/v1` endpoint; no custom HTTP layer needed.
- **CSV/Excel PDF pipeline (3 passes)**: (1) pdfjs text layer; (2) DeepSeek-OCR per page; (3) Gemma4 reconciliation → tool-use CSV generation.
- **Word PDF pipeline (4 passes per page)**: pdfjs text + DeepSeek-OCR + Gemma4 direct vision + Gemma4 corroboration. Content in 2+ sources trusted; single-source content verified against page image.
- **Resumable Word conversion**: progress written to `<output>.progress.json` after each page. Re-running skips completed pages.
- **Arabic RTL in Word**: paragraphs/cells with Arabic Unicode get `bidirectional: true` + `AlignmentType.RIGHT` + `rightToLeft: true` on text runs.
- **pdfjs-dist v5 worker**: `workerSrc` must point to actual legacy worker file via `_require.resolve(...)`.

## Product

CLI harness: pass an image or PDF → get `.csv`, `.xlsx`, or `.docx`. Word output preserves document structure with Arabic + Latin support and is resumable on interruption.

## User preferences

- Models served by LM Studio locally:
  - OCR: `mlx-community/DeepSeek-OCR-8bit`
  - dots.ocr: `mlx-community/dots.ocr-bf16`
  - GLM-OCR: `mlx-community/GLM-OCR-bf16`
  - Chandra-OCR: `jwindle47/chandra-ocr-2-8bit-mlx`
  - Qwen VLM: `mlx-community/Qwen2.5-VL-7B-Instruct-8bit`
  - Structurer: `zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine`
- Default LM Studio URL: `http://localhost:1234/v1`

## Gotchas

- Both models must be loaded in LM Studio before running
- Gemma4 must support tool use / function calling (for CSV/Excel mode)
- `pdftoppm` or `mutool` required for PDF page rendering (all OCR passes)
- pdfjs-dist v5: `workerSrc` must be a real file path, not `""`
- `--word` and `--excel` are mutually exclusive; `--word` requires PDF input

## Pointers

- `pnpm-workspace` skill — workspace structure and TypeScript setup
- LM Studio docs: https://lm-studio.ai/docs
