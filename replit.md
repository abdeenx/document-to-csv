# Document-to-CSV / Excel / Word Harness

A typesafe TypeScript CLI that uses DeepSeek-OCR (vision) + Gemma4 (tool use) via LM Studio to extract text from document images and PDFs, then produce structured CSV, Excel, or Word files. Supports Arabic (RTL) and Latin text.

## Run & Operate

- `pnpm --filter @workspace/scripts run document-to-csv <file> [options]` — main CLI
- `pnpm --filter @workspace/scripts run typecheck` — typecheck scripts package
- `pnpm run typecheck` — full typecheck across all packages

Key flags:
- `--excel` — write `.xlsx`; embeds rendered PDF pages as image sheets
- `--word` — write `.docx` from PDF; 4-pass per-page extraction + corroboration; Arabic RTL; resumable
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

## Where things live

- `scripts/src/types.ts` — all Zod schemas and TypeScript types (incl. `CliArgsSchema`, `WordProgressSchema`, `PageExtraction`)
- `scripts/src/lm-studio-client.ts` — typed OpenAI client factory
- `scripts/src/ocr.ts` — `callOcrModel()` with optional system prompt override; `OCR_SYSTEM_PROMPT_CSV` + `OCR_SYSTEM_PROMPT_WORD`
- `scripts/src/pdf.ts` — pdfjs extraction + OCR pass + `extractPdfjsPageTextsRaw()` (word) + `extractPdfPageLayouts()` (excel)
- `scripts/src/csv-generator.ts` — `verifyDocumentWithGemma()` + `generateCsvWithGemma()` tool-use loop + RFC 4180 sanitizer
- `scripts/src/excel-generator.ts` — `generateExcel()`: styled data sheet + layout-faithful document sheet
- `scripts/src/pdf-to-word.ts` — 4-pass per-page pipeline (pdfjs/OCR/Gemma4/corroborate) + progress tracking
- `scripts/src/word-generator.ts` — `generateWordDoc()`: docx with Arabic RTL, headings, tables, page breaks
- `scripts/src/document-to-csv.ts` — CLI entry; routes PDF/image and --excel/--word/--csv

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
