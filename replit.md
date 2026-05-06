# Document-to-CSV / Excel Harness

A typesafe TypeScript CLI that uses DeepSeek-OCR (vision) + Gemma4 (tool use) via LM Studio to extract text from document images and PDFs, then produce structured CSV or Excel files.

## Run & Operate

- `pnpm --filter @workspace/scripts run document-to-csv <file> [options]` — main CLI
- `pnpm --filter @workspace/scripts run typecheck` — typecheck scripts package
- `pnpm run typecheck` — full typecheck across all packages

Key flags:
- `--excel` — write `.xlsx` instead of `.csv`; embeds rendered PDF pages as image sheets
- `--verbose` — step-by-step logging
- `--output <path>` — custom output path

PDF renderer (required for OCR + `--excel` image embedding):
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
- Validation: Zod (`zod/v4`)

## Where things live

- `scripts/src/types.ts` — all Zod schemas and TypeScript types (including `CliArgsSchema`)
- `scripts/src/lm-studio-client.ts` — typed OpenAI client factory
- `scripts/src/ocr.ts` — `callOcrModel()` + `extractTextWithOcr()`; base64 encoding + DeepSeek-OCR call
- `scripts/src/pdf.ts` — pdfjs text layer extraction + DeepSeek-OCR pass on rendered pages
- `scripts/src/csv-generator.ts` — `verifyDocumentWithGemma()` pre-pass + `generateCsvWithGemma()` tool-use loop + RFC 4180 sanitizer
- `scripts/src/excel-generator.ts` — `generateExcel()`: styled data sheet + per-page image sheets via exceljs
- `scripts/src/document-to-csv.ts` — CLI entry; routes PDF vs image; orchestrates all steps

## Architecture decisions

- **LM Studio via OpenAI SDK**: OpenAI-compatible `/v1` endpoint; no custom HTTP layer needed.
- **Three-model-call PDF pipeline**: (1) DeepSeek-OCR renders+extracts each page visually; (2) Gemma4 `verifyDocumentWithGemma` reconciles page images + pdfjs text + OCR text into one clean tab-delimited text; (3) Gemma4 tool-use loop generates final CSV/Excel.
- **Agentic tool-use loop**: Gemma4 runs up to 8 iterations calling `write_csv`. Re-prompted automatically if it stops without calling the tool.
- **Zod for tool argument validation**: `WriteCsvToolArgsSchema` validates every `write_csv` call before accepting it.
- **pdfjs-dist v5 worker**: `workerSrc` must point to the actual legacy worker file (`_require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs")`). Empty string no longer works in v5 — it removed the in-process fake worker.
- **Excel image embedding**: `exceljs` `addImage()` with `ext: { width, height }` anchored at `tl: { col: 0, row: 0 }`; one worksheet per PDF page.

## Product

CLI harness: pass an image or PDF, get a structured `.csv` or styled `.xlsx` file. Excel output embeds the original rendered document pages alongside the extracted data table. Supports invoices, tables, reports, forms, spreadsheet screenshots, and multi-page PDFs.

## User preferences

- Models served by LM Studio locally:
  - OCR: `mlx-community/DeepSeek-OCR-8bit`
  - Structurer: `zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine`
- Default LM Studio URL: `http://localhost:1234/v1`

## Gotchas

- Both models must be loaded in LM Studio before running
- Gemma4 must support tool use / function calling (verify model template in LM Studio)
- `pdftoppm` or `mutool` required for PDF page rendering (OCR pass + Excel image embed)
- pdfjs-dist v5: `workerSrc` must be a real file path, not `""`

## Pointers

- `pnpm-workspace` skill — workspace structure and TypeScript setup
- LM Studio docs: https://lm-studio.ai/docs
