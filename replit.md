# Document-to-CSV Harness

A typesafe TypeScript CLI that uses DeepSeek-OCR (vision) + Gemma4 (tool use) via LM Studio to extract text from document images and convert them into structured CSV files.

## Run & Operate

- `pnpm --filter @workspace/scripts run document-to-csv <file> [options]` — main CLI (image or PDF)
- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from the OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- Required env: `DATABASE_URL` — Postgres connection string

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API: Express 5
- DB: PostgreSQL + Drizzle ORM
- Validation: Zod (`zod/v4`), `drizzle-zod`
- API codegen: Orval (from OpenAPI spec)
- Build: esbuild (CJS bundle)
- LLM client: `openai` SDK (LM Studio OpenAI-compatible endpoint)

## Where things live

- `scripts/src/types.ts` — all Zod schemas and TypeScript types
- `scripts/src/lm-studio-client.ts` — typed OpenAI client factory pointing at LM Studio
- `scripts/src/ocr.ts` — base64 image encoding + DeepSeek-OCR vision call; returns preprocessed image for Gemma4 visual pass
- `scripts/src/pdf.ts` — pdfjs-dist text extraction; reconstructs rows by Y-position grouping
- `scripts/src/csv-generator.ts` — Gemma4 tool-use agentic loop (`write_csv` tool) + RFC 4180 sanitizer
- `scripts/src/document-to-csv.ts` — CLI entry point; routes by extension (.pdf vs image)

## Architecture decisions

- **LM Studio via OpenAI SDK**: LM Studio exposes an OpenAI-compatible `/v1` endpoint, so the standard `openai` npm package is used directly with a custom `baseURL`. No custom HTTP layer needed.
- **Two-model pipeline**: DeepSeek-OCR handles vision/extraction (temperature 0, layout-preserving prompt); Gemma4 handles structure inference via tool use (separate concerns, separate models).
- **Agentic tool-use loop**: The Gemma4 step runs up to 8 iterations. It calls the `write_csv` tool once it has reasoned about the document structure. If it stops without calling the tool it is re-prompted automatically.
- **Zod for tool argument validation**: Gemma4's `write_csv` tool call arguments are parsed through `WriteCsvToolArgsSchema` before being accepted, ensuring the CSV content is never silently malformed.
- **Union type narrowing for tool calls**: OpenAI SDK's `ChatCompletionMessageToolCall` is a union; tool calls are narrowed via `type === "function"` guard before accessing `.function.name/.arguments`.
- **Dual input modes**: images go through DeepSeek-OCR + Gemma4 vision (image forwarded to Gemma4 for column-alignment verification); PDFs go through pdfjs-dist text extraction + Gemma4 text-only (no OCR needed).
- **pdfjs-dist legacy build via createRequire**: main ESM build requires DOM globals absent in Node.js; loading the legacy CJS build via `createRequire` works in Node.js 22+ which allows `require()` of `.mjs` modules.

## Product

CLI harness: pass an image or PDF, get a `.csv` file whose structure mirrors the original document. Supports invoices, tables, reports, forms, spreadsheet screenshots, and multi-page PDF reports.

## User preferences

- Models served by LM Studio locally:
  - OCR: `mlx-community/DeepSeek-OCR-8bit`
  - Structurer: `zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine`
- Default LM Studio URL: `http://localhost:1234/v1`

## Gotchas

- Both models must be loaded in LM Studio before running the CLI
- The Gemma4 model must support tool use / function calling — verify in LM Studio that the model template includes tool-use support
- LM Studio must have the vision model loaded when running OCR; the structurer model can be swapped in or both can be loaded simultaneously if VRAM allows

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
- LM Studio docs: https://lm-studio.ai/docs
