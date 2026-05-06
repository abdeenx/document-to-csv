#!/usr/bin/env node
/**
 * document-to-csv
 *
 * Converts a document image or PDF into a structured CSV (or Excel) file.
 *
 * Image path:
 *   1. DeepSeek-OCR (vision) extracts text
 *   2. Gemma4 (tool use) structures it into CSV / Excel
 *      The original image is forwarded to Gemma4 for visual column-alignment verification.
 *
 * PDF path:
 *   1. pdfjs-dist extracts the embedded text layer (structural / positional)
 *      DeepSeek-OCR runs on each rendered page for a second visual pass
 *   2. Gemma4 (non-tool) reconciles the page images + both text layers into
 *      a single clean, tab-delimited text (visual verification step)
 *   3. Gemma4 (tool use) structures the verified text into CSV / Excel
 *      Page images are forwarded again so Gemma4 can verify column alignment
 *
 * Usage:
 *   pnpm --filter @workspace/scripts run document-to-csv <file> [options]
 *
 * Options:
 *   --output <path>           Output path (default: <basename>.csv or .xlsx with --excel)
 *   --excel                   Write an Excel (.xlsx) file instead of CSV;
 *                             embeds rendered PDF pages as image sheets
 *   --lm-studio-url <url>     LM Studio base URL (default: http://localhost:1234/v1)
 *   --ocr-model <id>          OCR model (images & PDF pages; default: mlx-community/DeepSeek-OCR-8bit)
 *   --structurer-model <id>   Structuring model (default: zecanard/...)
 *   --verbose                 Enable verbose logging
 */

import { writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { parseArgs } from "node:util";
import { CliArgsSchema } from "./types.js";
import { createLmStudioClient } from "./lm-studio-client.js";
import { extractTextWithOcr } from "./ocr.js";
import { extractTextFromPdf } from "./pdf.js";
import { verifyDocumentWithGemma, generateCsvWithGemma } from "./csv-generator.js";
import { generateExcel } from "./excel-generator.js";

const PDF_EXTENSION = ".pdf";
const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp"]);

function parseCliArgs(): ReturnType<typeof CliArgsSchema.parse> {
  const { values, positionals } = parseArgs({
    args: process.argv.slice(2),
    options: {
      output: { type: "string", short: "o" },
      excel: { type: "boolean", default: false },
      "lm-studio-url": { type: "string" },
      "ocr-model": { type: "string" },
      "structurer-model": { type: "string" },
      verbose: { type: "boolean", short: "v", default: false },
      help: { type: "boolean", short: "h", default: false },
    },
    allowPositionals: true,
  });

  if (values.help || positionals.length === 0) {
    printUsage();
    process.exit(values.help ? 0 : 1);
  }

  const inputPath = positionals[0]!;
  const ext = extname(inputPath).toLowerCase();

  if (ext !== PDF_EXTENSION && !IMAGE_EXTENSIONS.has(ext)) {
    console.error(
      `Unsupported file type "${ext}". Supported: pdf, jpg, jpeg, png, gif, webp`,
    );
    process.exit(1);
  }

  const inputBase = basename(inputPath, extname(inputPath));
  // INIT_CWD is set by pnpm/npm to the directory where the user ran the command.
  const callerCwd = process.env["INIT_CWD"] ?? process.cwd();
  const outputExt = values.excel ? ".xlsx" : ".csv";
  const defaultOutput = resolve(callerCwd, `${inputBase}${outputExt}`);

  const raw = {
    imagePath: resolve(callerCwd, inputPath),
    outputPath: values.output ? resolve(callerCwd, values.output) : defaultOutput,
    lmStudioUrl: values["lm-studio-url"] ?? "http://localhost:1234/v1",
    ocrModel: values["ocr-model"] ?? "mlx-community/DeepSeek-OCR-8bit",
    structurerModel:
      values["structurer-model"] ??
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    verbose: values.verbose ?? false,
    excel: values.excel ?? false,
  };

  const result = CliArgsSchema.safeParse(raw);
  if (!result.success) {
    console.error("Invalid arguments:");
    for (const issue of result.error.issues) {
      console.error(`  ${issue.path.join(".")}: ${issue.message}`);
    }
    process.exit(1);
  }

  return result.data;
}

function printUsage(): void {
  console.log(`
document-to-csv — Convert a document image or PDF to a structured CSV or Excel file

Usage:
  pnpm --filter @workspace/scripts run document-to-csv <file> [options]

Arguments:
  <file>                       Path to the input file
                               Images: jpg, jpeg, png, gif, webp
                                 → OCR via DeepSeek-OCR, then structured by Gemma4
                               PDF:    pdf
                                 → pdfjs text layer + DeepSeek-OCR per page,
                                   then Gemma4 visual verification + structuring

Options:
  -o, --output <path>          Output file path
                               Default: <basename>.csv (or .xlsx with --excel)
  --excel                      Write an Excel (.xlsx) file instead of CSV
                               PDF pages are embedded as image sheets (requires
                               pdftoppm or mutool for PDF rendering)
  --lm-studio-url <url>        LM Studio base URL (default: http://localhost:1234/v1)
  --ocr-model <id>             OCR model ID (default: mlx-community/DeepSeek-OCR-8bit)
                               Used for images and for each rendered PDF page.
  --structurer-model <id>      Structuring model ID
                               (default: zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine)
  -v, --verbose                Enable step-by-step logging
  -h, --help                   Show this help message

PDF renderer (required for OCR + --excel image embedding):
  brew install poppler          # provides pdftoppm  ← recommended
  brew install mupdf-tools      # provides mutool    ← fallback

Examples:
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --output ./data/report.xlsx --excel
  pnpm --filter @workspace/scripts run document-to-csv ./screenshot.png --lm-studio-url http://192.168.1.10:1234/v1
`);
}

async function main(): Promise<void> {
  const args = parseCliArgs();

  const isPdf = extname(args.imagePath).toLowerCase() === PDF_EXTENSION;

  // For a PDF with images: extract(1) → verify(2) → generate(3)
  // For a PDF text-only or image: extract(1) → generate(2)
  // We don't know if images are available until after extraction, so we use
  // dynamic step labels printed after extraction.

  console.log("document-to-csv");
  console.log("================");
  console.log(`  Input:        ${args.imagePath}`);
  console.log(`  Output:       ${args.outputPath}  ${args.excel ? "(Excel)" : "(CSV)"}`);
  if (isPdf) {
    console.log(`  Mode:         PDF  →  pdfjs text layer + DeepSeek-OCR on rendered pages`);
    console.log(`  OCR model:    ${args.ocrModel}  (requires pdftoppm or mutool)`);
  } else {
    console.log(`  Mode:         Image  →  DeepSeek-OCR`);
    console.log(`  OCR model:    ${args.ocrModel}`);
  }
  console.log(`  Struct model: ${args.structurerModel}`);
  console.log(`  LM Studio:    ${args.lmStudioUrl}`);
  console.log("");

  const client = createLmStudioClient({
    baseUrl: args.lmStudioUrl,
    apiKey: "lm-studio",
  });

  // ── Step 1: Extract ────────────────────────────────────────────────────────
  console.log(
    isPdf
      ? "Step 1 — Extracting text from PDF (pdfjs + DeepSeek-OCR)..."
      : "Step 1 — Extracting text with DeepSeek-OCR...",
  );

  let ocrResult = isPdf
    ? await extractTextFromPdf(args.imagePath, args.verbose, client, args.ocrModel)
    : await extractTextWithOcr(client, args.imagePath, args.ocrModel, args.verbose);

  if (args.verbose) {
    console.log("\n--- EXTRACTED TEXT ---");
    console.log(ocrResult.rawText);
    console.log("--- END EXTRACTED TEXT ---\n");
  }

  console.log(`  Done. Extracted ${ocrResult.rawText.length} characters.`);
  console.log("");

  // ── Step 2 (PDF + images only): Visual verification ────────────────────────
  const hasPdfImages = isPdf && (ocrResult.pageImages?.length ?? 0) > 0;

  if (hasPdfImages) {
    console.log("Step 2 — Visual verification with Gemma4 (reconciling image + text layers)...");
    const verifiedText = await verifyDocumentWithGemma(
      client,
      ocrResult,
      args.structurerModel,
      args.verbose,
    );
    ocrResult = { ...ocrResult, rawText: verifiedText };
    console.log(`  Done. Verified text: ${verifiedText.length} characters.`);
    console.log("");
  }

  // ── Step 2 or 3: Generate CSV via Gemma4 tool-use ─────────────────────────
  const csvStepNum = hasPdfImages ? 3 : 2;
  console.log(
    `Step ${csvStepNum} — Generating ${args.excel ? "Excel" : "CSV"} structure with Gemma4 tool use...`,
  );

  const csvResult = await generateCsvWithGemma(
    client,
    ocrResult,
    args.structurerModel,
    args.outputPath!,
    args.verbose,
  );

  // ── Write output ───────────────────────────────────────────────────────────
  if (args.excel) {
    await generateExcel({
      csvContent: csvResult.csvContent,
      outputPath: csvResult.outputPath,
      ocrResult,
      verbose: args.verbose,
    });
  } else {
    await writeFile(csvResult.outputPath, csvResult.csvContent, "utf-8");
  }

  const dataRows = csvResult.csvContent.split("\n").filter(Boolean).length - 1;

  console.log("");
  console.log("Done!");
  console.log(
    `  ${args.excel ? "Excel" : "CSV"} saved to: ${csvResult.outputPath}`,
  );
  console.log(`  Rows: ${dataRows} (excluding header)`);
  if (args.verbose) {
    console.log(`  Reasoning: ${csvResult.reasoning}`);
  }
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error(`\nError: ${message}`);
  if (err instanceof Error && err.stack && process.env["VERBOSE"]) {
    console.error(err.stack);
  }
  process.exit(1);
});
