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
 * With --excel on a PDF:
 *   A layout-faithful "Document" sheet is added to the workbook. Text items
 *   are placed at their PDF-grid positions, and any images detected in the PDF
 *   operator list are cropped from the rendered page renders and embedded at
 *   the matching cell anchor.
 *
 * Usage:
 *   pnpm --filter @workspace/scripts run document-to-csv <file> [options]
 */

import { basename, extname, resolve } from "node:path";
import { writeFile } from "node:fs/promises";
import { parseArgs } from "node:util";
import { CliArgsSchema } from "./types.js";
import { createLmStudioClient } from "./lm-studio-client.js";
import { extractTextWithOcr } from "./ocr.js";
import { extractTextFromPdf, extractPdfPageLayouts } from "./pdf.js";
import { verifyDocumentWithGemma, generateCsvWithGemma } from "./csv-generator.js";
import { generateExcel } from "./excel-generator.js";

const PDF_EXTENSION    = ".pdf";
const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp"]);

function parseCliArgs(): ReturnType<typeof CliArgsSchema.parse> {
  const { values, positionals } = parseArgs({
    args: process.argv.slice(2),
    options: {
      output:             { type: "string",  short: "o" },
      excel:              { type: "boolean", default: false },
      "lm-studio-url":    { type: "string" },
      "ocr-model":        { type: "string" },
      "structurer-model": { type: "string" },
      verbose:            { type: "boolean", short: "v", default: false },
      help:               { type: "boolean", short: "h", default: false },
    },
    allowPositionals: true,
  });

  if (values.help || positionals.length === 0) {
    printUsage();
    process.exit(values.help ? 0 : 1);
  }

  const inputPath = positionals[0]!;
  const ext       = extname(inputPath).toLowerCase();

  if (ext !== PDF_EXTENSION && !IMAGE_EXTENSIONS.has(ext)) {
    console.error(`Unsupported file type "${ext}". Supported: pdf, jpg, jpeg, png, gif, webp`);
    process.exit(1);
  }

  const inputBase = basename(inputPath, extname(inputPath));
  const callerCwd = process.env["INIT_CWD"] ?? process.cwd();
  const outputExt = values.excel ? ".xlsx" : ".csv";
  const defaultOutput = resolve(callerCwd, `${inputBase}${outputExt}`);

  const raw = {
    imagePath:       resolve(callerCwd, inputPath),
    outputPath:      values.output ? resolve(callerCwd, values.output) : defaultOutput,
    lmStudioUrl:     values["lm-studio-url"]    ?? "http://localhost:1234/v1",
    ocrModel:        values["ocr-model"]         ?? "mlx-community/DeepSeek-OCR-8bit",
    structurerModel: values["structurer-model"]  ??
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    verbose: values.verbose ?? false,
    excel:   values.excel   ?? false,
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
  --excel                      Write an Excel (.xlsx) file instead of CSV.
                               For PDF inputs the workbook contains:
                                 "Data"      — clean structured table
                                 "Document"  — layout-faithful sheet where
                                               text lands at its PDF position
                                               and PDF images are embedded at
                                               the matching cell location
  --lm-studio-url <url>        LM Studio base URL (default: http://localhost:1234/v1)
  --ocr-model <id>             OCR model (default: mlx-community/DeepSeek-OCR-8bit)
  --structurer-model <id>      Structuring model
                               (default: zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine)
  -v, --verbose                Enable step-by-step logging
  -h, --help                   Show this help message

PDF renderer (required for OCR + image embedding in --excel mode):
  brew install poppler          # provides pdftoppm  ← recommended
  brew install mupdf-tools      # provides mutool    ← fallback

Examples:
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --output ./data/report.xlsx --excel
`);
}

async function main(): Promise<void> {
  const args  = parseCliArgs();
  const isPdf = extname(args.imagePath).toLowerCase() === PDF_EXTENSION;

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

  const client = createLmStudioClient({ baseUrl: args.lmStudioUrl, apiKey: "lm-studio" });

  // ── Step 1: Extract text ──────────────────────────────────────────────────
  console.log(
    isPdf
      ? "Step 1 — Extracting text from PDF (pdfjs + DeepSeek-OCR)..."
      : "Step 1 — Extracting text with DeepSeek-OCR...",
  );

  // Run extraction and (if PDF + excel) layout analysis in parallel.
  // extractPdfPageLayouts does a second, independent pdfjs load — this is
  // intentional so the two paths remain decoupled. pdfjs loading is fast.
  const [ocrResultRaw, pageLayouts] = await Promise.all([
    isPdf
      ? extractTextFromPdf(args.imagePath, args.verbose, client, args.ocrModel)
      : extractTextWithOcr(client, args.imagePath, args.ocrModel, args.verbose),
    isPdf && args.excel
      ? extractPdfPageLayouts(args.imagePath, args.verbose)
      : Promise.resolve(undefined),
  ]);

  let ocrResult = ocrResultRaw;

  if (args.verbose) {
    console.log("\n--- EXTRACTED TEXT ---");
    console.log(ocrResult.rawText);
    console.log("--- END EXTRACTED TEXT ---\n");
  }
  console.log(`  Done. Extracted ${ocrResult.rawText.length} characters.`);
  console.log("");

  // ── Step 2 (PDF + images only): Visual verification ───────────────────────
  const hasPdfImages = isPdf && (ocrResult.pageImages?.length ?? 0) > 0;

  if (hasPdfImages) {
    console.log("Step 2 — Visual verification with Gemma4 (reconciling image + text layers)...");
    const verifiedText = await verifyDocumentWithGemma(
      client, ocrResult, args.structurerModel, args.verbose,
    );
    ocrResult = { ...ocrResult, rawText: verifiedText };
    console.log(`  Done. Verified text: ${verifiedText.length} characters.`);
    console.log("");
  }

  // ── Step 2 or 3: Generate CSV ─────────────────────────────────────────────
  const csvStep = hasPdfImages ? 3 : 2;
  console.log(`Step ${csvStep} — Generating ${args.excel ? "Excel" : "CSV"} structure with Gemma4...`);

  const csvResult = await generateCsvWithGemma(
    client, ocrResult, args.structurerModel, args.outputPath!, args.verbose,
  );

  // ── Write output ──────────────────────────────────────────────────────────
  if (args.excel) {
    await generateExcel({
      csvContent:  csvResult.csvContent,
      outputPath:  csvResult.outputPath,
      ocrResult,
      pageLayouts: pageLayouts ?? undefined,
      verbose:     args.verbose,
    });
  } else {
    await writeFile(csvResult.outputPath, csvResult.csvContent, "utf-8");
  }

  const dataRows = csvResult.csvContent.split("\n").filter(Boolean).length - 1;

  console.log("");
  console.log("Done!");
  console.log(`  ${args.excel ? "Excel" : "CSV"} saved to: ${csvResult.outputPath}`);
  console.log(`  Rows: ${dataRows} (excluding header)`);
  if (args.verbose) console.log(`  Reasoning: ${csvResult.reasoning}`);
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  console.error(`\nError: ${message}`);
  if (err instanceof Error && err.stack && process.env["VERBOSE"]) console.error(err.stack);
  process.exit(1);
});
