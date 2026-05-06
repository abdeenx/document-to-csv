#!/usr/bin/env node
/**
 * document-to-csv
 *
 * Converts a document image or PDF into a structured CSV file.
 *
 * Image path: DeepSeek-OCR (vision) extracts text, then Gemma4 (tool use)
 *             structures it into CSV. The image is also sent to Gemma4 so it
 *             can correct any OCR column-alignment errors visually.
 *
 * PDF path:   pdfjs-dist extracts the embedded text layer (no LLM needed for
 *             extraction), then Gemma4 structures it into CSV. Works for
 *             text-based PDFs; scanned PDFs should be exported as images first.
 *
 * Usage:
 *   pnpm --filter @workspace/scripts run document-to-csv <file> [options]
 *
 * Options:
 *   --output <path>           Output CSV path (default: <basename>.csv)
 *   --lm-studio-url <url>     LM Studio base URL (default: http://localhost:1234/v1)
 *   --ocr-model <id>          OCR model (images only; default: mlx-community/DeepSeek-OCR-8bit)
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
import { generateCsvWithGemma } from "./csv-generator.js";

const PDF_EXTENSION = ".pdf";
const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp"]);

function parseCliArgs(): ReturnType<typeof CliArgsSchema.parse> {
  const { values, positionals } = parseArgs({
    args: process.argv.slice(2),
    options: {
      output: { type: "string", short: "o" },
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
  // INIT_CWD is set by pnpm/npm to the directory where the user ran the command,
  // regardless of which package directory the script executes from.
  const callerCwd = process.env["INIT_CWD"] ?? process.cwd();
  const defaultOutput = resolve(callerCwd, `${inputBase}.csv`);

  const raw = {
    imagePath: resolve(callerCwd, inputPath),
    outputPath: values.output ? resolve(callerCwd, values.output) : defaultOutput,
    lmStudioUrl: values["lm-studio-url"] ?? "http://localhost:1234/v1",
    ocrModel: values["ocr-model"] ?? "mlx-community/DeepSeek-OCR-8bit",
    structurerModel:
      values["structurer-model"] ??
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    verbose: values.verbose ?? false,
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
document-to-csv — Convert a document image or PDF to a structured CSV

Usage:
  pnpm --filter @workspace/scripts run document-to-csv <file> [options]

Arguments:
  <file>                       Path to the input file
                               Images: jpg, jpeg, png, gif, webp
                                 → OCR via DeepSeek-OCR, then structured by Gemma4
                               PDF:    pdf
                                 → Text extracted directly, then structured by Gemma4
                                 → Scanned PDFs: export pages as images first

Options:
  -o, --output <path>          Output CSV file path (default: <basename>.csv)
  --lm-studio-url <url>        LM Studio base URL (default: http://localhost:1234/v1)
  --ocr-model <id>             OCR model ID for images (default: mlx-community/DeepSeek-OCR-8bit)
  --structurer-model <id>      Structuring model ID
                               (default: zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine)
  -v, --verbose                Enable step-by-step logging
  -h, --help                   Show this help message

Examples:
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --output ./data/report.csv --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./screenshot.png --lm-studio-url http://192.168.1.10:1234/v1
`);
}

async function main(): Promise<void> {
  const args = parseCliArgs();

  const isPdf =
    extname(args.imagePath).toLowerCase() === PDF_EXTENSION;

  console.log("document-to-csv");
  console.log("================");
  console.log(`  Input:        ${args.imagePath}`);
  console.log(`  Output:       ${args.outputPath}`);
  console.log(`  Mode:         ${isPdf ? "PDF (pdfjs text extraction)" : "Image (DeepSeek-OCR)"}`);
  if (!isPdf) {
    console.log(`  OCR model:    ${args.ocrModel}`);
  }
  console.log(`  Struct model: ${args.structurerModel}`);
  console.log(`  LM Studio:    ${args.lmStudioUrl}`);
  console.log("");

  const client = createLmStudioClient({
    baseUrl: args.lmStudioUrl,
    apiKey: "lm-studio",
  });

  let step1Label: string;
  if (isPdf) {
    step1Label = "Step 1/2 — Extracting text from PDF (pdfjs)...";
  } else {
    step1Label = "Step 1/2 — Extracting text with DeepSeek-OCR...";
  }
  console.log(step1Label);

  const ocrResult = isPdf
    ? await extractTextFromPdf(args.imagePath, args.verbose)
    : await extractTextWithOcr(client, args.imagePath, args.ocrModel, args.verbose);

  if (args.verbose) {
    console.log("\n--- EXTRACTED TEXT ---");
    console.log(ocrResult.rawText);
    console.log("--- END EXTRACTED TEXT ---\n");
  }

  console.log(`  Done. Extracted ${ocrResult.rawText.length} characters.`);
  console.log("");

  console.log("Step 2/2 — Generating CSV structure with Gemma4 tool use...");
  const csvResult = await generateCsvWithGemma(
    client,
    ocrResult,
    args.structurerModel,
    args.outputPath!,
    args.verbose,
  );

  await writeFile(csvResult.outputPath, csvResult.csvContent, "utf-8");

  console.log("");
  console.log("Done!");
  console.log(`  CSV saved to: ${csvResult.outputPath}`);
  console.log(
    `  Rows: ${csvResult.csvContent.split("\n").filter(Boolean).length - 1} (excluding header)`,
  );
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
