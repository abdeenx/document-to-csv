#!/usr/bin/env node
/**
 * document-to-csv
 *
 * Takes a document image, extracts text via DeepSeek-OCR (LM Studio),
 * then uses Gemma4's tool-use to produce a CSV that mirrors the document structure.
 *
 * Usage:
 *   pnpm --filter @workspace/scripts run document-to-csv <image-path> [options]
 *
 * Options:
 *   --output <path>           Output CSV path (default: <image-basename>.csv)
 *   --lm-studio-url <url>     LM Studio base URL (default: http://localhost:1234/v1)
 *   --ocr-model <id>          OCR model ID (default: mlx-community/DeepSeek-OCR-8bit)
 *   --structurer-model <id>   Structurer model ID (default: zecanard/...)
 *   --verbose                 Enable verbose logging
 */

import { writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { parseArgs } from "node:util";
import { CliArgsSchema } from "./types.js";
import { createLmStudioClient } from "./lm-studio-client.js";
import { extractTextWithOcr } from "./ocr.js";
import { generateCsvWithGemma } from "./csv-generator.js";

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

  const imagePath = positionals[0]!;
  const imageBase = basename(imagePath, extname(imagePath));
  const defaultOutput = resolve(process.cwd(), `${imageBase}.csv`);

  const raw = {
    imagePath: resolve(imagePath),
    outputPath: values.output ? resolve(values.output) : defaultOutput,
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
document-to-csv — Convert a document image to a structured CSV

Usage:
  pnpm --filter @workspace/scripts run document-to-csv <image-path> [options]

Arguments:
  <image-path>                 Path to the image file (jpg, jpeg, png, gif, webp)

Options:
  -o, --output <path>          Output CSV file path (default: <image-basename>.csv)
  --lm-studio-url <url>        LM Studio base URL (default: http://localhost:1234/v1)
  --ocr-model <id>             OCR model ID (default: mlx-community/DeepSeek-OCR-8bit)
  --structurer-model <id>      Structurer model ID
                               (default: zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine)
  -v, --verbose                Enable verbose step-by-step logging
  -h, --help                   Show this help message

Examples:
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png
  pnpm --filter @workspace/scripts run document-to-csv ./report.jpg --output ./data/report.csv --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./screenshot.png --lm-studio-url http://192.168.1.10:1234/v1
`);
}

async function main(): Promise<void> {
  const args = parseCliArgs();

  console.log("document-to-csv");
  console.log("================");
  console.log(`  Image:      ${args.imagePath}`);
  console.log(`  Output:     ${args.outputPath}`);
  console.log(`  OCR model:  ${args.ocrModel}`);
  console.log(`  Struct model: ${args.structurerModel}`);
  console.log(`  LM Studio:  ${args.lmStudioUrl}`);
  console.log("");

  const client = createLmStudioClient({
    baseUrl: args.lmStudioUrl,
    apiKey: "lm-studio",
  });

  console.log("Step 1/2 — Extracting text with DeepSeek-OCR...");
  const ocrResult = await extractTextWithOcr(
    client,
    args.imagePath,
    args.ocrModel,
    args.verbose,
  );

  if (args.verbose) {
    console.log("\n--- OCR OUTPUT ---");
    console.log(ocrResult.rawText);
    console.log("--- END OCR OUTPUT ---\n");
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
  console.log(`  Rows: ${csvResult.csvContent.split("\n").filter(Boolean).length - 1} (excluding header)`);
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
