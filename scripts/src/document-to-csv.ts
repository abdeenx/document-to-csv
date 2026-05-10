#!/usr/bin/env node
/**
 * document-to-csv
 *
 * Converts a document image or PDF into a structured CSV, Excel, or Word file.
 *
 * ── Mode 1: Image → CSV / Excel ──────────────────────────────────────────────
 *   1. DeepSeek-OCR (vision) extracts text
 *   2. Gemma4 (tool use) structures it into CSV / Excel
 *      The original image is forwarded to Gemma4 for visual column-alignment
 *      verification.
 *
 * ── Mode 2: PDF → CSV / Excel ────────────────────────────────────────────────
 *   1. pdfjs-dist extracts the embedded text layer (structural / positional)
 *      DeepSeek-OCR runs on each rendered page for a second visual pass
 *   2. Gemma4 (non-tool) reconciles the page images + both text layers into
 *      a single clean, tab-delimited text (visual verification step)
 *   3. Gemma4 (tool use) structures the verified text into CSV / Excel
 *      Page images are forwarded again so Gemma4 can verify column alignment
 *
 *   With --excel on a PDF:
 *     A layout-faithful "Document" sheet is added to the workbook. Text items
 *     are placed at their PDF-grid positions, and any images detected in the PDF
 *     operator list are cropped from the rendered page renders and embedded at
 *     the matching cell anchor.
 *
 * ── Mode 3: PDF → Word (.docx) ──────────────────────────────────────────────
 *   Per page, four OCR models run in parallel then Gemma4 corroborates:
 *     1. DeepSeek-OCR, dots.ocr, GLM-OCR, Chandra-OCR (parallel)
 *     2. Gemma4 corroborate — reconciles 4 sources into the most accurate text
 *
 *   Arabic (RTL) and Latin text are both preserved.
 *   Progress is saved after every page — resume safely if interrupted.
 *
 * ── Mode 4: PDF → Word via Qwen2.5-VL (`--qwen-word`) ───────────────────────
 *   Single-model pipeline — one Qwen2.5-VL call per page, no corroboration.
 *   Faster and simpler than Mode 3; strong Arabic + Latin accuracy.
 *   Progress is saved after every page — resume safely if interrupted.
 *
 * ── Mode 5: PDF → Word via Qwen3-VL (`--qwen3-word`) ────────────────────────
 *   Identical pipeline to Mode 4 but uses Qwen3-VL-8B as the vision model.
 *   Progress file uses .qwen3-progress.json to distinguish from Mode 4.
 *
 * ── Mode 6: PDF + Word → corrected Word (`--qwen3-review`) ──────────────────
 *   Takes a PDF (visual ground truth) and an existing Word document.
 *   For each page, Qwen3-VL compares the rendered PDF page against the
 *   extracted Word text and fixes every discrepancy. Produces a corrected docx.
 *   Requires: --review-word <word-file.docx>
 *   Progress is saved after every page — resume safely if interrupted.
 *
 * Usage:
 *   pnpm --filter @workspace/scripts run document-to-csv <file> [options]
 */

import { basename, extname, resolve, dirname } from "node:path";
import { writeFile } from "node:fs/promises";
import { parseArgs } from "node:util";
import { CliArgsSchema } from "./types.js";
import { createLmStudioClient } from "./lm-studio-client.js";
import { extractTextWithOcr } from "./ocr.js";
import { extractTextFromPdf, extractPdfPageLayouts } from "./pdf.js";
import { verifyDocumentWithGemma, generateCsvWithGemma } from "./csv-generator.js";
import { generateExcel } from "./excel-generator.js";
import { convertPdfToWord, enhanceProgressFile } from "./pdf-to-word.js";
import { convertPdfToWordWithQwen } from "./qwen-pdf-to-word.js";
import { reviewWordWithQwen3 } from "./qwen3-review.js";

const PDF_EXTENSION    = ".pdf";
const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp"]);

function parseCliArgs(): ReturnType<typeof CliArgsSchema.parse> {
  const { values, positionals } = parseArgs({
    args: process.argv.slice(2),
    options: {
      output:             { type: "string",  short: "o" },
      excel:              { type: "boolean", default: false },
      word:               { type: "boolean", default: false },
      enhance:            { type: "boolean", default: false },
      "qwen-word":          { type: "boolean", default: false },
      "lm-studio-url":      { type: "string" },
      "ocr-model":          { type: "string" },
      "dots-ocr-model":     { type: "string" },
      "glm-ocr-model":      { type: "string" },
      "chandra-ocr-model":  { type: "string" },
      "structurer-model":   { type: "string" },
      "qwen-model":         { type: "string" },
      "qwen3-word":         { type: "boolean", default: false },
      "qwen3-review":       { type: "boolean", default: false },
      "review-word":        { type: "string" },
      "qwen3-model":        { type: "string" },
      verbose:            { type: "boolean", short: "v", default: false },
      help:               { type: "boolean", short: "h", default: false },
    },
    allowPositionals: true,
  });

  if (values.help || positionals.length === 0) {
    printUsage();
    process.exit(values.help ? 0 : 1);
  }

  const modeCount = [
    values.excel,
    values.word,
    values["qwen-word"],
    values["qwen3-word"],
    values["qwen3-review"],
  ].filter(Boolean).length;
  if (modeCount > 1) {
    console.error(
      "Error: --excel, --word, --qwen-word, --qwen3-word, and --qwen3-review are mutually exclusive.",
    );
    process.exit(1);
  }

  if (values["qwen3-review"] && !values["review-word"]) {
    console.error("Error: --qwen3-review requires --review-word <path-to-word-file.docx>");
    process.exit(1);
  }

  if (values.enhance && !values.word) {
    console.error("Error: --enhance requires --word.");
    process.exit(1);
  }

  const inputPath = positionals[0]!;
  const ext       = extname(inputPath).toLowerCase();

  if (ext !== PDF_EXTENSION && !IMAGE_EXTENSIONS.has(ext)) {
    console.error(`Unsupported file type "${ext}". Supported: pdf, jpg, jpeg, png, gif, webp`);
    process.exit(1);
  }

  if (values.word && ext !== PDF_EXTENSION) {
    console.error("Error: --word mode requires a PDF input file.");
    process.exit(1);
  }

  if (values["qwen-word"] && ext !== PDF_EXTENSION) {
    console.error("Error: --qwen-word mode requires a PDF input file.");
    process.exit(1);
  }

  if (values["qwen3-word"] && ext !== PDF_EXTENSION) {
    console.error("Error: --qwen3-word mode requires a PDF input file.");
    process.exit(1);
  }

  if (values["qwen3-review"] && ext !== PDF_EXTENSION) {
    console.error("Error: --qwen3-review mode requires a PDF input file.");
    process.exit(1);
  }

  const inputBase = basename(inputPath, extname(inputPath));
  const callerCwd = process.env["INIT_CWD"] ?? process.cwd();

  let outputExt: string;
  if (values.excel) outputExt = ".xlsx";
  else if (values.word || values["qwen-word"] || values["qwen3-word"] || values["qwen3-review"])
    outputExt = ".docx";
  else outputExt = ".csv";

  // --qwen3-review defaults to <basename>-reviewed.docx to avoid overwriting the original.
  const defaultOutput = values["qwen3-review"]
    ? resolve(callerCwd, `${inputBase}-reviewed.docx`)
    : resolve(callerCwd, `${inputBase}${outputExt}`);

  const raw = {
    imagePath:       resolve(callerCwd, inputPath),
    outputPath:      values.output ? resolve(callerCwd, values.output) : defaultOutput,
    lmStudioUrl:     values["lm-studio-url"]    ?? "http://localhost:1234/v1",
    ocrModel:        values["ocr-model"]          ?? "mlx-community/DeepSeek-OCR-8bit",
    dotsOcrModel:    values["dots-ocr-model"]     ?? "mlx-community/dots.ocr-bf16",
    glmOcrModel:     values["glm-ocr-model"]      ?? "mlx-community/GLM-OCR-bf16",
    chandraOcrModel: values["chandra-ocr-model"]  ?? "jwindle47/chandra-ocr-2-8bit-mlx",
    structurerModel: values["structurer-model"]  ??
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    qwenModel:      values["qwen-model"]    ?? "mlx-community/Qwen2.5-VL-7B-Instruct-8bit",
    qwen3Model:     values["qwen3-model"]   ?? "qwen/qwen3-vl-8b",
    verbose:        values.verbose          ?? false,
    excel:          values.excel            ?? false,
    word:           values.word             ?? false,
    enhance:        values.enhance          ?? false,
    qwenWord:       values["qwen-word"]     ?? false,
    qwen3Word:      values["qwen3-word"]    ?? false,
    qwen3Review:    values["qwen3-review"]  ?? false,
    reviewWordPath: values["review-word"],
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
document-to-csv — Convert a document image or PDF to CSV, Excel, or Word

Usage:
  pnpm --filter @workspace/scripts run document-to-csv <file> [options]

Arguments:
  <file>                       Path to the input file
                               Images: jpg, jpeg, png, gif, webp
                                 → OCR via DeepSeek-OCR, then structured by Gemma4
                               PDF:    pdf
                                 → Three modes available (see flags below)

Options:
  -o, --output <path>          Output file path
                               Default: <basename>.csv / .xlsx / .docx
  --excel                      Write an Excel (.xlsx) file (PDF or image input).
                               For PDF inputs the workbook contains:
                                 "Data"      — clean structured table
                                 "Document"  — layout-faithful sheet (text at PDF
                                               positions, embedded images)
  --word                       Write a Word (.docx) file (PDF input only).
                               Runs 4 OCR models in parallel then Gemma4
                               corroborates to produce the most accurate text.
                               Arabic (RTL) and Latin text are both preserved.
                               Progress is saved after every page — re-run the
                               same command to resume if interrupted.
  --enhance                    Re-run OCR on weak pages in an existing --word
                               progress file, then regenerate the docx.
  --qwen-word                  Write a Word (.docx) file (PDF input only) using
                               a single Qwen2.5-VL pass per page — no multi-
                               model corroboration required. Faster and lighter
                               than --word while retaining strong Arabic accuracy.
                               Progress is saved after every page — resumable.
  --qwen3-word                 Same as --qwen-word but uses Qwen3-VL-8B.
                               Progress file: <output>.qwen3-progress.json
  --qwen3-review               Review an existing Word document against a PDF.
                               For each page, Qwen3-VL compares the rendered
                               PDF page image against the Word text and fixes
                               every discrepancy. Produces a corrected .docx.
                               Requires: --review-word <word-file.docx>
                               Default output: <pdf-basename>-reviewed.docx
                               Progress is saved after every page — resumable.
  --review-word <path>         Path to the Word (.docx) file used as input for
                               --qwen3-review.
  --lm-studio-url <url>        LM Studio base URL (default: http://localhost:1234/v1)
  --ocr-model <id>             OCR model (default: mlx-community/DeepSeek-OCR-8bit)
  --structurer-model <id>      Structuring / corroboration model
                               (default: zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine)
  --qwen-model <id>            Qwen vision model used by --qwen-word
                               (default: mlx-community/Qwen2.5-VL-7B-Instruct-8bit)
  --qwen3-model <id>           Qwen3 vision model used by --qwen3-word and --qwen3-review
                               (default: qwen/qwen3-vl-8b)
  -v, --verbose                Enable step-by-step logging
  -h, --help                   Show this help message

PDF renderer (required for OCR passes and --excel image embedding):
  brew install poppler          # provides pdftoppm  ← recommended
  brew install mupdf-tools      # provides mutool    ← fallback

Examples:
  pnpm --filter @workspace/scripts run document-to-csv ./invoice.png
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --excel --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word --verbose
  pnpm --filter @workspace/scripts run document-to-csv ./contract.pdf --word --enhance
  pnpm --filter @workspace/scripts run document-to-csv ./report.pdf --output ./data/report.xlsx --excel
  pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen-word
  pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen-word --output ./out/arabic.docx
  pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen3-word
  pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen3-review --review-word ./arabic.docx
  pnpm --filter @workspace/scripts run document-to-csv ./arabic.pdf --qwen3-review --review-word ./draft.docx --output ./corrected.docx
`);
}

async function main(): Promise<void> {
  const args  = parseCliArgs();
  const isPdf = extname(args.imagePath).toLowerCase() === PDF_EXTENSION;

  // ── Mode 6: PDF + Word → corrected Word via Qwen3-VL ────────────────────
  if (args.qwen3Review) {
    const wordPath = args.reviewWordPath!;
    const progressPath = args.outputPath!.replace(/\.docx$/i, ".qwen3-review-progress.json");

    console.log("document-to-csv (PDF + Word → corrected Word / Qwen3-VL review mode)");
    console.log("========================================================================");
    console.log(`  PDF input:      ${args.imagePath}`);
    console.log(`  Word input:     ${wordPath}`);
    console.log(`  Output:         ${args.outputPath}`);
    console.log(`  Progress file:  ${progressPath}`);
    console.log(`  Qwen3 model:    ${args.qwen3Model}`);
    console.log(`  LM Studio:      ${args.lmStudioUrl}`);
    console.log("");

    const client = createLmStudioClient({ baseUrl: args.lmStudioUrl, apiKey: "lm-studio" });

    await reviewWordWithQwen3({
      pdfPath:       args.imagePath,
      wordPath,
      outputPath:    args.outputPath!,
      progressPath,
      client,
      qwen3ModelId:  args.qwen3Model,
      verbose:       args.verbose,
    });

    return;
  }

  // ── Mode 5: PDF → Word via Qwen3-VL (single-model) ───────────────────────
  if (args.qwen3Word) {
    const progressPath = args.outputPath!.replace(/\.docx$/i, ".qwen3-progress.json");

    console.log("document-to-csv (PDF → Word / Qwen3-VL mode)");
    console.log("==============================================");
    console.log(`  Input:          ${args.imagePath}`);
    console.log(`  Output:         ${args.outputPath}`);
    console.log(`  Progress file:  ${progressPath}`);
    console.log(`  Qwen3 model:    ${args.qwen3Model}`);
    console.log(`  LM Studio:      ${args.lmStudioUrl}`);
    console.log("");

    const client = createLmStudioClient({ baseUrl: args.lmStudioUrl, apiKey: "lm-studio" });

    await convertPdfToWordWithQwen({
      pdfPath:      args.imagePath,
      outputPath:   args.outputPath!,
      progressPath,
      client,
      qwenModelId:  args.qwen3Model,
      modelLabel:   "Qwen3-VL",
      verbose:      args.verbose,
    });

    return;
  }

  // ── Mode 4: PDF → Word via Qwen2.5-VL (single-model) ────────────────────
  if (args.qwenWord) {
    const progressPath = args.outputPath!.replace(/\.docx$/i, ".qwen-progress.json");

    console.log("document-to-csv (PDF → Word / Qwen2.5-VL mode)");
    console.log("=================================================");
    console.log(`  Input:          ${args.imagePath}`);
    console.log(`  Output:         ${args.outputPath}`);
    console.log(`  Progress file:  ${progressPath}`);
    console.log(`  Qwen model:     ${args.qwenModel}`);
    console.log(`  LM Studio:      ${args.lmStudioUrl}`);
    console.log("");

    const client = createLmStudioClient({ baseUrl: args.lmStudioUrl, apiKey: "lm-studio" });

    await convertPdfToWordWithQwen({
      pdfPath:      args.imagePath,
      outputPath:   args.outputPath!,
      progressPath,
      client,
      qwenModelId:  args.qwenModel,
      modelLabel:   "Qwen2.5-VL",
      verbose:      args.verbose,
    });

    return;
  }

  // ── Mode 3: PDF → Word (normal or enhance) ───────────────────────────────
  if (args.word) {
    const progressPath = args.outputPath!.replace(/\.docx$/i, ".progress.json");

    const modeLabel = args.enhance ? "PDF → Word (enhance mode)" : "PDF → Word mode";
    console.log(`document-to-csv (${modeLabel})`);
    console.log("==================================");
    console.log(`  Input:          ${args.imagePath}`);
    console.log(`  Output:         ${args.outputPath}`);
    console.log(`  Progress file:  ${progressPath}`);
    console.log(`  DeepSeek-OCR:   ${args.ocrModel}`);
    console.log(`  dots.ocr:       ${args.dotsOcrModel}`);
    console.log(`  GLM-OCR:        ${args.glmOcrModel}`);
    console.log(`  Chandra-OCR:    ${args.chandraOcrModel}`);
    console.log(`  Struct model:   ${args.structurerModel}`);
    console.log(`  LM Studio:      ${args.lmStudioUrl}`);
    console.log("");

    const client = createLmStudioClient({ baseUrl: args.lmStudioUrl, apiKey: "lm-studio" });

    const wordArgs = {
      pdfPath:            args.imagePath,
      outputPath:         args.outputPath!,
      progressPath,
      client,
      ocrModelId:         args.ocrModel,
      dotsOcrModelId:     args.dotsOcrModel,
      glmOcrModelId:      args.glmOcrModel,
      chandraOcrModelId:  args.chandraOcrModel,
      structurerModelId:  args.structurerModel,
      verbose:            args.verbose,
    };

    if (args.enhance) {
      await enhanceProgressFile(wordArgs);
    } else {
      await convertPdfToWord(wordArgs);
    }

    return;
  }

  // ── Mode 1 & 2: Image/PDF → CSV / Excel ──────────────────────────────────
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
