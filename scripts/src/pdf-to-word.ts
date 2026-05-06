/**
 * pdf-to-word.ts
 *
 * PDF → Word (.docx) conversion pipeline.
 *
 * Per-page, three independent text extraction passes are run:
 *   1. pdfjs text layer   — structural (tabs = column boundaries)
 *   2. DeepSeek-OCR       — visual extraction from a rendered page image
 *   3. Gemma4 direct      — Gemma4's own vision extraction pass on the page image
 *
 * These are then reconciled by a fourth Gemma4 corroboration call that uses
 * all three sources plus the page image to produce the most accurate text.
 *
 * Arabic (RTL) and Latin text are preserved throughout. The word-generator
 * renders Arabic paragraphs as right-to-left in the output document.
 *
 * Progress tracking:
 *   A JSON progress file is written after every completed page. If the process
 *   is interrupted, re-running the same command automatically resumes from the
 *   last completed page — no work is repeated.
 *
 *   Progress file location: <outputPath with .docx replaced by .progress.json>
 *   Delete the progress file to force a full re-run from scratch.
 */

import { readFile, writeFile, mkdir, rm, access } from "node:fs/promises";
import { join, dirname } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import type OpenAI from "openai";
import type {
  ChatCompletionMessageParam,
  ChatCompletionContentPart,
} from "openai/resources/chat/completions";
import {
  detectRenderer,
  renderPageToJpeg,
  extractPdfjsPageTextsRaw,
  type Renderer,
} from "./pdf.js";
import { callOcrModel, OCR_SYSTEM_PROMPT_WORD } from "./ocr.js";
import { WordProgressSchema, type WordProgress, type PageExtraction } from "./types.js";
import { generateWordDoc } from "./word-generator.js";

// ---------------------------------------------------------------------------
// Prompts
// ---------------------------------------------------------------------------

const GEMMA_EXTRACT_SYSTEM_PROMPT = [
  "You are a multimodal document reader with high accuracy for both Arabic and Latin text.",
  "Look at the document page image provided and extract all visible text.",
  "",
  "STRUCTURE RULES:",
  "- Headings and titles: place on their own line, with a blank line before and after.",
  "- Body paragraphs: separate with blank lines.",
  "- Tables: use tab characters (\\t) between columns, newlines between rows.",
  "  Include the header row. Do not skip any column.",
  "- Numbered or bulleted lists: preserve list markers (1., 2., •, -, etc.).",
  "- Key-value pairs: Key: Value, one per line.",
  "",
  "LANGUAGE:",
  "- Arabic text: output each Arabic word or phrase exactly as it appears.",
  "  Maintain the correct reading direction for each word (right-to-left within words).",
  "- Latin text, numbers, symbols: output exactly as shown.",
  "- Do not translate or transliterate any text.",
  "",
  "OUTPUT:",
  "- Only the extracted text. No commentary, no explanations, no markdown fences.",
  "- Preserve blank lines that reflect the document's paragraph and section structure.",
].join("\n");

const CORROBORATE_SYSTEM_PROMPT = [
  "You are a document accuracy expert specializing in Arabic and Latin text.",
  "You will receive:",
  "  1. A rendered image of a PDF page (use this as the visual ground truth).",
  "  2. Five independently extracted text versions of that page:",
  "       Source 1 — PDF text layer (pdfjs, structural/positional)",
  "       Source 2 — DeepSeek-OCR (visual OCR model)",
  "       Source 3 — dots.ocr (visual OCR model)",
  "       Source 4 — GLM-OCR (visual OCR model)",
  "       Source 5 — Gemma4 vision (AI visual reading)",
  "",
  "Your task: produce a single, maximally accurate version of the page's text content.",
  "",
  "RECONCILIATION RULES:",
  "- Use the page image to visually verify text when sources disagree.",
  "- Content appearing in 3 or more sources is almost certainly correct — include it.",
  "- Content in 2 sources: include if the page image does not contradict it.",
  "- Content in only 1 source: include only if the page image clearly confirms it.",
  "- If all five sources are wrong on something, correct it using the image.",
  "",
  "LANGUAGE:",
  "- Arabic text: preserve every Arabic word exactly. Maintain RTL character order within words.",
  "- Latin text, numbers, punctuation: preserve exactly.",
  "- Do not translate or transliterate.",
  "",
  "STRUCTURE:",
  "- Headings/titles: on their own line, blank lines before and after.",
  "- Paragraphs: separated by blank lines.",
  "- Tables: tab (\\t) between columns, newlines between rows.",
  "- Lists: preserve numbering and bullet markers.",
  "",
  "OUTPUT:",
  "- Only the final corroborated text. No commentary, no markdown fences.",
].join("\n");

// ---------------------------------------------------------------------------
// Progress file helpers
// ---------------------------------------------------------------------------

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function loadProgress(
  progressPath: string,
  pdfPath: string,
  totalPages: number,
): Promise<WordProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = WordProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath === pdfPath &&
        parsed.data.totalPages === totalPages
      ) {
        const completed = Object.keys(parsed.data.pages).length;
        if (completed > 0) {
          console.log(
            `[Word] Resuming: ${completed}/${totalPages} page(s) already done — loading from progress file.`,
          );
        }
        return parsed.data;
      }
      console.log(
        "[Word] Progress file found but does not match this PDF — starting fresh.",
      );
    } catch {
      console.log("[Word] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, totalPages, pages: {} };
}

async function saveProgress(
  progressPath: string,
  progress: WordProgress,
): Promise<void> {
  await mkdir(dirname(progressPath), { recursive: true });
  await writeFile(progressPath, JSON.stringify(progress, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Strip model thinking / reasoning traces
//
// Some models (e.g. Gemma4 with reasoning enabled) wrap their chain-of-thought
// in one of these patterns before the actual answer:
//
//   <|channel>thought  ...reasoning...  <channel|>
//   <thinking>         ...reasoning...  </thinking>
//   <think>            ...reasoning...  </think>
//
// We strip all such blocks and then trim, so only the final answer text
// reaches the Word document.
// ---------------------------------------------------------------------------

function stripThinking(raw: string): string {
  let text = raw;

  // <|channel>thought ... <channel|>  (Gemma4 / LM Studio reasoning format)
  text = text.replace(/<\|channel>thought[\s\S]*?<channel\|>/g, "");

  // <thinking> ... </thinking>
  text = text.replace(/<thinking>[\s\S]*?<\/thinking>/gi, "");

  // <think> ... </think>
  text = text.replace(/<think>[\s\S]*?<\/think>/gi, "");

  return text.trim();
}

// ---------------------------------------------------------------------------
// Gemma4 direct vision extraction (pass 3)
// ---------------------------------------------------------------------------

async function extractWithGemma4(
  client: OpenAI,
  modelId: string,
  pageImage: { base64: string; mimeType: string },
  pageNum: number,
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(
      `[Word/Gemma4-Extract] Page ${pageNum}: sending ${Math.round((pageImage.base64.length * 0.75) / 1024)} KB...`,
    );
  }

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: GEMMA_EXTRACT_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: { url: `data:${pageImage.mimeType};base64,${pageImage.base64}` },
          },
          {
            type: "text" as const,
            text: `Extract all text content from this document page (page ${pageNum}). Include Arabic and Latin text exactly as shown.`,
          },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0.1,
  });

  const text = stripThinking(response.choices[0]?.message.content ?? "");

  if (verbose) {
    console.log(
      `[Word/Gemma4-Extract] Page ${pageNum}: extracted ${text.length} chars.`,
    );
  }

  return text;
}

// ---------------------------------------------------------------------------
// Gemma4 corroboration (pass 4)
// ---------------------------------------------------------------------------

async function corroboratePage(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  pdfjsText: string,
  ocrText: string,
  dotsOcrText: string,
  glmOcrText: string,
  gemmaText: string,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(`[Word/Corroborate] Page ${pageNum}: reconciling 5 sources...`);
  }

  const textBody = [
    `=== SOURCE 1: PDF TEXT LAYER (structural embedding) ===`,
    pdfjsText || "(no text extracted from text layer)",
    ``,
    `=== SOURCE 2: DEEPSEEK-OCR EXTRACTION (visual OCR) ===`,
    ocrText || "(no text extracted by DeepSeek-OCR)",
    ``,
    `=== SOURCE 3: DOTS.OCR EXTRACTION (visual OCR) ===`,
    dotsOcrText || "(no text extracted by dots.ocr)",
    ``,
    `=== SOURCE 4: GLM-OCR EXTRACTION (visual OCR) ===`,
    glmOcrText || "(no text extracted by GLM-OCR)",
    ``,
    `=== SOURCE 5: GEMMA4 VISION EXTRACTION (AI visual reading) ===`,
    gemmaText || "(no text extracted by Gemma4)",
    ``,
    `Using the page image above as visual ground truth, produce the single most accurate version of this page's text.`,
  ].join("\n");

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: CORROBORATE_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: { url: `data:${pageImage.mimeType};base64,${pageImage.base64}` },
          },
          { type: "text" as const, text: textBody },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0.05,
  });

  const result = stripThinking(response.choices[0]?.message.content ?? "");

  if (verbose) {
    console.log(
      `[Word/Corroborate] Page ${pageNum}: corroborated text = ${result.length} chars.`,
    );
  }

  // If the model returned something meaningful, use it. Otherwise fall back
  // to the best single source: prefer Gemma4 > OCR > pdfjs.
  if (result.length > 20) return result;
  return gemmaText || ocrText || pdfjsText;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface ConvertPdfToWordArgs {
  pdfPath: string;
  outputPath: string;
  progressPath: string;
  client: OpenAI;
  ocrModelId: string;
  dotsOcrModelId: string;
  glmOcrModelId: string;
  structurerModelId: string;
  verbose: boolean;
}

/**
 * Convert a PDF to a Word document using three-source extraction and
 * Gemma4-powered corroboration, with per-page progress tracking.
 *
 * If the process is interrupted, re-running the same command will resume
 * from the last completed page using the saved progress file.
 */
export async function convertPdfToWord(args: ConvertPdfToWordArgs): Promise<void> {
  const {
    pdfPath,
    outputPath,
    progressPath,
    client,
    ocrModelId,
    dotsOcrModelId,
    glmOcrModelId,
    structurerModelId,
    verbose,
  } = args;

  // ── Step 1: pdfjs text layer (all pages at once) ──────────────────────────
  console.log("[Word] Step 1 — Extracting PDF text layer (pdfjs)...");
  const { pageTexts: pdfjsPageTexts, numPages, hasAnyText } =
    await extractPdfjsPageTextsRaw(pdfPath, verbose);

  if (!hasAnyText) {
    console.log(
      "[Word] Warning: No embedded text layer found in this PDF (possibly scanned).",
    );
    console.log(
      "         OCR and Gemma4 vision passes will still run on each rendered page.",
    );
  }

  console.log(`[Word]   ${numPages} page(s) found.`);
  console.log("");

  // ── Load progress ─────────────────────────────────────────────────────────
  const progress = await loadProgress(progressPath, pdfPath, numPages);

  // ── Detect renderer ───────────────────────────────────────────────────────
  const renderer: Renderer | null = await detectRenderer();
  if (!renderer) {
    console.log(
      "[Word] Warning: No PDF renderer found (pdftoppm / mutool).",
    );
    console.log(
      "         OCR and Gemma4 vision passes will be skipped.",
    );
    console.log(
      "         Install via: brew install poppler  (provides pdftoppm)",
    );
    console.log("");
  }

  // ── Per-page extraction loop ──────────────────────────────────────────────
  const tmpDir = join(tmpdir(), `doc2word-${randomUUID()}`);
  await mkdir(tmpDir, { recursive: true });

  try {
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const pageKey = String(pageNum);

      if (progress.pages[pageKey]) {
        if (verbose) {
          console.log(`[Word] Page ${pageNum}/${numPages} — already done, skipping.`);
        }
        continue;
      }

      console.log(`[Word] Page ${pageNum}/${numPages}:`);

      const pdfjsText = pdfjsPageTexts[pageNum - 1] ?? "";

      // ── Render page ───────────────────────────────────────────────────────
      let pageImage: { base64: string; mimeType: "image/jpeg" } | null = null;
      if (renderer) {
        try {
          process.stdout.write(`       Rendering...`);
          pageImage = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
          process.stdout.write(` done (${Math.round((pageImage.base64.length * 0.75) / 1024)} KB)\n`);
        } catch (renderErr) {
          process.stdout.write(` failed\n`);
          console.log(
            `       Render error: ${renderErr instanceof Error ? renderErr.message : String(renderErr)}`,
          );
        }
      }

      // ── Passes 2–5: All four model extractions run in parallel ────────────
      // DeepSeek-OCR, dots.ocr, GLM-OCR, and Gemma4-direct are independent
      // of each other, so we fire them all at once and wait for all results.
      let ocrText = "";
      let dotsOcrText = "";
      let glmOcrText = "";
      let gemmaText = "";

      if (pageImage) {
        // Print a header line, then each model prints its own result line the
        // instant it finishes — output appears in completion order, not start order.
        console.log(`       Extracting (4 models in parallel):`);
        const t0 = Date.now();

        function liveModel(label: string, p: Promise<string>): Promise<string> {
          const pad = label.padEnd(16);
          return p.then(
            (text) => {
              const s = ((Date.now() - t0) / 1000).toFixed(1);
              process.stdout.write(`         ✓ ${pad} ${text.length} chars  (${s}s)\n`);
              return text;
            },
            (err: unknown) => {
              const s = ((Date.now() - t0) / 1000).toFixed(1);
              const msg = err instanceof Error ? err.message.slice(0, 100) : String(err);
              process.stdout.write(`         ✗ ${pad} error: ${msg}  (${s}s)\n`);
              return "";
            },
          );
        }

        [ocrText, dotsOcrText, glmOcrText, gemmaText] = await Promise.all([
          liveModel("DeepSeek-OCR", callOcrModel(client, pageImage.base64, pageImage.mimeType, ocrModelId,     verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("dots.ocr",     callOcrModel(client, pageImage.base64, pageImage.mimeType, dotsOcrModelId, verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("GLM-OCR",      callOcrModel(client, pageImage.base64, pageImage.mimeType, glmOcrModelId,  verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("Gemma4",       extractWithGemma4(client, structurerModelId, pageImage, pageNum, verbose)),
        ]);
      }

      // ── Pass 6: Corroboration ─────────────────────────────────────────────
      let corroborated: string;
      if (pageImage) {
        try {
          process.stdout.write(`       Corroborating (5 sources)...`);
          corroborated = await corroboratePage(
            client,
            structurerModelId,
            pageNum,
            pdfjsText,
            ocrText,
            dotsOcrText,
            glmOcrText,
            gemmaText,
            pageImage,
            verbose,
          );
          process.stdout.write(` done (${corroborated.length} chars)\n`);
        } catch (corrobErr) {
          process.stdout.write(` failed — using best available source\n`);
          console.log(
            `       Corroboration error: ${corrobErr instanceof Error ? corrobErr.message : String(corrobErr)}`,
          );
          corroborated = gemmaText || glmOcrText || dotsOcrText || ocrText || pdfjsText;
        }
      } else {
        // No image available — fall back to the pdfjs text layer only
        corroborated = pdfjsText;
        console.log(`       No renderer — using pdfjs text layer only.`);
      }

      // ── Save progress ─────────────────────────────────────────────────────
      const extraction: PageExtraction = {
        pdfjsText,
        ocrText,
        dotsOcrText,
        glmOcrText,
        gemmaText,
        corroborated,
      };
      progress.pages[pageKey] = extraction;
      await saveProgress(progressPath, progress);
    }
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }

  console.log("");

  // ── Generate Word document ────────────────────────────────────────────────
  console.log("[Word] Generating Word document...");
  const orderedPages = Array.from({ length: numPages }, (_, i) => ({
    pageNum: i + 1,
    text: progress.pages[String(i + 1)]?.corroborated ?? "",
  }));

  await generateWordDoc(orderedPages, outputPath, verbose);

  console.log("");
  console.log("Done!");
  console.log(`  Word document: ${outputPath}`);
  console.log(`  Progress file: ${progressPath}`);
  console.log(`  (Delete the progress file to re-process from scratch next time.)`);
}
