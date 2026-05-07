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
import { callOcrModel, OCR_SYSTEM_PROMPT_WORD, stripThinking } from "./ocr.js";
import { WordProgressSchema, type WordProgress, type PageExtraction } from "./types.js";
import { generateWordDoc } from "./word-generator.js";

// ---------------------------------------------------------------------------
// Prompts
// ---------------------------------------------------------------------------

const CORROBORATE_SYSTEM_PROMPT = [
  "You are a document accuracy expert specializing in Arabic and Latin text.",
  "You will receive:",
  "  1. A rendered image of a PDF page (use this as the visual ground truth).",
  "  2. Four independently extracted text versions of that page:",
  "       Source 1 — DeepSeek-OCR (visual OCR model)",
  "       Source 2 — dots.ocr (visual OCR model)",
  "       Source 3 — GLM-OCR (visual OCR model)",
  "       Source 4 — Chandra-OCR (visual OCR model)",
  "",
  "Your task: produce a single, maximally accurate version of the page's text content.",
  "",
  "RECONCILIATION RULES:",
  "- Use the page image to visually verify text when sources disagree.",
  "- Content appearing in 3 or more sources is almost certainly correct — include it.",
  "- Content in 2 sources: include if the page image does not contradict it.",
  "- Content in only 1 source: include only if the page image clearly confirms it.",
  "- If all four sources are wrong on something, correct it using the image.",
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
  "- Only the final corroborated text. No commentary, no markdown fences, no reasoning traces.",
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
// Gemma4 corroboration
// ---------------------------------------------------------------------------

async function corroboratePage(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  ocrText: string,
  dotsOcrText: string,
  glmOcrText: string,
  chandraOcrText: string,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(`[Word/Corroborate] Page ${pageNum}: reconciling 4 OCR sources...`);
  }

  const textBody = [
    `=== SOURCE 1: DEEPSEEK-OCR EXTRACTION (visual OCR) ===`,
    ocrText || "(no text extracted by DeepSeek-OCR)",
    ``,
    `=== SOURCE 2: DOTS.OCR EXTRACTION (visual OCR) ===`,
    dotsOcrText || "(no text extracted by dots.ocr)",
    ``,
    `=== SOURCE 3: GLM-OCR EXTRACTION (visual OCR) ===`,
    glmOcrText || "(no text extracted by GLM-OCR)",
    ``,
    `=== SOURCE 4: CHANDRA-OCR EXTRACTION (visual OCR) ===`,
    chandraOcrText || "(no text extracted by Chandra-OCR)",
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
  // to the best single OCR source: prefer Chandra > GLM > dots > DeepSeek.
  if (result.length > 20) return result;
  return chandraOcrText || glmOcrText || dotsOcrText || ocrText;
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
  chandraOcrModelId: string;
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
    chandraOcrModelId,
    structurerModelId,
    verbose,
  } = args;

  // ── Step 1: count pages via pdfjs (text layer discarded) ─────────────────
  console.log("[Word] Step 1 — Counting pages...");
  const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
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

  const docStart = Date.now();

  try {
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const pageKey = String(pageNum);

      if (progress.pages[pageKey]) {
        if (verbose) {
          console.log(`[Word] Page ${pageNum}/${numPages} — already done, skipping.`);
        }
        continue;
      }

      const pageStart = Date.now();
      console.log(`[Word] Page ${pageNum}/${numPages}:`);

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

      // ── Passes 2–5: Four OCR models run in parallel ──────────────────────
      // All four are independent — fire at once, print as each finishes.
      let ocrText = "";
      let dotsOcrText = "";
      let glmOcrText = "";
      let chandraOcrText = "";

      if (pageImage) {
        console.log(`       Extracting (4 OCR models in parallel):`);
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

        [ocrText, dotsOcrText, glmOcrText, chandraOcrText] = await Promise.all([
          liveModel("DeepSeek-OCR",  callOcrModel(client, pageImage.base64, pageImage.mimeType, ocrModelId,        verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("dots.ocr",      callOcrModel(client, pageImage.base64, pageImage.mimeType, dotsOcrModelId,    verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("GLM-OCR",       callOcrModel(client, pageImage.base64, pageImage.mimeType, glmOcrModelId,     verbose, OCR_SYSTEM_PROMPT_WORD)),
          liveModel("Chandra-OCR",   callOcrModel(client, pageImage.base64, pageImage.mimeType, chandraOcrModelId, verbose, OCR_SYSTEM_PROMPT_WORD)),
        ]);
      }

      // ── Pass 5: Corroboration ─────────────────────────────────────────────
      let corroborated: string;
      if (pageImage) {
        try {
          process.stdout.write(`       Corroborating (4 OCR sources)...`);
          const corrobStart = Date.now();
          corroborated = await corroboratePage(
            client,
            structurerModelId,
            pageNum,
            ocrText,
            dotsOcrText,
            glmOcrText,
            chandraOcrText,
            pageImage,
            verbose,
          );
          const corrobSec = ((Date.now() - corrobStart) / 1000).toFixed(1);
          process.stdout.write(` done (${corroborated.length} chars, ${corrobSec}s)\n`);
        } catch (corrobErr) {
          process.stdout.write(` failed — using best available source\n`);
          console.log(
            `       Corroboration error: ${corrobErr instanceof Error ? corrobErr.message : String(corrobErr)}`,
          );
          corroborated = chandraOcrText || glmOcrText || dotsOcrText || ocrText;
        }
      } else {
        // No renderer — nothing can be extracted
        corroborated = "";
        console.log(`       No renderer — page will be blank.`);
      }

      // ── Save progress ─────────────────────────────────────────────────────
      const extraction: PageExtraction = {
        pdfjsText: "",   // pdfjs text no longer used as LLM source; kept for schema compat
        ocrText,
        dotsOcrText,
        glmOcrText,
        chandraOcrText,
        gemmaText: "",   // kept for progress-file backward compatibility
        corroborated,
      };
      progress.pages[pageKey] = extraction;
      await saveProgress(progressPath, progress);

      const pageSec = ((Date.now() - pageStart) / 1000).toFixed(1);
      console.log(`       Page ${pageNum} done in ${pageSec}s`);
    }
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }

  console.log("");

  // ── Generate Word document ────────────────────────────────────────────────
  console.log("[Word] Generating Word document...");
  const docGenStart = Date.now();
  const orderedPages = Array.from({ length: numPages }, (_, i) => ({
    pageNum: i + 1,
    text: progress.pages[String(i + 1)]?.corroborated ?? "",
  }));

  await generateWordDoc(orderedPages, outputPath, verbose);

  const totalSec = ((Date.now() - docStart) / 1000).toFixed(1);
  const docGenSec = ((Date.now() - docGenStart) / 1000).toFixed(1);

  console.log("");
  console.log("Done!");
  console.log(`  Word document:  ${outputPath}`);
  console.log(`  Progress file:  ${progressPath}`);
  console.log(`  (Delete the progress file to re-process from scratch next time.)`);
  console.log("");
  console.log(`  Timing:`);
  console.log(`    Document build:  ${docGenSec}s`);
  console.log(`    Total elapsed:   ${totalSec}s`);
}

// ---------------------------------------------------------------------------
// Enhance mode
// ---------------------------------------------------------------------------

/** Minimum character length for an OCR result to be considered "substantial". */
const SUBSTANTIAL_THRESHOLD = 20;

function isSubstantial(text: string): boolean {
  return text.trim().length > SUBSTANTIAL_THRESHOLD;
}

/**
 * Enhance an existing conversion by re-running OCR on pages where fewer than
 * 3 of the 4 OCR models produced substantial text.
 *
 * For each such page all weak models are retried in parallel (up to
 * MAX_RETRY_ATTEMPTS sequential attempts per model). Results print to the
 * terminal in real time as each attempt finishes, regardless of order.
 *
 * Backward compatible: existing progress files without `chandraOcrText` parse
 * fine because the field defaults to "".
 */
export async function enhanceProgressFile(args: ConvertPdfToWordArgs): Promise<void> {
  const {
    pdfPath,
    outputPath,
    progressPath,
    client,
    ocrModelId,
    dotsOcrModelId,
    glmOcrModelId,
    chandraOcrModelId,
    structurerModelId,
    verbose,
  } = args;

  const MAX_RETRY_ATTEMPTS = 3;
  const MIN_SUBSTANTIAL_SOURCES = 3;

  // ── Load existing progress file ───────────────────────────────────────────
  console.log(`[Enhance] Progress file: ${progressPath}`);

  if (!(await fileExists(progressPath))) {
    console.error("[Enhance] No progress file found at the expected path.");
    console.error("          Run the conversion first (--word) before using --enhance.");
    process.exit(1);
  }

  let progress: WordProgress;
  try {
    const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
    const parsed = WordProgressSchema.safeParse(raw);
    if (!parsed.success) {
      console.error("[Enhance] Progress file failed validation:", parsed.error.message);
      process.exit(1);
    }
    progress = parsed.data;
  } catch {
    console.error("[Enhance] Failed to read or parse progress file.");
    process.exit(1);
  }

  const { totalPages } = progress;
  const completedCount = Object.keys(progress.pages).length;

  if (completedCount === 0) {
    console.log("[Enhance] No completed pages in progress file — nothing to enhance.");
    return;
  }

  console.log(`[Enhance] ${completedCount}/${totalPages} pages completed in progress file.`);
  console.log("[Enhance] Scanning OCR coverage...");
  console.log("");

  // ── Require a renderer ────────────────────────────────────────────────────
  const renderer: Renderer | null = await detectRenderer();
  if (!renderer) {
    console.error("[Enhance] No PDF renderer found (pdftoppm / mutool).");
    console.error("          Install via: brew install poppler");
    process.exit(1);
  }

  // OCR model definitions — order determines fallback priority
  const ocrModels: Array<{ key: keyof PageExtraction; label: string; modelId: string }> = [
    { key: "ocrText",        label: "DeepSeek-OCR", modelId: ocrModelId },
    { key: "dotsOcrText",    label: "dots.ocr",     modelId: dotsOcrModelId },
    { key: "glmOcrText",     label: "GLM-OCR",      modelId: glmOcrModelId },
    { key: "chandraOcrText", label: "Chandra-OCR",  modelId: chandraOcrModelId },
  ];

  const tmpDir = join(tmpdir(), `doc2word-enhance-${randomUUID()}`);
  await mkdir(tmpDir, { recursive: true });

  const enhanceStart = Date.now();
  let pagesImproved = 0;
  let pagesAlreadyOk = 0;

  // ── Helper: retry one OCR model up to maxAttempts times, printing live ────
  // Sequential attempts per model, but models run concurrently with each other
  // via Promise.all below. Each attempt prints immediately on completion.
  function retryModel(
    label: string,
    modelId: string,
    image: { base64: string; mimeType: string },
    t0: number,
  ): Promise<string> {
    const pad = label.padEnd(14);
    async function attempt(n: number): Promise<string> {
      try {
        const text = await callOcrModel(
          client,
          image.base64,
          image.mimeType,
          modelId,
          verbose,
          OCR_SYSTEM_PROMPT_WORD,
        );
        const s = ((Date.now() - t0) / 1000).toFixed(1);
        if (isSubstantial(text)) {
          process.stdout.write(
            `           ✓ ${pad} attempt ${n}/${MAX_RETRY_ATTEMPTS}  ${text.length} chars  (${s}s)\n`,
          );
          return text;
        }
        process.stdout.write(
          `           ✗ ${pad} attempt ${n}/${MAX_RETRY_ATTEMPTS}  ${text.length} chars (too short)  (${s}s)\n`,
        );
      } catch (err) {
        const s = ((Date.now() - t0) / 1000).toFixed(1);
        const msg = err instanceof Error ? err.message.slice(0, 60) : String(err);
        process.stdout.write(
          `           ✗ ${pad} attempt ${n}/${MAX_RETRY_ATTEMPTS}  error: ${msg}  (${s}s)\n`,
        );
      }
      if (n < MAX_RETRY_ATTEMPTS) return attempt(n + 1);
      return "";
    }
    return attempt(1);
  }

  try {
    for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
      const pct = Math.round((pageNum / totalPages) * 100);
      const pageKey = String(pageNum);
      const pageData = progress.pages[pageKey];

      if (!pageData) {
        if (verbose) {
          console.log(
            `[Enhance] Page ${pageNum}/${totalPages} (${pct}%): not yet extracted — skipping.`,
          );
        }
        continue;
      }

      // ── Count how many OCR sources have substantial text ──────────────────
      const weakModels = ocrModels.filter(
        (m) => !isSubstantial(pageData[m.key] as string),
      );
      const substantialCount = ocrModels.length - weakModels.length;

      if (substantialCount >= MIN_SUBSTANTIAL_SOURCES) {
        pagesAlreadyOk++;
        if (verbose) {
          console.log(
            `[Enhance] Page ${pageNum}/${totalPages} (${pct}%): ${substantialCount}/4 sources OK — skipping.`,
          );
        }
        continue;
      }

      console.log(
        `[Enhance] Page ${pageNum}/${totalPages} (${pct}%): ${substantialCount}/4 OCR sources — enhancing...`,
      );

      // ── Render ────────────────────────────────────────────────────────────
      let pageImage: { base64: string; mimeType: "image/jpeg" } | null = null;
      try {
        process.stdout.write(`         Rendering...`);
        pageImage = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
        process.stdout.write(
          ` done (${Math.round((pageImage.base64.length * 0.75) / 1024)} KB)\n`,
        );
      } catch (renderErr) {
        process.stdout.write(` failed — skipping page\n`);
        console.log(
          `         Render error: ${renderErr instanceof Error ? renderErr.message : String(renderErr)}`,
        );
        continue;
      }

      // ── All weak models retry in parallel, printing results as they arrive ─
      const updatedData: PageExtraction = { ...pageData };

      if (weakModels.length > 0) {
        console.log(
          `         Retrying ${weakModels.length} weak model(s) in parallel:`,
        );
        const t0 = Date.now();

        const retryResults = await Promise.all(
          weakModels.map((m) => retryModel(m.label, m.modelId, pageImage!, t0)),
        );

        for (let i = 0; i < weakModels.length; i++) {
          const text = retryResults[i]!;
          if (isSubstantial(text)) {
            updatedData[weakModels[i]!.key] = text;
          }
        }
      }

      // ── Re-corroborate with all available sources ─────────────────────────
      const newSubstantialCount = ocrModels.filter(
        (m) => isSubstantial(updatedData[m.key] as string),
      ).length;
      process.stdout.write(
        `         → ${newSubstantialCount}/4 sources populated. Re-corroborating...`,
      );

      let anyImproved = false;
      try {
        const corrStart = Date.now();
        const corroborated = await corroboratePage(
          client,
          structurerModelId,
          pageNum,
          updatedData.ocrText,
          updatedData.dotsOcrText,
          updatedData.glmOcrText,
          updatedData.chandraOcrText,
          pageImage,
          verbose,
        );
        const corrSec = ((Date.now() - corrStart) / 1000).toFixed(1);
        process.stdout.write(
          ` done (${corroborated.length} chars, ${corrSec}s)\n`,
        );
        updatedData.corroborated = corroborated;
        anyImproved = true;
      } catch (corrobErr) {
        const msg =
          corrobErr instanceof Error ? corrobErr.message.slice(0, 80) : String(corrobErr);
        process.stdout.write(` failed: ${msg} — keeping existing text.\n`);
      }

      // ── Persist (resumable) ───────────────────────────────────────────────
      if (anyImproved) {
        updatedData.pdfjsText = "";  // ensure pdfjs text is never re-introduced
        progress.pages[pageKey] = updatedData;
        await saveProgress(progressPath, progress);
        pagesImproved++;
      }
    }
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }

  console.log("");
  console.log(
    `[Enhance] Complete. ${pagesImproved} page(s) improved, ${pagesAlreadyOk} page(s) already had sufficient coverage.`,
  );
  console.log("");

  // ── Regenerate Word document from updated progress ────────────────────────
  console.log("[Enhance] Regenerating Word document...");
  const docGenStart = Date.now();
  const orderedPages = Array.from({ length: totalPages }, (_, i) => ({
    pageNum: i + 1,
    text: progress.pages[String(i + 1)]?.corroborated ?? "",
  }));
  await generateWordDoc(orderedPages, outputPath, verbose);
  const docGenSec = ((Date.now() - docGenStart) / 1000).toFixed(1);
  const totalSec = ((Date.now() - enhanceStart) / 1000).toFixed(1);

  console.log("");
  console.log("Done!");
  console.log(`  Word document:  ${outputPath}`);
  console.log(`  Progress file:  ${progressPath}`);
  console.log("");
  console.log(`  Timing:`);
  console.log(`    Document build:  ${docGenSec}s`);
  console.log(`    Total elapsed:   ${totalSec}s`);
}
