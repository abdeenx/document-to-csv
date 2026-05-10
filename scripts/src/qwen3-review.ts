/**
 * qwen3-review.ts
 *
 * PDF + Word → corrected Word pipeline using Qwen3-VL.
 *
 * Takes an existing Word document and a PDF (the visual ground truth).
 * For each page, the rendered PDF page image is compared against the
 * extracted Word text by Qwen3-VL, which returns a corrected version.
 * A new Word document is assembled from the corrected page texts.
 *
 * Pipeline per page:
 *   1. Extract Word text for this page (from word/document.xml)
 *   2. Render PDF page to JPEG
 *   3. Qwen3-VL reviews image vs Word text → corrected text
 *   4. Strip thinking traces, save to progress file (resumable)
 *
 * After all pages: rebuild Word document from corrected texts.
 *
 * Progress file: <output>.qwen3-review-progress.json
 */

import { readFile, writeFile, mkdir, rm, access } from "node:fs/promises";
import { join, dirname } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import type OpenAI from "openai";
import type { ChatCompletionContentPart } from "openai/resources/chat/completions";
import {
  detectRenderer,
  renderPageToJpeg,
  extractPdfjsPageTextsRaw,
  type Renderer,
} from "./pdf.js";
import { stripThinking } from "./ocr.js";
import { Qwen3ReviewProgressSchema, type Qwen3ReviewProgress } from "./types.js";
import { generateWordDoc } from "./word-generator.js";
import { extractDocxPages } from "./docx-reader.js";

// ---------------------------------------------------------------------------
// Review prompt
// ---------------------------------------------------------------------------

const REVIEW_SYSTEM_PROMPT = [
  "You are a document accuracy expert specializing in Arabic and Latin text.",
  "You will receive:",
  "  1. A rendered image of a PDF page (the visual ground truth).",
  "  2. The extracted text from the corresponding Word document page.",
  "",
  "Your task: compare the Word text against the PDF page image and produce a corrected version.",
  "",
  "CORRECTION RULES:",
  "- Fix any words, characters, or punctuation in the Word text that differ from the image.",
  "- Fix Arabic text that is garbled, missing, contains wrong characters, or is in the wrong order.",
  "- Fix Latin text with wrong characters, incorrect spacing, or wrong punctuation.",
  "- Add content clearly visible in the image but absent from the Word text.",
  "- Remove content in the Word text that is not visible in the image.",
  "- Preserve the overall document structure: headings, paragraphs, tables, lists.",
  "",
  "LANGUAGE:",
  "- Arabic text: reproduce every Arabic word exactly as shown in the image.",
  "- Latin text, numbers, punctuation, symbols: reproduce exactly as shown in the image.",
  "- Do not translate, transliterate, rephrase, or summarize anything.",
  "",
  "OUTPUT:",
  "- The corrected text only. No commentary, no explanations, no diff markers, no reasoning traces.",
  "- Preserve paragraph breaks and document structure from the image.",
].join("\n");

// ---------------------------------------------------------------------------
// Timing helpers
// ---------------------------------------------------------------------------

function fmtSec(ms: number): string {
  return `${(ms / 1000).toFixed(1)}s`;
}

function fmtEta(ms: number): string {
  const s = Math.round(ms / 1000);
  if (s < 60) return `~${s}s`;
  const m = Math.floor(s / 60);
  const rem = s % 60;
  return `~${m}m ${rem < 10 ? "0" : ""}${rem}s`;
}

function fmtClock(ts: number): string {
  const d = new Date(ts);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

const DIVIDER = "─".repeat(56);

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
  wordPath: string,
  totalPages: number,
): Promise<Qwen3ReviewProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = Qwen3ReviewProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath === pdfPath &&
        parsed.data.wordPath === wordPath &&
        parsed.data.totalPages === totalPages
      ) {
        const completed = Object.keys(parsed.data.pages).length;
        if (completed > 0) {
          console.log(
            `[Review] Resuming: ${completed}/${totalPages} page(s) already done.`,
          );
        }
        return parsed.data;
      }
      console.log("[Review] Progress file doesn't match inputs — starting fresh.");
    } catch {
      console.log("[Review] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, wordPath, totalPages, pages: {} };
}

async function saveProgress(
  progressPath: string,
  progress: Qwen3ReviewProgress,
): Promise<void> {
  await mkdir(dirname(progressPath), { recursive: true });
  await writeFile(progressPath, JSON.stringify(progress, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Single-page review call
// ---------------------------------------------------------------------------

async function reviewPage(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  wordText: string,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(`[Review] Page ${pageNum}: sending to ${modelId}...`);
  }

  const textBody = [
    `=== WORD DOCUMENT TEXT FOR THIS PAGE ===`,
    wordText || "(no text found for this page)",
    ``,
    `Compare the Word text above against the PDF page image.`,
    `Fix every discrepancy and return the corrected text.`,
  ].join("\n");

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: REVIEW_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: {
              url: `data:${pageImage.mimeType};base64,${pageImage.base64}`,
            },
          },
          { type: "text" as const, text: textBody },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const raw = response.choices[0]?.message.content ?? "";
  const text = stripThinking(raw);

  if (verbose) {
    const tokens = response.usage?.total_tokens ?? "unknown";
    console.log(`[Review] Page ${pageNum}: ${text.length} chars, ${tokens} tokens`);
  }

  return text;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface Qwen3ReviewArgs {
  pdfPath: string;
  wordPath: string;
  outputPath: string;
  progressPath: string;
  client: OpenAI;
  qwen3ModelId: string;
  verbose: boolean;
}

/**
 * Compare a Word document against the PDF and produce a corrected Word file.
 * Fully resumable — progress is saved after every page.
 */
export async function reviewWordWithQwen3(args: Qwen3ReviewArgs): Promise<void> {
  const {
    pdfPath,
    wordPath,
    outputPath,
    progressPath,
    client,
    qwen3ModelId,
    verbose,
  } = args;

  const docStart = Date.now();

  // ── Step 1: count PDF pages ───────────────────────────────────────────────
  {
    const t0 = Date.now();
    process.stdout.write(`[Review] ${fmtClock(t0)}  Step 1 — Counting PDF pages...`);
    const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
    process.stdout.write(` done  ${fmtSec(Date.now() - t0)}  (${numPages} page(s))\n`);

    // ── Step 2: extract Word pages ──────────────────────────────────────────
    const t1 = Date.now();
    process.stdout.write(`[Review] ${fmtClock(t1)}  Step 2 — Extracting Word pages...`);
    let wordPages: string[];
    try {
      wordPages = await extractDocxPages(wordPath);
      process.stdout.write(
        ` done  ${fmtSec(Date.now() - t1)}  (${wordPages.length} page section(s) found)\n`,
      );
    } catch (err) {
      process.stdout.write(` failed  ${fmtSec(Date.now() - t1)}\n`);
      throw err;
    }

    if (wordPages.length !== numPages) {
      console.log(
        `[Review] Warning: Word has ${wordPages.length} page section(s), PDF has ${numPages} page(s).`,
      );
      console.log(`         Matching by index; extra pages will use empty Word text.`);
    }
    console.log("");

    // ── Load progress ───────────────────────────────────────────────────────
    const progress = await loadProgress(progressPath, pdfPath, wordPath, numPages);

    // ── Detect renderer ─────────────────────────────────────────────────────
    const tr = Date.now();
    process.stdout.write(`[Review] ${fmtClock(tr)}  Detecting PDF renderer...`);
    const renderer: Renderer | null = await detectRenderer();
    if (renderer) {
      process.stdout.write(` found  ${fmtSec(Date.now() - tr)}  (${renderer})\n`);
    } else {
      process.stdout.write(` none found  ${fmtSec(Date.now() - tr)}\n`);
      console.log("         Warning: No PDF renderer (pdftoppm / mutool).");
      console.log("         Review pass will be skipped — pages will be unchanged.");
      console.log("         Install via: brew install poppler");
    }
    console.log("");

    // ── Per-page review loop ────────────────────────────────────────────────
    const tmpDir = join(tmpdir(), `qwen3-review-${randomUUID()}`);
    await mkdir(tmpDir, { recursive: true });

    const pageTimes: number[] = [];

    try {
      for (let pageNum = 1; pageNum <= numPages; pageNum++) {
        const pageKey = String(pageNum);
        const pct = Math.round((pageNum / numPages) * 100);

        if (progress.pages[pageKey]) {
          if (verbose) {
            console.log(
              `[Review] Page ${pageNum}/${numPages} (${pct}%) — already done, skipping.`,
            );
          }
          continue;
        }

        const pageStart = Date.now();
        const wordText = wordPages[pageNum - 1] ?? "";
        console.log(`[Review] ${fmtClock(pageStart)}  Page ${pageNum}/${numPages} (${pct}%):`);
        console.log(
          `         Word text:  ${wordText.length} chars  (${wordText.split("\n").length} line(s))`,
        );

        // ── Render PDF page ───────────────────────────────────────────────
        let pageImage: { base64: string; mimeType: "image/jpeg" } | null = null;
        let renderMs = 0;
        if (renderer) {
          const t = Date.now();
          process.stdout.write(`         ${fmtClock(t)}  Rendering to JPEG...`);
          try {
            pageImage = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
            renderMs = Date.now() - t;
            const kb = Math.round((pageImage.base64.length * 0.75) / 1024);
            process.stdout.write(` done  ${fmtSec(renderMs)}  (${kb} KB)\n`);
          } catch (renderErr) {
            renderMs = Date.now() - t;
            process.stdout.write(` failed  ${fmtSec(renderMs)}\n`);
            console.log(
              `         Error: ${renderErr instanceof Error ? renderErr.message : String(renderErr)}`,
            );
          }
        }

        // ── Qwen3 review ──────────────────────────────────────────────────
        let correctedText = wordText; // default: keep original if no image
        let reviewMs = 0;
        if (pageImage) {
          const t = Date.now();
          process.stdout.write(`         ${fmtClock(t)}  Qwen3-VL review...`);
          try {
            correctedText = await reviewPage(
              client,
              qwen3ModelId,
              pageNum,
              wordText,
              pageImage,
              verbose,
            );
            reviewMs = Date.now() - t;
            const delta = correctedText.length - wordText.length;
            const deltaStr = delta >= 0 ? `+${delta}` : `${delta}`;
            process.stdout.write(
              ` done  ${fmtSec(reviewMs)}  (${correctedText.length} chars, Δ${deltaStr})\n`,
            );
          } catch (reviewErr) {
            reviewMs = Date.now() - t;
            process.stdout.write(` failed  ${fmtSec(reviewMs)}  — keeping original\n`);
            console.log(
              `         Error: ${reviewErr instanceof Error ? reviewErr.message : String(reviewErr)}`,
            );
            correctedText = wordText;
          }
        } else {
          console.log(`         No renderer — keeping original Word text for this page.`);
        }

        // ── Save progress ─────────────────────────────────────────────────
        {
          const t = Date.now();
          process.stdout.write(`         ${fmtClock(t)}  Saving progress...`);
          progress.pages[pageKey] = { wordText, correctedText };
          await saveProgress(progressPath, progress);
          const saveMs = Date.now() - t;
          process.stdout.write(` done  ${fmtSec(saveMs)}\n`);
        }

        // ── Page summary + ETA ────────────────────────────────────────────
        const pageMs = Date.now() - pageStart;
        pageTimes.push(pageMs);

        const avgMs = pageTimes.reduce((a, b) => a + b, 0) / pageTimes.length;
        const remaining = numPages - pageNum;

        console.log(`         ${DIVIDER}`);
        const parts = [`render ${fmtSec(renderMs)}`, `review ${fmtSec(reviewMs)}`];

        if (remaining > 0) {
          console.log(
            `         Page ${pageNum} done in ${fmtSec(pageMs)}  [${parts.join("  ·  ")}]`,
          );
          console.log(
            `         avg ${fmtSec(avgMs)}/page  ·  ${remaining} page(s) remaining  ·  ETA ${fmtEta(avgMs * remaining)}`,
          );
        } else {
          console.log(
            `         Page ${pageNum} done in ${fmtSec(pageMs)}  [${parts.join("  ·  ")}]`,
          );
        }
        console.log("");
      }
    } finally {
      await rm(tmpDir, { recursive: true, force: true });
    }

    // ── Build corrected Word document ───────────────────────────────────────
    {
      const t = Date.now();
      process.stdout.write(`[Review] ${fmtClock(t)}  Building corrected Word document...`);
      const orderedPages = Array.from({ length: numPages }, (_, i) => ({
        pageNum: i + 1,
        text: progress.pages[String(i + 1)]?.correctedText ?? wordPages[i] ?? "",
      }));
      await generateWordDoc(orderedPages, outputPath, verbose);
      const buildMs = Date.now() - t;
      process.stdout.write(` done  ${fmtSec(buildMs)}\n`);
    }

    // ── Final summary ─────────────────────────────────────────────────────
    const totalMs = Date.now() - docStart;
    const processed = pageTimes.length;
    const avgPerPage =
      processed > 0
        ? fmtSec(pageTimes.reduce((a, b) => a + b, 0) / processed)
        : "n/a";

    console.log("");
    console.log(`${"═".repeat(56)}`);
    console.log(`  Done!`);
    console.log(`  Corrected output: ${outputPath}`);
    console.log(`  Progress file:    ${progressPath}`);
    console.log(`  (Delete progress file to re-review from scratch.)`);
    console.log("");
    console.log(`  Pages reviewed:  ${processed} of ${numPages}`);
    console.log(`  Avg per page:    ${avgPerPage}`);
    console.log(`  Total elapsed:   ${fmtSec(totalMs)}`);
    console.log(`${"═".repeat(56)}`);
  }
}
