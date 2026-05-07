/**
 * qwen-pdf-to-word.ts
 *
 * Single-model PDF → Word (.docx) pipeline using Qwen2.5-VL.
 *
 * This is a lean, single-pass alternative to the 4-model corroboration pipeline
 * (`pdf-to-word.ts`). Instead of running four OCR models in parallel and then a
 * fifth corroboration call, every page is sent once to Qwen2.5-VL-7B-Instruct,
 * which is a multimodal vision-language model strong enough to extract Arabic and
 * Latin text accurately in a single call.
 *
 * Pipeline per page:
 *   1. Render page to JPEG (pdftoppm / mutool)
 *   2. Qwen2.5-VL extracts all text from the rendered image
 *   3. Strip any thinking/reasoning traces from the response
 *   4. Save to progress file (resumable)
 *
 * After all pages are done, the Word document is assembled by the shared
 * word-generator with full Arabic RTL support.
 *
 * Progress file: <output>.qwen-progress.json
 * Delete it to force a full re-run from scratch.
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
import { QwenProgressSchema, type QwenProgress } from "./types.js";
import { generateWordDoc } from "./word-generator.js";

// ---------------------------------------------------------------------------
// Qwen2.5-VL extraction prompt
// ---------------------------------------------------------------------------

const QWEN_OCR_SYSTEM_PROMPT = [
  "You are a precise document text extractor specializing in Arabic and Latin text.",
  "Extract all text from the document page image exactly as it appears.",
  "",
  "STRUCTURE RULES:",
  "- Headings and titles: place on their own line, preceded and followed by a blank line.",
  "- Body paragraphs: separate with blank lines.",
  "- Tables: use tab characters (\\t) between columns, newlines between rows. Include the header row.",
  "- Lists: preserve markers (1., 2., •, -, etc.) and indentation.",
  "- Key-value pairs: Key: Value, one per line.",
  "",
  "LANGUAGE:",
  "- Arabic text: preserve every word exactly. Maintain the correct right-to-left character order within each word.",
  "- Latin text, numbers, punctuation, and symbols: preserve exactly as shown in the image.",
  "- Do not translate, transliterate, or summarize anything.",
  "",
  "OUTPUT:",
  "- Extracted text only. No commentary, no explanations, no markdown code fences, no reasoning traces.",
  "- Preserve the blank lines that reflect the document's logical paragraph and section structure.",
].join("\n");

// ---------------------------------------------------------------------------
// Timing helpers
// ---------------------------------------------------------------------------

/** Format a duration in milliseconds as "X.Xs". */
function fmtSec(ms: number): string {
  return `${(ms / 1000).toFixed(1)}s`;
}

/**
 * Format an ETA in milliseconds as a human-readable string.
 * Examples: "~8s", "~1m 24s", "~12m 03s"
 */
function fmtEta(ms: number): string {
  const s = Math.round(ms / 1000);
  if (s < 60) return `~${s}s`;
  const m = Math.floor(s / 60);
  const rem = s % 60;
  return `~${m}m ${rem < 10 ? "0" : ""}${rem}s`;
}

/**
 * Format a wall-clock timestamp as "HH:MM:SS".
 * Used to mark when each step started.
 */
function fmtClock(ts: number): string {
  const d = new Date(ts);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

/** Divider line used between page footer items. */
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

async function loadQwenProgress(
  progressPath: string,
  pdfPath: string,
  totalPages: number,
): Promise<QwenProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = QwenProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath === pdfPath &&
        parsed.data.totalPages === totalPages
      ) {
        const completed = Object.keys(parsed.data.pages).length;
        if (completed > 0) {
          console.log(
            `[Qwen] Resuming: ${completed}/${totalPages} page(s) already done — loading from progress file.`,
          );
        }
        return parsed.data;
      }
      console.log("[Qwen] Progress file found but does not match this PDF — starting fresh.");
    } catch {
      console.log("[Qwen] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, totalPages, pages: {} };
}

async function saveQwenProgress(
  progressPath: string,
  progress: QwenProgress,
): Promise<void> {
  await mkdir(dirname(progressPath), { recursive: true });
  await writeFile(progressPath, JSON.stringify(progress, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Single-pass Qwen extraction
// ---------------------------------------------------------------------------

async function extractPageWithQwen(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(`[Qwen] Page ${pageNum}: sending to ${modelId}...`);
  }

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: QWEN_OCR_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: {
              url: `data:${pageImage.mimeType};base64,${pageImage.base64}`,
            },
          },
          {
            type: "text" as const,
            text: "Extract all text from this document page.",
          },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const raw = response.choices[0]?.message.content ?? "";
  const text = stripThinking(raw);

  if (verbose) {
    const tokensUsed = response.usage?.total_tokens ?? "unknown";
    console.log(`[Qwen] Page ${pageNum}: ${text.length} chars, ${tokensUsed} tokens`);
  }

  return text;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface QwenConvertArgs {
  pdfPath: string;
  outputPath: string;
  progressPath: string;
  client: OpenAI;
  qwenModelId: string;
  verbose: boolean;
}

/**
 * Convert a PDF to a Word document using a single Qwen2.5-VL pass per page.
 * Progress is saved after every page and the conversion is fully resumable.
 */
export async function convertPdfToWordWithQwen(args: QwenConvertArgs): Promise<void> {
  const { pdfPath, outputPath, progressPath, client, qwenModelId, verbose } = args;

  const docStart = Date.now();

  // ── Step 1: count pages ───────────────────────────────────────────────────
  {
    const t0 = Date.now();
    process.stdout.write(`[Qwen] ${fmtClock(t0)}  Step 1 — Counting pages...`);
    const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
    const elapsed = fmtSec(Date.now() - t0);
    process.stdout.write(` done  ${elapsed}  (${numPages} page(s))\n`);
    console.log("");

    // ── Load progress ───────────────────────────────────────────────────────
    const progress = await loadQwenProgress(progressPath, pdfPath, numPages);

    // ── Detect renderer ─────────────────────────────────────────────────────
    {
      const tr = Date.now();
      process.stdout.write(`[Qwen] ${fmtClock(tr)}  Detecting PDF renderer...`);
      const renderer: Renderer | null = await detectRenderer();
      const rendererElapsed = fmtSec(Date.now() - tr);
      if (renderer) {
        process.stdout.write(` found  ${rendererElapsed}  (${renderer})\n`);
      } else {
        process.stdout.write(` none found  ${rendererElapsed}\n`);
        console.log("         Warning: No PDF renderer (pdftoppm / mutool).");
        console.log("         OCR pass will be skipped — pages will be blank.");
        console.log("         Install via: brew install poppler  (provides pdftoppm)");
      }
      console.log("");

      // ── Per-page extraction loop ──────────────────────────────────────────
      const tmpDir = join(tmpdir(), `qwen-word-${randomUUID()}`);
      await mkdir(tmpDir, { recursive: true });

      // Rolling page-time tracker for ETA estimation.
      const pageTimes: number[] = [];

      try {
        for (let pageNum = 1; pageNum <= numPages; pageNum++) {
          const pageKey = String(pageNum);
          const pct = Math.round((pageNum / numPages) * 100);

          if (progress.pages[pageKey]) {
            if (verbose) {
              console.log(
                `[Qwen] Page ${pageNum}/${numPages} (${pct}%) — already done, skipping.`,
              );
            }
            continue;
          }

          const pageStart = Date.now();
          console.log(`[Qwen] ${fmtClock(pageStart)}  Page ${pageNum}/${numPages} (${pct}%):`);

          // ── Render page ─────────────────────────────────────────────────────
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

          // ── Qwen inference ──────────────────────────────────────────────────
          let text = "";
          let inferMs = 0;
          if (pageImage) {
            const t = Date.now();
            process.stdout.write(`         ${fmtClock(t)}  Qwen2.5-VL inference...`);
            try {
              text = await extractPageWithQwen(client, qwenModelId, pageNum, pageImage, verbose);
              inferMs = Date.now() - t;
              process.stdout.write(` done  ${fmtSec(inferMs)}  (${text.length} chars)\n`);
            } catch (extractErr) {
              inferMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(inferMs)}\n`);
              console.log(
                `         Error: ${extractErr instanceof Error ? extractErr.message : String(extractErr)}`,
              );
            }
          } else {
            console.log(`         No renderer — page will be blank.`);
          }

          // ── Strip thinking + save progress ──────────────────────────────────
          {
            const t = Date.now();
            process.stdout.write(`         ${fmtClock(t)}  Saving progress...`);
            progress.pages[pageKey] = { text };
            await saveQwenProgress(progressPath, progress);
            const saveMs = Date.now() - t;
            process.stdout.write(` done  ${fmtSec(saveMs)}\n`);
          }

          // ── Page summary + ETA ──────────────────────────────────────────────
          const pageMs = Date.now() - pageStart;
          pageTimes.push(pageMs);

          const avgMs = pageTimes.reduce((a, b) => a + b, 0) / pageTimes.length;
          const remaining = numPages - pageNum;

          console.log(`         ${DIVIDER}`);

          const parts = [
            `render ${fmtSec(renderMs)}`,
            `inference ${fmtSec(inferMs)}`,
          ];

          if (remaining > 0) {
            const etaStr = fmtEta(avgMs * remaining);
            console.log(
              `         Page ${pageNum} done in ${fmtSec(pageMs)}  [${parts.join("  ·  ")}]`,
            );
            console.log(
              `         avg ${fmtSec(avgMs)}/page  ·  ${remaining} page(s) remaining  ·  ETA ${etaStr}`,
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

      // ── Generate Word document ────────────────────────────────────────────
      {
        const t = Date.now();
        process.stdout.write(`[Qwen] ${fmtClock(t)}  Building Word document...`);
        const orderedPages = Array.from({ length: numPages }, (_, i) => ({
          pageNum: i + 1,
          text: progress.pages[String(i + 1)]?.text ?? "",
        }));
        await generateWordDoc(orderedPages, outputPath, verbose);
        const buildMs = Date.now() - t;
        process.stdout.write(` done  ${fmtSec(buildMs)}\n`);
      }

      // ── Final summary ─────────────────────────────────────────────────────
      const totalMs = Date.now() - docStart;
      const processedPages = pageTimes.length;
      const avgPerPage =
        processedPages > 0
          ? fmtSec(pageTimes.reduce((a, b) => a + b, 0) / processedPages)
          : "n/a";

      console.log("");
      console.log(`${"═".repeat(56)}`);
      console.log(`  Done!`);
      console.log(`  Output:          ${outputPath}`);
      console.log(`  Progress file:   ${progressPath}`);
      console.log(`  (Delete the progress file to re-process from scratch.)`);
      console.log("");
      console.log(`  Pages processed: ${processedPages} of ${numPages}`);
      console.log(`  Avg per page:    ${avgPerPage}`);
      console.log(`  Total elapsed:   ${fmtSec(totalMs)}`);
      console.log(`${"═".repeat(56)}`);
    }
  }
}
