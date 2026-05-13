/**
 * omlx-pdf-to-word.ts
 *
 * Arabic-book-optimised PDF → Word (.docx) pipeline for the omlx server
 * running Qwen3-VL-30B-A3B-Instruct-MLX-8bit (or any Qwen3-VL compatible model).
 *
 * ═══════════════════════════════════════════════════════════════
 * Prompt engineering — techniques from the Qwen3-VL OCR cookbook
 * ═══════════════════════════════════════════════════════════════
 *
 * 1. NO SYSTEM PROMPT
 *    The official Qwen3-VL cookbook explicitly comments out the system prompt
 *    in every OCR example. A system prompt can distract the model from pure
 *    text extraction and introduce hallucinations or refusals. This pipeline
 *    sends NO system message.
 *
 * 2. IMAGE BEFORE TEXT
 *    The cookbook places the image token first in the user content array,
 *    then the text prompt. This matches the model's training distribution.
 *
 * 3. SHORT, DIRECT PROMPT (cookbook technique for multilingual / Arabic OCR)
 *    The cookbook uses prompts like:
 *      "Read all the text in the image."
 *      "Please output only the text content from the image without any
 *       additional descriptions."
 *    We extend this minimally for Arabic books to ask for paragraph
 *    preservation and RTL reading order, while keeping the instruction
 *    extremely concise.
 *
 * 4. HIGH-CONTEXT OUTPUT (large max_tokens)
 *    A dense Arabic page can produce 1,000–3,000 tokens of text. The 30B
 *    model can handle it; we give it plenty of room with 8192 max_tokens.
 *
 * 5. TEMPERATURE 0 — deterministic, reproducible OCR output.
 *
 * ═══════════════════════════════════════════════════════════════
 * Pipeline
 * ═══════════════════════════════════════════════════════════════
 *   1. Render each PDF page to JPEG (pdftoppm / mutool, existing renderer)
 *   2. Send: [image, short Arabic OCR prompt] — NO system prompt
 *   3. Strip any residual thinking traces from the response
 *   4. Save text to progress file (fully resumable)
 *   5. Assemble Word document with Arabic RTL support (word-generator)
 *
 * Progress file:  <output>.omlx-progress.json
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
// OCR prompt  (Qwen3-VL cookbook techniques)
// ---------------------------------------------------------------------------

/**
 * Technique from cookbook section 2 — "Full Page OCR for Multilingual text":
 *   "Please output only the text content from the image without any additional
 *    descriptions."
 *
 * Extended minimally for Arabic scanned books:
 *   • Preserve paragraph structure (blank lines between logical blocks)
 *   • Preserve reading order (RTL for Arabic)
 *   • No commentary, no translation, no summaries
 *
 * NO SYSTEM PROMPT — following cookbook best-practice of skipping it entirely
 * for OCR tasks.
 */
const ARABIC_BOOK_OCR_PROMPT = [
  "Please output only the text content from this Arabic book page.",
  "Preserve paragraph structure and the natural reading order of the text.",
  "Do not add descriptions, commentary, translations, or any content not",
  "visible in the image.",
].join(" ");

// ---------------------------------------------------------------------------
// Timing helpers (shared pattern across all pipeline files)
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
  return [d.getHours(), d.getMinutes(), d.getSeconds()]
    .map((n) => String(n).padStart(2, "0"))
    .join(":");
}

const DIVIDER = "─".repeat(56);

// ---------------------------------------------------------------------------
// Progress file helpers
// ---------------------------------------------------------------------------

async function fileExists(p: string): Promise<boolean> {
  try {
    await access(p);
    return true;
  } catch {
    return false;
  }
}

async function loadProgress(
  progressPath: string,
  pdfPath: string,
  totalPages: number,
): Promise<QwenProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw    = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = QwenProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath    === pdfPath &&
        parsed.data.totalPages === totalPages
      ) {
        const done = Object.keys(parsed.data.pages).length;
        if (done > 0)
          console.log(`[OmlxWord] Resuming: ${done}/${totalPages} page(s) already done.`);
        return parsed.data;
      }
      console.log("[OmlxWord] Progress file doesn't match this PDF — starting fresh.");
    } catch {
      console.log("[OmlxWord] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, totalPages, pages: {} };
}

async function saveProgress(
  progressPath: string,
  progress: QwenProgress,
): Promise<void> {
  await mkdir(dirname(progressPath), { recursive: true });
  await writeFile(progressPath, JSON.stringify(progress, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Single-page model call  (cookbook technique: image first, no system prompt)
// ---------------------------------------------------------------------------

async function extractPageWithOmlx(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  // ── Cookbook technique 2: image BEFORE text in the content array ──────────
  const response = await client.chat.completions.create({
    model: modelId,
    // ── NO system prompt (cookbook explicitly skips it for OCR) ──────────────
    messages: [
      {
        role: "user",
        content: [
          // Image first — matches model's training distribution
          {
            type: "image_url" as const,
            image_url: {
              url: `data:${pageImage.mimeType};base64,${pageImage.base64}`,
            },
          },
          // Short, direct prompt — no meta-instructions, no role description
          {
            type: "text" as const,
            text: ARABIC_BOOK_OCR_PROMPT,
          },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    // Large context budget — dense Arabic page ≈ 1 000–3 000 tokens
    max_tokens: 8192,
    // Temperature 0 for deterministic, reproducible OCR
    temperature: 0,
  });

  const raw  = response.choices[0]?.message.content ?? "";
  const text = stripThinking(raw);

  if (verbose) {
    const tokens = response.usage?.total_tokens ?? "?";
    console.log(`[OmlxWord] Page ${pageNum}: ${text.length} chars, ${tokens} tokens`);
  }

  return text;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface OmlxWordArgs {
  pdfPath:      string;
  outputPath:   string;
  progressPath: string;
  client:       OpenAI;
  modelId:      string;
  verbose:      boolean;
}

/**
 * Convert a scanned Arabic PDF book to a Word document using Qwen3-VL-30B
 * served by the local omlx server.
 *
 * Applies the Qwen3-VL OCR cookbook techniques:
 *   • No system prompt
 *   • Image before text in content array
 *   • Short, direct Arabic OCR user prompt
 *   • temperature=0, max_tokens=8192
 *
 * Progress is saved after every page — fully resumable.
 */
export async function convertPdfToWordWithOmlx(args: OmlxWordArgs): Promise<void> {
  const { pdfPath, outputPath, progressPath, client, modelId, verbose } = args;
  const docStart = Date.now();

  // ── Step 1: Count pages ──────────────────────────────────────────────────
  {
    const t0 = Date.now();
    process.stdout.write(`[OmlxWord] ${fmtClock(t0)}  Step 1 — Counting pages...`);
    const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
    process.stdout.write(` done  ${fmtSec(Date.now() - t0)}  (${numPages} page(s))\n`);
    console.log("");

    const progress = await loadProgress(progressPath, pdfPath, numPages);

    // ── Detect renderer ───────────────────────────────────────────────────
    {
      const tr = Date.now();
      process.stdout.write(`[OmlxWord] ${fmtClock(tr)}  Detecting PDF renderer...`);
      const renderer: Renderer | null = await detectRenderer();
      if (renderer) {
        process.stdout.write(` found  ${fmtSec(Date.now() - tr)}  (${renderer})\n`);
      } else {
        process.stdout.write(` none found  ${fmtSec(Date.now() - tr)}\n`);
        console.log("           Warning: No PDF renderer (pdftoppm / mutool).");
        console.log("           OCR will be skipped — pages will be blank.");
        console.log("           Install: brew install poppler   (provides pdftoppm)");
      }
      console.log("");

      // ── Prompt in use ─────────────────────────────────────────────────────
      if (verbose) {
        console.log("[OmlxWord] Prompt (no system message):");
        console.log(`           "${ARABIC_BOOK_OCR_PROMPT}"`);
        console.log("");
      }

      // ── Per-page extraction loop ──────────────────────────────────────────
      const tmpDir = join(tmpdir(), `omlx-word-${randomUUID()}`);
      await mkdir(tmpDir, { recursive: true });
      const pageTimes: number[] = [];

      try {
        for (let pageNum = 1; pageNum <= numPages; pageNum++) {
          const pageKey = String(pageNum);
          const pct     = Math.round((pageNum / numPages) * 100);

          if (progress.pages[pageKey]) {
            if (verbose)
              console.log(`[OmlxWord] Page ${pageNum}/${numPages} (${pct}%) — already done, skipping.`);
            continue;
          }

          const pageStart = Date.now();
          console.log(`[OmlxWord] ${fmtClock(pageStart)}  Page ${pageNum}/${numPages} (${pct}%):`);

          // ── Render ─────────────────────────────────────────────────────────
          let pageImage: { base64: string; mimeType: "image/jpeg" } | null = null;
          let renderMs = 0;
          if (renderer) {
            const t = Date.now();
            process.stdout.write(`           ${fmtClock(t)}  Rendering to JPEG...`);
            try {
              pageImage = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
              renderMs  = Date.now() - t;
              const kb  = Math.round((pageImage.base64.length * 0.75) / 1024);
              process.stdout.write(` done  ${fmtSec(renderMs)}  (${kb} KB)\n`);
            } catch (err) {
              renderMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(renderMs)}\n`);
              console.log(`           Error: ${err instanceof Error ? err.message : String(err)}`);
            }
          }

          // ── Qwen3-VL inference (no system prompt, image first) ─────────────
          let text    = "";
          let inferMs = 0;
          if (pageImage) {
            const t = Date.now();
            process.stdout.write(`           ${fmtClock(t)}  Qwen3-VL OCR (no sys prompt, image-first)...`);
            try {
              text    = await extractPageWithOmlx(client, modelId, pageNum, pageImage, verbose);
              inferMs = Date.now() - t;
              process.stdout.write(` done  ${fmtSec(inferMs)}  (${text.length} chars)\n`);
            } catch (err) {
              inferMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(inferMs)}\n`);
              console.log(`           Error: ${err instanceof Error ? err.message : String(err)}`);
            }
          } else {
            console.log(`           No renderer — page will be blank in the Word document.`);
          }

          // ── Save progress ─────────────────────────────────────────────────
          {
            const t = Date.now();
            process.stdout.write(`           ${fmtClock(t)}  Saving progress...`);
            progress.pages[pageKey] = { text };
            await saveProgress(progressPath, progress);
            process.stdout.write(` done  ${fmtSec(Date.now() - t)}\n`);
          }

          // ── Page summary + ETA ────────────────────────────────────────────
          const pageMs   = Date.now() - pageStart;
          pageTimes.push(pageMs);
          const avgMs    = pageTimes.reduce((a, b) => a + b, 0) / pageTimes.length;
          const remaining = numPages - pageNum;

          console.log(`           ${DIVIDER}`);
          console.log(
            `           Page ${pageNum} done in ${fmtSec(pageMs)}  [render ${fmtSec(renderMs)}  ·  inference ${fmtSec(inferMs)}]`,
          );
          if (remaining > 0) {
            console.log(
              `           avg ${fmtSec(avgMs)}/page  ·  ${remaining} page(s) remaining  ·  ETA ${fmtEta(avgMs * remaining)}`,
            );
          }
          console.log("");
        }
      } finally {
        await rm(tmpDir, { recursive: true, force: true });
      }

      // ── Assemble Word document ────────────────────────────────────────────
      {
        const t = Date.now();
        process.stdout.write(`[OmlxWord] ${fmtClock(t)}  Building Word document...`);
        const orderedPages = Array.from({ length: numPages }, (_, i) => ({
          pageNum: i + 1,
          text:    progress.pages[String(i + 1)]?.text ?? "",
        }));
        await generateWordDoc(orderedPages, outputPath, verbose);
        process.stdout.write(` done  ${fmtSec(Date.now() - t)}\n`);
      }

      // ── Final summary ─────────────────────────────────────────────────────
      const totalMs    = Date.now() - docStart;
      const processed  = pageTimes.length;
      const avgPerPage =
        processed > 0
          ? fmtSec(pageTimes.reduce((a, b) => a + b, 0) / processed)
          : "n/a";

      console.log("");
      console.log(`${"═".repeat(56)}`);
      console.log(`  Done!`);
      console.log(`  Word document:   ${outputPath}`);
      console.log(`  Progress file:   ${progressPath}`);
      console.log(`  (Delete the progress file to re-convert from scratch.)`);
      console.log("");
      console.log(`  Pages processed: ${processed} of ${numPages}`);
      console.log(`  Avg per page:    ${avgPerPage}`);
      console.log(`  Total elapsed:   ${fmtSec(totalMs)}`);
      console.log(`${"═".repeat(56)}`);
    }
  }
}
