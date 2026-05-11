/**
 * qwen3-epub.ts
 *
 * PDF → EPUB pipeline using Qwen3-VL.
 *
 * For each page the model is asked to return a clean HTML5 fragment
 * (headings, paragraphs, tables, lists) with dir="rtl"/"ltr" attributes
 * on every block element. The fragments are assembled into a standard
 * EPUB 3 file by epub-generator.ts — one XHTML file per PDF page —
 * making it straightforward for a human reviewer to open any page,
 * compare it against the source PDF, and edit the HTML directly.
 *
 * Pipeline per page:
 *   1. Render page to JPEG (pdftoppm / mutool)
 *   2. Qwen3-VL returns an HTML fragment for the page
 *   3. Clean and normalise the fragment (strip fences, fix void elements)
 *   4. Save to progress file (resumable)
 *
 * After all pages: assemble the EPUB package via generateEpub().
 *
 * Progress file: <output>.epub-progress.json
 * Delete it to force a full re-run.
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
import { EpubProgressSchema, type EpubProgress } from "./types.js";
import { generateEpub, cleanHtmlFragment } from "./epub-generator.js";

// ---------------------------------------------------------------------------
// Extraction prompt
// ---------------------------------------------------------------------------

const EPUB_SYSTEM_PROMPT = [
  "You are a document page-to-HTML converter specializing in Arabic and Latin text.",
  "Your output will be embedded directly into an EPUB XHTML page file.",
  "",
  "Convert the document page image to a clean HTML5 content fragment.",
  "",
  "STRICT OUTPUT FORMAT:",
  "- Output ONLY the HTML content — no <html>, <head>, <body>, or <!DOCTYPE> tags.",
  "- No markdown fences (` ``` `), no XML declarations, no reasoning traces.",
  "",
  "ELEMENTS TO USE (use only these):",
  "  Headings:   <h2 dir=\"rtl\"> main title/chapter head | <h3 dir=\"rtl\"> sub-heading",
  "              (use dir=\"ltr\" for Latin-only headings)",
  "  Paragraphs: <p dir=\"rtl\"> Arabic  |  <p dir=\"ltr\"> Latin/English/numbers",
  "  Tables:     <table><thead><tr><th dir=\"rtl\">…</th></tr></thead>",
  "                     <tbody><tr><td dir=\"rtl\">…</td></tr></tbody></table>",
  "  Lists:      <ul dir=\"rtl\"><li>…</li></ul>  or  <ol dir=\"rtl\"><li>…</li></ol>",
  "  Key-value:  <p dir=\"rtl\"><strong>Key:</strong> Value</p>",
  "  Divider:    <hr/>",
  "",
  "DIRECTION (required on EVERY block element):",
  "  dir=\"rtl\" — any element containing Arabic characters",
  "  dir=\"ltr\" — any element with only Latin, English, or numeric content",
  "",
  "ACCURACY (non-negotiable):",
  "  - Reproduce every Arabic character exactly as visible in the image.",
  "  - Reproduce every Latin character, number, and punctuation mark exactly.",
  "  - Do NOT translate, transliterate, paraphrase, or summarize anything.",
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
  return [d.getHours(), d.getMinutes(), d.getSeconds()]
    .map((n) => String(n).padStart(2, "0"))
    .join(":");
}

const DIVIDER = "─".repeat(56);

// ---------------------------------------------------------------------------
// Progress file helpers
// ---------------------------------------------------------------------------

async function fileExists(p: string): Promise<boolean> {
  try { await access(p); return true; } catch { return false; }
}

async function loadProgress(
  progressPath: string,
  pdfPath: string,
  totalPages: number,
): Promise<EpubProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = EpubProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath === pdfPath &&
        parsed.data.totalPages === totalPages
      ) {
        const done = Object.keys(parsed.data.pages).length;
        if (done > 0)
          console.log(`[EPUB] Resuming: ${done}/${totalPages} page(s) already done.`);
        return parsed.data;
      }
      console.log("[EPUB] Progress file doesn't match this PDF — starting fresh.");
    } catch {
      console.log("[EPUB] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, totalPages, pages: {} };
}

async function saveProgress(path: string, p: EpubProgress): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, JSON.stringify(p, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Single-page Qwen3 HTML extraction
// ---------------------------------------------------------------------------

async function extractPageHtml(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: EPUB_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: { url: `data:${pageImage.mimeType};base64,${pageImage.base64}` },
          },
          {
            type: "text" as const,
            text: "Convert this document page to an HTML fragment for an EPUB file.",
          },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const raw = response.choices[0]?.message.content ?? "";
  const html = stripThinking(raw);

  if (verbose) {
    const tokens = response.usage?.total_tokens ?? "?";
    console.log(`[EPUB] Page ${pageNum}: ${html.length} chars, ${tokens} tokens`);
  }

  return html;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface QwenEpubArgs {
  pdfPath: string;
  outputPath: string;
  progressPath: string;
  client: OpenAI;
  modelId: string;
  verbose: boolean;
}

/**
 * Convert a PDF to an EPUB using Qwen3-VL for per-page HTML extraction.
 * Progress is saved after every page — fully resumable.
 */
export async function convertPdfToEpub(args: QwenEpubArgs): Promise<void> {
  const { pdfPath, outputPath, progressPath, client, modelId, verbose } = args;

  const docStart = Date.now();

  // ── Step 1: count pages ───────────────────────────────────────────────────
  {
    const t0 = Date.now();
    process.stdout.write(`[EPUB] ${fmtClock(t0)}  Step 1 — Counting pages...`);
    const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
    process.stdout.write(` done  ${fmtSec(Date.now() - t0)}  (${numPages} page(s))\n`);
    console.log("");

    // ── Load progress ───────────────────────────────────────────────────────
    const progress = await loadProgress(progressPath, pdfPath, numPages);

    // ── Detect renderer ─────────────────────────────────────────────────────
    {
      const tr = Date.now();
      process.stdout.write(`[EPUB] ${fmtClock(tr)}  Detecting PDF renderer...`);
      const renderer: Renderer | null = await detectRenderer();
      if (renderer) {
        process.stdout.write(` found  ${fmtSec(Date.now() - tr)}  (${renderer})\n`);
      } else {
        process.stdout.write(` none found  ${fmtSec(Date.now() - tr)}\n`);
        console.log("         Warning: No PDF renderer (pdftoppm / mutool).");
        console.log("         OCR pass skipped — pages will be empty.");
        console.log("         Install: brew install poppler");
      }
      console.log("");

      // ── Per-page extraction loop ──────────────────────────────────────────
      const tmpDir = join(tmpdir(), `qwen3-epub-${randomUUID()}`);
      await mkdir(tmpDir, { recursive: true });

      const pageTimes: number[] = [];

      try {
        for (let pageNum = 1; pageNum <= numPages; pageNum++) {
          const pageKey = String(pageNum);
          const pct = Math.round((pageNum / numPages) * 100);

          if (progress.pages[pageKey]) {
            if (verbose)
              console.log(`[EPUB] Page ${pageNum}/${numPages} (${pct}%) — already done, skipping.`);
            continue;
          }

          const pageStart = Date.now();
          console.log(`[EPUB] ${fmtClock(pageStart)}  Page ${pageNum}/${numPages} (${pct}%):`);

          // ── Render ─────────────────────────────────────────────────────────
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
            } catch (err) {
              renderMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(renderMs)}\n`);
              console.log(`         Error: ${err instanceof Error ? err.message : String(err)}`);
            }
          }

          // ── Qwen3 HTML extraction ───────────────────────────────────────────
          let html = "";
          let inferMs = 0;
          if (pageImage) {
            const t = Date.now();
            process.stdout.write(`         ${fmtClock(t)}  Qwen3-VL → HTML...`);
            try {
              const raw = await extractPageHtml(client, modelId, pageNum, pageImage, verbose);
              html = cleanHtmlFragment(raw);
              inferMs = Date.now() - t;
              process.stdout.write(` done  ${fmtSec(inferMs)}  (${html.length} chars)\n`);
            } catch (err) {
              inferMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(inferMs)}\n`);
              console.log(`         Error: ${err instanceof Error ? err.message : String(err)}`);
              html = `<p dir="rtl"><!-- extraction failed for page ${pageNum} --></p>`;
            }
          } else {
            html = `<p dir="rtl"><!-- no renderer — page ${pageNum} skipped --></p>`;
            console.log(`         No renderer — placeholder inserted.`);
          }

          // ── Save progress ───────────────────────────────────────────────────
          {
            const t = Date.now();
            process.stdout.write(`         ${fmtClock(t)}  Saving progress...`);
            progress.pages[pageKey] = { html };
            await saveProgress(progressPath, progress);
            const saveMs = Date.now() - t;
            process.stdout.write(` done  ${fmtSec(saveMs)}\n`);
          }

          // ── Page summary + ETA ──────────────────────────────────────────────
          const pageMs = Date.now() - pageStart;
          pageTimes.push(pageMs);

          const avgMs = pageTimes.reduce((a, b) => a + b, 0) / pageTimes.length;
          const remaining = numPages - pageNum;

          console.log(`         ${DIVIDER}`);
          const parts = [`render ${fmtSec(renderMs)}`, `inference ${fmtSec(inferMs)}`];

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

      // ── Assemble EPUB ─────────────────────────────────────────────────────
      {
        const t = Date.now();
        process.stdout.write(`[EPUB] ${fmtClock(t)}  Assembling EPUB package...`);

        const orderedPages = Array.from({ length: numPages }, (_, i) => ({
          pageNum: i + 1,
          html: progress.pages[String(i + 1)]?.html ?? "",
        }));

        await generateEpub(orderedPages, outputPath, pdfPath, verbose);
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
      console.log(`  EPUB:            ${outputPath}`);
      console.log(`  Progress file:   ${progressPath}`);
      console.log(`  (Delete progress file to re-convert from scratch.)`);
      console.log("");
      console.log(`  Pages:           ${processed} extracted of ${numPages}`);
      console.log(`  Avg per page:    ${avgPerPage}`);
      console.log(`  Total elapsed:   ${fmtSec(totalMs)}`);
      console.log(`${"═".repeat(56)}`);
    }
  }
}
