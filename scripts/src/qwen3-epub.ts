/**
 * qwen3-epub.ts
 *
 * PDF → EPUB pipeline using Qwen3-VL.
 *
 * For every page the model returns a SINGLE HTML fragment that combines
 * text elements AND visual elements in natural reading order.
 * Visual content (covers, illustrations, inline images) is represented as:
 *
 *   <figure data-region="x1,y1,x2,y2">
 *     <figcaption dir="rtl">optional caption</figcaption>
 *   </figure>
 *
 * where x1,y1 / x2,y2 are the top-left / bottom-right corners of the visual
 * element expressed as percentages of the page (0–100).
 *
 * After the model responds, processInlineImages() uses sharp to crop each
 * declared region from the rendered JPEG, saves the crop to the images
 * directory, and rewrites the figure to contain a proper <img> element.
 * Regions that cover ≥ 90% of the page get class="page-image";
 * smaller embedded visuals get class="inline-image".
 *
 * The final, processed HTML (with real <img> tags) is stored in the progress
 * file. The images directory persists across resume cycles.
 *
 * Pipeline per page:
 *   1. Render PDF page to JPEG  (pdftoppm / mutool)
 *   2. Qwen3-VL → HTML fragment (text + <figure data-region="…"> for visuals)
 *   3. processInlineImages()    (sharp crops; rewrites figures with <img>)
 *   4. Save processed HTML to progress file
 *
 * After all pages: assemble the EPUB via generateEpub().
 *
 * Progress file:   <output>.epub-progress.json
 * Images folder:   <output>.epub-images/
 * Delete both to force a full re-run.
 */

import { readFile, writeFile, mkdir, rm, access } from "node:fs/promises";
import { join, dirname } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import sharp from "sharp";
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
  "Convert the document page to a single HTML5 fragment that captures ALL content",
  "— text AND visuals — in natural reading order (top to bottom, right to left for Arabic).",
  "",
  "═══════════════════════════════════════════════════════════",
  "VISUAL ELEMENTS  (images, illustrations, photos, maps, diagrams, decorative art)",
  "═══════════════════════════════════════════════════════════",
  "For EVERY visual element on the page, insert a <figure> at the correct reading",
  "position relative to any surrounding text:",
  "",
  "  <figure data-region=\"x1,y1,x2,y2\">",
  "    <figcaption dir=\"rtl\">caption text if visible in the image</figcaption>",
  "  </figure>",
  "",
  "• x1,y1 = top-left corner of the visual element",
  "• x2,y2 = bottom-right corner of the visual element",
  "• All four values are INTEGER PERCENTAGES of the page dimensions (0–100).",
  "• Include a <figcaption> only if a caption is actually printed on the page.",
  "",
  "Examples:",
  "  Full-page cover with no text:  <figure data-region=\"0,0,100,100\"></figure>",
  "  Illustration in the top half:  <figure data-region=\"5,5,95,50\"></figure>",
  "  Small map embedded in text:    <figure data-region=\"10,40,60,65\"></figure>",
  "",
  "Include <figure> elements for: covers, illustrations, photographs, maps,",
  "diagrams, charts, ornamental borders/dividers, and any other non-text visual.",
  "",
  "═══════════════════════════════════════════════════════════",
  "TEXT ELEMENTS",
  "═══════════════════════════════════════════════════════════",
  "Use only these elements:",
  "  <h2 dir=\"rtl\"> / <h2 dir=\"ltr\">   main headings",
  "  <h3 dir=\"rtl\"> / <h3 dir=\"ltr\">   sub-headings",
  "  <p  dir=\"rtl\"> / <p  dir=\"ltr\">   paragraphs",
  "  <table> with <thead>/<tbody>, <tr>, <th dir=\"rtl\">, <td dir=\"rtl\">",
  "  <ul dir=\"rtl\"><li>…</li></ul>   or  <ol dir=\"rtl\"><li>…</li></ol>",
  "  <hr/>   decorative rule (text only, not for visual replacements)",
  "",
  "DIRECTION on every block element (required):",
  "  dir=\"rtl\" — element contains Arabic characters",
  "  dir=\"ltr\" — element contains only Latin / numeric content",
  "",
  "═══════════════════════════════════════════════════════════",
  "OUTPUT FORMAT",
  "═══════════════════════════════════════════════════════════",
  "• HTML content ONLY — no <html>, <head>, <body>, <!DOCTYPE>, no markdown fences.",
  "• No reasoning traces, no XML declarations.",
  "• Reproduce every Arabic character exactly as visible.",
  "• Reproduce every Latin character, number, and punctuation exactly.",
  "• Do NOT translate, transliterate, paraphrase, or summarize.",
].join("\n");

// ---------------------------------------------------------------------------
// Image region helpers
// ---------------------------------------------------------------------------

/**
 * Parse "x1,y1,x2,y2" (percentages) into normalised [0,1] fractions.
 * Returns null if the string is invalid.
 */
function parseRegion(
  attr: string,
): { left: number; top: number; width: number; height: number } | null {
  const parts = attr.split(",").map((s) => parseFloat(s.trim()));
  if (parts.length !== 4 || parts.some(isNaN)) return null;
  const [x1r, y1r, x2r, y2r] = parts as [number, number, number, number];
  const x1 = Math.max(0, Math.min(99, x1r));
  const y1 = Math.max(0, Math.min(99, y1r));
  const x2 = Math.max(x1 + 1, Math.min(100, x2r));
  const y2 = Math.max(y1 + 1, Math.min(100, y2r));
  return {
    left:   x1 / 100,
    top:    y1 / 100,
    width:  (x2 - x1) / 100,
    height: (y2 - y1) / 100,
  };
}

/**
 * Find all <figure data-region="…">…</figure> blocks in `html`.
 * For each one:
 *   • Crop the declared rectangle from the page JPEG using sharp.
 *   • Save the crop to `<imgFolder>/page_NNN_imgM.jpg`.
 *   • Replace the data-region figure with a proper <figure><img…/></figure>.
 *   • Regions covering ≥ 90% of the page in both dimensions → class="page-image".
 *   • Smaller regions → class="inline-image".
 *
 * Returns the processed HTML and the number of images successfully extracted.
 */
async function processInlineImages(
  html: string,
  pageJpegBase64: string,
  pageNum: number,
  imgFolder: string,
): Promise<{ html: string; count: number }> {
  const figRe = /<figure([^>]*?)data-region="([^"]+)"([^>]*)>([\s\S]*?)<\/figure>/gi;
  const matches = [...html.matchAll(figRe)];
  if (matches.length === 0) return { html, count: 0 };

  await mkdir(imgFolder, { recursive: true });
  const jpegBuffer = Buffer.from(pageJpegBase64, "base64");
  const meta = await sharp(jpegBuffer).metadata();
  const imgW = meta.width  ?? 1000;
  const imgH = meta.height ?? 1000;

  let result = html;
  let imgIndex = 0;
  let successCount = 0;

  for (const match of matches) {
    const fullBlock  = match[0]!;
    const regionStr  = match[2]!;
    const innerRaw   = (match[4] ?? "").trim();

    const region = parseRegion(regionStr);
    if (!region) {
      console.log(`         Warning: invalid data-region "${regionStr}" on page ${pageNum} — skipping.`);
      continue;
    }

    imgIndex++;
    const imgName = `page_${String(pageNum).padStart(3, "0")}_img${imgIndex}.jpg`;
    const imgPath = join(imgFolder, imgName);

    // Pixel coordinates (clamped to image bounds)
    const left   = Math.max(0, Math.round(region.left  * imgW));
    const top    = Math.max(0, Math.round(region.top   * imgH));
    const width  = Math.max(1, Math.min(imgW - left, Math.round(region.width  * imgW)));
    const height = Math.max(1, Math.min(imgH - top,  Math.round(region.height * imgH)));

    // A region covering ≥ 90% of the page in both dimensions is "full-page"
    const isFullPage = region.width >= 0.9 && region.height >= 0.9;
    const figClass   = isFullPage ? "page-image" : "inline-image";

    try {
      await sharp(jpegBuffer)
        .extract({ left, top, width, height })
        .jpeg({ quality: 88 })
        .toFile(imgPath);

      const imgEl   = `<img src="images/${imgName}" alt="illustration page ${pageNum}"/>`;
      const caption = innerRaw ? `\n  ${innerRaw}` : "";
      const newBlock = `<figure class="${figClass}">\n  ${imgEl}${caption}\n</figure>`;

      // Use a replacer function so $ in newBlock is treated literally
      result = result.replace(fullBlock, () => newBlock);
      successCount++;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.log(`         Warning: crop failed for page ${pageNum} img ${imgIndex}: ${msg}`);
      // Keep the figure but without <img> so the human reviewer can see the placeholder
      const fallback = `<figure class="${figClass}">\n  <!-- crop failed: ${msg} -->${innerRaw ? "\n  " + innerRaw : ""}\n</figure>`;
      result = result.replace(fullBlock, () => fallback);
    }
  }

  return { html: result, count: successCount };
}

/**
 * Directory where cropped JPEGs are stored.
 * Lives alongside the progress file so it survives resume cycles.
 *   /path/to/book.epub-progress.json  →  /path/to/book.epub-images/
 */
function imagesDir(progressPath: string): string {
  return progressPath.replace(/\.epub-progress\.json$/i, ".epub-images");
}

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
// Single-page Qwen3 extraction
// ---------------------------------------------------------------------------

async function extractPageRaw(
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
    console.log(`[EPUB] Page ${pageNum}: raw ${html.length} chars, ${tokens} tokens`);
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
 * Convert a PDF to an EPUB using Qwen3-VL for per-page extraction.
 *
 * Text and visual content are extracted together. Visuals declared by the
 * model with <figure data-region="…"> are cropped from the rendered JPEG
 * using sharp and embedded as real images in the EPUB.
 *
 * Progress is saved after every page — fully resumable.
 */
export async function convertPdfToEpub(args: QwenEpubArgs): Promise<void> {
  const { pdfPath, outputPath, progressPath, client, modelId, verbose } = args;
  const imgDir = imagesDir(progressPath);
  const docStart = Date.now();

  // ── Step 1: count pages ───────────────────────────────────────────────────
  {
    const t0 = Date.now();
    process.stdout.write(`[EPUB] ${fmtClock(t0)}  Step 1 — Counting pages...`);
    const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
    process.stdout.write(` done  ${fmtSec(Date.now() - t0)}  (${numPages} page(s))\n`);
    console.log("");

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
        console.log("         OCR and image extraction skipped — pages will be empty.");
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

          // ── Qwen3 extraction ───────────────────────────────────────────────
          let html = "";
          let inferMs = 0;
          let imgCount = 0;
          let imgMs = 0;

          if (pageImage) {
            // Step A: model produces HTML with <figure data-region="…"> for visuals
            const t = Date.now();
            process.stdout.write(`         ${fmtClock(t)}  Qwen3-VL → HTML + figures...`);
            try {
              const raw = await extractPageRaw(client, modelId, pageNum, pageImage, verbose);
              html = cleanHtmlFragment(raw);
              inferMs = Date.now() - t;
              process.stdout.write(` done  ${fmtSec(inferMs)}  (${html.length} chars)\n`);
            } catch (err) {
              inferMs = Date.now() - t;
              process.stdout.write(` failed  ${fmtSec(inferMs)}\n`);
              console.log(`         Error: ${err instanceof Error ? err.message : String(err)}`);
              html = `<p dir="rtl"><!-- extraction failed for page ${pageNum} --></p>`;
            }

            // Step B: crop declared image regions using sharp
            if (html.includes("data-region=")) {
              const ti = Date.now();
              process.stdout.write(`         ${fmtClock(ti)}  Cropping image regions...`);
              try {
                const result = await processInlineImages(html, pageImage.base64, pageNum, imgDir);
                html = result.html;
                imgCount = result.count;
                imgMs = Date.now() - ti;
                process.stdout.write(` done  ${fmtSec(imgMs)}  (${imgCount} image(s))\n`);
              } catch (err) {
                imgMs = Date.now() - ti;
                process.stdout.write(` failed  ${fmtSec(imgMs)}\n`);
                console.log(`         Error: ${err instanceof Error ? err.message : String(err)}`);
              }
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
          const parts = [
            `render ${fmtSec(renderMs)}`,
            `inference ${fmtSec(inferMs)}`,
            ...(imgCount > 0 ? [`${imgCount} image(s) ${fmtSec(imgMs)}`] : []),
          ];
          console.log(`         Page ${pageNum} done in ${fmtSec(pageMs)}  [${parts.join("  ·  ")}]`);
          if (remaining > 0) {
            console.log(
              `         avg ${fmtSec(avgMs)}/page  ·  ${remaining} page(s) remaining  ·  ETA ${fmtEta(avgMs * remaining)}`,
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

        await generateEpub(orderedPages, outputPath, pdfPath, imgDir, verbose);
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
      console.log(`  Images folder:   ${imgDir}`);
      console.log(`  Progress file:   ${progressPath}`);
      console.log(`  (Delete both to re-convert from scratch.)`);
      console.log("");
      console.log(`  Pages:           ${processed} extracted of ${numPages}`);
      console.log(`  Avg per page:    ${avgPerPage}`);
      console.log(`  Total elapsed:   ${fmtSec(totalMs)}`);
      console.log(`${"═".repeat(56)}`);
    }
  }
}
