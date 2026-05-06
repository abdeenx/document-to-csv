/**
 * PDF text extraction + optional OCR pass.
 *
 * Two-source strategy for maximum accuracy:
 *  1. pdfjs text layer  — structural layout (tabs = column boundaries)
 *  2. DeepSeek-OCR pass — visual extraction from rendered page images
 *
 * Both text sources AND the rendered page images are forwarded to Gemma4 so it
 * can reconcile all three signals: structural text, visual text, and the raw image.
 *
 * The OCR pass requires a system PDF renderer. Install one via:
 *   macOS:  brew install poppler        → provides pdftoppm
 *   macOS:  brew install mupdf-tools    → provides mutool
 *   Linux:  apt install poppler-utils   → provides pdftoppm
 *
 * If no renderer is found the CLI falls back to text-layer-only mode and prints
 * an install hint.
 */

import { readFile, readdir, mkdir, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import { execFile as execFileCb } from "node:child_process";
import { promisify } from "node:util";
import { createRequire } from "node:module";
import sharp from "sharp";
import type * as PdfJsLib from "pdfjs-dist";
import type { TextItem, TextMarkedContent } from "pdfjs-dist/types/src/display/api.js";
import type OpenAI from "openai";
import type { OcrResult, PdfPageLayout, PdfTextItem, PdfImageRegion } from "./types.js";
import { callOcrModel } from "./ocr.js";

const execFile = promisify(execFileCb);

const _require = createRequire(import.meta.url);
// Load the legacy CJS-compatible build. The main ESM build requires browser
// globals (DOMMatrix etc.) that don't exist in Node.js. Node.js 22+ allows
// require() of ESM .mjs modules that contain no top-level await.
const pdfjsLib = _require(
  "pdfjs-dist/legacy/build/pdf.mjs",
) as typeof PdfJsLib;

// pdfjs-dist v5 removed the in-process "fake worker" — workerSrc must point
// to the actual worker file. Node.js spawns it via worker_threads.
pdfjsLib.GlobalWorkerOptions.workerSrc = _require.resolve(
  "pdfjs-dist/legacy/build/pdf.worker.mjs",
);

const MAX_DIMENSION = 1600;
const RENDER_TIMEOUT_MS = 60_000;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function isTextItem(item: TextItem | TextMarkedContent): item is TextItem {
  return "str" in item && typeof item.str === "string";
}

type Renderer = "pdftoppm" | "mutool";

async function detectRenderer(): Promise<Renderer | null> {
  for (const tool of ["pdftoppm", "mutool"] as const) {
    try {
      await execFile("which", [tool], { timeout: 2_000 });
      return tool;
    } catch {
      // not available, try next
    }
  }
  return null;
}

async function resizeForOcr(
  buffer: Buffer,
  inputMime: string,
): Promise<{ base64: string; mimeType: "image/jpeg" }> {
  const meta = await sharp(buffer).metadata();
  const w = meta.width ?? 0;
  const h = meta.height ?? 0;
  if (w <= MAX_DIMENSION && h <= MAX_DIMENSION && inputMime === "image/jpeg") {
    return { base64: buffer.toString("base64"), mimeType: "image/jpeg" };
  }
  const out = await sharp(buffer)
    .resize(MAX_DIMENSION, MAX_DIMENSION, { fit: "inside", withoutEnlargement: true })
    .jpeg({ quality: 92 })
    .toBuffer();
  return { base64: out.toString("base64"), mimeType: "image/jpeg" };
}

/**
 * Render a single PDF page to a JPEG buffer using a system renderer.
 * Uses a unique prefix per page so there is no ambiguity with padded filenames.
 */
async function renderPageToJpeg(
  pdfPath: string,
  pageNum: number,
  renderer: Renderer,
  tmpDir: string,
  verbose: boolean,
): Promise<{ base64: string; mimeType: "image/jpeg" }> {
  if (renderer === "pdftoppm") {
    const prefix = join(tmpDir, `pg${pageNum}`);
    await execFile(
      "pdftoppm",
      ["-r", "150", "-jpeg", "-jpegopt", "quality=92", "-f", String(pageNum), "-l", String(pageNum), pdfPath, prefix],
      { timeout: RENDER_TIMEOUT_MS },
    );
    const files = (await readdir(tmpDir)).filter(
      (f) => f.startsWith(`pg${pageNum}`) && f.endsWith(".jpg"),
    );
    if (files.length === 0) {
      throw new Error(`pdftoppm produced no output for page ${pageNum}`);
    }
    const imgBuf = await readFile(join(tmpDir, files[0]!));
    await rm(join(tmpDir, files[0]!));
    if (verbose) {
      console.log(`[PDF] Page ${pageNum} rendered via pdftoppm (${Math.round(imgBuf.length / 1024)} KB)`);
    }
    return resizeForOcr(imgBuf, "image/jpeg");
  } else {
    // mutool convert writes PNG
    const outPath = join(tmpDir, `pg${pageNum}.png`);
    await execFile(
      "mutool",
      ["convert", "-F", "png", "-O", "resolution=150", "-o", outPath, pdfPath, String(pageNum)],
      { timeout: RENDER_TIMEOUT_MS },
    );
    const imgBuf = await readFile(outPath);
    await rm(outPath);
    if (verbose) {
      console.log(`[PDF] Page ${pageNum} rendered via mutool (${Math.round(imgBuf.length / 1024)} KB)`);
    }
    return resizeForOcr(imgBuf, "image/png");
  }
}

// ---------------------------------------------------------------------------
// PDF operator list constants for image detection
// pdfjs-dist may export OPS — we prefer that but fall back to known values.
// ---------------------------------------------------------------------------

const _OPS = (pdfjsLib as unknown as Record<string, unknown>)["OPS"] as
  | Record<string, number>
  | undefined;

const PDF_OPS_SAVE               = _OPS?.["save"]                ?? 13;
const PDF_OPS_RESTORE            = _OPS?.["restore"]             ?? 14;
const PDF_OPS_TRANSFORM          = _OPS?.["transform"]           ?? 109;
const PDF_OPS_PAINT_IMAGE_XOBJ   = _OPS?.["paintImageXObject"]   ?? 85;
const PDF_OPS_PAINT_JPEG_XOBJ    = _OPS?.["paintJpegXObject"]    ?? 82;
const PDF_OPS_PAINT_IMGMASK_XOBJ = _OPS?.["paintImageMaskXObject"] ?? 88;

/**
 * Multiply two PDF 6-element matrices [a,b,c,d,e,f].
 * result = m1 × m2  (m1 is already current; m2 is the new transform).
 */
function mulMat6(m1: number[], m2: number[]): number[] {
  const [a, b, c, d, e, f] = m1 as [number, number, number, number, number, number];
  const [ap, bp, cp, dp, ep, fp] = m2 as [number, number, number, number, number, number];
  return [
    a * ap + c * bp,
    b * ap + d * bp,
    a * cp + c * dp,
    b * cp + d * dp,
    a * ep + c * fp + e,
    b * ep + d * fp + f,
  ];
}

// ---------------------------------------------------------------------------
// pdfjs text layer extraction
// ---------------------------------------------------------------------------

type PdfDoc = Awaited<ReturnType<typeof pdfjsLib.getDocument>["promise"]>;

async function extractPdfjsText(
  doc: PdfDoc,
  numPages: number,
  verbose: boolean,
): Promise<{ pageTexts: string[]; hasAnyText: boolean }> {
  const pageTexts: string[] = [];
  let hasAnyText = false;

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const textContent = await page.getTextContent();

    const items = textContent.items
      .filter(isTextItem)
      .filter((item) => item.str.trim().length > 0)
      .map((item) => ({
        text: item.str,
        // transform = [scaleX, skewY, skewX, scaleY, translateX, translateY]
        x: item.transform[4] ?? 0,
        y: item.transform[5] ?? 0,
        size: Math.abs(item.transform[3] ?? 12),
      }));

    if (items.length > 0) hasAnyText = true;

    // PDF Y origin is bottom-left → sort descending (top of page first),
    // then ascending by X so items read left-to-right within each row.
    items.sort((a, b) => b.y - a.y || a.x - b.x);

    const rows: { text: string; x: number }[][] = [];
    let currentRow: { text: string; x: number }[] = [];
    let currentY: number | null = null;

    for (const item of items) {
      const threshold = Math.max(item.size * 0.5, 2);
      if (currentY === null || Math.abs(item.y - currentY) <= threshold) {
        currentRow.push({ text: item.text, x: item.x });
        if (currentY === null) currentY = item.y;
      } else {
        if (currentRow.length > 0) {
          currentRow.sort((a, b) => a.x - b.x);
          rows.push(currentRow);
        }
        currentRow = [{ text: item.text, x: item.x }];
        currentY = item.y;
      }
    }
    if (currentRow.length > 0) {
      currentRow.sort((a, b) => a.x - b.x);
      rows.push(currentRow);
    }

    const pageText = rows.map((row) => row.map((c) => c.text).join("\t")).join("\n");
    pageTexts.push(numPages > 1 ? `=== PAGE ${pageNum} ===\n${pageText}` : pageText);

    if (verbose) {
      console.log(`[PDF] Text layer page ${pageNum}: ${rows.length} rows, ${pageText.length} chars`);
    }
  }

  return { pageTexts, hasAnyText };
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Extract text from a PDF file and optionally run DeepSeek-OCR on each
 * rendered page for a two-source, visually verified result.
 *
 * @param pdfPath    Absolute path to the PDF file.
 * @param verbose    Enable step-by-step logging.
 * @param ocrClient  OpenAI-compatible client for the OCR model (optional).
 * @param ocrModelId Model ID for the OCR model (optional).
 */
export async function extractTextFromPdf(
  pdfPath: string,
  verbose: boolean,
  ocrClient?: OpenAI,
  ocrModelId?: string,
): Promise<OcrResult> {
  if (verbose) console.log(`[PDF] Loading: ${pdfPath}`);

  const buffer = await readFile(pdfPath);
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
  const numPages = doc.numPages;

  if (verbose) console.log(`[PDF] ${numPages} page(s) found`);

  // ── Step 1: pdfjs text layer ──────────────────────────────────────────────
  const { pageTexts, hasAnyText } = await extractPdfjsText(doc, numPages, verbose);

  if (!hasAnyText) {
    throw new Error(
      "This PDF has no embedded text (it appears to be scanned). " +
        "Export the pages as PNG or JPEG and pass each image to the CLI instead.",
    );
  }

  const pdfjsRawText = pageTexts.join("\n\n");

  // ── Step 2: OCR pass (optional) ───────────────────────────────────────────
  if (ocrClient && ocrModelId) {
    const renderer = await detectRenderer();

    if (!renderer) {
      console.log(
        "[PDF] No PDF renderer found — OCR pass skipped.\n" +
          "      Install one to enable visual OCR on rendered pages:\n" +
          "        macOS:  brew install poppler      (provides pdftoppm)\n" +
          "        macOS:  brew install mupdf-tools  (provides mutool)\n" +
          "        Linux:  apt install poppler-utils\n" +
          "[PDF] Continuing with text-layer extraction only.",
      );
    } else {
      if (verbose) {
        console.log(
          `[PDF] Renderer: ${renderer}. Running OCR pass on ${numPages} page(s) with model: ${ocrModelId}`,
        );
      }

      const tmpDir = join(tmpdir(), `doc2csv-${randomUUID()}`);
      await mkdir(tmpDir, { recursive: true });

      try {
        const pageImages: Array<{ base64: string; mimeType: string }> = [];
        const ocrPageTexts: string[] = [];

        for (let pageNum = 1; pageNum <= numPages; pageNum++) {
          const img = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
          pageImages.push(img);

          if (verbose) {
            console.log(`[PDF] Sending page ${pageNum}/${numPages} to DeepSeek-OCR...`);
          }
          const ocrText = await callOcrModel(
            ocrClient,
            img.base64,
            img.mimeType,
            ocrModelId,
            verbose,
          );
          ocrPageTexts.push(numPages > 1 ? `=== PAGE ${pageNum} ===\n${ocrText}` : ocrText);
        }

        // Combine both text sources. Gemma4 reconciles them with the page images.
        const combinedText = [
          "=== PDF TEXT LAYER (structural — tabs mark column boundaries) ===",
          pdfjsRawText,
          "",
          "=== OCR TEXT LAYER (visual — use to verify values and fill blanks) ===",
          ocrPageTexts.join("\n\n"),
        ].join("\n");

        if (verbose) {
          console.log(
            `[PDF] Both layers combined: ${combinedText.length} chars across ${numPages} page(s). ` +
              `Forwarding ${pageImages.length} page image(s) to Gemma4.`,
          );
        }

        return {
          rawText: combinedText,
          model: `pdfjs + ${ocrModelId}`,
          pageImages,
        };
      } finally {
        await rm(tmpDir, { recursive: true, force: true });
      }
    }
  }

  // ── Fallback: text-layer only ─────────────────────────────────────────────
  if (verbose) {
    console.log(`[PDF] Text-layer only: ${pdfjsRawText.length} chars`);
  }

  return {
    rawText: pdfjsRawText,
    model: "pdfjs-text-extract",
  };
}

// ---------------------------------------------------------------------------
// Layout extraction for Excel output
// ---------------------------------------------------------------------------

/**
 * Extract per-page layout data (text item positions + image bounding boxes)
 * from a PDF. Called only when `--excel` is used so there is no overhead in
 * the normal CSV path.
 *
 * Image bounding boxes are detected by replaying the PDF operator list and
 * tracking the Current Transformation Matrix (CTM) through save/restore/
 * transform ops.  When a paint-image op is encountered the CTM encodes the
 * image's position and size in PDF user-space.
 */
export async function extractPdfPageLayouts(
  pdfPath: string,
  verbose: boolean,
): Promise<PdfPageLayout[]> {
  if (verbose) console.log(`[PDF-Layout] Loading: ${pdfPath}`);

  const buffer = await readFile(pdfPath);
  const doc    = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;

  const layouts: PdfPageLayout[] = [];

  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page     = await doc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });
    const pageWidth  = viewport.width;
    const pageHeight = viewport.height;

    // ── Text items ──────────────────────────────────────────────────────────
    const textContent = await page.getTextContent();
    const textItems: PdfTextItem[] = [];

    for (const raw of textContent.items) {
      if (!isTextItem(raw) || !raw.str.trim()) continue;
      const t = raw.transform; // [a, b, c, d, e, f]
      // Font size ≈ Euclidean norm of the first column of the matrix
      const fontSize = Math.sqrt((t[0] ?? 0) ** 2 + (t[1] ?? 0) ** 2);
      textItems.push({
        x:        t[4] ?? 0,
        y:        t[5] ?? 0,       // PDF baseline from bottom
        str:      raw.str,
        fontSize: fontSize || 10,
        width:    raw.width ?? 0,
      });
    }

    // ── Image regions via operator list ─────────────────────────────────────
    // We replay every drawing op, tracking the CTM via save/restore/transform.
    // When a paint-image op fires the CTM encodes position + size.
    const opList = await page.getOperatorList();
    const imageRegions: PdfImageRegion[] = [];

    let   ctm: number[]   = [1, 0, 0, 1, 0, 0];  // identity
    const ctmStack: number[][] = [];

    for (let i = 0; i < opList.fnArray.length; i++) {
      const fn   = opList.fnArray[i]!;
      const args = opList.argsArray[i] as number[] | undefined;

      if (fn === PDF_OPS_SAVE) {
        ctmStack.push([...ctm]);
      } else if (fn === PDF_OPS_RESTORE) {
        if (ctmStack.length > 0) ctm = ctmStack.pop()!;
      } else if (fn === PDF_OPS_TRANSFORM && args && args.length >= 6) {
        ctm = mulMat6(ctm, args);
      } else if (
        fn === PDF_OPS_PAINT_IMAGE_XOBJ ||
        fn === PDF_OPS_PAINT_JPEG_XOBJ  ||
        fn === PDF_OPS_PAINT_IMGMASK_XOBJ
      ) {
        // CTM [a,b,c,d,e,f]: image unit square → parallelogram
        // For axis-aligned images (no rotation): a=width, d=height, e=x, f=y (bottom-left).
        // Guard against degenerate or tiny transforms (artefacts / hairlines).
        const imgLeft  = Math.min(ctm[4]!, ctm[4]! + ctm[0]!);
        const imgRight = Math.max(ctm[4]!, ctm[4]! + ctm[0]!);
        const imgBot   = Math.min(ctm[5]!, ctm[5]! + ctm[3]!);
        const imgTop   = Math.max(ctm[5]!, ctm[5]! + ctm[3]!);

        const w = imgRight - imgLeft;
        const h = imgTop   - imgBot;

        // Skip tiny artefacts (< 5 pt) and regions outside the page
        if (w < 5 || h < 5 || imgLeft < 0 || imgBot < 0) continue;
        if (imgRight > pageWidth * 1.05 || imgTop > pageHeight * 1.05) continue;

        imageRegions.push({ x: imgLeft, y: imgBot, width: w, height: h });

        if (verbose) {
          console.log(
            `[PDF-Layout] Page ${pageNum}: image @ x=${imgLeft.toFixed(0)} ` +
            `y=${imgBot.toFixed(0)} w=${w.toFixed(0)} h=${h.toFixed(0)}`,
          );
        }
      }
    }

    page.cleanup();

    if (verbose) {
      console.log(
        `[PDF-Layout] Page ${pageNum}: ${textItems.length} text item(s), ` +
        `${imageRegions.length} image region(s)`,
      );
    }

    layouts.push({ pageWidth, pageHeight, textItems, imageRegions });
  }

  return layouts;
}
