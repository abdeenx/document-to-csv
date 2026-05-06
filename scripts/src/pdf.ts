/**
 * PDF text extraction using pdfjs-dist (legacy build, no DOM required).
 * Groups text items by Y-position to reconstruct table rows, then returns
 * an OcrResult-compatible object so the same Gemma4 CSV pipeline applies.
 *
 * Note: works for text-based PDFs only. Scanned PDFs (image-only) have no
 * embedded text layer — users should export pages as images and use the
 * image path instead.
 */

import { readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import type * as PdfJsLib from "pdfjs-dist";
import type { TextItem, TextMarkedContent } from "pdfjs-dist/types/src/display/api.js";
import type { OcrResult } from "./types.js";

const _require = createRequire(import.meta.url);
// Load the legacy CJS-compatible build. The main ESM build requires browser
// globals (DOMMatrix etc.) that don't exist in Node.js. Node.js 22+ supports
// require() of ESM .mjs modules that contain no top-level await.
const pdfjsLib = _require(
  "pdfjs-dist/legacy/build/pdf.mjs",
) as typeof PdfJsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = "";

function isTextItem(item: TextItem | TextMarkedContent): item is TextItem {
  return "str" in item && typeof item.str === "string";
}

export async function extractTextFromPdf(
  pdfPath: string,
  verbose: boolean,
): Promise<OcrResult> {
  if (verbose) {
    console.log(`[PDF] Loading: ${pdfPath}`);
  }

  const buffer = await readFile(pdfPath);
  const doc = await pdfjsLib
    .getDocument({ data: new Uint8Array(buffer) })
    .promise;
  const numPages = doc.numPages;

  if (verbose) {
    console.log(`[PDF] ${numPages} page(s) found`);
  }

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
        // Font height from the scaleY component
        size: Math.abs(item.transform[3] ?? 12),
      }));

    if (items.length > 0) hasAnyText = true;

    // PDF Y origin is bottom-left → sort descending (top of page first),
    // then ascending by X so items read left-to-right within each row.
    items.sort((a, b) => b.y - a.y || a.x - b.x);

    // Group items into logical rows. Items whose Y differs by less than half
    // the font height (min 2pt) are considered on the same line.
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

    // Join items in a row with a tab so the structuring model sees clear
    // column boundaries, then join rows with newlines.
    const pageText = rows.map((row) => row.map((c) => c.text).join("\t")).join("\n");

    pageTexts.push(numPages > 1 ? `=== PAGE ${pageNum} ===\n${pageText}` : pageText);

    if (verbose) {
      console.log(
        `[PDF] Page ${pageNum}: ${rows.length} rows, ${pageText.length} chars`,
      );
    }
  }

  if (!hasAnyText) {
    throw new Error(
      "This PDF has no embedded text (it appears to be scanned). " +
        "Export the pages as PNG or JPEG and pass each image to the CLI instead.",
    );
  }

  const rawText = pageTexts.join("\n\n");

  if (verbose) {
    console.log(
      `[PDF] Extracted ${rawText.length} chars across ${numPages} page(s)`,
    );
  }

  return {
    rawText,
    model: "pdfjs-text-extract",
    // No image for text PDFs — Gemma4 falls back to text-only mode automatically
    // (imageBase64 is absent, so no vision content block is added to the message)
  };
}
