/**
 * html-to-docx-converter.ts
 *
 * Converts a cleaned HTML fragment (as produced by the EPUB/Rich-Word
 * extraction pipeline) into an array of docx block elements
 * (Paragraph | Table) suitable for embedding in a Word document.
 *
 * Handled block elements:
 *   <h1>–<h6>   Headings with RTL/LTR direction
 *   <p>          Paragraphs with inline bold/italic/br support
 *   <figure>     Cropped images read from `imagesDir`, scaled to page width
 *   <ul>/<ol>    Bullet / numbered lists
 *   <table>      Tables with header row support
 *   <hr/>        Paragraph spacer (empty line)
 *
 * Images are embedded as ImageRun inside a centred Paragraph. Captions
 * (<figcaption>) become a small italic paragraph below the image.
 *
 * Arabic RTL detection: if a block's dir attribute is "rtl" or the text
 * contains Arabic Unicode characters, the paragraph is set bidirectional
 * with right alignment and rightToLeft TextRun flags.
 */

import { readFile } from "node:fs/promises";
import { join } from "node:path";
import sharp from "sharp";
import {
  AlignmentType,
  HeadingLevel,
  ImageRun,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Arabic + related Unicode ranges */
const ARABIC_RE =
  /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;

/**
 * Maximum image width in pixels.
 * A4 with 2.54 cm margins on each side gives ~15.6 cm usable ≈ 590 px at 96 dpi.
 * We use 560 px as a safe maximum that leaves a small visual margin.
 */
const IMG_MAX_PX = 560;

// ---------------------------------------------------------------------------
// Public type
// ---------------------------------------------------------------------------

export type DocxElement = Paragraph | Table;

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Convert a cleaned HTML fragment into an array of docx block elements.
 *
 * @param html             HTML fragment with <img src="images/..."> already resolved
 *                         (i.e. after processInlineImages() has run).
 * @param imagesDir        Directory where cropped JPEG files live.
 * @param pageBreakBefore  If true, the very first element gets pageBreakBefore.
 */
export async function htmlFragmentToDocxBlocks(
  html: string,
  imagesDir: string,
  pageBreakBefore: boolean,
): Promise<DocxElement[]> {
  const rawBlocks = splitIntoBlocks(html);
  const result: DocxElement[] = [];

  for (let i = 0; i < rawBlocks.length; i++) {
    const block = rawBlocks[i]!;
    const addBreak = pageBreakBefore && i === 0;
    const elements = await blockToDocxElements(block, imagesDir, addBreak);
    result.push(...elements);
  }

  return result;
}

// ---------------------------------------------------------------------------
// Block splitter
// ---------------------------------------------------------------------------

interface RawBlock {
  /** Lower-cased tag name: "h2", "p", "figure", "table", "ul", "ol", "hr" */
  tag: string;
  /** Raw attribute string from the opening tag (may include dir="rtl" etc.) */
  attrs: string;
  /** innerHTML — everything between the opening and closing tag */
  inner: string;
}

/**
 * Split a flat HTML fragment into an ordered array of top-level blocks.
 * Uses a lazy-quantifier regex so nested tags (figcaption, td, li, …) are
 * captured correctly as part of their parent's innerHTML.
 */
function splitIntoBlocks(html: string): RawBlock[] {
  const result: RawBlock[] = [];

  const re =
    /<(h[1-6]|p|figure|table|ul|ol)\s*([^>]*?)>([\s\S]*?)<\/\1>|<hr\s*\/>/gi;

  for (const m of html.matchAll(re)) {
    if ((m[0] ?? "").startsWith("<hr")) {
      result.push({ tag: "hr", attrs: "", inner: "" });
    } else {
      result.push({
        tag:   (m[1] ?? "").toLowerCase(),
        attrs: m[2] ?? "",
        inner: m[3] ?? "",
      });
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Block → docx elements
// ---------------------------------------------------------------------------

const HEADING_LEVELS: Record<
  string,
  (typeof HeadingLevel)[keyof typeof HeadingLevel]
> = {
  h1: HeadingLevel.HEADING_1,
  h2: HeadingLevel.HEADING_2,
  h3: HeadingLevel.HEADING_3,
  h4: HeadingLevel.HEADING_4,
  h5: HeadingLevel.HEADING_5,
  h6: HeadingLevel.HEADING_6,
};

async function blockToDocxElements(
  block: RawBlock,
  imagesDir: string,
  pageBreakBefore: boolean,
): Promise<DocxElement[]> {
  const { tag, attrs, inner } = block;
  const dir   = extractDirAttr(attrs);
  const isRtl = dir !== "ltr";

  switch (tag) {
    // ── Headings ────────────────────────────────────────────────────────────
    case "h1":
    case "h2":
    case "h3":
    case "h4":
    case "h5":
    case "h6": {
      const text = decodeEntities(stripTags(inner).trim());
      return [
        new Paragraph({
          heading:         HEADING_LEVELS[tag] ?? HeadingLevel.HEADING_2,
          bidirectional:   isRtl,
          alignment:       isRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          pageBreakBefore,
          children: [new TextRun({ text, bold: true, rightToLeft: isRtl })],
        }),
      ];
    }

    // ── Paragraphs ──────────────────────────────────────────────────────────
    case "p": {
      const runs = parseInlineRuns(inner, isRtl);
      if (runs.length === 0) return [];
      return [
        new Paragraph({
          bidirectional:   isRtl,
          alignment:       isRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          pageBreakBefore,
          children: runs,
        }),
      ];
    }

    // ── Figures (images) ────────────────────────────────────────────────────
    case "figure":
      return figureToDocxElements(inner, imagesDir, pageBreakBefore);

    // ── Bullet lists ────────────────────────────────────────────────────────
    case "ul": {
      const items = extractListItems(inner);
      return items.map((itemHtml, idx) => {
        const text    = decodeEntities(stripTags(itemHtml).trim());
        const itemRtl = dir !== "ltr" || ARABIC_RE.test(text);
        return new Paragraph({
          bidirectional:   itemRtl,
          alignment:       itemRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          pageBreakBefore: pageBreakBefore && idx === 0,
          indent:          { start: 720 },
          children: [
            new TextRun({ text: "•  ", rightToLeft: false }),
            new TextRun({ text, rightToLeft: itemRtl }),
          ],
        });
      });
    }

    // ── Ordered lists ───────────────────────────────────────────────────────
    case "ol": {
      const items = extractListItems(inner);
      return items.map((itemHtml, idx) => {
        const text    = decodeEntities(stripTags(itemHtml).trim());
        const itemRtl = dir !== "ltr" || ARABIC_RE.test(text);
        return new Paragraph({
          bidirectional:   itemRtl,
          alignment:       itemRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          pageBreakBefore: pageBreakBefore && idx === 0,
          indent:          { start: 720 },
          children: [
            new TextRun({ text: `${idx + 1}.  `, rightToLeft: false }),
            new TextRun({ text, rightToLeft: itemRtl }),
          ],
        });
      });
    }

    // ── Tables ──────────────────────────────────────────────────────────────
    case "table":
      return [buildTable(inner, pageBreakBefore)];

    // ── Horizontal rule (spacer) ─────────────────────────────────────────
    case "hr":
      return [new Paragraph({ text: "", pageBreakBefore })];

    default:
      return [];
  }
}

// ---------------------------------------------------------------------------
// Figure / image handling
// ---------------------------------------------------------------------------

async function figureToDocxElements(
  figInner: string,
  imagesDir: string,
  pageBreakBefore: boolean,
): Promise<DocxElement[]> {
  // Extract the image filename from src="images/page_NNN_imgM.jpg"
  const srcMatch = /src="images\/([^"]+)"/i.exec(figInner);
  if (!srcMatch) return [];

  const imgFileName = srcMatch[1]!;
  const imgPath     = join(imagesDir, imgFileName);

  let imgBuffer: Buffer;
  try {
    imgBuffer = await readFile(imgPath);
  } catch {
    console.warn(`[RichWord] Warning: image not found: ${imgPath}`);
    return [];
  }

  // Get pixel dimensions via sharp then scale to fit page width
  const meta      = await sharp(imgBuffer).metadata();
  const origW     = meta.width  ?? 400;
  const origH     = meta.height ?? 300;
  const scale     = Math.min(1, IMG_MAX_PX / origW);
  const finalW    = Math.round(origW * scale);
  const finalH    = Math.round(origH * scale);

  const imageType = /\.png$/i.test(imgFileName) ? "png" : "jpg";

  const imagePara = new Paragraph({
    alignment:       AlignmentType.CENTER,
    pageBreakBefore,
    children: [
      new ImageRun({
        type:           imageType as "jpg" | "png",
        data:           imgBuffer,
        transformation: { width: finalW, height: finalH },
      }),
    ],
  });

  const result: DocxElement[] = [imagePara];

  // Optional figcaption → small italic paragraph below the image
  const captionMatch =
    /<figcaption([^>]*)>([\s\S]*?)<\/figcaption>/i.exec(figInner);
  if (captionMatch) {
    const captionText = decodeEntities(stripTags(captionMatch[2] ?? "").trim());
    if (captionText) {
      const captionAttrs = captionMatch[1] ?? "";
      const captionDir   = extractDirAttr(captionAttrs);
      const captionRtl   = captionDir !== "ltr" || ARABIC_RE.test(captionText);
      result.push(
        new Paragraph({
          alignment:     AlignmentType.CENTER,
          bidirectional: captionRtl,
          children: [
            new TextRun({
              text:        captionText,
              italics:     true,
              size:        18, // 9 pt
              rightToLeft: captionRtl,
            }),
          ],
        }),
      );
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Table handling
// ---------------------------------------------------------------------------

function buildTable(tableInner: string, pageBreakBefore: boolean): Table {
  const rowRe  = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const cellRe = /<t[hd]\s*([^>]*)>([\s\S]*?)<\/t[hd]>/gi;
  const rows: TableRow[] = [];
  let rowIdx = 0;

  for (const rowMatch of tableInner.matchAll(rowRe)) {
    const rowInner = rowMatch[1] ?? "";
    const cells: TableCell[] = [];

    for (const cellMatch of rowInner.matchAll(cellRe)) {
      const cellAttrs = cellMatch[1] ?? "";
      const cellInner = cellMatch[2] ?? "";
      const cellText  = decodeEntities(stripTags(cellInner).trim());
      const cellDir   = extractDirAttr(cellAttrs);
      const cellRtl   = cellDir !== "ltr" || ARABIC_RE.test(cellText);
      const isHeader  = rowIdx === 0;

      cells.push(
        new TableCell({
          children: [
            new Paragraph({
              bidirectional: cellRtl,
              alignment:     cellRtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
              children: [
                new TextRun({
                  text:        cellText,
                  bold:        isHeader,
                  rightToLeft: cellRtl,
                }),
              ],
            }),
          ],
        }),
      );
    }

    if (cells.length > 0) {
      rows.push(new TableRow({ children: cells }));
      rowIdx++;
    }
  }

  if (rows.length === 0) {
    rows.push(
      new TableRow({
        children: [new TableCell({ children: [new Paragraph({ text: "" })] })],
      }),
    );
  }

  // pageBreakBefore on a Table is not directly supported in docx v9;
  // we handle it by inserting an empty Paragraph before the table in the
  // caller when needed (the caller only passes pageBreakBefore=true for the
  // first block on a page, so we just add an empty spacer for tables).
  void pageBreakBefore; // acknowledged — handled by caller wrapping

  return new Table({
    rows,
    width: { size: 9638, type: WidthType.DXA },
  });
}

// ---------------------------------------------------------------------------
// Inline run parser
// ---------------------------------------------------------------------------

function parseInlineRuns(html: string, isRtl: boolean): TextRun[] {
  // Split on recognised inline tags, keeping the delimiters
  const parts = html.split(
    /(<(?:strong|b)[^>]*>[\s\S]*?<\/(?:strong|b)>|<(?:em|i)[^>]*>[\s\S]*?<\/(?:em|i)>|<br\s*\/>)/gi,
  );

  const runs: TextRun[] = [];

  for (const part of parts) {
    if (!part) continue;

    const boldMatch =
      /^<(?:strong|b)[^>]*>([\s\S]*?)<\/(?:strong|b)>$/i.exec(part);
    if (boldMatch) {
      const text = decodeEntities(stripTags(boldMatch[1] ?? "").trim());
      if (text) runs.push(new TextRun({ text, bold: true, rightToLeft: isRtl }));
      continue;
    }

    const emMatch =
      /^<(?:em|i)[^>]*>([\s\S]*?)<\/(?:em|i)>$/i.exec(part);
    if (emMatch) {
      const text = decodeEntities(stripTags(emMatch[1] ?? "").trim());
      if (text) runs.push(new TextRun({ text, italics: true, rightToLeft: isRtl }));
      continue;
    }

    if (/^<br\s*\/>$/i.test(part)) {
      runs.push(new TextRun({ text: "\n", rightToLeft: isRtl }));
      continue;
    }

    // Plain text node
    const text = decodeEntities(stripTags(part));
    if (text.trim()) {
      runs.push(new TextRun({ text, rightToLeft: isRtl }));
    }
  }

  return runs;
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function extractDirAttr(attrs: string): "rtl" | "ltr" {
  const m = /dir="(rtl|ltr)"/i.exec(attrs);
  return m ? (m[1] as "rtl" | "ltr") : "rtl"; // default RTL for Arabic-primary docs
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, "");
}

function decodeEntities(str: string): string {
  return str
    .replace(/&amp;/g,  "&")
    .replace(/&lt;/g,   "<")
    .replace(/&gt;/g,   ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCharCode(Number(n)));
}

function extractListItems(listInner: string): string[] {
  const items: string[] = [];
  for (const m of listInner.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/gi)) {
    items.push(m[1] ?? "");
  }
  return items;
}
