/**
 * docx-reader.ts
 *
 * Extracts per-page plain text from a .docx file by reading the internal XML.
 *
 * A .docx file is a ZIP archive. The document body lives in word/document.xml.
 * Text is stored in <w:t> elements.
 *
 * Page boundaries are detected from two marker types:
 *
 *   1. <w:pageBreakBefore/>  (primary)
 *      Written by the `docx` library when `pageBreakBefore: true` is set on a
 *      Paragraph. This is what our own word-generator produces — an empty
 *      paragraph with this flag marks the start of every page beyond the first.
 *
 *   2. <w:br w:type="page"/>  (fallback)
 *      An explicit inline page-break element used by some other tools.
 *
 * For files with no explicit page breaks, the entire document is returned as
 * a single-element array.
 */

import { readFile } from "node:fs/promises";
import JSZip from "jszip";

// ---------------------------------------------------------------------------
// XML entity decoding
// ---------------------------------------------------------------------------

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

// ---------------------------------------------------------------------------
// Text extraction from an XML segment
// ---------------------------------------------------------------------------

/**
 * Extract plain text from a slice of word/document.xml.
 *
 * Paragraphs are separated by newlines; tab characters in <w:tab/> are
 * preserved. XML entities are decoded.
 */
function extractTextFromSegment(xml: string): string {
  const lines: string[] = [];

  // Split at paragraph boundaries
  const paragraphs = xml.split(/<\/w:p>/i);

  for (const para of paragraphs) {
    // Replace <w:tab/> with a literal tab
    const withTabs = para.replace(/<w:tab[^>]*\/?>/gi, "\t");

    // Extract all <w:t> text content
    let paraText = "";
    const tPattern = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gi;
    let m: RegExpExecArray | null;
    while ((m = tPattern.exec(withTabs)) !== null) {
      paraText += decodeXmlEntities(m[1] ?? "");
    }

    if (paraText) {
      lines.push(paraText);
    }
  }

  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

// ---------------------------------------------------------------------------
// Page-split position detection
// ---------------------------------------------------------------------------

/**
 * Walk backwards in `xml` from `startIdx` to find the opening `<w:p` of the
 * paragraph that contains the marker at `startIdx`.
 *
 * We match `<w:p` only when the 5th character is `>` or whitespace, so that
 * `<w:pPr>`, `<w:pStyle>`, etc. are correctly skipped.
 *
 * Returns the index of the `<` character, or -1 if not found.
 */
function findParaStart(xml: string, startIdx: number): number {
  for (let i = startIdx - 1; i >= 0; i--) {
    if (xml[i] === "<") {
      // Must be <w:p> or <w:p ...> — the character after <w:p must be '>' or whitespace.
      const next5 = xml.slice(i, i + 5);
      if (next5.startsWith("<w:p")) {
        const ch5 = next5[4];
        if (ch5 === ">" || ch5 === " " || ch5 === "\t" || ch5 === "\r" || ch5 === "\n") {
          return i;
        }
      }
    }
  }
  return -1;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Extract per-page text from a .docx file.
 *
 * Returns an array where index 0 = page 1 text, index 1 = page 2 text, etc.
 * If no page breaks are detected, the entire document is returned as a
 * single-element array.
 */
export async function extractDocxPages(docxPath: string): Promise<string[]> {
  const buffer = await readFile(docxPath);
  const zip = await JSZip.loadAsync(buffer);

  const xmlFile = zip.file("word/document.xml");
  if (!xmlFile) {
    throw new Error(`word/document.xml not found inside "${docxPath}"`);
  }

  const xml = await xmlFile.async("string");

  // ── Collect split positions ──────────────────────────────────────────────
  const splitPositions = new Set<number>();

  // Strategy 1: <w:pageBreakBefore/> inside <w:pPr>
  // Each occurrence marks the paragraph that *starts* a new page.
  // Walk backwards from the marker to find the enclosing <w:p.
  const pageBreakBeforeRe = /<w:pageBreakBefore[^>]*\/?>/gi;
  let m: RegExpExecArray | null;
  while ((m = pageBreakBeforeRe.exec(xml)) !== null) {
    const pos = findParaStart(xml, m.index);
    if (pos !== -1) {
      splitPositions.add(pos);
    }
  }

  // Strategy 2: explicit <w:br w:type="page"/> — split after the closing </w:p>
  const brPageRe = /<w:br[^>]+w:type=["']page["'][^>]*\/?>/gi;
  while ((m = brPageRe.exec(xml)) !== null) {
    const closePara = xml.indexOf("</w:p>", m.index);
    if (closePara !== -1) {
      splitPositions.add(closePara + "</w:p>".length);
    }
  }

  if (splitPositions.size === 0) {
    // No page breaks found — return the whole document as one page.
    return [extractTextFromSegment(xml)];
  }

  const positions = [...splitPositions].sort((a, b) => a - b);

  // ── Slice into page segments ─────────────────────────────────────────────
  const pages: string[] = [];

  // Page 1: everything before the first split position
  pages.push(extractTextFromSegment(xml.slice(0, positions[0])));

  // Pages 2…N: between consecutive split positions
  for (let i = 0; i < positions.length; i++) {
    const start = positions[i]!;
    const end = i + 1 < positions.length ? positions[i + 1]! : xml.length;
    pages.push(extractTextFromSegment(xml.slice(start, end)));
  }

  return pages;
}
