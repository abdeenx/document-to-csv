/**
 * word-generator.ts
 *
 * Converts an array of per-page corroborated text strings into a Word (.docx)
 * document, preserving document structure and supporting both Arabic (RTL) and
 * Latin (LTR) text.
 *
 * Structure detection heuristics:
 *   - Headings   : short standalone line (< 80 chars) surrounded by blank lines,
 *                  or the first non-empty line of a page.
 *   - Tables     : consecutive lines that contain at least one tab character.
 *                  Columns are tab-separated; the first row is treated as the header.
 *   - Paragraphs : everything else.
 *
 * Arabic RTL support:
 *   Paragraphs/cells that contain Arabic Unicode characters are automatically
 *   marked as bidirectional with right alignment so Word renders them correctly.
 */

import { writeFile, mkdir } from "node:fs/promises";
import { dirname } from "node:path";
import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

// ---------------------------------------------------------------------------
// Arabic detection
// Covers: Arabic, Arabic Supplement, Arabic Extended-A/B, Arabic Presentation
// Forms-A and -B, and Arabic Mathematical Alphabetic Symbols.
// ---------------------------------------------------------------------------

const ARABIC_RE =
  /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;

function containsArabic(text: string): boolean {
  return ARABIC_RE.test(text);
}

// ---------------------------------------------------------------------------
// Element builders
// ---------------------------------------------------------------------------

function makeTextRun(text: string, bold = false): TextRun {
  const rtl = containsArabic(text);
  return new TextRun({ text, bold, rightToLeft: rtl });
}

function makeParagraph(
  text: string,
  options: {
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel];
    pageBreakBefore?: boolean;
  } = {},
): Paragraph {
  const rtl = containsArabic(text);
  return new Paragraph({
    heading: options.heading,
    bidirectional: rtl,
    alignment: rtl ? AlignmentType.RIGHT : undefined,
    pageBreakBefore: options.pageBreakBefore ?? false,
    children: [makeTextRun(text, !!options.heading)],
  });
}

function buildTable(lines: string[]): Table {
  const rows = lines.map((line, rowIdx) => {
    const cells = line.split("\t");
    return new TableRow({
      children: cells.map((cell) => {
        const cellText = cell.trim();
        const rtl = containsArabic(cellText);
        return new TableCell({
          children: [
            new Paragraph({
              bidirectional: rtl,
              alignment: rtl ? AlignmentType.RIGHT : undefined,
              children: [
                new TextRun({
                  text: cellText,
                  bold: rowIdx === 0,
                  rightToLeft: rtl,
                }),
              ],
            }),
          ],
        });
      }),
    });
  });

  return new Table({
    rows,
    width: { size: 9638, type: WidthType.DXA },
  });
}

// ---------------------------------------------------------------------------
// Heading heuristics
// ---------------------------------------------------------------------------

function isLikelyHeading(
  line: string,
  isFirstNonEmpty: boolean,
  prevLine: string,
  nextLine: string,
): boolean {
  if (!line.trim() || line.length > 100) return false;
  if (line.includes("\t")) return false; // table rows are never headings

  // First non-empty line of a page → heading
  if (isFirstNonEmpty) return true;

  // Short line surrounded by blank lines
  const prevBlank = prevLine.trim() === "";
  const nextBlank = nextLine.trim() === "";
  if (prevBlank && nextBlank && line.length < 80) return true;

  // All-caps short line (common for section headers)
  if (line.length < 60 && line === line.toUpperCase() && /[A-Z]/.test(line)) {
    return true;
  }

  return false;
}

// ---------------------------------------------------------------------------
// Per-page element builder
// ---------------------------------------------------------------------------

type DocElement = Paragraph | Table;

function buildPageElements(text: string): DocElement[] {
  const elements: DocElement[] = [];
  const lines = text.split("\n");
  let i = 0;
  let firstNonEmpty = true;

  while (i < lines.length) {
    const line = lines[i] ?? "";

    // Skip empty lines
    if (line.trim() === "") {
      i++;
      continue;
    }

    // Table block: collect consecutive lines that contain tabs
    if (line.includes("\t")) {
      const tableLines: string[] = [];
      while (i < lines.length && (lines[i] ?? "").includes("\t")) {
        tableLines.push(lines[i]!);
        i++;
      }
      elements.push(buildTable(tableLines));
      firstNonEmpty = false;
      continue;
    }

    const prevLine = i > 0 ? (lines[i - 1] ?? "") : "";
    const nextLine = i < lines.length - 1 ? (lines[i + 1] ?? "") : "";

    if (isLikelyHeading(line, firstNonEmpty, prevLine, nextLine)) {
      elements.push(makeParagraph(line, { heading: HeadingLevel.HEADING_2 }));
    } else {
      elements.push(makeParagraph(line));
    }

    firstNonEmpty = false;
    i++;
  }

  return elements;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Generate a Word (.docx) document from an ordered array of per-page texts.
 *
 * Each page starts on a new page in the output document (except page 1).
 * Arabic paragraphs are automatically rendered RTL; Latin paragraphs LTR.
 */
export async function generateWordDoc(
  pages: Array<{ pageNum: number; text: string }>,
  outputPath: string,
  verbose: boolean,
): Promise<void> {
  if (verbose) {
    console.log(`[Word] Building document from ${pages.length} page(s)...`);
  }

  const allElements: DocElement[] = [];

  for (const { pageNum, text } of pages) {
    if (verbose) {
      console.log(`[Word] Building page ${pageNum}/${pages.length}...`);
    }

    const pageElements = buildPageElements(text);

    if (pageNum > 1) {
      // Insert a page-break paragraph before this page's content.
      // If the page has elements, mark the first one with pageBreakBefore.
      // Otherwise insert a standalone page-break paragraph.
      if (pageElements.length > 0 && pageElements[0] instanceof Paragraph) {
        // Re-create the first element with pageBreakBefore: true
        const first = pageElements[0] as Paragraph;
        // Extract text and RTL from the first paragraph's first run
        // Simpler: just prepend a page-break-only paragraph
        allElements.push(
          new Paragraph({
            pageBreakBefore: true,
            children: [new TextRun({ text: "" })],
          }),
        );
        allElements.push(...pageElements);
      } else {
        // First element is a table or page is empty
        allElements.push(
          new Paragraph({
            pageBreakBefore: true,
            children: [new TextRun({ text: "" })],
          }),
        );
        allElements.push(...pageElements);
      }
    } else {
      allElements.push(...pageElements);
    }

    // Ensure there is at least one element per page (empty pages still need a node)
    if (pageElements.length === 0) {
      allElements.push(makeParagraph(""));
    }
  }

  // Fallback: document must have at least one paragraph
  if (allElements.length === 0) {
    allElements.push(makeParagraph(""));
  }

  const doc = new Document({
    sections: [
      {
        children: allElements,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  await mkdir(dirname(outputPath), { recursive: true });
  await writeFile(outputPath, buffer);

  if (verbose) {
    console.log(
      `[Word] Written: ${outputPath} (${Math.round(buffer.length / 1024)} KB)`,
    );
  }
}
