/**
 * epub-generator.ts
 *
 * Assembles an EPUB 3 file from an ordered array of per-page HTML fragments.
 *
 * Output structure (one XHTML file per PDF page):
 *
 *   output.epub (ZIP)
 *   ├── mimetype                       ← stored uncompressed, must be first
 *   ├── META-INF/
 *   │   └── container.xml
 *   └── EPUB/
 *       ├── content.opf                ← package manifest + spine
 *       ├── nav.xhtml                  ← EPUB 3 navigation document
 *       ├── toc.ncx                    ← EPUB 2 NCX (broad reader compatibility)
 *       ├── style.css                  ← Arabic RTL + Latin LTR stylesheet
 *       ├── page_001.xhtml             ← PDF page 1
 *       ├── page_002.xhtml             ← PDF page 2
 *       └── …
 *
 * Each XHTML file is self-contained and editable in any text editor.
 * A human reviewer can open page_069.xhtml and compare it with PDF page 69
 * side-by-side, then edit the HTML directly and re-pack as an EPUB.
 */

import { readFile, writeFile, mkdir } from "node:fs/promises";
import { dirname, basename, extname, join } from "node:path";
import { randomUUID } from "node:crypto";
import JSZip from "jszip";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

export function pageFileName(pageNum: number): string {
  return `page_${String(pageNum).padStart(3, "0")}.xhtml`;
}

function pageId(pageNum: number): string {
  return `p${String(pageNum).padStart(3, "0")}`;
}

function isoNow(): string {
  return new Date().toISOString().replace(/\.\d+Z$/, "Z");
}

export function deriveTitleFromPath(pdfPath: string): string {
  const base = basename(pdfPath, extname(pdfPath));
  return base.replace(/[-_]/g, " ").trim() || "Untitled";
}

// ---------------------------------------------------------------------------
// Static assets
// ---------------------------------------------------------------------------

const STYLE_CSS = `/* EPUB stylesheet — Arabic RTL primary, Latin LTR override */

body {
  font-family: "Amiri", "Traditional Arabic", "Arabic Typesetting",
               "Arial Unicode MS", serif;
  direction: rtl;
  text-align: right;
  line-height: 1.9;
  margin: 1.5em 2em;
  color: #1a1a1a;
  background: #ffffff;
}

/* Latin / LTR content */
[dir="ltr"] {
  direction: ltr;
  text-align: left;
  font-family: "Georgia", "Times New Roman", serif;
}

h1, h2, h3, h4, h5, h6 {
  margin: 1.2em 0 0.4em;
  line-height: 1.4;
  font-weight: bold;
}

h2 { font-size: 1.4em; }
h3 { font-size: 1.2em; }
h4 { font-size: 1.1em; }

p {
  margin: 0.5em 0;
  text-align: justify;
}

table {
  border-collapse: collapse;
  width: 100%;
  margin: 1em 0;
  font-size: 0.95em;
}

th, td {
  border: 1px solid #bbbbbb;
  padding: 0.3em 0.6em;
  vertical-align: top;
}

th {
  background-color: #f0f0f0;
  font-weight: bold;
}

ul, ol {
  margin: 0.5em 0;
  padding-inline-start: 2em;
}

li {
  margin: 0.2em 0;
}

dl { margin: 0.5em 0; }
dt { font-weight: bold; }
dd {
  margin-inline-start: 2em;
  margin-bottom: 0.3em;
}

hr {
  border: none;
  border-top: 1px solid #dddddd;
  margin: 1em 0;
}

/* Useful utility classes for manual post-editing */
.rtl { direction: rtl; text-align: right; }
.ltr { direction: ltr; text-align: left; }
.center { text-align: center; }
.bold { font-weight: bold; }

/* Full-page images (covers, illustrations) */
.page-image {
  margin: 0;
  padding: 0;
  text-align: center;
}

.page-image img {
  max-width: 100%;
  max-height: 95vh;
  height: auto;
  display: block;
  margin: 0 auto;
}

.page-image figcaption {
  font-size: 0.85em;
  color: #666;
  margin-top: 0.4em;
}
`;

// ---------------------------------------------------------------------------
// XML document builders
// ---------------------------------------------------------------------------

function buildContainer(): string {
  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">`,
    `  <rootfiles>`,
    `    <rootfile full-path="EPUB/content.opf"`,
    `              media-type="application/oebps-package+xml"/>`,
    `  </rootfiles>`,
    `</container>`,
  ].join("\n");
}

function buildOpf(
  pages: Array<{ pageNum: number; imageOnly?: boolean }>,
  title: string,
  uuid: string,
): string {
  // Determine the first image-only page (gets cover-image property)
  const firstImagePage = pages.find((p) => p.imageOnly)?.pageNum ?? -1;

  // Image manifest items
  const imageItems = pages
    .filter((p) => p.imageOnly)
    .map(({ pageNum }) => {
      const imgFile = pageFileName(pageNum).replace(".xhtml", ".jpg");
      const isCover = pageNum === firstImagePage;
      return (
        `    <item id="img_${pageId(pageNum)}" href="images/${imgFile}"` +
        ` media-type="image/jpeg"` +
        (isCover ? ` properties="cover-image"` : ``) +
        `/>`
      );
    });

  const manifest = [
    `    <item id="nav" href="nav.xhtml"`,
    `          media-type="application/xhtml+xml" properties="nav"/>`,
    `    <item id="ncx" href="toc.ncx"`,
    `          media-type="application/x-dtbncx+xml"/>`,
    `    <item id="css" href="style.css" media-type="text/css"/>`,
    ...imageItems,
    ...pages.map(
      ({ pageNum }) =>
        `    <item id="${pageId(pageNum)}" href="${pageFileName(pageNum)}"` +
        ` media-type="application/xhtml+xml"/>`,
    ),
  ].join("\n");

  const spine = pages
    .map(({ pageNum }) => `    <itemref idref="${pageId(pageNum)}"/>`)
    .join("\n");

  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<package xmlns="http://www.idpf.org/2007/opf"`,
    `         xmlns:dc="http://purl.org/dc/elements/1.1/"`,
    `         unique-identifier="book-id"`,
    `         version="3.0"`,
    `         xml:lang="ar">`,
    `  <metadata>`,
    `    <dc:identifier id="book-id">urn:uuid:${uuid}</dc:identifier>`,
    `    <dc:title>${escapeXml(title)}</dc:title>`,
    `    <dc:language>ar</dc:language>`,
    `    <meta property="dcterms:modified">${isoNow()}</meta>`,
    `  </metadata>`,
    `  <manifest>`,
    manifest,
    `  </manifest>`,
    `  <spine toc="ncx">`,
    spine,
    `  </spine>`,
    `</package>`,
  ].join("\n");
}

function buildNav(
  pages: Array<{ pageNum: number }>,
  title: string,
): string {
  const items = pages
    .map(
      ({ pageNum }) =>
        `      <li><a href="${pageFileName(pageNum)}">` +
        `Page ${pageNum}</a></li>`,
    )
    .join("\n");

  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<!DOCTYPE html>`,
    `<html xmlns="http://www.w3.org/1999/xhtml"`,
    `      xmlns:epub="http://www.idpf.org/2007/ops"`,
    `      xml:lang="ar">`,
    `<head>`,
    `  <meta charset="UTF-8"/>`,
    `  <title>${escapeXml(title)}</title>`,
    `</head>`,
    `<body>`,
    `  <nav epub:type="toc" id="toc">`,
    `    <h1>Pages</h1>`,
    `    <ol>`,
    items,
    `    </ol>`,
    `  </nav>`,
    `</body>`,
    `</html>`,
  ].join("\n");
}

function buildNcx(
  pages: Array<{ pageNum: number }>,
  title: string,
  uuid: string,
): string {
  const navPoints = pages
    .map(
      ({ pageNum }) =>
        `    <navPoint id="${pageId(pageNum)}" playOrder="${pageNum}">\n` +
        `      <navLabel><text>Page ${pageNum}</text></navLabel>\n` +
        `      <content src="${pageFileName(pageNum)}"/>\n` +
        `    </navPoint>`,
    )
    .join("\n");

  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">`,
    `  <head>`,
    `    <meta name="dtb:uid" content="urn:uuid:${uuid}"/>`,
    `    <meta name="dtb:depth" content="1"/>`,
    `    <meta name="dtb:totalPageCount" content="${pages.length}"/>`,
    `    <meta name="dtb:maxPageNumber" content="${pages.length}"/>`,
    `  </head>`,
    `  <docTitle><text>${escapeXml(title)}</text></docTitle>`,
    `  <navMap>`,
    navPoints,
    `  </navMap>`,
    `</ncx>`,
  ].join("\n");
}

function buildPageXhtml(
  pageNum: number,
  totalPages: number,
  html: string,
  title: string,
  imageOnly?: boolean,
): string {
  const body = imageOnly
    ? [
        `  <!-- PDF page ${pageNum} of ${totalPages} — image -->`,
        `  <figure class="page-image">`,
        `    <img src="images/${pageFileName(pageNum).replace(".xhtml", ".jpg")}"`,
        `         alt="Page ${pageNum} illustration"/>`,
        `  </figure>`,
      ].join("\n")
    : [
        `  <!-- PDF page ${pageNum} of ${totalPages} -->`,
        `  ${html.trim()}`,
      ].join("\n");

  return [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<!DOCTYPE html>`,
    `<html xmlns="http://www.w3.org/1999/xhtml"`,
    `      xmlns:epub="http://www.idpf.org/2007/ops"`,
    `      xml:lang="ar">`,
    `<head>`,
    `  <meta charset="UTF-8"/>`,
    `  <title>${escapeXml(title)} — Page ${pageNum}</title>`,
    `  <link rel="stylesheet" type="text/css" href="style.css"/>`,
    `</head>`,
    `<body>`,
    body,
    `</body>`,
    `</html>`,
  ].join("\n");
}

// ---------------------------------------------------------------------------
// XML/HTML utilities
// ---------------------------------------------------------------------------

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Normalise a raw HTML fragment from the LLM so it embeds cleanly in XHTML.
 *
 *  1. Strip markdown code fences (```html … ```)
 *  2. Strip <html>/<head>/<body> wrapper tags if the model added them
 *  3. Fix non-self-closing void elements (br, hr, img, input)
 *  4. If no HTML tags remain, treat the whole thing as plain text and
 *     wrap each paragraph in <p> with the appropriate dir attribute
 */
export function cleanHtmlFragment(raw: string): string {
  // 1. Strip code fences
  let html = raw
    .replace(/^```(?:html|xhtml|xml)?\s*/i, "")
    .replace(/\s*```\s*$/i, "")
    .trim();

  // 2. Strip outer HTML/head/body wrappers
  html = html
    .replace(/<!DOCTYPE[^>]*>/gi, "")
    .replace(/<html(?:\s[^>]*)?>/gi, "")
    .replace(/<\/html>/gi, "")
    .replace(/<head(?:\s[^>]*)?>[\s\S]*?<\/head>/gi, "")
    .replace(/<body(?:\s[^>]*)?>/gi, "")
    .replace(/<\/body>/gi, "")
    .trim();

  // 3. Normalise void elements to self-closing (XHTML requirement)
  html = html
    .replace(/<br\s*>/gi, "<br/>")
    .replace(/<hr\s*>/gi, "<hr/>")
    .replace(/<img([^>]*[^/])>/gi, "<img$1/>")
    .replace(/<input([^>]*[^/])>/gi, "<input$1/>");

  // 4. Plain-text fallback
  if (!/<[a-z]/i.test(html)) {
    const isArabic = /[\u0600-\u06FF]/.test(html);
    const dir = isArabic ? "rtl" : "ltr";
    html = html
      .split(/\n{2,}/)
      .map((block) => block.trim())
      .filter(Boolean)
      .map((block) => {
        const escaped = block
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/\n/g, "<br/>");
        return `<p dir="${dir}">${escaped}</p>`;
      })
      .join("\n");
  }

  return html || `<p dir="rtl"><!-- empty page --></p>`;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface EpubPage {
  pageNum: number;
  /** Extracted HTML fragment (text pages) or empty string (image-only pages). */
  html: string;
  /**
   * When true, the page is a full-page image (cover, illustration).
   * The JPEG must exist at `<imagesDir>/page_NNN.jpg`.
   * The XHTML file will contain an `<img>` instead of the extracted HTML.
   */
  imageOnly?: boolean;
}

/**
 * Assemble an EPUB 3 file from per-page HTML fragments and/or images.
 *
 * @param pages      Ordered array of { pageNum, html, imageOnly? }.
 *                   For image-only pages the JPEG must exist at
 *                   `<imagesDir>/page_NNN.jpg`.
 * @param outputPath Destination .epub file path.
 * @param pdfPath    Source PDF path (used to derive the book title).
 * @param imagesDir  Directory containing `page_NNN.jpg` files for image-only pages.
 *                   Pass `null` / `undefined` when no image pages exist.
 * @param verbose    If true, log progress to stdout.
 */
export async function generateEpub(
  pages: EpubPage[],
  outputPath: string,
  pdfPath: string,
  imagesDir: string | null | undefined,
  verbose: boolean,
): Promise<void> {
  const title = deriveTitleFromPath(pdfPath);
  const uuid = randomUUID();
  const total = pages.length;

  const imagePages = pages.filter((p) => p.imageOnly);
  if (verbose) {
    console.log(
      `[EPUB] Building "${title}" — ${total} page(s)` +
      (imagePages.length ? `, ${imagePages.length} image page(s)` : "") +
      `...`,
    );
  }

  const zip = new JSZip();

  // mimetype MUST be first and stored without compression (EPUB spec §3.3)
  zip.file("mimetype", "application/epub+zip", { compression: "STORE" });

  // META-INF
  zip.file("META-INF/container.xml", buildContainer());

  // EPUB package documents
  zip.file("EPUB/content.opf", buildOpf(pages, title, uuid));
  zip.file("EPUB/nav.xhtml",   buildNav(pages, title));
  zip.file("EPUB/toc.ncx",     buildNcx(pages, title, uuid));
  zip.file("EPUB/style.css",   STYLE_CSS);

  // Embed JPEG files for image-only pages
  if (imagesDir && imagePages.length > 0) {
    for (const { pageNum } of imagePages) {
      const imgName = `page_${String(pageNum).padStart(3, "0")}.jpg`;
      const imgPath = join(imagesDir, imgName);
      try {
        const imgData = await readFile(imgPath);
        zip.file(`EPUB/images/${imgName}`, imgData);
        if (verbose) {
          console.log(`[EPUB]   images/${imgName} — ${Math.round(imgData.length / 1024)} KB`);
        }
      } catch {
        console.warn(`[EPUB]   Warning: image not found for page ${pageNum}: ${imgPath}`);
      }
    }
  }

  // One XHTML file per page
  for (const { pageNum, html, imageOnly } of pages) {
    const cleanedHtml = imageOnly ? "" : cleanHtmlFragment(html);
    const xhtml = buildPageXhtml(pageNum, total, cleanedHtml, title, imageOnly);
    zip.file(`EPUB/${pageFileName(pageNum)}`, xhtml);

    if (verbose && !imageOnly) {
      console.log(`[EPUB]   page_${String(pageNum).padStart(3, "0")}.xhtml — ${cleanedHtml.length} chars`);
    } else if (verbose && imageOnly) {
      console.log(`[EPUB]   page_${String(pageNum).padStart(3, "0")}.xhtml — image`);
    }
  }

  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  await mkdir(dirname(outputPath), { recursive: true });
  await writeFile(outputPath, buffer);

  if (verbose) {
    console.log(`[EPUB] Written: ${outputPath} (${Math.round(buffer.length / 1024)} KB)`);
  }
}
