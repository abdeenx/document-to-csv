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
 *       ├── images/                    ← all cropped/extracted images
 *       │   ├── page_001_img1.jpg      ← full-page cover
 *       │   ├── page_042_img1.jpg      ← inline illustration from page 42
 *       │   └── …
 *       ├── page_001.xhtml             ← PDF page 1
 *       ├── page_002.xhtml             ← PDF page 2
 *       └── …
 *
 * Each XHTML file is self-contained and editable in any text editor.
 * Images are referenced as <img src="images/page_NNN_imgM.jpg"> within
 * the normal HTML flow — full-page images are <figure class="page-image">,
 * inline illustrations are <figure class="inline-image">.
 */

import { readFile, writeFile, mkdir, readdir } from "node:fs/promises";
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

/* Full-page images (book covers, purely decorative pages) */
.page-image {
  margin: 0;
  padding: 0;
  text-align: center;
  page-break-inside: avoid;
  break-inside: avoid;
}

.page-image img {
  max-width: 100%;
  max-height: 95vh;
  height: auto;
  display: block;
  margin: 0 auto;
}

/* Inline illustrations embedded within text */
.inline-image {
  margin: 1em auto;
  text-align: center;
  page-break-inside: avoid;
  break-inside: avoid;
}

.inline-image img {
  max-width: 100%;
  height: auto;
  display: block;
  margin: 0 auto;
}

figcaption {
  font-size: 0.85em;
  color: #555;
  margin-top: 0.4em;
  font-style: italic;
}

/* Useful utility classes for manual post-editing */
.rtl { direction: rtl; text-align: right; }
.ltr { direction: ltr; text-align: left; }
.center { text-align: center; }
.bold { font-weight: bold; }
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
  pages: Array<{ pageNum: number }>,
  title: string,
  uuid: string,
  imageFileNames: string[],
): string {
  // The first full-page image file (page_NNN.jpg without _imgM suffix) is the cover.
  const coverFile = imageFileNames.find((f) => /^page_\d{3}\.jpg$/.test(f));

  const imageItems = imageFileNames.map((fname) => {
    const id = `img_${fname.replace(/[^a-zA-Z0-9]/g, "_")}`;
    const isCover = fname === coverFile;
    return (
      `    <item id="${id}" href="images/${fname}"` +
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
): string {
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
    `  <!-- PDF page ${pageNum} of ${totalPages} -->`,
    `  ${html.trim()}`,
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
 *
 * Note: <figure data-region="…"> elements are handled by processInlineImages()
 * in qwen3-epub.ts BEFORE the HTML reaches this function, so by the time
 * cleanHtmlFragment() is called the figures already contain proper <img> tags.
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
  /**
   * Final HTML fragment for this page.
   * May contain <figure class="page-image|inline-image"><img src="images/…"/></figure>
   * elements for visual content extracted by processInlineImages() in qwen3-epub.ts.
   */
  html: string;
}

/**
 * Assemble an EPUB 3 file from per-page HTML fragments.
 *
 * Visual content (covers, illustrations, inline images) is referenced via
 * <img src="images/page_NNN_imgM.jpg"> in the HTML; the actual JPEG files
 * are read from `imagesDir` and embedded in the EPUB automatically.
 *
 * @param pages      Ordered array of { pageNum, html }.
 * @param outputPath Destination .epub file path.
 * @param pdfPath    Source PDF path (used to derive the book title).
 * @param imagesDir  Directory containing cropped JPEG files (`page_NNN_imgM.jpg`).
 *                   Pass `null` / `undefined` when no images were extracted.
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

  // Scan the images directory — include every JPEG found, sorted by name.
  const imageFileNames: string[] = [];
  if (imagesDir) {
    try {
      const files = await readdir(imagesDir);
      for (const f of files.sort()) {
        if (/\.(jpg|jpeg|png|gif|webp)$/i.test(f)) {
          imageFileNames.push(f);
        }
      }
    } catch {
      // directory doesn't exist — no images were extracted
    }
  }

  if (verbose) {
    console.log(
      `[EPUB] Building "${title}" — ${total} page(s)` +
      (imageFileNames.length ? `, ${imageFileNames.length} image file(s)` : "") +
      `...`,
    );
  }

  const zip = new JSZip();

  // mimetype MUST be first and stored without compression (EPUB spec §3.3)
  zip.file("mimetype", "application/epub+zip", { compression: "STORE" });

  // META-INF + package documents
  zip.file("META-INF/container.xml", buildContainer());
  zip.file("EPUB/content.opf", buildOpf(pages, title, uuid, imageFileNames));
  zip.file("EPUB/nav.xhtml",   buildNav(pages, title));
  zip.file("EPUB/toc.ncx",     buildNcx(pages, title, uuid));
  zip.file("EPUB/style.css",   STYLE_CSS);

  // Embed all image files from the images directory
  for (const fname of imageFileNames) {
    try {
      const imgData = await readFile(join(imagesDir!, fname));
      zip.file(`EPUB/images/${fname}`, imgData);
      if (verbose) {
        console.log(`[EPUB]   images/${fname} — ${Math.round(imgData.length / 1024)} KB`);
      }
    } catch {
      console.warn(`[EPUB]   Warning: could not read image ${fname}`);
    }
  }

  // One XHTML file per page
  for (const { pageNum, html } of pages) {
    const cleanedHtml = cleanHtmlFragment(html);
    const xhtml = buildPageXhtml(pageNum, total, cleanedHtml, title);
    zip.file(`EPUB/${pageFileName(pageNum)}`, xhtml);

    if (verbose) {
      console.log(`[EPUB]   page_${String(pageNum).padStart(3, "0")}.xhtml — ${cleanedHtml.length} chars`);
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
