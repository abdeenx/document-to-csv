/**
 * Excel generator — layout-faithful xlsx output.
 *
 * When PDF page-layout data is available (--excel on a PDF input) the output
 * mirrors the PDF structure:
 *   • "Data" sheet   — clean styled table from the Gemma4-generated CSV
 *   • "Page N" sheet — one per PDF page, grid-mapped so text lands at the
 *                      correct relative position and any images detected in
 *                      the PDF are cropped from the rendered page render and
 *                      embedded inline at the matching cell anchor.
 *
 * Grid resolution (tunable via constants below):
 *   GRID_COL_PTS  PDF points per Excel column  (default 6 pt ≈ 1 char wide)
 *   GRID_ROW_PTS  PDF points per Excel row      (default 12 pt = 12 pt tall)
 */

import { writeFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import sharp from "sharp";
import type { OcrResult, PdfPageLayout } from "./types.js";

// ---------------------------------------------------------------------------
// Grid constants
// ---------------------------------------------------------------------------

const GRID_COL_PTS  = 6;   // PDF pts → 1 Excel column
const GRID_ROW_PTS  = 12;  // PDF pts → 1 Excel row
const EXCEL_COL_W   = 1.0; // chars (≈ 6 pt at Calibri 11)
const EXCEL_ROW_H   = 9;   // points (slightly below GRID_ROW_PTS so text lines up)

// ---------------------------------------------------------------------------
// RFC 4180 CSV parser (local copy — avoids circular dep on csv-generator)
// ---------------------------------------------------------------------------

function parseCsvRow(line: string): string[] {
  const fields: string[] = [];
  let current = "";
  let inQuotes = false;
  let i = 0;
  while (i < line.length) {
    const ch = line[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') { current += '"'; i += 2; }
        else { inQuotes = false; i++; }
      } else { current += ch; i++; }
    } else {
      if (ch === '"') { inQuotes = true; i++; }
      else if (ch === ",") { fields.push(current); current = ""; i++; }
      else { current += ch; i++; }
    }
  }
  fields.push(current);
  return fields;
}

function parseCsvContent(csv: string): string[][] {
  return csv.split("\n").filter((l) => l.trim() !== "").map(parseCsvRow);
}

// ---------------------------------------------------------------------------
// Styled "Data" sheet (always written)
// ---------------------------------------------------------------------------

function addDataSheet(workbook: ExcelJS.Workbook, csvContent: string): void {
  const sheet = workbook.addWorksheet("Data");
  const rows  = parseCsvContent(csvContent);
  if (rows.length === 0) return;

  const headerRow = rows[0]!;
  const colCount  = headerRow.length;

  // Auto-size columns
  for (let c = 1; c <= colCount; c++) {
    const maxLen = rows.reduce((m, r) => Math.max(m, (r[c - 1] ?? "").length), 0);
    sheet.getColumn(c).width = Math.min(Math.max(maxLen + 2, 12), 50);
  }

  // Header row
  const xlsxHdr = sheet.addRow(headerRow);
  xlsxHdr.font      = { bold: true };
  xlsxHdr.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E4F7" } };
  xlsxHdr.border    = { bottom: { style: "thin", color: { argb: "FF8EB4E3" } } };
  xlsxHdr.alignment = { wrapText: false };

  // Data rows
  for (let r = 1; r < rows.length; r++) {
    const dr = sheet.addRow(rows[r]!);
    dr.alignment = { wrapText: true, vertical: "top" };
    if (r % 2 === 0) {
      dr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
    }
  }

  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

// ---------------------------------------------------------------------------
// Layout-faithful "Page N" sheets
// ---------------------------------------------------------------------------

function pdfColFor(x: number): number {
  return Math.max(1, Math.floor(x / GRID_COL_PTS) + 1); // 1-indexed
}

function pdfRowFor(yFromTop: number): number {
  return Math.max(1, Math.floor(yFromTop / GRID_ROW_PTS) + 1); // 1-indexed
}

async function addLayoutSheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  layout: PdfPageLayout,
  renderedBase64: string | undefined,
  renderedMimeType: string | undefined,
  verbose: boolean,
): Promise<void> {
  const sheet = workbook.addWorksheet(sheetName);
  const { pageWidth, pageHeight } = layout;

  const totalCols = Math.ceil(pageWidth  / GRID_COL_PTS) + 10;
  const totalRows = Math.ceil(pageHeight / GRID_ROW_PTS) + 10;

  // Set uniform grid dimensions
  for (let c = 1; c <= totalCols; c++) sheet.getColumn(c).width  = EXCEL_COL_W;
  for (let r = 1; r <= totalRows; r++) sheet.getRow(r).height = EXCEL_ROW_H;

  // ── Place text items ──────────────────────────────────────────────────────
  for (const item of layout.textItems) {
    // PDF y is baseline from bottom → convert to "from top"
    const yFromTop = pageHeight - item.y;
    const col = pdfColFor(item.x);
    const row = pdfRowFor(yFromTop);

    const cell = sheet.getCell(row, col);
    // Append if already occupied (two runs that rounded to the same cell)
    cell.value = cell.value ? `${cell.value as string} ${item.str}` : item.str;
    const fs = Math.round(item.fontSize);
    cell.font = { size: fs > 4 && fs < 72 ? fs : 10 };
  }

  // ── Embed image regions ───────────────────────────────────────────────────
  if (renderedBase64 && layout.imageRegions.length > 0) {
    const imgBuf  = Buffer.from(renderedBase64, "base64");
    const meta    = await sharp(imgBuf).metadata();
    const imgW    = meta.width  ?? 1;
    const imgH    = meta.height ?? 1;

    // Scale factors: rendered px → PDF pts
    const scaleX = imgW / pageWidth;
    const scaleY = imgH / pageHeight;

    for (const region of layout.imageRegions) {
      // Convert PDF bottom-origin → top-origin for image cropping
      const topFromTop = pageHeight - region.y - region.height;

      const cropLeft = Math.max(0, Math.round(region.x     * scaleX));
      const cropTop  = Math.max(0, Math.round(topFromTop   * scaleY));
      const cropW    = Math.round(region.width  * scaleX);
      const cropH    = Math.round(region.height * scaleY);

      // Clamp to image boundaries
      const safeLeft = Math.min(cropLeft, imgW - 1);
      const safeTop  = Math.min(cropTop,  imgH - 1);
      const safeW    = Math.min(cropW, imgW - safeLeft);
      const safeH    = Math.min(cropH, imgH - safeTop);

      if (safeW < 2 || safeH < 2) continue;

      let cropped: Buffer;
      try {
        cropped = await sharp(imgBuf)
          .extract({ left: safeLeft, top: safeTop, width: safeW, height: safeH })
          .png()
          .toBuffer();
      } catch {
        if (verbose) console.log(`[Excel] Skipping image crop (extract failed): ${JSON.stringify(region)}`);
        continue;
      }

      const imageId = workbook.addImage({ base64: cropped.toString("base64"), extension: "png" });

      // Map image bounding box → Excel cell anchors (0-indexed for addImage)
      const startCol = Math.max(0, pdfColFor(region.x) - 1);
      const startRow = Math.max(0, pdfRowFor(topFromTop) - 1);
      const endCol   = startCol + Math.max(1, Math.ceil(region.width  / GRID_COL_PTS));
      const endRow   = startRow + Math.max(1, Math.ceil(region.height / GRID_ROW_PTS));

      // exceljs v4 types declare Anchor with native units; cast to satisfy them.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      sheet.addImage(imageId, {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        tl: { col: startCol, row: startRow } as any,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        br: { col: endCol,   row: endRow   } as any,
        editAs: "oneCell",
      });

      if (verbose) {
        console.log(
          `[Excel] Embedded image region x=${region.x.toFixed(0)} y=${region.y.toFixed(0)} ` +
          `w=${region.width.toFixed(0)} h=${region.height.toFixed(0)} → ` +
          `cols ${startCol}-${endCol}, rows ${startRow}-${endRow}`,
        );
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface ExcelGeneratorOptions {
  csvContent: string;
  outputPath: string;
  ocrResult: OcrResult;
  pageLayouts?: PdfPageLayout[];
  verbose: boolean;
}

export async function generateExcel(opts: ExcelGeneratorOptions): Promise<void> {
  const { csvContent, outputPath, ocrResult, pageLayouts, verbose } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "document-to-csv";
  workbook.created = new Date();

  // ── Data sheet (always present) ───────────────────────────────────────────
  addDataSheet(workbook, csvContent);

  // ── Layout sheets (PDF + --excel only) ───────────────────────────────────
  if (pageLayouts && pageLayouts.length > 0) {
    const pageImages = ocrResult.pageImages ?? [];
    for (let i = 0; i < pageLayouts.length; i++) {
      const layout    = pageLayouts[i]!;
      const img       = pageImages[i];
      const sheetName = pageLayouts.length === 1 ? "Document" : `Page ${i + 1}`;

      if (verbose) {
        console.log(
          `[Excel] Building layout sheet "${sheetName}" — ` +
          `${layout.textItems.length} text item(s), ` +
          `${layout.imageRegions.length} image region(s)`,
        );
      }

      await addLayoutSheet(workbook, sheetName, layout, img?.base64, img?.mimeType, verbose);
    }
  } else {
    // Fallback for image inputs: embed the source image on a "Document" sheet
    const imgSheet = workbook.addWorksheet("Document");
    const src =
      ocrResult.pageImages?.[0] ??
      (ocrResult.imageBase64 && ocrResult.imageMimeType
        ? { base64: ocrResult.imageBase64, mimeType: ocrResult.imageMimeType }
        : null);

    if (src) {
      const ext = src.mimeType === "image/png" ? "png" : "jpeg";
      const imageId = workbook.addImage({ base64: src.base64, extension: ext });
      imgSheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 680, height: 960 } });
      imgSheet.getColumn(1).width = 100;
      for (let r = 1; r <= 75; r++) imgSheet.getRow(r).height = 14;
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  await writeFile(outputPath, Buffer.from(buffer));

  if (verbose) console.log(`[Excel] Workbook written: ${outputPath}`);
}
