import { writeFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import type { OcrResult } from "./types.js";

// ---------------------------------------------------------------------------
// RFC 4180 CSV parser (duplicated locally so excel-generator has no circular
// dependency on csv-generator's internal sanitizeCsv)
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
        if (line[i + 1] === '"') {
          current += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        current += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === ",") {
        fields.push(current);
        current = "";
        i++;
      } else {
        current += ch;
        i++;
      }
    }
  }
  fields.push(current);
  return fields;
}

function parseCsvContent(csvContent: string): string[][] {
  return csvContent
    .split("\n")
    .filter((l) => l.trim() !== "")
    .map(parseCsvRow);
}

// ---------------------------------------------------------------------------
// Excel generator
// ---------------------------------------------------------------------------

export interface ExcelGeneratorOptions {
  csvContent: string;
  outputPath: string;
  ocrResult: OcrResult;
  verbose: boolean;
}

export async function generateExcel(opts: ExcelGeneratorOptions): Promise<void> {
  const { csvContent, outputPath, ocrResult, verbose } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "document-to-csv";
  workbook.created = new Date();

  // ── Sheet 1: Data ──────────────────────────────────────────────────────────
  const dataSheet = workbook.addWorksheet("Data");
  const rows = parseCsvContent(csvContent);

  if (rows.length > 0) {
    const headerRow = rows[0]!;
    const colCount = headerRow.length;

    // Set column widths: auto-size heuristic (min 12, max 50 chars)
    for (let c = 1; c <= colCount; c++) {
      const maxLen = rows.reduce((max, row) => {
        const cell = row[c - 1] ?? "";
        return Math.max(max, cell.length);
      }, headerRow[c - 1]?.length ?? 10);
      dataSheet.getColumn(c).width = Math.min(Math.max(maxLen + 2, 12), 50);
    }

    // Header row — bold, light blue fill
    const xlsxHeader = dataSheet.addRow(headerRow);
    xlsxHeader.font = { bold: true };
    xlsxHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD6E4F7" },
    };
    xlsxHeader.border = {
      bottom: { style: "thin", color: { argb: "FF8EB4E3" } },
    };
    xlsxHeader.alignment = { wrapText: false };

    // Data rows — alternating light grey
    for (let r = 1; r < rows.length; r++) {
      const dataRow = dataSheet.addRow(rows[r]!);
      dataRow.alignment = { wrapText: true, vertical: "top" };
      if (r % 2 === 0) {
        dataRow.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF5F5F5" },
        };
      }
    }

    // Freeze the header row
    dataSheet.views = [{ state: "frozen", ySplit: 1 }];
  }

  // ── Per-page image sheets ──────────────────────────────────────────────────
  const pageImages = ocrResult.pageImages ?? [];

  if (pageImages.length > 0) {
    if (verbose) {
      console.log(`[Excel] Embedding ${pageImages.length} page image(s) into workbook...`);
    }

    for (let idx = 0; idx < pageImages.length; idx++) {
      const page = pageImages[idx]!;
      const sheetName = pageImages.length === 1 ? "Document" : `Page ${idx + 1}`;
      const imgSheet = workbook.addWorksheet(sheetName);

      // Detect extension from mimeType
      const ext = page.mimeType === "image/png" ? "png" : "jpeg";

      const imageId = workbook.addImage({
        base64: page.base64,
        extension: ext,
      });

      // Display the page at a comfortable reading size.
      // ExcelJS ext uses pixels. We target ~680px wide × 960px tall (portrait A4-ish).
      // Wider sheets (landscape) will be scaled proportionally by most viewers.
      const W = 680;
      const H = 960;

      imgSheet.addImage(imageId, {
        tl: { col: 0, row: 0 },
        ext: { width: W, height: H },
      });

      // Set column A wide enough to "hold" the image width so it displays cleanly
      imgSheet.getColumn(1).width = Math.ceil(W / 7);

      // Set enough rows with sufficient height to not clip the image.
      // Excel row height is in points; 1pt ≈ 1.333px
      const rowCount = Math.ceil(H / 14);
      for (let r = 1; r <= rowCount + 5; r++) {
        imgSheet.getRow(r).height = 14;
      }
    }
  } else if (ocrResult.imageBase64 && ocrResult.imageMimeType) {
    // Single image input (non-PDF path)
    const imgSheet = workbook.addWorksheet("Document");
    const ext = ocrResult.imageMimeType === "image/png" ? "png" : "jpeg";
    const imageId = workbook.addImage({
      base64: ocrResult.imageBase64,
      extension: ext,
    });
    imgSheet.addImage(imageId, {
      tl: { col: 0, row: 0 },
      ext: { width: 680, height: 960 },
    });
    imgSheet.getColumn(1).width = Math.ceil(680 / 7);
    for (let r = 1; r <= 75; r++) imgSheet.getRow(r).height = 14;
  }

  // Write to disk
  const buffer = await workbook.xlsx.writeBuffer();
  await writeFile(outputPath, Buffer.from(buffer));

  if (verbose) {
    console.log(`[Excel] Workbook written: ${outputPath}`);
  }
}
