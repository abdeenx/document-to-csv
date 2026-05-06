import { z } from "zod";

export const LmStudioConfigSchema = z.object({
  baseUrl: z.string().url().default("http://localhost:1234/v1"),
  apiKey: z.string().default("lm-studio"),
});
export type LmStudioConfig = z.infer<typeof LmStudioConfigSchema>;

export const ModelIdSchema = z.object({
  ocr: z.string().default("mlx-community/DeepSeek-OCR-8bit"),
  structurer: z
    .string()
    .default(
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    ),
});
export type ModelIds = z.infer<typeof ModelIdSchema>;

export const CliArgsSchema = z.object({
  imagePath: z.string().min(1, "Image path is required"),
  outputPath: z.string().optional(),
  lmStudioUrl: z.string().url().default("http://localhost:1234/v1"),
  ocrModel: z.string().default("mlx-community/DeepSeek-OCR-8bit"),
  structurerModel: z
    .string()
    .default(
      "zecanard/gemma-4-e4b-it-ultra-uncensored-heretic-mlx-int5-affine",
    ),
  verbose: z.boolean().default(false),
  excel: z.boolean().default(false),
  word: z.boolean().default(false),
});
export type CliArgs = z.infer<typeof CliArgsSchema>;

export const OcrResultSchema = z.object({
  rawText: z.string(),
  model: z.string(),
  tokensUsed: z.number().optional(),
  // Single image (from image input path) forwarded to Gemma4 for visual verification.
  imageBase64: z.string().optional(),
  imageMimeType: z.string().optional(),
  // Per-page rendered images (from PDF OCR pass) — one entry per page.
  // When set, these take priority over imageBase64/imageMimeType in the Gemma4 message.
  pageImages: z.array(z.object({ base64: z.string(), mimeType: z.string() })).optional(),
});
export type OcrResult = z.infer<typeof OcrResultSchema>;

export const WriteCsvToolArgsSchema = z.object({
  csv_content: z
    .string()
    .min(1)
    .describe("The full CSV content to write, including headers"),
  reasoning: z
    .string()
    .describe(
      "Brief explanation of the structure choices made (column mapping, row layout, etc.)",
    ),
});
export type WriteCsvToolArgs = z.infer<typeof WriteCsvToolArgsSchema>;

export const CsvResultSchema = z.object({
  csvContent: z.string(),
  outputPath: z.string(),
  reasoning: z.string(),
});
export type CsvResult = z.infer<typeof CsvResultSchema>;

// ---------------------------------------------------------------------------
// PDF layout types (used for layout-faithful Excel output)
// ---------------------------------------------------------------------------

/** A single text run extracted from a PDF page via pdfjs getTextContent(). */
export interface PdfTextItem {
  /** Left edge in PDF points, from left of page. */
  x: number;
  /** Baseline Y in PDF points, from BOTTOM of page (PDF convention). */
  y: number;
  str: string;
  /** Approximate font size in points. */
  fontSize: number;
  /** Run width in PDF points. */
  width: number;
}

/**
 * A bounding box for a painted image detected via the PDF operator list.
 * All coordinates are in PDF points using PDF convention: origin bottom-left.
 */
export interface PdfImageRegion {
  /** Left edge in PDF points from left of page. */
  x: number;
  /** Bottom edge in PDF points from bottom of page (PDF coords, not flipped). */
  y: number;
  width: number;
  height: number;
}

export interface PdfPageLayout {
  pageWidth: number;   // PDF points
  pageHeight: number;  // PDF points
  textItems: PdfTextItem[];
  imageRegions: PdfImageRegion[];
}

// ---------------------------------------------------------------------------
// Word output / progress tracking
// ---------------------------------------------------------------------------

/**
 * Per-page extraction results from all three sources plus the final
 * corroborated text that goes into the Word document.
 */
export interface PageExtraction {
  /** Text from the pdfjs embedded text layer. */
  pdfjsText: string;
  /** Text from the DeepSeek-OCR visual pass. */
  ocrText: string;
  /** Text from the Gemma4 direct vision extraction pass. */
  gemmaText: string;
  /** Final corroborated text (reconciled by Gemma4). */
  corroborated: string;
}

/**
 * Progress file written after every completed page so the conversion can
 * resume from where it left off if interrupted.
 */
export const WordProgressSchema = z.object({
  version: z.literal(1),
  pdfPath: z.string(),
  totalPages: z.number().int().positive(),
  pages: z.record(
    z.string().regex(/^\d+$/),
    z.object({
      pdfjsText: z.string(),
      ocrText: z.string(),
      gemmaText: z.string(),
      corroborated: z.string(),
    }),
  ),
});
export type WordProgress = z.infer<typeof WordProgressSchema>;

// ---------------------------------------------------------------------------

export const SupportedImageMimeSchema = z.enum([
  "image/jpeg",
  "image/png",
  "image/gif",
  "image/webp",
]);
export type SupportedImageMime = z.infer<typeof SupportedImageMimeSchema>;

export const IMAGE_EXTENSION_TO_MIME: Record<string, SupportedImageMime> = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  webp: "image/webp",
};
