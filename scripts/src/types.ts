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
});
export type CliArgs = z.infer<typeof CliArgsSchema>;

export const OcrResultSchema = z.object({
  rawText: z.string(),
  model: z.string(),
  tokensUsed: z.number().optional(),
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
