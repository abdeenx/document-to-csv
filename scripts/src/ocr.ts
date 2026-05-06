import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import type OpenAI from "openai";
import sharp from "sharp";
import {
  IMAGE_EXTENSION_TO_MIME,
  type OcrResult,
  type SupportedImageMime,
} from "./types.js";
import { buildDataUrl } from "./lm-studio-client.js";

const MAX_DIMENSION = 1600;

// ---------------------------------------------------------------------------
// Shared OCR system prompt (used by both image and PDF OCR passes)
// ---------------------------------------------------------------------------

const OCR_SYSTEM_PROMPT = [
  "You are a precise OCR engine for structured documents.",
  "Extract every piece of text visible in the image. Follow these rules exactly:",
  "",
  "LAYOUT:",
  "- For tables and grids: use tab characters (\\t) between columns and newlines between rows.",
  "- Column headers that visually wrap across two lines belong to a single column — join them with a space (e.g. 'AUTO\\nRENEW' → 'AUTO RENEW').",
  "- Cell values that visually wrap across two lines belong to one cell — join them with a space (e.g. 'Apr 29,\\n2028' → 'Apr 29, 2028').",
  "",
  "INTERACTIVE ELEMENTS:",
  "- Toggle switches / pill switches: inspect the visual state of each toggle carefully.",
  "  A toggle that is colored (teal, green, blue, any non-grey color) is ON — output 'Yes'.",
  "  A toggle that is grey, white, or hollow is OFF — output 'No'.",
  "  CRITICAL: Never leave a toggle cell blank. Every row with a toggle column MUST have 'Yes' or 'No' — a blank toggle cell is wrong.",
  "- Checkboxes: checked = 'Yes', unchecked = 'No'.",
  "- Action buttons and links (e.g. INFO, EDIT, SETUP, DELETE) that appear in every row as row-level actions: skip them entirely — do not include them as column values or extra columns.",
  "- Tooltip popups, notification banners, modal overlays, and dropdown menus that appear on top of the content: skip them entirely.",
  "",
  "OUTPUT:",
  "- Output ONLY the extracted text. No commentary, no summaries, no explanations.",
  "- Include every header, label, email, number, date, and status value exactly as shown.",
].join("\n");

// ---------------------------------------------------------------------------
// Core OCR API call — shared by image and PDF paths
// ---------------------------------------------------------------------------

/**
 * Send a pre-encoded image to the OCR model and return the extracted text.
 * Callers are responsible for resizing the image before calling this function.
 */
export async function callOcrModel(
  client: OpenAI,
  base64: string,
  mimeType: string,
  modelId: string,
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(
      `[OCR] Sending ${Math.round((base64.length * 0.75) / 1024)} KB to ${modelId}...`,
    );
  }

  const dataUrl = buildDataUrl(base64, mimeType);

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: OCR_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          { type: "image_url", image_url: { url: dataUrl } },
          { type: "text", text: "Extract all text from this image." },
        ],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const text = response.choices[0]?.message.content;
  if (!text) throw new Error("[OCR] Model returned an empty response.");

  if (verbose) {
    console.log(
      `[OCR] Extracted ${text.length} chars. Tokens: ${response.usage?.total_tokens ?? "unknown"}`,
    );
  }

  return text;
}

// ---------------------------------------------------------------------------
// Image loading + preprocessing
// ---------------------------------------------------------------------------

async function loadAndPreprocess(imagePath: string): Promise<{
  base64: string;
  mimeType: SupportedImageMime;
  originalWidth: number;
  originalHeight: number;
  finalWidth: number;
  finalHeight: number;
  resized: boolean;
}> {
  const ext = extname(imagePath).slice(1).toLowerCase();
  const mimeType = IMAGE_EXTENSION_TO_MIME[ext];

  if (!mimeType) {
    throw new Error(
      `Unsupported image extension ".${ext}". Supported: ${Object.keys(IMAGE_EXTENSION_TO_MIME).join(", ")}`,
    );
  }

  const imageBuffer = await readFile(imagePath);
  const metadata = await sharp(imageBuffer).metadata();
  const originalWidth = metadata.width ?? 0;
  const originalHeight = metadata.height ?? 0;

  const needsResize =
    originalWidth > MAX_DIMENSION || originalHeight > MAX_DIMENSION;

  if (!needsResize) {
    return {
      base64: imageBuffer.toString("base64"),
      mimeType,
      originalWidth,
      originalHeight,
      finalWidth: originalWidth,
      finalHeight: originalHeight,
      resized: false,
    };
  }

  const resizedBuffer = await sharp(imageBuffer)
    .resize(MAX_DIMENSION, MAX_DIMENSION, {
      fit: "inside",
      withoutEnlargement: true,
    })
    .jpeg({ quality: 92 })
    .toBuffer();

  const resizedMeta = await sharp(resizedBuffer).metadata();

  return {
    base64: resizedBuffer.toString("base64"),
    mimeType: "image/jpeg",
    originalWidth,
    originalHeight,
    finalWidth: resizedMeta.width ?? 0,
    finalHeight: resizedMeta.height ?? 0,
    resized: true,
  };
}

// ---------------------------------------------------------------------------
// Public entry point for image files
// ---------------------------------------------------------------------------

export async function extractTextWithOcr(
  client: OpenAI,
  imagePath: string,
  modelId: string,
  verbose: boolean,
): Promise<OcrResult> {
  if (verbose) {
    console.log(`[OCR] Loading image: ${imagePath}`);
  }

  const result = await loadAndPreprocess(imagePath);

  if (verbose) {
    if (result.resized) {
      console.log(
        `[OCR] Resized ${result.originalWidth}×${result.originalHeight} → ${result.finalWidth}×${result.finalHeight} (JPEG q92)`,
      );
    } else {
      console.log(
        `[OCR] Image is ${result.originalWidth}×${result.originalHeight} — no resize needed`,
      );
    }
  }

  const rawText = await callOcrModel(client, result.base64, result.mimeType, modelId, verbose);

  return {
    rawText,
    model: modelId,
    // Pass the preprocessed image to the structuring step so Gemma4 can use
    // visual reasoning to verify and correct OCR column alignment.
    imageBase64: result.base64,
    imageMimeType: result.mimeType,
  };
}
