import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import type OpenAI from "openai";
import {
  IMAGE_EXTENSION_TO_MIME,
  type OcrResult,
  type SupportedImageMime,
} from "./types.js";
import { buildDataUrl } from "./lm-studio-client.js";

async function encodeImageToBase64(imagePath: string): Promise<{
  base64: string;
  mimeType: SupportedImageMime;
}> {
  const ext = extname(imagePath).slice(1).toLowerCase();
  const mimeType = IMAGE_EXTENSION_TO_MIME[ext];

  if (!mimeType) {
    throw new Error(
      `Unsupported image extension ".${ext}". Supported: ${Object.keys(IMAGE_EXTENSION_TO_MIME).join(", ")}`,
    );
  }

  const imageBuffer = await readFile(imagePath);
  const base64 = imageBuffer.toString("base64");

  return { base64, mimeType };
}

export async function extractTextWithOcr(
  client: OpenAI,
  imagePath: string,
  modelId: string,
  verbose: boolean,
): Promise<OcrResult> {
  if (verbose) {
    console.log(`[OCR] Encoding image: ${imagePath}`);
  }

  const { base64, mimeType } = await encodeImageToBase64(imagePath);
  const dataUrl = buildDataUrl(base64, mimeType);

  if (verbose) {
    console.log(
      `[OCR] Image encoded (${Math.round(base64.length * 0.75 / 1024)} KB). Sending to ${modelId}...`,
    );
  }

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      {
        role: "system",
        content: [
          "You are a precise OCR engine for structured documents.",
          "Extract every piece of text visible in the image. Follow these rules exactly:",
          "",
          "LAYOUT:",
          "- For tables and grids: use tab characters (\\t) between columns and newlines between rows.",
          "- Column headers that visually wrap across two lines belong to a single column — join them with a space (e.g. 'AUTO\\nRENEW' → 'AUTO RENEW').",
          "- Cell values that visually wrap across two lines belong to one cell — join them with a space (e.g. 'Apr 29,\\n2028' → 'Apr 29, 2028').",
          "",
          "INTERACTIVE ELEMENTS:",
          "- Toggle switches: output 'Yes' if ON (colored/filled) or 'No' if OFF (grey/empty).",
          "- Action buttons and links (e.g. INFO, EDIT, SETUP, DELETE) that appear in every row as row-level actions: skip them entirely — do not include them as column values or extra columns.",
          "- Tooltip popups, notification banners, modal overlays, and dropdown menus that appear on top of the content: skip them entirely.",
          "",
          "OUTPUT:",
          "- Output ONLY the extracted text. No commentary, no summaries, no explanations.",
          "- Include every header, label, email, number, date, and status value exactly as shown.",
        ].join("\n"),
      },
      {
        role: "user",
        content: [
          {
            type: "image_url",
            image_url: { url: dataUrl },
          },
          {
            type: "text",
            text: "Extract all text from this image.",
          },
        ],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const rawText = response.choices[0]?.message.content;

  if (!rawText) {
    throw new Error("[OCR] Model returned an empty response.");
  }

  if (verbose) {
    console.log(
      `[OCR] Extracted ${rawText.length} characters. Tokens used: ${response.usage?.total_tokens ?? "unknown"}`,
    );
  }

  return {
    rawText,
    model: modelId,
    tokensUsed: response.usage?.total_tokens,
  };
}
