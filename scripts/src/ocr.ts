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
          "You are a precise OCR engine.",
          "Your only job is to extract every piece of text visible in the image — nothing else.",
          "Rules:",
          "- Output ONLY the text that appears in the image. No commentary, no summaries, no explanations.",
          "- Preserve the spatial layout: maintain rows, columns, and indentation as they appear.",
          "- For tables and grids, use tab characters between columns and newlines between rows.",
          "- Include every header, label, value, number, symbol, and special character exactly as shown.",
          "- Do not paraphrase, infer, or add anything not visible in the image.",
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
