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
        role: "user",
        content: [
          {
            type: "image_url",
            image_url: { url: dataUrl },
          },
          {
            type: "text",
            text: [
              "You are a precise OCR engine. Extract ALL text from this image exactly as it appears.",
              "Preserve:",
              "- The original layout: rows, columns, tables, lists, and indentation",
              "- All numbers, punctuation, symbols, and special characters",
              "- Column/row headers and their alignment relationships",
              "- Any structured groupings (sections, sub-sections, nested rows)",
              "",
              "Output ONLY the extracted text. Do not add commentary, summaries, or formatting beyond what is in the image.",
              "If the image contains a table or spreadsheet, represent it with tab-separated values to preserve column alignment.",
            ].join("\n"),
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
