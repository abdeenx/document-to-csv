import OpenAI from "openai";
import type { LmStudioConfig } from "./types.js";

export function createLmStudioClient(config: LmStudioConfig): OpenAI {
  return new OpenAI({
    baseURL: config.baseUrl,
    apiKey: config.apiKey,
    defaultHeaders: {
      "Content-Type": "application/json",
    },
  });
}

export function buildDataUrl(
  base64Data: string,
  mimeType: string,
): `data:${string};base64,${string}` {
  return `data:${mimeType};base64,${base64Data}`;
}
