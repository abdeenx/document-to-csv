/**
 * qwen-pdf-to-word.ts
 *
 * Single-model PDF → Word (.docx) pipeline using Qwen2.5-VL.
 *
 * This is a lean, single-pass alternative to the 4-model corroboration pipeline
 * (`pdf-to-word.ts`). Instead of running four OCR models in parallel and then a
 * fifth corroboration call, every page is sent once to Qwen2.5-VL-7B-Instruct,
 * which is a multimodal vision-language model strong enough to extract Arabic and
 * Latin text accurately in a single call.
 *
 * Pipeline per page:
 *   1. Render page to JPEG (pdftoppm / mutool)
 *   2. Qwen2.5-VL extracts all text from the rendered image
 *   3. Strip any thinking/reasoning traces from the response
 *   4. Save to progress file (resumable)
 *
 * After all pages are done, the Word document is assembled by the shared
 * word-generator with full Arabic RTL support.
 *
 * Progress file: <output>.qwen-progress.json
 * Delete it to force a full re-run from scratch.
 */

import { readFile, writeFile, mkdir, rm, access } from "node:fs/promises";
import { join, dirname } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import type OpenAI from "openai";
import type { ChatCompletionContentPart } from "openai/resources/chat/completions";
import {
  detectRenderer,
  renderPageToJpeg,
  extractPdfjsPageTextsRaw,
  type Renderer,
} from "./pdf.js";
import { stripThinking } from "./ocr.js";
import { QwenProgressSchema, type QwenProgress } from "./types.js";
import { generateWordDoc } from "./word-generator.js";

// ---------------------------------------------------------------------------
// Qwen2.5-VL extraction prompt
// ---------------------------------------------------------------------------

const QWEN_OCR_SYSTEM_PROMPT = [
  "You are a precise document text extractor specializing in Arabic and Latin text.",
  "Extract all text from the document page image exactly as it appears.",
  "",
  "STRUCTURE RULES:",
  "- Headings and titles: place on their own line, preceded and followed by a blank line.",
  "- Body paragraphs: separate with blank lines.",
  "- Tables: use tab characters (\\t) between columns, newlines between rows. Include the header row.",
  "- Lists: preserve markers (1., 2., •, -, etc.) and indentation.",
  "- Key-value pairs: Key: Value, one per line.",
  "",
  "LANGUAGE:",
  "- Arabic text: preserve every word exactly. Maintain the correct right-to-left character order within each word.",
  "- Latin text, numbers, punctuation, and symbols: preserve exactly as shown in the image.",
  "- Do not translate, transliterate, or summarize anything.",
  "",
  "OUTPUT:",
  "- Extracted text only. No commentary, no explanations, no markdown code fences, no reasoning traces.",
  "- Preserve the blank lines that reflect the document's logical paragraph and section structure.",
].join("\n");

// ---------------------------------------------------------------------------
// Progress file helpers
// ---------------------------------------------------------------------------

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function loadQwenProgress(
  progressPath: string,
  pdfPath: string,
  totalPages: number,
): Promise<QwenProgress> {
  if (await fileExists(progressPath)) {
    try {
      const raw = JSON.parse(await readFile(progressPath, "utf-8")) as unknown;
      const parsed = QwenProgressSchema.safeParse(raw);
      if (
        parsed.success &&
        parsed.data.pdfPath === pdfPath &&
        parsed.data.totalPages === totalPages
      ) {
        const completed = Object.keys(parsed.data.pages).length;
        if (completed > 0) {
          console.log(
            `[Qwen] Resuming: ${completed}/${totalPages} page(s) already done — loading from progress file.`,
          );
        }
        return parsed.data;
      }
      console.log("[Qwen] Progress file found but does not match this PDF — starting fresh.");
    } catch {
      console.log("[Qwen] Progress file unreadable — starting fresh.");
    }
  }
  return { version: 1, pdfPath, totalPages, pages: {} };
}

async function saveQwenProgress(
  progressPath: string,
  progress: QwenProgress,
): Promise<void> {
  await mkdir(dirname(progressPath), { recursive: true });
  await writeFile(progressPath, JSON.stringify(progress, null, 2), "utf-8");
}

// ---------------------------------------------------------------------------
// Single-pass Qwen extraction
// ---------------------------------------------------------------------------

async function extractPageWithQwen(
  client: OpenAI,
  modelId: string,
  pageNum: number,
  pageImage: { base64: string; mimeType: string },
  verbose: boolean,
): Promise<string> {
  if (verbose) {
    console.log(`[Qwen] Page ${pageNum}: sending to ${modelId}...`);
  }

  const response = await client.chat.completions.create({
    model: modelId,
    messages: [
      { role: "system", content: QWEN_OCR_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          {
            type: "image_url" as const,
            image_url: {
              url: `data:${pageImage.mimeType};base64,${pageImage.base64}`,
            },
          },
          {
            type: "text" as const,
            text: "Extract all text from this document page.",
          },
        ] satisfies ChatCompletionContentPart[],
      },
    ],
    max_tokens: 4096,
    temperature: 0,
  });

  const raw = response.choices[0]?.message.content ?? "";
  const text = stripThinking(raw);

  if (verbose) {
    console.log(`[Qwen] Page ${pageNum}: extracted ${text.length} chars.`);
  }

  return text;
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

export interface QwenConvertArgs {
  pdfPath: string;
  outputPath: string;
  progressPath: string;
  client: OpenAI;
  qwenModelId: string;
  verbose: boolean;
}

/**
 * Convert a PDF to a Word document using a single Qwen2.5-VL pass per page.
 * Progress is saved after every page and the conversion is fully resumable.
 */
export async function convertPdfToWordWithQwen(args: QwenConvertArgs): Promise<void> {
  const { pdfPath, outputPath, progressPath, client, qwenModelId, verbose } = args;

  // ── Step 1: count pages ───────────────────────────────────────────────────
  console.log("[Qwen] Step 1 — Counting pages...");
  const { numPages } = await extractPdfjsPageTextsRaw(pdfPath, verbose);
  console.log(`[Qwen]   ${numPages} page(s) found.`);
  console.log("");

  // ── Load progress ─────────────────────────────────────────────────────────
  const progress = await loadQwenProgress(progressPath, pdfPath, numPages);

  // ── Detect renderer ───────────────────────────────────────────────────────
  const renderer: Renderer | null = await detectRenderer();
  if (!renderer) {
    console.log("[Qwen] Warning: No PDF renderer found (pdftoppm / mutool).");
    console.log("         OCR pass will be skipped — pages will be blank.");
    console.log("         Install via: brew install poppler  (provides pdftoppm)");
    console.log("");
  }

  // ── Per-page extraction loop ──────────────────────────────────────────────
  const tmpDir = join(tmpdir(), `qwen-word-${randomUUID()}`);
  await mkdir(tmpDir, { recursive: true });

  const docStart = Date.now();

  try {
    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      const pageKey = String(pageNum);

      if (progress.pages[pageKey]) {
        if (verbose) {
          console.log(`[Qwen] Page ${pageNum}/${numPages} — already done, skipping.`);
        }
        continue;
      }

      const pageStart = Date.now();
      console.log(`[Qwen] Page ${pageNum}/${numPages}:`);

      // ── Render page ───────────────────────────────────────────────────────
      let pageImage: { base64: string; mimeType: "image/jpeg" } | null = null;
      if (renderer) {
        try {
          process.stdout.write(`       Rendering...`);
          pageImage = await renderPageToJpeg(pdfPath, pageNum, renderer, tmpDir, verbose);
          process.stdout.write(
            ` done (${Math.round((pageImage.base64.length * 0.75) / 1024)} KB)\n`,
          );
        } catch (renderErr) {
          process.stdout.write(` failed\n`);
          console.log(
            `       Render error: ${renderErr instanceof Error ? renderErr.message : String(renderErr)}`,
          );
        }
      }

      // ── Qwen extraction ───────────────────────────────────────────────────
      let text = "";
      if (pageImage) {
        try {
          const t0 = Date.now();
          process.stdout.write(`       Extracting with Qwen2.5-VL...`);
          text = await extractPageWithQwen(client, qwenModelId, pageNum, pageImage, verbose);
          const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
          process.stdout.write(` done (${text.length} chars, ${elapsed}s)\n`);
        } catch (extractErr) {
          process.stdout.write(` failed\n`);
          console.log(
            `       Extraction error: ${extractErr instanceof Error ? extractErr.message : String(extractErr)}`,
          );
        }
      } else {
        console.log(`       No renderer — page will be blank.`);
      }

      // ── Save progress ─────────────────────────────────────────────────────
      progress.pages[pageKey] = { text };
      await saveQwenProgress(progressPath, progress);

      const pageSec = ((Date.now() - pageStart) / 1000).toFixed(1);
      console.log(`       Page ${pageNum} done in ${pageSec}s`);
    }
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }

  console.log("");

  // ── Generate Word document ────────────────────────────────────────────────
  console.log("[Qwen] Generating Word document...");
  const docGenStart = Date.now();

  const orderedPages = Array.from({ length: numPages }, (_, i) => ({
    pageNum: i + 1,
    text: progress.pages[String(i + 1)]?.text ?? "",
  }));

  await generateWordDoc(orderedPages, outputPath, verbose);

  const totalSec = ((Date.now() - docStart) / 1000).toFixed(1);
  const docGenSec = ((Date.now() - docGenStart) / 1000).toFixed(1);

  console.log("");
  console.log("Done!");
  console.log(`  Word document:  ${outputPath}`);
  console.log(`  Progress file:  ${progressPath}`);
  console.log(`  (Delete the progress file to re-process from scratch next time.)`);
  console.log("");
  console.log(`  Timing:`);
  console.log(`    Document build:  ${docGenSec}s`);
  console.log(`    Total elapsed:   ${totalSec}s`);
}
