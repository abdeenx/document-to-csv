import type OpenAI from "openai";
import type {
  ChatCompletionMessageParam,
  ChatCompletionTool,
} from "openai/resources/chat/completions";
import {
  WriteCsvToolArgsSchema,
  type WriteCsvToolArgs,
  type CsvResult,
  type OcrResult,
} from "./types.js";

const CSV_TOOL: ChatCompletionTool = {
  type: "function",
  function: {
    name: "write_csv",
    description: [
      "Write the final CSV representation of the document.",
      "Call this exactly once when you have determined the optimal CSV structure.",
      "The CSV must faithfully represent the structure and layout of the original document,",
      "including all rows, columns, headers, and data values. Use RFC 4180 compliant CSV.",
      "Wrap field values containing commas, newlines, or quotes in double-quotes.",
      "Escape internal double-quotes by doubling them.",
    ].join(" "),
    parameters: {
      type: "object",
      properties: {
        csv_content: {
          type: "string",
          description:
            "The complete, valid CSV content including headers on the first row",
        },
        reasoning: {
          type: "string",
          description:
            "Brief explanation of the structural choices: column mapping, multi-value handling, merged cell resolution, etc.",
        },
      },
      required: ["csv_content", "reasoning"],
    },
  },
};

const SYSTEM_PROMPT = `You are a precise document structure analyst and CSV converter.

You will receive text extracted from a document (via OCR or PDF text layer), and optionally an image of that document.
Produce a well-formed CSV that faithfully represents the data — not UI chrome or navigation elements around it.

IF AN IMAGE IS PROVIDED — USE IT AS VISUAL GROUND TRUTH:
- The image is the authoritative source. Use it to determine the true column layout and the correct value for every cell.
- The extracted text may have column alignment errors — cross-reference every value's position with what you see in the image.
- For every column, verify that each row's value is semantically correct for that column's header (e.g. a date must go in a date column, a Yes/No in a toggle column). Correct any misalignment you detect.
- Toggle/switch controls: look at the image to determine ON (colored) vs OFF (grey). Output "Yes" for ON, "No" for OFF. Never leave a toggle cell blank.

IF NO IMAGE IS PROVIDED (text-only mode, e.g. PDF text extraction):
- Trust the structural layout: tabs (\\t) mark column boundaries within a row; newlines separate rows.
- For multi-page documents: the table header appears once; rows that continue on subsequent pages belong to the same table — do not repeat the header.
- Infer the expected data type for each column from its header name and verify that each cell's value makes sense for that type.

COLUMN STRUCTURE:
- Identify the true data columns from the table headers. The number of CSV columns must match the number of logical headers exactly.
- Multi-word headers joined with a space (e.g. "AUTO RENEW") are a single column — do not split them.
- UI action elements that appear in every row (e.g. "INFO", "EDIT", "SETUP", "DELETE", "VIEW") are not data columns and must not appear in the header or any row.

DATA ROWS:
- Each data row maps exactly to one record in the source document.
- Tooltip text, overlay banners, popup notifications, and modal dialogs are not data rows — skip them.
- Empty cells must still be represented as empty fields to keep column counts consistent.

CSV FORMATTING — CRITICAL:
- Every field value containing a comma, a double-quote, or a newline MUST be wrapped in double-quotes.
- Example: "Apr 29, 2028" contains a comma → write it as "Apr 29, 2028" in the CSV.
- Escape internal double-quotes by doubling them.

QUALITY CHECKS (run before calling write_csv):
1. Every row has exactly the same number of fields as the header row.
2. No row contains UI button labels (INFO, EDIT, SETUP, etc.) as field values.
3. Any field containing a comma is wrapped in double-quotes.
4. No tooltip or overlay text appears as a row.
5. Toggle columns have "Yes" or "No" in every row — never blank.
6. Date values appear in date columns, not in toggle/boolean columns.

GENERAL:
- For forms/invoices with key-value pairs: use "Field" and "Value" columns.
- For multi-section documents: add a "Section" column.

Think step-by-step:
1. If an image is provided, examine it first to understand the true column layout and visual values (toggles, checkboxes, etc.); otherwise use the tab-delimited text layout.
2. Determine the exact header row, ignoring UI action columns (INFO, EDIT, etc.).
3. For each data row, assign values to columns correctly — using the image when available, or semantic inference in text-only mode.
4. Run the quality checks above.
5. Call write_csv exactly once with the final result.`;

// ---------------------------------------------------------------------------
// RFC 4180 CSV sanitizer
// Parses the LLM's CSV output and re-serializes it with guaranteed correct
// quoting and consistent column counts. This catches any case where the model
// forgot to quote a field that contains a comma (e.g. "Apr 29, 2028").
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

function serializeCsvField(value: string): string {
  if (
    value.includes(",") ||
    value.includes('"') ||
    value.includes("\n") ||
    value.includes("\r")
  ) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}

function sanitizeCsv(csvContent: string): string {
  const lines = csvContent.split("\n").filter((l) => l.trim() !== "");
  if (lines.length === 0) return csvContent;

  const rows = lines.map(parseCsvRow);
  const headerCount = rows[0]!.length;

  const serialized = rows.map((row) => {
    // Pad short rows so every row has exactly headerCount fields
    while (row.length < headerCount) row.push("");
    return row
      .slice(0, headerCount)
      .map(serializeCsvField)
      .join(",");
  });

  return serialized.join("\n") + "\n";
}

// ---------------------------------------------------------------------------

function resolveToolCall(
  toolCall: { type: string; id: string; function?: { name: string; arguments: string } },
  verbose: boolean,
): WriteCsvToolArgs | null {
  if (toolCall.type !== "function" || !toolCall.function) {
    if (verbose) {
      console.log(`[Gemma4] Non-function tool call type "${toolCall.type}" — ignoring`);
    }
    return null;
  }

  if (toolCall.function.name !== "write_csv") {
    if (verbose) {
      console.log(`[Gemma4] Unknown tool call: ${toolCall.function.name} — ignoring`);
    }
    return null;
  }

  let rawArgs: unknown;
  try {
    rawArgs = JSON.parse(toolCall.function.arguments);
  } catch {
    throw new Error(
      `[Gemma4] Failed to parse tool call arguments: ${toolCall.function.arguments}`,
    );
  }

  const parsed = WriteCsvToolArgsSchema.safeParse(rawArgs);
  if (!parsed.success) {
    throw new Error(
      `[Gemma4] Invalid write_csv arguments: ${parsed.error.message}`,
    );
  }

  return parsed.data;
}

export async function generateCsvWithGemma(
  client: OpenAI,
  ocrResult: OcrResult,
  modelId: string,
  outputPath: string,
  verbose: boolean,
): Promise<CsvResult> {
  if (verbose) {
    console.log(`[Gemma4] Starting tool-use loop with model: ${modelId}`);
  }

  const ocrTextBlock = [
    "Here is the OCR-extracted text from the document image.",
    "Use the image above as visual ground truth — correct any column misalignment you see in the OCR text.",
    "",
    "--- BEGIN OCR TEXT ---",
    ocrResult.rawText,
    "--- END OCR TEXT ---",
    "",
    "Analyze the image and the OCR text together, then call write_csv with the final CSV content.",
  ].join("\n");

  // Build the first user message. If we have the preprocessed image, send it
  // alongside the OCR text so Gemma4 can use visual reasoning to verify column
  // alignment rather than blindly trusting the OCR's positional output.
  const firstUserContent: ChatCompletionMessageParam["content"] =
    ocrResult.imageBase64 && ocrResult.imageMimeType
      ? [
          {
            type: "image_url" as const,
            image_url: {
              url: `data:${ocrResult.imageMimeType};base64,${ocrResult.imageBase64}`,
            },
          },
          { type: "text" as const, text: ocrTextBlock },
        ]
      : ocrTextBlock;

  if (verbose) {
    console.log(
      ocrResult.imageBase64
        ? "[Gemma4] Sending image + OCR text for visual column-alignment verification"
        : "[Gemma4] No image available — sending OCR text only",
    );
  }

  const messages: ChatCompletionMessageParam[] = [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "user", content: firstUserContent },
  ];

  const MAX_ITERATIONS = 8;
  let iterations = 0;

  while (iterations < MAX_ITERATIONS) {
    iterations++;

    if (verbose) {
      console.log(`[Gemma4] Iteration ${iterations}/${MAX_ITERATIONS}...`);
    }

    const response = await client.chat.completions.create({
      model: modelId,
      messages,
      tools: [CSV_TOOL],
      tool_choice: iterations === 1 ? "auto" : "required",
      max_tokens: 8192,
      temperature: 0.1,
    });

    const choice = response.choices[0];
    if (!choice) {
      throw new Error("[Gemma4] No response choice returned from model.");
    }

    const { message } = choice;
    messages.push(message as ChatCompletionMessageParam);

    if (verbose && message.content) {
      console.log(`[Gemma4] Model reasoning: ${message.content.slice(0, 200)}${message.content.length > 200 ? "..." : ""}`);
    }

    if (choice.finish_reason === "tool_calls" && message.tool_calls?.length) {
      for (const toolCall of message.tool_calls) {
        if (verbose) {
          const toolName = toolCall.type === "function" ? toolCall.function.name : toolCall.type;
          console.log(`[Gemma4] Tool call: ${toolName}`);
        }

        const result = resolveToolCall(toolCall, verbose);

        messages.push({
          role: "tool",
          tool_call_id: toolCall.id,
          content: result
            ? JSON.stringify({ success: true, message: "CSV captured successfully." })
            : JSON.stringify({ success: false, message: "Unknown tool." }),
        });

        if (result) {
          if (verbose) {
            console.log(`[Gemma4] Reasoning: ${result.reasoning}`);
          }
          // Programmatically sanitize the LLM's output for guaranteed RFC 4180
          // compliance — this fixes unquoted commas in field values (e.g. dates)
          // and ensures every row has the same column count as the header.
          const sanitized = sanitizeCsv(result.csv_content);
          return {
            csvContent: sanitized,
            outputPath,
            reasoning: result.reasoning,
          };
        }
      }
      continue;
    }

    if (choice.finish_reason === "stop") {
      if (verbose) {
        console.log("[Gemma4] Model stopped without calling write_csv — prompting again...");
      }
      messages.push({
        role: "user",
        content:
          "You must call the write_csv tool to provide the final CSV. Please do so now.",
      });
      continue;
    }

    throw new Error(`[Gemma4] Unexpected finish_reason: ${choice.finish_reason}`);
  }

  throw new Error(
    `[Gemma4] Exceeded maximum iterations (${MAX_ITERATIONS}) without a write_csv call.`,
  );
}
