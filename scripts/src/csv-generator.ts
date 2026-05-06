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

You will receive OCR-extracted text from a document image. Produce a well-formed CSV that faithfully represents the data — not the UI chrome around it.

COLUMN STRUCTURE:
- Identify the true data columns from the table headers. The number of CSV columns must match the number of logical headers exactly.
- Multi-word headers that were joined with a space (e.g. "AUTO RENEW") are a single column — do not split them.
- UI action elements that appear in every row (e.g. "INFO", "EDIT", "SETUP", "DELETE", "VIEW") are not data columns and must not appear in the header or any row.

DATA ROWS:
- Each data row maps exactly to one record in the source document.
- Tooltip text, overlay banners, popup notifications, and modal dialogs visible in the OCR are not data rows — skip them.
- If a date value appears as "Apr 29, 2028" keep it as one cell value, not two separate columns.
- Toggle values ("Yes"/"No") belong in their column as-is.
- Empty cells must still be represented as empty fields (consecutive commas) to keep column counts consistent.

QUALITY CHECKS (run before calling write_csv):
1. Every row has exactly the same number of comma-separated fields as the header row.
2. No row contains UI button labels (INFO, EDIT, SETUP, etc.) as field values.
3. Dates are not split across two fields.
4. No tooltip or overlay text appears as a row.

GENERAL:
- For forms/invoices with key-value pairs: use "Field" and "Value" columns.
- For multi-section documents: add a "Section" column.
- Use RFC 4180 compliant CSV: quote any field containing a comma, newline, or double-quote; escape internal double-quotes by doubling them.

Think step-by-step:
1. Identify the document type and its true data columns (ignoring UI actions).
2. Determine the exact header row.
3. Map every data record to a row, filling empty cells with empty strings.
4. Run the quality checks above.
5. Call write_csv exactly once with the final result.`;

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

  const messages: ChatCompletionMessageParam[] = [
    { role: "system", content: SYSTEM_PROMPT },
    {
      role: "user",
      content: [
        "Here is the OCR-extracted text from the document image. Convert it to a CSV file.",
        "",
        "--- BEGIN OCR TEXT ---",
        ocrResult.rawText,
        "--- END OCR TEXT ---",
        "",
        "Analyze the structure, then call write_csv with the final CSV content.",
      ].join("\n"),
    },
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
          return {
            csvContent: result.csv_content,
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
