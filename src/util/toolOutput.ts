import type { Converted } from "../content/convert.js";

/**
 * A tool result with images the model should see, not read as base64: MCP
 * carries them as image content next to the JSON. Plain results need none of
 * this; the server serialises anything else as JSON text.
 */
export class ToolOutput {
  constructor(
    readonly data: unknown,
    readonly images: ReadonlyArray<{ data: string; mimeType: string }> = [],
  ) {}
}

/** A converted file as a tool result: images as images, text as text, else base64. */
export function fileOutput(meta: Record<string, unknown>, c: Converted): ToolOutput {
  switch (c.kind) {
    case "image":
      return new ToolOutput({ ...meta, mimeType: c.mimeType, shown: "as an image" }, [{ data: c.data, mimeType: c.mimeType }]);
    case "text":
      return new ToolOutput({ ...meta, mimeType: c.mimeType, ...(c.note ? { note: c.note } : {}), text: c.text });
    case "binary":
      return new ToolOutput({ ...meta, mimeType: c.mimeType, encoding: "base64", base64: c.base64 });
  }
}
