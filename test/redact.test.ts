import { describe, it, expect } from "vitest";
import { serializeResult } from "../src/util/redact.js";

// Regression: every listing handed the model pre-authenticated download links
// (tempauth tokens), usable by anyone for about an hour.
describe("serializeResult", () => {
  it("drops pre-authenticated download links at any depth, keeping everything else", () => {
    const listing = {
      value: [
        {
          id: "1",
          name: "a.docx",
          "@microsoft.graph.downloadUrl": "https://x/download.aspx?tempauth=SECRET",
          "@microsoft.graph.downloadUrlNoAuth": "https://x/download.aspx",
        },
      ],
      // graph_batch_get nests Graph bodies; delta may use the older key.
      responses: [{ body: { value: [{ id: "2", "@content.downloadUrl": "https://y?tempauth=SECRET" }] } }],
    };
    const out = serializeResult(listing);
    expect(out).not.toContain("SECRET");
    expect(JSON.parse(out)).toEqual({
      value: [{ id: "1", name: "a.docx", "@microsoft.graph.downloadUrlNoAuth": "https://x/download.aspx" }],
      responses: [{ body: { value: [{ id: "2" }] } }],
    });
  });

  // Found live: 228 of one newsletter's 255 preview characters were padding.
  it("tidies message previews wherever they appear, and nothing else", () => {
    const pad = "‌ ͏­ ‌ ";
    const listing = {
      value: [{ subject: `Keep${pad}this`, bodyPreview: `Offer${pad}${pad}inside` }],
      responses: [{ body: { value: [{ bodyPreview: `Nested${pad}preview` }] } }],
    };
    expect(JSON.parse(serializeResult(listing))).toEqual({
      value: [{ subject: `Keep${pad}this`, bodyPreview: "Offer inside" }],
      responses: [{ body: { value: [{ bodyPreview: "Nested preview" }] } }],
    });
  });

  it("serializes ordinary results unchanged", () => {
    expect(serializeResult({ a: [1, "b"], c: null })).toBe(JSON.stringify({ a: [1, "b"], c: null }, null, 2));
  });
});
