import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { getPageHtml } from "../src/graph/onenote.js";
import { onenoteTools } from "../src/tools/onenote/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

/** Behaves like the real SDK: text/html comes back as a stream unless text is requested. */
function sdkLike(html: string) {
  const seen: string[] = [];
  const graph = {
    api: (url: string) => {
      seen.push(url);
      let type: string | undefined;
      const req = {
        responseType(t: string) {
          type = t;
          return req;
        },
        get: async () => (type === "text" ? html : new ReadableStream()),
      };
      return req;
    },
  } as unknown as GraphClient;
  return { graph, seen };
}

// Regression, found live: onenote_get_page_content returned "[object ReadableStream]".
describe("getPageHtml", () => {
  it("asks for text, so the page arrives as a string", async () => {
    const { graph, seen } = sdkLike("<html><body><p>hi</p></body></html>");
    expect(await getPageHtml(graph, "1-abc!2")).toBe("<html><body><p>hi</p></body></html>");
    expect(seen).toEqual(["/me/onenote/pages/1-abc!2/content"]);
  });
});

describe("creating notebooks and sections", () => {
  const notebook = findTool(onenoteTools, "onenote_create_notebook");
  const section = findTool(onenoteTools, "onenote_create_section");

  it("creates a notebook, then a section inside it", async () => {
    const ctx = makeContext();
    await callTool(notebook, ctx, { displayName: "  Project notes " });
    await callTool(section, ctx, { notebookId: "1-nb", displayName: "Meetings" });
    expect(ctx.mock.calls).toEqual([
      expect.objectContaining({ method: "POST", path: "/me/onenote/notebooks", body: { displayName: "Project notes" } }),
      expect.objectContaining({ method: "POST", path: "/me/onenote/notebooks/1-nb/sections", body: { displayName: "Meetings" } }),
    ]);
  });

  it("refuses names OneNote would reject, before calling Graph", async () => {
    const ctx = makeContext();
    for (const bad of ["Q3/Q4", "R&D", "a\\b", "50%", "", "   "]) {
      await expect(callTool(notebook, ctx, { displayName: bad })).rejects.toThrow();
    }
    await expect(callTool(section, ctx, { notebookId: "1-nb", displayName: "x".repeat(51) })).rejects.toThrow();
    await expect(callTool(notebook, ctx, { displayName: "x".repeat(129) })).rejects.toThrow();
    expect(ctx.mock.calls).toHaveLength(0);
  });
});
