import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { getPageHtml } from "../src/graph/onenote.js";

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
