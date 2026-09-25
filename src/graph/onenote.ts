import { ResponseType, type Client as GraphClient } from "@microsoft/microsoft-graph-client";

/**
 * A OneNote page's HTML. Graph serves it as text/html, which the SDK does not
 * turn into a string on its own: without ResponseType.TEXT, get() resolves to
 * a ReadableStream, and callers that stringified it returned
 * "[object ReadableStream]" (found live; mocks returning a string hid it).
 */
export async function getPageHtml(graph: GraphClient, pageId: string): Promise<string> {
  const body: unknown = await graph
    .api(`/me/onenote/pages/${encodeURIComponent(pageId)}/content`)
    .responseType(ResponseType.TEXT)
    .get();
  if (typeof body !== "string") throw new Error(`OneNote page content came back as ${typeof body}, not text.`);
  return body;
}
