import { tidyText } from "../util/text.js";
import { htmlTitle, htmlToText } from "./html.js";
import { enc, uriFor, values, type ResourceKind } from "./types.js";

export const onenoteResource: ResourceKind = {
  key: "onenote",
  requiresTool: "onenote_get_page_content",
  template: {
    uriTemplate: "m365://onenote/{pageId}",
    name: "onenote-page",
    title: "OneNote page",
    description: "A OneNote page as plain text.",
    mimeType: "text/plain",
  },
  recent: {
    id: "onenote",
    url: "/me/onenote/pages?$top=20&$select=id,title,lastModifiedDateTime&$orderby=lastModifiedDateTime%20desc",
  },
  toEntries: (body) =>
    values(body).map((p) => ({
      uri: uriFor("onenote", String(p.id)),
      name: `onenote: ${(p.title as string) || "(untitled page)"}`,
      description: `Modified ${p.lastModifiedDateTime ?? ""}`.trim(),
      mimeType: "text/plain",
    })),
  async read(graph, [id], uri) {
    const raw: unknown = await graph.api(`/me/onenote/pages/${enc(id!)}/content`).get();
    const html = typeof raw === "string" ? raw : String(raw);
    const title = htmlTitle(html);
    const text = tidyText(htmlToText(html));
    return { uri, mimeType: "text/plain", text: title && !text.startsWith(title) ? `${title}\n\n${text}` : text };
  },
};
