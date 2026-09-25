import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { mailTools, tidyMessage } from "../src/tools/mail/index.js";
import type { ToolContext } from "../src/types.js";

const PAD = "‌ ͏­ ‌ ";
const msg = () => ({
  id: "m1",
  subject: "News",
  bodyPreview: `Offer${PAD}${PAD}inside`,
  body: { contentType: "text", content: `Hello${PAD}world\n\n\n\nBye` },
});

// Found live: 41% of a newsletter's text body was invisible padding.
describe("mail_get_message tidying", () => {
  // The preview is tidied for every tool by serializeResult (see redact.test.ts).
  it("strips padding from the text body", () => {
    const out = tidyMessage(msg(), "text");
    expect(out.body?.content).toBe("Hello world\n\nBye");
    expect(out.subject).toBe("News");
  });

  it("leaves an HTML body as sent", () => {
    const html = { ...msg(), body: { contentType: "html", content: "<pre>a   b\n\n\n c</pre>" } };
    const out = tidyMessage(html, "html");
    expect(out.body.content).toBe("<pre>a   b\n\n\n c</pre>");
  });

  it("is applied by the tool, which still asks Outlook for the requested format", async () => {
    const headers: Record<string, string> = {};
    const graph = {
      api: () => ({
        header(k: string, v: string) {
          headers[k] = v;
          return this;
        },
        get: async () => msg(),
      }),
    } as unknown as GraphClient;
    const tool = mailTools.find((t) => t.name === "mail_get_message")!;
    const out = (await tool.handler({ id: "m1", bodyFormat: "text" }, { graph } as unknown as ToolContext)) as ReturnType<typeof msg>;
    expect(headers.Prefer).toBe('outlook.body-content-type="text"');
    expect(out.body.content).toBe("Hello world\n\nBye");
  });
});
