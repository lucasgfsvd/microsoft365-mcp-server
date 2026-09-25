import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { Document, Packer, Paragraph } from "docx";
import { availableKinds, listRecent, RESOURCE_KINDS } from "../src/resources/index.js";
import { parseUri, uriFor } from "../src/resources/types.js";
import { htmlTitle, htmlToText } from "../src/resources/html.js";
import { tidyText } from "../src/util/text.js";
import { formatMessage } from "../src/resources/mail.js";
import { driveResource, fileContent } from "../src/resources/drive.js";
import { buildDeck } from "../src/ooxml/newDeck.js";

describe("resource URIs", () => {
  it("round-trips ids containing characters Graph ids really use", () => {
    const uri = uriFor("drive", "b!Pp_KJ6/U4==", "01APU2L5!YTM");
    expect(uri).toBe("m365://drive/b!Pp_KJ6%2FU4%3D%3D/01APU2L5!YTM");
    expect(parseUri(uri)).toEqual({ key: "drive", segments: ["b!Pp_KJ6/U4==", "01APU2L5!YTM"] });
  });

  it("rejects anything malformed", () => {
    for (const bad of ["m365://mail/", "m365://mail", "https://x/mail/1", "m365://mail//1", "m365://mail/%E0%A4%A"]) {
      expect(parseUri(bad)).toBeUndefined();
    }
  });
});

describe("text conversion", () => {
  it("turns OneNote HTML into readable text", () => {
    const html =
      "<html><head><title>Plan &amp; notes</title><style>p{}</style></head><body>" +
      "<h1>Goals</h1><p>First&nbsp;line<br/>second</p><ul><li>one</li><li>two &#8364;5</li></ul></body></html>";
    expect(htmlTitle(html)).toBe("Plan & notes");
    expect(htmlToText(html)).toBe("Goals\nFirst line\nsecond\n\n- one\n- two €5");
  });

  // Found live: OneNote serves indented HTML with CRLF line endings, whose
  // source whitespace leaked into the text as \r and blank lines.
  it("ignores source whitespace, as a browser does", () => {
    const html = "<body>\r\n  <h1>Heading</h1>\r\n  <p>First\r\n  paragraph</p>\r\n  <ul>\r\n    <li>item one</li>\r\n    <li>item two</li>\r\n  </ul>\r\n</body>";
    expect(tidyText(htmlToText(html))).toBe("Heading\nFirst paragraph\n\n- item one\n- item two");
  });

  it("normalises CRLF in plain text too (Outlook text bodies use it)", () => {
    expect(tidyText("line one\r\nline two\r\n\r\n\r\nend")).toBe("line one\nline two\n\nend");
  });

  // Found live: 41% of a real newsletter's text body was this padding.
  it("drops invisible preview-line padding and collapses whitespace", () => {
    const padded = `Hello‌ ͏­ ‌   world\n\n\n\n  next  `;
    expect(tidyText(padded)).toBe("Hello world\n\nnext");
  });

  it("formats a message the way one would paste it", () => {
    const text = formatMessage({
      subject: "Q3",
      from: { emailAddress: { name: "Ada", address: "ada@example.com" } },
      toRecipients: [{ emailAddress: { address: "me@example.com" } }],
      receivedDateTime: "2026-09-25T10:00:00Z",
      hasAttachments: true,
      body: { content: "Hi‌‌ there" },
    });
    expect(text).toBe(
      "Subject: Q3\nFrom: Ada <ada@example.com>\nTo: me@example.com\nDate: 2026-09-25T10:00:00Z\n" +
        "Attachments: yes (see mail_list_attachments)\n\nHi there\n",
    );
  });
});

describe("file content", () => {
  it("returns Word and PowerPoint as their text, not base64 of a zip", async () => {
    const docx = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Hello doc")] }] }));
    expect(fileContent("u", docx, undefined, "a.docx")).toMatchObject({ mimeType: "text/plain", text: expect.stringContaining("Hello doc") });
    const pptx = await buildDeck([{ title: "Deck title", bullets: ["b"] }]);
    expect(fileContent("u", pptx, undefined, "a.pptx")).toMatchObject({ text: "--- Slide 1 ---\nDeck title b" });
  });

  it("keeps binaries as blobs and text as text", () => {
    expect(fileContent("u", Buffer.from("a,b"), "text/csv", "x.csv")).toEqual({ uri: "u", mimeType: "text/csv", text: "a,b" });
    expect(fileContent("u", Buffer.from([0, 1]), "application/pdf", "x.pdf")).toEqual({ uri: "u", mimeType: "application/pdf", blob: "AAE=" });
  });

  it("lists files shared from other drives under their own drive, and skips folders", () => {
    const entries = driveResource.toEntries({
      value: [
        { id: "1", name: "mine.docx", file: {}, parentReference: { driveId: "d1" } },
        { id: "2", name: "shared.xlsx", remoteItem: { id: "r2", name: "shared.xlsx", file: {}, parentReference: { driveId: "d2" } } },
        { id: "3", name: "Folder", folder: {}, parentReference: { driveId: "d1" } },
      ],
    });
    expect(entries.map((e) => e.uri)).toEqual(["m365://drive/d1/1", "m365://drive/d2/r2"]);
  });
});

describe("availability", () => {
  it("offers a kind only when its read tool is visible", () => {
    expect(availableKinds(new Set(["mail_get_message"])).map((k) => k.key)).toEqual(["mail"]);
    expect(availableKinds(new Set()).length).toBe(0);
    expect(availableKinds(new Set(RESOURCE_KINDS.map((k) => k.requiresTool))).length).toBe(3);
  });

  it("lists every kind in one batch, leaving out a kind whose listing fails", async () => {
    const posted: unknown[] = [];
    const graph = {
      api: (path: string) => ({
        post: async (body: { requests: Array<{ id: string }> }) => {
          posted.push({ path, n: body.requests.length });
          return {
            responses: [
              { id: "mail", status: 200, body: { value: [{ id: "m1", subject: "Hi" }] } },
              { id: "drive", status: 200, body: { value: [] } },
              { id: "onenote", status: 403, body: { error: { code: "Forbidden" } } },
            ],
          };
        },
      }),
    } as unknown as GraphClient;
    const entries = await listRecent(graph, RESOURCE_KINDS);
    expect(posted).toEqual([{ path: "/$batch", n: 3 }]);
    expect(entries.map((e) => e.uri)).toEqual(["m365://mail/m1"]);
  });
});
