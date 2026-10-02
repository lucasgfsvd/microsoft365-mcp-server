import { describe, it, expect } from "vitest";
import { replyDraftTools, withReplyText } from "../src/tools/mail/replyDraft.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = findTool(replyDraftTools, "mail_create_reply_draft");
const htmlBody = { contentType: "html", content: '<html><head></head><body dir="ltr"><hr><div id="divRplyFwdMsg">quoted</div></body></html>' };
const draft = {
  id: "draft-1",
  subject: "RE: Q3 numbers",
  toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
  conversationId: "conv-1",
  isDraft: true,
  webLink: "https://outlook.office365.com/owa/?ItemID=draft-1",
  body: htmlBody,
};

describe("withReplyText", () => {
  // Found live: Graph's own `comment` reads as HTML, so "a\nb" came out as "a b".
  it("writes plain text into an HTML body, line breaks kept and markup escaped", () => {
    expect(withReplyText(htmlBody, "Sounds good.\nSee <you> & Bob").content).toBe(
      '<html><head></head><body dir="ltr"><div>Sounds good.<br>See &lt;you&gt; &amp; Bob</div><hr><div id="divRplyFwdMsg">quoted</div></body></html>',
    );
  });

  it("puts the text above a plain-text quote", () => {
    expect(withReplyText({ contentType: "text", content: "> quoted" }, "Thanks")).toEqual({ contentType: "text", content: "Thanks\n\n> quoted" });
  });

  it("copes with HTML that has no body tag", () => {
    expect(withReplyText({ contentType: "HTML", content: "<p>quoted</p>" }, "Hi").content).toBe("<div>Hi</div><p>quoted</p>");
  });
});

describe("mail_create_reply_draft", () => {
  it("asks Outlook for the reply, so it stays in the thread, then writes the text into it", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    ctx.mock.on("/me/messages/draft-1", (call) => ({ ...draft, body: (call.body as { body: unknown }).body }));
    await callTool(tool, ctx, { id: "msg-1", comment: "Sounds good" });
    expect(ctx.mock.calls.map((c) => [c.method, c.path])).toEqual([
      ["POST", "/me/messages/msg-1/createReply"],
      ["PATCH", "/me/messages/draft-1"],
    ]);
    expect(ctx.mock.calls[1]!.body).toEqual({ body: withReplyText(htmlBody, "Sounds good") });
  });

  it("uses createReplyAll for reply-all", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReplyAll", draft);
    await callTool(tool, ctx, { id: "msg-1", comment: "Thanks all", replyAll: true });
    expect(ctx.mock.calls[0]!.path).toBe("/me/messages/msg-1/createReplyAll");
  });

  it("returns the draft without the quoted thread", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    ctx.mock.on("/me/messages/draft-1", draft);
    const out = await callTool(tool, ctx, { id: "msg-1", comment: "Sounds good" });
    expect(out).toEqual({ ...draft, body: undefined });
    expect(out).not.toHaveProperty("body");
  });

  it("removes the empty draft when the text cannot be written", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    ctx.mock.on("/me/messages/draft-1", () => {
      throw new Error("ErrorItemNotFound");
    });
    await expect(callTool(tool, ctx, { id: "msg-1", comment: "Sounds good" })).rejects.toThrow(/ErrorItemNotFound/);
    expect(ctx.mock.calls.at(-1)).toMatchObject({ method: "DELETE", path: "/me/messages/draft-1" });
  });

  it("is a write, so it is hidden unless writes are enabled", () => {
    expect(tool.mutating).toBe(true);
    expect(tool.requiredScopes).toEqual(["Mail.ReadWrite"]);
  });
});
