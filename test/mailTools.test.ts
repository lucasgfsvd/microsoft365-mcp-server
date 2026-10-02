import { describe, it, expect } from "vitest";
import { listFilter, mailTools } from "../src/tools/mail/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

describe("mail_list_messages date filters", () => {
  const list = findTool(mailTools, "mail_list_messages");

  it("leads with receivedDateTime, which Graph requires when ordering by it", async () => {
    const ctx = makeContext();
    await callTool(list, ctx, { receivedAfter: "2026-09-28", receivedBefore: "2026-10-05T00:00:00+02:00", unreadOnly: true });
    expect(ctx.mock.calls[0]!.query.filter).toBe(
      "receivedDateTime ge 2026-09-28T00:00:00.000Z and receivedDateTime lt 2026-10-04T22:00:00.000Z and isRead eq false",
    );
    expect(ctx.mock.calls[0]!.query.orderby).toBe("receivedDateTime desc");
  });

  it("filters on one bound alone, or not at all", () => {
    expect(listFilter({ receivedAfter: "2026-09-28T08:00:00Z" })).toBe("receivedDateTime ge 2026-09-28T08:00:00.000Z");
    expect(listFilter({ unreadOnly: true })).toBe("isRead eq false");
    expect(listFilter({})).toBeUndefined();
  });

  it("refuses dates it cannot read, and an empty window", async () => {
    const ctx = makeContext();
    await expect(callTool(list, ctx, { receivedAfter: "last week" })).rejects.toThrow();
    await expect(callTool(list, ctx, { receivedAfter: "2026-09-28T08:00:00" })).rejects.toThrow(); // no offset: ambiguous
    await expect(callTool(list, ctx, { receivedAfter: "2026-10-05", receivedBefore: "2026-09-28" })).rejects.toThrow(/earlier/);
    expect(ctx.mock.calls).toHaveLength(0);
  });
});

describe("mail_create_reply_draft", () => {
  const tool = findTool(mailTools, "mail_create_reply_draft");
  const draft = {
    id: "draft-1",
    subject: "RE: Q3 numbers",
    toRecipients: [{ emailAddress: { address: "alice@example.com" } }],
    conversationId: "conv-1",
    isDraft: true,
    webLink: "https://outlook.office365.com/owa/?ItemID=draft-1",
    body: { contentType: "html", content: "<p>Sounds good</p><hr>…the whole quoted thread…" },
  };

  it("asks Outlook for the reply, so it stays in the thread", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    await callTool(tool, ctx, { id: "msg-1", comment: "Sounds good" });
    expect(ctx.mock.calls[0]).toMatchObject({ method: "POST", path: "/me/messages/msg-1/createReply", body: { comment: "Sounds good" } });
  });

  it("uses createReplyAll for reply-all", async () => {
    const ctx = makeContext();
    await callTool(tool, ctx, { id: "msg-1", comment: "Thanks all", replyAll: true });
    expect(ctx.mock.calls[0]!.path).toBe("/me/messages/msg-1/createReplyAll");
  });

  it("returns the draft without the quoted thread", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    const out = (await callTool(tool, ctx, { id: "msg-1", comment: "Sounds good" })) as Record<string, unknown>;
    expect(out).toEqual({
      id: "draft-1",
      subject: "RE: Q3 numbers",
      toRecipients: draft.toRecipients,
      conversationId: "conv-1",
      isDraft: true,
      webLink: draft.webLink,
    });
  });

  it("is a write, so it is hidden unless writes are enabled", () => {
    expect(tool.mutating).toBe(true);
    expect(tool.requiredScopes).toEqual(["Mail.ReadWrite"]);
  });
});
