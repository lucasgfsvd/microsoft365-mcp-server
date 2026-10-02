import { describe, it, expect, afterEach, vi } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { mailTools } from "../src/tools/mail/index.js";
import { SIMPLE_ATTACHMENT_LIMIT } from "../src/graph/mailAttachments.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = (name: string) => findTool(mailTools, name);
const steps = (ctx: ReturnType<typeof makeContext>) => ctx.mock.calls.map((c) => `${c.method} ${c.path}`);
const draft = { id: "d1", subject: "RE: x", body: { contentType: "html", content: "<html><body><hr>quoted</body></html>" } };

afterEach(() => vi.unstubAllGlobals());

describe("mail_send_message", () => {
  it("sends in one call when there is nothing to attach", async () => {
    const ctx = makeContext();
    await callTool(tool("mail_send_message"), ctx, { to: [{ address: "a@x.com" }], subject: "s", body: "b" });
    expect(steps(ctx)).toEqual(["POST /me/sendMail"]);
  });

  it("builds a draft, attaches, then sends, when there are attachments", async () => {
    const ctx = makeContext();
    ctx.mock.on("/me/messages", { id: "m1" });
    const out = await callTool(tool("mail_send_message"), ctx, {
      to: [{ address: "a@x.com" }], subject: "s", body: "b",
      attachments: [{ name: "notes.txt", contentBase64: Buffer.from("hello").toString("base64") }],
    });
    expect(steps(ctx)).toEqual(["POST /me/messages", "POST /me/messages/m1/attachments", "POST /me/messages/m1/send"]);
    expect(ctx.mock.calls[1]!.body).toMatchObject({ "@odata.type": "#microsoft.graph.fileAttachment", name: "notes.txt", contentBytes: "aGVsbG8=" });
    expect(out).toEqual({ ok: true, attached: [{ name: "notes.txt", size: 5 }] });
  });

  it("sends a large file from disk through an attachment upload session", async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), "m365-attach-"));
    try {
      const big = Buffer.alloc(SIMPLE_ATTACHMENT_LIMIT + 2 * 1024 * 1024, 7); // two chunks
      await fs.writeFile(path.join(dir, "deck.pdf"), big);
      const puts: number[] = [];
      vi.stubGlobal("fetch", async (_u: string, init: RequestInit) => {
        puts.push((init.body as Buffer).length);
        const sent = puts.reduce((a, b) => a + b, 0);
        // As Outlook answers (found live): 200 with the next range while more is
        // wanted, then 201 with no body. Taking that 200 as done lost the attachment.
        return sent === big.length
          ? new Response(null, { status: 201, headers: { Location: "att" } })
          : new Response(JSON.stringify({ nextExpectedRanges: [String(sent)] }), { status: 200 });
      });
      const ctx = makeContext({ uploadDir: dir });
      ctx.mock.on("/me/messages", { id: "m1" });
      ctx.mock.on("/createUploadSession", { uploadUrl: "https://outlook.office.com/upload" });
      await callTool(tool("mail_create_draft"), ctx, { to: [{ address: "a@x.com" }], subject: "s", body: "b", attachments: [{ localPath: "deck.pdf" }] });
      expect(ctx.mock.calls[1]).toMatchObject({ path: "/me/messages/m1/attachments/createUploadSession", body: { AttachmentItem: { attachmentType: "file", name: "deck.pdf", size: big.length } } });
      expect(puts.reduce((a, b) => a + b, 0)).toBe(big.length);
      expect(steps(ctx)).not.toContain("POST /me/messages/m1/send");
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });

  it("removes the draft, and sends nothing, when an attachment fails", async () => {
    const ctx = makeContext(); // no MCP_UPLOAD_DIR
    ctx.mock.on("/me/messages", { id: "m1" });
    await expect(callTool(tool("mail_send_message"), ctx, { to: [{ address: "a@x.com" }], subject: "s", body: "b", attachments: [{ localPath: "x.pdf" }] })).rejects.toThrow(/MCP_UPLOAD_DIR/);
    expect(steps(ctx)).toEqual(["POST /me/messages", "DELETE /me/messages/m1"]);
  });

  it("wants exactly one source per attachment", async () => {
    const ctx = makeContext();
    await expect(callTool(tool("mail_send_message"), ctx, { to: [{ address: "a@x.com" }], subject: "s", body: "b", attachments: [{ localPath: "a", driveItemId: "b" }] })).rejects.toThrow(/exactly one/);
    await expect(callTool(tool("mail_send_message"), ctx, { to: [{ address: "a@x.com" }], subject: "s", body: "b", attachments: [{ contentBase64: "aGk=" }] })).rejects.toThrow(/needs a name/);
  });
});

describe("replies and forwards", () => {
  it("reply: Outlook's reply, the text written in as typed, then sent", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    await callTool(tool("mail_reply_message"), ctx, { id: "m", comment: "Yes.\nThanks" });
    expect(steps(ctx)).toEqual(["POST /me/messages/m/createReply", "PATCH /me/messages/d1", "POST /me/messages/d1/send"]);
    expect(JSON.stringify(ctx.mock.calls[1]!.body)).toContain("Yes.<br>Thanks");
  });

  it("reply: a failed send leaves no draft behind", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createReply", draft);
    ctx.mock.on("/send", () => {
      throw new Error("ErrorSendAsDenied");
    });
    await expect(callTool(tool("mail_reply_message"), ctx, { id: "m", comment: "x" })).rejects.toThrow(/SendAsDenied/);
    expect(steps(ctx).at(-1)).toBe("DELETE /me/messages/d1");
  });

  it("forward: recipients and text set on Outlook's forward; draft: true stops before sending", async () => {
    const ctx = makeContext();
    ctx.mock.on("/createForward", { ...draft, id: "f1" });
    ctx.mock.on("/me/messages/f1", (c) => ({ id: "f1", ...(c.body as object) }));
    const out = await callTool(tool("mail_forward_message"), ctx, { id: "m", to: [{ address: "b@x.com" }], comment: "FYI", draft: true });
    expect(steps(ctx)).toEqual(["POST /me/messages/m/createForward", "PATCH /me/messages/f1"]);
    expect(ctx.mock.calls[1]!.body).toMatchObject({ toRecipients: [{ emailAddress: { address: "b@x.com" } }] });
    expect(out).toMatchObject({ id: "f1", toRecipients: [{ emailAddress: { address: "b@x.com" } }] });
  });

  it("send_draft sends what is in Drafts", async () => {
    const ctx = makeContext();
    expect(await callTool(tool("mail_send_draft"), ctx, { id: "d9" })).toEqual({ ok: true });
    expect(steps(ctx)).toEqual(["POST /me/messages/d9/send"]);
  });
});

describe("organising mail", () => {
  it("moves to a well-known folder, reporting the new id", async () => {
    const ctx = makeContext();
    ctx.mock.on("/move", { id: "new", parentFolderId: "arch", subject: "s" });
    expect(await callTool(tool("mail_move_message"), ctx, { id: "m", destination: "archive" })).toEqual({ id: "new", parentFolderId: "arch", subject: "s" });
    expect(ctx.mock.calls[0]!.body).toEqual({ destinationId: "archive" });
  });

  it("marks, flags and categorises in one update; [] clears categories", async () => {
    const ctx = makeContext();
    await callTool(tool("mail_update_message"), ctx, { id: "m", isRead: false, flag: "flagged", categories: [] });
    expect(ctx.mock.calls[0]!.body).toEqual({ isRead: false, flag: { flagStatus: "flagged" }, categories: [] });
    await expect(callTool(tool("mail_update_message"), ctx, { id: "m" })).rejects.toThrow(/at least one/);
  });
});

describe("mail_get_attachment", () => {
  const meta = (o: object) => ({ name: "a.txt", contentType: "text/plain", size: 5, ...o });

  it("reads a file attachment through the converter", async () => {
    const ctx = makeContext();
    ctx.mock.on("/attachments/a1", meta({ "@odata.type": "#microsoft.graph.fileAttachment" }));
    ctx.mock.on("/$value", Buffer.from("hello"));
    const out = (await callTool(tool("mail_get_attachment"), ctx, { messageId: "m", attachmentId: "a1" })) as { data: unknown };
    expect(out.data).toMatchObject({ name: "a.txt", text: "hello" });
  });

  it("gives an attached email as its text", async () => {
    const ctx = makeContext();
    ctx.mock.on("/attachments/a1", meta({ "@odata.type": "#microsoft.graph.itemAttachment", name: "Fwd" }));
    ctx.mock.on("/attachments/a1", { item: { subject: "Original", body: { contentType: "html", content: "<p>Hi <b>there</b></p>" } } });
    expect(await callTool(tool("mail_get_attachment"), ctx, { messageId: "m", attachmentId: "a1" })).toEqual({ name: "Fwd", item: { subject: "Original", body: "Hi there" } });
    expect(ctx.mock.calls[1]!.query.expand).toBe("microsoft.graph.itemattachment/item");
  });

  it("refuses a large one inline and points to saveToDisk", async () => {
    const ctx = makeContext();
    ctx.mock.on("/attachments/a1", meta({ "@odata.type": "#microsoft.graph.fileAttachment", size: 50 * 1024 * 1024 }));
    await expect(callTool(tool("mail_get_attachment"), ctx, { messageId: "m", attachmentId: "a1" })).rejects.toThrow(/saveToDisk/);
  });
});

describe("automatic replies", () => {
  it("turns them on for a period, messages written as HTML that reads the same", async () => {
    const ctx = makeContext();
    await callTool(tool("mail_set_automatic_replies"), ctx, {
      status: "scheduled", internalMessage: "Away <until> Monday\nAsk Bob",
      start: { dateTime: "2026-12-22T09:00:00", timeZone: "Europe/Madrid" }, end: { dateTime: "2027-01-07T09:00:00", timeZone: "Europe/Madrid" },
    });
    expect(ctx.mock.calls[0]).toMatchObject({
      method: "PATCH", path: "/me/mailboxSettings",
      body: { automaticRepliesSetting: { status: "scheduled", internalReplyMessage: "Away &lt;until&gt; Monday<br>Ask Bob" } },
    });
    await expect(callTool(tool("mail_set_automatic_replies"), ctx, { status: "scheduled" })).rejects.toThrow(/start and end/);
  });
});
