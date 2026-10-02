import { describe, it, expect } from "vitest";
import { teamsTools } from "../src/tools/teams/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = (n: string) => findTool(teamsTools, n);
const directory = (people: Array<{ id: string; displayName: string } | null>) => ({
  responses: people.map((p, i) =>
    p ? { id: String(i), status: 200, body: p } : { id: String(i), status: 404, body: { error: { message: "Resource not found" } } },
  ),
});

describe("mentions", () => {
  it("turns @Name in the text into a real mention, escaping the rest", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", directory([{ id: "u-ana", displayName: "Ana Ruiz" }]));
    await callTool(tool("teams_post_channel_message"), ctx, { teamId: "t", channelId: "c", body: "Thanks @Ana Ruiz, <3 & more", mentions: ["ana@x.com"] });
    expect(ctx.mock.calls[1]!.body).toEqual({
      body: { contentType: "html", content: 'Thanks <at id="0">Ana Ruiz</at>, &lt;3 &amp; more' },
      mentions: [{ id: 0, mentionText: "Ana Ruiz", mentioned: { user: { id: "u-ana", displayName: "Ana Ruiz", userIdentityType: "aadUser" } } }],
    });
  });

  it("puts a mention not written in the text at the start", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", directory([{ id: "u-bob", displayName: "Bob" }]));
    await callTool(tool("teams_post_chat_message"), ctx, { chatId: "ch", body: "please review", mentions: ["bob@x.com"] });
    expect((ctx.mock.calls[1]!.body as { body: { content: string } }).body.content).toBe('<at id="0">Bob</at> please review');
  });

  it("refuses to post when someone to mention is not found", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", directory([null]));
    await expect(callTool(tool("teams_post_chat_message"), ctx, { chatId: "ch", body: "hi", mentions: ["ghost@x.com"] })).rejects.toThrow(/Cannot mention ghost@x.com/);
    expect(ctx.mock.calls.some((c) => c.path.endsWith("/messages"))).toBe(false);
  });

  it("leaves a plain message as it was", async () => {
    const ctx = makeContext();
    await callTool(tool("teams_post_chat_message"), ctx, { chatId: "ch", body: "hi" });
    expect(ctx.mock.calls[0]!.body).toEqual({ body: { contentType: "text", content: "hi" } });
  });
});

describe("direct messages", () => {
  it("opens (or finds) the one-to-one chat, then sends", async () => {
    const ctx = makeContext();
    ctx.mock.on("/me", { id: "me-id" });
    ctx.mock.on("/chats", { id: "19:chat" });
    ctx.mock.on("/chats/19:chat/messages", { id: "m1" });
    const out = await callTool(tool("teams_send_direct_message"), ctx, { user: "o'brien@x.com", body: "Lunch?" });
    expect(ctx.mock.calls[1]).toMatchObject({
      method: "POST",
      path: "/chats",
      body: {
        chatType: "oneOnOne",
        members: [
          expect.objectContaining({ "user@odata.bind": "https://graph.microsoft.com/v1.0/users('me-id')" }),
          expect.objectContaining({ "user@odata.bind": "https://graph.microsoft.com/v1.0/users('o''brien@x.com')" }),
        ],
      },
    });
    expect(out).toMatchObject({ chatId: "19:chat", messageId: "m1" });
  });

  it("edits and deletes a message of yours", async () => {
    const ctx = makeContext();
    await callTool(tool("teams_update_chat_message"), ctx, { chatId: "c", messageId: "m", body: "fixed" });
    await callTool(tool("teams_delete_chat_message"), ctx, { chatId: "c", messageId: "m" });
    expect(ctx.mock.calls.map((x) => `${x.method} ${x.path}`)).toEqual(["PATCH /chats/c/messages/m", "POST /me/chats/c/messages/m/softDelete"]);
  });
});
