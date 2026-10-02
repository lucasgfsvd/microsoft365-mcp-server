import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { MessageBody, messageBody } from "./mentions.js";

const member = (userRef: string) => ({
  "@odata.type": "#microsoft.graph.aadUserConversationMember",
  roles: ["owner"],
  "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userRef.replace(/'/g, "''")}')`,
});

/** Chats with people: starting one, and editing or deleting what you said. */
export const chatTools: ToolDefinition[] = [
  {
    name: "teams_send_direct_message",
    surface: "teams",
    description:
      "Send a Teams message to one person, by email, in your one-to-one chat with them (started if you have " +
      "never chatted). Returns the chat id for follow-ups.",
    mutating: true,
    requiredScopes: ["Chat.ReadWrite", "ChatMessage.Send"],
    inputSchema: z.object({ user: z.string().min(3).describe("Their email address or directory id."), ...MessageBody }),
    handler: async (input, ctx) => {
      const me = (await ctx.graph.api(`/me`).select("id").get()) as { id: string };
      // Creating a one-to-one chat that exists returns the existing one.
      const chat = (await ctx.graph.api(`/chats`).post({ chatType: "oneOnOne", members: [member(me.id), member(input.user)] })) as { id: string };
      const msg = (await ctx.graph.api(`/chats/${chat.id}/messages`).post(await messageBody(ctx.graph, input))) as { id: string; createdDateTime?: string };
      return { chatId: chat.id, messageId: msg.id, createdDateTime: msg.createdDateTime };
    },
  },
  {
    name: "teams_update_chat_message",
    surface: "teams",
    description: "Edit a message you sent in a chat.",
    mutating: true,
    requiredScopes: ["Chat.ReadWrite"],
    inputSchema: z.object({ chatId: z.string(), messageId: z.string(), body: z.string(), contentType: z.enum(["text", "html"]).default("text") }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/chats/${input.chatId}/messages/${input.messageId}`).patch({ body: { contentType: input.contentType, content: input.body } });
      return { ok: true };
    },
  },
  {
    name: "teams_delete_chat_message",
    surface: "teams",
    description: "Delete a message you sent in a chat; Teams shows it as deleted.",
    mutating: true,
    requiredScopes: ["Chat.ReadWrite"],
    inputSchema: z.object({ chatId: z.string(), messageId: z.string() }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/me/chats/${input.chatId}/messages/${input.messageId}/softDelete`).post({});
      return { ok: true };
    },
  },
];
