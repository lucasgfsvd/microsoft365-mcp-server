import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

export const teamsTools: ToolDefinition[] = [
  {
    name: "teams_list_joined",
    surface: "teams",
    description: "List teams the user has joined.",
    requiredScopes: ["Team.ReadBasic.All"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/joinedTeams`, input),
  },
  {
    name: "teams_list_channels",
    surface: "teams",
    description: "List channels in a team.",
    requiredScopes: ["Channel.ReadBasic.All"],
    inputSchema: PaginationInput.extend({ teamId: z.string() }),
    handler: async (input, ctx) => fetchPage(ctx.graph, `/teams/${input.teamId}/channels`, input),
  },
  {
    name: "teams_list_channel_messages",
    surface: "teams",
    description: "List recent messages in a channel.",
    requiredScopes: ["ChannelMessage.Read.All"],
    inputSchema: PaginationInput.extend({ teamId: z.string(), channelId: z.string() }),
    handler: async (input, ctx) =>
      fetchPage(ctx.graph, `/teams/${input.teamId}/channels/${input.channelId}/messages`, input),
  },
  {
    name: "teams_get_message_replies",
    surface: "teams",
    description: "List replies on a channel message.",
    requiredScopes: ["ChannelMessage.Read.All"],
    inputSchema: PaginationInput.extend({
      teamId: z.string(),
      channelId: z.string(),
      messageId: z.string(),
    }),
    handler: async (input, ctx) =>
      fetchPage(
        ctx.graph,
        `/teams/${input.teamId}/channels/${input.channelId}/messages/${input.messageId}/replies`,
        input,
      ),
  },
  {
    name: "teams_list_chats",
    surface: "teams",
    description: "List the user's 1:1 and group chats.",
    requiredScopes: ["Chat.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/chats`, input),
  },
  {
    name: "teams_list_chat_messages",
    surface: "teams",
    description: "List messages in a chat.",
    requiredScopes: ["Chat.Read"],
    inputSchema: PaginationInput.extend({ chatId: z.string() }),
    handler: async (input, ctx) => fetchPage(ctx.graph, `/chats/${input.chatId}/messages`, input),
  },
  {
    name: "teams_post_channel_message",
    surface: "teams",
    description: "Post a message to a channel.",
    mutating: true,
    requiredScopes: ["ChannelMessage.Send"],
    inputSchema: z.object({
      teamId: z.string(),
      channelId: z.string(),
      body: z.string(),
      contentType: z.enum(["text", "html"]).default("text"),
    }),
    handler: async (input, ctx) =>
      ctx.graph
        .api(`/teams/${input.teamId}/channels/${input.channelId}/messages`)
        .post({ body: { contentType: input.contentType, content: input.body } }),
  },
  {
    name: "teams_reply_channel_message",
    surface: "teams",
    description: "Reply to a channel message.",
    mutating: true,
    requiredScopes: ["ChannelMessage.Send"],
    inputSchema: z.object({
      teamId: z.string(),
      channelId: z.string(),
      messageId: z.string(),
      body: z.string(),
      contentType: z.enum(["text", "html"]).default("text"),
    }),
    handler: async (input, ctx) =>
      ctx.graph
        .api(`/teams/${input.teamId}/channels/${input.channelId}/messages/${input.messageId}/replies`)
        .post({ body: { contentType: input.contentType, content: input.body } }),
  },
  {
    name: "teams_post_chat_message",
    surface: "teams",
    description: "Send a message to a chat.",
    mutating: true,
    requiredScopes: ["ChatMessage.Send"],
    inputSchema: z.object({
      chatId: z.string(),
      body: z.string(),
      contentType: z.enum(["text", "html"]).default("text"),
    }),
    handler: async (input, ctx) =>
      ctx.graph
        .api(`/chats/${input.chatId}/messages`)
        .post({ body: { contentType: input.contentType, content: input.body } }),
  },
];
