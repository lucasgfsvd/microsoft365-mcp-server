import { z } from "zod";
import type { ToolDefinition } from "../../types.js";

/** Sorting mail the way a person does in Outlook: filing it, marking it, flagging it. */
export const organizeTools: ToolDefinition[] = [
  {
    name: "mail_move_message",
    surface: "mail",
    description:
      "Move a message to another folder: a well-known name (archive, inbox, junkemail, deleteditems, " +
      "drafts, sentitems) or a folder id from mail_list_folders. The message gets a new id in its new folder.",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z.object({
      id: z.string(),
      destination: z.string().min(1).describe("Folder id, or a well-known name such as 'archive'."),
    }),
    handler: async ({ id, destination }, ctx) => {
      const moved = (await ctx.graph.api(`/me/messages/${id}/move`).post({ destinationId: destination })) as Record<string, unknown>;
      return { id: moved.id, parentFolderId: moved.parentFolderId, subject: moved.subject };
    },
  },
  {
    name: "mail_update_message",
    surface: "mail",
    description: "Mark a message read or unread, flag it (or mark the flag complete), set its categories or importance.",
    mutating: true,
    requiredScopes: ["Mail.ReadWrite"],
    inputSchema: z
      .object({
        id: z.string(),
        isRead: z.boolean().optional(),
        flag: z.enum(["flagged", "complete", "notFlagged"]).optional(),
        categories: z.array(z.string()).optional().describe("Replaces the message's categories; [] clears them."),
        importance: z.enum(["low", "normal", "high"]).optional(),
      })
      .refine((m) => [m.isRead, m.flag, m.categories, m.importance].some((x) => x !== undefined), {
        message: "Give at least one of isRead, flag, categories or importance",
      }),
    handler: async (input, ctx) => {
      const patch: Record<string, unknown> = {};
      if (input.isRead !== undefined) patch.isRead = input.isRead;
      if (input.flag) patch.flag = { flagStatus: input.flag };
      if (input.categories) patch.categories = input.categories; // [] clears them
      if (input.importance) patch.importance = input.importance;
      const m = ((await ctx.graph.api(`/me/messages/${input.id}`).patch(patch)) ?? {}) as Record<string, unknown>;
      return { id: input.id, isRead: m.isRead, flag: m.flag, categories: m.categories, importance: m.importance };
    },
  },
];
