import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { textToHtml } from "../../util/html.js";

const When = z.object({
  dateTime: z.string().describe("Local date-time, e.g. 2026-12-22T09:00:00"),
  timeZone: z.string().default("UTC").describe("IANA or Windows time zone, e.g. Europe/Madrid"),
});

/** Out-of-office: the automatic replies Outlook sends while the user is away. */
export const mailSettingsTools: ToolDefinition[] = [
  {
    name: "mail_get_automatic_replies",
    surface: "mail",
    description: "Show the out-of-office (automatic replies) setting: whether it is on, when, and the messages.",
    requiredScopes: ["MailboxSettings.Read"],
    inputSchema: z.object({}),
    handler: async (_input, ctx) => ctx.graph.api(`/me/mailboxSettings/automaticRepliesSetting`).get(),
  },
  {
    name: "mail_set_automatic_replies",
    surface: "mail",
    description:
      "Turn out-of-office (automatic replies) on, off, or on for a period. Messages are plain text; line " +
      "breaks are kept. People outside the organisation get the external message, if externalAudience allows.",
    mutating: true,
    requiredScopes: ["MailboxSettings.ReadWrite"],
    inputSchema: z
      .object({
        status: z.enum(["disabled", "alwaysEnabled", "scheduled"]),
        internalMessage: z.string().optional().describe("Reply to colleagues."),
        externalMessage: z.string().optional().describe("Reply to people outside the organisation."),
        externalAudience: z.enum(["none", "contactsOnly", "all"]).optional(),
        start: When.optional().describe("Required when scheduled."),
        end: When.optional().describe("Required when scheduled."),
      })
      .refine((s) => s.status !== "scheduled" || (s.start && s.end), { message: "A scheduled reply needs start and end" }),
    handler: async (input, ctx) => {
      const setting: Record<string, unknown> = { status: input.status };
      if (input.internalMessage !== undefined) setting.internalReplyMessage = textToHtml(input.internalMessage);
      if (input.externalMessage !== undefined) setting.externalReplyMessage = textToHtml(input.externalMessage);
      if (input.externalAudience) setting.externalAudience = input.externalAudience;
      if (input.start) setting.scheduledStartDateTime = input.start;
      if (input.end) setting.scheduledEndDateTime = input.end;
      await ctx.graph.api(`/me/mailboxSettings`).patch({ automaticRepliesSetting: setting });
      return ctx.graph.api(`/me/mailboxSettings/automaticRepliesSetting`).get();
    },
  },
];
