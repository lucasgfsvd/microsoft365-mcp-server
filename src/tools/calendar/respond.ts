import { z } from "zod";
import type { ToolDefinition } from "../../types.js";

const When = z.object({ dateTime: z.string(), timeZone: z.string().default("UTC") });

const ACTIONS = { accept: "accept", tentative: "tentativelyAccept", decline: "decline" } as const;

/** Answering invitations, and calling off meetings the user organises. */
export const respondTools: ToolDefinition[] = [
  {
    name: "calendar_respond_to_event",
    surface: "calendar",
    description:
      "Accept, tentatively accept or decline a meeting invitation, optionally with a note to the organiser. " +
      "With tentative or decline you can propose another time, where the organiser allows proposals.",
    mutating: true,
    requiredScopes: ["Calendars.ReadWrite"],
    inputSchema: z
      .object({
        id: z.string().describe("The event, from calendar_list_events."),
        response: z.enum(["accept", "tentative", "decline"]),
        comment: z.string().optional().describe("Note to the organiser."),
        sendResponse: z.boolean().default(true).describe("false answers without telling the organiser."),
        proposedNewTime: z.object({ start: When, end: When }).optional(),
      })
      .refine((r) => !r.proposedNewTime || (r.response !== "accept" && r.sendResponse), {
        message: "A new time can be proposed only when declining or tentatively accepting, with sendResponse",
      }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/me/events/${input.id}/${ACTIONS[input.response as keyof typeof ACTIONS]}`).post({
        sendResponse: input.sendResponse,
        ...(input.comment ? { comment: input.comment } : {}),
        ...(input.proposedNewTime ? { proposedNewTime: input.proposedNewTime } : {}),
      });
      return { ok: true, response: input.response };
    },
  },
  {
    name: "calendar_cancel_event",
    surface: "calendar",
    description:
      "Call off a meeting you organise: every attendee gets a cancellation, with your note, and it leaves " +
      "the calendars. For a meeting someone else organises, decline it instead.",
    mutating: true,
    requiredScopes: ["Calendars.ReadWrite"],
    inputSchema: z.object({ id: z.string(), comment: z.string().optional().describe("Shown to attendees in the cancellation.") }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`/me/events/${input.id}/cancel`).post(input.comment ? { comment: input.comment } : {});
      return { ok: true };
    },
  },
];
