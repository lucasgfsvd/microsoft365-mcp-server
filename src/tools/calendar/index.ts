import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput, trimEmpty } from "../../util/schema.js";

const DateTimeTZ = z.object({
  dateTime: z.string().describe("ISO-8601, e.g. 2026-04-20T14:00:00"),
  timeZone: z.string().default("UTC"),
});

const Attendee = z.object({
  email: z.string().email(),
  name: z.string().optional(),
  type: z.enum(["required", "optional", "resource"]).default("required"),
});

export const calendarTools: ToolDefinition[] = [
  {
    name: "calendar_list_events",
    surface: "calendar",
    description: "List events in the default calendar, optionally constrained to a time window.",
    requiredScopes: ["Calendars.Read"],
    inputSchema: PaginationInput.extend({
      startDateTime: z.string().optional(),
      endDateTime: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const path = input.startDateTime && input.endDateTime
        ? `/me/calendarView?startDateTime=${encodeURIComponent(input.startDateTime)}&endDateTime=${encodeURIComponent(input.endDateTime)}`
        : `/me/events`;
      return fetchPage(ctx.graph, path, {
        ...input,
        orderBy: "start/dateTime",
        select: ["id", "subject", "start", "end", "organizer", "attendees", "location", "bodyPreview", "isOnlineMeeting"],
      });
    },
  },
  {
    name: "calendar_get_event",
    surface: "calendar",
    description: "Fetch a single event with full details.",
    requiredScopes: ["Calendars.Read"],
    inputSchema: z.object({ id: z.string() }),
    handler: async ({ id }, ctx) => ctx.graph.api(`/me/events/${id}`).get(),
  },
  {
    name: "calendar_list_calendars",
    surface: "calendar",
    description: "List the user's calendars.",
    requiredScopes: ["Calendars.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/calendars`, input),
  },
  {
    name: "calendar_find_meeting_times",
    surface: "calendar",
    description: "Ask Graph to suggest meeting times given attendees and constraints.",
    requiredScopes: ["Calendars.Read.Shared", "Calendars.Read"],
    inputSchema: z.object({
      attendees: z.array(Attendee).min(1),
      durationMinutes: z.number().int().min(5).max(480).default(30),
      timeWindowStart: z.string().describe("ISO-8601 window start"),
      timeWindowEnd: z.string().describe("ISO-8601 window end"),
      timeZone: z.string().default("UTC"),
    }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/findMeetingTimes`).post({
        attendees: input.attendees.map((a: z.infer<typeof Attendee>) => ({
          type: a.type,
          emailAddress: { address: a.email, name: a.name },
        })),
        meetingDuration: `PT${input.durationMinutes}M`,
        timeConstraint: {
          timeslots: [
            {
              start: { dateTime: input.timeWindowStart, timeZone: input.timeZone },
              end: { dateTime: input.timeWindowEnd, timeZone: input.timeZone },
            },
          ],
        },
      }),
  },
  {
    name: "calendar_get_free_busy",
    surface: "calendar",
    description: "Get free/busy schedule for a list of people.",
    requiredScopes: ["Calendars.Read.Shared"],
    inputSchema: z.object({
      schedules: z.array(z.string().email()).min(1),
      startDateTime: z.string(),
      endDateTime: z.string(),
      timeZone: z.string().default("UTC"),
      availabilityViewIntervalMinutes: z.number().int().min(5).max(1440).default(60),
    }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/calendar/getSchedule`).post({
        schedules: input.schedules,
        startTime: { dateTime: input.startDateTime, timeZone: input.timeZone },
        endTime: { dateTime: input.endDateTime, timeZone: input.timeZone },
        availabilityViewInterval: input.availabilityViewIntervalMinutes,
      }),
  },
  {
    name: "calendar_create_event",
    surface: "calendar",
    description: "Create a new calendar event.",
    mutating: true,
    requiredScopes: ["Calendars.ReadWrite"],
    inputSchema: z.object({
      subject: z.string(),
      start: DateTimeTZ,
      end: DateTimeTZ,
      body: z.string().optional(),
      bodyType: z.enum(["Text", "HTML"]).default("Text"),
      location: z.string().optional(),
      attendees: z.array(Attendee).optional(),
      isOnlineMeeting: z.boolean().default(false),
    }),
    handler: async (input, ctx) =>
      ctx.graph.api(`/me/events`).post(
        trimEmpty({
          subject: input.subject,
          start: input.start,
          end: input.end,
          body: input.body ? { contentType: input.bodyType, content: input.body } : undefined,
          location: input.location ? { displayName: input.location } : undefined,
          attendees: input.attendees?.map((a: z.infer<typeof Attendee>) => ({
            type: a.type,
            emailAddress: { address: a.email, name: a.name },
          })),
          isOnlineMeeting: input.isOnlineMeeting,
        }),
      ),
  },
  {
    name: "calendar_update_event",
    surface: "calendar",
    description: "Update an existing event (partial).",
    mutating: true,
    requiredScopes: ["Calendars.ReadWrite"],
    inputSchema: z.object({
      id: z.string(),
      subject: z.string().optional(),
      start: DateTimeTZ.optional(),
      end: DateTimeTZ.optional(),
      location: z.string().optional(),
    }),
    handler: async (input, ctx) => {
      const { id, ...rest } = input;
      const patch = trimEmpty({
        subject: rest.subject,
        start: rest.start,
        end: rest.end,
        location: rest.location ? { displayName: rest.location } : undefined,
      });
      return ctx.graph.api(`/me/events/${id}`).patch(patch);
    },
  },
  {
    name: "calendar_delete_event",
    surface: "calendar",
    description: "Cancel/delete an event.",
    mutating: true,
    requiredScopes: ["Calendars.ReadWrite"],
    inputSchema: z.object({ id: z.string() }),
    handler: async ({ id }, ctx) => {
      await ctx.graph.api(`/me/events/${id}`).delete();
      return { ok: true };
    },
  },
];
