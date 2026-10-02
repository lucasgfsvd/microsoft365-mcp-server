import { describe, it, expect } from "vitest";
import type { ServerConfig } from "../src/types.js";
import { calendarTools } from "../src/tools/calendar/index.js";
import { RecurrenceInput, toPatternedRecurrence } from "../src/tools/calendar/recurrence.js";
import { allTools } from "../src/tools/index.js";
import { ToolRegistry } from "../src/tools/registry.js";
import { assertAllowed } from "../src/util/writeGuard.js";
import { loadConfig } from "../src/config.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = (n: string) => findTool(calendarTools, n);
const rec = (r: object) => RecurrenceInput.parse(r);

describe("recurrence", () => {
  // 2026-10-06 is a Tuesday.
  it("reads 'every other Tuesday until Christmas' from the start date", () => {
    expect(toPatternedRecurrence(rec({ pattern: "weekly", interval: 2, until: "2026-12-25" }), "2026-10-06T09:00:00", "Europe/Madrid")).toEqual({
      pattern: { type: "weekly", interval: 2, daysOfWeek: ["tuesday"] },
      range: { type: "endDate", startDate: "2026-10-06", endDate: "2026-12-25", recurrenceTimeZone: "Europe/Madrid" },
    });
  });

  it("covers weekdays, monthly by date or by weekday, and yearly", () => {
    expect(toPatternedRecurrence(rec({ pattern: "weekly", daysOfWeek: ["monday", "tuesday", "wednesday", "thursday", "friday"], occurrences: 10 }), "2026-10-06T09:00:00", "UTC").range)
      .toMatchObject({ type: "numbered", numberOfOccurrences: 10 });
    expect(toPatternedRecurrence(rec({ pattern: "absoluteMonthly" }), "2026-10-06T09:00:00", "UTC").pattern).toEqual({ type: "absoluteMonthly", interval: 1, dayOfMonth: 6 });
    expect(toPatternedRecurrence(rec({ pattern: "relativeMonthly", weekIndex: "first" }), "2026-10-06T09:00:00", "UTC").pattern)
      .toEqual({ type: "relativeMonthly", interval: 1, daysOfWeek: ["tuesday"], index: "first" });
    expect(toPatternedRecurrence(rec({ pattern: "absoluteYearly" }), "2026-10-06T09:00:00", "UTC")).toMatchObject({
      pattern: { dayOfMonth: 6, month: 10 },
      range: { type: "noEnd" },
    });
  });

  it("refuses contradictions", () => {
    expect(() => rec({ pattern: "daily", until: "2026-12-01", occurrences: 3 })).toThrow(/not both/);
    expect(() => rec({ pattern: "relativeMonthly" })).toThrow(/weekIndex/);
  });
});

describe("calendar_create_event", () => {
  it("books a room as a resource attendee, and names it as the location", async () => {
    const ctx = makeContext();
    await callTool(tool("calendar_create_event"), ctx, {
      subject: "Planning", start: { dateTime: "2026-10-06T09:00:00" }, end: { dateTime: "2026-10-06T10:00:00" },
      attendees: [{ email: "bob@x.com" }], room: { email: "room-3a@x.com", name: "Room 3A" },
      recurrence: { pattern: "weekly", occurrences: 4 },
    });
    expect(ctx.mock.calls[0]!.body).toMatchObject({
      attendees: [
        { type: "required", emailAddress: { address: "bob@x.com" } },
        { type: "resource", emailAddress: { address: "room-3a@x.com", name: "Room 3A" } },
      ],
      location: { displayName: "Room 3A", locationEmailAddress: "room-3a@x.com" },
      recurrence: { pattern: { type: "weekly", daysOfWeek: ["tuesday"] }, range: { type: "numbered", numberOfOccurrences: 4 } },
    });
  });
});

describe("answering and cancelling", () => {
  it("declines with a proposed time, telling the organiser", async () => {
    const ctx = makeContext();
    const when = { start: { dateTime: "2026-10-07T15:00:00", timeZone: "UTC" }, end: { dateTime: "2026-10-07T15:30:00", timeZone: "UTC" } };
    await callTool(tool("calendar_respond_to_event"), ctx, { id: "e1", response: "decline", comment: "Clash", proposedNewTime: when });
    expect(ctx.mock.calls[0]).toMatchObject({ path: "/me/events/e1/decline", body: { sendResponse: true, comment: "Clash", proposedNewTime: when } });
  });

  it("maps tentative to Graph's action, and refuses a proposal on an accept", async () => {
    const ctx = makeContext();
    await callTool(tool("calendar_respond_to_event"), ctx, { id: "e1", response: "tentative" });
    expect(ctx.mock.calls[0]!.path).toBe("/me/events/e1/tentativelyAccept");
    const when = { start: { dateTime: "2026-10-07T15:00:00" }, end: { dateTime: "2026-10-07T15:30:00" } };
    await expect(callTool(tool("calendar_respond_to_event"), ctx, { id: "e1", response: "accept", proposedNewTime: when })).rejects.toThrow(/declining or tentatively/);
  });

  it("cancels with a note to attendees", async () => {
    const ctx = makeContext();
    await callTool(tool("calendar_cancel_event"), ctx, { id: "e1", comment: "Moved to next week" });
    expect(ctx.mock.calls[0]).toMatchObject({ method: "POST", path: "/me/events/e1/cancel", body: { comment: "Moved to next week" } });
  });
});

describe("rooms, behind the admin-consent flag", () => {
  const registry = new ToolRegistry();
  registry.registerAll(allTools());
  const base = { authMode: "device-code", disabledTools: new Set(), enableWrites: false, perSurfaceWrites: {} } as unknown as ServerConfig;

  it("is hidden until MCP_ENABLE_ADMIN_SCOPES, and says why", () => {
    expect(registry.list(base).some((t) => t.name === "calendar_list_rooms")).toBe(false);
    expect(registry.list({ ...base, adminScopes: true }).some((t) => t.name === "calendar_list_rooms")).toBe(true);
    expect(() => assertAllowed(registry.get("calendar_list_rooms")!, base)).toThrow(/MCP_ENABLE_ADMIN_SCOPES/);
  });

  it("asks for the admin scopes only when enabled", () => {
    const saved = process.env.MCP_ENABLE_ADMIN_SCOPES;
    try {
      delete process.env.MCP_ENABLE_ADMIN_SCOPES;
      expect(loadConfig(["node", "idx"]).scopes).not.toContain("Place.Read.All");
      process.env.MCP_ENABLE_ADMIN_SCOPES = "true";
      expect(loadConfig(["node", "idx"]).scopes).toEqual(expect.arrayContaining(["Place.Read.All", "User.Read.All"]));
    } finally {
      if (saved === undefined) delete process.env.MCP_ENABLE_ADMIN_SCOPES;
      else process.env.MCP_ENABLE_ADMIN_SCOPES = saved;
    }
  });

  it("lists rooms with what matters for booking, filtered", async () => {
    const ctx = makeContext();
    ctx.mock.on("/places/microsoft.graph.room", {
      value: [
        { displayName: "3A", emailAddress: "3a@x.com", capacity: 8, building: "HQ", floorNumber: 3, videoDeviceName: "Teams Room" },
        { displayName: "Booth", emailAddress: "b@x.com", capacity: 2, building: "HQ" },
      ],
    });
    const out = (await callTool(tool("calendar_list_rooms"), ctx, { minCapacity: 4 })) as { value: unknown[] };
    expect(out.value).toEqual([{ name: "3A", email: "3a@x.com", capacity: 8, building: "HQ", floor: 3, features: ["video"] }]);
  });
});
