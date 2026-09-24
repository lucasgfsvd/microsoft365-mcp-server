import { describe, it, expect } from "vitest";
import { availablePrompts, renderPrompt } from "../src/prompts/index.js";
import { dayWindow, parseDay, parseSince, type PromptContext } from "../src/prompts/types.js";
import { briefRequests, taskListUrl } from "../src/prompts/dailyBrief.js";

const NOW = new Date("2026-09-24T20:00:00Z");
const ALL = new Set([
  "graph_batch_get",
  "graph_search",
  "mail_list_messages",
  "mail_get_message",
  "mail_create_draft",
  "calendar_list_events",
  "calendar_get_event",
  "contacts_people_search",
]);
const ctx = (tools: Set<string> = ALL): PromptContext => ({ tools, now: NOW, timeZone: "Europe/Madrid" });

describe("time helpers", () => {
  it("reads relative and absolute lookbacks, and refuses anything else", () => {
    expect(parseSince(undefined, NOW).toISOString()).toBe("2026-09-23T20:00:00.000Z");
    expect(parseSince("3d", NOW).toISOString()).toBe("2026-09-21T20:00:00.000Z");
    expect(parseSince("6h", NOW).toISOString()).toBe("2026-09-24T14:00:00.000Z");
    expect(parseSince("2026-09-01", NOW).toISOString()).toBe("2026-09-01T00:00:00.000Z");
    expect(() => parseSince("last week", NOW)).toThrow(/24h/);
  });

  it("builds a day window in the user's zone, across daylight-saving changes", () => {
    expect(dayWindow("2026-06-09", "Europe/Madrid")).toEqual({
      start: "2026-06-08T22:00:00.000Z",
      end: "2026-06-09T22:00:00.000Z",
    });
    expect(dayWindow("2026-01-15", "Europe/Madrid").start).toBe("2026-01-14T23:00:00.000Z");
    // Clocks go forward on 29 March 2026: that day is 23 hours long.
    const { start, end } = dayWindow("2026-03-29", "Europe/Madrid");
    expect((Date.parse(end) - Date.parse(start)) / 3_600_000).toBe(23);
  });

  it("defaults the day to today in the user's zone, not UTC", () => {
    // 23:30 UTC is already the next day in Madrid.
    expect(parseDay(undefined, new Date("2026-09-24T23:30:00Z"), "Europe/Madrid")).toBe("2026-09-25");
    expect(() => parseDay("tomorrow", NOW, "Europe/Madrid")).toThrow(/2026-09-24/);
  });
});

describe("availablePrompts", () => {
  it("lists only prompts whose tools are enabled", () => {
    expect(availablePrompts(ALL).map((p) => p.name)).toEqual(["daily-brief", "inbox-triage", "meeting-prep"]);
    expect(availablePrompts(new Set(["mail_get_message"])).map((p) => p.name)).toEqual(["inbox-triage"]);
    expect(() => renderPrompt("daily-brief", {}, ctx(new Set()))).toThrow(/unavailable/);
  });
});

describe("rendering", () => {
  it("never lets a workflow send mail", () => {
    for (const p of availablePrompts(ALL)) {
      expect(renderPrompt(p.name, {}, ctx())).toMatch(/never call them/);
    }
  });

  it("drafts replies only when drafting is enabled", () => {
    expect(renderPrompt("inbox-triage", {}, ctx())).toContain("mail_create_draft");
    const readOnly = new Set([...ALL].filter((t) => t !== "mail_create_draft"));
    expect(renderPrompt("inbox-triage", {}, ctx(readOnly))).toContain("Writes are disabled");
  });

  it("falls back to the list tool when batching is disabled", () => {
    const noBatch = new Set([...ALL].filter((t) => t !== "graph_batch_get"));
    expect(renderPrompt("inbox-triage", {}, ctx(noBatch))).toContain("mail_list_messages");
  });

  it("keeps Graph search types in separate calls", () => {
    const text = renderPrompt("meeting-prep", {}, ctx());
    expect(text).toContain('entityTypes ["message"]');
    expect(text).toContain('entityTypes ["driveItem"]');
    expect(text).not.toMatch(/\["message",\s*"/);
  });

  it("treats a long base64 argument as an event id, anything else as a subject", () => {
    const id = "AAMkADUzZWI0ZjgyLTA4N2UtNDc1MS04NTY3LWY2M2NjYjY3MTFiYgBGAAAAAABR7nI3ivxgS6zcEfb77IrJBwDg";
    expect(renderPrompt("meeting-prep", { meeting: id }, ctx())).toContain("calendar_get_event");
    expect(renderPrompt("meeting-prep", { meeting: "thermal review" }, ctx())).toContain('contains "thermal review"');
  });
});

// Found against a live tenant: To Do answers "Invalid request" to a multi-field
// $select on lists, and to any $filter or $select on tasks inside $batch.
describe("To Do queries", () => {
  it("does not send query options To Do rejects", () => {
    const lists = briefRequests("2026-09-24", "Europe/Madrid", NOW).find((r) => r.id === "todo-lists")!;
    expect(lists.url).toBe("/me/todo/lists");
    expect(taskListUrl("abc")).toBe("/me/todo/lists/abc/tasks?$top=100");
  });
});
