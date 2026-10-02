import { describe, it, expect } from "vitest";
import { allTools } from "../src/tools/index.js";
import { annotationsFor, writeKind } from "../src/tools/annotations.js";

const tools = allTools();
const of = (name: string) => annotationsFor(tools.find((t) => t.name === name)!);

describe("tool annotations", () => {
  it("classifies every write tool, so a new one cannot slip through as unknown", () => {
    const unknown = tools.filter((t) => t.mutating && !writeKind(t.name)).map((t) => t.name);
    expect(unknown).toEqual([]);
  });

  it("marks reads as read-only, and nothing else as read-only", () => {
    for (const t of tools) expect(annotationsFor(t).readOnlyHint).toBe(!t.mutating);
    expect(of("mail_search_messages")).toEqual({ readOnlyHint: true, openWorldHint: false });
  });

  it("tells sending apart from deleting", () => {
    expect(of("mail_send_message")).toEqual({ readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true });
    expect(of("files_delete")).toEqual({ readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false });
    expect(of("mail_create_reply_draft")).toMatchObject({ destructiveHint: false, openWorldHint: false });
  });

  it("flags what other people see or receive", () => {
    for (const name of ["mail_reply_message", "teams_post_chat_message", "calendar_update_event", "planner_delete_task", "files_share"]) {
      expect(of(name).openWorldHint, name).toBe(true);
    }
    for (const name of ["todo_create_task", "word_append_paragraph", "onenote_create_page"]) {
      expect(of(name).openWorldHint, name).toBe(false);
    }
  });

  it("knows a file 'create' replaces, and a delete by position is not repeatable", () => {
    expect(of("excel_create_workbook")).toMatchObject({ destructiveHint: true, idempotentHint: true });
    expect(of("word_delete_paragraph")).toMatchObject({ destructiveHint: true, idempotentHint: false });
    expect(of("files_create_folder")).toMatchObject({ destructiveHint: false, idempotentHint: false });
  });

  it("assumes the worst of a write it cannot classify", () => {
    expect(annotationsFor({ name: "x_frobnicate_thing", mutating: true })).toMatchObject({ destructiveHint: true, idempotentHint: false });
  });
});
