import { describe, it, expect } from "vitest";
import { allTools } from "../src/tools/index.js";
import { SURFACES } from "../src/types.js";

describe("tool catalog", () => {
  const tools = allTools();

  it("registers tools across all 10 surfaces", () => {
    const present = new Set(tools.map((t) => t.surface));
    for (const s of SURFACES) expect(present.has(s)).toBe(true);
  });

  it("has unique tool names", () => {
    const names = tools.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("gives each mutating tool a conventional prefix", () => {
    const mutating = tools.filter((t) => t.mutating);
    const prefixes = /^(create_|update_|delete_|send_|post_|reply_|upload_|share_|add_|complete_|replace_|append_|run_|close_|mail_send|mail_reply|mail_delete|mail_create|teams_post|teams_reply|files_|calendar_create|calendar_update|calendar_delete|contacts_create|contacts_update|contacts_delete|todo_create|todo_complete|planner_create|planner_complete|onenote_create|onenote_delete|excel_|word_|powerpoint_)/;
    for (const t of mutating) expect(t.name).toMatch(prefixes);
  });
});
