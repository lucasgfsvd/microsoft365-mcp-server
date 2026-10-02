import { describe, it, expect } from "vitest";
import { listFilter, mailTools } from "../src/tools/mail/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

describe("mail_list_messages date filters", () => {
  const list = findTool(mailTools, "mail_list_messages");

  it("leads with receivedDateTime, which Graph requires when ordering by it", async () => {
    const ctx = makeContext();
    await callTool(list, ctx, { receivedAfter: "2026-09-28", receivedBefore: "2026-10-05T00:00:00+02:00", unreadOnly: true });
    expect(ctx.mock.calls[0]!.query.filter).toBe(
      "receivedDateTime ge 2026-09-28T00:00:00.000Z and receivedDateTime lt 2026-10-04T22:00:00.000Z and isRead eq false",
    );
    expect(ctx.mock.calls[0]!.query.orderby).toBe("receivedDateTime desc");
  });

  it("filters on one bound alone, or not at all", () => {
    expect(listFilter({ receivedAfter: "2026-09-28T08:00:00Z" })).toBe("receivedDateTime ge 2026-09-28T08:00:00.000Z");
    expect(listFilter({ unreadOnly: true })).toBe("isRead eq false");
    expect(listFilter({})).toBeUndefined();
  });

  it("refuses dates it cannot read, and an empty window", async () => {
    const ctx = makeContext();
    await expect(callTool(list, ctx, { receivedAfter: "last week" })).rejects.toThrow();
    await expect(callTool(list, ctx, { receivedAfter: "2026-09-28T08:00:00" })).rejects.toThrow(); // no offset: ambiguous
    await expect(callTool(list, ctx, { receivedAfter: "2026-10-05", receivedBefore: "2026-09-28" })).rejects.toThrow(/earlier/);
    expect(ctx.mock.calls).toHaveLength(0);
  });
});
