import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { assertGraphCursor, fetchDelta, initialDeltaUrl } from "../src/graph/delta.js";

const G = "https://graph.microsoft.com/v1.0";

/** Serves pages keyed by URL, recording what was requested. */
function mockGraph(pages: Record<string, unknown>) {
  const requested: string[] = [];
  const graph = {
    api: (url: string) => ({
      header() {
        return this;
      },
      get: () => {
        requested.push(url);
        if (!(url in pages)) return Promise.reject(new Error(`unexpected ${url}`));
        return Promise.resolve(pages[url]);
      },
    }),
  } as unknown as GraphClient;
  return { graph, requested };
}

describe("initialDeltaUrl", () => {
  it("syncs the inbox by default, without bodies", () => {
    const url = initialDeltaUrl("mail");
    expect(url).toMatch(/^\/me\/mailFolders\/inbox\/messages\/delta\?/);
    expect(decodeURIComponent(url)).toContain("$select=id,subject");
    expect(url).not.toContain("body,");
  });

  it("puts the window on calendar requests and refuses without one", () => {
    expect(initialDeltaUrl("calendar", { startDateTime: "2026-09-01T00:00:00Z", endDateTime: "2026-10-01T00:00:00Z" }))
      .toBe("/me/calendarView/delta?startDateTime=2026-09-01T00%3A00%3A00Z&endDateTime=2026-10-01T00%3A00%3A00Z");
    expect(() => initialDeltaUrl("calendar")).toThrow(/startDateTime/);
  });

  it("needs a list id for todo, and encodes it", () => {
    expect(() => initialDeltaUrl("todo")).toThrow(/listId/);
    expect(initialDeltaUrl("todo", { listId: "a/b=" })).toBe("/me/todo/lists/a%2Fb%3D/tasks/delta");
  });

  it("refuses select where Graph ignores or rejects it", () => {
    expect(() => initialDeltaUrl("todo", { listId: "l", select: ["title"] })).toThrow(/select/);
    expect(initialDeltaUrl("drive", { select: ["id", "name"] })).toBe("/me/drive/root/delta?%24select=id%2Cname");
  });

  // Drive ignores Prefer: odata.maxpagesize (verified live) but honours $top.
  it("sizes drive pages with $top, and only drive", () => {
    expect(initialDeltaUrl("drive", {}, 20)).toBe("/me/drive/root/delta?%24top=20");
    expect(initialDeltaUrl("contacts", {}, 20)).toBe("/me/contacts/delta");
  });
});

describe("assertGraphCursor", () => {
  it("accepts Graph delta links", () => {
    expect(() => assertGraphCursor(`${G}/me/drive/root/delta(token='abc')`)).not.toThrow();
    expect(() => assertGraphCursor(`${G}/me/mailFolders('x')/messages/delta?$deltatoken=t`)).not.toThrow();
  });

  // The Graph client attaches the bearer token to whatever it is given.
  it("refuses anything that would send the token elsewhere", () => {
    expect(() => assertGraphCursor("https://evil.example/me/messages/delta")).toThrow();
    expect(() => assertGraphCursor("https://graph.microsoft.com.evil.example/v1.0/me/contacts/delta")).toThrow();
    expect(() => assertGraphCursor("http://graph.microsoft.com/v1.0/me/contacts/delta")).toThrow();
    expect(() => assertGraphCursor(`${G}/me/messages`)).toThrow();
    expect(() => assertGraphCursor("not a url")).toThrow();
  });
});

describe("fetchDelta", () => {
  it("follows nextLinks to the deltaLink, separating removals from changes", async () => {
    const { graph, requested } = mockGraph({
      "/me/contacts/delta": {
        value: [{ id: "c1", displayName: "Ada", "@odata.etag": "W/1" }],
        "@odata.nextLink": `${G}/me/contacts/delta?$skiptoken=2`,
      },
      [`${G}/me/contacts/delta?$skiptoken=2`]: {
        value: [{ id: "c2", "@removed": { reason: "deleted" } }],
        "@odata.deltaLink": `${G}/me/contacts/delta?$deltatoken=d`,
      },
    });
    const out = await fetchDelta(graph, { resource: "contacts" });
    expect(requested).toHaveLength(2);
    expect(out).toEqual({
      changed: [{ id: "c1", displayName: "Ada" }],
      removed: [{ id: "c2", reason: "deleted" }],
      complete: true,
      deltaLink: `${G}/me/contacts/delta?$deltatoken=d`,
      pages: 2,
    });
  });

  it("reports drive deletions, which use a facet rather than @removed", async () => {
    const { graph } = mockGraph({
      "/me/drive/root/delta?%24top=500": {
        value: [{ id: "f1", name: "gone.txt", deleted: { state: "deleted" } }, { id: "f2", name: "kept.txt" }],
        "@odata.deltaLink": `${G}/me/drive/root/delta(token='t')`,
      },
    });
    const out = await fetchDelta(graph, { resource: "drive" });
    expect(out.removed).toEqual([{ id: "f1", reason: "deleted" }]);
    expect(out.changed.map((i) => i.id)).toEqual(["f2"]);
  });

  it("stops at maxItems and hands back the nextLink, marked incomplete", async () => {
    const next = `${G}/me/contacts/delta?$skiptoken=2`;
    const { graph, requested } = mockGraph({
      "/me/contacts/delta": { value: [{ id: "a" }, { id: "b" }], "@odata.nextLink": next },
    });
    const out = await fetchDelta(graph, { resource: "contacts" }, { maxItems: 2 });
    expect(requested).toHaveLength(1);
    expect(out).toMatchObject({ complete: false, nextLink: next });
    expect(out.deltaLink).toBeUndefined();
  });

  it("resumes from a cursor", async () => {
    const cursor = `${G}/me/contacts/delta?$deltatoken=d`;
    const { graph, requested } = mockGraph({
      [cursor]: { value: [], "@odata.deltaLink": `${G}/me/contacts/delta?$deltatoken=e` },
    });
    const out = await fetchDelta(graph, { cursor });
    expect(requested).toEqual([cursor]);
    expect(out).toMatchObject({ complete: true, changed: [], removed: [] });
  });

  it("refuses a foreign cursor before making any request", async () => {
    const { graph, requested } = mockGraph({});
    await expect(fetchDelta(graph, { cursor: "https://evil.example/delta" })).rejects.toThrow();
    expect(requested).toHaveLength(0);
  });

  it("fails loudly on a page with neither link", async () => {
    const { graph } = mockGraph({ "/me/contacts/delta": { value: [] } });
    await expect(fetchDelta(graph, { resource: "contacts" })).rejects.toThrow(/neither/);
  });
});
