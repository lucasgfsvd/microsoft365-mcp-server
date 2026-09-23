import { describe, it, expect } from "vitest";
import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";
import { assertCombinableEntityTypes, searchQuery } from "../src/graph/search.js";

function mockGraph(reply: unknown) {
  const posted: unknown[] = [];
  const graph = {
    api: (path: string) => ({
      post: (body: unknown) => {
        posted.push({ path, body });
        return Promise.resolve(reply);
      },
    }),
  } as unknown as GraphClient;
  return { graph, posted };
}

describe("assertCombinableEntityTypes", () => {
  it("allows the SharePoint/OneDrive family together", () => {
    expect(() => assertCombinableEntityTypes(["driveItem", "site", "listItem", "list"])).not.toThrow();
  });

  // A live tenant rejects this pairing, despite both being "mail-ish".
  it("rejects message + event", () => {
    expect(() => assertCombinableEntityTypes(["message", "event"])).toThrow(/mail|calendar/);
  });

  it("rejects mixing groups, naming them", () => {
    expect(() => assertCombinableEntityTypes(["message", "driveItem"])).toThrow(/mail|files & sites/);
  });

  it("rejects chatMessage or person mixed with anything else", () => {
    expect(() => assertCombinableEntityTypes(["chatMessage", "message"])).toThrow();
    expect(() => assertCombinableEntityTypes(["person", "driveItem"])).toThrow();
  });

  it("allows a single type", () => {
    expect(() => assertCombinableEntityTypes(["person"])).not.toThrow();
  });
});

describe("searchQuery", () => {
  const reply = {
    value: [
      {
        hitsContainers: [
          {
            total: 2,
            moreResultsAvailable: true,
            hits: [
              {
                summary: "…thermal…",
                resource: {
                  "@odata.type": "#microsoft.graph.driveItem",
                  id: "d1",
                  name: "report.docx",
                  webUrl: "https://x/report.docx",
                  lastModifiedDateTime: "2026-09-01T00:00:00Z",
                },
              },
              {
                resource: { "@odata.type": "#microsoft.graph.message", id: "m1", subject: "Re: review", webLink: "https://x/m1" },
              },
            ],
          },
        ],
      },
    ],
  };

  it("posts a single search request with the query and paging", async () => {
    const { graph, posted } = mockGraph(reply);
    await searchQuery(graph, { query: "thermal", entityTypes: ["driveItem"], size: 10, from: 5 });
    expect(posted[0]).toEqual({
      path: "/search/query",
      body: { requests: [{ entityTypes: ["driveItem"], query: { queryString: "thermal" }, from: 5, size: 10 }] },
    });
  });

  it("defaults paging when not given", async () => {
    const { graph, posted } = mockGraph(reply);
    await searchQuery(graph, { query: "x", entityTypes: ["message"] });
    const body = (posted[0] as { body: { requests: Array<{ from: number; size: number }> } }).body;
    expect(body.requests[0]).toMatchObject({ from: 0, size: 25 });
  });

  it("flattens hits and strips the odata type prefix", async () => {
    const { graph } = mockGraph(reply);
    const out = await searchQuery(graph, { query: "thermal", entityTypes: ["driveItem"] });
    expect(out.total).toBe(2);
    expect(out.moreAvailable).toBe(true);
    expect(out.hits[0]).toMatchObject({ entityType: "driveItem", name: "report.docx", webUrl: "https://x/report.docx" });
  });

  it("names a hit from whichever field the entity uses", async () => {
    const { graph } = mockGraph(reply);
    const out = await searchQuery(graph, { query: "x", entityTypes: ["message"] });
    expect(out.hits[1]).toMatchObject({ entityType: "message", name: "Re: review", webUrl: "https://x/m1" });
  });

  it("refuses a mixed-group query before calling Graph", async () => {
    const { graph, posted } = mockGraph(reply);
    await expect(searchQuery(graph, { query: "x", entityTypes: ["message", "driveItem"] })).rejects.toThrow();
    expect(posted).toHaveLength(0);
  });

  it("survives an empty result set", async () => {
    const { graph } = mockGraph({ value: [{ hitsContainers: [{ total: 0, hits: [] }] }] });
    const out = await searchQuery(graph, { query: "zzz", entityTypes: ["message"] });
    expect(out.hits).toEqual([]);
  });
});
