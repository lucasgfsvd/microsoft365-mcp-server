import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

/**
 * SharePoint lists: the trackers, registers and request logs teams keep in a
 * site. Items are their fields, by internal column name (Title, Status, ...).
 */

const Site = z.string().min(1).describe("Site id, from sites_search.");
const Fields = z.record(z.string(), z.unknown()).describe("Column values by internal name, e.g. { Title: 'Laptop', Status: 'Ordered' }.");
const base = (siteId: string, listId: string) => `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}`;

export const siteListTools: ToolDefinition[] = [
  {
    name: "sites_list_lists",
    surface: "files",
    description: "List the SharePoint lists in a site (document libraries left out).",
    requiredScopes: ["Sites.Read.All"],
    inputSchema: z.object({ siteId: Site }),
    handler: async ({ siteId }, ctx) => {
      const res = (await ctx.graph.api(`/sites/${encodeURIComponent(siteId)}/lists`).select("id,displayName,description,webUrl,list").get()) as {
        value?: Array<{ id: string; displayName?: string; description?: string; webUrl?: string; list?: { template?: string } }>;
      };
      return (res.value ?? [])
        .filter((l) => l.list?.template !== "documentLibrary")
        .map((l) => ({ id: l.id, name: l.displayName, description: l.description, template: l.list?.template, webUrl: l.webUrl }));
    },
  },
  {
    name: "sites_get_list_items",
    surface: "files",
    description:
      "Read a SharePoint list's items with their field values. filter is OData on fields, e.g. " +
      "\"fields/Status eq 'Open'\"; on a large list it works only on indexed columns.",
    requiredScopes: ["Sites.Read.All"],
    inputSchema: PaginationInput.extend({ siteId: Site, listId: z.string().min(1), filter: z.string().optional() }),
    handler: async (input, ctx) => {
      const page = await fetchPage<{ id: string; fields?: Record<string, unknown>; webUrl?: string }>(
        ctx.graph,
        `${base(input.siteId, input.listId)}/items?$expand=fields`,
        input,
      );
      // SharePoint's own bookkeeping columns are noise in a conversation.
      const own = (f: Record<string, unknown> = {}) => Object.fromEntries(Object.entries(f).filter(([k]) => !k.startsWith("@") && !k.startsWith("_") && !/^(LinkTitle|LinkTitleNoMenu|ContentType|Edit|ItemChildCount|FolderChildCount|AppAuthor|AppEditor|Attachments)$/.test(k)));
      return { ...page, value: page.value.map((i) => ({ id: i.id, fields: own(i.fields), webUrl: i.webUrl })) };
    },
  },
  {
    name: "sites_create_list_item",
    surface: "files",
    description: "Add an item to a SharePoint list.",
    mutating: true,
    requiredScopes: ["Sites.ReadWrite.All"],
    inputSchema: z.object({ siteId: Site, listId: z.string().min(1), fields: Fields }),
    handler: async (input, ctx) => {
      const item = (await ctx.graph.api(`${base(input.siteId, input.listId)}/items`).post({ fields: input.fields })) as { id: string; fields?: unknown };
      return { id: item.id, fields: item.fields };
    },
  },
  {
    name: "sites_update_list_item",
    surface: "files",
    description: "Change field values of a SharePoint list item; fields not given are left as they are.",
    mutating: true,
    requiredScopes: ["Sites.ReadWrite.All"],
    inputSchema: z.object({ siteId: Site, listId: z.string().min(1), itemId: z.string().min(1), fields: Fields }),
    handler: async (input, ctx) =>
      ctx.graph.api(`${base(input.siteId, input.listId)}/items/${encodeURIComponent(input.itemId)}/fields`).patch(input.fields),
  },
];
