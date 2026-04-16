import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

export const onenoteTools: ToolDefinition[] = [
  {
    name: "onenote_list_notebooks",
    surface: "onenote",
    description: "List OneNote notebooks.",
    requiredScopes: ["Notes.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/onenote/notebooks`, input),
  },
  {
    name: "onenote_list_sections",
    surface: "onenote",
    description: "List sections (optionally within a notebook).",
    requiredScopes: ["Notes.Read"],
    inputSchema: PaginationInput.extend({ notebookId: z.string().optional() }),
    handler: async (input, ctx) => {
      const path = input.notebookId
        ? `/me/onenote/notebooks/${input.notebookId}/sections`
        : `/me/onenote/sections`;
      return fetchPage(ctx.graph, path, input);
    },
  },
  {
    name: "onenote_list_pages",
    surface: "onenote",
    description: "List pages (optionally within a section).",
    requiredScopes: ["Notes.Read"],
    inputSchema: PaginationInput.extend({ sectionId: z.string().optional() }),
    handler: async (input, ctx) => {
      const path = input.sectionId
        ? `/me/onenote/sections/${input.sectionId}/pages`
        : `/me/onenote/pages`;
      return fetchPage(ctx.graph, path, input);
    },
  },
  {
    name: "onenote_get_page_content",
    surface: "onenote",
    description: "Get the HTML content of a OneNote page.",
    requiredScopes: ["Notes.Read"],
    inputSchema: z.object({ pageId: z.string() }),
    handler: async ({ pageId }, ctx) => {
      const html = await ctx.graph.api(`/me/onenote/pages/${pageId}/content`).get();
      return { html: typeof html === "string" ? html : String(html) };
    },
  },
  {
    name: "onenote_create_page",
    surface: "onenote",
    description: "Create a new OneNote page in a section (HTML body).",
    mutating: true,
    requiredScopes: ["Notes.ReadWrite"],
    inputSchema: z.object({
      sectionId: z.string(),
      title: z.string(),
      html: z.string().describe("Page HTML body; a <title> will be injected."),
    }),
    handler: async (input, ctx) => {
      const body = `<!DOCTYPE html><html><head><title>${escapeHtml(input.title)}</title></head><body>${input.html}</body></html>`;
      return ctx.graph
        .api(`/me/onenote/sections/${input.sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(body);
    },
  },
  {
    name: "onenote_delete_page",
    surface: "onenote",
    description: "Delete a OneNote page.",
    mutating: true,
    requiredScopes: ["Notes.ReadWrite"],
    inputSchema: z.object({ pageId: z.string() }),
    handler: async ({ pageId }, ctx) => {
      await ctx.graph.api(`/me/onenote/pages/${pageId}`).delete();
      return { ok: true };
    },
  },
];

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
