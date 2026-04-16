import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

const EmailAddr = z.object({ address: z.string().email(), name: z.string().optional() });

export const contactsTools: ToolDefinition[] = [
  {
    name: "contacts_list",
    surface: "contacts",
    description: "List personal contacts.",
    requiredScopes: ["Contacts.Read"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/contacts`, input),
  },
  {
    name: "contacts_search",
    surface: "contacts",
    description: "Search personal contacts by name or email.",
    requiredScopes: ["Contacts.Read"],
    inputSchema: PaginationInput.extend({ query: z.string().min(1) }),
    handler: async (input, ctx) =>
      fetchPage(ctx.graph, `/me/contacts`, { ...input, search: input.query }),
  },
  {
    name: "contacts_people_search",
    surface: "contacts",
    description: "Search the People API (mix of contacts, directory, and frequent collaborators).",
    requiredScopes: ["People.Read"],
    inputSchema: PaginationInput.extend({ query: z.string().min(1) }),
    handler: async (input, ctx) =>
      fetchPage(ctx.graph, `/me/people`, { ...input, search: input.query }),
  },
  {
    name: "contacts_create",
    surface: "contacts",
    description: "Create a personal contact.",
    mutating: true,
    requiredScopes: ["Contacts.ReadWrite"],
    inputSchema: z.object({
      givenName: z.string(),
      surname: z.string().optional(),
      emailAddresses: z.array(EmailAddr).optional(),
      companyName: z.string().optional(),
      mobilePhone: z.string().optional(),
    }),
    handler: async (input, ctx) => ctx.graph.api(`/me/contacts`).post(input),
  },
  {
    name: "contacts_update",
    surface: "contacts",
    description: "Update a contact (partial).",
    mutating: true,
    requiredScopes: ["Contacts.ReadWrite"],
    inputSchema: z.object({
      id: z.string(),
      patch: z.record(z.any()).describe("Fields to update, e.g. {mobilePhone:'+1...'}"),
    }),
    handler: async ({ id, patch }, ctx) => ctx.graph.api(`/me/contacts/${id}`).patch(patch),
  },
  {
    name: "contacts_delete",
    surface: "contacts",
    description: "Delete a contact.",
    mutating: true,
    requiredScopes: ["Contacts.ReadWrite"],
    inputSchema: z.object({ id: z.string() }),
    handler: async ({ id }, ctx) => {
      await ctx.graph.api(`/me/contacts/${id}`).delete();
      return { ok: true };
    },
  },
];
