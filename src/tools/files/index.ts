import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { DrivePath, Filename, PaginationInput } from "../../util/schema.js";
import { uploadContent } from "../../graph/upload.js";
import { drivePrefix, ScopeInput } from "./scope.js";
import { filesDownloadTools } from "./download.js";

export const filesTools: ToolDefinition[] = [
  {
    name: "files_list_children",
    surface: "files",
    description: "List the contents of a OneDrive or SharePoint folder.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: PaginationInput.merge(ScopeInput).extend({
      itemId: z.string().optional().describe("Folder item id. Omit for root."),
      path: DrivePath.optional().describe("Path relative to root, e.g. '/Reports/2026'"),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const target = input.itemId
        ? `${base}/items/${input.itemId}/children`
        : input.path
          ? `${base}/root:${input.path}:/children`
          : `${base}/root/children`;
      return fetchPage(ctx.graph, target, input);
    },
  },
  {
    name: "files_get_item",
    surface: "files",
    description: "Get metadata for a single file or folder by id or path.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string().optional(),
      path: DrivePath.optional(),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const target = input.itemId
        ? `${base}/items/${input.itemId}`
        : `${base}/root:${input.path ?? ""}`;
      return ctx.graph.api(target).get();
    },
  },
  {
    name: "files_search",
    surface: "files",
    description: "Search files across OneDrive or a SharePoint drive.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: PaginationInput.merge(ScopeInput).extend({ query: z.string().min(1) }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return fetchPage(ctx.graph, `${base}/root/search(q='${encodeURIComponent(input.query)}')`, input);
    },
  },
  {
    name: "sites_search",
    surface: "files",
    description: "Search SharePoint sites the user can access.",
    requiredScopes: ["Sites.Read.All"],
    inputSchema: PaginationInput.extend({ query: z.string().min(1) }),
    handler: async (input, ctx) =>
      fetchPage(ctx.graph, `/sites?search=${encodeURIComponent(input.query)}`, input),
  },
  {
    name: "files_list_drives",
    surface: "files",
    description: "List the drives available to the user (OneDrive + SharePoint document libraries).",
    requiredScopes: ["Files.Read.All"],
    inputSchema: PaginationInput,
    handler: async (input, ctx) => fetchPage(ctx.graph, `/me/drives`, input),
  },
  {
    name: "files_upload",
    surface: "files",
    description:
      "Upload (or overwrite) a file. Any size: files over 4 MB are sent through an upload session " +
      "automatically. Content travels base64-encoded inside the tool call, which caps a file at " +
      "roughly 47 MB under the default MCP_MAX_MESSAGE_MB of 64; very large files are better " +
      "synced by the OneDrive client.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      parentPath: DrivePath.describe("Parent folder path, e.g. '/Reports'"),
      filename: Filename,
      contentBase64: z.string().describe("File bytes, base64-encoded."),
    }),
    handler: async (input, ctx) => {
      const buf = Buffer.from(input.contentBase64, "base64");
      return uploadContent(ctx.graph, drivePrefix(input), input, buf);
    },
  },
  {
    name: "files_create_folder",
    surface: "files",
    description: "Create a folder.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      parentPath: DrivePath,
      name: Filename,
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const target = `${base}/root:${input.parentPath}:/children`;
      return ctx.graph.api(target).post({
        name: input.name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      });
    },
  },
  {
    name: "files_delete",
    surface: "files",
    description: "Delete a file or folder (moved to recycle bin).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({ itemId: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      await ctx.graph.api(`${base}/items/${input.itemId}`).delete();
      return { ok: true };
    },
  },
  {
    name: "files_copy",
    surface: "files",
    description:
      "Copy a file to a new location under a new name. Works across OneDrive folders and SharePoint " +
      "sites (same-drive only for now). Accepts the source by itemId OR path.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string().optional(),
      path: DrivePath.optional().describe("Source path, e.g. '/Templates/proposal.docx'"),
      destinationParentPath: DrivePath.describe("Target folder path, e.g. '/Reports/2026'"),
      destinationName: Filename.describe("New filename"),
    }).refine((d) => Boolean(d.itemId) || Boolean(d.path), {
      message: "Provide either itemId or path",
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const sourceUrl = input.itemId
        ? `${base}/items/${input.itemId}/content`
        : `${base}/root:${input.path}:/content`;
      const stream: NodeJS.ReadableStream = await ctx.graph.api(sourceUrl).getStream();
      const chunks: Buffer[] = [];
      for await (const c of stream) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
      const buf = Buffer.concat(chunks);
      return uploadContent(ctx.graph, base, { parentPath: input.destinationParentPath, filename: input.destinationName }, buf);
    },
  },
  {
    name: "files_share",
    surface: "files",
    description: "Create a sharing link for a file.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string(),
      type: z.enum(["view", "edit", "embed"]).default("view"),
      scope: z.enum(["anonymous", "organization"]).default("organization"),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      return ctx.graph.api(`${base}/items/${input.itemId}/createLink`).post({
        type: input.type,
        scope: input.scope,
      });
    },
  },
  ...filesDownloadTools,
];
