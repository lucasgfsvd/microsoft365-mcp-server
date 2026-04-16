import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

function drivePrefix(input: { driveId?: string; siteId?: string }): string {
  if (input.driveId) return `/drives/${input.driveId}`;
  if (input.siteId) return `/sites/${input.siteId}/drive`;
  return `/me/drive`;
}

const ScopeInput = z.object({
  driveId: z.string().optional().describe("Drive id. Omit to use the user's OneDrive."),
  siteId: z.string().optional().describe("SharePoint site id (use sites_search to discover)."),
});

export const filesTools: ToolDefinition[] = [
  {
    name: "files_list_children",
    surface: "files",
    description: "List the contents of a OneDrive or SharePoint folder.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: PaginationInput.merge(ScopeInput).extend({
      itemId: z.string().optional().describe("Folder item id. Omit for root."),
      path: z.string().optional().describe("Path relative to root, e.g. '/Reports/2026'"),
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
      path: z.string().optional(),
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
    name: "files_download",
    surface: "files",
    description: "Download a file's bytes (base64-encoded in the response).",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: ScopeInput.extend({ itemId: z.string() }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const ab: ArrayBuffer = await ctx.graph.api(`${base}/items/${input.itemId}/content`).getStream().then(streamToArrayBuffer);
      return { base64: Buffer.from(ab).toString("base64"), byteLength: ab.byteLength };
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
    description: "Upload (or overwrite) a file. Use for files ≤ 4 MB; larger files require an upload session (not yet exposed).",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      parentPath: z.string().describe("Parent folder path, e.g. '/Reports'"),
      filename: z.string(),
      contentBase64: z.string().describe("File bytes, base64-encoded."),
    }),
    handler: async (input, ctx) => {
      const base = drivePrefix(input);
      const buf = Buffer.from(input.contentBase64, "base64");
      if (buf.byteLength > 4 * 1024 * 1024) {
        throw new Error("files_upload only supports <=4 MB. Large file upload sessions are not yet implemented.");
      }
      const path = `${base}/root:${input.parentPath.replace(/\/$/, "")}/${input.filename}:/content`;
      return ctx.graph.api(path).put(buf);
    },
  },
  {
    name: "files_create_folder",
    surface: "files",
    description: "Create a folder.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      parentPath: z.string(),
      name: z.string(),
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
      path: z.string().optional().describe("Source path, e.g. '/Templates/proposal.docx'"),
      destinationParentPath: z.string().describe("Target folder path, e.g. '/Reports/2026'"),
      destinationName: z.string().describe("New filename"),
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
      if (buf.byteLength > 4 * 1024 * 1024) {
        throw new Error(
          "files_copy currently downloads+uploads and is capped at 4 MB. Use Graph's async copy for larger files.",
        );
      }
      const parent = input.destinationParentPath.replace(/\/$/, "");
      const targetPath = `${base}/root:${parent}/${input.destinationName}:/content`;
      return ctx.graph.api(targetPath).put(buf);
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
];

async function streamToArrayBuffer(stream: NodeJS.ReadableStream): Promise<ArrayBuffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of stream) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  const buf = Buffer.concat(chunks);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}
