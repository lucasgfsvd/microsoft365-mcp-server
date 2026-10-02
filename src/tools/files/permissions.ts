import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { drivePrefix, ScopeInput } from "./scope.js";

interface Permission {
  id: string;
  roles?: string[];
  link?: { type?: string; scope?: string; webUrl?: string };
  grantedToV2?: { user?: { displayName?: string; email?: string }; siteUser?: { displayName?: string; email?: string }; group?: { displayName?: string } };
  grantedToIdentitiesV2?: Array<{ user?: { displayName?: string; email?: string } }>;
  inheritedFrom?: { path?: string };
  invitation?: { email?: string };
}

/** Who can open a file, in plain terms: a person, a group, or a link, and with what rights. */
export function describePermission(p: Permission) {
  const who = p.grantedToV2?.user ?? p.grantedToV2?.siteUser;
  const linkHolders = (p.grantedToIdentitiesV2 ?? []).map((g) => g.user?.email ?? g.user?.displayName).filter(Boolean);
  return {
    permissionId: p.id,
    roles: p.roles,
    ...(p.link
      ? { via: "link", linkType: p.link.type, linkScope: p.link.scope, ...(linkHolders.length ? { usedBy: linkHolders } : {}) }
      : { via: "direct", who: who?.email ?? who?.displayName ?? p.grantedToV2?.group?.displayName ?? p.invitation?.email }),
    ...(p.inheritedFrom?.path ? { inheritedFrom: p.inheritedFrom.path } : {}),
  };
}

export const filePermissionTools: ToolDefinition[] = [
  {
    name: "files_list_permissions",
    surface: "files",
    description: "Who can open a file or folder: people and groups given access, and sharing links, with their rights.",
    requiredScopes: ["Files.Read.All", "Sites.Read.All"],
    inputSchema: ScopeInput.extend({ itemId: z.string() }),
    handler: async (input, ctx) => {
      const res = (await ctx.graph.api(`${drivePrefix(input)}/items/${input.itemId}/permissions`).get()) as { value?: Permission[] };
      return (res.value ?? []).map(describePermission);
    },
  },
  {
    name: "files_invite",
    surface: "files",
    description:
      "Share a file or folder with specific people (by email), to view or to edit. They get an email with " +
      "your message, unless sendInvitation is false. For a link anyone in the organisation can use, see files_share.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({
      itemId: z.string(),
      recipients: z.array(z.email()).min(1).max(50),
      role: z.enum(["read", "write"]).default("read"),
      message: z.string().max(2000).optional(),
      sendInvitation: z.boolean().default(true),
    }),
    handler: async (input, ctx) => {
      const res = (await ctx.graph.api(`${drivePrefix(input)}/items/${input.itemId}/invite`).post({
        recipients: input.recipients.map((email: string) => ({ email })),
        roles: [input.role],
        requireSignIn: true,
        sendInvitation: input.sendInvitation,
        ...(input.message ? { message: input.message } : {}),
      })) as { value?: Permission[] };
      return (res.value ?? []).map(describePermission);
    },
  },
  {
    name: "files_remove_permission",
    surface: "files",
    description: "Take away a person's access, or disable a sharing link, using a permissionId from files_list_permissions.",
    mutating: true,
    requiredScopes: ["Files.ReadWrite.All", "Sites.ReadWrite.All"],
    inputSchema: ScopeInput.extend({ itemId: z.string(), permissionId: z.string() }),
    handler: async (input, ctx) => {
      await ctx.graph.api(`${drivePrefix(input)}/items/${input.itemId}/permissions/${input.permissionId}`).delete();
      return { ok: true };
    },
  },
];
