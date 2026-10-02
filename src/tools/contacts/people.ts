import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { batchGet } from "../../graph/batch.js";

/** Colleagues, from the directory: who manages whom, who someone is, whether they are free. */

const Person = z.string().min(3).describe("Email address (UPN) or directory id of a colleague.");
const PROFILE = "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,businessPhones,mobilePhone";
const userPath = (u: string | undefined) => (u ? `/users/${encodeURIComponent(u)}` : "/me");

export const peopleTools: ToolDefinition[] = [
  {
    name: "people_get_manager",
    surface: "contacts",
    description: "Who someone's manager is (yours, by default).",
    requiredScopes: ["User.Read"],
    inputSchema: z.object({ user: Person.optional() }),
    handler: async ({ user }, ctx) => ctx.graph.api(`${userPath(user)}/manager`).select(PROFILE).get(),
  },
  {
    name: "people_list_direct_reports",
    surface: "contacts",
    description: "Who reports to someone (to you, by default).",
    requiredScopes: ["User.ReadBasic.All"],
    inputSchema: z.object({ user: Person.optional() }),
    handler: async ({ user }, ctx) => ctx.graph.api(`${userPath(user)}/directReports`).select(PROFILE).get(),
  },
  {
    name: "people_get_profile",
    surface: "contacts",
    description:
      "A colleague's directory profile: name, email, and where the tenant allows it, job title, department, " +
      "office and phone (those need User.Read.All, an admin-approved permission, for anyone but yourself).",
    requiredScopes: ["User.ReadBasic.All"],
    inputSchema: z.object({ user: Person }),
    handler: async ({ user }, ctx) => ctx.graph.api(userPath(user)).select(PROFILE).get(),
  },
  {
    name: "people_get_presence",
    surface: "contacts",
    description:
      "Whether people are available now, as Teams shows it: Available, Busy, DoNotDisturb, Away, Offline, " +
      "with what they are doing (InACall, InAMeeting, Presenting...). Yours, if no one is named.",
    requiredScopes: ["Presence.Read.All", "User.ReadBasic.All"],
    inputSchema: z.object({ users: z.array(Person).max(20).optional() }),
    handler: async ({ users }, ctx) => {
      if (!users?.length) return ctx.graph.api(`/me/presence`).get();
      // Presence is looked up by directory id; resolve addresses in one batch.
      const found = await batchGet(ctx.graph, users.map((u: string, i: number) => ({ id: String(i), url: `${userPath(u)}?$select=id,displayName,mail` })));
      const people = found.map((r, i) => ({ asked: users[i]!, ...(r.body as { id?: string; displayName?: string } | undefined), error: r.error?.message }));
      const ids = people.flatMap((p) => (p.id ? [p.id] : []));
      const presences = ids.length
        ? (((await ctx.graph.api(`/communications/getPresencesByUserId`).post({ ids })) as { value?: Array<Record<string, unknown>> }).value ?? [])
        : [];
      return people.map((p) => {
        const pr = presences.find((x) => x.id === p.id);
        return p.id
          ? { user: p.asked, name: p.displayName, availability: pr?.availability, activity: pr?.activity, statusMessage: (pr?.statusMessage as { message?: { content?: string } } | undefined)?.message?.content }
          : { user: p.asked, error: p.error ?? "not found in the directory" };
      });
    },
  },
];
