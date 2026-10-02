import { describe, it, expect } from "vitest";
import { peopleTools } from "../src/tools/contacts/people.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = (n: string) => findTool(peopleTools, n);

describe("people", () => {
  it("finds your manager, or a colleague's, with the useful fields", async () => {
    const ctx = makeContext();
    await callTool(tool("people_get_manager"), ctx, {});
    await callTool(tool("people_get_manager"), ctx, { user: "bob@contoso.com" });
    expect(ctx.mock.calls.map((c) => c.path)).toEqual(["/me/manager", "/users/bob%40contoso.com/manager"]);
    expect(String(ctx.mock.calls[0]!.query.select)).toContain("jobTitle");
  });

  it("lists direct reports and looks up a profile", async () => {
    const ctx = makeContext();
    await callTool(tool("people_list_direct_reports"), ctx, {});
    await callTool(tool("people_get_profile"), ctx, { user: "ana@contoso.com" });
    expect(ctx.mock.calls.map((c) => c.path)).toEqual(["/me/directReports", "/users/ana%40contoso.com"]);
  });

  it("gives your own presence when no one is named", async () => {
    const ctx = makeContext();
    await callTool(tool("people_get_presence"), ctx, {});
    expect(ctx.mock.calls[0]!.path).toBe("/me/presence");
  });

  it("resolves addresses to ids in one batch, then asks presence for all at once", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", {
      responses: [
        { id: "0", status: 200, body: { id: "u-bob", displayName: "Bob" } },
        { id: "1", status: 404, body: { error: { code: "Request_ResourceNotFound", message: "No such user" } } },
      ],
    });
    ctx.mock.on("/communications/getPresencesByUserId", {
      value: [{ id: "u-bob", availability: "Busy", activity: "InAMeeting", statusMessage: { message: { content: "Back at 3" } } }],
    });
    const out = await callTool(tool("people_get_presence"), ctx, { users: ["bob@contoso.com", "ghost@contoso.com"] });
    expect(ctx.mock.calls[1]!.body).toEqual({ ids: ["u-bob"] });
    expect(out).toEqual([
      { user: "bob@contoso.com", name: "Bob", availability: "Busy", activity: "InAMeeting", statusMessage: "Back at 3" },
      { user: "ghost@contoso.com", error: "No such user" },
    ]);
  });
});
