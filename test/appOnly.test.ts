import { describe, it, expect } from "vitest";
import { ClientCertificateCredential } from "@azure/identity";
import type { ServerConfig } from "../src/types.js";
import { graphForUser, retarget, NO_USER_MESSAGE } from "../src/graph/targetUser.js";
import { batchGet } from "../src/graph/batch.js";
import { allTools } from "../src/tools/index.js";
import { mailTools } from "../src/tools/mail/index.js";
import { ToolRegistry } from "../src/tools/registry.js";
import { assertAllowed } from "../src/util/writeGuard.js";
import { offersMailbox, takeMailbox, withMailboxParam } from "../src/tools/mailbox.js";
import { buildCredential } from "../src/auth/index.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const alice = { user: "alice@contoso.com", required: true };

describe("retarget", () => {
  it("points /me at the named user, and leaves everything else alone", () => {
    expect(retarget("/me", alice)).toBe("/users/alice%40contoso.com");
    expect(retarget("/me/messages?$top=5", alice)).toBe("/users/alice%40contoso.com/messages?$top=5");
    expect(retarget("/me?$select=mail", alice)).toBe("/users/alice%40contoso.com?$select=mail");
    for (const other of ["/meetings", "/teams/1/channels", "/planner/tasks/x", "/users/bob/messages"]) {
      expect(retarget(other, alice)).toBe(other);
    }
  });

  it("keeps /me for a signed-in user, and refuses it when there is none", () => {
    expect(retarget("/me/drive", { required: false })).toBe("/me/drive");
    expect(() => retarget("/me/drive", { required: true })).toThrow(NO_USER_MESSAGE);
    expect(retarget("/sites?search=x", { required: true })).toBe("/sites?search=x");
  });
});

describe("graphForUser", () => {
  it("rewrites what tools ask for, through the same client", async () => {
    const ctx = makeContext();
    ctx.graph = graphForUser(ctx.graph, alice);
    await callTool(findTool(mailTools, "mail_list_messages"), ctx, { folder: "Inbox" });
    expect(ctx.mock.calls[0]!.path).toBe("/users/alice%40contoso.com/mailFolders/Inbox/messages");
  });

  // The prompts and resources build /me URLs that travel inside a $batch body.
  it("rewrites $batch sub-requests too, and reports them under the URL asked for", async () => {
    const ctx = makeContext();
    ctx.mock.on("/$batch", (call) => ({
      responses: (call.body as { requests: Array<{ id: string; url: string }> }).requests.map((r) => ({ id: r.id, status: 200, body: { url: r.url } })),
    }));
    const [r] = await batchGet(graphForUser(ctx.graph, alice), [{ id: "a", url: "/me/messages?$top=1" }]);
    expect(r).toMatchObject({ url: "/me/messages?$top=1", body: { url: "/users/alice%40contoso.com/messages?$top=1" } });
  });
});

const config = (over: Partial<ServerConfig> = {}): ServerConfig => ({
  authMode: "device-code",
  tenantId: "contoso.com",
  clientId: "cid",
  tokenCachePath: "/tmp/tc.json",
  scopes: [],
  enableWrites: true,
  perSurfaceWrites: {},
  disabledTools: new Set(),
  logLevel: "silent",
  listTools: false,
  logout: false,
  sharedMailboxes: false,
  ...over,
});

describe("app-only mode", () => {
  const registry = new ToolRegistry();
  registry.registerAll(allTools());

  it("hides the tools Graph refuses without a signed-in user", () => {
    const appOnly = new Set(registry.list(config({ authMode: "client-credentials", user: "a@contoso.com" })).map((t) => t.name));
    for (const gone of ["onenote_create_page", "teams_post_chat_message", "graph_search"]) expect(appOnly.has(gone), gone).toBe(false);
    for (const kept of ["mail_send_message", "calendar_create_event", "files_upload", "teams_list_channel_messages"]) expect(appOnly.has(kept), kept).toBe(true);
    expect(registry.list(config()).some((t) => t.name === "graph_search")).toBe(true);
  });

  it("says why a hidden tool cannot run, rather than calling it a write", () => {
    const tool = registry.get("onenote_list_notebooks")!;
    expect(() => assertAllowed(tool, config({ authMode: "client-credentials" }))).toThrow(/app-only/);
    expect(() => assertAllowed(registry.get("mail_list_messages")!, config({ disabledTools: new Set(["mail_list_messages"]) }))).toThrow(/MCP_DISABLED_TOOLS/);
    expect(() => assertAllowed(registry.get("mail_send_message")!, config({ enableWrites: false }))).toThrow(/write operation/);
  });

  it("signs in with a certificate when one is given", async () => {
    const cred = await buildCredential(config({ authMode: "client-credentials", clientCertificatePath: "/certs/app.pem", clientSecret: "also-set" }));
    expect(cred).toBeInstanceOf(ClientCertificateCredential);
  });
});

describe("the mailbox argument", () => {
  const send = findTool(mailTools, "mail_send_message");
  const on = config({ sharedMailboxes: true });

  it("is offered on mail and calendar tools only, and only when enabled", () => {
    expect(offersMailbox(send, on)).toBe(true);
    expect(offersMailbox({ surface: "calendar" }, on)).toBe(true);
    expect(offersMailbox({ surface: "files" }, on)).toBe(false);
    expect(offersMailbox(send, config())).toBe(false);
    expect(withMailboxParam({ type: "object", properties: { id: {} } }).properties).toHaveProperty("mailbox");
  });

  it("is split from the tool's own arguments", () => {
    expect(takeMailbox(send, on, { mailbox: " support@contoso.com ", subject: "x" })).toEqual({ mailbox: "support@contoso.com", args: { subject: "x" } });
    expect(takeMailbox(send, config(), { subject: "x" })).toEqual({ args: { subject: "x" } });
  });

  // Ignoring it would quietly act on the user's own mailbox instead.
  it("is refused, never ignored, where it is not offered", () => {
    expect(() => takeMailbox(send, config(), { mailbox: "support@contoso.com" })).toThrow(/MCP_ENABLE_SHARED_MAILBOXES/);
    expect(() => takeMailbox({ name: "files_upload", surface: "files" }, on, { mailbox: "s@contoso.com" })).toThrow(/takes no mailbox/);
    expect(() => takeMailbox(send, config({ authMode: "client-credentials" }), { mailbox: "s@contoso.com" })).toThrow(/MCP_USER only/);
    expect(() => takeMailbox(send, on, { mailbox: "../../users/ceo" })).toThrow(/email address/);
  });
});
