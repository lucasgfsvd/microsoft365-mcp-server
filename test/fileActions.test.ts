import { describe, it, expect } from "vitest";
import { filesTools } from "../src/tools/files/index.js";
import { describePermission } from "../src/tools/files/permissions.js";
import { parentPathOf } from "../src/tools/files/manage.js";
import { callTool, findTool, makeContext } from "./helpers/mockGraph.js";

const tool = (n: string) => findTool(filesTools, n);

describe("files_move", () => {
  it("moves into a folder found by path, and renames, refusing to replace", async () => {
    const ctx = makeContext();
    ctx.mock.on("root:/Archive/2025", { id: "fold", folder: {} });
    ctx.mock.on("/items/i1", { id: "i1", name: "old-q3.xlsx", parentReference: { path: "/drive/root:/Archive/2025" } });
    const out = await callTool(tool("files_move"), ctx, { itemId: "i1", destinationParentPath: "/Archive/2025", newName: "old-q3.xlsx" });
    expect(ctx.mock.calls[1]).toMatchObject({
      method: "PATCH",
      path: "/me/drive/items/i1",
      body: { "@microsoft.graph.conflictBehavior": "fail", name: "old-q3.xlsx", parentReference: { id: "fold" } },
    });
    expect(out).toMatchObject({ id: "i1", folder: "/Archive/2025" });
  });

  it("refuses a destination that is not a folder, and a call that changes nothing", async () => {
    const ctx = makeContext();
    ctx.mock.on("root:/notes.txt", { id: "f", file: {} });
    await expect(callTool(tool("files_move"), ctx, { itemId: "i1", destinationParentPath: "/notes.txt" })).rejects.toThrow(/not a folder/);
    await expect(callTool(tool("files_move"), ctx, { itemId: "i1" })).rejects.toThrow(/destinationParentPath, newName/);
  });

  it("reads the folder of an item, root included, decoded", () => {
    expect(parentPathOf({ id: "x", name: "a", parentReference: { path: "/drive/root:/Q3%20Reports" } })).toBe("/Q3 Reports");
    expect(parentPathOf({ id: "x", name: "a", parentReference: { path: "/drive/root:" } })).toBe("/");
  });
});

describe("files_export_pdf", () => {
  it("converts with ?format=pdf and saves the PDF next to the original", async () => {
    const ctx = makeContext();
    ctx.mock.on("/items/i1", { id: "i1", name: "Board deck.pptx", size: 10, parentReference: { path: "/drive/root:/Decks" } });
    ctx.mock.on("format=pdf", Buffer.from("%PDF-1.7"));
    ctx.mock.on("Board deck.pdf", { id: "p1", name: "Board deck.pdf", size: 8 });
    const out = await callTool(tool("files_export_pdf"), ctx, { itemId: "i1" });
    expect(ctx.mock.calls.map((c) => `${c.method} ${c.path}`)).toEqual([
      "GET /me/drive/items/i1",
      "GET /me/drive/items/i1/content?format=pdf",
      "PUT /me/drive/root:/Decks/Board deck.pdf:/content",
    ]);
    expect(out).toMatchObject({ driveItemId: "p1", name: "Board deck.pdf", folder: "/Decks" });
  });

  it("says when OneDrive cannot convert the file", async () => {
    const ctx = makeContext();
    ctx.mock.on("/items/i1", { id: "i1", name: "photo.png", parentReference: {} });
    await expect(callTool(tool("files_export_pdf"), ctx, { itemId: "i1" })).rejects.toThrow(/cannot be converted/);
  });
});

describe("permissions", () => {
  it("describes access in plain terms: a person, a link, an inherited grant", () => {
    expect(describePermission({ id: "p1", roles: ["write"], grantedToV2: { user: { displayName: "Bob", email: "bob@x.com" } } })).toEqual({
      permissionId: "p1", roles: ["write"], via: "direct", who: "bob@x.com",
    });
    expect(describePermission({ id: "p2", roles: ["read"], link: { type: "view", scope: "organization" }, inheritedFrom: { path: "/drive/root:/Team" } })).toEqual({
      permissionId: "p2", roles: ["read"], via: "link", linkType: "view", linkScope: "organization", inheritedFrom: "/drive/root:/Team",
    });
  });

  it("invites people by email to view or edit, with a message", async () => {
    const ctx = makeContext();
    ctx.mock.on("/invite", { value: [{ id: "p9", roles: ["write"], grantedToV2: { user: { email: "ana@x.com" } } }] });
    const out = await callTool(tool("files_invite"), ctx, { itemId: "i1", recipients: ["ana@x.com"], role: "write", message: "Please review" });
    expect(ctx.mock.calls[0]!.body).toEqual({ recipients: [{ email: "ana@x.com" }], roles: ["write"], requireSignIn: true, sendInvitation: true, message: "Please review" });
    expect(out).toEqual([{ permissionId: "p9", roles: ["write"], via: "direct", who: "ana@x.com" }]);
  });

  it("removes one permission by id", async () => {
    const ctx = makeContext();
    await callTool(tool("files_remove_permission"), ctx, { itemId: "i1", permissionId: "p9" });
    expect(ctx.mock.calls[0]).toMatchObject({ method: "DELETE", path: "/me/drive/items/i1/permissions/p9" });
  });
});

describe("SharePoint lists", () => {
  it("lists lists without document libraries", async () => {
    const ctx = makeContext();
    ctx.mock.on("/lists", { value: [{ id: "l1", displayName: "Assets", list: { template: "genericList" } }, { id: "d", displayName: "Documents", list: { template: "documentLibrary" } }] });
    expect(await callTool(tool("sites_list_lists"), ctx, { siteId: "s1" })).toEqual([{ id: "l1", name: "Assets", description: undefined, template: "genericList", webUrl: undefined }]);
  });

  it("reads items as their own fields, without SharePoint's bookkeeping", async () => {
    const ctx = makeContext();
    ctx.mock.on("/items", { value: [{ id: "1", fields: { "@odata.etag": "x", Title: "Laptop", Status: "Ordered", _UIVersionString: "1.0", LinkTitle: "Laptop" } }] });
    const out = (await callTool(tool("sites_get_list_items"), ctx, { siteId: "s1", listId: "l1", filter: "fields/Status eq 'Ordered'" })) as { value: unknown[] };
    expect(out.value).toEqual([{ id: "1", fields: { Title: "Laptop", Status: "Ordered" }, webUrl: undefined }]);
    expect(ctx.mock.calls[0]).toMatchObject({ path: "/sites/s1/lists/l1/items?$expand=fields", query: { filter: "fields/Status eq 'Ordered'" } });
  });

  it("adds and updates items by field", async () => {
    const ctx = makeContext();
    await callTool(tool("sites_create_list_item"), ctx, { siteId: "s1", listId: "l1", fields: { Title: "Monitor" } });
    await callTool(tool("sites_update_list_item"), ctx, { siteId: "s1", listId: "l1", itemId: "7", fields: { Status: "Delivered" } });
    expect(ctx.mock.calls.map((c) => [c.method, c.path, c.body])).toEqual([
      ["POST", "/sites/s1/lists/l1/items", { fields: { Title: "Monitor" } }],
      ["PATCH", "/sites/s1/lists/l1/items/7/fields", { Status: "Delivered" }],
    ]);
  });
});
