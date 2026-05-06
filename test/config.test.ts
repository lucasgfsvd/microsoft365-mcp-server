import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { loadConfig } from "../src/config.js";

const ENV_KEYS = [
  "MCP_AUTH_MODE",
  "MCP_TENANT_ID",
  "MCP_CLIENT_ID",
  "MCP_CLIENT_SECRET",
  "MCP_ENABLE_WRITES",
  "MCP_ENABLE_MAIL_WRITE",
  "MCP_DISABLED_TOOLS",
  "MCP_SCOPES",
];

describe("loadConfig", () => {
  const saved: Record<string, string | undefined> = {};
  beforeEach(() => {
    for (const k of ENV_KEYS) {
      saved[k] = process.env[k];
      delete process.env[k];
    }
  });
  afterEach(() => {
    for (const k of ENV_KEYS) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
  });

  it("defaults to device-code, read-only, common tenant", () => {
    const c = loadConfig(["node", "idx"]);
    expect(c.authMode).toBe("device-code");
    expect(c.enableWrites).toBe(false);
    expect(c.tenantId).toBe("common");
  });

  it("respects MCP_ENABLE_WRITES=true", () => {
    process.env.MCP_ENABLE_WRITES = "true";
    const c = loadConfig(["node", "idx"]);
    expect(c.enableWrites).toBe(true);
  });

  it("parses per-surface MCP_ENABLE_MAIL_WRITE=true", () => {
    process.env.MCP_ENABLE_MAIL_WRITE = "true";
    const c = loadConfig(["node", "idx"]);
    expect(c.perSurfaceWrites.mail).toBe(true);
  });

  it("uses .default scope for client-credentials", () => {
    process.env.MCP_AUTH_MODE = "client-credentials";
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toEqual(["https://graph.microsoft.com/.default"]);
  });

  it("parses --disabled-tools CSV", () => {
    const c = loadConfig(["node", "idx", "--disabled-tools", "mail_send_message,mail_delete_message"]);
    expect(c.disabledTools.has("mail_send_message")).toBe(true);
    expect(c.disabledTools.has("mail_delete_message")).toBe(true);
  });

  it("defaults to read-only scopes when writes are disabled", () => {
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toContain("Mail.Read");
    expect(c.scopes).toContain("Files.Read.All");
    expect(c.scopes).not.toContain("Mail.Send");
    expect(c.scopes).not.toContain("Files.ReadWrite.All");
  });

  it("adds write scopes when MCP_ENABLE_WRITES=true", () => {
    process.env.MCP_ENABLE_WRITES = "true";
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toContain("Mail.Send");
    expect(c.scopes).toContain("Files.ReadWrite.All");
    expect(c.scopes).toContain("Mail.Read");
  });

  it("adds write scopes when any per-surface write flag is enabled", () => {
    process.env.MCP_ENABLE_MAIL_WRITE = "true";
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toContain("Mail.Send");
  });

  it("MCP_SCOPES fully overrides the computed defaults", () => {
    process.env.MCP_SCOPES = "User.Read,Mail.Read";
    process.env.MCP_ENABLE_WRITES = "true";
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toEqual(["User.Read", "Mail.Read"]);
  });
});
