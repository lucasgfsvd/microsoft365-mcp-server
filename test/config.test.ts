import { describe, it, expect, beforeEach, afterEach } from "vitest";
import path from "node:path";
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
  "MCP_MAX_MESSAGE_MB",
  "MCP_UPLOAD_DIR",
  "MCP_TOKEN_CACHE_KEY",
  "MCP_USER",
  "MCP_CLIENT_CERTIFICATE_PATH",
  "MCP_CLIENT_CERTIFICATE_PASSWORD",
  "MCP_ENABLE_SHARED_MAILBOXES",
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

  it("takes the token cache key from the environment, and refuses a short one", () => {
    expect(loadConfig(["node", "idx"]).tokenCacheKey).toBeUndefined();
    process.env.MCP_TOKEN_CACHE_KEY = "too-short";
    expect(() => loadConfig(["node", "idx"])).toThrow(/at least 16 characters/);
    process.env.MCP_TOKEN_CACHE_KEY = "k".repeat(32);
    expect(loadConfig(["node", "idx"]).tokenCacheKey).toBe("k".repeat(32));
  });

  it("leaves local uploads off unless MCP_UPLOAD_DIR names a folder, resolved to an absolute path", () => {
    expect(loadConfig(["node", "idx"]).uploadDir).toBeUndefined();
    process.env.MCP_UPLOAD_DIR = "uploads";
    expect(loadConfig(["node", "idx"]).uploadDir).toBe(path.resolve("uploads"));
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
    process.env.MCP_TENANT_ID = "contoso.onmicrosoft.com";
    process.env.MCP_USER = "alice@contoso.com";
    const c = loadConfig(["node", "idx"]);
    expect(c.scopes).toEqual(["https://graph.microsoft.com/.default"]);
    expect(c.user).toBe("alice@contoso.com");
  });

  // App-only has no signed-in user, so /me means nothing without one named.
  it("refuses client-credentials without a user or with a multi-tenant authority", () => {
    process.env.MCP_AUTH_MODE = "client-credentials";
    process.env.MCP_TENANT_ID = "contoso.onmicrosoft.com";
    expect(() => loadConfig(["node", "idx"])).toThrow(/MCP_USER/);
    process.env.MCP_USER = "alice@contoso.com";
    process.env.MCP_TENANT_ID = "common";
    expect(() => loadConfig(["node", "idx"])).toThrow(/tenant id/);
  });

  it("takes a certificate for client-credentials, by flag or environment", () => {
    const base = ["node", "idx", "--auth", "client-credentials", "--tenant", "contoso.com", "--user", "a@contoso.com"];
    expect(loadConfig([...base, "--client-certificate", "/certs/app.pem"]).clientCertificatePath).toBe("/certs/app.pem");
    process.env.MCP_CLIENT_CERTIFICATE_PATH = "/certs/env.pem";
    process.env.MCP_CLIENT_CERTIFICATE_PASSWORD = "pw";
    const c = loadConfig(base);
    expect([c.clientCertificatePath, c.clientCertificatePassword]).toEqual(["/certs/env.pem", "pw"]);
  });

  it("adds the shared-mailbox scopes only when asked, and only with sign-in", () => {
    expect(loadConfig(["node", "idx"]).sharedMailboxes).toBe(false);
    process.env.MCP_ENABLE_SHARED_MAILBOXES = "true";
    const read = loadConfig(["node", "idx"]);
    expect(read.sharedMailboxes).toBe(true);
    expect(read.scopes).toContain("Mail.Read.Shared");
    expect(read.scopes).not.toContain("Mail.Send.Shared");
    process.env.MCP_ENABLE_WRITES = "true";
    expect(loadConfig(["node", "idx"]).scopes).toEqual(expect.arrayContaining(["Mail.ReadWrite.Shared", "Mail.Send.Shared", "Calendars.ReadWrite.Shared"]));
    const appOnly = loadConfig(["node", "idx", "--auth", "client-credentials", "--tenant", "contoso.com", "--user", "a@contoso.com"]);
    expect(appOnly.sharedMailboxes).toBe(false);
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

  // The SDK's 10 MiB stdio default capped base64 uploads near 7.5 MB, and an
  // over-limit message ended the server.
  it("accepts 64 MiB messages by default, configurable", () => {
    expect(loadConfig(["node", "x"]).maxMessageBytes).toBe(64 * 1024 * 1024);
    process.env.MCP_MAX_MESSAGE_MB = "128";
    expect(loadConfig(["node", "x"]).maxMessageBytes).toBe(128 * 1024 * 1024);
  });

  it("rejects a message limit that is not a positive whole number", () => {
    for (const bad of ["0", "-5", "1.5", "lots"]) {
      process.env.MCP_MAX_MESSAGE_MB = bad;
      expect(() => loadConfig(["node", "x"])).toThrow(/MCP_MAX_MESSAGE_MB/);
    }
  });
});
