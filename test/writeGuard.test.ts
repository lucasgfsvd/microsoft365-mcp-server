import { describe, it, expect } from "vitest";
import { z } from "zod";
import { isToolAllowed, WriteBlockedError, assertAllowed } from "../src/util/writeGuard.js";
import type { ServerConfig, ToolDefinition } from "../src/types.js";

const baseConfig: ServerConfig = {
  authMode: "device-code",
  tenantId: "common",
  clientId: "x",
  tokenCachePath: "/tmp/tc.json",
  scopes: [],
  enableWrites: false,
  perSurfaceWrites: {},
  disabledTools: new Set(),
  logLevel: "info",
};

const readTool: ToolDefinition = {
  name: "mail_list_messages",
  surface: "mail",
  description: "",
  inputSchema: z.object({}),
  handler: async () => ({}),
};

const writeTool: ToolDefinition = {
  name: "mail_send_message",
  surface: "mail",
  description: "",
  inputSchema: z.object({}),
  mutating: true,
  handler: async () => ({}),
};

describe("writeGuard", () => {
  it("allows read tools by default", () => {
    expect(isToolAllowed(readTool, baseConfig)).toBe(true);
  });

  it("blocks write tools by default", () => {
    expect(isToolAllowed(writeTool, baseConfig)).toBe(false);
    expect(() => assertAllowed(writeTool, baseConfig)).toThrow(WriteBlockedError);
  });

  it("unlocks writes when enableWrites is true", () => {
    expect(isToolAllowed(writeTool, { ...baseConfig, enableWrites: true })).toBe(true);
  });

  it("unlocks writes per-surface", () => {
    expect(
      isToolAllowed(writeTool, { ...baseConfig, perSurfaceWrites: { mail: true } }),
    ).toBe(true);
  });

  it("honors per-surface disablement of specific tools", () => {
    expect(
      isToolAllowed(readTool, {
        ...baseConfig,
        disabledTools: new Set(["mail_list_messages"]),
      }),
    ).toBe(false);
  });
});
