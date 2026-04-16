import { describe, it, expect } from "vitest";
import { z } from "zod";
import { ToolRegistry } from "../src/tools/registry.js";
import type { ServerConfig, ToolDefinition } from "../src/types.js";

const cfg: ServerConfig = {
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

const mk = (name: string, mutating = false): ToolDefinition => ({
  name,
  surface: "mail",
  description: "",
  mutating,
  inputSchema: z.object({}),
  handler: async () => ({}),
});

describe("ToolRegistry", () => {
  it("registers and lists tools, filtering writes by default", () => {
    const r = new ToolRegistry();
    r.registerAll([mk("a"), mk("b_send", true)]);
    const names = r.list(cfg).map((t) => t.name);
    expect(names).toEqual(["a"]);
    expect(r.list({ ...cfg, enableWrites: true }).map((t) => t.name)).toEqual(["a", "b_send"]);
  });

  it("rejects duplicates", () => {
    const r = new ToolRegistry();
    r.register(mk("a"));
    expect(() => r.register(mk("a"))).toThrow();
  });
});
