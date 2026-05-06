import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    coverage: {
      reporter: ["text", "lcov"],
      // Conservative starting thresholds — ratchet up as we add tests.
      // Focus on the orchestration layer (config/registry/writeGuard/auth/util)
      // and the OOXML helpers; surface-specific tool handlers are exercised
      // via the integration tests once those land.
      include: ["src/config.ts", "src/util/**", "src/tools/registry.ts", "src/auth/tokenCache.ts", "src/ooxml/**"],
      thresholds: {
        lines: 50,
        functions: 50,
        statements: 50,
        branches: 40,
      },
    },
  },
});
