import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    coverage: {
      reporter: ["text", "lcov"],
      // Floors sit just under the current actuals (87/72/89/89) so a drop is
      // caught, with headroom for normal churn. Scope is the orchestration
      // layer plus the OOXML helpers; surface tool handlers go through the
      // integration tests instead.
      include: ["src/config.ts", "src/util/**", "src/tools/registry.ts", "src/auth/tokenCache.ts", "src/ooxml/**"],
      thresholds: {
        lines: 85,
        functions: 85,
        statements: 85,
        branches: 65,
      },
    },
  },
});
