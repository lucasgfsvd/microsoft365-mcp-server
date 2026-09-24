import { describe, it, expect } from "vitest";
import { itemByPath } from "../src/tools/files/scope.js";

// Regression, found live: "/" became `root:/:`, which Graph rejects, so
// creating a folder at (or listing) the drive root failed.
describe("itemByPath", () => {
  it("addresses the root as root, however it is written", () => {
    for (const p of [undefined, "", "/", "//"]) expect(itemByPath("/me/drive", p)).toBe("/me/drive/root");
  });

  it("uses the path form for anything else, without a trailing slash", () => {
    expect(itemByPath("/me/drive", "/Reports/2026")).toBe("/me/drive/root:/Reports/2026:");
    expect(itemByPath("/me/drive", "/Reports/")).toBe("/me/drive/root:/Reports:");
    expect(itemByPath("/drives/d", "Reports")).toBe("/drives/d/root:/Reports:");
  });
});
