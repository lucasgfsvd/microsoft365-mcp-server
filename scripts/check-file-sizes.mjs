#!/usr/bin/env node
/**
 * Mechanical line-count gate: keeps files small enough to hold in your head.
 *
 * Per-type limits, plus a baseline recording the files already over their limit
 * as ceilings: CI is green today, but no file may grow past its ceiling and no
 * new file may land over its limit. Split along a natural seam instead of
 * raising a ceiling.
 *
 *   node scripts/check-file-sizes.mjs            # check (CI)
 *   node scripts/check-file-sizes.mjs --update   # rewrite the baseline (on purpose)
 */
import { existsSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";

const ROOT = process.cwd();
const BASELINE_PATH = join(ROOT, "scripts", "file-size-baseline.json");
const SCAN_DIRS = ["src", "test", "scripts"];
const IGNORE_DIRS = new Set(["node_modules", "dist", "coverage"]);
const LIMIT = 300;

function walk(dir, acc = []) {
  for (const entry of readdirSync(dir)) {
    if (IGNORE_DIRS.has(entry)) continue;
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) walk(full, acc);
    else if (/\.(ts|mjs|js)$/.test(entry)) acc.push(full);
  }
  return acc;
}

const counts = {};
for (const dir of SCAN_DIRS.filter((d) => existsSync(join(ROOT, d)))) {
  for (const file of walk(join(ROOT, dir))) {
    counts[relative(ROOT, file).split("\\").join("/")] = readFileSync(file, "utf8").split("\n").length;
  }
}

if (process.argv.includes("--update")) {
  const over = Object.entries(counts)
    .filter(([, lines]) => lines > LIMIT)
    .sort(([a], [b]) => a.localeCompare(b));
  writeFileSync(BASELINE_PATH, JSON.stringify(Object.fromEntries(over), null, 2) + "\n");
  console.log(`Wrote baseline: ${over.length} over-limit file(s) recorded.`);
  process.exit(0);
}

const baseline = existsSync(BASELINE_PATH) ? JSON.parse(readFileSync(BASELINE_PATH, "utf8")) : {};
const violations = Object.entries(counts).flatMap(([rel, lines]) => {
  const ceiling = baseline[rel] ?? LIMIT;
  if (lines <= ceiling) return [];
  return [`  ${rel}: ${lines} lines (> ${rel in baseline ? `baseline ceiling ${ceiling}` : `limit ${LIMIT}`})`];
});

if (violations.length) {
  console.error("File-size gate FAILED:\n" + violations.join("\n"));
  console.error(
    "\nSplit into smaller, single-concern modules. If a baselined file shrank, or a larger " +
      "file is genuinely intended, run: node scripts/check-file-sizes.mjs --update",
  );
  process.exit(1);
}
console.log(`File-size gate passed (${Object.keys(counts).length} files scanned).`);
