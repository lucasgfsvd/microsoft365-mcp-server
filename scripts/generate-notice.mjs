#!/usr/bin/env node
// Generate NOTICE — a third-party-attribution file listing every production
// dependency and its declared license. Run via `npm run docs:notice`.
//
// Implementation note: this walks production deps transitively by reading the
// resolved `dependencies` tree from `package-lock.json`. We don't ship anything
// listed only under devDependencies (per the npm package's `files` field), so
// those are intentionally omitted.

import { readFile, writeFile } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");

const pkg = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
const lock = JSON.parse(await readFile(resolve(root, "package-lock.json"), "utf8"));

// Collect every package node_modules entry that's in the production tree.
// package-lock v3 uses the `packages` map keyed by install path. Skip the root
// entry (`""`) and any entry flagged `dev: true`.
const entries = [];
for (const [installPath, meta] of Object.entries(lock.packages ?? {})) {
  if (!installPath || installPath === "") continue;
  if (meta.dev || meta.devOptional) continue;
  // Read the actual installed package.json for the most accurate metadata.
  const pkgJsonPath = resolve(root, installPath, "package.json");
  let installed;
  try {
    installed = JSON.parse(await readFile(pkgJsonPath, "utf8"));
  } catch {
    continue; // bin-only or missing — skip
  }
  entries.push({
    name: installed.name ?? installPath.split("node_modules/").pop(),
    version: installed.version ?? meta.version ?? "?",
    license: licenseFromPkg(installed) ?? meta.license ?? "UNKNOWN",
    homepage:
      typeof installed.homepage === "string"
        ? installed.homepage
        : typeof installed.repository === "string"
          ? installed.repository
          : installed.repository?.url ?? "",
    author: authorString(installed.author),
  });
}

// De-dupe: same name+version may appear under different install paths.
const dedup = new Map();
for (const e of entries) dedup.set(`${e.name}@${e.version}`, e);
const sorted = [...dedup.values()].sort((a, b) => a.name.localeCompare(b.name));

const header = `# Third-party software notices

This file lists the open-source software included in distributions of \`${pkg.name}\` (version ${pkg.version}).

It is auto-generated from \`package-lock.json\` — do not edit by hand. Regenerate with:

\`\`\`bash
npm run docs:notice
\`\`\`

Each entry below is the property of its respective copyright holder, used here under its declared license. Copies of the underlying license texts are available alongside each package in \`node_modules/<name>/LICENSE\` (or equivalent).

`;

const tableHeader = `| Package | Version | License | Source |
| --- | --- | --- | --- |
`;

const rows = sorted
  .map(
    (e) =>
      `| \`${escapePipe(e.name)}\` | ${escapePipe(e.version)} | ${escapePipe(String(e.license))} | ${escapePipe(cleanUrl(e.homepage))} |`,
  )
  .join("\n");

const footer = `

---

Total: ${sorted.length} package${sorted.length === 1 ? "" : "s"}.

Microsoft, Microsoft 365, Outlook, OneDrive, SharePoint, Teams, OneNote, Excel, Word, and PowerPoint are trademarks of Microsoft Corporation. This is an independent open-source project; it is not affiliated with, endorsed by, or sponsored by Microsoft.
`;

await writeFile(resolve(root, "NOTICE.md"), header + tableHeader + rows + footer);
process.stdout.write(`Wrote NOTICE.md (${sorted.length} packages)\n`);

function licenseFromPkg(p) {
  if (typeof p.license === "string") return p.license;
  if (p.license && typeof p.license === "object" && "type" in p.license) return p.license.type;
  if (Array.isArray(p.licenses) && p.licenses[0]?.type) return p.licenses[0].type;
  return undefined;
}

function authorString(a) {
  if (typeof a === "string") return a;
  if (a && typeof a === "object" && "name" in a) return a.name;
  return "";
}

function cleanUrl(s) {
  if (!s) return "";
  return String(s)
    .replace(/^git\+/, "")
    .replace(/^ssh:\/\/git@github\.com\//, "https://github.com/")
    .replace(/\.git$/, "");
}

function escapePipe(s) {
  return String(s).replace(/\|/g, "\\|");
}
