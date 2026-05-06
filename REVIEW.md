# Project Review: microsoft365-mcp-server

## Review Metadata

- **Date**: 2026-05-06
- **Project**: `@microsoft365-mcp/server` v0.1.0 — Model Context Protocol server wrapping Microsoft Graph for Mail, Calendar, Files, Teams, OneNote, Excel, Word, PowerPoint, Tasks, Contacts (~92 tools across 10 surfaces).
- **Tech Stack**: TypeScript 5 (strict, `noUncheckedIndexedAccess`), Node 20+, Zod, `@azure/identity` + `@azure/msal-node`, `@microsoft/microsoft-graph-client`, `@modelcontextprotocol/sdk`, PizZip + `docx` + `pptxgenjs` + ExcelJS for OOXML, Pino logging, Commander CLI. Built with tsup, tested with Vitest, packaged as distroless Docker.
- **Maturity**: Early alpha. Single commit on `main`. Per the README, only ~15 unit tests + one tool (`mail_list_folders`) verified end-to-end against a real tenant.
- **Reviewer Perspectives**: 8/8 expert panels.
- **Severity calibration**: Early-alpha — missing tests, retry logic, monitoring rated 🟡 not 🔴. Security-affecting gaps held to standard severity.

## Executive Summary

This is a remarkably mature alpha. The architecture has clear boundaries (one folder per surface, shared `auth/` + `graph/` + `util/`), the write-protection model uses defense in depth (list-time filtering plus handler-time `assertAllowed`), and the README is unusually honest about what's tested vs. only code-reviewed. For 0.1.0, the discipline visible in [src/server.ts](src/server.ts), [src/util/writeGuard.ts](src/util/writeGuard.ts), and [src/util/logger.ts](src/util/logger.ts) is well above average.

The risks are concentrated in the OOXML mutation paths — [src/tools/word/index.ts](src/tools/word/index.ts) and [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts) string-replace XML directly, are the longest files, and have zero test coverage. The README acknowledges this as the riskiest surface; the QA gap is the largest gap in the project. Secondary concerns: 16 unmerged Dependabot PRs accumulating drift, an over-broad default scope set that contradicts the README's "least privilege" guidance, and a few small documentation/UX fit-and-finish items (broken link to `docs/tools.md`, missing `--version` flag, missing `.env.example`).

For an MVP/alpha aimed at individual users on personal accounts, the current state is shippable. Before tagging `0.2.0` I would want OOXML fixture tests, the Dependabot backlog triaged, and the default scope set tightened.

## Findings Summary Dashboard

| Perspective         | 🔴 | 🟠 | 🟡 | 🟢 | ℹ️ |
| ------------------- | -: | -: | -: | -: | -: |
| System Architect    |  0 |  0 |  2 |  1 |  2 |
| Senior Developer    |  0 |  2 |  4 |  3 |  0 |
| Cybersecurity       |  0 |  2 |  4 |  2 |  0 |
| QA Engineer         |  0 |  1 |  4 |  1 |  0 |
| Product Manager     |  0 |  0 |  0 |  2 |  1 |
| UX Designer         |  0 |  0 |  1 |  1 |  1 |
| Business Analyst    |  0 |  0 |  0 |  2 |  1 |
| Legal Advisor       |  0 |  0 |  1 |  2 |  0 |
| **Total**           |  **0** |  **5** | **16** | **14** |  **5** |

## Detailed Findings by Perspective

### 1. System Architecture Review

**Strengths**

- Clean modular monolith. Each surface is isolated under [src/tools/](src/tools/), shared concerns (`auth/`, `graph/`, `util/`) are pulled out, and dependencies point inward through the `ToolContext` injected by [src/server.ts:113](src/server.ts:113).
- `ToolRegistry` ([src/tools/registry.ts](src/tools/registry.ts)) rejects duplicate names — collisions fail fast at startup instead of silently overriding.
- Two-layer write protection: `tools/list` filters by `isToolAllowed` ([src/server.ts:78-86](src/server.ts:78)), and the call handler re-checks via `assertAllowed` ([src/server.ts:111](src/server.ts:111)). Defense in depth.
- Auth warmup at startup ([src/server.ts:62-77](src/server.ts:62)) so device-code prompts fire before the first tool call rather than blocking it.

**Findings**

- 🟡 **OOXML mutation has no abstraction layer.** [src/tools/word/index.ts:79-92](src/tools/word/index.ts:79) and [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts) both reach directly into PizZip + string-replace XML inside their tool handlers. Factor a `src/ooxml/` module so each surface uses one hardened "open → mutate → save" helper. Hardening once is cheaper than hardening per-handler.
- 🟡 **No retry/backoff layer around the Graph client.** README explicitly notes "no built-in retry on Graph 429/503." This belongs around [src/graph/client.ts:6-12](src/graph/client.ts:6) (a middleware on `GraphClient.initWithMiddleware`), not inside each tool. Already on the roadmap.
- 🟢 **No metrics or distributed tracing.** Pino-to-stderr ([src/util/logger.ts:23](src/util/logger.ts:23)) is the only observability. For an alpha, fine; before any hosted variant, add request-count and Graph-call latency.
- ℹ️ **`isomorphic-fetch` is imported as a side-effect.** [src/graph/client.ts:1](src/graph/client.ts:1). On Node 20+, `fetch` is built-in. The dep is technically removable.
- ℹ️ **Excel sessions can leak.** `excel_create_session` ([src/tools/excel/index.ts:67-81](src/tools/excel/index.ts:67)) returns a session id; if the caller forgets `excel_close_session`, the session lingers until Graph times it out (~7 min). Acceptable, but document the lifetime expectation in the tool description.

### 2. Senior Developer Review

**Strengths**

- TypeScript `strict: true` plus `noUncheckedIndexedAccess: true` ([tsconfig.json:7-8](tsconfig.json:7)) — strong type discipline. Most tool handlers are narrow and type-safe.
- Zod input schemas at every tool boundary; consistent use of `PaginationInput` ([src/util/schema.ts:3-9](src/util/schema.ts:3)) keeps surface signatures uniform.
- `normalizeGraphError` ([src/graph/errors.ts:15-28](src/graph/errors.ts:15)) gives every Graph failure a uniform shape that the dispatcher in [src/server.ts:118-125](src/server.ts:118) maps to MCP `isError` responses.
- Logger redacts common token fields with deep-path globs ([src/util/logger.ts:18-21](src/util/logger.ts:18)).

**Findings**

- 🟠 **`escapeXml` does not escape quotes.** [src/tools/word/index.ts:407-409](src/tools/word/index.ts:407) and [src/tools/powerpoint/index.ts:446-448](src/tools/powerpoint/index.ts:446) escape only `&<>`. Today every insertion lands inside element text where this is sufficient, but if a future change ever inserts user input into an XML attribute (e.g., `<a:rPr lang="..."/>`), single/double quotes can break the document or inject attributes. Add `'` → `&apos;` and `"` → `&quot;` defensively.
- 🟠 **Broken link to `docs/tools.md`.** Referenced from [README.md:42](README.md:42) ("the full reference is in docs/tools.md") and [CONTRIBUTING.md:26](CONTRIBUTING.md:26) ("Update docs/tools.md"). The file does not exist. Either generate it (a small script over `allTools()` would do it) or remove the references.
- 🟡 **Long surface files.** [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts) (449 lines), [src/tools/excel/index.ts](src/tools/excel/index.ts) (426 lines), [src/tools/word/index.ts](src/tools/word/index.ts) (410 lines). Above the project's CLAUDE.md soft limit. Each individual tool is fine; what's bloating the file is the OOXML helpers (`downloadDocx`, `mutateDocumentXml`, `escapeXml`, `appendMinimalSlide`). Pulling those into shared OOXML utilities (see Architect finding) splits the concerns naturally.
- 🟡 **No `.env.example` shipped.** README has a comprehensive env-var table at [README.md:631-649](README.md:631), but contributors cloning the repo have no template to copy. Add a short `.env.example` listing all `MCP_*` vars with placeholder values and comments.
- 🟡 **Logger redact list misses `id_token`.** [src/util/logger.ts:3-12](src/util/logger.ts:3) covers `access_token`/`refresh_token`/`client_secret`/`Authorization`. If `id_token` ever lands in a logged object (it currently doesn't, but a careless `logger.info({ tokenResponse })` would expose it), it leaks. Add `id_token`/`idToken` and `client_assertion`.
- 🟡 **Default `MCP_CLIENT_ID` is the public Microsoft client.** [src/config.ts:60](src/config.ts:60) defaults to Microsoft Graph PowerShell's public client (`14d82eec-…`). README warns about this in [README.md:702](README.md:702), but the warning lives in docs, not at runtime. Emit a one-line `logger.warn` at startup when the default is in use ("Using default public client; override MCP_CLIENT_ID for production / audit-log clarity").
- 🟢 **`contacts_update.patch` is `z.record(z.any())`.** [src/tools/contacts/index.ts:58](src/tools/contacts/index.ts:58) bypasses Zod's value validation. Enumerate the patchable fields (`emailAddresses`, `mobilePhone`, `companyName`, etc.) for type safety.
- 🟢 **Dead exported function.** `clearLastDeviceCodePrompt` ([src/auth/index.ts:31-33](src/auth/index.ts:31)) is exported but never imported anywhere. Either wire it into a "logout" flow or delete it.
- 🟢 **Inconsistent return shapes across mail tools.** `mail_send_message` and `mail_reply_message` return `{ ok: true }`, but `mail_create_draft` returns the raw Graph response ([src/tools/mail/index.ts:117-123](src/tools/mail/index.ts:117)). Either standardize on `{ ok: true, ...graphFields }` or document why the draft is special.

### 3. Cybersecurity Review

**Strengths**

- Three-channel device-code prompt UX ([src/server.ts:48-57](src/server.ts:48), [src/server.ts:96-108](src/server.ts:96), [src/auth/index.ts:60-73](src/auth/index.ts:60)) so the sign-in code is visible whether the MCP client renders stderr, MCP logging, or only tool errors.
- Token cache is `chmod 600` on POSIX ([src/auth/tokenCache.ts:18-26](src/auth/tokenCache.ts:18)) and uses the OS keychain on platforms with `keytar` (made `optionalDependencies`).
- CodeQL weekly + on PR ([.github/workflows/codeql.yml](.github/workflows/codeql.yml)) — baseline SAST.
- Distroless nonroot Docker runtime ([Dockerfile:21-35](Dockerfile:21)).
- Release pipeline uses `npm publish --provenance` ([.github/workflows/release.yml:24](.github/workflows/release.yml:24)) — supply-chain attestation.
- Defense in depth on writes (see Architect).

**Findings**

- 🟠 **Default scope set requests everything.** [src/config.ts:6-26](src/config.ts:6) lists ~19 scopes by default — the union of all surfaces, including `Mail.ReadWrite`, `Files.ReadWrite.All`, `ChannelMessage.Send`. This contradicts the README's "least-privilege scopes" claim ([README.md:702](README.md:702)) and the per-tool scope tables. The default should be `User.Read` + `offline_access`; surfaces should require explicit opt-in via `MCP_SCOPES`. Or, at minimum, default to read-only scopes and require additional ones for writes.
- 🟠 **16 unmerged Dependabot PRs.** Visible as `dependabot/*` branches on `origin`: `msal-node 5.1.4`, `commander 14.0.3`, `eslint 10.2.1`, `nock 14.0.13`, `pino 10.3.1`, `pptx-automizer 0.8.1`, `types/node 25.6.0`, `typescript 6.0.3`, `vitest 4.1.5`, `zod 4.3.6`, plus 5 GitHub Actions and 1 Docker base image bump. Several are major versions. Triage and merge in batches; the longer they sit, the more painful the eventual integration.
- 🟡 **`unsafeAllowUnencryptedStorage: true` in cache plugin options.** [src/auth/index.ts:49](src/auth/index.ts:49). On Linux without `keytar`, the token cache falls back silently to plaintext-on-disk. README acknowledges, but for unattended/server deployments consider erroring out rather than silently degrading. At minimum, log a `warn` on first cache write when keytar is absent.
- 🟡 **`chmod` is best-effort with silent failure.** [src/auth/tokenCache.ts:22-26](src/auth/tokenCache.ts:22) catches and ignores chmod errors. On Windows the call is a no-op (Windows ACLs are different), and the comment says "ignore" — but the cache write succeeded despite the permission tightening failing. Should `logger.warn` so an operator notices in stderr.
- 🟡 **`files_*` paths are interpolated unchecked.** [src/tools/files/index.ts:51](src/tools/files/index.ts:51), [src/tools/files/index.ts:111](src/tools/files/index.ts:111), and the `parentPath` parameters in [src/tools/word/index.ts](src/tools/word/index.ts), [src/tools/excel/index.ts](src/tools/excel/index.ts), [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts) get spliced raw into Graph URLs. Graph itself enforces auth, but a path containing `:`, `?`, `#`, or sequences like `..` could construct an unintended URL. Validate `parentPath` and `path` start with `/` and reject `..` segments.
- 🟡 **No `npm audit` in CI.** [.github/workflows/ci.yml](.github/workflows/ci.yml) runs lint + typecheck + test + build but doesn't surface known CVEs in dependencies. Add `npm audit --omit=dev --audit-level=high` (allow it to fail without breaking the build initially, then promote once the backlog is clean).
- 🟢 **OneNote `html` body stored verbatim.** [src/tools/onenote/index.ts:64](src/tools/onenote/index.ts:64) injects user-supplied HTML directly. Self-XSS only — these are the user's own pages — but if pages are ever shared, an attacker who controls the input could ride along. Document the trust boundary in the tool description, or sanitize.
- 🟢 **No Dockerfile `HEALTHCHECK`.** Even a stdio process can expose a basic liveness check (e.g., responding to an MCP `ping` request). Optional; primarily helpful for orchestrators.

### 4. QA Engineering Review

**Strengths**

- 4 test files, ~15 tests, all using clean AAA patterns. [test/config.test.ts:14-27](test/config.test.ts:14) correctly saves and restores env vars in `beforeEach`/`afterEach`. [test/writeGuard.test.ts](test/writeGuard.test.ts) covers all four allow/block paths.
- CI runs lint + typecheck + test + build on every push and PR ([.github/workflows/ci.yml:14-24](.github/workflows/ci.yml:14)), and Docker build runs after tests.
- `nock` is in devDependencies ([package.json:76](package.json:76)) — the integration-test fixture is set up; just unused.

**Findings**

- 🟠 **Zero test coverage on OOXML mutators.** [src/tools/word/index.ts](src/tools/word/index.ts) and [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts) — flagged as the riskiest surface by the README itself — have no tests. A small fixture suite (a 1KB `.docx` and `.pptx` in `test/fixtures/`) running through `word_replace_text`, `word_append_paragraph`, `word_insert_paragraph_at`, `word_delete_paragraph`, `powerpoint_add_slide`, `powerpoint_delete_slide`, `powerpoint_replace_text` would catch >90% of the regressions that the README warns "produce a file that opens but rendered slightly off."
- 🟡 **No nock-based integration tests.** `nock` is installed but unused. One nock test per surface that asserts the constructed Graph URL + method + body for a representative read and write would catch most "we changed the URL template and didn't notice" regressions.
- 🟡 **No coverage threshold.** [vitest.config.ts:7](vitest.config.ts:7) configures coverage reporters (`text`, `lcov`) but no `thresholds` field. Coverage is reported but never enforced. CI doesn't upload coverage either. Add a low starting threshold (50% on `src/util/`, `src/graph/`, `src/tools/registry.ts`, `src/auth/tokenCache.ts`) and ratchet up.
- 🟡 **`fetchPage` pagination helper untested.** [src/graph/pagination.ts](src/graph/pagination.ts) handles the `nextLink` round-trip and search-string escaping ([src/graph/pagination.ts:22](src/graph/pagination.ts:22): `replace(/"/g, '\\"')`). Both branches are easy to exercise with a nock test.
- 🟡 **Tool dispatcher untested.** [src/server.ts:88-126](src/server.ts:88) — unknown tool name, schema validation failure, `WriteBlockedError` formatting, the warmup-pending fallback at [src/server.ts:97-108](src/server.ts:97). All easily testable.
- 🟢 **No smoke test on built `dist/index.js`.** CI's Docker build verifies it compiles into a runnable image, but `node dist/index.js --help` (or `--version`, once it exists) exiting 0 in CI would catch shebang/banner issues.

### 5. Product Management Review

**Strengths**

- Honest alpha framing ([README.md:46-56](README.md:46)) — "what's tested" section explicitly distinguishes verified-end-to-end vs. code-reviewed-only. Builds trust with security-aware users.
- 92 tools across 10 surfaces is a competitive surface area for an open-source MCP server.
- README has three Quickstarts (local build, npx-coming-soon, Docker-coming-soon) tailored to different audiences.
- Rollout-to-a-team section ([README.md:262-324](README.md:262)) is unusually thorough — Model A (delegated, per-user) vs. Model B (application, automation), with a "what IT does once" / "what each user does once" breakdown that reads like a vendor doc.
- Cross-surface example prompts ([README.md:448-461](README.md:448)) demonstrate the differentiator (orchestrating multiple Microsoft surfaces in one prompt) instead of just listing tools.

**Findings**

- 🟢 **No runtime tool discovery for operators.** Users running the server have no way to see which tools loaded except by listing them through the MCP client. A `--list-tools` CLI flag that prints the surface/name/scopes table (and which are filtered out by current config) would help operators sanity-check their config.
- 🟢 **`docs/tools.md` referenced but missing.** Already counted under Senior Developer; from a PM angle, this is also incomplete onboarding — the README points users to a missing doc.
- ℹ️ **No telemetry.** Privacy-positive (the server never reports back to anyone) but means the maintainer can't answer "is feature X used?" — fine for a local OSS tool; revisit if a hosted variant ever ships.

### 6. UX Design Review

The server has no UI. UX evaluation applies to CLI / error messages / first-run experience.

**Strengths**

- `WriteBlockedError` ([src/util/writeGuard.ts:18-26](src/util/writeGuard.ts:18)) tells the user the exact env var to flip — no scavenger hunt through docs.
- Three-channel device-code prompt UX (see Cybersecurity strengths) — sign-in code reaches the user regardless of how the MCP client renders log vs. tool-error output.
- Tool errors normalized to `<code>: <message>` ([src/server.ts:122](src/server.ts:122)) — short, stable, pattern-matchable from the client.

**Findings**

- 🟡 **No `--version` flag.** [CONTRIBUTING.md:37](CONTRIBUTING.md:37) tells users to run `npx @microsoft365-mcp/server --version` for bug reports, but [src/config.ts:35-49](src/config.ts:35) never calls `commander.version(...)`. Add `program.version("0.1.0")` (or read from package.json) so the documented bug-report flow works.
- 🟢 **Generic `--help` text.** Commander's default `--help` lists options but doesn't tell the user how to set `MCP_CLIENT_ID` or what the env var equivalents are. A short `--help-config` (or a banner at the top of `--help`) showing env-var equivalents would help first-run debugging.
- ℹ️ **(Positive)** Three-channel device-code prompt UX is a thoughtful pattern worth preserving across future flows.

### 7. Business Analysis Review

**Strengths**

- MIT licensed ([LICENSE](LICENSE)). Standard, permissive, low friction for adopters.
- Repository metadata complete: `repository`, `bugs`, `homepage`, `keywords` all filled in [package.json:19-26](package.json:19) — surfaces well in npm/GitHub search once published.
- Roadmap documented in README ([README.md:737-746](README.md:737)) — adopters can see what's coming.

**Findings**

- 🟢 **No "not affiliated with Microsoft" disclaimer.** "Microsoft 365" is a Microsoft trademark; the project name uses it (`@microsoft365-mcp/server`). Microsoft is generally permissive for clearly third-party tooling, but a one-line README footer ("Microsoft, Microsoft 365, and Office are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.") is the standard safe move.
- 🟢 **No `--logout` CLI flag.** Documented "log out" path is "delete `~/.microsoft365-mcp/tokencache.json`" ([docs/troubleshooting.md:21-24](docs/troubleshooting.md:21)). A `--logout` flag would be friendlier and could also clear the keytar entry.
- ℹ️ **Bus factor = 1.** Single maintainer in [package.json:18](package.json:18) and SECURITY.md. Common at v0.1 but worth flagging — if the project gets traction, recruit a co-maintainer before depending on it for production workflows.

### 8. Legal & Compliance Review

**Strengths**

- LICENSE file (MIT) present and declared in package.json.
- SECURITY.md ([SECURITY.md:1-39](SECURITY.md:1)) defines a vulnerability disclosure path (GitHub private advisories or email), 72h ack target, 30d coordinated disclosure for high-severity. Within industry norms.
- CODE_OF_CONDUCT.md adopts Contributor Covenant v2.1 — standard.
- Default client ID is Microsoft's own public "Graph Command Line Tools" app — no licensing question; Microsoft permits its use.

**Findings**

- 🟡 **No NOTICE / third-party attribution file.** With ~17 runtime dependencies (incl. `@microsoft/microsoft-graph-client`, MIT/Apache mix), an attribution file is a defensive add that some downstream commercial users will require. `npx license-checker --production --summary` will scaffold one. Add to the release pipeline so it stays current.
- 🟢 **No data-flow disclosure.** Server is local and only talks to Microsoft Graph (under user auth) and the MCP client. The README's prompt examples ("Summarize my unread emails") imply Graph data flows through the user's AI client. The user — not this server — is responsible for what their AI client does with that data, but a one-paragraph "what data flows where" note in the Security section would set the disclosure baseline.
- 🟢 **Trademark disclaimer absent.** Already counted under Business Analyst.

> **Disclaimer**: This review identifies areas for legal attention but does not constitute legal advice. Consult qualified counsel for jurisdiction-specific compliance.

## Cross-Cutting Concerns

- **OOXML mutation strategy is fragile and untested.** Touches Architect (boundary), Senior Developer (regex correctness, escape coverage, file size), QA (zero coverage on the riskiest surface), Cybersecurity (escape coverage). One coordinated workstream (factor `src/ooxml/`, harden `escapeXml`, add fixture tests) addresses all four.
- **Default scope breadth contradicts the documented stance.** The README, [SECURITY.md](SECURITY.md), and the per-tool scope tables all push "least privilege"; the runtime default does the opposite. Tightening the default to `User.Read` + `offline_access` (with users explicitly opting into surfaces via `MCP_SCOPES`) realigns the runtime with the docs.
- **Dependency hygiene.** 16 Dependabot PRs sitting in the queue are simultaneously a Senior Developer concern (drift) and a Security concern (CVE exposure window). Triage in one batch.
- **Test pyramid is bottom-empty for the riskiest paths.** Architect, Senior Developer, and QA all converge on the same fix: fixture-based tests for the OOXML mutators, nock-based tests for Graph URL construction.
- **Documentation references a missing file.** `docs/tools.md` is linked from README + CONTRIBUTING but doesn't exist. Either generate from `allTools()` or remove the references.

## Prioritized Action Plan

### Immediate (High - This Sprint)

1. **Tighten `escapeXml` to also escape `'` and `"`.** ([src/tools/word/index.ts:407](src/tools/word/index.ts:407), [src/tools/powerpoint/index.ts:446](src/tools/powerpoint/index.ts:446)). Two lines. Defensive against future attribute-context insertions. *(Senior Developer / Cybersecurity, ~15 min.)*
2. **Triage and merge the Dependabot backlog.** Major versions (eslint 10, typescript 6, vitest 4, commander 14, msal-node 5.1.4, zod 4.3.6, types/node 25.6.0) deserve a typecheck + test pass; minor/patch can go in batches. *(Cybersecurity / Senior Developer, ~half day.)*
3. **Tighten default scopes in [src/config.ts:6-26](src/config.ts:6).** Default to `User.Read` + `offline_access` only; require explicit `MCP_SCOPES` for surfaces. Add a startup `logger.info` showing the active scope set. *(Cybersecurity, ~1 hour.)*
4. **Fix the broken `docs/tools.md` link** by generating the file from `allTools()` (a 30-line script using `zod-to-json-schema`) and wiring it into a `docs:tools` npm script that CI runs. *(Senior Developer, ~1 hour.)*
5. **Add `program.version(...)` to [src/config.ts:35](src/config.ts:35)** so the documented bug-report flow (`--version`) works. *(UX, ~5 min.)*

### Short-term (Medium - Next 2-4 Sprints)

6. **OOXML fixture-test suite.** Add `test/fixtures/empty.docx` and `test/fixtures/empty.pptx` (1KB each, generated by `docx` / `pptxgenjs` once and committed). Cover `word_replace_text`, `word_append_paragraph`, `word_insert_paragraph_at`, `word_delete_paragraph`, `word_append_heading`, `powerpoint_add_slide`, `powerpoint_delete_slide`, `powerpoint_replace_text` end-to-end (PizZip → mutate → re-open with PizZip → assert). *(QA / Senior Developer, ~1 day.)*
7. **Factor `src/ooxml/` shared module.** Move `downloadDocx`/`uploadDocx`/`mutateDocumentXml`/`escapeXml`/`escapeRegex`/`appendMinimalSlide` into `src/ooxml/{docx,pptx,zip}.ts`. Word and PowerPoint surfaces drop to ~250 lines each, helpers are testable in isolation. *(System Architect, ~half day.)*
8. **Nock-based integration tests for one read + one write per surface.** ~10 tests total. Catches Graph URL drift and verifies the auth header is present. *(QA, ~half day.)*
9. **Add `npm audit --audit-level=high` to CI** ([.github/workflows/ci.yml](.github/workflows/ci.yml)). Initially `continue-on-error: true`; promote to a hard gate once the queue is clean. *(Cybersecurity, ~15 min.)*
10. **Validate `path`/`parentPath` inputs.** [src/tools/files/index.ts](src/tools/files/index.ts), [src/tools/word/index.ts](src/tools/word/index.ts), [src/tools/excel/index.ts](src/tools/excel/index.ts), [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts). Reject paths that don't start with `/`, contain `..`, or contain `?#`. Single Zod refinement applied through a shared `DrivePath` schema. *(Cybersecurity, ~30 min.)*
11. **Ship `.env.example`** with all `MCP_*` vars and inline comments. *(Senior Developer, ~10 min.)*
12. **Add `--list-tools` and `--logout` CLI flags.** Both pure-CLI, ~30 lines together. *(UX / PM, ~1 hour.)*

### Long-term (Low/Info - Backlog)

13. **Coverage threshold in [vitest.config.ts](vitest.config.ts).** Start at 50% on `src/util/`, `src/graph/`, `src/tools/registry.ts`, `src/auth/tokenCache.ts`; ratchet up.
14. **Retry/backoff middleware on the Graph client** ([src/graph/client.ts](src/graph/client.ts)). On the README roadmap.
15. **Generate NOTICE / third-party-attribution file** from `npm ls --production` + `license-checker`; bake into release pipeline.
16. **Trademark and data-flow disclaimers in README** (one short paragraph each).
17. **Logger redact list:** add `id_token`, `idToken`, `client_assertion` ([src/util/logger.ts:3-12](src/util/logger.ts:3)).
18. **Explicit warn when running with the default public client** ([src/config.ts:60](src/config.ts:60)).
19. **Drop `isomorphic-fetch`** ([src/graph/client.ts:1](src/graph/client.ts:1), [package.json:60](package.json:60)) — Node 20 has fetch built-in.
20. **Delete dead `clearLastDeviceCodePrompt`** or wire it to a logout flow ([src/auth/index.ts:31-33](src/auth/index.ts:31)).

## Appendix

### A. Files Reviewed

**Source**: [src/index.ts](src/index.ts), [src/server.ts](src/server.ts), [src/config.ts](src/config.ts), [src/types.ts](src/types.ts), [src/auth/index.ts](src/auth/index.ts), [src/auth/tokenCache.ts](src/auth/tokenCache.ts), [src/graph/client.ts](src/graph/client.ts), [src/graph/errors.ts](src/graph/errors.ts), [src/graph/pagination.ts](src/graph/pagination.ts), [src/util/writeGuard.ts](src/util/writeGuard.ts), [src/util/logger.ts](src/util/logger.ts), [src/util/schema.ts](src/util/schema.ts), [src/tools/registry.ts](src/tools/registry.ts), [src/tools/index.ts](src/tools/index.ts), [src/tools/mail/index.ts](src/tools/mail/index.ts), [src/tools/calendar/index.ts](src/tools/calendar/index.ts), [src/tools/contacts/index.ts](src/tools/contacts/index.ts), [src/tools/files/index.ts](src/tools/files/index.ts), [src/tools/teams/index.ts](src/tools/teams/index.ts), [src/tools/tasks/index.ts](src/tools/tasks/index.ts), [src/tools/onenote/index.ts](src/tools/onenote/index.ts), [src/tools/excel/index.ts](src/tools/excel/index.ts), [src/tools/word/index.ts](src/tools/word/index.ts), [src/tools/powerpoint/index.ts](src/tools/powerpoint/index.ts).

**Tests**: [test/allTools.test.ts](test/allTools.test.ts), [test/config.test.ts](test/config.test.ts), [test/registry.test.ts](test/registry.test.ts), [test/writeGuard.test.ts](test/writeGuard.test.ts).

**Build / packaging**: [package.json](package.json), [tsconfig.json](tsconfig.json), [tsup.config.ts](tsup.config.ts), [vitest.config.ts](vitest.config.ts), [eslint.config.js](eslint.config.js), [.prettierrc](.prettierrc), [Dockerfile](Dockerfile), [.dockerignore](.dockerignore), [.gitignore](.gitignore).

**CI / governance**: [.github/workflows/ci.yml](.github/workflows/ci.yml), [.github/workflows/codeql.yml](.github/workflows/codeql.yml), [.github/workflows/release.yml](.github/workflows/release.yml), [.github/dependabot.yml](.github/dependabot.yml), [SECURITY.md](SECURITY.md), [CONTRIBUTING.md](CONTRIBUTING.md), [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md), [LICENSE](LICENSE).

**Docs**: [README.md](README.md), [docs/auth.md](docs/auth.md), [docs/permissions.md](docs/permissions.md), [docs/troubleshooting.md](docs/troubleshooting.md).

### B. Methods

- Static read of every source/config/doc file listed above (no execution, no network calls).
- `git status`/`git remote -v`/`git fetch --all`/`git branch -a` for sync verification.
- The 8 expert perspectives were applied per the `/app-project-review` skill specification.
- Severity calibration: alpha/MVP — `missing tests` and similar maturity gaps held to 🟡; security-affecting gaps held to standard severity.

### C. Severity Scale

| Severity | Label    | Meaning                                        |
| -------- | -------- | ---------------------------------------------- |
| 🔴       | Critical | Blocks deployment or poses immediate risk      |
| 🟠       | High     | Significant issue, fix before next release     |
| 🟡       | Medium   | Should fix, but not blocking                   |
| 🟢       | Low      | Nice-to-have improvement                       |
| ℹ️       | Info     | Observation or suggestion                      |
