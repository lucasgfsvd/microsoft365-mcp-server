# Contributing

Thanks for your interest! This project welcomes contributions of all sizes — from typo fixes to new surface coverage.

## Getting set up

```bash
git clone https://github.com/lucasgfsvd/microsoft365-mcp-server
cd microsoft365-mcp-server
npm install
npm run build
npm test

# Enable the repo's git hooks (once per clone)
git config core.hooksPath .githooks
```

### The pre-commit hook

`.githooks/pre-commit` runs [gitleaks](https://github.com/gitleaks/gitleaks) against your **staged** changes and blocks the commit if it finds a secret.

CI scans too, but only after a push — by then a secret is already on GitHub, in a public repository's history. The hook is the only check that runs while it is still on your machine, so please enable it.

```bash
choco install gitleaks      # Windows, from an elevated shell
brew install gitleaks       # macOS
```

Without gitleaks installed the hook warns and lets the commit through, so it never blocks a clone that hasn't set it up. False positive? Add an allowlist entry to [`.gitleaks.toml`](./.gitleaks.toml) rather than reaching for `--no-verify`.

## Branching and commits

- Branch from `main`.
- We use [Conventional Commits](https://www.conventionalcommits.org/) — PR titles and commit messages should follow `type(scope): summary`, e.g. `feat(mail): add batch mark-as-read tool`.
- Every PR needs a passing CI run: lint, typecheck, tests **with coverage thresholds**, build, `npm audit` (blocking on high-severity advisories in production dependencies), a gitleaks secret scan of full history, and the Docker build.

## Adding a new tool

1. Add it to the appropriate surface under `src/tools/<surface>/`. Most surfaces are a single `index.ts`; larger ones are split by sub-concern with `index.ts` as a barrel (see `src/tools/excel/`). If a surface file is approaching ~300 lines, split it along its natural seams rather than appending.
2. Give it a Zod input schema, a concise description, and correct `requiredScopes` / `mutating` flags.
3. Mutating tools must use a verb prefix (`create_`, `update_`, `delete_`, `send_`, `post_`, `reply_`, `upload_`, `share_`, …) so tests pass.
4. Update `docs/tools.md` (tables in the README) and `docs/permissions.md` if scopes changed.
5. Add a test in `test/` — at minimum that the tool registers with a valid schema.

## Adding a new surface

Create `src/tools/<surface>/index.ts`, add it to `src/types.ts` (`Surface` union and `SURFACES`), and register it in `src/tools/index.ts`. Update the README feature matrix, `docs/permissions.md`, and add tests.

## Reporting bugs

Use GitHub issues. Include:
- MCP client name + version
- Server version (`npx @microsoft365-mcp/server --version`)
- Auth mode
- The exact tool call and (redacted) response
- Relevant lines from stderr

## Security issues

See [SECURITY.md](./SECURITY.md) — please don't file public issues for security reports.
