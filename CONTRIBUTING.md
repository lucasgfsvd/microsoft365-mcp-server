# Contributing

Thanks for your interest! This project welcomes contributions of all sizes — from typo fixes to new surface coverage.

## Getting set up

```bash
git clone https://github.com/lucasgfsvd/microsoft365-mcp-server
cd microsoft365-mcp-server
npm install
npm run build
npm test
```

## Branching and commits

- Branch from `main`.
- We use [Conventional Commits](https://www.conventionalcommits.org/) — PR titles and commit messages should follow `type(scope): summary`, e.g. `feat(mail): add batch mark-as-read tool`.
- Every PR needs a passing CI run (lint + typecheck + tests + build + Docker).

## Adding a new tool

1. Add it to the appropriate `src/tools/<surface>/index.ts`.
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
