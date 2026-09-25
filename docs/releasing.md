# Releasing

Pushing a tag `vX.Y.Z` runs [`.github/workflows/release.yml`](../.github/workflows/release.yml):

1. **verify**: the tag must equal `v` + the `package.json` version, then lint,
   typecheck, tests with coverage, build, production `npm audit` and
   `npm pack --dry-run`. Nothing is published if any of these fail.
2. **npm**: `npm publish --provenance --access public`, under the dist-tag
   `latest`, or `next` for a pre-release.
3. **docker**: a multi-arch image (amd64, arm64) pushed to
   `ghcr.io/lucasgfsvd/microsoft365-mcp-server`, tagged `X.Y.Z` and `X.Y`, plus
   `latest` for a full release.

A hyphen marks a pre-release: `v0.2.0-rc.1` publishes to npm under `next` and
does not move Docker's `latest`, so `npx` and `docker pull` users keep getting
the last full release.

---

## Do this now, whether or not you publish

**Claim the npm scope.** The docs already tell people to run
`npx -y @microsoft365-mcp/server`, and no one owns `@microsoft365-mcp` yet.
Until someone does, anyone could register it and publish a package under that
exact name, and people following the README would run it. Creating the
organisation is free and publishes nothing:

```bash
npm login
npm org create microsoft365-mcp
```

Note that the unscoped name `microsoft365-mcp-server` belongs to an unrelated
package (1.2.6 as of September 2026). Keep the scoped name, and never tell
anyone to run the unscoped one.

---

## One-time setup before the first release

- [ ] **npm scope owned** (above).
- [ ] **Publish credential.** Either:
  - *Trusted publishing (preferred, no stored secret):* on npmjs.com, add this
    repository and `release.yml` as a trusted publisher for the package. It
    needs npm 11.5.1 or later in the workflow, so move `setup-node` to Node 24
    (or add `npm install -g npm@latest`) and drop `NODE_AUTH_TOKEN`.
  - *Token:* create a granular access token that can publish to the
    `@microsoft365-mcp` scope, and add it as the repository secret `NPM_TOKEN`.
- [ ] **GHCR visibility.** After the first image push, set the
  `microsoft365-mcp-server` package to public in the GitHub package settings,
  and link it to this repository. GHCR packages start private.
- [ ] **arm64 image.** CI builds and smoke-tests amd64 only. Build arm64 once
  locally (`docker buildx build --platform linux/arm64 .`) before the first
  multi-arch release.
- [ ] **README.** Remove the "coming once we publish" banners from the `npx` and
  Docker quickstarts once the first release is out.

---

## Each release

1. Run the live tests against a real tenant, which catch what unit tests
   cannot (see [`scripts/live/`](../scripts/live/README.md)):
   `office.mjs`, `pim.mjs`, and `onenote.mjs` with a test section.
2. Bump the version without tagging, and commit:
   `npm version X.Y.Z --no-git-tag-version`.
3. Tag and push: `git tag vX.Y.Z && git push origin vX.Y.Z`.
4. Watch the Release workflow. If `verify` fails, nothing was published: fix it,
   delete the tag, and tag again.

## If a release is bad

- **npm:** `npm deprecate @microsoft365-mcp/server@X.Y.Z "reason"` and publish a
  fix. `npm unpublish` only works within 72 hours and frees nothing: the
  version number can never be reused.
- **Docker:** push a fixed tag. `latest` follows the newest full release. Delete
  the bad version from GHCR's package settings if it must not be pulled.
