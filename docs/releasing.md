# Releasing

The plan is to **publish `0.1.0` to npm once** and not maintain it after that
(see the README's *Maintenance status*). No container image is published: an
unmaintained image accumulates base-image vulnerabilities, and the `Dockerfile`
is there for anyone who wants one.

Pushing a tag `vX.Y.Z` runs [`.github/workflows/release.yml`](../.github/workflows/release.yml):

1. **verify**: the tag must equal `v` + the `package.json` version, then lint,
   typecheck, tests with coverage, build, production `npm audit` and
   `npm pack --dry-run`. Nothing is published if any of these fail.
2. **npm**: `npm publish --provenance --access public`, under the dist-tag
   `latest`, or `next` for a pre-release (a tag with a hyphen, like `v0.2.0-rc.1`).

---

## Publishing 0.1.0

Steps 1 to 3 need the owner's npm account; step 1 matters even if nothing is
ever published.

1. **Claim the scope.** The README already shows `npx -y @microsoft365-mcp/server`.
   Until someone owns `@microsoft365-mcp`, anyone could register it and publish
   a package under that exact name, which people following the README would
   then run. Sign in at npmjs.com and create the organisation `microsoft365-mcp`
   on the free plan at <https://www.npmjs.com/org/create>; there is no CLI
   command for this. The unscoped name `microsoft365-mcp-server` belongs to an
   unrelated package: never point anyone at it.
2. **Give the workflow a way to publish.** Either:
   - *Trusted publishing (preferred, no stored secret):* in the package settings
     on npmjs.com, add this repository and `release.yml` as a trusted publisher.
     It needs npm 11.5.1 or later in the workflow: move `setup-node` to Node 24
     (or add `npm install -g npm@latest`) and drop `NODE_AUTH_TOKEN`. On a
     brand-new package this may only be possible after the first publish, in
     which case use a token once.
   - *Token:* create a granular access token that can publish to the
     `@microsoft365-mcp` scope, add it as the repository secret `NPM_TOKEN`, and
     delete the token after the release.
3. **Run the live tests** against a real tenant (see
   [`scripts/live/`](../scripts/live/README.md); the sandbox ids are in the
   handover), then tag and push: `git tag v0.1.0 && git push origin v0.1.0`.
   If `verify` fails, nothing was published: fix it, delete the tag, tag again.
4. **Afterwards:** remove the "Not published yet" banner from the README's `npx`
   quickstart, and check the package page shows the provenance badge.

## When it should stop being used

- `npm deprecate @microsoft365-mcp/server "Unmaintained; see <fork or link>"`
  shows a warning on every install without breaking existing users.
- Archive the GitHub repository (Settings → Archive): it becomes read-only, and
  issues and pull requests close.
- `npm unpublish` only works within 72 hours of a release and frees nothing
  (the version number can never be reused), so deprecation is the right tool
  after that.
