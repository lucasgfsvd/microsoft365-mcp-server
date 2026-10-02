# Releasing

The plan is to **publish `0.1.0` to npm once**, now that the "Before 0.1.0"
items in [roadmap.md](./roadmap.md) are done, and not maintain it after that (see the
README's *Maintenance status*). No container image is published: an
unmaintained image accumulates base-image vulnerabilities, and the `Dockerfile`
is there for anyone who wants one.

Pushing a tag `vX.Y.Z` runs [`.github/workflows/release.yml`](../.github/workflows/release.yml):

1. **verify**: the tag must equal `v` + the `package.json` version, then lint,
   typecheck, tests with coverage, build, production `npm audit` and
   `npm pack --dry-run`. Nothing is published if any of these fail.
2. **npm**: `npm publish --provenance --access public`, under the dist-tag
   `latest`, or `next` for a pre-release (a tag with a hyphen, like `v0.2.0-rc.1`).
3. **registry** (full releases only): publishes `server.json` to the official
   MCP Registry, authenticated as the repository owner through GitHub OIDC (no
   secret). The registry checks `mcpName` in the published `package.json`, so
   `server.json`'s name and versions must match it; the job checks that first.
4. **github-release**: a GitHub Release whose notes are the version's
   `CHANGELOG.md` entry. A version with no entry fails here, after npm.

The README ships inside the package and is what npmjs.com shows, so it must be
final in the tagged commit: no "not published yet" wording.

---

## Publishing 0.1.0

0. ~~**The roadmap's "Before 0.1.0" items are done**~~, each with its live test,
   and listed in the `CHANGELOG.md` entry. Done.
1. ~~**Claim the scope.**~~ Done: the npm organisation `microsoft365-mcp` exists,
   owned by `lucasgfsvd`, so no one else can publish `@microsoft365-mcp/server`.
   (The unscoped `microsoft365-mcp-server` belongs to an unrelated package:
   never point anyone at it.)
2. ~~**Turn on two-factor authentication**~~ Done on the account. Optionally also
   *Enable 2FA Enforcement* on the organisation.
3. **Give the workflow a way to publish, once.** npm is retiring tokens that
   bypass 2FA: they lost sensitive account operations in August 2026 and lose
   direct publishing around **January 2027**
   ([GitHub changelog](https://github.blog/changelog/2026-07-08-npm-install-time-security-and-gat-bypass2fa-deprecation/)).
   For a single release before then, a short-lived token is simplest, and
   publishing from CI is what earns the provenance badge (proof the package was
   built from this repository):
   - On npmjs.com, Access Tokens → Generate New Token → *Granular*: publish
     access to the `@microsoft365-mcp` scope only, expiry 7 days, *bypass 2FA*
     allowed (CI cannot type a code).
   - Add it as the repository secret `NPM_TOKEN`.
   - **Delete the token and the secret** as soon as the release is out.

   After January 2027 this route is gone: use *trusted publishing* (OIDC; set up
   in the package's settings once it exists, needs npm 11.5.1+ in the workflow,
   so Node 24 or `npm install -g npm@latest`, and no `NODE_AUTH_TOKEN`), or
   *staged publishing*, where CI stages and a person approves with 2FA. A
   one-off `npm publish --access public` from your own machine also works,
   without provenance.
4. **Make the README final**: it ships in the package and is what npmjs.com
   shows. Remove both "Not published yet" notes and the "Being finished" wording
   in *Maintenance status*. Date the `CHANGELOG.md` heading. If the version is
   not `0.1.0`, bump `package.json`, both versions in `server.json` and the `npx`
   pins. Then **run the live tests** against a real tenant (see
   [`scripts/live/`](../scripts/live/README.md); the sandbox ids are in the
   handover), then tag and push: `git tag v0.1.0 && git push origin v0.1.0`.
   If `verify` fails, nothing was published: fix it, delete the tag, tag again.
5. **Afterwards:** check the npm page (provenance badge, README), the registry
   entry (`https://registry.modelcontextprotocol.io/v0/servers?search=microsoft365`)
   and the GitHub Release, then delete the token and the `NPM_TOKEN` secret.

## When it should stop being used

- `npm deprecate @microsoft365-mcp/server "Unmaintained; see <fork or link>"`
  shows a warning on every install without breaking existing users.
- Archive the GitHub repository (Settings → Archive): it becomes read-only, and
  issues and pull requests close.
- `npm unpublish` only works within 72 hours of a release and frees nothing
  (the version number can never be reused), so deprecation is the right tool
  after that.
