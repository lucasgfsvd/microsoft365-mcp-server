# Handover

State of play and what to pick up next. Written for whoever continues this work — human or agent.

---

## Where things stand

97 tools across 12 surfaces. 113 tests. CI gates lint, typecheck, coverage thresholds, `npm audit` (blocking, high severity, production deps), a full-history gitleaks scan, and the Docker build.

Recently landed and worth knowing about:

- **Per-path token stores.** `MCP_TOKEN_CACHE_PATH` now selects a token store of its own; before, every process shared one whatever the path said.
- **Lazy auth.** The server starts silent and signs in only when `auth_sign_in` is called. It used to warm the credential at startup, which minted a device code on *every* client launch that nobody entered, and which expired unused ~15 minutes later. The README's "Auth & token cache" section has the full reasoning.
- **`graph_batch_get`** — up to 20 Graph GETs in one `$batch`. Measured 2.4× faster than four sequential tool calls against a live tenant; the gap widens with more requests.
- **`graph_search`** — one relevance-ranked Microsoft Search query across mail, files, SharePoint, Teams or people.
- **`graph_delta`** — incremental sync for mail, calendar, drive, contacts and To Do; the `deltaLink` goes back to the caller. See [tools.md](./tools.md#-graph--3-tools).

---

## Sign-in defect: not reproducible, probably a code handed to the wrong process

**Status: closed unless it recurs.** Kept here because the reasoning is what to reuse if it does.

### The symptom

After `auth_sign_in`, `auth_status` said `signedIn: true` while every Graph call failed with `AuthenticationRequiredError: Automatic authentication has been disabled`. A restart fixed it.

### What was tested

- **One process, its own store, first-ever sign-in.** The server started with no auth record, a sign-in was completed within that running process, and it called Graph immediately and again 60 seconds later without restarting. All six calls succeeded; all seven silent token requests after sign-in succeeded. This was the exact scenario the old "credential built without a record" hypothesis blamed, and it does not fail. That matches the library: `authenticate()` keeps the account in the credential's own state (`state.cachedAccount`, `@azure/identity` 4.x `msalClient.js`), so the record on disk only matters to a *new* process.
- **Contention.** Three servers on one signed-in store, 45 interleaved Graph calls, no failures and no lock errors.

`SwappableCredential`, the fix that was tried and reverted, is therefore not needed.

### The explanation that fits

At the time, every server warmed its credential at startup, so every process printed its own device code, and several processes were always running. Entering one process's code signs in *that* process. The process being queried kept no account in its state and failed with exactly this error. `auth_status` said otherwise because of the since-fixed stale-flag bug. A restart helped because the new process read the auth record the *other* process had written. And a lone server never failed because it had only one code.

Both halves are already closed: codes are only issued by an explicit `auth_sign_in`, and `status()` checks the credential. The one remaining edge: a server that is running when *another* process signs in does not pick that up until it restarts. If that ever matters, the fix is to re-read `authrecord.json` when a token request finds no account, not to wrap the credential.

To reproduce it on purpose: start two servers on one fresh `MCP_TOKEN_CACHE_PATH`, call `auth_sign_in` on both, enter only the first code, then call Graph on the second.

---

## Delta: what live testing has and has not covered

All five resources were exercised against a live tenant: initial sync, resuming from a `nextLink`, and a follow-up from the `deltaLink` returning nothing new. Page sizing is resource-specific — Outlook resources honour `Prefer: odata.maxpagesize`, drive ignores it and needs `$top` instead.

**Removals, live:** both forms are confirmed. A contact and a OneDrive file were each created, synced, deleted and synced again; each came back in `removed` with its id and reason `deleted`, and not in `changed`. On drive, the parent folder also shows up in `changed`, because a child changing modifies it — expected, not a leak. To Do has no delete tool, so its removals cannot be tested through the server, but it shares the contacts' `@removed` form.

---

## The container image

Verified running. The distroless runtime has no `libsecret`, so the cache plugin fails to load (`libsecret-1.so.0: cannot open shared object file`) and the server falls back to an in-memory token cache. It starts, serves tools and answers `auth_status`. CI now runs the image and requires it to answer `initialize`, so a regression to a static import fails the build.

**Consequence for device-code users:** in the container a sign-in does not survive a restart. Client-credentials mode, the natural fit for a container, needs no cache. Making device-code persist there would mean installing `libsecret` in the runtime image (not available on distroless) or adding a file-based cache for this case.

---

## `--logout` leaves tokens in the store

`runLogout` removes `authrecord.json` (and a `tokencache.json` that has never existed). The tokens live in the identity plugin's store (`%LOCALAPPDATA%\.IdentityService\<name>` on Windows, keychain/keyring elsewhere) and are not touched. The next start is signed out, so this is not a functional bug, but a refresh token stays on disk until it expires. The README says so. Fixing it means deleting the store per platform, including the keychain entry, and the `.cae` variant the plugin may create.

---

## Other candidates

From the roadmap, roughly in value order:

- **Large-file upload sessions (>4 MB)** — `files_upload` returns 413 above the single-shot limit today. Self-contained and well-specified.
- **Streaming downloads** — downloads are base64 inside the tool result, which is memory-heavy for anything sizeable. MCP Resources would be the idiomatic fix, and are also on the roadmap.
- **Prompt templates** — triage, meeting-prep. Low effort, and the place where multi-surface workflows become discoverable instead of something the caller has to invent.
- **Per-tool retry policy** — currently the SDK default: 3 retries, 3s base delay, honours `Retry-After`.

---

## Working notes

Things that cost time to learn.

**Test against a live tenant, not just the mocks.** `graph_search` was first written with a wrong model of which entity types Graph will combine: `message` + `event` is rejected, even though both are mail-ish. The unit tests encoded the same wrong assumption and passed happily. Only a live call caught it. Mocks verify the model, not reality.

**`dist/` is somebody's running server.** It is gitignored, so rebuilding on a feature branch silently replaces the binary a connected MCP client is using. Build in a worktree, or expect to break the thing you are testing against.

**Driving the server in tests.** Spawn it, write JSON-RPC lines to stdin, read responses from stdout: `initialize`, then `notifications/initialized`, then `tools/call`. Read stdout as UTF-8 explicitly — real mail subjects contain non-ASCII, and a default Windows codec will throw mid-run.

**Configuration.** A desktop client may rewrite its own config file from memory on exit, silently discarding edits made while it was running. Edit those files with the client fully closed. Where more than one config file exists, confirm which the process actually reads before assuming: check the client id in the startup log, or read the `app_displayname` claim out of a Graph download URL.

**The pre-commit hook** runs gitleaks against staged changes and blocks the commit. Enable it once per clone with `git config core.hooksPath .githooks`. Without gitleaks installed it warns and lets the commit through, so it never blocks a fresh clone.
