# Handover

State of play and what to pick up next. Written for whoever continues this work — human or agent.

---

## Where things stand

97 tools across 12 surfaces. 113 tests. CI gates lint, typecheck, coverage thresholds, `npm audit` (blocking, high severity, production deps), a full-history gitleaks scan, and the Docker build.

Recently landed and worth knowing about:

- **Lazy auth.** The server starts silent and signs in only when `auth_sign_in` is called. It used to warm the credential at startup, which minted a device code on *every* client launch that nobody entered, and which expired unused ~15 minutes later. The README's "Auth & token cache" section has the full reasoning.
- **`graph_batch_get`** — up to 20 Graph GETs in one `$batch`. Measured 2.4× faster than four sequential tool calls against a live tenant; the gap widens with more requests.
- **`graph_search`** — one relevance-ranked Microsoft Search query across mail, files, SharePoint, Teams or people.
- **`graph_delta`** — incremental sync for mail, calendar, drive, contacts and To Do; the `deltaLink` goes back to the caller. See [tools.md](./tools.md#-graph--3-tools).

---

## Open defect: token acquisition can fail after an in-process sign-in

**The most valuable thing in this document. Do not "fix" it without a reproduction.**

### Symptom

After completing `auth_sign_in` in a running server, `auth_status` reported `signedIn: true` while every Graph call failed with:

```
AuthenticationRequiredError: Automatic authentication has been disabled.
You may call the authentication() method.
```

Restarting the process fixed it immediately, and the startup log then read `using cached Microsoft 365 credentials`.

### Leading hypothesis

The credential is constructed once in `startServer`, reading `authrecord.json` from disk. On a first-ever sign-in that file does not exist yet, so the credential is built **without** an `authenticationRecord`. `authenticate()` then succeeds and writes the record to disk — but the live credential instance still has none, and a credential with `disableAutomaticAuthentication: true` and no record cannot select an account from the persistent cache silently.

That would explain a restart fixing it: the next start does find the record.

### Why it is not fixed

An attempt was made — a `SwappableCredential` wrapper, so `AuthSession` could replace the credential with one built around the new record. It typechecked, passed every test, then failed a live Graph call. The subsequent "revert" test then *passed* with the same code still in the working tree, because uncommitted changes had followed a branch switch. Same code, same test, opposite results.

So the failure is **intermittent**, and the wrapper was never shown to be either the cause or the cure. It was reverted rather than shipped on a guess.

A later attempt at a standalone reproduction also failed to converge: a script constructing a credential with the same options as `buildCredential` could not acquire a token, while the server could, minutes apart, against the same cache.

### What is known

- Consent is not the problem. Every scope group — including the admin-consent Teams ones, and `.default` — acquires a token silently when the credential has a record.
- Several server processes were running concurrently during every failed observation (a client-connected one plus test instances). **Contention on the OS credential store between processes sharing one cache is untested but plausible.**
- A single server, running alone, has never been observed to fail.

### Suggested approach

1. Reproduce first, with **exactly one** process touching the cache. Delete the token cache directory, start one server, sign in, then call a Graph tool *without* restarting.
2. If it fails there, the hypothesis holds and `SwappableCredential` is the right shape of fix. Reinstate it and verify live, not just in tests.
3. If it does not fail, suspect multi-process cache contention and test that directly: two servers, one cache, interleaved `getToken`.

`status()` now verifies against the credential instead of trusting its cached flag, so the server at least reports this honestly rather than claiming to be signed in while failing.

---

## Delta: what live testing has and has not covered

All five resources were exercised against a live tenant: initial sync, resuming from a `nextLink`, and a follow-up from the `deltaLink` returning nothing new. Page sizing is resource-specific — Outlook resources honour `Prefer: odata.maxpagesize`, drive ignores it and needs `$top` instead.

**Removals, live:** both forms are confirmed. A contact and a OneDrive file were each created, synced, deleted and synced again; each came back in `removed` with its id and reason `deleted`, and not in `changed`. On drive, the parent folder also shows up in `changed`, because a child changing modifies it — expected, not a leak. To Do has no delete tool, so its removals cannot be tested through the server, but it shares the contacts' `@removed` form.

---

## The container image

Verified running. The distroless runtime has no `libsecret`, so the cache plugin fails to load (`libsecret-1.so.0: cannot open shared object file`) and the server falls back to an in-memory token cache. It starts, serves tools and answers `auth_status`. CI now runs the image and requires it to answer `initialize`, so a regression to a static import fails the build.

**Consequence for device-code users:** in the container a sign-in does not survive a restart. Client-credentials mode, the natural fit for a container, needs no cache. Making device-code persist there would mean installing `libsecret` in the runtime image (not available on distroless) or adding a file-based cache for this case.

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
