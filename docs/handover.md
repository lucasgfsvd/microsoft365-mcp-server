# Handover

State of play and what to pick up next. Written for whoever continues this work — human or agent.

---

## Where things stand

102 tools across 12 surfaces, three prompt templates, and three resource types. 243 unit tests. CI gates lint, typecheck, coverage thresholds, `npm audit` (blocking, high severity, production deps), a full-history gitleaks scan, and builds *and starts* the Docker image.

**Live coverage: all 102 tools pass against a real tenant** (`scripts/live/`), with Word, PowerPoint and Excel output also opened by independent parsers. No open defects are known.

What the server can now do that it could not before, grouped by concern (git has the history):

- **Sync.** `graph_delta` gives incremental changes, deletions included, for mail, calendar, drive, contacts and To Do; the caller keeps the `deltaLink`. `graph_search` and `graph_batch_get` cover cross-surface search and parallel reads.
- **Large files.** Uploads over 4 MB use Graph upload sessions (all seven write paths, Office tools included); `files_upload` with `localPath` streams a file from `MCP_UPLOAD_DIR` a chunk at a time (verified live at 12 MB), confined to that folder by its real path. Uploads and downloads to disk report MCP progress and stop on cancellation (`ctx.progress`, `ctx.signal`), verified live with a 48 MB upload cancelled partway: no file, server still answering. Downloads over 5 MB stream to `MCP_DOWNLOAD_DIR` in constant memory instead of entering the conversation. `MCP_MAX_MESSAGE_MB` (default 64) bounds a single MCP message.
- **Safety.** A `POST` from a mutating tool is retried only on 429, so a send is never repeated. Pre-authenticated `downloadUrl` links are stripped from every result. `--logout` removes this app's tokens from the store, leaving other apps' alone. Sign-in is lazy and explicit.
- **Portability.** The token-cache plugin loads lazily, so the server runs where `libsecret` is missing (headless Linux, the distroless image); there `MCP_TOKEN_CACHE_KEY` gives it an encrypted file cache, and without it tokens are held in memory. `MCP_TOKEN_CACHE_PATH` gets its own token store on Windows and with the Linux file fallback.
- **Workflows.** `daily-brief`, `inbox-triage` and `meeting-prep` prompts, with Graph queries computed server-side and verified live.
- **Attachable context.** Mail, files and OneNote pages as MCP resources (`src/resources/`): recent items in one `$batch`, any item by `m365://` URI, returned as text (Word and PowerPoint extracted, Excel as CSV per sheet). Gated on the matching read tool, like the prompts.

**Next up: the second "Before 0.1.0" round in [roadmap.md](./roadmap.md).** A review of what users ask for in chat and what companies need added 37 tools (139 in all), standard annotations, progress and cancellation, app-only mode for a named user, shared mailboxes, and files the model can read (PDFs as text, images as images). Their live tests are `scripts/live/features-*.mjs`, `shared.mjs` and `apponly.mjs`; the last needs an app registration with admin consent. What has passed live and what is pending is in the roadmap table. **Outlook attachment upload sessions answer 200 while they want more** (drive sessions answer 202); taking that 200 as done silently dropped every attachment over 3 MB until the live pass caught it. Publishing waits for the owner; the release is prepared ([releasing.md](./releasing.md)). After `0.1.0` the project is left unmaintained, by the owner's choice.

**Scopes that need an admin are opt-in.** Asking for a scope no admin has approved fails the whole sign-in, so `Place.Read.All` and `User.Read.All` are requested only with `MCP_ENABLE_ADMIN_SCOPES`, and the tools needing them (`needsAdminConsent`) are hidden otherwise. Keep the default scope set user-consentable. Adding any scope to it makes every existing user consent again at their next sign-in.

**`/me` is rewritten in one place** (`src/graph/targetUser.ts`): to `/users/{MCP_USER}` in app-only mode and to `/users/{mailbox}` for a `mailbox` argument, for direct requests and `$batch` sub-requests alike. New tools should keep addressing `/me`; tools Graph refuses app-only belong in `DELEGATED_ONLY` (`src/tools/appOnly.ts`).

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

All three halves are closed: codes are only issued by an explicit `auth_sign_in`, `status()` checks the credential, and a server running when *another* process signs in now picks that up. A credential takes its account from the record it is built with and cannot be given a new one, so `ReloadingCredential` (`src/auth/reloadingCredential.ts`) delegates to it and, when a token request finds no account, re-reads `authrecord.json` and rebuilds the credential if another process has written a new record. It reloads only on that error, and never twice for the same record.

To reproduce it on purpose: start two servers on one fresh `MCP_TOKEN_CACHE_PATH`, call `auth_sign_in` on both, enter only the first code, then call Graph on the second.

---

## Delta: what live testing has and has not covered

All five resources were exercised against a live tenant: initial sync, resuming from a `nextLink`, and a follow-up from the `deltaLink` returning nothing new. Page sizing is resource-specific — Outlook resources honour `Prefer: odata.maxpagesize`, drive ignores it and needs `$top` instead.

**Removals, live:** both forms are confirmed. A contact and a OneDrive file were each created, synced, deleted and synced again; each came back in `removed` with its id and reason `deleted`, and not in `changed`. On drive, the parent folder also shows up in `changed`, because a child changing modifies it — expected, not a leak. To Do removals have not been run through delta live (`todo_delete_task` now makes that possible); they share the contacts' `@removed` form.

---

## The container image

Verified running. The distroless runtime has no `libsecret`, so the OS-store plugin fails to load (`libsecret-1.so.0: cannot open shared object file`). It starts, serves tools and answers `auth_status`. CI runs the image and requires it to answer `initialize`, so a regression to a static import fails the build.

**Device-code sign-in persists with `MCP_TOKEN_CACHE_KEY`.** The tokens then live in `MCP_TOKEN_CACHE_PATH` (`/data/tokencache.json` in the image, on a volume), encrypted with AES-256-GCM under a scrypt-stretched key (`src/auth/encryptedFile.ts`). The cache plugin (`src/auth/encryptedCache.ts`) talks to MSAL's cache-plugin interface directly, because `@azure/msal-node-extensions` cannot be imported there at all (its index loads keytar). A lock file serialises servers sharing the file. Without the key the container still holds tokens in memory only. Client-credentials mode needs no cache.

---

## Other candidates

From the roadmap, roughly in value order:

- **The live sandbox** is reusable for pre-release runs: a private team `mcp-live-test` (only the account owner; group/team id `e118e4d7-1a8c-4238-9f47-05dffdfbee9f`, channel `tests`), a Planner plan `mcp-live-test` in it (`ethENaSyq0qyllQYJelKF5gAER_m`), and a OneNote notebook `mcp-live-test` with a section `tests` (`1-041ccae9-c331-409f-80f8-ddcc828c3d78`). Deleting the team needs Group.ReadWrite.All, which the app does not have: do it in Teams.
- **Subscriptions (webhooks)**: designed, not built; see [webhooks.md](./webhooks.md). The public endpoint is the smaller problem: a stdio server exists only while a client runs, and nothing wakes a model when data changes. Recommendation: keep `graph_delta`; if real-time is ever needed, use Azure Event Hubs delivery (no inbound endpoint, holds messages while the server is off). Needs an owner decision on consumer, Azure subscription and scope.
- **Publishing** is prepared and deliberately waiting for the roadmap: the npm scope `@microsoft365-mcp` is claimed (org owned by `lucasgfsvd`, 2FA on), `NPM_TOKEN` is set, and the release workflow publishes to npm, then the MCP Registry (`server.json`, `mcpName`, GitHub OIDC), then a GitHub Release from `CHANGELOG.md`. npm retires 2FA-bypass tokens for publishing around January 2027; if the release slips past that, see [releasing.md](./releasing.md).

---

## Working notes

Things that cost time to learn.

**Graph answers some writes with 204 No Content** (a Planner PATCH, for one). The tool then resolves to `undefined`, which `serializeResult` now turns into `{"ok": true}`; before, it produced an invalid MCP message and the client reported a failure for an action that had succeeded. **And `/me/planner/plans` omits group plans**, which is to say nearly all of them: `listPlans` (`src/graph/planner.ts`) also asks each joined team. Both were invisible to the unit tests and found by the sandbox run.

**Graph reads `createReply`'s `comment` as HTML**: line breaks collapse and markup passes through (seen live). `mail_create_reply_draft` therefore creates the reply empty and writes the text into the draft's body itself, escaped for that body's format. **Graph cannot delete a OneNote notebook**; each one is a folder under `/Notebooks` in the owner's OneDrive, which is how the live test removes the one it creates.

**Marketing mail is mostly padding.** Newsletters fill their preview line with zero-width non-joiners, combining grapheme joiners and soft hyphens: 41% of one real message's text body, and 228 of its 255-character `bodyPreview`. `tidyText` (`src/util/text.ts`) strips them from mail resources and from `mail_get_message`'s plain-text body (an HTML body is left as sent, since collapsing its whitespace could change how it renders). `serializeResult` tidies every `bodyPreview` at any depth, so list, search, delta and raw batch results are covered without each tool opting in.

**Live tests live in `scripts/live/`** and are the first thing to rerun after touching a tool; see its README for what they write and clean up. Index arguments are 1-based throughout (`slideIndex`, `paragraphIndex`; `after: 0` means the top), which is easy to get wrong in a test and look like a tool bug. python-docx reports `None` as the style of unstyled paragraphs in documents made by the `docx` library, which declares no default paragraph style; that is harmless, since Word falls back to the document defaults.

**A retried POST can act twice.** The SDK retries 429/503/504 for any JSON-bodied request, POSTs included, and a 503 or 504 can arrive after Graph already sent the mail or posted the message. `graphForTool` (`src/graph/retry.ts`) gives each tool call a client whose requests carry that tool's retry options; for a mutating tool, a POST is retried only on 429, which Graph returns before acting. Keep `mutating: true` accurate on new tools: it now decides retries as well as the write guard. `test/retry.test.ts` runs the SDK's real RetryHandler under nock, including a baseline showing the double send without the policy.

**To Do rejects query options that work elsewhere.** Seen live: `/me/todo/lists` answers "Invalid request" to a multi-field `$select` (a single field is silently ignored), and tasks inside `$batch` reject any `$filter` or `$select`; only `$top` works. The prompt templates (`src/prompts/`) fetch tasks plainly and filter in the model. Any new Graph query in a prompt should be run live before it ships: the prompts' queries were, and this is what it caught.

**Tool results must go through `serializeResult`** (`src/util/redact.ts`). Graph attaches `@microsoft.graph.downloadUrl` to drive items: a `tempauth` link that downloads the file without authentication for about an hour. Results land in the model's context, transcripts and logs, so those keys are stripped at any depth, which covers listings, search, delta and batch alike.

**Large MCP messages end the server.** The SDK's stdio transport rejects any message over its buffer limit by closing the transport, and the process then exits with code 0. The server now sets that limit from `MCP_MAX_MESSAGE_MB` (default 64) and logs the error and the close. Anything that puts big payloads in a tool call, `files_upload` above all, is bounded by it. Upload sessions (`src/graph/upload.ts`) have been verified live at 6, 25 and 45 MB.

**Nothing big goes back through a tool result either.** A result lands in the model's context and has to fit the *client's* inbound message limit (10 MiB by the SDK default), so `files_download` refuses more than 5 MB inline. Big files stream to `MCP_DOWNLOAD_DIR` instead (`src/graph/download.ts`), verified live at 45 MB with flat memory. MCP resources were considered and not built: each read is still one message, so they would not lift the size limit.

**The token store is not ours, and not always private.** `@azure/identity-cache-persistence` owns it and offers no clear/delete API; `src/auth/tokenStore.ts` mirrors its platform selection to reach it. On macOS and Linux with a keyring it is one keychain item (`Microsoft.Developer.IdentityService`/`MSALCache`) shared by every app on the machine using that plugin, whatever name is passed; even the Windows file held tokens for two client ids on the development machine. So anything that edits it must remove only entries with our `client_id`, never the whole item. If the plugin changes where it stores things, `tokenStore.ts` has to follow.

**Test against a live tenant, not just the mocks.** `graph_search` was first written with a wrong model of which entity types Graph will combine: `message` + `event` is rejected, even though both are mail-ish. The unit tests encoded the same wrong assumption and passed happily. Only a live call caught it. Mocks verify the model, not reality.

**`dist/` is somebody's running server.** It is gitignored, so rebuilding on a feature branch silently replaces the binary a connected MCP client is using. Build in a worktree, or expect to break the thing you are testing against.

**Driving the server in tests.** Spawn it, write JSON-RPC lines to stdin, read responses from stdout: `initialize`, then `notifications/initialized`, then `tools/call`. Read stdout as UTF-8 explicitly — real mail subjects contain non-ASCII, and a default Windows codec will throw mid-run.

**Configuration.** A desktop client may rewrite its own config file from memory on exit, silently discarding edits made while it was running. Edit those files with the client fully closed. Where more than one config file exists, confirm which the process actually reads before assuming: check the client id in the startup log, or read the `app_displayname` claim out of a Graph download URL.

**The pre-commit hook** runs gitleaks against staged changes and blocks the commit. Enable it once per clone with `git config core.hooksPath .githooks`. Without gitleaks installed it warns and lets the commit through, so it never blocks a fresh clone.
