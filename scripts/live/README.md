# Live tests

Drive a separate server over MCP against a **real** tenant: every tool is called
the way a client would call it, and the result is checked, not just accepted.
Mocks verify our model of Graph; these verify Graph. Every pass so far has found
something the unit tests could not.

```bash
npm run build
node scripts/live/office.mjs "$PWD/dist/index.js" /tmp/live   # files, Word, PowerPoint, Excel
node scripts/live/pim.mjs    "$PWD/dist/index.js" /tmp/live   # mail, calendar, contacts, To Do, Planner, OneNote, Teams
node scripts/live/onenote.mjs "$PWD/dist/index.js" /tmp/live <sectionId>          # OneNote page writes and resource
node scripts/live/teams-planner.mjs "$PWD/dist/index.js" /tmp/live <teamId> <planId>  # channel posts, Planner tasks
node scripts/live/signin.mjs "$PWD/dist/index.js" /tmp/live [image]   # encrypted cache, sign-in seen by other servers
node scripts/live/transfers.mjs "$PWD/dist/index.js" /tmp/live          # progress and cancellation of large transfers
node scripts/live/shared.mjs "$PWD/dist/index.js" /tmp/live <mailbox>  # another mailbox you can open (MCP_ENABLE_SHARED_MAILBOXES)
node scripts/live/apponly.mjs "$PWD/dist/index.js" /tmp/live          # app-only mode; needs MCP_TENANT_ID, MCP_CLIENT_ID, MCP_USER and a certificate in the environment
```

The last two need a sandbox: a OneNote section, and a private team with a Planner
plan, that nobody else uses. The ids of the current sandbox are in
[docs/handover.md](../../docs/handover.md).

The second argument is a scratch folder for downloads and `results-*.json`.
Needs a signed-in token cache; runs with writes enabled.

**What they touch, and clean up:**

- `office.mjs` works inside one new OneDrive folder, `mcp-live-test-<timestamp>`,
  deleted at the end (to the recycle bin), including a 12 MB file uploaded from
  `<scratch>/up` with `localPath`. It creates an organization-only
  sharing link on a file in it. Edited documents are downloaded and opened with
  python-docx, python-pptx and openpyxl (`pip install python-docx python-pptx openpyxl`).
- `pim.mjs` sends one mail **to the signed-in user only** and replies to it, creates
  a draft and a reply draft, an event with no attendees, a contact, a OneNote page
  and a To Do task, deleting each afterwards. It posts one message to the user's own
  Teams notes chat (`48:notes`), which only they can see. Everything carries an
  `[mcp-live-test …]` tag.
- `onenote.mjs` creates and deletes one page in the given section, then creates a
  notebook `mcp-live-test-<timestamp>` with a section and a page, and deletes the
  notebook's folder (`/Notebooks/…` in OneDrive) to the recycle bin; Graph has no
  way to delete a notebook itself.
- `signin.mjs` needs a person: it prints a device code to enter. It signs in once
  into a fresh token cache encrypted with a random `MCP_TOKEN_CACHE_KEY`, checks
  that a server already running sees that sign-in, that a restarted one is still
  signed in and a wrong key is not, and, given a Docker image, that a container
  starts signed in from the same file. It ends with `--logout` and deletes its
  folder.
- `transfers.mjs` uploads 24 MB from `<scratch>/up` and downloads it again, both with
  progress, then cancels a 48 MB upload partway, all inside one new OneDrive folder
  `mcp-live-test-<timestamp>`, deleted at the end.
- `shared.mjs` needs a person: it signs in (device code, consenting to the `.Shared`
  scopes) to a throwaway encrypted cache, then lists folders, mail and events in the
  given mailbox and creates and deletes one draft there.
- `apponly.mjs` runs as an app registration with application permissions, working
  for `MCP_USER`: it reads mail and events, creates and deletes a draft, an event, a
  To Do task and a OneDrive folder with a file and a workbook. It sends nothing.
- `teams-planner.mjs` posts a message and a reply in the sandbox team's channel, and
  creates, completes and deletes a Planner task. Channel messages cannot be deleted
  through the server, so they stay in the sandbox. Never point it at a team other
  people are in.
