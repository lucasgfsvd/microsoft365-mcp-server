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
```

The last two need a sandbox: a OneNote section, and a private team with a Planner
plan, that nobody else uses. The ids of the current sandbox are in
[docs/handover.md](../../docs/handover.md).

The second argument is a scratch folder for downloads and `results-*.json`.
Needs a signed-in token cache; runs with writes enabled.

**What they touch, and clean up:**

- `office.mjs` works inside one new OneDrive folder, `mcp-live-test-<timestamp>`,
  deleted at the end (to the recycle bin). It creates an organization-only
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
- `teams-planner.mjs` posts a message and a reply in the sandbox team's channel, and
  creates, completes and deletes a Planner task. Channel messages cannot be deleted
  through the server, so they stay in the sandbox. Never point it at a team other
  people are in.
