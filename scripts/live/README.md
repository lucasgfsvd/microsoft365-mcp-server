# Live tests

Drive a separate server over MCP against a **real** tenant: every tool is called
the way a client would call it, and the result is checked, not just accepted.
Mocks verify our model of Graph; these verify Graph. Every pass so far has found
something the unit tests could not.

```bash
npm run build
node scripts/live/office.mjs "$PWD/dist/index.js" /tmp/live   # files, Word, PowerPoint, Excel
node scripts/live/pim.mjs    "$PWD/dist/index.js" /tmp/live   # mail, calendar, contacts, To Do, Planner, OneNote, Teams
```

The second argument is a scratch folder for downloads and `results-*.json`.
Needs a signed-in token cache; runs with writes enabled.

**What they touch, and clean up:**

- `office.mjs` works inside one new OneDrive folder, `mcp-live-test-<timestamp>`,
  deleted at the end (to the recycle bin). It creates an organization-only
  sharing link on a file in it. Edited documents are downloaded and opened with
  python-docx, python-pptx and openpyxl (`pip install python-docx python-pptx openpyxl`).
- `pim.mjs` sends one mail **to the signed-in user only** and replies to it, creates
  a draft, an event with no attendees, a contact and a OneNote page, deleting each
  afterwards. It completes a To Do task (there is no delete tool, so it stays,
  completed) and posts one message to the user's own Teams notes chat (`48:notes`),
  which only they can see. Everything carries an `[mcp-live-test …]` tag.
- Nothing is posted to Teams channels or Planner plans, which other people see.
