# Change notifications (webhooks): design options

**Status: not built. This is a decision document.** It sets out what real-time
change notifications from Microsoft Graph would take for this server, and
recommends a direction. Facts about Graph are from Microsoft's documentation as
of September 2026; limits change, so re-check them before building.

---

## The real constraint is not the endpoint

The obvious problem is that Graph delivers notifications by POSTing to a
**publicly reachable HTTPS URL**, which a local stdio server does not have. The
deeper one is that this server is a **stdio process that lives only while an MCP
client has it open**, and nothing in MCP wakes a model when data changes.

So a notification needs somewhere to land while the server is off, and
something that acts on it when it arrives. Any design has to answer both:

- **Where do notifications go when no server is running?** A Graph subscription
  outlives the process that created it (days, see below). Webhook deliveries to
  a dead endpoint are retried for up to 4 hours, then lost.
- **Who consumes them?** MCP lets a server tell a client that a resource changed
  (`notifications/resources/updated`, for clients that subscribe), but whether a
  client re-reads it or tells the model is up to the client. A chat client will
  not start a turn on its own.

Real-time delivery is worth building for long-running, agent-style clients that
act on events. For a person chatting with an assistant, polling at the start of
a task gives the same answer.

---

## What Graph requires

**Subscriptions** (`POST /subscriptions`) name a `resource`, the `changeType`s
(`created`, `updated`, `deleted`), a `notificationUrl`, an `expirationDateTime`,
and a `clientState` secret that comes back on every notification so the
receiver can reject forged ones. A second subscription with the same resource
and change types is refused with `409 Conflict`.

**Maximum lifetimes** (so renewal is always needed):

| Resource | Maximum |
|---|---|
| Outlook message, event, contact | 10,080 min (under 7 days); 1,440 min with resource data |
| OneDrive `driveItem`, SharePoint `list` | 42,300 min (under 30 days) |
| Teams `chatMessage`, `chat`, `channel` | 4,320 min (3 days) |
| To Do `todoTask` | 4,230 min (under 3 days) |

Anything under 45 minutes is raised to 45. Teams subscriptions longer than an
hour also need a `lifecycleNotificationUrl`, for `reauthorizationRequired`,
`subscriptionRemoved` and `missed` events.

**Webhook endpoint rules:**

- **Validation.** On creation, Graph POSTs `?validationToken=…` to the URL. The
  endpoint must answer within **10 seconds** with `200`, `text/plain`, and the
  decoded token as the body, or the subscription is not created.
- **Delivery.** Each notification must get a `2xx` within **3 seconds**
  (`202` if it is queued for later). Failures are retried with backoff for
  **up to 4 hours**, with a 10-second timeout on retries.
- **Throttling.** More than 10% of responses over 3 s in 10 minutes marks the
  endpoint *slow* (deliveries delayed 10 minutes). More than 15% over 10 s marks
  it *drop*: notifications are discarded for 10 minutes and **cannot be
  recovered**.
- **Latency.** Typically under a minute for mail, calendar and contacts
  (3 minutes at most). Drive changes are usually under a minute but can take
  up to 6 hours.

A notification carries the changed item's id, not its content, unless the
subscription asks for resource data. That also requires an encryption
certificate and shortens Outlook lifetimes to a day. Either way the server
fetches what changed, and `graph_delta` already does that well.

---

## Options

| | What it takes | Server off? | Exposure | Fit |
|---|---|---|---|---|
| **A. Poll with `graph_delta`** (built) | Nothing | Nothing lost: the `deltaLink` picks up where it left off | None | Chat clients; any "what changed since…" |
| **B. Local listener + dev tunnel** | A tunnel (Microsoft dev tunnels, ngrok) and a small HTTP listener in the server | Deliveries fail, are retried for 4 h, then lost | A public URL into the user's machine while running | Development only |
| **C. Azure Event Hubs delivery** | Azure subscription, an event hub, and the *Microsoft Graph Change Tracking* app granted *Event Hubs Data Sender*; `notificationUrl` becomes `EventHub:https://<ns>.servicebus.windows.net/eventhubname/<hub>?tenantId=<domain>` | The hub keeps messages for its retention period; the server reads them when it next runs | **None inbound**: the server pulls | Real-time for agents without exposing anything |
| **D. Hosted relay** (e.g. an Azure Function receiving webhooks into a queue) | Hosting, deployment and secrets to run | Queue holds them | A public endpoint, but not the user's machine | Teams or organisations running shared infrastructure |

Event Grid partner topics are a variant of C, with similar trade-offs.

---

## Recommendation

1. **Keep `graph_delta` as the answer for now.** It handles deletions, survives
   the server being off, needs no infrastructure, and matches how chat clients
   actually work: ask what changed when a task starts.
2. **If real-time is needed, choose C (Event Hubs).** It removes the public
   endpoint problem outright and holds notifications while the server is not
   running. The shape would be:
   - opt-in config (`MCP_EVENTHUB_*`), off by default;
   - tools to create, list, renew and delete subscriptions, with the server
     renewing its own subscriptions before they expire;
   - a consumer (`@azure/event-hubs`, a new dependency) that checks
     `clientState`, ignores validation messages, and records the ids that changed;
   - a `graph_changes` tool to drain them, plus
     `notifications/resources/updated` for clients that subscribe to resources.
3. **Do not ship B** beyond a development aid: a tunnel into the user's machine
   is the wrong default for a tool that reads their mail.

## Decisions this needs from the owner

- Is there a real consumer, meaning an agent that acts on events, or is polling enough?
- If C: which Azure subscription and tenant own the event hub, and who pays for it
  (the Basic tier is inexpensive, but it is not free)?
- Which resources matter? Mail and calendar are the usual pair; Teams brings the
  lifecycle-notification requirement with it.
