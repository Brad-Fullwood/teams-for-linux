---
id: 032-mcp-server
---

# ADR 032: Local Read-Only MCP Server

## Status

✅ Accepted (2026-09-08)

## Context

Users increasingly work alongside AI coding assistants. A request such as "the customer emailed about X, fix it" needs the assistant to read the email or Teams thread. Doing that normally requires a Microsoft Entra app registration with `Mail.Read` / `Chat.Read` consent, which many corporate tenants do not grant to individuals.

Teams for Linux already holds a working Graph token via the Teams web app's own auth provider ([ADR-030](030-graph-api-teams-session-token.md)), and the Graph API research documented that `/me/messages` works with that token. The roadmap has carried "Graph API Enhanced Features: mail preview" under *Awaiting User Feedback*; this is that request.

Options considered:

1. **External tool over Chrome DevTools Protocol.** `electronCLIFlags` can expose `--remote-debugging-port`; an outside process could call `teamsForLinuxReactHandler.acquireToken` and query Graph itself. Works, but it is a side door into the renderer, needs the debug port open permanently, and lives outside the project's tests and docs.
2. **A generic plugin system.** Would let this and other integrations live out of tree. Requires a plugin contract, a security model for third-party code in the main process, and versioning. Too large a change for one use case; conflicts with the *start simple* principle.
3. **A first-class opt-in module, like MQTT and Graph API.** One folder, one nested config block, one wiring call, a few read-only Graph methods.

## Decision

Option 3. `app/mcp/` runs a [Model Context Protocol](https://modelcontextprotocol.io) server over Streamable HTTP inside the main process, bound to `127.0.0.1`, disabled by default, and dependent on `graphApi.enabled`.

Constraints adopted:

- **Read-only by construction.** Tools call only `GET` methods on `GraphApiClient` and `ChatServiceClient`. Unit tests fail if a write method name or a non-GET verb appears under `app/mcp/` or `app/chatService/`. Graph's `/search/query` (a POST) is therefore not used; Teams search scans recent conversations client-side.
- **Teams data comes from the chat service, not Graph.** Validation in a tenant with Conditional Access session policies showed every Graph Teams-messaging endpoint (`/me/chats`, `/chats/{id}/members`, `/me/joinedTeams`) answering `401 Additional claims required` (a `capolids` claims challenge) to the Teams web token, and the Teams web MSAL wrapper does not forward a `claims` request. The token also lacks `Chat.Read`. The Teams web client itself reads conversations from a regional chat service (`https://<region>.ng.msg.teams.microsoft.com`) with a Skype token; `ChatServiceClient` reads that token and URL from the web client's discovery service through the ReactHandler and issues the same GET calls from the main process. Mail stays on Graph, where `Mail.Read` works. The teams/channels directory comes from the chat aggregator (`api/csa`), and 1:1 chat partners are named through Graph `/users/{id}` (`User.ReadBasic.All`). The chat service enforces about 15 calls per 10 s, so the client serialises requests and backs off on 429. Meeting transcripts are read from the recording's SharePoint media transcript (the API Stream uses), reached through the recording link the meeting chat already contains, Graph's share resolution, and a SharePoint token from the same auth provider; Graph's transcript endpoints need scopes the Teams token lacks.
- **Loopback only, origin-checked, optional bearer token.** No option to bind to other interfaces.
- **No new IPC.** The module talks to `GraphApiClient` directly in the main process.
- **Stateless transport.** A fresh `McpServer` per request keeps the module free of session bookkeeping; the cost is negligible for a local single-user tool.
- **No persistence.** Tokens stay in the existing in-memory cache; results are not cached.

Dependencies added: `@modelcontextprotocol/sdk`, `zod`. The SDK transitively brings `express` (unused here; the Node `http` transport is used). That dependency tree includes a package directory named `ipaddr.js`, which the shell-expanded `eslint **/*.js` glob passes to ESLint as an ignored path, aborting the run; the lint script now adds `--no-error-on-unmatched-pattern`, which keeps the linted file set unchanged.

## Consequences

- Users get assistant access to mail and chat with zero tenant changes. Scope availability is tenant-dependent; a 403 surfaces in the tool's error text.
- Attack surface: a local process can read the user's mail and chats through the port. Mitigations: opt-in, loopback, origin check, optional token, documented in the user guide.
- Two new runtime dependencies to keep updated.
- If a write capability is ever wanted it should be a separate, explicitly named opt-in, not an extension of these tools.
- The chat service is an internal Teams API, like the ReactHandler paths already relied on. A Teams web update can change the discovery service shape; the failure mode is a clear tool error, not a crash.

## Related

- [ADR-030](030-graph-api-teams-session-token.md) Graph API access via the Teams session token
- [ADR-007](007-embedded-mqtt-broker.md) MQTT integration precedent
- User guide: [MCP Integration](../../mcp-integration.md)
- Module: `app/mcp/README.md`
