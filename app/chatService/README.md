# Chat Service Module

Read-only clients for the Teams conversation data the MCP server (`app/mcp/`) exposes: chats,
channels, messages, inline images and meeting transcripts. Both clients reuse the signed-in Teams
session and issue `GET` requests only. See
[ADR-032](../../docs-site/docs/development/adr/032-mcp-server.md) for why conversations come from the
chat service instead of Graph.

## Files

- `index.js`: `ChatServiceClient`, for the regional chat service the Teams web client uses
  (`https://<region>.ng.msg.teams.microsoft.com`) and the chat aggregator
  (`teams.microsoft.com/api/csa`).
- `transcripts.js`: `TranscriptClient`, which finds meeting recordings in a meeting chat and reads
  their transcripts through SharePoint's media API, plus `formatTranscript`, which renders entries as
  `[h:mm:ss] Speaker: text`.

## ChatServiceClient

Authentication: `acquireSession()` runs code in the main window's renderer that asks the Teams web
client's discovery service for its Skype token and the regional chat service URL. The token is kept in
memory until 5 minutes before its expiry. A `401` clears it and retries once with a fresh one.

Requests run from the main process, one at a time and at least 700 ms apart, because the chat service
allows about 15 calls per 10 seconds. A `429` waits for `Retry-After` (or 3, 6, 9 seconds) and retries
up to 3 times. Each request times out after 30 seconds.

| Method | Reads |
|--------|-------|
| `getConversations(pageSize, cursor)` | Recent chats, group chats, meeting chats and channels, most recent first |
| `getMessages(conversationId, { pageSize, startTime, cursor })` | Messages in one conversation, newest first. `cursor` continues into older history |
| `getThread(conversationId)` | Thread details including members, used to name 1:1 and group chats |
| `getTeamsAndChannels()` | Every team and channel from the chat aggregator, using the Graph token from `GraphApiClient` plus the Skype token |
| `getImage(sourceUrl)` | An inline image, only from `asm.skype.com/v1/objects/` URLs |
| `get(path, params)` | Generic `GET` under the chat service base URL, used by the methods above |

`isEnabled()` follows `graphApi.enabled`: the client has no setting of its own.

## TranscriptClient

A recorded meeting leaves a recording message in its meeting chat. `findRecordings(chatId)` scans up
to 3 pages of that chat for those messages. `resolveItem(url)` turns the recording's SharePoint link
into a drive item through Graph, and `listTranscripts` and `getTranscript` read the transcript from
SharePoint's media API. SharePoint tokens come from the Teams web client's auth provider, are cached
per site origin in memory, and are refreshed 5 minutes before expiry. Meetings that were transcribed
but not recorded leave no recording message, so their transcripts are not reachable.

## Read-only guarantee and tokens

`tests/unit/chatServiceClient.test.js` fails if any file in this folder uses a `POST`, `PATCH`,
`PUT` or `DELETE` method, and checks that requests are sent as `GET`. Tokens stay in memory: they
are not written to disk, returned from an MCP tool or included in log lines.

## Wiring

`initializeMcpService()` in `app/index.js` creates both clients when `mcp.enabled` is true and passes
them to `McpService.initialize(graphApiClient, chatServiceClient, transcriptClient)` in
`app/mcp/index.js`.
