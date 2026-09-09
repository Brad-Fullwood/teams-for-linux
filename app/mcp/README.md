# MCP Module

A local, read-only [Model Context Protocol](https://modelcontextprotocol.io) server. It lets AI assistants such as Claude Code read your Teams chats and Outlook mail through the Teams session you are already signed in to. Nothing in this module can send, modify or delete anything.

## How it works

- **Mail** uses the existing `GraphApiClient` (`app/graphApi/`). Tokens come from the Teams web app's own auth provider via `teamsForLinuxReactHandler.acquireToken`, so no app registration or extra consent is needed.
- **Teams conversations** use `ChatServiceClient` (`app/chatService/`), which talks to the same regional chat service the Teams web client uses (`https://<region>.ng.msg.teams.microsoft.com`). Its Skype token and URL are read from the web client's discovery service. The teams/channels directory comes from the chat aggregator (`teams.microsoft.com/api/csa`). Requests are serialised 700 ms apart with 429 backoff because the service allows roughly 15 calls per 10 s. Graph's Teams-messaging endpoints are not used: in tenants with Conditional Access session policies they answer `401 Additional claims required` to the Teams web token, and the token lacks `Chat.Read` anyway.
- Serves MCP over Streamable HTTP on `http://127.0.0.1:<port>/mcp` using `@modelcontextprotocol/sdk`. Stateless mode: a fresh server per request.
- Binds to `127.0.0.1` only, rejects non-loopback `Origin` headers, and optionally requires a bearer token.
- Every tool issues `GET` requests only. Unit tests (`tests/unit/mcpTools.test.js`, `tests/unit/chatServiceClient.test.js`) fail if any write method or verb appears under `app/mcp/` or `app/chatService/`.

## Configuration

```json
{
  "graphApi": { "enabled": true },
  "mcp": {
    "enabled": true,
    "port": 3040,
    "authToken": ""
  }
}
```

| Option | Default | Description |
|--------|---------|-------------|
| `mcp.enabled` | `false` | Start the server. Requires `graphApi.enabled`. |
| `mcp.port` | `3040` | Localhost port. |
| `mcp.authToken` | `""` | If set, clients must send `Authorization: Bearer <token>`. |

Restart the app after changing these.

## Connecting a client

Claude Code:

```bash
claude mcp add --transport http --scope user teams http://127.0.0.1:3040/mcp
# with a token:
claude mcp add --transport http --scope user teams http://127.0.0.1:3040/mcp --header "Authorization: Bearer <token>"
```

## Tools

| Tool | Reads |
|------|-------|
| `me` | Signed-in user's name and email |
| `teams_list_chats` | Recent chats, meeting chats, channels and notes-to-self with last-message preview, UNREAD marker and an `older:` cursor |
| `teams_list_channels` | Every team with all its channels, followed or not, with unread markers |
| `teams_get_meeting_transcript` | Transcript of a recorded meeting, from the recording's SharePoint media transcript |
| `teams_get_chat_messages` | Messages in one conversation (chat, meeting chat or channel); returns an `older:` cursor for history |
| `teams_search_messages` | Text search across recent conversations (client-side scan) |
| `mail_list_messages` | Messages in a folder with unread/flagged/sender/subject/age filters |
| `mail_get_message` | One message with plain-text body |
| `mail_search` | KQL search across all folders |
| `mail_list_attachments` | Attachment names and sizes (no download) |
| `files_shared_with_me` | Documents shared with you on OneDrive/SharePoint |
| `files_get_content` | Text of a shared Word, Excel or plain-text file from its link |
| `triage_digest` | Unread and flagged mail plus chats awaiting a reply, grouped by person, oldest first |

## Scope notes

Mail relies on the Teams web token carrying `Mail.Read` (it does in the tenants tested). Teams conversations rely on the chat service session, which exists whenever the client is signed in. A tool that fails reports the HTTP status in its error text.

## Files

- `index.js` — `McpService`: HTTP listener, origin and auth checks, per-request MCP server
- `tools/format.js` — HTML to text, OData escaping, result shaping
- `tools/teams.js`, `tools/mail.js`, `tools/triage.js` — tool definitions
- `../chatService/index.js` — `ChatServiceClient`: Skype token + regional chat service, GET only
- `../chatService/transcripts.js` — `TranscriptClient`: recording link → Graph drive item → SharePoint media transcript, GET only
