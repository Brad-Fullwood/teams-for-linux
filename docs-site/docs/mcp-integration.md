# MCP Integration

:::info Feature Status
This feature is **disabled by default**. You must explicitly enable it in your configuration. It also requires `graphApi.enabled`.
:::

Teams for Linux can run a local, **read-only** [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server. AI assistants such as Claude Code connect to it to read your Teams chats and Outlook mail through the Teams session you are already signed in to.

Typical prompts once connected:

- "A customer emailed about the invoice mismatch. Find the email and summarise what they need."
- "Joe sent me the details on Teams yesterday. Pull them and let's fix it."
- "What am I being chased on that is an easy win?"

## How it works

- The server lives inside the Teams for Linux main process. Mail comes from Microsoft Graph through the existing [Graph API integration](development/adr/030-graph-api-teams-session-token.md); Teams conversations come from the same chat service the Teams web client uses. Both reuse the signed-in session, so there is **no app registration, no extra consent and no credentials to store**.
- It listens on `http://127.0.0.1:<port>/mcp` only. Requests from non-loopback origins are rejected. An optional bearer token can be required.
- **Every tool is a `GET`.** Nothing can send a message, mark anything as read, or change anything. Unit tests enforce this.
- Nothing is cached to disk; each tool call goes to Graph.

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

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `mcp.enabled` | boolean | `false` | Start the MCP server on launch |
| `mcp.port` | number | `3040` | Localhost port |
| `mcp.authToken` | string | `""` | If set, clients must send `Authorization: Bearer <token>` |

Restart Teams for Linux after changing these.

## Connecting Claude Code

```bash
claude mcp add --transport http --scope user teams http://127.0.0.1:3040/mcp
```

With a token:

```bash
claude mcp add --transport http --scope user teams http://127.0.0.1:3040/mcp \
  --header "Authorization: Bearer <token>"
```

Any MCP client that supports Streamable HTTP can connect the same way.

## Tools

| Tool | What it reads |
|------|---------------|
| `me` | Your name and email |
| `teams_list_chats` | Recent chats, meeting chats, channels and notes-to-self, with an unread marker and a cursor for older ones |
| `teams_list_channels` | Every team and channel you belong to, with unread markers |
| `teams_get_meeting_transcript` | Transcript of a recorded meeting, with speaker names and times |
| `teams_get_chat_messages` | Messages in one conversation, with a cursor for older history |
| `teams_search_messages` | Text search over recent conversations |
| `mail_list_messages` | A mail folder, with unread / flagged / sender / subject / age filters |
| `mail_get_message` | One message with a plain-text body |
| `mail_search` | Full-text mail search (Outlook KQL) |
| `mail_list_attachments` | Attachment names and sizes |
| `files_shared_with_me` | Documents people shared with you |
| `files_get_content` | Read a shared Word, Excel or text file |
| `triage_digest` | Unread and flagged mail plus chats waiting on you, grouped by person |

## Troubleshooting

**A mail tool returns `HTTP 403`.** The Teams web token does not carry `Mail.Read` in your tenant.

**A Teams tool returns "Teams discovery service not available".** The Teams page is not fully loaded yet, or a Teams web update changed its internals. Wait and retry; if it persists, open an issue.

**`HTTP 401`/"Failed to acquire token".** Teams is not fully signed in yet. Wait for the app to finish loading and retry.

**Port already in use.** Change `mcp.port`; the app logs `[MCP] Failed to listen` with the error code.

**Privacy.** The server is reachable by any process on your machine that can open a loopback connection. Set `mcp.authToken` on shared machines. Nothing is exposed on the network.

## Security notes

- Binds to `127.0.0.1` only; there is no option to bind elsewhere.
- Rejects requests whose `Origin` header is not `localhost`/`127.0.0.1` (browser-based DNS rebinding guard).
- Read-only by construction; see `app/mcp/README.md` for the enforcement test.

## Original files and folders

`files_list_folder(url, limit, cursor)` lists immediate SharePoint or OneDrive folder children using the existing Teams session. Pass the returned cursor with the same folder URL for the next page.

`files_download(url, saveDir)` saves original file bytes, including Word custom XML and content controls, without conversion. Files are limited to 25 MB. The default destination is a new temporary directory; existing files are never overwritten. `files_get_content` retains its text extraction behaviour. No browser login or additional authentication flow is introduced.

## Project file discovery

Project investigations must check **files as well as chats**. Start with `sites_search(query)` and `files_search(query)`, follow project-site links, enumerate `files_list_libraries(siteId)`, and use `files_search(query, siteId)` across every library, including **Delivery Documents**. Search CDD identifiers separately if a combined project/document query finds nothing. Read candidate documents with `files_get_content`; authoritative CDD requirements take precedence over background solution designs.

All discovery uses the existing Teams authentication and Microsoft Graph GET endpoints: `/sites?search=`, `/sites/{id}/drives`, `/me/drive/search(q=...)`, `/drives/{id}/root/search(q=...)`, and `/drives/{id}/items/{id}/search(q=...)`. [Graph file search](https://learn.microsoft.com/en-us/graph/api/driveitem-search?view=graph-rest-1.0).

Follow every returned cursor. Site search covers every enumerated library, not only the default drive. File-search cursors are opaque, expire after 30 minutes or application restart, and must be used with identical arguments. Results include names, URLs, locations, modification timestamps, explicit completion and access errors. Indexed search cannot prove absence: list the project folders if necessary. Denied scopes and interrupted pagination must be reported as incomplete. No query text, document content or sensitive request URLs are logged.
