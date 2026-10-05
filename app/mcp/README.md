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
| `teams_get_chat_messages` | Messages in one conversation (chat, meeting chat or channel); returns an `older:` cursor for history. Shows reactions, quoted-reply context, adaptive/rich card text and, in channels, thread linkage — see below |
| `teams_search_messages` | Text search across recent conversations (client-side scan) |
| `mail_list_messages` | Messages in a folder with unread/flagged/sender/subject/age filters |
| `mail_get_message` | One message with plain-text body |
| `mail_search` | KQL search across all folders; `cursor` pages via `@odata.nextLink`, `receivedAfter`/`receivedBefore` filter the page client-side |
| `mail_list_attachments` | Attachment names and sizes (no download) |
| `mail_get_attachment` | Text of one attachment: Word, Excel, PDF and PowerPoint converted to text; images saved to disk |
| `mail_download_attachment` | Save one attachment's original bytes to disk, no conversion (25 MB limit, no overwrite) |
| `calendar_list_events` | Outlook calendar events for a date range in local time, recurring meetings expanded, with all-day, tentative and cancelled markers, organiser and Teams link |
| `calendar_get_event` | One event with attendees and their responses, and the agenda as plain text |
| `files_shared_with_me` | Documents shared with you on OneDrive/SharePoint |
| `files_get_content` | Text of a shared Word, Excel, PDF, PowerPoint or plain-text file from its link; images saved to disk |
| `triage_digest` | Unread and flagged mail plus chats awaiting a reply, grouped by person, oldest first |

## Scope notes

Mail relies on the Teams web token carrying `Mail.Read` (it does in the tenants tested). Teams conversations rely on the chat service session, which exists whenever the client is signed in. A tool that fails reports the HTTP status in its error text.

## Files

- `index.js` — `McpService`: HTTP listener, origin and auth checks, per-request MCP server
- `tools/format.js` — HTML to text, OData escaping, result shaping
- `tools/teams.js`, `tools/mail.js`, `tools/calendar.js`, `tools/triage.js` — tool definitions
- `../chatService/index.js` — `ChatServiceClient`: Skype token + regional chat service, GET only
- `../chatService/transcripts.js` — `TranscriptClient`: recording link → Graph drive item → SharePoint media transcript, GET only

## Original files and folders

`files_list_folder(url, limit, cursor)` lists immediate SharePoint or OneDrive folder children using the existing Teams session. Pass the returned cursor with the same folder URL for the next page.

`files_download(url, saveDir)` saves original file bytes, including Word custom XML and content controls, without conversion. Files are limited to 25 MB. The default destination is a new temporary directory; existing files are never overwritten. `files_get_content` retains its text extraction behaviour. No browser login or additional authentication flow is introduced.

Inline Teams images are now listed as `image:` URLs in conversation messages.
Use `teams_download_image` to save an original image using the existing Teams session.
The download is GET-only, limited to HTTPS Teams AMS object URLs, refuses redirects,
checks the image MIME type, and limits the response to 25 MB. It never exposes the session token.

## Richer chat message rendering

`teams_get_chat_messages` accepts `maxChars` (default 2000, min 200, max 20000) to control how much of each message body is kept; this replaces what used to be a fixed 2000-character cut.

Beyond the base `[time] Author: text` line, a message can carry extra lines:

- `reactions: like×2 (Name A, Name B); heart×1 (Name C)` — parsed from `message.properties.emotions` (JSON string or array of `{key, users:[{mri, time}]}`). MRIs are resolved to display names with the same Graph `/users/{id}?$select=displayName` lookup used for 1:1 chat naming, sharing its cache; lookups are capped at 50 new users per call, and a reaction whose users could not be resolved is still shown as a bare count (`heart×1`).
- `card: <text>` — visible text pulled out of adaptive/rich card JSON (`message.properties.cards`, or `message.attachments[].content`): `TextBlock`/`RichTextBlock` text, `FactSet` title/value pairs, walked recursively through `Container`/`ColumnSet`/`Column`; images and link-only elements are skipped. Bounded by `maxChars`. This is what makes Praise and Viva Engage card bodies visible instead of just the card title.
- `  > quoting <Author> (<time>): <preview>` — when a reply quotes an earlier message (`<blockquote itemtype="http://schema.skype.com/Reply">…</blockquote>` in the HTML), the quote is pulled out onto its own line and removed from the reply's own text, so the main line contains only the replier's words. Missing author/time/preview are handled gracefully; a blockquote with neither an author nor a preview is dropped rather than shown empty.
- `[id <short>]` / ` (reply in thread <short>)` — in channel conversations (heuristically, ids ending `@thread.tacv2` or `@thread.skype`), a root post is tagged with its own short id and a reply is tagged with the short id of the message it replies to, so replies can be grouped without extra calls. The root/parent id is read from `message.properties.rootMessageId`, `message.rootMessageId`, `message.parentMessageId`, or a `;messageid=<id>` suffix on `message.conversationLink` — whichever is present. **Open gap:** this only surfaces threading when the chat service inlines it on the reply message itself; if a tenant's channel replies do not carry any of those fields inline, no thread tag appears and there is currently no separate call to fetch thread structure (none was confirmed in the existing chat service surface). The channel-id heuristic and the exact field the live service uses have not been verified against real channel traffic yet — check after restart.

## Mail search paging and date filtering

`mail_search` now accepts `cursor` (follows `@odata.nextLink` from the previous page; validated to be the same Graph host and `/me/messages` path before it is used) and prints `more: cursor=<value>` when Graph reports another page. `receivedAfter`/`receivedBefore` (ISO dates) are applied client-side to the page already returned, because Graph does not reliably combine `$search` with `$filter`; `limit` stays capped at 50 per page.

## Project file discovery

Project investigations must check **files as well as chats**. Start with `sites_search(query)` and `files_search(query)`, follow project-site links, enumerate `files_list_libraries(siteId)`, and use `files_search(query, siteId)` across every library, including **Delivery Documents**. Search CDD identifiers separately if a combined project/document query finds nothing. Read candidate documents with `files_get_content`; authoritative CDD requirements take precedence over background solution designs.

All discovery uses the existing Teams authentication and Microsoft Graph GET endpoints: `/sites?search=`, `/sites/{id}/drives`, `/me/drive/search(q=...)`, `/drives/{id}/root/search(q=...)`, and `/drives/{id}/items/{id}/search(q=...)`. [Graph file search](https://learn.microsoft.com/en-us/graph/api/driveitem-search?view=graph-rest-1.0).

Follow every returned cursor. Site search covers every enumerated library, not only the default drive. File-search cursors are opaque, expire after 30 minutes or application restart, and must be used with identical arguments. Results include names, URLs, locations, modification timestamps, explicit completion and access errors. Indexed search cannot prove absence: list the project folders if necessary. Denied scopes and interrupted pagination must be reported as incomplete. No query text, document content or sensitive request URLs are logged.
