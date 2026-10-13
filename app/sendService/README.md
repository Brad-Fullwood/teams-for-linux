# Send service

A local HTTP endpoint that sends one Teams chat message or one Outlook mail reply per request, and only
with an approval token minted by a companion app when the user presses Send. It is the one write path
in this fork. The MCP server (`app/mcp/`) and the chat service client (`app/chatService/`) stay
read-only, and tests check that neither loads this module.

Decision record: [ADR 033](../../docs-site/docs/development/adr/033-gated-send-service.md).

## When it runs

`app/index.js` starts it next to the MCP server. It listens only when
`~/.config/work-assistant/send-secret` exists, is mode 600 and holds at least 32 bytes once surrounding
whitespace is trimmed. Otherwise it logs one line (`[SEND] Send service not started`) and does nothing.
It binds to `127.0.0.1:3041` and accepts no `Origin` or a loopback `Origin`, the same rule as the MCP.

## Request

`POST /send` with a JSON body:

| Field | Type | Meaning |
|---|---|---|
| `tool` | `teams_send_message` or `mail_send_reply` | What to send |
| `target` | string | Chat id (`19:...@...`, or `48:notes` for notes to self), or a mail message id |
| `body` | string | Plain text |
| `to`, `cc` | list of `"Name <addr>"` or bare addresses | Mail recipients. `to` is required for mail |
| `expiry` | integer, Unix seconds | Not in the past, at most 10 minutes ahead |
| `nonce` | 16 lowercase hex characters | One use per process |
| `token` | 64 lowercase hex characters | See below |

`token = HMAC-SHA256(secret, tool + "\n" + target + "\n" + sha256hex(body) + "\n" + str(expiry) + "\n" + nonce)`,
with the secret as the trimmed file content in UTF-8 and `body` hashed as UTF-8.

Checks, in order: JSON shape and content (400), target format and recipients (400), token, expiry and
nonce (403), 10 sends in a rolling hour (429). Then the send. Success is `200 {"ok": true, "id": ...}`,
anything else `{"error": "<short text>"}` with the status. Upstream failures are 502, a timeout is 504
(check the thread before sending again), and mail with no usable token is 501.

## Sends

- **Teams**: `POST {chat service}/v1/users/ME/conversations/{id}/messages` with the Skype token the reads
  use, through `ChatServiceClient.withSession`, so it shares the 700 ms gap, the 429 backoff and the 401
  refresh. Body `content` (HTML), `messagetype: "RichText/Html"`, `contenttype: "text"`, a random
  19 digit `clientmessageid`, and `imdisplayname` from Graph `/me`. Channel targets (`@thread.tacv2`, or a
  `;messageid=` suffix) are refused. No mentions.
- **Mail**: Graph `POST /me/messages/{id}/reply` when the Graph token carries `Mail.Send`. The Teams web
  token does not carry it, so in practice the service uses Outlook REST
  `POST https://outlook.office.com/api/v2.0/me/messages/{id}/reply` with a token for the
  `https://outlook.office.com` audience from the same Teams auth provider, which carries `Mail.Send`.
  The recipients are exactly `to` and `cc`. Neither token sending gives 501.

Plain text becomes HTML by escaping `& < > "`, one `<p>` per blank line separated paragraph and `<br>`
for single line breaks.

## Logging

One line per request: tool, status and the target cut to 30 characters. Never the body, recipients or
a token.

## Tests

`tests/unit/sendService.test.js`: HMAC against a vector computed in Python, expiry window, nonce reuse,
rate cap, HTML conversion, recipient parsing, both sends against stubbed `fetch` and clients, the HTTP
layer, secret file checks, and the guard that keeps this module out of `app/mcp` and `app/chatService`.
