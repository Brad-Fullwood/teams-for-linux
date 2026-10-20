---
id: 034-gated-send-service
---

# ADR 034: Gated Send Service

## Status

✅ Accepted (2026-10-13). Fork only, not part of the upstream MCP proposal.

## Context

The MCP server ([ADR 033](033-mcp-server.md)) is read-only. An assistant built on it drafts Teams
messages and mail replies, and the user copies each one into Teams or Outlook to send it. The user
decided on 2026-10-13 that sending is allowed when every message is a draft he edits in his own panel
and nothing goes out until he clicks Send and confirms.

Assistants that read through the MCP must still be unable to send, by mistake or because a message
they read tells them to. Adding a send tool to the MCP would put it in every assistant's tool list.

A scope probe on 2026-10-13 decoded the tokens the Teams web auth provider issues in this tenant. The
Graph token has `Mail.Read` and `Mail.ReadWrite` but no `Mail.Send`. A token for the
`https://outlook.office.com` audience from the same provider has `Mail.Send`, and Outlook REST v2.0
answers with it and uses the same message ids as Graph.

## Decision

A separate module, `app/sendService/`, on its own listener (`127.0.0.1:3041`, path `/send`):

- Each request carries an approval token: an HMAC-SHA256 over the tool, target, a SHA-256 of the body,
  an expiry at most 10 minutes ahead and a one-use nonce, keyed with a secret file
  (`~/.config/work-assistant/send-secret`, mode 600). Only the companion app's Send route mints it.
  Any change to the text, the target or the tool breaks the token.
- The listener starts only when that file exists, so a normal install never opens it.
- Sends are capped at 10 an hour per process.
- Teams messages go to the chat service through `ChatServiceClient.withSession`, which hands the
  session to a callback inside the existing request queue. `app/chatService` itself still issues GET
  only.
- Mail replies use Graph when its token has `Mail.Send`, else Outlook REST v2.0 with the Outlook-audience
  token, else 501.
- Version 1 sends plain text converted to HTML, to chats only (no channel posts), with no mentions and
  no attachments.
- Unit tests keep `app/mcp` and `app/chatService` free of the module and of write verbs.

## Consequences

- The gate stops an assistant sending without the user, because no assistant has a send tool or the
  companion app's panel key. Everything runs as one Unix user, so a process that sets out to read the
  secret file can still mint a token. A native confirm dialog in this app would close that gap at the
  cost of a second click.
- Sends from the main process skip the Defender for Cloud Apps session proxy, as the reads already do.
- Outlook REST v2.0 is a legacy API. If Microsoft turns it off for this client, mail sends return 502
  and the user falls back to copying the draft.
- A chat service POST from the main process is new ground for this fork. Upstream recorded a 201 that
  did not reach the recipient on a different send path (ADR 015). The first live test reads the message
  back.
