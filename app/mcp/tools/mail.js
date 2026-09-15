const { z } = require('zod');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { docxToText, xlsxBufferToText } = require('./files');
const { pdfToolResult, pptxToText } = require('./documents');
const { validateNext } = require('./discovery');
const { htmlToText, truncate, formatDate, isoDaysAgo, odataString, toolResult, personName, personAddress } = require('./format');

const LIST_SELECT = 'id,subject,from,toRecipients,receivedDateTime,isRead,flag,bodyPreview,hasAttachments,importance,conversationId,webLink';

function renderMessageLine(m) {
  const flags = [
    m.isRead ? '' : 'UNREAD',
    m.flag?.flagStatus === 'flagged' ? 'FLAGGED' : '',
    m.hasAttachments ? 'ATT' : '',
    m.importance === 'high' ? 'HIGH' : '',
  ].filter(Boolean).join(',');
  const from = `${personName(m.from)}${personAddress(m.from)}`;
  return [
    `- [${formatDate(m.receivedDateTime)}] ${m.subject || '(no subject)'}${flags ? ` [${flags}]` : ''}`,
    `  from: ${from}`,
    `  id: ${m.id}`,
    `  preview: ${truncate((m.bodyPreview || '').replaceAll(/\s+/g, ' '), 200)}`,
  ].join('\n');
}

function renderMessageList(data) {
  const items = data?.value ?? [];
  if (items.length === 0) return '';
  const footer = data?.truncated ? '\n(more results available; narrow the query)' : '';
  return `${items.length} message(s):\n${items.map(renderMessageLine).join('\n')}${footer}`;
}

// Graph only accepts $filter together with $orderby=receivedDateTime when the
// filter's first clause is on receivedDateTime, so every filter starts with one.
const EPOCH = '1970-01-01T00:00:00Z';

function buildMailFilter({ unreadOnly, from, sinceDays, subjectContains, flaggedOnly }) {
  const clauses = [`receivedDateTime ge ${sinceDays ? isoDaysAgo(sinceDays) : EPOCH}`];
  if (unreadOnly) clauses.push('isRead eq false');
  if (flaggedOnly) clauses.push("flag/flagStatus eq 'flagged'");
  if (from) {
    const f = odataString(from);
    clauses.push(`(from/emailAddress/address eq '${f}' or contains(from/emailAddress/name,'${f}') or contains(from/emailAddress/address,'${f}'))`);
  }
  if (subjectContains) clauses.push(`contains(subject,'${odataString(subjectContains)}')`);
  return clauses.join(' and ');
}

/**
 * Register read-only mail tools.
 * @param {import('@modelcontextprotocol/sdk/server/mcp.js').McpServer} server
 * @param {object} graph - GraphApiClient
 */
function registerMailTools(server, graph) {
  server.registerTool('mail_list_messages', {
    title: 'List mail messages',
    description: 'List recent Outlook messages in a folder, newest first. Filter by unread, sender, subject, flagged, or age. Returns ids for mail_get_message.',
    inputSchema: {
      folder: z.string().default('inbox').describe("Well-known folder name (inbox, sentitems, drafts, archive, junkemail, deleteditems) or a folder id"),
      limit: z.number().int().min(1).max(50).default(25),
      unreadOnly: z.boolean().default(false),
      flaggedOnly: z.boolean().default(false),
      from: z.string().optional().describe('Sender email address or a fragment of the sender display name'),
      sinceDays: z.number().int().min(1).max(365).optional().describe('Only messages received within the last N days'),
      subjectContains: z.string().optional(),
    },
  }, async (args) => {
    const options = { top: args.limit, select: LIST_SELECT, orderby: 'receivedDateTime desc', filter: buildMailFilter(args) };
    const result = await graph.getMailFolderMessages(args.folder, options);
    return toolResult(result, renderMessageList);
  });

  server.registerTool('mail_get_message', {
    title: 'Get a mail message',
    description: 'Read one Outlook message in full (headers and body as plain text). Use the id from mail_list_messages or mail_search.',
    inputSchema: {
      id: z.string(),
      maxBodyChars: z.number().int().min(200).max(50_000).default(8000),
    },
  }, async ({ id, maxBodyChars }) => {
    const result = await graph.getMailMessage(id, {
      select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,flag,hasAttachments,importance,conversationId,webLink,body',
    });
    return toolResult(result, (m) => {
      const recipients = (list) => (list ?? []).map((r) => `${personName(r)}${personAddress(r)}`).join(', ');
      const body = m.body?.contentType === 'html' ? htmlToText(m.body.content) : (m.body?.content ?? '');
      return [
        `Subject: ${m.subject || '(no subject)'}`,
        `From: ${personName(m.from)}${personAddress(m.from)}`,
        `To: ${recipients(m.toRecipients)}`,
        m.ccRecipients?.length ? `Cc: ${recipients(m.ccRecipients)}` : null,
        `Received: ${formatDate(m.receivedDateTime)}`,
        `Status: ${m.isRead ? 'read' : 'unread'}${m.flag?.flagStatus === 'flagged' ? ', flagged' : ''}${m.hasAttachments ? ', has attachments (use mail_list_attachments)' : ''}`,
        `Conversation: ${m.conversationId}`,
        `Link: ${m.webLink}`,
        '',
        truncate(body, maxBodyChars),
      ].filter((line) => line !== null).join('\n');
    });
  });

  server.registerTool('mail_search', {
    title: 'Search mail',
    description: 'Full-text search across all mail folders using Outlook KQL syntax, e.g. "invoice 4412", "from:alice subject:outage", "hasattachment:true". Results are relevance-ordered. Pass back a "more: cursor" value to fetch the next page. receivedAfter/receivedBefore (ISO dates) are applied client-side to the returned page, since Graph cannot reliably combine $search with $filter.',
    inputSchema: {
      query: z.string().min(1),
      limit: z.number().int().min(1).max(50).default(20),
      cursor: z.string().optional().describe('Continue into the next page from a previous result\'s "more: cursor" value'),
      receivedAfter: z.string().optional().describe('ISO date; drop messages received before this (applied to this page only)'),
      receivedBefore: z.string().optional().describe('ISO date; drop messages received on/after this (applied to this page only)'),
    },
  }, async ({ query, limit, cursor, receivedAfter, receivedBefore }) => {
    let result;
    if (cursor) {
      let validated;
      try {
        validated = validateNext(cursor, '/me/messages');
      } catch {
        return { isError: true, content: [{ type: 'text', text: 'Invalid or expired cursor.' }] };
      }
      result = await graph.makeRequest(validated);
    } else {
      const escaped = query.replaceAll('"', String.raw`\"`);
      result = await graph.getMailMessages({ search: `"${escaped}"`, top: limit, select: LIST_SELECT });
    }
    if (!result.success) return toolResult(result, () => '');
    const after = receivedAfter ? new Date(receivedAfter).getTime() : null;
    const before = receivedBefore ? new Date(receivedBefore).getTime() : null;
    const items = (result.data?.value ?? []).filter((m) => {
      if (!Number.isFinite(after) && !Number.isFinite(before)) return true;
      const received = new Date(m.receivedDateTime).getTime();
      if (Number.isNaN(received)) return true; // defensive: never silently drop on an unparsable date
      if (Number.isFinite(after) && received < after) return false;
      if (Number.isFinite(before) && received >= before) return false;
      return true;
    });
    const next = result.data?.['@odata.nextLink'];
    return toolResult({ success: true, data: items }, (data) => {
      const lines = [];
      const rendered = renderMessageList({ value: data });
      if (rendered) lines.push(rendered);
      if (next) lines.push(`more: cursor=${next}`);
      return lines.join('\n');
    });
  });

  server.registerTool('mail_get_attachment', {
    title: 'Read a mail attachment',
    description: 'Content of one attachment from mail_list_attachments: Word (.docx), Excel (.xlsx), PDF and PowerPoint (.pptx) become text; plain text is returned as is; other types (images) are saved to saveDir and the path is returned. Read-only against Microsoft.',
    inputSchema: {
      messageId: z.string(),
      attachmentId: z.string(),
      saveDir: z.string().optional().describe('Directory to save binary attachments into (default: the OS temp dir)'),
      maxChars: z.number().int().min(1000).max(400_000).default(60_000),
      maxPages: z.number().int().min(1).max(2000).default(200).describe('PDF only: maximum pages to extract text from'),
    },
  }, async ({ messageId, attachmentId, saveDir, maxChars, maxPages }) => {
    const meta = await graph.makeRequest(`/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}?$select=id,name,contentType,size`, { sensitive: true });
    if (!meta.success) return toolResult(meta, () => '');
    const name = String(meta.data?.name ?? 'attachment');
    const content = await graph.makeRequest(`/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`, { raw: true, sensitive: true });
    if (!content.success) return toolResult(content, () => '');
    const buffer = content.data;
    const ext = name.toLowerCase().split('.').pop();
    const header = `${name} (${Math.round(buffer.length / 1024)} KB)\n`;
    if (ext === 'docx') {
      const text = docxToText(buffer);
      if (text !== null) return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (ext === 'xlsx' || ext === 'xlsm') {
      const text = xlsxBufferToText(buffer, 500);
      if (text !== null) return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (ext === 'pdf' || String(meta.data?.contentType ?? '') === 'application/pdf') {
      return pdfToolResult(buffer, { header, name, maxPages, maxChars, saveDir });
    }
    if (ext === 'pptx') {
      const text = pptxToText(buffer);
      if (text !== null) return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (['txt', 'md', 'csv', 'json', 'log', 'xml', 'al', 'yml', 'yaml', 'html', 'htm'].includes(ext) || String(meta.data?.contentType ?? '').startsWith('text/')) {
      const raw = buffer.toString('utf8');
      return { content: [{ type: 'text', text: header + '\n' + truncate(ext === 'html' || ext === 'htm' ? htmlToText(raw) : raw, maxChars) }] };
    }
    const dir = saveDir || os.tmpdir();
    fs.mkdirSync(dir, { recursive: true });
    const target = path.join(dir, name.replaceAll(/[\\/:*?"<>|]/g, '_'));
    fs.writeFileSync(target, buffer);
    return { content: [{ type: 'text', text: `${header}Saved to ${target} (${meta.data?.contentType || 'unknown type'}).` }] };
  });

  server.registerTool('mail_list_attachments', {
    title: 'List mail attachments',
    description: 'List attachment names, types and sizes for a message. Does not download content.',
    inputSchema: { messageId: z.string() },
  }, async ({ messageId }) => {
    const result = await graph.getMailAttachments(messageId, { select: 'id,name,contentType,size,isInline' });
    return toolResult(result, (data) => (data?.value ?? [])
      .map((a) => `- ${a.name} (${a.contentType || 'unknown type'}, ${Math.round((a.size || 0) / 1024)} KB${a.isInline ? ', inline' : ''}) id: ${a.id}`)
      .join('\n'));
  });
}

module.exports = { registerMailTools, buildMailFilter, renderMessageLine, LIST_SELECT };
