const { z } = require('zod');
const { truncate, formatDate, toolResult, personName } = require('./format');
const { buildMailFilter, LIST_SELECT } = require('./mail');
const { chatLabel, isUnread, messageAuthor, messageText, selectConversations, resolvePartnerNames } = require('./teams');

/**
 * Register the triage digest: "what am I being chased on".
 * @param {import('@modelcontextprotocol/sdk/server/mcp.js').McpServer} server
 * @param {object} graph - GraphApiClient (mail)
 * @param {object} chat - ChatServiceClient (Teams)
 */
function registerTriageTools(server, graph, chat) {
  server.registerTool('triage_digest', {
    title: 'Triage digest',
    description: 'What needs attention: unread and flagged inbox mail plus unread Teams conversations, grouped by person, oldest-waiting first. Start here for "what am I being chased on".',
    inputSchema: {
      sinceDays: z.number().int().min(1).max(60).default(7),
      mailLimit: z.number().int().min(1).max(50).default(40),
      chatLimit: z.number().int().min(1).max(100).default(40),
    },
  }, async ({ sinceDays, mailLimit, chatLimit }) => {
    const [unread, flagged, conversations] = await Promise.all([
      graph.getMailFolderMessages('inbox', { top: mailLimit, select: LIST_SELECT, orderby: 'receivedDateTime desc', filter: buildMailFilter({ unreadOnly: true, sinceDays }) }),
      graph.getMailFolderMessages('inbox', { top: mailLimit, select: LIST_SELECT, orderby: 'receivedDateTime desc', filter: buildMailFilter({ flaggedOnly: true }) }),
      chat.getConversations(Math.max(chatLimit * 3, 100)),
    ]);

    const errors = [unread, flagged, conversations].filter((r) => !r.success).map((r) => r.error);
    const since = Date.now() - sinceDays * 86_400_000;
    const bySender = new Map();
    const add = (name, item) => {
      if (!bySender.has(name)) bySender.set(name, []);
      bySender.get(name).push(item);
    };

    const seenMail = new Set();
    for (const m of [...(unread.data?.value ?? []), ...(flagged.data?.value ?? [])]) {
      if (seenMail.has(m.id)) continue;
      seenMail.add(m.id);
      add(personName(m.from), {
        when: m.receivedDateTime,
        line: `mail ${m.isRead ? '' : 'UNREAD '}${m.flag?.flagStatus === 'flagged' ? 'FLAGGED ' : ''}"${m.subject || '(no subject)'}" — ${truncate((m.bodyPreview || '').replaceAll(/\s+/g, ' '), 140)} (id: ${m.id})`,
      });
    }

    const list = selectConversations(conversations.data?.conversations ?? [], { kinds: ['chat', 'meeting', 'channel'], unreadOnly: true }).slice(0, chatLimit);
    await resolvePartnerNames(list, graph);
    for (const c of list) {
      const last = c.lastMessage;
      if (!last?.composetime || new Date(last.composetime).getTime() < since) continue;
      if (!isUnread(c)) continue;
      add(messageAuthor(last), {
        when: last.composetime,
        line: `chat "${chatLabel(c)}": ${truncate(messageText(last).replaceAll(/\s+/g, ' '), 160)} (chatId: ${c.id})`,
      });
    }

    const groups = [...bySender.entries()]
      .map(([name, items]) => ({ name, items: items.sort((a, b) => new Date(a.when) - new Date(b.when)) }))
      .sort((a, b) => new Date(a.items[0].when) - new Date(b.items[0].when));

    return toolResult({ success: true, data: groups }, (all) => {
      const body = all.map((g) => [
        `## ${g.name} (${g.items.length}, oldest ${formatDate(g.items[0].when)})`,
        ...g.items.map((i) => `- [${formatDate(i.when)}] ${i.line}`),
      ].join('\n')).join('\n\n');
      const warn = errors.length ? `\n\nWarnings: ${errors.join('; ')}` : '';
      return `${body || 'Nothing waiting.'}${warn}`;
    });
  });
}

module.exports = { registerTriageTools };
