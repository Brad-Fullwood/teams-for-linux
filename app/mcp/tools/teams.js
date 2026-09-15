const { z } = require('zod');
const { htmlToText, truncate, formatDate, isoDaysAgo, toolResult } = require('./format');
const { formatTranscript } = require('../../chatService/transcripts');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

// Conversation kinds worth showing. Streams (mentions, call logs, notifications, drafts) are noise.
const KINDS = {
  chat: 'chat',        // 1:1 and group chats (threadType "chat")
  meeting: 'meeting',  // meeting chats
  channel: 'channel',  // team General channels ("space") and standard channels ("topic")
  notes: 'notes',      // the user's own notes-to-self thread
};

function kindOf(conversation) {
  const type = conversation.threadProperties?.threadType;
  if (type === 'chat') return KINDS.chat;
  if (type === 'meeting') return KINDS.meeting;
  if (type === 'space' || type === 'topic') return KINDS.channel;
  if (type === 'streamofnotes') return KINDS.notes;
  return null;
}

// Display names by user object id, learned from message authors and Graph lookups.
const nameCache = new Map();
let myOid = null;

function oidFromMri(mri) {
  return String(mri ?? '').match(/8:orgid:([0-9a-f-]{36})/i)?.[1] ?? null;
}

function rememberAuthor(message) {
  const oid = oidFromMri(message.from);
  if (oid && message.imdisplayname) nameCache.set(oid, message.imdisplayname);
}

/** The other participant's object id in a 1:1 chat id "19:<oid>_<oid>@unq.gbl.spaces". */
function partnerOid(conversation) {
  const ids = String(conversation.id ?? '').match(/^19:([0-9a-f-]{36})_([0-9a-f-]{36})@unq\.gbl\.spaces$/i);
  if (!ids) return null;
  if (myOid && ids[1].toLowerCase() === myOid) return ids[2];
  if (myOid && ids[2].toLowerCase() === myOid) return ids[1];
  return null;
}

/**
 * Fill nameCache for 1:1 chats whose partner is unknown, using Graph /users/{id}
 * (User.ReadBasic.All is in the Teams token). Bounded so a listing stays fast.
 */
async function resolvePartnerNames(conversations, graph, maxLookups = 10) {
  if (!graph) return;
  if (!myOid) {
    const me = await graph.makeRequest('/me?$select=id');
    myOid = me.success ? String(me.data?.id ?? '').toLowerCase() : null;
    if (!myOid) return;
  }
  let lookups = 0;
  for (const c of conversations) {
    if (c.lastMessage) rememberAuthor(c.lastMessage);
    const oid = partnerOid(c);
    if (!oid || nameCache.has(oid)) continue;
    if (lookups >= maxLookups) break;
    lookups += 1;
    const user = await graph.makeRequest(`/users/${encodeURIComponent(oid)}?$select=displayName`);
    nameCache.set(oid, user.success && user.data?.displayName ? user.data.displayName : null);
  }
}

function chatLabel(conversation) {
  const tp = conversation.threadProperties ?? {};
  if (tp.threadType === 'streamofnotes') return 'Notes to self';
  if (tp.spaceThreadTopic) return tp.spaceThreadTopic;
  if (tp.topic) return tp.topic;
  if (tp.productThreadType === 'OneToOneChat') {
    const partner = nameCache.get(partnerOid(conversation) ?? '');
    if (partner) return `1:1 with ${partner}`;
    const last = conversation.lastMessage;
    const lastOid = oidFromMri(last?.from);
    if (last?.imdisplayname && (!lastOid || lastOid.toLowerCase() !== myOid)) return `1:1 with ${last.imdisplayname}`;
    return '1:1 chat';
  }
  return `(${tp.productThreadType || tp.threadType || 'conversation'})`;
}

/** Sender display name; falls back to a cached name for the sender id, then a neutral label. */
function messageAuthor(message) {
  if (message.imdisplayname) return message.imdisplayname;
  if (message.fromDisplayNameInToken) return message.fromDisplayNameInToken;
  const oid = oidFromMri(message.from);
  if (oid && nameCache.get(oid)) return nameCache.get(oid);
  return oid ? 'participant' : 'system';
}

/** Chat service marks read position as "<messageId>;<timestamp>;<clientId>". */
function isUnread(conversation) {
  const horizon = String(conversation.properties?.consumptionhorizon ?? '').split(';')[0];
  const lastId = Number(conversation.lastMessage?.id);
  if (!lastId) return false;
  if (!horizon) return true;
  return lastId > Number(horizon);
}

// Message types that carry something worth showing. Roster/topic/picture events do not.
const MEDIA_LABELS = {
  'RichText/Media_CallRecording': 'call recording',
  'RichText/Media_CallTranscript': 'call transcript',
  'RichText/Media_CallLogRecording': 'call recording',
  'RichText/Media_CallLogTranscript': 'call transcript',
  'RichText/Media_CallLogVoicemail': 'voicemail',
  'RichText/Media_GenericFile': 'file',
  'RichText/Media_Card': 'card',
  'RichText/Media_Video': 'video',
  'RichText/Media_Audio': 'audio',
};

function isContentMessage(message) {
  const type = String(message.messagetype ?? '');
  if (type === 'Event/Call') return true;
  if (MEDIA_LABELS[type]) return true;
  return (type === 'Text' || type === 'RichText' || type === 'RichText/Html') && Boolean(message.content);
}

function messageText(message) {
  const type = String(message.messagetype ?? '');
  if (type === 'Event/Call') {
    const content = String(message.content ?? '');
    const count = content.match(/count="(\d+)"/)?.[1];
    const duration = content.match(/<duration>([\d.]+)<\/duration>/)?.[1];
    const state = content.includes('<ended/>') ? 'ended' : content.includes('<started/>') ? 'started' : 'event';
    const minutes = duration ? `, ${Math.round(Number(duration) / 60)} min` : '';
    return `[call ${state}${count ? `, ${count} participants` : ''}${minutes}]`;
  }
  if (MEDIA_LABELS[type]) {
    const content = String(message.content ?? '');
    const title = content.match(/<Title>([^<]*)<\/Title>/i)?.[1] || content.match(/title="([^"]*)"/)?.[1];
    let link;
    try {
      link = JSON.parse(message.properties?.atp ?? '[]')[0]?.URL ?? '';
    } catch {
      link = '';
    }
    return `[${MEDIA_LABELS[type]}${title ? `: ${title}` : ''}${link ? ` ${link}` : ''}]`;
  }
  if (type === 'Text' || type === 'RichText') return String(message.content ?? '');
  return htmlToText(extractQuote(message.content).html);
}

function attachmentNames(message) {
  try {
    const files = JSON.parse(message.properties?.files ?? '[]');
    return files.map((f) => {
      const name = f.fileName || f.title;
      const url = f.objectUrl || f.fileInfo?.fileUrl || f.fileUrl || f.itemid;
      return name ? (url && /^https?:/.test(url) ? `${name} (${url})` : name) : null;
    }).filter(Boolean);
  } catch {
    return [];
  }
}

// --- Reactions (message.properties.emotions) -------------------------------------

/** Parse a message's reactions defensively: JSON string or array, tolerant of missing/unknown fields. */
function parseEmotions(message) {
  let raw = message?.properties?.emotions;
  if (!raw) return [];
  if (typeof raw === 'string') {
    try {
      raw = JSON.parse(raw);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(raw)) return [];
  return raw
    .filter((e) => e && typeof e === 'object' && e.key)
    .map((e) => ({ key: String(e.key), users: Array.isArray(e.users) ? e.users : [] }));
}

/**
 * Resolve reaction MRIs to display names via the same Graph /users/{id} lookup used
 * for 1:1 chat naming, sharing its cache. Bounded so a page of messages stays fast.
 */
async function resolveReactionNames(messages, graph, maxLookups = 50) {
  if (!graph) return;
  let lookups = 0;
  for (const m of messages ?? []) {
    for (const emotion of parseEmotions(m)) {
      for (const user of emotion.users) {
        const oid = oidFromMri(user?.mri);
        if (!oid || nameCache.has(oid)) continue;
        if (lookups >= maxLookups) return;
        lookups += 1;
        const resolved = await graph.makeRequest(`/users/${encodeURIComponent(oid)}?$select=displayName`);
        nameCache.set(oid, resolved.success && resolved.data?.displayName ? resolved.data.displayName : null);
      }
    }
  }
}

function reactionsLine(message) {
  const emotions = parseEmotions(message);
  const parts = [];
  for (const emotion of emotions) {
    if (!emotion.users.length) continue;
    const names = emotion.users
      .map((u) => { const oid = oidFromMri(u?.mri); return oid ? nameCache.get(oid) : null; })
      .filter(Boolean);
    parts.push(`${emotion.key}×${emotion.users.length}${names.length ? ` (${names.join(', ')})` : ''}`);
  }
  return parts.length ? `  reactions: ${parts.join('; ')}` : null;
}

// --- Adaptive / rich cards (message.properties.cards, message.attachments) -------

/** Root node whose children hold the visible card content, whatever shape it was wrapped in. */
function cardRoot(card) {
  if (card?.content?.body) return card.content.body;
  if (card?.body) return card.body;
  if (card?.content) return card.content;
  return card;
}

/** Recursively collect visible text from adaptive-card-shaped JSON, skipping images and bare URLs. */
function extractCardText(node, out, depth = 0) {
  if (!node || typeof node !== 'object' || depth > 12) return;
  if (Array.isArray(node)) {
    for (const item of node) extractCardText(item, out, depth + 1);
    return;
  }
  const type = node.type;
  if (type === 'Image') return;
  if (type === 'TextBlock') {
    if (typeof node.text === 'string' && node.text.trim()) out.push(node.text.trim());
    return;
  }
  if (type === 'RichTextBlock') {
    const parts = (Array.isArray(node.inlines) ? node.inlines : [])
      .map((i) => (typeof i === 'string' ? i : i?.text))
      .filter(Boolean);
    if (parts.length) out.push(parts.join(''));
    return;
  }
  if (type === 'FactSet') {
    for (const fact of Array.isArray(node.facts) ? node.facts : []) {
      const line = `${fact?.title ?? ''}: ${fact?.value ?? ''}`.trim();
      if (line && line !== ':') out.push(line);
    }
    return;
  }
  // Container/ColumnSet/Column and anything else that might nest content: walk known child arrays.
  for (const key of ['items', 'columns', 'body']) {
    if (Array.isArray(node[key])) extractCardText(node[key], out, depth + 1);
  }
}

/** Gather card-shaped JSON from properties.cards (string or array) and any message.attachments. */
function collectCardSources(message) {
  const sources = [];
  let raw = message?.properties?.cards;
  if (raw) {
    if (typeof raw === 'string') {
      try {
        raw = JSON.parse(raw);
      } catch {
        raw = null;
      }
    }
    for (const c of Array.isArray(raw) ? raw : (raw ? [raw] : [])) {
      if (typeof c === 'string') {
        try { sources.push(JSON.parse(c)); } catch { /* not JSON, skip */ }
      } else if (c && typeof c === 'object') {
        sources.push(c);
      }
    }
  }
  for (const attachment of Array.isArray(message?.attachments) ? message.attachments : []) {
    if (attachment?.content) sources.push(attachment.content);
  }
  return sources;
}

function cardText(message, maxChars) {
  const out = [];
  for (const card of collectCardSources(message)) extractCardText(cardRoot(card), out);
  const joined = out.join(' | ').replaceAll(/\s+/g, ' ').trim();
  return joined ? truncate(joined, maxChars) : null;
}

// --- Quoted replies (schema.skype.com/Reply blockquote in HTML content) ---------

const QUOTE_RE = /<blockquote\b[^>]*itemtype="http:\/\/schema\.skype\.com\/Reply"[^>]*>([\s\S]*?)<\/blockquote>/i;

/** Strip an inline quoted-reply blockquote out of Teams HTML and return its parts. */
function extractQuote(html) {
  const str = String(html ?? '');
  const match = str.match(QUOTE_RE);
  if (!match) return { html: str, quote: null };
  const block = match[1];
  const authorRaw = block.match(/<strong\b[^>]*itemprop="mri"[^>]*>([\s\S]*?)<\/strong>/i)?.[1];
  const timeRaw = block.match(/<span\b[^>]*itemprop="time"[^>]*itemid="(\d+)"/i)?.[1];
  const previewRaw = block.match(/<p\b[^>]*itemprop="preview"[^>]*>([\s\S]*?)<\/p>/i)?.[1];
  const quote = {
    author: authorRaw ? htmlToText(authorRaw).trim() || null : null,
    time: timeRaw ? Number(timeRaw) : null,
    preview: previewRaw ? htmlToText(previewRaw).trim() || null : null,
  };
  const stripped = str.slice(0, match.index) + str.slice(match.index + match[0].length);
  return { html: stripped, quote };
}

function quoteLine(quote) {
  if (!quote || (!quote.author && !quote.preview)) return null;
  const author = quote.author || 'someone';
  const time = Number.isFinite(quote.time) ? formatDate(new Date(quote.time).toISOString()) : null;
  const preview = quote.preview ? `: ${quote.preview}` : '';
  return `  > quoting ${author}${time ? ` (${time})` : ''}${preview}`;
}

// --- Channel threading (properties.rootMessageId / parentMessageId / conversationLink) ---

/** The id of the message this one replies to, from whichever field carries it. */
function threadRootId(message) {
  const root = message?.properties?.rootMessageId ?? message?.rootMessageId ?? message?.parentMessageId;
  if (root) return String(root);
  const fromLink = String(message?.conversationLink ?? '').match(/;messageid=(\d+)/i)?.[1];
  return fromLink ?? null;
}

/** Compact id for grouping replies in output; not a substitute for the real message id. */
function shortId(id) {
  const s = String(id ?? '');
  return s.length > 8 ? s.slice(-8) : s;
}

/** Heuristic: standard/private channel thread ids end in @thread.tacv2 or @thread.skype. */
function isChannelConversationId(chatId) {
  return /@thread\.(tacv2|skype)$/i.test(String(chatId ?? ''));
}

function renderChatMessage(m, opts = {}) {
  rememberAuthor(m);
  const maxChars = opts.maxChars ?? 2000;
  const isChannel = Boolean(opts.isChannel);
  const rootId = threadRootId(m);
  const isReply = Boolean(rootId) && rootId !== String(m.id ?? '');
  const idTag = isChannel && !isReply && m.id ? `[id ${shortId(m.id)}] ` : '';
  const threadTag = isReply ? ` (reply in thread ${shortId(rootId)})` : '';
  const lines = [`[${formatDate(m.composetime || m.originalarrivaltime)}] ${idTag}${messageAuthor(m)}: ${truncate(messageText(m), maxChars)}${threadTag}`];
  const files = attachmentNames(m);
  if (files.length) lines.push(`  attachments: ${files.join(', ')}`);
  const images = [...String(m.content ?? '').matchAll(/<img\b[^>]*\bsrc=["']([^"']+)["'][^>]*>/gi)].map(match => match[1].replaceAll('&amp;', '&'));
  for (const url of images) lines.push(`  image: ${url}`);
  const card = cardText(m, maxChars);
  if (card) lines.push(`  card: ${card}`);
  const reactions = reactionsLine(m);
  if (reactions) lines.push(reactions);
  const quote = extractQuote(m.content).quote;
  const quoted = quoteLine(quote);
  if (quoted) lines.push(quoted);
  if (m.properties?.deletetime) lines.push('  (deleted)');
  return lines.join('\n');
}

function renderConversation(c) {
  const last = c.lastMessage;
  const text = last ? truncate(messageText(last).replaceAll(/\s+/g, ' '), 160) : '';
  const preview = last ? `${messageAuthor(last)}: ${text || '(preview not available for this kind; use teams_get_chat_messages)'}` : '(no messages)';
  return [
    `- ${chatLabel(c)} [${kindOf(c)}${isUnread(c) ? ', UNREAD' : ''}]`,
    `  id: ${c.id}`,
    `  last: [${formatDate(last?.composetime)}] ${preview}`,
  ].join('\n');
}

function selectConversations(list, { kinds, unreadOnly }) {
  return list
    .filter((c) => kindOf(c) && kinds.includes(kindOf(c)))
    .filter((c) => !unreadOnly || isUnread(c))
    .sort((a, b) => new Date(b.lastMessage?.composetime ?? 0) - new Date(a.lastMessage?.composetime ?? 0));
}

const kindsSchema = z.array(z.enum(['chat', 'meeting', 'channel', 'notes'])).default(['chat', 'meeting']);

/**
 * Register read-only Teams tools backed by the Teams chat service.
 * @param {import('@modelcontextprotocol/sdk/server/mcp.js').McpServer} server
 * @param {object} chat - ChatServiceClient
 * @param {object} graph - GraphApiClient (the `me` tool and 1:1 chat naming)
 * @param {object} [transcripts] - TranscriptClient for meeting recordings
 */
function registerTeamsTools(server, chat, graph, transcripts) {
  server.registerTool('teams_download_image', {
    title: 'Download a Teams message image',
    description: 'Save an inline Teams image from an image URL returned by teams_get_chat_messages, using the existing Teams session. Read-only.',
    inputSchema: { url: z.string().url(), saveDir: z.string().optional() },
  }, async ({ url, saveDir }) => {
    const result = await chat.getImage(url);
    if (!result.success) return { isError: true, content: [{ type: 'text', text: result.error }] };
    const dir = saveDir ? path.resolve(saveDir) : fs.mkdtempSync(path.join(os.tmpdir(), 'teams-image-'));
    fs.mkdirSync(dir, { recursive: true });
    const extension = result.mimeType === 'image/jpeg' ? 'jpg' : result.mimeType.split('/')[1];
    const dest = path.join(dir, `teams-image-${Date.now()}.${extension}`);
    fs.writeFileSync(dest, result.data, { flag: 'wx' });
    return { content: [{ type: 'text', text: `Saved ${result.data.length} bytes: ${dest}` }] };
  });
  server.registerTool('me', {
    title: 'Who am I',
    description: 'Display name, email and job title of the signed-in user.',
    inputSchema: {},
  }, async () => {
    const result = await graph.makeRequest('/me?$select=displayName,mail,userPrincipalName,jobTitle,officeLocation');
    return toolResult(result, (u) => `${u.displayName} <${u.mail || u.userPrincipalName}>${u.jobTitle ? `, ${u.jobTitle}` : ''}${u.officeLocation ? `, ${u.officeLocation}` : ''}`);
  });

  server.registerTool('teams_list_chats', {
    title: 'List Teams conversations',
    description: 'Recent chats, meeting chats and channels with the last message preview and an UNREAD marker, most recently active first. Returns ids for teams_get_chat_messages.',
    inputSchema: {
      limit: z.number().int().min(1).max(100).default(25),
      kinds: kindsSchema.describe('Which conversation kinds to include'),
      unreadOnly: z.boolean().default(false),
      cursor: z.string().optional().describe('Continue into less recently active conversations from a previous result'),
    },
  }, async ({ limit, kinds, unreadOnly, cursor }) => {
    // Streams and other kinds are filtered client-side, so fetch well beyond `limit`.
    const result = await chat.getConversations(Math.max(limit * 3, 100), cursor);
    if (!result.success) return toolResult(result, () => '');
    const selected = selectConversations(result.data?.conversations ?? [], { kinds, unreadOnly }).slice(0, limit);
    await resolvePartnerNames(selected, graph);
    return toolResult(result, (data) => {
      const lines = selected.map(renderConversation);
      const older = data?._metadata?.backwardLink;
      if (older && (data?.conversations ?? []).length > 0) lines.push(`older: cursor=${older}`);
      return lines.join('\n');
    });
  });

  server.registerTool('teams_get_chat_messages', {
    title: 'Get conversation messages',
    description: 'Messages in one chat, meeting chat or channel, oldest first. Use the id from teams_list_chats. For older history, pass back the "older: cursor" value from a previous call. Shows reactions, quoted-reply context, adaptive/rich card text, and (in channels) thread linkage when present.',
    inputSchema: {
      chatId: z.string(),
      limit: z.number().int().min(1).max(200).default(30),
      sinceDays: z.number().int().min(1).max(3650).optional(),
      cursor: z.string().optional().describe('Continue into older messages from a previous result'),
      maxChars: z.number().int().min(200).max(20_000).default(2000).describe('Maximum characters kept per message body'),
    },
  }, async ({ chatId, limit, sinceDays, cursor, maxChars }) => {
    // Member/topic/picture events share the feed with real messages; fetch extra so `limit` means content.
    const options = { pageSize: Math.min(limit * 2, 200), cursor };
    if (sinceDays) options.startTime = Date.now() - sinceDays * 86_400_000;
    const result = await chat.getMessages(chatId, options);
    if (!result.success) return toolResult(result, () => '');
    const messages = (result.data?.messages ?? []).filter(isContentMessage).slice(0, limit).reverse();
    await resolveReactionNames(messages, graph);
    const isChannel = isChannelConversationId(chatId);
    return toolResult(result, (data) => {
      const older = data?._metadata?.backwardLink;
      const lines = messages.map((m) => renderChatMessage(m, { maxChars, isChannel }));
      if (older && (data?.messages ?? []).length > 0) lines.push(`older: cursor=${older}`);
      return lines.join('\n');
    });
  });

  server.registerTool('teams_list_channels', {
    title: 'List teams and channels',
    description: 'Every team you belong to with all its channels (followed or not), with unread markers. Channel ids work with teams_get_chat_messages. teams_list_chats only shows channels with recent activity.',
    inputSchema: {
      teamFilter: z.string().optional().describe('Only teams whose name contains this'),
    },
  }, async ({ teamFilter }) => {
    const result = await chat.getTeamsAndChannels();
    return toolResult(result, (data) => (data?.teams ?? [])
      .filter((t) => !t.isDeleted && (!teamFilter || String(t.displayName).toLowerCase().includes(teamFilter.toLowerCase())))
      .map((t) => [
        `- ${t.displayName}${t.isFavorite ? ' [favourite]' : ''} (${(t.channels ?? []).length} channels) team id: ${t.id}`,
        ...(t.channels ?? []).filter((c) => !c.isDeleted).map((c) => `    - ${c.displayName}${c.isGeneral ? ' (General)' : ''}${c.isMessageRead === false ? ' [UNREAD]' : ''}${c.isFollowed ? '' : ' (not followed)'} id: ${c.id}`),
      ].join('\n'))
      .join('\n'));
  });

  if (transcripts) {
    server.registerTool('teams_get_meeting_transcript', {
      title: 'Get meeting transcript',
      description: 'Transcript of a recorded meeting, as "[time] Speaker: text". Give the meeting chat id (kind meeting in teams_list_chats). If the meeting was recorded more than once, the first call lists the recordings and you pick one with recordingIndex. Only works for meetings that were recorded; transcript-only meetings leave no link in the chat.',
      inputSchema: {
        chatId: z.string(),
        recordingIndex: z.number().int().min(0).default(0).describe('Which recording, newest first'),
        maxChars: z.number().int().min(1000).max(400_000).default(60_000),
      },
    }, async ({ chatId, recordingIndex, maxChars }) => {
      const found = await transcripts.findRecordings(chatId);
      if (!found.success) return toolResult(found, () => '');
      const recordings = found.data;
      if (recordings.length === 0) {
        return { isError: true, content: [{ type: 'text', text: 'No recording link found in this chat. The meeting was not recorded, or the recording has been deleted.' }] };
      }
      if (!recordings[recordingIndex]) {
        return { isError: true, content: [{ type: 'text', text: `recordingIndex out of range; ${recordings.length} recording(s): ${recordings.map((r, i) => `${i}: ${r.title} (${formatDate(r.when)})`).join('; ')}` }] };
      }

      // Recordings expire or get deleted; walk forward from the requested one and report what was skipped.
      const skipped = [];
      for (let i = recordingIndex; i < recordings.length; i++) {
        const chosen = recordings[i];
        const item = await transcripts.resolveItem(chosen.url);
        if (!item.success) {
          skipped.push(`${i}: ${chosen.title} (${formatDate(chosen.when)}) — ${item.status === 404 ? 'recording deleted or expired' : item.error}`);
          continue;
        }
        const list = await transcripts.listTranscripts(item.data);
        if (!list.success) return toolResult(list, () => '');
        const entries = list.data?.value ?? [];
        const pick = entries.find((t) => t.isDefault) ?? entries[0];
        if (!pick) {
          skipped.push(`${i}: ${chosen.title} (${formatDate(chosen.when)}) — no transcript`);
          continue;
        }
        const transcript = await transcripts.getTranscript({ ...item.data, transcriptId: pick.id });
        return toolResult(transcript, (data) => {
          const others = recordings.map((r, j) => j === i ? null : `${j}: ${formatDate(r.when)}`).filter(Boolean);
          const header = [
            `${chosen.title} — recorded ${formatDate(chosen.when)}, ${(data?.entries ?? []).length} entries, language ${pick.languageTag ?? '?'}`,
            recordings.length > 1 ? `(recording ${i} of ${recordings.length}; others: ${others.join(', ')})` : '',
            skipped.length ? `Skipped: ${skipped.join('; ')}` : '',
          ].filter(Boolean).join('\n');
          return `${header}\n\n${formatTranscript(data?.entries, maxChars)}`;
        });
      }
      return { isError: true, content: [{ type: 'text', text: `No usable transcript. ${skipped.join('; ')}` }] };
    });
  }

  server.registerTool('teams_search_messages', {
    title: 'Search Teams messages',
    description: 'Case-insensitive text search over messages in the most recently active conversations (a bounded client-side scan, not a server index). Use for "what did X say about Y".',
    inputSchema: {
      query: z.string().min(2),
      sinceDays: z.number().int().min(1).max(730).default(14).describe('How far back to look; long windows page through history (about 0.7 s per 200 messages)'),
      chatLimit: z.number().int().min(1).max(100).default(15).describe('How many recently active conversations to scan'),
      kinds: kindsSchema,
      from: z.string().optional().describe('Only messages whose author name contains this'),
      maxHits: z.number().int().min(1).max(200).default(40),
    },
  }, async ({ query, sinceDays, chatLimit, kinds, from, maxHits }) => {
    const list = await chat.getConversations(Math.max(chatLimit * 3, 100));
    if (!list.success) return toolResult(list, () => '');
    await resolvePartnerNames(list.data?.conversations ?? [], graph, 0);

    const needle = query.toLowerCase();
    const author = from?.toLowerCase();
    const since = Date.now() - sinceDays * 86_400_000;
    const hits = [];
    let scanned = 0;

    for (const conversation of selectConversations(list.data?.conversations ?? [], { kinds, unreadOnly: false }).slice(0, chatLimit)) {
      if (new Date(conversation.lastMessage?.composetime ?? 0).getTime() < since) continue;
      // Page backwards through the conversation until the window is exhausted.
      let cursor;
      for (let page = 0; page < 40; page++) {
        const messages = await chat.getMessages(conversation.id, { pageSize: 200, startTime: since, cursor });
        if (!messages.success) break;
        const batch = messages.data?.messages ?? [];
        scanned += batch.length;
        for (const m of batch) {
          if (!isContentMessage(m)) continue;
          const text = messageText(m);
          const files = attachmentNames(m).join(' ');
          if (!text.toLowerCase().includes(needle) && !files.toLowerCase().includes(needle)) continue;
          if (author && !messageAuthor(m).toLowerCase().includes(author)) continue;
          hits.push({ conversation, m, text, files });
        }
        const oldest = batch.length ? new Date(batch[batch.length - 1].composetime).getTime() : 0;
        cursor = messages.data?._metadata?.backwardLink;
        if (!cursor || batch.length === 0 || oldest < since) break;
      }
    }

    hits.sort((a, b) => new Date(b.m.composetime) - new Date(a.m.composetime));
    return toolResult({ success: true, data: hits }, (all) => {
      const lines = all.slice(0, maxHits).map(({ conversation, m, text, files }) => [
        `- [${formatDate(m.composetime)}] ${messageAuthor(m)} in "${chatLabel(conversation)}"`,
        `  chatId: ${conversation.id}`,
        `  ${truncate(text.replaceAll(/\s+/g, ' '), 600)}`,
        files ? `  attachments: ${files}` : null,
      ].filter(Boolean).join('\n'));
      lines.push(`(${all.length} hit(s) in ${scanned} messages scanned over ${sinceDays} days; showing ${Math.min(all.length, maxHits)})`);
      return lines.join('\n');
    });
  });
}

module.exports = {
  registerTeamsTools, chatLabel, kindOf, isUnread, isContentMessage, messageAuthor, messageText, renderChatMessage,
  selectConversations, resolvePartnerNames, rememberAuthor, isoDaysAgo,
  parseEmotions, resolveReactionNames, reactionsLine, cardText, extractCardText, extractQuote, quoteLine,
  threadRootId, shortId, isChannelConversationId,
};
