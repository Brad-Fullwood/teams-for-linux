'use strict';

const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');

const {
  renderChatMessage, messageText, parseEmotions, resolveReactionNames, reactionsLine,
  cardText, extractQuote, quoteLine, threadRootId, shortId, isChannelConversationId,
  rememberAuthor,
} = require('../../app/mcp/tools/teams');

describe('A. teams_get_chat_messages maxChars', () => {
  it('defaults to a 2000-character cut and honours a smaller maxChars', () => {
    const long = 'x'.repeat(2500);
    const m = { messagetype: 'Text', content: long, composetime: '2026-09-08T10:00:00Z', imdisplayname: 'A' };
    assert.equal(renderChatMessage(m).split(': ')[1].length, 2000); // default truncate(…, 2000) incl. ellipsis
    const short = renderChatMessage(m, { maxChars: 200 });
    assert.equal(short.split(': ')[1].length, 200);
  });

  it('keeps short messages intact regardless of maxChars', () => {
    const m = { messagetype: 'Text', content: 'hi there', composetime: '2026-09-08T10:00:00Z', imdisplayname: 'A' };
    assert.match(renderChatMessage(m, { maxChars: 300 }), /: hi there$/);
  });
});

describe('B. reactions', () => {
  beforeEach(() => {
    // Force a fresh author cache entry so name resolution assertions are deterministic per test file run.
    rememberAuthor({ from: 'https://x/contacts/8:orgid:11111111-1111-1111-1111-111111111111', imdisplayname: 'Author' });
  });

  it('parses emotions from a JSON string and from an array', () => {
    const asString = { properties: { emotions: JSON.stringify([{ key: 'like', users: [{ mri: '8:orgid:aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', time: 1 }] }]) } };
    const asArray = { properties: { emotions: [{ key: 'heart', users: [] }] } };
    assert.deepEqual(parseEmotions(asString), [{ key: 'like', users: [{ mri: '8:orgid:aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', time: 1 }] }]);
    assert.deepEqual(parseEmotions(asArray), [{ key: 'heart', users: [] }]);
  });

  it('parses defensively: missing field, malformed JSON, non-array, missing users, unknown keys', () => {
    assert.deepEqual(parseEmotions({}), []);
    assert.deepEqual(parseEmotions({ properties: { emotions: '{not json' } }), []);
    assert.deepEqual(parseEmotions({ properties: { emotions: '{}' } }), []);
    assert.deepEqual(parseEmotions({ properties: { emotions: [{ key: 'mystery' }] } }), [{ key: 'mystery', users: [] }]);
  });

  it('resolves MRIs to names via Graph, bounded by maxLookups, and caches across calls', async () => {
    const oidA = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
    const oidB = 'cccccccc-cccc-cccc-cccc-cccccccccccc';
    const calls = [];
    const graph = {
      makeRequest: async (ep) => {
        calls.push(ep);
        return { success: true, data: { displayName: ep.includes(oidA) ? 'Alice' : 'Bob' } };
      },
    };
    const messages = [
      { properties: { emotions: JSON.stringify([{ key: 'like', users: [{ mri: `8:orgid:${oidA}` }, { mri: `8:orgid:${oidB}` }] }]) } },
    ];
    await resolveReactionNames(messages, graph, 1); // bounded to 1 lookup
    assert.equal(calls.length, 1);
    const line = reactionsLine(messages[0]);
    assert.match(line, /reactions: like×2 \(Alice\)/); // only the resolved one has a name

    await resolveReactionNames(messages, graph, 50);
    assert.equal(calls.length, 2);
    assert.match(reactionsLine(messages[0]), /reactions: like×2 \(Alice, Bob\)/);

    await resolveReactionNames(messages, graph, 50);
    assert.equal(calls.length, 2, 'cached, no new lookups');
  });

  it('shows the count without names when resolution is unavailable', () => {
    const m = { properties: { emotions: JSON.stringify([{ key: 'laugh', users: [{ mri: '8:orgid:dddddddd-dddd-dddd-dddd-dddddddddddd' }] }]) } };
    assert.equal(reactionsLine(m), '  reactions: laugh×1');
  });

  it('renders multiple reaction kinds on one line, and nothing when there are none', () => {
    const m = {
      properties: {
        emotions: JSON.stringify([
          { key: 'like', users: [{ mri: '8:orgid:11111111-1111-1111-1111-111111111111' }] },
          { key: 'heart', users: [] },
        ]),
      },
    };
    assert.match(reactionsLine(m), /^ {2}reactions: like×1 \(Author\)$/);
    assert.equal(reactionsLine({ properties: {} }), null);
  });
});

describe('C. adaptive / rich cards', () => {
  it('extracts TextBlock, RichTextBlock, FactSet and nested Container/ColumnSet text', () => {
    const card = {
      type: 'AdaptiveCard',
      body: [
        { type: 'TextBlock', text: 'Great job on the release!' },
        { type: 'Image', url: 'https://example.com/badge.png' },
        { type: 'ColumnSet', columns: [{ type: 'Column', items: [{ type: 'TextBlock', text: 'From Alice' }] }] },
        { type: 'FactSet', facts: [{ title: 'Value', value: 'Teamwork' }] },
        { type: 'RichTextBlock', inlines: [{ text: 'Nice ' }, { text: 'work' }] },
      ],
    };
    const m = { properties: { cards: JSON.stringify([{ content: card }]) } };
    assert.equal(cardText(m, 2000), 'Great job on the release! | From Alice | Value: Teamwork | Nice work');
  });

  it('reads message.attachments with content.body as an alternative source', () => {
    const m = { attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: { body: [{ type: 'TextBlock', text: 'Praise: excellent work' }] } }] };
    assert.equal(cardText(m, 2000), 'Praise: excellent work');
  });

  it('is bounded by maxChars and returns null when there is no visible text', () => {
    const m = { properties: { cards: JSON.stringify([{ body: [{ type: 'TextBlock', text: 'x'.repeat(50) }] }]) } };
    assert.equal(cardText(m, 10).length, 10);
    assert.equal(cardText({ properties: { cards: JSON.stringify([{ body: [{ type: 'Image', url: 'https://x/y.png' }] }]) } }, 2000), null);
    assert.equal(cardText({}, 2000), null);
  });

  it('is defensive against malformed cards JSON', () => {
    assert.equal(cardText({ properties: { cards: 'not json' } }, 2000), null);
    assert.equal(cardText({ properties: { cards: '["also not json"' } }, 2000), null);
  });

  it('renderChatMessage renders a card line for a Praise-shaped message', () => {
    const m = {
      messagetype: 'RichText/Media_Card', content: '<URIObject><Title>Praise</Title></URIObject>',
      composetime: '2026-09-08T10:00:00Z', imdisplayname: 'Boss',
      properties: { cards: JSON.stringify([{ body: [{ type: 'TextBlock', text: 'Amazing work this sprint' }] }]) },
    };
    assert.match(renderChatMessage(m), /^ {2}card: Amazing work this sprint$/m);
  });
});

describe('D. quoted replies', () => {
  const html = '<div><blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="123"><strong itemprop="mri" itemid="8:orgid:x">Jane Doe</strong><span itemprop="time" itemid="1700000000000">t</span><p itemprop="preview">original message text</p></blockquote>My reply text</div>';

  it('strips the quote from the message text, leaving only the reply', () => {
    const m = { messagetype: 'RichText/Html', content: html };
    assert.equal(messageText(m), 'My reply text');
  });

  it('extracts author, time and preview from the blockquote', () => {
    const { quote, html: stripped } = extractQuote(html);
    assert.equal(quote.author, 'Jane Doe');
    assert.equal(quote.time, 1700000000000);
    assert.equal(quote.preview, 'original message text');
    assert.doesNotMatch(stripped, /blockquote/);
  });

  it('renders a separate quoting line in renderChatMessage for a reply with a quote', () => {
    const m = { messagetype: 'RichText/Html', content: html, composetime: '2026-09-08T10:00:00Z', imdisplayname: 'Replier' };
    const rendered = renderChatMessage(m);
    assert.match(rendered, /^\[.*\] Replier: My reply text$/m);
    assert.match(rendered, /^ {2}> quoting Jane Doe \(.*\): original message text$/m);
  });

  it('renders no quote line for a plain reply without a quote', () => {
    const m = { messagetype: 'RichText/Html', content: '<p>just a normal reply</p>', composetime: '2026-09-08T10:00:00Z', imdisplayname: 'Replier' };
    const rendered = renderChatMessage(m);
    assert.doesNotMatch(rendered, /quoting/);
    assert.equal(quoteLine(extractQuote(m.content).quote), null);
  });

  it('handles a quote with missing parts gracefully', () => {
    const partial = '<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="1"><p itemprop="preview">only a preview</p></blockquote>reply';
    const { quote } = extractQuote(partial);
    assert.equal(quote.author, null);
    assert.equal(quote.time, null);
    assert.equal(quoteLine(quote), '  > quoting someone: only a preview');

    const noPreviewNoAuthor = '<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="1"><span itemprop="time" itemid="1700000000000">t</span></blockquote>reply';
    assert.equal(quoteLine(extractQuote(noPreviewNoAuthor).quote), null);
  });
});

describe('E. channel threads', () => {
  it('detects a root/parent id from properties.rootMessageId, rootMessageId, parentMessageId, or conversationLink', () => {
    assert.equal(threadRootId({ properties: { rootMessageId: '111' } }), '111');
    assert.equal(threadRootId({ rootMessageId: '222' }), '222');
    assert.equal(threadRootId({ parentMessageId: '333' }), '333');
    assert.equal(threadRootId({ conversationLink: 'https://x/v1/threads/19:abc@thread.tacv2;messageid=444' }), '444');
    assert.equal(threadRootId({}), null);
  });

  it('shortens ids for compact display', () => {
    assert.equal(shortId('1690000000123'), '00000123');
    assert.equal(shortId('42'), '42');
    assert.equal(shortId(undefined), '');
  });

  it('identifies standard/private channel conversation ids heuristically', () => {
    assert.equal(isChannelConversationId('19:abcdef@thread.tacv2'), true);
    assert.equal(isChannelConversationId('19:abcdef@thread.skype'), true);
    assert.equal(isChannelConversationId('19:oid1_oid2@unq.gbl.spaces'), false);
    assert.equal(isChannelConversationId('19:abcdef@thread.v2'), false);
  });

  it('renders "(reply in thread <short>)" for a reply and "[id <short>]" for a channel root post', () => {
    const reply = { id: '1690000000999', messagetype: 'Text', content: 'a reply', composetime: '2026-09-08T10:00:00Z', imdisplayname: 'A', properties: { rootMessageId: '1690000000111' } };
    const rendered = renderChatMessage(reply, { isChannel: true });
    assert.match(rendered, /\(reply in thread 00000111\)$/);
    assert.doesNotMatch(rendered, /\[id /);

    const root = { id: '1690000000222', messagetype: 'Text', content: 'a root post', composetime: '2026-09-08T10:00:00Z', imdisplayname: 'A' };
    const renderedRoot = renderChatMessage(root, { isChannel: true });
    assert.match(renderedRoot, /\[id 00000222\] A:/);
  });

  it('does not tag ids or threads for non-channel conversations', () => {
    const m = { id: '1690000000222', messagetype: 'Text', content: 'hi', composetime: '2026-09-08T10:00:00Z', imdisplayname: 'A' };
    assert.doesNotMatch(renderChatMessage(m, { isChannel: false }), /\[id /);
  });
});
