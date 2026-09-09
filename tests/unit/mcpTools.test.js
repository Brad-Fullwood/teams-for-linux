'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const { htmlToText, odataString, toolResult } = require('../../app/mcp/tools/format');
const { buildMailFilter, renderMessageLine } = require('../../app/mcp/tools/mail');
const { chatLabel, renderChatMessage, isUnread, kindOf, selectConversations, isContentMessage, messageText, resolvePartnerNames } = require('../../app/mcp/tools/teams');

const MCP_DIR = path.join(__dirname, '..', '..', 'app', 'mcp');

// The MCP module is documented as read-only. Every Graph write method on
// GraphApiClient, and every HTTP verb other than GET, must stay out of app/mcp/.
const WRITE_METHODS = ['createCalendarEvent', 'updateCalendarEvent', 'deleteCalendarEvent', 'sendChatMessageToUser', 'resolveConversation'];
const WRITE_VERBS = /method:\s*['"`](POST|PATCH|PUT|DELETE)['"`]/;

function* jsFiles(dir) {
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name);
		if (entry.isDirectory()) yield* jsFiles(full);
		else if (entry.name.endsWith('.js')) yield full;
	}
}

describe('mcp read-only guard', () => {
	it('never references a Graph write method or a non-GET verb', () => {
		for (const file of jsFiles(MCP_DIR)) {
			const src = fs.readFileSync(file, 'utf8');
			for (const name of WRITE_METHODS) {
				assert.ok(!src.includes(name), `${path.relative(MCP_DIR, file)} references ${name}`);
			}
			assert.ok(!WRITE_VERBS.test(src), `${path.relative(MCP_DIR, file)} uses a write HTTP verb`);
		}
	});
});

describe('htmlToText', () => {
	it('flattens Teams message HTML including mentions and entities', () => {
		const html = '<p>Hi <at id="0">Brad Fullwood</at>,<br>can you look at &quot;ticket&quot; &#35;42?</p><ul><li>one</li><li>two</li></ul>';
		assert.strictEqual(htmlToText(html), 'Hi @Brad Fullwood,\ncan you look at "ticket" #42?\n- one\n- two');
	});

	it('keeps link targets so shared documents stay reachable', () => {
		assert.strictEqual(htmlToText('<p>See <a href="https://x.sharepoint.com/:x:/r/doc">Astonish scenario testing</a> now</p>'), 'See Astonish scenario testing (https://x.sharepoint.com/:x:/r/doc) now');
		assert.strictEqual(htmlToText('<a href="https://x/y">https://x/y</a>'), 'https://x/y');
		assert.strictEqual(htmlToText('<a href="mailto:a@b.c">mail me</a>'), 'mail me');
	});

	it('drops style and script blocks', () => {
		assert.strictEqual(htmlToText('<style>p{}</style><script>x()</script><div>ok</div>'), 'ok');
	});

	it('handles empty input', () => {
		assert.strictEqual(htmlToText(undefined), '');
	});
});

describe('buildMailFilter', () => {
	it('always leads with receivedDateTime so Graph accepts the $orderby', () => {
		assert.strictEqual(buildMailFilter({}), 'receivedDateTime ge 1970-01-01T00:00:00Z');
		assert.match(buildMailFilter({ sinceDays: 1 }), /^receivedDateTime ge \d{4}-\d{2}-\d{2}T[\d:.]+Z$/);
	});

	it('appends the other clauses with and', () => {
		const filter = buildMailFilter({ unreadOnly: true, flaggedOnly: true, from: "o'neil", subjectContains: 'x' });
		assert.match(filter, /^receivedDateTime ge .* and isRead eq false and flag\/flagStatus eq 'flagged' and \(from\/emailAddress\/address eq 'o''neil' or contains\(from\/emailAddress\/name,'o''neil'\) or contains\(from\/emailAddress\/address,'o''neil'\)\) and contains\(subject,'x'\)$/);
	});

	it('escapes single quotes for OData', () => {
		assert.strictEqual(odataString("it's"), "it''s");
	});
});

describe('renderers', () => {
	it('renders a mail line with flags', () => {
		const line = renderMessageLine({
			id: 'AAA', subject: 'Outage', isRead: false, hasAttachments: true, importance: 'high',
			flag: { flagStatus: 'flagged' }, receivedDateTime: '2026-09-08T09:30:00Z',
			from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } }, bodyPreview: 'Line one\nLine two',
		});
		assert.match(line, /\[2026-09-08 09:30Z\] Outage \[UNREAD,FLAGGED,ATT,HIGH\]/);
		assert.match(line, /from: Alice <alice@example.com>/);
		assert.match(line, /preview: Line one Line two/);
	});

	it('labels a conversation by channel topic, chat topic, or 1:1 partner', () => {
		assert.strictEqual(chatLabel({ threadProperties: { threadType: 'space', spaceThreadTopic: 'Business Apps' } }), 'Business Apps');
		assert.strictEqual(chatLabel({ threadProperties: { threadType: 'chat', topic: 'Project X' } }), 'Project X');
		assert.strictEqual(chatLabel({ threadProperties: { threadType: 'chat', productThreadType: 'OneToOneChat' }, lastMessage: { imdisplayname: 'Ann' } }), '1:1 with Ann');
		assert.strictEqual(chatLabel({ threadProperties: { threadType: 'meeting', productThreadType: 'Meeting' } }), '(Meeting)');
	});

	it('names 1:1 chats from the partner id via Graph, then from cache', async () => {
		const me = '11111111-1111-1111-1111-111111111111';
		const other = '22222222-2222-2222-2222-222222222222';
		const calls = [];
		const graph = { makeRequest: async (ep) => { calls.push(ep); return ep.startsWith('/me') ? { success: true, data: { id: me } } : { success: true, data: { displayName: 'Ann Other' } }; } };
		const conv = { id: `19:${me}_${other}@unq.gbl.spaces`, threadProperties: { threadType: 'chat', productThreadType: 'OneToOneChat' }, lastMessage: { imdisplayname: 'Me', from: `https://x/contacts/8:orgid:${me}` } };
		await resolvePartnerNames([conv], graph);
		assert.strictEqual(chatLabel(conv), '1:1 with Ann Other');
		await resolvePartnerNames([conv], graph);
		assert.strictEqual(calls.filter((c) => c.startsWith('/users/')).length, 1, 'second call served from cache');
	});

	it('classifies kinds and drops notification streams', () => {
		assert.strictEqual(kindOf({ threadProperties: { threadType: 'space' } }), 'channel');
		assert.strictEqual(kindOf({ threadProperties: { threadType: 'topic' } }), 'channel');
		assert.strictEqual(kindOf({ threadProperties: { threadType: 'streamofnotes' } }), 'notes');
		assert.strictEqual(chatLabel({ threadProperties: { threadType: 'streamofnotes' } }), 'Notes to self');
		assert.strictEqual(kindOf({ threadProperties: { threadType: 'meeting' } }), 'meeting');
		assert.strictEqual(kindOf({ threadProperties: { threadType: 'streamofmentions' } }), null);
		const list = [
			{ id: 'a', threadProperties: { threadType: 'chat' }, lastMessage: { id: '2', composetime: '2026-09-08T10:00:00Z' }, properties: { consumptionhorizon: '1;1;1' } },
			{ id: 'b', threadProperties: { threadType: 'streamofcalllogs' }, lastMessage: { id: '9', composetime: '2026-09-08T12:00:00Z' } },
			{ id: 'c', threadProperties: { threadType: 'chat' }, lastMessage: { id: '3', composetime: '2026-09-08T11:00:00Z' }, properties: { consumptionhorizon: '3;1;1' } },
		];
		assert.deepStrictEqual(selectConversations(list, { kinds: ['chat'], unreadOnly: false }).map((c) => c.id), ['c', 'a']);
		assert.deepStrictEqual(selectConversations(list, { kinds: ['chat'], unreadOnly: true }).map((c) => c.id), ['a']);
	});

	it('derives unread from the consumption horizon', () => {
		assert.strictEqual(isUnread({ lastMessage: { id: '1788882060523' }, properties: { consumptionhorizon: '1788882060523;1788882399699;142' } }), false);
		assert.strictEqual(isUnread({ lastMessage: { id: '1788882060600' }, properties: { consumptionhorizon: '1788882060523;1788882399699;142' } }), true);
		assert.strictEqual(isUnread({ lastMessage: { id: '5' }, properties: {} }), true);
		assert.strictEqual(isUnread({ properties: {} }), false);
	});

	it('summarises call events and media messages instead of dumping their XML', () => {
		assert.strictEqual(messageText({ messagetype: 'Event/Call', content: '<ended/><partlist count="4"><part><duration>1234.5</duration></part></partlist>' }), '[call ended, 4 participants, 21 min]');
		assert.strictEqual(messageText({ messagetype: 'RichText/Media_CallRecording', content: '<URIObject url="x"><Title>Weekly sync</Title></URIObject>', properties: { atp: JSON.stringify([{ URL: 'https://x.sharepoint.com/rec' }]) } }), '[call recording: Weekly sync https://x.sharepoint.com/rec]');
		assert.strictEqual(messageText({ messagetype: 'RichText/Media_CallLogVoicemail', content: 'Voicemail Call Logs for Call 1' }), '[voicemail]');
		assert.strictEqual(isContentMessage({ messagetype: 'ThreadActivity/AddMember', content: '<addmember/>' }), false);
		assert.strictEqual(isContentMessage({ messagetype: 'RichText', content: '' }), false);
		assert.strictEqual(isContentMessage({ messagetype: 'Text', content: 'hi' }), true);
	});

	it('renders a chat service message with author and attachments', () => {
		const text = renderChatMessage({
			composetime: '2026-09-08T10:00:00.0000000Z', messagetype: 'RichText/Html', imdisplayname: 'Joe Bloggs',
			content: '<p>see attached <span itemtype="http://schema.skype.com/Mention">Brad</span></p>',
			properties: { files: JSON.stringify([{ fileName: 'spec.pdf' }]) },
		});
		assert.strictEqual(text, '[2026-09-08 10:00Z] Joe Bloggs: see attached Brad\n  attachments: spec.pdf');
	});
});

describe('truncate', () => {
	const { truncate } = require('../../app/mcp/tools/format');

	it('strips zero-width padding and collapses runs of spaces', () => {
		assert.strictEqual(truncate('See more\u034f \u200c \u034f \u200c   now'), 'See more now');
	});

	it('truncates with an ellipsis', () => {
		assert.strictEqual(truncate('abcdefgh', 5), 'abcd…');
	});
});

describe('toolResult', () => {
	it('turns a failed Graph result into an isError result', () => {
		const r = toolResult({ success: false, error: 'Forbidden', status: 403 }, () => 'x');
		assert.strictEqual(r.isError, true);
		assert.match(r.content[0].text, /HTTP 403.*Forbidden/);
	});

	it('renders success and substitutes a placeholder for empty output', () => {
		assert.strictEqual(toolResult({ success: true, data: [] }, () => '').content[0].text, '(no results)');
	});
});

describe('files', () => {
	const { docxToText, zipEntry, shareId } = require('../../app/mcp/tools/files');
	const zlib = require('node:zlib');

	// Build a tiny zip in memory: one deflated entry, central directory, EOCD.
	function makeZip(name, content) {
		const data = zlib.deflateRawSync(Buffer.from(content));
		const nameBuf = Buffer.from(name);
		const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(8, 8); local.writeUInt32LE(data.length, 18); local.writeUInt32LE(content.length, 22); local.writeUInt16LE(nameBuf.length, 26);
		const central = Buffer.alloc(46); central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(8, 10); central.writeUInt32LE(data.length, 20); central.writeUInt32LE(content.length, 24); central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(0, 42);
		const cdOffset = local.length + nameBuf.length + data.length;
		const eocd = Buffer.alloc(22); eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(1, 8); eocd.writeUInt16LE(1, 10); eocd.writeUInt32LE(central.length + nameBuf.length, 12); eocd.writeUInt32LE(cdOffset, 16);
		return Buffer.concat([local, nameBuf, data, central, nameBuf, eocd]);
	}

	it('reads a deflated zip entry', () => {
		assert.strictEqual(zipEntry(makeZip('a.txt', 'hello zip'), 'a.txt').toString(), 'hello zip');
		assert.strictEqual(zipEntry(makeZip('a.txt', 'x'), 'missing'), null);
	});

	it('flattens a Word document body into lines and tab-separated table rows', () => {
		const xml = '<w:document><w:body><w:p><w:r><w:t>Title</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>';
		assert.strictEqual(docxToText(makeZip('word/document.xml', xml)), 'Title\nA\tB');
	});

	it('encodes sharing links the way Graph expects', () => {
		assert.strictEqual(shareId('https://x/a?b=c'), 'u!aHR0cHM6Ly94L2E_Yj1j');
	});
});
