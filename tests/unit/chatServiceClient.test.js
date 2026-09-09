'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const ChatServiceClient = require('../../app/chatService');

// The chat service client is read-only: no method may issue anything but GET,
// and the module must never persist or log the Skype token.
describe('ChatServiceClient', () => {
	it('only issues GET requests and persists nothing', () => {
		const dir = path.join(__dirname, '..', '..', 'app', 'chatService');
		for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.js'))) {
			const src = fs.readFileSync(path.join(dir, file), 'utf8');
			assert.ok(!/method:\s*['"`](POST|PATCH|PUT|DELETE)['"`]/.test(src), `${file} uses a write verb`);
			assert.ok(!/fs\.|writeFile|localStorage/.test(src), `${file} must not persist anything`);
		}
	});

	it('is disabled unless graphApi.enabled', async () => {
		const off = new ChatServiceClient({ graphApi: { enabled: false } });
		assert.strictEqual(off.isEnabled(), false);
		assert.deepStrictEqual(await off.getConversations(), { success: false, error: 'Graph API is disabled' });
	});

	it('fails cleanly without a main window', async () => {
		const client = new ChatServiceClient({ graphApi: { enabled: true } });
		const result = await client.getConversations();
		assert.strictEqual(result.success, false);
		assert.match(result.error, /Main window/);
	});

	it('builds conversation and message URLs with the session from the renderer', async () => {
		const client = new ChatServiceClient({ graphApi: { enabled: true } });
		client.initialize({
			webContents: {
				executeJavaScript: async () => ({ success: true, skypeToken: 'tok', expiration: Date.now() + 3_600_000, chatService: 'https://uk.ng.msg.teams.microsoft.com/' }),
			},
		});
		const calls = [];
		const originalFetch = globalThis.fetch;
		globalThis.fetch = async (url, init) => {
			calls.push({ url, init });
			return { ok: true, status: 200, text: async () => JSON.stringify({ messages: [] }) };
		};
		try {
			const result = await client.getMessages('19:abc@thread.v2', { pageSize: 5, startTime: 123 });
			assert.strictEqual(result.success, true);
			assert.strictEqual(calls[0].url, 'https://uk.ng.msg.teams.microsoft.com/v1/users/ME/conversations/19%3Aabc%40thread.v2/messages?view=msnp24Equivalent&pageSize=5&startTime=123');
			assert.strictEqual(calls[0].init.method, 'GET');
			assert.strictEqual(calls[0].init.headers.Authentication, 'skypetoken=tok');
			const older = await client.getMessages('19:abc@thread.v2', { cursor: 'https://uk.ng.msg.teams.microsoft.com/v1/users/ME/conversations/x/messages?syncState=abc' });
			assert.strictEqual(older.success, true);
			assert.strictEqual(calls[1].url, 'https://uk.ng.msg.teams.microsoft.com/v1/users/ME/conversations/x/messages?syncState=abc');
			const foreign = await client.getMessages('19:abc@thread.v2', { cursor: 'https://evil.example/v1/x' });
			assert.strictEqual(foreign.success, false);
			assert.strictEqual(calls.length, 2, 'foreign cursor must not be fetched');
			await client.getMessages('19:abc@thread.v2', { pageSize: 9999 });
			assert.match(calls[2].url, /pageSize=200$/);
		} finally {
			globalThis.fetch = originalFetch;
		}
	});
});

describe('transcripts', () => {
	const { TranscriptClient, formatTranscript } = require('../../app/chatService/transcripts');

	it('formats entries as timed speaker turns and merges consecutive lines', () => {
		const text = formatTranscript([
			{ speakerDisplayName: 'Ann', text: 'Hello.', startOffset: '00:00:03.6340146' },
			{ speakerDisplayName: 'Ann', text: 'Are we all here?', startOffset: '00:00:05.1' },
			{ speakerDisplayName: 'Bob', text: 'Yes.', startOffset: '00:00:07.0' },
			{ speakerDisplayName: 'Bob', text: '   ', startOffset: '00:00:08.0' },
		]);
		assert.strictEqual(text, '[00:00:03] Ann: Hello. Are we all here?\n[00:00:07] Bob: Yes.');
	});

	it('truncates at maxChars with a note', () => {
		const text = formatTranscript([{ speakerDisplayName: 'A', text: 'x'.repeat(2000), startOffset: '0' }], 1000);
		assert.match(text, /truncated at 1000 chars; 1 turns/);
	});

	it('finds de-duplicated recording links across message pages', async () => {
		const pages = [
			{ success: true, data: { messages: [
				{ messagetype: 'RichText/Media_CallRecording', content: '<URIObject><Title>Sync</Title></URIObject>', composetime: '2026-09-01T10:00:00Z', properties: { atp: JSON.stringify([{ URL: 'https://x.sharepoint.com/a' }]) } },
				{ messagetype: 'RichText/Media_CallRecording', content: '', composetime: '2026-09-01T09:00:00Z', properties: {} },
				{ messagetype: 'RichText/Html', content: '<p>hi</p>' },
			], _metadata: { backwardLink: 'https://x/next' } } },
			{ success: true, data: { messages: [
				{ messagetype: 'RichText/Media_CallRecording', content: '<URIObject><Title>Sync</Title></URIObject>', composetime: '2026-09-01T10:00:00Z', properties: { atp: JSON.stringify([{ URL: 'https://x.sharepoint.com/a' }]) } },
				{ messagetype: 'RichText/Media_CallRecording', content: '<URIObject><Title>Older</Title></URIObject>', composetime: '2026-08-01T10:00:00Z', properties: { atp: JSON.stringify([{ URL: 'https://x.sharepoint.com/b' }]) } },
			], _metadata: {} } },
		];
		const chat = { getMessages: async () => pages.shift() };
		const client = new TranscriptClient({ graphApiClient: {}, chatServiceClient: chat });
		const found = await client.findRecordings('19:m@thread.v2');
		assert.deepStrictEqual(found.data.map((r) => [r.title, r.url]), [['Sync', 'https://x.sharepoint.com/a'], ['Older', 'https://x.sharepoint.com/b']]);
	});

	it('resolves a sharing link through Graph into site, drive and item ids', async () => {
		const seen = [];
		const graph = { makeRequest: async (ep) => { seen.push(ep); return { success: true, data: { id: 'ITEM', name: 'rec.mp4', parentReference: { driveId: 'DRIVE' }, sharepointIds: { siteUrl: 'https://x-my.sharepoint.com/personal/u' } } }; } };
		const client = new TranscriptClient({ graphApiClient: graph, chatServiceClient: {} });
		const item = await client.resolveItem('https://x-my.sharepoint.com/:v:/g/personal/u/abc');
		assert.deepStrictEqual(item.data, { siteUrl: 'https://x-my.sharepoint.com/personal/u', driveId: 'DRIVE', itemId: 'ITEM', name: 'rec.mp4' });
		assert.match(seen[0], /^\/shares\/u!aHR0cHM6Ly94LW15LnNoYXJlcG9pbnQuY29tLzp2Oi9nL3BlcnNvbmFsL3UvYWJj\/driveItem/);
	});
});
