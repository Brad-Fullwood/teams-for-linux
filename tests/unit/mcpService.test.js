'use strict';

const { describe, it, before, after } = require('node:test');
const assert = require('node:assert');
const { McpService } = require('../../app/mcp');

// Minimal stand-ins for GraphApiClient and ChatServiceClient: only the read methods the tools use.
function fakeGraph() {
	return {
		makeRequest: async (endpoint) => ({ success: true, data: endpoint.startsWith('/me') ? { id: '00000000-0000-0000-0000-000000000000', displayName: 'Test User', mail: 'test@example.com' } : {} }),
		getMailFolderMessages: async () => ({ success: true, data: { value: [] } }),
		getMailMessages: async () => ({ success: true, data: { value: [] } }),
		getMailMessage: async () => ({ success: false, error: 'not found', status: 404 }),
		getMailAttachments: async () => ({ success: true, data: { value: [] } }),
	};
}

function fakeChat() {
	return {
		getConversations: async () => ({ success: true, data: { conversations: [] } }),
		getMessages: async () => ({ success: true, data: { messages: [] } }),
		getTeamsAndChannels: async () => ({ success: true, data: { teams: [] } }),
	};
}

const PORT = 3999;
const URL_ = `http://127.0.0.1:${PORT}/mcp`;
const HEADERS = { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' };

function rpc(method, params = {}, id = 1) {
	return JSON.stringify({ jsonrpc: '2.0', id, method, params });
}

async function initialize(extraHeaders = {}) {
	return fetch(URL_, {
		method: 'POST',
		headers: { ...HEADERS, ...extraHeaders },
		body: rpc('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '0' } }),
	});
}

describe('McpService', () => {
	let service;

	before(async () => {
		service = new McpService({ mcp: { enabled: true, port: PORT, authToken: 'secret' } });
		await service.initialize(fakeGraph(), fakeChat(), { findRecordings: async () => ({ success: true, data: [] }) });
	});

	after(async () => {
		await service.shutdown();
	});

	it('does nothing when disabled', async () => {
		const off = new McpService({ mcp: { enabled: false } });
		await off.initialize(fakeGraph(), fakeChat());
		await off.shutdown();
	});

	it('rejects requests without the bearer token', async () => {
		const res = await initialize();
		assert.strictEqual(res.status, 401);
	});

	it('rejects non-loopback origins', async () => {
		const res = await initialize({ Authorization: 'Bearer secret', Origin: 'https://evil.example' });
		assert.strictEqual(res.status, 403);
	});

	it('refuses GET and DELETE (stateless mode has no stream or session)', async () => {
		const get = await fetch(URL_, { headers: { Authorization: 'Bearer secret', Accept: 'text/event-stream' } });
		assert.strictEqual(get.status, 405);
		const del = await fetch(URL_, { method: 'DELETE', headers: { Authorization: 'Bearer secret' } });
		assert.strictEqual(del.status, 405);
	});

	it('returns 404 off the /mcp path', async () => {
		const res = await fetch(`http://127.0.0.1:${PORT}/other`, { headers: { Authorization: 'Bearer secret' } });
		assert.strictEqual(res.status, 404);
	});

	it('lists only read-only tools and executes one', async () => {
		const auth = { Authorization: 'Bearer secret', Origin: 'http://localhost' };
		const init = await initialize(auth);
		assert.strictEqual(init.status, 200);

		const list = await fetch(URL_, { method: 'POST', headers: { ...HEADERS, ...auth }, body: rpc('tools/list', {}, 2) });
		const listBody = await list.json();
		const names = listBody.result.tools.map((t) => t.name).sort();
		assert.deepStrictEqual(names, [
			'files_get_content', 'files_shared_with_me', 'mail_get_attachment', 'mail_get_message', 'mail_list_attachments', 'mail_list_messages', 'mail_search', 'me',
			'teams_get_chat_messages', 'teams_get_meeting_transcript', 'teams_list_channels', 'teams_list_chats', 'teams_search_messages', 'triage_digest',
		]);

		const call = await fetch(URL_, { method: 'POST', headers: { ...HEADERS, ...auth }, body: rpc('tools/call', { name: 'me', arguments: {} }, 3) });
		const callBody = await call.json();
		assert.strictEqual(callBody.result.content[0].text, 'Test User <test@example.com>');

		const failing = await fetch(URL_, { method: 'POST', headers: { ...HEADERS, ...auth }, body: rpc('tools/call', { name: 'mail_get_message', arguments: { id: 'x' } }, 4) });
		const failingBody = await failing.json();
		assert.strictEqual(failingBody.result.isError, true);
		assert.match(failingBody.result.content[0].text, /HTTP 404/);
	});
});
