'use strict';

const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { approvalToken, NonceStore, RateLimiter, checkShape, checkApproval } = require('../../app/sendService/gate');
const { textToHtml, parseRecipient, parseRecipients } = require('../../app/sendService/format');
const { checkTarget, scopesOf, sendTeamsMessage, sendMailReply } = require('../../app/sendService/senders');
const { SendService, loadSecret } = require('../../app/sendService');
const ChatServiceClient = require('../../app/chatService');

// Test key only, used to check the HMAC against a vector computed the way the bridge does it.
const KEY = '0123456789abcdef'.repeat(4); // gitleaks:allow (test fixture)
const NOW_S = 1_760_000_000;

function signed(fields, key = KEY) {
	const request = { to: [], cc: [], expiry: NOW_S + 120, nonce: '00112233aabbccdd', ...fields };
	return { ...request, token: approvalToken(key, request) };
}

function fakeJwt(claims) {
	const part = (obj) => Buffer.from(JSON.stringify(obj)).toString('base64url');
	return `${part({ alg: 'none' })}.${part(claims)}.sig`;
}

function withFetch(handler) {
	const calls = [];
	const original = globalThis.fetch;
	globalThis.fetch = async (url, init) => {
		calls.push({ url, init, body: init?.body ? JSON.parse(init.body) : undefined });
		return handler(url, init, calls.length);
	};
	return { calls, restore: () => { globalThis.fetch = original; } };
}

const response = (status, body = '', headers = {}) => ({
	ok: status >= 200 && status < 300,
	status,
	headers: { get: (name) => headers[name.toLowerCase()] ?? null },
	text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
});

describe('send gate: approval token', () => {
	it('matches the HMAC the bridge computes in Python', () => {
		const token = approvalToken(KEY, { tool: 'teams_send_message', target: '48:notes', body: 'Hello £ ✓\n\nTwo', expiry: 1760000000, nonce: '00112233aabbccdd' });
		assert.strictEqual(token, 'e024c48321bf32e9c311d0c656f57955845acfb2eed685d04e93f78da508a033'); // gitleaks:allow (expected HMAC)
	});

	it('accepts a valid approval and rejects any changed field', () => {
		const request = signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' });
		assert.strictEqual(checkApproval(Buffer.from(KEY), request, new NonceStore(), NOW_S), null);
		for (const change of [{ body: 'hi!' }, { target: '19:x@thread.v2' }, { tool: 'mail_send_reply' }, { expiry: NOW_S + 121 }, { nonce: '00112233aabbccde' }, { token: 'abc' }]) {
			const result = checkApproval(Buffer.from(KEY), { ...request, ...change }, new NonceStore(), NOW_S);
			assert.deepStrictEqual(result, { status: 403, error: 'approval token does not match' }, JSON.stringify(change));
		}
		const otherKey = signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' }, 'x'.repeat(64));
		assert.strictEqual(checkApproval(Buffer.from(KEY), otherKey, new NonceStore(), NOW_S).status, 403);
	});

	it('enforces the expiry window: not past, at most 10 minutes ahead', () => {
		const at = (expiry) => checkApproval(Buffer.from(KEY), signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi', expiry }), new NonceStore(), NOW_S);
		assert.deepStrictEqual(at(NOW_S - 1), { status: 403, error: 'approval expired' });
		assert.strictEqual(at(NOW_S), null);
		assert.strictEqual(at(NOW_S + 600), null);
		assert.deepStrictEqual(at(NOW_S + 601), { status: 403, error: 'approval expiry too far ahead' });
	});

	it('refuses a reused nonce and prunes nonces once their approval expires', () => {
		const nonces = new NonceStore();
		const request = signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' });
		assert.strictEqual(checkApproval(Buffer.from(KEY), request, nonces, NOW_S), null);
		assert.deepStrictEqual(checkApproval(Buffer.from(KEY), request, nonces, NOW_S), { status: 403, error: 'nonce already used' });
		assert.strictEqual(nonces.size, 1);
		nonces.claim('ffffffffffffffff', NOW_S + 500, NOW_S + 200);
		assert.strictEqual(nonces.size, 1, 'the first nonce expired at NOW_S + 120 and is gone');
	});
});

describe('send gate: request shape', () => {
	const base = { tool: 'teams_send_message', target: '48:notes', body: 'hi', expiry: NOW_S, nonce: '00112233aabbccdd' };

	it('rejects unknown tools, empty target or body, and mail without to', () => {
		assert.strictEqual(checkShape(base), null);
		assert.deepStrictEqual(checkShape({ ...base, tool: 'delete_everything' }), { status: 400, error: 'unknown tool' });
		assert.deepStrictEqual(checkShape({ ...base, target: '  ' }), { status: 400, error: 'target is empty' });
		assert.deepStrictEqual(checkShape({ ...base, body: '\n' }), { status: 400, error: 'body is empty' });
		assert.deepStrictEqual(checkShape({ ...base, tool: 'mail_send_reply', to: [] }), { status: 400, error: 'mail reply has no to' });
		assert.deepStrictEqual(checkShape({ ...base, tool: 'mail_send_reply' }), { status: 400, error: 'mail reply has no to' });
		assert.strictEqual(checkShape({ ...base, nonce: 'xyz' }).status, 400);
		assert.strictEqual(checkShape({ ...base, expiry: '1' }).status, 400);
		assert.strictEqual(checkShape({ ...base, to: 'a@b.c' }).status, 400);
		assert.strictEqual(checkShape(null).status, 400);
	});

	it('takes chats and notes to self, refuses channels and bad ids', () => {
		assert.strictEqual(checkTarget('teams_send_message', '48:notes'), null);
		assert.strictEqual(checkTarget('teams_send_message', '19:abc_def@unq.gbl.spaces'), null);
		assert.strictEqual(checkTarget('teams_send_message', '19:meeting_xyz@thread.v2'), null);
		assert.deepStrictEqual(checkTarget('teams_send_message', '19:abc@thread.tacv2'), { status: 400, error: 'channel threads are not supported' });
		assert.deepStrictEqual(checkTarget('teams_send_message', '19:abc@thread.tacv2;messageid=1700000000000'), { status: 400, error: 'channel threads are not supported' });
		assert.strictEqual(checkTarget('teams_send_message', '19:abc@thread.v2/../x').status, 400);
		assert.strictEqual(checkTarget('mail_send_reply', 'AAMkAGI2THVSAAA=', ['a@b.co']), null);
		assert.strictEqual(checkTarget('mail_send_reply', 'AAMk AGI2THVSAAA=', ['a@b.co']).status, 400);
		assert.deepStrictEqual(checkTarget('mail_send_reply', 'AAMkAGI2THVSAAA=', ['not an address']), { status: 400, error: 'invalid recipient' });
	});
});

describe('rate cap', () => {
	it('allows 10 sends in a rolling hour', () => {
		const rate = new RateLimiter();
		for (let i = 0; i < 10; i++) assert.ok(rate.take(i * 1000));
		assert.strictEqual(rate.take(59 * 60 * 1000), false);
		assert.ok(rate.take(60 * 60 * 1000 + 1));
	});
});

describe('textToHtml', () => {
	it('escapes, splits paragraphs and keeps single line breaks', () => {
		assert.strictEqual(textToHtml('Hi Sam,\n\nLine one\nLine <two> & "three"\n\n\nThanks 👍'), '<p>Hi Sam,</p><p>Line one<br>Line &lt;two&gt; &amp; &quot;three&quot;</p><p>Thanks 👍</p>');
		assert.strictEqual(textToHtml('a\r\n\r\nb'), '<p>a</p><p>b</p>');
		assert.strictEqual(textToHtml('\n\nonly\n\n'), '<p>only</p>');
		assert.strictEqual(textToHtml('x\n   \ny'), '<p>x</p><p>y</p>');
	});
});

describe('recipient parsing', () => {
	it('reads named, quoted and bare addresses', () => {
		assert.deepStrictEqual(parseRecipient('Sam Jones <sam.jones@example.com>'), { name: 'Sam Jones', address: 'sam.jones@example.com' });
		assert.deepStrictEqual(parseRecipient('"Jones, Sam" <sam@example.co.uk>'), { name: 'Jones, Sam', address: 'sam@example.co.uk' });
		assert.deepStrictEqual(parseRecipient(' sam@example.com '), { name: '', address: 'sam@example.com' });
		assert.deepStrictEqual(parseRecipient('<sam@example.com>'), { name: '', address: 'sam@example.com' });
		assert.strictEqual(parseRecipient('Sam Jones'), null);
		assert.strictEqual(parseRecipient('sam@localhost'), null);
		assert.deepStrictEqual(parseRecipients(['a@b.co', 'bad']), { recipients: [{ name: '', address: 'a@b.co' }], invalid: 1 });
	});
});

describe('teams_send_message', () => {
	function client() {
		const c = new ChatServiceClient({ graphApi: { enabled: true } });
		let sessions = 0;
		c.initialize({ webContents: { executeJavaScript: async () => ({ success: true, skypeToken: `tok${++sessions}`, expiration: Date.now() + 3_600_000, chatService: 'https://uk.ng.msg.teams.microsoft.com/' }) } });
		return c;
	}

	it('posts RichText/Html to the conversation with the session token', async () => {
		const stub = withFetch(() => response(201, { OriginalArrivalTime: 1760000000123 }, { location: 'https://uk.ng.msg.teams.microsoft.com/v1/users/ME/conversations/48:notes/messages/1760000000123' }));
		try {
			const result = await sendTeamsMessage({ chatServiceClient: client(), getDisplayName: async () => 'Brad Fullwood' }, { target: '48:notes', body: 'test <1>\n\nsecond' });
			assert.deepStrictEqual(result, { status: 200, id: '1760000000123' });
			const [call] = stub.calls;
			assert.strictEqual(call.url, 'https://uk.ng.msg.teams.microsoft.com/v1/users/ME/conversations/48%3Anotes/messages');
			assert.strictEqual(call.init.method, 'POST');
			assert.strictEqual(call.init.headers.Authentication, 'skypetoken=tok1');
			assert.strictEqual(call.body.content, '<p>test &lt;1&gt;</p><p>second</p>');
			assert.strictEqual(call.body.messagetype, 'RichText/Html');
			assert.strictEqual(call.body.contenttype, 'text');
			assert.strictEqual(call.body.imdisplayname, 'Brad Fullwood');
			assert.match(call.body.clientmessageid, /^[1-9]\d{18}$/);
		} finally {
			stub.restore();
		}
	});

	it('refreshes the session once on 401 and reports other failures as 502', async () => {
		const stub = withFetch((url, init, n) => (n === 1 ? response(401) : response(403, { message: 'Forbidden' })));
		try {
			const result = await sendTeamsMessage({ chatServiceClient: client(), getDisplayName: async () => 'B' }, { target: '48:notes', body: 'x' });
			assert.deepStrictEqual(result, { status: 502, error: 'chat service HTTP 403: Forbidden' });
			assert.strictEqual(stub.calls.length, 2);
			assert.strictEqual(stub.calls[1].init.headers.Authentication, 'skypetoken=tok2');
		} finally {
			stub.restore();
		}
	});

	it('reports a timeout as 504 without retrying', async () => {
		const stub = withFetch(() => { const e = new Error('timeout'); e.name = 'TimeoutError'; throw e; });
		try {
			const result = await sendTeamsMessage({ chatServiceClient: client(), getDisplayName: async () => 'B' }, { target: '48:notes', body: 'x' });
			assert.strictEqual(result.status, 504);
			assert.strictEqual(stub.calls.length, 1);
		} finally {
			stub.restore();
		}
	});
});

describe('mail_send_reply', () => {
	const request = { target: 'AAMkAGI2THVSAAA=', body: 'Thanks\n\nBrad', to: ['Sam <sam@example.com>'], cc: ['lee@example.com'] };

	it('uses Graph /reply when the Graph token carries Mail.Send', async () => {
		const seen = [];
		const graphApiClient = {
			acquireToken: async () => ({ success: true, token: fakeJwt({ scp: 'Mail.Read Mail.Send' }) }),
			makeRequest: async (endpoint, options) => { seen.push({ endpoint, options }); return { success: true, data: null }; },
		};
		const result = await sendMailReply({ graphApiClient, outlookTokens: { get: async () => assert.fail('Outlook not needed') } }, request);
		assert.deepStrictEqual(result, { status: 200, id: null });
		assert.strictEqual(seen[0].endpoint, '/me/messages/AAMkAGI2THVSAAA%3D/reply');
		assert.strictEqual(seen[0].options.method, 'POST');
		assert.deepStrictEqual(seen[0].options.body, {
			comment: '<p>Thanks</p><p>Brad</p>',
			message: { toRecipients: [{ emailAddress: { name: 'Sam', address: 'sam@example.com' } }], ccRecipients: [{ emailAddress: { address: 'lee@example.com' } }] },
		});
	});

	it('falls back to Outlook REST when only the Outlook token carries Mail.Send', async () => {
		const graphApiClient = { acquireToken: async () => ({ success: true, token: fakeJwt({ scp: 'Mail.Read Mail.ReadWrite' }) }), makeRequest: async () => assert.fail('Graph must not be called') };
		const stub = withFetch(() => response(202));
		try {
			const result = await sendMailReply({ graphApiClient, outlookTokens: { get: async () => ({ success: true, token: fakeJwt({ scp: 'Mail.ReadWrite Mail.Send' }) }) } }, request);
			assert.deepStrictEqual(result, { status: 200, id: null });
			const [call] = stub.calls;
			assert.strictEqual(call.url, 'https://outlook.office.com/api/v2.0/me/messages/AAMkAGI2THVSAAA%3D/reply');
			assert.strictEqual(call.init.method, 'POST');
			assert.match(call.init.headers.Authorization, /^Bearer /);
			assert.deepStrictEqual(call.body, {
				Comment: '<p>Thanks</p><p>Brad</p>',
				Message: { ToRecipients: [{ EmailAddress: { Name: 'Sam', Address: 'sam@example.com' } }], CcRecipients: [{ EmailAddress: { Address: 'lee@example.com' } }] },
			});
		} finally {
			stub.restore();
		}
	});

	it('returns 501 when no token can send mail', async () => {
		const graphApiClient = { acquireToken: async () => ({ success: true, token: fakeJwt({ scp: 'Mail.Read' }) }) };
		const noScope = await sendMailReply({ graphApiClient, outlookTokens: { get: async () => ({ success: true, token: fakeJwt({ scp: 'Mail.Read' }) }) } }, request);
		assert.strictEqual(noScope.status, 501);
		assert.match(noScope.error, /^mail send not available: /);
		const noToken = await sendMailReply({ graphApiClient, outlookTokens: { get: async () => ({ success: false, error: 'Auth provider not found' }) } }, request);
		assert.strictEqual(noToken.status, 501);
	});

	it('reads scope names from a token', () => {
		assert.deepStrictEqual([...scopesOf(fakeJwt({ scp: 'a b' }))], ['a', 'b']);
		assert.strictEqual(scopesOf('not-a-jwt').size, 0);
	});
});

describe('SendService', () => {
	let service;
	let sent;

	beforeEach(async () => {
		sent = [];
		service = new SendService({
			port: 0,
			secret: KEY,
			now: () => NOW_S * 1000,
			senders: {
				teams: async (r) => { sent.push(['teams', r.target]); return { status: 200, id: '42' }; },
				mail: async (r) => { sent.push(['mail', r.target]); return { status: 200 }; },
			},
		});
		await service.initialize();
	});

	afterEach(async () => {
		await service.shutdown();
	});

	const post = async (body, headers = {}, method = 'POST', pathname = '/send') => {
		const { port } = service.address();
		const payload = typeof body === 'string' ? body : JSON.stringify(body);
		const res = await fetch(`http://127.0.0.1:${port}${pathname}`, { method, headers: { 'Content-Type': 'application/json', ...headers }, body: method === 'POST' ? payload : undefined });
		return { status: res.status, body: await res.json() };
	};

	it('listens on loopback and sends an approved request once', async () => {
		assert.strictEqual(service.address().address, '127.0.0.1');
		const request = signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' });
		assert.deepStrictEqual(await post(request), { status: 200, body: { ok: true, id: '42' } });
		assert.deepStrictEqual(await post(request), { status: 403, body: { error: 'nonce already used' } });
		const mail = signed({ tool: 'mail_send_reply', target: 'AAMkAGI2THVSAAA=', body: 'x', to: ['a@b.co'], nonce: '0000000000000001' });
		assert.deepStrictEqual(await post(mail), { status: 200, body: { ok: true, id: null } });
		assert.deepStrictEqual(sent, [['teams', '48:notes'], ['mail', 'AAMkAGI2THVSAAA=']]);
	});

	it('rejects bad requests before sending anything', async () => {
		assert.strictEqual((await post({}, {}, 'GET')).status, 405);
		assert.strictEqual((await post({}, {}, 'POST', '/mcp')).status, 404);
		assert.strictEqual((await post(signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' }), { Origin: 'https://evil.example' })).status, 403);
		assert.strictEqual((await post('{not json')).status, 400);
		assert.strictEqual((await post({ ...signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi' }), body: 'changed' })).status, 403);
		assert.strictEqual((await post(signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi', expiry: NOW_S - 5 }))).status, 403);
		assert.strictEqual((await post(signed({ tool: 'mail_send_reply', target: 'AAMkAGI2THVSAAA=', body: 'x' }))).status, 400);
		assert.deepStrictEqual(sent, []);
	});

	it('caps sends at 10 an hour with 429', async () => {
		for (let i = 0; i < 10; i++) {
			const r = await post(signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi', nonce: i.toString(16).padStart(16, '0') }));
			assert.strictEqual(r.status, 200);
		}
		const over = await post(signed({ tool: 'teams_send_message', target: '48:notes', body: 'hi', nonce: 'aaaaaaaaaaaaaaaa' }));
		assert.deepStrictEqual(over, { status: 429, body: { error: 'send limit reached (10 an hour)' } });
		assert.strictEqual(sent.length, 10);
	});
});

describe('send secret file', () => {
	const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'send-secret-'));
	const file = (name, content, mode) => {
		const p = path.join(dir, name);
		fs.writeFileSync(p, content, { mode });
		fs.chmodSync(p, mode);
		return p;
	};

	it('loads a mode 600 file of 32 bytes or more, trimmed', () => {
		const loaded = loadSecret(file('ok', `${KEY}\n`, 0o600));
		assert.strictEqual(loaded.secret.toString(), KEY);
		assert.deepStrictEqual(loadSecret(path.join(dir, 'missing')), { reason: 'secret file missing' });
		assert.deepStrictEqual(loadSecret(file('open', KEY, 0o644)), { reason: 'secret file must be mode 600' });
		assert.deepStrictEqual(loadSecret(file('short', 'abc', 0o600)), { reason: 'secret shorter than 32 bytes' });
	});

	it('does not listen without the file', async () => {
		const service = new SendService({ secretPath: path.join(dir, 'missing'), port: 0 });
		await service.initialize({ graphApiClient: {}, chatServiceClient: {} });
		assert.strictEqual(service.address(), null);
	});
});

describe('send service stays out of the read-only modules', () => {
	it('app/mcp and app/chatService do not load the send service', () => {
		const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]));
		for (const sub of ['mcp', 'chatService']) {
			for (const f of walk(path.join(__dirname, '..', '..', 'app', sub)).filter((x) => x.endsWith('.js'))) {
				assert.ok(!/sendService/.test(fs.readFileSync(f, 'utf8')), `${f} references sendService`);
			}
		}
	});
});
