const { test } = require('node:test');
const assert = require('node:assert/strict');
const ChatServiceClient = require('../../app/chatService');
const { renderChatMessage } = require('../../app/mcp/tools/teams');

test('inline images remain discoverable in a text message', () => {
  const rendered = renderChatMessage({ messagetype: 'RichText/Html', content: '<p>Reference</p><img src="https://eu.asm.skype.com/v1/objects/1/views/imgpsh_fullsize?a=1&amp;b=2">' });
  assert.match(rendered, /image: https:\/\/eu.asm.skype.com\/v1\/objects\/1\/views\/imgpsh_fullsize\?a=1&b=2/);
});

test('media client limits credential use to Teams image hosts and returns original bytes', async () => {
  const client = new ChatServiceClient({ graphApi: { enabled: true } });
  client.initialize({ webContents: { executeJavaScript: async () => ({ success: true, skypeToken: 'test-token', expiration: Date.now() + 3_600_000, chatService: 'https://uk.ng.msg.teams.microsoft.com' }) } });
  const original = globalThis.fetch;
  const calls = [];
  globalThis.fetch = async (url, init) => { calls.push({ url, init }); return new Response(Buffer.from([137, 80, 78, 71]), { headers: { 'content-type': 'image/png' } }); };
  try {
    for (const url of ['https://evil.example/v1/objects/x', 'https://asm.skype.com.evil.example/v1/objects/x', 'https://user@asm.skype.com/v1/objects/x', 'http://asm.skype.com/v1/objects/x', 'https://asm.skype.com/private']) assert.equal((await client.getImage(url)).success, false);
    assert.equal(calls.length, 0);
    const result = await client.getImage('https://eu.asm.skype.com/v1/objects/x/views/imgpsh_fullsize');
    assert.equal(result.success, true);
    assert.deepEqual(result.data, Buffer.from([137, 80, 78, 71]));
    assert.equal(calls[0].init.redirect, 'error');
    assert.equal(calls[0].init.method, 'GET');
  } finally { globalThis.fetch = original; }
});
