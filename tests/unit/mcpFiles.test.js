'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { registerFileTools } = require('../../app/mcp/tools/files');

function harness(responses) {
  const handlers = {};
  const calls = [];
  registerFileTools({ registerTool: (name, schema, handler) => { handlers[name] = handler; } }, {
    makeRequest: async (...args) => { assert.equal(args[1]?.sensitive, true, "File requests must redact private URLs and responses"); calls.push(args); return responses.shift(); },
  });
  return { handlers, calls };
}
const ok = (data) => ({ success: true, data });
const folder = { id: 'folder', folder: {}, parentReference: { driveId: 'drive' } };
const url = 'https://example.sharepoint.com/folder';

test('folder listing returns children and resumes the Graph cursor', async () => {
  const cursor = 'https://graph.microsoft.com/v1.0/drives/drive/items/folder/children?$skiptoken=abc';
  const h = harness([ok(folder), ok({ value: [{ name: 'Invoice.docx', file: {}, size: 42, webUrl: url }], '@odata.nextLink': cursor }), ok(folder), ok({ value: [] })]);
  const result = await h.handlers.files_list_folder({ url, limit: 1 });
  assert.match(result.content[0].text, /Invoice.docx/);
  assert.ok(result.content[0].text.includes(cursor));
  await h.handlers.files_list_folder({ url, limit: 1, cursor });
  assert.equal(h.calls[3][0], cursor);
});

test('folder cursor cannot redirect authenticated requests to another host or folder', async () => {
  for (const cursor of ['https://evil.example/v1.0/drives/drive/items/folder/children', 'https://graph.microsoft.com/v1.0/me', 'invalid']) {
    const h = harness([ok(folder)]);
    assert.equal((await h.handlers.files_list_folder({ url, limit: 10, cursor })).isError, true);
    assert.equal(h.calls.length, 1);
  }
});

test('original bytes survive download, unsafe names are contained, and collisions do not overwrite', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-files-test-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const bytes = Buffer.from('PK\u0000original custom XML and content controls\u00ff', 'utf8');
  const meta = { id: 'file', name: '../Invoice.docx', file: {}, size: bytes.length, parentReference: { driveId: 'drive' } };
  const h = harness([ok(meta), ok(bytes), ok(meta), ok(Buffer.alloc(bytes.length))]);
  assert.equal((await h.handlers.files_download({ url, saveDir: dir })).isError, undefined);
  const names = fs.readdirSync(dir);
  assert.equal(names.length, 1);
  assert.deepEqual(fs.readFileSync(path.join(dir, names[0])), bytes);
  assert.equal((await h.handlers.files_download({ url, saveDir: dir })).isError, true);
  assert.deepEqual(fs.readFileSync(path.join(dir, names[0])), bytes);
  assert.deepEqual(h.calls[1][1], { raw: true, sensitive: true });
});

test('download rejects folders, oversize metadata, oversized bodies and incomplete bodies', async () => {
  const meta = { id: 'file', name: 'file.docx', file: {}, size: 3, parentReference: { driveId: 'drive' } };
  for (const responses of [[ok(folder)], [ok({ ...meta, size: 26 * 1024 * 1024 })], [ok(meta), ok(Buffer.alloc(26 * 1024 * 1024))], [ok(meta), ok(Buffer.from('x'))]]) {
    const h = harness(responses);
    assert.equal((await h.handlers.files_download({ url })).isError, true);
  }
});

test('authentication failures remain actionable tool errors', async () => {
  for (const name of ['files_download', 'files_list_folder']) {
    const h = harness([{ success: false, status: 401, error: 'Additional claims required' }]);
    const result = await h.handlers[name]({ url, limit: 10 });
    assert.equal(result.isError, true);
    assert.match(result.content[0].text, /HTTP 401.*Additional claims required/);
  }
});
