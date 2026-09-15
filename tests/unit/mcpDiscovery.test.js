const { test } = require('node:test');
const assert = require('node:assert/strict');
const { registerDiscoveryTools } = require('../../app/mcp/tools/discovery');
function harness(responses) {
  const handlers = {}, calls = [];
  registerDiscoveryTools({ registerTool: (n, s, h) => { handlers[n] = h; } }, { makeRequest: async (...args) => { calls.push(args); return responses.shift(); } });
  return { handlers, calls };
}
const ok = (value, next) => ({ success: true, data: { value, ...(next ? { '@odata.nextLink': next } : {}) } });
const data = (r) => JSON.parse(r.content[0].text);
const file = (id) => ({ id, name: `${id}.docx`, webUrl: `https://tenant.sharepoint.com/${id}`, parentReference: { driveId: 'delivery' }, lastModifiedDateTime: '2026-09-12' });
test('site search visits paged non-default libraries and deduplicates remote results', async () => {
  const h = harness([ok([{ id: 'default' }], 'https://graph.microsoft.com/v1.0/sites/site/drives?$skiptoken=2'), ok([]), ok([{ id: 'delivery', name: 'Delivery Documents' }]), ok([file('CDD004'), { remoteItem: file('CDD004') }, file('CDD005')])]);
  const r = data(await h.handlers.files_search({ query: 'CDD', siteId: 'site' }));
  assert.equal(r.complete, true); assert.equal(r.searchedLibraries, 2); assert.equal(r.items.length, 2);
  assert.match(h.calls[3][0], /drives\/delivery\/root\/search/);
  for (const [, opts] of h.calls) assert.deepEqual(opts, { method: 'GET', sensitive: true });
});
test('search cursors resume pages, retain deduplication, and reject changed scope/replay', async () => {
  const h = harness([ok([file('a')], "https://graph.microsoft.com/v1.0/me/drive/search(q='CDD')?$skiptoken=2"), ok([file('a'), file('b')])]);
  const args = { query: 'CDD', limit: 1 };
  const first = data(await h.handlers.files_search(args)); assert.equal(first.complete, false);
  assert.equal((await h.handlers.files_search({ ...args, query: 'other', cursor: first.cursor })).isError, true);
  const second = data(await h.handlers.files_search({ ...args, cursor: first.cursor }));
  assert.equal(second.complete, true); assert.deepEqual(second.items.map(x => x.id), ['b']);
  assert.equal((await h.handlers.files_search({ ...args, cursor: first.cursor })).isError, true);
});
test('permission errors preserve successful results and declare incompleteness', async () => {
  const h = harness([ok([{ id: 'a' }, { id: 'b' }]), { success: false, status: 403, error: 'private query' }, ok([file('CDD')])]);
  const r = data(await h.handlers.files_search({ query: 'CDD', siteId: 's' }));
  assert.equal(r.complete, false); assert.equal(r.items.length, 1); assert.match(r.errors[0].error, /Access denied/); assert.doesNotMatch(JSON.stringify(r), /private query/);
});
test('host, path, userinfo, fragment and loop pagination are rejected without following them', async () => {
  for (const next of ['https://evil.example/v1.0/me/drive/search', 'https://graph.microsoft.com/v1.0/me/messages', "https://u@graph.microsoft.com/v1.0/me/drive/search(q='CDD')", "https://graph.microsoft.com/v1.0/me/drive/search(q='CDD')#x"]) {
    const h = harness([ok([], next)]);
    const r = data(await h.handlers.files_search({ query: 'CDD' }));
    assert.equal(r.complete, false); assert.match(r.errors[0].error, /Invalid pagination/); assert.equal(h.calls.length, 1);
  }
  const next = "https://graph.microsoft.com/v1.0/me/drive/search(q='CDD')?$skiptoken=1";
  const h = harness([ok([], next), ok([], next)]);
  assert.match(data(await h.handlers.files_search({ query: 'CDD' })).errors[0].error, /loop/);
});
test('global, folder and escaped query scopes use supported GET endpoints', async () => {
  const h = harness([{ success: true, data: { id: 'folder', folder: {}, parentReference: { driveId: 'd' } } }, ok([])]);
  assert.equal(data(await h.handlers.files_search({ query: "O'Brien & CDD", folderUrl: 'https://tenant.sharepoint.com/folder' })).complete, true);
  assert.match(h.calls[1][0], /\/drives\/d\/items\/folder\/search\(q='O%27%27Brien%20%26%20CDD'\)/);
  assert.equal((await h.handlers.files_search({ query: 'a', siteId: 's', folderUrl: 'https://a.com' })).isError, true);
});
test('site and library tools expose pagination and auth failures', async () => {
  const h = harness([ok([{ id: 's' }], 'https://graph.microsoft.com/v1.0/sites?search=Tahira&$skiptoken=1'), { success: false, status: 401 }, ok([{ id: 'd', name: 'Delivery Documents' }])]);
  const first = data(await h.handlers.sites_search({ query: 'Tahira' })); assert.equal(first.complete, false);
  assert.equal((await h.handlers.sites_search({ query: 'Tahira', cursor: first.cursor })).isError, true);
  assert.equal(data(await h.handlers.files_list_libraries({ siteId: 's' })).items[0].name, 'Delivery Documents');
  assert.equal((await h.handlers.files_list_libraries({ siteId: 's', cursor: 'https://graph.microsoft.com/v1.0/sites/other/drives' })).isError, true);
});

test('site cursors cannot silently change the requested project search', async () => {
  const h = harness([]);
  const r = await h.handlers.sites_search({ query: 'Tahira', cursor: 'https://graph.microsoft.com/v1.0/sites?search=Other&$skiptoken=2' });
  assert.equal(r.isError, true);
  assert.equal(h.calls.length, 0);
});
