'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const { registerMailTools } = require('../../app/mcp/tools/mail');

function harness(overrides = {}) {
  const calls = [];
  const graph = {
    getMailMessages: async (options) => { calls.push(['getMailMessages', options]); return overrides.searchResult ?? { success: true, data: { value: [] } }; },
    makeRequest: async (endpoint) => { calls.push(['makeRequest', endpoint]); return overrides.cursorResult ?? { success: true, data: { value: [] } }; },
  };
  const handlers = {};
  registerMailTools({ registerTool: (name, _schema, handler) => { handlers[name] = handler; } }, graph);
  return { handlers, calls };
}

const msg = (id, receivedDateTime, subject = 'Subject') => ({
  id, subject, receivedDateTime, isRead: true,
  from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
});

describe('F. mail_search cursor and paging', () => {
  it('searches without a cursor and shows no "more:" line when there is no next page', async () => {
    const { handlers, calls } = harness({ searchResult: { success: true, data: { value: [msg('1', '2026-09-08T09:00:00Z')] } } });
    const result = await handlers.mail_search({ query: 'invoice', limit: 20 });
    assert.equal(calls[0][0], 'getMailMessages');
    assert.match(calls[0][1].search, /"invoice"/);
    assert.doesNotMatch(result.content[0].text, /more:/);
    assert.match(result.content[0].text, /Subject/);
  });

  it('prints "more: cursor=<value>" from @odata.nextLink when present', async () => {
    const next = 'https://graph.microsoft.com/v1.0/me/messages?$search=%22x%22&$skiptoken=abc';
    const { handlers } = harness({ searchResult: { success: true, data: { value: [msg('1', '2026-09-08T09:00:00Z')], '@odata.nextLink': next } } });
    const result = await handlers.mail_search({ query: 'x', limit: 20 });
    assert.match(result.content[0].text, new RegExp(`more: cursor=${next.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`));
  });

  it('follows a valid cursor via makeRequest instead of rebuilding the search', async () => {
    const cursor = 'https://graph.microsoft.com/v1.0/me/messages?$search=%22x%22&$skiptoken=abc';
    const { handlers, calls } = harness({ cursorResult: { success: true, data: { value: [msg('2', '2026-09-08T09:00:00Z')] } } });
    const result = await handlers.mail_search({ query: 'x', limit: 20, cursor });
    assert.equal(calls[0][0], 'makeRequest');
    assert.equal(new URL(calls[0][1]).pathname, '/v1.0/me/messages');
    assert.equal(new URL(calls[0][1]).searchParams.get('$skiptoken'), 'abc');
    assert.match(result.content[0].text, /Subject/);
  });

  it('rejects a cursor pointing at a different host or path', async () => {
    const { handlers: h1, calls: c1 } = harness();
    const badHost = await h1.mail_search({ query: 'x', cursor: 'https://evil.example/v1.0/me/messages?$search=x' });
    assert.equal(badHost.isError, true);
    assert.equal(c1.length, 0);

    const { handlers: h2, calls: c2 } = harness();
    const badPath = await h2.mail_search({ query: 'x', cursor: 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages' });
    assert.equal(badPath.isError, true);
    assert.equal(c2.length, 0);
  });

  it('applies receivedAfter/receivedBefore client-side to the returned page', async () => {
    const items = [msg('1', '2026-09-01T00:00:00Z', 'Old'), msg('2', '2026-09-10T00:00:00Z', 'Mid'), msg('3', '2026-09-20T00:00:00Z', 'New')];
    const { handlers } = harness({ searchResult: { success: true, data: { value: items } } });
    const result = await handlers.mail_search({ query: 'x', receivedAfter: '2026-09-05T00:00:00Z', receivedBefore: '2026-09-15T00:00:00Z' });
    assert.match(result.content[0].text, /Mid/);
    assert.doesNotMatch(result.content[0].text, /Old/);
    assert.doesNotMatch(result.content[0].text, /New/);
  });

  it('keeps an item whose receivedDateTime cannot be parsed rather than silently dropping it', async () => {
    const items = [{ id: '1', subject: 'Weird', receivedDateTime: 'not-a-date', isRead: true, from: { emailAddress: { name: 'A', address: 'a@b.c' } } }];
    const { handlers } = harness({ searchResult: { success: true, data: { value: items } } });
    const result = await handlers.mail_search({ query: 'x', receivedAfter: '2026-09-05T00:00:00Z' });
    assert.match(result.content[0].text, /Weird/);
  });

  it('reports a Graph failure as an isError result', async () => {
    const { handlers } = harness({ searchResult: { success: false, status: 403, error: 'Forbidden' } });
    const result = await handlers.mail_search({ query: 'x' });
    assert.equal(result.isError, true);
  });
});
