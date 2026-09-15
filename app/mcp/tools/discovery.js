const { z } = require('zod');
const { randomUUID, createHash } = require('node:crypto');

// Kept outside the per-request MCP server. Cursors expire and cannot supply URLs.
const searches = new Map();
const TTL = 30 * 60 * 1000;
const enc = (s) => encodeURIComponent(s).replaceAll("'", '%27');
const result = (data, isError = false) => ({ ...(isError ? { isError: true } : {}), content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] });
const failure = (r) => r.status === 403 ? 'Access denied (HTTP 403)' : r.status === 401 ? 'Authentication required (HTTP 401)' : `Request failed${r.status ? ` (HTTP ${r.status})` : ''}`;

function validateNext(next, base) {
  const url = new URL(next);
  const expected = new URL(base, 'https://graph.microsoft.com/v1.0/');
  // Relative Graph paths are relative to the version, not to the hostname.
  const pathname = base.startsWith('/') ? `/v1.0${base.split('?')[0]}` : expected.pathname;
  if (url.origin !== 'https://graph.microsoft.com' || url.username || url.password || url.hash || url.pathname !== pathname) throw new Error('Invalid pagination URL');
  return url.href;
}

async function read(graph, endpoint) {
  try { return await graph.makeRequest(endpoint, { method: 'GET', sensitive: true }); }
  catch { return { success: false }; }
}

const metadata = (item, location) => {
  const r = item.remoteItem ?? item;
  return { id: r.id, name: r.name, url: r.webUrl, location: r.parentReference?.path ?? location, driveId: r.parentReference?.driveId, modified: r.lastModifiedDateTime, type: r.folder ? 'folder' : 'file' };
};

function registerDiscoveryTools(server, graph) {
  const limit = z.number().int().min(1).max(200).default(100);
  const cursor = z.string().max(16000).optional();
  const annotations = { readOnlyHint: true, destructiveHint: false };
  async function list(base, endpoint, cursorValue) {
    try {
      const expectedQuery = new URL(endpoint, 'https://graph.microsoft.com').searchParams.get('search');
      const validate = (next) => {
        const checked = validateNext(next, base);
        if (expectedQuery !== null && new URL(checked).searchParams.get('search') !== expectedQuery) throw new Error('Cursor belongs to a different site query');
        return checked;
      };
      const r = await read(graph, cursorValue ? validate(cursorValue) : endpoint);
      if (!r.success) return result({ complete: false, error: failure(r) }, true);
      const next = r.data?.['@odata.nextLink'];
      return result({ items: r.data?.value ?? [], complete: !next, cursor: next ? validate(next) : null });
    } catch { return result({ complete: false, error: 'Invalid pagination URL. Restart the listing.' }, true); }
  }
  server.registerTool('sites_search', {
    title: 'Find project SharePoint sites', annotations,
    description: 'Search accessible SharePoint sites by project name. Project investigations must check files as well as chats. Follow project-site links and enumerate all libraries, including Delivery Documents. GET-only; denied access and pagination are explicit.',
    inputSchema: { query: z.string().trim().min(1).max(500), limit, cursor },
  }, async ({ query, limit = 100, cursor }) => list('/sites', `/sites?search=${enc(query)}&$top=${limit}&$select=id,displayName,webUrl,lastModifiedDateTime`, cursor));
  server.registerTool('files_list_libraries', {
    title: 'List every site document library', annotations,
    description: 'Enumerate a SharePoint site document libraries, including non-default libraries such as Delivery Documents. Follow every cursor. Use the site ID from sites_search. GET-only.',
    inputSchema: { siteId: z.string().min(1).max(1000), limit, cursor },
  }, async ({ siteId, limit = 100, cursor }) => {
    const base = `/sites/${enc(siteId)}/drives`;
    return list(base, `${base}?$top=${limit}&$select=id,name,webUrl,driveType,lastModifiedDateTime`, cursor);
  });
  server.registerTool('files_search', {
    title: 'Find accessible project files', annotations,
    description: 'Search accessible files globally, across EVERY library of a supplied site, or under a folder URL. Returns names, URLs, locations, modified dates and an opaque continuation cursor. Use CDDs as authoritative requirements; solution designs are background. Follow cursors and report incomplete/denied scopes. Global Graph search is indexed and is not exhaustive proof of absence; follow project sites and search their libraries. GET-only.',
    inputSchema: { query: z.string().trim().min(1).max(500), siteId: z.string().min(1).max(1000).optional(), folderUrl: z.string().url().optional(), limit, cursor },
  }, async ({ query, siteId, folderUrl, limit = 100, cursor }) => {
    if (siteId && folderUrl) return result({ complete: false, error: 'Supply siteId or folderUrl, not both.' }, true);
    const scope = createHash('sha256').update(JSON.stringify([query, siteId, folderUrl, limit])).digest('hex');
    for (const [key, value] of searches) if (Date.now() - value.created > TTL) searches.delete(key);
    let state = cursor ? searches.get(cursor) : null;
    if (cursor && (!state || state.scope !== scope)) return result({ complete: false, error: 'Invalid or expired cursor. Restart with the same search scope.' }, true);
    if (state?.busy) return result({ complete: false, error: 'Cursor is already in use.' }, true);
    const q = enc(query.replaceAll("'", "''"));
    const searchPath = (base) => `${base}/search(q='${q}')`;
    if (!state) {
      state = { created: Date.now(), scope, tasks: [], seen: new Set(), visited: new Set(), errors: [], libraries: new Set() };
      if (siteId) {
        const base = `/sites/${enc(siteId)}/drives`;
        state.tasks.push({ kind: 'libraries', base, endpoint: `${base}?$top=200`, location: siteId });
      } else if (folderUrl) {
        const share = `u!${Buffer.from(folderUrl).toString('base64url')}`;
        const r = await read(graph, `/shares/${share}/driveItem?$select=id,folder,parentReference`);
        if (!r.success) return result({ complete: false, error: failure(r) }, true);
        const f = r.data;
        if (!f?.folder || !f.id || !f.parentReference?.driveId) return result({ complete: false, error: 'The link does not resolve to a folder.' }, true);
        const base = searchPath(`/drives/${enc(f.parentReference.driveId)}/items/${enc(f.id)}`);
        state.tasks.push({ base, endpoint: `${base}?$top=${limit}`, location: folderUrl });
      } else {
        const base = searchPath('/me/drive');
        state.tasks.push({ base, endpoint: `${base}?$top=${limit}`, location: 'Accessible indexed files' });
      }
    }
    state.busy = true;
    const items = [];
    // Bound each call, preserving all unfinished libraries/pages in the cursor.
    for (let count = 0; state.tasks.length && count < 30 && items.length < limit; count++) {
      const task = state.tasks.shift();
      if (state.visited.has(task.endpoint)) { state.errors.push({ location: task.location, error: 'Pagination loop' }); continue; }
      state.visited.add(task.endpoint);
      const r = await read(graph, task.endpoint);
      if (!r.success) { state.errors.push({ location: task.location, error: failure(r) }); continue; }
      if (!Array.isArray(r.data?.value)) { state.errors.push({ location: task.location, error: 'Invalid response' }); continue; }
      for (const entry of r.data.value) {
        if (task.kind === 'libraries') {
          if (!entry.id || state.libraries.has(entry.id)) continue;
          state.libraries.add(entry.id);
          const base = searchPath(`/drives/${enc(entry.id)}/root`);
          state.tasks.push({ base, endpoint: `${base}?$top=${limit}`, location: entry.webUrl ?? entry.name, driveId: entry.id });
        } else {
          const item = metadata(entry, task.location);
          const key = `${item.driveId ?? task.driveId ?? ''}:${item.id ?? item.url}`;
          if (!state.seen.has(key)) { state.seen.add(key); items.push(item); }
        }
      }
      if (r.data['@odata.nextLink']) {
        try { state.tasks.push({ ...task, endpoint: validateNext(r.data['@odata.nextLink'], task.base) }); }
        catch { state.errors.push({ location: task.location, error: 'Invalid pagination URL' }); }
      }
    }
    if (cursor) searches.delete(cursor);
    state.busy = false;
    const next = state.tasks.length ? randomUUID() : null;
    if (next) {
      if (searches.size >= 100) searches.delete(searches.keys().next().value);
      searches.set(next, state);
    }
    return result({ items, complete: !next && !state.errors.length, cursor: next, searchedLibraries: state.libraries.size, errors: state.errors, scope: siteId ? 'all site libraries' : folderUrl ? 'folder hierarchy' : 'accessible indexed files; follow project sites to verify coverage' });
  });
}
module.exports = { registerDiscoveryTools, validateNext };
