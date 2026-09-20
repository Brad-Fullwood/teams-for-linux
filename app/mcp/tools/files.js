const { z } = require('zod');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { htmlToText, truncate, formatDate, toolResult } = require('./format');
const { zipEntry, pdfToolResult, pptxToText } = require('./documents');

const { registerDiscoveryTools } = require('./discovery');

const MAX_DOWNLOAD_BYTES = 25 * 1024 * 1024;

function fileError(text) {
  return { isError: true, content: [{ type: 'text', text }] };
}

/** Encode a sharing URL for Graph's /shares/{id} segment. */
function shareId(url) {
  return `u!${Buffer.from(url).toString('base64').replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '')}`;
}

/** Word document body as plain text: paragraphs become lines, tabs kept, tables become tab-separated rows. */
function docxToText(buffer) {
  const xml = zipEntry(buffer, 'word/document.xml');
  if (!xml) return null;
  // Paragraphs end rows; inside a table each cell's paragraphs are joined with spaces and cells with tabs.
  const text = xml.toString('utf8')
    .replaceAll(/<w:tab\/>/g, '\t')
    .replaceAll(/<w:tc>[\s\S]*?<\/w:tc>/g, (cell) => `${cell.replaceAll(/<\/w:p>/g, ' ').replaceAll(/<[^>]+>/g, '').trim()}\t`)
    .replaceAll(/<\/w:tr>/g, '<br>')
    .replaceAll(/<\/w:p>/g, '<br>');
  return htmlToText(text).split('\n').map((l) => l.replace(/\t+$/, '')).join('\n');
}

/** Excel workbook from raw bytes: shared strings + each sheet's cells as tab-separated rows. */
function xlsxBufferToText(buffer, maxRowsPerSheet = 500) {
  const workbook = zipEntry(buffer, 'xl/workbook.xml');
  if (!workbook) return null;
  const shared = [];
  const sst = zipEntry(buffer, 'xl/sharedStrings.xml');
  if (sst) {
    for (const m of sst.toString('utf8').matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      shared.push(htmlToText(m[1].replaceAll(/<[^>]+>/g, '')));
    }
  }
  const sheets = [...workbook.toString('utf8').matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)];
  const rels = zipEntry(buffer, 'xl/_rels/workbook.xml.rels')?.toString('utf8') ?? '';
  const parts = [];
  for (const [, name, rid] of sheets) {
    const target = rels.match(new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`))?.[1] ?? rels.match(new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`))?.[1];
    if (!target) continue;
    const xml = zipEntry(buffer, `xl/${target.replace(/^\//, '').replace(/^xl\//, '')}`)?.toString('utf8');
    if (!xml) continue;
    const rows = [];
    for (const row of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells = [];
      for (const cell of row[1].matchAll(/<c ([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attrs = cell[1]; const inner = cell[2];
        const v = inner.match(/<v>([^<]*)<\/v>/)?.[1] ?? '';
        const t = inner.match(/<t[^>]*>([^<]*)<\/t>/)?.[1] ?? '';
        cells.push(/t="s"/.test(attrs) ? (shared[Number(v)] ?? '') : (/t="inlineStr"/.test(attrs) ? htmlToText(t) : v));
      }
      if (cells.some((c) => String(c).trim() !== '')) rows.push(cells.map((c) => String(c).replaceAll(/\s+/g, ' ')).join('\t'));
    }
    parts.push(`## ${name} (${rows.length} rows${rows.length > maxRowsPerSheet ? `, first ${maxRowsPerSheet} shown` : ''})\n${rows.slice(0, maxRowsPerSheet).join('\n')}`);
  }
  return parts.join('\n\n');
}

/** Excel workbook via Graph's workbook API: every worksheet's used range as tab-separated rows. */
async function xlsxToText(graph, driveId, itemId, maxRowsPerSheet) {
  const base = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/worksheets`;
  const sheets = await graph.makeRequest(`${base}?$select=id,name`, { sensitive: true });
  if (!sheets.success) return sheets;
  const parts = [];
  for (const sheet of sheets.data?.value ?? []) {
    const range = await graph.makeRequest(`${base}/${encodeURIComponent(sheet.id)}/usedRange(valuesOnly=true)?$select=address,text`, { sensitive: true });
    if (!range.success) {
      parts.push(`## ${sheet.name}\n(could not read: ${range.error})`);
      continue;
    }
    const rows = (range.data?.text ?? []).filter((r) => r.some((c) => String(c).trim() !== ''));
    const shown = rows.slice(0, maxRowsPerSheet).map((r) => r.map((c) => String(c).replaceAll(/\s+/g, ' ')).join('\t'));
    parts.push(`## ${sheet.name} (${rows.length} rows${rows.length > maxRowsPerSheet ? `, first ${maxRowsPerSheet} shown` : ''})\n${shown.join('\n')}`);
  }
  return { success: true, data: parts.join('\n\n') };
}

/**
 * Register read-only file tools (OneDrive/SharePoint via Graph).
 * @param {import('@modelcontextprotocol/sdk/server/mcp.js').McpServer} server
 * @param {object} graph - GraphApiClient
 */
function registerFileTools(server, graph) {
  registerDiscoveryTools(server, graph);
  server.registerTool('files_list_folder', {
    title: 'List a shared folder',
    description: 'List immediate children of a OneDrive/SharePoint folder using the signed-in Teams session. Returns original file links and a cursor for the next page. Read-only.',
    inputSchema: {
      url: z.string().url().describe('Sharing or web link of the folder'),
      limit: z.number().int().min(1).max(200).default(100),
      cursor: z.string().optional().describe('Next-page cursor from this tool for the same folder'),
    },
  }, async ({ url, limit, cursor }) => {
    const item = await graph.makeRequest(`/shares/${shareId(url)}/driveItem?$select=id,name,folder,parentReference`, { sensitive: true });
    if (!item.success) return toolResult(item, () => '');
    const meta = item.data ?? {};
    if (!meta.folder || !meta.id || !meta.parentReference?.driveId) return fileError('The link does not resolve to a folder.');
    const base = `/drives/${encodeURIComponent(meta.parentReference.driveId)}/items/${encodeURIComponent(meta.id)}/children`;
    let endpoint = `${base}?$top=${limit}&$select=id,name,size,file,folder,webUrl,lastModifiedDateTime`;
    if (cursor) {
      try {
        const next = new URL(cursor);
        if (next.origin !== 'https://graph.microsoft.com' || next.pathname !== `/v1.0${base}` || next.username || next.password || next.hash) return fileError('Invalid folder cursor. Use the cursor returned for this folder.');
        endpoint = next.href;
      } catch {
        return fileError('Invalid folder cursor.');
      }
    }
    const result = await graph.makeRequest(endpoint, { sensitive: true });
    return toolResult(result, (data) => {
      const lines = (data?.value ?? []).map((entry) => `- ${entry.name}${entry.folder ? '/' : ''} [${entry.folder ? 'folder' : entry.file?.mimeType ?? 'file'}, ${entry.size ?? 0} bytes]\n  link: ${entry.webUrl ?? ''}\n  modified: ${formatDate(entry.lastModifiedDateTime)}`);
      if (data?.['@odata.nextLink']) lines.push(`next: cursor=${data['@odata.nextLink']}`);
      return lines.join('\n');
    });
  });

  server.registerTool('files_download', {
    title: 'Download an original shared file',
    description: 'Save original OneDrive/SharePoint file bytes using the signed-in Teams session, including DOCX XML bindings. No conversion. Maximum 25 MB; existing local files are never overwritten. Read-only against Microsoft.',
    inputSchema: {
      url: z.string().url().describe('Sharing or web link of the file'),
      saveDir: z.string().optional().describe('Local destination directory; defaults to a new temporary directory'),
    },
  }, async ({ url, saveDir }) => {
    const item = await graph.makeRequest(`/shares/${shareId(url)}/driveItem?$select=id,name,size,file,folder,parentReference`, { sensitive: true });
    if (!item.success) return toolResult(item, () => '');
    const meta = item.data ?? {};
    if (meta.folder || !meta.file || !meta.id || !meta.parentReference?.driveId) return fileError('The link does not resolve to a downloadable file.');
    if (meta.size > MAX_DOWNLOAD_BYTES) return fileError('File too large to download (limit 25 MB).');
    const content = await graph.makeRequest(`/drives/${encodeURIComponent(meta.parentReference.driveId)}/items/${encodeURIComponent(meta.id)}/content`, { raw: true, sensitive: true });
    if (!content.success) return toolResult(content, () => '');
    if (!Buffer.isBuffer(content.data)) return fileError('Download did not return file bytes.');
    if (content.data.length > MAX_DOWNLOAD_BYTES) return fileError('File too large to save (limit 25 MB).');
    if (typeof meta.size === 'number' && content.data.length !== meta.size) return fileError('Download size changed or is incomplete; retry the download.');
    const name = String(meta.name ?? 'download').replaceAll(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/[. ]+$/, '') || 'download';
    try {
      const dir = saveDir ? path.resolve(saveDir) : fs.mkdtempSync(path.join(os.tmpdir(), 'o365-download-'));
      fs.mkdirSync(dir, { recursive: true });
      const target = path.join(dir, name);
      fs.writeFileSync(target, content.data, { flag: 'wx', mode: 0o600 });
      return { content: [{ type: 'text', text: `Saved original file to ${target} (${content.data.length} bytes).` }] };
    } catch (error) {
      return fileError(error.code === 'EEXIST' ? 'Destination already exists; choose a different saveDir. Nothing was overwritten.' : `Could not save file (${error.code ?? 'local filesystem error'}).`);
    }
  });

  server.registerTool('files_shared_with_me', {
    title: 'Files shared with me',
    description: 'Documents other people shared with you (OneDrive/SharePoint), newest first. Returns links usable with files_get_content.',
    inputSchema: {
      nameContains: z.string().optional(),
      limit: z.number().int().min(1).max(100).default(30),
    },
  }, async ({ nameContains, limit }) => {
    const result = await graph.makeRequest('/me/drive/sharedWithMe?$top=200', { sensitive: true });
    return toolResult(result, (data) => (data?.value ?? [])
      .filter((i) => !nameContains || String(i.name ?? i.remoteItem?.name ?? '').toLowerCase().includes(nameContains.toLowerCase()))
      .sort((a, b) => new Date(b.remoteItem?.shared?.sharedDateTime ?? b.lastModifiedDateTime ?? 0) - new Date(a.remoteItem?.shared?.sharedDateTime ?? a.lastModifiedDateTime ?? 0))
      .slice(0, limit)
      .map((i) => {
        const r = i.remoteItem ?? i;
        const by = r.shared?.sharedBy?.user?.displayName || r.createdBy?.user?.displayName || '?';
        return `- ${r.name}${r.folder ? '/' : ''} — shared by ${by} ${formatDate(r.shared?.sharedDateTime)}, modified ${formatDate(r.lastModifiedDateTime)}\n  link: ${r.webUrl}`;
      })
      .join('\n'));
  });

  server.registerTool('files_get_content', {
    title: 'Read a shared file',
    description: 'Content of a OneDrive/SharePoint file from its link: Word (.docx), Excel (.xlsx), PDF and PowerPoint (.pptx) are converted to text, plain text/markdown/csv/json returned as is, anything else (images) is saved to saveDir and the path returned. Read-only against Microsoft.',
    inputSchema: {
      url: z.string().url().describe('The sharing or web link of the file'),
      maxChars: z.number().int().min(1000).max(400_000).default(60_000),
      maxRowsPerSheet: z.number().int().min(10).max(5000).default(500),
      maxPages: z.number().int().min(1).max(2000).default(200).describe('PDF only: maximum pages to extract text from'),
      saveDir: z.string().optional().describe('Directory to save files that are not converted to text (images); default: the OS temp dir'),
    },
  }, async ({ url, maxChars, maxRowsPerSheet, maxPages, saveDir }) => {
    const item = await graph.makeRequest(`/shares/${shareId(url)}/driveItem?$select=id,name,size,file,parentReference,webUrl,lastModifiedDateTime`, { sensitive: true });
    if (!item.success) return toolResult(item, () => '');
    const meta = item.data ?? {};
    const name = String(meta.name ?? '');
    const ext = name.toLowerCase().split('.').pop();
    const driveId = meta.parentReference?.driveId;
    const header = `${name} (${Math.round((meta.size ?? 0) / 1024)} KB, modified ${formatDate(meta.lastModifiedDateTime)})\n`;

    if (ext === 'xlsx' || ext === 'xlsm') {
      const text = await xlsxToText(graph, driveId, meta.id, maxRowsPerSheet);
      return toolResult(text, (t) => header + '\n' + truncate(t, maxChars));
    }

    if ((meta.size ?? 0) > MAX_DOWNLOAD_BYTES) {
      return { isError: true, content: [{ type: 'text', text: `${header}File too large to read here (limit ${MAX_DOWNLOAD_BYTES / 1024 / 1024} MB).` }] };
    }
    const content = await graph.makeRequest(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(meta.id)}/content`, { raw: true, sensitive: true });
    if (!content.success) return toolResult(content, () => '');
    const buffer = content.data;

    if (ext === 'docx') {
      const text = docxToText(buffer);
      if (text === null) return { isError: true, content: [{ type: 'text', text: `${header}Could not unpack the document.` }] };
      return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (ext === 'pdf' || (meta.file?.mimeType ?? '') === 'application/pdf') {
      return pdfToolResult(buffer, { header, name, maxPages, maxChars, saveDir });
    }
    if (ext === 'pptx') {
      const text = pptxToText(buffer);
      if (text !== null) return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (['txt', 'md', 'csv', 'json', 'log', 'xml', 'al', 'yml', 'yaml'].includes(ext) || (meta.file?.mimeType ?? '').startsWith('text/')) {
      return { content: [{ type: 'text', text: header + '\n' + truncate(buffer.toString('utf8'), maxChars) }] };
    }
    const dir = saveDir || os.tmpdir();
    fs.mkdirSync(dir, { recursive: true });
    const target = path.join(dir, name.replaceAll(/[\\/:*?"<>|]/g, '_'));
    fs.writeFileSync(target, buffer);
    return { content: [{ type: 'text', text: `${header}Type "${ext}" is not converted to text here; saved to ${target}.` }] };
  });
}

// zipEntry now lives in ./documents; re-exported here so existing imports keep working.
module.exports = { registerFileTools, docxToText, xlsxBufferToText, zipEntry, shareId, fileError, MAX_DOWNLOAD_BYTES };
