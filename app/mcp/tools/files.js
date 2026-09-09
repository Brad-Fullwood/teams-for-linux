const { z } = require('zod');
const zlib = require('node:zlib');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { htmlToText, truncate, formatDate, toolResult } = require('./format');

const MAX_DOWNLOAD_BYTES = 25 * 1024 * 1024;

/** Encode a sharing URL for Graph's /shares/{id} segment. */
function shareId(url) {
  return `u!${Buffer.from(url).toString('base64').replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '')}`;
}

/**
 * Minimal ZIP reader (central directory + deflate) so .docx/.xlsx/.pptx can be read
 * without a dependency. Returns the named entry as a Buffer, or null.
 */
function zipEntry(buffer, name) {
  const eocd = buffer.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (eocd < 0) return null;
  const count = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) return null;
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const entryName = buffer.toString('utf8', offset + 46, offset + 46 + nameLength);
    if (entryName === name) {
      const localNameLength = buffer.readUInt16LE(localOffset + 26);
      const localExtraLength = buffer.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + localNameLength + localExtraLength;
      const data = buffer.subarray(start, start + compressedSize);
      if (method === 0) return Buffer.from(data);
      if (method === 8) return zlib.inflateRawSync(data);
      return null;
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return null;
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
  const sheets = await graph.makeRequest(`${base}?$select=id,name`);
  if (!sheets.success) return sheets;
  const parts = [];
  for (const sheet of sheets.data?.value ?? []) {
    const range = await graph.makeRequest(`${base}/${encodeURIComponent(sheet.id)}/usedRange(valuesOnly=true)?$select=address,text`);
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
  server.registerTool('files_shared_with_me', {
    title: 'Files shared with me',
    description: 'Documents other people shared with you (OneDrive/SharePoint), newest first. Returns links usable with files_get_content.',
    inputSchema: {
      nameContains: z.string().optional(),
      limit: z.number().int().min(1).max(100).default(30),
    },
  }, async ({ nameContains, limit }) => {
    const result = await graph.makeRequest('/me/drive/sharedWithMe?$top=200');
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
    description: 'Content of a OneDrive/SharePoint file from its link: Word (.docx) and Excel (.xlsx) are converted to text, plain text/markdown/csv/json returned as is, anything else (PDF, images) is saved to saveDir and the path returned. Read-only against Microsoft.',
    inputSchema: {
      url: z.string().url().describe('The sharing or web link of the file'),
      maxChars: z.number().int().min(1000).max(400_000).default(60_000),
      maxRowsPerSheet: z.number().int().min(10).max(5000).default(500),
      saveDir: z.string().optional().describe('Directory to save files that are not converted to text (PDF, images); default: the OS temp dir'),
    },
  }, async ({ url, maxChars, maxRowsPerSheet, saveDir }) => {
    const item = await graph.makeRequest(`/shares/${shareId(url)}/driveItem?$select=id,name,size,file,parentReference,webUrl,lastModifiedDateTime`);
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
    const content = await graph.makeRequest(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(meta.id)}/content`, { raw: true });
    if (!content.success) return toolResult(content, () => '');
    const buffer = content.data;

    if (ext === 'docx') {
      const text = docxToText(buffer);
      if (text === null) return { isError: true, content: [{ type: 'text', text: `${header}Could not unpack the document.` }] };
      return { content: [{ type: 'text', text: header + '\n' + truncate(text, maxChars) }] };
    }
    if (['txt', 'md', 'csv', 'json', 'log', 'xml', 'al', 'yml', 'yaml'].includes(ext) || (meta.file?.mimeType ?? '').startsWith('text/')) {
      return { content: [{ type: 'text', text: header + '\n' + truncate(buffer.toString('utf8'), maxChars) }] };
    }
    const dir = saveDir || os.tmpdir();
    fs.mkdirSync(dir, { recursive: true });
    const target = path.join(dir, name.replaceAll(/[\\/:*?"<>|]/g, '_'));
    fs.writeFileSync(target, buffer);
    return { content: [{ type: 'text', text: `${header}Type "${ext}" is not converted to text here; saved to ${target}. PDFs can be read with pdftotext.` }] };
  });
}

module.exports = { registerFileTools, docxToText, xlsxBufferToText, zipEntry, shareId };
