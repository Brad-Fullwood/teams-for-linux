const zlib = require('node:zlib');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { htmlToText, truncate } = require('./format');

/**
 * Tool response for PDF bytes shared by files_get_content and mail_get_attachment.
 * Text is returned directly; a local copy is saved only when no text could be read
 * (scanned/image-only, encrypted or corrupt PDFs), so readable documents never land on disk.
 */
async function pdfToolResult(buffer, { header, name, maxPages, maxChars, saveDir }) {
  const result = await pdfToText(buffer, { maxPages });
  if (!result.error && result.hasText) return { content: [{ type: 'text', text: header + '\n' + truncate(result.text, maxChars) }] };
  const dir = saveDir || os.tmpdir();
  fs.mkdirSync(dir, { recursive: true });
  const target = path.join(dir, String(name).replaceAll(/[\\/:*?"<>|]/g, '_'));
  fs.writeFileSync(target, buffer);
  if (result.error) return { isError: true, content: [{ type: 'text', text: `${header}${result.error} Saved to ${target}.` }] };
  return { content: [{ type: 'text', text: `${header}No text layer found (probably a scanned or image-only PDF); saved to ${target}.` }] };
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

// unpdf is ESM-only; load it once via a cached dynamic import so this CommonJS
// module can still call it.
let unpdfPromise;
function loadUnpdf() {
  if (!unpdfPromise) unpdfPromise = import('unpdf');
  return unpdfPromise;
}

/** Collapse runs of spaces/tabs per line and drop blank padding, keeping line breaks. */
function normalisePageText(raw) {
  return String(raw ?? '')
    .split('\n')
    .map((line) => line.replaceAll(/[ \t]+/g, ' ').trim())
    .join('\n')
    .replaceAll(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * PDF bytes to text via unpdf (serverless pdf.js), one page at a time.
 * Never throws: encrypted/corrupt files come back as { hasText: false, error }.
 * @returns {Promise<{ text: string, pages: number, pagesRead: number, hasText: boolean, error?: string }>}
 */
async function pdfToText(buffer, { maxPages = 200 } = {}) {
  const empty = { text: '', pages: 0, pagesRead: 0, hasText: false };
  let unpdf;
  try {
    unpdf = await loadUnpdf();
  } catch (error) {
    return { ...empty, error: `Could not load the PDF engine (${error.message || 'unknown error'}).` };
  }
  const { getDocumentProxy, extractText, getMeta } = unpdf;
  let pdf;
  let pages;
  let totalPages;
  try {
    pdf = await getDocumentProxy(new Uint8Array(buffer));
    ({ totalPages, text: pages } = await extractText(pdf, { mergePages: false }));
  } catch (error) {
    return { ...empty, error: `Could not read the PDF (${error.message || 'corrupt or encrypted file'}).` };
  }

  let metaLine = '';
  try {
    const meta = await getMeta(pdf);
    const title = meta?.info?.Title?.trim();
    const author = meta?.info?.Author?.trim();
    if (title || author) metaLine = [title && `Title: ${title}`, author && `Author: ${author}`].filter(Boolean).join('\n') + '\n\n';
  } catch {
    // Metadata is a nice-to-have; ignore failures.
  }

  const pagesRead = Math.min(pages.length, maxPages);
  let hasText = false;
  const parts = [];
  for (let i = 0; i < pagesRead; i++) {
    const normalised = normalisePageText(pages[i]);
    if (normalised) hasText = true;
    parts.push(`--- page ${i + 1} ---\n${normalised}`);
  }
  const truncatedNote = totalPages > pagesRead ? `\n\n(truncated: showing ${pagesRead} of ${totalPages} pages)` : '';
  return { text: metaLine + parts.join('\n\n') + truncatedNote, pages: totalPages, pagesRead, hasText };
}

/** Text inside a slide/notes-slide XML part: each </a:p> ends a line, <a:t> runs are joined and entity-decoded. */
function paragraphsToText(xml) {
  return xml
    .split(/<\/a:p>/)
    .map((para) => htmlToText([...para.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join('')))
    .filter((line) => line !== '')
    .join('\n');
}

/** Id -> Target map for a .rels part, attribute-order independent. */
function relationshipMap(rels) {
  const map = {};
  for (const tag of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = tag[0].match(/\sId="([^"]+)"/)?.[1];
    const target = tag[0].match(/\sTarget="([^"]+)"/)?.[1];
    if (id && target) map[id] = target;
  }
  return map;
}

/** Slide part paths (ppt/slides/slideN.xml) in presentation order, or null if that can't be determined. */
function orderedSlideFiles(buffer) {
  const presentation = zipEntry(buffer, 'ppt/presentation.xml')?.toString('utf8');
  const rels = zipEntry(buffer, 'ppt/_rels/presentation.xml.rels')?.toString('utf8');
  if (!presentation || !rels) return null;
  const relMap = relationshipMap(rels);
  const ids = [...presentation.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
  const files = ids.map((id) => relMap[id]).filter(Boolean).map((target) => path.posix.normalize(`ppt/${target.replace(/^\.?\//, '')}`));
  return files.length ? files : null;
}

/** Slide part paths in numeric order, reading sequentially until one is missing (cap 500). */
function numericSlideFiles(buffer) {
  const files = [];
  for (let i = 1; i <= 500; i++) {
    const name = `ppt/slides/slide${i}.xml`;
    if (!zipEntry(buffer, name)) break;
    files.push(name);
  }
  return files;
}

/** Speaker notes text for a slide part, if its rels point to a notesSlide part. */
function notesForSlide(buffer, slidePath) {
  const relsPath = `${path.posix.dirname(slidePath)}/_rels/${path.posix.basename(slidePath)}.rels`;
  const rels = zipEntry(buffer, relsPath)?.toString('utf8');
  if (!rels) return null;
  for (const tag of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    if (!/Type="[^"]*\/notesSlide"/.test(tag[0])) continue;
    const target = tag[0].match(/\sTarget="([^"]+)"/)?.[1];
    if (!target) continue;
    const notesPath = path.posix.normalize(path.posix.join(path.posix.dirname(slidePath), target));
    const xml = zipEntry(buffer, notesPath)?.toString('utf8');
    if (!xml) return null;
    const text = paragraphsToText(xml);
    return text || null;
  }
  return null;
}

/** PowerPoint slide text (and speaker notes) as plain text, or null if this isn't a readable .pptx. */
function pptxToText(buffer) {
  const files = orderedSlideFiles(buffer) ?? numericSlideFiles(buffer);
  if (!files || files.length === 0) return null;
  const parts = [];
  for (let i = 0; i < files.length; i++) {
    const xml = zipEntry(buffer, files[i])?.toString('utf8');
    if (!xml) continue;
    const text = paragraphsToText(xml);
    let notes;
    try {
      notes = notesForSlide(buffer, files[i]);
    } catch {
      notes = null; // keep it simple: skip notes rather than fail the slide
    }
    parts.push(`--- slide ${i + 1} ---\n${text}${notes ? `\nnotes: ${notes}` : ''}`);
  }
  return parts.join('\n\n');
}

module.exports = { zipEntry, pdfToText, pdfToolResult, pptxToText };
