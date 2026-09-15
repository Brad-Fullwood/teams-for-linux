'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const zlib = require('node:zlib');
const { pdfToText, pptxToText, zipEntry } = require('../../app/mcp/tools/documents');

// --- Minimal hand-built PDF fixture -----------------------------------------------------
// One object per page/content/font, a correct xref table with offsets computed as the
// body is written out, matching what a real (if tiny) PDF producer would emit.
function makePdf(pagesText) {
  const objects = [];
  const pageCount = pagesText.length;
  const pageObjNums = pagesText.map((_, i) => 3 + i);
  const fontObjNum = 3 + pageCount;
  const contentObjNums = pagesText.map((_, i) => fontObjNum + 1 + i);

  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[2] = `<< /Type /Pages /Kids [${pageObjNums.map((n) => `${n} 0 R`).join(' ')}] /Count ${pageCount} >>`;
  for (let i = 0; i < pageCount; i++) {
    objects[pageObjNums[i]] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 ${fontObjNum} 0 R >> >> /Contents ${contentObjNums[i]} 0 R >>`;
  }
  objects[fontObjNum] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';
  for (let i = 0; i < pageCount; i++) {
    const text = pagesText[i];
    const stream = text ? `BT /F1 12 Tf 10 100 Td (${text.replaceAll(/([()\\])/g, String.raw`\$1`)}) Tj ET` : '';
    objects[contentObjNums[i]] = { stream };
  }

  const maxObjNum = contentObjNums[contentObjNums.length - 1];
  let out = '%PDF-1.4\n';
  const offsets = new Array(maxObjNum + 1).fill(0);
  for (let n = 1; n <= maxObjNum; n++) {
    offsets[n] = Buffer.byteLength(out, 'latin1');
    const obj = objects[n];
    out += typeof obj === 'string'
      ? `${n} 0 obj\n${obj}\nendobj\n`
      : `${n} 0 obj\n<< /Length ${Buffer.byteLength(obj.stream, 'latin1')} >>\nstream\n${obj.stream}\nendstream\nendobj\n`;
  }
  const xrefOffset = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${maxObjNum + 1}\n0000000000 65535 f \n`;
  for (let n = 1; n <= maxObjNum; n++) out += `${String(offsets[n]).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${maxObjNum + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return Buffer.from(out, 'latin1');
}

test('pdfToText reads a real two-page PDF with page markers', async () => {
  const buffer = makePdf(['Hello PDF page one', 'Hello PDF page two']);
  const result = await pdfToText(buffer);
  assert.equal(result.pages, 2);
  assert.equal(result.pagesRead, 2);
  assert.equal(result.hasText, true);
  assert.match(result.text, /--- page 1 ---\nHello PDF page one/);
  assert.match(result.text, /--- page 2 ---\nHello PDF page two/);
});

test('pdfToText respects maxPages and notes truncation', async () => {
  const buffer = makePdf(['One', 'Two', 'Three']);
  const result = await pdfToText(buffer, { maxPages: 1 });
  assert.equal(result.pages, 3);
  assert.equal(result.pagesRead, 1);
  assert.match(result.text, /--- page 1 ---\nOne/);
  assert.doesNotMatch(result.text, /page 2/);
  assert.match(result.text, /truncated: showing 1 of 3 pages/);
});

test('pdfToText reports no text layer for a page with no content', async () => {
  const buffer = makePdf(['']);
  const result = await pdfToText(buffer);
  assert.equal(result.hasText, false);
  assert.equal(result.error, undefined);
});

test('pdfToText handles corrupt/garbage bytes without throwing', async () => {
  const result = await pdfToText(Buffer.from('not a pdf, just some random bytes 12345'));
  assert.equal(result.hasText, false);
  assert.ok(result.error);
  assert.doesNotMatch(result.error, /\n\s+at /); // no stack trace leaking through
});

// --- pdfToolResult: readable PDFs are returned as text and never written to disk -----------
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pdfToolResult } = require('../../app/mcp/tools/documents');

function toolArgs(saveDir) {
  return { header: 'doc.pdf (1 KB)\n', name: 'doc.pdf', maxPages: 200, maxChars: 60_000, saveDir };
}

test('pdfToolResult returns text and saves nothing when the PDF has text', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-tool-'));
  const result = await pdfToolResult(makePdf(['Readable page']), toolArgs(dir));
  assert.equal(result.isError, undefined);
  assert.match(result.content[0].text, /--- page 1 ---\nReadable page/);
  assert.deepEqual(fs.readdirSync(dir), []);
});

test('pdfToolResult saves a copy and says so when there is no text layer', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-tool-'));
  const result = await pdfToolResult(makePdf(['']), toolArgs(dir));
  assert.match(result.content[0].text, /No text layer found .*saved to /);
  assert.deepEqual(fs.readdirSync(dir), ['doc.pdf']);
});

test('pdfToolResult reports an error and saves a copy for unreadable bytes', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-tool-'));
  const result = await pdfToolResult(Buffer.from('not a pdf'), toolArgs(dir));
  assert.equal(result.isError, true);
  assert.deepEqual(fs.readdirSync(dir), ['doc.pdf']);
});

// --- Minimal in-memory PPTX fixture (STORED/uncompressed zip) --------------------------
function crc32(buf) {
  if (typeof zlib.crc32 === 'function') return zlib.crc32(buf);
  const table = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c;
  }
  let crc = 0xffffffff;
  for (const byte of buf) crc = table[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** Build a minimal STORED-method (uncompressed) zip: local headers, central directory, EOCD. */
function buildZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const { name, content } of entries) {
    const data = Buffer.from(content, 'utf8');
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    localParts.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuf);

    offset += local.length + nameBuf.length + data.length;
  }
  const centralStart = offset;
  const centralBuf = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(centralStart, 16);
  return Buffer.concat([...localParts, centralBuf, eocd]);
}

function slideXml(text) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>`;
}

function makePptx() {
  const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>
</p:presentation>`;
  const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`;
  const slide1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>`;
  const notesSlide1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Speaker note for slide one</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:notes>`;

  return buildZip([
    { name: 'ppt/presentation.xml', content: presentationXml },
    { name: 'ppt/_rels/presentation.xml.rels', content: presentationRels },
    // Written out of numeric order to prove ordering comes from sldIdLst, not file order.
    { name: 'ppt/slides/slide2.xml', content: slideXml('Second slide title') },
    { name: 'ppt/slides/slide1.xml', content: slideXml('First slide &amp; title') },
    { name: 'ppt/slides/_rels/slide1.xml.rels', content: slide1Rels },
    { name: 'ppt/notesSlides/notesSlide1.xml', content: notesSlide1 },
  ]);
}

test('zipEntry still reads a STORED entry (shared with files.js/mail.js)', () => {
  const zip = buildZip([{ name: 'hello.txt', content: 'hi there' }]);
  assert.equal(zipEntry(zip, 'hello.txt').toString(), 'hi there');
  assert.equal(zipEntry(zip, 'missing.txt'), null);
});

test('pptxToText orders slides via sldIdLst, decodes entities and includes notes', () => {
  const text = pptxToText(makePptx());
  const slide1 = text.indexOf('--- slide 1 ---');
  const slide2 = text.indexOf('--- slide 2 ---');
  assert.ok(slide1 >= 0 && slide2 > slide1);
  assert.match(text, /--- slide 1 ---\nFirst slide & title\nnotes: Speaker note for slide one/);
  assert.match(text, /--- slide 2 ---\nSecond slide title/);
});

test('pptxToText falls back to numeric slide order without a presentation.xml', () => {
  const zip = buildZip([
    { name: 'ppt/slides/slide1.xml', content: slideXml('Only slide') },
  ]);
  const text = pptxToText(zip);
  assert.equal(text, '--- slide 1 ---\nOnly slide');
});

test('pptxToText returns null for something that is not a pptx', () => {
  const zip = buildZip([{ name: 'word/document.xml', content: '<w:document/>' }]);
  assert.equal(pptxToText(zip), null);
});
