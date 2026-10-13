/** Escape the characters that break Teams and Outlook HTML. */
const escapeHtml = (text) => text
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

/**
 * Plain text to HTML: one <p> per paragraph (blank line separated), <br> for single line breaks.
 * Emoji and other Unicode pass through unchanged.
 */
function textToHtml(text) {
  return text
    .replace(/\r\n?/g, '\n')
    .split(/\n[ \t]*\n+/)
    .map((paragraph) => paragraph.replace(/^\n+|\n+$/g, ''))
    .filter((paragraph) => paragraph.trim())
    .map((paragraph) => `<p>${paragraph.split('\n').map(escapeHtml).join('<br>')}</p>`)
    .join('');
}

const ADDRESS = /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/;

/**
 * Parse "Name <addr>", "\"Last, First\" <addr>" or a bare address.
 * @returns {{name: string, address: string} | null} null when no valid address is found
 */
function parseRecipient(entry) {
  const text = String(entry).trim();
  const angled = /^(.*?)\s*<([^<>]+)>$/.exec(text);
  const address = (angled ? angled[2] : text).trim();
  if (!ADDRESS.test(address)) return null;
  const name = angled ? angled[1].trim().replace(/^"(.*)"$/, '$1').trim() : '';
  return { name, address };
}

/**
 * Parse a recipient list.
 * @returns {{recipients: Array<{name: string, address: string}>, invalid: number}}
 */
function parseRecipients(list = []) {
  const recipients = [];
  let invalid = 0;
  for (const entry of list) {
    const parsed = parseRecipient(entry);
    if (parsed) recipients.push(parsed);
    else invalid++;
  }
  return { recipients, invalid };
}

module.exports = { escapeHtml, textToHtml, parseRecipient, parseRecipients };
