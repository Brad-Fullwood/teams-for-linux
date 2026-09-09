/**
 * Small formatting helpers shared by the MCP tools.
 * Output goes to an AI assistant, so favour compact, scannable text.
 */

const ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
};

/** Convert Teams/Outlook HTML into plain text. Good enough for reading, not a full parser. */
function htmlToText(html) {
  if (!html) return '';
  return String(html)
    .replaceAll(/<style[\s\S]*?<\/style>/gi, '')
    .replaceAll(/<script[\s\S]*?<\/script>/gi, '')
    .replaceAll(/<br\s*\/?>/gi, '\n')
    .replaceAll(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, '\n')
    .replaceAll(/<li[^>]*>/gi, '- ')
    .replaceAll(/<at[^>]*>([^<]*)<\/at>/gi, '@$1')
    // Keep link targets: "text (url)" unless the text already is the url.
    .replaceAll(/<a\s[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, (m, href, inner) => {
      const text = inner.replaceAll(/<[^>]+>/g, '').trim();
      if (!/^https?:/i.test(href)) return text;
      return text && text !== href && !href.startsWith(text) ? `${text} (${href})` : href;
    })
    .replaceAll(/<[^>]+>/g, '')
    .replaceAll(/&(#x?[0-9a-f]+|[a-z]+);/gi, (match, code) => {
      if (code[0] === '#') {
        const num = code[1].toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10);
        return Number.isNaN(num) ? match : String.fromCodePoint(num);
      }
      return ENTITIES[code.toLowerCase()] ?? match;
    })
    .replaceAll(/[ \t]+\n/g, '\n')
    .replaceAll(/\n{3,}/g, '\n\n')
    .trim();
}

// Zero-width and joiner characters that newsletters pad previews with.
const INVISIBLE = /[\u034f\u200b-\u200d\u2060\ufeff]/g;

function truncate(text, max = 400) {
  const clean = (text ?? '').replaceAll(INVISIBLE, '').replaceAll(/[ \t]{2,}/g, ' ');
  if (clean.length <= max) return clean;
  return `${clean.slice(0, max - 1)}…`;
}

function formatDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? iso : d.toISOString().replace('T', ' ').slice(0, 16) + 'Z';
}

function isoDaysAgo(days) {
  return new Date(Date.now() - days * 86_400_000).toISOString();
}

/** Escape a value for use inside an OData single-quoted string literal. */
function odataString(value) {
  return String(value).replaceAll("'", "''");
}

/** Shape a Graph result into MCP tool content. Errors become isError results, never throws. */
function toolResult(result, render) {
  if (!result?.success) {
    const status = result?.status ? ` (HTTP ${result.status})` : '';
    return {
      isError: true,
      content: [{ type: 'text', text: `Graph request failed${status}: ${result?.error ?? 'unknown error'}` }],
    };
  }
  const text = render(result.data);
  return { content: [{ type: 'text', text: text || '(no results)' }] };
}

function personName(person) {
  return person?.emailAddress?.name || person?.emailAddress?.address || person?.user?.displayName || person?.application?.displayName || 'unknown';
}

function personAddress(person) {
  return person?.emailAddress?.address ? ` <${person.emailAddress.address}>` : '';
}

module.exports = { htmlToText, truncate, formatDate, isoDaysAgo, odataString, toolResult, personName, personAddress };
