const { randomInt } = require('node:crypto');
const { textToHtml, parseRecipients } = require('./format');

const REQUEST_TIMEOUT_MS = 30000;
const OUTLOOK_RESOURCE = 'https://outlook.office.com';
const OUTLOOK_REPLY_URL = 'https://outlook.office.com/api/v2.0/me/messages';
const TOKEN_EXPIRY_BUFFER_S = 5 * 60;
const NOTES_TO_SELF = '48:notes';
const CHAT_ID = /^19:[^\s/?#;@]+@[a-z0-9.]+$/i;
const MAIL_ID = /^[A-Za-z0-9+/=_-]{16,512}$/;

/** Decode a JWT payload. Returns {} when the token is not a JWT. The token itself is not kept. */
function jwtClaims(token) {
  try {
    const part = String(token).split('.')[1];
    return JSON.parse(Buffer.from(part.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
  } catch {
    return {};
  }
}

const scopesOf = (token) => new Set(String(jwtClaims(token).scp || '').split(' ').filter(Boolean));

/**
 * Target checks per tool.
 * @returns {{status: number, error: string} | null}
 */
function checkTarget(tool, target, to = [], cc = []) {
  if (tool === 'teams_send_message') {
    if (/@thread\.tacv2|;messageid=/i.test(target)) return { status: 400, error: 'channel threads are not supported' };
    if (target !== NOTES_TO_SELF && !CHAT_ID.test(target)) return { status: 400, error: 'target is not a chat id' };
    return null;
  }
  if (!MAIL_ID.test(target)) return { status: 400, error: 'target is not a mail message id' };
  if (parseRecipients(to).invalid || parseRecipients(cc).invalid) return { status: 400, error: 'invalid recipient' };
  return null;
}

const randomDigits = (count) => Array.from({ length: count }, (_, i) => randomInt(i === 0 ? 1 : 0, 10)).join('');

async function readJson(response) {
  const text = await response.text();
  try {
    return text ? JSON.parse(text) : null;
  } catch {
    return null;
  }
}

const upstreamError = (service, status, data) => {
  const detail = data?.error?.message || data?.message || data?.errorCode || '';
  return { status: 502, error: `${service} HTTP ${status}${detail ? `: ${String(detail).slice(0, 200)}` : ''}` };
};

const networkError = (error) => error.name === 'TimeoutError' || error.name === 'AbortError'
  ? { status: 504, error: 'send timed out, check the thread before sending again' }
  : { status: 502, error: error.message };

/**
 * Post a chat message through the Teams chat service, inside ChatServiceClient's throttle and session.
 * @param {{chatServiceClient: object, getDisplayName: () => Promise<string>}} deps
 */
async function sendTeamsMessage(deps, { target, body }) {
  const payload = {
    content: textToHtml(body),
    messagetype: 'RichText/Html',
    contenttype: 'text',
    clientmessageid: randomDigits(19),
    imdisplayname: (await deps.getDisplayName()) || '',
  };
  const result = await deps.chatServiceClient.withSession(async ({ skypeToken, baseUrl }) => {
    try {
      const response = await fetch(`${baseUrl}/v1/users/ME/conversations/${encodeURIComponent(target)}/messages`, {
        method: 'POST',
        headers: { Authentication: `skypetoken=${skypeToken}`, 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      const data = await readJson(response);
      // The real status goes back so withSession can retry on 401 and 429.
      if (!response.ok) return { status: response.status, error: upstreamError('chat service', response.status, data).error };
      const location = response.headers?.get?.('location') || '';
      const id = location.split('/').pop() || (data?.OriginalArrivalTime ? String(data.OriginalArrivalTime) : null);
      return { status: 200, id };
    } catch (error) {
      return networkError(error);
    }
  });
  if (result.status === 200) return result;
  return { status: result.status === 504 ? 504 : 502, error: result.error || 'send failed' };
}

/** Outlook-audience token from the Teams auth provider, cached in memory until 5 minutes before expiry. */
class OutlookTokenSource {
  #getWebContents;
  #token = null;
  #expiry = 0;

  constructor(getWebContents) {
    this.#getWebContents = getWebContents;
  }

  async get() {
    if (this.#token && this.#expiry - Date.now() / 1000 > TOKEN_EXPIRY_BUFFER_S) return { success: true, token: this.#token };
    const webContents = this.#getWebContents();
    if (!webContents) return { success: false, error: 'main window not initialized' };
    try {
      const result = await webContents.executeJavaScript(`
        (async () => {
          const handler = window.teamsForLinuxReactHandler;
          if (!handler?.acquireToken) return { success: false, error: 'ReactHandler not available' };
          return await handler.acquireToken(${JSON.stringify(OUTLOOK_RESOURCE)}, {});
        })()
      `);
      if (!result?.success || !result.token) return { success: false, error: result?.error || 'no Outlook token' };
      this.#token = result.token;
      this.#expiry = Number(jwtClaims(result.token).exp) || Date.now() / 1000 + 3600;
      return { success: true, token: this.#token };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
}

/**
 * Reply to a mail message with exactly the given recipients. Uses Graph when its token carries Mail.Send,
 * else Outlook REST v2.0 with an Outlook-audience token that carries Mail.Send, else 501.
 * @param {{graphApiClient: object, outlookTokens: {get: () => Promise<object>}}} deps
 */
async function sendMailReply(deps, { target, body, to, cc }) {
  const html = textToHtml(body);
  const toList = parseRecipients(to).recipients;
  const ccList = parseRecipients(cc).recipients;

  const graphToken = await deps.graphApiClient.acquireToken();
  if (graphToken?.success && scopesOf(graphToken.token).has('Mail.Send')) {
    const recipient = ({ name, address }) => ({ emailAddress: name ? { name, address } : { address } });
    const result = await deps.graphApiClient.makeRequest(`/me/messages/${encodeURIComponent(target)}/reply`, {
      method: 'POST',
      sensitive: true,
      body: { comment: html, message: { toRecipients: toList.map(recipient), ccRecipients: ccList.map(recipient) } },
    });
    if (result.success) return { status: 200, id: null };
    return upstreamError('Graph', result.status ?? 'error', { message: result.error });
  }

  const outlook = await deps.outlookTokens.get();
  if (!outlook.success) return { status: 501, error: `mail send not available: Graph token has no Mail.Send and no Outlook token (${outlook.error})` };
  if (!scopesOf(outlook.token).has('Mail.Send')) return { status: 501, error: 'mail send not available: neither the Graph nor the Outlook token has Mail.Send' };

  const recipient = ({ name, address }) => ({ EmailAddress: name ? { Name: name, Address: address } : { Address: address } });
  try {
    const response = await fetch(`${OUTLOOK_REPLY_URL}/${encodeURIComponent(target)}/reply`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${outlook.token}`, 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify({ Comment: html, Message: { ToRecipients: toList.map(recipient), CcRecipients: ccList.map(recipient) } }),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
    if (!response.ok) return upstreamError('Outlook', response.status, await readJson(response));
    return { status: 200, id: null };
  } catch (error) {
    return networkError(error);
  }
}

module.exports = { jwtClaims, scopesOf, checkTarget, sendTeamsMessage, sendMailReply, OutlookTokenSource, NOTES_TO_SELF };
