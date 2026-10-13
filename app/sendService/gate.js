const { createHash, createHmac, timingSafeEqual } = require('node:crypto');

const TOOLS = new Set(['teams_send_message', 'mail_send_reply']);
const MAX_AHEAD_SECONDS = 10 * 60;
const SENDS_PER_HOUR = 10;
const HOUR_MS = 60 * 60 * 1000;

const sha256Hex = (text) => createHash('sha256').update(text, 'utf8').digest('hex');

/** HMAC-SHA256(secret, tool \n target \n sha256hex(body) \n expiry \n nonce) as lowercase hex. */
function approvalToken(secret, { tool, target, body, expiry, nonce }) {
  const message = [tool, target, sha256Hex(body), String(expiry), nonce].join('\n');
  return createHmac('sha256', secret).update(message, 'utf8').digest('hex');
}

function tokenMatches(secret, request) {
  if (typeof request.token !== 'string' || !/^[0-9a-f]{64}$/.test(request.token)) return false;
  const expected = Buffer.from(approvalToken(secret, request), 'hex');
  return timingSafeEqual(expected, Buffer.from(request.token, 'hex'));
}

/** Nonces seen in this process, each kept until its approval expires. */
class NonceStore {
  #seen = new Map();

  /** Records the nonce. Returns false when it was used before. */
  claim(nonce, expiry, nowSeconds) {
    for (const [key, until] of this.#seen) {
      if (until < nowSeconds) this.#seen.delete(key);
    }
    if (this.#seen.has(nonce)) return false;
    this.#seen.set(nonce, expiry);
    return true;
  }

  get size() {
    return this.#seen.size;
  }
}

/** Rolling one-hour cap on send attempts. */
class RateLimiter {
  #times = [];

  constructor(limit = SENDS_PER_HOUR) {
    this.limit = limit;
  }

  take(nowMs) {
    this.#times = this.#times.filter((t) => nowMs - t < HOUR_MS);
    if (this.#times.length >= this.limit) return false;
    this.#times.push(nowMs);
    return true;
  }
}

const isStringList = (value) => value === undefined || (Array.isArray(value) && value.every((v) => typeof v === 'string'));

/**
 * Shape and content checks that need no secret.
 * @returns {{status: number, error: string} | null} null when the request is well formed
 */
function checkShape(request) {
  if (!request || typeof request !== 'object' || Array.isArray(request)) return { status: 400, error: 'body must be a JSON object' };
  const { tool, target, body, to, cc, expiry, nonce } = request;
  if (typeof tool !== 'string' || !TOOLS.has(tool)) return { status: 400, error: 'unknown tool' };
  if (typeof target !== 'string' || !target.trim()) return { status: 400, error: 'target is empty' };
  if (typeof body !== 'string' || !body.trim()) return { status: 400, error: 'body is empty' };
  if (!isStringList(to) || !isStringList(cc)) return { status: 400, error: 'to and cc must be lists of strings' };
  if (!Number.isInteger(expiry)) return { status: 400, error: 'expiry must be an integer' };
  if (typeof nonce !== 'string' || !/^[0-9a-f]{16}$/.test(nonce)) return { status: 400, error: 'nonce must be 16 hex characters' };
  if (tool === 'mail_send_reply' && !(to?.length > 0)) return { status: 400, error: 'mail reply has no to' };
  return null;
}

/**
 * Approval checks. Consumes the nonce when everything else passes.
 * @returns {{status: number, error: string} | null} null when approved
 */
function checkApproval(secret, request, nonces, nowSeconds) {
  if (!tokenMatches(secret, request)) return { status: 403, error: 'approval token does not match' };
  if (request.expiry < nowSeconds) return { status: 403, error: 'approval expired' };
  if (request.expiry > nowSeconds + MAX_AHEAD_SECONDS) return { status: 403, error: 'approval expiry too far ahead' };
  if (!nonces.claim(request.nonce, request.expiry, nowSeconds)) return { status: 403, error: 'nonce already used' };
  return null;
}

module.exports = { TOOLS, SENDS_PER_HOUR, sha256Hex, approvalToken, tokenMatches, NonceStore, RateLimiter, checkShape, checkApproval };
