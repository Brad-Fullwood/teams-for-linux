const { createHash, createHmac, timingSafeEqual } = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const TOOLS = new Set(['teams_send_message', 'mail_send_reply', 'mail_create_draft']);
const WORK_ASSISTANT = path.join(os.homedir(), 'Projects', 'Personal', 'work-assistant');
/** mail_create_draft attaches files from these folders only, compared after symlinks are resolved. */
const ATTACHMENT_ROOTS = [path.join(WORK_ASSISTANT, 'attachments'), path.join(WORK_ASSISTANT, 'tmp')];
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;
const MAX_AHEAD_SECONDS = 10 * 60;
const SENDS_PER_HOUR = 10;
const HOUR_MS = 60 * 60 * 1000;

const sha256Hex = (data) => createHash('sha256').update(data, typeof data === 'string' ? 'utf8' : undefined).digest('hex');

/**
 * The text the approval HMAC signs.
 * Sends: tool \n target \n sha256hex(body) \n expiry \n nonce.
 * mail_create_draft: tool \n sha256hex(to joined by \n) \n sha256hex(subject) \n sha256hex(body)
 * \n sha256hex(attachment bytes) \n expiry \n nonce.
 */
function approvalMessage({ tool, target, body, to, subject, attachmentSha256, expiry, nonce }) {
  if (tool === 'mail_create_draft') {
    return [tool, sha256Hex((to ?? []).join('\n')), sha256Hex(subject ?? ''), sha256Hex(body), attachmentSha256 ?? '', String(expiry), nonce].join('\n');
  }
  return [tool, target, sha256Hex(body), String(expiry), nonce].join('\n');
}

/** HMAC-SHA256(secret, approvalMessage(fields)) as lowercase hex. */
function approvalToken(secret, fields) {
  return createHmac('sha256', secret).update(approvalMessage(fields), 'utf8').digest('hex');
}

/**
 * @param {object} [server] - Values the service computed itself. attachmentSha256 replaces any value in the request.
 */
function tokenMatches(secret, request, { attachmentSha256 } = {}) {
  if (typeof request.token !== 'string' || !/^[0-9a-f]{64}$/.test(request.token)) return false;
  const expected = Buffer.from(approvalToken(secret, { ...request, attachmentSha256 }), 'hex');
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

const isUnder = (file, root) => file.startsWith(root.endsWith(path.sep) ? root : root + path.sep);

/**
 * Resolve an attachment path, symlinks followed, and require a regular file inside one of the roots.
 * @returns {{path: string, size: number} | {status: number, error: string}}
 */
function resolveAttachment(attachmentPath, roots = ATTACHMENT_ROOTS) {
  if (typeof attachmentPath !== 'string' || !path.isAbsolute(attachmentPath)) return { status: 400, error: 'attachment_path must be an absolute path' };
  let real;
  try {
    real = fs.realpathSync(attachmentPath);
  } catch {
    return { status: 400, error: 'attachment not found' };
  }
  const realRoots = roots.flatMap((root) => {
    try {
      return [fs.realpathSync(root)];
    } catch {
      return [];
    }
  });
  if (!realRoots.some((root) => isUnder(real, root))) return { status: 400, error: 'attachment is outside the allowed folders' };
  const stat = fs.statSync(real);
  if (!stat.isFile()) return { status: 400, error: 'attachment is not a file' };
  return { path: real, size: stat.size };
}

/**
 * Resolve and read an attachment, refusing files over 3 MB.
 * @returns {{name: string, bytes: Buffer, sha256: string} | {status: number, error: string}}
 */
function readAttachment(attachmentPath, roots = ATTACHMENT_ROOTS) {
  const resolved = resolveAttachment(attachmentPath, roots);
  if (resolved.error) return resolved;
  if (resolved.size > MAX_ATTACHMENT_BYTES) return { status: 400, error: 'attachment is over 3 MB' };
  const bytes = fs.readFileSync(resolved.path);
  if (bytes.length > MAX_ATTACHMENT_BYTES) return { status: 400, error: 'attachment is over 3 MB' };
  return { name: path.basename(resolved.path), bytes, sha256: sha256Hex(bytes) };
}

/**
 * Shape and content checks that need no secret.
 * @returns {{status: number, error: string} | null} null when the request is well formed
 */
function checkShape(request, { attachmentRoots = ATTACHMENT_ROOTS } = {}) {
  if (!request || typeof request !== 'object' || Array.isArray(request)) return { status: 400, error: 'body must be a JSON object' };
  const { tool, target, body, to, cc, expiry, nonce } = request;
  if (typeof tool !== 'string' || !TOOLS.has(tool)) return { status: 400, error: 'unknown tool' };
  if (tool === 'mail_create_draft') return checkDraftShape(request, attachmentRoots);
  if (typeof target !== 'string' || !target.trim()) return { status: 400, error: 'target is empty' };
  if (typeof body !== 'string' || !body.trim()) return { status: 400, error: 'body is empty' };
  if (!isStringList(to) || !isStringList(cc)) return { status: 400, error: 'to and cc must be lists of strings' };
  if (!Number.isInteger(expiry)) return { status: 400, error: 'expiry must be an integer' };
  if (typeof nonce !== 'string' || !/^[0-9a-f]{16}$/.test(nonce)) return { status: 400, error: 'nonce must be 16 hex characters' };
  if (tool === 'mail_send_reply' && !(to?.length > 0)) return { status: 400, error: 'mail reply has no to' };
  return null;
}

function checkDraftShape({ to, cc, subject, body, attachment_path: attachmentPath, expiry, nonce }, attachmentRoots) {
  if (!Array.isArray(to) || to.length === 0 || !to.every((v) => typeof v === 'string')) return { status: 400, error: 'draft needs to as a list of one or more addresses' };
  if (to.some((v) => /[\r\n]/.test(v))) return { status: 400, error: 'invalid recipient' };
  if (cc !== undefined && !(Array.isArray(cc) && cc.length === 0)) return { status: 400, error: 'draft does not take cc' };
  if (typeof subject !== 'string' || !subject.trim()) return { status: 400, error: 'subject is empty' };
  if (typeof body !== 'string' || !body.trim()) return { status: 400, error: 'body is empty' };
  if (!Number.isInteger(expiry)) return { status: 400, error: 'expiry must be an integer' };
  if (typeof nonce !== 'string' || !/^[0-9a-f]{16}$/.test(nonce)) return { status: 400, error: 'nonce must be 16 hex characters' };
  const attachment = resolveAttachment(attachmentPath, attachmentRoots);
  return attachment.error ? attachment : null;
}

/**
 * Approval checks. Consumes the nonce when everything else passes.
 * @param {object} [server] - { attachmentSha256 } computed by the service from the file it will attach
 * @returns {{status: number, error: string} | null} null when approved
 */
function checkApproval(secret, request, nonces, nowSeconds, server = {}) {
  if (!tokenMatches(secret, request, server)) return { status: 403, error: 'approval token does not match' };
  if (request.expiry < nowSeconds) return { status: 403, error: 'approval expired' };
  if (request.expiry > nowSeconds + MAX_AHEAD_SECONDS) return { status: 403, error: 'approval expiry too far ahead' };
  if (!nonces.claim(request.nonce, request.expiry, nowSeconds)) return { status: 403, error: 'nonce already used' };
  return null;
}

module.exports = {
  TOOLS, SENDS_PER_HOUR, ATTACHMENT_ROOTS, MAX_ATTACHMENT_BYTES, sha256Hex, approvalMessage, approvalToken, tokenMatches,
  NonceStore, RateLimiter, resolveAttachment, readAttachment, checkShape, checkApproval,
};
