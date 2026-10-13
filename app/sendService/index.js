const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');
const logger = require('electron-log');
const { TOOLS, NonceStore, RateLimiter, checkShape, checkApproval } = require('./gate');
const { checkTarget, sendTeamsMessage, sendMailReply, OutlookTokenSource } = require('./senders');

const DEFAULT_SECRET_PATH = path.join(os.homedir(), '.config', 'work-assistant', 'send-secret');
const DEFAULT_PORT = 3041;
const SEND_PATH = '/send';
const MAX_BODY_BYTES = 256 * 1024;
const MIN_SECRET_BYTES = 32;
const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '[::1]']);

/**
 * Read the shared secret. It must be a regular file readable by the owner only, with at least
 * 32 bytes once surrounding whitespace is trimmed.
 * @returns {{secret?: Buffer, reason?: string}}
 */
function loadSecret(secretPath) {
  let stat;
  try {
    stat = fs.statSync(secretPath);
  } catch {
    return { reason: 'secret file missing' };
  }
  if (!stat.isFile()) return { reason: 'secret path is not a file' };
  if ((stat.mode & 0o077) !== 0) return { reason: 'secret file must be mode 600' };
  const secret = Buffer.from(fs.readFileSync(secretPath, 'utf8').trim(), 'utf8');
  if (secret.length < MIN_SECRET_BYTES) return { reason: 'secret shorter than 32 bytes' };
  return { secret };
}

/**
 * Gated send endpoint. Sends one Teams chat message or one mail reply per request, and only with an
 * approval token minted by the work-assistant bridge when the user presses Send in the panel.
 * Listens on 127.0.0.1 only. Kept apart from app/mcp so the MCP stays read-only.
 */
class SendService {
  #secret = null;
  #port;
  #secretPath;
  #httpServer = null;
  #nonces = new NonceStore();
  #rate = new RateLimiter();
  #senders;
  #now;

  /**
   * @param {object} [options]
   * @param {string} [options.secretPath]
   * @param {number} [options.port]
   * @param {Buffer|string} [options.secret] - Use this secret instead of reading the file (tests)
   * @param {object} [options.senders] - { teams(request), mail(request) } returning { status, id?, error? }
   * @param {() => number} [options.now] - Clock in milliseconds
   */
  constructor({ secretPath = DEFAULT_SECRET_PATH, port = DEFAULT_PORT, secret, senders, now = Date.now } = {}) {
    this.#secretPath = secretPath;
    this.#port = port;
    this.#secret = secret ? Buffer.from(secret) : null;
    this.#senders = senders ?? null;
    this.#now = now;
  }

  /** Start the listener when the secret file exists. Logs one line and returns when it does not. */
  async initialize({ graphApiClient, chatServiceClient, mainWindow } = {}) {
    if (!this.#secret) {
      const loaded = loadSecret(this.#secretPath);
      if (!loaded.secret) {
        logger.info('[SEND] Send service not started', { reason: loaded.reason });
        return;
      }
      this.#secret = loaded.secret;
    }
    if (!this.#senders) {
      if (!graphApiClient || !chatServiceClient) {
        logger.warn('[SEND] Send service not started', { reason: 'Graph or chat service client missing' });
        return;
      }
      this.#senders = this.#defaultSenders({ graphApiClient, chatServiceClient, mainWindow });
    }

    this.#httpServer = http.createServer((req, res) => {
      this.#handleHttpRequest(req, res).catch((error) => {
        logger.error('[SEND] Unhandled request error', { message: error.message });
        if (!res.headersSent) this.#reply(res, 500, { error: 'internal error' });
      });
    });
    await new Promise((resolve, reject) => {
      this.#httpServer.once('error', reject);
      this.#httpServer.listen(this.#port, '127.0.0.1', () => {
        this.#httpServer.off('error', reject);
        resolve();
      });
    }).catch((error) => {
      logger.error('[SEND] Failed to listen', { code: error.code, port: this.#port });
      this.#httpServer = null;
    });
    if (this.#httpServer) logger.info('[SEND] Send service listening', { port: this.#port, path: SEND_PATH });
  }

  /** Bound address while listening (tests use port 0). */
  address() {
    return this.#httpServer?.address() ?? null;
  }

  async shutdown() {
    if (!this.#httpServer) return;
    await new Promise((resolve) => this.#httpServer.close(resolve));
    this.#httpServer = null;
    logger.info('[SEND] Send service stopped');
  }

  #defaultSenders({ graphApiClient, chatServiceClient, mainWindow }) {
    let displayName = null;
    const getDisplayName = async () => {
      if (displayName) return displayName;
      const profile = await graphApiClient.getUserProfile();
      displayName = profile?.success ? profile.data?.displayName ?? null : null;
      return displayName;
    };
    const outlookTokens = new OutlookTokenSource(() => mainWindow?.webContents ?? null);
    return {
      teams: (request) => sendTeamsMessage({ chatServiceClient, getDisplayName }, request),
      mail: (request) => sendMailReply({ graphApiClient, outlookTokens }, request),
    };
  }

  /**
   * Check and carry out one send request.
   * @param {object} request - Parsed JSON body
   * @returns {Promise<{status: number, body: object}>}
   */
  async handle(request) {
    const result = await this.#check(request);
    const tool = TOOLS.has(request?.tool) ? request.tool : 'unknown';
    const target = typeof request?.target === 'string' ? request.target.slice(0, 30) : '';
    logger.info('[SEND] Send request', { tool, status: result.status, target });
    if (result.status === 200) return { status: 200, body: { ok: true, id: result.id ?? null } };
    return { status: result.status, body: { error: result.error } };
  }

  async #check(request) {
    const shape = checkShape(request);
    if (shape) return shape;
    const target = checkTarget(request.tool, request.target, request.to, request.cc);
    if (target) return target;
    const nowMs = this.#now();
    const approval = checkApproval(this.#secret, request, this.#nonces, Math.floor(nowMs / 1000));
    if (approval) return approval;
    if (!this.#rate.take(nowMs)) return { status: 429, error: 'send limit reached (10 an hour)' };
    try {
      return request.tool === 'teams_send_message' ? await this.#senders.teams(request) : await this.#senders.mail(request);
    } catch (error) {
      return { status: 502, error: error.message };
    }
  }

  async #handleHttpRequest(req, res) {
    const url = new URL(req.url, 'http://127.0.0.1');
    if (url.pathname !== SEND_PATH) return this.#reply(res, 404, { error: 'not found' });
    if (!this.#isAllowedOrigin(req.headers.origin)) {
      logger.warn('[SEND] Rejected request from disallowed origin');
      return this.#reply(res, 403, { error: 'forbidden origin' });
    }
    if (req.method !== 'POST') return this.#reply(res, 405, { error: 'method not allowed' }, { Allow: 'POST' });

    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) return this.#reply(res, 413, { error: 'request too large' });
      chunks.push(chunk);
    }
    let request;
    try {
      request = JSON.parse(Buffer.concat(chunks).toString('utf8'));
    } catch {
      return this.#reply(res, 400, { error: 'body is not JSON' });
    }
    const { status, body } = await this.handle(request);
    return this.#reply(res, status, body);
  }

  #reply(res, status, body, headers = {}) {
    res.writeHead(status, { 'Content-Type': 'application/json', ...headers });
    res.end(JSON.stringify(body));
  }

  /** Same rule as the MCP server: no Origin, or a loopback Origin. */
  #isAllowedOrigin(origin) {
    if (!origin) return true;
    try {
      return LOOPBACK_HOSTS.has(new URL(origin).hostname);
    } catch {
      return false;
    }
  }
}

module.exports = { SendService, loadSecret, DEFAULT_SECRET_PATH, DEFAULT_PORT };
