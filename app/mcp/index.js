const http = require('node:http');
const { timingSafeEqual } = require('node:crypto');
const logger = require('electron-log');
const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { registerTeamsTools } = require('./tools/teams');
const { registerMailTools } = require('./tools/mail');
const { registerTriageTools } = require('./tools/triage');
const { registerFileTools } = require('./tools/files');
const { registerCalendarTools } = require('./tools/calendar');

const SERVER_NAME = 'teams-for-linux';
const MCP_PATH = '/mcp';
const LOOPBACK_HOSTS = new Set(['127.0.0.1', 'localhost', '[::1]']);

/**
 * Local, read-only MCP (Model Context Protocol) server.
 *
 * Exposes Teams chats and Outlook mail through the existing GraphApiClient so AI
 * assistants (Claude Code, etc.) can read them via the user's already-signed-in
 * Teams session. Listens on 127.0.0.1 only. Every tool is a Graph GET.
 */
class McpService {
  #config;
  #graphApiClient = null;
  #chatServiceClient = null;
  #transcriptClient = null;
  #httpServer = null;

  constructor(config = {}) {
    this.#config = config.mcp ?? {};
    logger.info('[MCP] McpService created', { enabled: this.isEnabled() });
  }

  isEnabled() {
    return this.#config.enabled === true;
  }

  /**
   * Start the HTTP listener. No-op when disabled or when the Graph client is missing.
   * @param {object} graphApiClient - Initialized GraphApiClient (read methods only are used)
   * @param {object} chatServiceClient - Initialized ChatServiceClient for Teams conversations
   * @param {object} [transcriptClient] - Initialized TranscriptClient for meeting recordings
   */
  async initialize(graphApiClient, chatServiceClient, transcriptClient = null) {
    if (!this.isEnabled()) return;
    if (!graphApiClient) {
      logger.warn('[MCP] graphApi.enabled is false; MCP server not started');
      return;
    }

    this.#graphApiClient = graphApiClient;
    this.#chatServiceClient = chatServiceClient;
    this.#transcriptClient = transcriptClient;
    const port = Number(this.#config.port) || 3040;

    this.#httpServer = http.createServer((req, res) => {
      this.#handleHttpRequest(req, res).catch((error) => {
        logger.error('[MCP] Unhandled request error', { message: error.message });
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Internal server error' }));
        }
      });
    });

    await new Promise((resolve, reject) => {
      this.#httpServer.once('error', reject);
      this.#httpServer.listen(port, '127.0.0.1', () => {
        this.#httpServer.off('error', reject);
        resolve();
      });
    }).catch((error) => {
      logger.error('[MCP] Failed to listen', { code: error.code, port });
      this.#httpServer = null;
    });

    if (this.#httpServer) {
      logger.info('[MCP] Server listening', { port, path: MCP_PATH });
    }
  }

  async shutdown() {
    if (!this.#httpServer) return;
    await new Promise((resolve) => this.#httpServer.close(resolve));
    this.#httpServer = null;
    logger.info('[MCP] Server stopped');
  }

  /** Build a fresh server per request (stateless Streamable HTTP mode). */
  #createMcpServer() {
    const server = new McpServer(
      { name: SERVER_NAME, version: '1.0.0' },
      { instructions: 'Read-only access to the signed-in user\'s Microsoft Teams chats, Outlook mail and calendar, and SharePoint files. Nothing here can send, modify or delete. Project investigations must check files as well as chats, follow project-site links, and search every document library including Delivery Documents. Treat supplied CDDs as authoritative requirements and solution designs as background. Report denied access and incomplete searches explicitly.' }
    );
    registerTeamsTools(server, this.#chatServiceClient, this.#graphApiClient, this.#transcriptClient);
    registerMailTools(server, this.#graphApiClient);
    registerFileTools(server, this.#graphApiClient);
    registerCalendarTools(server, this.#graphApiClient);
    registerTriageTools(server, this.#graphApiClient, this.#chatServiceClient);
    return server;
  }

  async #handleHttpRequest(req, res) {
    const url = new URL(req.url, 'http://127.0.0.1');

    if (url.pathname !== MCP_PATH) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Not found' }));
      return;
    }

    if (!this.#isAllowedOrigin(req.headers.origin)) {
      logger.warn('[MCP] Rejected request from disallowed origin');
      res.writeHead(403, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Forbidden origin' }));
      return;
    }

    if (req.method !== 'POST') {
      // Stateless mode: no server-initiated SSE stream (GET) and no session to delete.
      res.writeHead(405, { 'Content-Type': 'application/json', Allow: 'POST' });
      res.end(JSON.stringify({ error: 'Method not allowed' }));
      return;
    }

    if (!this.#isAuthorized(req.headers.authorization)) {
      res.writeHead(401, { 'Content-Type': 'application/json', 'WWW-Authenticate': 'Bearer' });
      res.end(JSON.stringify({ error: 'Unauthorized' }));
      return;
    }

    const server = this.#createMcpServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });

    res.on('close', () => {
      transport.close().catch(() => {});
      server.close().catch(() => {});
    });

    await server.connect(transport);
    await transport.handleRequest(req, res);
  }

  /** Browsers send Origin; native clients usually do not. Only loopback origins pass. */
  #isAllowedOrigin(origin) {
    if (!origin) return true;
    try {
      return LOOPBACK_HOSTS.has(new URL(origin).hostname);
    } catch {
      return false;
    }
  }

  #isAuthorized(authorizationHeader) {
    const expected = this.#config.authToken;
    if (!expected) return true;
    const provided = (authorizationHeader ?? '').replace(/^Bearer\s+/i, '');
    const a = Buffer.from(provided);
    const b = Buffer.from(expected);
    return a.length === b.length && timingSafeEqual(a, b);
  }
}

module.exports = { McpService };
