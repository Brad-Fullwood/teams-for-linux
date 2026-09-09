const logger = require('electron-log');

const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000;
const REQUEST_TIMEOUT_MS = 30000;
const MAX_MESSAGE_PAGE = 200;      // service returns 400 above this
const MAX_CONVERSATION_PAGE = 500;
// The service enforces roughly 15 calls per 10 seconds per user; space requests out
// and back off on 429 so a search scan or a busy assistant does not trip it.
const MIN_REQUEST_GAP_MS = 700;
const MAX_429_RETRIES = 3;
// Teams/channels directory ("CSA"): needs an AAD token for the chat aggregator plus the Skype token.
const CSA_TEAMS_URL = 'https://teams.microsoft.com/api/csa/api/v1/teams/users/me?isPrefetch=false&enableMembershipSummary=true';
const CSA_RESOURCE = 'https://chatsvcagg.teams.microsoft.com';

/**
 * Read-only client for the Teams chat service (the "chatsvc" API the Teams web
 * client itself uses for conversations and messages).
 *
 * Authentication reuses the signed-in Teams session: the web client's discovery
 * service already holds a Skype token and the regional chat service URL, and
 * both are read from the renderer via the ReactHandler. Requests are then made
 * from the main process. Only GET is supported.
 */
class ChatServiceClient {
  #mainWindow = null;
  #skypeToken = null;
  #tokenExpiry = 0;
  #baseUrl = null;
  #queue = Promise.resolve();
  #lastRequestAt = 0;

  constructor(config = {}) {
    this.enabled = config.graphApi?.enabled ?? false;
  }

  initialize(mainWindow) {
    this.#mainWindow = mainWindow;
  }

  isEnabled() {
    return this.enabled;
  }

  /**
   * Fetch (or reuse) the Skype token and chat service URL from the Teams renderer.
   * @returns {Promise<{success: boolean, error?: string}>}
   */
  async acquireSession(forceRefresh = false) {
    if (!this.#mainWindow?.webContents) {
      return { success: false, error: 'Main window not initialized' };
    }

    if (!forceRefresh && this.#skypeToken && this.#tokenExpiry - Date.now() > TOKEN_EXPIRY_BUFFER_MS) {
      return { success: true };
    }

    try {
      const result = await this.#mainWindow.webContents.executeJavaScript(`
        (async () => {
          const handler = window.teamsForLinuxReactHandler;
          const coreServices = handler?._getTeams2CoreServices?.();
          const discover = coreServices?.gtmRegistry?.discover;
          const authService = coreServices?.authenticationService;
          if (!discover?.getSkypeTokenInfoFromDiscoveryResponse || !authService?.getUser) {
            return { success: false, error: 'Teams discovery service not available' };
          }
          let user = authService.getUser();
          if (user && typeof user.then === 'function') user = await user;
          if (!user) return { success: false, error: 'No signed-in user' };
          const info = await discover.getSkypeTokenInfoFromDiscoveryResponse({
            userIdentifier: user,
            correlation: { id: crypto.randomUUID(), source: 'teams-for-linux:chatService', scenarioName: 'GetSkypeToken' }
          });
          if (!info?.skypeToken || !info?.regionGtms?.chatService) {
            return { success: false, error: 'Discovery response missing token or chat service URL' };
          }
          return { success: true, skypeToken: info.skypeToken, expiration: info.expiration, chatService: info.regionGtms.chatService };
        })()
      `);

      if (!result?.success) {
        logger.warn('[CHAT_SVC] Session acquisition failed', { error: result?.error });
        return { success: false, error: result?.error || 'Unknown error' };
      }

      this.#skypeToken = result.skypeToken;
      this.#tokenExpiry = Number(result.expiration) || Date.now() + 60 * 60 * 1000;
      this.#baseUrl = String(result.chatService).replace(/\/+$/, '');
      logger.debug('[CHAT_SVC] Session acquired', { expiresInMin: Math.round((this.#tokenExpiry - Date.now()) / 60000) });
      return { success: true };
    } catch (error) {
      logger.error('[CHAT_SVC] Session acquisition error', { message: error.message });
      return { success: false, error: error.message };
    }
  }

  /**
   * GET a chat service path. Retries once with a fresh session on 401.
   * @param {string} path - Path under the regional chat service, e.g. /v1/users/ME/conversations
   * @param {URLSearchParams|object} [params]
   * @returns {Promise<{success: boolean, data?: object, error?: string, status?: number}>}
   */
  async get(path, params = {}, retry = true) {
    if (!this.enabled) {
      return { success: false, error: 'Graph API is disabled' };
    }

    const session = await this.acquireSession();
    if (!session.success) return session;

    for (let attempt = 0; ; attempt++) {
      const result = await this.#throttled(() => this.#fetchOnce(path, params, retry));
      if (result.status !== 429 || attempt >= MAX_429_RETRIES) return result;
      const waitMs = (Number(result.retryAfter) || 3 * (attempt + 1)) * 1000;
      logger.info('[CHAT_SVC] Rate limited, backing off', { waitMs, attempt });
      await new Promise((resolve) => setTimeout(resolve, waitMs));
    }
  }

  /** Serialize requests with a minimum gap between them. */
  #throttled(fn) {
    const run = this.#queue.then(async () => {
      const wait = this.#lastRequestAt + MIN_REQUEST_GAP_MS - Date.now();
      if (wait > 0) await new Promise((resolve) => setTimeout(resolve, wait));
      this.#lastRequestAt = Date.now();
      return await fn();
    });
    this.#queue = run.catch(() => {});
    return run;
  }

  async #fetchOnce(path, params, retry) {
    let url;
    if (/^https?:/i.test(path)) {
      // A link the service handed back (paging cursor). Only follow it on our own host.
      if (!path.startsWith(`${this.#baseUrl}/`)) {
        return { success: false, error: 'Cursor does not belong to the chat service' };
      }
      url = path;
    } else {
      const query = new URLSearchParams({ view: 'msnp24Equivalent', ...params }).toString();
      url = `${this.#baseUrl}${path}?${query}`;
    }
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

    try {
      const response = await fetch(url, {
        method: 'GET',
        headers: { Authentication: `skypetoken=${this.#skypeToken}`, Accept: 'application/json' },
        signal: controller.signal,
      });

      if (response.status === 401 && retry) {
        logger.info('[CHAT_SVC] Unauthorized, refreshing session');
        this.#skypeToken = null;
        const refreshed = await this.acquireSession(true);
        if (!refreshed.success) return refreshed;
        return await this.#fetchOnce(path, params, false);
      }
      if (response.status === 429) {
        return { success: false, status: 429, retryAfter: response.headers?.get?.('retry-after'), error: 'Rate limited by chat service' };
      }

      const text = await response.text();
      let data = null;
      try {
        data = text ? JSON.parse(text) : null;
      } catch {
        data = null;
      }

      if (!response.ok) {
        logger.warn('[CHAT_SVC] Request failed', { status: response.status, path });
        return { success: false, status: response.status, error: data?.message || data?.errorCode || `HTTP ${response.status}` };
      }
      return { success: true, data };
    } catch (error) {
      const message = error.name === 'AbortError' ? 'Request timed out' : error.message;
      logger.error('[CHAT_SVC] Request error', { message, path });
      return { success: false, error: message };
    } finally {
      clearTimeout(timeout);
    }
  }

  /** Recent conversations (chats, group chats, meeting chats and channels), most recent first. */
  async getConversations(pageSize = 25, cursor) {
    if (cursor) return await this.get(cursor);
    return await this.get('/v1/users/ME/conversations', { pageSize: String(Math.min(pageSize, MAX_CONVERSATION_PAGE)) });
  }

  /**
   * Messages in a conversation, newest first. Pass `cursor` (the `_metadata.backwardLink`
   * of a previous page) to continue into older history.
   */
  async getMessages(conversationId, { pageSize = 30, startTime, cursor } = {}) {
    if (cursor) return await this.get(cursor);
    const params = { pageSize: String(Math.min(pageSize, MAX_MESSAGE_PAGE)) };
    if (startTime) params.startTime = String(startTime);
    return await this.get(`/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`, params);
  }

  /**
   * All teams the user belongs to with every channel (followed or not), from the
   * chat aggregator directory the Teams client loads at startup. Read-only.
   */
  async getTeamsAndChannels() {
    if (!this.enabled) return { success: false, error: 'Graph API is disabled' };
    const session = await this.acquireSession();
    if (!session.success) return session;
    if (!this.#mainWindow?.webContents) return { success: false, error: 'Main window not initialized' };

    let aad;
    try {
      aad = await this.#mainWindow.webContents.executeJavaScript(`
        (async () => {
          const handler = window.teamsForLinuxReactHandler;
          if (!handler?.acquireToken) return { success: false, error: 'ReactHandler not available' };
          return await handler.acquireToken(${JSON.stringify(CSA_RESOURCE)}, {});
        })()
      `);
    } catch (error) {
      return { success: false, error: error.message };
    }
    if (!aad?.success || !aad.token) return { success: false, error: aad?.error || 'Failed to acquire directory token' };

    return await this.#throttled(async () => {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
      try {
        const response = await fetch(CSA_TEAMS_URL, {
          method: 'GET',
          headers: { Authorization: `Bearer ${aad.token}`, 'x-skypetoken': this.#skypeToken, Accept: 'application/json' },
          signal: controller.signal,
        });
        const text = await response.text();
        let data = null;
        try {
          data = text ? JSON.parse(text) : null;
        } catch {
          data = null;
        }
        if (!response.ok) {
          logger.warn('[CHAT_SVC] Directory request failed', { status: response.status });
          return { success: false, status: response.status, error: data?.message || `HTTP ${response.status}` };
        }
        return { success: true, data };
      } catch (error) {
        return { success: false, error: error.name === 'AbortError' ? 'Request timed out' : error.message };
      } finally {
        clearTimeout(timeout);
      }
    });
  }

  /** Thread details including members (used to name 1:1 and group chats). */
  async getThread(conversationId) {
    return await this.get(`/v1/threads/${encodeURIComponent(conversationId)}`);
  }
}

module.exports = ChatServiceClient;
