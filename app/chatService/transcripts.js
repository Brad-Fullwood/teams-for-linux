const logger = require('electron-log');

const REQUEST_TIMEOUT_MS = 60000;
const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000;
const RECORDING_TYPE = 'RichText/Media_CallRecording';
const MAX_MESSAGE_PAGES = 3;

/**
 * Read-only access to meeting transcripts.
 *
 * A recorded meeting leaves a recording message in the meeting chat whose `atp`
 * property carries the SharePoint sharing link of the video. Graph resolves that
 * link to a drive item, and SharePoint's media API (the one Stream uses) lists and
 * serves the transcript for it. Tokens for SharePoint come from the Teams web
 * client's auth provider, like the Graph token. Every call is a GET.
 */
class TranscriptClient {
  #mainWindow = null;
  #graph;
  #chat;
  #spTokens = new Map(); // origin -> { token, expiry }

  constructor({ graphApiClient, chatServiceClient }) {
    this.#graph = graphApiClient;
    this.#chat = chatServiceClient;
  }

  initialize(mainWindow) {
    this.#mainWindow = mainWindow;
  }

  /** Recording links posted in a meeting chat, newest first, de-duplicated. */
  async findRecordings(chatId) {
    const recordings = [];
    const seen = new Set();
    let cursor;
    for (let page = 0; page < MAX_MESSAGE_PAGES; page++) {
      const result = await this.#chat.getMessages(chatId, { pageSize: 200, cursor });
      if (!result.success) return result;
      const messages = result.data?.messages ?? [];
      for (const m of messages) {
        if (m.messagetype !== RECORDING_TYPE) continue;
        let url;
        try {
          url = JSON.parse(m.properties?.atp ?? '[]')[0]?.URL;
        } catch {
          url = undefined;
        }
        if (!url || seen.has(url)) continue;
        seen.add(url);
        const title = String(m.content ?? '').match(/<Title>([^<]*)<\/Title>/i)?.[1] ?? 'Recording';
        recordings.push({ url, title, when: m.composetime });
      }
      cursor = result.data?._metadata?.backwardLink;
      if (!cursor || messages.length === 0) break;
    }
    return { success: true, data: recordings };
  }

  /** Resolve a sharing link to the drive item and the SharePoint site that hosts it. */
  async resolveItem(url) {
    const encoded = `u!${Buffer.from(url).toString('base64').replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '')}`;
    const result = await this.#graph.makeRequest(`/shares/${encoded}/driveItem?$select=id,name,parentReference,sharepointIds`);
    if (!result.success) return result;
    const item = result.data ?? {};
    const siteUrl = item.sharepointIds?.siteUrl;
    const driveId = item.parentReference?.driveId;
    if (!siteUrl || !driveId || !item.id) return { success: false, error: 'Drive item is missing SharePoint identifiers' };
    return { success: true, data: { siteUrl, driveId, itemId: item.id, name: item.name } };
  }

  async #acquireSharePointToken(origin) {
    const cached = this.#spTokens.get(origin);
    if (cached && cached.expiry - Date.now() > TOKEN_EXPIRY_BUFFER_MS) return { success: true, token: cached.token };
    if (!this.#mainWindow?.webContents) return { success: false, error: 'Main window not initialized' };
    try {
      const result = await this.#mainWindow.webContents.executeJavaScript(`
        (async () => {
          const handler = window.teamsForLinuxReactHandler;
          if (!handler?.acquireToken) return { success: false, error: 'ReactHandler not available' };
          return await handler.acquireToken(${JSON.stringify(origin)}, {});
        })()
      `);
      if (!result?.success || !result.token) return { success: false, error: result?.error || 'Failed to acquire SharePoint token' };
      const expiry = result.expiry ? new Date(result.expiry).getTime() : Date.now() + 30 * 60 * 1000;
      this.#spTokens.set(origin, { token: result.token, expiry });
      return { success: true, token: result.token };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async #spGet(siteUrl, path) {
    const origin = new URL(siteUrl).origin;
    const auth = await this.#acquireSharePointToken(origin);
    if (!auth.success) return auth;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(`${siteUrl}/_api/v2.1/${path}`, {
        method: 'GET',
        headers: { Authorization: `Bearer ${auth.token}`, Accept: 'application/json' },
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
        logger.warn('[TRANSCRIPT] SharePoint request failed', { status: response.status });
        return { success: false, status: response.status, error: data?.error?.message || `HTTP ${response.status}` };
      }
      return { success: true, data };
    } catch (error) {
      return { success: false, error: error.name === 'AbortError' ? 'Request timed out' : error.message };
    } finally {
      clearTimeout(timeout);
    }
  }

  /** Transcripts attached to a recording. */
  async listTranscripts({ siteUrl, driveId, itemId }) {
    return await this.#spGet(siteUrl, `drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/media/transcripts`);
  }

  /** Transcript JSON: { entries: [{ text, speakerDisplayName, startOffset, endOffset }], events } */
  async getTranscript({ siteUrl, driveId, itemId, transcriptId }) {
    return await this.#spGet(siteUrl, `drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/media/transcripts/${encodeURIComponent(transcriptId)}/streamContent?format=json`);
  }
}

/** Render transcript entries as "[h:mm:ss] Speaker: text", merging consecutive lines by the same speaker. */
function formatTranscript(entries, maxChars = 60000) {
  const lines = [];
  let current = null;
  for (const e of entries ?? []) {
    const speaker = e.speakerDisplayName || 'Unknown';
    const text = String(e.text ?? '').trim();
    if (!text) continue;
    if (current && current.speaker === speaker) {
      current.text += ` ${text}`;
      continue;
    }
    current = { speaker, text, at: String(e.startOffset ?? '').replace(/\.\d+$/, '') };
    lines.push(current);
  }
  let out = '';
  for (const l of lines) {
    const line = `[${l.at}] ${l.speaker}: ${l.text}\n`;
    if (out.length + line.length > maxChars) {
      out += `… (truncated at ${maxChars} chars; ${lines.length} turns in total)`;
      break;
    }
    out += line;
  }
  return out.trim();
}

module.exports = { TranscriptClient, formatTranscript };
