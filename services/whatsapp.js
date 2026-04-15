'use strict';

const { Client, LocalAuth } = require('whatsapp-web.js');
const QRCode  = require('qrcode');
const path    = require('path');
const fs      = require('fs');
const EventEmitter = require('events');

/**
 * Singleton WhatsApp client service.
 * Emits: 'status_change', 'qr'
 *
 * Group loading strategy:
 *  - When 'ready' fires, a background load via client.getChats() starts immediately.
 *  - All concurrent callers of getGroups() share the same in-flight promise —
 *    no stampede, no duplicate fetches.
 *  - Results are cached for 5 minutes.  Empty results are never cached.
 *  - No artificial timeout is imposed on getChats(); on a fresh session it
 *    legitimately takes 1–3 minutes to serialise all chats on a VPS.
 */
class WhatsAppService extends EventEmitter {
  constructor() {
    super();
    this._client          = null;
    this._status          = 'off';   // off|init|qr|authenticated|ready|auth_fail|disconnected|error
    this._qrUrl           = null;    // base64 data-URL of current QR image
    this._info            = null;    // connected account info
    this._lastError       = null;    // last error message for display
    this._groupsCache     = null;    // cached group list
    this._groupsCacheAt   = 0;       // timestamp of last successful cache
    this._groupsLoadPromise = null;  // shared in-flight promise (prevents stampede)
  }

  get status()    { return this._status;    }
  get qrUrl()     { return this._qrUrl;     }
  get info()      { return this._info;      }
  get lastError() { return this._lastError; }
  get isReady()   { return this._status === 'ready'; }

  // ── Reset all group state ─────────────────────────────────────────────────────
  _resetGroups() {
    this._groupsCache       = null;
    this._groupsCacheAt     = 0;
    this._groupsLoadPromise = null;
  }

  // ── Start client ──────────────────────────────────────────────────────────────
  connect() {
    if (this._client) return;   // already running

    this._setStatus('init');
    this._lastError = null;
    this._resetGroups();

    this._client = new Client({
      authStrategy: new LocalAuth({
        dataPath: path.join(__dirname, '..', '.wwebjs_auth'),
      }),
      puppeteer: {
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-gpu',
          '--disable-extensions',
          '--disable-background-networking',
          '--disable-default-apps',
        ],
      },
    });

    this._client.on('qr', async (qrString) => {
      this._qrUrl = await QRCode.toDataURL(qrString, { width: 300 });
      this._setStatus('qr');
      this.emit('qr', this._qrUrl);
    });

    this._client.on('authenticated', () => {
      this._qrUrl = null;
      this._setStatus('authenticated');
    });

    this._client.on('auth_failure', () => {
      this._client = null;
      this._resetGroups();
      this._setStatus('auth_fail');
      this.clearSessionData();
    });

    this._client.on('ready', () => {
      this._info = this._client.info;
      this._resetGroups();
      this._setStatus('ready');
      // Kick off background group load so groups are ready when user opens the page
      this._fetchGroups().catch(() => {});
    });

    this._client.on('disconnected', async (reason) => {
      console.log('[WA] Disconnected:', reason);
      this._info  = null;
      this._qrUrl = null;
      this._resetGroups();
      const c = this._client;
      this._client = null;
      this._setStatus('disconnected');
      await this._destroyClient(c);
      this.clearSessionData();
    });

    this._client.initialize().catch(async err => {
      console.error('[WA] Init error:', err.message);
      this._lastError = err.message;
      const c = this._client;
      this._client = null;
      this._resetGroups();
      this._setStatus('error');
      await this._destroyClient(c);
    });
  }

  // ── Internal: fetch & cache groups (shared promise, no stampede) ──────────────
  _fetchGroups() {
    // Return existing in-flight promise if one is already running
    if (this._groupsLoadPromise) return this._groupsLoadPromise;

    this._groupsLoadPromise = (async () => {
      try {
        console.log('[WA] Fetching groups via getChats()…');
        const chats = await this._client.getChats();
        const groups = chats
          .filter(c => c.isGroup)
          .map(c => ({
            id:      c.id._serialized,
            name:    c.name || c.id.user || '',
            members: c.participants?.length || 0,
          }))
          .filter(c => c.id && c.name)
          .sort((a, b) => a.name.localeCompare(b.name));

        if (groups.length) {
          this._groupsCache   = groups;
          this._groupsCacheAt = Date.now();
          console.log(`[WA] Groups cached: ${groups.length} found`);
        } else {
          console.log('[WA] getChats() returned 0 groups — not caching');
        }
        return groups;
      } catch (err) {
        console.warn('[WA] Group fetch failed:', err.message);
        return [];
      } finally {
        this._groupsLoadPromise = null;
      }
    })();

    return this._groupsLoadPromise;
  }

  // ── Public: list all groups ───────────────────────────────────────────────────
  async getGroups() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');

    // Return valid cache (5-minute TTL)
    if (this._groupsCache && this._groupsCache.length &&
        (Date.now() - this._groupsCacheAt) < 300000) {
      return this._groupsCache;
    }

    // Wait for the in-flight fetch (or start one)
    return this._fetchGroups();
  }

  // ── Send a text message ───────────────────────────────────────────────────────
  async sendText(chatId, text) {
    if (!this.isReady) throw new Error('WhatsApp is not connected. Please scan the QR code first.');
    await this._client.sendMessage(chatId, text);
  }

  // ── Send message to a group by name (case-insensitive) ───────────────────────
  async sendToGroup(groupName, text) {
    if (!this.isReady) return;
    try {
      const groups = await this.getGroups();
      const group  = groups.find(g => g.name.toLowerCase() === groupName.toLowerCase());
      if (!group) {
        console.warn(`[WA] Group "${groupName}" not found`);
        return;
      }
      await new Promise(r => setTimeout(r, 1500)); // small anti-spam delay
      await this._client.sendMessage(group.id, text);
      console.log(`[WA] Sent to group "${groupName}"`);
    } catch (err) {
      console.error('[WA] sendToGroup error:', err.message);
    }
  }

  // ── List contacts ─────────────────────────────────────────────────────────────
  async getContacts() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');
    const contacts = await this._client.getContacts();
    return contacts
      .filter(c => c.isMyContact && !c.isGroup && c.name)
      .map(c => ({ id: c.id._serialized, name: c.name, number: c.number }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  // ── Disconnect ────────────────────────────────────────────────────────────────
  async disconnect() {
    const c = this._client;
    this._client  = null;
    this._info    = null;
    this._qrUrl   = null;
    this._resetGroups();
    this._setStatus('off');
    await this._destroyClient(c);
    this.clearSessionData();
  }

  // ── Delete stored session so next connect always shows a fresh QR ─────────────
  clearSessionData() {
    const sessionDir = path.join(__dirname, '..', '.wwebjs_auth');
    try {
      if (fs.existsSync(sessionDir)) {
        fs.rmSync(sessionDir, { recursive: true, force: true });
        console.log('[WA] Session data cleared');
      }
    } catch (err) {
      console.warn('[WA] Could not clear session data:', err.message);
    }
  }

  // ── Safely close the puppeteer browser and destroy the client ─────────────────
  async _destroyClient(client) {
    if (!client) return;
    try {
      if (client.pupBrowser) {
        await client.pupBrowser.close().catch(() => {});
        await new Promise(r => setTimeout(r, 800));
      }
      await client.destroy().catch(() => {});
    } catch (_) {}
  }

  _setStatus(s) {
    this._status = s;
    this.emit('status_change', s);
    console.log(`[WA] Status: ${s}`);
  }
}

module.exports = new WhatsAppService();
