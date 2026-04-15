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
 * Group-loading design
 * ─────────────────────
 *  • On 'ready', _fetchGroups() starts in the background (fire-and-forget).
 *  • _fetchGroups() is a single shared promise — concurrent callers don't
 *    trigger duplicate fetches.
 *  • getGroups() polls the cache for up to 30 s, then returns [] so the
 *    HTTP request never hangs the browser.  The background fetch continues
 *    and the next call will hit the warm cache.
 *  • A sessionId stamp prevents a stale fetch (started before a disconnect)
 *    from writing into the new session's cache.
 */
class WhatsAppService extends EventEmitter {
  constructor() {
    super();
    this._client            = null;
    this._status            = 'off';  // off|init|qr|authenticated|ready|auth_fail|disconnected|error
    this._qrUrl             = null;
    this._info              = null;
    this._lastError         = null;
    this._groupsCache       = null;
    this._groupsCacheAt     = 0;
    this._groupsLoadPromise = null;   // shared in-flight fetch promise
    this._sessionId         = 0;      // incremented on every connect()
  }

  get status()    { return this._status;    }
  get qrUrl()     { return this._qrUrl;     }
  get info()      { return this._info;      }
  get lastError() { return this._lastError; }
  get isReady()   { return this._status === 'ready'; }

  // ── Connect ───────────────────────────────────────────────────────────────────
  connect() {
    if (this._client) return;

    this._sessionId++;
    this._setStatus('init');
    this._lastError = null;
    this._clearGroupState();

    this._client = new Client({
      authStrategy: new LocalAuth({ dataPath: path.join(__dirname, '..', '.wwebjs_auth') }),
      puppeteer: {
        headless: true,
        args: [
          '--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage',
          '--disable-gpu', '--disable-extensions', '--disable-background-networking',
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
      this._clearGroupState();
      this._setStatus('auth_fail');
      this.clearSessionData();
    });

    this._client.on('ready', () => {
      this._info = this._client.info;
      this._clearGroupState();
      this._setStatus('ready');
      // Kick off background load — don't await
      this._fetchGroups().catch(() => {});
    });

    this._client.on('disconnected', async (reason) => {
      console.log('[WA] Disconnected:', reason);
      this._info  = null;
      this._qrUrl = null;
      this._clearGroupState();
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
      this._clearGroupState();
      this._setStatus('error');
      await this._destroyClient(c);
    });
  }

  // ── Disconnect ────────────────────────────────────────────────────────────────
  async disconnect() {
    const c = this._client;
    this._client = null;
    this._info   = null;
    this._qrUrl  = null;
    this._clearGroupState();
    this._setStatus('off');
    await this._destroyClient(c);
    this.clearSessionData();
  }

  // ── Public: list all groups ───────────────────────────────────────────────────
  async getGroups() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');

    // Cache hit (5-minute TTL)
    if (this._groupsCache && this._groupsCache.length &&
        (Date.now() - this._groupsCacheAt) < 300000) {
      return this._groupsCache;
    }

    // Ensure a background fetch is running, but don't block on it
    this._fetchGroups().catch(() => {});

    // Poll the cache for up to 30 s so the HTTP request doesn't hang the browser
    for (let i = 0; i < 60; i++) {
      await new Promise(r => setTimeout(r, 500));
      if (!this.isReady) throw new Error('WhatsApp disconnected while waiting for groups.');
      if (this._groupsCache && this._groupsCache.length) return this._groupsCache;
    }

    // Still not ready — return empty; client should show a "retry" button
    console.log('[WA] Groups not ready within 30 s — returning empty, client should retry');
    return [];
  }

  // ── Internal: single shared fetch (prevents stampede) ────────────────────────
  _fetchGroups() {
    if (this._groupsLoadPromise) return this._groupsLoadPromise;

    const sid    = this._sessionId;
    const client = this._client;

    this._groupsLoadPromise = (async () => {
      try {
        console.log('[WA] Fetching groups via getChats()…');

        // Hard 5-minute cap — if getChats() hangs past this, something is broken
        const chats = await Promise.race([
          client.getChats(),
          new Promise((_, rej) =>
            setTimeout(() => rej(new Error('getChats timed out (5 min)')), 300000)
          ),
        ]);

        // Discard result if session changed while we were waiting
        if (this._sessionId !== sid) {
          console.log('[WA] Session changed during fetch — discarding stale result');
          return [];
        }

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
          console.log(`[WA] Groups loaded: ${groups.length}`);
        } else {
          console.log('[WA] getChats() returned 0 groups — not caching');
        }
        return groups;

      } catch (err) {
        console.warn('[WA] Group fetch failed:', err.message);
        return [];
      } finally {
        // Only clear the promise if this session is still active
        if (this._sessionId === sid) this._groupsLoadPromise = null;
      }
    })();

    return this._groupsLoadPromise;
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
      if (!group) { console.warn(`[WA] Group "${groupName}" not found`); return; }
      await new Promise(r => setTimeout(r, 1500));
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

  // ── Delete stored session so next connect always shows a fresh QR ─────────────
  clearSessionData() {
    const dir = path.join(__dirname, '..', '.wwebjs_auth');
    try {
      if (fs.existsSync(dir)) {
        fs.rmSync(dir, { recursive: true, force: true });
        console.log('[WA] Session data cleared');
      }
    } catch (err) {
      console.warn('[WA] Could not clear session data:', err.message);
    }
  }

  // ── Helpers ───────────────────────────────────────────────────────────────────
  _clearGroupState() {
    this._groupsCache       = null;
    this._groupsCacheAt     = 0;
    this._groupsLoadPromise = null;
  }

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
