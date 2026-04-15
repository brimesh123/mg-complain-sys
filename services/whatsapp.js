'use strict';

const { Client, LocalAuth } = require('whatsapp-web.js');
const QRCode  = require('qrcode');
const path    = require('path');
const fs      = require('fs');
const EventEmitter = require('events');

/**
 * Singleton WhatsApp client service.
 * Emits: 'status_change', 'qr', 'groups_ready'
 *
 * Group loading strategy (fastest first):
 *   1. Serve from 5-minute cache immediately if warm.
 *   2. Read window.Store.Chat models directly in the browser page —
 *      instant when the store is populated, no network round-trip.
 *      Group IDs always end with @g.us — reliable cross-version filter.
 *   3. If the store is still empty after 10 s (fresh session syncing),
 *      fall back to client.getChats() — one shared promise so concurrent
 *      callers don't stampede.
 *   4. getGroups() always returns within 30 s so HTTP requests never hang.
 *   5. When cache is finally populated, 'groups_ready' is emitted so the
 *      frontend can refresh the group list without user interaction.
 */
class WhatsAppService extends EventEmitter {
  constructor() {
    super();
    this._client            = null;
    this._status            = 'off';
    this._qrUrl             = null;
    this._info              = null;
    this._lastError         = null;
    this._groupsCache       = null;
    this._groupsCacheAt     = 0;
    this._groupsLoadPromise = null;
    this._sessionId         = 0;      // incremented on every connect() call
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
      // Start background group warm-up immediately on ready
      this._warmGroupCache().catch(() => {});
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

  // ── Public: get groups ────────────────────────────────────────────────────────
  async getGroups() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');

    // 1. Warm cache hit
    if (this._groupsCache && this._groupsCache.length &&
        (Date.now() - this._groupsCacheAt) < 300000) {
      return this._groupsCache;
    }

    // 2. Try the in-page store immediately (fast path)
    const storeGroups = await this._getGroupsFromStore();
    if (storeGroups && storeGroups.length) {
      this._setCache(storeGroups);
      return storeGroups;
    }

    // 3. Ensure background warm-up is running
    if (!this._groupsLoadPromise) {
      this._warmGroupCache().catch(() => {});
    }

    // 4. Poll cache for up to 30 s — HTTP request never hangs the browser
    for (let i = 0; i < 60; i++) {
      await new Promise(r => setTimeout(r, 500));
      if (!this.isReady) throw new Error('WhatsApp disconnected while loading groups.');
      if (this._groupsCache && this._groupsCache.length) return this._groupsCache;
    }

    // 5. Still empty — return [] so client can show "retry" button
    console.log('[WA] Groups not ready within 30 s — returning empty');
    return [];
  }

  // ── Send a text message ───────────────────────────────────────────────────────
  async sendText(chatId, text) {
    if (!this.isReady) throw new Error('WhatsApp is not connected. Please scan the QR code first.');
    await this._client.sendMessage(chatId, text);
  }

  // ── Send to a group by name ───────────────────────────────────────────────────
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

  // ── Clear stored session so next connect shows a fresh QR ────────────────────
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

  // ── Private: read groups directly from WhatsApp Web's in-memory store ─────────
  // This is instant (<100 ms) when the store is populated.
  // WhatsApp group IDs always end with @g.us — reliable across all WW.js versions.
  async _getGroupsFromStore() {
    if (!this._client?.pupPage) return null;
    try {
      return await this._client.pupPage.evaluate(async () => {
        const store = window.Store?.Chat;
        if (!store) return null;

        let models = [];
        try {
          // getModelsArray may be async in newer WW.js versions
          if (typeof store.getModelsArray === 'function') {
            models = await store.getModelsArray();
          } else if (typeof store.getModels === 'function') {
            models = store.getModels() || [];
          } else if (store.models) {
            models = Object.values(store.models);
          }
        } catch (_) { return null; }

        const groups = (models || [])
          .filter(m => m?.id?._serialized?.endsWith('@g.us'))
          .map(m => ({
            id:      m.id._serialized,
            name:    m.name || m.formattedTitle || m.id?.user || '',
            members: m.groupMetadata?.participants?.length || m.participants?.length || 0,
          }))
          .filter(g => g.id && g.name)
          .sort((a, b) => a.name.localeCompare(b.name));

        return groups.length ? groups : null;
      });
    } catch (_) {
      return null;
    }
  }

  // ── Private: warm group cache — store first, then getChats() fallback ─────────
  async _warmGroupCache() {
    const sid = this._sessionId;

    // Phase 1: poll the in-memory store every 500 ms for up to 10 s
    console.log('[WA] Warming group cache (store poll)…');
    for (let i = 0; i < 20; i++) {
      if (this._sessionId !== sid || !this.isReady) return;
      const groups = await this._getGroupsFromStore();
      if (groups && groups.length) {
        this._setCache(groups);
        console.log(`[WA] Groups from store: ${groups.length}`);
        this.emit('groups_ready', groups);
        return;
      }
      await new Promise(r => setTimeout(r, 500));
    }

    // Phase 2: store still empty — fall back to getChats() (one shared promise)
    if (this._groupsLoadPromise || this._sessionId !== sid) return;
    console.log('[WA] Store empty after 10 s, falling back to getChats()…');
    const client = this._client;

    this._groupsLoadPromise = (async () => {
      try {
        // 5-minute hard cap — if getChats() hangs past this something is broken
        const chats = await Promise.race([
          client.getChats(),
          new Promise((_, rej) =>
            setTimeout(() => rej(new Error('getChats timed out (5 min)')), 300000)
          ),
        ]);

        if (this._sessionId !== sid) return []; // session changed, discard

        const groups = chats
          .filter(c => c.isGroup)
          .map(c => ({
            id:      c.id._serialized,
            name:    c.name || c.id.user || '',
            members: c.participants?.length || 0,
          }))
          .filter(g => g.id && g.name)
          .sort((a, b) => a.name.localeCompare(b.name));

        if (groups.length) {
          this._setCache(groups);
          console.log(`[WA] Groups from getChats(): ${groups.length}`);
          this.emit('groups_ready', groups);
        } else {
          console.log('[WA] getChats() returned 0 groups — account may have no groups');
        }
        return groups;
      } catch (err) {
        console.warn('[WA] getChats() failed:', err.message);
        return [];
      } finally {
        if (this._sessionId === sid) this._groupsLoadPromise = null;
      }
    })();

    return this._groupsLoadPromise;
  }

  // ── Private helpers ───────────────────────────────────────────────────────────
  _setCache(groups) {
    this._groupsCache   = groups;
    this._groupsCacheAt = Date.now();
  }

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
