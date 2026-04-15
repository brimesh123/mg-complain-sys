'use strict';

const { Client, LocalAuth } = require('whatsapp-web.js');
const QRCode  = require('qrcode');
const path    = require('path');
const fs      = require('fs');
const EventEmitter = require('events');

/**
 * Singleton WhatsApp client service.
 * Emits: 'status_change', 'qr'
 */
class WhatsAppService extends EventEmitter {
  constructor() {
    super();
    this._client    = null;
    this._status    = 'off';      // off | init | qr | ready | auth_fail | disconnected | error
    this._qrUrl     = null;       // base64 data-URL of current QR image
    this._info      = null;       // connected account info
    this._lastError = null;       // last error message for display
  }

  get status()    { return this._status;    }
  get qrUrl()     { return this._qrUrl;     }
  get info()      { return this._info;      }
  get lastError() { return this._lastError; }
  get isReady()   { return this._status === 'ready'; }

  // ── Start client ────────────────────────────────────────────────────────────
  connect() {
    if (this._client) return;   // already running

    this._setStatus('init');

    this._lastError = null;
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
      this._setStatus('auth_fail');
      this.clearSessionData();  // bad session — wipe it so QR appears fresh next time
    });

    this._client.on('ready', async () => {
      this._info        = this._client.info;
      this._groupsCache = null; // force fresh fetch after connect
      this._setStatus('ready');
    });

    this._client.on('disconnected', async (reason) => {
      console.log('[WA] Disconnected:', reason);
      this._info        = null;
      this._qrUrl       = null;
      this._groupsCache = null;
      const c = this._client;
      this._client = null;
      this._setStatus('disconnected');
      // Kill browser process so file locks are released before clearing session
      await this._destroyClient(c);
      this.clearSessionData();  // wipe auth so next connect always shows a fresh QR
    });

    this._client.initialize().catch(async err => {
      console.error('[WA] Init error:', err.message);
      this._lastError = err.message;
      const c = this._client;
      this._client = null;
      this._setStatus('error');
      await this._destroyClient(c);
    });
  }

  // ── Delete stored session files so next connect shows a fresh QR ─────────────
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

  // ── Safely destroy a client + close its browser ───────────────────────────────
  async _destroyClient(client) {
    if (!client) return;
    try {
      // Close puppeteer browser first to release file locks (e.g. first_party_sets.db)
      if (client.pupBrowser) {
        await client.pupBrowser.close().catch(() => {});
        await new Promise(r => setTimeout(r, 800)); // let OS release file handles
      }
      await client.destroy().catch(() => {});
    } catch (_) {}
  }

  // ── Disconnect ───────────────────────────────────────────────────────────────
  async disconnect() {
    const c = this._client;
    this._client = null;
    this._info   = null;
    this._qrUrl  = null;
    this._groupsCache = null;
    this._setStatus('off');
    await this._destroyClient(c);
    this.clearSessionData();  // wipe saved session so next connect shows a fresh QR
  }

  // ── Send a text message ───────────────────────────────────────────────────────
  async sendText(chatId, text) {
    if (!this.isReady) throw new Error('WhatsApp is not connected. Please scan the QR code first.');
    await this._client.sendMessage(chatId, text);
  }

  // ── Send message to a group by name (case-insensitive) ───────────────────────
  async sendToGroup(groupName, text) {
    if (!this.isReady) return;   // silently skip if WA not connected
    try {
      const groups = await this.getGroups();
      const group = groups.find(c =>
        c.name.toLowerCase() === groupName.toLowerCase());
      if (!group) {
        console.warn(`[WA] Group "${groupName}" not found`);
        return;
      }
      await new Promise(r => setTimeout(r, 1500)); // small delay — anti-spam
      await this._client.sendMessage(group.id, text);
      console.log(`[WA] Sent to group "${groupName}"`);
    } catch (err) {
      console.error('[WA] sendToGroup error:', err.message);
    }
  }

  // ── List all groups this account is in ───────────────────────────────────────
  async getGroups() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');
    // Return cached list if fetched within last 60 seconds (only non-empty results)
    const now = Date.now();
    if (this._groupsCache && this._groupsCache.length && (now - this._groupsCacheAt) < 60000) {
      return this._groupsCache;
    }
    // Read directly from WhatsApp Web's in-memory store.
    // After a fresh QR scan the store syncs asynchronously — wait up to 45s.
    const groups = await this._client.pupPage.evaluate(async () => {
      try {
        // Phase 1: wait for the Store object itself to exist (up to 15s)
        for (let i = 0; i < 30; i++) {
          const hasStore = window.Store?.Chat?.getModelsArray ||
                           window.Store?.Chat?.models;
          if (hasStore) break;
          await new Promise(r => setTimeout(r, 500));
        }
        // Phase 2: wait for chats to actually populate (up to 30 more seconds)
        let models = [];
        for (let i = 0; i < 60; i++) {
          if (window.Store?.Chat?.getModelsArray) {
            models = await window.Store.Chat.getModelsArray();
          } else if (window.Store?.Chat?.models) {
            models = Object.values(window.Store.Chat.models);
          }
          if (models.length) break;
          await new Promise(r => setTimeout(r, 500));
        }
        return models
          .filter(c => c.isGroup)
          .map(c => ({
            id:      c.id?._serialized || '',
            name:    c.name || c.formattedTitle || c.id?.user || '',
            members: c.groupMetadata?.participants?.length ||
                     c.participants?.length || 0,
          }))
          .filter(c => c.id)
          .sort((a, b) => a.name.localeCompare(b.name));
      } catch (e) {
        return [];
      }
    });
    if (!groups.length) {
      // Store still syncing — do NOT cache empty results so next request retries
      console.log('[WA] Groups not ready yet (store still syncing), returning empty');
      return [];
    }
    this._groupsCache   = groups;
    this._groupsCacheAt = now;
    return groups;
  }

  // ── List contacts (for sending to individual engineers) ───────────────────────
  async getContacts() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');
    const contacts = await this._client.getContacts();
    return contacts
      .filter(c => c.isMyContact && !c.isGroup && c.name)
      .map(c => ({ id: c.id._serialized, name: c.name, number: c.number }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  _setStatus(s) {
    this._status = s;
    this.emit('status_change', s);
    console.log(`[WA] Status: ${s}`);
  }
}

module.exports = new WhatsAppService();
