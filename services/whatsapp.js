'use strict';

const { Client, LocalAuth } = require('whatsapp-web.js');
const QRCode  = require('qrcode');
const path    = require('path');
const EventEmitter = require('events');

/**
 * Singleton WhatsApp client service.
 * Emits: 'status_change', 'qr'
 */
class WhatsAppService extends EventEmitter {
  constructor() {
    super();
    this._client  = null;
    this._status  = 'off';      // off | init | qr | ready | auth_fail | disconnected
    this._qrUrl   = null;       // base64 data-URL of current QR image
    this._info    = null;       // connected account info
  }

  get status()  { return this._status; }
  get qrUrl()   { return this._qrUrl;  }
  get info()    { return this._info;   }
  get isReady() { return this._status === 'ready'; }

  // ── Start client ────────────────────────────────────────────────────────────
  connect() {
    if (this._client) return;   // already running

    this._setStatus('init');

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
        ],
      },
      webVersionCache: {
        type: 'remote',
        remotePath: 'https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.2412.54.html',
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
      this._setStatus('auth_fail');
      this._client = null;
    });

    this._client.on('ready', async () => {
      this._info = this._client.info;
      this._setStatus('ready');
    });

    this._client.on('disconnected', (reason) => {
      console.log('[WA] Disconnected:', reason);
      this._info   = null;
      this._qrUrl  = null;
      this._client = null;
      this._setStatus('disconnected');
    });

    this._client.initialize().catch(err => {
      console.error('[WA] Init error:', err.message);
      this._client = null;
      this._setStatus('error');
    });
  }

  // ── Disconnect ───────────────────────────────────────────────────────────────
  async disconnect() {
    if (!this._client) return;
    try { await this._client.destroy(); } catch(_) {}
    this._client = null;
    this._info   = null;
    this._qrUrl  = null;
    this._setStatus('off');
  }

  // ── Send a text message ───────────────────────────────────────────────────────
  async sendText(chatId, text) {
    if (!this.isReady) throw new Error('WhatsApp is not connected. Please scan the QR code first.');
    await this._client.sendMessage(chatId, text);
  }

  // ── List all groups this account is in ───────────────────────────────────────
  async getGroups() {
    if (!this.isReady) throw new Error('WhatsApp is not connected.');
    const chats = await this._client.getChats();
    return chats
      .filter(c => c.isGroup)
      .map(c => ({ id: c.id._serialized, name: c.name, members: c.participants?.length || 0 }))
      .sort((a, b) => a.name.localeCompare(b.name));
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
