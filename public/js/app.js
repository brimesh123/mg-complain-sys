/* ─── Mogal Complaint System ── Frontend SPA ───────────────────────────────── */
'use strict';

// ─── State ────────────────────────────────────────────────────────────────────
const S = {
  page: 'dashboard',
  engineers: [],
  complaintTypes: [],
  areas: [],
  customers:  { data: [], total: 0, page: 1, search: '', area: '', status: '' },
  complaints: { data: [], total: 0, page: 1, search: '', status: '', engineer_id: '', date_from: '', date_to: '' },
  selectedCustomer: null,
};

const WA   = { status: 'off', target_id: null, target_name: null, account: null, hasQr: false, lastError: null };
const SYNC = { lastExport: null, xlsxExists: false };

// ─── API ──────────────────────────────────────────────────────────────────────
const API = {
  async req(method, url, body) {
    const opts = { method, headers: { 'Content-Type': 'application/json' } };
    if (body !== undefined) opts.body = JSON.stringify(body);
    const r = await fetch(url, opts);
    let json;
    try { json = await r.json(); }
    catch (_) {
      throw new Error(
        r.status === 404 ? `Endpoint not found. Please restart the server (${url}).`
        : `Server error ${r.status} — please restart start.bat`
      );
    }
    if (!json.success) throw new Error(json.message || 'Request failed');
    return json;
  },
  get:    url       => API.req('GET',    url),
  post:   (url, b)  => API.req('POST',   url, b),
  put:    (url, b)  => API.req('PUT',    url, b),
  delete: (url, b)  => API.req('DELETE', url, b),
};

// ─── Utilities ────────────────────────────────────────────────────────────────
function esc(s) {
  if (s == null) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
                  .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' });
}
function fmtDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return fmtDate(iso) + ' ' + d.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit', hour12:true });
}
function statusBadge(s) {
  const m = { Open:'info', 'In Progress':'warning', Resolved:'success', Closed:'gray' };
  return `<span class="badge badge-${m[s]||'gray'}">${esc(s)}</span>`;
}
function priorityBadge(p) {
  const m = { Low:'success', Normal:'info', High:'warning', Urgent:'danger' };
  return `<span class="badge badge-${m[p]||'gray'}">${esc(p)}</span>`;
}
function connBadge(s) {
  return s === 'ON'
    ? `<span class="badge badge-success">ON</span>`
    : `<span class="badge badge-danger">OFF</span>`;
}
function warrantyBadge(installDate) {
  if (!installDate) return `<span class="badge badge-gray">No install date</span>`;
  const warrantyEnd = new Date(new Date(installDate).getTime() + 365 * 24 * 60 * 60 * 1000);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (today <= warrantyEnd) {
    const days = Math.ceil((warrantyEnd - today) / 86400000);
    return `<span class="badge badge-success">In Warranty &middot; ${days}d left</span>`;
  }
  const days = Math.floor((today - warrantyEnd) / 86400000);
  return `<span class="badge badge-danger">Out of Warranty &middot; ${days}d ago</span>`;
}
function animateCounter(el, target, duration = 900) {
  if (!target) { el.textContent = '0'; return; }
  const start = performance.now();
  const tick = now => {
    const t = Math.min((now - start) / duration, 1);
    const eased = 1 - Math.pow(1 - t, 3);
    el.textContent = Math.round(eased * target);
    if (t < 1) requestAnimationFrame(tick);
    else el.textContent = target;
  };
  requestAnimationFrame(tick);
}
function debounce(fn, ms) {
  let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function toast(msg, type = 'info') {
  const icons = { success:'check-circle', danger:'x-circle', warning:'alert-triangle', info:'info' };
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `
    <div class="toast-icon"><i data-lucide="${icons[type]||'info'}"></i></div>
    <div class="toast-msg">${esc(msg)}</div>
    <button class="toast-close icon-btn" aria-label="Dismiss"><i data-lucide="x"></i></button>`;
  el.querySelector('.toast-close').onclick = () => dismissToast(el);
  document.getElementById('toastContainer').prepend(el);
  lucide.createIcons({ nodes: [el] });
  setTimeout(() => dismissToast(el), 4500);
}
function dismissToast(el) {
  el.style.opacity = '0';
  el.style.transform = 'translateX(20px)';
  setTimeout(() => el.remove(), 250);
}

// ─── Confirm Dialog ───────────────────────────────────────────────────────────
function confirmDialog(title, msg, opts = {}) {
  return new Promise(resolve => {
    const uid = '_c' + Date.now();
    const div = document.createElement('div');
    div.className = 'confirm-overlay';
    div.innerHTML = `
      <div class="confirm-box" role="alertdialog" aria-modal="true">
        <div class="confirm-icon ${opts.danger ? 'ci-danger' : 'ci-warning'}">
          <i data-lucide="${opts.danger ? 'trash-2' : 'alert-triangle'}"></i>
        </div>
        <h3 class="confirm-title">${esc(title)}</h3>
        <p  class="confirm-msg">${esc(msg)}</p>
        <div class="confirm-btns">
          <button class="btn btn-secondary" id="${uid}_no">Cancel</button>
          <button class="btn ${opts.danger ? 'btn-danger' : 'btn-primary'}" id="${uid}_yes">
            ${esc(opts.confirmText || 'Confirm')}
          </button>
        </div>
      </div>`;
    document.body.appendChild(div);
    lucide.createIcons({ nodes: [div] });
    const done = v => { div.classList.add('cd-out'); setTimeout(() => div.remove(), 200); resolve(v); };
    document.getElementById(`${uid}_yes`).onclick = () => done(true);
    document.getElementById(`${uid}_no`).onclick  = () => done(false);
    div.addEventListener('click', e => { if (e.target === div) done(false); });
  });
}

// ─── Modal ────────────────────────────────────────────────────────────────────
const Modal = {
  open(title, html, opts = {}) {
    const box = document.getElementById('modalBox');
    box.className = 'modal-box' + (opts.wide ? ' wide' : '');
    document.getElementById('modalTitle').textContent = title || '';
    document.getElementById('modalBody').innerHTML = html || '';
    document.getElementById('modalBackdrop').classList.add('is-open');
    document.getElementById('modal').classList.add('is-open');
    lucide.createIcons({ nodes: [document.getElementById('modal')] });
    if (opts.onOpen) opts.onOpen();
  },
  close() {
    document.getElementById('modal').classList.remove('is-open');
    document.getElementById('modalBackdrop').classList.remove('is-open');
    document.getElementById('modalBody').innerHTML = '';
    document.getElementById('modalTitle').textContent = '';
  },
};
document.getElementById('modalClose').onclick    = Modal.close;
document.getElementById('modalBackdrop').onclick = Modal.close;

// ─── Clock ────────────────────────────────────────────────────────────────────
function startClock() {
  const el = document.getElementById('topbarTime');
  const tick = () => {
    const n = new Date();
    el.textContent =
      n.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' }) +
      '  ' + n.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit', second:'2-digit', hour12:true });
  };
  tick(); setInterval(tick, 1000);
}

// ─── WhatsApp helpers ─────────────────────────────────────────────────────────
async function waPoll() {
  try {
    const { data } = await API.get('/api/whatsapp/status');
    Object.assign(WA, data);
    const dot = document.getElementById('navWaDot');
    if (dot) dot.classList.toggle('visible', data.status === 'ready');
  } catch (_) {}
}

// ─── Router ───────────────────────────────────────────────────────────────────
const PAGES = {
  dashboard, customers, newComplaint, complaints,
  engineers, whatsapp: whatsappPage, logs: activityLogs, reports: reportsPage,
};
const TITLES = {
  dashboard: 'Dashboard', customers: 'Customers',
  'new-complaint': 'New Complaint', complaints: 'Complaints',
  engineers: 'Engineers', whatsapp: 'WhatsApp Integration',
  logs: 'Activity Log', reports: 'Reports',
};

function navigate(page) {
  S.page = page;
  document.querySelectorAll('.nav-item').forEach(el =>
    el.classList.toggle('active', el.dataset.page === page));
  document.getElementById('breadcrumb').textContent = TITLES[page] || page;
  const content = document.getElementById('pageContent');
  content.innerHTML = `<div class="page-loading"><div class="spinner"></div></div>`;
  const fn = PAGES[page === 'new-complaint' ? 'newComplaint' : page];
  if (fn) fn(content); else content.innerHTML = '<p>Page not found.</p>';
}

const sidebar   = document.getElementById('sidebar');
const mainWrap  = document.getElementById('mainWrapper');
const sbBackdrop = document.getElementById('sidebarBackdrop');

const closeSidebar = () => {
  sidebar.classList.remove('open');
  sidebar.classList.add('collapsed');
  mainWrap.classList.add('expanded');
  sbBackdrop.classList.remove('visible');
};
const openSidebar = () => {
  sidebar.classList.add('open');
  sidebar.classList.remove('collapsed');
  mainWrap.classList.remove('expanded');
  sbBackdrop.classList.add('visible');
};

document.getElementById('sidebarToggle').onclick  = closeSidebar;
document.getElementById('sidebarOpenBtn').onclick = openSidebar;
sbBackdrop.addEventListener('click', closeSidebar);

// Nav item click — navigate + auto-close on mobile
document.querySelectorAll('.nav-item').forEach(el =>
  el.addEventListener('click', e => {
    e.preventDefault();
    navigate(el.dataset.page);
    if (window.innerWidth <= 768) closeSidebar();
  }));

// ─── Preload ──────────────────────────────────────────────────────────────────
async function preload() {
  const [eng, types, areas] = await Promise.all([
    API.get('/api/engineers'),
    API.get('/api/complaint-types'),
    API.get('/api/areas'),
  ]);
  S.engineers      = eng.data;
  S.complaintTypes = types.data;
  S.areas          = areas.data;
}
function engineerOptions(sel = '') {
  return S.engineers.filter(e => e.active)
    .map(e => `<option value="${e.id}"${e.id == sel ? ' selected':''}>${esc(e.name)}</option>`).join('');
}
function complaintTypeOptions(sel = '') {
  return S.complaintTypes
    .map(t => `<option value="${esc(t.name)}"${t.name === sel ? ' selected':''}>${esc(t.name)}</option>`).join('');
}
function areaOptions(sel = '') {
  return S.areas
    .map(a => `<option value="${esc(a)}"${a === sel ? ' selected':''}>${esc(a)}</option>`).join('');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
async function dashboard(el) {
  try {
    const { data } = await API.get('/api/dashboard/stats');
    const { stats, recent } = data;
    el.innerHTML = `
      <div class="page-header">
        <div>
          <div class="page-header-title">Dashboard</div>
          <div class="page-header-sub">Overview of your complaint management system</div>
        </div>
        <button class="btn btn-primary" onclick="navigate('new-complaint')">
          <i data-lucide="plus"></i> New Complaint
        </button>
      </div>

      <div class="stats-grid">
        ${[
          ['total_customers',   'Total Customers',   'users',          '#2563eb','#eff6ff'],
          ['active_customers',  'Active (ON)',        'wifi',           '#059669','#d1fae5'],
          ['inactive_customers','Inactive (OFF)',     'wifi-off',       '#dc2626','#fee2e2'],
          ['open_complaints',   'Open Complaints',    'alert-circle',   '#d97706','#fef3c7'],
          ['total_complaints',  'Total Complaints',   'clipboard-list', '#0369a1','#e0f2fe'],
        ].map(([k, label, icon, color, bg]) => `
          <div class="stat-card">
            <div class="stat-icon" style="background:${bg};color:${color}">
              <i data-lucide="${icon}"></i>
            </div>
            <div class="stat-info">
              <div class="stat-value" data-counter="${stats[k] ?? 0}">0</div>
              <div class="stat-label">${label}</div>
            </div>
          </div>`).join('')}
      </div>

      <div class="card">
        <div class="card-header">
          <span class="card-title">Recent Complaints</span>
          <button class="btn btn-sm btn-ghost" onclick="navigate('complaints')">
            View All <i data-lucide="arrow-right"></i>
          </button>
        </div>
        <div class="table-wrapper">
          ${!recent.length ? `
            <div class="empty-state">
              <i data-lucide="clipboard"></i>
              <h3>No complaints yet</h3>
              <p>Log your first complaint to get started.</p>
            </div>` : `
          <table>
            <thead><tr>
              <th>No.</th><th>Customer</th><th>Complaint Type</th><th>Status</th><th>Date</th>
            </tr></thead>
            <tbody>
              ${recent.map(c => `
                <tr class="row-link" onclick="openComplaintDetail(${c.id})">
                  <td><span class="mono-tag">${esc(c.complaint_no)}</span></td>
                  <td>
                    <div class="fw-600">${esc(c.new_party_name)}</div>
                    <div class="td-muted">NSN: ${c.nsn}${c.area ? ' · ' + esc(c.area) : ''}</div>
                  </td>
                  <td>${esc(c.complaint_type)}</td>
                  <td>${statusBadge(c.status)}</td>
                  <td class="td-muted">${fmtDate(c.created_at)}</td>
                </tr>`).join('')}
            </tbody>
          </table>`}
        </div>
      </div>`;
    lucide.createIcons({ nodes: [el] });
    el.querySelectorAll('[data-counter]').forEach(num =>
      animateCounter(num, parseInt(num.dataset.counter)));
  } catch(e) {
    el.innerHTML = `<div class="empty-state"><p class="text-danger">${esc(e.message)}</p></div>`;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// NEW COMPLAINT — Bulk Queue
// ═══════════════════════════════════════════════════════════════════════════════
const COMPLAINT_TYPES = [
  'Pressure','Over detection','Image problem','Light setting','Vacuum',
  'Pixel set / Model','Bottle change','Dundi Marvi','Size setting','GT card',
  'Open / Close','Dip out','Kuwait Nakhvi','Shifting','Focus levu','Cleaning',
  'Liquid vadhare aave','Blure','Machine issue','bubble nathi udta','Formate',
  'Camera hang','PC hang','New update','ASV update','Extra pc programme',
  'programme install','ball valve kholvo','rod nakhvo','error problem',
  'file problem','jump','hitter change','sensor change','leakage',
];

let _ncQueue      = [];   // [{ customer, types:[], description }]
let _ncCustomer   = null;
let _ncSelTypes   = [];   // currently selected types (before adding to queue)

// ── Complaint type multi-select ───────────────────────────────────────────────
function _nctRender() {
  const chips = document.getElementById('nctChips');
  const ph    = document.getElementById('nctPlaceholder');
  const list  = document.getElementById('nctList');
  if (!chips) return;

  // Update chips display
  const existing = chips.querySelectorAll('.nct-chip');
  existing.forEach(c => c.remove());
  if (_ncSelTypes.length) {
    ph && (ph.style.display = 'none');
    _ncSelTypes.forEach(t => {
      const chip = document.createElement('span');
      chip.className = 'nct-chip';
      chip.innerHTML = `${esc(t)} <button class="nct-chip-x" data-t="${esc(t)}" tabindex="-1">×</button>`;
      chip.querySelector('.nct-chip-x').onclick = (e) => {
        e.stopPropagation();
        _ncSelTypes = _ncSelTypes.filter(x => x !== t);
        _nctRender();
        _nctPopulateList();
      };
      chips.insertBefore(chip, ph || null);
    });
  } else {
    ph && (ph.style.display = '');
  }

  // Update list checkboxes if open
  if (list && list.childElementCount) _nctPopulateList();
}

function _nctPopulateList(filter) {
  const list = document.getElementById('nctList');
  if (!list) return;
  const q = (filter ?? (document.getElementById('nctSearch')?.value || '')).toLowerCase();
  const filtered = COMPLAINT_TYPES.filter(t => !q || t.toLowerCase().includes(q));
  list.innerHTML = filtered.length
    ? filtered.map(t => `
        <div class="nct-item ${_ncSelTypes.includes(t) ? 'selected' : ''}" data-t="${esc(t)}">
          <span class="nct-item-check">${_ncSelTypes.includes(t) ? '✓' : ''}</span>
          ${esc(t)}
        </div>`).join('')
    : `<div class="nct-empty">No match</div>`;
  list.querySelectorAll('.nct-item').forEach(item => {
    item.onclick = () => {
      const t = item.dataset.t;
      if (_ncSelTypes.includes(t)) _ncSelTypes = _ncSelTypes.filter(x => x !== t);
      else _ncSelTypes.push(t);
      _nctRender();
      _nctPopulateList();
    };
  });
}

function _nctOpen() {
  const dd = document.getElementById('nctDropdown');
  const arrow = document.getElementById('nctArrow');
  if (!dd) return;
  dd.style.display = '';
  arrow && arrow.classList.add('open');
  _nctPopulateList('');
  setTimeout(() => document.getElementById('nctSearch')?.focus(), 30);
}

function _nctClose() {
  const dd = document.getElementById('nctDropdown');
  const arrow = document.getElementById('nctArrow');
  if (!dd) return;
  dd.style.display = 'none';
  arrow && arrow.classList.remove('open');
}

function _nctInit() {
  const ctrl = document.getElementById('nctControl');
  const wrap = document.getElementById('nctWrap');
  const search = document.getElementById('nctSearch');
  if (!ctrl) return;

  ctrl.addEventListener('click', () => {
    const dd = document.getElementById('nctDropdown');
    if (dd && dd.style.display === 'none') _nctOpen(); else _nctClose();
  });
  ctrl.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); _nctOpen(); }
    if (e.key === 'Escape') _nctClose();
  });
  if (search) {
    search.addEventListener('input', () => _nctPopulateList(search.value));
    search.addEventListener('keydown', e => { if (e.key === 'Escape') _nctClose(); });
  }
  document.addEventListener('click', function onDocClick(e) {
    if (wrap && !wrap.contains(e.target)) { _nctClose(); }
  });
}

function newComplaint(el) {
  _ncQueue    = [];
  _ncCustomer = null;
  _ncSelTypes = [];

  el.innerHTML = `
    <div class="page-header" style="margin-bottom:16px">
      <div>
        <div class="page-header-title">Log Complaints</div>
        <div class="page-header-sub">Add multiple, then submit all at once to WhatsApp</div>
      </div>
      <span class="nc-queue-pill" id="ncPill" style="display:none">
        <i data-lucide="layers"></i> <span id="ncPillCount">0</span> in queue
      </span>
    </div>

    <div class="nc-layout">
      <!-- Left: Search + Form -->
      <div class="nc-left card">
        <div class="nc-section-head"><i data-lucide="user-search"></i> Find Customer</div>
        <div style="display:flex;gap:8px;align-items:flex-start">
          <div class="lookup-search" id="lookupSearch" style="flex:1;position:relative">
            <i data-lucide="search" class="lookup-search-icon"></i>
            <input id="lookupInput" type="text" class="form-control lookup-input"
              placeholder="Serial No., Name or OSN…" autocomplete="off"/>
            <div class="lookup-spinner" id="lookupSpinner"></div>
            <div class="lookup-results" id="lookupResults"></div>
          </div>
          <button class="btn btn-secondary" style="flex-shrink:0;height:42px" onclick="openCustomerForm()" title="Add new party">
            <i data-lucide="user-plus"></i>
          </button>
        </div>

        <div class="customer-card" id="customerCard"></div>

        <div class="nc-complaint-area" id="ncComplaintArea" style="display:none">
          <div class="nc-section-head"><i data-lucide="tag"></i> Complaint Type <span class="nc-req-badge">required</span></div>

          <!-- Multi-select types dropdown -->
          <div class="nct-wrap" id="nctWrap">
            <div class="nct-control" id="nctControl" tabindex="0">
              <div class="nct-chips" id="nctChips">
                <span class="nct-placeholder" id="nctPlaceholder">Search type… (select one or more)</span>
              </div>
              <i data-lucide="chevron-down" class="nct-arrow" id="nctArrow"></i>
            </div>
            <div class="nct-dropdown" id="nctDropdown" style="display:none">
              <div class="nct-search-wrap">
                <i data-lucide="search" class="nct-search-icon"></i>
                <input class="nct-search" id="nctSearch" type="text" placeholder="Search complaint types…" autocomplete="off"/>
              </div>
              <div class="nct-list" id="nctList"></div>
            </div>
          </div>

          <div class="nc-section-head" style="margin-top:14px"><i data-lucide="message-square"></i> Additional Notes <span class="nc-opt-badge">optional</span></div>
          <textarea class="form-control nc-textarea" id="ncRemarks" rows="2"
            placeholder="Extra details… (Enter to add, Shift+Enter for newline)"></textarea>
          <button class="btn btn-primary nc-add-btn" id="ncAddBtn">
            <i data-lucide="plus-circle"></i> Add to Queue
          </button>
        </div>
      </div>

      <!-- Right: Queue -->
      <div class="nc-right card">
        <div class="nc-queue-hdr">
          <div class="nc-queue-title"><i data-lucide="layers"></i> Queue</div>
          <span class="badge badge-info nc-queue-count" id="ncQueueCount">0</span>
        </div>

        <div class="nc-queue-body" id="ncQueueBody">
          <div class="nc-queue-empty">
            <i data-lucide="inbox"></i>
            <p>Find a customer, write the complaint,<br>then click <b>Add to Queue</b></p>
          </div>
        </div>

        <div class="nc-queue-actions" id="ncQueueActions" style="display:none">
          <button class="btn btn-success" id="ncSubmitBtn">
            <i data-lucide="send"></i> <span id="ncSubmitLabel">Submit &amp; Send to WhatsApp</span>
          </button>
          <button class="btn btn-secondary" id="ncCopyAllBtn" style="display:none">
            <i data-lucide="copy"></i> Copy Combined Message
          </button>
        </div>

        <!-- Success State -->
        <div class="nc-success" id="ncSuccess" style="display:none">
          <div class="nc-success-icon"><i data-lucide="check-circle-2"></i></div>
          <div class="nc-success-title" id="ncSuccessTitle">Complaints Logged!</div>
          <div class="nc-success-cmps" id="ncSuccessCmps"></div>
          <div class="nc-section-head" style="margin:14px 0 6px;width:100%">
            <i data-lucide="message-circle"></i> Combined Message
          </div>
          <div class="nc-success-msg-wrap" id="ncSuccessMsg"></div>
          <div class="nc-wa-status-wrap" id="ncWaStatus" style="display:none"></div>
          <div class="nc-success-btns" id="ncSuccessBtns"></div>
        </div>
      </div>
    </div>`;

  lucide.createIcons({ nodes: [el] });
  initLookup();
  _nctInit();

  document.getElementById('ncAddBtn').addEventListener('click', _ncAddToQueue);
  document.getElementById('ncRemarks').addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); _ncAddToQueue(); }
  });
  document.getElementById('ncSubmitBtn').addEventListener('click', _ncSubmitAll);
  _ncRefreshSubmitBtn();
}

function initLookup() {
  const input   = document.getElementById('lookupInput');
  const results = document.getElementById('lookupResults');
  const spinner = document.getElementById('lookupSpinner');

  const doSearch = debounce(async q => {
    if (!q) { results.style.display = 'none'; return; }
    spinner.style.display = 'block';
    try {
      const { data } = await API.get(`/api/customers/lookup?q=${encodeURIComponent(q)}`);
      if (!data.length) {
        results.innerHTML = `
          <div class="lookup-result-item" style="flex-direction:column;align-items:flex-start;gap:8px">
            <span class="text-muted" style="font-size:13px">No customer found for "<b>${esc(q)}</b>"</span>
            <button class="btn btn-sm btn-primary" onclick="openCustomerForm();document.getElementById('lookupResults').style.display='none'">
              <i data-lucide="user-plus"></i> Add "${esc(q)}" as New Party
            </button>
          </div>`;
      } else {
        results.innerHTML = data.map(c => `
          <div class="lookup-result-item" data-id="${c.id}" tabindex="0" role="option">
            <span class="lri-nsn">NSN ${c.nsn}</span>
            <div>
              <div class="lri-name">${esc(c.new_party_name || c.party_name)}</div>
              <div class="lri-meta">OSN: ${esc(c.osn||'—')} · ${esc(c.area||'')} · ${connBadge(c.status)}</div>
            </div>
          </div>`).join('');
        results.querySelectorAll('[data-id]').forEach(item =>
          item.onclick = () => selectCustomer(data.find(c => c.id == item.dataset.id)));
      }
      lucide.createIcons({ nodes: [results] });
      results.style.display = 'block';
    } catch(e) {
      results.innerHTML = `<div class="lookup-result-item text-danger">${esc(e.message)}</div>`;
      results.style.display = 'block';
    } finally { spinner.style.display = 'none'; }
  }, 280);

  input.addEventListener('input', e => doSearch(e.target.value.trim()));
  document.addEventListener('click', e => {
    if (!document.getElementById('lookupSearch')?.contains(e.target))
      results.style.display = 'none';
  });
}

function selectCustomer(c) {
  if (!c) return;
  _ncCustomer = c;
  document.getElementById('lookupResults').style.display = 'none';
  document.getElementById('lookupInput').value = `${c.new_party_name || c.party_name} (NSN: ${c.nsn})`;

  const card = document.getElementById('customerCard');
  card.innerHTML = `
    <div class="cc-header">
      <div>
        <div class="cc-name">${esc(c.new_party_name || c.party_name)}</div>
        <div class="cc-sn">NSN: <b>${c.nsn}</b> &nbsp;|&nbsp; OSN: <b>${esc(c.osn||'—')}</b></div>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        ${connBadge(c.status)}
        <button class="icon-btn" title="Clear selection" onclick="clearSelection()">
          <i data-lucide="x"></i>
        </button>
      </div>
    </div>
    <div class="cc-grid">
      <div class="cc-field"><span class="cc-field-label">Contact</span><span class="cc-field-value">${esc(c.contact_no||'—')}</span></div>
      <div class="cc-field"><span class="cc-field-label">Area</span><span class="cc-field-value">${esc(c.area||'—')}</span></div>
      <div class="cc-field" style="grid-column:1/-1"><span class="cc-field-label">Address</span><span class="cc-field-value">${esc(c.address||'—')}</span></div>
    </div>`;
  card.classList.add('visible');
  lucide.createIcons({ nodes: [card] });

  const area = document.getElementById('ncComplaintArea');
  if (area) {
    area.style.display = '';
    _ncSelTypes = [];
    _nctRender();
    setTimeout(() => _nctOpen(), 80);
  }
}

window.clearSelection = () => {
  _ncCustomer = null;
  _ncSelTypes = [];
  const card = document.getElementById('customerCard');
  if (card) card.classList.remove('visible');
  const area = document.getElementById('ncComplaintArea');
  if (area) area.style.display = 'none';
  _nctClose();
  const input = document.getElementById('lookupInput');
  if (input) { input.value = ''; input.focus(); }
};

// ── Queue helpers ─────────────────────────────────────────────────────────────
function _ncRefreshSubmitBtn() {
  const btn  = document.getElementById('ncSubmitBtn');
  const copy = document.getElementById('ncCopyAllBtn');
  if (!btn) return;
  if (WA.status === 'ready' && WA.target_id) {
    document.getElementById('ncSubmitLabel').textContent = `Submit & Send to ${WA.target_name || 'WhatsApp'}`;
    copy && (copy.style.display = 'none');
  } else {
    document.getElementById('ncSubmitLabel').textContent = 'Submit Complaints';
    copy && (copy.style.display = '');
    copy && (copy.onclick = _ncCopyAll);
  }
}

function _ncAddToQueue() {
  if (!_ncCustomer) { toast('Select a customer first', 'warning'); return; }
  if (!_ncSelTypes.length) {
    toast('Select at least one complaint type', 'warning');
    _nctOpen();
    return;
  }
  const remarks = (document.getElementById('ncRemarks')?.value || '').trim();

  _ncQueue.push({ customer: { ..._ncCustomer }, types: [..._ncSelTypes], description: remarks });
  _ncRenderQueue();

  // Clear for next entry
  document.getElementById('ncRemarks').value = '';
  _ncSelTypes = [];
  _nctClose();
  clearSelection();
  toast(`Added (${_ncQueue.length} in queue)`, 'success');
}

function _ncRenderQueue() {
  const body    = document.getElementById('ncQueueBody');
  const actions = document.getElementById('ncQueueActions');
  const count   = document.getElementById('ncQueueCount');
  const pill    = document.getElementById('ncPill');
  if (!body) return;

  const n = _ncQueue.length;
  if (count) count.textContent = n;
  if (pill)  pill.style.display = n ? '' : 'none';
  if (document.getElementById('ncPillCount')) document.getElementById('ncPillCount').textContent = n;

  if (!n) {
    body.innerHTML = `<div class="nc-queue-empty">
      <i data-lucide="inbox"></i>
      <p>Find a customer, write the complaint,<br>then click <b>Add to Queue</b></p>
    </div>`;
    if (actions) actions.style.display = 'none';
    lucide.createIcons({ nodes: [body] });
    return;
  }

  body.innerHTML = _ncQueue.map((item, i) => `
    <div class="nc-qi">
      <div class="nc-qi-num">${i + 1}</div>
      <div class="nc-qi-body">
        <div class="nc-qi-name">${esc(item.customer.new_party_name || item.customer.party_name)}</div>
        <div class="nc-qi-meta">SN: <b>${item.customer.nsn}</b>${item.customer.area ? ' · ' + esc(item.customer.area) : ''}</div>
        <div class="nc-qi-types">${(item.types||[]).map(t => `<span class="nc-qi-type-chip">${esc(t)}</span>`).join('')}</div>
        ${item.description ? `<div class="nc-qi-desc">${esc(item.description)}</div>` : ''}
      </div>
      <button class="nc-qi-remove" onclick="_ncRemove(${i})" title="Remove"><i data-lucide="x"></i></button>
    </div>`).join('');

  if (actions) actions.style.display = '';
  _ncRefreshSubmitBtn();
  lucide.createIcons({ nodes: [body] });
}

window._ncRemove = (idx) => { _ncQueue.splice(idx, 1); _ncRenderQueue(); };

function _ncCopyAll() {
  const msg = _ncBuildMsg(_ncQueue.map(q => ({
    nsn: q.customer.nsn, new_party_name: q.customer.new_party_name,
    address: q.customer.address, types: q.types, remarks: q.description,
  })));
  navigator.clipboard.writeText(msg).then(() => toast('Copied!', 'success'));
}

function _ncBuildMsg(results) {
  const now = new Date();
  const timeStr = now.toLocaleString('en-IN', {
    day:'2-digit', month:'short', year:'numeric',
    hour:'2-digit', minute:'2-digit', hour12:true,
  });
  const lines = [`Complaints — ${timeStr}`, ''];
  results.forEach((r, i) => {
    lines.push(`${i + 1}. SN ${r.nsn} — ${r.new_party_name || ''}`);
    if (r.address) lines.push(`Address: ${r.address}`);
    if (r.types && r.types.length) lines.push(`Type: ${r.types.join(', ')}`);
    if (r.remarks) lines.push(`Note: ${r.remarks}`);
    if (i < results.length - 1) lines.push('');
  });
  return lines.join('\n');
}

async function _ncSubmitAll() {
  if (!_ncQueue.length) { toast('Queue is empty — add at least one complaint', 'warning'); return; }
  const btn = document.getElementById('ncSubmitBtn');
  btn.disabled = true;
  btn.innerHTML = '<i data-lucide="loader"></i> Submitting…';
  lucide.createIcons({ nodes: [btn] });

  try {
    const queueSnapshot = _ncQueue.map(q => ({ types: q.types, description: q.description }));
    const { data: results } = await API.post('/api/complaints/bulk', {
      complaints: _ncQueue.map(q => ({
        customer_id: q.customer.id,
        complaint_type: (q.types||[]).join(', ') || 'General',
        remarks: [q.types&&q.types.length ? q.types.join(', ') : null, q.description||null].filter(Boolean).join(' — ') || null,
      })),
    });
    await preload();
    // Merge types/description back into results for WA message
    results.forEach((r, i) => {
      if (queueSnapshot[i]) { r.types = queueSnapshot[i].types; r.remarks = queueSnapshot[i].description; }
    });
    _ncShowSuccess(results);
  } catch(err) {
    toast(err.message, 'danger');
    btn.disabled = false;
    _ncRefreshSubmitBtn();
    lucide.createIcons({ nodes: [btn] });
  }
}

function _ncShowSuccess(results) {
  // Hide queue, show success panel
  const qBody   = document.getElementById('ncQueueBody');
  const qAct    = document.getElementById('ncQueueActions');
  const success = document.getElementById('ncSuccess');
  if (!success) return;
  if (qBody)  qBody.style.display   = 'none';
  if (qAct)   qAct.style.display    = 'none';
  success.style.display = '';

  document.getElementById('ncSuccessTitle').textContent =
    `${results.length} Complaint${results.length !== 1 ? 's' : ''} Logged!`;

  document.getElementById('ncSuccessCmps').innerHTML =
    results.map(r => `<span class="mono-tag" style="font-size:13px;font-weight:700">${esc(r.complaint_no)}</span>`).join('');

  const msg = _ncBuildMsg(results);
  document.getElementById('ncSuccessMsg').textContent = msg;

  // Build action buttons
  const btnsEl = document.getElementById('ncSuccessBtns');
  const waReady = WA.status === 'ready' && WA.target_id;

  btnsEl.innerHTML = `
    <button class="btn btn-secondary" id="ncCopyFinalBtn"><i data-lucide="copy"></i> Copy Message</button>
    <button class="btn btn-primary" onclick="navigate('new-complaint')"><i data-lucide="plus"></i> Log More</button>
    <button class="btn btn-ghost" onclick="navigate('complaints')"><i data-lucide="list"></i> View All</button>`;
  lucide.createIcons({ nodes: [btnsEl] });

  document.getElementById('ncCopyFinalBtn').onclick = () =>
    navigator.clipboard.writeText(msg).then(() => toast('Copied!', 'success'));

  // WA status indicator (not a button — auto-sends silently)
  const waStatus = document.getElementById('ncWaStatus');
  if (waReady && waStatus) {
    waStatus.style.display = '';
    waStatus.innerHTML = `<span class="nc-wa-status sending"><i data-lucide="loader"></i> Sending to ${esc(WA.target_name || 'group')}…</span>`;
    lucide.createIcons({ nodes: [waStatus] });
    API.post('/api/whatsapp/send', { chat_id: WA.target_id, message: msg })
      .then(() => {
        waStatus.innerHTML = `<span class="nc-wa-status sent"><i data-lucide="check-circle"></i> Sent to ${esc(WA.target_name || 'group')}</span>`;
        lucide.createIcons({ nodes: [waStatus] });
        toast('Sent to WhatsApp group!', 'success');
      })
      .catch(e => {
        waStatus.innerHTML = `<span class="nc-wa-status error"><i data-lucide="alert-circle"></i> ${esc(e.message)}</span>`;
        lucide.createIcons({ nodes: [waStatus] });
      });
  } else if (waStatus) {
    waStatus.style.display = 'none';
  }

  lucide.createIcons({ nodes: [success] });
  _ncQueue = [];
  const pill = document.getElementById('ncPill');
  if (pill) pill.style.display = 'none';
}

// ═══════════════════════════════════════════════════════════════════════════════
// CUSTOMERS
// ═══════════════════════════════════════════════════════════════════════════════
async function customers(el) {
  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">Customers</div>
        <div class="page-header-sub">Manage customer database</div>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-secondary" id="syncFromXlsxBtn" title="Import changes from sync.xlsx on disk">
          <i data-lucide="refresh-cw"></i> Sync from XLSX
        </button>
        <button class="btn btn-secondary" id="importXlsxBtn" title="Import customers from any Excel file">
          <i data-lucide="file-down"></i> Import XLSX
        </button>
        <a class="btn btn-secondary" href="/api/export/customers" download title="Download customers as Excel">
          <i data-lucide="file-up"></i> Export XLSX
        </a>
        <button class="btn btn-primary" id="addCustomerBtn">
          <i data-lucide="user-plus"></i> Add Customer
        </button>
      </div>
    </div>
    <input type="file" id="xlsxFileInput" accept=".xlsx,.xls" style="display:none" />
    <div class="card">
      <div class="card-header" style="flex-wrap:wrap;gap:10px">
        <div class="filter-bar" style="flex:1">
          <div class="search-box" style="flex:1;min-width:220px">
            <i data-lucide="search" class="search-icon"></i>
            <input id="custSearch" type="text" class="form-control" placeholder="Search name, NSN, OSN…" value="${esc(S.customers.search)}" />
          </div>
          <select class="form-control" id="custArea" style="width:150px">
            <option value="">All Areas</option>${areaOptions(S.customers.area)}
          </select>
          <select class="form-control" id="custStatus" style="width:120px">
            <option value="">All Status</option>
            <option value="ON"${S.customers.status==='ON'?' selected':''}>ON</option>
            <option value="OFF"${S.customers.status==='OFF'?' selected':''}>OFF</option>
          </select>
        </div>
      </div>
      <div id="custBulkBar" style="display:none;padding:10px 16px;background:var(--primary-light);border-bottom:1px solid var(--primary-border);display:none;align-items:center;gap:12px">
        <span id="custBulkCount" style="font-size:13px;font-weight:600;color:var(--primary)"></span>
        <button class="btn btn-sm btn-danger" id="custBulkDeleteBtn">
          <i data-lucide="trash-2"></i> Delete Selected
        </button>
        <button class="btn btn-sm btn-secondary" id="custBulkClearBtn">Clear Selection</button>
      </div>
      <div id="custTableWrap"><div class="page-loading"><div class="spinner"></div></div></div>
    </div>`;
  lucide.createIcons({ nodes: [el] });

  const load = async () => {
    const wrap = document.getElementById('custTableWrap');
    wrap.innerHTML = `<div class="page-loading"><div class="spinner"></div></div>`;
    try {
      const { data, total } = await API.get(
        `/api/customers?search=${encodeURIComponent(S.customers.search)}&area=${encodeURIComponent(S.customers.area)}&status=${S.customers.status}&page=${S.customers.page}&limit=50`
      );
      S.customers.data = data; S.customers.total = total;
      renderCustomerTable(wrap, data, total);
    } catch(e) {
      wrap.innerHTML = `<div class="empty-state"><p class="text-danger">${esc(e.message)}</p></div>`;
    }
  };

  const search = debounce(() => { S.customers.page = 1; load(); }, 320);
  document.getElementById('custSearch').addEventListener('input',  e => { S.customers.search = e.target.value; search(); });
  document.getElementById('custArea').addEventListener('change',   e => { S.customers.area   = e.target.value; S.customers.page=1; load(); });
  document.getElementById('custStatus').addEventListener('change', e => { S.customers.status = e.target.value; S.customers.page=1; load(); });
  document.getElementById('addCustomerBtn').onclick = () => openCustomerForm();

  // Sync from XLSX (reads sync.xlsx on disk → upserts into DB)
  document.getElementById('syncFromXlsxBtn').onclick = async () => {
    const btn = document.getElementById('syncFromXlsxBtn');
    btn.disabled = true;
    btn.innerHTML = '<i data-lucide="loader"></i> Syncing…';
    lucide.createIcons({ nodes: [btn] });
    try {
      const { data: r } = await API.post('/api/sync/import');
      const total = (r.custChanged || 0) + (r.engChanged || 0);
      toast(
        total > 0
          ? `Synced from XLSX: ${r.custChanged} customers, ${r.engChanged} engineers updated`
          : 'sync.xlsx is already up to date — no changes found',
        total > 0 ? 'success' : 'info'
      );
      await syncPoll();
      if (r.custChanged > 0) { S.customers.page = 1; load(); }
    } catch(e) { toast(e.message, 'danger'); }
    finally {
      btn.disabled = false;
      btn.innerHTML = '<i data-lucide="refresh-cw"></i> Sync from XLSX';
      lucide.createIcons({ nodes: [btn] });
    }
  };

  // XLSX Import
  const fileInput = document.getElementById('xlsxFileInput');
  document.getElementById('importXlsxBtn').onclick = () => fileInput.click();
  fileInput.onchange = async () => {
    const file = fileInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async ev => {
      const b64 = ev.target.result.split(',')[1];
      const btn = document.getElementById('importXlsxBtn');
      btn.disabled = true; btn.innerHTML = '<i data-lucide="loader"></i> Importing…';
      lucide.createIcons({ nodes: [btn] });
      try {
        const { data: r } = await API.post('/api/import/customers', { data: b64 });
        toast(`Import done: ${r.added} added, ${r.updated} updated, ${r.skipped} skipped`, r.added+r.updated > 0 ? 'success' : 'info');
        await preload();
        S.customers.page = 1;
        load();
      } catch(e) { toast(e.message, 'danger'); }
      finally {
        btn.disabled = false;
        btn.innerHTML = '<i data-lucide="file-down"></i> Import XLSX';
        lucide.createIcons({ nodes: [btn] });
        fileInput.value = '';
      }
    };
    reader.readAsDataURL(file);
  };

  load();
}

function renderCustomerTable(wrap, data, total) {
  if (!data.length) {
    wrap.innerHTML = `<div class="empty-state"><i data-lucide="users"></i><h3>No customers found</h3><p>Adjust filters or add a customer.</p></div>`;
    lucide.createIcons({ nodes: [wrap] }); return;
  }
  wrap.innerHTML = `
    <div class="table-wrapper">
      <table>
        <thead><tr>
          <th style="width:36px"><input type="checkbox" id="custSelectAll" title="Select all" /></th>
          <th>NSN</th><th>OSN</th><th>Party Name</th><th>Contact</th>
          <th>Area</th><th>Status</th><th>Complaints</th><th>Actions</th>
        </tr></thead>
        <tbody>
          ${data.map(c => `
            <tr data-id="${c.id}">
              <td><input type="checkbox" class="cust-chk" value="${c.id}" /></td>
              <td><span class="mono-tag">${c.nsn}</span></td>
              <td class="td-muted">${esc(c.osn||'—')}</td>
              <td>
                <div class="fw-600">${esc(c.new_party_name||c.party_name)}</div>
                ${c.new_party_name && c.new_party_name!==c.party_name
                  ? `<div class="td-muted" style="font-size:11.5px">Orig: ${esc(c.party_name)}</div>`:''}
              </td>
              <td>${esc(c.contact_no||'—')}</td>
              <td>${esc(c.area||'—')}</td>
              <td>${connBadge(c.status)}</td>
              <td>
                <span class="badge ${c.open_count>0?'badge-warning':'badge-gray'}">
                  ${c.open_count} open / ${c.complaint_count}
                </span>
              </td>
              <td>
                <div class="gap-10">
                  <button class="btn btn-sm btn-ghost" title="New Complaint" onclick="quickComplaint(${c.id})">
                    <i data-lucide="plus-circle"></i>
                  </button>
                  <button class="btn btn-sm btn-ghost" title="Edit" onclick="openCustomerForm(${c.id})">
                    <i data-lucide="pencil"></i>
                  </button>
                  <button class="btn btn-sm btn-ghost text-danger" title="Delete" onclick="deleteCustomer(${c.id},'${esc(c.new_party_name||c.party_name)}')">
                    <i data-lucide="trash-2"></i>
                  </button>
                </div>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>
    <div class="pagination">
      <span>Showing ${data.length} of ${total} customers</span>
      <div class="pagination-btns">
        <button class="btn btn-sm btn-secondary" ${S.customers.page<=1?'disabled':''} onclick="custPage(${S.customers.page-1})">
          <i data-lucide="chevron-left"></i> Prev
        </button>
        <button class="btn btn-sm btn-secondary" ${S.customers.page*50>=total?'disabled':''} onclick="custPage(${S.customers.page+1})">
          Next <i data-lucide="chevron-right"></i>
        </button>
      </div>
    </div>`;
  lucide.createIcons({ nodes: [wrap] });

  // ── Bulk selection logic ─────────────────────────────────────────────────────
  const bulkBar   = document.getElementById('custBulkBar');
  const bulkCount = document.getElementById('custBulkCount');
  const selectAll = document.getElementById('custSelectAll');

  const updateBulkBar = () => {
    const checked = document.querySelectorAll('.cust-chk:checked');
    if (checked.length > 0) {
      bulkBar.style.display = 'flex';
      bulkCount.textContent = `${checked.length} customer${checked.length>1?'s':''} selected`;
    } else {
      bulkBar.style.display = 'none';
    }
    selectAll.indeterminate = checked.length > 0 && checked.length < data.length;
    selectAll.checked = checked.length === data.length;
  };

  selectAll.addEventListener('change', () => {
    document.querySelectorAll('.cust-chk').forEach(cb => cb.checked = selectAll.checked);
    updateBulkBar();
  });
  document.querySelectorAll('.cust-chk').forEach(cb =>
    cb.addEventListener('change', updateBulkBar));

  document.getElementById('custBulkClearBtn').onclick = () => {
    document.querySelectorAll('.cust-chk').forEach(cb => cb.checked = false);
    selectAll.checked = false;
    updateBulkBar();
  };

  document.getElementById('custBulkDeleteBtn').onclick = async () => {
    const ids = [...document.querySelectorAll('.cust-chk:checked')].map(cb => parseInt(cb.value));
    if (!ids.length) return;
    const ok = await confirmDialog(
      'Delete Customers',
      `Delete ${ids.length} selected customer${ids.length>1?'s':''}? This cannot be undone.`,
      { danger: true, confirmText: `Delete ${ids.length}` }
    );
    if (!ok) return;
    try {
      await API.delete('/api/customers/bulk', { ids });
      toast(`${ids.length} customer${ids.length>1?'s':''} deleted`, 'success');
      S.customers.page = 1;
      // reload page
      navigate('customers');
    } catch(e) { toast(e.message, 'danger'); }
  };
}

window.custPage = p => { S.customers.page = p; navigate('customers'); };

window.quickComplaint = customerId => {
  navigate('new-complaint');
  setTimeout(async () => {
    try {
      const { data } = await API.get(`/api/customers/${customerId}`);
      selectCustomer(data);
    } catch(e) { toast(e.message, 'danger'); }
  }, 120);
};

function openCustomerForm(id = null) {
  const isEdit   = id !== null;
  const existing = isEdit ? S.customers.data.find(c => c.id === id) : null;
  const v = f => existing ? esc(existing[f] || '') : '';

  Modal.open(isEdit ? 'Edit Customer' : 'Add New Party', `
    <form id="custForm" style="display:flex;flex-direction:column;gap:16px">
      <div class="form-grid form-grid-2">
        <div class="form-group">
          <label class="form-label">New Serial No. (NSN) <span class="req">*</span></label>
          <input class="form-control" name="nsn" type="number" placeholder="e.g. 1025" value="${v('nsn')}" required />
        </div>
        <div class="form-group">
          <label class="form-label">Old Serial No. (OSN)</label>
          <input class="form-control" name="osn" type="text" placeholder="e.g. ZA" value="${v('osn')}" />
        </div>
      </div>
      <div class="form-group">
        <label class="form-label">Original Party Name <span class="req">*</span></label>
        <input class="form-control" name="party_name" type="text" placeholder="Original party name at install time" value="${v('party_name')}" required />
      </div>
      <div class="form-group">
        <label class="form-label">Current Party Name <span class="form-hint-inline">(if transferred)</span></label>
        <input class="form-control" name="new_party_name" type="text" placeholder="Leave blank if same as original" value="${v('new_party_name')}" />
      </div>
      <div class="form-grid form-grid-2">
        <div class="form-group">
          <label class="form-label">Contact No.</label>
          <input class="form-control" name="contact_no" type="text" placeholder="9XXXXXXXXX" value="${v('contact_no')}" />
        </div>
        <div class="form-group">
          <label class="form-label">Area</label>
          <input class="form-control" name="area" type="text" list="areaList" placeholder="e.g. Katargam" value="${v('area')}" />
          <datalist id="areaList">${S.areas.map(a=>`<option value="${esc(a)}">`).join('')}</datalist>
        </div>
      </div>
      <div class="form-group">
        <label class="form-label">Address</label>
        <textarea class="form-control" name="address" rows="2" placeholder="Full address…">${v('address')}</textarea>
      </div>
      <div class="form-grid form-grid-2">
        <div class="form-group">
          <label class="form-label">Install Date</label>
          <input class="form-control" name="install_date" type="date"
            value="${existing?.install_date ? existing.install_date.split('T')[0] : ''}" />
          ${isEdit ? `<div style="margin-top:6px">${warrantyBadge(existing?.install_date)}</div>` : ''}
        </div>
        <div class="form-group">
          <label class="form-label">Connection Status</label>
          <select class="form-control" name="status">
            <option value="ON"${(!existing||existing.status==='ON')?' selected':''}>ON</option>
            <option value="OFF"${existing?.status==='OFF'?' selected':''}>OFF</option>
          </select>
        </div>
      </div>
      <div class="form-group">
        <label class="form-label">Notes</label>
        <textarea class="form-control" name="notes" rows="2" placeholder="Internal notes…">${v('notes')}</textarea>
      </div>
      <div class="modal-footer" style="padding:0;border:none">
        <button type="button" class="btn btn-secondary" onclick="Modal.close()">Cancel</button>
        <button type="submit" class="btn btn-primary">${isEdit ? 'Save Changes' : 'Add Customer'}</button>
      </div>
    </form>`, { wide: true });

  document.getElementById('custForm').onsubmit = async e => {
    e.preventDefault();
    const fd  = new FormData(e.target);
    const btn = e.target.querySelector('[type=submit]');
    btn.disabled = true;
    try {
      if (isEdit) await API.put(`/api/customers/${id}`, Object.fromEntries(fd));
      else        await API.post('/api/customers', Object.fromEntries(fd));
      toast(isEdit ? 'Customer updated' : 'Customer added', 'success');
      Modal.close();
      await preload();
      navigate('customers');
    } catch(err) { toast(err.message, 'danger'); }
    finally { btn.disabled = false; }
  };
}

/** @type {any} */ (window).deleteCustomer = async (id, name) => {
  const ok = await confirmDialog('Delete Customer', `Delete "${name}"? This cannot be undone.`, { danger: true, confirmText: 'Delete' });
  if (!ok) return;
  try { await API.delete(`/api/customers/${id}`); toast('Customer deleted', 'success'); navigate('customers'); }
  catch(e) { toast(e.message, 'danger'); }
};

// ═══════════════════════════════════════════════════════════════════════════════
// COMPLAINTS
// ═══════════════════════════════════════════════════════════════════════════════
async function complaints(el) {
  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">Complaints</div>
        <div class="page-header-sub">View and manage all complaints</div>
      </div>
      <div style="display:flex;gap:8px">
        <a class="btn btn-secondary" href="/api/export/complaints" download title="Export to Excel">
          <i data-lucide="file-up"></i> Export XLSX
        </a>
        <button class="btn btn-primary" onclick="navigate('new-complaint')">
          <i data-lucide="plus"></i> New Complaint
        </button>
      </div>
    </div>
    <div class="period-stats-row" id="cmpPeriodStats">
      <div class="page-loading" style="min-height:72px"><div class="spinner" style="width:24px;height:24px;border-width:2px"></div></div>
    </div>

    <div class="card">
      <div class="card-header" style="flex-wrap:wrap;gap:10px">
        <div class="filter-bar" style="flex:1">
          <div class="search-box" style="flex:1;min-width:200px">
            <i data-lucide="search" class="search-icon"></i>
            <input id="cmpSearch" type="text" class="form-control" placeholder="Search complaint no., customer…" value="${esc(S.complaints.search)}" />
          </div>
          <select class="form-control" id="cmpStatus" style="width:140px">
            <option value="">All Status</option>
            ${['Open','In Progress','Resolved','Closed'].map(s =>
              `<option value="${s}"${S.complaints.status===s?' selected':''}>${s}</option>`).join('')}
          </select>
          <select class="form-control" id="cmpEngineer" style="width:150px">
            <option value="">All Engineers</option>
            ${S.engineers.map(e=>`<option value="${e.id}"${S.complaints.engineer_id==e.id?' selected':''}>${esc(e.name)}</option>`).join('')}
          </select>
          <input type="date" class="form-control" id="cmpFrom" style="width:140px" value="${S.complaints.date_from}" title="From date" />
          <input type="date" class="form-control" id="cmpTo"   style="width:140px" value="${S.complaints.date_to}"   title="To date" />
        </div>
      </div>
      <div id="cmpBulkBar" style="display:none;padding:10px 16px;background:var(--primary-light);border-bottom:1px solid var(--primary-border);align-items:center;gap:12px">
        <span id="cmpBulkCount" style="font-size:13px;font-weight:600;color:var(--primary)"></span>
        <button class="btn btn-sm btn-danger" id="cmpBulkDeleteBtn">
          <i data-lucide="trash-2"></i> Delete Selected
        </button>
        <button class="btn btn-sm btn-secondary" id="cmpBulkClearBtn">Clear Selection</button>
      </div>
      <div id="cmpTableWrap"><div class="page-loading"><div class="spinner"></div></div></div>
    </div>`;
  lucide.createIcons({ nodes: [el] });

  const load = async () => {
    const wrap = document.getElementById('cmpTableWrap');
    wrap.innerHTML = `<div class="page-loading"><div class="spinner"></div></div>`;
    try {
      const p = new URLSearchParams({
        search: S.complaints.search, status: S.complaints.status,
        engineer_id: S.complaints.engineer_id, date_from: S.complaints.date_from,
        date_to: S.complaints.date_to, page: S.complaints.page, limit: 50,
      });
      const { data, total } = await API.get(`/api/complaints?${p}`);
      S.complaints.data = data; S.complaints.total = total;
      renderComplaintTable(wrap, data, total);
    } catch(e) {
      wrap.innerHTML = `<div class="empty-state"><p class="text-danger">${esc(e.message)}</p></div>`;
    }
  };

  const search = debounce(() => { S.complaints.page=1; load(); }, 320);
  document.getElementById('cmpSearch').addEventListener('input',    e => { S.complaints.search      = e.target.value; search(); });
  document.getElementById('cmpStatus').addEventListener('change',   e => { S.complaints.status      = e.target.value; S.complaints.page=1; load(); });
  document.getElementById('cmpEngineer').addEventListener('change', e => { S.complaints.engineer_id = e.target.value; S.complaints.page=1; load(); });
  document.getElementById('cmpFrom').addEventListener('change',     e => { S.complaints.date_from   = e.target.value; S.complaints.page=1; load(); });
  document.getElementById('cmpTo').addEventListener('change',       e => { S.complaints.date_to     = e.target.value; S.complaints.page=1; load(); });
  load();

  // Period stats cards
  API.get('/api/complaints/stats').then(({ data: s }) => {
    const ps = document.getElementById('cmpPeriodStats');
    if (!ps) return;
    ps.innerHTML = [
      ['today',      "Today",       'calendar-check',  '#7c3aed','#ede9fe'],
      ['yesterday',  "Yesterday",   'calendar-minus',  '#0369a1','#e0f2fe'],
      ['this_month', "This Month",  'calendar',        '#059669','#d1fae5'],
      ['total',      "Total",       'layers',          '#d97706','#fef3c7'],
    ].map(([k, label, icon, color, bg]) => `
      <div class="period-stat-card">
        <div class="ps-icon" style="background:${bg};color:${color}"><i data-lucide="${icon}"></i></div>
        <div class="ps-body">
          <div class="ps-value" data-counter="${s[k] ?? 0}">0</div>
          <div class="ps-label">${label}</div>
        </div>
      </div>`).join('');
    lucide.createIcons({ nodes: [ps] });
    ps.querySelectorAll('[data-counter]').forEach(el => animateCounter(el, parseInt(el.dataset.counter)));
  }).catch(() => {
    const ps = document.getElementById('cmpPeriodStats');
    if (ps) ps.remove();
  });
}

function renderComplaintTable(wrap, data, total) {
  if (!data.length) {
    wrap.innerHTML = `<div class="empty-state"><i data-lucide="clipboard-list"></i><h3>No complaints found</h3><p>Adjust filters or log a new complaint.</p></div>`;
    lucide.createIcons({ nodes: [wrap] }); return;
  }
  wrap.innerHTML = `
    <div class="table-wrapper">
      <table>
        <thead><tr>
          <th style="width:36px"><input type="checkbox" id="cmpSelectAll" title="Select all" /></th>
          <th>Complaint No</th><th>Customer</th><th>Type</th>
          <th>Engineer</th><th>Priority</th><th>Status</th><th>Date</th><th>Actions</th>
        </tr></thead>
        <tbody>
          ${data.map(c => `
            <tr>
              <td><input type="checkbox" class="cmp-chk" value="${c.id}" /></td>
              <td><span class="mono-tag row-link" onclick="openComplaintDetail(${c.id})">${esc(c.complaint_no)}</span></td>
              <td>
                <div class="fw-600">${esc(c.new_party_name)}</div>
                <div class="td-muted">NSN: ${c.nsn} · ${esc(c.area||'')}</div>
              </td>
              <td>${esc(c.complaint_type)}</td>
              <td>${esc(c.engineer_name||'—')}</td>
              <td>${priorityBadge(c.priority)}</td>
              <td>${statusBadge(c.status)}</td>
              <td class="td-muted">${fmtDateTime(c.created_at)}</td>
              <td>
                <div class="gap-10">
                  <button class="btn btn-sm btn-ghost" title="View" onclick="openComplaintDetail(${c.id})"><i data-lucide="eye"></i></button>
                  <button class="btn btn-sm btn-ghost" title="Edit" onclick="openComplaintEdit(${c.id})"><i data-lucide="pencil"></i></button>
                  <button class="btn btn-sm btn-ghost text-danger" title="Delete" onclick="deleteComplaint(${c.id},'${esc(c.complaint_no)}')"><i data-lucide="trash-2"></i></button>
                </div>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>
    <div class="pagination">
      <span>Showing ${data.length} of ${total} complaints</span>
      <div class="pagination-btns">
        <button class="btn btn-sm btn-secondary" ${S.complaints.page<=1?'disabled':''} onclick="cmpPage(${S.complaints.page-1})">
          <i data-lucide="chevron-left"></i> Prev
        </button>
        <button class="btn btn-sm btn-secondary" ${S.complaints.page*50>=total?'disabled':''} onclick="cmpPage(${S.complaints.page+1})">
          Next <i data-lucide="chevron-right"></i>
        </button>
      </div>
    </div>`;
  lucide.createIcons({ nodes: [wrap] });

  // ── Bulk selection logic ─────────────────────────────────────────────────────
  const bulkBar   = document.getElementById('cmpBulkBar');
  const bulkCount = document.getElementById('cmpBulkCount');
  const selectAll = document.getElementById('cmpSelectAll');

  const updateBulkBar = () => {
    const checked = document.querySelectorAll('.cmp-chk:checked');
    bulkBar.style.display = checked.length > 0 ? 'flex' : 'none';
    if (checked.length) bulkCount.textContent = `${checked.length} complaint${checked.length>1?'s':''} selected`;
    selectAll.indeterminate = checked.length > 0 && checked.length < data.length;
    selectAll.checked = checked.length === data.length;
  };

  selectAll.addEventListener('change', () => {
    document.querySelectorAll('.cmp-chk').forEach(cb => cb.checked = selectAll.checked);
    updateBulkBar();
  });
  document.querySelectorAll('.cmp-chk').forEach(cb =>
    cb.addEventListener('change', updateBulkBar));

  document.getElementById('cmpBulkClearBtn').onclick = () => {
    document.querySelectorAll('.cmp-chk').forEach(cb => cb.checked = false);
    selectAll.checked = false;
    updateBulkBar();
  };

  document.getElementById('cmpBulkDeleteBtn').onclick = async () => {
    const ids = [...document.querySelectorAll('.cmp-chk:checked')].map(cb => parseInt(cb.value));
    if (!ids.length) return;
    const ok = await confirmDialog(
      'Delete Complaints',
      `Permanently delete ${ids.length} selected complaint${ids.length>1?'s':''}? This cannot be undone.`,
      { danger: true, confirmText: `Delete ${ids.length}` }
    );
    if (!ok) return;
    try {
      await API.delete('/api/complaints/bulk', { ids });
      toast(`${ids.length} complaint${ids.length>1?'s':''} deleted`, 'success');
      S.complaints.page = 1;
      navigate('complaints');
    } catch(e) { toast(e.message, 'danger'); }
  };
}

window.cmpPage = p => { S.complaints.page = p; navigate('complaints'); };

window.openComplaintDetail = async id => {
  try {
    const { data: c } = await API.get(`/api/complaints/${id}`);
    Modal.open(`Complaint ${c.complaint_no}`, `
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:4px">
        ${statusBadge(c.status)} ${priorityBadge(c.priority)} ${connBadge(c.connection_status)}
      </div>
      <div class="divider"></div>
      <div class="section-label">Customer</div>
      <div class="detail-grid" style="margin-bottom:14px">
        <div class="detail-row"><span class="detail-label">Name</span><span class="detail-value fw-600">${esc(c.new_party_name)}</span></div>
        <div class="detail-row"><span class="detail-label">NSN / OSN</span><span class="detail-value">${c.nsn} / ${esc(c.osn||'—')}</span></div>
        <div class="detail-row"><span class="detail-label">Contact</span><span class="detail-value">${esc(c.contact_no||'—')}</span></div>
        <div class="detail-row"><span class="detail-label">Area</span><span class="detail-value">${esc(c.area||'—')}</span></div>
        <div class="detail-row" style="grid-column:1/-1"><span class="detail-label">Address</span><span class="detail-value">${esc(c.address||'—')}</span></div>
      </div>
      <div class="divider"></div>
      <div class="section-label">Complaint</div>
      <div class="detail-grid">
        <div class="detail-row"><span class="detail-label">Type</span><span class="detail-value fw-600">${esc(c.complaint_type)}</span></div>
        <div class="detail-row"><span class="detail-label">Engineer</span><span class="detail-value">${esc(c.engineer_name||'Unassigned')}</span></div>
        <div class="detail-row"><span class="detail-label">Logged On</span><span class="detail-value">${fmtDateTime(c.created_at)}</span></div>
        <div class="detail-row"><span class="detail-label">Resolved On</span><span class="detail-value">${fmtDateTime(c.resolved_at)}</span></div>
        ${c.remarks?`<div class="detail-row" style="grid-column:1/-1"><span class="detail-label">Remarks</span><span class="detail-value">${esc(c.remarks)}</span></div>`:''}
      </div>
      <div class="divider"></div>
      <div class="section-label">Update Status</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        ${['Open','In Progress','Resolved','Closed'].map(s =>
          `<button class="btn btn-sm ${c.status===s?'btn-primary':'btn-secondary'}" onclick="quickStatus(${c.id},'${s}')">${s}</button>`
        ).join('')}
      </div>`, { wide: true });
  } catch(e) { toast(e.message, 'danger'); }
};

window.quickStatus = async (id, status) => {
  try {
    const { data: c } = await API.get(`/api/complaints/${id}`);
    await API.put(`/api/complaints/${id}`, {
      complaint_type: c.complaint_type, remarks: c.remarks,
      engineer_id: c.engineer_id, status, priority: c.priority,
    });
    toast(`Status → "${status}"`, 'success');
    Modal.close();
    navigate('complaints');
  } catch(e) { toast(e.message, 'danger'); }
};

window.openComplaintEdit = async id => {
  try {
    const { data: c } = await API.get(`/api/complaints/${id}`);
    Modal.open(`Edit ${c.complaint_no}`, `
      <form id="editCmpForm" style="display:flex;flex-direction:column;gap:16px">
        <div class="form-grid form-grid-2">
          <div class="form-group">
            <label class="form-label">Complaint Type <span class="req">*</span></label>
            <select class="form-control" name="complaint_type" required>
              ${S.complaintTypes.map(t=>`<option value="${esc(t.name)}"${t.name===c.complaint_type?' selected':''}>${esc(t.name)}</option>`).join('')}
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Priority</label>
            <select class="form-control" name="priority">
              ${['Low','Normal','High','Urgent'].map(p=>`<option value="${p}"${c.priority===p?' selected':''}>${p}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="form-grid form-grid-2">
          <div class="form-group">
            <label class="form-label">Status</label>
            <select class="form-control" name="status">
              ${['Open','In Progress','Resolved','Closed'].map(s=>`<option value="${s}"${c.status===s?' selected':''}>${s}</option>`).join('')}
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Engineer</label>
            <select class="form-control" name="engineer_id">
              <option value="">Unassigned</option>${engineerOptions(c.engineer_id)}
            </select>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Remarks</label>
          <textarea class="form-control" name="remarks">${esc(c.remarks||'')}</textarea>
        </div>
        <div class="modal-footer" style="padding:0;border:none">
          <button type="button" class="btn btn-secondary" onclick="Modal.close()">Cancel</button>
          <button type="submit" class="btn btn-primary">Save Changes</button>
        </div>
      </form>`);
    document.getElementById('editCmpForm').onsubmit = async e => {
      e.preventDefault();
      const btn = e.target.querySelector('[type=submit]');
      btn.disabled = true;
      try {
        await API.put(`/api/complaints/${id}`, Object.fromEntries(new FormData(e.target)));
        toast('Complaint updated', 'success');
        Modal.close();
        navigate('complaints');
      } catch(err) { toast(err.message, 'danger'); }
      finally { btn.disabled = false; }
    };
  } catch(e) { toast(e.message, 'danger'); }
};

window.deleteComplaint = async (id, no) => {
  const ok = await confirmDialog('Delete Complaint', `Delete "${no}"? This cannot be undone.`, { danger: true, confirmText: 'Delete' });
  if (!ok) return;
  try { await API.delete(`/api/complaints/${id}`); toast('Complaint deleted', 'success'); navigate('complaints'); }
  catch(e) { toast(e.message, 'danger'); }
};

// ═══════════════════════════════════════════════════════════════════════════════
// ENGINEERS
// ═══════════════════════════════════════════════════════════════════════════════
async function engineers(el) {
  let data = [];
  try { data = (await API.get('/api/engineers')).data; } catch(_) {}

  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">Engineers</div>
        <div class="page-header-sub">Manage field engineers and technicians</div>
      </div>
      <button class="btn btn-primary" id="addEngBtn">
        <i data-lucide="user-plus"></i> Add Engineer
      </button>
    </div>
    <div class="card">
      ${!data.length ? `
        <div class="empty-state">
          <i data-lucide="hard-hat"></i><h3>No engineers added yet</h3>
          <p>Add engineers to assign complaints.</p>
        </div>` : `
      <div class="table-wrapper">
        <table>
          <thead><tr><th>Name</th><th>Contact</th><th>Open</th><th>Total</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${data.map(e => `
              <tr>
                <td class="fw-600">${esc(e.name)}</td>
                <td>${esc(e.contact||'—')}</td>
                <td><span class="badge ${e.open_complaints>0?'badge-warning':'badge-gray'}">${e.open_complaints}</span></td>
                <td>${e.total_complaints}</td>
                <td>${e.active?'<span class="badge badge-success">Active</span>':'<span class="badge badge-gray">Inactive</span>'}</td>
                <td>
                  <div class="gap-10">
                    <button class="btn btn-sm btn-ghost" onclick="openEngineerForm(${e.id})"><i data-lucide="pencil"></i></button>
                    <button class="btn btn-sm btn-ghost ${e.active?'text-danger':''}" onclick="toggleEngineer(${e.id},${e.active?0:1},'${esc(e.name)}')">
                      <i data-lucide="${e.active?'user-x':'user-check'}"></i>
                    </button>
                  </div>
                </td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>`}
    </div>`;
  lucide.createIcons({ nodes: [el] });
  document.getElementById('addEngBtn').onclick = () => openEngineerForm();
}

window.openEngineerForm = (id = null) => {
  const existing = id ? S.engineers.find(e => e.id === id) : null;
  Modal.open(existing ? 'Edit Engineer' : 'Add Engineer', `
    <form id="engForm" style="display:flex;flex-direction:column;gap:16px">
      <div class="form-group">
        <label class="form-label">Full Name <span class="req">*</span></label>
        <input class="form-control" name="name" type="text" placeholder="Engineer name"
          value="${esc(existing?.name||'')}" required autofocus />
      </div>
      <div class="form-group">
        <label class="form-label">Contact No.</label>
        <input class="form-control" name="contact" type="text" placeholder="Mobile number"
          value="${esc(existing?.contact||'')}" />
      </div>
      <div class="modal-footer" style="padding:0;border:none">
        <button type="button" class="btn btn-secondary" onclick="Modal.close()">Cancel</button>
        <button type="submit" class="btn btn-primary">${existing?'Save Changes':'Add Engineer'}</button>
      </div>
    </form>`);
  document.getElementById('engForm').onsubmit = async e => {
    e.preventDefault();
    const btn = e.target.querySelector('[type=submit]');
    btn.disabled = true;
    try {
      const fd = Object.fromEntries(new FormData(e.target));
      if (existing) await API.put(`/api/engineers/${id}`, { ...fd, active: existing.active });
      else          await API.post('/api/engineers', fd);
      toast(existing ? 'Engineer updated' : 'Engineer added', 'success');
      Modal.close();
      await preload();
      navigate('engineers');
    } catch(err) { toast(err.message, 'danger'); }
    finally { btn.disabled = false; }
  };
};

window.toggleEngineer = async (id, active, name) => {
  const action = active ? 'activate' : 'deactivate';
  const ok = await confirmDialog(
    `${action.charAt(0).toUpperCase()+action.slice(1)} Engineer`,
    `${action.charAt(0).toUpperCase()+action.slice(1)} "${name}"?`,
    { confirmText: action.charAt(0).toUpperCase()+action.slice(1) }
  );
  if (!ok) return;
  try {
    const eng = S.engineers.find(e => e.id === id);
    await API.put(`/api/engineers/${id}`, { name: eng.name, contact: eng.contact, active });
    toast(`Engineer ${action}d`, 'success');
    await preload();
    navigate('engineers');
  } catch(e) { toast(e.message, 'danger'); }
};

// ═══════════════════════════════════════════════════════════════════════════════
// WHATSAPP PAGE
// ═══════════════════════════════════════════════════════════════════════════════
async function whatsappPage(el) {
  await waPoll();
  renderWaPage(el);
}

function renderWaPage(el) {
  const s = WA.status;
  const statusLabel = {
    off:'Not Connected', init:'Starting up…', authenticated:'Authenticated…',
    qr:'Scan QR Code', ready:'Connected', disconnected:'Disconnected',
    auth_fail:'Auth Failed', error:'Error',
  };
  const statusCls = { ready:'badge-success', qr:'badge-warning', init:'badge-info',
    authenticated:'badge-info', disconnected:'badge-danger', auth_fail:'badge-danger', error:'badge-danger' };

  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">WhatsApp Integration</div>
        <div class="page-header-sub">Connect your WhatsApp to auto-send complaints to your engineer group</div>
      </div>
      <button class="btn btn-secondary" id="waRefreshBtn">
        <i data-lucide="refresh-cw"></i> Refresh Status
      </button>
    </div>
    <div style="max-width:680px;display:flex;flex-direction:column;gap:20px">

      <div class="card">
        <div class="card-header">
          <span class="card-title">Connection Status</span>
          <span class="badge ${statusCls[s]||'badge-gray'}">${statusLabel[s]||s}</span>
        </div>
        <div class="card-body" id="waConnBody">${renderWaConnBody(s)}</div>
      </div>

      ${s === 'ready' ? `
      <div class="card">
        <div class="card-header">
          <span class="card-title">Notification Target</span>
          ${WA.target_name?`<span class="badge badge-success">${esc(WA.target_name)}</span>`:''}
        </div>
        <div class="card-body">
          <p style="font-size:13px;color:var(--text-secondary);margin-bottom:14px">
            Select which group receives the complaint message automatically.
          </p>
          <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
            <div class="form-group" style="flex:1;min-width:200px">
              <label class="form-label">Select Group</label>
              <select class="form-control" id="waTargetSelect"><option value="">Loading groups…</option></select>
            </div>
            <button class="btn btn-secondary" id="waRetryGroups" style="flex-shrink:0;display:none">
              <i data-lucide="refresh-cw"></i> Retry
            </button>
            <button class="btn btn-primary" id="waSetTargetBtn" style="flex-shrink:0">
              <i data-lucide="save"></i> Set Default
            </button>
          </div>
          ${WA.target_name?`
          <div class="info-banner" style="margin-top:12px">
            <i data-lucide="check-circle-2"></i>
            Messages will be sent to: <b>${esc(WA.target_name)}</b>
          </div>`:''}
        </div>
      </div>` : ''}

      <div class="card">
        <div class="card-header"><span class="card-title">How It Works</span></div>
        <div class="card-body" style="display:flex;flex-direction:column;gap:14px">
          ${[
            ['message-circle', 'Click <b>Connect WhatsApp</b> — a QR code will appear on screen.'],
            ['smartphone',     'Open WhatsApp on your phone → tap <b>⋮ Menu → Linked Devices → Link a Device</b> → scan the QR.'],
            ['check-circle-2', 'After scanning, the page will <b>automatically update</b> to Connected within a few seconds.'],
            ['users',          'Once Connected, select your engineer group from the dropdown and click <b>Set Default</b>.'],
            ['send',           'From now on, every new complaint will <b>automatically</b> send a message to that group.'],
            ['shield',         'Session is saved — you only scan QR once. If disconnected, just click Connect again.'],
          ].map(([, text], i) => `
            <div style="display:flex;gap:12px;align-items:flex-start">
              <div class="step-num">${i+1}</div>
              <div style="font-size:13px;color:var(--text-secondary);padding-top:2px">${text}</div>
            </div>`).join('')}
        </div>
      </div>
    </div>`;
  lucide.createIcons({ nodes: [el] });

  // Refresh Status button — re-polls then re-renders
  document.getElementById('waRefreshBtn').addEventListener('click', async () => {
    const btn = document.getElementById('waRefreshBtn');
    btn.disabled = true;
    await waPoll();
    S.page = null;
    navigate('whatsapp');
  });

  document.getElementById('waConnectBtn')?.addEventListener('click', async () => {
    const btn = document.getElementById('waConnectBtn');
    btn.disabled = true; btn.textContent = 'Starting…';
    try {
      await API.post('/api/whatsapp/connect');
      startWaStatusPolling(el);
      // Re-render to show spinner immediately
      S.page = null; navigate('whatsapp');
    } catch(e) {
      toast(e.message, 'danger');
      btn.disabled = false; btn.textContent = 'Connect WhatsApp';
    }
  });
  document.getElementById('waDisconnectBtn')?.addEventListener('click', async () => {
    const ok = await confirmDialog('Disconnect WhatsApp', 'You will need to scan the QR code again to reconnect.');
    if (!ok) return;
    try {
      await API.post('/api/whatsapp/disconnect');
      Object.assign(WA, { status:'off', target_id:null, target_name:null });
      toast('WhatsApp disconnected', 'info');
      S.page = null; navigate('whatsapp');
    } catch(e) { toast(e.message, 'danger'); }
  });
  if (s === 'ready') {
    loadWaTargets();
    document.getElementById('waSetTargetBtn').onclick = saveWaTarget;
    document.getElementById('waRetryGroups').onclick = () => loadWaTargets();
  }
  // Always poll when on this page so any status change auto-updates
  if (!['ready','off'].includes(s)) startWaStatusPolling(el);
  else if (s === 'off') {
    // Poll every 2s — handle both direct 'ready' jump and intermediate states
    if (_waPollTimer) clearInterval(_waPollTimer);
    _waPollTimer = setInterval(async () => {
      await waPoll();
      if (WA.status === 'off') return; // still off, keep waiting
      clearInterval(_waPollTimer); _waPollTimer = null;
      if (S.page !== 'whatsapp') return;
      if (WA.status === 'ready') {
        // jumped straight to ready (session restored)
        S.page = null; navigate('whatsapp');
      } else {
        // in-progress state (init/qr/authenticated) — start fast poll
        startWaStatusPolling(el);
        S.page = null; navigate('whatsapp');
      }
    }, 2000);
  }
}

function renderWaConnBody(s) {
  if (s === 'ready') return `
    <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap">
      <div style="width:48px;height:48px;background:var(--success-bg);border-radius:50%;display:flex;align-items:center;justify-content:center;color:var(--success)">
        <i data-lucide="check-circle-2"></i>
      </div>
      <div>
        <div class="fw-700" style="font-size:15px">WhatsApp Connected</div>
        <div class="td-muted">${WA.account ? `Logged in as: ${esc(WA.account)}` : 'Session active'}</div>
      </div>
      <button class="btn btn-secondary btn-sm text-danger" id="waDisconnectBtn" style="margin-left:auto">
        <i data-lucide="log-out"></i> Disconnect
      </button>
    </div>`;
  if (s === 'qr') return `
    <div style="text-align:center;padding:8px 0">
      <p style="font-size:13px;color:var(--text-secondary);margin-bottom:16px">
        Open WhatsApp on your phone → Settings → Linked Devices → Link a Device → Scan this QR
      </p>
      <div style="display:inline-block;padding:14px;background:#fff;border:2px solid var(--border);border-radius:var(--radius-lg)">
        <img id="waQrImg" src="" alt="QR Code" style="width:240px;height:240px;display:block" />
      </div>
      <p class="td-muted" style="margin-top:10px;font-size:12px">QR code refreshes automatically</p>
    </div>`;
  if (s === 'init' || s === 'authenticated') return `
    <div style="display:flex;align-items:center;gap:12px;padding:6px 0">
      <div class="spinner" style="width:22px;height:22px;border-width:2px;flex-shrink:0"></div>
      <span style="font-size:13.5px;color:var(--text-secondary)">
        ${s === 'init' ? 'Starting WhatsApp Web, please wait…' : 'Authenticating…'}
      </span>
    </div>`;
  const errMsg = WA.lastError || null;
  return `
    <div>
      ${s === 'error' && errMsg ? `
        <div style="background:var(--danger-bg);border:1px solid #fca5a5;border-radius:var(--radius);padding:10px 14px;margin-bottom:14px;font-size:12.5px;color:var(--danger)">
          <b>Error:</b> ${esc(errMsg)}
        </div>` : ''}
      <p style="font-size:13.5px;color:var(--text-secondary);margin-bottom:14px">
        Connect your WhatsApp to send complaint notifications directly to your engineer group.
        ${['auth_fail','disconnected'].includes(s) ? '<br><span style="color:var(--danger);font-size:12.5px">Previous session ended. Please connect again.</span>' : ''}
        ${s === 'error' && !errMsg ? '<br><span style="color:var(--danger);font-size:12.5px">Failed to start. Please try again.</span>' : ''}
      </p>
      <button class="btn btn-primary" id="waConnectBtn">
        <i data-lucide="message-circle"></i> ${s === 'error' ? 'Retry Connect' : 'Connect WhatsApp'}
      </button>
    </div>`;
}

// ─── Server-Sent Events — real-time WA status (no polling needed) ────────────
function setupWaSSE() {
  const es = new EventSource('/api/whatsapp/stream');

  es.addEventListener('status', e => {
    const d = JSON.parse(e.data);
    const prev = WA.status;
    Object.assign(WA, d);

    // Keep nav dot in sync
    const dot = document.getElementById('navWaDot');
    if (dot) dot.classList.toggle('visible', d.status === 'ready');

    // Re-render WA page instantly on any status change
    if (d.status !== prev && S.page === 'whatsapp') {
      S.page = null; navigate('whatsapp');
    }

    // Update QR image in-place (avoid full re-render during QR phase)
    if (d.hasQr && S.page === 'whatsapp') {
      const img = document.getElementById('waQrImg');
      if (img) {
        API.get('/api/whatsapp/qr').then(({ data }) => {
          const imgNow = document.getElementById('waQrImg');
          if (imgNow) imgNow.src = data.qr;
        }).catch(() => {});
      } else if (d.status !== prev) {
        S.page = null; navigate('whatsapp');
      }
    }
  });

  es.addEventListener('groups_ready', () => {
    // Groups just loaded on the server — refresh the dropdown only if it's still
    // showing a loading/syncing placeholder (not if user already has groups open)
    if (S.page === 'whatsapp' && WA.status === 'ready') {
      const sel = document.getElementById('waTargetSelect');
      const hasRealGroups = sel && sel.options.length > 1 && sel.options[1]?.value;
      if (sel && !hasRealGroups) loadWaTargets();
    }
  });

  es.onerror = () => { /* EventSource auto-reconnects — no action needed */ };
}

let _waPollTimer = null;
// Fallback polling — only used when SSE is unavailable (e.g., old proxy config)
function startWaStatusPolling(_el) {
  if (_waPollTimer) clearInterval(_waPollTimer);
  _waPollTimer = setInterval(async () => {
    await waPoll();
    if (WA.hasQr) {
      try {
        const { data } = await API.get('/api/whatsapp/qr');
        const img = document.getElementById('waQrImg');
        if (img) img.src = data.qr;
        else if (S.page === 'whatsapp') { clearInterval(_waPollTimer); _waPollTimer = null; navigate('whatsapp'); }
      } catch(_) {}
    }
    if (['ready','off','auth_fail','error'].includes(WA.status)) {
      clearInterval(_waPollTimer); _waPollTimer = null;
      if (S.page === 'whatsapp') { S.page = null; navigate('whatsapp'); }
    }
  }, 2000);
}

async function loadWaTargets() {
  const sel = document.getElementById('waTargetSelect');
  const retryBtn = document.getElementById('waRetryGroups');
  if (!sel) return;
  sel.innerHTML = `<option value="">Loading groups…</option>`;
  sel.disabled = true;
  if (retryBtn) retryBtn.style.display = 'none';
  try {
    const { data: groups } = await API.get('/api/whatsapp/groups');
    sel.disabled = false;
    if (!groups.length) {
      // Groups still syncing — show static message and schedule a silent retry.
      // We do NOT touch innerHTML during the countdown so an open dropdown stays open.
      sel.innerHTML = `<option value="">WhatsApp syncing groups…</option>`;
      if (retryBtn) {
        retryBtn.style.display = '';
        retryBtn.onclick = () => loadWaTargets();
      }
      // Auto-retry after 8 s, but skip if the dropdown is open (focused)
      setTimeout(() => {
        const s = document.getElementById('waTargetSelect');
        if (!s || WA.status !== 'ready') return;
        if (document.activeElement === s) {
          // Defer until user closes the dropdown
          s.addEventListener('blur', () => loadWaTargets(), { once: true });
        } else {
          loadWaTargets();
        }
      }, 8000);
      return;
    }
    sel.innerHTML = `
      <option value="">— Choose a group —</option>
      ${groups.map(g =>
        `<option value="${esc(g.id)}" ${WA.target_id===g.id?'selected':''}>${esc(g.name)} (${g.members} members)</option>`
      ).join('')}`;
  } catch(e) {
    sel.disabled = false;
    sel.innerHTML = `<option value="">Error loading groups — click Retry</option>`;
    if (retryBtn) retryBtn.style.display = '';
    console.error('[WA groups]', e.message);
  }
}

async function saveWaTarget() {
  const sel = document.getElementById('waTargetSelect');
  if (!sel?.value) { toast('Please select a group or contact', 'warning'); return; }
  const btn = document.getElementById('waSetTargetBtn');
  btn.disabled = true;
  try {
    const name = sel.options[sel.selectedIndex].text;
    await API.post('/api/whatsapp/target', { id: sel.value, name });
    WA.target_id = sel.value; WA.target_name = name;
    toast(`Target set: ${name}`, 'success');
    navigate('whatsapp');
  } catch(e) { toast(e.message, 'danger'); btn.disabled = false; }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ACTIVITY LOGS
// ═══════════════════════════════════════════════════════════════════════════════
let _logsRefreshTimer = null;

async function activityLogs(el) {
  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">Activity Log</div>
        <div class="page-header-sub">All system events — auto-refreshes every 10 seconds</div>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn btn-secondary" id="clearLogsBtn">
          <i data-lucide="trash-2"></i> Clear Log
        </button>
        <button class="btn btn-secondary" id="refreshLogsBtn">
          <i data-lucide="refresh-cw"></i> Refresh
        </button>
      </div>
    </div>
    <div class="card" id="logsCard">
      <div id="logsWrap"><div class="page-loading"><div class="spinner"></div></div></div>
    </div>`;
  lucide.createIcons({ nodes: [el] });

  const load = async (silent = false) => {
    const wrap = document.getElementById('logsWrap');
    if (!wrap) return;
    if (!silent) wrap.innerHTML = `<div class="page-loading"><div class="spinner"></div></div>`;
    try {
      const { data } = await API.get('/api/logs');
      renderLogs(wrap, data);
    } catch(e) {
      wrap.innerHTML = `<div class="empty-state"><p class="text-danger">${esc(e.message)}</p></div>`;
    }
  };

  document.getElementById('refreshLogsBtn')?.addEventListener('click', () => load());
  document.getElementById('clearLogsBtn')?.addEventListener('click', async () => {
    const ok = await confirmDialog('Clear Activity Log', 'Delete all log entries? This cannot be undone.', { danger: true, confirmText: 'Clear All' });
    if (!ok) return;
    try {
      await API.delete('/api/logs');
      toast('Activity log cleared', 'success');
      load();
    } catch(e) { toast(e.message, 'danger'); }
  });

  if (_logsRefreshTimer) clearInterval(_logsRefreshTimer);
  _logsRefreshTimer = setInterval(() => {
    if (S.page !== 'logs') { clearInterval(_logsRefreshTimer); _logsRefreshTimer = null; return; }
    load(true);
  }, 10000);

  load();
}

function renderLogs(wrap, logs) {
  if (!logs.length) {
    wrap.innerHTML = `
      <div class="empty-state">
        <i data-lucide="activity"></i>
        <h3>No activity yet</h3>
        <p>Events appear here as you log complaints, add customers, etc.</p>
      </div>`;
    lucide.createIcons({ nodes: [wrap] });
    return;
  }

  const iconMap = {
    'complaint-new':    ['file-plus',      'success'],
    'complaint-update': ['file-pen-line',  'warning'],
    'complaint-delete': ['file-x',         'danger' ],
    'customer-add':     ['user-plus',      'success'],
    'customer-update':  ['user-check',     'info'   ],
    'customer-delete':  ['user-x',         'danger' ],
    'engineer-add':     ['hard-hat',       'success'],
    'engineer-update':  ['settings-2',     'info'   ],
    'import':           ['upload',         'purple' ],
    'sync':             ['refresh-cw',     'info'   ],
  };

  // Group by local date
  const grouped = {};
  for (const log of logs) {
    const key = new Date(log.created_at).toLocaleDateString('en-IN', { day:'2-digit', month:'long', year:'numeric' });
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(log);
  }

  wrap.innerHTML = `<div class="log-feed">` +
    Object.entries(grouped).map(([date, entries]) => `
      <div class="log-date-sep">
        <i data-lucide="calendar" style="width:12px;height:12px"></i> ${date}
        <span class="log-date-count">${entries.length} event${entries.length !== 1 ? 's' : ''}</span>
      </div>
      ${entries.map(log => {
        const [icon, type] = iconMap[log.event_type] || ['activity', 'info'];
        return `
          <div class="log-entry">
            <div class="log-icon log-icon-${type}"><i data-lucide="${icon}"></i></div>
            <div class="log-body">
              <div class="log-desc">${esc(log.description)}</div>
              <div class="log-time">${fmtDateTime(log.created_at)}</div>
            </div>
          </div>`;
      }).join('')}
    `).join('') + `</div>`;
  lucide.createIcons({ nodes: [wrap] });
}

// ═══════════════════════════════════════════════════════════════════════════════
// REPORTS PAGE
// ═══════════════════════════════════════════════════════════════════════════════
let _rptCustomer = null; // selected customer object

async function reportsPage(el) {
  _rptCustomer = null;

  el.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-header-title">Reports</div>
        <div class="page-header-sub">Filter by customer, area, date range or engineer</div>
      </div>
      <a class="btn btn-secondary" id="rptExportBtn" href="#" style="pointer-events:none;opacity:.4">
        <i data-lucide="file-up"></i> Export XLSX
      </a>
    </div>

    <div class="card rpt-filter-card">
      <div class="card-body" style="padding:16px 20px">
        <div class="rpt-filter-grid">

          <!-- Customer -->
          <div class="form-group" style="margin:0;position:relative">
            <label class="form-label">Customer</label>
            <div class="rpt-ac-box">
              <i data-lucide="search" class="rpt-ac-icon"></i>
              <input class="form-control rpt-ac-input" id="rptCustInput"
                placeholder="Name or serial no…" autocomplete="off"/>
              <button class="rpt-ac-clear" id="rptCustClear" style="display:none" title="Clear">
                <i data-lucide="x"></i>
              </button>
            </div>
            <div class="rpt-ac-dropdown" id="rptCustDropdown"></div>
          </div>

          <!-- Status -->
          <div class="form-group" style="margin:0">
            <label class="form-label">Status</label>
            <select class="form-control" id="rptStatus">
              <option value="">All Status</option>
              ${['Open','In Progress','Resolved','Closed'].map(s=>`<option value="${s}">${s}</option>`).join('')}
            </select>
          </div>

          <!-- Area -->
          <div class="form-group" style="margin:0">
            <label class="form-label">Area</label>
            <select class="form-control" id="rptArea">
              <option value="">All Areas</option>
              ${(S.areas||[]).map(a=>`<option value="${esc(a)}">${esc(a)}</option>`).join('')}
            </select>
          </div>

          <!-- Quick range -->
          <div class="form-group" style="margin:0">
            <label class="form-label">Date Range</label>
            <select class="form-control" id="rptRange">
              <option value="">All Time</option>
              <option value="today">Today</option>
              <option value="yesterday">Yesterday</option>
              <option value="week">This Week</option>
              <option value="month">This Month</option>
              <option value="custom">Custom Range…</option>
            </select>
          </div>

          <!-- From / To (custom) -->
          <div class="form-group" id="rptFromWrap" style="margin:0;display:none">
            <label class="form-label">From</label>
            <input class="form-control" id="rptFrom" type="date"/>
          </div>
          <div class="form-group" id="rptToWrap" style="margin:0;display:none">
            <label class="form-label">To</label>
            <input class="form-control" id="rptTo" type="date"/>
          </div>

          <!-- Actions -->
          <div style="display:flex;align-items:flex-end;gap:8px" class="rpt-filter-actions">
            <button class="btn btn-primary" id="rptSearchBtn" style="white-space:nowrap">
              <i data-lucide="refresh-cw"></i> Refresh
            </button>
            <button class="btn btn-ghost" id="rptClearBtn" style="white-space:nowrap">
              <i data-lucide="x"></i> Clear
            </button>
          </div>
        </div>

        <!-- Selected customer chip -->
        <div id="rptCustChip" style="display:none;margin-top:10px"></div>
      </div>
    </div>

    <div id="rptResults" style="margin-top:16px"></div>`;

  lucide.createIcons({ nodes: [el] });

  // ── Date range toggle ────────────────────────────────────────────────────────
  const rangeSelect = document.getElementById('rptRange');
  rangeSelect.addEventListener('change', () => {
    const show = rangeSelect.value === 'custom';
    document.getElementById('rptFromWrap').style.display = show ? '' : 'none';
    document.getElementById('rptToWrap').style.display   = show ? '' : 'none';
  });

  const getDateRange = () => {
    const today = new Date();
    const fmt = d => d.toISOString().slice(0,10);
    switch (rangeSelect.value) {
      case 'today':     return { from: fmt(today), to: fmt(today) };
      case 'yesterday': { const y = new Date(today); y.setDate(y.getDate()-1); return { from:fmt(y), to:fmt(y) }; }
      case 'week':      { const w = new Date(today); w.setDate(w.getDate()-6); return { from:fmt(w), to:fmt(today) }; }
      case 'month':     return { from: fmt(new Date(today.getFullYear(), today.getMonth(), 1)), to: fmt(today) };
      case 'custom':    return { from: document.getElementById('rptFrom').value, to: document.getElementById('rptTo').value };
      default:          return { from:'', to:'' };
    }
  };

  // ── Customer autocomplete ────────────────────────────────────────────────────
  const custInput    = document.getElementById('rptCustInput');
  const custDropdown = document.getElementById('rptCustDropdown');
  const custClear    = document.getElementById('rptCustClear');
  const custChip     = document.getElementById('rptCustChip');

  const rptSelectCustomer = (c) => {
    _rptCustomer = c;
    custInput.value = `${c.new_party_name} — NSN ${c.nsn}`;
    custDropdown.innerHTML = ''; custDropdown.style.display = 'none';
    custClear.style.display = '';
    custChip.style.display = '';
    custChip.innerHTML = `
      <div class="rpt-cust-chip">
        <i data-lucide="user-check"></i>
        <span><b>${esc(c.new_party_name)}</b> · NSN: ${c.nsn}${c.area ? ' · '+esc(c.area) : ''}</span>
      </div>`;
    lucide.createIcons({ nodes: [custChip] });
    debouncedLoad();
  };

  const rptClearCustomer = () => {
    _rptCustomer = null;
    custInput.value = '';
    custClear.style.display = 'none';
    custChip.style.display = 'none';
    custDropdown.style.display = 'none';
  };

  custClear.addEventListener('click', () => { rptClearCustomer(); custInput.focus(); });

  const searchRptCustomers = debounce(async (q) => {
    if (!q) { custDropdown.style.display = 'none'; return; }
    try {
      const { data } = await API.get(`/api/customers?search=${encodeURIComponent(q)}&limit=8`);
      custDropdown.innerHTML = data.length
        ? data.map((c,i) => `
            <div class="rpt-ac-item" data-i="${i}">
              <div class="rpt-ac-name">${esc(c.new_party_name)}</div>
              <div class="rpt-ac-meta">NSN: ${c.nsn} · ${esc(c.area||'—')}</div>
            </div>`).join('')
        : `<div class="rpt-ac-empty">No customers found</div>`;
      custDropdown.querySelectorAll('.rpt-ac-item').forEach(item =>
        item.addEventListener('click', () => rptSelectCustomer(data[+item.dataset.i])));
      custDropdown.style.display = 'block';
    } catch(_) {}
  }, 250);

  custInput.addEventListener('input',  e => { if (_rptCustomer) rptClearCustomer(); searchRptCustomers(e.target.value.trim()); });
  custInput.addEventListener('blur',   ()  => setTimeout(() => { custDropdown.style.display = 'none'; }, 200));
  custInput.addEventListener('focus',  e  => { if (e.target.value && !_rptCustomer) searchRptCustomers(e.target.value); });

  // ── Clear all filters ─────────────────────────────────────────────────────────
  document.getElementById('rptClearBtn').addEventListener('click', () => {
    rptClearCustomer();
    document.getElementById('rptStatus').value = '';
    document.getElementById('rptArea').value   = '';
    rangeSelect.value = '';
    document.getElementById('rptFromWrap').style.display = 'none';
    document.getElementById('rptToWrap').style.display   = 'none';
    load();
  });

  // ── Generate ─────────────────────────────────────────────────────────────────
  const load = async () => {
    const wrap = document.getElementById('rptResults');
    wrap.innerHTML = `<div class="page-loading"><div class="spinner"></div></div>`;
    const { from, to } = getDateRange();
    const p = new URLSearchParams({
      nsn:       _rptCustomer ? _rptCustomer.nsn : '',
      search:    _rptCustomer ? '' : custInput.value.trim(),
      status:    document.getElementById('rptStatus').value,
      area:      document.getElementById('rptArea').value,
      date_from: from, date_to: to,
    });
    try {
      const { data } = await API.get(`/api/reports/complaints?${p}`);
      renderReportResults(wrap, data, _rptCustomer);
      const exportBtn = document.getElementById('rptExportBtn');
      if (data.length) {
        exportBtn.style.opacity = '1'; exportBtn.style.pointerEvents = '';
        exportBtn.href = `/api/export/complaints?${p}`;
      } else {
        exportBtn.style.opacity = '.4'; exportBtn.style.pointerEvents = 'none'; exportBtn.href = '#';
      }
    } catch(e) {
      wrap.innerHTML = `<div class="empty-state"><p class="text-danger">${esc(e.message)}</p></div>`;
    }
  };

  const debouncedLoad = debounce(load, 400);

  document.getElementById('rptSearchBtn').addEventListener('click', load);
  custInput.addEventListener('keydown', e => { if (e.key === 'Enter') load(); });

  // Auto-generate whenever any filter changes
  document.getElementById('rptStatus').addEventListener('change', debouncedLoad);
  document.getElementById('rptArea').addEventListener('change', debouncedLoad);
  rangeSelect.addEventListener('change', () => {
    if (rangeSelect.value !== 'custom') debouncedLoad();
  });
  document.getElementById('rptFrom').addEventListener('change', debouncedLoad);
  document.getElementById('rptTo').addEventListener('change', debouncedLoad);

  // Auto-load on open (show all complaints)
  load();
}

function renderReportResults(wrap, data, customer) {
  if (!data.length) {
    wrap.innerHTML = `
      <div class="empty-state">
        <i data-lucide="file-search"></i>
        <h3>No complaints found</h3>
        <p>Try adjusting the filters or selecting a different customer.</p>
      </div>`;
    lucide.createIcons({ nodes: [wrap] }); return;
  }

  const total    = data.length;
  const open     = data.filter(c => c.status === 'Open').length;
  const progress = data.filter(c => c.status === 'In Progress').length;
  const resolved = data.filter(c => c.status === 'Resolved' || c.status === 'Closed').length;
  const customerCard = customer ? `
    <div class="card rpt-cust-card">
      <div class="card-body rpt-cust-body">
        <div class="rpt-cust-avatar"><i data-lucide="user"></i></div>
        <div class="rpt-cust-info">
          <div class="rpt-cust-name">${esc(customer.new_party_name)}</div>
          <div class="rpt-cust-details">
            <span><i data-lucide="hash"></i> NSN: <b>${customer.nsn}</b></span>
            ${customer.contact_no ? `<span><i data-lucide="phone"></i> ${esc(customer.contact_no)}</span>` : ''}
            ${customer.area       ? `<span><i data-lucide="map-pin"></i> ${esc(customer.area)}</span>` : ''}
          </div>
          ${customer.address ? `<div class="rpt-cust-addr">${esc(customer.address)}</div>` : ''}
        </div>
        <div class="rpt-cust-badge">${warrantyBadge(customer.install_date)}</div>
      </div>
    </div>` : '';

  wrap.innerHTML = `
    ${customerCard}
    <div class="rpt-stats-bar">
      ${[
        { label: 'Total',       val: total,    icon: 'layers',       cls: 'blue'   },
        { label: 'Open',        val: open,     icon: 'alert-circle', cls: 'amber'  },
        { label: 'In Progress', val: progress, icon: 'clock',        cls: 'purple' },
        { label: 'Resolved',    val: resolved, icon: 'check-circle', cls: 'green'  },
      ].map(s => `
        <div class="rpt-stat-card rpt-stat-${s.cls}">
          <div class="rpt-stat-icon"><i data-lucide="${s.icon}"></i></div>
          <div class="rpt-stat-val">${s.val}</div>
          <div class="rpt-stat-label">${s.label}</div>
        </div>`).join('')}
    </div>

    <div class="card">
      <div class="card-header">
        <span class="card-title">Complaint History</span>
        <span class="badge badge-info">${total} record${total !== 1 ? 's' : ''}</span>
      </div>
      <div class="table-wrapper">
        <table>
          <thead><tr>
            <th>Ref No</th>
            ${!customer ? '<th>Customer</th>' : ''}
            <th>Area</th>
            <th>Complaint</th>
            <th>Status</th>
            <th>Logged On</th>
            <th>Resolved On</th>
          </tr></thead>
          <tbody>
            ${data.map(c => `
              <tr>
                <td><span class="mono-tag">${esc(c.complaint_no)}</span></td>
                ${!customer ? `
                  <td>
                    <div class="fw-600">${esc(c.new_party_name)}</div>
                    <div class="td-muted">NSN: ${c.nsn}</div>
                  </td>` : ''}
                <td class="td-muted">${esc(c.area || '—')}</td>
                <td style="max-width:220px;white-space:normal;line-height:1.4">${esc(c.remarks || '—')}</td>
                <td>${statusBadge(c.status)}</td>
                <td class="td-muted" style="white-space:nowrap">${fmtDateTime(c.created_at)}</td>
                <td class="td-muted" style="white-space:nowrap">${c.resolved_at ? fmtDateTime(c.resolved_at) : '—'}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  lucide.createIcons({ nodes: [wrap] });
}

// ─── Sync Status ──────────────────────────────────────────────────────────────
function timeAgo(iso) {
  if (!iso) return null;
  const diff = Date.now() - new Date(iso).getTime();
  const mins = Math.floor(diff / 60000);
  if (mins < 1)  return 'just now';
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(mins / 60);
  if (hrs < 24)  return `${hrs}h ago`;
  return `${Math.floor(hrs / 24)}d ago`;
}

function updateSyncPill() {
  const pill = document.getElementById('syncPill');
  const text = document.getElementById('syncPillText');
  if (!pill || !text) return;
  if (!SYNC.xlsxExists) {
    text.textContent = 'No backup yet';
    pill.className = 'sync-pill';
    return;
  }
  const ago = timeAgo(SYNC.lastExport);
  text.textContent = ago ? `Synced ${ago}` : 'Sync ready';
  pill.className = 'sync-pill ok';
}

async function syncPoll() {
  try {
    const { data } = await API.get('/api/sync/status');
    SYNC.lastExport  = data.lastExport;
    SYNC.xlsxExists  = data.xlsxExists;
    updateSyncPill();
  } catch(_) {}
}

// ─── Globals ──────────────────────────────────────────────────────────────────
window.Modal    = Modal;
window.navigate = navigate;
window.S        = S;
window.WA       = WA;
window.openCustomerForm = openCustomerForm;

// ─── Boot ─────────────────────────────────────────────────────────────────────
(async () => {
  lucide.createIcons();
  startClock();
  await preload();
  await Promise.all([waPoll(), syncPoll()]);
  navigate('dashboard');
  setupWaSSE(); // real-time WA status via Server-Sent Events
  // Fallback poll every 10 s — catches status if SSE drops (proxy restart etc.)
  let _globalWaStatus = WA.status;
  setInterval(async () => {
    await waPoll();
    if (WA.status !== _globalWaStatus) {
      _globalWaStatus = WA.status;
      if (S.page === 'whatsapp') { S.page = null; navigate('whatsapp'); }
    }
  }, 10000);
  setInterval(syncPoll, 30000);
})();
