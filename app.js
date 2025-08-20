/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL    = 'https://script.google.com/macros/s/AKfycbwnw1pImnH46jD9F5bXwQGdxVu4YFyN9gvdtC1QU8YsD8wKTM5VPp8LZDeStx_JG_GVtQ/exec'; // <-- REPLACE
const API_SHARED_TOKEN= 't9x_93HDa8nL0PQ6RvzX4wqZ'; // <-- REPLACE
/* ===================================================== */

const STATUS_ORDER = [
  'PENDING AVAILABILITY',
  'NOT AVAILABLE',
  'LONG DAY',
  'NIGHT',
  'LONG DAY/NIGHT'
];
const STATUS_TO_CODE = {
  'PENDING AVAILABILITY': '',     // blank (white)
  'NOT AVAILABLE':        'N/A',  // red
  'LONG DAY':             'LD',   // yellow
  'NIGHT':                'N',    // yellow
  'LONG DAY/NIGHT':       'LD/N'  // yellow
};
const CODE_TO_STATUS = Object.entries(STATUS_TO_CODE).reduce((acc,[k,v]) => (acc[v]=k, acc), {});

const els = {
  grid: document.getElementById('grid'),
  footer: document.getElementById('footer'),
  draftMsg: document.getElementById('draftMsg'),
  submitBtn: document.getElementById('submitBtn'),
  clearBtn: document.getElementById('clearBtn'),
  refreshBtn: document.getElementById('refreshBtn'),
  lastLoaded: document.getElementById('lastLoaded'),
  toast: document.getElementById('toast'),
  installBtn: document.getElementById('installBtn'),
  // Optional hooks (safe if missing in HTML)
  candidateName: document.getElementById('candidateName'),
  helpMsg: document.getElementById('helpMsg') // if you add a help banner element
};

// State
let baseline = null;          // server response (tiles, lastLoadedAt, candidateName, candidate)
let draft = {};               // ymd -> status (only diffs from baseline)
let identity = {};            // { k?:string, msisdn?:string }
let lastHiddenAt = null;      // timestamp when page becomes hidden
const DRAFT_KEY = 'rota_avail_draft_v1';
const LAST_LOADED_KEY = 'rota_avail_last_loaded_v1';

// ------- Loading overlay (auto-created, no HTML changes required) -------
let _loadingCount = 0;
function ensureLoadingOverlay() {
  if (document.getElementById('loadingOverlay')) return;
  const style = document.createElement('style');
  style.textContent = `
    #loadingOverlay {
      position: fixed; inset: 0; z-index: 9999;
      display: flex; align-items: center; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
      color: #e7ecf3; font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      transition: opacity .15s ease;
    }
    #loadingOverlay.hidden { opacity: 0; pointer-events: none; }
    #loadingBox {
      display: inline-flex; align-items: center; gap: .7rem;
      border: 1px solid #2a3446; background: #131926;
      padding: .75rem 1rem; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,.35);
      font-weight: 600;
    }
    #loadingSpinner {
      width: 18px; height: 18px; border-radius: 50%;
      border: 3px solid rgba(255,255,255,.2); border-top-color: #4f8cff;
      animation: spin 0.9s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    /* Past Shifts overlay */
    #pastOverlay {
      position: fixed; inset: 0; z-index: 9998;
      display: none; align-items: stretch; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
    }
    #pastOverlay.show { display: flex; }
    #pastSheet {
      background: #0f1115; color: #e7ecf3;
      border-top-left-radius: 16px; border-top-right-radius: 16px;
      width: min(720px, 96vw);
      margin-top: min(8vh, 80px);
      margin-bottom: 0;
      box-shadow: 0 10px 30px rgba(0,0,0,.45);
      display: flex; flex-direction: column; max-height: calc(100vh - 12vh);
      border: 1px solid #222936;
    }
    #pastHeader {
      display:flex; align-items:center; gap:.6rem; padding:.7rem .9rem;
      border-bottom:1px solid #222936; background: rgba(15,17,21,.9); backdrop-filter: blur(6px);
    }
    #pastTitle { font-weight:800; font-size:1.05rem; margin:0; }
    #pastClose {
      margin-left:auto; appearance:none; border:0; background:transparent; color:#a7b0c0;
      font-weight:700; font-size:1rem; cursor:pointer; padding:.25rem .5rem; border-radius:8px;
    }
    #pastClose:hover { background:#1e2532; color:#fff; }
    #pastList {
      overflow:auto; padding:.7rem .9rem 1rem; display:flex; flex-direction:column; gap:.6rem;
    }
    .past-item {
      border:1px solid #222a36; background:#131926; border-radius:12px; padding:.6rem .7rem;
      line-height:1.35;
    }
    .past-date { font-weight:800; margin-bottom:.15rem; color:#e7ecf3; }
    .muted { color:#a7b0c0; }
  `;
  document.head.appendChild(style);

  const wrap = document.createElement('div');
  wrap.id = 'loadingOverlay';
  wrap.className = 'hidden';
  wrap.innerHTML = `
    <div id="loadingBox" role="status" aria-live="polite">
      <div id="loadingSpinner"></div>
      <span id="loadingText">Loading...</span>
    </div>
  `;
  document.body.appendChild(wrap);

  // Past Shifts overlay container
  const pastOverlay = document.createElement('div');
  pastOverlay.id = 'pastOverlay';
  pastOverlay.innerHTML = `
    <div id="pastSheet" role="dialog" aria-modal="true" aria-label="Past 14 days of shifts">
      <div id="pastHeader">
        <div id="pastTitle">Past 14 days</div>
        <button id="pastClose" aria-label="Close">✕</button>
      </div>
      <div id="pastList" tabindex="0"></div>
    </div>
  `;
  document.body.appendChild(pastOverlay);

  pastOverlay.addEventListener('click', (e) => {
    if (e.target === pastOverlay) closePastShifts(); // click backdrop to close
  });
  document.getElementById('pastClose').addEventListener('click', closePastShifts);
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && pastOverlay.classList.contains('show')) closePastShifts();
  });
}
function showLoading(msg = 'Loading...') {
  ensureLoadingOverlay();
  const overlay = document.getElementById('loadingOverlay');
  const text = document.getElementById('loadingText');
  if (text) text.textContent = msg;
  _loadingCount++;
  overlay.classList.remove('hidden');
}
function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (!overlay) return;
  _loadingCount = Math.max(0, _loadingCount - 1);
  if (_loadingCount === 0) overlay.classList.add('hidden');
}

// Install prompt (PWA)
let deferredPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredPrompt = e;
  els.installBtn && els.installBtn.classList.remove('hidden');
});
els.installBtn && els.installBtn.addEventListener('click', async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  els.installBtn.classList.add('hidden');
});

// Visibility handling for auto-clear after 3 minutes away
document.addEventListener('visibilitychange', () => {
  if (document.hidden) {
    lastHiddenAt = Date.now();
    persistDraft();
  } else {
    if (lastHiddenAt && Date.now() - lastHiddenAt > 3 * 60 * 1000) {
      // >3 minutes away → auto clear and reload latest
      if (Object.keys(draft).length) {
        draft = {};
        persistDraft();
        showToast('Unsaved changes were cleared after inactivity.');
      }
      showLoading();
      loadFromServer({ force: true })
        .catch(err => showToast('Reload failed: ' + err.message))
        .finally(hideLoading);
    } else {
      // Auto reload when returning if not making changes
      if (!Object.keys(draft).length) {
        showLoading();
        loadFromServer().catch(()=>{}).finally(hideLoading);
      }
    }
    lastHiddenAt = null;
  }
});

// Parse query
function parseQuery() {
  const q = new URLSearchParams(location.search);
  const k = q.get('k')?.trim();
  const msisdn = q.get('msisdn')?.trim();
  const t = q.get('t')?.trim(); // allow passing t in URL for testing
  // prefer k when present
  identity = k ? { k } : (msisdn ? { msisdn } : {});
  if (t) sessionStorage.setItem('api_t_override', t);
}

// Helpers
function statusClass(s) {
  switch (s) {
    case 'BOOKED': return 'status-booked';
    case 'BLOCKED': return 'status-blocked';
    case 'NOT AVAILABLE': return 'status-na';
    case 'LONG DAY': return 'status-ld';
    case 'NIGHT': return 'status-n';
    case 'LONG DAY/NIGHT': return 'status-ldn';
    default: return 'status-pending';
  }
}
function nextStatus(s) {
  const idx = STATUS_ORDER.indexOf(s);
  const next = STATUS_ORDER[(idx + 1) % STATUS_ORDER.length];
  return next;
}
function showFooterIfNeeded() {
  const dirty = Object.keys(draft).length > 0;
  els.footer.classList.toggle('hidden', !dirty);
  // When footer appears/disappears, recompute sizing (no-scroll layout)
  sizeGrid();
}
function showToast(msg, ms=2800) {
  if (!els.toast) return;
  els.toast.textContent = msg;
  els.toast.classList.remove('hidden');
  window.clearTimeout(showToast._t);
  showToast._t = window.setTimeout(() => els.toast.classList.add('hidden'), ms);
}
function persistDraft() {
  const save = { draft, identity, at: Date.now() };
  localStorage.setItem(DRAFT_KEY, JSON.stringify(save));
}
function loadDraft() {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (!raw) return;
    const { draft: d0, identity: i0 } = JSON.parse(raw);
    if (i0 && ((i0.k && identity.k && i0.k === identity.k) || (i0.msisdn && identity.msisdn && i0.msisdn === identity.msisdn))) {
      draft = d0 || {};
    } else {
      // different user -> ignore stale draft
      draft = {};
      localStorage.removeItem(DRAFT_KEY);
    }
  } catch { /* noop */ }
}
function saveLastLoaded(tsIso) {
  localStorage.setItem(LAST_LOADED_KEY, tsIso);
  if (!els.lastLoaded) return;
  if (tsIso) {
    const d = new Date(tsIso);
    els.lastLoaded.textContent = 'Updated ' + d.toLocaleString();
    els.lastLoaded.classList.remove('hidden');
  } else {
    els.lastLoaded.classList.add('hidden');
  }
}

// Read a CSS variable in px (fallback to 0 on parse issues)
function getCssVarPx(name) {
  const val = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const num = parseFloat(val);
  return Number.isFinite(num) ? num : 0;
}

/**
 * Compute and apply exact sizing so the grid occupies the viewport
 * without any scrollbars, regardless of footer visibility.
 * We publish --header-h, --footer-h, and --rows to CSS.
 */
function sizeGrid() {
  if (!els.grid) return;

  // Stabilize mobile address bar changes (optional)
  const vh = window.innerHeight * 0.01;
  document.documentElement.style.setProperty('--vh', `${vh}px`);

  // Measure header/footer heights
  const header = document.querySelector('header');
  const headerH = header ? Math.round(header.getBoundingClientRect().height) : 0;

  // If footer is hidden with display:none, its height will be 0 — fall back to CSS var --footer-h
  let footerH = 0;
  if (els.footer) {
    const rectH = Math.round(els.footer.getBoundingClientRect().height);
    footerH = rectH > 0 ? rectH : Math.round(getCssVarPx('--footer-h'));
  }

  // Publish to CSS so main/grid can size themselves
  document.documentElement.style.setProperty('--header-h', `${headerH}px`);
  document.documentElement.style.setProperty('--footer-h', `${footerH}px`);

  // Determine how many columns are actually in use, then rows
  const cs = getComputedStyle(els.grid);
  const templateCols = cs.gridTemplateColumns.split(' ').filter(Boolean);
  const cols = Math.max(1, templateCols.length);
  const numTiles = (baseline && baseline.tiles) ? baseline.tiles.length : 14;
  const rows = Math.max(1, Math.ceil(numTiles / cols));

  // Tell CSS the row count so it can compute exact row heights via grid-auto-rows
  els.grid.style.setProperty('--rows', rows);
}

/**
 * OPTIONAL: show a one-line instruction above the grid, if your HTML includes #helpMsg.
 * Falls back silently if you haven't added the element yet.
 */
function ensureHelpMessageVisible() {
  if (!els.helpMsg) return;
  els.helpMsg.textContent = "Please tap any date to change your availability — don't forget to tap Submit when done.";
  els.helpMsg.classList.remove('hidden');
}

// --------- Past Shifts (overlay + fetch) ----------
function ensurePastShiftsButton() {
  const header = document.querySelector('header');
  if (!header || document.getElementById('pastBtn')) return;
  const btn = document.createElement('button');
  btn.id = 'pastBtn';
  btn.textContent = 'Past Shifts';
  btn.setAttribute('aria-haspopup', 'dialog');
  header.insertBefore(btn, els.refreshBtn || header.lastChild);
  btn.addEventListener('click', openPastShifts);
}
function openPastShifts() {
  ensureLoadingOverlay();
  const overlay = document.getElementById('pastOverlay');
  const list = document.getElementById('pastList');
  if (!overlay || !list) return;

  list.innerHTML = '';
  overlay.classList.add('show');
  overlay.dataset.prevFocus = document.activeElement && document.activeElement.id;
  document.getElementById('pastClose').focus();

  fetchPastShifts()
    .then(items => {
      if (!items || !items.length) {
        list.innerHTML = `<div class="muted">No shifts in the last 14 days.</div>`;
        return;
      }
      const frag = document.createDocumentFragment();
      items.forEach(it => {
        const card = document.createElement('div');
        card.className = 'past-item';
        // Time line: Notes overrides; else include LONG/NIGHT with canonical times
        const timeLine = (it.notes && it.notes.trim())
          ? it.notes.trim()
          : ((it.shiftType === 'NIGHT') ? 'NIGHT 19:30–08:00' : 'LONG DAY 07:30–20:00');
        const roleLine = (it.jobTitle && it.jobTitle.trim()) ? `\nRole: ${it.jobTitle.trim()}` : '';
        const refLine = (it.bookingRef && it.bookingRef.trim())
          ? `\nRef: ${it.bookingRef.trim()}`
          : `\nAwaiting Booking Reference number`;

        card.innerHTML = `
          <div class="past-date">${it.dateDisplay}</div>
          <div>${timeLine}</div>
          <div>${it.hospital || ''}</div>
          <div>Ward: ${it.ward || ''}</div>
          ${roleLine ? `<div>${roleLine.slice(1)}</div>` : ``}
          <div>${refLine.startsWith('\n') ? refLine.slice(1) : refLine}</div>
        `;
        frag.appendChild(card);
      });
      list.appendChild(frag);
    })
    .catch(err => {
      list.innerHTML = `<div class="muted">Could not load past shifts: ${err.message || err}</div>`;
    });
}
function closePastShifts() {
  const overlay = document.getElementById('pastOverlay');
  if (!overlay) return;
  overlay.classList.remove('show');
  const prev = overlay.dataset.prevFocus;
  if (prev) {
    const el = document.getElementById(prev);
    if (el) el.focus();
  } else {
    const btn = document.getElementById('pastBtn');
    if (btn) btn.focus();
  }
}
async function fetchPastShifts() {
  if (!Object.keys(identity).length) throw new Error('Missing identity');
  const params = new URLSearchParams({ t: authToken(), view: 'past14' });
  if (identity.k) params.set('k', identity.k);
  if (identity.msisdn) params.set('msisdn', identity.msisdn);

  const url = `${API_BASE_URL}?${params.toString()}`;
  const res = await fetch(url, { method: 'GET', credentials: 'omit' });
  const json = await res.json().catch(() => ({}));
  if (!res.ok || !json || json.ok === false) {
    const msg = (json && json.error) || `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return json.items || [];
}

// Rendering
function renderTiles() {
  els.grid.innerHTML = '';
  if (!baseline || !baseline.tiles) return;

  // Ensure help message (if present in DOM) and Past Shifts button
  ensureHelpMessageVisible();
  ensurePastShiftsButton();

  const tiles = baseline.tiles.map(t => {
    const effectiveStatus = draft[t.ymd] ?? t.status; // show draft if exists
    return { ...t, effectiveStatus };
  });

  for (const t of tiles) {
    const card = document.createElement('div');
    card.className = 'tile';
    card.dataset.ymd = t.ymd;

    // Header: stacked day / date (weekday directly above date)
    const header = document.createElement('div');
    header.className = 'tile-header';
    const day = document.createElement('div');
    day.className = 'tile-day';
    day.textContent = t.displayDay;          // e.g., MON
    const date = document.createElement('div');
    date.className = 'tile-date';
    date.textContent = t.displayDate;        // e.g., 1 August
    header.append(day, date);

    const status = document.createElement('div');
    status.className = `tile-status ${statusClass(t.effectiveStatus)}`;

    // Status text rules:
    // - BOOKED: show "BOOKED"
    // - BLOCKED: show "BLOCKED"
    // - Non-booked availability:
    //     LONG DAY -> "LONG DAY 07:30–20:00"
    //     NIGHT -> "NIGHT 19:30–08:00"
    //     LONG DAY/NIGHT -> "LD 07:30–20:00 or N 19:30–08:00"
    //     others -> unchanged
    if (t.booked) {
      status.textContent = 'BOOKED';
    } else if (t.effectiveStatus === 'LONG DAY') {
      status.textContent = 'LONG DAY 07:30–20:00';
    } else if (t.effectiveStatus === 'NIGHT') {
      status.textContent = 'NIGHT 19:30–08:00';
    } else if (t.effectiveStatus === 'LONG DAY/NIGHT') {
      status.textContent = 'LD 07:30–20:00 or N 19:30–08:00';
    } else {
      status.textContent = t.effectiveStatus;
    }

    const sub = document.createElement('div');
    sub.className = 'tile-sub';

    if (t.booked) {
      // Notes override: if shiftInfo present and looks like Notes, use verbatim.
      // If shiftInfo is just "LONG DAY"/"NIGHT", show canonical line including label + times.
      const line1 = (() => {
        const si = (t.shiftInfo || '').trim();
        if (si && si.toUpperCase() !== 'LONG DAY' && si.toUpperCase() !== 'NIGHT') {
          return si; // treat as Notes verbatim
        }
        // fallback to canonical with label
        const shiftType = (t.shiftInfo || '').toUpperCase().includes('NIGHT') ? 'NIGHT' : 'LONG DAY';
        return shiftType === 'NIGHT' ? 'NIGHT 19:30–08:00' : 'LONG DAY 07:30–20:00';
      })();

      const hospital = (t.hospital || t.location || '').split(' – ')[0] || (t.hospital || t.location || '');
      const ward = (t.ward || (t.location && t.location.includes(' – ') ? t.location.split(' – ')[1] : '') || '').trim();
      const roleLine = (t.jobTitle && t.jobTitle.trim()) ? `Role: ${t.jobTitle.trim()}` : '';
      const refLine = (t.bookingRef && t.bookingRef.trim()) ? `Ref: ${t.bookingRef.trim()}` : '';

      // Build multi-line subtext with <br>, no layout changes
      const lines = [];
      lines.push(line1);
      if (hospital) lines.push(hospital);
      lines.push(`Ward: ${ward || ''}`);
      if (roleLine) lines.push(roleLine);
      if (refLine) lines.push(refLine);

      sub.innerHTML = lines.map(escapeHtml).join('<br>');
    } else if (t.effectiveStatus === 'BLOCKED') {
      // Show fixed explanatory line
      sub.textContent = 'You cannot work this shift as it will breach the 6 days in a row rule.';
    } else if (t.effectiveStatus !== t.status) {
      sub.innerHTML = `<span class="edit-note">Pending change</span>`;
    } else {
      sub.textContent = '';
    }

    card.append(header, status, sub);

    // Interactions
    const editable = (!t.booked && t.status !== 'BLOCKED' && t.editable !== false);
    if (editable) {
      card.style.cursor = 'pointer';
      card.title = 'Tap to change: Pending → N/A → LD → N → LD/N → Pending';
      card.addEventListener('click', () => {
        const cur = draft[t.ymd] ?? t.status;
        const next = nextStatus(cur);
        // update draft (store only diffs from baseline)
        if (next === t.status) {
          delete draft[t.ymd];
        } else {
          draft[t.ymd] = next;
        }
        persistDraft();
        renderTiles();
        showFooterIfNeeded();
      });
    } else {
      card.title = t.booked ? 'BOOKED (locked)' : (t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable');
      card.style.opacity = t.booked || t.status === 'BLOCKED' ? .9 : 1;
    }

    els.grid.appendChild(card);
  }

  // After tiles render, compute sizing so the grid fills the viewport precisely
  sizeGrid();
}
function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

// API
function authToken() {
  return sessionStorage.getItem('api_t_override') || API_SHARED_TOKEN;
}

async function loadFromServer({ force=false } = {}) {
  if (!Object.keys(identity).length) {
    throw new Error('Missing identity: provide ?k= or ?msisdn=');
  }

  showLoading();

  try {
    const params = new URLSearchParams({ t: authToken() });
    if (identity.k) params.set('k', identity.k);
    if (identity.msisdn) params.set('msisdn', identity.msisdn);

    const url = `${API_BASE_URL}?${params.toString()}`;
    const res = await fetch(url, { method: 'GET', credentials: 'omit' });
    const json = await res.json().catch(() => ({}));
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && json.error) || `HTTP ${res.status}`;
      throw new Error(msg);
    }

    baseline = json;

    // Display candidate name if provided (bigger and at the top in HTML)
    const name =
      (json.candidateName && String(json.candidateName).trim()) ||
      (json.candidate && [json.candidate.firstName, json.candidate.surname].filter(Boolean).join(' ').trim()) ||
      '';
    if (els.candidateName) {
      if (name) {
        els.candidateName.textContent = name;
        els.candidateName.classList.remove('hidden');
      } else {
        els.candidateName.textContent = '';
        els.candidateName.classList.add('hidden');
      }
    }

    saveLastLoaded(json.lastLoadedAt || new Date().toISOString());

    // keep only draft keys that still exist in 14-day window
    const validYmd = new Set(baseline.tiles.map(x => x.ymd));
    for (const k of Object.keys(draft)) {
      if (!validYmd.has(k)) delete draft[k];
    }

    renderTiles();
    showFooterIfNeeded();
  } finally {
    hideLoading();
  }
}

async function submitChanges() {
  if (!baseline) return;

  const changes = [];
  for (const [ymd, newStatus] of Object.entries(draft)) {
    const code = STATUS_TO_CODE[newStatus];
    if (code === undefined) continue;
    changes.push({ ymd, code });
  }
  if (!changes.length) {
    showToast('No changes to submit.');
    return;
  }

  showLoading('Saving.....');

  try {
    const body = {
      t: authToken(),
      changes
    };
    if (identity.k) body.k = identity.k;
    if (identity.msisdn) body.msisdn = identity.msisdn;

    // IMPORTANT: use a "simple" request to avoid CORS preflight (no custom headers beyond simple types)
    const res = await fetch(API_BASE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      credentials: 'omit',
      body: JSON.stringify(body)
    });

    const text = await res.text();
    let json = {};
    try { json = JSON.parse(text); } catch { /* keep text for diagnostics */ }

    if (!res.ok || !json || json.ok === false) {
      const msg = (json && json.error) || `HTTP ${res.status}`;
      const extra = text && text !== '{}' ? ` – ${text.slice(0,200)}` : '';
      throw new Error(msg + extra);
    }

    // Summarise results (server may reject some ymds e.g., BOOKED/BLOCKED)
    const applied = (json.results || []).filter(r => r.applied).length;
    const rejected = (json.results || []).filter(r => !r.applied);
    if (rejected.length) {
      const reasons = rejected.slice(0,3).map(r => `${r.ymd}: ${r.reason}`).join(', ');
      showToast(`Saved ${applied}. Skipped ${rejected.length} (${reasons}${rejected.length>3?', …':''}).`, 4200);
    } else {
      showToast(`Saved ${applied} change${applied===1?'':'s'}.`);
    }

    // Clear draft & reload fresh baseline
    draft = {};
    persistDraft();

    // NOTE: Do NOT call showLoading() here (avoids double spinner).
    await loadFromServer({ force: true });
  } catch (err) {
    showToast('Submit failed: ' + (err.message || err));
  } finally {
    hideLoading();
  }
}

function clearChanges() {
  draft = {};
  persistDraft();
  renderTiles();
  showFooterIfNeeded();
  showToast('Cleared pending changes.');
}

// UI events
els.submitBtn && els.submitBtn.addEventListener('click', () => {
  els.submitBtn.disabled = true;
  submitChanges()
    .finally(() => { els.submitBtn.disabled = false; });
});
els.clearBtn && els.clearBtn.addEventListener('click', clearChanges);
els.refreshBtn && els.refreshBtn.addEventListener('click', async () => {
  if (Object.keys(draft).length) {
    if (!confirm('You have unsaved changes. Reload and lose them?')) return;
    clearChanges();
  }
  showLoading();
  try {
    await loadFromServer({ force: true });
  } catch (e) {
    showToast('Reload failed: ' + e.message);
  } finally {
    hideLoading();
  }
});

// Resize/orientation handlers to keep grid filling the screen
window.addEventListener('resize', sizeGrid);
window.addEventListener('orientationchange', () => {
  // wait a tick for viewport to settle
  setTimeout(sizeGrid, 250);
});

// Boot
(async function init() {
  parseQuery();
  loadDraft();
  showLoading(); // show immediately on first load
  try {
    await loadFromServer({ force: true });
  } catch (e) {
    showToast('Load failed: ' + e.message, 5000);
  } finally {
    // Ensure initial sizing even if load fails
    sizeGrid();
    hideLoading();
    // Ensure overlays and Past Shifts button are wired
    ensureLoadingOverlay();
    ensurePastShiftsButton();
  }
})();
