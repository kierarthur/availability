/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL     = 'https://script.google.com/macros/s/AKfycbxEZIaMHuP6-QUe7ITk68JNLEE9W-DlYhs-BWimkYyqmPHk0z3R7ZD9OrxNbEf7lgo2Qw/exec'; // <-- REPLACE
const API_SHARED_TOKEN = 't9x_93HDa8nL0PQ6RvzX4wqZ'; // <-- REPLACE
/* ===================================================== */

const PENDING_LABEL_DEFAULT = 'PLEASE PROVIDE YOUR AVAILABILITY';
const PENDING_LABEL_WRAPPED = 'NOT SURE OF MY AVAILABILITY YET';

const STATUS_ORDER = [
  PENDING_LABEL_DEFAULT,
  'NOT AVAILABLE',
  'LONG DAY',
  'NIGHT',
  'LONG DAY/NIGHT'
];
const STATUS_TO_CODE = {
  [PENDING_LABEL_DEFAULT]: '',     // blank (white)
  'NOT AVAILABLE':        'N/A',   // red
  'LONG DAY':             'LD',    // yellow
  'NIGHT':                'N',     // yellow
  'LONG DAY/NIGHT':       'LD/N'   // yellow
};
const CODE_TO_STATUS = Object.entries(STATUS_TO_CODE).reduce((acc,[k,v]) => (acc[v]=k, acc), {});

const els = {
  grid: document.getElementById('grid'),
  footer: document.getElementById('footer'),
  draftMsg: document.getElementById('draftMsg'),
  submitBtn: document.getElementById('submitBtn'),
  clearBtn: document.getElementById('clearBtn'),
  refreshBtn: document.getElementById('refreshBtn'),
  lastLoaded: document.getElementById('lastLoaded'), // optional in HTML
  toast: document.getElementById('toast'),
  installBtn: document.getElementById('installBtn'),
  candidateName: document.getElementById('candidateName'),
  helpMsg: document.getElementById('helpMsg')
};

// ===== Auth guard (fatal overlay) =====
let AUTH_DENIED = false;
function showAuthError(message = 'Not an authorised user') {
  if (AUTH_DENIED) return;
  AUTH_DENIED = true;

  try { els.footer && els.footer.classList.add('hidden'); } catch {}
  try { els.grid && (els.grid.innerHTML = ''); } catch {}
  try { els.helpMsg && els.helpMsg.classList.add('hidden'); } catch {}
  try { els.submitBtn && (els.submitBtn.disabled = true); } catch {}
  try { els.clearBtn && (els.clearBtn.disabled = true); } catch {}
  try { els.refreshBtn && (els.refreshBtn.disabled = true); } catch {}

  hideLoading();

  const div = document.createElement('div');
  div.id = 'authErrorOverlay';
  div.setAttribute('role', 'alertdialog');
  div.setAttribute('aria-modal', 'true');
  Object.assign(div.style, {
    position:'fixed', inset:'0', zIndex:'10000',
    display:'flex', alignItems:'center', justifyContent:'center',
    background:'rgba(10,12,16,0.85)', backdropFilter:'blur(2px)'
  });
  div.innerHTML = `
    <div style="
      border:1px solid #2a3446;
      background:#131926;
      color:#e7ecf3;
      padding:1rem 1.2rem;
      border-radius:12px;
      font-weight:800;
      font-family: system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
      box-shadow:0 12px 30px rgba(0,0,0,.4);
      max-width: 90vw;
      text-align:center;
    ">
      ${message}
    </div>
  `;
  document.body.appendChild(div);
}

// ---------- State ----------
let baseline = null;          // server response (tiles, lastLoadedAt, candidateName, candidate, newUserHint?, hadK?)
let draft = {};               // ymd -> status LABEL (only diffs from baseline). Pending stored as BASE label.
let identity = {};            // { k?:string, msisdn?:string }
let lastHiddenAt = null;      // timestamp when page becomes hidden
const DRAFT_KEY = 'rota_avail_draft_v1';
const LAST_LOADED_KEY = 'rota_avail_last_loaded_v1';
const SAVED_IDENTITY_KEY = 'rota_avail_identity_v1'; // persisted identity
const SAVED_TOKEN_KEY = 'rota_avail_token_v1';       // persisted token (for iOS A2HS)

// Per-session memory of tiles the user has tapped at least once this session.
// Used to decide whether the pending wording should be "wrapped" (NOT SURE...).
const toggledThisSession = new Set();

// Submit-nudge timer state (1-minute reminder)
let submitNudgeTimer = null;
let lastEditAt = 0;
let nudgeShown = false;

// ---------- Simple API helpers ----------
async function apiGET(paramsObj) {
  const params = new URLSearchParams({ t: authToken(), ...(paramsObj || {}) });
  if (identity.k) params.set('k', identity.k);
  if (identity.msisdn) params.set('msisdn', identity.msisdn);
  const url = `${API_BASE_URL}?${params.toString()}`;
  const res = await fetch(url, { method: 'GET', credentials: 'omit' });
  const text = await res.text();
  let json = {};
  try { json = JSON.parse(text); } catch {}
  return { res, json, text };
}
async function apiPOST(bodyObj) {
  const body = { t: authToken(), ...(bodyObj || {}) };
  if (identity.k) body.k = identity.k;
  if (identity.msisdn) body.msisdn = identity.msisdn;
  const res = await fetch(API_BASE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // simple request (no preflight)
    credentials: 'omit',
    body: JSON.stringify(body)
  });
  const text = await res.text();
  let json = {};
  try { json = JSON.parse(text); } catch {}
  return { res, json, text };
}

// Saved-identity helpers
function saveIdentity(id) {
  try {
    if (id && (id.k || id.msisdn)) {
      localStorage.setItem(SAVED_IDENTITY_KEY, JSON.stringify(id));
    }
  } catch {}
}
function loadSavedIdentity() {
  try {
    const raw = localStorage.getItem(SAVED_IDENTITY_KEY);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    if (obj && (obj.k || obj.msisdn)) return obj;
  } catch {}
  return {};
}
function clearSavedIdentity() {
  try { localStorage.removeItem(SAVED_IDENTITY_KEY); } catch {}
  try { localStorage.removeItem(SAVED_TOKEN_KEY); } catch {}
  try { sessionStorage.removeItem('api_t_override'); } catch {}
}

// ------- Loading overlay + shared styles -------
let _loadingCount = 0;
function ensureLoadingOverlay() {
  if (document.getElementById('loadingOverlay')) return;

  const style = document.createElement('style');
  style.textContent = `
    #loadingOverlay {
      position: fixed; inset: 0; z-index: 9999;
      display: flex; align-items: center; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
      color: #e7ecf3; font-family: system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
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
    #pastOverlay, #contentOverlay, #tokenOverlay, #welcomeOverlay {
      position: fixed; inset: 0; z-index: 9998;
      display: none; align-items: stretch; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
    }
    #pastOverlay.show, #contentOverlay.show, #tokenOverlay.show, #welcomeOverlay.show { display: flex; }
    .sheet {
      background: #0f1115; color: #e7ecf3;
      border-top-left-radius: 16px; border-top-right-radius: 16px;
      width: min(820px, 96vw);
      margin-top: min(8vh, 80px);
      margin-bottom: 0;
      box-shadow: 0 10px 30px rgba(0,0,0,.45);
      display: flex; flex-direction: column; max-height: calc(100vh - 12vh);
      border: 1px solid #222936;
    }
    .sheet-header {
      display:flex; align-items:center; gap:.6rem; padding:.7rem .9rem;
      border-bottom:1px solid #222936; background: rgba(15,17,21,.9); backdrop-filter: blur(6px);
    }
    .sheet-title { font-weight:800; font-size:1.05rem; margin:0; }
    .sheet-close {
      margin-left:auto; appearance:none; border:0; background:transparent; color:#a7b0c0;
      font-weight:700; font-size:1rem; cursor:pointer; padding:.25rem .5rem; border-radius:8px;
    }
    .sheet-close:hover { background:#1e2532; color:#fff; }
    .sheet-body {
      overflow:auto; padding:.7rem .9rem 1rem; display:flex; flex-direction:column; gap:.6rem;
    }
    .past-item {
      border:1px solid #222a36; background:#131926; border-radius:12px; padding:.6rem .7rem;
      line-height:1.35;
    }
    .past-date { font-weight:800; margin-bottom:.15rem; color:#e7ecf3; }
    .muted { color:#a7b0c0; }

    /* Menu dropdown */
    #menuWrap { position: relative; margin-left: auto; }
    #menuBtn {
      appearance: none; border: 1px solid #2a3446; background: #131926; color: #e7ecf3;
      border-radius: 8px; padding: .35rem .6rem; cursor: pointer; font-weight: 700;
    }
    #menuList {
      position: absolute; right: 0; top: calc(100% + 6px); z-index: 9997;
      background: #0f1115; border: 1px solid #222936; border-radius: 10px;
      min-width: 220px; padding: .35rem; display: none;
      box-shadow: 0 10px 24px rgba(0,0,0,.45);
    }
    #menuList.show { display: block; }
    .menu-item {
      display: block; width: 100%; text-align: left; background: transparent; border: 0;
      color: #e7ecf3; padding: .45rem .6rem; border-radius: 8px; cursor: pointer;
      font-weight: 600;
    }
    .menu-item:hover { background: #1b2331; }

    /* Attention pulses (respect reduced motion) */
    @media (prefers-reduced-motion: no-preference) {
      @keyframes softPulse {
        0%   { opacity: 1; transform: scale(1); }
        50%  { opacity: .78; transform: scale(0.99); }
        100% { opacity: 1; transform: scale(1); }
      }
      @keyframes strongPulse {
        0%   { opacity: 1; transform: scale(1); }
        50%  { opacity: .72; transform: scale(1.03); }
        100% { opacity: 1; transform: scale(1); }
      }
      .needs-attention { animation: softPulse 1.6s ease-in-out infinite; }
      .btn-attention   { animation: strongPulse 1.0s ease-in-out infinite; }
    }

    /* Thin red border for attention tiles */
    .attention-border { border-color: var(--danger) !important; }
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
    <div id="pastSheet" class="sheet" role="dialog" aria-modal="true" aria-label="Past 14 days of shifts">
      <div class="sheet-header">
        <div class="sheet-title">Past 14 days</div>
        <button id="pastClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div id="pastList" class="sheet-body" tabindex="0"></div>
    </div>
  `;
  document.body.appendChild(pastOverlay);

  // Content overlay for server HTML (addresses, accommodation)
  const contentOverlay = document.createElement('div');
  contentOverlay.id = 'contentOverlay';
  contentOverlay.innerHTML = `
    <div id="contentSheet" class="sheet" role="dialog" aria-modal="true" aria-label="Information">
      <div class="sheet-header">
        <div id="contentTitle" class="sheet-title">Info</div>
        <button id="contentClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div id="contentBody" class="sheet-body" tabindex="0"></div>
    </div>
  `;
  document.body.appendChild(contentOverlay);

  // Token claim overlay (iPhone paste flow)
  const tokenOverlay = document.createElement('div');
  tokenOverlay.id = 'tokenOverlay';
  tokenOverlay.innerHTML = `
    <div id="tokenSheet" class="sheet" role="dialog" aria-modal="true" aria-label="iPhone Token Claim">
      <div class="sheet-header">
        <div class="sheet-title">Unlock your link on iPhone</div>
        <button id="tokenClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div class="sheet-body">
        <p>If you opened this from an iPhone “Add to Home Screen” app, the security token may be missing.</p>
        <p>Please paste the <strong>full URL</strong> from the secure text/email you received:</p>
        <textarea id="tokenTextarea" rows="4" style="width:100%;resize:vertical;border-radius:8px;border:1px solid #222936;background:#0b0e14;color:#e7ecf3;padding:.6rem;"></textarea>
        <div style="display:flex;gap:.6rem;margin-top:.6rem;">
          <button id="tokenSubmit" class="menu-item" style="border:1px solid #2a3446;">Claim Token</button>
          <button id="tokenHelp" class="menu-item" style="border:1px solid #2a3446;">What is this?</button>
        </div>
        <div id="tokenError" class="muted" role="alert" style="margin-top:.4rem;"></div>
      </div>
    </div>
  `;
  document.body.appendChild(tokenOverlay);

  // Welcome/Alert overlay for newUserHint
  const welcomeOverlay = document.createElement('div');
  welcomeOverlay.id = 'welcomeOverlay';
  welcomeOverlay.innerHTML = `
    <div id="welcomeSheet" class="sheet" role="dialog" aria-modal="true" aria-label="Welcome">
      <div class="sheet-header">
        <div id="welcomeTitle" class="sheet-title">Welcome</div>
        <button id="welcomeClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div id="welcomeBody" class="sheet-body" tabindex="0"></div>
      <div style="display:flex;justify-content:flex-end;gap:.5rem;padding:.5rem .9rem 1rem;">
        <button id="welcomeDismiss" class="menu-item" style="border:1px solid #2a3446;">Got it</button>
      </div>
    </div>
  `;
  document.body.appendChild(welcomeOverlay);

  // Shared overlay wiring (click-outside & esc)
  ;[
    { overlayId:'pastOverlay', closeId:'pastClose' },
    { overlayId:'contentOverlay', closeId:'contentClose' },
    { overlayId:'tokenOverlay', closeId:'tokenClose' },
    { overlayId:'welcomeOverlay', closeId:'welcomeClose' },
  ].forEach(({overlayId, closeId}) => {
    const overlay = document.getElementById(overlayId);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeOverlay(overlayId); });
    document.getElementById(closeId).addEventListener('click', () => closeOverlay(overlayId));
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      ['pastOverlay','contentOverlay','tokenOverlay','welcomeOverlay'].forEach(id => {
        const ov = document.getElementById(id);
        if (ov && ov.classList.contains('show')) closeOverlay(id);
      });
    }
  });
}
function openOverlay(overlayId, focusSel) {
  ensureLoadingOverlay();
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;
  overlay.classList.add('show');
  overlay.dataset.prevFocus = document.activeElement && document.activeElement.id ? document.activeElement.id : '';
  const focusEl = focusSel ? overlay.querySelector(focusSel) : overlay.querySelector('.sheet-close');
  focusEl && focusEl.focus();
  trapFocus(overlay);
}
function closeOverlay(overlayId) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;
  overlay.classList.remove('show');
  releaseFocusTrap(overlay);
  const prev = overlay.dataset.prevFocus;
  if (prev) {
    const el = document.getElementById(prev);
    if (el) el.focus();
  }
}
function trapFocus(container) {
  const selectors = [
    'a[href]','area[href]','input:not([disabled])','select:not([disabled])','textarea:not([disabled])',
    'button:not([disabled])','iframe','[tabindex]:not([tabindex="-1"])','[contenteditable="true"]'
  ];
  const focusables = Array.from(container.querySelectorAll(selectors.join(','))).filter(el => el.offsetParent !== null);
  if (!focusables.length) return;
  function loop(e) {
    if (e.key !== 'Tab') return;
    const first = focusables[0];
    const last = focusables[focusables.length - 1];
    if (e.shiftKey) {
      if (document.activeElement === first) { last.focus(); e.preventDefault(); }
    } else {
      if (document.activeElement === last) { first.focus(); e.preventDefault(); }
    }
  }
  container._focusHandler = loop;
  container.addEventListener('keydown', loop);
}
function releaseFocusTrap(container) {
  if (container && container._focusHandler) {
    container.removeEventListener('keydown', container._focusHandler);
    container._focusHandler = null;
  }
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

// ---------- PWA install prompt ----------
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

// ---------- Visibility handling (auto-clear after 3 minutes away) ----------
document.addEventListener('visibilitychange', () => {
  if (document.hidden) {
    // pause submit-nudge checks while hidden
    cancelSubmitNudge();
    lastHiddenAt = Date.now();
    persistDraft();
  } else {
    if (lastHiddenAt && Date.now() - lastHiddenAt > 3 * 60 * 1000) {
      if (Object.keys(draft).length) {
        draft = {};
        persistDraft();
        showToast('Unsaved changes were cleared after inactivity.');
      }
      showLoading();
      loadFromServer({ force: true })
        .catch(err => { if (!AUTH_DENIED) showToast('Reload failed: ' + err.message); })
        .finally(hideLoading);
    } else if (!Object.keys(draft).length) {
      showLoading();
      loadFromServer().catch(()=>{}).finally(hideLoading);
    }
    // If still dirty and no reminder yet, resume countdown
    if (Object.keys(draft).length) scheduleSubmitNudge();
    lastHiddenAt = null;
  }
});

// ---------- Query parsing ----------
function parseQuery() {
  // Read from both search and hash to support iOS A2HS which often preserves # better than ?.
  const search = new URLSearchParams(location.search);
  const hash = new URLSearchParams(location.hash && location.hash.startsWith('#') ? location.hash.slice(1) : '');

  const pick = (key) => (search.get(key)?.trim() || hash.get(key)?.trim() || '');

  const t = pick('t');
  const k = pick('k');
  const msisdn = pick('msisdn');

  // Persist token in BOTH sessionStorage and localStorage (iOS A2HS needs localStorage on first launch).
  if (t) {
    try { sessionStorage.setItem('api_t_override', t); } catch {}
    try { localStorage.setItem(SAVED_TOKEN_KEY, t); } catch {}
  }

  if (k || msisdn) {
    identity = k ? { k } : { msisdn };
    saveIdentity(identity); // persist latest identity
  } else {
    // restore from saved identity when URL has none
    const saved = loadSavedIdentity();
    identity = (saved.k || saved.msisdn) ? saved : {};
  }
}

// Ensure the current URL (especially when user taps "Add to Home Screen") carries what we need.
function ensureCanonicalURLWithToken() {
  try {
    const token = authToken();
    const params = new URLSearchParams(location.hash && location.hash.startsWith('#') ? location.hash.slice(1) : '');
    let changed = false;
    if (token && params.get('t') !== token) { params.set('t', token); changed = true; }
    if (identity.k && params.get('k') !== identity.k) { params.set('k', identity.k); changed = true; }
    if (identity.msisdn && params.get('msisdn') !== identity.msisdn) { params.set('msisdn', identity.msisdn); changed = true; }
    if (changed) {
      history.replaceState(null, '', location.pathname + location.search + '#' + params.toString());
    }
  } catch {}
}

// ---------- Helpers ----------
function statusClass(label) {
  switch (label) {
    case 'BOOKED': return 'status-booked';
    case 'BLOCKED': return 'status-blocked';
    case 'NOT AVAILABLE': return 'status-na';
    case 'LONG DAY': return 'status-ld';
    case 'NIGHT': return 'status-n';
    case 'LONG DAY/NIGHT': return 'status-ldn';
    default: return 'status-pending';
  }
}
function nextStatus(currentLabel) {
  const idx = STATUS_ORDER.indexOf(currentLabel);
  return STATUS_ORDER[(idx + 1) % STATUS_ORDER.length];
}
function showFooterIfNeeded() {
  const dirty = Object.keys(draft).length > 0;
  els.footer.classList.toggle('hidden', !dirty);

  // Submit button pulse when dirty (stronger blink)
  if (els.submitBtn) {
    if (dirty) els.submitBtn.classList.add('btn-attention');
    else els.submitBtn.classList.remove('btn-attention');
  }

  sizeGrid(); // harmless with scrolling layout; keeps CSS vars fresh
}
function showToast(msg, ms=2800) {
  if (!els.toast) return;
  els.toast.textContent = msg;
  els.toast.classList.remove('hidden');
  els.toast.setAttribute('aria-live','polite');
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
      draft = {};
      localStorage.removeItem(DRAFT_KEY);
    }
  } catch {}
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
function getCssVarPx(name) {
  const val = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  const num = parseFloat(val);
  return Number.isFinite(num) ? num : 0;
}

// ---- Code ↔ Label normalization (CRITICAL FIX) ----
function codeToLabel(v) {
  // Already a known label?
  if (v && STATUS_TO_CODE.hasOwnProperty(v)) return v;
  if (v === 'BOOKED' || v === 'BLOCKED') return v;
  // Map codes ('' | 'N/A' | 'LD' | 'N' | 'LD/N') → labels
  if (v === undefined || v === null) return PENDING_LABEL_DEFAULT;
  const mapped = CODE_TO_STATUS[v];
  return mapped !== undefined ? mapped : PENDING_LABEL_DEFAULT;
}

/**
 * Decide which pending wording to display for a tile.
 * - If effective label is pending:
 *   - If baseline was pending AND user hasn't tapped this session → base wording
 *   - Else → wrapped wording
 * - Otherwise return the effective non-pending label
 */
function displayPendingLabelForTile(t) {
  if (t.effectiveLabel !== PENDING_LABEL_DEFAULT) return t.effectiveLabel;
  const serverPending = (t.baselineLabel === PENDING_LABEL_DEFAULT);
  const touched = toggledThisSession.has(t.ymd);
  return (serverPending && !touched) ? PENDING_LABEL_DEFAULT : PENDING_LABEL_WRAPPED;
}

/**
 * Build the cycle hint. When we show pending in the sequence, use the visual that will appear.
 */
function buildCycleHint(t) {
  const pendingVisual =
    (t.baselineLabel === PENDING_LABEL_DEFAULT && !toggledThisSession.has(t.ymd))
      ? PENDING_LABEL_DEFAULT
      : PENDING_LABEL_WRAPPED;
  const seq = [pendingVisual, 'N/A', 'LD', 'N', 'LD/N', pendingVisual];
  return 'Tap to change: ' + seq.join(' → ');
}

/**
 * Fit status text: no mid-word breaks, allow wrapping between words, and
 * only shrink if a single word would overflow even on its own line.
 */
function fitStatusLabel(el) {
  if (!el) return;
  el.style.wordBreak = 'keep-all';
  el.style.overflowWrap = 'normal';
  el.style.fontSize = '';

  const containerWidth = el.clientWidth;
  if (!containerWidth) return;

  const MIN_SCALE = 0.85;
  let scale = 1.0;

  const text = el.textContent || '';
  const words = text.split(/\s+/).filter(Boolean);

  if (el.scrollWidth <= el.clientWidth) return;

  const measurer = document.createElement('span');
  Object.assign(measurer.style, {
    visibility:'hidden', whiteSpace:'nowrap', position:'absolute',
    pointerEvents:'none', wordBreak:'keep-all', overflowWrap:'normal'
  });
  el.appendChild(measurer);

  const baseSize = parseFloat(getComputedStyle(el).fontSize) || 16;

  const longestWordWidth = () => {
    let max = 0;
    for (const w of words) {
      measurer.textContent = w;
      const wWidth = measurer.getBoundingClientRect().width;
      if (wWidth > max) max = wWidth;
    }
    return max;
  };

  let longest = longestWordWidth();
  while (longest > containerWidth && scale > MIN_SCALE + 0.001) {
    scale = Math.max(MIN_SCALE, scale - 0.05);
    el.style.fontSize = (baseSize * scale) + 'px';
    longest = longestWordWidth();
  }
  measurer.remove();
}

// ---------- Submit nudge (1-minute reminder) ----------
function cancelSubmitNudge() {
  if (submitNudgeTimer) {
    clearTimeout(submitNudgeTimer);
    submitNudgeTimer = null;
  }
}
function scheduleSubmitNudge() {
  if (nudgeShown) return;
  if (!Object.keys(draft).length) return;
  cancelSubmitNudge();
  const elapsed = Date.now() - (lastEditAt || Date.now());
  const remaining = Math.max(0, 60_000 - elapsed);
  submitNudgeTimer = window.setTimeout(() => {
    if (document.visibilityState === 'visible' && Object.keys(draft).length) {
      showToast('You have unsaved changes — tap Submit to save.', 4200);
      nudgeShown = true;
    }
  }, remaining);
}
function registerUserEdit() {
  lastEditAt = Date.now();
  if (!nudgeShown) scheduleSubmitNudge();
}

// ---------- Help banner ----------
function ensureHelpMessageVisible() {
  if (!els.helpMsg) return;
  els.helpMsg.textContent = "Please tap any date to change your availability — don't forget to tap Submit when done.";
  els.helpMsg.classList.remove('hidden');
}

// ---------- Past Shifts (overlay + fetch) ----------
function ensurePastShiftsButton() {
  const header = document.querySelector('header');
  if (!header) return;
  // Ensure menu exists (top-right actions)
  ensureMenu();

  // Also keep the dedicated Past Shifts button if desired (optional)
  if (!document.getElementById('pastBtn')) {
    const btn = document.createElement('button');
    btn.id = 'pastBtn';
    btn.textContent = 'Past Shifts';
    btn.setAttribute('aria-haspopup', 'dialog');
    header.insertBefore(btn, els.refreshBtn || header.lastChild);
    btn.addEventListener('click', openPastShifts);
  }
}
function openPastShifts() {
  ensureLoadingOverlay();
  const list = document.getElementById('pastList');
  if (!list) return;

  list.innerHTML = '';
  openOverlay('pastOverlay', '#pastClose');

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
async function fetchPastShifts() {
  if (!Object.keys(identity).length) throw new Error('Missing identity');
  const { res, json } = await apiGET({ view: 'past14' });
  if (!res.ok || !json || json.ok === false) {
    const msg = (json && json.error) || `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return json.items || [];
}

// ---------- Content overlay (Addresses / Accommodation) ----------
async function openContent(kind, titleFallback) {
  ensureLoadingOverlay();
  const titleEl = document.getElementById('contentTitle');
  const bodyEl  = document.getElementById('contentBody');
  if (!titleEl || !bodyEl) return;
  bodyEl.innerHTML = `<div class="muted">Loading...</div>`;
  titleEl.textContent = titleFallback || 'Information';
  openOverlay('contentOverlay', '#contentClose');

  try {
    const { res, json } = await apiGET({ view:'content', kind });
    if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
      bodyEl.innerHTML = `<div class="muted">Server is busy, please try again shortly.</div>`;
      return;
    }
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && json.error) || `HTTP ${res.status}`;
      bodyEl.innerHTML = `<div class="muted">Could not load content: ${msg}</div>`;
      return;
    }
    titleEl.textContent = json.title || titleFallback || 'Information';
    // Server returns sanitized HTML
    bodyEl.innerHTML = json.html || '<div class="muted">No content.</div>';
  } catch (e) {
    bodyEl.innerHTML = `<div class="muted">Could not load content: ${e.message || e}</div>`;
  }
}

// ---------- Menu (top-right) ----------
function ensureMenu() {
  const header = document.querySelector('header');
  if (!header) return;
  if (document.getElementById('menuWrap')) return;

  const wrap = document.createElement('div');
  wrap.id = 'menuWrap';
  const btn = document.createElement('button');
  btn.id = 'menuBtn';
  btn.type = 'button';
  btn.textContent = 'Menu';
  const list = document.createElement('div');
  list.id = 'menuList';
  list.innerHTML = `
    <button class="menu-item" id="miPast">Past Shifts</button>
    <button class="menu-item" id="miHosp">Hospital Addresses</button>
    <button class="menu-item" id="miAccom">Accommodation Contacts</button>
    <button class="menu-item" id="miTimesheet">Send Timesheet</button>
    <button class="menu-item" id="miLogout">Log out</button>
  `;
  wrap.append(btn, list);
  header.appendChild(wrap);

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    list.classList.toggle('show');
  });
  document.addEventListener('click', () => list.classList.remove('show'));
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') list.classList.remove('show');
  });

  document.getElementById('miPast').addEventListener('click', () => { list.classList.remove('show'); openPastShifts(); });
  document.getElementById('miHosp').addEventListener('click', () => { list.classList.remove('show'); openContent('HOSPITAL','Hospital Addresses'); });
  document.getElementById('miAccom').addEventListener('click', () => { list.classList.remove('show'); openContent('ACCOMMODATION','Accommodation Contacts'); });
  document.getElementById('miTimesheet').addEventListener('click', async () => {
    list.classList.remove('show');
    try {
      showLoading('Sending timesheet...');
      const { res, json } = await apiPOST({ action:'SEND_TIMESHEET' });
      if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
        showToast('Server is busy, please try again.');
        return;
      }
      if (!res.ok || !json || json.ok === false) {
        const err = (json && json.error) || `HTTP ${res.status}`;
        showToast(err);
        return;
      }
      showToast('Timesheet sent.');
    } catch (e) {
      showToast('Failed to send: ' + (e.message || e));
    } finally { hideLoading(); }
  });
  document.getElementById('miLogout').addEventListener('click', () => {
    list.classList.remove('show');
    clearSavedIdentity();
    showAuthError('Logged out. Please reopen your secure link.');
  });
}

// ---------- Welcome / newUserHint ----------
function maybeShowWelcome() {
  if (!baseline || !baseline.newUserHint) return;
  const hint = baseline.newUserHint; // { title, html }
  const titleEl = document.getElementById('welcomeTitle');
  const bodyEl  = document.getElementById('welcomeBody');
  if (!titleEl || !bodyEl) return;

  titleEl.textContent = (hint.title || 'Welcome').trim();
  bodyEl.innerHTML = hint.html || '';
  openOverlay('welcomeOverlay', '#welcomeDismiss');

  const dismiss = document.getElementById('welcomeDismiss');
  dismiss.onclick = async () => {
    try {
      await apiPOST({ action:'MARK_MESSAGE_SEEN' });
    } catch {}
    closeOverlay('welcomeOverlay');
  };
}

// ---------- Token claim (iPhone paste full URL flow) ----------
function openTokenClaimDialog(prefillText='') {
  const textarea = document.getElementById('tokenTextarea');
  const errorEl  = document.getElementById('tokenError');
  const helpBtn  = document.getElementById('tokenHelp');
  const submit   = document.getElementById('tokenSubmit');
  if (!textarea || !errorEl || !helpBtn || !submit) return;
  textarea.value = prefillText;
  errorEl.textContent = '';
  openOverlay('tokenOverlay', '#tokenTextarea');

  helpBtn.onclick = () => {
    errorEl.innerHTML = 'Tip: Open the original secure SMS or email, press and hold the link, choose “Copy”, then paste it above.';
  };
  submit.onclick = async () => {
    const userText = textarea.value.trim();
    if (!userText) {
      errorEl.textContent = 'Please paste the full secure URL first.';
      return;
    }
    try {
      showLoading('Claiming token...');
      const { res, json } = await apiPOST({ action:'TOKEN_CLAIM', userText });
      if (!res.ok || !json || json.ok === false) {
        const msg = (json && json.error) || `HTTP ${res.status}`;
        if (msg === 'TOKEN_NOT_FOUND_IN_TEXT') {
          errorEl.textContent = 'We could not find a token in that text. Please paste the full URL.';
        } else if (msg === 'TOKEN_NOT_ACTIVE_OR_UNKNOWN') {
          errorEl.textContent = 'That token is not active or unknown. Please check the link and try again.';
        } else {
          errorEl.textContent = msg;
        }
        return;
      }
      // Success: server returns { ok:true, token, msisdn }
      if (json.token) {
        try { localStorage.setItem(SAVED_TOKEN_KEY, json.token); } catch {}
        try { sessionStorage.setItem('api_t_override', json.token); } catch {}
      }
      identity = { msisdn: json.msisdn };
      saveIdentity(identity);

      // Clear any bad k
      const params = new URLSearchParams(location.hash && location.hash.startsWith('#') ? location.hash.slice(1) : '');
      params.delete('k');
      params.set('t', authToken());
      params.set('msisdn', identity.msisdn);
      history.replaceState(null, '', location.pathname + location.search + '#' + params.toString());

      closeOverlay('tokenOverlay');

      // Reload tiles with new token
      loadFromServer({ force: true }).catch(e => showToast(e.message || e));
    } catch (e) {
      errorEl.textContent = e.message || String(e);
    } finally { hideLoading(); }
  };
}

// ---------- Rendering ----------
function renderTiles() {
  els.grid.innerHTML = '';
  if (!baseline || !baseline.tiles) return;

  ensureHelpMessageVisible();
  ensurePastShiftsButton();

  const tiles = (baseline.tiles || []).map(t => {
    const baselineLabel = t.booked
      ? 'BOOKED'
      : (t.status === 'BLOCKED' ? 'BLOCKED' : codeToLabel(t.status));
    const effectiveLabel = draft[t.ymd] ?? baselineLabel;
    return { ...t, baselineLabel, effectiveLabel };
  });

  for (const t of tiles) {
    const card = document.createElement('div');
    card.className = 'tile';
    card.dataset.ymd = t.ymd;

    const header = document.createElement('div');
    header.className = 'tile-header';

    const [Y, M, D] = t.ymd.split('-').map(Number);
    const d = new Date(Y, M - 1, D);
    const formatted = d
      .toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' })
      .replace(/,/g, '')
      .toUpperCase();

    const dateLine = document.createElement('div');
    dateLine.className = 'tile-day';
    dateLine.textContent = formatted;
    header.appendChild(dateLine);

    const status = document.createElement('div');
    status.className = `tile-status ${statusClass(t.effectiveLabel)}`;

    const sub = document.createElement('div');
    sub.className = 'tile-sub';

    // Content rules
    if (t.booked) {
      status.textContent = 'BOOKED';
      const line1 = (() => {
        const si = (t.shiftInfo || '').trim();
        if (si && si.toUpperCase() !== 'LONG DAY' && si.toUpperCase() !== 'NIGHT') return si;
        const shiftType = (t.shiftInfo || '').toUpperCase().includes('NIGHT') ? 'NIGHT' : 'LONG DAY';
        return shiftType === 'NIGHT' ? 'NIGHT 19:30–08:00' : 'LONG DAY 07:30–20:00';
      })();

      const hospital = (t.hospital || t.location || '').split(' – ')[0] || (t.hospital || t.location || '');
      const ward = (t.ward || (t.location && t.location.includes(' – ') ? t.location.split(' – ')[1] : '') || '').trim();
      const roleLine = (t.jobTitle && t.jobTitle.trim()) ? `Role: ${t.jobTitle.trim()}` : '';
      const refLine = (t.bookingRef && t.bookingRef.trim()) ? `Ref: ${t.bookingRef.trim()}` : '';

      const lines = [];
      lines.push(line1);
      if (hospital) lines.push(hospital);
      lines.push(`Ward: ${ward || ''}`);
      if (roleLine) lines.push(roleLine);
      if (refLine) lines.push(refLine);

      sub.innerHTML = lines.map(escapeHtml).join('<br>');
    } else if (t.status === 'BLOCKED' || t.effectiveLabel === 'BLOCKED') {
      status.textContent = 'BLOCKED';
      sub.textContent = 'You cannot work this shift as it will breach the 6 days in a row rule.';
    } else if (t.effectiveLabel === 'NOT AVAILABLE') {
      status.textContent = 'NOT AVAILABLE';
      if (t.effectiveLabel !== t.baselineLabel) {
        sub.innerHTML = `<span class="edit-note">Pending change</span>`;
      } else {
        sub.textContent = '';
      }
    } else if (
      t.effectiveLabel === 'LONG DAY' ||
      t.effectiveLabel === 'NIGHT' ||
      t.effectiveLabel === 'LONG DAY/NIGHT'
    ) {
      status.textContent = 'AVAILABLE FOR';
      const availability =
        t.effectiveLabel === 'LONG DAY' ? 'LONG DAY ONLY' :
        t.effectiveLabel === 'NIGHT' ? 'NIGHT ONLY' : 'LONG DAY AND NIGHT';
      if (t.effectiveLabel !== t.baselineLabel) {
        sub.innerHTML = `<span class="edit-note">Pending change</span><br>${escapeHtml(availability)}`;
      } else {
        sub.textContent = availability;
      }
    } else {
      // Pending variants (display-only swap)
      const displayLabel = displayPendingLabelForTile(t);
      status.textContent = displayLabel;

      // Mark pending tiles as needing attention (pulse + thin red border)
      if (!t.booked && t.status !== 'BLOCKED' && t.editable !== false) {
        card.classList.add('needs-attention');
        card.classList.add('attention-border'); // base and wrapped both show border while pending
      }

      if (t.effectiveLabel !== t.baselineLabel) {
        sub.innerHTML = `<span class="edit-note">Pending change</span>`;
      } else {
        sub.textContent = '';
      }
    }

    card.append(header, status, sub);

    els.grid.appendChild(card);
    // Fit status text
    fitStatusLabel(status);

    // Ensure non-attention tiles get a white border, per design
    if (!card.classList.contains('attention-border')) {
      card.style.borderColor = '#ffffff';
    }

    // Interactions
    const editable = (!t.booked && t.status !== 'BLOCKED' && t.editable !== false);
    if (editable) {
      card.style.cursor = 'pointer';
      card.title = buildCycleHint(t);
      card.addEventListener('click', () => {
        const curLabel = draft[t.ymd] ?? t.baselineLabel;   // always labels
        const nextLabel = nextStatus(curLabel);             // always labels

        toggledThisSession.add(t.ymd); // mark as user-touched this session

        if (nextLabel === t.baselineLabel) {
          delete draft[t.ymd];
        } else {
          // Store BASE pending label in draft (maps to blank code); display logic will wrap when needed
          draft[t.ymd] = nextLabel;
        }
        persistDraft();

        // user just edited — (re)start 1-minute reminder window
        registerUserEdit();

        renderTiles();
        showFooterIfNeeded();
      });
    } else {
      card.title = t.booked ? 'BOOKED (locked)' : (t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable');
      card.style.opacity = t.booked || t.status === 'BLOCKED' ? .9 : 1;
    }
  }

  // Equal heights: match tallest tile
  equalizeTileHeights();

  sizeGrid();
}
function equalizeTileHeights() {
  const tiles = Array.from(els.grid.querySelectorAll('.tile'));
  if (!tiles.length) return;
  tiles.forEach(t => t.style.minHeight = ''); // reset
  let max = 0;
  tiles.forEach(t => { max = Math.max(max, Math.ceil(t.getBoundingClientRect().height)); });
  tiles.forEach(t => { t.style.minHeight = max + 'px'; });
}
function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

// ---------- API ----------
function authToken() {
  // NEW: prefer session override, then saved token (localStorage), then shared fallback
  try {
    return (
      sessionStorage.getItem('api_t_override') ||
      localStorage.getItem(SAVED_TOKEN_KEY) ||
      API_SHARED_TOKEN
    );
  } catch {
    return API_SHARED_TOKEN;
  }
}

async function loadFromServer({ force=false } = {}) {
  if (!Object.keys(identity).length) {
    showAuthError('Not an authorised user');
    throw new Error('__AUTH_STOP__');
  }

  showLoading();

  try {
    const { res, json } = await apiGET({});
    // Dedicated busy handling
    if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
      showToast('Server is busy, please try again.');
      return;
    }

    const errCode = (json && json.error) || '';
    if (!res.ok) {
      if (res.status === 400 || res.status === 401 || res.status === 403) {
        // Special: UNAUTHORIZED_TOKEN → iPhone paste flow (only if server hints hadK true)
        if (errCode === 'UNAUTHORIZED_TOKEN' || (json && json.error === 'UNAUTHORIZED_TOKEN')) {
          openTokenClaimDialog();
          return;
        }
        showAuthError('Not an authorised user');
        throw new Error('__AUTH_STOP__');
      }
      const msg = errCode || `HTTP ${res.status}`;
      throw new Error(msg);
    }

    if (json && json.ok === false) {
      const unauthErrors = new Set([
        'INVALID_OR_UNKNOWN_IDENTITY',
        'NOT_IN_CANDIDATE_LIST',
        'FORBIDDEN'
      ]);
      if (json.error === 'UNAUTHORIZED_TOKEN') {
        openTokenClaimDialog();
        return;
      }
      if (unauthErrors.has(errCode)) {
        clearSavedIdentity();
        showAuthError('Not an authorised user');
        throw new Error('__AUTH_STOP__');
      }
      throw new Error(errCode || 'SERVER_ERROR');
    }

    baseline = json;

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

    const validYmd = new Set((baseline.tiles || []).map(x => x.ymd));
    for (const k of Object.keys(draft)) {
      if (!validYmd.has(k)) delete draft[k];
    }

    renderTiles();
    showFooterIfNeeded();

    // Show welcome/alert banner if provided
    maybeShowWelcome();
  } finally {
    hideLoading();
  }
}

async function submitChanges() {
  if (AUTH_DENIED) return;
  if (!baseline) return;

  const changes = [];
  for (const [ymd, newStatusLabel] of Object.entries(draft)) {
    const code = STATUS_TO_CODE[newStatusLabel];
    if (code === undefined) continue;
    changes.push({ ymd, code });
  }
  if (!changes.length) {
    showToast('No changes to submit.');
    return;
  }

  showLoading('Saving...');

  try {
    const { res, json, text } = await apiPOST({ changes });

    if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
      showToast('Server is busy, please try again.');
      return;
    }

    if (!res.ok) {
      if (res.status === 400 || res.status === 401 || res.status === 403) {
        if (json && json.error === 'UNAUTHORIZED_TOKEN') {
          openTokenClaimDialog();
          return;
        }
        showAuthError('Not an authorised user');
        throw new Error('__AUTH_STOP__');
      }
    }
    if (!res.ok || !json || json.ok === false) {
      const unauthErrors = new Set([
        'INVALID_OR_UNKNOWN_IDENTITY',
        'NOT_IN_CANDIDATE_LIST',
        'FORBIDDEN'
      ]);
      if (json && unauthErrors.has(json.error)) {
        clearSavedIdentity();
        showAuthError('Not an authorised user');
        throw new Error('__AUTH_STOP__');
      }
      const msg = (json && json.error) || `HTTP ${res.status}`;
      const extra = text && text !== '{}' ? ` – ${text.slice(0,200)}` : '';
      throw new Error(msg + extra);
    }

    const results = json.results || [];
    const applied = results.filter(r => r.applied).length;
    const rejected = results.filter(r => !r.applied);
    if (rejected.length) {
      const reasons = rejected.slice(0,3).map(r => `${r.ymd}: ${r.reason}`).join(', ');
      showToast(`Saved ${applied}. Skipped ${rejected.length} (${reasons}${rejected.length>3?', …':''}).`, 4200);
    } else {
      showToast(`Saved ${applied} change${applied===1?'':'s'}.`);
    }

    // Clear draft and any pending nudge state
    draft = {};
    persistDraft();
    cancelSubmitNudge();
    nudgeShown = false;
    lastEditAt = 0;

    // Reset per-session tap memory so post-save render mimics a cold load
    toggledThisSession.clear();

    await loadFromServer({ force: true });
  } catch (err) {
    if (String(err && err.message) !== '__AUTH_STOP__') {
      showToast('Submit failed: ' + (err.message || err));
    }
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
  // Reset submit-nudge state
  cancelSubmitNudge();
  nudgeShown = false;
  lastEditAt = 0;
}

// ---------- UI events ----------
els.submitBtn && els.submitBtn.addEventListener('click', () => {
  els.submitBtn.disabled = true;
  submitChanges().finally(() => { els.submitBtn.disabled = false; });
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
    if (!AUTH_DENIED) showToast('Reload failed: ' + e.message);
  } finally {
    hideLoading();
  }
});

// ---------- Keep CSS vars fresh ----------
window.addEventListener('resize', () => { sizeGrid(); equalizeTileHeights(); }, { passive: true });
window.addEventListener('orientationchange', () => setTimeout(() => { sizeGrid(); equalizeTileHeights(); }, 250));

/**
 * Light sizing helper: keeps --vh fresh and republishes footer height.
 * Works fine with the new scrolling layout (no fixed grid math).
 */
function sizeGrid() {
  if (!els.grid) return;

  // Stabilize mobile address bar changes
  const vh = window.innerHeight * 0.01;
  document.documentElement.style.setProperty('--vh', `${vh}px`);

  // Measure header/footer heights (footer may be visibility:hidden but measurable)
  const header = document.querySelector('header');
  const headerH = header ? Math.round(header.getBoundingClientRect().height) : 0;

  let footerH = 0;
  if (els.footer) {
    const rectH = Math.round(els.footer.getBoundingClientRect().height);
    footerH = rectH > 0 ? rectH : Math.round(getCssVarPx('--footer-h'));
  }

  document.documentElement.style.setProperty('--header-h', `${headerH}px`);
  document.documentElement.style.setProperty('--footer-h', `${footerH}px`);
}

// ---------- Boot ----------
(async function init() {
  parseQuery();

  // Ensure the URL carries token/identity so A2HS keeps them when added to Home Screen.
  ensureCanonicalURLWithToken();

  // If no k or msisdn provided in URL or storage → deny immediately
  if (!identity.k && !identity.msisdn) {
    showAuthError('Not an authorised user');
    return;
  }

  ensureLoadingOverlay(); // build overlays & menu styles early
  ensureMenu();           // add top-right menu

  loadDraft();
  showLoading(); // show immediately on first load
  try {
    await loadFromServer({ force: true });
  } catch (e) {
    if (String(e && e.message) !== '__AUTH_STOP__') {
      showToast('Load failed: ' + e.message, 5000);
    }
  } finally {
    sizeGrid();           // ensure initial CSS vars
    equalizeTileHeights();
    hideLoading();
    ensurePastShiftsButton();
  }

  // Wire token dialog buttons (created in ensureLoadingOverlay)
  const tokenClose = document.getElementById('tokenClose');
  if (tokenClose) tokenClose.addEventListener('click', () => closeOverlay('tokenOverlay'));
})();
