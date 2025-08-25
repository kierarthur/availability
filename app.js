/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL     = 'https://script.google.com/macros/s/AKfycbztdWnm2db9aRRFz2zfXSgcIbKRtlYSbA87pYbJ-ZvDZLFFz4HZBiStHk5UW8KV4ITAdQ/exec'; // <-- REPLACE
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
let baseline = null;          // server response (tiles, lastLoadedAt, candidateName, candidate, newUserHint?)
let draft = {};               // ymd -> status LABEL (diffs from baseline)
let identity = {};            // { msisdn?:string }
let lastHiddenAt = null;      // timestamp when page becomes hidden

const DRAFT_KEY = 'rota_avail_draft_v1';
const LAST_LOADED_KEY = 'rota_avail_last_loaded_v1';
const SAVED_IDENTITY_KEY = 'rota_avail_identity_v1';
const LAST_EMAIL_KEY = 'rota_last_login_email_v1';     // for prefilling login/forgot
// Note: we do NOT store passwords. Password managers handle that with biometrics.

// Per-session memory of tiles the user has tapped at least once this session.
const toggledThisSession = new Set();

// Submit-nudge timer state (1-minute reminder)
let submitNudgeTimer = null;
let lastEditAt = 0;
let nudgeShown = false;

// ---------- Simple API helpers ----------
async function apiGET(paramsObj) {
  const params = new URLSearchParams({ t: authToken(), ...(paramsObj || {}) });
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

// ---- Auth API wrappers ----
async function apiAuthLogin(email, password) {
  return apiPOST({ action: 'AUTH_LOGIN', email, password });
}
async function apiForgotPassword(email) {
  return apiPOST({ action: 'AUTH_FORGOT_PASSWORD', email });
}
async function apiResetPassword(k, newPassword) {
  return apiPOST({ action: 'AUTH_RESET_PASSWORD', k, newPassword });
}

// Saved-identity helpers
function saveIdentity(id) {
  try {
    if (id && id.msisdn) {
      localStorage.setItem(SAVED_IDENTITY_KEY, JSON.stringify(id));
    }
  } catch {}
}
function loadSavedIdentity() {
  try {
    const raw = localStorage.getItem(SAVED_IDENTITY_KEY);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    if (obj && obj.msisdn) return obj;
  } catch {}
  return {};
}
function clearSavedIdentity() {
  try { localStorage.removeItem(SAVED_IDENTITY_KEY); } catch {}
  // Keep LAST_EMAIL_KEY to help with login UX.
}

function rememberEmailLocal(email) {
  try { localStorage.setItem(LAST_EMAIL_KEY, email || ''); } catch {}
}
function getRememberedEmail() {
  try { return localStorage.getItem(LAST_EMAIL_KEY) || ''; } catch { return ''; }
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

    /* Overlays */
    #pastOverlay, #contentOverlay, #welcomeOverlay,
    #loginOverlay, #forgotOverlay, #resetOverlay {
      position: fixed; inset: 0; z-index: 9998;
      display: none; align-items: stretch; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
    }
    #pastOverlay.show, #contentOverlay.show, #welcomeOverlay.show,
    #loginOverlay.show, #forgotOverlay.show, #resetOverlay.show { display: flex; }

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
    .sheet-close:hover { background: #1e2532; color:#fff; }
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

    @media (prefers-reduced-motion: no-preference) {
      @keyframes softPulse {
        0%   { opacity: 1; transform: scale(1); }
        50%  { opacity: .78; transform: scale(0.99); }
        100% { opacity: 1; transform: scale(1); }
      }
      .needs-attention { animation: softPulse 1.6s ease-in-out infinite; }
      .btn-attention   { animation: strongPulse 1.0s ease-in-out infinite; }
    }

    .attention-border { border-color: var(--danger) !important; }

    /* Simple form inputs */
    .sheet-body label { font-weight:700; font-size:.9rem; margin-top:.2rem; }
    .sheet-body input {
      border-radius:8px; border:1px solid #222936; background:#0b0e14; color:#e7ecf3;
      padding:.6rem; width:100%;
    }
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

  // Past Shifts overlay
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

  // Content overlay
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

  // Welcome/Alert overlay
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

  // Login overlay
  const loginOverlay = document.createElement('div');
  loginOverlay.id = 'loginOverlay';
  loginOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Sign in">
      <div class="sheet-header">
        <div class="sheet-title">Sign in</div>
        <button id="loginClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div class="sheet-body">
        <form id="loginForm" autocomplete="on">
          <label for="loginEmail">Email</label>
          <input id="loginEmail" name="username" type="email"
                 inputmode="email" autocapitalize="none" spellcheck="false"
                 autocomplete="username" required />

          <label for="loginPassword" style="margin-top:.6rem">Password</label>
          <input id="loginPassword" name="password" type="password"
                 autocomplete="current-password" required enterkeyhint="go" />

          <div style="display:flex;gap:.6rem;margin-top:.8rem;">
            <button id="loginSubmit" class="menu-item" style="border:1px solid #2a3446;">Sign in</button>
            <button id="loginForgot" type="button" class="menu-item" style="border:1px solid #2a3446;">Forgot password?</button>
          </div>

          <div id="loginErr" class="muted" role="alert" style="margin-top:.4rem;"></div>
        </form>
      </div>
    </div>
  `;
  document.body.appendChild(loginOverlay);

  // Forgot overlay
  const forgotOverlay = document.createElement('div');
  forgotOverlay.id = 'forgotOverlay';
  forgotOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Forgot password">
      <div class="sheet-header">
        <div class="sheet-title">Forgot password</div>
        <button id="forgotClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div class="sheet-body">
        <form id="forgotForm" autocomplete="on">
          <label for="forgotEmail">Email</label>
          <input id="forgotEmail" name="username" type="email"
                 inputmode="email" autocapitalize="none" spellcheck="false"
                 autocomplete="username" required />
          <div style="display:flex;gap:.6rem;margin-top:.8rem;">
            <button id="forgotSubmit" class="menu-item" style="border:1px solid #2a3446;">Send reset link</button>
          </div>
          <div id="forgotMsg" class="muted" role="status" style="margin-top:.4rem;"></div>
        </form>
      </div>
    </div>
  `;
  document.body.appendChild(forgotOverlay);

  // Reset overlay (for ?k=… links)
  const resetOverlay = document.createElement('div');
  resetOverlay.id = 'resetOverlay';
  resetOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Set new password">
      <div class="sheet-header">
        <div class="sheet-title">Set a new password</div>
        <button id="resetClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div class="sheet-body">
        <form id="resetForm" autocomplete="on">
          <label for="resetPassword">New password</label>
          <input id="resetPassword" type="password" autocomplete="new-password"
                 minlength="8" required />
          <div class="muted" style="margin:.2rem 0 .6rem">
            Use at least 8 characters with uppercase, lowercase, and a number.
          </div>
          <div style="display:flex;gap:.6rem;">
            <button id="resetSubmit" class="menu-item" style="border:1px solid #2a3446;">Update password</button>
          </div>
          <div id="resetErr" class="muted" role="alert" style="margin-top:.4rem;"></div>
        </form>
      </div>
    </div>
  `;
  document.body.appendChild(resetOverlay);

  // Shared overlay wiring (click-outside & esc)
  ;[
    { overlayId:'pastOverlay', closeId:'pastClose' },
    { overlayId:'contentOverlay', closeId:'contentClose' },
    { overlayId:'welcomeOverlay', closeId:'welcomeClose' },
    { overlayId:'loginOverlay', closeId:'loginClose' },
    { overlayId:'forgotOverlay', closeId:'forgotClose' },
    { overlayId:'resetOverlay', closeId:'resetClose' },
  ].forEach(({overlayId, closeId}) => {
    const overlay = document.getElementById(overlayId);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeOverlay(overlayId); });
    document.getElementById(closeId).addEventListener('click', () => closeOverlay(overlayId));
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      ['pastOverlay','contentOverlay','welcomeOverlay','loginOverlay','forgotOverlay','resetOverlay'].forEach(id => {
        const ov = document.getElementById(id);
        if (ov && ov.classList.contains('show')) closeOverlay(id);
      });
    }
  });

  // Wire auth forms
  wireAuthForms();
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
    if (Object.keys(draft).length) scheduleSubmitNudge();
    lastHiddenAt = null;
  }
});

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
  if (els.footer) els.footer.classList.toggle('hidden', !dirty);

  if (els.submitBtn) {
    if (dirty) els.submitBtn.classList.add('btn-attention');
    else els.submitBtn.classList.remove('btn-attention');
  }

  sizeGrid();
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
    if (i0 && (i0.msisdn && identity.msisdn && i0.msisdn === identity.msisdn)) {
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

// ---- Code ↔ Label normalization ----
function codeToLabel(v) {
  if (v && STATUS_TO_CODE.hasOwnProperty(v)) return v;
  if (v === 'BOOKED' || v === 'BLOCKED') return v;
  if (v === undefined || v === null) return PENDING_LABEL_DEFAULT;
  const mapped = CODE_TO_STATUS[v];
  return mapped !== undefined ? mapped : PENDING_LABEL_DEFAULT;
}

function displayPendingLabelForTile(t) {
  if (t.effectiveLabel !== PENDING_LABEL_DEFAULT) return t.effectiveLabel;
  const serverPending = (t.baselineLabel === PENDING_LABEL_DEFAULT);
  const touched = toggledThisSession.has(t.ymd);
  return (serverPending && !touched) ? PENDING_LABEL_DEFAULT : PENDING_LABEL_WRAPPED;
}
function buildCycleHint(t) {
  const pendingVisual =
    (t.baselineLabel === PENDING_LABEL_DEFAULT && !toggledThisSession.has(t.ymd))
      ? PENDING_LABEL_DEFAULT
      : PENDING_LABEL_WRAPPED;
  const seq = [pendingVisual, 'N/A', 'LD', 'N', 'LD/N', pendingVisual];
  return 'Tap to change: ' + seq.join(' → ');
}
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

// ---------- Submit nudge ----------
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

// ---------- Past Shifts ----------
function ensurePastShiftsButton() {
  ensureMenu();
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
  const { res, json } = await apiGET({ view: 'past14' });
  if (!res.ok || !json || json.ok === false) {
    const msg = (json && json.error) || `HTTP ${res.status}`;
    throw new Error(msg);
  }
  return json.items || [];
}

// ---------- Content overlay ----------
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
    <button class="menu-item" id="miTimesheet">Send a timesheet by email</button>
    <hr style="border:none;border-top:1px solid #222936;margin:.35rem 0;">
    <button class="menu-item" id="miChangePw">Change password</button>
    <button class="menu-item" id="miLogout">Logout</button>
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

  // Change password: send a reset link to the known email
  document.getElementById('miChangePw').addEventListener('click', async () => {
    list.classList.remove('show');
    const email = (baseline && baseline.candidate && baseline.candidate.email) ||
                  getRememberedEmail();
    if (!email) {
      // If we don’t know the email, open the Forgot flow
      openForgotOverlay();
      return;
    }
    try {
      showLoading('Requesting password reset link...');
      await apiForgotPassword(email); // always returns ok UX
      showToast('If your email exists, a reset link has been sent.');
    } catch (e) {
      showToast('Could not start password change: ' + (e.message || e));
    } finally {
      hideLoading();
    }
  });

  document.getElementById('miLogout').addEventListener('click', () => {
    list.classList.remove('show');
    clearSavedIdentity();
    draft = {}; persistDraft();
    baseline = null;
    els.grid && (els.grid.innerHTML = '');
    openLoginOverlay();
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
    try { await apiPOST({ action:'MARK_MESSAGE_SEEN' }); } catch {}
    closeOverlay('welcomeOverlay');
  };
}

// ---------- Auth overlays logic ----------
function openLoginOverlay() {
  const email = getRememberedEmail();
  const le = document.getElementById('loginEmail');
  const lerr = document.getElementById('loginErr');
  if (le) le.value = email;
  if (lerr) lerr.textContent = '';
  openOverlay('loginOverlay', '#loginEmail');
}
function openForgotOverlay() {
  const fe = document.getElementById('forgotEmail');
  const fmsg = document.getElementById('forgotMsg');
  if (fe) fe.value = getRememberedEmail();
  if (fmsg) fmsg.textContent = '';
  openOverlay('forgotOverlay', '#forgotEmail');
}
function openResetOverlay() {
  const rerr = document.getElementById('resetErr');
  if (rerr) rerr.textContent = '';
  openOverlay('resetOverlay', '#resetPassword');
}

function wireAuthForms() {
  // Login
  const lf  = document.getElementById('loginForm');
  const le  = document.getElementById('loginEmail');
  const lp  = document.getElementById('loginPassword');
  const ls  = document.getElementById('loginSubmit');
  const lfg = document.getElementById('loginForgot');
  const lerr= document.getElementById('loginErr');

  lf?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = (le.value || '').trim().toLowerCase();
    const pw    = lp.value || '';
    if (!email || !pw) return;

    try {
      ls.disabled = true; showLoading('Signing in...');
      const { res, json } = await apiAuthLogin(email, pw);
      if (!res.ok || !json || json.ok === false) {
        const msg = (json && json.error) || `HTTP ${res.status}`;
        lerr.textContent = (msg === 'INVALID_CREDENTIALS') ? 'Email or password is incorrect.' : (msg || 'Sign-in failed.');
        return;
      }
      identity = { msisdn: json.msisdn }; saveIdentity(identity);
      rememberEmailLocal(email);
      await storeCredentialIfSupported(email, pw); // Chrome/Android

      // Clear password field
      lp.value = '';
      closeOverlay('loginOverlay');

      await loadFromServer({ force: true });
    } finally {
      hideLoading(); ls.disabled = false;
    }
  });
  lfg?.addEventListener('click', () => { closeOverlay('loginOverlay'); openForgotOverlay(); });

  // Forgot
  const ff   = document.getElementById('forgotForm');
  const fe   = document.getElementById('forgotEmail');
  const fs   = document.getElementById('forgotSubmit');
  const fmsg = document.getElementById('forgotMsg');
  ff?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = (fe.value || '').trim().toLowerCase();
    if (!email) return;
    try {
      fs.disabled = true; showLoading('Sending link...');
      await apiForgotPassword(email); // always ok UX
      rememberEmailLocal(email);
      fmsg.textContent = 'If your email exists, a reset link has been sent.';
      setTimeout(() => { closeOverlay('forgotOverlay'); openLoginOverlay(); }, 1200);
    } finally {
      hideLoading(); fs.disabled = false;
    }
  });

  // Reset
  const rf   = document.getElementById('resetForm');
  const rp   = document.getElementById('resetPassword');
  const rs   = document.getElementById('resetSubmit');
  const rerr = document.getElementById('resetErr');
  rf?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const pw = rp.value || '';
    if (!pw || pw.length < 8 || !/[a-z]/.test(pw) || !/[A-Z]/.test(pw) || !/[0-9]/.test(pw)) {
      rerr.textContent = 'Use at least 8 chars with uppercase, lowercase, and a number.';
      return;
    }
    const k = new URLSearchParams(location.search).get('k') || '';
    if (!k) { rerr.textContent = 'This reset link is invalid or missing.'; return; }

    try {
      rs.disabled = true; showLoading('Updating...');
      const { res, json } = await apiResetPassword(k, pw);
      if (!res.ok || !json || json.ok === false) {
        const msg = (json && json.error) || `HTTP ${res.status}`;
        if (msg === 'INVALID_OR_EXPIRED_RESET') {
          rerr.textContent = 'This link has expired. Please request a new one.';
        } else {
          rerr.textContent = msg || 'Could not update password.';
        }
        return;
      }
      // Strip ?k=… from URL (reset is one-time)
      const clean = location.pathname + (location.hash || '');
      history.replaceState(null, '', clean);
      closeOverlay('resetOverlay');
      showToast('Password updated. Please sign in.');
      openLoginOverlay();
    } finally {
      hideLoading(); rs.disabled = false;
    }
  });
}

// ---------- Credential Management API (Android/Chrome) ----------
async function storeCredentialIfSupported(email, password) {
  try {
    if ('credentials' in navigator && 'PasswordCredential' in window) {
      const cred = new PasswordCredential({ id: email, password, name: email });
      await navigator.credentials.store(cred);
    }
  } catch {}
}
async function tryAutoLoginViaCredentialsAPI() {
  try {
    if (!('credentials' in navigator)) return false;
    const cred = await navigator.credentials.get({ password: true, mediation: 'optional' });
    if (cred && cred.id && cred.password) {
      const { res, json } = await apiAuthLogin(cred.id, cred.password);
      if (res.ok && json && json.ok && json.msisdn) {
        identity = { msisdn: json.msisdn }; saveIdentity(identity);
        rememberEmailLocal(cred.id);
        await loadFromServer({ force: true });
        return true;
      }
    }
  } catch {}
  return false;
}

// ---------- Rendering ----------
function renderTiles() {
  if (!els.grid) { console.warn('renderTiles(): #grid not found'); return; }
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
      const displayLabel = displayPendingLabelForTile(t);
      status.textContent = displayLabel;

      if (!t.booked && t.status !== 'BLOCKED' && t.editable !== false) {
        card.classList.add('needs-attention');
        card.classList.add('attention-border');
      }

      if (t.effectiveLabel !== t.baselineLabel) {
        sub.innerHTML = `<span class="edit-note">Pending change</span>`;
      } else {
        sub.textContent = '';
      }
    }

    card.append(header, status, sub);
    els.grid.appendChild(card);
    fitStatusLabel(status);

    if (!card.classList.contains('attention-border')) {
      card.style.borderColor = '#ffffff';
    }

    const editable = (!t.booked && t.status !== 'BLOCKED' && t.editable !== false);
    if (editable) {
      card.style.cursor = 'pointer';
      card.title = buildCycleHint(t);
      card.addEventListener('click', () => {
        const curLabel = draft[t.ymd] ?? t.baselineLabel;
        const nextLabel = nextStatus(curLabel);

        toggledThisSession.add(t.ymd);

        if (nextLabel === t.baselineLabel) {
          delete draft[t.ymd];
        } else {
          draft[t.ymd] = nextLabel;
        }
        persistDraft();
        registerUserEdit();

        renderTiles();
        showFooterIfNeeded();
      });
    } else {
      card.title = t.booked ? 'BOOKED (locked)' : (t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable');
      card.style.opacity = t.booked || t.status === 'BLOCKED' ? .9 : 1;
    }
  }

  equalizeTileHeights();
  sizeGrid();
}
function equalizeTileHeights() {
  const tiles = Array.from(els.grid.querySelectorAll('.tile'));
  if (!tiles.length) return;
  tiles.forEach(t => t.style.minHeight = '');
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

// ---------- API shared token ----------
function authToken() {
  // Allow override via #t or ?t if you ever need it, else fallback to shared token.
  try {
    const search = new URLSearchParams(location.search);
    const hash   = new URLSearchParams(location.hash && location.hash.startsWith('#') ? location.hash.slice(1) : '');
    return search.get('t') || hash.get('t') || API_SHARED_TOKEN;
  } catch {
    return API_SHARED_TOKEN;
  }
}

// ---------- Load + Submit ----------
async function loadFromServer({ force=false } = {}) {
  if (!authToken()) {
    showAuthError('Missing or invalid token');
    throw new Error('__AUTH_STOP__');
  }
  if (!identity || !identity.msisdn) {
    // Not signed in yet; show login if not already visible
    openLoginOverlay();
    return;
  }

  showLoading();

  try {
    const { res, json } = await apiGET({});
    if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
      showToast('Server is busy, please try again.');
      return;
    }

    const errCode = (json && json.error) || '';
    if (!res.ok) {
      if (res.status === 400 || res.status === 401 || res.status === 403) {
        clearSavedIdentity();
        showAuthError('Not an authorised user');
        throw new Error('__AUTH_STOP__');
      }
      throw new Error(errCode || `HTTP ${res.status}`);
    }

    if (json && json.ok === false) {
      const unauthErrors = new Set([
        'INVALID_OR_UNKNOWN_IDENTITY',
        'NOT_IN_CANDIDATE_LIST',
        'FORBIDDEN'
      ]);
      if (unauthErrors.has(errCode)) {
        clearSavedIdentity();
        openLoginOverlay();
        return;
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

    // capture candidate email to assist Change Password flow
    if (baseline && baseline.candidate && baseline.candidate.email) {
      rememberEmailLocal(String(baseline.candidate.email).trim().toLowerCase());
    }

    saveLastLoaded(json.lastLoadedAt || new Date().toISOString());

    const validYmd = new Set((baseline.tiles || []).map(x => x.ymd));
    for (const k of Object.keys(draft)) {
      if (!validYmd.has(k)) delete draft[k];
    }

    renderTiles();
    showFooterIfNeeded();
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

    if (!res.ok || !json || json.ok === false) {
      const unauthErrors = new Set([
        'INVALID_OR_UNKNOWN_IDENTITY',
        'NOT_IN_CANDIDATE_LIST',
        'FORBIDDEN'
      ]);
      if (json && unauthErrors.has(json.error)) {
        clearSavedIdentity();
        openLoginOverlay();
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

    draft = {};
    persistDraft();
    cancelSubmitNudge();
    nudgeShown = false;
    lastEditAt = 0;
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

function sizeGrid() {
  if (!els.grid) return;
  const vh = window.innerHeight * 0.01;
  document.documentElement.style.setProperty('--vh', `${vh}px`);

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

// ---------- Simple router for auth flows ----------
function routeFromURL() {
  const hash = location.hash || '';
  const hasMsisdn = !!(identity && identity.msisdn);

  // If a reset token is present in the QUERY (?k=...), show reset
  const k = new URLSearchParams(location.search).get('k');
  if (k) { openResetOverlay(); return; }

  if (/^#\/login/.test(hash))  { openLoginOverlay(); return; }
  if (/^#\/forgot/.test(hash)) { openForgotOverlay(); return; }
  if (/^#\/reset/.test(hash))  { openResetOverlay(); return; }

  if (!hasMsisdn) { openLoginOverlay(); return; }
}
window.addEventListener('hashchange', routeFromURL);

// ---------- Boot ----------
(async function init() {
  // Identity from storage
  identity = loadSavedIdentity();

  // Build overlays/menu early
  ensureLoadingOverlay();
  ensureMenu();

  // If reset link present, show reset
  const hasK = new URLSearchParams(location.search).has('k');
  if (hasK) {
    openResetOverlay();
  } else if (!identity.msisdn) {
    // Attempt silent auto‑sign‑in on Android/Chrome
    const autoOK = await tryAutoLoginViaCredentialsAPI();
    if (!autoOK) openLoginOverlay();
  }

  // Draft + initial load (only if already signed in)
  loadDraft();
  if (identity && identity.msisdn) {
    showLoading();
    try {
      await loadFromServer({ force: true });
    } catch (e) {
      if (String(e && e.message) !== '__AUTH_STOP__') {
        showToast('Load failed: ' + e.message, 5000);
      }
    } finally {
      sizeGrid();
      equalizeTileHeights();
      hideLoading();
      ensurePastShiftsButton();
    }
  }

  routeFromURL();
})();
