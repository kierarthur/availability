/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL     = 'https://script.google.com/macros/s/AKfycbw9X71BbvC55s2iJhyfGU5PoBtOxC9jvFsdKpd8BvrNghMfw8Yy9X-iSZ1Xcw9oTAPMcA/exec'; // <-- REPLACE
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
  emergencyBtn: document.getElementById('emergencyBtn'), // new red EMERGENCY button
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
  try { els.emergencyBtn && (els.emergencyBtn.disabled = true); } catch {} // NEW

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
let baseline = null;          // server response (tiles, lastLoadedAt, candidateName, candidate, newUserHint? / iOS/Android variants)
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

// ===== Overlay config (dismiss/blocks-other) =====
const OVERLAY_CONFIG = {
  pastOverlay:     { dismissible: true,  blocking: false },
  contentOverlay:  { dismissible: true,  blocking: false },
  welcomeOverlay:  { dismissible: false, blocking: true  }, // MUST press “Got it”
  // Login is truly blocking & non-dismissable
  loginOverlay:    { dismissible: false, blocking: true  },
  forgotOverlay:   { dismissible: true,  blocking: false },
  resetOverlay:    { dismissible: false, blocking: true  }, // MUST complete/reset flow
  emergencyOverlay:{ dismissible: true,  blocking: false }, // new EMERGENCY flow wizard
  alertOverlay:    { dismissible: false, blocking: true  }  // single-OK blocking alert
};

function isBlockingOverlayOpen() {
  const ids = Object.keys(OVERLAY_CONFIG).filter(id => OVERLAY_CONFIG[id].blocking === true);
  return ids.some(id => {
    const el = document.getElementById(id);
    return el && el.classList.contains('show');
  });
}

// ------- Loading overlay + shared styles -------
let _loadingCount = 0;
function ensureLoadingOverlay() {
  if (document.getElementById('loadingOverlay')) return;

  const style = document.createElement('style');
  style.textContent = `
    :root{
      /* Brand theming for login + reset */
      --brand-blue-dk: #0b254a;        /* adjust to your logo blue */
      --brand-blue-dk-95: #0d2b56;     /* slightly lighter for headers/bands */
      --border-white: #ffffff;
      --text-on-blue: #eaf0f7;
      --muted-on-blue: #c7d3e5;
    }

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
#loginOverlay, #forgotOverlay, #resetOverlay, #alertOverlay, #emergencyOverlay {
  position: fixed; inset: 0; z-index: 9998;
  display: none; align-items: stretch; justify-content: center;
  background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
}
#pastOverlay.show, #contentOverlay.show, #welcomeOverlay.show,
#loginOverlay.show, #forgotOverlay.show, #resetOverlay.show, #alertOverlay.show, #emergencyOverlay.show {
  display: flex;
}

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

    /* ===== Themed look for LOGIN + RESET sheets (dark blue + white borders) ===== */
    #loginOverlay .sheet,
    #resetOverlay .sheet {
      background: var(--brand-blue-dk);
      color: var(--text-on-blue);
      border: 2px solid var(--border-white);
      box-shadow: 0 16px 40px rgba(0,0,0,.5);
    }
    #loginOverlay .sheet-header,
    #resetOverlay .sheet-header {
      background: var(--brand-blue-dk-95);
      border-bottom: 1px solid rgba(255,255,255,0.4);
    }
    #loginOverlay .sheet-title,
    #resetOverlay .sheet-title {
      color: #ffffff;
    }
    #loginOverlay .sheet-body,
    #resetOverlay .sheet-body {
      color: var(--text-on-blue);
    }
    #loginOverlay .muted,
    #resetOverlay .muted {
      color: var(--muted-on-blue);
    }
    #loginOverlay .sheet-body label,
    #resetOverlay .sheet-body label {
      color: var(--text-on-blue);
    }
    #loginOverlay .sheet-body input,
    #resetOverlay .sheet-body input {
      background: rgba(0,0,0,0.25);
      border: 1px solid rgba(255,255,255,0.55);
      color: #ffffff;
    }
    #loginOverlay .sheet-body input:focus,
    #resetOverlay .sheet-body input:focus {
      outline: 2px solid #ffffff;
      outline-offset: 1px;
      border-color: #ffffff;
    }
    #loginOverlay .icon-btn,
    #resetOverlay .icon-btn,
    #loginOverlay .menu-item,
    #resetOverlay .menu-item {
      border-color: #ffffff !important;
      color: #ffffff !important;
      background: transparent;
    }
    #loginOverlay .menu-item:hover,
    #resetOverlay .menu-item:hover,
    #loginOverlay .icon-btn[aria-pressed="true"],
    #resetOverlay .icon-btn[aria-pressed="true"] {
      background: rgba(255,255,255,0.12) !important;
    }

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

    /* Simple form inputs (default look for non-themed sheets) */
    .sheet-body label { font-weight:700; font-size:.9rem; margin-top:.2rem; }
    .sheet-body input {
      border-radius:8px; border:1px solid #222936; background:#0b0e14; color:#e7ecf3;
      padding:.6rem; width:100%;
    }

    /* Input-with-button row for reveal eye */
    .input-row {
      display:flex; gap:.4rem; align-items:center;
    }
    .input-row input {
      flex:1 1 auto;
    }
    .icon-btn {
      appearance:none; border:1px solid #2a3446; background:#131926; color:#e7ecf3;
      border-radius:8px; padding:.45rem .6rem; font-weight:700; cursor:pointer;
    }
    .icon-btn[aria-pressed="true"] { background:#1b2331; }
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

  // Past Shifts overlay (dismissible)
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

  // Content overlay (dismissible)
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

  // Welcome/Alert overlay (NON-dismissable: no close button)
  const welcomeOverlay = document.createElement('div');
  welcomeOverlay.id = 'welcomeOverlay';
  welcomeOverlay.innerHTML = `
    <div id="welcomeSheet" class="sheet" role="dialog" aria-modal="true" aria-label="Welcome">
      <div class="sheet-header">
        <div id="welcomeTitle" class="sheet-title">Welcome</div>
      </div>
      <div id="welcomeBody" class="sheet-body" tabindex="0"></div>
      <div style="display:flex;justify-content:flex-end;gap:.5rem;padding:.5rem .9rem 1rem;">
        <button id="welcomeDismiss" class="menu-item" style="border:1px solid #2a3446;">Got it</button>
      </div>
    </div>
  `;
  document.body.appendChild(welcomeOverlay);

  // Blocking single-OK alert overlay (NON-dismissable)
  const alertOverlay = document.createElement('div');
  alertOverlay.id = 'alertOverlay';
  alertOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Alert">
      <div class="sheet-header">
        <div id="alertTitle" class="sheet-title">Notice</div>
      </div>
      <div id="alertBody" class="sheet-body"></div>
      <div style="display:flex;justify-content:flex-end;gap:.5rem;padding:.5rem .9rem 1rem;">
        <button id="alertOk" class="menu-item" style="border:1px solid #2a3446;">OKAY</button>
      </div>
    </div>
  `;
  document.body.appendChild(alertOverlay);

  // Login overlay (NON-dismissable + BLOCKING) — NO close button
  const loginOverlay = document.createElement('div');
  loginOverlay.id = 'loginOverlay';
  loginOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Sign in">
      <div class="sheet-header">
        <div class="sheet-title">Sign in</div>
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

  // Forgot overlay (dismissible)
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

  // Reset overlay (NON-dismissable: no close button)
  const resetOverlay = document.createElement('div');
  resetOverlay.id = 'resetOverlay';
  resetOverlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Set new password">
      <div class="sheet-header">
        <div class="sheet-title">Set a new password</div>
      </div>
      <div class="sheet-body">
        <form id="resetForm" autocomplete="on">
          <label for="resetPassword">New password</label>
          <div class="input-row">
            <input id="resetPassword" type="password" autocomplete="new-password" minlength="8" required />
            <button id="resetPwReveal" type="button" class="icon-btn" aria-pressed="false" aria-label="Show password" title="Show password">👁</button>
          </div>
          <div class="muted" style="margin:.2rem 0 .6rem">
            Use at least 8 characters with uppercase, lowercase, and a number.
          </div>

          <label for="resetConfirm">Confirm new password</label>
          <div class="input-row">
            <input id="resetConfirm" type="password" autocomplete="new-password" minlength="8" required />
            <button id="resetConfirmReveal" type="button" class="icon-btn" aria-pressed="false" aria-label="Show password" title="Show password">👁</button>
          </div>

          <div style="display:flex;gap:.6rem;margin-top:.8rem;">
            <button id="resetSubmit" class="menu-item" style="border:1px solid #2a3446;">Update password</button>
          </div>
          <div id="resetErr" class="muted" role="alert" style="margin-top:.4rem;"></div>
        </form>
      </div>
    </div>
  `;
  document.body.appendChild(resetOverlay);

  // Shared overlay wiring — respect dismissibility
  [
  { overlayId:'pastOverlay',    closeId:'pastClose' },
  { overlayId:'contentOverlay', closeId:'contentClose' },
  // welcomeOverlay: NO closeId (non-dismissable)
  // loginOverlay: NON-dismissable (no close button, do NOT wire)
  { overlayId:'forgotOverlay',  closeId:'forgotClose' },
  // resetOverlay: NO closeId (non-dismissable)
  // alertOverlay: NON-dismissable (wired via showBlockingAlert)
  { overlayId:'emergencyOverlay', closeId:'emergencyClose' } // NEW
].forEach(({overlayId, closeId}) => {
  const overlay = document.getElementById(overlayId);
  const cfg = OVERLAY_CONFIG[overlayId] || { dismissible:true, blocking:false };
  if (!overlay) return;

  // Backdrop click → only if dismissible
  overlay.addEventListener('click', (e) => {
    if (e.target !== overlay) return;
    if (!cfg.dismissible) { e.stopPropagation(); return; }
    closeOverlay(overlayId);
  });

  // Close button → only wired if exists and dismissible
  if (closeId) {
    const btn = document.getElementById(closeId);
    if (btn) {
      btn.addEventListener('click', () => {
        if (!cfg.dismissible) return;
        closeOverlay(overlayId);
      });
    }
  }
});

// Global Escape → only closes dismissible overlays
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return;
  [
    'pastOverlay',
    'contentOverlay',
    'welcomeOverlay',
    'loginOverlay',
    'forgotOverlay',
    'resetOverlay',
    'alertOverlay',
    'emergencyOverlay' // NEW
  ].forEach(id => {
    const ov = document.getElementById(id);
    const cfg = OVERLAY_CONFIG[id] || { dismissible:true, blocking:false };
    if (ov && ov.classList.contains('show') && cfg.dismissible) closeOverlay(id);
  });
});


  // Wire auth forms
  wireAuthForms();
}

// ----- Overlay open/close with config awareness -----
function openOverlay(overlayId, focusSel) {
  ensureLoadingOverlay();
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  const cfg = OVERLAY_CONFIG[overlayId] || { dismissible:true, blocking:false };
  overlay.dataset.dismissible = String(!!cfg.dismissible);
  overlay.dataset.blocking = String(!!cfg.blocking);

  overlay.classList.add('show');
  overlay.dataset.prevFocus = document.activeElement && document.activeElement.id ? document.activeElement.id : '';
  const focusEl = focusSel ? overlay.querySelector(focusSel) : overlay.querySelector('.sheet-close') || overlay.querySelector('[tabindex],input,button,select,textarea');
  focusEl && focusEl.focus();

  // When blocking is active, prevent other overlays from opening later
  trapFocus(overlay);
}

function closeOverlay(overlayId, force=false) {
  const overlay = document.getElementById(overlayId);
  if (!overlay) return;

  const dismissible = overlay.dataset.dismissible !== 'false';
  if (!dismissible && !force) return; // honor non-dismissable unless force=true (e.g., Got it)

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

// ----- Blocking alert helper (OK-only) -----
function showBlockingAlert(message, onOk) {
  ensureLoadingOverlay();
  const body = document.getElementById('alertBody');
  const ok = document.getElementById('alertOk');
  const title = document.getElementById('alertTitle');
  if (!body || !ok) return;
  title.textContent = 'Notice';
  body.innerHTML = `<div class="muted">${message || ''}</div>`;
  // Replace previous click
  const newOk = ok.cloneNode(true);
  ok.parentNode.replaceChild(newOk, ok);
  newOk.addEventListener('click', () => {
    closeOverlay('alertOverlay', /*force*/true);
    if (typeof onOk === 'function') onOk();
  });
  openOverlay('alertOverlay', '#alertOk');
}

/* ---------- PWA install prompt (improved) ---------- */
let deferredPrompt = null;

function isStandalone() {
  try {
    return (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) ||
           (typeof navigator !== 'undefined' && 'standalone' in navigator && navigator.standalone === true);
  } catch { return false; }
}

function isiOS() {
  const ua = navigator.userAgent || navigator.vendor || window.opera || '';
  return /iphone|ipad|ipod/i.test(ua);
}
function isAndroid() {
  const ua = navigator.userAgent || navigator.vendor || window.opera || '';
  return /android/i.test(ua);
}

// Fallback overlay with quick instructions (iOS / no BIP event)
function openInstallInstructions() {
  if (isBlockingOverlayOpen()) return;
  ensureLoadingOverlay();
  const titleEl = document.getElementById('contentTitle');
  const bodyEl  = document.getElementById('contentBody');
  if (!titleEl || !bodyEl) return;
  titleEl.textContent = 'Install this app';
  const iosSteps = `
    <ol>
      <li>Tap the <strong>Share</strong> icon in Safari.</li>
      <li>Choose <strong>Add to Home Screen</strong>.</li>
      <li>Tap <strong>Add</strong>.</li>
    </ol>
  `;
  const androidSteps = `
    <ol>
      <li>Open browser menu (⋮).</li>
      <li>Tap <strong>Install app</strong> or <strong>Add to Home screen</strong>.</li>
      <li>Confirm.</li>
    </ol>
  `;
  bodyEl.innerHTML = `
    <div class="muted">${isiOS()
      ? 'On iPhone/iPad, install by adding to Home Screen:'
      : 'Install this app from your browser:'}</div>
    ${isiOS() ? iosSteps : androidSteps}
  `;
  openOverlay('contentOverlay', '#contentClose');
}

// Listen for the installability event (Android/Chrome)
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredPrompt = e;
  // Show CTA only when installable on Android/Chromium and not standalone
  if (els.installBtn && !isStandalone() && isAndroid()) {
    els.installBtn.classList.remove('hidden');
    els.installBtn.disabled = false;
  }
});

// Click handler for Install button
els.installBtn && els.installBtn.addEventListener('click', async () => {
  if (isStandalone()) { els.installBtn.classList.add('hidden'); return; }

  // iOS → always show instructions
  if (isiOS()) {
    openInstallInstructions();
    return;
  }

  // Android/Chromium: only prompt if we captured BIP
  if (isAndroid() && deferredPrompt) {
    try {
      deferredPrompt.prompt();
      await deferredPrompt.userChoice;
    } catch {}
    deferredPrompt = null;
    els.installBtn.classList.add('hidden');
    return;
  }

  // Other/Not ready yet on Android
  showToast('Install not available yet — try again shortly.');
});

// Hide or show CTA based on platform + state
function syncInstallCTAVisibility() {
  if (!els.installBtn) return;
  if (isStandalone()) {
    els.installBtn.classList.add('hidden');
    els.installBtn.disabled = true;
    return;
  }
  // iOS: visible (instructions)
  if (isiOS()) {
    els.installBtn.classList.remove('hidden');
    els.installBtn.disabled = false;
    return;
  }
  // Android/Chromium: visible only when BIP captured
  if (isAndroid() && deferredPrompt) {
    els.installBtn.classList.remove('hidden');
    els.installBtn.disabled = false;
  } else {
    els.installBtn.classList.add('hidden');
    els.installBtn.disabled = true;
  }
}

// When the app gets installed
window.addEventListener('appinstalled', () => {
  if (els.installBtn) els.installBtn.classList.add('hidden');
  showToast('Installed — open it from your Home Screen.');
});

// Initial CTA sync (in case BIP never fires, e.g., iOS)
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (els.installBtn) {
      syncInstallCTAVisibility();
    }
  }, 400);
});
// ───────────────────────── CHANGED ─────────────────────────
// Single, minimal place to refresh on visibility change.
document.addEventListener('visibilitychange', () => {
  if (document.hidden) {
    cancelSubmitNudge();
    lastHiddenAt = Date.now();
    persistDraft();
    return;
  }

  const wasAwayLong = lastHiddenAt && (Date.now() - lastHiddenAt > 3 * 60 * 1000);
  lastHiddenAt = null;

  // If there were local edits, clear them consistently (same as before)
  if (wasAwayLong && Object.keys(draft).length) {
    draft = {};
    persistDraft();
    showToast('Unsaved changes were cleared after inactivity.');
  }

  // Only one path to refresh: the orchestrator. Coalesces via loadFromServer() guard.
  const needsRefresh = wasAwayLong || !Object.keys(draft).length;
  if (needsRefresh) {
    showLoading();
    refreshAppState({ force: true })
      .catch(err => { if (!AUTH_DENIED) showToast('Reload failed: ' + err.message); })
      .finally(() => { hideLoading(); });
  }

  if (Object.keys(draft).length) scheduleSubmitNudge();
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

// ───────────────────────── NEW (background single-flight) ─────────────────────────
async function prefetchEmergencyEligibility({ priority = 'normal', onReady } = {}) {
  // If we already have eligibility cached and non-empty, nothing to do
  if (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible) && emergencyEligibilityCache.eligible.length) {
    // Ensure button reflects current cache
    updateEmergencyButtonFromCache();
    if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
    return;
  }

  // Single-flight guard: avoid duplicate background calls
  if (prefetchEmergencyEligibility._inflight) {
    try { await prefetchEmergencyEligibility._inflight; } catch (_) {}
    // After existing inflight finishes, reflect result & notify
    updateEmergencyButtonFromCache();
    if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
    return;
  }

  prefetchEmergencyEligibility._inflight = (async () => {
    try {
      const { ok, eligible } = await apiGetEmergencyWindow();
      if (ok && Array.isArray(eligible)) {
        emergencyEligibilityCache = { eligible };
      } else {
        // Keep an empty cache on failure; UI will remain disabled, which is safe
        emergencyEligibilityCache = { eligible: [] };
      }
      updateEmergencyButtonFromCache();
      if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
    } finally {
      prefetchEmergencyEligibility._inflight = null;
    }
  })();

  return prefetchEmergencyEligibility._inflight;
}

function showFooterIfNeeded() {
  const dirty = Object.keys(draft).length > 0;
  if (els.footer) els.footer.classList.toggle('hidden', !dirty);

  if (els.submitBtn) {
    if (dirty) els.submitBtn.classList.add('btn-attention');
    else els.submitBtn.classList.remove('btn-attention');
  }

  // Important: do NOT call sizeGrid() here to avoid scroll jolts when footer appears.
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
  if (isBlockingOverlayOpen()) return; // guard against blocking modals
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
  if (isBlockingOverlayOpen()) return; // guard against blocking modals
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
    if (isBlockingOverlayOpen()) return; // don't open dropdown over blocking modal
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
    if (isBlockingOverlayOpen()) return;
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
    if (isBlockingOverlayOpen()) return;
    const email = (baseline && baseline.candidate && baseline.candidate.email) ||
                  getRememberedEmail();
    if (!email) {
      openForgotOverlay();
      return;
    }
    try {
      showLoading('Requesting password reset link...');
      await apiForgotPassword(email); // always returns ok UX
      // Requested exact wording:
      showToast('Password Reset request has been sent by email if email exists');
    } catch (e) {
      showToast('Could not start password change: ' + (e.message || e));
    } finally {
      hideLoading();
    }
  });

  document.getElementById('miLogout').addEventListener('click', () => {
  list.classList.remove('show');
  if (isBlockingOverlayOpen()) return;

  clearSavedIdentity();
  draft = {}; persistDraft();
  baseline = null;
  els.grid && (els.grid.innerHTML = '');

  // --- NEW: clear emergency state + hide/disable the button + close overlay ---
  try {
    // Clear any cached emergency state your new code keeps
    window._emergencyEligibleShifts = [];
    window._emergencyState = null;
  } catch {}

  // Hide/disable the EMERGENCY button and un-wire it
  try {
  if (els.emergencyBtn) {
    els.emergencyBtn.hidden = true;
    els.emergencyBtn.disabled = true;
    // No need to remove event listeners here; wireEmergencyButton/openEmergencyOverlay handles click logic.
  }
} catch {}


  // If the emergency overlay is open, close it defensively
  try { closeOverlay('emergencyOverlay'); } catch {}

  openLoginOverlay();
});

}

// ---------- Welcome / newUserHint with platform targeting ----------
function maybeShowWelcome() {
  if (!baseline) return;

  // Server can provide device-specific hints; we pick best match.
  const isIOS = isiOS();
  const isAnd = isAndroid();

  // Priority: platform-specific → generic fallback
  const hint =
    (isIOS && baseline.newUserHintIos) ? baseline.newUserHintIos :
    (isAnd && baseline.newUserHintAndroid) ? baseline.newUserHintAndroid :
    baseline.newUserHint;

  if (!hint) return;

  const titleEl = document.getElementById('welcomeTitle');
  const bodyEl  = document.getElementById('welcomeBody');
  if (!titleEl || !bodyEl) return;

  titleEl.textContent = (hint.title || 'Welcome').trim();
  bodyEl.innerHTML = hint.html || '';
  openOverlay('welcomeOverlay', '#welcomeDismiss');

  const dismiss = document.getElementById('welcomeDismiss');
  // Make CTA more prominent
  dismiss.style.background = '#1b2331';
  dismiss.style.fontWeight = '800';
  dismiss.style.border = '1px solid #2a3446';
  dismiss.style.padding = '.6rem 1rem';
  dismiss.style.borderRadius = '10px';

  dismiss.onclick = async () => {
    try { await apiPOST({ action:'MARK_MESSAGE_SEEN' }); } catch {}
    closeOverlay('welcomeOverlay', /* force */ true);
  };
}

// ---------- Auth overlays logic ----------
function openLoginOverlay() {
  if (isBlockingOverlayOpen()) return; // cannot open over blocking modals
  const email = getRememberedEmail();
  const le = document.getElementById('loginEmail');
  const lp = document.getElementById('loginPassword');
  const lerr = document.getElementById('loginErr');
  if (le) le.value = email;
  if (lerr) lerr.textContent = '';
  openOverlay('loginOverlay', email ? '#loginPassword' : '#loginEmail');
}
function openForgotOverlay() {
  if (isBlockingOverlayOpen()) return; // cannot open over blocking modals
  const fe = document.getElementById('forgotEmail');
  const fmsg = document.getElementById('forgotMsg');
  if (fe) fe.value = getRememberedEmail();
  if (fmsg) fmsg.textContent = '';
  openOverlay('forgotOverlay', '#forgotEmail');
}
function openResetOverlay() {
  // Non-dismissable by config; keep user inside until completion
  const rerr = document.getElementById('resetErr');
  const rp = document.getElementById('resetPassword');
  const rc = document.getElementById('resetConfirm');
  const rs = document.getElementById('resetSubmit');
  if (rerr) rerr.textContent = '';
  if (rp) rp.value = '';
  if (rc) rc.value = '';
  if (rs) rs.disabled = true; // disabled until validated
  openOverlay('resetOverlay', '#resetPassword');

  // If token is missing → immediately exit to blocking alert, then login
  const k = new URLSearchParams(location.search).get('k');
  if (!k) {
    closeOverlay('resetOverlay', /*force*/true);
    showBlockingAlert('This reset link is invalid or missing.', () => {
      const clean = location.pathname + (location.hash || '');
      history.replaceState(null, '', clean);
      openLoginOverlay();
    });
  }
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
      // *** FIX: force-close the blocking login overlay after successful auth ***
      closeOverlay('loginOverlay', /* force */ true);

      await loadFromServer({ force: true });
    } finally {
      hideLoading(); ls.disabled = false;
    }
  });
  lfg?.addEventListener('click', () => { closeOverlay('loginOverlay', /* force */ true); openForgotOverlay(); });

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
      await apiForgotPassword(email); // always ok UX (neutral; no enumeration)
      rememberEmailLocal(email);
      // *** Requested exact wording for unauthenticated flow ***
      fmsg.textContent = 'Password Reset request has been sent by email if email exists';
      setTimeout(() => { closeOverlay('forgotOverlay'); openLoginOverlay(); }, 1200);
    } finally {
      hideLoading(); fs.disabled = false;
    }
  });

  // Reset — confirm + validation + reveal toggles
  const rf   = document.getElementById('resetForm');
  const rp   = document.getElementById('resetPassword');
  const rc   = document.getElementById('resetConfirm');
  const rs   = document.getElementById('resetSubmit');
  const rerr = document.getElementById('resetErr');
  const pwEye = document.getElementById('resetPwReveal');
  const cfEye = document.getElementById('resetConfirmReveal');

  function meetsPolicy(pw) {
    return !!(pw && pw.length >= 8 && /[a-z]/.test(pw) && /[A-Z]/.test(pw) && /[0-9]/.test(pw));
  }
  function same(p, c) {
    return p === c && p.length > 0;
  }
  function validateReset() {
    const k = new URLSearchParams(location.search).get('k') || '';
    const p = rp ? rp.value : '';
    const c = rc ? rc.value : '';
    let ok = true;
    let msg = '';

    if (!k) {
      ok = false;
      msg = 'This reset link is invalid or missing.';
    } else if (!meetsPolicy(p)) {
      ok = false;
      msg = 'Use at least 8 chars with uppercase, lowercase, and a number.';
    } else if (!same(p, c)) {
      ok = false;
      msg = "Passwords don't match.";
    }

    if (rerr) rerr.textContent = msg;
    if (rs) rs.disabled = !ok;
    return ok;
  }

  // Live validation on input
  rp?.addEventListener('input', validateReset);
  rc?.addEventListener('input', validateReset);

  // Eye toggles
  pwEye?.addEventListener('click', () => {
    const isText = rp.type === 'text';
    rp.type = isText ? 'password' : 'text';
    pwEye.setAttribute('aria-pressed', String(!isText));
    pwEye.setAttribute('aria-label', isText ? 'Show password' : 'Hide password');
    pwEye.title = isText ? 'Show password' : 'Hide password';
    rp.focus({ preventScroll: true });
  });
  cfEye?.addEventListener('click', () => {
    const isText = rc.type === 'text';
    rc.type = isText ? 'password' : 'text';
    cfEye.setAttribute('aria-pressed', String(!isText));
    cfEye.setAttribute('aria-label', isText ? 'Show password' : 'Hide password');
    cfEye.title = isText ? 'Show password' : 'Hide password';
    rc.focus({ preventScroll: true });
  });

  // Submit handler
  rf?.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!validateReset()) return;

    const pw = rp.value || '';
    const k = new URLSearchParams(location.search).get('k') || '';
    if (!k) {
      // Defensive: should be handled earlier
      closeOverlay('resetOverlay', /*force*/true);
      showBlockingAlert('This reset link is invalid or missing.', () => {
        const clean = location.pathname + (location.hash || '');
        history.replaceState(null, '', clean);
        openLoginOverlay();
      });
      return;
    }

    try {
      rs.disabled = true; showLoading('Updating...');
      const { res, json } = await apiResetPassword(k, pw);
      if (!res.ok || !json || json.ok === false) {
        const msg = (json && json.error) || `HTTP ${res.status}`;
        if (msg === 'INVALID_OR_EXPIRED_RESET') {
          // Close reset → blocking OK → clean URL → to Login
          closeOverlay('resetOverlay', /*force*/true);
          showBlockingAlert('This link has expired. Please request a new one.', () => {
            const clean = location.pathname + (location.hash || '');
            history.replaceState(null, '', clean);
            openLoginOverlay();
          });
        } else if (msg === 'WEAK_PASSWORD') {
          rerr.textContent = 'Use at least 8 chars with uppercase, lowercase, and a number.';
        } else {
          rerr.textContent = msg || 'Could not update password.';
        }
        return;
      }
      // Strip ?k=… from URL (reset is one-time)
      const clean = location.pathname + (location.hash || '');
      history.replaceState(null, '', clean);

      // Force-close non-dismissable overlay after success
      closeOverlay('resetOverlay', /* force */ true);
      showToast('Password updated. Please sign in.');
      openLoginOverlay();

      // Clear fields
      rp.value = '';
      rc.value = '';
      rs.disabled = true;
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
// ───────────────────────── CHANGED ─────────────────────────
function renderTiles() {
  if (!els.grid) { console.warn('renderTiles(): #grid not found'); return; }
  els.grid.innerHTML = '';
  if (!baseline || !baseline.tiles) {
    // No fetches from here; rendering must be side-effect free.
    return;
  }

  ensureHelpMessageVisible();
  ensurePastShiftsButton();

  const tiles = (baseline.tiles || []).map(t => {
    const baselineLabel = t.booked
      ? 'BOOKED'
      : (t.status === 'BLOCKED' ? 'BLOCKED' : codeToLabel(t.status));
    const effectiveLabel = draft[t.ymd] ?? baselineLabel;
    return { ...t, baselineLabel, effectiveLabel };
  });

  function updateCardUI(card, ymd, baselineLabel, newLabel, isBooked, isBlocked, editable) {
    const statusEl = card.querySelector('.tile-status');
    const subEl    = card.querySelector('.tile-sub');
    statusEl.className = 'tile-status ' + statusClass(newLabel);

    if (isBooked) {
      statusEl.textContent = 'BOOKED';
      return;
    }

    if (isBlocked || newLabel === 'BLOCKED') {
      statusEl.textContent = 'BLOCKED';
      subEl.textContent = 'You cannot work this shift as it will breach the 6 days in a row rule.';
      card.classList.remove('needs-attention', 'attention-border');
    } else if (newLabel === 'NOT AVAILABLE') {
      statusEl.textContent = 'NOT AVAILABLE';
      subEl.innerHTML = (newLabel !== baselineLabel) ? `<span class="edit-note">Pending change</span>` : '';
      card.classList.remove('needs-attention', 'attention-border');
    } else if (newLabel === 'LONG DAY' || newLabel === 'NIGHT' || newLabel === 'LONG DAY/NIGHT') {
      statusEl.textContent = 'AVAILABLE FOR';
      const availability =
        newLabel === 'LONG DAY' ? 'LONG DAY ONLY' :
        newLabel === 'NIGHT' ? 'NIGHT ONLY' : 'LONG DAY AND NIGHT';
      subEl.innerHTML = (newLabel !== baselineLabel)
        ? `<span class="edit-note">Pending change</span><br>${escapeHtml(availability)}`
        : availability;
      card.classList.remove('needs-attention', 'attention-border');
    } else {
      const displayLabel = (function () {
        if (newLabel !== PENDING_LABEL_DEFAULT) return newLabel;
        const serverPending = (baselineLabel === PENDING_LABEL_DEFAULT);
        const touched = toggledThisSession.has(ymd);
        return (serverPending && !touched) ? PENDING_LABEL_DEFAULT : PENDING_LABEL_WRAPPED;
      })();
      statusEl.textContent = displayLabel;

      if (newLabel !== baselineLabel) {
        subEl.innerHTML = `<span class="edit-note">Pending change</span>`;
      } else {
        subEl.textContent = '';
      }

      if (editable) {
        card.classList.add('needs-attention', 'attention-border');
      } else {
        card.classList.remove('needs-attention', 'attention-border');
      }
    }

    if (!card.classList.contains('attention-border')) {
      card.style.borderColor = '#ffffff';
    }

    fitStatusLabel(statusEl);
  }

  for (const t of tiles) {
    const card = document.createElement('div');
    card.className = 'tile';
    card.dataset.ymd = t.ymd;
    card.dataset.baselineLabel = t.baselineLabel;
    card.dataset.blocked = String(t.status === 'BLOCKED');
    card.dataset.booked = String(!!t.booked);
    card.dataset.editable = String(!t.booked && t.status !== 'BLOCKED' && t.editable !== false);

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
      sub.innerHTML = (t.effectiveLabel !== t.baselineLabel) ? `<span class="edit-note">Pending change</span>` : '';
    } else if (
      t.effectiveLabel === 'LONG DAY' ||
      t.effectiveLabel === 'NIGHT' ||
      t.effectiveLabel === 'LONG DAY/NIGHT'
    ) {
      status.textContent = 'AVAILABLE FOR';
      const availability =
        t.effectiveLabel === 'LONG DAY' ? 'LONG DAY ONLY' :
        t.effectiveLabel === 'NIGHT' ? 'NIGHT ONLY' : 'LONG DAY AND NIGHT';
      sub.innerHTML = (t.effectiveLabel !== t.baselineLabel)
        ? `<span class="edit-note">Pending change</span><br>${escapeHtml(availability)}`
        : availability;
    } else {
      const displayLabel = displayPendingLabelForTile(t);
      status.textContent = displayLabel;

      if (!t.booked && t.status !== 'BLOCKED' && t.editable !== false) {
        card.classList.add('needs-attention', 'attention-border');
      }

      sub.innerHTML = (t.effectiveLabel !== t.baselineLabel) ? `<span class="edit-note">Pending change</span>` : '';
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
        const ymd = card.dataset.ymd;
        const baselineLabel = card.dataset.baselineLabel;
        const isBooked  = card.dataset.booked === 'true';
        const isBlocked = card.dataset.blocked === 'true';
        const canEdit   = card.dataset.editable === 'true';
        if (!canEdit) return;

        const currentLabel = draft[ymd] ?? baselineLabel;
        const nextLabel = nextStatus(currentLabel);
        toggledThisSession.add(ymd);

        if (nextLabel === baselineLabel) {
          delete draft[ymd];
        } else {
          draft[ymd] = nextLabel;
        }
        persistDraft();
        registerUserEdit();

        const effective = draft[ymd] ?? baselineLabel;
        updateCardUI(card, ymd, baselineLabel, effective, isBooked, isBlocked, canEdit);

        const pseudoTile = { ymd, baselineLabel, effectiveLabel: effective };
        card.title = buildCycleHint(pseudoTile);

        showFooterIfNeeded();
      });
    } else {
      card.title = t.booked ? 'BOOKED (locked)' : (t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable');
      card.style.opacity = t.booked || t.status === 'BLOCKED' ? .9 : 1;
    }
  }

  // Equalize only on initial/full render, not on tap
  equalizeTileHeights({ reset: true });

  // No network calls here anymore; button state comes from cache set in loadFromServer().
  updateEmergencyButtonFromCache();
}


let cachedTileHeight = 0; // global cache

function equalizeTileHeights({ reset = false } = {}) {
  const tiles = Array.from(els.grid.querySelectorAll('.tile'));
  if (!tiles.length) return;

  // On a full render / resize we recompute the max
  if (reset || cachedTileHeight === 0) {
    // Clear previous minHeight so we can measure natural heights
    tiles.forEach(t => t.style.minHeight = '');

    let max = 0;
    tiles.forEach(t => {
      max = Math.max(max, Math.ceil(t.getBoundingClientRect().height));
    });

    cachedTileHeight = max;
  }

  // Apply cached uniform height to all tiles
  tiles.forEach(t => {
    t.style.minHeight = cachedTileHeight + 'px';
  });
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
// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
async function loadFromServer({ force=false } = {}) {
  // Single-flight guard (coalesce bursts)
  if (loadFromServer._inflight && !force) return loadFromServer._inflight;

  if (!authToken()) {
    try { els.emergencyBtn && (els.emergencyBtn.hidden = true, els.emergencyBtn.disabled = true); } catch {}
    showAuthError('Missing or invalid token');
    throw new Error('__AUTH_STOP__');
  }
  if (!identity || !identity.msisdn) {
    try { els.emergencyBtn && (els.emergencyBtn.hidden = true, els.emergencyBtn.disabled = true); } catch {}
    if (!isBlockingOverlayOpen()) openLoginOverlay();
    return;
  }

  // Safe default: button invisible until eligibility arrives
  try {
    if (els.emergencyBtn) {
      els.emergencyBtn.hidden   = true;  // keep hidden while we background-load
      els.emergencyBtn.disabled = true;  // defensive; won’t be seen anyway
    }
  } catch {}

  showLoading();

  const work = (async () => {
    try {
      // ── fetch BASELINE ONLY; do not block on emergency ──
      const { res, json } = await apiGET({});

      if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
        try { els.emergencyBtn && (els.emergencyBtn.hidden = true, els.emergencyBtn.disabled = true); } catch {}
        showToast('Server is busy, please try again.');
        return;
      }

      const errCode = (json && json.error) || '';
      if (!res.ok) {
        if (res.status === 400 || res.status === 401 || res.status === 403) {
          try { els.emergencyBtn && (els.emergencyBtn.hidden = true, els.emergencyBtn.disabled = true); } catch {}
          clearSavedIdentity();
          showAuthError('Not an authorised user');
          throw new Error('__AUTH_STOP__');
        }
        throw new Error(errCode || `HTTP ${res.status}`);
      }

      if (json && json.ok === false) {
        const unauthErrors = new Set(['INVALID_OR_UNKNOWN_IDENTITY','NOT_IN_CANDIDATE_LIST','FORBIDDEN']);
        if (unauthErrors.has(errCode)) {
          try { els.emergencyBtn && (els.emergencyBtn.hidden = true, els.emergencyBtn.disabled = true); } catch {}
          clearSavedIdentity();
          if (!isBlockingOverlayOpen()) openLoginOverlay();
          return;
        }
        throw new Error(errCode || 'SERVER_ERROR');
      }

      // ------- baseline UI/data -------
      baseline = json;

      const name =
        (json.candidateName && String(json.candidateName).trim()) ||
        (json.candidate && [json.candidate.firstName, json.candidate.surname].filter(Boolean).join(' ').trim()) || '';
      if (els.candidateName) {
        if (name) {
          els.candidateName.textContent = name;
          els.candidateName.classList.remove('hidden');
        } else {
          els.candidateName.textContent = '';
          els.candidateName.classList.add('hidden');
        }
      }

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

      // Reflect whatever cache exists now (likely empty → stays hidden)
      updateEmergencyButtonFromCache();

      // ── kick off EMERGENCY eligibility in the background; do not await ──
      prefetchEmergencyEligibility().catch(() => { /* fail closed; stays hidden */ });

    } finally {
      hideLoading();
    }
  })();

  loadFromServer._inflight = work.finally(() => { loadFromServer._inflight = null; });
  return loadFromServer._inflight;
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

  // Map for quick lookup of the code we attempted to set per day
  const changeMap = new Map(changes.map(c => [c.ymd, c.code]));

  showLoading('Saving...');

  try {
    const { res, json, text } = await apiPOST({ changes });

    if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
      showToast('Server is busy, please try again.');
      return;
    }

    if (!res.ok || !json || json.ok === false) {
      const unauthErrors = new Set(['INVALID_OR_UNKNOWN_IDENTITY','NOT_IN_CANDIDATE_LIST','FORBIDDEN']);
      if (json && unauthErrors.has(json.error)) {
        clearSavedIdentity();
        if (!isBlockingOverlayOpen()) openLoginOverlay();
        throw new Error('__AUTH_STOP__');
      }
      const msg = (json && json.error) || `HTTP ${res.status}`;
      const extra = text && text !== '{}' ? ` – ${text.slice(0,200)}` : '';
      throw new Error(msg + extra);
    }

    const results  = json.results || [];
    const applied  = results.filter(r => r.applied);
    const rejected = results.filter(r => !r.applied);

    // Optimistically update local baseline for successfully applied days
    if (baseline && Array.isArray(baseline.tiles) && applied.length) {
      const byYmd = new Map(baseline.tiles.map(t => [t.ymd, t]));
      for (const r of applied) {
        const tile = byYmd.get(r.ymd);
        if (!tile) continue;
        // Prefer code from server result if present, else fall back to what we sent
        const newCode = (r.code !== undefined && r.code !== null) ? r.code : changeMap.get(r.ymd);
        if (newCode !== undefined) {
          tile.status = newCode; // keep booked/blocked flags as-is; server rules will sync on next real refresh
        }
      }
      // Bump the "last loaded" timestamp locally to avoid stale-looking UI
      saveLastLoaded(new Date().toISOString());
    }

    // User feedback
    if (rejected.length) {
      const reasons = rejected.slice(0,3).map(r => `${r.ymd}: ${r.reason}`).join(', ');
      showToast(`Saved ${applied.length}. Skipped ${rejected.length} (${reasons}${rejected.length>3?', …':''}).`, 4200);
    } else {
      showToast(`Saved ${applied.length} change${applied.length===1?'':'s'}.`);
    }

    // Clear local edit state
    draft = {};
    persistDraft();
    cancelSubmitNudge();
    nudgeShown = false;
    lastEditAt = 0;
    toggledThisSession.clear();

    // Re-render locally (no full reload)
    renderTiles();
    showFooterIfNeeded();
  } catch (err) {
    if (String(err && err.message) !== '__AUTH_STOP__') {
      showToast('Submit failed: ' + (err.message || err));
    }
  } finally {
  hideLoading();
  try { refreshEmergencyEligibility({ silent:true }); } catch {}

}

}

// ───────────────────────── CHANGED ─────────────────────────
function clearChanges() {
  draft = {};
  persistDraft();
  renderTiles();
  showFooterIfNeeded();
  showToast('Cleared pending changes.');
  cancelSubmitNudge();
  nudgeShown = false;
  lastEditAt = 0;

  // No network fetch here. Button state derives from existing cache.
  updateEmergencyButtonFromCache();
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
    await refreshAppState({ force: true }); // single path: coalesces and fills caches
  } catch (e) {
    if (!AUTH_DENIED) showToast('Reload failed: ' + e.message);
  } finally {
    hideLoading();
  }
});


// ---------- Keep CSS vars fresh ----------
window.addEventListener('resize', () => { 
  sizeGrid(); 
  equalizeTileHeights({ reset: true }); 
}, { passive: true });

window.addEventListener('orientationchange', () => 
  setTimeout(() => { 
    sizeGrid(); 
    equalizeTileHeights({ reset: true }); 
  }, 250)
);

/* =========================
 * NEW STATE (module-level)
 * ========================= */
let emergencyState = null;            // Wizard state (see resetEmergencyState)
let emergencyEligibilityCache = null; // Last successful result from apiGetEmergencyWindow()



/** POST { action:'RUNNING_LATE_SEND', context, eta_label } */
async function apiPostRunningLateSend(shift, etaLabel) {
  try {
    // `shift` may actually be the full response from OPTIONS/PREVIEW, or just the context itself.
    const context = shift && shift.context ? shift.context : shift;

    const { res, json } = await apiPOST({ action: 'RUNNING_LATE_SEND', context, eta_label: etaLabel });
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && (json.error || json.message)) || `HTTP ${res.status}`;
      return { ok: false, sent: null, tokens: null, error: msg || 'UNKNOWN_ERROR' };
    }

    // Server returns { ok, sent: {sameShift, previousShift, emergencyContacts}, tokens: {...} }
    return {
      ok: true,
      sent: json.sent || null,
      tokens: json.tokens || null
    };
  } catch (e) {
    return { ok: false, sent: null, tokens: null, error: String(e && e.message) || 'NETWORK_ERROR' };
  }
}

/** POST { action:'EMERGENCY_RAISE', shift, issue:{...} } */


async function apiPostEmergencyRaise(shift, issueType, issuePayload) {
  try {
    const it = String(issueType || '').toUpperCase();

    let body;
    if (it === 'DNA') {
      // New DNA shape
      body = {
        action: 'EMERGENCY_RAISE',
        shift: shift || {},
        issue: {
          type: 'DNA',
          subject_name: (issuePayload && issuePayload.subject_name) || '',
          // Only include if present
          ...(issuePayload && issuePayload.subject_msisdn ? { subject_msisdn: issuePayload.subject_msisdn } : {}),
          ...(issuePayload && issuePayload.reason_text ? { reason_text: issuePayload.reason_text } : {})
        }
      };
    } else {
      // Legacy shape for RL/CA/LE
      const etaOrLeave = (issuePayload && issuePayload.eta_or_leave_time_label) || 'N/A';
      const reason = (issuePayload && issuePayload.reason_text) || '';
      body = {
        action: 'EMERGENCY_RAISE',
        shift: shift || {},
        issue: {
          issue_type: it,                 // "CANNOT_ATTEND" | "LEAVE_EARLY" | (older clients)
          eta_or_leave_time_label: etaOrLeave,
          reason_text: reason
        }
      };
    }

    const { res, json } = await apiPOST(body);
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && (json.error || json.message)) || `HTTP ${res.status}`;
      return { ok: false, message: '', error: msg || 'UNKNOWN_ERROR' };
    }
    return { ok: true, message: json.message || '', alert_id: json.alert_id || null, notifications: json.notifications || null };
  } catch (e) {
    return { ok: false, message: '', error: String(e && e.message) || 'NETWORK_ERROR' };
  }
}


/* =========================
 * NEW UI ENTRYPOINTS
 * ========================= */

/** Create the bright red EMERGENCY button next to Refresh (once). Hidden by default. */
function ensureEmergencyButton() {
  const existing = document.getElementById('emergencyBtn');
  if (existing) { 
    els.emergencyBtn = existing; // keep els in sync if already present
    return; 
  }

  // Find header and refresh button
  const header = document.querySelector('header');
  if (!header || !els.refreshBtn) return;

  // Create button
  const btn = document.createElement('button');
  btn.id = 'emergencyBtn';
  btn.type = 'button';
  btn.textContent = 'EMERGENCY';
  btn.setAttribute('aria-label', 'Report an emergency');
  btn.className = ''; // do not inherit standard .btn styling
  Object.assign(btn.style, {
    background: '#e06666',
    color: '#ffffff',
    fontWeight: '900',
    borderRadius: '10px',
    padding: '.5rem .7rem',
    border: '0',
    cursor: 'pointer',
    fontSize: '.9rem',
    outline: '0',
    transition: 'transform .08s ease, opacity .15s ease',
    boxShadow: '0 0 0 1px rgba(0,0,0,.2) inset',
  });
  btn.onpointerdown = () => { btn.style.transform = 'scale(0.98)'; };
  btn.onpointerup = () => { btn.style.transform = 'scale(1)'; };

  // Hidden by default
  btn.hidden = true;

  // Insert immediately after Refresh
  els.refreshBtn.insertAdjacentElement('afterend', btn);
  els.emergencyBtn = btn; // <-- keep els in sync

  // Wire click
  btn.addEventListener('click', () => {
    if (btn.dataset.busy === '1') return;
    openEmergencyOverlay();
  });
}

/** Show/hide the EMERGENCY button based on eligible shifts. */
function syncEmergencyButtonVisibility(eligible) {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;
  const hasEligible = Array.isArray(eligible) && eligible.length > 0;
  btn.hidden = !hasEligible;
}

/** Open the Emergency overlay, seeding state from cached/fetched eligibility. */


/** Close the Emergency overlay and reset state. */

/** Reset the wizard state but keep overlay open. */

function resetEmergencyState() {
  emergencyState = {
    eligible: [],
    selectedShift: null,
    issueType: null,              // 'RUNNING_LATE' | 'CANNOT_ATTEND' | 'LEAVE_EARLY' | 'DNA'
    lateOptions: [],              // [{ label, minutes }]
    runningLateContext: null,     // ensure default for RL preview/send
    selectedLateLabel: null,
    reasonText: '',
    previewHtml: '',
    // DNA-specific
    dnaOptions: [],               // [{ name, role, msisdn07?, id? }] – server-provided cohort/peers list
    dnaAbsenteeName: null,
    dnaAbsenteeMsisdn: null,      // not shown in UI list; may be present from server
    dnaTriedCalling: false,       // gate for Continue
    step: 'PICK_SHIFT',
    __confirmSubmitting: false,   // double-submit guard reset
    confirmError: ''              // inline confirm warning
  };
}



/* =========================
 * NEW OVERLAY RENDERER
 * ========================= */

/** Build and manage the Emergency overlay's current step. */




/* =========================
 * NEW STEP HANDLERS
 * ========================= */

function handleEmergencyPickShift() {
  const body = document.getElementById('emergencyBody');
  if (!body || !emergencyState) return;

  const chosen = body.querySelector('input[name="emShift"]:checked');
  if (!chosen || !chosen._shift) {
    showToast('Please select a shift.');
    return;
  }
  emergencyState.selectedShift = chosen._shift;
  emergencyState.issueType = null;
  emergencyState.step = 'PICK_ISSUE';
  renderEmergencyStep();
}

function handleEmergencyPickIssue() {
  const body = document.getElementById('emergencyBody');
  if (!body || !emergencyState) return;

  const chosen = body.querySelector('input[name="emIssue"]:checked');
  if (!chosen) {
    showToast('Please select an option.');
    return;
  }
  const issue = String(chosen.value || '').toUpperCase();
  emergencyState.issueType = issue;
  emergencyState.selectedLateLabel = null;
  emergencyState.reasonText = '';
  emergencyState.previewHtml = '';

  // reset DNA-specific state on entry
  if (issue === 'DNA') {
    emergencyState.dnaAbsenteeName = null;
    emergencyState.dnaAbsenteeMsisdn = null;
    emergencyState.dnaTriedCalling = false;
  }

  emergencyState.step = 'DETAILS';
  renderEmergencyStep();
}



/* =========================
 * NEW ERROR / UX UTILITIES
 * ========================= */

function mapServerErrorToMessage(err) {
  const s = String(err || '').toUpperCase();
  if (!s) return 'Something went wrong. Please try again.';
  if (s.includes('NOT_ELIGIBLE')) return 'This shift is not eligible for emergency actions.';
  if (s.includes('SHIFT_NOT_FOUND')) return 'We could not find that shift.';
  if (s.includes('OUT_OF_WINDOW')) return 'This action is outside the allowed time window.';
  if (s.includes('RATE_LIMIT')) return 'Please wait a moment and try again.';
  if (s.includes('TEMPORARILY_BUSY_TRY_AGAIN')) return 'Server is busy, please try again.';
  if (s.includes('FORBIDDEN') || s.includes('UNAUTH') || s.includes('INVALID_OR_UNKNOWN_IDENTITY')) return 'You are not authorised for this action.';
  if (/HTTP\s*\d+/.test(s)) return `Server error: ${err}`;
  return 'Could not complete the request. Please try again.';
}

/** Refresh eligibility and toggle button; failure is silent by default. */
/** Refresh eligibility and toggle button; failure is silent by default. */

/* =========================
 * NEW LIFECYCLE HOOKS
 * ========================= */


/** Call at the end of Refresh button flow to re-check eligibility. */
async function onRefreshClickForEmergency() {
  await refreshEmergencyEligibility({ silent: true });
}


/* =========================
 * NEW OVERLAY FACTORY
 * ========================= */

/** Lazily create the Emergency overlay DOM (uses existing overlay system). */

function ensureEmergencyOverlay() {
  if (document.getElementById('emergencyOverlay')) return;

  ensureLoadingOverlay(); // ensures shared styles and focus-trap helpers exist

  const overlay = document.createElement('div');
  overlay.id = 'emergencyOverlay';
  overlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Emergency">
      <div class="sheet-header">
        <div id="emergencyTitle" class="sheet-title">Emergency</div>
        <button id="emergencyClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div id="emergencyBody" class="sheet-body"></div>
      <div id="emergencyFooter" style="display:flex;justify-content:flex-end;gap:.5rem;padding:.5rem .9rem 1rem;"></div>
    </div>
  `;
  document.body.appendChild(overlay);

  // Close button — emergency overlay is dismissible
  const closeBtn = overlay.querySelector('#emergencyClose');
  if (closeBtn) closeBtn.addEventListener('click', () => closeEmergencyOverlay(false));

  // Backdrop click to close (consistent with dismissible overlays)
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closeEmergencyOverlay(false);
  });
}


function populateEmergencyModalStart(eligible) {
  // Don’t open on top of a blocking overlay (e.g., reset/login/alert).
  if (typeof isBlockingOverlayOpen === 'function' && isBlockingOverlayOpen()) return;

  // Ensure the Emergency overlay DOM exists.
  if (typeof ensureEmergencyOverlay === 'function') ensureEmergencyOverlay();

  // Resolve eligible shifts: prefer the argument; otherwise fall back to cache; else empty.
  const list =
    (Array.isArray(eligible) && eligible) ||
    (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible) && emergencyEligibilityCache.eligible) ||
    [];

  // Reset wizard state and seed with eligibility.
  if (typeof resetEmergencyState === 'function') resetEmergencyState();
  if (!emergencyState) {
    // Safety: create a minimal state if reset helper isn’t available for some reason.
    emergencyState = { eligible: [], selectedShift: null, issueType: null, lateOptions: [], selectedLateLabel: null, reasonText: '', previewHtml: '', step: 'PICK_SHIFT' };
  }
  emergencyState.eligible = list;
  emergencyState.step = 'PICK_SHIFT';

  // Open overlay and render first step (shift selection or “no eligible shifts”).
  if (typeof openOverlay === 'function') openOverlay('emergencyOverlay', '#emergencyClose');
  if (typeof renderEmergencyStep === 'function') renderEmergencyStep();
}

function syncEmergencyEligibilityUI(eligible) {
  try {
    const btn = document.getElementById('emergencyBtn');
    if (!btn) return;

    // Prefer explicit arg → cache → global fallback (tolerates older call sites)
    const list =
      Array.isArray(eligible) ? eligible :
      (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible))
        ? emergencyEligibilityCache.eligible
        : (window._emergencyEligibleShifts || []);

    const hasEligible = Array.isArray(list) && list.length > 0;
    btn.hidden = !hasEligible;     // ← use attribute, not class
    btn.disabled = !hasEligible;

  } catch {
    // Defensive: hide if anything goes wrong
    try {
      const btn = document.getElementById('emergencyBtn');
      if (btn) { btn.hidden = true; btn.disabled = true; }  // ← use attribute, not class
    } catch {}
  }
}



function wireEmergencyButton() {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;

  // Remove any old click handlers to avoid duplicates
  const clone = btn.cloneNode(true);
  btn.replaceWith(clone);
  const freshBtn = document.getElementById('emergencyBtn');

  // Delegate to the canonical entrypoint
  freshBtn.addEventListener('click', () => {
    if (freshBtn.dataset.busy === '1') return; // respect busy state if set
    openEmergencyOverlay(); // this handles cache/fetch, overlay open, and rendering
  });
}

// =========================
// SERVER-AUTHORITATIVE FLOW
// =========================

async function refreshEmergencyEligibility({ silent = false } = {}) {
  const { ok, eligible, error } = await apiGetEmergencyWindow();
  if (!ok) {
    // Fail soft: clear caches, hide button, keep app usable
    emergencyEligibilityCache = null;
    window._emergencyEligibleShifts = [];
    syncEmergencyButtonVisibility([]);
    try { typeof syncEmergencyEligibilityUI === 'function' && syncEmergencyEligibilityUI([]); } catch {}
    if (!silent) showToast(mapServerErrorToMessage(error));
    return;
  }
  // Cache the entire server payload once; no client decisions
  emergencyEligibilityCache = { eligible };
  window._emergencyEligibleShifts = eligible;
  syncEmergencyButtonVisibility(eligible);
  try { typeof syncEmergencyEligibilityUI === 'function' && syncEmergencyEligibilityUI(eligible); } catch {}
}

// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
async function openEmergencyOverlay() {
  if (isBlockingOverlayOpen()) return;

  // Only open if eligibility already exists (button is visible only in that case anyway)
  const haveEligible =
    emergencyEligibilityCache &&
    Array.isArray(emergencyEligibilityCache.eligible) &&
    emergencyEligibilityCache.eligible.length > 0;

  if (!haveEligible) {
    // Ensure a high-priority background fetch; overlay stays closed
    prefetchEmergencyEligibility({ priority: 'high' }).catch(() => {});
    try { showToast('No eligible shifts right now.'); } catch {}
    return;
  }

  ensureEmergencyOverlay();
  const btn = document.getElementById('emergencyBtn');
  if (btn) { btn.dataset.busy = '1'; btn.style.opacity = '.9'; btn.style.pointerEvents = 'none'; }

  try {
    resetEmergencyState();
    emergencyState.eligible = emergencyEligibilityCache.eligible || [];
    emergencyState.step = 'PICK_SHIFT';

    openOverlay('emergencyOverlay', '#emergencyClose');
    renderEmergencyStep();
  } finally {
    if (btn) { delete btn.dataset.busy; btn.style.opacity = ''; btn.style.pointerEvents = ''; }
  }
}



function renderEmergencyStep() {
  const body = document.getElementById('emergencyBody');
  const title = document.getElementById('emergencyTitle');
  const footer = document.getElementById('emergencyFooter');
  if (!body || !title || !footer || !emergencyState) return;

  function setFooter({ showContinue = true, continueText = 'Continue', continueDisabled = false, showCancel = true, cancelText = 'Cancel' } = {}) {
    footer.innerHTML = '';
    if (showCancel) {
      const cancel = document.createElement('button');
      cancel.id = 'emergencyCancel';
      cancel.type = 'button';
      cancel.textContent = cancelText;
      cancel.className = 'menu-item';
      cancel.style.border = '1px solid #2a3446';
      cancel.addEventListener('click', () => closeEmergencyOverlay(false));
      footer.appendChild(cancel);
    }
    if (showContinue) {
      const cont = document.createElement('button');
      cont.id = 'emergencyContinue';
      cont.type = 'button';
      cont.textContent = continueText;
      cont.className = 'menu-item';
      cont.style.border = '1px solid #2a3446';
      cont.disabled = !!continueDisabled;
      footer.appendChild(cont);
    }
  }

  const state = emergencyState;
  body.innerHTML = '';

  if (state.step === 'PICK_SHIFT') {
    title.textContent = 'Which shift is this emergency concerning?';

    if (!Array.isArray(state.eligible) || state.eligible.length === 0) {
      body.innerHTML = `<div class="muted">No eligible shifts right now.</div>`;
      setFooter({ showContinue: false, showCancel: true, cancelText: 'Close' });
      return;
    }

    const list = document.createElement('div');
    list.setAttribute('role', 'list');
    list.style.display = 'flex';
    list.style.flexDirection = 'column';
    list.style.gap = '.5rem';

    state.eligible.forEach((shift, idx) => {
      const id = `emShift_${idx}`;
      const item = document.createElement('label');
      item.setAttribute('role', 'listitem');
      item.style.border = '1px solid #222a36';
      item.style.background = '#131926';
      item.style.borderRadius = '10px';
      item.style.padding = '.5rem .6rem';
      item.style.display = 'flex';
      item.style.alignItems = 'center';
      item.style.gap = '.5rem';
      item.innerHTML = `
        <input type="radio" name="emShift" id="${id}" value="${escapeHtml(shift.ymd || String(idx))}">
        <div style="font-weight:700;">${escapeHtml(buildEmergencyShiftLabel(shift))}</div>
      `;
      item.querySelector('input')._shift = shift; // stash whole server object
      list.appendChild(item);
    });

    body.appendChild(list);
    setFooter({ continueDisabled: true });

    const cont = document.getElementById('emergencyContinue');
    body.addEventListener('change', (e) => {
      if (e.target && e.target.name === 'emShift') {
        cont && (cont.disabled = false);
      }
    });
    cont && cont.addEventListener('click', handleEmergencyPickShift);
  }

  else if (state.step === 'PICK_ISSUE') {
    title.textContent = 'Please select your emergency';

    const s = state.selectedShift || {};
    let allowed = [];

    if (Array.isArray(s.allowed_issues) && s.allowed_issues.length) {
      allowed = s.allowed_issues.map(x => String(x).toUpperCase());
    } else if (Array.isArray(s.allowedActions) && s.allowedActions.length) {
      allowed = s.allowedActions.map(x => String(x).toUpperCase());
    } else {
      if (s.canRunLate === true)     allowed.push('RUNNING_LATE');
      if (s.canCancel === true)      allowed.push('CANNOT_ATTEND');
      if (s.canLeaveEarly === true)  allowed.push('LEAVE_EARLY');
      if (s.canDNA === true)         allowed.push('DNA'); // tolerant fallback
    }

    const choices = document.createElement('div');
    choices.setAttribute('role', 'group');
    choices.style.display = 'flex';
    choices.style.flexDirection = 'column';
    choices.style.gap = '.5rem';

    function addOption(key, label, help) {
      const id = `emIssue_${key}`;
      const row = document.createElement('label');
      row.style.border = '1px solid #222a36';
      row.style.background = '#131926';
      row.style.borderRadius = '10px';
      row.style.padding = '.5rem .6rem';
      row.style.display = 'grid';
      row.style.gridTemplateColumns = 'auto 1fr';
      row.style.gap = '.6rem';
      row.innerHTML = `
        <input type="radio" name="emIssue" id="${id}" value="${key}">
        <div>
          <div style="font-weight:800;">${escapeHtml(label)}</div>
          ${help ? `<div class="muted" style="margin-top:.15rem">${escapeHtml(help)}</div>` : ''}
        </div>
      `;
      choices.appendChild(row);
    }

    if (allowed.includes('RUNNING_LATE')) addOption('RUNNING_LATE', 'Running late', 'If the shift is within 4 hours or has already started.');
    if (allowed.includes('CANNOT_ATTEND')) addOption('CANNOT_ATTEND', 'Cannot attend shift', 'Only if your shift has not started yet.');
    if (allowed.includes('LEAVE_EARLY')) addOption('LEAVE_EARLY', 'Need to leave current shift early', 'Only if you are currently on shift.');
    if (allowed.includes('DNA')) addOption('DNA', 'Someone hasn’t arrived on shift yet', 'Select the person who has not arrived.');

    if (!choices.children.length) {
      body.innerHTML = `<div class="muted">No emergency options are available for this shift.</div>`;
      setFooter({ showCancel: true, cancelText: 'Close', showContinue: false });
      return;
    }

    body.appendChild(choices);
    setFooter({ continueDisabled: true });

    const cont = document.getElementById('emergencyContinue');
    body.addEventListener('change', (e) => {
      if (e.target && e.target.name === 'emIssue') {
        cont && (cont.disabled = false);
      }
    });
    cont && cont.addEventListener('click', handleEmergencyPickIssue);
  }

  else if (state.step === 'DETAILS') {
    // ───────────────────────── RUNNING_LATE ─────────────────────────
    if (state.issueType === 'RUNNING_LATE') {
      title.textContent = 'How late will you be? Please note that is how late you will arrive after the shift time is due to start.';

      const container = document.createElement('div');
      container.style.display = 'flex';
      container.style.flexDirection = 'column';
      container.style.gap = '.5rem';

      const optsWrap = document.createElement('div');
      container.appendChild(optsWrap);
      body.appendChild(container);
      setFooter({ continueDisabled: true });

      const s = state.selectedShift || {};
      const serverOpts = Array.isArray(s.late_options) ? s.late_options : null;
      const serverCtx  = s.context || null;

      (async () => {
        if (serverOpts && serverOpts.length) {
          state.lateOptions = serverOpts;
          state.runningLateContext = serverCtx || s;
        } else {
          optsWrap.innerHTML = `<div class="muted">Loading options…</div>`;
          const { ok, options, context, eligible, reason, error } = await apiPostRunningLateOptions(s);
          if (!ok) { optsWrap.innerHTML = `<div class="muted">${escapeHtml(mapServerErrorToMessage(error))}</div>`; return; }
          if (eligible === false) { optsWrap.innerHTML = `<div class="muted">${escapeHtml(reason || 'Not eligible right now.')}</div>`; return; }
          state.lateOptions = options;
          state.runningLateContext = context || s;
        }

        optsWrap.innerHTML = '';
        const group = document.createElement('div');
        group.setAttribute('role', 'group');
        group.style.display = 'flex';
        group.style.flexDirection = 'column';
        group.style.gap = '.5rem';

        (state.lateOptions || []).forEach((opt, idx) => {
          const id = `lateOpt_${idx}`;
          const row = document.createElement('label');
          row.style.border = '1px solid #222a36';
          row.style.background = '#131926';
          row.style.borderRadius = '10px';
          row.style.padding = '.5rem .6rem';
          row.style.display = 'flex';
          row.style.alignItems = 'center';
          row.style.gap = '.6rem';
          const labelText = typeof opt === 'string' ? opt : (opt.label || '');
          row.innerHTML = `
            <input type="radio" name="emLate" id="${id}" value="${escapeHtml(labelText)}">
            <div style="font-weight:700;">${escapeHtml(labelText)}</div>
          `;
          group.appendChild(row);
        });

        optsWrap.appendChild(group);

        const cont = document.getElementById('emergencyContinue');
        body.addEventListener('change', (e) => {
          if (e.target && e.target.name === 'emLate') {
            cont && (cont.disabled = false);
            state.selectedLateLabel = e.target.value || null;
            if (state.selectedLateLabel) state.confirmError = '';
          }
        });
      })();

      const cont = document.getElementById('emergencyContinue');
      cont && cont.addEventListener('click', handleEmergencyCollectDetails);
    }

    // ───────────────────────── CANNOT_ATTEND ─────────────────────────
    else if (state.issueType === 'CANNOT_ATTEND') {
      title.textContent = 'Please provide a reason for your very late cancellation';
      const p = document.createElement('div');
      p.className = 'muted';
      p.style.marginBottom = '.5rem';
      p.textContent = 'We will notify our office and they will be called on their phones. If this is not an emergency you will be removed from all future care packages.';
      const ta = document.createElement('textarea');
      ta.id = 'emergencyReason';
      ta.rows = 4;
      ta.style.width = '100%';
      ta.style.borderRadius = '10px';
      ta.style.border = '1px solid #222a36';
      ta.style.background = '#0b0e14';
      ta.style.color = '#e7ecf3';
      ta.style.padding = '.6rem';
      ta.placeholder = 'Type your reason…';
      body.append(p, ta);

      setFooter({ continueDisabled: true });

      const cont = document.getElementById('emergencyContinue');
      ta.addEventListener('input', () => {
        emergencyState.reasonText = (ta.value || '').trim();
        const valid = emergencyState.reasonText.length >= 3;
        cont && (cont.disabled = !valid);
        if (valid) emergencyState.confirmError = '';
      });
      cont && cont.addEventListener('click', handleEmergencyCollectDetails);
    }

    // ───────────────────────── LEAVE_EARLY ─────────────────────────
    else if (state.issueType === 'LEAVE_EARLY') {
      title.textContent = 'Reason for leaving early';
      const p = document.createElement('div');
      p.className = 'muted';
      p.style.marginBottom = '.5rem';
      p.textContent = 'Our office staff will be called on their mobile phones immediately regarding this.';
      const ta = document.createElement('textarea');
      ta.id = 'emergencyReason';
      ta.rows = 4;
      ta.style.width = '100%';
      ta.style.borderRadius = '10px';
      ta.style.border = '1px solid #222a36';
      ta.style.background = '#0b0e14';
      ta.style.color = '#e7ecf3';
      ta.style.padding = '.6rem';
      ta.placeholder = 'Type your reason…';

      // time selector
      const timeWrap = document.createElement('div');
      timeWrap.style.marginTop = '.6rem';
      timeWrap.style.border = '1px solid #222a36';
      timeWrap.style.background = '#131926';
      timeWrap.style.borderRadius = '10px';
      timeWrap.style.padding = '.6rem';

      const legend = document.createElement('div');
      legend.style.fontWeight = '800';
      legend.style.marginBottom = '.35rem';
      legend.textContent = 'When do you need to leave?';
      timeWrap.appendChild(legend);

      const help = document.createElement('div');
      help.className = 'muted';
      help.style.marginBottom = '.4rem';
      help.textContent = 'Choose NOW or enter a 24-hour time (HH MM).';
      timeWrap.appendChild(help);

      const rowNow = document.createElement('label');
      rowNow.style.display = 'flex';
      rowNow.style.alignItems = 'center';
      rowNow.style.gap = '.5rem';
      rowNow.style.marginBottom = '.35rem';
      rowNow.innerHTML = `<input type="radio" name="leaveTimeMode" value="NOW"> <div>NOW / ASAP</div>`;

      const rowAt = document.createElement('div');
      rowAt.style.display = 'grid';
      rowAt.style.gridTemplateColumns = 'auto auto auto 1fr';
      rowAt.style.alignItems = 'center';
      rowAt.style.gap = '.4rem';

      const atLabel = document.createElement('label');
      atLabel.style.display = 'flex';
      atLabel.style.alignItems = 'center';
      atLabel.style.gap = '.5rem';
      atLabel.innerHTML = `<input type="radio" name="leaveTimeMode" value="AT"> <div>At time:</div>`;

      const hh = document.createElement('input');
      hh.type = 'text';
      hh.inputMode = 'numeric';
      hh.maxLength = 2;
      hh.placeholder = 'HH';
      hh.style.width = '3.5rem';
      hh.style.textAlign = 'center';
      hh.style.borderRadius = '8px';
      hh.style.border = '1px solid #222a36';
      hh.style.background = '#0b0e14';
      hh.style.color = '#e7ecf3';
      hh.style.padding = '.35rem';

      const mm = document.createElement('input');
      mm.type = 'text';
      mm.inputMode = 'numeric';
      mm.maxLength = 2;
      mm.placeholder = 'MM';
      mm.style.width = '3.5rem';
      mm.style.textAlign = 'center';
      mm.style.borderRadius = '8px';
      mm.style.border = '1px solid #222a36';
      mm.style.background = '#0b0e14';
      mm.style.color = '#e7ecf3';
      mm.style.padding = '.35rem';

      const err = document.createElement('div');
      err.className = 'muted';
      err.style.color = '#ff8a80';
      err.style.marginTop = '.3rem';
      err.style.display = 'none';
      err.textContent = 'Enter a valid 24-hour time (00–23 and 00–59).';

      rowAt.appendChild(atLabel);
      rowAt.appendChild(hh);
      rowAt.appendChild(mm);

      timeWrap.appendChild(rowNow);
      timeWrap.appendChild(rowAt);
      timeWrap.appendChild(err);

      body.append(p, ta, timeWrap);

      setFooter({ continueDisabled: true });

      const cont = document.getElementById('emergencyContinue');
      emergencyState.selectedLateLabel = null;

      function isNumInRange(v, min, max) {
        const n = Number(v);
        return String(v).trim() !== '' && Number.isInteger(n) && n >= min && n <= max;
      }
      function pad2(n) { n = Number(n); return (n < 10 ? '0' + n : String(n)); }

      function updateTimeSelection() {
        const modeNow = body.querySelector('input[name="leaveTimeMode"][value="NOW"]');
        const modeAt  = body.querySelector('input[name="leaveTimeMode"][value="AT"]');
        const nowChecked = modeNow && modeNow.checked;
        const atChecked  = modeAt && modeAt.checked;

        let ok = false;
        let label = null;

        if (nowChecked) {
          label = 'NOW';
          ok = true;
          err.style.display = 'none';
        } else if (atChecked) {
          const H = hh.value.replace(/\D+/g, '');
          const M = mm.value.replace(/\D+/g, '');
          const hOk = isNumInRange(H, 0, 23);
          const mOk = isNumInRange(M, 0, 59);
          ok = hOk && mOk;
          if (ok) {
            label = `${pad2(H)}:${pad2(M)}`;
            err.style.display = 'none';
          } else {
            if (H !== '' || M !== '') err.style.display = '';
          }
        } else {
          err.style.display = 'none';
        }

        emergencyState.selectedLateLabel = ok ? label : null;

        const reasonOK = (emergencyState.reasonText || '').length >= 3;
        cont && (cont.disabled = !(reasonOK && !!emergencyState.selectedLateLabel));

        if (reasonOK && !!emergencyState.selectedLateLabel) {
          emergencyState.confirmError = '';
        }
      }

      body.addEventListener('change', (e) => {
        if (e.target && e.target.name === 'leaveTimeMode') {
          updateTimeSelection();
        }
      });
      [hh, mm].forEach(inp => {
        inp.addEventListener('input', updateTimeSelection);
        inp.addEventListener('blur', updateTimeSelection);
      });

      ta.addEventListener('input', () => {
        emergencyState.reasonText = (ta.value || '').trim();
        updateTimeSelection();
      });

      const modeNowRadio = body.querySelector('input[name="leaveTimeMode"][value="NOW"]');
      modeNowRadio && (modeNowRadio.checked = false);
      updateTimeSelection();

      cont && cont.addEventListener('click', handleEmergencyCollectDetails);
    }

    // ───────────────────────── DNA ─────────────────────────
    else if (state.issueType === 'DNA') {
  title.textContent = 'Who hasn’t arrived?';

  const s = state.selectedShift || {};

  const peers =
    (Array.isArray(s.cohort) && s.cohort) ||
    (Array.isArray(s.cohort_peers) && s.cohort_peers) ||
    (Array.isArray(s.peers) && s.peers) ||
    (Array.isArray(s.shiftPeers) && s.shiftPeers) ||
    [];

  state.dnaOptions = Array.isArray(state.dnaOptions) && state.dnaOptions.length ? state.dnaOptions : peers;

  const info = document.createElement('div');
  info.className = 'muted';
  info.style.marginBottom = '.5rem';
  info.textContent = 'Select the co-worker who has not arrived.';
  body.appendChild(info);

  const list = document.createElement('div');
  list.setAttribute('role', 'group');
  list.style.display = 'flex';
  list.style.flexDirection = 'column';
  list.style.gap = '.5rem';

  function getDisplayName(p) {
    if (!p) return '';
    if (p.displayName) return String(p.displayName);
    const fn = p.firstName || p.firstname || '';
    const ln = p.surname || p.lastName || p.lastname || p.last_name || '';
    const nameKey = p.nameKey || '';
    const full = [fn, ln].filter(Boolean).join(' ').trim();
    return full || p.name || nameKey || '';
  }

  (state.dnaOptions || []).forEach((p, idx) => {
    const id = `dnaOpt_${idx}`;
    const name = getDisplayName(p);
    const role = p.role || p.jobTitle || p.title || '';
    const ms = p.msisdn07 || p.msisdn || p.mobile07 || p.mobile || null;

    const row = document.createElement('label');
    row.style.border = '1px solid #222a36';
    row.style.background = '#131926';
    row.style.borderRadius = '10px';
    row.style.padding = '.5rem .6rem';
    row.style.display = 'grid';
    row.style.gridTemplateColumns = 'auto 1fr';
    row.style.gap = '.6rem';
    row.innerHTML = `
      <input type="radio" name="dnaAbsentee" id="${id}" value="${escapeHtml(name)}">
      <div>
        <div style="font-weight:800;">${escapeHtml(name)}</div>
        ${role ? `<div class="muted" style="margin-top:.15rem">${escapeHtml(role)}</div>` : ''}
      </div>
    `;
    row.querySelector('input')._abs = { name, msisdn: ms || null };
    list.appendChild(row);
  });

  if (!list.children.length) {
    body.innerHTML = `<div class="muted">We can’t find your co-workers for this shift yet. Please go back and try again shortly.</div>`;
    setFooter({ showCancel: true, cancelText: 'Back', showContinue: false });
    return;
  }

  body.appendChild(list);

  // Instruction block (populated when absentee is selected)
  const callInfo = document.createElement('div');
  callInfo.className = 'muted';
  callInfo.style.marginTop = '.6rem';
  body.appendChild(callInfo);
// Tried-calling checkbox
const chkWrap = document.createElement('label');
chkWrap.style.display = 'flex';
chkWrap.style.alignItems = 'center';
chkWrap.style.gap = '.5rem';
chkWrap.style.marginTop = '.6rem';
chkWrap.innerHTML = `<input type="checkbox" id="dnaTriedCalling"> <div>I tried calling them and could not reach them</div>`;
body.appendChild(chkWrap);

setFooter({ continueDisabled: true });

const cont = document.getElementById('emergencyContinue');
const chk = chkWrap.querySelector('#dnaTriedCalling');


function syncContinueGate() {
  const chosen = body.querySelector('input[name="dnaAbsentee"]:checked');
  if (chosen && chosen._abs) {
    emergencyState.dnaAbsenteeName = chosen._abs.name || null;
    emergencyState.dnaAbsenteeMsisdn = chosen._abs.msisdn || null;

    const dispNum = formatMsisdn07Display(chosen._abs.msisdn);
    if (dispNum) {
      callInfo.textContent = `Please call ${chosen._abs.name} on ${dispNum}. If you can’t get a response after trying a couple of times, tick below and continue.`;
    } else {
      callInfo.textContent = `Please try calling ${chosen._abs.name} via your usual contact method. If you can’t get a response after trying a couple of times, tick below and continue.`;
    }
  }

  const ok = !!chosen && !!chk.checked;
  cont && (cont.disabled = !ok);
  emergencyState.dnaTriedCalling = !!chk.checked;
  emergencyState.confirmError = '';
}

  body.addEventListener('change', (e) => {
    if (e.target && (e.target.name === 'dnaAbsentee')) {
      syncContinueGate();
    }
    if (e.target && e.target.id === 'dnaTriedCalling') {
      syncContinueGate();
    }
  });

  cont && cont.addEventListener('click', handleEmergencyCollectDetails);
}

  }

  else if (state.step === 'CONFIRM') {
    let msg = '';
    if (state.issueType === 'RUNNING_LATE') {
      msg = "Your colleagues on shift will be informed and so will the Arthur Rai office. Do you want to proceed? You cannot cancel after this step.";
    } else if (state.issueType === 'CANNOT_ATTEND') {
      msg = "Warning – Our office staff will now be called on their mobile phones regarding your cancellation. Are you absolutely sure you want to continue?";
    } else if (state.issueType === 'LEAVE_EARLY') {
      msg = "Are you sure you wish to continue? Our office staff will be called on their mobile phones immediately regarding this.";
    } else if (state.issueType === 'DNA') {
      const who = emergencyState.dnaAbsenteeName || 'this person';
      msg = `Are you sure you want to raise an emergency and contact Arthur Rai to inform them that ${who} hasn’t arrived on shift?`;
    }

    title.textContent = 'Please confirm';
    const box = document.createElement('div');
    box.style.border = '1px solid #2a3446';
    box.style.background = '#131926';
    box.style.borderRadius = '10px';
    box.style.padding = '.6rem';
    box.style.lineHeight = '1.35';
    box.innerHTML = `<div class="muted">${escapeHtml(msg)}</div>`;

    if (emergencyState.previewHtml && emergencyState.issueType === 'RUNNING_LATE') {
      const prev = document.createElement('div');
      prev.style.marginTop = '.6rem';
      prev.innerHTML = emergencyState.previewHtml;
      box.appendChild(prev);
    }

    if (state.confirmError) {
      const warn = document.createElement('div');
      warn.style.marginTop = '.6rem';
      warn.style.fontWeight = '700';
      warn.style.color = '#ff8a80';
      warn.textContent = state.confirmError;
      box.appendChild(warn);
    }

    body.appendChild(box);

    setFooter({ continueText: 'Yes, I am sure', cancelText: 'Cancel', continueDisabled: !!state.confirmError });
    const cont = document.getElementById('emergencyContinue');
    cont && cont.addEventListener('click', handleEmergencyConfirm);
  }
}



// Prefer server-provided label/time; fall back to old label composition.
function buildEmergencyShiftLabel(shift) {
  if (shift && typeof shift.display_label === 'string' && shift.display_label.trim()) {
    return shift.display_label.trim();
  }
  const datePart =
    (shift && shift.date_label) ||
    (() => {
      const ymd = String(shift && shift.ymd || '');
      const [Y, M, D] = ymd.split('-').map(n => parseInt(n, 10));
      const dt = (Y && M && D) ? new Date(Y, M - 1, D) : null;
      return dt ? dt.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' }).replace(/,/g, '') : (ymd || 'Date');
    })();

  const type =
    (shift && (shift.shift_type_label || shift.shift_type_raw || shift.shift_type)) ||
    ((String(shift && shift.shiftInfo || '').toUpperCase().includes('NIGHT')) ? 'NIGHT' : 'LONG DAY');

  const timeRange =
    (shift && (shift.time_range_label || shift.timeLabel)) || // server-supplied pretty range preferred
    '';

  const hosp = (shift && (shift.hospital || shift.location)) || '';
  const ward = (shift && shift.ward) || '';
  const ref  = (shift && (shift.booking_ref || shift.bookingRef)) || '';

  const parts = [
    String(datePart || '').trim(),
    String(type || '').trim(),
    timeRange ? String(timeRange).trim() : '',
    hosp ? String(hosp).trim() : '',
    ward ? `(${String(ward).trim()})` : '',
    ref ? `Ref: ${String(ref).trim()}` : ''
  ];
  return parts.filter(Boolean).join(' — ');
}

async function apiGetEmergencyWindow() {
  try {
    const { res, json } = await apiGET({ view: 'emergency_window' });
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && (json.error || json.message)) || `HTTP ${res.status}`;
      return { ok: false, eligible: [], error: msg || 'UNKNOWN_ERROR' };
    }
    // Pass through whatever the server decided (eligible items may include
    // allowed_issues, late_options, context, labels, etc.)
    const eligible = Array.isArray(json.eligible) ? json.eligible : [];
    return { ok: true, eligible, serverTime: json.serverTime || null };
  } catch (e) {
    return { ok: false, eligible: [], error: String(e && e.message) || 'NETWORK_ERROR' };
  }
}

async function handleEmergencyCollectDetails() {
  if (!emergencyState) return;

  if (emergencyState.issueType === 'RUNNING_LATE') {
    if (!emergencyState.selectedLateLabel) {
      emergencyState.confirmError = 'Please pick how late you will be.';
      emergencyState.step = 'CONFIRM';
      renderEmergencyStep();
      return;
    }

    const s = emergencyState.selectedShift || {};
    const prePreview = (s.previewHtml || (s.preview && s.preview.html)) || '';

    if (prePreview) {
      emergencyState.previewHtml = prePreview;
      emergencyState.step = 'CONFIRM';
      renderEmergencyStep();
      return;
    }

    const ctxForPreview =
      (emergencyState.runningLateContext && (emergencyState.runningLateContext.context || emergencyState.runningLateContext)) ||
      (s && (s.context || s));

    if (ctxForPreview) {
      const { ok, previewHtml } = await apiPostRunningLatePreview(ctxForPreview, emergencyState.selectedLateLabel);
      emergencyState.previewHtml = ok ? (previewHtml || '') : '';
    } else {
      emergencyState.previewHtml = '';
    }

    emergencyState.confirmError = '';
    emergencyState.step = 'CONFIRM';
    renderEmergencyStep();
    return;
  }

  if (emergencyState.issueType === 'DNA') {
    const nameOK = !!(emergencyState.dnaAbsenteeName && emergencyState.dnaAbsenteeName.trim());
    const calledOK = !!emergencyState.dnaTriedCalling;
    if (!nameOK || !calledOK) {
      emergencyState.confirmError = !nameOK
        ? 'Please select who has not arrived.'
        : 'Please confirm you tried calling them before continuing.';
      emergencyState.step = 'CONFIRM';
      renderEmergencyStep();
      return;
    }
    emergencyState.confirmError = '';
    emergencyState.step = 'CONFIRM';
    renderEmergencyStep();
    return;
  }

  // CANNOT_ATTEND / LEAVE_EARLY
  if (!emergencyState.reasonText || emergencyState.reasonText.length < 3) {
    emergencyState.confirmError = 'Please enter a reason before continuing.';
    emergencyState.step = 'CONFIRM';
    renderEmergencyStep();
    return;
  }
  emergencyState.confirmError = '';
  emergencyState.step = 'CONFIRM';
  renderEmergencyStep();
}

function formatMsisdn07Display(msisdn) {
  if (!msisdn) return '';
  let s = String(msisdn).replace(/\s+/g, '').trim();

  // +447xxxxxxxxx → 07xxxxxxxxx
  if (/^\+447\d{9}$/.test(s)) return '0' + s.slice(3);

  // 447xxxxxxxxx → 07xxxxxxxxx
  if (/^447\d{9}$/.test(s)) return '0' + s.slice(2);

  // Already in 07xxxxxxxxx form
  if (/^07\d{9}$/.test(s)) return s;

  // Fallback: strip non-digits and try again
  const digits = s.replace(/[^\d]/g, '');
  if (digits.startsWith('447') && digits.length === 12) {
    return '0' + digits.slice(2);
  }
  if (digits.startsWith('07') && digits.length === 11) {
    return digits;
  }

  // Couldn’t normalise → return raw input
  return s;
}


async function handleEmergencyConfirm() {
  if (!emergencyState || !emergencyState.selectedShift || !emergencyState.issueType) return;

  // --- double-submit guard ---
  if (emergencyState.__confirmSubmitting) return;
  emergencyState.__confirmSubmitting = true;

  const cont = document.getElementById('emergencyContinue');
  const cancel = document.getElementById('emergencyCancel');
  const origText = cont ? cont.textContent : '';

  if (cont) { cont.disabled = true; cont.textContent = 'Sending…'; }
  if (cancel) cancel.disabled = true;

  function reenable() {
    if (cont) { cont.disabled = false; cont.textContent = origText || 'Continue'; }
    if (cancel) cancel.disabled = false;
    emergencyState.__confirmSubmitting = false;
  }

  try {
    if (emergencyState.issueType === 'RUNNING_LATE') {
      const s = emergencyState.selectedShift || {};
      const ctx =
        (emergencyState.runningLateContext && (emergencyState.runningLateContext.context || emergencyState.runningLateContext)) ||
        (s && (s.context || s));

      if (!emergencyState.selectedLateLabel || !ctx) {
        emergencyState.confirmError = 'Please pick how late you will be.';
        emergencyState.step = 'CONFIRM';
        reenable();
        renderEmergencyStep();
        return;
      }

      await submitEmergencyRunningLate(ctx, emergencyState.selectedLateLabel);

    } else if (emergencyState.issueType === 'LEAVE_EARLY') {
      const reasonOk = (emergencyState.reasonText || '').trim().length >= 3;
      const lbl = (emergencyState.selectedLateLabel || '').trim();
      const isNow = lbl === 'NOW';
      let isValidTime = false;

      if (/^\d{2}:\d{2}$/.test(lbl)) {
        const parts = lbl.split(':');
        const hh = parseInt(parts[0], 10);
        const mm = parseInt(parts[1], 10);
        isValidTime = Number.isInteger(hh) && Number.isInteger(mm) && hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59;
      }

      if (!reasonOk) {
        emergencyState.confirmError = 'Please enter a reason before continuing.';
        emergencyState.step = 'CONFIRM';
        reenable();
        renderEmergencyStep();
        return;
      }
      if (!(isNow || isValidTime)) {
        emergencyState.confirmError = 'Please select NOW or enter a valid time in 24-hour format (HH:MM).';
        emergencyState.step = 'CONFIRM';
        reenable();
        renderEmergencyStep();
        return;
      }

      await submitEmergencyRaise(emergencyState.selectedShift, 'LEAVE_EARLY', emergencyState.reasonText || '', emergencyState.selectedLateLabel);

    } else if (emergencyState.issueType === 'DNA') {
      const nameOK = !!(emergencyState.dnaAbsenteeName && emergencyState.dnaAbsenteeName.trim());
      const calledOK = !!emergencyState.dnaTriedCalling;
      if (!nameOK || !calledOK) {
        emergencyState.confirmError = !nameOK
          ? 'Please select who has not arrived.'
          : 'Please confirm you tried calling them before continuing.';
        emergencyState.step = 'CONFIRM';
        reenable();
        renderEmergencyStep();
        return;
      }

      await submitEmergencyRaise(emergencyState.selectedShift, 'DNA', '', 'N/A');

    } else {
      // CANNOT_ATTEND
      const reasonOk = (emergencyState.reasonText || '').trim().length >= 3;
      if (!reasonOk) {
        emergencyState.confirmError = 'Please enter a reason before continuing.';
        emergencyState.step = 'CONFIRM';
        reenable();
        renderEmergencyStep();
        return;
      }
      await submitEmergencyRaise(emergencyState.selectedShift, 'CANNOT_ATTEND', emergencyState.reasonText || '', 'N/A');
    }
  } catch (err) {
    emergencyState.confirmError = 'Something went wrong. Please try again.';
    emergencyState.step = 'CONFIRM';
    reenable();
    renderEmergencyStep();
    return;
  }

  // Success → keep disabled; submitters will replace footer/overlay
}

async function submitEmergencyRaise(shift, issueType, reasonText, etaOrLeaveTimeLabel) {
  const bodyEl = document.getElementById('emergencyBody');
  const title = document.getElementById('emergencyTitle');
  const footer = document.getElementById('emergencyFooter');
  if (!bodyEl || !title || !footer || !emergencyState) return;

  // Compute robust defaults
  const it = String(issueType || '').toUpperCase();
  const leaveLabel = (it === 'LEAVE_EARLY') ? (emergencyState.selectedLateLabel || etaOrLeaveTimeLabel || 'N/A') : 'N/A';
  const reason = String(reasonText || emergencyState.reasonText || '').trim();

  // Build a tolerant issue payload for apiPostEmergencyRaise
  let issuePayload;
  if (it === 'DNA') {
    issuePayload = {
      type: 'DNA',
      subject_name: emergencyState.dnaAbsenteeName || '',
      // Optional but recommended; server will normalise if present
      subject_msisdn: emergencyState.dnaAbsenteeMsisdn || undefined,
      // Optional – server also auto-composes DNA reason text if omitted
      reason_text: reason || ''
    };
  } else {
    issuePayload = {
      issue_type: it, // legacy shape
      eta_or_leave_time_label: leaveLabel || '',
      reason_text: reason || ''
    };
  }

  showLoading('Sending…');
  try {
    const { ok, error } = await apiPostEmergencyRaise(shift, it, issuePayload);

    if (!ok) {
      bodyEl.innerHTML = `<div class="muted">${escapeHtml(mapServerErrorToMessage(error))}</div>`;
      footer.innerHTML = '';
      const cancel = document.createElement('button');
      cancel.type = 'button';
      cancel.textContent = 'Cancel';
      cancel.className = 'menu-item';
      cancel.style.border = '1px solid #2a3446';
      cancel.addEventListener('click', () => closeEmergencyOverlay(false));
      const retry = document.createElement('button');
      retry.type = 'button';
      retry.textContent = 'Try again';
      retry.className = 'menu-item';
      retry.style.border = '1px solid #2a3446';
      retry.addEventListener('click', handleEmergencyConfirm);
      footer.append(cancel, retry);
      return;
    }

    title.textContent = 'Sent';
    let successMsg = 'Our office have been informed and will call you shortly.';
    if (it === 'CANNOT_ATTEND') {
      successMsg = 'Our office have been informed. Please expect a phonecall shortly.';
    } else if (it === 'DNA') {
      successMsg = 'This emergency has been raised with Arthur Rai and someone will be in contact shortly.';
    }
    bodyEl.innerHTML = `<div>${escapeHtml(successMsg)}</div>`;
    footer.innerHTML = '';

    const x = document.getElementById('emergencyClose');
    if (x) {
      const newX = x.cloneNode(true);
      x.parentNode.replaceChild(newX, x);
      newX.addEventListener('click', () => closeEmergencyOverlay(true));
    }

    const close = document.createElement('button');
    close.type = 'button';
    close.textContent = 'Close';
    close.className = 'menu-item';
    close.style.border = '1px solid #2a3446';
    close.addEventListener('click', () => closeEmergencyOverlay(true));
    footer.appendChild(close);
  } finally {
    hideLoading();
  }
}


function closeEmergencyOverlay(reload = false) {
  closeOverlay('emergencyOverlay', /*force*/true);
  emergencyState = null;

  if (reload === true) {
    // Guard against multiple rapid clicks
    if (!window.__reloadingFromEmergency) {
      window.__reloadingFromEmergency = true;
      // Allow overlay close animation to finish before reload
      setTimeout(() => { window.location.reload(); }, 150);
    }
  }
}

async function submitEmergencyRunningLate(ctx, label) {
  const body = document.getElementById('emergencyBody');
  const title = document.getElementById('emergencyTitle');
  const footer = document.getElementById('emergencyFooter');
  if (!body || !title || !footer || !emergencyState) return;

  const resolvedCtx =
    (ctx && (ctx.context || ctx)) ||
    (emergencyState.runningLateContext && (emergencyState.runningLateContext.context || emergencyState.runningLateContext)) ||
    (emergencyState.selectedShift && (emergencyState.selectedShift.context || emergencyState.selectedShift));

  const resolvedLabel = label || emergencyState.selectedLateLabel;

  if (!resolvedCtx || !resolvedLabel) {
    showToast('Please pick how late you will be.');
    return;
  }

  showLoading('Sending…');
  try {
    const { ok, error } = await apiPostRunningLateSend(resolvedCtx, resolvedLabel);
    if (!ok) {
      body.innerHTML = `<div class="muted">${escapeHtml(mapServerErrorToMessage(error))}</div>`;
      footer.innerHTML = '';
      const cancel = document.createElement('button');
      cancel.type = 'button';
      cancel.textContent = 'Cancel';
      cancel.className = 'menu-item';
      cancel.style.border = '1px solid #2a3446';
      cancel.addEventListener('click', () => closeEmergencyOverlay(false));
      const retry = document.createElement('button');
      retry.type = 'button';
      retry.textContent = 'Try again';
      retry.className = 'menu-item';
      retry.style.border = '1px solid #2a3446';
      retry.addEventListener('click', handleEmergencyConfirm);
      footer.append(cancel, retry);
      return;
    }

    title.textContent = 'Sent';
    body.innerHTML = `<div> Your colleagues on shift and the Arthur Rai office have been informed.</div>`;
    footer.innerHTML = '';

    // Make the header ✕ (close) also trigger a reload after success
    const x = document.getElementById('emergencyClose');
    if (x) {
      const newX = x.cloneNode(true);
      x.parentNode.replaceChild(newX, x);
      newX.addEventListener('click', () => closeEmergencyOverlay(true));
    }

    const close = document.createElement('button');
    close.type = 'button';
    close.textContent = 'Close';
    close.className = 'menu-item';
    close.style.border = '1px solid #2a3446';
    close.addEventListener('click', () => closeEmergencyOverlay(true));
    footer.appendChild(close);
  } finally {
    hideLoading();
  }
}




// Fallback endpoints kept for backward compatibility (only used if the server
// did NOT include late options/context in the initial emergency_window payload)
async function apiPostRunningLateOptions(shift) {
  try {
    const { res, json } = await apiPOST({ action: 'RUNNING_LATE_OPTIONS', shift });
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && (json.error || json.message)) || `HTTP ${res.status}`;
      return { ok: false, options: [], context: null, eligible: false, error: msg || 'UNKNOWN_ERROR' };
    }
    const options  = Array.isArray(json.options) ? json.options : [];
    const context  = json.context || null;
    const eligible = !!json.eligible;
    const reason   = json.reason || '';
    return { ok: true, options, context, eligible, reason };
  } catch (e) {
    return { ok: false, options: [], context: null, eligible: false, error: String(e && e.message) || 'NETWORK_ERROR' };
  }
}

async function apiPostRunningLatePreview(shiftOrContext, etaLabel) {
  try {
    const context = shiftOrContext && shiftOrContext.context ? shiftOrContext.context : shiftOrContext;
    const { res, json } = await apiPOST({ action: 'RUNNING_LATE_PREVIEW', context, eta_label: etaLabel });
    if (!res.ok || !json || json.ok === false) {
      const msg = (json && (json.error || json.message)) || `HTTP ${res.status}`;
      return { ok: false, previewHtml: '', context: null, minutes: 0, arrivalByLabel: '', error: msg || 'UNKNOWN_ERROR' };
    }
    return {
      ok: true,
      previewHtml: json.previewHtml || '',
      context: json.context || context || null,
      minutes: Number(json.minutes || 0),
      arrivalByLabel: (json.preview && json.preview.arrival_by_label) || ''
    };
  } catch (e) {
    return { ok: false, previewHtml: '', context: null, minutes: 0, arrivalByLabel: '', error: String(e && e.message) || 'NETWORK_ERROR' };
  }
}

// ───────────────────────── NEW helper ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
function updateEmergencyButtonFromCache() {
  try {
    const hasEligible = !!(emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible) && emergencyEligibilityCache.eligible.length);
    if (els.emergencyBtn) {
      // Only show the button when eligibility exists; otherwise keep it invisible
      els.emergencyBtn.hidden   = !hasEligible;
      els.emergencyBtn.disabled = !hasEligible; // optional but safe
    }
  } catch {}
}


// ───────────────────────── NEW orchestrator ─────────────────────────
// Centralised refresh: invalidate caches, then call loadFromServer() once.
// ───────────────────────── CHANGED ─────────────────────────
async function refreshAppState({ force = true } = {}) {
  // Invalidate emergency cache so we’ll refresh it (in the background) after baseline
  emergencyEligibilityCache = null;

  // Refresh baseline (single-flight inside loadFromServer)
  const out = await loadFromServer({ force });

  // Fire-and-forget: refresh emergency eligibility in background after baseline refresh
  prefetchEmergencyEligibility().catch(() => {});

  return out;
}

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

  // If a blocking overlay is active, ignore route changes
  if (isBlockingOverlayOpen()) return;

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
// ───────────────────────── CHANGED (IIFE) ─────────────────────────
(async function init() {
  identity = loadSavedIdentity();

  ensureLoadingOverlay();
  ensureMenu();
  ensureEmergencyButton(); // idempotent creation of the button

  const hasK = new URLSearchParams(location.search).has('k');
  if (hasK) {
    openResetOverlay();
  } else if (!identity.msisdn) {
    try { if (els.emergencyBtn) { els.emergencyBtn.hidden = true; els.emergencyBtn.disabled = true; } } catch {}
    const autoOK = await tryAutoLoginViaCredentialsAPI();
    if (!autoOK && !isBlockingOverlayOpen()) openLoginOverlay();
  }

  loadDraft();
  if (identity && identity.msisdn) {
    showLoading();
    try {
      await loadFromServer({ force: true }); // single orchestrated fetch
    } catch (e) {
      if (String(e && e.message) !== '__AUTH_STOP__') {
        showToast('Load failed: ' + e.message, 5000);
      }
    } finally {
      sizeGrid();
      equalizeTileHeights({ reset: true });
      hideLoading();

      // Button wiring only; visibility already set from cache by loadFromServer()
      try { typeof wireEmergencyButton === 'function' && wireEmergencyButton(); } catch {}
      ensurePastShiftsButton();
    }
  }

  routeFromURL();
})();
