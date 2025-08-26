/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL     = 'https://script.google.com/macros/s/AKfycbwGWn3JxvP7oze5sw_CuUNEChqm0E6nZwRa_g_iX21_lkhbWJDINibKBUwcPlGiR3Nnyw/exec'; // <-- REPLACE
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
    #loginOverlay, #forgotOverlay, #resetOverlay, #alertOverlay {
      position: fixed; inset: 0; z-index: 9998;
      display: none; align-items: stretch; justify-content: center;
      background: rgba(10,12,16,0.55); backdrop-filter: blur(2px);
    }
    #pastOverlay.show, #contentOverlay.show, #welcomeOverlay.show,
    #loginOverlay.show, #forgotOverlay.show, #resetOverlay.show, #alertOverlay.show { display: flex; }

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
    { overlayId:'forgotOverlay',  closeId:'forgotClose' }
    // resetOverlay: NO closeId (non-dismissable)
    // alertOverlay: NON-dismissable (wired via showBlockingAlert)
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
    ['pastOverlay','contentOverlay','welcomeOverlay','loginOverlay','forgotOverlay','resetOverlay','alertOverlay'].forEach(id => {
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

function renderTiles() {
  if (!els.grid) { console.warn('renderTiles(): #grid not found'); return; }
  els.grid.innerHTML = '';
  if (!baseline || !baseline.tiles) return;

  ensureHelpMessageVisible();
  ensurePastShiftsButton();

  // Build once from baseline + current draft
  const tiles = (baseline.tiles || []).map(t => {
    const baselineLabel = t.booked
      ? 'BOOKED'
      : (t.status === 'BLOCKED' ? 'BLOCKED' : codeToLabel(t.status));
    const effectiveLabel = draft[t.ymd] ?? baselineLabel;
    return { ...t, baselineLabel, effectiveLabel };
  });

  // Helper: update only one card's UI from a label
  function updateCardUI(card, ymd, baselineLabel, newLabel, isBooked, isBlocked, editable) {
    const statusEl = card.querySelector('.tile-status');
    const subEl    = card.querySelector('.tile-sub');

    // Reset status class then apply the new one
    statusEl.className = 'tile-status ' + statusClass(newLabel);

    if (isBooked) {
      // Booked tiles never change via taps (not editable), but keep content stable
      statusEl.textContent = 'BOOKED';
      // subEl remains as originally rendered
      return;
    }

    if (isBlocked || newLabel === 'BLOCKED') {
      statusEl.textContent = 'BLOCKED';
      subEl.textContent = 'You cannot work this shift as it will breach the 6 days in a row rule.';
      card.classList.remove('needs-attention', 'attention-border');
    } else if (newLabel === 'NOT AVAILABLE') {
      statusEl.textContent = 'NOT AVAILABLE';
      if (newLabel !== baselineLabel) {
        subEl.innerHTML = `<span class="edit-note">Pending change</span>`;
      } else {
        subEl.textContent = '';
      }
      card.classList.remove('needs-attention', 'attention-border');
    } else if (newLabel === 'LONG DAY' || newLabel === 'NIGHT' || newLabel === 'LONG DAY/NIGHT') {
      statusEl.textContent = 'AVAILABLE FOR';
      const availability =
        newLabel === 'LONG DAY' ? 'LONG DAY ONLY' :
        newLabel === 'NIGHT' ? 'NIGHT ONLY' : 'LONG DAY AND NIGHT';
      if (newLabel !== baselineLabel) {
        subEl.innerHTML = `<span class="edit-note">Pending change</span><br>${escapeHtml(availability)}`;
      } else {
        subEl.textContent = availability;
      }
      card.classList.remove('needs-attention', 'attention-border');
    } else {
      // Pending state (two wordings based on first-touch visual)
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

    // Keep border fallback for non-attention tiles
    if (!card.classList.contains('attention-border')) {
      card.style.borderColor = '#ffffff';
    }

    // Keep the label visually fitting inside its container
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

    // Initial content (same logic as before)
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
        card.classList.add('needs-attention', 'attention-border');
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

    // Tap handler → incremental update only
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

        // Cycle label, persist, and update only this card
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

        // Refresh card UI only (no grid rebuild, no equalization here)
        const effective = draft[ymd] ?? baselineLabel;
        updateCardUI(card, ymd, baselineLabel, effective, isBooked, isBlocked, canEdit);

        // Refresh tooltip/hint for the new cycle state
        const pseudoTile = { ymd, baselineLabel, effectiveLabel: effective };
        card.title = buildCycleHint(pseudoTile);

        // Footer visibility/animation can update safely
        showFooterIfNeeded();
      });
    } else {
      card.title = t.booked ? 'BOOKED (locked)' : (t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable');
      card.style.opacity = t.booked || t.status === 'BLOCKED' ? .9 : 1;
    }
  }

  // Equalize only on initial/full render, not on tap
 equalizeTileHeights({ reset: true });
  // No sizeGrid() here; avoids layout shifts when footer toggles.
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
async function loadFromServer({ force=false } = {}) {
  if (!authToken()) {
    showAuthError('Missing or invalid token');
    throw new Error('__AUTH_STOP__');
  }
  if (!identity || !identity.msisdn) {
    // Not signed in yet; show login if not already visible (but not over blocking modals)
    if (!isBlockingOverlayOpen()) openLoginOverlay();
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
        if (!isBlockingOverlayOpen()) openLoginOverlay();
        return;
      }
      throw new Error(errCode || 'SERVER_ERROR');
    }

    baseline = json;

    // Update UI name
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

    // Last loaded
    saveLastLoaded(json.lastLoadedAt || new Date().toISOString());

    // Drop draft for unknown tiles
    const validYmd = new Set((baseline.tiles || []).map(x => x.ymd));
    for (const k of Object.keys(draft)) {
      if (!validYmd.has(k)) delete draft[k];
    }

    renderTiles();
    showFooterIfNeeded();
    maybeShowWelcome(); // uses ios/android-specific messages if provided
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
        if (!isBlockingOverlayOpen()) openLoginOverlay();
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
(async function init() {
  // Identity from storage
  identity = loadSavedIdentity();

  // Build overlays/menu early
  ensureLoadingOverlay();
  ensureMenu();

  // If reset link present, show reset (blocking)
  const hasK = new URLSearchParams(location.search).has('k');
  if (hasK) {
    openResetOverlay();
  } else if (!identity.msisdn) {
    // Attempt silent auto-sign-in on Android/Chrome
    const autoOK = await tryAutoLoginViaCredentialsAPI();
    if (!autoOK && !isBlockingOverlayOpen()) openLoginOverlay();
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
      equalizeTileHeights({ reset: true });
      hideLoading();
      ensurePastShiftsButton();
    }
  }

  routeFromURL();
})();

