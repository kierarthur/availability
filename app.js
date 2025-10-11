/* ====== CONFIG — replace these with your values ====== */
const API_BASE_URL     = 'https://script.google.com/macros/s/AKfycbw9X71BbvC55s2iJhyfGU5PoBtOxC9jvFsdKpd8BvrNghMfw8Yy9X-iSZ1Xcw9oTAPMcA/exec'; // <-- REPLACE
const API_SHARED_TOKEN = 't9x_93HDa8nL0PQ6RvzX4wqZ'; // <-- REPLACE
/* ===================================================== */

const PENDING_LABEL_DEFAULT = 'PLEASE PROVIDE YOUR AVAILABILITY';
const PENDING_LABEL_WRAPPED = 'NOT SURE OF MY AVAILABILITY YET';

// 30s poll (coalesces)
const DRAFT_KEY = 'rota_avail_draft_v1';
const LAST_LOADED_KEY = 'rota_avail_last_loaded_v1';
const SAVED_IDENTITY_KEY = 'rota_avail_identity_v1';
const LAST_EMAIL_KEY = 'rota_last_login_email_v1';
const REFRESH_SNOOZE_UNTIL_KEY = 'rota_refresh_snooze_until_v1'; // ← hoisted
window.CONFIG = { TIMESHEET_FEATURE_ENABLED: true, TIMESHEET_TESTMODE: true };

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
  // If we've already shown it, bail early.
  if (window.AUTH_DENIED === true || AUTH_DENIED === true) return;

  // Sync both the global and the module-scoped flags.
  window.AUTHDENIED = false; // legacy safety if referenced elsewhere
  window.AUTH_DENIED = true;
  AUTH_DENIED = true;

  // Hide/disable interactive UI
  try { els.footer && els.footer.classList.add('hidden'); } catch {}
  try { els.grid && (els.grid.innerHTML = ''); } catch {}
  try { els.helpMsg && els.helpMsg.classList.add('hidden'); } catch {}
  try { els.submitBtn && (els.submitBtn.disabled = true); } catch {}
  try { els.clearBtn && (els.clearBtn.disabled = true); } catch {}
  try { els.refreshBtn && (els.refreshBtn.disabled = true); } catch {}
  try { els.emergencyBtn && (els.emergencyBtn.disabled = true); } catch {}

  hideLoading();

  // Remove any previous auth overlay before adding a new one
  try {
    const prev = document.getElementById('authErrorOverlay');
    if (prev && prev.parentNode) prev.parentNode.removeChild(prev);
  } catch {}

  // Non-blocking overlay: allow clicks to the Login overlay (pointer-events: none)
  const div = document.createElement('div');
  div.id = 'authErrorOverlay';
  div.setAttribute('role', 'alert');
  div.setAttribute('aria-live', 'assertive');
  Object.assign(div.style, {
    position: 'fixed',
    inset: '0',
    zIndex: '9000',                 // below modal layers (e.g., 9998/9999)
    pointerEvents: 'none',          // never block clicks to Login overlay
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'rgba(10,12,16,0.85)',
    backdropFilter: 'blur(2px)'
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

// --- UPDATED ---
function clearSavedIdentity() {
  // Clear localStorage copy…
  try { localStorage.removeItem(SAVED_IDENTITY_KEY); } catch {}
  // …and also clear the in-memory identity so subsequent requests
  // don’t keep sending a stale msisdn.
  try { identity = {}; } catch {}
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
      --brand-blue-dk: #0b254a;
      --brand-blue-dk-95: #0d2b56;
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

    /* Menu dropdown — CSS keeps it collapsed by default; JS toggles 'show' */
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

    .sheet-body label { font-weight:700; font-size:.9rem; margin-top:.2rem; }
    .sheet-body input {
      border-radius:8px; border:1px solid #222936; background:#0b0e14; color:#e7ecf3;
      padding:.6rem; width:100%;
    }

    .input-row {
      display:flex; gap:.4rem; align-items:center;
    }
    .input-row input { flex:1 1 auto; }
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

  // Login overlay (NON-dismissable + BLOCKING)
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
           <button id="forgotSubmit" type="submit" class="menu-item" style="border:1px solid #2a3446;">Send reset link</button>
          </div>
          <div id="forgotMsg" class="muted" role="status" style="margin-top:.4rem;"></div>
        </form>
      </div>
    </div>
  `;
  document.body.appendChild(forgotOverlay);

  // Reset overlay (NON-dismissable)
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
    // welcomeOverlay: non-dismissable
    // loginOverlay: non-dismissable
    { overlayId:'forgotOverlay',  closeId:'forgotClose' },
    // resetOverlay: non-dismissable
    // alertOverlay: non-dismissable
    { overlayId:'emergencyOverlay', closeId:'emergencyClose' }
  ].forEach(({overlayId, closeId}) => {
    const overlay = document.getElementById(overlayId);
    const cfg = OVERLAY_CONFIG[overlayId] || { dismissible:true, blocking:false };
    if (!overlay) return;

    // Backdrop click → only if dismissible
    overlay.addEventListener('click', (e) => {
      if (e.target !== overlay) return;
      if (!cfg.dismissible) { e.stopPropagation(); return; }
      if (overlayId === 'emergencyOverlay') {
        closeEmergencyOverlay(true);
      } else {
        closeOverlay(overlayId);
      }
    });

    // Close button → only wired if exists and dismissible
    if (closeId) {
      const btn = document.getElementById(closeId);
      if (btn) {
        btn.addEventListener('click', () => {
          if (!cfg.dismissible) return;
          if (overlayId === 'emergencyOverlay') {
            closeEmergencyOverlay(true);
          } else {
            closeOverlay(overlayId);
          }
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
      'emergencyOverlay'
    ].forEach(id => {
      const ov = document.getElementById(id);
      const cfg = OVERLAY_CONFIG[id] || { dismissible:true, blocking:false };
      if (ov && ov.classList.contains('show') && cfg.dismissible) {
        if (id === 'emergencyOverlay') {
          closeEmergencyOverlay(true);
        } else {
          closeOverlay(id);
        }
      }
    });
  });

  // Wire auth forms early so auth UI works regardless of when cache render happens
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
// Close Emergency overlay, then either Tiles→Emergency (if tiles due within 1 min) or just Emergency.
// Cadence restart is handled by startTilesThenEmergencyChain when available.
// Close Emergency overlay, then either Tiles→Emergency (if tiles due within 1 min) or just Emergency.
// Cadence restart is handled by startTilesThenEmergencyChain when available.

async function closeEmergencyOverlay(recheck = true) {
  // Close immediately
  closeOverlay('emergencyOverlay', /*force*/ true);
  emergencyState = null;

  if (!recheck) return;

  // Determine whether tiles are due within the next minute
  let tilesDueSoon = false;
  try {
    const lastAtStr = localStorage.getItem(LAST_LOADED_KEY);
    const lastAt = lastAtStr ? Date.parse(lastAtStr) : 0;
    const nextDue = lastAt ? (lastAt + 3 * 60 * 1000) : 0;
    tilesDueSoon = !!(nextDue && (nextDue - Date.now() <= 60 * 1000));
  } catch {}

  if (tilesDueSoon) {
    if (typeof startTilesThenEmergencyChain === 'function') {
      // Preferred: single-flight chain; it will advance the cadence in its finally
      await startTilesThenEmergencyChain({ force: true });
    } else {
      // Fallback: manual tiles then emergency, then advance cadence
      try { await loadFromServer({ force: true }); } catch {}
      await refreshEmergencyEligibility({ silent: true }).catch(() => {});
      try { if (typeof scheduleNextTilesRefresh === 'function') scheduleNextTilesRefresh(Date.now()); } catch {}
    }
  } else {
    // Only emergency refresh
    await refreshEmergencyEligibility({ silent: true }).catch(() => {});
  }
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

  // Replace previous click; primary button semantics: enable → on click disable + "Please wait…" until close
  const newOk = ok.cloneNode(true);
  ok.parentNode.replaceChild(newOk, ok);
  newOk.disabled = false;
  const normal = (newOk.textContent || 'OK').trim();
  newOk.textContent = normal;

  newOk.addEventListener('click', async () => {
    newOk.disabled = true;
    newOk.textContent = 'Please wait…';
    closeOverlay('alertOverlay', /*force*/true);
    if (typeof onOk === 'function') onOk();
  });

  openOverlay('alertOverlay', '#alertOk');
}


async function onRefreshClickForEmergency() {
  await refreshEmergencyEligibility({ silent: true, hideDuringRefresh: true });
}


function equalizeTileHeights({ reset = false } = {}) {
  const tiles = Array.from(els.grid.querySelectorAll('.tile'));
  if (!tiles.length) return;

  if (reset || cachedTileHeight === 0) {
    tiles.forEach(t => t.style.minHeight = '');
    let max = 0;
    tiles.forEach(t => {
      max = Math.max(max, Math.ceil(t.getBoundingClientRect().height));
    });
    cachedTileHeight = max;
  }

  tiles.forEach(t => {
    t.style.minHeight = cachedTileHeight + 'px';
  });
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
// Initial CTA sync (in case BIP never fires, e.g., iOS)
document.addEventListener('DOMContentLoaded', () => {
  // If this page was opened via a password-reset link (?k=...), go straight to the reset overlay
  const k = new URLSearchParams(location.search).get('k');
  if (k) {
    ensureLoadingOverlay();   // ensures the reset overlay DOM exists
    openResetOverlay();       // non-dismissable; validates k and drives the reset flow
    return;                   // prevent running non-essential startup UI below during reset
  }

  // Normal startup path (no reset token present)
  setTimeout(() => {
    if (els.installBtn) {
      syncInstallCTAVisibility();
    }
  }, 400);
});

// ───────────────────────── CHANGED ─────────────────────────
// Single, minimal place to refresh on visibility change.
// ── Resume logic (unchanged) + periodic visible refresh ──

document.addEventListener('visibilitychange', () => {
  if (document.hidden) {
    // Going to background
    cancelSubmitNudge();
    lastHiddenAt = Date.now();
    persistDraft();
    return;
  }

  // Coming to foreground
  const awayMs = lastHiddenAt ? (Date.now() - lastHiddenAt) : 0;
  const wasAwayLong = lastHiddenAt && (awayMs >= 10 * 60 * 1000); // >= 10 minutes
  lastHiddenAt = null;

  // Clear local edits after long away (existing behaviour)
  if (wasAwayLong && Object.keys(draft).length) {
    draft = {};
    persistDraft();
    showToast('Unsaved changes were cleared after inactivity.');
  }

  // If Emergency overlay is open → pause everything (no cadence work on resume)
  if (typeof emergencyOpen === 'function' ? emergencyOpen() : false) {
    if (Object.keys(draft).length) scheduleSubmitNudge();
    return;
  }

  // NEW: If we were away for >= 10 minutes, force-grey the Emergency button until a fresh server result arrives.
  if (wasAwayLong) {
    window._emergencyForceGreyUntilFresh = true;
    try { setEmergencyButtonBusy(true); } catch {}
  }

  // Resume behaviour (non-blocking): background tiles → emergency chain; loader is gated inside loadFromServer for cold start
  startTilesThenEmergencyChain({ force: false }).catch(() => {});

  // If drafts exist (post-clear), continue nudging
  if (Object.keys(draft).length) scheduleSubmitNudge();
});



// ---------- Helpers ----------
// Map status to semantic class; backgrounds/text tones are applied in renderTiles via tile--bg-* and tile--tone-*
function statusClass(label) {
  switch (label) {
    case 'BOOKED':           return 'status-booked';
    case 'BLOCKED':          return 'status-blocked';
    case 'NOT AVAILABLE':    return 'status-na';
    case 'LONG DAY':         return 'status-ld';
    case 'NIGHT':            return 'status-n';
    case 'LONG DAY/NIGHT':   return 'status-ldn';
    default:                 return 'status-pending';
  }
}

function nextStatus(currentLabel) {
  const idx = STATUS_ORDER.indexOf(currentLabel);
  return STATUS_ORDER[(idx + 1) % STATUS_ORDER.length];
}

// ───────────────────────── NEW (background single-flight) ─────────────────────────
async function prefetchEmergencyEligibility({ priority = 'normal', onReady } = {}) {
  // If we already have eligibility cached (even empty), still reflect UI; only skip net call if non-empty.
  if (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible)) {
    syncEmergencyButtonVisibility(emergencyEligibilityCache.eligible);
    if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
    if (emergencyEligibilityCache.eligible.length) return;
  }

  // Avoid duplicate calls
  if (prefetchEmergencyEligibility._inflight) {
    try { await prefetchEmergencyEligibility._inflight; } catch {}
    syncEmergencyButtonVisibility(emergencyEligibilityCache && emergencyEligibilityCache.eligible || []);
    if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
    return;
  }

  prefetchEmergencyEligibility._inflight = (async () => {
    await refreshEmergencyEligibility({ silent: true, hideDuringRefresh: true });
    if (typeof onReady === 'function') try { onReady(emergencyEligibilityCache); } catch {}
  })();

  try { await prefetchEmergencyEligibility._inflight; } finally { prefetchEmergencyEligibility._inflight = null; }
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

      const TS = window.Timesheets;
     const tsFeatureOn = !!(TS && TS.isTimesheetFeatureEnabled && TS.isTimesheetFeatureEnabled(baseline || window.identity || ''));

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

        // NEW: TS ✓ badge and click to open submitted-timesheet modal
        if (tsFeatureOn && it.timesheet_authorised === true) {
          // lightweight badge
          const badge = document.createElement('span');
          badge.textContent = 'TS ✓';
          Object.assign(badge.style, {
            marginLeft: '8px',
            padding: '2px 6px',
            borderRadius: '10px',
            background: '#0a7f43',
            color: '#fff',
            fontSize: '11px',
            fontWeight: '600',
          });
          const dateEl = card.querySelector('.past-date');
          if (dateEl) dateEl.appendChild(badge);

          // open the submitted-timesheet modal on click
          card.style.cursor = 'pointer';
          card.addEventListener('click', () => {
            const TS = window.Timesheets;
            if (!TS || !TS.openSubmittedTimesheetModal) return;
            // Build a minimal "tile-like" object; booking_id should be present after augment
            const tileForModal = {
              booking_id: it.booking_id,
              hospital: it.hospital,
              ward: it.ward,
              job_title: it.jobTitle || it.role,
              timesheet_authorised: true
            };
            TS.openSubmittedTimesheetModal(tileForModal, window.identity, baseline);
          });
        }

        frag.appendChild(card);
      });
      list.appendChild(frag);
    })
    .catch(err => {
      const msg = (err && err.message) ? err.message : err;
      document.getElementById('pastList').innerHTML = `<div class="muted">Could not load past shifts: ${msg}</div>`;
    });
}
async function fetchPastShifts() {
  const { res, json } = await apiGET({ view: 'past14' });
  if (!res.ok || !json || json.ok === false) {
    const msg = (json && json.error) || `HTTP ${res.status}`;
    throw new Error(msg);
  }

  const items = json.items || [];

  // NEW: augment past shifts with timesheet authorisation (TS ✓) via the broker
  try {
    const TS = window.Timesheets;
    const subject = window.baseline || window.identity || '';
    const tsFeatureOn = !!(TS && TS.isTimesheetFeatureEnabled && TS.isTimesheetFeatureEnabled(subject));
    if (tsFeatureOn && TS && TS.augmentHistoryWithTimesheetStatus) {
      await TS.augmentHistoryWithTimesheetStatus(items, baseline);
      // items now have .timesheet_authorised and .booking_id where determinable
    }
  } catch (e) {
    // Non-fatal: history still renders without TS ✓
    if (window.CONFIG?.TIMESHEET_TESTMODE) {
      const name = _candidateNameFrom(window.identity, window.baseline);
      if (window.Timesheets?.isTimesheetTestUser?.(name)) {
        console.warn('Timesheet status augmentation failed:', e);
      }
    }
  }

  return items;
}

function _normName(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function _candidateNameFrom(identity, baseline, hint) {
  // Prefer explicit hint if provided
  if (typeof hint === 'string' && hint) return hint;

  // Prefer baseline-provided name
  if (baseline && typeof baseline === 'object') {
    if (baseline.candidateName) return baseline.candidateName;
    if (baseline.candidate && (baseline.candidate.firstName || baseline.candidate.surname)) {
      return [baseline.candidate.firstName || '', baseline.candidate.surname || ''].join(' ').trim();
    }
  }

  // Fall back to identity fields if available
  if (identity && typeof identity === 'object') {
    if (identity.candidateName) return identity.candidateName;
    if (identity.firstName || identity.surname) {
      return [identity.firstName || '', identity.surname || ''].join(' ').trim();
    }
  }

  // Last resort: try globals if present
  if (window?.baseline) {
    if (window.baseline.candidateName) return window.baseline.candidateName;
    if (window.baseline.candidate && (window.baseline.candidate.firstName || window.baseline.candidate.surname)) {
      return [window.baseline.candidate.firstName || '', window.baseline.candidate.surname || ''].join(' ').trim();
    }
  }
  return '';
}
function _isAllowlistedName(name) {
  const n = _normName(name);
  return n === 'kier arthur' || n === 'arthur kier';
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
  // Make sure base styles/overlays are injected early so the menu has its CSS on first paint
  try { ensureLoadingOverlay(); } catch {}

  const header = document.querySelector('header');
  if (!header) return;
  if (document.getElementById('menuWrap')) return;

  const wrap = document.createElement('div');
  wrap.id = 'menuWrap';

  const btn = document.createElement('button');
  btn.id = 'menuBtn';
  btn.type = 'button';
  btn.textContent = 'Menu';
  // ARIA wiring
  btn.setAttribute('aria-haspopup', 'true');
  btn.setAttribute('aria-expanded', 'false');
  btn.setAttribute('aria-controls', 'menuList');

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

  // Collapsed-by-default on first paint (do NOT rely on CSS that may not be injected yet)
  list.style.display = 'none';
  list.setAttribute('hidden', 'hidden');
  list.classList.remove('show');

  wrap.append(btn, list);
  header.appendChild(wrap);

  // Toggle handler that works even before CSS is injected
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    if (isBlockingOverlayOpen()) return;
    const willShow = !list.classList.contains('show');
    list.classList.toggle('show', willShow);
    if (willShow) {
      list.removeAttribute('hidden');
      list.style.display = '';
      btn.setAttribute('aria-expanded', 'true');
    } else {
      list.setAttribute('hidden', 'hidden');
      list.style.display = 'none';
      btn.setAttribute('aria-expanded', 'false');
    }
  });

  // Global close handlers keep it collapsed and idempotent
  const closeMenu = () => {
    if (!list.classList.contains('show')) return;
    list.classList.remove('show');
    list.setAttribute('hidden', 'hidden');
    list.style.display = 'none';
    btn.setAttribute('aria-expanded', 'false');
  };
  document.addEventListener('click', closeMenu);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); });

  document.getElementById('miPast').addEventListener('click', () => { closeMenu(); openPastShifts(); });
  document.getElementById('miHosp').addEventListener('click', () => { closeMenu(); openContent('HOSPITAL','Hospital Addresses'); });
  document.getElementById('miAccom').addEventListener('click', () => { closeMenu(); openContent('ACCOMMODATION','Accommodation Contacts'); });

  document.getElementById('miTimesheet').addEventListener('click', async () => {
    closeMenu();
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
    closeMenu();
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
      showToast('Password Reset request has been sent by email if email exists');
    } catch (e) {
      showToast('Could not start password change: ' + (e.message || e));
    } finally {
      hideLoading();
    }
  });

  document.getElementById('miLogout').addEventListener('click', () => {
    closeMenu();
    if (isBlockingOverlayOpen()) return;

    clearSavedIdentity();
    draft = {}; persistDraft();
    baseline = null;
    els.grid && (els.grid.innerHTML = '');

    // Clear emergency state + hide/disable the button + close overlay
    try {
      window._emergencyEligibleShifts = [];
      window._emergencyState = null;
    } catch {}
    try {
      if (els.emergencyBtn) {
        els.emergencyBtn.hidden = true;
        els.emergencyBtn.disabled = true;
      }
    } catch {}
    try { closeOverlay('emergencyOverlay'); } catch {}

    openLoginOverlay();
  });

  // Final safety: enforce collapsed state once more
  try { ensureTopBarStatus(); } catch {}
}


// ---------- Welcome / newUserHint with platform targeting ----------
function maybeShowWelcome() {
  if (!baseline) return;

  const isIOS = isiOS();
  const isAnd = isAndroid();

  const hint =
    (isIOS && baseline.newUserHintIos) ? baseline.newUserHintIos :
    (isAnd && baseline.newUserHintAndroid) ? baseline.newUserHintAndroid :
    baseline.newUserHint;

  if (!hint) return;

  const titleEl = document.getElementById('welcomeTitle');
  const bodyEl  = document.getElementById('welcomeBody');
  const cta     = document.getElementById('welcomeDismiss');
  if (!titleEl || !bodyEl || !cta) return;

  titleEl.textContent = (hint.title || 'Welcome').trim();
  bodyEl.innerHTML = hint.html || '';

  // Primary button behaviour: enabled with normal label; on click → disable + "Please wait…" until overlay closes
  const normalLabel = (cta.textContent || 'OK').trim();
  cta.disabled = false;
  cta.textContent = normalLabel;
  cta.style.background = '#1b2331';
  cta.style.fontWeight = '800';
  cta.style.border = '1px solid #2a3446';
  cta.style.padding = '.6rem 1rem';
  cta.style.borderRadius = '10px';

  // Replace handler
  const newCta = cta.cloneNode(true);
  cta.parentNode.replaceChild(newCta, cta);
  newCta.addEventListener('click', async () => {
    newCta.disabled = true;
    newCta.textContent = 'Please wait…';
    try { await apiPOST({ action:'MARK_MESSAGE_SEEN' }); } catch {}
    closeOverlay('welcomeOverlay', /* force */ true);
    // On next open, label is restored above
  });

  openOverlay('welcomeOverlay', '#welcomeDismiss');
}

// ---------- Auth overlays logic ----------
function openLoginOverlay() {
  // Don’t open on top of blocking flows (reset/alert/login already open)
  if (isBlockingOverlayOpen && isBlockingOverlayOpen()) return;

  // If a persisted identity exists (even if in-memory hasn’t hydrated yet), avoid popping Login
  // UNLESS we have just shown an auth error (AUTH_DENIED true) – in that case, force login.
  try {
    const saved = (!identity || !identity.msisdn) ? loadSavedIdentity() : null;
    if (saved && saved.msisdn && !AUTH_DENIED) {
      identity = identity && identity.msisdn ? identity : saved;
      try {
        if (/^#\/login/.test(location.hash)) {
          history.replaceState(null, '', location.pathname + location.search);
        }
      } catch {}
      return;
    }
  } catch {}

  if (identity && identity.msisdn && !AUTH_DENIED) {
    try {
      if (/^#\/login/.test(location.hash)) {
        history.replaceState(null, '', location.pathname + location.search);
      }
    } catch {}
    return;
  }

  // Normal login open path (pre-fill remembered email)
  const email = getRememberedEmail();
  const le = document.getElementById('loginEmail');
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
    const email = (le?.value || '').trim().toLowerCase();
    const pw    = lp?.value || '';
    if (!email || !pw) return;

    try {
      if (ls) ls.disabled = true;
      showLoading('Signing in...');
      const { res, json } = await apiAuthLogin(email, pw);
      if (!res.ok || !json || json.ok === false) {
        const msg = (json && json.error) || `HTTP ${res.status}`;
        if (lerr) {
          lerr.textContent = (msg === 'INVALID_CREDENTIALS')
            ? 'Email or password is incorrect.'
            : (msg || 'Sign-in failed.');
        }
        return;
      }

      // Persist identity (stateless pattern) & remembered email; store creds (Chrome/Android)
      identity = { msisdn: json.msisdn }; saveIdentity(identity);
      rememberEmailLocal(email);
      await storeCredentialIfSupported(email, pw);
      try { window.AUTH_DENIED = false; AUTH_DENIED = false; } catch {}
      // Clear password field for safety
      if (lp) lp.value = '';

      // Force-close the blocking Login overlay so it cannot sit over a signed-in UI
      closeOverlay('loginOverlay', /* force */ true);

      // Clear any #/login hash so router can’t reopen it on next route/change
      try {
        if (/^#\/(login|forgot|reset)/.test(location.hash || '')) {
          history.replaceState(null, '', location.pathname + (location.search || ''));
        }
      } catch {}

      // Hard refresh baseline as an authenticated user (coalesced inside loadFromServer)
      await loadFromServer({ force: true });

      // (Optional) refresh emergency eligibility silently
      try { refreshEmergencyEligibility({ silent: true }); } catch {}
    } finally {
      hideLoading();
      if (ls) ls.disabled = false;
    }
  });

  // “Forgot password” from login
  lfg?.addEventListener('click', () => {
    closeOverlay('loginOverlay', /* force */ true);
    openForgotOverlay();
  });

  // Forgot password form submission (prevents page navigation and calls API)
  const ff   = document.getElementById('forgotForm');
  const fe   = document.getElementById('forgotEmail');
  const fs   = document.getElementById('forgotSubmit');
  const fmsg = document.getElementById('forgotMsg');

  ff?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = (fe?.value || '').trim().toLowerCase();
    if (!email) {
      if (fmsg) fmsg.textContent = 'Please enter your email address.';
      return;
    }

    try {
      if (fs) fs.disabled = true;
      showLoading('Sending reset link...');
      await apiForgotPassword(email); // privacy-safe: don’t branch on existence
      rememberEmailLocal(email);

      // Show confirmation and KEEP the Forgot overlay open
      if (fmsg) fmsg.textContent = 'If this email exists, we’ve sent a reset link.';

      // Suppress any auth-error UI for the next 10s so the confirmation remains visible
      try { window.AUTH_DENIED = true; AUTH_DENIED = true; } catch {}
      try {
        // Remove any lingering auth overlay if it flashed up already
        const prev = document.getElementById('authErrorOverlay');
        if (prev && prev.parentNode) prev.parentNode.removeChild(prev);
      } catch {}

      // After 10 seconds, restore auth handling, close Forgot, and return to Login
      setTimeout(() => {
        try { window.AUTH_DENIED = false; AUTH_DENIED = false; } catch {}
        closeOverlay('forgotOverlay');
        openLoginOverlay();
      }, 10_000);
    } catch {
      if (fmsg) fmsg.textContent = 'Could not send reset link. Please try again.';
    } finally {
      hideLoading();
      if (fs) fs.disabled = false;
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
  function same(p, c) { return p === c && p.length > 0; }

  function validateReset() {
    const k = new URLSearchParams(location.search).get('k') || '';
    const p = rp ? rp.value : '';
    const c = rc ? rc.value : '';
    let ok = true, msg = '';

    if (!k) {
      ok = false; msg = 'This reset link is invalid or missing.';
    } else if (!meetsPolicy(p)) {
      ok = false; msg = 'Use at least 8 chars with uppercase, lowercase, and a number.';
    } else if (!same(p, c)) {
      ok = false; msg = "Passwords don't match.";
    }
    if (rerr) rerr.textContent = msg;
    if (rs) rs.disabled = !ok;
    return ok;
  }

  rp?.addEventListener('input', validateReset);
  rc?.addEventListener('input', validateReset);

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

  rf?.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!validateReset()) return;

    const pw = rp.value || '';
    const k = new URLSearchParams(location.search).get('k') || '';
    if (!k) {
      // Defensive: already handled above, but keep UX smooth
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

      // Strip ?k=… from URL (one-time link)
      const clean = location.pathname + (location.hash || '');
      history.replaceState(null, '', clean);

      // Close reset, prompt sign-in
      closeOverlay('resetOverlay', /* force */ true);
      showToast('Password updated. Please sign in.');
      openLoginOverlay();

      // Clear fields
      rp.value = '';
      rc.value = '';
      rs.disabled = true;
    } finally {
      hideLoading(); rs && (rs.disabled = false);
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

    // If identity already hydrated, nothing to do.
    if (identity && identity.msisdn) return true;

    const cred = await navigator.credentials.get({ password: true, mediation: 'optional' });
    if (cred && cred.id && cred.password) {
      const { res, json } = await apiAuthLogin(cred.id, cred.password);
      if (res.ok && json && json.ok && json.msisdn) {
        identity = { msisdn: json.msisdn }; saveIdentity(identity);
        rememberEmailLocal(cred.id);

        // Remove any stale auth hash and close overlay if it was open
        try {
          if (/^#\/(login|forgot|reset)/.test(location.hash || '')) {
            history.replaceState(null, '', location.pathname + (location.search || ''));
          }
        } catch {}
        try { closeOverlay('loginOverlay', /*force*/ true); } catch {}

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

  // Feature flag (safe guards if Timesheets not present)
  // CHANGED: pass baseline/identity (name-based gating inside Timesheets)
  const TS = window.Timesheets;
  const tsFeatureOn = !!(
    TS &&
    TS.isTimesheetFeatureEnabled &&
    TS.isTimesheetFeatureEnabled(baseline || window.identity || '')
  );

  // NEW: visually confirm name-based TESTMODE match by turning the name red
  try {
    const testMode = !!(window.CONFIG && window.CONFIG.TIMESHEET_TESTMODE);
    const allowed = !!(TS && typeof TS.isTestModeAllowed === 'function' && TS.isTestModeAllowed(window.identity));
    if (els.candidateName) {
      if (testMode && tsFeatureOn && allowed) {
        // iOS-style system red; tweak as desired or swap to a CSS class.
        els.candidateName.style.color = '#ff3b30';
        els.candidateName.title = 'Test mode: name matched';
      } else {
        // reset if previously coloured
        els.candidateName.style.color = '';
        if (els.candidateName.title === 'Test mode: name matched') els.candidateName.title = '';
      }
    }
  } catch {}

  ensureHelpMessageVisible();
  ensurePastShiftsButton();

  const tiles = (baseline.tiles || []).map(t => {
    const baselineLabel = t.booked
      ? 'BOOKED'
      : (t.status === 'BLOCKED' ? 'BLOCKED' : codeToLabel(t.status));
    const effectiveLabel = draft[t.ymd] ?? baselineLabel;
    return { ...t, baselineLabel, effectiveLabel };
  });

  // Helper: set background + text tone per new colour rules
  function applyTileTheme(card, label, isBooked, isBlocked) {
    // Reset
    card.classList.remove(
      'tile--bg-booked','tile--bg-blocked','tile--bg-notavail',
      'tile--bg-available','tile--bg-pending','tile--tone-dark','tile--tone-light'
    );

    let bgClass = 'tile--bg-pending';
    let tone = 'tile--tone-light'; // default for black

    if (isBooked) {
      bgClass = 'tile--bg-booked';     // vivid blue
      tone   = 'tile--tone-light';     // white text for contrast
    } else if (isBlocked || label === 'BLOCKED') {
      bgClass = 'tile--bg-blocked';    // light red
      tone   = 'tile--tone-dark';      // dark text
    } else if (label === 'NOT AVAILABLE') {
      bgClass = 'tile--bg-notavail';   // bright red
      tone   = 'tile--tone-light';     // white text
    } else if (
      label === 'LONG DAY' || label === 'NIGHT' || label === 'LONG DAY/NIGHT' ||
      label === 'AVAILABLE' || label === 'AVAILABLE FOR'
    ) {
      bgClass = 'tile--bg-available';  // availability
      tone   = 'tile--tone-dark';      // dark text
    } else {
      bgClass = 'tile--bg-pending';    // black
      tone   = 'tile--tone-light';     // white text
    }

    card.classList.add(bgClass, tone);
  }

  function updateCardUI(card, ymd, baselineLabel, newLabel, isBooked, isBlocked, editable) {
    const statusEl = card.querySelector('.tile-status');
    const subEl    = card.querySelector('.tile-sub');
    statusEl.className = 'tile-status ' + statusClass(newLabel);

    // Apply background + text tone every time state changes
    applyTileTheme(card, newLabel, isBooked, isBlocked);

    if (isBooked) {
      statusEl.textContent = 'BOOKED';
      subEl && (subEl.innerHTML = subEl.innerHTML || '');
      card.classList.remove('needs-attention', 'attention-border');
      fitStatusLabel(statusEl);
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

    // Seed content as before
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

    // Apply the new theme classes (bg + tone) based on state
    applyTileTheme(card, t.effectiveLabel, !!t.booked, t.status === 'BLOCKED');

    card.append(header, status, sub);
    els.grid.appendChild(card);
    fitStatusLabel(status);

    if (!card.classList.contains('attention-border')) {
      card.style.borderColor = '#ffffff';
    }

    // NEW: booked-tile click routing for Timesheets (wizard or submitted modal)
    if (t.booked && tsFeatureOn) {
      card.style.cursor = 'pointer';
      card.title = 'Tap for timesheet';
      card.addEventListener('click', (ev) => {
        // Let the Timesheets module decide: submitted modal vs wizard vs no-op
        try {
          const handled = TS.onTileClickTimesheetBranch && TS.onTileClickTimesheetBranch(t, window.identity, baseline);
          if (handled) {
            ev.stopPropagation();
            ev.preventDefault();
          }
        } catch (e) {
          // fail-safe: ignore and fall back to no-op
        }
      });

      // (A) draw TS ✓ badge on booked tiles if authorised
      try { TS.decorateTileWithTSBadge && TS.decorateTileWithTSBadge(card, t); } catch {}

      // (B) attention state for eligible, not-yet-authorised tiles
      try {
        const showCTA =
          TS.shouldShowTimesheetCTA &&
          TS.shouldShowTimesheetCTA(t, window.identity);

        if (showCTA) {
          TS.decorateTileWithTSAttention && TS.decorateTileWithTSAttention(card, t);
          TS.startTileContentAlternator && TS.startTileContentAlternator(card, {
            originalHtml: sub.innerHTML,
            altText: 'Please touch to sign your timesheet',
            periodMs: 4000
          });
        } else {
          TS.stopTileContentAlternator && TS.stopTileContentAlternator(card);
          // Ensure attention visuals are cleared if previously applied
          card.classList.remove('tile--bg-ts-cta');
          const attn = card.querySelector('.ts-badge-attn');
          if (attn) attn.remove();
        }
      } catch {}
    }

    // Keep existing edit-cycle click for editable NON-booked tiles
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
    } else if (!t.booked) {
      card.title = t.status === 'BLOCKED' ? 'BLOCKED by 6-day rule' : 'Not editable';
      card.style.opacity = t.status === 'BLOCKED' ? .9 : 1;
    } else {
      // booked + feature off: keep current “locked” tooltip
      if (!tsFeatureOn) {
        card.title = 'BOOKED (locked)';
      }
    }
  }

  // Equalize only on initial/full render, not on tap
  equalizeTileHeights({ reset: true });

  // Button state comes from cache and the chained eligibility refresh.
  updateEmergencyButtonFromCache();
}


let cachedTileHeight = 0; // global cache

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

// --- UPDATED ---

async function loadFromServer({ force = false } = {}) {
  // Single-flight guard (coalesce bursts)
  if (loadFromServer._inflight && !force) return loadFromServer._inflight;

  // Make sure the top bar is in a safe initial state before any immediate render.
  try { ensureTopBarStatus(); } catch {}

  // Ensure in-memory identity is hydrated from storage BEFORE any network work
  try {
    if (!identity || !identity.msisdn) {
      const saved = loadSavedIdentity();
      if (saved && saved.msisdn) identity = saved;
    }
  } catch {}

  // Hydrate baseline from local storage (per-user) so tiles render immediately on cold/hard start
  try {
    if (!baseline || !Array.isArray(baseline.tiles) || baseline.tiles.length === 0) {
      const cacheKey = (identity && identity.msisdn) ? `BASELINE_CACHE_${identity.msisdn}` : null;
      if (cacheKey) {
        const raw = localStorage.getItem(cacheKey);
        if (raw) {
          const cached = JSON.parse(raw);
          if (cached && Array.isArray(cached.tiles) && cached.tiles.length > 0) {
            baseline = cached;
            try { renderTiles(); } catch {}
            try { showFooterIfNeeded(); } catch {}
            // IMPORTANT: if we're "force-grey-until-fresh", do NOT repaint from cache here.
            try { if (!window._emergencyForceGreyUntilFresh) updateEmergencyButtonFromCache(); } catch {}
          }
        }
      }
    }
  } catch {}

  // Do we already have tiles on screen? (controls loader UX only)
  const hadBaseline = !!(baseline && Array.isArray(baseline.tiles) && baseline.tiles.length > 0);

  // Token missing → definitive unauth → show login (regardless of tiles)
  if (!authToken()) {
    try { els.emergencyBtn && (els.emergencyBtn.disabled = true, els.emergencyBtn.hidden = false); } catch {}
    showAuthError('Missing or invalid token');
    if (!isBlockingOverlayOpen()) openLoginOverlay();
    throw new Error('__AUTH_STOP__');
  }

  // Keep Emergency button visible but disabled while (re)establishing state
  try {
    if (els.emergencyBtn) {
      els.emergencyBtn.hidden = false;
      els.emergencyBtn.disabled = true;
    }
  } catch {}

  // Loader only for cold-start (no tiles yet). Background otherwise.
  const shouldShowLoader = !hadBaseline;
  let showedLoader = false;
  if (shouldShowLoader) {
    showLoading('Updating…');
    showedLoader = true;
  }

  // Always show a subtle top-bar updating indicator during the fetch (non-blocking)
  try { setTopBarUpdating(true); } catch {}

  const work = (async () => {
    try {
      // ── fetch BASELINE ONLY ──
      const doFetch = async (withMsisdn = false) => {
        const hasId = !!(identity && identity.msisdn);
        return withMsisdn && hasId ? await apiGET({ msisdn: identity.msisdn }) : await apiGET({});
      };

      let { res, json } = await doFetch(false);

      // One-time recovery for the race: if server says identity is unknown AND we do have a saved identity, rehydrate and retry once.
      if (
        (res.status === 400 && (!json || json.error === 'INVALID_OR_UNKNOWN_IDENTITY')) ||
        (json && json.ok === false && json.error === 'INVALID_OR_UNKNOWN_IDENTITY')
      ) {
        const saved = loadSavedIdentity();
        if (saved && saved.msisdn) {
          identity = saved; // rehydrate
          ({ res, json } = await doFetch(true)); // retry once with explicit msisdn
        }
      }

      // Transient busy → toast, no login
      if (res.status === 503 || (json && json.error === 'TEMPORARILY_BUSY_TRY_AGAIN')) {
        showToast('Server is busy, please try again.');
        return;
      }

      const errCode = (json && json.error) || '';

      // Definitive auth failures → clear identity + open Login (even if tiles exist)
      if (!res.ok) {
        const isUnauth400 = (res.status === 400) && (errCode === 'INVALID_OR_UNKNOWN_IDENTITY' || errCode === 'NOT_IN_CANDIDATE_LIST');
        if (res.status === 401 || res.status === 403 || isUnauth400) {
          try { els.emergencyBtn && (els.emergencyBtn.disabled = true, els.emergencyBtn.hidden = false); } catch {}
          clearSavedIdentity();
          showAuthError('Not an authorised user');
          if (!isBlockingOverlayOpen()) openLoginOverlay();
          throw new Error('__AUTH_STOP__');
        }
        // Any other non-OK (e.g., 400 for non-auth reasons) → throw but do NOT open login here
        throw new Error(errCode || `HTTP ${res.status}`);
      }

      if (json && json.ok === false) {
        const unauthErrors = new Set(['FORBIDDEN', 'INVALID_OR_UNKNOWN_IDENTITY', 'NOT_IN_CANDIDATE_LIST']);
        if (unauthErrors.has(errCode)) {
          try { els.emergencyBtn && (els.emergencyBtn.disabled = true, els.emergencyBtn.hidden = false); } catch {}
          clearSavedIdentity();
          showAuthError('Not an authorised user');
          if (!isBlockingOverlayOpen()) openLoginOverlay();
          throw new Error('__AUTH_STOP__');
        }
        throw new Error(errCode || 'SERVER_ERROR');
      }

      // ------- baseline UI/data -------
      const prevBaseline = baseline; // for draft merge
      let nextBaseline = json;

      // Short merge/guard: preserve very recent optimistic changes
      try {
        const ra = (typeof window !== 'undefined') ? window.__recentlyApplied : null;
        const now = Date.now();
        if (ra && ra.until && now < ra.until && ra.codes && nextBaseline && Array.isArray(nextBaseline.tiles)) {
          const codes = ra.codes;
          const byYmdNew = new Map((nextBaseline.tiles || []).map(t => [t.ymd, t]));
          for (const [ymd, code] of Object.entries(codes)) {
            const t = byYmdNew.get(ymd);
            if (!t) continue;

            const isBooked  = !!t.booked;
            const isBlocked = String(t.status || '').toUpperCase() === 'BLOCKED';
            const editable  = t.editable !== false;

            // Honor server supremacy for BOOKED/BLOCKED; otherwise preserve optimistic code
            if (!isBooked && !isBlocked && editable) {
              t.status = code;
            }
          }
        }
        // Clear guard when expired
        if (!ra || now >= ra.until) { try { if (window && window.__recentlyApplied) delete window.__recentlyApplied; } catch {} }
      } catch {}

      baseline = nextBaseline;

      // Persist fresh baseline to per-user cache for instant next-boot render
      try {
        if (identity && identity.msisdn) {
          localStorage.setItem(`BASELINE_CACHE_${identity.msisdn}`, JSON.stringify(baseline));
        }
      } catch {}

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

      // Cull drafts ONLY where server tile changed (unchanged-tile merge rule)
      try {
        if (prevBaseline && prevBaseline.tiles && baseline && baseline.tiles) {
          const oldMap = new Map((prevBaseline.tiles || []).map(t => [t.ymd, JSON.stringify({
            ymd: t.ymd,
            booked: !!t.booked,
            blocked: t.status === 'BLOCKED',
            status: t.status,
            shiftInfo: t.shiftInfo || '',
            hospital: t.hospital || t.location || '',
            ward: t.ward || '',
            jobTitle: t.jobTitle || '',
            bookingRef: t.bookingRef || '',
            display: t.display_label || '',
            time: t.time_range_label || '',
            type: t.shift_type_label || '',
            editable: t.editable !== false
          })]));
          const newMap = new Map((baseline.tiles || []).map(t => [t.ymd, JSON.stringify({
            ymd: t.ymd,
            booked: !!t.booked,
            blocked: t.status === 'BLOCKED',
            status: t.status,
            shiftInfo: t.shiftInfo || '',
            hospital: t.hospital || t.location || '',
            ward: t.ward || '',
            jobTitle: t.jobTitle || '',
            bookingRef: t.bookingRef || '',
            display: t.display_label || '',
            time: t.time_range_label || '',
            type: t.shift_type_label || '',
            editable: t.editable !== false
          })]));

          for (const ymd of Object.keys(draft)) {
            const before = oldMap.get(ymd);
            const after  = newMap.get(ymd);
            if (!before || !after || before !== after) {
              delete draft[ymd];
            }
          }
          persistDraft();
        } else {
          const validYmd = new Set((baseline.tiles || []).map(x => x.ymd));
          for (const k of Object.keys(draft)) {
            if (!validYmd.has(k)) delete draft[k];
          }
        }
      } catch {
        const validYmd = new Set((baseline.tiles || []).map(x => x.ymd));
        for (const k of Object.keys(draft)) {
          if (!validYmd.has(k)) delete draft[k];
        }
      }

      renderTiles();
      showFooterIfNeeded();
      maybeShowWelcome();

      // Keep button visible; eligibility refresh will decide final state.
      // IMPORTANT: don't repaint from cache if we're in a "force-grey-until-fresh" phase.
      if (!window._emergencyForceGreyUntilFresh) {
        updateEmergencyButtonFromCache();
      }

      // We are definitely signed in at this point → ensure no login sheet is lingering
      // We are definitely signed in at this point → ensure no auth/login overlays are lingering
try { closeOverlay && closeOverlay('loginOverlay', /* force */ true); } catch {}
try {
  if (identity && identity.msisdn && /^#\/(login|forgot|reset)/.test(location.hash || '')) {
    history.replaceState(null, '', location.pathname + (location.search || ''));
  }
} catch {}

// Clear any previous "auth denied" dimmer and reset the flag
try {
  const dim = document.getElementById('authErrorOverlay');
  if (dim && dim.parentNode) dim.parentNode.removeChild(dim);
} catch {}
try { window.AUTH_DENIED = false; AUTH_DENIED = false; } catch {}
try { window.AUTH_DENIED = false; } catch {}
    } finally {
      if (showedLoader) hideLoading();
      try { setTopBarUpdating(false); } catch {}
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

  const changeMap = new Map(changes.map(c => [c.ymd, c.code]));

  showLoading('Saving…');

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

    // Optimistically update local baseline
    if (baseline && Array.isArray(baseline.tiles) && applied.length) {
      const byYmd = new Map(baseline.tiles.map(t => [t.ymd, t]));
      const recentCodes = {};

      for (const r of applied) {
        const tile = byYmd.get(r.ymd);
        if (!tile) continue;
        const newCode = (r.code !== undefined && r.code !== null) ? r.code : changeMap.get(r.ymd);
        if (newCode !== undefined) {
          tile.status = newCode;
          recentCodes[r.ymd] = newCode; // remember for short merge/guard
        }
      }

      saveLastLoaded(new Date().toISOString());

      // Persist the updated baseline immediately so UI doesn't revert on next paint
      try {
        if (identity && identity.msisdn) {
          localStorage.setItem(`BASELINE_CACHE_${identity.msisdn}`, JSON.stringify(baseline));
        }
      } catch {}

      // Start a short "recently applied" window so an immediate fetch won't clobber optimistic values
      try {
        if (Object.keys(recentCodes).length) {
          window.__recentlyApplied = {
            until: Date.now() + 20_000, // ~20s guard
            codes: recentCodes
          };
        }
      } catch {}
    }

    if (rejected.length) {
      const reasons = rejected.slice(0,3).map(r => `${r.ymd}: ${r.reason}`).join(', ');
      showToast(`Saved ${applied.length}. Skipped ${rejected.length} (${reasons}${rejected.length>3?', …':''}).`, 4200);
    } else {
      showToast(`Saved ${applied.length} change${applied.length===1?'':'s'}.`);
    }

    draft = {};
    persistDraft();
    cancelSubmitNudge();
    nudgeShown = false;
    lastEditAt = 0;
    toggledThisSession.clear();

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
  // Apply base styles (red theme); actual state (grey/red & enabled/disabled) is controlled elsewhere.
  function applyStyles(btn) {
    if (!btn) return;
    btn.className = 'btn btn-danger';
    btn.style.background = '#d32f2f';
    btn.style.color = '#fff';
    btn.style.border = '1px solid #b71c1c';

    // keep it visually and interactively separate from Refresh
    btn.style.marginLeft = '.5rem';
    btn.style.position = 'relative';
    btn.style.zIndex = '3';
    btn.style.pointerEvents = 'auto';
  }

  function finalize(btn) {
    els.emergencyBtn = btn;

    // Wire first (wireEmergencyButton clones the node)
    try { wireEmergencyButton(); } catch {}

    // Rebind to the live (possibly cloned) node, then style + baseline state
    const live = document.getElementById('emergencyBtn') || btn;
    applyStyles(live);
    els.emergencyBtn = live;

    // Default state on app start/resume: visible but grey & disabled until a fresh server check completes.
    live.hidden = false;
    live.disabled = true;
    live.style.filter = 'grayscale(1)';
    live.style.opacity = '.85';

    // IMPORTANT: Do NOT repaint from cache here.
    // We want the button to remain grey until the next server result arrives.
  }

  const existing = document.getElementById('emergencyBtn');
  if (existing) { finalize(existing); return; }

  const tryFindAnchor = () => {
    if (els.refreshBtn && els.refreshBtn instanceof HTMLElement) return els.refreshBtn;
    return (
      document.querySelector('#header .actions, .header .actions, .topbar .actions, .toolbar, #menuRight') ||
      document.getElementById('header') ||
      document.querySelector('#app, body')
    );
  };

  const createBtn = (anchor) => {
    if (!anchor) return false;

    const btn = document.createElement('button');
    btn.id = 'emergencyBtn';
    btn.type = 'button';
    btn.textContent = 'Emergency';
    btn.hidden = false;     // always visible
    btn.disabled = true;    // greyed until eligibility loads
    btn.style.filter = 'grayscale(1)';
    btn.style.opacity = '.85';

    if (anchor === els.refreshBtn && anchor.parentElement) {
      anchor.parentElement.insertBefore(btn, anchor.nextSibling);
    } else {
      anchor.appendChild(btn);
    }

    finalize(btn);
    return true;
  };

  if (createBtn(tryFindAnchor())) return;

  let attempts = 0;
  const maxAttempts = 20; // ~5s at 250ms
  const timer = setInterval(() => {
    attempts += 1;
    if (createBtn(tryFindAnchor()) || attempts >= maxAttempts) clearInterval(timer);
  }, 250);
}



/** Show/hide the EMERGENCY button based on eligible shifts. */
function syncEmergencyButtonVisibility(eligible) {
  // Revised: never hide; reflect GREYED vs BRIGHT RED based on eligibility array
  try {
    const btn = document.getElementById('emergencyBtn');
    if (!btn) return;

    // Cache if provided so the rest of the app stays consistent
    if (Array.isArray(eligible)) {
      emergencyEligibilityCache = { eligible };
    }

    btn.hidden = false; // always visible

    const hasEligible = Array.isArray(eligible) && eligible.length > 0;
    if (btn.dataset.busy === '1') {
      // Busy → greyed & disabled
      btn.disabled = true;
      btn.style.filter = 'grayscale(1)';
      btn.style.opacity = '.85';
      return;
    }

    if (hasEligible) {
      // BRIGHT RED & enabled
      btn.disabled = false;
      btn.style.filter = '';
      btn.style.opacity = '';
      btn.style.background = '#d32f2f';
      btn.style.border = '1px solid #b71c1c';
      btn.style.color = '#fff';
    } else {
      // Greyed & disabled
      btn.disabled = true;
      btn.style.filter = 'grayscale(1)';
      btn.style.opacity = '.85';
    }
  } catch {}
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




/* =========================
 * NEW OVERLAY FACTORY
 * ========================= */

/** Lazily create the Emergency overlay DOM (uses existing overlay system). */
function ensureEmergencyOverlay() {
  if (document.getElementById('emergencyOverlay')) return;

  // Make sure base overlay utilities exist (focus trap, loading, etc.)
  ensureLoadingOverlay();

  // ─────────────────────────
  // Inject Emergency-specific CSS once
  // ─────────────────────────
  if (!document.getElementById('emergencyOverlayStyles')) {
    const style = document.createElement('style');
    style.id = 'emergencyOverlayStyles';
    style.textContent = `
      /* Constrain the sheet so content never runs off small screens (e.g., iPhone) */
      #emergencyOverlay .sheet {
        width: min(720px, 96vw);
        margin-top: min(8vh, 80px);
        margin-bottom: 0;
        max-height: calc(100vh - 12vh);
        display: flex;
        flex-direction: column;
      }

      /* Scroll body if content is long; no horizontal scroll */
      #emergencyBody {
        overflow-y: auto;
        overflow-x: hidden;
      }

      /* A consistent, compact row layout for ALL option lists (shift/issue/late/DNA) */
      #emergencyBody .em-list {
        display: flex;
        flex-direction: column;
        gap: .5rem;
        max-width: 100%;
      }

      /* Force labels (rows) to a simple flex row — overrides any inline 'grid' used elsewhere */
      #emergencyBody label,
      #emergencyBody label[role],
      #emergencyBody .em-row {
        display: flex !important;
        align-items: center;
        gap: .6rem;
        padding: .55rem .7rem;
        border: 1px solid #222a36;
        background: #131926;
        border-radius: 10px;
        max-width: 100%;
        box-sizing: border-box;
      }

      /* Keep the radio a sensible, consistent size and aligned */
      #emergencyBody input[type="radio"] {
        flex: 0 0 auto;
        width: 18px;
        height: 18px;
        accent-color: #4f8cff; /* supported on modern mobile browsers */
        margin: 0; /* remove default iOS margin that can push content */
      }

      /* Checkbox (DNA tried-calling) tidy sizing too */
      #emergencyBody input[type="checkbox"] {
        width: 18px;
        height: 18px;
        accent-color: #4f8cff;
        margin: 0;
      }

      /* Text container next to the radio: fill remaining width, wrap nicely */
      #emergencyBody label > div,
      #emergencyBody .em-row > .em-text {
        flex: 1 1 auto;
        min-width: 0;                 /* prevents narrow columns that cause 2-3 words per line */
        white-space: normal;          /* allow natural wrapping */
        overflow-wrap: anywhere;      /* break long tokens/refs gracefully */
        word-break: normal;
        line-height: 1.35;
      }

      /* Headline text inside a row (e.g., shift label) */
      #emergencyBody .em-title {
        font-weight: 800;
        margin: 0;
      }

      /* Subtext/help under a title inside a row */
      #emergencyBody .em-help,
      #emergencyBody .muted {
        color: #a7b0c0;
      }

      /* Running-late option rows should be compact, not "massive" */
      #emergencyBody .em-late-row {
        padding: .5rem .7rem;
        font-size: clamp(.95rem, 2.5vw, 1rem);
      }

      /* DNA instruction block and other inline notes keep to the viewport */
      #emergencyBody .em-note {
        margin-top: .6rem;
        color: #a7b0c0;
        max-width: 100%;
        overflow-wrap: anywhere;
      }

      /* Footer buttons stay visible and tidy on small screens */
      #emergencyFooter {
        display: flex;
        justify-content: flex-end;
        gap: .5rem;
        padding: .5rem .9rem 1rem;
        flex: 0 0 auto;
      }

      /* Prevent any accidental horizontal scrolling on the overlay container */
      #emergencyOverlay {
        overflow: hidden;
      }
    `;
    document.head.appendChild(style);
  }

  // ─────────────────────────
  // Build overlay DOM
  // ─────────────────────────
  const overlay = document.createElement('div');
  overlay.id = 'emergencyOverlay';
  overlay.innerHTML = `
    <div class="sheet" role="dialog" aria-modal="true" aria-label="Emergency">
      <div class="sheet-header">
        <div id="emergencyTitle" class="sheet-title">Emergency</div>
        <button id="emergencyClose" class="sheet-close" aria-label="Close">✕</button>
      </div>
      <div id="emergencyBody" class="sheet-body"></div>
      <div id="emergencyFooter"></div>
    </div>
  `;
  document.body.appendChild(overlay);

  // Close button — emergency overlay is dismissible
  const closeBtn = overlay.querySelector('#emergencyClose');
  if (closeBtn) closeBtn.addEventListener('click', () => closeEmergencyOverlay(true)); // <-- recheck on close

  // Backdrop click to close (consistent with dismissible overlays)
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closeEmergencyOverlay(true); // <-- recheck on close
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

function wireEmergencyButton() {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;

  const clone = btn.cloneNode(true);
  btn.replaceWith(clone);

  els.emergencyBtn = document.getElementById('emergencyBtn');

  // Ensure it’s fully clickable (no leftover busy/disabled state)
  setEmergencyButtonBusy(false);

  els.emergencyBtn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    openEmergencyOverlay();
  });
}

// =========================
// SERVER-AUTHORITATIVE FLOW
// =========================
async function refreshEmergencyEligibility({ silent = false /*, hideDuringRefresh ignored */ } = {}) {
  // Pause while the Emergency overlay is open (it's dismissible, so don't use isBlockingOverlayOpen)
  if (typeof emergencyOpen === 'function' && emergencyOpen()) return;

  // Single-flight + cooldown
  if (window.emergencyInFlight) return;
  const now = Date.now();
  if (window.emergencyCooldownUntil && now < window.emergencyCooldownUntil) return;

  window.emergencyInFlight = true;
  window.emergencyCooldownUntil = now + 8000; // ~8s

  // Decide whether to grey the button during this fetch:
  // - If we're in a "force-grey-until-fresh" phase (e.g., after submit or long-away resume), grey it.
  // - If we have NOT previously been eligible (i.e., not red), grey it.
  // - Otherwise (currently red/eligible), DO NOT grey during routine checks to avoid flicker.
  const forceGrey = !!window._emergencyForceGreyUntilFresh;
  const hadEligible = Array.isArray(window._emergencyEligibleShifts) && window._emergencyEligibleShifts.length > 0;
  const weSetBusy = (forceGrey || !hadEligible);

  if (weSetBusy) {
    try { setEmergencyButtonBusy(true); } catch {}
  }

  let ok = false, eligible = [], error = null;

  try {
    const res = await apiGetEmergencyWindow();
    ok = !!res.ok;
    eligible = Array.isArray(res.eligible) ? res.eligible : [];
    error = res.error || null;
  } catch (e) {
    ok = false;
    error = e && e.message;
  }

  if (!ok) {
    emergencyEligibilityCache = { eligible: [] };
    window._emergencyEligibleShifts = [];
    try { syncEmergencyButtonVisibility([]); } catch {}
    if (!silent) showToast(mapServerErrorToMessage(error));
  } else {
    emergencyEligibilityCache = { eligible };
    window._emergencyEligibleShifts = eligible;
    try { syncEmergencyButtonVisibility(eligible); } catch {}
  }

  // Fresh server answer → clear any force-grey phase.
  window._emergencyForceGreyUntilFresh = false;

  // Only un-busy if we set it busy here (avoids extra repaint from cache when we never greyed)
  if (weSetBusy) {
    try { setEmergencyButtonBusy(false); } catch {}
  }

  window.emergencyInFlight = false;
}


// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────

async function openEmergencyOverlay() {
  if (isBlockingOverlayOpen()) return;

  // Use cached eligibility; do not trigger background refresh while opening.
  const haveEligible = !!(emergencyEligibilityCache &&
    Array.isArray(emergencyEligibilityCache.eligible) &&
    emergencyEligibilityCache.eligible.length);

  if (!haveEligible) {
    // Defensive: if clicked while not eligible, soft prompt and background check once.
    refreshEmergencyEligibility({ silent: true, hideDuringRefresh: true }).catch(() => {});
    showToast('No eligible shifts right now.');
    return;
  }

  ensureEmergencyOverlay();

  // Keep button interactive state predictable (brief busy for UX tightness; it stays visible & greyed if busy)
  setEmergencyButtonBusy(true);
  try {
    resetEmergencyState();
    emergencyState.eligible = emergencyEligibilityCache.eligible || [];
    emergencyState.step = 'PICK_SHIFT';
    openOverlay('emergencyOverlay', '#emergencyClose');
    renderEmergencyStep();
  } finally {
    setEmergencyButtonBusy(false);
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
      cancel.addEventListener('click', () => closeEmergencyOverlay(true));

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
    list.className = 'em-list';

    state.eligible.forEach((shift, idx) => {
      const id = `emShift_${idx}`;
      const item = document.createElement('label');
      item.setAttribute('role', 'listitem');
      item.className = 'em-row';
      item.innerHTML = `
        <input type="radio" name="emShift" id="${id}" value="${escapeHtml(shift.ymd || String(idx))}">
        <div class="em-text"><div class="em-title">${escapeHtml(buildEmergencyShiftLabel(shift))}</div></div>
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

    // 1) Start with server-declared allowed list (if present)
    if (Array.isArray(s.allowed_issues) && s.allowed_issues.length) {
      allowed = s.allowed_issues.map(x => String(x).toUpperCase());
    } else if (Array.isArray(s.allowedActions) && s.allowedActions.length) {
      allowed = s.allowedActions.map(x => String(x).toUpperCase());
    }

    // 2) Merge in tolerant signals (DON'T replace — union instead)
    const merged = new Set(allowed);

    // Running Late should appear if server signalled *either* canRunLate
    // or provided concrete late_options (even if not listed in allowed_issues)
    if (s.canRunLate === true || (Array.isArray(s.late_options) && s.late_options.length > 0)) {
      merged.add('RUNNING_LATE');
    }

    // Keep these tolerant merges too (backend sometimes sends booleans)
    if (s.canCancel === true)     merged.add('CANNOT_ATTEND');
    if (s.canLeaveEarly === true) merged.add('LEAVE_EARLY');
    if (s.canDNA === true || s.canDna === true) merged.add('DNA');

    allowed = Array.from(merged);

    const choices = document.createElement('div');
    choices.setAttribute('role', 'group');
    choices.className = 'em-list';

    function addOption(key, label, help) {
      const id = `emIssue_${key}`;
      const row = document.createElement('label');
      row.className = 'em-row';
      row.innerHTML = `
        <input type="radio" name="emIssue" id="${id}" value="${key}">
        <div class="em-text">
          <div class="em-title">${escapeHtml(label)}</div>
          ${help ? `<div class="em-help" style="margin-top:.15rem">${escapeHtml(help)}</div>` : ''}
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
        group.className = 'em-list';

        (state.lateOptions || []).forEach((opt, idx) => {
          const id = `lateOpt_${idx}`;
          const row = document.createElement('label');
          row.className = 'em-row em-late-row';
          const labelText = typeof opt === 'string' ? opt : (opt.label || '');
          row.innerHTML = `
            <input type="radio" name="emLate" id="${id}" value="${escapeHtml(labelText)}">
            <div class="em-text"><div class="em-title">${escapeHtml(labelText)}</div></div>
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
      rowNow.className = 'em-row';
      rowNow.style.marginBottom = '.35rem';
      rowNow.innerHTML = `<input type="radio" name="leaveTimeMode" value="NOW"> <div class="em-text">NOW / ASAP</div>`;

      const rowAt = document.createElement('div');
      rowAt.style.display = 'grid';
      rowAt.style.gridTemplateColumns = 'auto auto auto 1fr';
      rowAt.style.alignItems = 'center';
      rowAt.style.gap = '.4rem';

      const atLabel = document.createElement('label');
      atLabel.className = 'em-row';
      atLabel.style.padding = '0';
      atLabel.style.border = '0';
      atLabel.style.background = 'transparent';
      atLabel.innerHTML = `<input type="radio" name="leaveTimeMode" value="AT"> <div class="em-text">At time:</div>`;

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
        // ───────────────────────── DNA (Step A: select person) ─────────────────────────
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
      list.className = 'em-list';

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
        row.className = 'em-row';
        row.innerHTML = `
          <input type="radio" name="dnaAbsentee" id="${id}" value="${escapeHtml(name)}">
          <div class="em-text">
            <div class="em-title">${escapeHtml(name)}</div>
            ${role ? `<div class="em-help" style="margin-top:.15rem">${escapeHtml(role)}</div>` : ''}
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
      setFooter({ continueDisabled: true });

      const cont = document.getElementById('emergencyContinue');
      body.addEventListener('change', (e) => {
        if (e.target && e.target.name === 'dnaAbsentee') {
          cont && (cont.disabled = false);
        }
      });

      // Move to the new clean page showing only the selected person's number
      cont && cont.addEventListener('click', () => {
        const chosen = body.querySelector('input[name="dnaAbsentee"]:checked');
        if (!chosen || !chosen._abs) return;
        emergencyState.dnaAbsenteeName = chosen._abs.name || null;
        emergencyState.dnaAbsenteeMsisdn = chosen._abs.msisdn || null;
        emergencyState.dnaTriedCalling = false;
        emergencyState.step = 'DNA_CALL';
        renderEmergencyStep();
      });
    }


  }
  else if (state.step === 'DNA_CALL') {
    const name = emergencyState.dnaAbsenteeName || '';
    const msisdn = emergencyState.dnaAbsenteeMsisdn || '';
    if (!name) {
      // Safety: if someone navigates here without selecting, bounce back to selection
      emergencyState.step = 'DETAILS';
      renderEmergencyStep();
      return;
    }

    title.textContent = 'Please call';

    const box = document.createElement('div');
    box.style.border = '1px solid #2a3446';
    box.style.background = '#131926';
    box.style.borderRadius = '10px';
    box.style.padding = '.6rem';

    const dispNum = formatMsisdn07Display(msisdn);
    const msg = dispNum
      ? `Please call ${name} on ${dispNum}.`
      : `Please try calling ${name} via your usual contact method.`;

    const p = document.createElement('div');
    p.textContent = msg;
    box.appendChild(p);
    body.appendChild(box);

    const chkWrap = document.createElement('label');
    chkWrap.className = 'em-row';
    chkWrap.style.marginTop = '.6rem';
    chkWrap.innerHTML = `<input type="checkbox" id="dnaTriedCalling"> <div class="em-text">I tried calling them and could not reach them</div>`;
    body.appendChild(chkWrap);

    setFooter({ continueDisabled: true });

    const cont = document.getElementById('emergencyContinue');
    const chk = chkWrap.querySelector('#dnaTriedCalling');

    chk.addEventListener('change', () => {
      emergencyState.dnaTriedCalling = !!chk.checked;
      cont && (cont.disabled = !emergencyState.dnaTriedCalling);
    });

    cont && cont.addEventListener('click', handleEmergencyCollectDetails);
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
  finalizeEmergencyStepLayout();
}
// === New helper functions for Emergency layout ===

// 1) Inject a tiny, targeted stylesheet (idempotent)
function injectEmergencyStyles() {
  if (document.getElementById('emergencyStyles')) return;

  const css = `
    /* Container */
    #emergencyOverlay .sheet-body { overflow:auto; }

    /* Lists & rows */
    .em-list {
      display:flex; flex-direction:column; gap:.5rem;
      max-width: min(780px, 96vw); margin: 0 auto;
    }
    .em-row {
      display:grid; grid-template-columns:auto 1fr; align-items:center;
      gap:.6rem; border:1px solid #222a36; background:#131926;
      border-radius:10px; padding:.55rem .65rem;
      width:100%;
    }
    .em-row > input[type="radio"],
    .em-row > input[type="checkbox"] {
      width:1.1rem; height:1.1rem; min-width:1.1rem; min-height:1.1rem;
      margin:0; transform:none; /* keep controls compact */
    }

    /* Text wrapper keeps content on-grid */
    .em-text { min-width:0; }
    .em-title {
      font-weight:800; line-height:1.28;
      word-break:break-word; overflow-wrap:anywhere; hyphens:auto;
    }
    .em-help { color:#a7b0c0; line-height:1.25; }
    .em-note { color:#a7b0c0; line-height:1.25; margin:.6rem 0; max-width:min(780px,96vw); margin-left:auto; margin-right:auto; }

    /* Make long titles behave on narrow iPhones */
    @media (max-width: 380px) {
      .em-title { font-size:.95rem; }
      .em-row { padding:.5rem .55rem; }
    }

    /* Keep the sheet within viewport width neatly */
    #emergencyOverlay .sheet {
      max-width: min(820px, 96vw);
      margin-left:auto; margin-right:auto;
    }
  `;

  const style = document.createElement('style');
  style.id = 'emergencyStyles';
  style.textContent = css;
  document.head.appendChild(style);
}

// 2) Clamp the Emergency sheet/body to viewport (safe on resize; idempotent)
function measureAndClampEmergencySheet() {
  const overlay = document.getElementById('emergencyOverlay');
  if (!overlay) return;

  const sheet = overlay.querySelector('.sheet');
  const body  = overlay.querySelector('.sheet-body');
  if (sheet) {
    sheet.style.maxWidth = 'min(820px, 96vw)';
    sheet.style.marginLeft = 'auto';
    sheet.style.marginRight = 'auto';
  }
  if (body) {
    body.style.maxWidth = 'min(780px, 96vw)';
    body.style.marginLeft = 'auto';
    body.style.marginRight = 'auto';
  }
}

// 3) Normalize text blocks inside current Emergency step (safe to call after render)
function normalizeEmergencyText() {
  const overlay = document.getElementById('emergencyOverlay');
  if (!overlay) return;

  overlay.querySelectorAll('.em-title, .em-help, .em-text, .em-note').forEach(el => {
    el.style.whiteSpace = 'normal';
    el.style.wordBreak = 'break-word';
    el.style.overflowWrap = 'anywhere';
    el.style.hyphens = 'auto';
  });
}

// 4) Apply list layout tweaks (compact radios/checkboxes, tidy alignment)
function applyEmergencyListLayout() {
  const overlay = document.getElementById('emergencyOverlay');
  if (!overlay) return;

  overlay.querySelectorAll('.em-row > input[type="radio"], .em-row > input[type="checkbox"]').forEach(inp => {
    inp.style.width = '1.1rem';
    inp.style.height = '1.1rem';
    inp.style.minWidth = '1.1rem';
    inp.style.minHeight = '1.1rem';
    inp.style.margin = '0';
  });

  // Ensure text cell aligns and can wrap neatly
  overlay.querySelectorAll('.em-row .em-text').forEach(div => {
    div.style.alignSelf = 'center';
    div.style.minWidth = '0';
  });
}

// 5) Scroll the Emergency body to top (useful when switching steps)
function scrollEmergencyToTop() {
  const body = document.querySelector('#emergencyOverlay .sheet-body');
  if (body) body.scrollTo({ top: 0, behavior: 'smooth' });
}

// 6) Hook for resize/orientation changes (optional; safe to call)
function onEmergencyResize() {
  measureAndClampEmergencySheet();
}

// (Optional) one-liner to run after each step render:
function finalizeEmergencyStepLayout() {
  injectEmergencyStyles();
  measureAndClampEmergencySheet();
  normalizeEmergencyText();
  applyEmergencyListLayout();
  scrollEmergencyToTop();
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
    if (!nameOK) {
      // No person chosen yet → go back to selection
      emergencyState.step = 'DETAILS';
      renderEmergencyStep();
      return;
    }
    if (!emergencyState.dnaTriedCalling) {
      // Person chosen but not ticked the tried-calling box → stay on the clean call page
      emergencyState.step = 'DNA_CALL';
      renderEmergencyStep();
      return;
    }
    // Ready to confirm
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

  // double-submit guard
  if (emergencyState.__confirmSubmitting) return;
  emergencyState.__confirmSubmitting = true;

  const cont = document.getElementById('emergencyContinue');
  const cancel = document.getElementById('emergencyCancel');
  const origText = cont ? cont.textContent : '';

  if (cont) { cont.disabled = true; cont.textContent = 'Please wait…'; }
  if (cancel) cancel.disabled = true;

  const reenable = () => {
    if (cont) { cont.disabled = false; cont.textContent = origText || 'Continue'; }
    if (cancel) cancel.disabled = false;
    emergencyState.__confirmSubmitting = false;
  };

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
        const [hh, mm] = lbl.split(':').map(v => parseInt(v, 10));
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

  // Success → keep disabled; submitters will replace footer/overlay and trigger refresh
}

// Raise emergency issue; fixes border string bug and keeps button/eligibility flow intact.
async function submitEmergencyRaise(shift, issueType, reasonText, etaOrLeaveTimeLabel) {
  const bodyEl = document.getElementById('emergencyBody');
  const title  = document.getElementById('emergencyTitle');
  const footer = document.getElementById('emergencyFooter');
  if (!bodyEl || !title || !footer || !emergencyState) return;

  const it = String(issueType || '').toUpperCase();
  const leaveLabel = (it === 'LEAVE_EARLY')
    ? (emergencyState.selectedLateLabel || etaOrLeaveTimeLabel || 'N/A')
    : 'N/A';
  const reason = String(reasonText || emergencyState.reasonText || '').trim();

  let issuePayload;
  if (it === 'DNA') {
    issuePayload = {
      type: 'DNA',
      subject_name:  emergencyState.dnaAbsenteeName || '',
      subject_msisdn: emergencyState.dnaAbsenteeMsisdn || undefined,
      reason_text:   reason || ''
    };
  } else {
    issuePayload = {
      issue_type: it,
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
      cancel.addEventListener('click', () => closeEmergencyOverlay(true));

      const retry = document.createElement('button');
      retry.type = 'button';
      retry.textContent = 'Try again';
      retry.className = 'menu-item';
      retry.style.border = '1px solid #2a3446';
      retry.addEventListener('click', handleEmergencyConfirm);

      footer.append(cancel, retry);
      return;
    }

    // ✅ Immediately force-grey the Emergency button until a fresh eligibility response arrives
    window._emergencyForceGreyUntilFresh = true;
    try { setEmergencyButtonBusy(true); } catch {}

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

    // Kick the eligibility refresh (no-op while overlay is open; close handler will also trigger it)
    refreshEmergencyEligibility({ silent: true }).catch(() => {});
  } finally {
    hideLoading();
  }
}



// Simple helper used by polling/guards
function emergencyOpen() {
  const ov = document.getElementById('emergencyOverlay');
  return !!(ov && ov.classList.contains('show'));
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
      cancel.addEventListener('click', () => closeEmergencyOverlay(true));

      const retry = document.createElement('button');
      retry.type = 'button';
      retry.textContent = 'Try again';
      retry.className = 'menu-item';
      retry.style.border = '1px solid #2a3446';
      retry.addEventListener('click', handleEmergencyConfirm);
      footer.append(cancel, retry);
      return;
    }

    // ✅ Immediately force-grey the Emergency button until a fresh eligibility response arrives
    window._emergencyForceGreyUntilFresh = true;
    try { setEmergencyButtonBusy(true); } catch {}

    title.textContent = 'Sent';
    body.innerHTML = `<div>Your colleagues on shift and the Arthur Rai office have been informed.</div>`;
    footer.innerHTML = '';

    // header X → also triggers a reload after success
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

    // Kick the eligibility refresh (this is a no-op while the overlay is open; the close handler will also trigger it)
    refreshEmergencyEligibility({ silent: true, hideDuringRefresh: true }).catch(() => {});
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
// --- NEW: tiny helper used throughout emergency flow ---
function setEmergencyButtonBusy(busy) {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;
  btn.dataset.busy = busy ? '1' : '0';
  // Reflect latest cache/busy state
  syncEmergencyButtonVisibility(
    (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible))
      ? emergencyEligibilityCache.eligible
      : []
  );
}

// --- FINISH: previously truncated function ---
function updateEmergencyButtonFromCache() {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;

  btn.hidden = false; // never hide

  // If we're in a "force-grey-until-fresh" phase (e.g., long-away resume or post-submit),
  // do NOT repaint from cache. Keep the button grey & disabled until a fresh server result arrives.
  if (window._emergencyForceGreyUntilFresh) {
    btn.disabled = true;
    btn.style.filter = 'grayscale(1)';
    btn.style.opacity = '.85';
    return;
  }

  const busy = btn.dataset.busy === '1';
  const eligible = (emergencyEligibilityCache && Array.isArray(emergencyEligibilityCache.eligible))
    ? emergencyEligibilityCache.eligible
    : [];

  if (busy) {
    btn.disabled = true;
    btn.style.filter = 'grayscale(1)';
    btn.style.opacity = '.85';
    return;
  }

  if (eligible.length > 0) {
    // Bright red & enabled
    btn.disabled = false;
    btn.style.filter = '';
    btn.style.opacity = '';
    btn.style.background = '#d32f2f';
    btn.style.border = '1px solid #b71c1c';
    btn.style.color = '#fff';
  } else {
    // Greyed & disabled
    btn.disabled = true;
    btn.style.filter = 'grayscale(1)';
    btn.style.opacity = '.85';
  }
}


// ───────────────────────── NEW helper ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────
// ───────────────────────── CHANGED ─────────────────────────


// ===== Cache helpers =====

const BASELINE_CACHE_KEY = 'baseline_cache_v1';
const EMERGENCY_CACHE_KEY = 'emergency_cache_v1';
const NEXT_TILES_DUE_AT_KEY = 'next_tiles_due_at';

/**
 * Load the last saved baseline (tiles) payload from localStorage.
 * Returns { baseline, lastLoadedAt } or null if absent/invalid.
 */
function loadBaselineCache() {
  try {
    const raw = localStorage.getItem(BASELINE_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    // Minimal sanity check: must have tiles array if present
    if (parsed.baseline && parsed.baseline.tiles && !Array.isArray(parsed.baseline.tiles)) return null;
    return parsed;
  } catch {
    return null;
  }
}

/**
 * Save the current baseline and timestamp to localStorage.
 * @param {object} baseline - server baseline payload (including tiles[])
 * @param {string} [lastLoadedAt=new Date().toISOString()]
 */
function saveBaselineCache(baseline, lastLoadedAt) {
  try {
    if (!baseline || typeof baseline !== 'object') return;
    const payload = {
      baseline,
      lastLoadedAt: lastLoadedAt || new Date().toISOString()
    };
    localStorage.setItem(BASELINE_CACHE_KEY, JSON.stringify(payload));
  } catch {}
}

/**
 * Load cached emergency eligibility.
 * Returns { eligible: [] } or null if absent/invalid.
 */
function loadEmergencyEligibilityCache() {
  try {
    const raw = localStorage.getItem(EMERGENCY_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.eligible)) return null;
    return parsed;
  } catch {
    return null;
  }
}

/**
 * Save emergency eligibility to localStorage.
 * @param {{eligible: Array}} cache
 */
function saveEmergencyEligibilityCache(cache) {
  try {
    if (!cache || !Array.isArray(cache.eligible)) return;
    localStorage.setItem(EMERGENCY_CACHE_KEY, JSON.stringify({ eligible: cache.eligible }));
  } catch {}
}


// ===== Draft merge helpers =====

/**
 * Compute a stable, comparable signature for a tile based on all
 * server-controlled attributes that, if changed, should discard drafts.
 * The signature MUST NOT include any local draft edits.
 * @param {object} t - tile from server baseline
 * @returns {string} JSON string signature
 */
function computeTileSignature(t) {
  if (!t) return '∅';
  // Normalise key fields that drive the card UI and editability
  const norm = {
    ymd: String(t.ymd || ''),
    booked: !!t.booked,
    statusCode: t.status ?? null,              // server status code
    blocked: (t.status === 'BLOCKED') || !!t.blocked, // be tolerant to shapes
    editable: (t.editable !== false) && !t.booked && t.status !== 'BLOCKED',

    // Card-visible fields (any change nukes draft)
    shiftInfo: (t.shiftInfo || '').trim(),
    hospital: (t.hospital || '').trim(),
    location: (t.location || '').trim(),
    ward: (t.ward || '').trim(),
    jobTitle: (t.jobTitle || '').trim(),
    bookingRef: (t.bookingRef || '').trim(),

    // Server-provided labels often used in UI
    display_label: (t.display_label || '').trim(),
    time_range_label: (t.time_range_label || '').trim(),
    shift_type_label: (t.shift_type_label || '').trim(),
  };

  return JSON.stringify(norm);
}

/**
 * Merge/cull drafts against a new baseline based on unchanged-tile rule.
 * Server wins on ANY change between old vs new tile signatures.
 *
 * @param {object|null} oldBaseline - previously rendered baseline (may be null)
 * @param {object} newBaseline - fresh baseline from server
 * @param {object} drafts - current draft map { [ymd]: 'LABEL' }
 * @returns {object} survivingDrafts - filtered draft map
 */
function mergeDraftsAgainstNewBaseline(oldBaseline, newBaseline, drafts) {
  if (!drafts || !newBaseline || !Array.isArray(newBaseline.tiles)) return {};
  const out = {};
  const oldTiles = (oldBaseline && Array.isArray(oldBaseline.tiles)) ? oldBaseline.tiles : [];

  const oldByYmd = new Map(oldTiles.map(t => [String(t.ymd), t]));
  const newByYmd = new Map(newBaseline.tiles.map(t => [String(t.ymd), t]));

  for (const [ymd, draftLabel] of Object.entries(drafts)) {
    const oldTile = oldByYmd.get(ymd) || null;
    const newTile = newByYmd.get(ymd) || null;

    // If the day left or entered the window → discard
    if (!oldTile || !newTile) continue;

    // If editability flipped to non-editable → discard
    const newEditable = (newTile.editable !== false) && !newTile.booked && newTile.status !== 'BLOCKED';
    if (!newEditable) continue;

    // Compare server → server signatures (not vs draft)
    const oldSig = computeTileSignature(oldTile);
    const newSig = computeTileSignature(newTile);

    if (oldSig === newSig) {
      out[ymd] = draftLabel; // tile unchanged → keep draft
    }
    // else: any change → discard draft
  }

  return out;
}


// ===== Cadence / orchestration helpers =====

/**
 * Is a tiles refresh due now, considering 3-minute cadence and any snooze?
 * Returns true only when overlay is closed and current time >= due time and >= snooze.
 */
function isTilesRefreshDue() {
  if (typeof isBlockingOverlayOpen === 'function' && isBlockingOverlayOpen('emergencyOverlay')) return false;

  const now = Date.now();
  let dueAt = 0;
  let snoozeUntil = 0;

  try {
    const dueStr = localStorage.getItem(NEXT_TILES_DUE_AT_KEY);
    const snoozeStr = localStorage.getItem(REFRESH_SNOOZE_UNTIL_KEY);
    dueAt = dueStr ? parseInt(dueStr, 10) : 0;
    snoozeUntil = snoozeStr ? parseInt(snoozeStr, 10) : 0;
  } catch {}

  if (snoozeUntil && now < snoozeUntil) return false;
  if (!dueAt) return true; // no schedule yet → allow a run
  return now >= dueAt;
}

// Cadence state

// ===== Cadence (Tiles drive; Emergency is chained) =====

// ===== Cadence (Tiles drive; Emergency is chained) =====
const NEXT_KEY   = (typeof NEXT_TILES_DUE_AT_KEY === 'string' ? NEXT_TILES_DUE_AT_KEY : 'next_tiles_due_at');
const SNOOZE_KEY = (typeof REFRESH_SNOOZE_UNTIL_KEY === 'string' ? REFRESH_SNOOZE_UNTIL_KEY : 'rota_refresh_snooze_until_v1');

let tilesInFlight  = false;
let nextTilesDueAt = 0; // epoch ms for next tiles refresh

// Driver chain: Tiles → Emergency → restart 3-minute clock
// --- NEW: single path for refresh UX ---
async function startTilesThenEmergencyChain({ force = false } = {}) {
  if (startTilesThenEmergencyChain._inflight) return startTilesThenEmergencyChain._inflight;

  // Ensure top bar state before we potentially render from cache (prevents menu spill)
  try { ensureTopBarStatus(); } catch {}

  // We still fetch during the short guard window; loadFromServer will merge recent optimistic values
  const work = (async () => {
    try { await loadFromServer({ force }); }
    catch (e) {
      if (String(e && e.message) !== '__AUTH_STOP__') {
        // Non-auth errors are already handled inside loadFromServer when relevant.
      }
      return;
    }
    try { await refreshEmergencyEligibility({ silent: true }); } catch {}
  })();

  startTilesThenEmergencyChain._inflight = work.finally(() => {
    // Advance the 3-minute cadence regardless of success/failure
    try { if (typeof scheduleNextTilesRefresh === 'function') scheduleNextTilesRefresh(Date.now()); } catch {}
    startTilesThenEmergencyChain._inflight = null;
  });

  return startTilesThenEmergencyChain._inflight;
}

function ensureTopBarStatus() {
  try {
    const top = document.getElementById('topBar') || document.querySelector('.topbar') || document.querySelector('header');
    if (!top) return;

    // Prefer explicit ids created by ensureMenu
    const menuList = document.getElementById('menuList') || top.querySelector('[data-role="menu-list"], .menu-items, nav ul');
    const menuBtn  = document.getElementById('menuBtn')   || document.getElementById('menuToggle') || top.querySelector('[data-role="menu-toggle"], .menu-toggle, button[aria-controls]');

    // Force collapsed-by-default even if CSS hasn’t been injected yet
    if (menuList) {
      menuList.classList.remove('show');
      menuList.setAttribute('hidden', 'hidden');
      menuList.style.display = 'none';
    }
    if (menuBtn) {
      menuBtn.setAttribute('aria-expanded', 'false');
    }
  } catch {}
}


function setTopBarUpdating(on) {
  try {
    const top = document.getElementById('topBar') || document.querySelector('.topbar') || document.querySelector('header');
    if (!top) return;

    // Find a sensible anchor: the Refresh button if present; else the emergency button; else append to the bar.
    const refreshBtn =
      document.getElementById('refreshBtn') ||
      top.querySelector('[data-role="refresh"], .refresh-btn, button.refresh');
    const anchor = refreshBtn?.parentNode || top;

    let badge = document.getElementById('topbarUpdating');
    if (on) {
      if (!badge) {
        badge = document.createElement('span');
        badge.id = 'topbarUpdating';
        badge.textContent = 'Updating…';
        // Subtle, non-blocking inline styling to avoid layout shifts
        badge.style.display = 'inline-flex';
        badge.style.alignItems = 'center';
        badge.style.fontSize = '0.85rem';
        badge.style.opacity = '0.7';
        badge.style.marginRight = '0.75rem';
        badge.style.whiteSpace = 'nowrap';

        if (refreshBtn && anchor) {
          anchor.insertBefore(badge, refreshBtn);
        } else {
          top.appendChild(badge);
        }
      } else {
        badge.style.display = 'inline-flex';
      }
    } else if (badge) {
      badge.style.display = 'none';
    }
  } catch {}
}

// Optional helper used by your closeEmergencyOverlay() fallback:



// 30s poller: only evaluates Tiles due-ness; NEVER calls emergency directly
(function ensureTilesPoller() {
  if (window._tilesPollerId) return; // avoid duplicate intervals on hot-reload
  window._tilesPollerId = setInterval(async () => {
    if (typeof emergencyOpen === 'function' && emergencyOpen()) return; // pause while overlay open
    if (tilesInFlight) return;

    const now   = Date.now();
    const dueAt = nextTilesDueAt || Number(localStorage.getItem(NEXT_KEY) || 0);
    if (dueAt && now < dueAt) return;

    const snoozeUntil = Number(localStorage.getItem(SNOOZE_KEY) || 0);
    if (snoozeUntil && now < snoozeUntil) return;

    await startTilesThenEmergencyChain({ force: false });
  }, 30000);
})();

// Bootstrap cadence on first load:
//  - Restore a persisted next-due (if any), otherwise allow immediate run.
//  - Kick the first chain right away (guards prevent double-runs).
// Bootstrap cadence on first load:
//  - Restore a persisted next-due (if any), otherwise leave as 0 to allow immediate chain kick,
//  - Kick the first chain right away; the chain will set next due = now + 3m on completion.
(function bootstrapCadence() {
  try {
    const restored = Number(localStorage.getItem(NEXT_KEY) || 0);
    if (restored) {
      nextTilesDueAt = restored;
    } else {
      nextTilesDueAt = 0; // allow immediate first run; chain will set +3m on completion
    }
  } catch {
    nextTilesDueAt = 0;
  }

  // Kick once now; subsequent runs are timer-driven
  startTilesThenEmergencyChain({ force: false }).catch(() => {});
})();


// Optional helper if you want to set next-due explicitly elsewhere.
function scheduleNextTilesRefresh(fromEpochMs) {
  const base = Number.isFinite(fromEpochMs) ? fromEpochMs : Date.now();
  const next = base + 3 * 60 * 1000;
  nextTilesDueAt = next;
  try { localStorage.setItem(NEXT_KEY, String(next)); } catch {}
  return next;
}


/**
 * Start the Tiles → Emergency chain with single-flight guards and cadence restart.
 * Use this for: resume, timer tick when due, and "bump forward" on emergency close.
 */

/**
 * Can we start an emergency refresh right now?
 * Enforces: overlay closed, not inflight, and not within cooldown.
 */
function canRunEmergencyRefreshNow() {
  if (typeof isBlockingOverlayOpen === 'function' && isBlockingOverlayOpen('emergencyOverlay')) return false;
  if (window.emergencyInFlight) return false;
  const now = Date.now();
  return !(window.emergencyCooldownUntil && now < window.emergencyCooldownUntil);
}

/**
 * Set a short cooldown to deduplicate overlapping emergency triggers.
 * @param {number} ms default 8000ms
 */
function markEmergencyCooldown(ms = 8000) {
  try { window.emergencyCooldownUntil = Date.now() + Math.max(0, ms); } catch {}
  return window.emergencyCooldownUntil;
}


// ===== Tile theming & emergency button state helpers =====

/**
 * Apply background + text tone to a tile/card based on the effective status.
 * @param {HTMLElement} card - the tile element
 * @param {string} statusLabel - one of BOOKED/BLOCKED/NOT AVAILABLE/LONG DAY/NIGHT/LONG DAY/NIGHT/PENDING...
 */
function applyTileThemeForStatus(card, statusLabel) {
  if (!card) return;

  // Clear prior tone classes
  card.classList.remove('tile-bg-blue', 'tile-bg-soft-red', 'tile-bg-bright-red', 'tile-bg-green', 'tile-bg-black');
  card.classList.remove('tile-text-dark', 'tile-text-light');

  const headline = card.querySelector('.tile-status');
  const subtext  = card.querySelector('.tile-sub');

  // Helper to set text tone across both
  const setTone = (tone /* 'dark' | 'light' */) => {
    const dark = tone === 'dark';
    card.classList.add(dark ? 'tile-text-dark' : 'tile-text-light');
    if (headline) headline.classList.toggle('text-light', !dark);
    if (headline) headline.classList.toggle('text-dark',  dark);
    if (subtext)  subtext.classList.toggle('text-light', !dark);
    if (subtext)  subtext.classList.toggle('text-dark',  dark);
  };

  const upper = String(statusLabel || '').toUpperCase();

  if (upper === 'BOOKED') {
    card.classList.add('tile-bg-blue');
    setTone('dark');
  } else if (upper === 'BLOCKED') {
    card.classList.add('tile-bg-soft-red');
    setTone('dark');
  } else if (upper === 'NOT AVAILABLE') {
    card.classList.add('tile-bg-bright-red');
    setTone('light');
  } else if (upper === 'LONG DAY' || upper === 'NIGHT' || upper === 'LONG DAY/NIGHT' || upper === 'AVAILABLE' || upper === 'AVAILABLE FOR') {
    card.classList.add('tile-bg-green');
    setTone('dark');
  } else if (upper === 'PLEASE PROVIDE YOUR AVAILABILITY' || upper === 'NOT SURE OF MY AVAILABILITY YET' || upper === 'PENDING') {
    card.classList.add('tile-bg-black');
    setTone('light');
  } else {
    // Default to pending/unknown → keep readable
    card.classList.add('tile-bg-black');
    setTone('light');
  }
}

/** Greyed state for emergency button (visible & disabled). */
function setEmergencyButtonStateGreyed() {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;
  btn.hidden = false;
  btn.disabled = true;
  btn.style.filter = 'grayscale(1)';
  btn.style.opacity = '.85';
  btn.style.background = '#6b6b6b';
  btn.style.border = '1px solid #5a5a5a';
  btn.style.color = '#fff';
}

/** BRIGHT RED enabled state for emergency button. */
function setEmergencyButtonStateRed() {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;
  btn.hidden = false;
  btn.disabled = false;
  btn.style.filter = '';
  btn.style.opacity = '';
  btn.style.background = '#d32f2f';
  btn.style.border = '1px solid #b71c1c';
  btn.style.color = '#fff';
}

/**
 * Overlay primary button busy semantics: disable + "Please wait…" until caller resets/closes.
 * @param {HTMLElement} btn
 */
function setOverlayPrimaryButtonBusy(btn) {
  if (!btn) return;
  btn.dataset._normalLabel = (btn.textContent || '').trim();
  btn.disabled = true;
  btn.textContent = 'Please wait…';
}

/**
 * Restore overlay primary button to its normal label and enabled state.
 * @param {HTMLElement} btn
 */
function resetOverlayPrimaryButton(btn) {
  if (!btn) return;
  const normal = btn.dataset._normalLabel || 'OK';
  delete btn.dataset._normalLabel;
  btn.disabled = false;
  btn.textContent = normal;
}


// ===== Emergency-close & resume helpers =====

/**
 * Decide if tiles should be "bumped forward" on emergency overlay close.
 * Return true when the next tiles refresh is due within 60s.
 */
function shouldBumpTilesOnEmergencyClose() {
  try {
    const lastAt = localStorage.getItem(LAST_LOADED_KEY);
    const nextDue = lastAt ? (Date.parse(lastAt) + 3 * 60 * 1000) : 0;
    if (!nextDue) return true; // no history → safe to run
    return (nextDue - Date.now()) <= 60 * 1000;
  } catch {
    return true;
  }
}

/**
 * Resume handler: when app/tab becomes visible and overlay is closed,
 * immediately start the Tiles → Emergency chain and restart cadence from completion.
 */
async function tilesResumeHandler() {
  if (typeof emergencyOpen === 'function' && emergencyOpen()) return;

  // On app return (tab visibilitychange), we must start from a known-safe state:
  //  - Force-grey the button until the server confirms eligibility again.
  //  - Mark a "force grey until fresh" phase so refreshEmergencyEligibility honours it.
  window._emergencyForceGreyUntilFresh = true;
  try { setEmergencyButtonBusy(true); } catch {}

  await startTilesThenEmergencyChain({ force: true });

  // Advance cadence right after resume-triggered refresh
  try { if (typeof scheduleNextTilesRefresh === 'function') scheduleNextTilesRefresh(Date.now()); } catch {}
}



// Tiles driver only. Do nothing while Emergency overlay is open.
// Emergency eligibility is handled by the cadence chain, not here.
async function refreshAppState({ force = false } = {}) {
  // Don’t refresh underneath the Emergency flow or any blocking modal.
  if (typeof emergencyOpen === 'function' && emergencyOpen()) return;
  if (typeof isBlockingOverlayOpen === 'function' && isBlockingOverlayOpen()) return;

  // Funnel through the single-flight chain to avoid races (pull-to-refresh, visibility, manual).
  // This mirrors the old “one inflight” pattern and prevents a stray call from deciding
  // “no identity yet → open login” while another succeeds.
  return startTilesThenEmergencyChain({ force });
}


// ─────────────────────────────────────────────────────────────────────────────
// Timesheets: new helpers for attention state (TS !), test-mode allowlist,
// and 4s content alternation. Pure frontend; no network.
// ─────────────────────────────────────────────────────────────────────────────
(function () {
  const TS = (window.Timesheets = window.Timesheets || {});

  // ---- tiny utils (private) ----
  function _normalizeMsisdn(n) {
    if (!n) return '';
    let s = String(n).replace(/[^\d+]/g, '');
    if (s.startsWith('+')) s = s.slice(1);
    // Common UK shapes → canonical 447XXXXXXXXX
    if (s.startsWith('07') && s.length === 11) return '44' + s.slice(1);
    if (s.startsWith('7') && s.length === 10) return '44' + s;
    if (s.startsWith('447') && s.length === 12) return s;
    if (s.startsWith('44') && s.length === 12) return s;
    return s; // best effort
  }
  function _isAllowlistedMsisdn(n) {
    const canon = _normalizeMsisdn(n);
    // Whitelist: 07545970472 / 7545970472 (any formatting)
    const tail = '7545970472';
    return canon.endsWith(tail) || canon === '44' + tail;
  }
  function _getConfig() {
    return window.CONFIG || {};
  }

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: isTestModeAllowed(identity)
  // - If test mode OFF → true
  // - If test mode ON  → true only for allowlisted msisdn
  // ───────────────────────────────────────────────────────────────────────────
 TS.isTestModeAllowed = function isTestModeAllowed(identity) {
  const cfg = _getConfig();
  const testMode = !!cfg.TIMESHEET_TESTMODE;
  if (!testMode) return true;

  const name = _candidateNameFrom(identity, window.baseline);
  return _isAllowlistedName(name);
};

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: shouldShowTimesheetCTA(tile, identity)
  // Predicate for the attention state (TS ! + orange + alternator).
  // Conditions:
  //   - feature enabled for user
  //   - test mode allowed (see above)
  //   - tile.booked === true
  //   - tile.timesheet_eligible === true
  //   - tile.timesheet_authorised !== true
  // ───────────────────────────────────────────────────────────────────────────
  TS.shouldShowTimesheetCTA = function shouldShowTimesheetCTA(tile, identity) {
  if (!tile) return false;

  // Pass identity (or it will pick up baseline/identity from globals)
  const enabled =
    typeof TS.isTimesheetFeatureEnabled === 'function'
      ? !!TS.isTimesheetFeatureEnabled(identity || window.baseline || '')
      : false;

  if (!enabled) return false;
  if (!TS.isTestModeAllowed(identity)) return false;

  const isBooked = !!tile.booked;
  const eligible = tile.timesheet_eligible === true;
  const notAuthorised = tile.timesheet_authorised !== true;

  return isBooked && eligible && notAuthorised;
};

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: applyCTATheme(cardEl)
  // Adds an orange attention background class. CSS should define .tile--bg-ts-cta
  // (kept separate so you can centralize theme toggling if desired).
  // ───────────────────────────────────────────────────────────────────────────
  TS.applyCTATheme = function applyCTATheme(cardEl) {
    if (!cardEl) return;
    cardEl.classList.add('tile--bg-ts-cta');
  };

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: decorateTileWithTSAttention(cardEl, tile)
  // - Adds "TS !" badge (top-right, like TS ✓)
  // - Applies orange bg via .tile--bg-ts-cta
  // - Does NOT start the alternator (that’s separate so callers can control)
  // ───────────────────────────────────────────────────────────────────────────
  TS.decorateTileWithTSAttention = function decorateTileWithTSAttention(cardEl /*, tile */) {
    if (!cardEl) return;

    // Apply orange theme
    TS.applyCTATheme(cardEl);

    // Ensure a single attention badge
    const existing = cardEl.querySelector('.ts-badge-attn');
    if (!existing) {
      const badge = document.createElement('div');
      badge.className = 'ts-badge ts-badge-attn';
      badge.textContent = 'TS !';
      // If you don’t already position via CSS, this ensures visible placement.
      // You can remove these styles if your CSS handles it.
      badge.style.position = 'absolute';
      badge.style.top = '6px';
      badge.style.right = '6px';
      badge.style.padding = '2px 6px';
      badge.style.borderRadius = '10px';
      badge.style.fontSize = '12px';
      badge.style.fontWeight = '700';
      badge.style.background = 'rgba(255, 165, 0, 0.95)';
      badge.style.color = '#000';
      badge.style.pointerEvents = 'none';
      cardEl.appendChild(badge);
    }
  };

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: startTileContentAlternator(cardEl, { originalHtml, altText, periodMs })
  // Flips the tile's .tile-sub innerHTML every 'periodMs' between originalHtml
  // and altText. Idempotent; restarts cleanly if already running.
  // ───────────────────────────────────────────────────────────────────────────
  TS.startTileContentAlternator = function startTileContentAlternator(
    cardEl,
    { originalHtml, altText = 'Please touch to sign your timesheet', periodMs = 4000 } = {}
  ) {
    if (!cardEl) return;

    // Clear any existing alternator first
    TS.stopTileContentAlternator(cardEl);

    const sub = cardEl.querySelector('.tile-sub');
    if (!sub) return;

    const orig = originalHtml != null ? String(originalHtml) : sub.innerHTML;
    cardEl.dataset.tsAltOriginalHtml = orig;

    let showingAlt = false;
    const id = window.setInterval(() => {
      try {
        if (!sub.isConnected) { TS.stopTileContentAlternator(cardEl); return; }
        showingAlt = !showingAlt;
        sub.innerHTML = showingAlt ? altText : orig;
      } catch (_) {
        // On any unexpected DOM failure, stop gracefully.
        TS.stopTileContentAlternator(cardEl);
      }
    }, Math.max(1000, Number(periodMs) || 4000));

    cardEl.dataset.tsAltTimerId = String(id);
  };

  // ───────────────────────────────────────────────────────────────────────────
  // NEW: stopTileContentAlternator(cardEl)
  // Clears the interval and restores the original sub innerHTML if we saved it.
  // ───────────────────────────────────────────────────────────────────────────
  TS.stopTileContentAlternator = function stopTileContentAlternator(cardEl) {
    if (!cardEl) return;

    const id = Number(cardEl.dataset.tsAltTimerId || 0);
    if (id) {
      try { window.clearInterval(id); } catch (_) {}
      delete cardEl.dataset.tsAltTimerId;
    }

    const sub = cardEl.querySelector('.tile-sub');
    const orig = cardEl.dataset.tsAltOriginalHtml;
    if (sub && orig != null) {
      sub.innerHTML = orig;
    }
    delete cardEl.dataset.tsAltOriginalHtml;

    // Remove attention visuals if present (caller may also do this)
    const attn = cardEl.querySelector('.ts-badge-attn');
    if (attn) attn.remove();
    cardEl.classList.remove('tile--bg-ts-cta');
  };
})();




// Busy helper for the Emergency button — never hides; only greys/disable while busy.
// On un-busy, defer final state to updateEmergencyButtonFromCache().
function setEmergencyButtonBusy(isBusy) {
  const btn = document.getElementById('emergencyBtn');
  if (!btn) return;

  if (isBusy) {
    btn.hidden = false;                 // never hide
    btn.disabled = true;
    btn.dataset.busy = '1';
    btn.setAttribute('aria-busy', 'true');
    btn.style.pointerEvents = 'none';
    btn.style.filter = 'grayscale(1)';
    btn.style.opacity = '.85';
  } else {
    delete btn.dataset.busy;
    btn.removeAttribute('aria-busy');
    btn.style.pointerEvents = '';
    btn.style.filter = '';
    btn.style.opacity = '';
    // Do not force-enable here; let the cache/state machine decide:
    updateEmergencyButtonFromCache();
  }
}

/* ============================================================================
 *  Timesheets Frontend Module (new code only)
 *  - Drop this file into your frontend repo (e.g., /public/js/timesheets.js)
 *  - Wire calls from your existing render / click handlers (see summaries)
 * ============================================================================ */

window.Timesheets = (function () {
  "use strict";

  /* ──────────────────────────────────────────────────────────────────────────
   * 0) Config (safe defaults; override via window.CONFIG)
   * ────────────────────────────────────────────────────────────────────────── */
  const CFG = {
    BROKER_BASE_URL:
      (window.CONFIG && window.CONFIG.BROKER_BASE_URL) ||
      "https://arthur-rai-broker.kier-88a.workers.dev",
    TIMESHEET_FEATURE_ENABLED:
      (window.CONFIG && window.CONFIG.TIMESHEET_FEATURE_ENABLED) !== undefined
        ? !!window.CONFIG.TIMESHEET_FEATURE_ENABLED
        : true,
    TIMESHEET_TESTMODE:
      (window.CONFIG && window.CONFIG.TIMESHEET_TESTMODE) !== undefined
        ? !!window.CONFIG.TIMESHEET_TESTMODE
        : false,
    LONDON_TZ: "Europe/London",
    UPLOAD_MAX_BYTES: 300000, // keep aligned with broker default
  };

  /* ──────────────────────────────────────────────────────────────────────────
   * 1) Identity helpers (TESTMODE gating)
   * ────────────────────────────────────────────────────────────────────────── */

  function normalizeMsisdn(msisdn) {
    if (!msisdn) return "";
    const digits = String(msisdn).replace(/\D+/g, "");
    if (digits.startsWith("44")) return digits;
    if (digits.startsWith("0")) return "44" + digits.slice(1);
    return digits;
  }
function isTimesheetTestUser(subject) {
  // subject may be a name string, an identity object, or a baseline object
  const name = _candidateNameFrom(
    typeof subject === 'object' ? subject : window.identity,
    typeof subject === 'object' ? subject : window.baseline,
    typeof subject === 'string' ? subject : ''
  );
  return _isAllowlistedName(name);
}

function isTimesheetFeatureEnabled(subject) {
  if (!CFG.TIMESHEET_FEATURE_ENABLED) return false;
  if (CFG.TIMESHEET_TESTMODE) {
    // TESTMODE true ⇒ only the NAME-allowlisted user sees features
    return isTimesheetTestUser(subject);
  }
  // TESTMODE false ⇒ everyone sees features
  return true;
}
  /* ──────────────────────────────────────────────────────────────────────────
   * 2) Occupant key + booking id helpers (for History/status lookups)
   * ────────────────────────────────────────────────────────────────────────── */

  function deriveOccupantKey(baseline) {
    // Preferred: surname + firstname
    const s =
      baseline?.candidate?.surname ||
      baseline?.candidate?.lastName ||
      baseline?.candidate?.last_name ||
      "";
    const f =
      baseline?.candidate?.firstName ||
      baseline?.candidate?.firstname ||
      baseline?.candidate?.given_name ||
      "";
    const joined = `${String(s).trim()} ${String(f).trim()}`.trim();
    if (joined) return joined;
    // fallback
    if (baseline?.candidateName) return String(baseline.candidateName).trim();
    return "";
  }



  function _normForBooking(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^\w\s\-@&\/,.:]/g, "");
  }

  async function _sha256Hex(str) {
    const data = new TextEncoder().encode(str);
    const buf = await crypto.subtle.digest("SHA-256", data);
    const arr = Array.from(new Uint8Array(buf));
    return arr.map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  /**
   * Make a booking_id deterministically (matches broker’s algorithm).
   * Use when your history item lacks booking_id but you still need to query status.
   */
  async function makeBookingId({
    occupant_key,
    date_start_local, // YYYY-MM-DD
    hospital,
    ward,
    job_title,
    shift_label = "",
  }) {
    const base = `${_normForBooking(occupant_key)}|${date_start_local}|${_normForBooking(
      hospital
    )}|${_normForBooking(ward)}|${_normForBooking(job_title)}|${_normForBooking(shift_label)}`;
    const hash = await _sha256Hex(base);
    return `bk_${hash.slice(0, 16)}`;
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 3) Broker client (fetch helpers)
   * ────────────────────────────────────────────────────────────────────────── */

  async function _requestBroker(path, { method = "GET", body = null, headers = {} } = {}) {
    const url = path.startsWith("http") ? path : `${CFG.BROKER_BASE_URL}${path}`;
    const opts = {
      method,
      headers: {
        "Content-Type": "application/json",
        ...headers,
      },
      credentials: "omit",
    };
    if (body) opts.body = JSON.stringify(body);

    let res, text;
    try {
      res = await fetch(url, opts);
      text = await res.text();
    } catch (e) {
      return { ok: false, status: 0, error: e.message || "Network error" };
    }
    let json = null;
    try {
      json = text ? JSON.parse(text) : null;
    } catch {
      // non-JSON; keep raw
    }
    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        json,
        error: (json && (json.error || json.message)) || text || `HTTP ${res.status}`,
      };
    }
    return { ok: true, status: res.status, json };
  }

  async function presignTimesheet(payload) {
    return _requestBroker("/timesheets/presign", { method: "POST", body: payload });
  }

  async function uploadSignaturePutUrl(putUrl, pngBlob) {
    const res = await fetch(putUrl, {
      method: "PUT",
      headers: { "Content-Type": "image/png" },
      body: pngBlob,
    }).catch((e) => ({ ok: false, status: 0, _err: e }));
    if (!res || !res.ok) {
      let text = "";
      try {
        text = await res.text();
      } catch {}
      return {
        ok: false,
        status: res?.status || 0,
        error: text || "Upload failed",
      };
    }
    return { ok: true, status: res.status };
  }

  async function submitTimesheet(payload) {
    return _requestBroker("/timesheets/submit", { method: "POST", body: payload });
  }

  async function getTimesheet(booking_id) {
    return _requestBroker(`/timesheets/${encodeURIComponent(booking_id)}`, { method: "GET" });
  }

  async function revokeTimesheet(booking_id, reason) {
    return _requestBroker("/timesheets/revoke", {
      method: "POST",
      body: { booking_id, reason: reason || "" },
    });
  }

  async function revokeAndPresign(body) {
    return _requestBroker("/timesheets/revoke-and-presign", {
      method: "POST",
      body,
    });
  }

  async function authorisedStatus(booking_ids) {
    return _requestBroker("/timesheets/authorised-status", {
      method: "POST",
      body: { booking_ids: booking_ids || [] },
    });
  }

  async function presignSignatureGet(booking_id, which = "nurse", expires_seconds = 180) {
    return _requestBroker("/signatures/presign-get", {
      method: "POST",
      body: { booking_id, which, expires_seconds },
    });
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 4) History augmentation (fetch authorised flags from the broker)
   * ────────────────────────────────────────────────────────────────────────── */

  /**
   * Given an array of history rows, ensure each has a booking_id, then call the broker
   * to learn which items are authorised. Returns a map { booking_id: boolean } and
   * also mutates rows to set row.timesheet_authorised = true/false.
   *
   * Expected minimal row fields per item:
   *   - booking_id (optional)
   *   - occupant_key (optional if you provide baseline separately)
   *   - ymd (YYYY-MM-DD, the shift date)
   *   - hospital, ward, job_title, shift_label? (required only if we must compute booking_id)
   */
  async function augmentHistoryWithTimesheetStatus(rows, baseline) {
    if (!Array.isArray(rows) || !rows.length) return {};

    // Build/collect booking_ids
    const ids = [];
    for (const r of rows) {
      if (!r.booking_id) {
        const occupant_key = r.occupant_key || deriveOccupantKey(baseline);
        r.booking_id = await makeBookingId({
          occupant_key,
          date_start_local: r.ymd || r.date_start_local,
          hospital: r.hospital,
          ward: r.ward,
          job_title: r.job_title || r.role || r.jobTitle,
          shift_label: r.shift_label || r.shift || r.shiftLabel || "",
        });
      }
      ids.push(r.booking_id);
    }

    // Query broker
    const resp = await authorisedStatus(ids);
    const map = {};
    if (resp.ok && resp.json && resp.json.statuses) {
      Object.assign(map, resp.json.statuses);
      // mutate rows for immediate UI
      for (const r of rows) {
        r.timesheet_authorised = !!map[r.booking_id];
      }
    } else {
      // broker unavailable; mark as false but keep function resilient
      for (const r of rows) r.timesheet_authorised = !!r.timesheet_authorised;
    }
    return map;
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 5) TS ✓ badge for tiles
   * ────────────────────────────────────────────────────────────────────────── */

  function decorateTileWithTSBadge(cardEl, tile) {
    try {
      if (!cardEl || !tile) return;
      const badgeClass = "ts-badge";
      // remove any existing badge first
      cardEl.querySelectorAll("." + badgeClass).forEach((n) => n.remove());
      if (tile.timesheet_authorised !== true) return;

      const badge = document.createElement("div");
      badge.className = badgeClass;
      badge.textContent = "TS ✓";
      // minimal styles for the badge if your CSS isn’t ready:
      Object.assign(badge.style, {
        position: "absolute",
        top: "6px",
        right: "6px",
        fontSize: "11px",
        lineHeight: "1",
        padding: "4px 6px",
        borderRadius: "10px",
        background: "#0a7f43",
        color: "#fff",
        fontWeight: "600",
        letterSpacing: "0.2px",
        pointerEvents: "none",
      });
      // ensure tile is position:relative
      const pos = getComputedStyle(cardEl).position;
      if (pos === "static") cardEl.style.position = "relative";
      cardEl.appendChild(badge);
    } catch {
      /* no-op */
    }
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 6) Click router (tile → Wizard or Submitted Modal)
   * ────────────────────────────────────────────────────────────────────────── */

function onTileClickTimesheetBranch(tile, identity, baseline) {
  // NAME-BASED gate
  if (!isTimesheetFeatureEnabled(baseline || identity || '')) return false;

  // Test-mode allowlist: only the NAME-whitelisted user may use the flow in test mode
  try {
    if (window.Timesheets && typeof window.Timesheets.isTestModeAllowed === 'function') {
      if (!window.Timesheets.isTestModeAllowed(identity)) return false;
    }
  } catch (_) {}

  if (tile?.timesheet_authorised === true) {
    openSubmittedTimesheetModal(tile, identity, baseline);
    return true;
  }

  if (tile?.timesheet_eligible === true) {
    startTimesheetWizard(tile, identity, baseline);
    return true;
  }

  // Not eligible; do nothing special (preserves existing behaviour)
  return false;
}

  /* ──────────────────────────────────────────────────────────────────────────
   * 7) Wizard (minimal self-contained overlay)
   *     - Step 1: Hours & Break
   *     - Step 2: Nurse signature
   *     - Step 3: Authoriser details + signature → Submit
   * ────────────────────────────────────────────────────────────────────────── */

  // Utility: make DOM
  function _el(tag, attrs = {}, ...children) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (k === "style" && v && typeof v === "object") Object.assign(n.style, v);
      else if (k.startsWith("on") && typeof v === "function") n[k] = v;
      else if (v !== null && v !== undefined) n.setAttribute(k, v);
    }
    for (const c of children) n.append(c?.nodeType ? c : document.createTextNode(String(c)));
    return n;
  }

  function _isoFromYmdHHMM(ymd, hhmm) {
    // Interpret as local time (Europe/London UX), then convert to ISO
    const [H, M] = String(hhmm || "00:00").split(":").map((x) => parseInt(x, 10));
    const [y, m, d] = String(ymd).split("-").map((x) => parseInt(x, 10));
    const dt = new Date(y, m - 1, d, H, M, 0, 0);
    return dt.toISOString();
  }

  function _minutesBetween(isoA, isoB) {
    const a = new Date(isoA).getTime();
    const b = new Date(isoB).getTime();
    return Math.round((b - a) / 60000);
  }

  function _todayHHMM() {
    const d = new Date();
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${hh}${mm}`;
  }

  function _buildSignaturePad(width = 300, height = 120) {
    const canvas = _el("canvas", { width, height, style: { width: "100%", height: "120px", border: "1px solid #ccc", borderRadius: "8px", background: "#fff" } });
    const ctx = canvas.getContext("2d");
    ctx.lineWidth = 2;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#111";
    let drawing = false;
    let last = null;

    function pos(e) {
      const rect = canvas.getBoundingClientRect();
      const x = (e.touches ? e.touches[0].clientX : e.clientX) - rect.left;
      const y = (e.touches ? e.touches[0].clientY : e.clientY) - rect.top;
      return { x: (x * canvas.width) / rect.width, y: (y * canvas.height) / rect.height };
    }

    function start(e) {
      drawing = true;
      last = pos(e);
      e.preventDefault();
    }
    function move(e) {
      if (!drawing) return;
      const p = pos(e);
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
      last = p;
      e.preventDefault();
    }
    function end(e) {
      drawing = false;
      e.preventDefault();
    }

    canvas.addEventListener("mousedown", start);
    canvas.addEventListener("mousemove", move);
    window.addEventListener("mouseup", end);
    canvas.addEventListener("touchstart", start, { passive: false });
    canvas.addEventListener("touchmove", move, { passive: false });
    canvas.addEventListener("touchend", end);

    return {
      el: canvas,
      isEmpty: () => {
        // crude: sample a few pixels
        const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        for (let i = 3; i < data.length; i += 4) if (data[i] !== 0) return false;
        return true;
      },
      toPNGBlob: () =>
        new Promise((resolve) => canvas.toBlob((b) => resolve(b), "image/png")),
      clear: () => ctx.clearRect(0, 0, canvas.width, canvas.height),
    };
  }
function _mountOverlay(titleText) {
  const scrim = _el("div", {
    class: "ts-overlay",
    style: {
      position: "fixed",
      inset: "0",
      background: "rgba(0,0,0,0.5)",
      zIndex: "9999",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
    },
  });

  const card = _el(
    "div",
    {
      style: {
        width: "min(600px, 94vw)",
        maxHeight: "90vh",
        overflow: "auto",
        background: "#fff",
        borderRadius: "12px",
        boxShadow: "0 10px 30px rgba(0,0,0,.2)",
        padding: "16px",
      },
    },
    _el("div", { style: { fontSize: "18px", fontWeight: "700", marginBottom: "8px" } }, titleText),
    _el("div", { class: "ts-contents" })
  );

  // Close-on-scrim click; if success path marked this overlay, refresh afterwards
  scrim.addEventListener("click", (e) => {
    if (e.target === scrim) {
      const shouldRefresh = scrim.dataset.tsRefreshOnClose === "1";
      document.body.removeChild(scrim);
      if (shouldRefresh) {
        try { if (typeof startTilesThenEmergencyChain === "function") startTilesThenEmergencyChain(); } catch (_) {}
      }
    }
  });

  card.addEventListener("click", (e) => e.stopPropagation());

  scrim.appendChild(card);
  document.body.appendChild(scrim);

  return { root: scrim, card, contents: card.querySelector(".ts-contents") };
}

  async function startTimesheetWizard(tile, identity, baseline) {
    const ov = _mountOverlay("Timesheet");

    // Shared state
    const state = {
      tile,
      ymd: tile.ymd || tile.date_start_local, // YYYY-MM-DD
      hospital: tile.hospital,
      ward: tile.ward,
      job_title: tile.job_title || tile.role || tile.jobTitle,
      shift_label: tile.shift_label || tile.shift || tile.shiftLabel || "",
      scheduled_start_hhmm: tile.shiftStartHHMM || "",
      scheduled_end_hhmm: tile.shiftEndHHMM || "",
      worked_start_hhmm: "",
      worked_end_hhmm: "",
      break_start_hhmm: "",
      break_end_hhmm: "",
      nurseSig: null,
      authSig: null,
      auth_name: "",
      auth_job_title: "",
      presign: null,
    };

    // Step 1 — Hours & Break
    const step1 = () => {
      const c = ov.contents;
      c.innerHTML = "";

      const form = _el("div", { style: { display: "grid", gap: "12px" } });
      form.append(
        _el("div", {}, `Hospital: ${state.hospital}`),
        _el("div", {}, `Ward: ${state.ward}`),
        _el("div", {}, `Job title: ${state.job_title}`),
        _el("label", {}, "Worked Start (HH:MM) ", _el("input", { type: "time", value: state.worked_start_hhmm, oninput: (e) => (state.worked_start_hhmm = e.target.value) })),
        _el("label", {}, "Worked End (HH:MM) ", _el("input", { type: "time", value: state.worked_end_hhmm, oninput: (e) => (state.worked_end_hhmm = e.target.value) })),
        _el("label", {}, "Break Start (HH:MM) ", _el("input", { type: "time", value: state.break_start_hhmm, oninput: (e) => (state.break_start_hhmm = e.target.value) })),
        _el("label", {}, "Break End (HH:MM) ", _el("input", { type: "time", value: state.break_end_hhmm, oninput: (e) => (state.break_end_hhmm = e.target.value) }))
      );

      const warn = _el("div", { style: { color: "#a15c00", display: "none" } }, "Warning: Break is not exactly 60 minutes.");
      const btns = _el(
        "div",
        { style: { display: "flex", gap: "8px", marginTop: "12px", justifyContent: "flex-end" } },
        _el("button", { type: "button", onclick: () => document.body.removeChild(ov.root) }, "Cancel"),
        _el("button", {
          type: "button",
          onclick: () => {
            // validate
            if (!state.worked_start_hhmm || !state.worked_end_hhmm || !state.break_start_hhmm || !state.break_end_hhmm) {
              alert("Please complete all times.");
              return;
            }
            const ws = _isoFromYmdHHMM(state.ymd, state.worked_start_hhmm);
            const we = _isoFromYmdHHMM(state.ymd, state.worked_end_hhmm);
            const bs = _isoFromYmdHHMM(state.ymd, state.break_start_hhmm);
            const be = _isoFromYmdHHMM(state.ymd, state.break_end_hhmm);
            const bm = _minutesBetween(bs, be);
            warn.style.display = bm === 60 ? "none" : "block";
            step2();
          },
        }, "Next")
      );

      c.append(form, warn, btns);
    };

    // Step 2 — Nurse signature
    const step2 = () => {
      ov.contents.innerHTML = "";
      const pad = _buildSignaturePad();
      const timeNow = _todayHHMM();

      ov.contents.append(
        _el("div", { style: { marginBottom: "6px" } }, "Nurse signature"),
        pad.el,
        _el("div", { style: { marginTop: "6px", color: "#666" } }, `Date/time now: ${timeNow}`),
        _el(
          "div",
          { style: { display: "flex", gap: "8px", marginTop: "12px", justifyContent: "flex-end" } },
          _el("button", { type: "button", onclick: step1 }, "Back"),
          _el("button", {
            type: "button",
            onclick: async () => {
              if (pad.isEmpty()) {
                alert("Please sign before continuing.");
                return;
              }
              state.nurseSig = await pad.toPNGBlob();
              step3();
            },
          }, "Next")
        )
      );
    };

    // Step 3 — Ward Authorisation (details + signature + submit)
    const step3 = () => {
      ov.contents.innerHTML = "";

      const wsISO = _isoFromYmdHHMM(state.ymd, state.worked_start_hhmm);
      const weISO = _isoFromYmdHHMM(state.ymd, state.worked_end_hhmm);
      const bsISO = _isoFromYmdHHMM(state.ymd, state.break_start_hhmm);
      const beISO = _isoFromYmdHHMM(state.ymd, state.break_end_hhmm);
      const bm = _minutesBetween(bsISO, beISO);

      const nameIn = _el("input", { type: "text", placeholder: "Authoriser full name", oninput: (e) => (state.auth_name = e.target.value) });
      const titleIn = _el("input", { type: "text", placeholder: "Authoriser job title", oninput: (e) => (state.auth_job_title = e.target.value) });

      const pad = _buildSignaturePad();

      ov.contents.append(
        _el("div", { style: { marginBottom: "8px", fontWeight: "600" } }, "Ward Authorisation — Signing Timesheet"),
        _el("div", {}, `Nurse: ${deriveOccupantKey(baseline) || "—"}`),
        _el("div", {}, `Date: ${state.ymd}`),
        _el("div", {}, `Hospital: ${state.hospital}`),
        _el("div", {}, `Ward: ${state.ward}`),
        _el("div", {}, `Job title: ${state.job_title}`),
        _el("div", {}, `Start: ${state.worked_start_hhmm}`),
        _el("div", {}, `End: ${state.worked_end_hhmm}`),
        _el("div", {}, `Break Start: ${state.break_start_hhmm}`),
        _el("div", {}, `Break End: ${state.break_end_hhmm}`),
        _el("div", {}, `Break length: ${bm} mins`),
        _el("hr"),
        _el("div", { style: { marginTop: "6px", marginBottom: "6px" } }, "Please enter your full name and job title"),
        nameIn,
        titleIn,
        _el("div", { style: { marginTop: "12px" } }, "Please sign to confirm hours worked are accurate and that you are authorised to sign this timesheet"),
        pad.el,
        _el(
          "div",
          { style: { display: "flex", gap: "8px", marginTop: "12px", justifyContent: "space-between" } },
          _el("div", {}, _el("button", { type: "button", onclick: step2 }, "Back")),
          _el("div", {},
            _el("button", { type: "button", onclick: () => document.body.removeChild(ov.root) }, "Cancel"),
            _el("button", {
              type: "button",
              style: { marginLeft: "8px" },
              onclick: async () => {
                if (!state.auth_name || !state.auth_job_title) {
                  alert("Please enter authoriser name and job title.");
                  return;
                }
                if (pad.isEmpty()) {
                  alert("Authoriser signature is required.");
                  return;
                }

                // 1) Presign
                const pres = await presignTimesheet({
                  occupant_key: deriveOccupantKey(baseline),
                  date_start_local: state.ymd,
                  hospital: state.hospital,
                  ward: state.ward,
                  job_title: state.job_title,
                  shift_label: state.shift_label,
                  scheduled_start_hhmm: state.scheduled_start_hhmm,
                  scheduled_end_hhmm: state.scheduled_end_hhmm,
                });

                if (!pres.ok) {
                  _debugOrToast("Presign failed", pres);
                  return;
                }
                state.presign = pres.json;

                // 2) Upload signatures
                const nurseUp = await uploadSignaturePutUrl(state.presign.upload.nurse.put_url, state.nurseSig);
                if (!nurseUp.ok) {
                  _debugOrToast("Nurse signature upload failed", nurseUp);
                  return;
                }
                const authBlob = await pad.toPNGBlob();
                const authUp = await uploadSignaturePutUrl(state.presign.upload.authoriser.put_url, authBlob);
                if (!authUp.ok) {
                  _debugOrToast("Authoriser signature upload failed", authUp);
                  return;
                }

                // 3) Submit
                const payload = {
                  booking_id: state.presign.booking_id,
                  scheduled_start_iso: state.scheduled_start_hhmm ? _isoFromYmdHHMM(state.ymd, state.scheduled_start_hhmm) : null,
                  scheduled_end_iso: state.scheduled_end_hhmm ? _isoFromYmdHHMM(state.ymd, state.scheduled_end_hhmm) : null,
                  worked_start_iso: wsISO,
                  worked_end_iso: weISO,
                  break_start_iso: bsISO,
                  break_end_iso: beISO,
                  auth_name: state.auth_name,
                  auth_job_title: state.auth_job_title,
                  nurse_key: state.presign.upload.nurse.key,
                  authoriser_key: state.presign.upload.authoriser.key,
                  idempotency_key: crypto.randomUUID(),
                  client_user_agent: navigator.userAgent,
                  client_timestamp: new Date().toISOString(),
                  occupant_key: deriveOccupantKey(baseline),
                  hospital: state.hospital,
                  ward: state.ward,
                  job_title: state.job_title,
                  shift_label: state.shift_label,
                };

                // Ensure a small status line in the overlay (under the button)
const statusEl = ensureTsStatusBox(ov.card);

// Disable button while submitting / retrying
btn.disabled = true;
btn.textContent = "Submitting…";

// Reuse the same idempotency key for all retries
if (!payload.idempotency_key) payload.idempotency_key = crypto.randomUUID();

const { ok, sub, cancelled } = await submitTimesheetWithRetry(payload, {
  maxAttempts: 5,
  updateStatus: (attempt, max, phase) => {
    // phase: "submit" | "retry" | "waiting"
    if (phase === "submit") {
      setTsStatus(statusEl, `Submitting… attempt ${attempt}/${max}`);
      btn.textContent = `Submitting… (${attempt}/${max})`;
    } else if (phase === "retry") {
      setTsStatus(statusEl, `Retrying… attempt ${attempt}/${max}`);
      btn.textContent = `Retrying… (${attempt}/${max})`;
    } else if (phase === "waiting") {
      setTsStatus(statusEl, `Retrying shortly… attempt ${attempt}/${max}`);
    }
  },
  overlayRoot: ov.root
});

// Re-enable the button (unless overlay has gone away)
if (document.body.contains(ov.root)) {
  btn.disabled = false;
  btn.textContent = "Authorise Timesheet";
}

if (!ok) {
  if (!cancelled) {
    setTsStatus(statusEl,
      "Please complete a paper timesheet as there are network problems. " +
      "If you need a timesheet emailed to you navigate to the menu on the Arthur Rai app and choose ‘Send a timesheet by email’."
    );
    _debugOrToast("Submit failed", sub);
  }
  return;
}

// Success: show big tick then close + refresh (handled in _successTick)
_successTick(ov.root);

              },
            }, "Authorise Timesheet")
          )
        )
      );
    };

    step1();
  }

  function _successTick(overlayRoot) {
  // Mark this overlay so closing it will trigger a server refresh
  try { if (overlayRoot) overlayRoot.dataset.tsRefreshOnClose = "1"; } catch (_) {}

  const pane = _el("div", {
    style: {
      position: "fixed",
      inset: "0",
      background: "rgba(255,255,255,0.95)",
      zIndex: "10000",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: "56px",
      color: "#0a7f43",
      fontWeight: "800",
    },
  }, "✓");

  document.body.appendChild(pane);

  setTimeout(() => {
    pane.remove();

    // Close the overlay (if still present)
    if (overlayRoot && overlayRoot.parentNode) {
      overlayRoot.parentNode.removeChild(overlayRoot);
    }

    // Refresh tiles from the server after successful authorisation & close
    try { if (typeof startTilesThenEmergencyChain === "function") startTilesThenEmergencyChain(); } catch (_) {}
  }, 5000);
}

function _debugOrToast(label, resp) {
  // Name-based test gate: subject can be baseline or identity
  const subject = window.baseline || window.identity || '';
  const test = !!CFG.TIMESHEET_TESTMODE && isTimesheetTestUser(subject);

  if (test) {
    const msg = `${label}\nstatus=${resp?.status}\nerror=${resp?.error || ""}\njson=${JSON.stringify(resp?.json || {}, null, 2)}`;
    alert(msg);
  } else if (window.showToast) {
    window.showToast("Something went wrong. Please try again.");
  } else {
    alert("Something went wrong. Please try again.");
  }
}


  /* ──────────────────────────────────────────────────────────────────────────
   * 8) Submitted Timesheet Modal (summary + revoke & resubmit)
   * ────────────────────────────────────────────────────────────────────────── */

  async function openSubmittedTimesheetModal(tile, identity, baseline) {
    // booking_id may be on the tile or computed:
    const booking_id =
      tile.booking_id ||
      (await makeBookingId({
        occupant_key: deriveOccupantKey(baseline),
        date_start_local: tile.ymd || tile.date_start_local,
        hospital: tile.hospital,
        ward: tile.ward,
        job_title: tile.job_title || tile.role || tile.jobTitle,
        shift_label: tile.shift_label || tile.shift || tile.shiftLabel || "",
      }));

    // Fetch details
    const resp = await getTimesheet(booking_id);
    let ts = null;
    if (resp.ok) ts = resp.json;

    const ov = _mountOverlay("Submitted Timesheet");
    const c = ov.contents;

    if (!ts) {
      c.append(_el("div", {}, "No details available for this timesheet."));
    } else {
      const lines = [
        `Hospital: ${ts.hospital || tile.hospital}`,
        `Ward: ${ts.ward || tile.ward}`,
        `Job title: ${ts.job_title || tile.job_title || tile.role || tile.jobTitle}`,
        `Worked Start: ${_fmtTime(ts.worked_start_iso)}`,
        `Worked End: ${_fmtTime(ts.worked_end_iso)}`,
        `Break Start: ${_fmtTime(ts.break_start_iso)}`,
        `Break End: ${_fmtTime(ts.break_end_iso)}`,
        `Break length: ${ts.break_minutes ?? "—"} mins`,
        `Authoriser: ${ts.authoriser_name || "—"} (${ts.authoriser_job_title || "—"})`,
        `Authorised at: ${_fmtDateTime(ts.authorised_at)}`
      ];
      c.append(_el("div", { style: { display: "grid", gap: "4px", marginBottom: "12px" } }, ...lines.map((t) => _el("div", {}, t))));
    }

    const btns = _el(
      "div",
      { style: { display: "flex", gap: "8px", justifyContent: "space-between", marginTop: "12px" } },
      _el("div", {}, _el("button", { onclick: () => document.body.removeChild(ov.root) }, "Close")),
      _el("div", {},
        _el("button", {
          style: { background: "#b21f1f", color: "#fff" },
          onclick: async () => {
            // Confirm revoke & presign
            const confirmMsg = "You have already submitted a timesheet for this shift. Do you want to revoke it and submit a new one?";
            if (!confirm(confirmMsg)) return;

            const body = {
              booking_id,
              reason: "user_resubmitted",
              occupant_key: deriveOccupantKey(baseline),
              date_start_local: tile.ymd || tile.date_start_local,
              hospital: tile.hospital,
              ward: tile.ward,
              job_title: tile.job_title || tile.role || tile.jobTitle,
              shift_label: tile.shift_label || tile.shift || tile.shiftLabel || "",
            };
            const rx = await revokeAndPresign(body);
            if (!rx.ok) {
              _debugOrToast("Revoke & presign failed", rx);
              return;
            }
            document.body.removeChild(ov.root);
            // Immediately start wizard on new presign (we’ll just pass the original tile;
            // the wizard will do a fresh presign anyway if it needs to refresh tokens).
            startTimesheetWizard(tile, identity, baseline);
          },
        }, "Revoke & Resubmit")
      )
    );

    c.append(btns);
  }

  function _fmtTime(iso) {
    if (!iso) return "—";
    const d = new Date(iso);
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${hh}:${mm}`;
  }
  function _fmtDateTime(iso) {
    if (!iso) return "—";
    const d = new Date(iso);
    return d.toLocaleString();
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 9) Outbox (optional stubs for offline queueing; safe to no-op)
   *    - You can wire these to your SW/IndexedDB later.
   * ────────────────────────────────────────────────────────────────────────── */
  async function queueSignatureUpload(/* item */) {
    // TODO: implement IDB queue if desired
    return true;
  }
  async function queueTimesheetSubmit(/* item */) {
    // TODO: implement IDB queue if desired
    return true;
  }
  async function flushTimesheetOutbox() {
    // TODO: pull from IDB and replay nurse → authoriser → submit
    return { uploads: 0, submits: 0 };
  }

  /* ──────────────────────────────────────────────────────────────────────────
   * 10) Public API
   * ────────────────────────────────────────────────────────────────────────── */
  return {
    // Config & gating
    normalizeMsisdn,
    isTimesheetFeatureEnabled,
    isTimesheetTestUser,
    deriveOccupantKey,

    // Booking ids
    makeBookingId,

    // Broker
    presignTimesheet,
    uploadSignaturePutUrl,
    submitTimesheet,
    getTimesheet,
    revokeTimesheet,
    revokeAndPresign,
    authorisedStatus,
    presignSignatureGet,

    // History
    augmentHistoryWithTimesheetStatus,

    // UI hooks
    decorateTileWithTSBadge,
    onTileClickTimesheetBranch,
    startTimesheetWizard,
    openSubmittedTimesheetModal,

    // Outbox stubs
    queueSignatureUpload,
    queueTimesheetSubmit,
    flushTimesheetOutbox,
  };
})();

// Simple DOM helpers to show status under the modal header/button
function ensureTsStatusBox(cardEl) {
  let el = cardEl.querySelector(".ts-status");
  if (!el) {
    el = document.createElement("div");
    el.className = "ts-status";
    el.style.margin = "8px 0 0 0";
    el.style.fontSize = "13px";
    el.style.color = "#444";
    cardEl.appendChild(el);
  }
  return el;
}
function setTsStatus(statusEl, text) {
  if (!statusEl) return;
  statusEl.textContent = String(text || "");
}

// Classify transient vs permanent (retryable) errors
function isTransientTimesheetError(result) {
  // result can be { ok:false, status, error } from submitTimesheet or a catch
  const s = Number(result && result.status);
  if (!isFinite(s) || s === 0) return true;     // network / fetch failure
  if (s === 408 || s === 429) return true;      // timeout / rate limited
  if (s >= 500 && s <= 504) return true;        // 5xx
  return false;                                  // treat others as permanent
}

// tiny delay
function delay(ms) { return new Promise(res => setTimeout(res, ms)); }

// Core retry wrapper (reuses idempotency_key)
async function submitTimesheetWithRetry(payload, { maxAttempts = 5, updateStatus, overlayRoot } = {}) {
  // idempotency_key must be stable across attempts
  if (!payload.idempotency_key) {
    payload.idempotency_key = (crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now());
  }

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    // If overlay has been closed, stop trying (user navigated away)
    if (!document.body.contains(overlayRoot)) {
      return { ok: false, cancelled: true };
    }

    updateStatus && updateStatus(attempt, maxAttempts, "submit");

    let res;
    try {
      res = await submitTimesheet(payload);
    } catch (err) {
      res = { ok: false, status: 0, error: err };
    }

    if (res && res.ok) return { ok: true, sub: res };

    // Decide whether to retry
    if (attempt >= maxAttempts || !isTransientTimesheetError(res)) {
      return { ok: false, sub: res };
    }

    // Backoff with jitter (1s, 2s, 4s, 8s … capped)
    const base = Math.min(15000, 1000 * Math.pow(2, attempt - 1));
    const jitter = Math.floor(Math.random() * 300);
    const waitMs = base + jitter;

    updateStatus && updateStatus(attempt + 1, maxAttempts, "waiting");
    await delay(waitMs);

    // Update UI state for next loop
    updateStatus && updateStatus(attempt + 1, maxAttempts, "retry");
  }

  // Should not reach here
  return { ok: false };
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
async function init() {
  // Rehydrate identity (do NOT open login here)
  identity = loadSavedIdentity();

  // Chrome + furniture
  ensureLoadingOverlay();
  ensureMenu();
  ensureEmergencyButton();

  // If we arrived on a password-reset link, open that flow (non-dismissable)
  if (new URLSearchParams(location.search).has('k')) {
    openResetOverlay();
  }

  // Restore any local draft
  loadDraft();

  // Try silent auto-login (Credentials API) if we don't have an identity.
  // Do NOT open login here; let loadFromServer() decide from the server.
  let autoOK = false;
  if (!identity.msisdn) {
    try { if (els.emergencyBtn) { els.emergencyBtn.disabled = true; els.emergencyBtn.hidden = false; } } catch {}
    try { autoOK = await tryAutoLoginViaCredentialsAPI(); } catch {}
  }

  try {
    if (!autoOK) {
      // Show loader only on true cold start (handled inside loadFromServer)
      await loadFromServer({ force: true });
    }
    // Kick emergency eligibility quietly in background
    try { refreshEmergencyEligibility({ silent: true }); } catch {}
  } catch (e) {
    // Only surface non-auth failures; auth failures show login inside loadFromServer
    if (String(e && e.message) !== '__AUTH_STOP__') {
      showToast('Load failed: ' + (e.message || e), 5000);
    }
  } finally {
    sizeGrid();
    equalizeTileHeights({ reset: true });

    // Ensure the Emergency button exists and is wired (idempotent)
    ensureEmergencyButton();
    try { typeof wireEmergencyButton === 'function' && wireEmergencyButton(); } catch {}

    ensurePastShiftsButton();
  }

  // If we’re signed in, scrub any leftover #/login|#/forgot|#/reset hash so a reload doesn’t reopen the modal.
  try {
    if (identity && identity.msisdn && /^#\/(login|forgot|reset)/.test(location.hash || '')) {
      history.replaceState(null, '', location.pathname + (location.search || ''));
    }
  } catch {}

  // Wire Refresh to the unified chain (no premature login)
  if (els.refreshBtn && !els.refreshBtn.__wired) {
    els.refreshBtn.__wired = true;
    els.refreshBtn.addEventListener('click', () => {
      startTilesThenEmergencyChain({ force: true }).catch(() => {});
    });
  }

  routeFromURL && routeFromURL();
}

window.addEventListener('hashchange', routeFromURL);

// ---------- Boot ----------
// ───────────────────────── CHANGED (IIFE) ─────────────────────────

// --- UPDATED ---
// --- UPDATED: init() ---
async function init() {
  // Rehydrate identity (do NOT open login here)
  identity = loadSavedIdentity();

  // Chrome + furniture
  ensureLoadingOverlay();
  ensureMenu();
  ensureEmergencyButton();

  // If we arrived on a password-reset link, open that flow (non-dismissable)
  if (new URLSearchParams(location.search).has('k')) {
    openResetOverlay();
  }

  // Restore any local draft
  loadDraft();

  // Try silent auto-login (Credentials API) if we don't have an identity.
  // Do NOT open login here; let loadFromServer() decide from the server.
  let autoOK = false;
  if (!identity.msisdn) {
    try { if (els.emergencyBtn) { els.emergencyBtn.disabled = true; els.emergencyBtn.hidden = false; } } catch {}
    try { autoOK = await tryAutoLoginViaCredentialsAPI(); } catch {}
  }

  try {
    if (!autoOK) {
      // Show loader only on true cold start (handled inside loadFromServer)
      await loadFromServer({ force: true });
    }
    // Kick emergency eligibility quietly in background
    try { refreshEmergencyEligibility({ silent: true }); } catch {}
  } catch (e) {
    // Only surface non-auth failures; auth failures show login inside loadFromServer
    if (String(e && e.message) !== '__AUTH_STOP__') {
      showToast('Load failed: ' + (e.message || e), 5000);
    }
  } finally {
    sizeGrid();
    equalizeTileHeights({ reset: true });

    // Ensure the Emergency button exists and is wired (idempotent)
    ensureEmergencyButton();
    try { typeof wireEmergencyButton === 'function' && wireEmergencyButton(); } catch {}

    ensurePastShiftsButton();
  }

  // If we’re signed in, scrub any leftover #/login|#/forgot|#/reset hash so a reload doesn’t reopen the modal.
  try {
    if (identity && identity.msisdn && /^#\/(login|forgot|reset)/.test(location.hash || '')) {
      history.replaceState(null, '', location.pathname + (location.search || ''));
    }
  } catch {}

  // Wire Refresh to the unified chain (no premature login)
  if (els.refreshBtn && !els.refreshBtn.__wired) {
    els.refreshBtn.__wired = true;
    els.refreshBtn.addEventListener('click', () => {
      startTilesThenEmergencyChain({ force: true }).catch(() => {});
    });
  }

  routeFromURL && routeFromURL();
}


// Ensure init runs exactly once on load.
if (document.readyState !== 'loading') {
  init();
} else {
  window.addEventListener('DOMContentLoaded', init, { once: true });
}
