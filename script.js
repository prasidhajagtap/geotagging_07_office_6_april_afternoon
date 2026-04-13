/*
 * ══════════════════════════════════════════════════════════════
 *  SEAMEX GEO-ATTENDANCE  |  script.js  |  v07
 *  Author : Prasidha Jagtap
 *  Role   : Assistant Manager – IT, Aditya Birla Group (Seamex)
 *  Office : Reliable Tech Park, Airoli, Maharashtra
 *
 *  Hello, future developer. 👋
 *  Maintained by Prasidha Jagtap. Test on real Android + iOS.
 *
 *  FUNCTION INDEX
 *  ─────────────────────────────────────────────────────────────
 *  INIT          DOMContentLoaded
 *  AMBIENT       setTimeOfDay()
 *  THEME         applyTheme() · cycleTheme() · isDay()
 *  CLOCK LOOP    startMainLoop() · stopMainLoop()
 *  SVG HANDS     tickHands() · startHandLoop() · stopHandLoop() · rot()
 *  TRANSITIONS   pageTransition()
 *  VALIDATION    isValidName · isValidId · isValidLoc · sanitize
 *  AUTH          setupLoginValidation()
 *  RENDER MAIN   renderMain()   ← session-restore entry point
 *  CLOCK IN      btn-ci listener
 *  CLOCK OUT     btn-co listener  ← GPS freeze fixed here (FIX-05)
 *  RENDER REVIEW renderReview()
 *  MODAL         showModal() · closeModal()
 *  FIX CLOCK-IN  promptFixIn()
 *  REDO OUT      promptRedoOut() · showRedoBanner()
 *  SUBMIT        btn-submit listener
 *  SUCCESS       renderSuccess()
 *  RECOVERY      markBusy() · clearBusy() · recoverState()
 *  GPS           getCoords()   ← Promise.race fix lives here (FIX-05)
 *  LOCATION CACHE getCached() · saveLoc() · renderChips() · setupAC() · renderDrop()
 *  TOAST         toast()
 *  HELPERS       g · setTx · hide · show · nowISO · rnd · save · isNewDay · fmt · msDur · duration
 *
 *  SECURITY NOTES (Prasidha)
 *  ─────────────────────────────────────────────────────────────
 *  · isValidName/isValidId/isValidLoc — strict regex on every input.
 *    Same pattern enforced on clock-in AND clock-out. Re-checked on click.
 *  · sanitize() — strips SQL/XSS chars before every DB write:
 *    < > " ' ` ; % ( ) { } \ / = + | # @
 *  · All DOM writes via .textContent — never .innerHTML (XSS).
 *  · Supabase anon key is public. Safe ONLY with RLS enabled:
 *    INSERT-only for anon. Deny SELECT/UPDATE/DELETE.
 *  · service_role key NEVER on client. Ever.
 *  · GPS maximumAge:0 — always fresh, never cached position.
 *  · localStorage shape validated before trusting on restore.
 *  · isSubmitting flag — double-submit race guard.
 *
 *  BUG FIXES
 *  ─────────────────────────────────────────────────────────────
 *  FIX-01  renderMain() enforces section visibility on every call.
 *          Refresh no longer drops user to login screen.
 *  FIX-02  Single rAF loop drives header clock + shift timer.
 *          Zero lag, zero drift between the two displays.
 *  FIX-03  Redo flow fully resets U.clockOut, lastActionDate,
 *          btn text. rAF timer resumes automatically.
 *  FIX-04  Ghost screen fully hidden + co-grp restored on GPS fail.
 *  FIX-05  ★ GPS freeze on redo: getCoords() uses true Promise.race.
 *          btn-co has a guaranteed `finally` block that re-enables
 *          the button on EVERY outcome — success, error, or timeout.
 *  FIX-06  Recovery button: appears after 5s stuck, DOM-only restore.
 *
 *  MIGRATION CHECKLIST
 *  ─────────────────────────────────────────────────────────────
 *  [ ] Azure AD SSO: replace #auth-sec with MSAL redirect.
 *      Map: displayName→U.name | employeeId→U.id
 *  [ ] GitHub Pages→SharePoint SPFx: bundle supabase locally.
 *      Change submitted_via to 'sharepoint' in payload.
 *  [ ] DB→Production: swap SUPABASE_URL + SUPABASE_KEY only.
 *  [ ] Multi-office: uncomment branch_code in payload.
 * ══════════════════════════════════════════════════════════════
 */

/* ── SUPABASE ────────────────────────────────────────────────
   Prasidha: Anon key is safe ONLY while RLS is active on 'attendance'.
   Policy: INSERT for anon. SELECT/UPDATE/DELETE: DENY.
   Rotate key immediately if RLS is ever disabled.
────────────────────────────────────────────────────────────── */
const SUPABASE_URL = 'https://svhbqvcabbzrxvndxtjm.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN2aGJxdmNhYmJ6cnh2bmR4dGptIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUyMTA0MjksImV4cCI6MjA5MDc4NjQyOX0.lYIsM5zN4uGKbP79avcKR_EaAlP5tu2N688OgZI6wZA';
const _db = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

/* ── STATE ───────────────────────────────────────────────────
   Prasidha: U is the single source of truth.
   Every change calls save() to mirror to localStorage.
   Shape validated on restore — tampered data cannot bypass auth.
────────────────────────────────────────────────────────────── */
let U = {
  name: '', id: '',
  clockIn: null,  clockInCoords: '',  clockInLoc: '',
  clockOut: null, clockOutCoords: '', clockOutLoc: '',
  isClockedIn: false, submitted: false, lastActionDate: null
};

let rafId        = null;  // requestAnimationFrame loop ID
let handIv       = null;  // setInterval ID for auth clock hands
let isSubmitting = false; // double-submit guard
let inRedoMode   = false; // true when user is redoing clock-out
let stuckTimer   = null;  // setTimeout ID for stuck-button detector

/* ── STORAGE KEYS ────────────────────────────────────────────
   Using smx_v06 key for backward compatibility so active shifts
   from v06 survive the upgrade. Change key only on schema break.
────────────────────────────────────────────────────────────── */
const KEY_USER  = 'smx_v06';
const KEY_ID    = 'smx_lastid';
const KEY_THEME = 'smx_theme';
const KEY_LOCS  = 'smx_locs';
const MAX_LOCS  = 6;

/* ── HELPERS (hoisted) ───────────────────────────────────────
   Prasidha: Declared as `function` so they are fully hoisted.
   `const` arrow functions are NOT hoisted — using them above
   their declaration throws ReferenceError (temporal dead zone).
   Keep these as `function` declarations. Do not convert to const.
────────────────────────────────────────────────────────────── */

/** g — Prasidha: shorthand getElementById, used everywhere */
function g(id) { return document.getElementById(id); }

/** setTx — Prasidha: safely sets textContent, no-ops if element missing */
function setTx(id, v) { const e = g(id); if (e) e.textContent = v; }

/** hide — Prasidha: adds .hidden class (display:none!important in CSS) */
function hide(id) { g(id)?.classList.add('hidden'); }

/** show — Prasidha: removes .hidden class */
function show(id) { g(id)?.classList.remove('hidden'); }

/** nowISO — Prasidha: current timestamp as ISO 8601 string */
function nowISO() { return new Date().toISOString(); }

/** rnd — Prasidha: random element from array */
function rnd(arr) { return arr[Math.floor(Math.random() * arr.length)]; }

/** save — Prasidha: serializes U state to localStorage */
function save() { localStorage.setItem(KEY_USER, JSON.stringify(U)); }

/* ── CONTENT ARRAYS ──────────────────────────────────────────
   Prasidha: All UI copy here. Edit freely without touching logic.
────────────────────────────────────────────────────────────── */
const GREETINGS = [
  'Rare Sunday warrior spotted. Respect. 🦁',
  'New week, fresh resolve. Make it count! 🌟',
  'Tuesday energy — steady and purposeful. 💪',
  'Midweek momentum. You are in the thick of it. ⚡',
  'Thursday: the unsung hero of the week.',
  'Friday vibes. Finish line is right there. 🎉',
  'Closing the week strong. That is the Seamex way. 🚀'
];
const GHOST_NOTES = [
  'Your timings are being stored, safe and sound.',
  'Attendance logged with care — just like clockwork.',
  'Professional work happening behind the scenes.',
  'The database is receiving your day\'s hard work.',
  'Shift data is making its way in. Hang tight.'
];
const LOADER_MSGS = [
  'Starting your day…', 'One moment please…',
  'Getting things ready…', 'Just a tick…'
];

/* ══════════════════════════════════════════════════════════════
   INIT
   ══════════════════════════════════════════════════════════════
   Prasidha: Order matters.
   1. Ambient time-of-day (CSS attr, no runtime cost after this)
   2. Saved theme (prevents flash of wrong theme)
   3. Daily greeting (day-of-week from GREETINGS)
   4. Pre-fill last Poornata ID (reduce login friction)
   5. Session restore OR fresh auth setup
   6. setupAC called ONCE here — not on every renderMain call
      (calling it multiple times stacks document.click listeners)
*/
window.addEventListener('DOMContentLoaded', () => {
  setTimeOfDay();

  const savedTheme = localStorage.getItem(KEY_THEME) || 'auto';
  _themeMode = savedTheme;
  applyTheme(_themeMode);

  setTx('daily-greet', GREETINGS[new Date().getDay()]);

  const lastId = localStorage.getItem(KEY_ID);
  if (lastId) { const el = g('inp-id'); if (el) el.value = lastId; }

  /* Session restore — Prasidha: validate shape before trusting */
  try {
    const raw = localStorage.getItem(KEY_USER);
    if (raw) {
      const p = JSON.parse(raw);
      const valid =
        p?.name && typeof p.name === 'string' &&
        p?.id   && typeof p.id   === 'string' &&
        typeof p.isClockedIn === 'boolean' &&
        !isNewDay(p.lastActionDate);
      if (valid) {
        U = p;
        /* FIX-01: renderMain enforces section visibility on every call */
        renderMain();
        setupAC('in-loc',  'in-drop',  'in-chips');
        setupAC('out-loc', 'out-drop', 'out-chips');
        return;
      }
      localStorage.removeItem(KEY_USER);
    }
  } catch { localStorage.removeItem(KEY_USER); }

  /* Fresh session */
  setupLoginValidation();
  setupAC('in-loc',  'in-drop',  'in-chips');
  setupAC('out-loc', 'out-drop', 'out-chips');
});

/* ══════════════════════════════════════════════════════════════
   TIME-OF-DAY AMBIENT
   Prasidha: Sets data-tod on body. CSS handles all animation.
   dawn(5-9am) · day(9-15) · dusk(15-18) · night(18-5)
*/
function setTimeOfDay() {
  const h = new Date().getHours();
  const tod = h >= 5  && h < 9  ? 'dawn'
            : h >= 9  && h < 15 ? 'day'
            : h >= 15 && h < 18 ? 'dusk'
            : 'night';
  document.body.setAttribute('data-tod', tod);
}

/* ══════════════════════════════════════════════════════════════
   THEME SYSTEM
   Prasidha: Modes: 'auto' (time-based) | 'light' | 'dark'.
   Preference saved to localStorage. All .t-ico and .t-lbl
   elements across both footers updated together.
*/
let _themeMode = 'auto';

/** applyTheme — Prasidha: resolves effective theme and updates DOM */
function applyTheme(mode) {
  const eff = mode === 'auto' ? (isDay() ? 'light' : 'dark') : mode;
  document.documentElement.setAttribute('data-theme', eff);
  const ico = eff === 'dark' ? '☀️' : '🌙';
  const lbl = eff === 'dark' ? 'Light' : 'Dark';
  document.querySelectorAll('.t-ico').forEach(e => e.textContent = ico);
  document.querySelectorAll('.t-lbl').forEach(e => e.textContent = lbl);
}

/** cycleTheme — Prasidha: called by onclick on theme buttons */
function cycleTheme() {
  const next = { auto: 'light', light: 'dark', dark: 'auto' };
  _themeMode = next[_themeMode] || 'auto';
  localStorage.setItem(KEY_THEME, _themeMode);
  applyTheme(_themeMode);
}

/** isDay — Prasidha: true between 6am and 6pm */
const isDay = () => { const h = new Date().getHours(); return h >= 6 && h < 18; };

/* ══════════════════════════════════════════════════════════════
   MAIN RAF LOOP — FIX-02 (Prasidha)
   Single requestAnimationFrame loop drives BOTH the live header
   clock AND the shift timer from the same Date.now() call.
   Zero lag, zero drift. Runs only while main-sec is visible.
*/

/** startMainLoop — Prasidha: starts unified clock + timer rAF loop */
function startMainLoop() {
  stopMainLoop();
  let lastSec = -1;

  function loop() {
    const d  = new Date();
    const sc = d.getSeconds();

    if (sc !== lastSec) {
      lastSec = sc;
      const hh = d.getHours()  .toString().padStart(2, '0');
      const mm = d.getMinutes().toString().padStart(2, '0');
      const ss = sc.toString().padStart(2, '0');

      setTx('hdr-clock', `${hh}:${mm}:${ss}`);
      setTx('hdr-date', d.toLocaleDateString('en-IN', {
        weekday: 'short', day: 'numeric', month: 'short', year: 'numeric'
      }));

      /* Shift timer — only while actively clocked in (U.clockOut is null) */
      if (U.isClockedIn && U.clockIn && !U.clockOut && !U.submitted) {
        setTx('timer-val', msDur(Date.now() - new Date(U.clockIn).getTime()));
      }
    }
    rafId = requestAnimationFrame(loop);
  }
  rafId = requestAnimationFrame(loop);
}

/** stopMainLoop — Prasidha: cancels the rAF loop */
function stopMainLoop() {
  if (rafId) { cancelAnimationFrame(rafId); rafId = null; }
}

/* ══════════════════════════════════════════════════════════════
   SVG CLOCK HANDS
   Prasidha: JS sets rotation angle once/second.
   CSS cubic-bezier transition does the smooth easing (GPU-composited).
*/

/** tickHands — Prasidha: sets H/M/S SVG hand rotation angles */
function tickHands(hId, mId, sId) {
  const d  = new Date();
  const sc = d.getSeconds();
  const mn = d.getMinutes() + sc / 60;
  const hr = (d.getHours() % 12) + mn / 60;
  rot(hId, hr * 30);
  rot(mId, mn * 6);
  rot(sId, sc * 6);
}

/** startHandLoop — Prasidha: starts 1s interval for auth clock */
function startHandLoop(hId = 'ac-h', mId = 'ac-m', sId = 'ac-s') {
  tickHands(hId, mId, sId);
  handIv = setInterval(() => tickHands(hId, mId, sId), 1000);
}

/** stopHandLoop — Prasidha: clears auth clock interval */
function stopHandLoop() { clearInterval(handIv); handIv = null; }

/** rot — Prasidha: applies CSS rotation to SVG element */
function rot(id, deg) {
  const el = g(id);
  if (el) el.style.transform = `rotate(${deg}deg)`;
}

/* ══════════════════════════════════════════════════════════════
   PAGE TRANSITION — CRED-style clock overlay
   Prasidha: Full-screen overlay with ticking clock + floating
   time emojis between screens. No dependencies.
*/
const PT_MSGS   = ['Switching view…', 'One moment…', 'Loading…', 'Just a tick…'];
const PT_EMOJIS = ['⏱', '⌛', '🕐', '⏳', '🕑', '⏰'];

/** pageTransition — Prasidha: shows animated overlay, runs fn() mid-fade */
function pageTransition(fn, msg) {
  const ov   = g('pg-tr');
  const msgEl = g('pt-msg');
  const parts = g('pt-parts');
  if (!ov) { if (fn) fn(); return; }

  if (msgEl) msgEl.textContent = msg || rnd(PT_MSGS);

  /* Spawn floating time emojis */
  if (parts) {
    parts.innerHTML = '';
    [...PT_EMOJIS].sort(() => Math.random() - .5).slice(0, 4).forEach((em, i) => {
      const p = document.createElement('div');
      p.className = 'pt-p';
      p.textContent = em;
      p.style.cssText = `left:${12 + i * 22}%;bottom:${8 + Math.random() * 18}px;animation-delay:${i * .14}s`;
      parts.appendChild(p);
    });
  }

  ov.classList.add('on');
  setTimeout(() => { if (fn) fn(); }, 260);
  setTimeout(() => ov.classList.remove('on'), 960);
}

/* ══════════════════════════════════════════════════════════════
   INPUT VALIDATION — Prasidha
   SECURITY: all three validators enforced consistently on every
   input, every page, re-checked on every button click.

   isValidName: letters + spaces, min 2 chars.
   isValidId  : digits only, 3–12 chars.
   isValidLoc : letters/digits/spaces/hyphens ONLY.
                Blocks all SQL-injection and HTML-injection chars.
   sanitize   : strips dangerous chars before any DB write.
                Supabase parameterized queries are the baseline.
                This is defence-in-depth.
*/
const isValidName = s => /^[a-zA-Z\s]{2,60}$/.test(s.trim());
const isValidId   = s => /^[0-9]{3,12}$/.test(s.trim());

/** isValidLoc — Prasidha: BLOCKS < > " ' ; = + # | \ / ( ) { } % @ ` */
const isValidLoc  = s => /^[a-zA-Z0-9 \-]{2,60}$/.test(s.trim());

/** sanitize — Prasidha: final strip before any DB write */
const sanitize = s =>
  s.trim().replace(/[<>"'`%;(){}\\\/=+|#@]/g, '').slice(0, 60);

/** showErrIf — Prasidha: helper to toggle inline error visibility */
const showErrIf = (id, cond) => {
  const el = g(id);
  if (el) el.style.display = cond ? 'block' : 'none';
};

/** validateLoc — Prasidha: validates location input, shows error, returns bool */
function validateLoc(raw, errId) {
  if (!raw.trim()) { toast('Please enter a location name.'); return false; }
  if (!isValidLoc(raw)) {
    showErrIf(errId, true);
    toast('Use letters, numbers, spaces or hyphens only.', 'err'); return false;
  }
  showErrIf(errId, false);
  return true;
}

/* ══════════════════════════════════════════════════════════════
   LOGIN VALIDATION
   Prasidha: Enables Start Day only when both fields pass.
*/
function setupLoginValidation() {
  const nIn = g('inp-name'), iIn = g('inp-id'), btn = g('btn-start');
  if (!nIn || !iIn || !btn) return;

  const check = () => {
    const n = nIn.value.trim(), i = iIn.value.trim();
    showErrIf('err-name', n && !isValidName(n));
    showErrIf('err-id',   i && !isValidId(i));
    btn.disabled = !(isValidName(n) && isValidId(i));
  };
  nIn.addEventListener('input', check);
  iIn.addEventListener('input', check);
}

/* ══════════════════════════════════════════════════════════════
   START DAY
   Prasidha: Re-validates on click (not just on input events).
*/
g('btn-start').addEventListener('click', () => {
  const name = g('inp-name').value.trim();
  const id   = g('inp-id').value.trim();

  /* SECURITY: Re-validate on click (Prasidha) */
  if (!isValidName(name) || !isValidId(id)) {
    toast('Check your details.', 'err'); return;
  }

  localStorage.setItem(KEY_ID, id);
  U.name = name; U.id = id;
  U.lastActionDate = nowISO(); save();

  show('inline-loader');
  const startBtn = g('btn-start');
  if (startBtn) startBtn.classList.add('hidden');
  setTx('ld-msg', rnd(LOADER_MSGS));
  startHandLoop('ac-h', 'ac-m', 'ac-s');

  setTimeout(() => {
    stopHandLoop();
    hide('inline-loader');
    if (startBtn) startBtn.classList.remove('hidden');
    pageTransition(() => renderMain(), 'Starting your day…');
  }, 1400);
});

/* ══════════════════════════════════════════════════════════════
   RENDER MAIN — central state router
   FIX-01 (Prasidha): enforces section visibility at every call.
   Called from btn-start AND session restore. Both paths show
   main-sec and hide auth-sec correctly.
*/
function renderMain() {
  /* FIX-01: Always enforce correct section visibility */
  hide('auth-sec');
  show('main-sec');

  setTx('disp-name', U.name);
  setTx('disp-id',   U.id);
  startMainLoop();

  if (U.submitted) { renderSuccess(); return; }

  if (U.isClockedIn) {
    hide('ci-grp');
    show('status-card');
    setTx('ci-time', fmt(U.clockIn));
    setTx('ci-loc',  U.clockInLoc);

    if (!U.clockOut) {
      show('co-grp');
      if (inRedoMode) showRedoBanner(null, null);
      renderChips('out-loc', 'out-chips', 'out-drop');
    } else {
      renderReview();
    }
  } else {
    show('ci-grp');
    renderChips('in-loc', 'in-chips', 'in-drop');
  }
}

/* ══════════════════════════════════════════════════════════════
   CLOCK IN — btn-ci click handler
   Prasidha: validates, GPS, saves, transitions to co-grp.
*/
g('btn-ci').addEventListener('click', async () => {
  const raw = g('in-loc').value;
  if (!validateLoc(raw, 'err-in-loc')) return;

  const btn = g('btn-ci');
  btn.disabled = true; btn.textContent = 'Fetching location…';

  /* FIX-06: mark busy — shows recovery button after 5s */
  markBusy();

  let coords = null;
  try {
    coords = await getCoords();
  } catch (e) {
    console.error('[Seamex|Prasidha] GPS exception:', e);
  } finally {
    /* GUARANTEED re-enable on every outcome (Prasidha: FIX-05) */
    clearBusy();
    btn.disabled = false; btn.textContent = 'Clock In';
  }

  if (!coords) return;

  const loc = sanitize(raw);
  saveLoc(loc);
  U.clockIn = nowISO(); U.clockInCoords = coords;
  U.clockInLoc = loc; U.isClockedIn = true;
  U.lastActionDate = nowISO(); save();

  show('status-card');
  setTx('ci-time', fmt(U.clockIn));
  setTx('ci-loc',  U.clockInLoc);

  pageTransition(() => {
    hide('ci-grp');
    show('co-grp');
    renderChips('out-loc', 'out-chips', 'out-drop');
  }, 'Clocked in! Shift started.');

  toast('Clocked in! Have a great shift.', 'ok');
});

/* ══════════════════════════════════════════════════════════════
   CLOCK OUT — btn-co click handler

   ★ FIX-05 (Prasidha) — THE FREEZE FIX ★
   Root cause of redo freeze:
     On iOS Safari, calling getCurrentPosition a second time in the
     same page session (the redo case) can cause the browser to
     silently stall — firing neither the success nor the error
     callback. The previous v06 code had the button re-enable ONLY
     inside the `if(!coords)` block, which is reached only if
     getCoords() returns. If getCoords() never returned, the button
     stayed frozen forever.

   Two-part fix applied here:
     1. getCoords() uses true Promise.race — two separate Promises.
        The safety promise resolves null at 9s guaranteed.
        Whichever resolves first wins.
     2. btn re-enable is in a `finally` block. This runs on EVERY
        outcome: GPS success, GPS error, GPS timeout, or any thrown
        exception. The button is ALWAYS re-enabled after this point.

   FIX-04 (from v06, retained): ghost fully hidden + co-grp restored
   on GPS fail. rAF timer continues — U.clockOut is still null.
*/
g('btn-co').addEventListener('click', async () => {
  const raw = g('out-loc').value;
  if (!validateLoc(raw, 'err-out-loc')) return;

  const btn = g('btn-co');
  btn.disabled = true; btn.textContent = 'Fetching location…';

  /* Show ghost, hide co-grp. Timer keeps running — U.clockOut is null */
  setTx('ghost-note', rnd(GHOST_NOTES));
  show('ghost-scr');
  hide('co-grp');

  /* FIX-06: mark busy — shows recovery button after 5s */
  markBusy();

  let coords = null;
  try {
    coords = await getCoords();
  } catch (e) {
    console.error('[Seamex|Prasidha] GPS exception:', e);
  } finally {
    /*
     * ★ FIX-05 CRITICAL ★ (Prasidha)
     * This finally block runs on EVERY code path:
     *   - GPS success → finally runs → then success path below
     *   - GPS error   → finally runs → coords is null → restore UI
     *   - GPS timeout → finally runs → coords is null → restore UI
     *   - JS exception→ finally runs → coords is null → restore UI
     * In all cases the button is re-enabled. The freeze is impossible.
     */
    clearBusy();
    btn.disabled = false; btn.textContent = 'Clock Out';
  }

  /* GPS failed — restore UI. Timer continues via rAF loop. */
  if (!coords) {
    hide('ghost-scr');
    show('co-grp');
    return;
  }

  /* GPS success */
  const loc = sanitize(raw);
  saveLoc(loc);
  U.clockOut = nowISO(); U.clockOutCoords = coords;
  U.clockOutLoc = loc; U.lastActionDate = nowISO();
  inRedoMode = false;
  hide('redo-banner');
  save();

  setTimeout(() => {
    hide('ghost-scr');
    pageTransition(() => renderReview(), 'Shift logged. Review below.');
    toast('Clocked out! Review your shift.', 'ok');
  }, 1500);
});

/* ══════════════════════════════════════════════════════════════
   RENDER REVIEW — Prasidha
   Shows the review card. Freezes timer at actual shift duration.
*/
function renderReview() {
  hide('co-grp'); hide('ghost-scr'); hide('ci-grp');
  setTx('timer-val', duration(U.clockIn, U.clockOut));
  show('status-card');
  setTx('ci-time', fmt(U.clockIn));
  setTx('ci-loc',  U.clockInLoc);
  setTx('r-in-t',  fmt(U.clockIn));
  setTx('r-in-l',  U.clockInLoc);
  setTx('r-out-t', fmt(U.clockOut));
  setTx('r-out-l', U.clockOutLoc);
  setTx('r-dur',   duration(U.clockIn, U.clockOut));
  show('review-wrap');
}

/* ══════════════════════════════════════════════════════════════
   MODAL SYSTEM — Prasidha
   Generic modal. All content via textContent (XSS safe).
   Buttons built with createElement.
*/

/**
 * showModal — Prasidha
 * @param {object} opts  { icon, title, body, buttons[], showEditIn }
 * buttons: [{ label, cls, fn }]
 */
function showModal({ icon, title, body, buttons, showEditIn = false }) {
  setTx('m-icon',  icon  || 'ℹ️');
  setTx('m-title', title || '');
  setTx('m-body',  body  || '');

  if (showEditIn) {
    show('m-edit-in');
    const inp = g('edit-in-inp');
    if (inp) { inp.value = U.clockInLoc; setTimeout(() => inp.focus(), 200); }
  } else {
    hide('m-edit-in');
  }

  const btnsEl = g('m-btns');
  btnsEl.innerHTML = '';
  buttons.forEach(b => {
    const btn = document.createElement('button');
    btn.className = `btn-full ${b.cls || 'btn-red'}`;
    btn.style.cssText = 'margin:0;flex:1';
    btn.textContent = b.label; /* SECURITY: textContent (Prasidha) */
    btn.addEventListener('click', () => { closeModal(); if (b.fn) b.fn(); });
    btnsEl.appendChild(btn);
  });

  g('modal-ov').classList.add('open');
}

/** closeModal — Prasidha */
function closeModal() { g('modal-ov').classList.remove('open'); }

g('modal-ov').addEventListener('click', e => {
  if (e.target === g('modal-ov')) closeModal();
});

/* ══════════════════════════════════════════════════════════════
   FIX CLOCK-IN — promptFixIn
   Prasidha: Clock-in GPS timestamp is sealed for audit integrity.
   Only location NAME is editable here.
   Time corrections require HR to update server-side.
*/
function promptFixIn() {
  showModal({
    icon: '✏️', title: 'Correct Clock-In Location',
    body: 'Only the location name can be updated here. To correct your clock-in time, please contact your HR team.',
    showEditIn: true,
    buttons: [
      {
        label: 'Update Location', cls: 'btn-ora',
        fn: () => {
          const val = g('edit-in-inp')?.value.trim() || '';
          if (!isValidLoc(val)) {
            showErrIf('edit-in-err', true);
            toast('Enter a valid location name.', 'err'); return;
          }
          showErrIf('edit-in-err', false);
          U.clockInLoc = sanitize(val); save();
          renderReview();
          toast('Clock-in location updated.', 'ok');
        }
      },
      { label: 'Cancel', cls: 'btn-edit', fn: null }
    ]
  });
}

/* ══════════════════════════════════════════════════════════════
   REDO CLOCK-OUT — promptRedoOut + showRedoBanner
   FIX-03 (Prasidha): Full state reset — U.clockOut=null,
   lastActionDate updated, btn text reset.
   rAF loop resumes automatically because U.clockOut is null.
*/
function promptRedoOut() {
  const prevTime = fmt(U.clockOut);
  const prevLoc  = U.clockOutLoc;

  showModal({
    icon: '↩', title: 'Redo Clock-Out?',
    body: `Your current clock-out (${prevTime} at ${prevLoc}) will be permanently overridden. Continue?`,
    buttons: [
      {
        label: 'Yes, Redo', cls: 'btn-cri',
        fn: () => {
          /* FIX-03: Full state reset (Prasidha) */
          const savedT = U.clockOut, savedL = U.clockOutLoc;
          U.clockOut = null; U.clockOutCoords = ''; U.clockOutLoc = '';
          U.lastActionDate = nowISO();
          inRedoMode = true; save();

          /* Reset btn-co text and state explicitly */
          const btn = g('btn-co');
          if (btn) { btn.disabled = false; btn.textContent = 'Clock Out'; }

          /* Reset out-loc input */
          const outInp = g('out-loc');
          if (outInp) outInp.value = '';
          showErrIf('err-out-loc', false);

          /* Restore status card with live timer */
          show('status-card');
          setTx('ci-time', fmt(U.clockIn));
          setTx('ci-loc',  U.clockInLoc);
          /* rAF loop is running — U.clockOut is null so timer resumes */
          setTx('timer-val', msDur(Date.now() - new Date(U.clockIn).getTime()));

          showRedoBanner(savedT, savedL);

          hide('review-wrap');
          show('co-grp');
          renderChips('out-loc', 'out-chips', 'out-drop');

          toast('Enter your new clock-out details. ↩');
        }
      },
      { label: 'Cancel', cls: 'btn-edit', fn: null }
    ]
  });
}

/**
 * showRedoBanner — Prasidha
 * Shows the redo context banner with previous clock-out details.
 * @param {string|null} prevT - Previous clock-out ISO timestamp
 * @param {string|null} prevL - Previous clock-out location
 */
function showRedoBanner(prevT, prevL) {
  const info = g('redo-prev');
  if (info && prevT && prevL) {
    info.textContent = `Previously: ${fmt(prevT)} at ${prevL}`;
  } else if (info) {
    info.textContent = '';
  }
  show('redo-banner');
}

/* ══════════════════════════════════════════════════════════════
   SUBMIT DAY
   Prasidha: isSubmitting guards against double-submit.
   Final sanitize() pass before DB write (defence-in-depth).
   SECURITY: raw Supabase error never shown to user.
*/
g('btn-submit').addEventListener('click', async () => {
  if (isSubmitting) return;
  isSubmitting = true;

  const btn = g('btn-submit');
  btn.disabled = true; btn.textContent = 'Submitting…';
  hide('review-wrap');
  setTx('ghost-note', "Almost there — submitting your day's work…");
  show('ghost-scr');

  /*
   * SECURITY: Final sanitization before DB write (Prasidha).
   * Supabase uses parameterized queries server-side.
   * Client sanitize() is defence-in-depth.
   * MIGRATION NOTE: Uncomment branch_code + app_version before go-live.
   */
  const payload = {
    user_name:               sanitize(U.name),
    employee_id:             U.id.replace(/\D/g, ''), /* strip non-digits */
    clock_in_time:           U.clockIn,
    clock_in_coords:         U.clockInCoords,
    clock_in_location_name:  sanitize(U.clockInLoc),
    clock_out_time:          U.clockOut,
    clock_out_coords:        U.clockOutCoords,
    clock_out_location_name: sanitize(U.clockOutLoc),
    status:                  'completed'
    /* branch_code:   'HO_AIROLI', */
    /* app_version:   '07',        */
    /* submitted_via: 'web',       */
  };

  try {
    const { error } = await _db.from('attendance').insert([payload]);
    hide('ghost-scr');

    if (!error) {
      U.submitted = true; save();
      pageTransition(() => renderSuccess(), 'Shift submitted!');
      toast('All done! See you tomorrow.', 'ok');
    } else {
      /* SECURITY: never surface raw DB error (Prasidha) */
      console.error('[Seamex|Prasidha] Supabase error:', error);
      toast('Submission failed. Please try again.', 'err');
      isSubmitting = false;
      btn.disabled = false; btn.textContent = 'Submit the Day ✓';
      renderReview();
    }
  } catch (e) {
    hide('ghost-scr');
    console.error('[Seamex|Prasidha] Network error:', e);
    toast('Connection issue. Check network and retry.', 'err');
    isSubmitting = false;
    btn.disabled = false; btn.textContent = 'Submit the Day ✓';
    renderReview();
  }
});

/* ══════════════════════════════════════════════════════════════
   RENDER SUCCESS — Prasidha
   Visible all day. Stops rAF loop (clock not needed on success).
*/
function renderSuccess() {
  stopMainLoop();
  ['ci-grp','co-grp','review-wrap','ghost-scr','status-card'].forEach(hide);
  const today = new Date();
  setTx('ss-in-t',  fmt(U.clockIn));
  setTx('ss-in-l',  U.clockInLoc);
  setTx('ss-out-t', fmt(U.clockOut));
  setTx('ss-out-l', U.clockOutLoc);
  setTx('ss-dur',   duration(U.clockIn, U.clockOut));
  setTx('ss-date',  today.toLocaleDateString('en-IN', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  }));
  show('suc-card');
}

/* ══════════════════════════════════════════════════════════════
   STUCK BUTTON DETECTOR + RECOVERY — FIX-06 (Prasidha)

   markBusy(): call when a GPS-dependent button is disabled.
   After 5s, the recovery button (↺) appears next to theme toggle
   with a spinning animation and a tooltip.

   clearBusy(): called in every finally block after GPS resolves.
   Hides the recovery button and stops the spin.

   recoverState(): DOM-only re-render. URL unchanged. Session kept.
   Resets all stuck buttons and re-renders from current U state.
*/

/** markBusy — Prasidha: starts 5s timer for stuck-button detection */
function markBusy() {
  clearBusy();
  stuckTimer = setTimeout(() => {
    document.querySelectorAll('.rcv-btn').forEach(b => {
      b.classList.remove('hidden');
      b.classList.add('rcv-spin');
    });
    document.querySelectorAll('.rcv-tip').forEach(t => {
      t.classList.remove('hidden');
      setTimeout(() => t.classList.add('hidden'), 3500);
    });
  }, 5000);
}

/** clearBusy — Prasidha: cancels stuck timer and hides recovery button */
function clearBusy() {
  if (stuckTimer) { clearTimeout(stuckTimer); stuckTimer = null; }
  document.querySelectorAll('.rcv-btn').forEach(b => {
    b.classList.add('hidden');
    b.classList.remove('rcv-spin');
  });
  document.querySelectorAll('.rcv-tip').forEach(t => t.classList.add('hidden'));
}

/**
 * recoverState — Prasidha (FIX-06)
 * DOM-only recovery. Re-renders from U. URL unchanged. Session kept.
 * Called by the ↺ recovery button in both footers.
 */
function recoverState() {
  clearBusy();
  closeModal();

  /* Reset all potentially stuck buttons */
  [
    { id: 'btn-start',  txt: 'Start Day'        },
    { id: 'btn-ci',     txt: 'Clock In'         },
    { id: 'btn-co',     txt: 'Clock Out'        },
    { id: 'btn-submit', txt: 'Submit the Day ✓' }
  ].forEach(({ id, txt }) => {
    const b = g(id);
    if (b) { b.disabled = false; b.textContent = txt; }
  });

  hide('ghost-scr');
  hide('inline-loader');
  const sb = g('btn-start');
  if (sb) sb.classList.remove('hidden');

  /* Re-render the correct view from U state */
  if (!g('main-sec') || g('main-sec').classList.contains('hidden')) {
    /* Stuck on auth — just re-enable the form */
    setupLoginValidation();
  } else {
    /* Stuck on main — re-render panels from U */
    ['ci-grp','co-grp','ghost-scr','review-wrap','suc-card','status-card'].forEach(hide);

    if (U.submitted) { renderSuccess(); }
    else if (U.isClockedIn) {
      show('status-card');
      setTx('ci-time', fmt(U.clockIn));
      setTx('ci-loc',  U.clockInLoc);

      if (!U.clockOut) {
        inRedoMode = false; hide('redo-banner');
        show('co-grp');
        setTx('timer-val', msDur(Date.now() - new Date(U.clockIn).getTime()));
        renderChips('out-loc', 'out-chips', 'out-drop');
      } else {
        renderReview();
      }
    } else {
      show('ci-grp');
      renderChips('in-loc', 'in-chips', 'in-drop');
    }
  }

  toast('Page recovered. ↺', 'ok');
}

/* ══════════════════════════════════════════════════════════════
   GPS — getCoords()
   ★ FIX-05 (Prasidha) — TRUE Promise.race ★

   Two completely separate Promises:
     geoPromise    : resolves via getCurrentPosition callback.
     safetyPromise : resolves null after 9s guaranteed.
   Promise.race returns whichever settles first.

   This is different from the v06 approach which had ONE Promise
   with an internal timer — the timer could be bypassed in edge
   cases on iOS WebKit where the geo callback neither fires
   success nor error, leaving the function awaiting indefinitely.

   With Promise.race, the maximum wait is always 9s.
   Combined with the `finally` block in btn-co and btn-ci,
   the button is guaranteed to re-enable in under 9 seconds.
*/
/* ══════════════════════════════════════════════════════════════
   ANDROID WEBVIEW LOCATION MODAL — Prasidha
   Shown when GPS permission is denied inside an Android WebView.
   WebView does not expose browser settings to the user, so the
   standard "Allow it in your browser settings" instruction is
   not actionable. This modal gives two concrete options instead.
   
   HOST APP NOTE: The permanent fix requires the host app developer
   to implement WebChromeClient.onGeolocationPermissionsShowPrompt()
   — see the comment in getCoords() for the exact code needed.
*/
function showWebViewModal() {
  showModal({
    icon: '📍',
    title: 'Location Access Needed',
    body: 'This app is running inside another app which is blocking GPS access. To clock in or out, please open this page directly in Chrome browser on your Android device.',
    buttons: [
      {
        label: 'Open in Chrome',
        cls: 'btn-ora',
        fn: () => {
          /* Build a Chrome intent URL — opens the current page in Chrome */
          const url = window.location.href;
          /* Try Chrome intent first, fall back to just opening the URL */
          const intentUrl = `intent://${url.replace(/^https?:\/\//, '')}#Intent;scheme=https;package=com.android.chrome;end`;
          try {
            window.location.href = intentUrl;
          } catch (e) {
            /* Fallback: copy URL to clipboard and prompt user */
            if (navigator.clipboard) {
              navigator.clipboard.writeText(url).then(() => {
                toast('URL copied. Paste it in Chrome to open.', 'ok');
              });
            } else {
              toast('Please open this URL in Chrome browser.', 'ok');
            }
          }
        }
      },
      { label: 'Cancel', cls: 'btn-edit', fn: null }
    ]
  });
}


async function getCoords() {
  /*
   * getCoords — Prasidha (FIX-05 v3)
   *
   * ANDROID WEBVIEW FIX (v08):
   * When the app runs inside an Android WebView (e.g. embedded in the
   * Leadership Dashboard app), Android OS does NOT automatically forward
   * location permission to the WebView — unlike iOS WKWebView which inherits
   * it from the host app automatically.
   *
   * Root cause: error code 1 (PERMISSION_DENIED) fires immediately inside
   * Android WebView even when the user would normally allow it, because the
   * host app has not called WebChromeClient.onGeolocationPermissionsShowPrompt().
   *
   * Two-part fix applied here:
   *   1. Detect Android WebView from user agent string.
   *   2. On error code 1 inside WebView: show a specific actionable message
   *      directing the user to open in Chrome instead of just "browser settings",
   *      because WebView settings are not user-accessible.
   *
   * HOST APP FIX (requires host app developer — Prasidha):
   * The Android host app's WebView must implement:
   *
   *   webView.setWebChromeClient(new WebChromeClient() {
   *     @Override
   *     public void onGeolocationPermissionsShowPrompt(
   *         String origin, GeolocationPermissions.Callback callback) {
   *       callback.invoke(origin, true, false);
   *     }
   *   });
   *
   * AND the host app's AndroidManifest.xml must include:
   *   <uses-permission android:name="android.permission.ACCESS_FINE_LOCATION"/>
   *   <uses-permission android:name="android.permission.ACCESS_COARSE_LOCATION"/>
   *
   * Until the host app is updated, users on Android WebView should open
   * the app URL directly in Chrome browser for full GPS functionality.
   *
   * SECURITY (Prasidha): Safe error messages only. Raw err object never exposed.
   */

  /* Detect Android WebView — Prasidha
     Android WebView UA contains "wv" token or "Version/X.X Chrome" without "Mobile Safari".
     This detection is used only for user-facing error message customisation. */
  const ua = navigator.userAgent || '';
  const isAndroidWebView = /Android/.test(ua) && (/wv\)/.test(ua) || /Version\/\d/.test(ua));

  let safetyId = null;  /* setTimeout reference — cleared when GPS settles    */
  let settled  = false; /* flag — ensures exactly ONE toast fires per call     */

  const geoPromise = new Promise(resolve => {

    /* Check if geolocation API is present at all */
    if (!navigator.geolocation) {
      if (!settled) {
        settled = true;
        toast('Geolocation is not supported on this device or browser.', 'err');
      }
      return resolve(null);
    }

    navigator.geolocation.getCurrentPosition(
      pos => {
        /* GPS success — cancel safety timer immediately (Prasidha: FIX-05 BUG-A) */
        clearTimeout(safetyId);
        settled = true;
        resolve(`${pos.coords.latitude.toFixed(6)},${pos.coords.longitude.toFixed(6)}`);
      },
      err => {
        /* GPS error — cancel safety timer, show exactly one message (FIX-05 BUG-B) */
        clearTimeout(safetyId);
        if (!settled) {
          settled = true;

          if (err.code === 1) {
            /* PERMISSION_DENIED — different message for WebView vs browser */
            if (isAndroidWebView) {
              /* WebView: user cannot change settings from inside the app —
                 direct them to open in Chrome instead (Prasidha: Android WebView fix) */
              showWebViewModal();
            } else {
              toast('Location access denied. Allow it in your browser settings.', 'err');
            }
          } else if (err.code === 2) {
            toast('Location unavailable. Please try again.', 'err');
          } else if (err.code === 3) {
            toast('Location timed out. Please try again.', 'err');
          } else {
            toast('Location error. Please try again.', 'err');
          }
        }
        resolve(null);
      },
      {
        enableHighAccuracy: true,
        timeout:    8000,
        maximumAge: 0     /* always fresh GPS — never a cached stale position (Prasidha) */
      }
    );
  });

  /* Safety promise — absolute 9s ceiling. Only shows toast if geo never fired. */
  const safetyPromise = new Promise(resolve => {
    safetyId = setTimeout(() => {
      if (!settled) {
        settled = true;
        toast('Location timed out. Please try again.', 'err');
      }
      resolve(null);
    }, 9000);
  });

  /* Race: whichever settles first wins. Safety timer always cancelled on geo settle. */
  return Promise.race([geoPromise, safetyPromise]);
}

/* ══════════════════════════════════════════════════════════════
   LOCATION CACHE
   Prasidha: localStorage cache for autocomplete and chips.
   SECURITY: entries sanitized on write. textContent on all renders.
   MIGRATION: swap localStorage with sessionStorage or userProfile
   service when moving to SharePoint SPFx.
*/

/** getCached — Prasidha: returns cached location array */
function getCached() {
  try { return JSON.parse(localStorage.getItem(KEY_LOCS) || '[]'); }
  catch { return []; }
}

/** saveLoc — Prasidha: prepends location, deduplicates, trims to MAX_LOCS */
function saveLoc(loc) {
  if (!loc) return;
  let list = getCached().filter(l => l.toLowerCase() !== loc.toLowerCase());
  list.unshift(loc);
  localStorage.setItem(KEY_LOCS, JSON.stringify(list.slice(0, MAX_LOCS)));
}

/**
 * renderChips — Prasidha
 * Renders last 3 locations as quick-select pill buttons.
 * SECURITY: textContent on all chip text — XSS safe.
 */
function renderChips(inputId, chipsId, dropId) {
  const locs = getCached().slice(0, 3);
  const wrap = g(chipsId); if (!wrap) return;
  wrap.innerHTML = '';
  locs.forEach(loc => {
    const c = document.createElement('button');
    c.type = 'button'; c.className = 'chip'; c.title = loc;
    c.textContent = '📍 ' + loc; /* SECURITY: textContent (Prasidha) */
    c.addEventListener('click', () => {
      const inp = g(inputId); if (inp) inp.value = loc;
      const drop = g(dropId); if (drop) drop.classList.remove('open');
    });
    wrap.appendChild(c);
  });
}

/**
 * setupAC — Prasidha
 * Sets up autocomplete for one input/dropdown pair.
 * MUST be called only once per input at DOMContentLoaded.
 * Multiple calls stack document.click listeners — memory leak.
 */
function setupAC(inputId, dropId, chipsId) {
  const input = g(inputId), drop = g(dropId);
  if (!input || !drop) return;
  input.addEventListener('focus', () => {
    renderChips(inputId, chipsId, dropId);
    renderDrop(inputId, dropId);
  });
  input.addEventListener('input', () => renderDrop(inputId, dropId));
  document.addEventListener('click', e => {
    if (!input.contains(e.target) && !drop.contains(e.target))
      drop.classList.remove('open');
  });
}

/**
 * renderDrop — Prasidha
 * Populates autocomplete dropdown. Max 5 suggestions.
 * SECURITY: all items via createElement/textContent — XSS safe.
 */
function renderDrop(inputId, dropId) {
  const q    = g(inputId)?.value.toLowerCase().trim();
  const drop = g(dropId); if (!drop) return;
  const hits = (q
    ? getCached().filter(l => l.toLowerCase().includes(q))
    : getCached()
  ).slice(0, 5);

  if (!hits.length) { drop.classList.remove('open'); return; }
  drop.innerHTML = '';
  hits.forEach(loc => {
    const item = document.createElement('div');
    item.className = 'ac-item';
    const ico = document.createElement('span'); ico.textContent = '📍';
    const txt = document.createElement('span'); txt.textContent = loc;
    item.append(ico, txt);
    item.addEventListener('click', () => {
      const inp = g(inputId); if (inp) inp.value = loc;
      drop.classList.remove('open');
    });
    drop.appendChild(item);
  });
  drop.classList.add('open');
}

/* ══════════════════════════════════════════════════════════════
   TOAST — Prasidha
   Centered blur-backdrop overlay. 1.5s auto-dismiss.
   SECURITY: msg via textContent — never innerHTML.
*/
let _toastTimer = null;

/**
 * toast — Prasidha
 * @param {string} msg - Message text
 * @param {'ok'|'err'|''} type - Colour variant
 */
function toast(msg, type = '') {
  const ov   = g('toast-ov');
  const pill = g('toast-pill');
  if (!ov || !pill) return;

  if (_toastTimer) { clearTimeout(_toastTimer); _toastTimer = null; }

  pill.textContent = msg; /* SECURITY: textContent (Prasidha) */
  pill.className = 'toast-pill' + (type ? ' ' + type : '');

  ov.classList.remove('hidden', 'to-out');

  _toastTimer = setTimeout(() => {
    ov.classList.add('to-out');
    setTimeout(() => {
      ov.classList.add('hidden');
      ov.classList.remove('to-out');
    }, 280);
  }, 1500);
}

/* ══════════════════════════════════════════════════════════════
   HELPERS — see top of file for hoisted function declarations
*/

/**
 * isNewDay — Prasidha
 * True if lastDate was on a previous calendar day.
 * Drives the midnight session reset.
 */
function isNewDay(lastDate) {
  if (!lastDate) return false;
  return new Date(lastDate).setHours(0,0,0,0) < new Date().setHours(0,0,0,0);
}

/**
 * fmt — Prasidha
 * Formats ISO timestamp as HH:MM AM/PM (en-IN locale).
 */
function fmt(iso) {
  if (!iso) return '--:--';
  return new Date(iso).toLocaleTimeString('en-IN', {
    hour: '2-digit', minute: '2-digit', hour12: true
  });
}

/**
 * msDur — Prasidha
 * Converts milliseconds to HH:MM:SS string.
 * Used by rAF loop for live timer and duration() for frozen display.
 */
function msDur(ms) {
  if (!ms || isNaN(ms) || ms < 0) return '00:00:00';
  return [
    Math.floor(ms / 3600000),
    Math.floor((ms % 3600000) / 60000),
    Math.floor((ms % 60000) / 1000)
  ].map(n => n.toString().padStart(2, '0')).join(':');
}

/**
 * duration — Prasidha
 * Computes and formats duration between two ISO timestamps.
 */
function duration(inISO, outISO) {
  if (!inISO || !outISO) return '--';
  return msDur(new Date(outISO) - new Date(inISO));
}

/*
 * ══════════════════════════════════════════════════════════════
 *  Prasidha Jagtap | IT · Aditya Birla Group (Seamex)
 *  Geo Attendance v07 — Definitive Golden Build
 *  Built for field teams. Maintained with care and intent.
 *  If you are reading this: keep the standards. 🚀
 * ══════════════════════════════════════════════════════════════
 */
