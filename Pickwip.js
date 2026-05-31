/* Picking WIP — Command Center Logic
   Updated for bulletproof backend:
   - Live WIP from payload.data
   - Movement check from payload.movement
   - Daily metrics from payload.dailyStats

   Correct metric definitions:
   - SF Productive Today = gain into SF Scan & Verify
   - FSV Productive Today = gain into FSV Scan & Verify
   - Breakage Count Today = drop from Breakage to Picking
   - Replenishment Count Today = drop from Replenishment
*/

const API = 'https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec';

const AUTO_REFRESH_MS = 60000;
const USE_DEMO_ON_ERROR = false;

const TRACKED = [
  'SF Scan & Verify',
  'FSV Scan & Verify',
  'Frame Only Scan & Verify',
  'Breakage to Picking',
  'Replenishment'
];

const COLORS = [
  '#19c8ff',
  '#40a4ff',
  '#ff9f1c',
  '#ffd21f',
  '#77889c',
  '#b44cff',
  '#21e57c'
];

const CLASS_MAP = {
  'SF Scan & Verify': 'sf',
  'FSV Scan & Verify': 'fsv',
  'Frame Only Scan & Verify': 'frame',
  'Breakage to Picking': 'brk',
  'Replenishment': 'brk'
};

let STATE = {
  wip: [],
  completion: [],
  history: [],
  movement: [],
  movementStatus: 'LOADING',
  movementMessage: 'Waiting for movement comparison...',
  health: {},
  dailyStats: {},
  activityByOperator: [],
  hourlyActivity: [],
  manualCounts: {},
  dailyActivityCounts: {},
  snapshotRecordCounts: {},
  recordCounts: {},
  totalWip: 0,
  totalPickingWip: 0,
  totalInventoryWip: 0,
  scanVerifyWip: 0,
  demo: false
};

let OP_STATION_VIEW = 'all';
let OP_DETAIL_STATION = '';
let OP_SELECTED_OPERATOR_KEY = '';
let PERSONAL_HOURLY_VIEW = 'all';

const PICKING_JPH_STORAGE_KEY = 'picking_operator_jph_config_v1';
const PICKING_JPH_DEFAULT_CONFIG = {
  sfTarget: 98,
  fsvTarget: 72,
  amberPct: 90,
  redBelowPct: 90
};

let PICKING_JPH_CONFIG = loadPickingJphConfig();

const DEMO = {
  wip: [
    { picking: 'Breakage to Picking', department: 'Inventory', total: 1, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'FSV Scan & Verify', department: 'Inventory', total: 44, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Out of Finish Queue', department: 'Inventory', total: 44, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Out of Surface Queue', department: 'Inventory', total: 47, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Re-Cal', department: 'Inventory', total: 17, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Release from Hold', department: 'Inventory', total: 1, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Replenishment', department: 'Inventory', total: 2, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'SF Scan & Verify', department: 'Inventory', total: 193, lastUpdated: '2026-04-30 13:31:39' }
  ],
  dailyStats: {
    sfProductivityToday: 33,
    fsvProductivityToday: 26,
    breakageCountToday: 6,
    replenishmentCountToday: 4,
    totalProductivityToday: 59
  },
  movement: [
    {
      subDepartment: 'Out of Surface Queue',
      previousTotal: 80,
      currentTotal: 47,
      difference: -33,
      movementType: 'DROP'
    },
    {
      subDepartment: 'SF Scan & Verify',
      previousTotal: 160,
      currentTotal: 193,
      difference: 33,
      movementType: 'GAIN'
    },
    {
      subDepartment: 'Out of Finish Queue',
      previousTotal: 70,
      currentTotal: 44,
      difference: -26,
      movementType: 'DROP'
    },
    {
      subDepartment: 'FSV Scan & Verify',
      previousTotal: 18,
      currentTotal: 44,
      difference: 26,
      movementType: 'GAIN'
    }
  ]
};

/* ──────────────────────────────────────────────────
   Error Log
────────────────────────────────────────────────── */

const ERROR_LOG = [];
let errIdCounter = 0;

function logError(level = 'error', title, detail = '', source = '') {
  const id = ++errIdCounter;
  const ts = new Date().toLocaleTimeString('en-US', {
    hour12: false,
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });

  ERROR_LOG.unshift({
    id,
    level,
    title,
    detail,
    source,
    ts
  });

  if (ERROR_LOG.length > 100) {
    ERROR_LOG.pop();
  }

  renderErrorLog();
}

function renderErrorLog() {
  const body = document.getElementById('errorLogBody');
  const empty = document.getElementById('errorLogEmpty');
  const badge = document.getElementById('errCountBadge');

  if (!body) return;

  const errCount = ERROR_LOG.filter(e => e.level === 'error').length;

  if (badge) {
    badge.textContent = errCount;
    badge.style.display = errCount > 0 ? 'inline-flex' : 'none';
  }

  body.querySelectorAll('.err-row').forEach(el => el.remove());

  if (ERROR_LOG.length === 0) {
    if (empty) empty.style.display = 'block';
    return;
  }

  if (empty) empty.style.display = 'none';

  ERROR_LOG.forEach(e => {
    const row = document.createElement('div');
    row.className = 'err-row';
    row.dataset.errId = e.id;

    row.innerHTML = `
      <div class="err-time">${e.ts}</div>
      <div class="err-level ${e.level}">${e.level.toUpperCase()}</div>
      <div class="err-msg">
        <strong>${escHtml(e.title)}</strong>
        ${e.detail ? `<em>${escHtml(e.detail)}</em>` : ''}
        ${e.source ? `<em style="color:#7dd4f8">@ ${escHtml(e.source)}</em>` : ''}
      </div>
      <button class="err-dismiss" title="Dismiss" onclick="dismissError(${e.id})">✕</button>
    `;

    body.appendChild(row);
  });
}

function dismissError(id) {
  const idx = ERROR_LOG.findIndex(e => e.id === id);

  if (idx !== -1) {
    ERROR_LOG.splice(idx, 1);
  }

  renderErrorLog();
}

function clearErrorLog() {
  ERROR_LOG.length = 0;
  renderErrorLog();
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ──────────────────────────────────────────────────
   Init
────────────────────────────────────────────────── */

/* ──────────────────────────────────────────────────
   Loading Screen Controller — Industrial Terminal
────────────────────────────────────────────────── */

function loaderTick() {
  const el = document.getElementById('loaderTime');
  if (el) el.textContent = new Date().toLocaleTimeString('en-US', { hour12: false });
}

function loaderSetBar(pct, label) {
  const fill  = document.getElementById('loaderBarFill');
  const lbl   = document.getElementById('loaderBarLabel');
  const num   = document.getElementById('ldPctNum');
  const fstat = document.getElementById('ldFooterStatus');
  if (fill) fill.style.width = pct + '%';
  if (lbl)  lbl.textContent  = label;
  if (num)  num.innerHTML    = `${pct}<span>%</span>`;
  if (fstat) fstat.textContent = pct >= 100 ? 'ONLINE' : 'BOOTING';
}

function loaderSetCheck(id, state, statusText) {
  const row  = document.getElementById(id);
  const fill = document.getElementById('lsf-' + id.replace('lc-', ''));
  const val  = document.getElementById('lsv-' + id.replace('lc-', ''));
  if (!row) return;
  row.className = 'ld-status-row ' + (state === 'ok' ? 'done' : state === 'loading' ? 'active' : '');
  if (fill) fill.style.width = state === 'ok' ? '100%' : state === 'loading' ? '55%' : '0%';
  if (val)  val.textContent  = statusText || 'STANDBY';
}

function loaderAppendLog(tag, tagClass, text, delay = 0) {
  const body = document.getElementById('ldLogBody');
  if (!body) return;
  setTimeout(() => {
    const line = document.createElement('div');
    line.className = 'ld-log-line';
    line.innerHTML = `<span class="ld-tag ${tagClass}">[${tag}]</span> ${text}`;
    body.appendChild(line);
    body.scrollTop = body.scrollHeight;
  }, delay);
}

function loaderSetCursor(text) {
  const el = document.getElementById('ldCursorText');
  if (el) el.textContent = text;
}

function dismissLoader() {
  const screen = document.getElementById('loaderScreen');
  const mainEl = document.querySelector('.main');

  if (mainEl) {
    mainEl.classList.remove('loading-active');
  }

  if (!screen || screen.classList.contains('hidden')) return;

  screen.classList.add('hidden');

  setTimeout(() => {
    if (screen && screen.parentNode) {
      screen.parentNode.removeChild(screen);
    }
  }, 450);
}

document.addEventListener('DOMContentLoaded', () => {
  const mainEl = document.querySelector('.main');
  const loaderEl = document.getElementById('loaderScreen');

  if (mainEl) {
    mainEl.classList.add('loading-active');
  }

  if (loaderEl) {
    loaderEl.classList.add('integrated-loader');
  }

  loaderTick();
  setInterval(loaderTick, 1000);

  loaderSetBar(15, 'Preparing Picking command center...');
  loaderSetCheck('lc-api', 'loading', 'connecting');

  setupTabs();
  setupOperatorControls();
  renderPickingJphConfigPanel();

  window.addEventListener('error', ev => {
    logError(
      'error',
      ev.message || 'Runtime Error',
      `${ev.filename || ''}:${ev.lineno}:${ev.colno}`,
      'window.onerror'
    );
  });

  window.addEventListener('unhandledrejection', ev => {
    const msg = ev.reason instanceof Error ? ev.reason.message : String(ev.reason);
    logError('error', 'Unhandled Promise Rejection', msg, 'window.unhandledrejection');
  });

  document.addEventListener('keydown', ev => {
    if (ev.key === 'Escape') closeOperatorStationDetail();
  });

  initParticles();

  fetchAll(true);
  setInterval(() => fetchAll(false), AUTO_REFRESH_MS);

  setTimeout(() => {
    document.querySelectorAll('.row-bar i, .progress-mini i, .q-card-bar i').forEach(el => {
      const w = el.style.width;
      el.style.width = '0';
      requestAnimationFrame(() => {
        el.style.width = w;
      });
    });
  }, 300);
});

function setupTabs() {
  document.querySelectorAll('.side-link[data-tab]').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;

      document.querySelectorAll('.side-link').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('tab-hidden'));

      const panel = document.getElementById('tab-' + target);
      if (panel) {
        panel.classList.remove('tab-hidden');
      }

      const titleEl = document.querySelector('h1');
      const labelEl = btn.querySelector('b');

      if (titleEl && labelEl) {
        titleEl.textContent = labelEl.textContent === 'Command Overview' ? 'Picking Dashboard' : 'Picking Floor — ' + labelEl.textContent;
      }
    });
  });
}

function setupOperatorControls() {
  const controls = [
    document.getElementById('operatorSearch'),
    document.getElementById('operatorStationFilter'),
    document.getElementById('operatorSortMode')
  ].filter(Boolean);

  controls.forEach(control => {
    const eventName = control.tagName === 'INPUT' ? 'input' : 'change';
    control.addEventListener(eventName, () => {
      if (control.id === 'operatorStationFilter') {
        OP_STATION_VIEW = control.value || 'all';
        OP_SELECTED_OPERATOR_KEY = '';
      }
      renderOperatorActivityTab();
    });
  });
}


/* ──────────────────────────────────────────────────
   Picking Operator JPH Configuration
────────────────────────────────────────────────── */

function getDefaultPickingJphConfig() {
  return Object.assign({}, PICKING_JPH_DEFAULT_CONFIG);
}

function loadPickingJphConfig() {
  try {
    const raw = localStorage.getItem(PICKING_JPH_STORAGE_KEY);
    if (!raw) return getDefaultPickingJphConfig();

    const parsed = JSON.parse(raw);
    return {
      sfTarget: Math.max(0, num(parsed.sfTarget ?? PICKING_JPH_DEFAULT_CONFIG.sfTarget)),
      fsvTarget: Math.max(0, num(parsed.fsvTarget ?? PICKING_JPH_DEFAULT_CONFIG.fsvTarget)),
      amberPct: Math.max(1, Math.min(100, num(parsed.amberPct ?? PICKING_JPH_DEFAULT_CONFIG.amberPct))),
      redBelowPct: Math.max(1, Math.min(100, num(parsed.redBelowPct ?? PICKING_JPH_DEFAULT_CONFIG.redBelowPct)))
    };
  } catch (err) {
    console.warn('Failed to load Picking JPH config. Using defaults.', err);
    return getDefaultPickingJphConfig();
  }
}

function savePickingJphConfig(config) {
  PICKING_JPH_CONFIG = Object.assign(getDefaultPickingJphConfig(), config || {});
  localStorage.setItem(PICKING_JPH_STORAGE_KEY, JSON.stringify(PICKING_JPH_CONFIG));
  renderPickingJphConfigPanel();
  renderOperatorActivityTab();
  renderPersonalHourlyPerformance();
}

function savePickingJphConfigFromUi() {
  const sfTarget = num(document.getElementById('cfgPickingSfTarget')?.value || PICKING_JPH_CONFIG.sfTarget);
  const fsvTarget = num(document.getElementById('cfgPickingFsvTarget')?.value || PICKING_JPH_CONFIG.fsvTarget);
  const amberPct = num(document.getElementById('cfgPickingAmberPct')?.value || PICKING_JPH_CONFIG.amberPct);

  savePickingJphConfig({
    sfTarget: Math.max(0, sfTarget),
    fsvTarget: Math.max(0, fsvTarget),
    amberPct: Math.max(1, Math.min(100, amberPct)),
    redBelowPct: Math.max(1, Math.min(100, amberPct))
  });

  const state = document.getElementById('cfgPickingSaveState');
  if (state) {
    state.textContent = 'Saved';
    state.classList.add('saved');
    clearTimeout(window.__pickCfgStateTimer);
    window.__pickCfgStateTimer = setTimeout(() => {
      state.textContent = 'Local config active';
      state.classList.remove('saved');
    }, 1800);
  }

  toast('Picking JPH configuration saved.');
}

function resetPickingJphConfig() {
  savePickingJphConfig(getDefaultPickingJphConfig());
  toast('Picking JPH configuration reset to defaults.');
}

function renderPickingJphConfigPanel() {
  const sfInput = document.getElementById('cfgPickingSfTarget');
  const fsvInput = document.getElementById('cfgPickingFsvTarget');
  const amberInput = document.getElementById('cfgPickingAmberPct');
  const sfAmber = document.getElementById('cfgPickingSfAmberLabel');
  const fsvAmber = document.getElementById('cfgPickingFsvAmberLabel');
  const sfRed = document.getElementById('cfgPickingSfRedLabel');
  const fsvRed = document.getElementById('cfgPickingFsvRedLabel');
  const summary = document.getElementById('cfgPickingSummary');

  if (sfInput && document.activeElement !== sfInput) sfInput.value = PICKING_JPH_CONFIG.sfTarget;
  if (fsvInput && document.activeElement !== fsvInput) fsvInput.value = PICKING_JPH_CONFIG.fsvTarget;
  if (amberInput && document.activeElement !== amberInput) amberInput.value = PICKING_JPH_CONFIG.amberPct;

  const sfAmberValue = Math.ceil(PICKING_JPH_CONFIG.sfTarget * (PICKING_JPH_CONFIG.amberPct / 100));
  const fsvAmberValue = Math.ceil(PICKING_JPH_CONFIG.fsvTarget * (PICKING_JPH_CONFIG.amberPct / 100));

  if (sfAmber) sfAmber.textContent = `${sfAmberValue}–${Math.max(PICKING_JPH_CONFIG.sfTarget - 1, sfAmberValue)} / hr`;
  if (fsvAmber) fsvAmber.textContent = `${fsvAmberValue}–${Math.max(PICKING_JPH_CONFIG.fsvTarget - 1, fsvAmberValue)} / hr`;
  if (sfRed) sfRed.textContent = `< ${sfAmberValue} / hr`;
  if (fsvRed) fsvRed.textContent = `< ${fsvAmberValue} / hr`;

  if (summary) {
    summary.textContent = `Green: 100%+ · Amber: ${PICKING_JPH_CONFIG.amberPct}%–99% · Red: below ${PICKING_JPH_CONFIG.redBelowPct}%`;
  }
}

function getStationJphConfig(station) {
  const name = displayPickingStationName(station || '').toUpperCase();
  const target = name.includes('FSV') ? PICKING_JPH_CONFIG.fsvTarget : PICKING_JPH_CONFIG.sfTarget;

  return {
    target: num(target),
    amberPct: num(PICKING_JPH_CONFIG.amberPct),
    redBelowPct: num(PICKING_JPH_CONFIG.redBelowPct)
  };
}

function isPickingJphTargetHour(hour) {
  const text = String(hour || '').trim().toUpperCase();
  return !/^(6:00 AM|6:00 PM|7:00 PM|8:00 PM)$/.test(text);
}

function getPickingJphResult(value, station, hour) {
  const actual = num(value);
  const cfg = getStationJphConfig(station);
  const hasTarget = isPickingJphTargetHour(hour) && cfg.target > 0;

  if (!hasTarget) {
    return {
      cls: actual > 0 ? 'neutral active' : 'neutral',
      status: 'NO_TARGET',
      target: 0,
      pct: null,
      pctText: 'No Target',
      targetLabel: 'No target'
    };
  }

  const pct = Math.round((actual / cfg.target) * 100);
  let cls = 'bad';
  let status = 'RED';

  if (pct >= 100) {
    cls = 'good';
    status = 'GREEN';
  } else if (pct >= cfg.amberPct) {
    cls = 'warn';
    status = 'AMBER';
  }

  return {
    cls,
    status,
    target: cfg.target,
    pct,
    pctText: `${pct}%`,
    targetLabel: `${cfg.target}/hr`
  };
}

function getPickingTargetLabel(hour, station) {
  return getPickingJphResult(0, station || OP_DETAIL_STATION || OP_STATION_VIEW, hour).targetLabel;
}

function getPickingStationTargetText(station) {
  const cfg = getStationJphConfig(station);
  const amberValue = Math.ceil(cfg.target * (cfg.amberPct / 100));
  return `${cfg.target}/hr target · Amber ${amberValue}-${Math.max(cfg.target - 1, amberValue)} · Red < ${amberValue}`;
}

/* ──────────────────────────────────────────────────
   API
────────────────────────────────────────────────── */

async function apiFetch(debug = true) {
  /*
   * IMPORTANT:
   * This dashboard must call the Picking-specific Production Flow API action.
   * If action=pickingDashboard is missing, the API returns the default
   * Surface/AR/Finish productionFlow payload and this page will show zeros.
   */
  const url = `${API}?action=pickingDashboard&debug=${debug ? 'true' : 'false'}&t=${Date.now()}`;

  const res = await fetch(url, {
    method: 'GET',
    redirect: 'follow',
    cache: 'no-store'
  });

  if (!res.ok) {
    const msg = `HTTP ${res.status}`;
    logError('error', 'API HTTP Error', msg, 'Picking Dashboard API');
    throw new Error(msg);
  }

  const responseText = await res.text();

  try {
    const payload = JSON.parse(responseText);

    if (payload && payload.action && payload.action !== 'pickingDashboard') {
      logError(
        'warning',
        'Unexpected API Action',
        `Expected pickingDashboard but received ${payload.action}`,
        'Picking Dashboard API'
      );
    }

    if (payload && payload.requestedArea && payload.requestedArea !== 'Picking') {
      logError(
        'warning',
        'Unexpected API Area',
        `Expected Picking but received ${payload.requestedArea}`,
        'Picking Dashboard API'
      );
    }

    return payload;
  } catch {
    logError('error', 'API JSON Parse Failed', responseText.slice(0, 160), 'Picking Dashboard API');
    throw new Error('Bad JSON from API');
  }
}

async function fetchAll(showToast = false) {
  /*
   * Loading should not clear the dashboard.
   * Keep last good numbers visible until a good API response arrives.
   */
  setLiveState('loading', 'LOADING');

  if (showToast) {
    toast('Refreshing Picking WIP data...');
  }

  // Loader: step 1 — hitting API
  loaderSetBar(30, 'Syncing WIP, activity, and operator data...');
  loaderSetCheck('lc-api', 'loading', 'FETCHING');
  loaderSetCursor('Querying picking floor data source');

  try {
    const payload = await apiFetch(true);

    if (payload.status !== 'success') {
      throw new Error(payload.message || 'API returned an error');
    }

    // Loader: step 2 — API OK
    loaderSetBar(55, 'Building Picking WIP split...');
    loaderSetCheck('lc-api', 'ok', 'ONLINE');
    loaderSetCheck('lc-wip', 'loading', 'LOADING');
    loaderSetCursor('Normalizing WIP queue records');
    loaderAppendLog('  OK  ', 'ok', 'API connection established');

    /*
     * Build next state first.
     * Do not replace STATE until payload is confirmed good.
     */
    const nextWip = normalizeWip(payload.wip || payload.data || []);
    const nextDailyStats = normalizeDailyStats(
      payload.dailyStats || {},
      payload.dailyActivityCounts || payload.snapshotRecordCounts || payload.recordCounts || {},
      payload.manualCounts || {}
    );

    console.log('Picking daily stats loaded:', nextDailyStats, 'dailyActivityCounts:', payload.dailyActivityCounts || {}, 'manualCounts:', payload.manualCounts || {});

    // Loader: step 3 — WIP normalized
    loaderSetBar(72, 'Checking queue movement and snapshots...');
    loaderSetCheck('lc-wip', 'ok', 'LOADED');
    loaderSetCheck('lc-movement', 'loading', 'CHECKING');
    loaderSetCursor('Running movement delta analysis');
    loaderAppendLog('  OK  ', 'ok', `WIP records loaded — ${nextWip.length} queue(s) found`);

    const nextState = {
      wip: nextWip,
      completion: normalizeCompletion(
        payload.completion && payload.completion.length
          ? payload.completion
          : buildCompletionFromDailyStats(nextDailyStats)
      ),
      history: payload.history || [],
      movement: normalizeMovement(payload.movement || []),
      movementStatus: payload.movementStatus || 'NORMAL',
      movementMessage: payload.movementMessage || 'Normal WIP movement.',
      health: payload.health || {},
      dailyStats: nextDailyStats,
      activityByOperator: normalizeOperatorActivity(payload.activityByOperator || []),
      hourlyActivity: normalizeHourlyActivity(payload.hourlyActivity || []),
      manualCounts: payload.manualCounts || {},
      dailyActivityCounts: payload.dailyActivityCounts || payload.snapshotRecordCounts || payload.recordCounts || {},
      snapshotRecordCounts: payload.snapshotRecordCounts || payload.dailyActivityCounts || payload.recordCounts || {},
      recordCounts: payload.dailyActivityCounts || payload.snapshotRecordCounts || payload.recordCounts || {},
      totalWip: getSafePickingWipTotal_(payload, nextWip),
      totalPickingWip: getSafePickingWipTotal_(payload, nextWip),
      totalInventoryWip: getSafeInventoryWipTotal_(payload, nextWip),
      scanVerifyWip: getSafeScanVerifyWipTotal_(payload, nextWip),
      wipBreakdown: payload.wipBreakdown || {},
      demo: false
    };

    // Loader: step 4 — all done
    loaderSetBar(90, 'Rendering scorecards and operator drilldowns...');
    loaderSetCheck('lc-movement', 'ok', 'VALID');
    loaderSetCheck('lc-stats', 'loading', 'LOADING');
    loaderSetCursor('Compiling daily productivity metrics');
    loaderAppendLog('  OK  ', 'ok', `Movement check — ${payload.movementStatus || 'NORMAL'}`);

    /*
     * Only now replace STATE and repaint the page.
     */
    STATE = nextState;

    setLiveState('ok', 'LIVE');
    renderAll();

    // Loader: complete — dismiss after brief hold
    loaderSetBar(100, 'Picking dashboard online.');
    loaderSetCheck('lc-stats', 'ok', 'READY');
    loaderSetCursor('Dashboard ready — launching');
    loaderAppendLog('  OK  ', 'ok', 'All systems online — launching dashboard');
    setTimeout(dismissLoader, 900);

    if (showToast) {
      toast('Live data loaded.');
    }

  } catch (err) {
    console.warn(err);
    logError('error', 'Fetch Failed', err.message, 'fetchAll → Picking WIP API');

    /*
     * Do not replace STATE.
     * Do not call renderAll().
     * Keep last good dashboard values visible.
     */
    setLiveState('error', 'ERROR');
    toast('API error. Keeping last good data on screen.');

    // Dismiss loader even on error — don't block the UI forever
    loaderSetBar(100, 'Connection error — showing cached data.');
    loaderAppendLog(' FAIL ', 'err', 'API error: ' + err.message.slice(0, 50));
    setTimeout(dismissLoader, 1200);
  }
}

/* ──────────────────────────────────────────────────
   Normalizers
────────────────────────────────────────────────── */

function normalizeWip(rows) {
  return rows.map(r => ({
    picking:
      r.subDepartment ||
      r.picking ||
      r.picking_type ||
      r.pickingType ||
      r.type ||
      r.name ||
      r.PickingType ||
      'Unknown',

    department:
      r.department ||
      r.Department ||
      'Inventory',

    total:
      num(
        r.currentJobTotal ??
        r.total ??
        r.wip ??
        r.count ??
        r.WIP ??
        r['WIP Count']
      ),

    lastUpdated:
      r.fileUpdatedTime ||
      r.importedAt ||
      r.lastUpdated ||
      r.last_updated ||
      r.updated ||
      r.LastUpdated ||
      r['Last Updated'] ||
      ''
  })).filter(r => r.picking !== 'Unknown');
}

function normalizeCompletion(rows) {
  return rows.map(r => ({
    picking:
      r.picking ||
      r.picking_type ||
      r.pickingType ||
      r.type ||
      r.name ||
      'Unknown',

    completed:
      num(
        r.completed ??
        r.jobs_completed ??
        r.jobs ??
        r.CompletedToday ??
        r['Completed Today']
      ),

    direction:
      String(r.direction || r.dir || 'completed').toLowerCase(),

    before:
      num(r.wip_before ?? r.before),

    after:
      num(r.wip_after ?? r.after),

    timestamp:
      r.timestamp ||
      r.ts ||
      r.lastUpdated ||
      r.last_updated ||
      ''
  })).filter(r => r.picking !== 'Unknown');
}

function normalizeMovement(rows) {
  return rows.map(r => {
    const diff = num(r.difference);

    return {
      subDepartment:
        r.subDepartment ||
        r.station ||
        r.name ||
        'Unknown',

      previousTotal:
        num(r.previousTotal),

      currentTotal:
        num(r.currentTotal),

      difference:
        diff,

      movementType:
        r.movementType ||
        (diff > 0 ? 'GAIN' : diff < 0 ? 'DROP' : 'NO_CHANGE'),

      currentFileName:
        r.currentFileName || '',

      currentFileTime:
        r.currentFileTime || '',

      previousFileName:
        r.previousFileName || '',

      previousFileTime:
        r.previousFileTime || ''
    };
  }).filter(r => r.subDepartment !== 'Unknown');
}

function normalizeDailyStats(stats, dailyActivityCounts = {}, manualCounts = {}) {
  const safeStats = stats || {};
  const safeDailyCounts = dailyActivityCounts || {};
  const safeManual = manualCounts || {};

  const sf = num(safeStats.sfProductivityToday ?? safeStats.surfaceScanActivityToday);
  const fsv = num(safeStats.fsvProductivityToday ?? safeStats.combinedFsvFrameActivityToday);
  const frameOnly = num(safeStats.frameOnlyProductivityToday ?? safeStats.framePickingActivityToday);
  const frameOnlyWip = num(safeStats.frameOnlyScanVerifyWip ?? safeStats.frameOnlyWip);
  const lensPicking = num(safeStats.lensPickingActivityToday);

  /*
   * Record counts now come from PICKING_DAILY_ACTIVITY_LOG.
   * Manual counts are only a fallback for old API payloads.
   * Do not let a stale manual sheet override the new automatic log.
   */
  const dailyBreakage = getNullableNumber_(safeDailyCounts.breakageCountToday);
  const dailyReplenishment = getNullableNumber_(safeDailyCounts.replenishmentCountToday);
  const manualRecordCounts = getManualRecordCounts_(safeManual);

  const breakage = dailyBreakage !== null
    ? dailyBreakage
    : getNullableNumber_(safeStats.breakageCountToday) !== null
      ? num(safeStats.breakageCountToday)
      : manualRecordCounts.breakageCountToday !== null
        ? manualRecordCounts.breakageCountToday
        : 0;

  const replenishment = dailyReplenishment !== null
    ? dailyReplenishment
    : getNullableNumber_(safeStats.replenishmentCountToday) !== null
      ? num(safeStats.replenishmentCountToday)
      : manualRecordCounts.replenishmentCountToday !== null
        ? manualRecordCounts.replenishmentCountToday
        : 0;

  const totalProductivity = num(
    safeStats.totalProductivityToday ||
    (lensPicking + frameOnly) ||
    (sf + fsv)
  );

  const recordCountToday = num(
    safeStats.recordCountToday ||
    safeDailyCounts.recordCountToday ||
    (breakage + replenishment)
  );

  return {
    sfProductivityToday: sf,
    fsvProductivityToday: fsv,
    frameOnlyProductivityToday: frameOnly,
    frameOnlyScanVerifyWip: frameOnlyWip,

    surfaceScanActivityToday: num(safeStats.surfaceScanActivityToday ?? sf),
    combinedFsvFrameActivityToday: num(safeStats.combinedFsvFrameActivityToday ?? fsv),
    outOfFinishQueueDropToday: num(safeStats.outOfFinishQueueDropToday),
    framePickingActivityToday: frameOnly,
    lensFsvActivityToday: num(safeStats.lensFsvActivityToday),
    lensPickingActivityToday: lensPicking || (sf + fsv - frameOnly),

    breakageCountToday: breakage,
    replenishmentCountToday: replenishment,
    recordCountToday: recordCountToday,
    totalProductivityToday: totalProductivity,
    activeOperators: num(safeStats.activeOperators),
    calculationMode: safeStats.calculationMode || safeDailyCounts.recordCountMode || '',
    calculationNote: safeStats.calculationNote || '',
    recordCountMode: safeStats.recordCountMode || safeDailyCounts.recordCountMode || '',
    recordCountSource: safeStats.recordCountSource || safeDailyCounts.recordCountSource || '',
    countRule: safeStats.countRule || safeDailyCounts.countRule || '',
    dateBasis: safeStats.dateBasis || safeDailyCounts.dateBasis || '',
    manualOverrideActive: Boolean(
      safeStats.manualOverrideActive ||
      manualRecordCounts.manualOverrideActive
    ),
    manualOverrideNotes: safeStats.manualOverrideNotes || manualRecordCounts.manualOverrideNotes || []
  };
}

function getNullableNumber_(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function getManualRecordCounts_(manualCounts) {
  const result = {
    breakageCountToday: null,
    replenishmentCountToday: null,
    manualOverrideActive: false,
    manualOverrideNotes: []
  };

  if (!manualCounts || typeof manualCounts !== 'object') {
    return result;
  }

  if (manualCounts.breakageCountToday !== null && manualCounts.breakageCountToday !== undefined && manualCounts.breakageCountToday !== '') {
    result.breakageCountToday = num(manualCounts.breakageCountToday);
    result.manualOverrideActive = true;
    result.manualOverrideNotes.push(`Breakage Count Today manually set to ${result.breakageCountToday}`);
  }

  if (manualCounts.replenishmentCountToday !== null && manualCounts.replenishmentCountToday !== undefined && manualCounts.replenishmentCountToday !== '') {
    result.replenishmentCountToday = num(manualCounts.replenishmentCountToday);
    result.manualOverrideActive = true;
    result.manualOverrideNotes.push(`Replenishment Count Today manually set to ${result.replenishmentCountToday}`);
  }

  if (Array.isArray(manualCounts.rows)) {
    manualCounts.rows.forEach(row => {
      const metric = String(row.metric || row.Metric || '').toUpperCase();
      const count = num(row.manualCount ?? row.ManualCount ?? row.count ?? row.Count);

      if (metric.includes('BREAKAGE')) {
        result.breakageCountToday = count;
        result.manualOverrideActive = true;
      }

      if (metric.includes('REPLENISH')) {
        result.replenishmentCountToday = count;
        result.manualOverrideActive = true;
      }
    });
  }

  return result;
}

function getSafePositiveNumber_(...values) {
  for (const value of values) {
    const n = Number(value);
    if (Number.isFinite(n) && n > 0) return n;
  }
  return 0;
}

function getSafePickingWipTotal_(payload, wipRows) {
  return getSafePositiveNumber_(
    payload?.totalPickingWip,
    payload?.totalWip,
    sum((wipRows || []).filter(r => !isTrackingOnlyQueue(r.picking)), 'total')
  );
}

function getSafeInventoryWipTotal_(payload, wipRows) {
  return getSafePositiveNumber_(
    payload?.totalInventoryWip,
    sum(wipRows || [], 'total')
  );
}

function getSafeScanVerifyWipTotal_(payload, wipRows) {
  return getSafePositiveNumber_(
    payload?.scanVerifyWip,
    sum((wipRows || []).filter(r => isTrackingOnlyQueue(r.picking)), 'total')
  );
}

function normalizeOperatorActivity(rows) {
  return (Array.isArray(rows) ? rows : []).map(r => {
    const rawStation = r.flowStation || r.FlowStation || r.station || r.Station || r.accessPoint || r.AccessPoint || '';
    const displayStation = displayPickingStationName(rawStation);

    return {
      reportDate: r.reportDate || r.ReportDate || '',
      area: r.area || r.Area || '',
      rawFlowStation: rawStation,
      flowStation: displayStation,
      accessPoint: displayPickingStationName(r.accessPoint || r.AccessPoint || rawStation),
      operator: r.operator || r.Operator || 'Unknown',
      total: num(r.total ?? r.Total),
      hourlyTotal: num(r.hourlyTotal ?? r.HourlyTotal ?? r.total ?? r.Total),
      bestHour: r.bestHour || r.BestHour || '',
      bestHourValue: num(r.bestHourValue ?? r.BestHourValue),
      lastActiveHour: r.lastActiveHour || r.LastActiveHour || '',
      hours: r.hours || r.Hours || {}
    };
  }).filter(r => r.operator && r.operator !== 'Unknown');
}

function displayPickingStationName(name) {
  const text = String(name || '').trim().replace(/\s+/g, ' ');
  const key = text.toUpperCase().replace(/&/g, 'AND').replace(/[^A-Z0-9]/g, '');

  if (
    key === 'FSVSCANVERIFYFRAMEONLY' ||
    key === 'FSVSCANANDVERIFYFRAMEONLY' ||
    key.includes('FSVSCANVERIFYFRAMEONLY') ||
    key.includes('FSVSCANANDVERIFYFRAMEONLY')
  ) {
    return 'FSV Scan & Verify';
  }

  return text || '--';
}

function normalizeHourlyActivity(rows) {
  return (Array.isArray(rows) ? rows : []).map(r => ({
    hour: r.hour || r.Hour || '',
    sfActivity: num(r.sfActivity ?? r.SFActivity),
    combinedFsvFrameActivity: num(r.combinedFsvFrameActivity ?? r.CombinedFsvFrameActivity),
    totalActivity: num(r.totalActivity ?? r.TotalActivity)
  })).filter(r => r.hour);
}

function buildCompletionFromDailyStats(stats) {
  const d = normalizeDailyStats(stats);

  return [
    {
      picking: 'SF Scan & Verify',
      completed: d.sfProductivityToday,
      direction: 'completed',
      before: 0,
      after: 0,
      timestamp: new Date().toISOString()
    },
    {
      picking: 'FSV Scan & Verify',
      completed: d.fsvProductivityToday,
      direction: 'completed',
      before: 0,
      after: 0,
      timestamp: new Date().toISOString()
    },
    {
      picking: 'Breakage to Picking',
      completed: d.breakageCountToday,
      direction: 'record',
      before: 0,
      after: 0,
      timestamp: new Date().toISOString()
    },
    {
      picking: 'Replenishment',
      completed: d.replenishmentCountToday,
      direction: 'record',
      before: 0,
      after: 0,
      timestamp: new Date().toISOString()
    }
  ];
}

/* ──────────────────────────────────────────────────
   Main Render
────────────────────────────────────────────────── */

function renderAll() {
  const wip = STATE.wip;
  const completion = STATE.completion;
  const daily = STATE.dailyStats || {};

  const totalWip = STATE.totalWip || sum(wip, 'total');

  const sfProductive = num(daily.sfProductivityToday);
  const fsvProductive = num(daily.fsvProductivityToday);
  const frameOnlyProductive = num(daily.frameOnlyProductivityToday || daily.framePickingActivityToday);
  const breakageCount = num(daily.breakageCountToday);
  const replenishmentCount = num(daily.replenishmentCountToday);
  const totalProductive = num(daily.totalProductivityToday);

  const latest =
    STATE.health?.lastImportTimestamp ||
    latestTimestamp(wip, completion);

  const sorted = [...wip].sort((a, b) => b.total - a.total);
  const highest = sorted[0] || { picking: '--', total: 0 };

  setText('lastUpdated', latest ? toEastern(latest) : '--');
  setText('sideSystemText', STATE.demo ? 'Demo Mode Active' : 'All Systems Operational');

  animateNumber('kpiTotalWip', totalWip);
  animateNumber('kpiSFProductive', sfProductive);
  animateNumber('kpiFSVProductive', fsvProductive);
  animateNumber('kpiFrameOnlyProductive', frameOnlyProductive);
  animateNumber('kpiBreakageCount', breakageCount);
  animateNumber('kpiReplenishmentCount', replenishmentCount);
  renderWipSplitCards();
  renderCommandHeroMetrics(wip, daily, latest);

  const trendInfo = getOverallTrend();

  renderMovementCheck();
  renderTrackingRows(wip, daily);
  renderWipRows(wip);
  renderDistribution(wip, totalWip);
  renderQueueBars(wip);
  renderFocus(wip, totalWip, highest, totalProductive);
  renderQueuesTab(wip, totalWip);
  renderActivityTab();
  renderOperatorActivityTab();
  renderPersonalTab();
  renderReportsTab();
  updateTicker(wip);
}


function renderWipSplitCards() {
  const pickingWip = Number(STATE.totalPickingWip || STATE.totalWip || 0);
  const inventoryWip = Number(STATE.totalInventoryWip || sum(STATE.wip || [], 'total'));
  const scanVerifyWip = Number(STATE.scanVerifyWip || sum((STATE.wip || []).filter(r => isTrackingOnlyQueue(r.picking)), 'total'));

  animateNumber('kpiPickingWipSplit', pickingWip);
  animateNumber('kpiInventoryWip', inventoryWip);
  animateNumber('kpiScanVerifyWip', scanVerifyWip);
}

function renderCommandHeroMetrics(wip, daily, latest) {
  const safeWip = Array.isArray(wip) ? wip : [];

  setText('cmdRecordSource', daily.recordCountSource || daily.recordCountMode || 'API');
  setText('cmdCountRule', daily.countRule || 'BOS baseline · positive increases');
  setText('cmdLastSnapshot', daily.dateBasis || (latest ? 'Live snapshot' : 'Waiting'));

  animateNumber('flowOutSurface', getWipFor('Out of Surface Queue', safeWip));
  animateNumber('flowSfWip', getWipFor('SF Scan & Verify', safeWip));
  animateNumber('flowOutFinish', getWipFor('Out of Finish Queue', safeWip));
  animateNumber('flowFsvWip', getWipFor('FSV Scan & Verify', safeWip));
  animateNumber('flowFrameOnlyWip', getWipFor('Frame Only Scan & Verify', safeWip));
  animateNumber('flowBreakage', getWipFor('Breakage to Picking', safeWip));
  animateNumber('flowReplenishment', getWipFor('Replenishment', safeWip));
}


function renderMovementCheck() {
  const statusEl = document.getElementById('movementStatus');
  const msgEl = document.getElementById('movementMessage');
  const rowsEl = document.getElementById('movementRows');

  if (!statusEl || !msgEl || !rowsEl) return;

  const status = STATE.movementStatus || 'NORMAL';
  const message = STATE.movementMessage || 'Normal WIP movement.';
  const rows = STATE.movement || [];

  statusEl.textContent = status;
  statusEl.classList.remove('valid', 'warning', 'error');

  msgEl.textContent = message;
  msgEl.classList.remove('valid', 'warning', 'error');

  if (
    status === 'VALID_MOVEMENT' ||
    status === 'NORMAL' ||
    status === 'BASELINE'
  ) {
    statusEl.classList.add('valid');
    msgEl.classList.add('valid');
  } else if (status === 'WARNING_REVIEW') {
    statusEl.classList.add('warning');
    msgEl.classList.add('warning');
  } else if (status === 'DEMO' || status === 'ERROR') {
    statusEl.classList.add('error');
    msgEl.classList.add('error');
  }

  if (!rows.length) {
    rowsEl.innerHTML = `
      <div class="movement-row no-change">
        <div class="movement-station">
          <span class="movement-station-dot"></span>
          No movement comparison yet
        </div>
        <span class="movement-prev">--</span>
        <span class="movement-current">--</span>
        <span class="movement-diff flat">0</span>
        <span><b class="movement-type">WAITING</b></span>
      </div>
    `;
    return;
  }

  rowsEl.innerHTML = rows.map((r, i) => {
    const diff = Number(r.difference || 0);
    const rowClass = diff > 0 ? 'gain' : diff < 0 ? 'drop' : 'no-change';
    const diffClass = diff > 0 ? 'gain' : diff < 0 ? 'drop' : 'flat';
    const diffText = diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString();

    return `
      <div class="movement-row ${rowClass}" style="animation-delay:${i * 0.035}s">
        <div class="movement-station">
          <span class="movement-station-dot"></span>
          ${escHtml(r.subDepartment)}
        </div>
        <span class="movement-prev">${Number(r.previousTotal || 0).toLocaleString()}</span>
        <span class="movement-current">${Number(r.currentTotal || 0).toLocaleString()}</span>
        <span class="movement-diff ${diffClass}">${diffText}</span>
        <span><b class="movement-type">${escHtml(r.movementType || '')}</b></span>
      </div>
    `;
  }).join('');
}

/* ──────────────────────────────────────────────────
   Command Center Renderers
────────────────────────────────────────────────── */

function renderTrackingRows(wip, daily) {
  const metrics = [
    {
      name: 'SF Scan & Verify',
      label: 'SF Productive Today',
      count: num(daily.sfProductivityToday),
      current: getWipFor('SF Scan & Verify', wip),
      type: 'PRODUCTIVITY',
      status: 'Sent to Surface Unbox',
      cls: 'sf',
      short: 'SF'
    },
    {
      name: 'FSV Scan & Verify',
      label: 'FSV Productive Today',
      count: num(daily.fsvProductivityToday),
      current: getWipFor('FSV Scan & Verify', wip),
      type: 'PRODUCTIVITY',
      status: 'Sent to Finish Unbox',
      cls: 'fsv',
      short: 'FSV'
    },
    {
      name: 'Frame Only Scan & Verify',
      label: 'Frame Only Scan & Verify WIP',
      count: num(daily.frameOnlyScanVerifyWip || getWipFor('Frame Only Scan & Verify', wip)),
      current: getWipFor('Frame Only Scan & Verify', wip),
      type: 'WIP TRACKING',
      status: 'Frame Only WIP',
      cls: 'frame',
      short: 'FO'
    },
    {
      name: 'Breakage to Picking',
      label: 'Breakage Count Today',
      count: num(daily.breakageCountToday),
      current: getWipFor('Breakage to Picking', wip),
      type: 'RECORD',
      status: 'Drop count from Breakage',
      cls: 'brk',
      short: 'BR'
    },
    {
      name: 'Replenishment',
      label: 'Replenishment Count',
      count: num(daily.replenishmentCountToday),
      current: getWipFor('Replenishment', wip),
      type: 'RECORD',
      status: 'Drop count from Replenishment',
      cls: 'brk',
      short: 'RP'
    }
  ];

  const maxCount = Math.max(...metrics.map(m => m.count), 1);

  const html = metrics.map((m, i) => {
    const bar = Math.min((m.count / maxCount) * 100, 100);

    return `
      <div class="tracking-row" style="animation-delay:${i * .05}s">
        <div class="type-cell ${m.cls}">
          <span class="type-badge">${m.short}</span>
          ${escHtml(m.label)}
        </div>

        <div>
          <span class="count ${m.cls}">${m.count.toLocaleString()}</span>
          <span class="progress-mini">
            <i style="width:${bar}%"></i>
          </span>
        </div>

        <div class="count ${m.cls}">${m.current.toLocaleString()}</div>
        <div class="last-change">${escHtml(m.type)}</div>

        <div class="trend-down">
          ● ${escHtml(m.status)}
          <span class="trend-pulse"></span>
        </div>
      </div>
    `;
  }).join('');

  const el = document.getElementById('trackingRows');
  if (el) el.innerHTML = html;
}

function renderWipRows(wip) {
  const order = [...wip].sort((a, b) => {
    const at = TRACKED.includes(a.picking);
    const bt = TRACKED.includes(b.picking);

    if (at !== bt) {
      return bt - at;
    }

    return b.total - a.total;
  });

  const max = Math.max(...order.map(r => r.total), 1);

  setText('tableMeta', `${order.length} rows`);

  const rowsEl = document.getElementById('wipRows');
  if (!rowsEl) return;

  if (!order.length) {
    rowsEl.innerHTML = `
      <div class="wip-row">
        <div class="wip-name">No WIP rows loaded yet</div>
        <div class="dept">--</div>
        <div class="wip-count"><strong>0</strong></div>
        <div class="wip-date">--</div>
      </div>
    `;
    return;
  }

  rowsEl.innerHTML = order.map((r, i) => {
    const tracked = TRACKED.includes(r.picking);
    const rowClass = tracked ? 'blue-row' : 'orange-row';

    const icon =
      r.picking === 'SF Scan & Verify'
        ? 'SF'
        : r.picking === 'FSV Scan & Verify'
          ? 'FSV'
          : r.picking === 'Breakage to Picking'
            ? 'BR'
            : r.picking === 'Replenishment'
              ? 'RP'
              : getQueueIcon(r.picking);

    const pct = Math.max(2, Math.round(r.total / max * 100));

    return `
      <div class="wip-row ${rowClass}" style="animation-delay:${i * .035}s">
        <div class="wip-name">
          <span class="wip-icon">${icon}</span>
          ${escHtml(r.picking)}
        </div>

        <div class="dept">${escHtml(r.department || 'Inventory')}</div>

        <div class="wip-count">
          <strong>${r.total.toLocaleString()}</strong>
          <div class="row-bar">
            <i style="width:${pct}%"></i>
          </div>
        </div>

        <div class="wip-date">${toEastern(r.lastUpdated)}</div>
      </div>
    `;
  }).join('');
}

function renderDistribution(wip, total) {
  const donut = document.getElementById('donutChart');
  const legend = document.getElementById('wipLegend');

  if (!donut || !legend) return;

  const displayRows = [...wip]
    .filter(r => Number(r.total || 0) > 0)
    .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));

  const distributionTotal = displayRows.reduce((sum, r) => {
    return sum + Number(r.total || 0);
  }, 0);

  const top = displayRows.slice(0, 5);
  const used = sum(top, 'total');
  const others = Math.max(distributionTotal - used, 0);
  const segments = others ? [...top, { picking: 'Others', total: others }] : top;

  let start = 0;

  const conic = segments.map((r, i) => {
    const deg = distributionTotal ? (r.total / distributionTotal) * 360 : 0;
    const end = start + deg;
    const part = `${COLORS[i % COLORS.length]} ${start}deg ${end}deg`;

    start = end;

    return part;
  }).join(', ');

  donut.style.background = `conic-gradient(${conic || '#192638 0deg 360deg'})`;

  setText('donutTotal', distributionTotal.toLocaleString());

  legend.innerHTML = segments.map((r, i) => {
    const pct = distributionTotal ? (r.total / distributionTotal * 100).toFixed(1) : '0.0';

    return `
      <div class="leg-row">
        <span class="leg-dot" style="background:${COLORS[i % COLORS.length]};color:${COLORS[i % COLORS.length]}"></span>
        <span>${escHtml(r.picking)}</span>
        <b>${r.total.toLocaleString()} (${pct}%)</b>
      </div>
    `;
  }).join('');
}
function renderQueueBars(wip) {
  const el = document.getElementById('queueBars');
  if (!el) return;

  const queues = wip
    .filter(r => !TRACKED.includes(r.picking))
    .sort((a, b) => b.total - a.total)
    .slice(0, 6);

  const max = Math.max(...queues.map(q => q.total), 1);

  if (!queues.length) {
    el.innerHTML = '<div class="error-log-empty"><span>◌</span>No queue data available</div>';
    return;
  }

  el.innerHTML = queues.map(q => {
    const height = Math.max(4, Math.round(q.total / max * 100));

    return `
      <div class="bar-item">
        <div class="bar-val">${q.total.toLocaleString()}</div>
        <div class="bar-col" style="height:${height}%"></div>
        <div class="bar-label">${wrapLabel(q.picking)}</div>
      </div>
    `;
  }).join('');
}

function renderFocus(wip, total, highest, totalProductive) {
  const focus = highest || [...wip].sort((a, b) => b.total - a.total)[0];

  if (!focus) {
    setText('focusTitle', '--');
    setText('focusText', 'Waiting for live data...');
    return;
  }

  const pct = total ? (focus.total / total * 100).toFixed(1) : '0.0';

  setText('focusTitle', focus.picking);
  setText(
    'focusText',
    `Highest WIP at ${focus.total.toLocaleString()} items (${pct}% of total). SF + FSV productivity today: ${Number(totalProductive || 0).toLocaleString()}.`
  );
}

/* ──────────────────────────────────────────────────
   Queues Tab
────────────────────────────────────────────────── */

function renderQueuesTab(wip, total) {
  /*
   * Queue tab rule:
   * - Hide zero queues from cards/chart/table.
   * - Total Queues = only queues with total > 0.
   * - Empty Queues = how many source rows are zero.
   * - Official Total WIP stays from backend total, so SF/FSV Only
   *   can show without inflating Total WIP.
   */

  const allQueues = Array.isArray(wip) ? wip : [];

  const activeQueues = allQueues
    .filter(r => Number(r.total || 0) > 0)
    .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));

  const emptyQueues = allQueues.filter(r => Number(r.total || 0) === 0).length;

  const max = Math.max(...activeQueues.map(r => Number(r.total || 0)), 1);

  const largest = activeQueues[0] || {
    picking: '--',
    total: 0
  };

  animateNumber('qTabTotalQueues', activeQueues.length);
  animateNumber('qTabTotalWip', Number(total || 0));
  animateNumber('qTabLargest', Number(largest.total || 0));
  setText('qTabLargestName', largest.picking || '--');
  animateNumber('qTabEmpty', emptyQueues);

  const cardGrid = document.getElementById('queueCardGrid');

  if (cardGrid) {
    if (!activeQueues.length) {
      cardGrid.innerHTML = '<div class="error-log-empty"><span>◌</span>No active queue data available</div>';
    } else {
      cardGrid.innerHTML = activeQueues.map(r => {
        const tracked = isTrackingOnlyQueue(r.picking);
        const pct = Math.max(2, Math.round(Number(r.total || 0) / max * 100));
        const cls = tracked ? 'tracked' : '';

        return `
          <div class="q-card ${cls}">
            <div class="q-card-name">${escHtml(r.picking)}</div>
            <div class="q-card-count">${Number(r.total || 0).toLocaleString()}</div>
            <div class="q-card-bar"><i style="width:${pct}%"></i></div>
            <div class="q-card-dept">${escHtml(r.department || 'Inventory')}</div>
          </div>
        `;
      }).join('');
    }
  }

  const chartEl = document.getElementById('queueChartBars');

  if (chartEl) {
    setText('qChartMeta', `${activeQueues.length} queues`);

    chartEl.innerHTML = activeQueues.length
      ? activeQueues.map(r => {
          const tracked = isTrackingOnlyQueue(r.picking);
          const pct = Math.max(1, Math.round(Number(r.total || 0) / max * 100));

          return `
            <div class="q-bar-row ${tracked ? 'tracked' : ''}">
              <div class="q-bar-row-label">${escHtml(r.picking)}</div>
              <div class="q-bar-track"><i style="width:${pct}%"></i></div>
              <div class="q-bar-val">${Number(r.total || 0).toLocaleString()}</div>
            </div>
          `;
        }).join('')
      : '<div class="error-log-empty"><span>◌</span>No active queue chart data available</div>';
  }

  const tableEl = document.getElementById('queueTableRows');

  if (tableEl) {
    setText('qTableMeta', `${activeQueues.length} rows`);

    tableEl.innerHTML = activeQueues.length
      ? activeQueues.map((r, i) => {
          const tracked = isTrackingOnlyQueue(r.picking);
          const pct = Math.max(2, Math.round(Number(r.total || 0) / max * 100));
          const rowCls = tracked ? 'blue-row' : 'orange-row';

          const priority = tracked
            ? '<span class="priority-high">● TRACK</span>'
            : Number(r.total || 0) >= 20
              ? '<span class="priority-med">● MED</span>'
              : '<span class="priority-low">○ LOW</span>';

          const icon = tracked
            ? getTrackingIcon(r.picking)
            : getQueueIcon(r.picking);

          return `
            <div class="wip-row ${rowCls}" style="grid-template-columns:1.8fr .9fr 1.1fr .8fr 1.25fr;animation-delay:${i * .03}s">
              <span>
                <b class="queue-mini-icon">${icon}</b>
                ${escHtml(r.picking)}
              </span>
              <span>${escHtml(r.department || 'Inventory')}</span>
              <span>
                <b>${Number(r.total || 0).toLocaleString()}</b>
                <em class="row-bar"><i style="width:${pct}%"></i></em>
              </span>
              <span>${priority}</span>
              <span>${escHtml(r.lastUpdated || r.importedAt || '--')}</span>
            </div>
          `;
        }).join('')
      : '<div class="error-log-empty"><span>◌</span>No active queue detail available</div>';
  }
}

/* ──────────────────────────────────────────────────
   Activity Log Tab
────────────────────────────────────────────────── */

function renderActivityTab() {
  const daily = STATE.dailyStats || {};
  const latest =
    STATE.health?.lastImportTimestamp ||
    new Date().toLocaleString();

  const sf = num(daily.sfProductivityToday);
  const fsv = num(daily.fsvProductivityToday);
  const brk = num(daily.breakageCountToday);
  const rep = num(daily.replenishmentCountToday);

  animateNumber('actSFProductive', sf);
  animateNumber('actFSVProductive', fsv);
  animateNumber('actBreakageCount', brk);
  animateNumber('actReplenishmentCount', rep);

  const rows = [
    {
      metric: 'SF Productive Today',
      type: 'PRODUCTIVITY',
      count: sf,
      source: 'Gain into SF Scan & Verify',
      status: 'Active',
      value: sf
    },
    {
      metric: 'FSV Productive Today',
      type: 'PRODUCTIVITY',
      count: fsv,
      source: 'Gain into FSV Scan & Verify',
      status: 'Active',
      value: fsv
    },
    {
      metric: 'Breakage Count Today',
      type: 'RECORD',
      count: brk,
      source: 'Drop from Breakage to Picking',
      status: 'Tracked',
      value: brk
    },
    {
      metric: 'Replenishment Count Today',
      type: 'RECORD',
      count: rep,
      source: 'Drop from Replenishment',
      status: 'Tracked',
      value: rep
    }
  ];

  const logEl = document.getElementById('actLogRows');
  if (!logEl) return;

  setText('actLogMeta', `${rows.length} metrics`);

  logEl.innerHTML = rows.map((r, i) => {
    const typeCls = r.type === 'PRODUCTIVITY' ? 'act-dir-completed' : 'act-dir-added';

    return `
      <div class="act-row" style="animation-delay:${Math.min(i, 25) * 0.04}s">
        <div class="act-type other">${escHtml(r.metric)}</div>
        <div class="${typeCls}">${escHtml(r.type)}</div>
        <div class="act-num">${r.count.toLocaleString()}</div>
        <div class="act-num">${escHtml(r.source)}</div>
        <div class="${typeCls}">${escHtml(r.status)}</div>
        <div class="act-ts">${toEastern(latest)}</div>
        <div class="act-num">${r.value.toLocaleString()}</div>
      </div>
    `;
  }).join('');
}


/* ──────────────────────────────────────────────────
   Operator Activity Tab
────────────────────────────────────────────────── */

function renderOperatorActivityTab() {
  const rows = Array.isArray(STATE.activityByOperator) ? STATE.activityByOperator : [];

  syncOperatorStationFilter(rows);
  renderPickingJphConfigPanel();
  renderOperatorStationScorecards(rows);

  const searchEl = document.getElementById('operatorSearch');
  const stationEl = document.getElementById('operatorStationFilter');
  const sortEl = document.getElementById('operatorSortMode');

  const search = String(searchEl?.value || '').trim().toLowerCase();
  const station = String(stationEl?.value || OP_STATION_VIEW || 'all');
  const sortMode = String(sortEl?.value || 'total');

  OP_STATION_VIEW = station || 'all';

  const filtered = rows.filter(r => {
    const displayStation = displayPickingStationName(r.flowStation || r.accessPoint || '');
    const operatorName = String(r.operator || '').toLowerCase();
    const matchesSearch = !search || operatorName.includes(search);
    const matchesStation = station === 'all' || displayStation === station;
    return matchesSearch && matchesStation;
  });

  const sortedRows = sortOperatorRows(filtered, sortMode);
  const totalJobs = filtered.reduce((s, r) => s + num(r.total), 0);
  const totalOperators = new Set(filtered.map(r => String(r.operator || '').trim()).filter(Boolean)).size;
  const top = [...filtered].sort((a, b) => num(b.total) - num(a.total))[0] || { operator: '--', total: 0, flowStation: '--' };
  const peak = getPeakOperatorHour(null, filtered);
  const avg = totalOperators ? Math.round(totalJobs / totalOperators) : 0;

  animateNumber('opTotalOperators', totalOperators);
  animateNumber('opTotalJobs', totalJobs);
  animateNumber('opTopOperatorTotal', num(top.total));
  setText('opTopOperatorName', top.operator || '--');
  animateNumber('opPeakHourTotal', num(peak.total));
  setText('opPeakHourName', peak.hour || '--');
  animateNumber('opAvgPerOperator', avg);
  setText('operatorHeroStatus', rows.length ? 'ONLINE' : 'WAITING');

  const note = document.getElementById('operatorDataNote');
  if (note) {
    const scope = station === 'all' ? 'all Scan & Verify activity' : station;
    note.textContent = rows.length
      ? `${filtered.length} visible row(s) for ${scope}. Click an operator row to open hourly detail.`
      : 'No operator activity returned by the API yet.';
  }

  const meta = document.getElementById('operatorTableMeta');
  if (meta) meta.textContent = `${filtered.length} rows`;

  renderOperatorLeaderboard(filtered);
  renderOperatorStationMix(rows);

  const visibleKeys = new Set(filtered.map(r => getOperatorRowKey(r)));
  if (OP_SELECTED_OPERATOR_KEY && !visibleKeys.has(OP_SELECTED_OPERATOR_KEY)) {
    OP_SELECTED_OPERATOR_KEY = '';
  }

  const selected = OP_SELECTED_OPERATOR_KEY
    ? filtered.find(r => getOperatorRowKey(r) === OP_SELECTED_OPERATOR_KEY)
    : null;

  renderSelectedOperatorDetail(selected, filtered);

  const body = document.getElementById('operatorActivityRows');
  if (body) {
    if (!filtered.length) {
      body.innerHTML = '<div class="op-empty">No operator activity matches the current scorecard/filter.</div>';
    } else {
      const maxTotal = Math.max(...filtered.map(r => num(r.total)), 1);
      body.innerHTML = sortedRows
        .map((r, i) => renderOperatorRow(r, maxTotal, i))
        .join('');
    }
  }

  // Hourly performance summary was moved to the Personal tab.
  // Operator Activity now focuses on scorecards, leaderboard, table, and selected-operator drilldown.

  const overlay = document.getElementById('operatorStationDetailOverlay');
  if (overlay && overlay.classList.contains('open')) {
    renderOperatorStationDetailModal();
  }
}

function sortOperatorRows(rows, sortMode) {
  const copy = [...rows];

  if (sortMode === 'operator') {
    return copy.sort((a, b) => String(a.operator || '').localeCompare(String(b.operator || '')) || num(b.total) - num(a.total));
  }

  if (sortMode === 'station') {
    return copy.sort((a, b) => String(a.flowStation || '').localeCompare(String(b.flowStation || '')) || num(b.total) - num(a.total));
  }

  if (sortMode === 'bestHour') {
    return copy.sort((a, b) => num(b.bestHourValue) - num(a.bestHourValue) || num(b.total) - num(a.total));
  }

  return copy.sort((a, b) => num(b.total) - num(a.total));
}

function syncOperatorStationFilter(rows) {
  const el = document.getElementById('operatorStationFilter');
  if (!el) return;

  const current = OP_STATION_VIEW || el.value || 'all';
  const stations = [...new Set(rows.map(r => displayPickingStationName(r.flowStation || r.accessPoint || '')).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));

  const nextHtml = ['<option value="all">All Stations</option>']
    .concat(stations.map(st => `<option value="${escHtml(st)}">${escHtml(st)}</option>`))
    .join('');

  if (el.dataset.optionsHtml !== nextHtml) {
    el.innerHTML = nextHtml;
    el.dataset.optionsHtml = nextHtml;
  }

  el.value = stations.includes(current) ? current : 'all';
  OP_STATION_VIEW = el.value;
}

function setOperatorStationView(station) {
  OP_STATION_VIEW = station || 'all';
  OP_SELECTED_OPERATOR_KEY = '';

  const stationEl = document.getElementById('operatorStationFilter');
  if (stationEl) {
    stationEl.value = OP_STATION_VIEW;
  }

  renderOperatorActivityTab();
}

function renderOperatorStationScorecards(rows) {
  const el = document.getElementById('operatorStationScorecards');
  const meta = document.getElementById('operatorScorecardMeta');
  if (!el) return;

  const stationDefs = [
    {
      label: 'SF Scan & Verify',
      key: 'SF Scan & Verify',
      type: 'sf',
      sub: 'Lens scan activity from Surface flow'
    },
    {
      label: 'FSV Scan & Verify',
      key: 'FSV Scan & Verify',
      type: 'fsv',
      sub: 'Finish scan activity shown without Frame Only label'
    }
  ];

  const cards = stationDefs.map(def => {
    const stationRows = rows.filter(r => displayPickingStationName(r.flowStation || r.accessPoint || '') === def.key);
    const total = stationRows.reduce((s, r) => s + num(r.total), 0);
    const operators = new Set(stationRows.map(r => String(r.operator || '').trim()).filter(Boolean));
    const top = [...stationRows].sort((a, b) => num(b.total) - num(a.total))[0] || { operator: '--', total: 0 };
    const peak = getPeakOperatorHour(null, stationRows);
    const targetText = getPickingStationTargetText(def.key);
    const active = OP_STATION_VIEW === def.key;

    return `
      <article class="op-big-scorecard ${def.type} ${active ? 'active' : ''}" onclick="openOperatorStationDetail('${escAttr(def.key)}')" role="button" tabindex="0">
        <div class="op-big-score-top">
          <span class="op-big-score-icon">${def.type === 'sf' ? 'SF' : 'FSV'}</span>
          <div>
            <h3>${escHtml(def.label)}</h3>
            <p>${escHtml(def.sub)}</p>
          </div>
        </div>

        <div class="op-big-score-main">
          <strong>${total.toLocaleString()}</strong>
          <span>Total Output Today</span>
        </div>

        <div class="op-big-score-grid">
          <div><span>Operators</span><b>${operators.size}</b></div>
          <div><span>Top Operator</span><b>${escHtml(top.operator || '--')}</b></div>
          <div><span>Top Output</span><b>${num(top.total).toLocaleString()}</b></div>
          <div><span>Peak Hour</span><b>${escHtml(peak.hour || '--')}</b></div>
          <div class="target-wide"><span>JPH Target</span><b>${escHtml(targetText)}</b></div>
        </div>

        <div class="op-big-score-footer">
          <span>${active ? 'Focused View Active' : 'Open station drilldown'}</span>
          <button type="button" onclick="event.stopPropagation();openOperatorStationDetail('${escAttr(def.key)}')">Open Detail</button>
        </div>
      </article>
    `;
  }).join('');

  el.innerHTML = cards;

  if (meta) {
    meta.textContent = OP_STATION_VIEW === 'all'
      ? 'Showing all operators. Click SF or FSV to focus.'
      : `Focused on ${OP_STATION_VIEW}.`;
  }
}


function openOperatorStationDetail(station) {
  OP_DETAIL_STATION = displayPickingStationName(station || '');
  OP_STATION_VIEW = OP_DETAIL_STATION || 'all';
  OP_SELECTED_OPERATOR_KEY = '';

  const stationEl = document.getElementById('operatorStationFilter');
  if (stationEl) stationEl.value = OP_STATION_VIEW;

  renderOperatorActivityTab();
  renderOperatorStationDetailModal();

  const overlay = document.getElementById('operatorStationDetailOverlay');
  if (overlay) {
    overlay.classList.add('open');
    overlay.setAttribute('aria-hidden', 'false');
    document.body.classList.add('modal-open');
  }
}

function closeOperatorStationDetail() {
  const overlay = document.getElementById('operatorStationDetailOverlay');
  if (overlay) {
    overlay.classList.remove('open');
    overlay.setAttribute('aria-hidden', 'true');
  }
  document.body.classList.remove('modal-open');
}

function renderOperatorStationDetailModal() {
  if (!OP_DETAIL_STATION) return;

  const allRows = Array.isArray(STATE.activityByOperator) ? STATE.activityByOperator : [];
  const stationRows = allRows
    .filter(r => displayPickingStationName(r.flowStation || r.accessPoint || '') === OP_DETAIL_STATION)
    .sort((a, b) => num(b.total) - num(a.total));

  const title = document.getElementById('opStationDetailTitle');
  const sub = document.getElementById('opStationDetailSub');
  const scope = document.getElementById('opStationDetailScope');

  if (title) title.textContent = OP_DETAIL_STATION;
  if (sub) sub.textContent = `${OP_DETAIL_STATION} · Picking Scan & Verify · ${getPickingStationTargetText(OP_DETAIL_STATION)}`;
  if (scope) scope.textContent = `${stationRows.length} rows`;

  const total = stationRows.reduce((s, r) => s + num(r.total), 0);
  const operators = new Set(stationRows.map(r => String(r.operator || '').trim()).filter(Boolean));
  const peak = getPeakOperatorHour(null, stationRows);
  const top = stationRows[0] || { operator: '--', total: 0 };

  setText('opDetailTotalOutput', total.toLocaleString());
  setText('opDetailOperators', operators.size.toLocaleString());
  setText('opDetailPeakHour', peak.hour || '--');
  setText('opDetailPeakCount', `${num(peak.total).toLocaleString()} jobs`);
  setText('opDetailTopOperator', top.operator || '--');
  setText('opDetailTopOperatorCount', `${num(top.total).toLocaleString()} jobs`);

  syncOperatorDetailFilter(stationRows);

  const filterEl = document.getElementById('opDetailOperatorFilter');
  const selectedOperator = String(filterEl?.value || 'all');
  const visibleRows = selectedOperator === 'all'
    ? stationRows
    : stationRows.filter(r => getOperatorRowKey(r) === selectedOperator);

  const rowsEl = document.getElementById('opDetailOperatorRows');
  if (!rowsEl) return;

  if (!visibleRows.length) {
    rowsEl.innerHTML = `
      <div class="op-empty">
        No operator output available for ${escHtml(OP_DETAIL_STATION)}.
      </div>
    `;
    return;
  }

  rowsEl.innerHTML = visibleRows.map(row => renderOperatorStationDetailCard(row)).join('') + `
    <div class="op-detail-footnote">
      Hour colors use the Picking JPH configuration. Green means 100% or better, amber means 90% to 99%, and red means below 90%. 6 AM and 6 PM–8 PM are treated as no-target hours until shift rules are expanded.
    </div>
  `;
}

function syncOperatorDetailFilter(rows) {
  const el = document.getElementById('opDetailOperatorFilter');
  if (!el) return;

  const current = el.value || 'all';
  const options = ['<option value="all">All operators</option>']
    .concat(rows.map(r => {
      const key = getOperatorRowKey(r);
      return `<option value="${escAttr(key)}">${escHtml(r.operator || '--')} · ${num(r.total).toLocaleString()}</option>`;
    }))
    .join('');

  if (el.dataset.optionsHtml !== options) {
    el.innerHTML = options;
    el.dataset.optionsHtml = options;
  }

  const values = new Set(['all'].concat(rows.map(r => getOperatorRowKey(r))));
  el.value = values.has(current) ? current : 'all';
}

function renderOperatorStationDetailCard(row) {
  const total = num(row.total);
  const hours = row.hours || row.Hours || {};
  const station = displayPickingStationName(row.flowStation || row.accessPoint || OP_DETAIL_STATION || '--');
  const hourEntries = normalizeOperatorHourEntries(hours);
  const peak = hourEntries.reduce((best, cur) => cur.value > best.value ? cur : best, { hour: '--', value: 0 });
  const activeHours = hourEntries.filter(h => h.value > 0).length;
  const status = getOperatorStatus(total);
  const avgPerActiveHour = activeHours ? Math.round(total / activeHours) : 0;
  const targetCfg = getStationJphConfig(station);
  const targetText = getPickingStationTargetText(station);
  const currentTargetHours = hourEntries.filter(h => isPickingJphTargetHour(h.hour));
  const greenHours = currentTargetHours.filter(h => getPickingJphResult(h.value, station, h.hour).status === 'GREEN').length;
  const amberHours = currentTargetHours.filter(h => getPickingJphResult(h.value, station, h.hour).status === 'AMBER').length;
  const redHours = currentTargetHours.filter(h => getPickingJphResult(h.value, station, h.hour).status === 'RED').length;

  return `
    <article class="op-detail-operator-card ${status.cls || ''}">
      <div class="op-detail-operator-top">
        <div class="op-detail-person">
          <span class="op-detail-avatar">${escHtml(getOperatorInitials(row.operator))}</span>
          <div class="op-detail-name">
            <strong>${escHtml(row.operator || '--')}</strong>
            <small>${escHtml(station)} · ${escHtml(row.area || 'Picking')} · Avg ${avgPerActiveHour.toLocaleString()}/active hr</small>
          </div>
          <span class="op-detail-jph-chip">Target ${targetCfg.target.toLocaleString()}/hr</span>
          <span class="op-detail-jph-chip neutral">${greenHours}G · ${amberHours}A · ${redHours}R</span>
        </div>
        <div class="op-detail-total">
          <span>Total Output</span>
          <strong>${total.toLocaleString()}</strong>
        </div>
      </div>

      <div class="op-target-note">${escHtml(targetText)}</div>

      <div class="op-detail-hour-grid">
        ${hourEntries.map(h => {
          const result = getPickingJphResult(h.value, station, h.hour);
          const pctHeight = result.target
            ? Math.max(h.value > 0 ? 8 : 3, Math.min(100, Math.round((h.value / result.target) * 100)))
            : Math.max(h.value > 0 ? 8 : 3, Math.min(100, Math.round((h.value / Math.max(peak.value, 1)) * 100)));

          return `
            <div class="op-detail-hour-cell ${result.cls}" title="${escHtml(h.hour)} · ${h.value.toLocaleString()} jobs · ${escHtml(result.targetLabel)} · ${escHtml(result.pctText)}">
              <span class="hr">${escHtml(shortHourLabel(h.hour))}</span>
              <span class="val">${h.value.toLocaleString()}</span>
              <div class="op-detail-hour-bar"><i style="height:${pctHeight}%"></i></div>
              <span class="target">${escHtml(result.targetLabel)}</span>
              <span class="pct">${escHtml(result.pctText)}</span>
            </div>
          `;
        }).join('')}
      </div>
    </article>
  `;
}

function normalizeOperatorHourEntries(hours) {
  const preferred = [
    '6:00 AM','7:00 AM','8:00 AM','9:00 AM','10:00 AM','11:00 AM','12:00 PM',
    '1:00 PM','2:00 PM','3:00 PM','4:00 PM','5:00 PM','6:00 PM','7:00 PM','8:00 PM'
  ];

  const map = {};
  Object.entries(hours || {}).forEach(([hour, value]) => {
    map[String(hour)] = num(value);
  });

  return preferred.map(hour => ({ hour, value: num(map[hour]) }));
}

function shortHourLabel(hour) {
  return String(hour || '')
    .replace(':00 ', '')
    .replace(' AM', 'A')
    .replace(' PM', 'P');
}

function getOperatorHourClass(value, stationOrMaxValue, hour) {
  if (typeof stationOrMaxValue === 'string') {
    return getPickingJphResult(value, stationOrMaxValue, hour).cls;
  }

  const v = num(value);
  const maxValue = num(stationOrMaxValue);
  if (v <= 0) return 'neutral';
  const ratio = maxValue ? v / maxValue : 0;
  if (ratio >= .65) return 'good';
  if (ratio >= .35) return 'warn';
  return 'bad';
}

function getCurrentOperatorHourValue(entries) {
  const currentHour = new Date().toLocaleString('en-US', {
    timeZone: 'America/New_York',
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  }).replace(/^0/, '');

  const found = (entries || []).find(e => String(e.hour) === currentHour);
  return found ? found.value : 0;
}

function escAttr(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/'/g, '&#39;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function renderOperatorLeaderboard(rows) {
  const el = document.getElementById('operatorLeaderboardCards');
  const meta = document.getElementById('operatorLeaderboardMeta');
  if (!el) return;

  const top = [...rows].sort((a, b) => num(b.total) - num(a.total)).slice(0, 3);
  const max = Math.max(...top.map(r => num(r.total)), 1);

  if (meta) meta.textContent = top.length ? `Top ${top.length} by output` : 'Waiting for data';

  if (!top.length) {
    el.innerHTML = '<div class="op-empty">No leaderboard data available.</div>';
    return;
  }

  el.innerHTML = top.map((r, i) => {
    const pct = Math.max(4, Math.round((num(r.total) / max) * 100));
    const station = displayPickingStationName(r.flowStation || r.accessPoint);
    return `
      <article class="op-leader-card">
        <div class="op-rank">#${i + 1}</div>
        <div class="op-leader-top">
          <span class="op-leader-avatar">${escHtml(getOperatorInitials(r.operator))}</span>
          <div>
            <div class="op-leader-name">${escHtml(r.operator)}</div>
            <div class="op-leader-station">${escHtml(station)}</div>
          </div>
        </div>
        <div class="op-leader-metric">
          <strong>${num(r.total).toLocaleString()}</strong>
          <span>jobs today</span>
        </div>
        <div class="op-leader-spark"><i style="width:${pct}%"></i></div>
      </article>
    `;
  }).join('');
}

function renderOperatorStationMix(rows) {
  const el = document.getElementById('operatorStationMix');
  const meta = document.getElementById('operatorStationMeta');
  if (!el) return;

  const map = {};
  rows.forEach(r => {
    const station = displayPickingStationName(r.flowStation || r.accessPoint || 'Unknown');
    if (!map[station]) map[station] = { station, total: 0, operators: new Set() };
    map[station].total += num(r.total);
    if (r.operator) map[station].operators.add(r.operator);
  });

  const stations = Object.values(map).sort((a, b) => b.total - a.total);
  const max = Math.max(...stations.map(s => s.total), 1);

  if (meta) meta.textContent = stations.length ? `${stations.length} station(s)` : 'No station mix';

  if (!stations.length) {
    el.innerHTML = '<div class="op-empty">No station activity available.</div>';
    return;
  }

  el.innerHTML = stations.map(s => {
    const pct = Math.max(3, Math.round((s.total / max) * 100));
    return `
      <div class="op-station-card">
        <div class="op-station-top">
          <strong>${escHtml(s.station)}</strong>
          <span>${s.total.toLocaleString()}</span>
        </div>
        <div class="op-station-bar"><i style="width:${pct}%"></i></div>
        <small>${s.operators.size} active operator(s)</small>
      </div>
    `;
  }).join('');
}

function renderOperatorRow(r, maxTotal, index) {
  const total = num(r.total);
  const initials = getOperatorInitials(r.operator);
  const bars = buildOperatorHourBars(r.hours || {}, total || maxTotal);
  const station = displayPickingStationName(r.flowStation || r.accessPoint || '--');
  const stationCls = station.toUpperCase().includes('FSV') ? 'fsv' : '';
  const status = getOperatorStatus(total);
  const key = getOperatorRowKey(r);
  const selected = key === OP_SELECTED_OPERATOR_KEY;
  const encodedKey = encodeURIComponent(key);

  return `
    <div class="operator-row ${selected ? 'selected' : ''}" style="animation-delay:${Math.min(index, 20) * .035}s" onclick="selectOperatorDetail('${encodedKey}')" role="button" tabindex="0" title="Click to view hourly output for ${escHtml(r.operator || '')}">
      <div class="op-name">
        <span class="op-avatar">${escHtml(initials)}</span>
        <span class="op-name-text">
          <strong>${escHtml(r.operator)}</strong>
          <small>${escHtml(r.area || 'Picking')}</small>
        </span>
      </div>
      <div><span class="op-station-pill ${stationCls}">${escHtml(station)}</span></div>
      <div class="op-total">${total.toLocaleString()}</div>
      <div class="op-best"><span class="op-best-badge">${escHtml(r.bestHour || '--')}${r.bestHourValue ? ` · ${num(r.bestHourValue).toLocaleString()}` : ''}</span></div>
      <div class="op-last">${escHtml(r.lastActiveHour || '--')}</div>
      <div class="op-hour-bars" title="Hourly activity pattern">${bars}</div>
      <div><span class="op-status-chip ${status.cls}">${escHtml(status.label)}</span></div>
    </div>
  `;
}

function selectOperatorDetail(encodedKey) {
  OP_SELECTED_OPERATOR_KEY = decodeURIComponent(String(encodedKey || ''));
  renderOperatorActivityTab();
}

function clearSelectedOperatorDetail() {
  OP_SELECTED_OPERATOR_KEY = '';
  renderOperatorActivityTab();
}

function getOperatorRowKey(row) {
  return [
    String(row?.operator || '').trim(),
    displayPickingStationName(row?.flowStation || row?.accessPoint || ''),
    String(row?.area || '').trim()
  ].join('||');
}

function renderSelectedOperatorDetail(selected, visibleRows) {
  const panel = document.getElementById('operatorSelectedPanel');
  if (!panel) return;

  if (!selected) {
    const top = [...(visibleRows || [])].sort((a, b) => num(b.total) - num(a.total))[0];

    panel.innerHTML = `
      <div class="op-selected-empty">
        <span>👤</span>
        <strong>Select an operator</strong>
        <small>Click any operator name/row to view their hourly output by hour.</small>
        ${top ? `<button type="button" onclick="selectOperatorDetail('${encodeURIComponent(getOperatorRowKey(top))}')">Open top operator ›</button>` : ''}
      </div>
    `;
    return;
  }

  const hours = selected.hours || selected.Hours || {};
  const entries = Object.entries(hours).map(([hour, value]) => ({ hour, value: num(value) }));
  const total = num(selected.total);
  const station = displayPickingStationName(selected.flowStation || selected.accessPoint || '--');
  const max = Math.max(...entries.map(e => e.value), 1);
  const activeHours = entries.filter(e => e.value > 0).length;
  const peak = entries.reduce((best, cur) => cur.value > best.value ? cur : best, { hour: '--', value: 0 });

  panel.innerHTML = `
    <div class="op-selected-head">
      <div class="op-selected-person">
        <span class="op-selected-avatar">${escHtml(getOperatorInitials(selected.operator))}</span>
        <div>
          <h3>${escHtml(selected.operator || '--')}</h3>
          <p>${escHtml(station)} · ${escHtml(selected.area || 'Picking')}</p>
        </div>
      </div>
      <button class="op-selected-close" type="button" onclick="clearSelectedOperatorDetail()">Clear</button>
    </div>

    <div class="op-selected-kpis">
      <div><span>Total Output</span><strong>${total.toLocaleString()}</strong></div>
      <div><span>Peak Hour</span><strong>${escHtml(peak.hour)}</strong></div>
      <div><span>Peak Output</span><strong>${peak.value.toLocaleString()}</strong></div>
      <div><span>Target</span><strong>${getStationJphConfig(station).target}/hr</strong></div>
    </div>

    <div class="op-selected-hours">
      ${entries.length ? entries.map(e => {
        const result = getPickingJphResult(e.value, station, e.hour);
        const pct = result.target
          ? Math.max(e.value > 0 ? 4 : 1, Math.min(100, Math.round((e.value / result.target) * 100)))
          : Math.max(e.value > 0 ? 4 : 1, Math.round((e.value / max) * 100));
        return `
          <div class="op-selected-hour ${result.cls} ${e.value === peak.value && e.value > 0 ? 'peak' : ''}" title="${escHtml(result.targetLabel)} · ${escHtml(result.pctText)}">
            <span>${escHtml(e.hour)}</span>
            <div><i style="width:${pct}%"></i></div>
            <b>${e.value.toLocaleString()} <em>${escHtml(result.pctText)}</em></b>
          </div>
        `;
      }).join('') : '<div class="op-empty">No hourly output available for this operator.</div>'}
    </div>
  `;
}

function getOperatorStatus(total) {
  const value = num(total);
  if (value >= 250) return { label: 'High', cls: '' };
  if (value >= 75) return { label: 'Active', cls: '' };
  if (value > 0) return { label: 'Low', cls: 'low' };
  return { label: 'Idle', cls: 'idle' };
}

function buildOperatorHourBars(hours, maxValue) {
  const entries = Object.entries(hours || {});
  const max = Math.max(...entries.map(([, v]) => num(v)), num(maxValue), 1);

  if (!entries.length) {
    return '<i style="height:2px"></i>';
  }

  return entries.map(([, value]) => {
    const h = Math.max(2, Math.round((num(value) / max) * 32));
    return `<i style="height:${h}px"></i>`;
  }).join('');
}

function renderOperatorHourlyRows(hourly, operatorRows) {
  const el = document.getElementById('operatorHourlyRows');
  if (!el) return;

  const meta = document.getElementById('operatorHourlyMeta');
  const safeOperatorRows = Array.isArray(operatorRows) ? operatorRows : [];
  let rows = Array.isArray(hourly) ? hourly.filter(Boolean) : [];

  if (!rows.length && safeOperatorRows.length) {
    const bucket = {};

    safeOperatorRows.forEach(op => {
      Object.entries(op.hours || op.Hours || {}).forEach(([hour, value]) => {
        if (!bucket[hour]) {
          bucket[hour] = {
            hour,
            sfActivity: 0,
            combinedFsvFrameActivity: 0,
            totalActivity: 0
          };
        }

        const val = num(value);
        bucket[hour].totalActivity += val;

        const station = String(op.flowStation || op.FlowStation || op.station || '').toUpperCase();

        if (station.includes('SF')) {
          bucket[hour].sfActivity += val;
        }

        if (station.includes('FSV')) {
          bucket[hour].combinedFsvFrameActivity += val;
        }
      });
    });

    rows = Object.values(bucket);
  }

  if (meta) {
    meta.textContent = `${rows.length} hour bucket(s) · SF cyan / FSV orange`;
  }

  renderHourlySplitRows(rows, 'operatorHourlyRows', 'operatorHourlySummary');
}

function getPeakOperatorHour(hourly, operatorRows) {
  if (hourly && hourly.length) {
    return hourly.reduce((best, row) => {
      const total = num(row.totalActivity);
      return total > num(best.total) ? { hour: row.hour, total } : best;
    }, { hour: '--', total: 0 });
  }

  const bucket = {};
  (operatorRows || []).forEach(op => {
    Object.entries(op.hours || op.Hours || {}).forEach(([hour, value]) => {
      bucket[hour] = (bucket[hour] || 0) + num(value);
    });
  });

  return Object.keys(bucket).reduce((best, hour) => {
    return bucket[hour] > best.total ? { hour, total: bucket[hour] } : best;
  }, { hour: '--', total: 0 });
}

function getOperatorInitials(name) {
  const parts = String(name || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (!parts.length) return '--';
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}


/* ──────────────────────────────────────────────────
   Personal Tab — JPH Config + Hourly Performance
────────────────────────────────────────────────── */

function renderPersonalTab() {
  renderPickingJphConfigPanel();
  renderPersonalHourlyPerformance();
}

function setPersonalHourlyStation(station) {
  PERSONAL_HOURLY_VIEW = station || 'all';

  document.querySelectorAll('[data-personal-hourly-filter]').forEach(btn => {
    btn.classList.toggle('active', String(btn.dataset.personalHourlyFilter || 'all') === PERSONAL_HOURLY_VIEW);
  });

  renderPersonalHourlyPerformance();
}

function getPersonalHourlyRows() {
  const raw = Array.isArray(STATE.hourlyActivity) ? STATE.hourlyActivity : [];
  const view = PERSONAL_HOURLY_VIEW || 'all';

  return raw.map(row => {
    const hour = row.hour || row.Hour || row.label || '--';
    const sf = num(row.sfActivity ?? row.SFActivity ?? row.sf ?? 0);
    const fsv = num(row.combinedFsvFrameActivity ?? row.fsvActivity ?? row.FSVActivity ?? row.fsv ?? 0);

    if (view === 'SF Scan & Verify') {
      return { hour, sfActivity: sf, combinedFsvFrameActivity: 0, totalActivity: sf };
    }

    if (view === 'FSV Scan & Verify') {
      return { hour, sfActivity: 0, combinedFsvFrameActivity: fsv, totalActivity: fsv };
    }

    return { hour, sfActivity: sf, combinedFsvFrameActivity: fsv, totalActivity: sf + fsv };
  });
}

function renderPersonalHourlyPerformance() {
  const el = document.getElementById('personalHourlyRows');
  if (!el) return;

  document.querySelectorAll('[data-personal-hourly-filter]').forEach(btn => {
    btn.classList.toggle('active', String(btn.dataset.personalHourlyFilter || 'all') === (PERSONAL_HOURLY_VIEW || 'all'));
  });

  const rows = getPersonalHourlyRows();
  const meta = document.getElementById('personalHourlyMeta');
  const view = PERSONAL_HOURLY_VIEW || 'all';

  if (meta) {
    meta.textContent = view === 'all'
      ? 'SF and FSV shown in separate colors'
      : `${view} hourly performance`;
  }

  renderHourlySplitRows(rows, 'personalHourlyRows', 'personalHourlySummary');
}

function renderHourlySplitRows(rows, rowsId, summaryId) {
  const el = document.getElementById(rowsId);
  const summary = document.getElementById(summaryId);
  if (!el) return;

  const safeRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
  const normalized = safeRows.map(row => {
    const sf = num(row.sfActivity ?? row.SFActivity ?? row.sf ?? 0);
    const fsv = num(row.combinedFsvFrameActivity ?? row.fsvActivity ?? row.FSVActivity ?? row.fsv ?? 0);
    const total = num(row.totalActivity ?? row.ActivityToday ?? row.total ?? row.Total ?? (sf + fsv));

    return {
      hour: row.hour || row.Hour || row.label || '--',
      sfActivity: sf,
      combinedFsvFrameActivity: fsv,
      totalActivity: total
    };
  });

  const max = Math.max(...normalized.map(r => num(r.totalActivity)), 1);
  const peak = normalized.reduce((best, row) => {
    return num(row.totalActivity) > num(best.totalActivity) ? row : best;
  }, { hour: '--', totalActivity: 0, sfActivity: 0, combinedFsvFrameActivity: 0 });

  const total = normalized.reduce((s, r) => s + num(r.totalActivity), 0);
  const sfTotal = normalized.reduce((s, r) => s + num(r.sfActivity), 0);
  const fsvTotal = normalized.reduce((s, r) => s + num(r.combinedFsvFrameActivity), 0);

  if (summary) {
    summary.innerHTML = `
      <div class="op-hourly-mini total"><span>Total Jobs</span><strong>${total.toLocaleString()}</strong></div>
      <div class="op-hourly-mini sf"><span>SF Output</span><strong>${sfTotal.toLocaleString()}</strong></div>
      <div class="op-hourly-mini fsv"><span>FSV Output</span><strong>${fsvTotal.toLocaleString()}</strong></div>
      <div class="op-hourly-mini peak"><span>Peak Hour</span><strong>${escHtml(peak.hour || '--')}</strong></div>
    `;
  }

  if (!normalized.length) {
    el.innerHTML = '<div class="op-empty">No hourly performance available yet.</div>';
    return;
  }

  el.innerHTML = normalized.map(row => {
    const sf = num(row.sfActivity);
    const fsv = num(row.combinedFsvFrameActivity);
    const total = num(row.totalActivity);
    const sfPct = total > 0 ? Math.round((sf / total) * 100) : 0;
    const fsvPct = total > 0 ? Math.round((fsv / total) * 100) : 0;
    const totalPct = Math.max(total > 0 ? 3 : 0, Math.round((total / max) * 100));
    const isPeak = row.hour === peak.hour && total > 0;
    const targetStatus = getHourlyTargetStatusForSplit(row);

    return `
      <div class="personal-hour-row ${isPeak ? 'peak' : ''} ${targetStatus}" title="SF ${sf.toLocaleString()} · FSV ${fsv.toLocaleString()} · Total ${total.toLocaleString()}">
        <div class="personal-hour-time">${escHtml(row.hour)}</div>

        <div class="personal-hour-body">
          <div class="personal-total-track">
            <i style="width:${totalPct}%"></i>
          </div>
          <div class="personal-split-track">
            <i class="sf" style="width:${sfPct}%"></i>
            <i class="fsv" style="width:${fsvPct}%"></i>
          </div>
        </div>

        <div class="personal-hour-metric sf"><span>SF</span><strong>${sf.toLocaleString()}</strong></div>
        <div class="personal-hour-metric fsv"><span>FSV</span><strong>${fsv.toLocaleString()}</strong></div>
        <div class="personal-hour-total"><span>Total</span><strong>${total.toLocaleString()}</strong></div>
      </div>
    `;
  }).join('');
}

function getHourlyTargetStatusForSplit(row) {
  const sf = num(row.sfActivity);
  const fsv = num(row.combinedFsvFrameActivity);
  const sfTarget = num(PICKING_JPH_CONFIG?.sfTarget || 98);
  const fsvTarget = num(PICKING_JPH_CONFIG?.fsvTarget || 72);
  const amberPct = num(PICKING_JPH_CONFIG?.amberPct || 90);

  const statuses = [];

  if (sf > 0 && sfTarget > 0) {
    const sfPct = Math.round((sf / sfTarget) * 100);
    statuses.push(sfPct >= 100 ? 'good' : sfPct >= amberPct ? 'warn' : 'bad');
  }

  if (fsv > 0 && fsvTarget > 0) {
    const fsvPct = Math.round((fsv / fsvTarget) * 100);
    statuses.push(fsvPct >= 100 ? 'good' : fsvPct >= amberPct ? 'warn' : 'bad');
  }

  if (!statuses.length) return 'idle';
  if (statuses.includes('bad')) return 'bad';
  if (statuses.includes('warn')) return 'warn';
  return 'good';
}


function renderReportOperatorHourlyPerformance() {
  const rowsEl = document.getElementById('rptOperatorHourlyRows');
  const summaryEl = document.getElementById('rptOperatorHourlySummary');
  const metaEl = document.getElementById('rptOperatorHourlyMeta');

  if (!rowsEl && !summaryEl && !metaEl) return;

  const rows = (Array.isArray(STATE.activityByOperator) ? STATE.activityByOperator : [])
    .filter(r => num(r.total) > 0)
    .map(r => ({
      ...r,
      flowStation: displayPickingStationName(r.flowStation || r.accessPoint || ''),
      accessPoint: displayPickingStationName(r.accessPoint || r.flowStation || '')
    }));

  const sortedRows = [...rows].sort((a, b) => num(b.total) - num(a.total));
  const totalOutput = rows.reduce((s, r) => s + num(r.total), 0);
  const sfOutput = rows
    .filter(r => displayPickingStationName(r.flowStation || r.accessPoint).includes('SF'))
    .reduce((s, r) => s + num(r.total), 0);
  const fsvOutput = rows
    .filter(r => displayPickingStationName(r.flowStation || r.accessPoint).includes('FSV'))
    .reduce((s, r) => s + num(r.total), 0);
  const activeOperators = new Set(rows.map(r => String(r.operator || '').trim()).filter(Boolean)).size;
  const peak = getPeakOperatorHour(null, rows);
  const top = sortedRows[0] || null;

  if (metaEl) {
    metaEl.textContent = rows.length
      ? `${activeOperators} operator(s) · ${rows.length} station row(s)`
      : 'Waiting for operator activity';
  }

  if (summaryEl) {
    summaryEl.innerHTML = `
      <div class="rpt-op-summary-card"><span>Total Operator Output</span><strong>${totalOutput.toLocaleString()}</strong><small>SF + FSV operator activity today</small></div>
      <div class="rpt-op-summary-card sf"><span>SF Scan & Verify</span><strong>${sfOutput.toLocaleString()}</strong><small>${PICKING_JPH_CONFIG.sfTarget}/hr target · per operator hour</small></div>
      <div class="rpt-op-summary-card fsv"><span>FSV Scan & Verify</span><strong>${fsvOutput.toLocaleString()}</strong><small>${PICKING_JPH_CONFIG.fsvTarget}/hr target · per operator hour</small></div>
      <div class="rpt-op-summary-card"><span>Active Operators</span><strong>${activeOperators.toLocaleString()}</strong><small>Unique operators with activity</small></div>
      <div class="rpt-op-summary-card"><span>Peak Hour</span><strong>${escHtml(peak.hour || '--')}</strong><small>${num(peak.total).toLocaleString()} jobs across operators</small></div>
    `;
  }

  if (!rowsEl) return;

  if (!sortedRows.length) {
    rowsEl.innerHTML = '<div class="rpt-op-empty">No operator hourly performance returned by the API yet.</div>';
    return;
  }

  rowsEl.innerHTML = sortedRows.map((r, i) => {
    const station = displayPickingStationName(r.flowStation || r.accessPoint || '--');
    const isFsv = station.toUpperCase().includes('FSV');
    const bestResult = getPickingJphResult(num(r.bestHourValue), station, r.bestHour);
    const statusClass = bestResult.cls || 'neutral';
    const statusText = bestResult.pctText || bestResult.status || 'No Target';
    const maxHour = Math.max(...Object.values(r.hours || {}).map(v => num(v)), 1);
    const bars = buildReportOperatorHourBars(r.hours || {}, maxHour, isFsv);

    return `
      <div class="rpt-op-row" style="animation-delay:${Math.min(i, 25) * .025}s">
        <div class="rpt-op-person">
          <span class="rpt-op-avatar">${escHtml(getOperatorInitials(r.operator))}</span>
          <div>
            <strong>${escHtml(r.operator || '--')}</strong>
            <small>${escHtml(r.area || 'Picking')}</small>
          </div>
        </div>
        <div><span class="rpt-op-station ${isFsv ? 'fsv' : ''}">${escHtml(station)}</span></div>
        <div class="rpt-op-total">${num(r.total).toLocaleString()}</div>
        <div class="rpt-op-best"><b>${escHtml(r.bestHour || '--')}</b>${r.bestHourValue ? ` · ${num(r.bestHourValue).toLocaleString()}` : ''}</div>
        <div class="rpt-op-last">${escHtml(r.lastActiveHour || '--')}</div>
        <div class="rpt-op-hour-strip ${isFsv ? 'fsv' : ''}" title="${escHtml(buildReportOperatorHourTitle(r.hours || {}))}">${bars}</div>
        <div><span class="rpt-op-status ${statusClass}">${escHtml(statusText)}</span></div>
      </div>
    `;
  }).join('');
}

function buildReportOperatorHourBars(hours, maxValue, isFsv) {
  const entries = Object.entries(hours || {});
  const max = Math.max(...entries.map(([, v]) => num(v)), num(maxValue), 1);

  if (!entries.length) {
    return '<i style="height:2px;opacity:.35"></i>';
  }

  return entries.map(([hour, value]) => {
    const val = num(value);
    const result = getPickingJphResult(val, isFsv ? 'FSV Scan & Verify' : 'SF Scan & Verify', hour);
    const h = Math.max(val > 0 ? 4 : 2, Math.round((val / max) * 30));
    const opacity = val > 0 ? 1 : .28;
    return `<i class="${escHtml(result.cls || '')}" style="height:${h}px;opacity:${opacity}"></i>`;
  }).join('');
}

function buildReportOperatorHourTitle(hours) {
  return Object.entries(hours || {})
    .map(([hour, value]) => `${hour}: ${num(value).toLocaleString()}`)
    .join(' · ');
}

/* ──────────────────────────────────────────────────
   Reports Tab
────────────────────────────────────────────────── */

function renderReportsTab() {
  const wip = STATE.wip || [];
  const daily = STATE.dailyStats || {};
  const totalWip = STATE.totalWip || sum(wip, 'total');
  const sorted = [...wip].sort((a, b) => b.total - a.total);

  const sf = num(daily.sfProductivityToday);
  const fsv = num(daily.fsvProductivityToday);
  const frame = num(daily.frameOnlyProductivityToday ?? daily.framePickingActivityToday);
  const brk = num(daily.breakageCountToday);
  const rep = num(daily.replenishmentCountToday);

  // Total Jobs of the Day must include all scan/verify production buckets.
  // Rule: SF + FSV + Frame Only Scan & Verify.
  const totalProductivity = sf + fsv + frame;

  const now = new Date().toLocaleString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });

  setText('rptGeneratedAt', `Generated ${now}`);

  const kpiGrid = document.getElementById('rptKpiGrid');

  if (kpiGrid) {
    const kpiData = [
      { label: 'Total WIP', val: totalWip, sub: 'Current live snapshot', cls: 'blue' },
      { label: 'Surface (SF) Total Jobs', val: sf, sub: 'Total Scan & Verify', cls: 'green' },
      { label: 'Finish (FSV) Total Jobs', val: fsv, sub: 'Total Scan & Verify', cls: 'orange' },
      { label: 'Frame Only Total Jobs', val: frame, sub: 'Frame Only Scan & Verify', cls: 'yellow' },
      { label: 'Total Jobs of the Day', val: totalProductivity, sub: 'SF + FSV + Frame Only', cls: 'purple' },
      { label: 'Breakage Count', val: brk, sub: 'Record count from Breakage', cls: 'orange' },
      { label: 'Replenishment Count', val: rep, sub: 'Record count from Replenishment', cls: 'blue' }
    ];

    kpiGrid.innerHTML = kpiData.map(k => `
      <div class="rpt-kpi-card ${k.cls}">
        <div class="rpt-kpi-val">${Number(k.val || 0).toLocaleString()}</div>
        <div class="rpt-kpi-label">${escHtml(k.label)}</div>
        <div class="rpt-kpi-sub">${escHtml(k.sub)}</div>
      </div>
    `).join('');
  }

  const trackedEl = document.getElementById('rptTrackedDetail');

  if (trackedEl) {
    const metrics = [
      {
        name: 'SF Productive Today',
        short: 'SF',
        cls: 'sf',
        current: getWipFor('SF Scan & Verify', wip),
        count: sf,
        type: 'PRODUCTIVITY',
        definition: 'Positive gain into SF Scan & Verify'
      },
      {
        name: 'FSV Productive Today',
        short: 'FSV',
        cls: 'fsv',
        current: getWipFor('FSV Scan & Verify', wip),
        count: fsv,
        type: 'PRODUCTIVITY',
        definition: 'Positive gain into FSV Scan & Verify'
      },
      {
        name: 'Breakage Count Today',
        short: 'BR',
        cls: 'brk',
        current: getWipFor('Breakage to Picking', wip),
        count: brk,
        type: 'RECORD',
        definition: 'Drops from Breakage to Picking'
      },
      {
        name: 'Replenishment Count Today',
        short: 'RP',
        cls: 'brk',
        current: getWipFor('Replenishment', wip),
        count: rep,
        type: 'RECORD',
        definition: 'Drops from Replenishment'
      }
    ];

    trackedEl.innerHTML = metrics.map(m => `
      <div class="rpt-tracked-row">
        <div class="rpt-tracked-badge ${m.cls}">${m.short}</div>
        <div class="rpt-tracked-name">${escHtml(m.name)}</div>
        <div class="rpt-tracked-stat"><span>Current WIP</span><strong class="${m.cls}">${m.current.toLocaleString()}</strong></div>
        <div class="rpt-tracked-stat"><span>Count Today</span><strong>${m.count.toLocaleString()}</strong></div>
        <div class="rpt-tracked-stat"><span>Type</span><strong style="color:var(--cyan)">${escHtml(m.type)}</strong></div>
        <div class="rpt-tracked-stat"><span>Definition</span><strong style="font-size:12px">${escHtml(m.definition)}</strong></div>
        <div class="rpt-tracked-stat"><span>Status</span><strong style="color:var(--green)">Tracked</strong></div>
      </div>
    `).join('');
  }

  setText('rptSnapshotMeta', `${sorted.length} picking types`);

  const rptTable = document.getElementById('rptWipTable');
  const max = Math.max(...sorted.map(r => r.total), 1);

  if (rptTable) {
    rptTable.innerHTML = sorted.length
      ? sorted.map((r, i) => {
          const tracked = TRACKED.includes(r.picking);
          const pct = totalWip ? (r.total / totalWip * 100).toFixed(1) : '0.0';
          const barPct = Math.max(2, Math.round(r.total / max * 100));

          const icon = tracked
            ? (r.picking === 'SF Scan & Verify' ? 'SF' : r.picking === 'FSV Scan & Verify' ? 'FSV' : r.picking === 'Breakage to Picking' ? 'BR' : 'RP')
            : getQueueIcon(r.picking);

          return `
            <div class="wip-row ${tracked ? 'blue-row' : 'orange-row'}" style="grid-template-columns:1.8fr .9fr 1fr .8fr 1.4fr;animation-delay:${i * .03}s">
              <div class="wip-name"><span class="wip-icon">${icon}</span>${escHtml(r.picking)}</div>
              <div class="dept">${escHtml(r.department || 'Inventory')}</div>
              <div class="wip-count"><strong>${r.total.toLocaleString()}</strong><div class="row-bar"><i style="width:${barPct}%"></i></div></div>
              <div style="color:var(--muted);font-family:'JetBrains Mono',monospace;font-size:13px">${pct}%</div>
              <div class="wip-date">${toEastern(r.lastUpdated)}</div>
            </div>
          `;
        }).join('')
      : '<div class="error-log-empty"><span>◌</span>No WIP snapshot available</div>';
  }

  const activitySummary = document.getElementById('rptActivitySummary');

  if (activitySummary) {
    activitySummary.innerHTML = `
      <div class="rpt-summary-grid">
        <div class="rpt-summary-head"><span>Metric</span><span>Type</span><span>Today</span><span>Current WIP</span><span>Definition</span></div>

        <div class="rpt-summary-row tracked">
          <span>SF Productive Today</span>
          <span style="color:var(--green)">PRODUCTIVITY</span>
          <span style="color:var(--green)">${sf.toLocaleString()}</span>
          <span>${getWipFor('SF Scan & Verify', wip).toLocaleString()}</span>
          <span>Gain into SF Scan & Verify</span>
        </div>

        <div class="rpt-summary-row tracked">
          <span>FSV Productive Today</span>
          <span style="color:var(--green)">PRODUCTIVITY</span>
          <span style="color:var(--green)">${fsv.toLocaleString()}</span>
          <span>${getWipFor('FSV Scan & Verify', wip).toLocaleString()}</span>
          <span>Gain into FSV Scan & Verify</span>
        </div>

        <div class="rpt-summary-row tracked">
          <span>Breakage Count Today</span>
          <span style="color:var(--orange)">RECORD</span>
          <span style="color:var(--orange)">${brk.toLocaleString()}</span>
          <span>${getWipFor('Breakage to Picking', wip).toLocaleString()}</span>
          <span>Drop from Breakage to Picking</span>
        </div>

        <div class="rpt-summary-row tracked">
          <span>Replenishment Count Today</span>
          <span style="color:var(--orange)">RECORD</span>
          <span style="color:var(--orange)">${rep.toLocaleString()}</span>
          <span>${getWipFor('Replenishment', wip).toLocaleString()}</span>
          <span>Drop from Replenishment</span>
        </div>
      </div>
    `;
  }

  renderReportOperatorHourlyPerformance();

  renderNarrativeReport(totalWip, sf, fsv, frame, brk, rep, totalProductivity);
}

function renderNarrativeReport(totalWip, sf, fsv, frame, brk, rep, totalProductivity) {
  const narrativeEl = document.getElementById('rptNarrative');
  if (!narrativeEl) return;

  setText(
    'rptNarrativeMeta',
    `${new Date().toLocaleDateString('en-US', {
      weekday: 'long',
      month: 'long',
      day: 'numeric',
      year: 'numeric'
    })}`
  );

  const today = new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric'
  });

  const timeNow = new Date().toLocaleTimeString('en-US', {
    hour: '2-digit',
    minute: '2-digit'
  });

  const otherQueues = STATE.wip
    .filter(r => !TRACKED.includes(r.picking) && r.total > 0)
    .sort((a, b) => b.total - a.total);

  const otherHighlight = otherQueues
    .slice(0, 3)
    .map(r => `<span class="narr-tag orange">${escHtml(r.picking)}: ${r.total.toLocaleString()}</span>`)
    .join(' ');

  narrativeEl.innerHTML = `
    <div class="narr-block">
      <div class="narr-dateline">📋 &nbsp; ${today} &nbsp;·&nbsp; As of ${timeNow}</div>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Overall Status</div>
      <p class="narr-text">
        As of <strong>${timeNow}</strong>, total picking WIP stands at
        <span class="narr-num cyan">${totalWip.toLocaleString()}</span> items across
        <span class="narr-num">${STATE.wip.filter(r => r.total > 0).length}</span> active queues.
        SF + FSV + Frame Only productive movement today is
        <span class="narr-num green">${totalProductivity.toLocaleString()}</span>.
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Productivity Counts</div>
      <p class="narr-text">
        SF Productive Today:
        <span class="narr-num cyan">${sf.toLocaleString()}</span>
        <span class="narr-tag cyan">Surface Unbox</span>
        <br>
        FSV Productive Today:
        <span class="narr-num orange">${fsv.toLocaleString()}</span>
        <span class="narr-tag orange">Finish Unbox</span>
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Record Counts</div>
      <p class="narr-text">
        Breakage Count Today:
        <span class="narr-num orange">${brk.toLocaleString()}</span>
        <span class="narr-tag orange">Drop from Breakage to Picking</span>
        <br>
        Replenishment Count Today:
        <span class="narr-num cyan">${rep.toLocaleString()}</span>
        <span class="narr-tag cyan">Drop from Replenishment</span>
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Movement Check</div>
      <p class="narr-text">
        ${escHtml(STATE.movementMessage || 'Normal WIP movement.')}
      </p>
    </div>

    ${otherQueues.length ? `
      <div class="narr-block">
        <div class="narr-section-title">Other Active Queues</div>
        <p class="narr-text">
          ${otherQueues.length} additional queue${otherQueues.length > 1 ? 's' : ''} currently active:
          ${otherHighlight}
          ${otherQueues.length > 3 ? `<span class="narr-tag gray">+${otherQueues.length - 3} more</span>` : ''}
        </p>
      </div>
    ` : ''}

    <div class="narr-block narr-footer">
      <span>⬡ Auto-generated by Zenni Lab — Picking Dashboard</span>
      <span>Report time: ${new Date().toLocaleString()}</span>
    </div>
  `;
}

/* ──────────────────────────────────────────────────
   Helpers
────────────────────────────────────────────────── */

function getWipFor(name, wip) {
  const row = wip.find(r => r.picking === name);
  return row ? row.total : 0;
}

function getOverallTrend() {
  if (STATE.movementStatus === 'VALID_MOVEMENT') {
    return {
      arrow: '↔',
      value: 0,
      text: 'Valid station movement'
    };
  }

  if (STATE.movementStatus === 'WARNING_REVIEW') {
    return {
      arrow: '⚠',
      value: 0,
      text: 'Movement review needed'
    };
  }

  return {
    arrow: '●',
    value: 0,
    text: 'Live WIP snapshot'
  };
}

function getQueueIcon(name) {
  if (/surface|finish/i.test(name)) return '◷';
  if (/re-cal/i.test(name)) return '⚒';
  if (/frame/i.test(name)) return '▢';
  if (/hold/i.test(name)) return '▣';
  if (/scan verify/i.test(name)) return 'SV';
  if (/out queue/i.test(name)) return 'OQ';
  if (/replenishment/i.test(name)) return 'RP';
  return '•';
}

function toEastern(ts) {
  if (!ts) return '--';

  try {
    const d = new Date(ts);

    if (isNaN(d)) {
      return ts;
    }

    return d.toLocaleString('en-US', {
      timeZone: 'America/New_York',
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true
    }) + ' ET';
  } catch {
    return ts;
  }
}

function latestTimestamp(wip, completion) {
  return [
    ...wip.map(r => r.lastUpdated),
    ...completion.map(e => e.timestamp)
  ].filter(Boolean).sort().pop() || '';
}

function wrapLabel(str) {
  return String(str)
    .replace(/ Queue$/, ' Queue')
    .replace(/ /g, '<br>');
}

function num(v) {
  if (v === null || v === undefined || v === '') return 0;

  const n = parseInt(String(v).replace(/,/g, ''), 10);

  return Number.isFinite(n) ? n : 0;
}

function sum(rows, key) {
  return rows.reduce((s, r) => s + num(r[key]), 0);
}

function setText(id, value) {
  const el = document.getElementById(id);

  if (el) {
    el.textContent = value;
  }
}

function animateNumber(id, target) {
  const el = document.getElementById(id);

  if (!el) return;

  target = Number(target || 0);

  const current = parseInt(String(el.textContent).replace(/,/g, ''), 10) || 0;

  if (current === target) {
    el.textContent = target.toLocaleString();
    return;
  }

  const card = el.closest('.kpi-card');

  if (card) {
    card.classList.remove('flash-update');
    void card.offsetWidth;
    card.classList.add('flash-update');
    setTimeout(() => card.classList.remove('flash-update'), 700);
  }

  const start = performance.now();
  const duration = 850;

  const tick = now => {
    const p = Math.min((now - start) / duration, 1);
    const e = 1 - Math.pow(1 - p, 3);

    el.textContent = Math.round(current + (target - current) * e).toLocaleString();

    if (p < 1) {
      requestAnimationFrame(tick);
    } else {
      el.classList.remove('num-pop');
      void el.offsetWidth;
      el.classList.add('num-pop');
      setTimeout(() => el.classList.remove('num-pop'), 500);
    }
  };

  requestAnimationFrame(tick);
}

function setLiveState(state, text) {
  const pill = document.getElementById('livePill');
  const txt = document.getElementById('liveText');

  if (!pill || !txt) return;

  pill.classList.remove('loading', 'error');

  if (state === 'loading') {
    pill.classList.add('loading');
  }

  if (state === 'error') {
    pill.classList.add('error');
  }

  txt.textContent = text;
}

function toast(msg) {
  const el = document.getElementById('toast');

  if (!el) return;

  el.textContent = msg;
  el.classList.add('show');

  clearTimeout(window.__toastTimer);

  window.__toastTimer = setTimeout(() => {
    el.classList.remove('show');
  }, 3000);
}

function isTrackingOnlyQueue(name) {
  const key = String(name || '')
    .toUpperCase()
    .replace(/&/g, 'AND')
    .replace(/[^A-Z0-9]/g, '');

  return (
    key === 'SFSCANANDVERIFY' ||
    key === 'FSVSCANANDVERIFY' ||
    key === 'FRAMEONLYSCANANDVERIFY'
  );
}

function getTrackingIcon(name) {
  const key = String(name || '').toUpperCase();

  if (key.includes('SF SCAN')) return 'SF';
  if (key.includes('FSV SCAN')) return 'FSV';
  if (key.includes('FRAME ONLY')) return 'FO';

  return 'TR';
}

/* ──────────────────────────────────────────────────
   Motion Layer
────────────────────────────────────────────────── */

function initParticles() {
  const count = 18;
  const colors = [
    'rgba(25,200,255,',
    'rgba(180,76,255,',
    'rgba(255,159,28,',
    'rgba(33,229,124,'
  ];

  for (let i = 0; i < count; i++) {
    const p = document.createElement('div');
    p.className = 'data-particle';

    const size = Math.random() * 3 + 1.5;
    const left = Math.random() * 100;
    const delay = Math.random() * 18;
    const dur = 14 + Math.random() * 18;
    const drift = (Math.random() - 0.5) * 120;
    const color = colors[Math.floor(Math.random() * colors.length)];
    const opacity = 0.15 + Math.random() * 0.3;

    p.style.cssText = `
      width:${size}px;
      height:${size}px;
      left:${left}%;
      bottom:-10px;
      background:${color}${opacity});
      box-shadow:0 0 ${size * 3}px ${color}0.5);
      --drift:${drift}px;
      animation-duration:${dur}s;
      animation-delay:${delay}s;
    `;

    document.body.appendChild(p);
  }
}

function updateTicker(wip) {
  const daily = STATE.dailyStats || {};

  const sf = num(daily.sfProductivityToday);
  const fsv = num(daily.fsvProductivityToday);
  const brk = num(daily.breakageCountToday);
  const rep = num(daily.replenishmentCountToday);

  const tot = STATE.totalWip || sum(wip, 'total');
  const lines = wip.filter(r => r.total > 0).length;

  ['tk-sf', 'tk-sf2'].forEach(id => setText(id, sf.toLocaleString()));
  ['tk-fsv', 'tk-fsv2'].forEach(id => setText(id, fsv.toLocaleString()));
  ['tk-brk', 'tk-brk2'].forEach(id => setText(id, brk.toLocaleString()));
  ['tk-rep', 'tk-rep2'].forEach(id => setText(id, rep.toLocaleString()));
  ['tk-total', 'tk-total2'].forEach(id => setText(id, tot.toLocaleString()));
  ['tk-lines', 'tk-lines2'].forEach(id => setText(id, lines));
}