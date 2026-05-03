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

const API = 'https://script.google.com/macros/s/AKfycbzaEX7HJODh0PhUm7GwDi4Htx9UYoUfVbIV9i20EM-Uef7JjfCBlGNVRK5enOhKYQpPCQ/exec';

const AUTO_REFRESH_MS = 60000;
const USE_DEMO_ON_ERROR = true;

const TRACKED = [
  'SF Scan & Verify',
  'FSV Scan & Verify',
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
  totalWip: 0,
  demo: false
};

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

document.addEventListener('DOMContentLoaded', () => {
  setupTabs();

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
        titleEl.textContent = 'Picking WIP — ' + labelEl.textContent;
      }
    });
  });
}

/* ──────────────────────────────────────────────────
   API
────────────────────────────────────────────────── */

async function apiFetch(debug = true) {
  const url = `${API}?debug=${debug ? 'true' : 'false'}&t=${Date.now()}`;

  const res = await fetch(url, {
    method: 'GET',
    redirect: 'follow'
  });

  if (!res.ok) {
    const msg = `HTTP ${res.status}`;
    logError('error', 'API HTTP Error', msg, 'Picking WIP API');
    throw new Error(msg);
  }

  const text = await res.text();

  try {
    return JSON.parse(text);
  } catch {
    logError('error', 'API JSON Parse Failed', text.slice(0, 160), 'Picking WIP API');
    throw new Error('Bad JSON from API');
  }
}

async function fetchAll(showToast = false) {
  setLiveState('loading', 'LOADING');

  if (showToast) {
    toast('Refreshing Picking WIP data...');
  }

  try {
    const payload = await apiFetch(true);

    if (payload.status !== 'success') {
      throw new Error(payload.message || 'API returned an error');
    }

    STATE = {
      wip: normalizeWip(payload.data || []),
      completion: normalizeCompletion(payload.completion || []),
      history: payload.history || [],
      movement: normalizeMovement(payload.movement || []),
      movementStatus: payload.movementStatus || 'NORMAL',
      movementMessage: payload.movementMessage || 'Normal WIP movement.',
      health: payload.health || {},
      dailyStats: normalizeDailyStats(payload.dailyStats || {}),
      totalWip: Number(payload.totalWip || 0),
      demo: false
    };

    setLiveState('ok', 'LIVE');

    if (showToast) {
      toast('Live data loaded.');
    }

  } catch (err) {
    console.warn(err);
    logError('error', 'Fetch Failed', err.message, 'fetchAll → Picking WIP API');

    if (USE_DEMO_ON_ERROR) {
      const demoWip = normalizeWip(DEMO.wip);

      STATE = {
        wip: demoWip,
        completion: buildCompletionFromDailyStats(DEMO.dailyStats),
        history: [],
        movement: normalizeMovement(DEMO.movement),
        movementStatus: 'DEMO',
        movementMessage: 'API unreachable — displaying demo data.',
        health: {},
        dailyStats: normalizeDailyStats(DEMO.dailyStats),
        totalWip: sum(demoWip, 'total'),
        demo: true
      };

      setLiveState('error', 'DEMO');
      logError('warn', 'Demo Mode Active', 'API unreachable — displaying demo data', 'fetchAll');
      toast('API not reachable. Showing demo layout so you can test the UI.');
    } else {
      setLiveState('error', 'ERROR');
      toast('API error: ' + err.message);
    }
  }

  renderAll();
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

function normalizeDailyStats(stats) {
  return {
    sfProductivityToday: num(stats.sfProductivityToday),
    fsvProductivityToday: num(stats.fsvProductivityToday),
    breakageCountToday: num(stats.breakageCountToday),
    replenishmentCountToday: num(stats.replenishmentCountToday),
    totalProductivityToday: num(stats.totalProductivityToday)
  };
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
  animateNumber('kpiBreakageCount', breakageCount);
  animateNumber('kpiReplenishmentCount', replenishmentCount);

  const trendInfo = getOverallTrend();

  renderMovementCheck();
  renderTrackingRows(wip, daily);
  renderWipRows(wip);
  renderDistribution(wip, totalWip);
  renderQueueBars(wip);
  renderFocus(wip, totalWip, highest, totalProductive);
  renderQueuesTab(wip, totalWip);
  renderActivityTab();
  renderReportsTab();
  updateTicker(wip);
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

  if (status === 'VALID_MOVEMENT' || status === 'NORMAL') {
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
      status: 'Moved to SF Scan Verify',
      cls: 'sf',
      short: 'SF'
    },
    {
      name: 'FSV Scan & Verify',
      label: 'FSV Productive Today',
      count: num(daily.fsvProductivityToday),
      current: getWipFor('FSV Scan & Verify', wip),
      type: 'PRODUCTIVITY',
      status: 'Moved to FSV Scan Verify',
      cls: 'fsv',
      short: 'FSV'
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

  const top = [...wip].sort((a, b) => b.total - a.total).slice(0, 5);
  const used = sum(top, 'total');
  const others = Math.max(total - used, 0);
  const segments = others ? [...top, { picking: 'Others', total: others }] : top;

  let start = 0;

  const conic = segments.map((r, i) => {
    const deg = total ? (r.total / total) * 360 : 0;
    const end = start + deg;
    const part = `${COLORS[i % COLORS.length]} ${start}deg ${end}deg`;

    start = end;

    return part;
  }).join(', ');

  donut.style.background = `conic-gradient(${conic || '#192638 0deg 360deg'})`;

  setText('donutTotal', total.toLocaleString());

  legend.innerHTML = segments.map((r, i) => {
    const pct = total ? (r.total / total * 100).toFixed(1) : '0.0';

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
  const sorted = [...wip].sort((a, b) => b.total - a.total);
  const max = Math.max(...sorted.map(r => r.total), 1);
  const empty = sorted.filter(r => r.total === 0).length;
  const largest = sorted[0] || { picking: '--', total: 0 };

  animateNumber('qTabTotalQueues', wip.length);
  animateNumber('qTabTotalWip', total);
  animateNumber('qTabLargest', largest.total);
  setText('qTabLargestName', largest.picking);
  animateNumber('qTabEmpty', empty);

  const cardGrid = document.getElementById('queueCardGrid');

  if (cardGrid) {
    if (!sorted.length) {
      cardGrid.innerHTML = '<div class="error-log-empty"><span>◌</span>No queue data available</div>';
    } else {
      cardGrid.innerHTML = sorted.map(r => {
        const tracked = TRACKED.includes(r.picking);
        const pct = Math.max(2, Math.round(r.total / max * 100));
        const cls = tracked ? 'tracked' : (r.total === 0 ? 'zero' : '');

        return `
          <div class="q-card ${cls}">
            <div class="q-card-name">${escHtml(r.picking)}</div>
            <div class="q-card-count">${r.total.toLocaleString()}</div>
            <div class="q-card-bar"><i style="width:${pct}%"></i></div>
            <div class="q-card-dept">${escHtml(r.department || 'Inventory')}</div>
          </div>
        `;
      }).join('');
    }
  }

  const chartEl = document.getElementById('queueChartBars');

  if (chartEl) {
    setText('qChartMeta', `${sorted.length} queues`);

    chartEl.innerHTML = sorted.length
      ? sorted.map(r => {
          const tracked = TRACKED.includes(r.picking);
          const pct = Math.max(1, Math.round(r.total / max * 100));

          return `
            <div class="q-bar-row ${tracked ? 'tracked' : ''}">
              <div class="q-bar-row-label">${escHtml(r.picking)}</div>
              <div class="q-bar-track"><i style="width:${pct}%"></i></div>
              <div class="q-bar-val">${r.total.toLocaleString()}</div>
            </div>
          `;
        }).join('')
      : '<div class="error-log-empty"><span>◌</span>No queue chart data available</div>';
  }

  const tableEl = document.getElementById('queueTableRows');

  if (tableEl) {
    setText('qTableMeta', `${sorted.length} rows`);

    tableEl.innerHTML = sorted.length
      ? sorted.map((r, i) => {
          const tracked = TRACKED.includes(r.picking);
          const pct = Math.max(2, Math.round(r.total / max * 100));
          const rowCls = tracked ? 'blue-row' : 'orange-row';

          const priority = tracked
            ? '<span class="priority-high">● HIGH</span>'
            : r.total >= 20
              ? '<span class="priority-med">● MED</span>'
              : '<span class="priority-low">○ LOW</span>';

          const icon = tracked
            ? (r.picking === 'SF Scan & Verify' ? 'SF' : r.picking === 'FSV Scan & Verify' ? 'FSV' : r.picking === 'Breakage to Picking' ? 'BR' : 'RP')
            : getQueueIcon(r.picking);

          return `
            <div class="wip-row ${rowCls}" style="grid-template-columns:1.8fr .9fr 1.1fr .8fr 1.25fr;animation-delay:${i * .03}s">
              <div class="wip-name"><span class="wip-icon">${icon}</span>${escHtml(r.picking)}</div>
              <div class="dept">${escHtml(r.department || 'Inventory')}</div>
              <div class="wip-count">
                <strong>${r.total.toLocaleString()}</strong>
                <div class="row-bar"><i style="width:${pct}%"></i></div>
              </div>
              <div>${priority}</div>
              <div class="wip-date">${toEastern(r.lastUpdated)}</div>
            </div>
          `;
        }).join('')
      : '<div class="error-log-empty"><span>◌</span>No queue table data available</div>';
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
   Reports Tab
────────────────────────────────────────────────── */

function renderReportsTab() {
  const wip = STATE.wip || [];
  const daily = STATE.dailyStats || {};
  const totalWip = STATE.totalWip || sum(wip, 'total');
  const sorted = [...wip].sort((a, b) => b.total - a.total);

  const sf = num(daily.sfProductivityToday);
  const fsv = num(daily.fsvProductivityToday);
  const brk = num(daily.breakageCountToday);
  const rep = num(daily.replenishmentCountToday);
  const totalProductivity = num(daily.totalProductivityToday);

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
      { label: 'SF Productive', val: sf, sub: 'Moved into SF Scan Verify', cls: 'green' },
      { label: 'FSV Productive', val: fsv, sub: 'Moved into FSV Scan Verify', cls: 'orange' },
      { label: 'Total Productivity', val: totalProductivity, sub: 'SF + FSV productive count', cls: 'purple' },
      { label: 'Breakage Count', val: brk, sub: 'Record count from Breakage', cls: 'orange' },
      { label: 'Replenishment', val: rep, sub: 'Record count from Replenishment', cls: 'blue' }
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

  renderNarrativeReport(totalWip, sf, fsv, brk, rep, totalProductivity);
}

function renderNarrativeReport(totalWip, sf, fsv, brk, rep, totalProductivity) {
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
        SF + FSV productive movement today is
        <span class="narr-num green">${totalProductivity.toLocaleString()}</span>.
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Productivity Counts</div>
      <p class="narr-text">
        SF Productive Today:
        <span class="narr-num cyan">${sf.toLocaleString()}</span>
        <span class="narr-tag cyan">Gain into SF Scan Verify</span>
        <br>
        FSV Productive Today:
        <span class="narr-num orange">${fsv.toLocaleString()}</span>
        <span class="narr-tag orange">Gain into FSV Scan Verify</span>
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
      <span>⬡ Auto-generated by Picking WIP Command Center</span>
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