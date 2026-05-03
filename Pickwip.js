/* Picking WIP — Command Center Logic
   Updated for new backend:
   payload.data
   payload.movement
   payload.totalWip
   payload.movementStatus
   payload.movementMessage
   payload.health
*/

const API = 'https://script.google.com/macros/s/AKfycbzaEX7HJODh0PhUm7GwDi4Htx9UYoUfVbIV9i20EM-Uef7JjfCBlGNVRK5enOhKYQpPCQ/exec';

const AUTO_REFRESH_MS = 60000;
const USE_DEMO_ON_ERROR = true;

const TRACKED = ['SF Scan & Verify', 'FSV Scan & Verify', 'Breakage to Picking'];

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
  'Breakage to Picking': 'brk'
};

let STATE = {
  wip: [],
  completion: [],
  history: [],
  movement: [],
  movementStatus: 'LOADING',
  movementMessage: 'Waiting for movement comparison...',
  health: {},
  totalWip: 0,
  demo: false
};

const DEMO = {
  wip: [
    { picking: 'SF Scan & Verify', department: 'Inventory', total: 182, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'FSV Scan & Verify', department: 'Inventory', total: 42, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Breakage to Picking', department: 'Inventory', total: 1, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Out Queue', department: 'Inventory', total: 61, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Scan Verify', department: 'Inventory', total: 40, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Re-Cal', department: 'Inventory', total: 17, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Frame Only Scan & Verify', department: 'Inventory', total: 2, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Replenishment', department: 'Inventory', total: 2, lastUpdated: '2026-04-30 13:31:39' },
    { picking: 'Release from Hold', department: 'Inventory', total: 1, lastUpdated: '2026-04-30 13:31:39' }
  ],
  completion: [
    { picking: 'SF Scan & Verify', completed: 98, direction: 'completed', wip_before: 181, wip_after: 182, timestamp: '2026-04-30 13:31:39' },
    { picking: 'FSV Scan & Verify', completed: 189, direction: 'completed', wip_before: 45, wip_after: 42, timestamp: '2026-04-30 13:31:39' },
    { picking: 'Breakage to Picking', completed: 73, direction: 'completed', wip_before: 8, wip_after: 1, timestamp: '2026-04-30 13:31:39' }
  ],
  movement: [
    {
      subDepartment: 'Out Queue',
      previousTotal: 800,
      currentTotal: 300,
      difference: -500,
      movementType: 'DROP'
    },
    {
      subDepartment: 'Scan Verify',
      previousTotal: 200,
      currentTotal: 698,
      difference: 498,
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

  const entry = {
    id,
    level,
    title,
    detail,
    source,
    ts
  };

  ERROR_LOG.unshift(entry);

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
  return String(str)
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
        completion: normalizeCompletion(DEMO.completion),
        history: [],
        movement: normalizeMovement(DEMO.movement),
        movementStatus: 'DEMO',
        movementMessage: 'API unreachable — displaying demo data.',
        health: {},
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

/* ──────────────────────────────────────────────────
   Main Render
────────────────────────────────────────────────── */

function renderAll() {
  const wip = STATE.wip;
  const completion = STATE.completion;

  const totalWip = STATE.totalWip || sum(wip, 'total');

  const completedTotal = completion
    .filter(e => e.direction === 'completed')
    .reduce((s, e) => s + e.completed, 0);

  const latest =
    STATE.health?.lastImportTimestamp ||
    latestTimestamp(wip, completion);

  const sorted = [...wip].sort((a, b) => b.total - a.total);
  const highest = sorted[0] || { picking: '--', total: 0 };
  const activeLines = wip.filter(r => r.total > 0).length;

  setText('lastUpdated', latest ? toEastern(latest) : '--');
  setText('sideSystemText', STATE.demo ? 'Demo Mode Active' : 'All Systems Operational');

  animateNumber('kpiTotalWip', totalWip);
  animateNumber('kpiCompleted', completedTotal);
  animateNumber('kpiHighestQueue', highest.total);
  setText('kpiHighestName', highest.picking);
  animateNumber('kpiActiveLines', activeLines);

  const trendInfo = getOverallTrend(completion);

  setText('kpiTrend', `${trendInfo.arrow} ${trendInfo.value}`);
  setText('kpiTrendText', STATE.movementStatus || trendInfo.text);

  renderMovementCheck();
  renderTrackingRows(wip, completion);
  renderWipRows(wip);
  renderDistribution(wip, totalWip);
  renderQueueBars(wip);
  renderFocus(wip, totalWip);
  renderQueuesTab(wip, totalWip);
  renderActivityTab();
  renderReportsTab();
  updateTicker(wip, completion);
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

function renderTrackingRows(wip, completion) {
  const maxDone = Math.max(...TRACKED.map(t => getCompletedFor(t, completion)), 1);

  const html = TRACKED.map((name, i) => {
    const cls = CLASS_MAP[name] || '';
    const short = name === 'SF Scan & Verify'
      ? 'SF'
      : name === 'FSV Scan & Verify'
        ? 'FSV'
        : 'BR';

    const done = getCompletedFor(name, completion);
    const current = getWipFor(name, wip);
    const evt = getLatestEventFor(name, completion);
    const change = getChange(evt, name);
    const bar = Math.min((done / maxDone) * 100, 100);

    return `
      <div class="tracking-row" style="animation-delay:${i * .05}s">
        <div class="type-cell ${cls}">
          <span class="type-badge">${short}</span>
          ${name}
        </div>

        <div>
          <span class="count ${cls}">${done.toLocaleString()}</span>
          <span class="progress-mini">
            <i style="width:${bar}%"></i>
          </span>
        </div>

        <div class="count ${cls}">${current.toLocaleString()}</div>
        <div class="last-change">${change.text}</div>

        <div class="${change.up ? 'trend-up' : 'trend-down'}">
          ${change.icon} ${change.value.toLocaleString()}
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

function renderFocus(wip, total) {
  const focus = [...wip].sort((a, b) => b.total - a.total)[0];

  if (!focus) {
    setText('focusTitle', '--');
    setText('focusText', 'Waiting for live data...');
    return;
  }

  const pct = total ? (focus.total / total * 100).toFixed(1) : '0.0';

  setText('focusTitle', focus.picking);
  setText(
    'focusText',
    `Highest WIP at ${focus.total.toLocaleString()} items (${pct}% of total). Monitor throughput and capacity to reduce backlog.`
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
            ? (r.picking === 'SF Scan & Verify' ? 'SF' : r.picking === 'FSV Scan & Verify' ? 'FSV' : 'BR')
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
  const completion = STATE.completion || [];
  const history = STATE.history || [];

  const filter = (document.getElementById('actFilter') || {}).value || '';

  let events = [...completion];

  if (history.length) {
    events = [...events, ...history];
  }

  events = events.filter(e => {
    const p = e.picking || '';
    if (!p || p === '--' || p === 'Unknown') return false;

    const hasBefore = (e.before ?? e.wip_before) != null && (e.before ?? e.wip_before) !== '';
    const hasAfter = (e.after ?? e.wip_after) != null && (e.after ?? e.wip_after) !== '';
    const hasJobs = e.completed != null && e.completed !== '' && e.completed !== 0;

    return hasBefore || hasAfter || hasJobs;
  });

  events = events.sort((a, b) => String(b.timestamp || b.ts || '').localeCompare(String(a.timestamp || a.ts || '')));

  if (filter) {
    events = events.filter(e => e.picking === filter);
  }

  const completed = completion
    .filter(e => e.direction === 'completed')
    .reduce((s, e) => s + e.completed, 0);

  const added = completion
    .filter(e => e.direction !== 'completed')
    .reduce((s, e) => s + e.completed, 0);

  const net = added - completed;

  animateNumber('actTabEntries', completion.length);
  animateNumber('actTabCompleted', completed);
  animateNumber('actTabAdded', added);
  animateNumber('actTabNet', Math.abs(net));
  setText('actTabNetText', net > 0 ? '▲ WIP increasing' : net < 0 ? '▼ WIP decreasing' : 'No net change');

  const logEl = document.getElementById('actLogRows');
  if (!logEl) return;

  setText('actLogMeta', `${events.length} events`);

  if (!events.length) {
    logEl.innerHTML = '<div class="error-log-empty"><span>◌</span>No activity recorded yet</div>';
    return;
  }

  logEl.innerHTML = events.map((e, i) => {
    const picking = e.picking || '--';

    const typeCls =
      picking === 'SF Scan & Verify'
        ? 'sf'
        : picking === 'FSV Scan & Verify'
          ? 'fsv'
          : picking === 'Breakage to Picking'
            ? 'brk'
            : 'other';

    const dir = (e.direction || 'completed').toLowerCase();
    const dirCls = dir === 'completed' ? 'act-dir-completed' : 'act-dir-added';
    const dirLabel = dir === 'completed' ? '▼ Completed' : '▲ Added';

    const beforeVal = e.before ?? e.wip_before;
    const afterVal = e.after ?? e.wip_after;

    const delta = afterVal != null && beforeVal != null
      ? afterVal - beforeVal
      : dir === 'completed'
        ? -Math.abs(e.completed || 0)
        : Math.abs(e.completed || 0);

    const deltaCls = delta < 0 ? 'act-delta-down' : delta > 0 ? 'act-delta-up' : 'act-delta-zero';
    const deltaStr = delta > 0 ? `+${delta}` : String(delta);
    const ts = e.timestamp || e.ts || '--';

    return `
      <div class="act-row" style="animation-delay:${Math.min(i, 25) * 0.04}s">
        <div class="act-ts">${escHtml(ts)}</div>
        <div class="act-type ${typeCls}">${escHtml(picking)}</div>
        <div class="act-num">${beforeVal ?? '--'}</div>
        <div class="act-num">${afterVal ?? '--'}</div>
        <div class="act-num">${e.completed ?? '--'}</div>
        <div class="${dirCls}">${dirLabel}</div>
        <div class="${deltaCls}">${deltaStr}</div>
      </div>
    `;
  }).join('');
}

/* ──────────────────────────────────────────────────
   Reports Tab
────────────────────────────────────────────────── */

function renderReportsTab() {
  const wip = STATE.wip || [];
  const completion = STATE.completion || [];
  const totalWip = STATE.totalWip || sum(wip, 'total');
  const sorted = [...wip].sort((a, b) => b.total - a.total);

  const validEvents = completion.filter(e => {
    const p = e.picking || '';
    if (!p || p === '--') return false;

    const hasBefore = (e.before ?? e.wip_before) != null;
    const hasAfter = (e.after ?? e.wip_after) != null;
    const hasJobs = e.completed != null && e.completed !== 0;

    return hasBefore || hasAfter || hasJobs;
  }).sort((a, b) => String(b.timestamp || '').localeCompare(String(a.timestamp || '')));

  const completedEvents = validEvents.filter(e => e.direction === 'completed');
  const addedEvents = validEvents.filter(e => e.direction !== 'completed');
  const totalCompleted = completedEvents.reduce((s, e) => s + (e.completed || 0), 0);
  const totalAdded = addedEvents.reduce((s, e) => s + (e.completed || 0), 0);
  const netChange = totalAdded - totalCompleted;

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
      { label: 'Total WIP', val: totalWip, sub: 'Across all queues', cls: 'blue' },
      { label: 'Completed Today', val: totalCompleted, sub: 'Jobs finished', cls: 'green' },
      { label: 'Added Today', val: totalAdded, sub: 'Jobs entered queue', cls: 'orange' },
      {
        label: 'Net Change',
        val: Math.abs(netChange),
        sub: netChange <= 0 ? '▼ WIP decreasing' : '▲ WIP increasing',
        cls: netChange <= 0 ? 'green' : 'orange'
      },
      { label: 'Active Lines', val: wip.filter(r => r.total > 0).length, sub: 'Non-zero queues', cls: 'purple' },
      { label: 'Movement', val: STATE.movementStatus === 'VALID_MOVEMENT' ? 1 : 0, sub: STATE.movementStatus || 'Snapshot status', cls: 'blue' }
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
    trackedEl.innerHTML = TRACKED.map(name => {
      const wipRow = wip.find(r => r.picking === name) || { total: 0 };
      const done = completedEvents.filter(e => e.picking === name).reduce((s, e) => s + (e.completed || 0), 0);
      const added = addedEvents.filter(e => e.picking === name).reduce((s, e) => s + (e.completed || 0), 0);
      const net = added - done;

      const cls = name === 'SF Scan & Verify'
        ? 'sf'
        : name === 'FSV Scan & Verify'
          ? 'fsv'
          : 'brk';

      const short = name === 'SF Scan & Verify'
        ? 'SF'
        : name === 'FSV Scan & Verify'
          ? 'FSV'
          : 'BR';

      const pct = totalWip ? (wipRow.total / totalWip * 100).toFixed(1) : '0.0';

      return `
        <div class="rpt-tracked-row">
          <div class="rpt-tracked-badge ${cls}">${short}</div>
          <div class="rpt-tracked-name">${escHtml(name)}</div>
          <div class="rpt-tracked-stat"><span>Current WIP</span><strong class="${cls}">${wipRow.total.toLocaleString()}</strong></div>
          <div class="rpt-tracked-stat"><span>% of Total</span><strong>${pct}%</strong></div>
          <div class="rpt-tracked-stat"><span>Completed</span><strong style="color:var(--green)">${done.toLocaleString()}</strong></div>
          <div class="rpt-tracked-stat"><span>Added</span><strong style="color:var(--orange)">${added.toLocaleString()}</strong></div>
          <div class="rpt-tracked-stat"><span>Net</span><strong style="color:${net <= 0 ? 'var(--green)' : 'var(--red)'}">${net > 0 ? '+' : ''}${net.toLocaleString()}</strong></div>
        </div>
      `;
    }).join('');
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
            ? (r.picking === 'SF Scan & Verify' ? 'SF' : r.picking === 'FSV Scan & Verify' ? 'FSV' : 'BR')
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
    const byType = {};

    validEvents.forEach(e => {
      if (!byType[e.picking]) {
        byType[e.picking] = {
          completed: 0,
          added: 0,
          events: 0
        };
      }

      if (e.direction === 'completed') {
        byType[e.picking].completed += e.completed || 0;
      } else {
        byType[e.picking].added += e.completed || 0;
      }

      byType[e.picking].events++;
    });

    const summaryRows = Object.entries(byType)
      .sort((a, b) => (b[1].completed + b[1].added) - (a[1].completed + a[1].added));

    activitySummary.innerHTML = summaryRows.length
      ? `
        <div class="rpt-summary-grid">
          <div class="rpt-summary-head"><span>Picking Type</span><span>Events</span><span>Completed</span><span>Added</span><span>Net</span></div>
          ${summaryRows.map(([name, d]) => {
            const net = d.added - d.completed;
            const tracked = TRACKED.includes(name);

            return `
              <div class="rpt-summary-row ${tracked ? 'tracked' : ''}">
                <span>${escHtml(name)}</span>
                <span style="color:var(--muted)">${d.events}</span>
                <span style="color:var(--green)">${d.completed.toLocaleString()}</span>
                <span style="color:var(--orange)">${d.added.toLocaleString()}</span>
                <span style="color:${net <= 0 ? 'var(--green)' : 'var(--red)'}">${net > 0 ? '+' : ''}${net.toLocaleString()}</span>
              </div>
            `;
          }).join('')}
        </div>
      `
      : '<div class="error-log-empty"><span>◌</span>No activity data available from current backend</div>';
  }

  renderNarrativeReport(validEvents, totalWip, totalCompleted, totalAdded, netChange);
}

function renderNarrativeReport(validEvents, totalWip, totalCompleted, totalAdded, netChange) {
  const narrativeEl = document.getElementById('rptNarrative');
  if (!narrativeEl) return;

  setText(
    'rptNarrativeMeta',
    `${validEvents.length} events · ${new Date().toLocaleDateString('en-US', {
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

  const overallTrend = netChange < 0 ? 'improving' : netChange === 0 ? 'stable' : 'building';
  const trendColor = netChange < 0 ? 'var(--green)' : netChange === 0 ? 'var(--cyan)' : 'var(--orange)';

  const otherQueues = STATE.wip
    .filter(r => !TRACKED.includes(r.picking) && r.total > 0)
    .sort((a, b) => b.total - a.total);

  const otherHighlight = otherQueues
    .slice(0, 3)
    .map(r => `<span class="narr-tag orange">${escHtml(r.picking)}: ${r.total.toLocaleString()}</span>`)
    .join(' ');

  narrativeEl.innerHTML = `
    <div class="narr-block">
      <div class="narr-dateline">📋 &nbsp; ${today} &nbsp;·&nbsp; As of ${timeNow} &nbsp;·&nbsp; ${validEvents.length} logged events</div>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Overall Status</div>
      <p class="narr-text">
        As of <strong>${timeNow}</strong>, total picking WIP stands at
        <span class="narr-num cyan">${totalWip.toLocaleString()}</span> items across
        <span class="narr-num">${STATE.wip.filter(r => r.total > 0).length}</span> active queues.
        The current movement status is
        <span class="narr-tag ${STATE.movementStatus === 'WARNING_REVIEW' ? 'orange' : STATE.movementStatus === 'DEMO' ? 'gray' : 'green'}">${escHtml(STATE.movementStatus || 'NORMAL')}</span>.
        Overall WIP is <strong style="color:${trendColor}">${overallTrend}</strong>.
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

function getCompletedFor(name, completion) {
  const direct = completion
    .filter(e => e.picking === name && e.direction === 'completed')
    .reduce((s, e) => s + e.completed, 0);

  if (direct) return direct;

  const demoRow = DEMO.completion.find(e => e.picking === name);

  return STATE.demo && demoRow ? demoRow.completed : 0;
}

function getWipFor(name, wip) {
  const row = wip.find(r => r.picking === name);
  return row ? row.total : 0;
}

function getLatestEventFor(name, completion) {
  return completion
    .filter(e => e.picking === name)
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)))[0];
}

function getChange(evt, name) {
  if (!evt && STATE.demo) {
    evt = DEMO.completion.find(e => e.picking === name);
  }

  if (!evt) {
    return {
      text: '<span class="neg">0</span> no change',
      value: 0,
      up: false,
      icon: '▼'
    };
  }

  const delta = evt.after && evt.before
    ? evt.after - evt.before
    : evt.direction === 'completed'
      ? -Math.abs(evt.completed)
      : Math.abs(evt.completed);

  const up = delta > 0;
  const val = Math.abs(delta || evt.completed || 0);
  const label = up ? 'added' : 'done';

  return {
    text: `<span class="${up ? 'pos' : 'neg'}">${up ? '+' : '-'} ${val.toLocaleString()}</span> ${label}`,
    value: val,
    up,
    icon: up ? '▲' : '▼'
  };
}

function getOverallTrend(completion) {
  if (!completion || completion.length === 0) {
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

  const latest = TRACKED.map(t => getChange(getLatestEventFor(t, completion), t));

  const up = latest
    .filter(x => x.up)
    .reduce((s, x) => s + x.value, 0);

  const down = latest
    .filter(x => !x.up)
    .reduce((s, x) => s + x.value, 0);

  if (up > down) {
    return {
      arrow: '▲',
      value: up - down,
      text: 'Backlog increasing'
    };
  }

  return {
    arrow: '▼',
    value: Math.max(0, down - up),
    text: 'Overall improving'
  };
}

function getQueueIcon(name) {
  if (/surface|finish/i.test(name)) return '◷';
  if (/re-cal/i.test(name)) return '⚒';
  if (/frame/i.test(name)) return '▢';
  if (/hold/i.test(name)) return '▣';
  if (/scan verify/i.test(name)) return 'SV';
  if (/out queue/i.test(name)) return 'OQ';
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

function updateTicker(wip, completion) {
  const sf = (wip.find(r => r.picking === 'SF Scan & Verify') || {}).total ?? '--';
  const fsv = (wip.find(r => r.picking === 'FSV Scan & Verify') || {}).total ?? '--';
  const brk = (wip.find(r => r.picking === 'Breakage to Picking') || {}).total ?? '--';

  const tot = STATE.totalWip || sum(wip, 'total');

  const done = completion
    .filter(e => e.direction === 'completed')
    .reduce((s, e) => s + e.completed, 0);

  const lines = wip.filter(r => r.total > 0).length;

  ['tk-sf', 'tk-sf2'].forEach(id => setText(id, sf));
  ['tk-fsv', 'tk-fsv2'].forEach(id => setText(id, fsv));
  ['tk-brk', 'tk-brk2'].forEach(id => setText(id, brk));
  ['tk-total', 'tk-total2'].forEach(id => setText(id, tot.toLocaleString()));
  ['tk-done', 'tk-done2'].forEach(id => setText(id, done.toLocaleString()));
  ['tk-lines', 'tk-lines2'].forEach(id => setText(id, lines));
}