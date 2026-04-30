/* Picking WIP — Command Center Logic */
const API = 'https://script.google.com/macros/s/AKfycbzaEX7HJODh0PhUm7GwDi4Htx9UYoUfVbIV9i20EM-Uef7JjfCBlGNVRK5enOhKYQpPCQ/exec';
const AUTO_REFRESH_MS = 60000;
const USE_DEMO_ON_ERROR = true;

const TRACKED = ['SF Scan & Verify', 'FSV Scan & Verify', 'Breakage to Picking'];
const COLORS = ['#19c8ff', '#40a4ff', '#ff9f1c', '#ffd21f', '#77889c', '#b44cff', '#21e57c'];

const CLASS_MAP = {
  'SF Scan & Verify': 'sf',
  'FSV Scan & Verify': 'fsv',
  'Breakage to Picking': 'brk'
};

let STATE = {
  wip: [],
  completion: [],
  history: [],
  demo: false
};

const DEMO = {
  wip: [
    { picking:'SF Scan & Verify', department:'Inventory', total:182, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'FSV Scan & Verify', department:'Inventory', total:42, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Breakage to Picking', department:'Inventory', total:1, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Out of Surface Queue', department:'Inventory', total:61, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Out of Finish Queue', department:'Inventory', total:40, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Re-Cal', department:'Inventory', total:17, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Frame Only Scan & Verify', department:'Inventory', total:2, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Replenishment', department:'Inventory', total:2, lastUpdated:'2026-04-30 13:31:39' },
    { picking:'Release from Hold', department:'Inventory', total:1, lastUpdated:'2026-04-30 13:31:39' }
  ],
  completion: [
    { picking:'SF Scan & Verify', completed:98, direction:'completed', wip_before:181, wip_after:182, timestamp:'2026-04-30 13:31:39' },
    { picking:'FSV Scan & Verify', completed:189, direction:'completed', wip_before:45, wip_after:42, timestamp:'2026-04-30 13:31:39' },
    { picking:'Breakage to Picking', completed:73, direction:'completed', wip_before:8, wip_after:1, timestamp:'2026-04-30 13:31:39' }
  ],
  history: []
};

/* ── Error Log ──────────────────────────────────── */
const ERROR_LOG = [];
let errIdCounter = 0;

function logError(level = 'error', title, detail = '', source = '') {
  const id = ++errIdCounter;
  const ts = new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const entry = { id, level, title, detail, source, ts };
  ERROR_LOG.unshift(entry);

  // cap at 100 entries
  if (ERROR_LOG.length > 100) ERROR_LOG.pop();

  renderErrorLog();
}

function renderErrorLog() {
  const body   = document.getElementById('errorLogBody');
  const empty  = document.getElementById('errorLogEmpty');
  const badge  = document.getElementById('errCountBadge');
  if (!body) return;

  const errCount = ERROR_LOG.filter(e => e.level === 'error').length;
  if (badge) {
    badge.textContent = errCount;
    badge.style.display = errCount > 0 ? 'inline-flex' : 'none';
  }

  if (ERROR_LOG.length === 0) {
    if (empty) empty.style.display = 'block';
    // remove all err-rows
    body.querySelectorAll('.err-row').forEach(el => el.remove());
    return;
  }

  if (empty) empty.style.display = 'none';

  // rebuild rows
  body.querySelectorAll('.err-row').forEach(el => el.remove());

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
  if (idx !== -1) ERROR_LOG.splice(idx, 1);
  renderErrorLog();
}

function clearErrorLog() {
  ERROR_LOG.length = 0;
  renderErrorLog();
}

function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

document.addEventListener('DOMContentLoaded', () => {
  // ── Sidebar tab switching ──────────────────────
  document.querySelectorAll('.side-link[data-tab]').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;

      // Update active state on buttons
      document.querySelectorAll('.side-link').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Show/hide tab content
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('tab-hidden'));
      const panel = document.getElementById('tab-' + target);
      if (panel) panel.classList.remove('tab-hidden');

      // Update header title
      const titleEl = document.querySelector('h1');
      const labelEl = btn.querySelector('b');
      if (titleEl && labelEl) {
        titleEl.textContent = 'Picking WIP — ' + labelEl.textContent;
      }
    });
  });

  // Global JS error capture → Error Log
  window.addEventListener('error', ev => {
    logError('error', ev.message || 'Runtime Error', `${ev.filename || ''}:${ev.lineno}:${ev.colno}`, 'window.onerror');
  });
  window.addEventListener('unhandledrejection', ev => {
    const msg = ev.reason instanceof Error ? ev.reason.message : String(ev.reason);
    logError('error', 'Unhandled Promise Rejection', msg, 'window.unhandledrejection');
  });

  fetchAll(true);
  setInterval(() => fetchAll(false), AUTO_REFRESH_MS);
});

async function apiFetch(action) {
  const url = `${API}?action=${action}&t=${Date.now()}`;

  const res = await fetch(url, {
    method: 'GET',
    redirect: 'follow'
  });

  if (!res.ok) {
    const msg = `HTTP ${res.status}`;
    logError('error', 'API HTTP Error', msg, `action=${action}`);
    throw new Error(msg);
  }

  const text = await res.text();

  try {
    return JSON.parse(text);
  } catch {
    logError('error', 'API JSON Parse Failed', text.slice(0, 120), `action=${action}`);
    throw new Error('Bad JSON from API');
  }
}

async function fetchAll(showToast = false) {
  setLiveState('loading', 'LOADING');

  if (showToast) {
    toast('Refreshing Picking WIP data...');
  }

  try {
    const dash = await apiFetch('getDashboard');

    STATE = {
      wip: normalizeWip(dash.wip || []),
      completion: normalizeCompletion(dash.completion || []),
      history: dash.history || [],
      demo: false
    };

    setLiveState('ok', 'LIVE');
    toast('Live data loaded.');
  } catch (err) {
    console.warn(err);
    logError('error', 'Fetch Failed', err.message, 'fetchAll → getDashboard');

    if (USE_DEMO_ON_ERROR) {
      STATE = {
        wip: normalizeWip(DEMO.wip),
        completion: normalizeCompletion(DEMO.completion),
        history: [],
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

function normalizeWip(rows) {
  return rows.map(r => ({
    picking: r.picking || r.picking_type || r.pickingType || r.type || r.name || r.PickingType || 'Unknown',
    department: r.department || r.Department || 'Inventory',
    total: num(r.total ?? r.wip ?? r.count ?? r.WIP ?? r['WIP Count']),
    lastUpdated: r.lastUpdated || r.last_updated || r.updated || r.LastUpdated || r['Last Updated'] || ''
  })).filter(r => r.picking !== 'Unknown');
}

function normalizeCompletion(rows) {
  return rows.map(r => ({
    picking: r.picking || r.picking_type || r.pickingType || r.type || r.name || 'Unknown',
    completed: num(r.completed ?? r.jobs_completed ?? r.jobs ?? r.CompletedToday ?? r['Completed Today']),
    direction: String(r.direction || r.dir || 'completed').toLowerCase(),
    before: num(r.wip_before ?? r.before),
    after: num(r.wip_after ?? r.after),
    timestamp: r.timestamp || r.ts || r.lastUpdated || r.last_updated || ''
  })).filter(r => r.picking !== 'Unknown');
}

function renderAll() {
  const wip = STATE.wip;
  const completion = STATE.completion;

  const totalWip = sum(wip, 'total');
  const completedTotal = completion
    .filter(e => e.direction === 'completed')
    .reduce((s, e) => s + e.completed, 0);

  const latest = latestTimestamp(wip, completion);
  const sorted = [...wip].sort((a, b) => b.total - a.total);
  const highest = sorted[0] || { picking:'--', total:0 };
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
  setText('kpiTrendText', trendInfo.text);

  renderTrackingRows(wip, completion);
  renderWipRows(wip);
  renderDistribution(wip, totalWip);
  renderQueueBars(wip);
  renderFocus(wip, totalWip);
  renderQueuesTab(wip, totalWip);
  renderActivityTab();
  renderReportsTab();

  // Motion layer — update ticker after data renders
  if (typeof updateTicker === 'function') updateTicker(wip, completion);
}

function renderTrackingRows(wip, completion) {
  const maxDone = Math.max(...TRACKED.map(t => getCompletedFor(t, completion)), 1);

  const html = TRACKED.map((name, i) => {
    const cls = CLASS_MAP[name] || '';
    const short = name === 'SF Scan & Verify' ? 'SF' : name === 'FSV Scan & Verify' ? 'FSV' : 'BR';
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
          <span class="count ${cls}">${done}</span>
          <span class="progress-mini">
            <i style="width:${bar}%"></i>
          </span>
        </div>

        <div class="count ${cls}">${current}</div>
        <div class="last-change">${change.text}</div>

        <div class="${change.up ? 'trend-up' : 'trend-down'}">
          ${change.icon} ${change.value}
          <span class="trend-pulse"></span>
        </div>
      </div>
    `;
  }).join('');

  document.getElementById('trackingRows').innerHTML = html;
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

  document.getElementById('wipRows').innerHTML = order.map((r, i) => {
    const tracked = TRACKED.includes(r.picking);
    const rowClass = tracked ? 'blue-row' : 'orange-row';

    const icon =
      r.picking === 'SF Scan & Verify' ? 'SF' :
      r.picking === 'FSV Scan & Verify' ? 'FSV' :
      r.picking === 'Breakage to Picking' ? 'BR' :
      getQueueIcon(r.picking);

    const pct = Math.max(2, Math.round(r.total / max * 100));

    return `
      <div class="wip-row ${rowClass}" style="animation-delay:${i * .035}s">
        <div class="wip-name">
          <span class="wip-icon">${icon}</span>
          ${r.picking}
        </div>

        <div class="dept">${r.department || 'Inventory'}</div>

        <div class="wip-count">
          <strong>${r.total}</strong>
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
  const top = [...wip].sort((a, b) => b.total - a.total).slice(0, 5);
  const used = sum(top, 'total');
  const others = Math.max(total - used, 0);
  const segments = others ? [...top, { picking:'Others', total:others }] : top;

  let start = 0;

  const conic = segments.map((r, i) => {
    const deg = total ? (r.total / total) * 360 : 0;
    const end = start + deg;
    const part = `${COLORS[i % COLORS.length]} ${start}deg ${end}deg`;

    start = end;

    return part;
  }).join(', ');

  document.getElementById('donutChart').style.background =
    `conic-gradient(${conic || '#192638 0deg 360deg'})`;

  setText('donutTotal', total);

  document.getElementById('wipLegend').innerHTML = segments.map((r, i) => {
    const pct = total ? (r.total / total * 100).toFixed(1) : '0.0';

    return `
      <div class="leg-row">
        <span class="leg-dot" style="background:${COLORS[i % COLORS.length]};color:${COLORS[i % COLORS.length]}"></span>
        <span>${r.picking}</span>
        <b>${r.total} (${pct}%)</b>
      </div>
    `;
  }).join('');
}

function renderQueueBars(wip) {
  const queues = wip
    .filter(r => !TRACKED.includes(r.picking))
    .sort((a, b) => b.total - a.total)
    .slice(0, 6);

  const max = Math.max(...queues.map(q => q.total), 1);

  document.getElementById('queueBars').innerHTML = queues.map(q => {
    const height = Math.max(4, Math.round(q.total / max * 100));

    return `
      <div class="bar-item">
        <div class="bar-val">${q.total}</div>
        <div class="bar-col" style="height:${height}%"></div>
        <div class="bar-label">${wrapLabel(q.picking)}</div>
      </div>
    `;
  }).join('');
}

function renderFocus(wip, total) {
  const focus = [...wip].sort((a, b) => b.total - a.total)[0];

  if (!focus) return;

  const pct = total ? (focus.total / total * 100).toFixed(1) : '0.0';

  setText('focusTitle', focus.picking);
  setText(
    'focusText',
    `Highest WIP at ${focus.total} items (${pct}% of total). Monitor throughput and capacity to reduce backlog.`
  );
}

/* ── Queues Tab ─────────────────────────────────── */
function renderQueuesTab(wip, total) {
  if (!wip || !wip.length) return;

  const sorted  = [...wip].sort((a, b) => b.total - a.total);
  const max     = Math.max(...sorted.map(r => r.total), 1);
  const empty   = sorted.filter(r => r.total === 0).length;
  const largest = sorted[0] || { picking:'--', total:0 };

  animateNumber('qTabTotalQueues', wip.length);
  animateNumber('qTabTotalWip', total);
  animateNumber('qTabLargest', largest.total);
  setText('qTabLargestName', largest.picking);
  animateNumber('qTabEmpty', empty);

  // ── Cards ──
  const cardGrid = document.getElementById('queueCardGrid');
  if (cardGrid) {
    cardGrid.innerHTML = sorted.map(r => {
      const tracked = TRACKED.includes(r.picking);
      const pct = Math.max(2, Math.round(r.total / max * 100));
      const cls = tracked ? 'tracked' : (r.total === 0 ? 'zero' : '');
      return `
        <div class="q-card ${cls}">
          <div class="q-card-name">${r.picking}</div>
          <div class="q-card-count">${r.total}</div>
          <div class="q-card-bar"><i style="width:${pct}%"></i></div>
          <div class="q-card-dept">${r.department || 'Inventory'}</div>
        </div>`;
    }).join('');
  }

  // ── Horizontal bar chart ──
  const chartEl = document.getElementById('queueChartBars');
  if (chartEl) {
    setText('qChartMeta', `${sorted.length} queues`);
    chartEl.innerHTML = sorted.map(r => {
      const tracked = TRACKED.includes(r.picking);
      const pct = Math.max(1, Math.round(r.total / max * 100));
      return `
        <div class="q-bar-row ${tracked ? 'tracked' : ''}">
          <div class="q-bar-row-label">${r.picking}</div>
          <div class="q-bar-track"><i style="width:${pct}%"></i></div>
          <div class="q-bar-val">${r.total}</div>
        </div>`;
    }).join('');
  }

  // ── Full table ──
  const tableEl = document.getElementById('queueTableRows');
  if (tableEl) {
    setText('qTableMeta', `${sorted.length} rows`);
    tableEl.innerHTML = sorted.map((r, i) => {
      const tracked = TRACKED.includes(r.picking);
      const pct     = Math.max(2, Math.round(r.total / max * 100));
      const rowCls  = tracked ? 'blue-row' : 'orange-row';
      const priority = tracked
        ? '<span class="priority-high">● HIGH</span>'
        : r.total >= 20
          ? '<span class="priority-med">● MED</span>'
          : '<span class="priority-low">○ LOW</span>';
      const icon = TRACKED.includes(r.picking)
        ? (r.picking === 'SF Scan & Verify' ? 'SF' : r.picking === 'FSV Scan & Verify' ? 'FSV' : 'BR')
        : getQueueIcon(r.picking);
      return `
        <div class="wip-row ${rowCls}" style="grid-template-columns:1.8fr .9fr 1.1fr .8fr 1.25fr;animation-delay:${i*.03}s">
          <div class="wip-name"><span class="wip-icon">${icon}</span>${r.picking}</div>
          <div class="dept">${r.department || 'Inventory'}</div>
          <div class="wip-count">
            <strong>${r.total}</strong>
            <div class="row-bar"><i style="width:${pct}%"></i></div>
          </div>
          <div>${priority}</div>
          <div class="wip-date">${toEastern(r.lastUpdated)}</div>
        </div>`;
    }).join('');
  }
}

/* ── Activity Log Tab ────────────────────────────── */
function renderActivityTab() {
  const completion = STATE.completion;
  const history    = STATE.history || [];

  // Merge completion events + history snapshots into one log
  const filter = (document.getElementById('actFilter') || {}).value || '';

  // Use completion events as our activity source
  let events = [...completion];
  if (history.length) events = [...events, ...history];

  // Filter out ghost rows — must have a real picking type and meaningful data
  events = events.filter(e => {
    const p = e.picking || '';
    if (!p || p === '--' || p === 'Unknown') return false;
    // Must have at least one real numeric field
    const hasBefore = (e.before ?? e.wip_before) != null && (e.before ?? e.wip_before) !== '';
    const hasAfter  = (e.after  ?? e.wip_after)  != null && (e.after  ?? e.wip_after)  !== '';
    const hasJobs   = e.completed != null && e.completed !== '' && e.completed !== 0;
    return hasBefore || hasAfter || hasJobs;
  });

  // Sort newest first
  events = events.sort((a, b) => String(b.timestamp || b.ts || '').localeCompare(String(a.timestamp || a.ts || '')));

  if (filter) events = events.filter(e => e.picking === filter);

  // KPIs
  const completed = completion.filter(e => e.direction === 'completed').reduce((s,e)=>s+e.completed,0);
  const added     = completion.filter(e => e.direction !== 'completed').reduce((s,e)=>s+e.completed,0);
  const net       = added - completed;

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
    const picking  = e.picking || '--';
    const typeCls  = picking === 'SF Scan & Verify' ? 'sf'
                   : picking === 'FSV Scan & Verify' ? 'fsv'
                   : picking === 'Breakage to Picking' ? 'brk' : 'other';
    const dir      = (e.direction || 'completed').toLowerCase();
    const dirCls   = dir === 'completed' ? 'act-dir-completed' : 'act-dir-added';
    const dirLabel = dir === 'completed' ? '▼ Completed' : '▲ Added';
    const delta    = (e.after ?? e.wip_after) != null && (e.before ?? e.wip_before) != null
                   ? (e.after ?? e.wip_after) - (e.before ?? e.wip_before)
                   : dir === 'completed' ? -(e.completed||0) : (e.completed||0);
    const deltaCls = delta < 0 ? 'act-delta-down' : delta > 0 ? 'act-delta-up' : 'act-delta-zero';
    const deltaStr = delta > 0 ? `+${delta}` : String(delta);
    const ts       = e.timestamp || e.ts || '--';

    return `
      <div class="act-row" style="animation-delay:${Math.min(i,.25)*0.04}s">
        <div class="act-ts">${ts}</div>
        <div class="act-type ${typeCls}">${picking}</div>
        <div class="act-num">${e.before ?? e.wip_before ?? '--'}</div>
        <div class="act-num">${e.after ?? e.wip_after ?? '--'}</div>
        <div class="act-num">${e.completed ?? '--'}</div>
        <div class="${dirCls}">${dirLabel}</div>
        <div class="${deltaCls}">${deltaStr}</div>
      </div>`;
  }).join('');
}

/* ── Reports Tab ─────────────────────────────────── */
function renderReportsTab() {
  const wip        = STATE.wip;
  const completion = STATE.completion;
  const totalWip   = sum(wip, 'total');
  const sorted     = [...wip].sort((a, b) => b.total - a.total);

  // Valid completion events only
  const validEvents = completion.filter(e => {
    const p = e.picking || '';
    if (!p || p === '--') return false;
    const hasBefore = (e.before ?? e.wip_before) != null;
    const hasAfter  = (e.after  ?? e.wip_after)  != null;
    const hasJobs   = e.completed != null && e.completed !== 0;
    return hasBefore || hasAfter || hasJobs;
  }).sort((a,b) => String(b.timestamp||'').localeCompare(String(a.timestamp||'')));

  const completedEvents = validEvents.filter(e => e.direction === 'completed');
  const addedEvents     = validEvents.filter(e => e.direction !== 'completed');
  const totalCompleted  = completedEvents.reduce((s,e) => s + (e.completed||0), 0);
  const totalAdded      = addedEvents.reduce((s,e) => s + (e.completed||0), 0);
  const netChange       = totalAdded - totalCompleted;

  // Timestamp
  const now = new Date().toLocaleString('en-US', { weekday:'long', year:'numeric', month:'long', day:'numeric', hour:'2-digit', minute:'2-digit', second:'2-digit' });
  setText('rptGeneratedAt', `Generated ${now}`);

  // ── Executive KPI grid ──
  const kpiData = [
    { label:'Total WIP',        val: totalWip,       sub:'Across all queues',       cls:'blue'   },
    { label:'Completed Today',  val: totalCompleted, sub:'Jobs finished',            cls:'green'  },
    { label:'Added Today',      val: totalAdded,     sub:'Jobs entered queue',       cls:'orange' },
    { label:'Net Change',       val: Math.abs(netChange), sub: netChange <= 0 ? '▼ WIP decreasing' : '▲ WIP increasing', cls: netChange <= 0 ? 'green' : 'orange' },
    { label:'Active Lines',     val: wip.filter(r=>r.total>0).length, sub:'Non-zero queues', cls:'purple' },
    { label:'Log Events',       val: validEvents.length, sub:'Valid entries today',  cls:'blue'   },
  ];

  document.getElementById('rptKpiGrid').innerHTML = kpiData.map(k => `
    <div class="rpt-kpi-card ${k.cls}">
      <div class="rpt-kpi-val">${k.val.toLocaleString()}</div>
      <div class="rpt-kpi-label">${k.label}</div>
      <div class="rpt-kpi-sub">${k.sub}</div>
    </div>`).join('');

  // ── Tracked types detail ──
  const trackedHtml = TRACKED.map(name => {
    const wipRow   = wip.find(r => r.picking === name) || { total: 0 };
    const done     = completedEvents.filter(e=>e.picking===name).reduce((s,e)=>s+(e.completed||0),0);
    const added    = addedEvents.filter(e=>e.picking===name).reduce((s,e)=>s+(e.completed||0),0);
    const net      = added - done;
    const cls      = name==='SF Scan & Verify'?'sf':name==='FSV Scan & Verify'?'fsv':'brk';
    const short    = name==='SF Scan & Verify'?'SF':name==='FSV Scan & Verify'?'FSV':'BR';
    const pct      = totalWip ? (wipRow.total/totalWip*100).toFixed(1) : '0.0';
    return `
      <div class="rpt-tracked-row">
        <div class="rpt-tracked-badge ${cls}">${short}</div>
        <div class="rpt-tracked-name">${name}</div>
        <div class="rpt-tracked-stat"><span>Current WIP</span><strong class="${cls}">${wipRow.total}</strong></div>
        <div class="rpt-tracked-stat"><span>% of Total</span><strong>${pct}%</strong></div>
        <div class="rpt-tracked-stat"><span>Completed</span><strong style="color:var(--green)">${done}</strong></div>
        <div class="rpt-tracked-stat"><span>Added</span><strong style="color:var(--orange)">${added}</strong></div>
        <div class="rpt-tracked-stat"><span>Net</span><strong style="color:${net<=0?'var(--green)':'var(--red)'}">${net>0?'+':''}${net}</strong></div>
      </div>`;
  }).join('');
  document.getElementById('rptTrackedDetail').innerHTML = trackedHtml;

  // ── Full WIP snapshot table ──
  setText('rptSnapshotMeta', `${sorted.length} picking types`);
  const max = Math.max(...sorted.map(r=>r.total), 1);
  document.getElementById('rptWipTable').innerHTML = sorted.map((r,i) => {
    const tracked = TRACKED.includes(r.picking);
    const pct     = totalWip ? (r.total/totalWip*100).toFixed(1) : '0.0';
    const barPct  = Math.max(2, Math.round(r.total/max*100));
    return `
      <div class="wip-row ${tracked?'blue-row':'orange-row'}" style="grid-template-columns:1.8fr .9fr 1fr .8fr 1.4fr;animation-delay:${i*.03}s">
        <div class="wip-name"><span class="wip-icon">${tracked?(r.picking==='SF Scan & Verify'?'SF':r.picking==='FSV Scan & Verify'?'FSV':'BR'):getQueueIcon(r.picking)}</span>${r.picking}</div>
        <div class="dept">${r.department||'Inventory'}</div>
        <div class="wip-count"><strong>${r.total}</strong><div class="row-bar"><i style="width:${barPct}%"></i></div></div>
        <div style="color:var(--muted);font-family:'JetBrains Mono',monospace;font-size:13px">${pct}%</div>
        <div class="wip-date">${r.lastUpdated||'--'}</div>
      </div>`;
  }).join('');

  // ── Activity summary by type ──
  const byType = {};
  validEvents.forEach(e => {
    if (!byType[e.picking]) byType[e.picking] = { completed:0, added:0, events:0 };
    if (e.direction==='completed') byType[e.picking].completed += e.completed||0;
    else byType[e.picking].added += e.completed||0;
    byType[e.picking].events++;
  });
  const summaryRows = Object.entries(byType).sort((a,b)=>(b[1].completed+b[1].added)-(a[1].completed+a[1].added));
  document.getElementById('rptActivitySummary').innerHTML = summaryRows.length ? `
    <div class="rpt-summary-grid">
      <div class="rpt-summary-head"><span>Picking Type</span><span>Events</span><span>Completed</span><span>Added</span><span>Net</span></div>
      ${summaryRows.map(([name, d]) => {
        const net = d.added - d.completed;
        const tracked = TRACKED.includes(name);
        return `<div class="rpt-summary-row ${tracked?'tracked':''}">
          <span>${name}</span>
          <span style="color:var(--muted)">${d.events}</span>
          <span style="color:var(--green)">${d.completed}</span>
          <span style="color:var(--orange)">${d.added}</span>
          <span style="color:${net<=0?'var(--green)':'var(--red)'}">${net>0?'+':''}${net}</span>
        </div>`;
      }).join('')}
    </div>` : '<div class="error-log-empty"><span>◌</span>No activity data available</div>';

  // ── Narrative Day Summary ──
  setText('rptNarrativeMeta', `${validEvents.length} events · ${new Date().toLocaleDateString('en-US',{weekday:'long',month:'long',day:'numeric',year:'numeric'})}`);

  const narrativeEl = document.getElementById('rptNarrative');
  if (!narrativeEl) return;

  const today = new Date().toLocaleDateString('en-US', { weekday:'long', month:'long', day:'numeric', year:'numeric' });
  const timeNow = new Date().toLocaleTimeString('en-US', { hour:'2-digit', minute:'2-digit' });

  // Per-type stats
  const typeStats = {};
  TRACKED.forEach(name => { typeStats[name] = { completed:0, added:0, events:0, peakWip:0, latestWip:0 }; });
  validEvents.forEach(e => {
    const n = e.picking;
    if (!typeStats[n]) typeStats[n] = { completed:0, added:0, events:0, peakWip:0, latestWip:0 };
    if (e.direction === 'completed') typeStats[n].completed += e.completed||0;
    else typeStats[n].added += e.completed||0;
    typeStats[n].events++;
    const wipVal = e.after ?? e.wip_after ?? 0;
    if (wipVal > typeStats[n].peakWip) typeStats[n].peakWip = wipVal;
  });

  // Current WIP from live state
  STATE.wip.forEach(r => {
    if (typeStats[r.picking]) typeStats[r.picking].latestWip = r.total;
  });

  // Build narrative blocks
  const sfS  = typeStats['SF Scan & Verify']    || {};
  const fsvS = typeStats['FSV Scan & Verify']   || {};
  const brkS = typeStats['Breakage to Picking'] || {};

  const sfNet  = (sfS.added||0)  - (sfS.completed||0);
  const fsvNet = (fsvS.added||0) - (fsvS.completed||0);
  const brkNet = (brkS.added||0) - (brkS.completed||0);

  const overallNet = totalAdded - totalCompleted;
  const overallTrend = overallNet < 0 ? 'improving' : overallNet === 0 ? 'stable' : 'building';
  const trendColor   = overallNet < 0 ? 'var(--green)' : overallNet === 0 ? 'var(--cyan)' : 'var(--orange)';

  // Find busiest hour
  const hourBuckets = {};
  validEvents.forEach(e => {
    const ts = e.timestamp || e.ts || '';
    const m = ts.match(/T(\d{2}):/);
    if (m) {
      const h = parseInt(m[1]);
      hourBuckets[h] = (hourBuckets[h]||0) + (e.completed||0);
    }
  });
  const busiestHour = Object.entries(hourBuckets).sort((a,b)=>b[1]-a[1])[0];
  const busiestLabel = busiestHour
    ? `${parseInt(busiestHour[0]) % 12 || 12}:00 ${parseInt(busiestHour[0]) >= 12 ? 'PM' : 'AM'} (${busiestHour[1]} jobs)`
    : 'N/A';

  // Non-tracked queue highlights
  const otherQueues = STATE.wip.filter(r => !TRACKED.includes(r.picking) && r.total > 0)
    .sort((a,b) => b.total - a.total);
  const otherHighlight = otherQueues.slice(0,3).map(r => `<span class="narr-tag orange">${r.picking}: ${r.total}</span>`).join(' ');

  narrativeEl.innerHTML = `
    <div class="narr-block">
      <div class="narr-dateline">📋 &nbsp; ${today} &nbsp;·&nbsp; As of ${timeNow} &nbsp;·&nbsp; ${validEvents.length} logged events</div>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Overall Status</div>
      <p class="narr-text">
        As of <strong>${timeNow}</strong>, total picking WIP stands at
        <span class="narr-num cyan">${totalWip}</span> items across
        <span class="narr-num">${STATE.wip.filter(r=>r.total>0).length}</span> active queues.
        The floor processed <span class="narr-num green">${totalCompleted}</span> completions
        and added <span class="narr-num orange">${totalAdded}</span> new jobs today,
        for a net change of <span class="narr-num" style="color:${trendColor}">${overallNet > 0 ? '+' : ''}${overallNet}</span>.
        Overall WIP is <strong style="color:${trendColor}">${overallTrend}</strong>.
        Peak throughput hour was <strong>${busiestLabel}</strong>.
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">SF Scan &amp; Verify</div>
      <p class="narr-text">
        Current WIP: <span class="narr-num cyan">${sfS.latestWip ?? '--'}</span>.
        Today's completions: <span class="narr-num green">${sfS.completed||0}</span> jobs finished,
        <span class="narr-num orange">${sfS.added||0}</span> added — net
        <span class="narr-num" style="color:${sfNet<=0?'var(--green)':'var(--red)'}">${sfNet>0?'+':''}${sfNet}</span>.
        ${sfS.events ? `Logged <strong>${sfS.events}</strong> change events.` : 'No events recorded.'}
        ${(sfS.latestWip??0) > 150 ? '<span class="narr-tag red">⚠ High WIP — monitor closely</span>' :
          (sfS.latestWip??0) === 0 ? '<span class="narr-tag green">✓ Queue cleared</span>' :
          '<span class="narr-tag cyan">✓ Within normal range</span>'}
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">FSV Scan &amp; Verify</div>
      <p class="narr-text">
        Current WIP: <span class="narr-num" style="color:#54a8ff">${fsvS.latestWip ?? '--'}</span>.
        Today's completions: <span class="narr-num green">${fsvS.completed||0}</span> jobs finished,
        <span class="narr-num orange">${fsvS.added||0}</span> added — net
        <span class="narr-num" style="color:${fsvNet<=0?'var(--green)':'var(--red)'}">${fsvNet>0?'+':''}${fsvNet}</span>.
        ${fsvS.events ? `Logged <strong>${fsvS.events}</strong> change events.` : 'No events recorded.'}
        ${(fsvS.latestWip??0) > 80 ? '<span class="narr-tag red">⚠ Elevated WIP</span>' :
          (fsvS.latestWip??0) === 0 ? '<span class="narr-tag green">✓ Queue cleared</span>' :
          '<span class="narr-tag cyan">✓ Normal</span>'}
      </p>
    </div>

    <div class="narr-block">
      <div class="narr-section-title">Breakage to Picking</div>
      <p class="narr-text">
        Current WIP: <span class="narr-num orange">${brkS.latestWip ?? '--'}</span>.
        Today's completions: <span class="narr-num green">${brkS.completed||0}</span> jobs finished,
        <span class="narr-num orange">${brkS.added||0}</span> added — net
        <span class="narr-num" style="color:${brkNet<=0?'var(--green)':'var(--red)'}">${brkNet>0?'+':''}${brkNet}</span>.
        ${brkS.events ? `Logged <strong>${brkS.events}</strong> change events.` : 'No events recorded.'}
        ${(brkS.latestWip??0) > 20 ? '<span class="narr-tag red">⚠ Review breakage volume</span>' :
          (brkS.latestWip??0) === 0 ? '<span class="narr-tag green">✓ No breakage backlog</span>' :
          '<span class="narr-tag cyan">✓ Low volume</span>'}
      </p>
    </div>

    ${otherQueues.length ? `
    <div class="narr-block">
      <div class="narr-section-title">Other Active Queues</div>
      <p class="narr-text">
        ${otherQueues.length} additional queue${otherQueues.length>1?'s':''} currently active:
        ${otherHighlight}
        ${otherQueues.length > 3 ? `<span class="narr-tag gray">+${otherQueues.length-3} more</span>` : ''}
      </p>
    </div>` : ''}

    <div class="narr-block narr-footer">
      <span>⬡ Auto-generated by Picking WIP Command Center</span>
      <span>Report time: ${new Date().toLocaleString()}</span>
    </div>
  `;
}


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
      text:'<span class="neg">0</span> no change',
      value:0,
      up:false,
      icon:'▼'
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
    text:`<span class="${up ? 'pos' : 'neg'}">${up ? '+' : '-'} ${val}</span> ${label}`,
    value:val,
    up,
    icon:up ? '▲' : '▼'
  };
}

function getOverallTrend(completion) {
  const latest = TRACKED.map(t => getChange(getLatestEventFor(t, completion), t));

  const up = latest
    .filter(x => x.up)
    .reduce((s, x) => s + x.value, 0);

  const down = latest
    .filter(x => !x.up)
    .reduce((s, x) => s + x.value, 0);

  if (up > down) {
    return {
      arrow:'▲',
      value:up - down,
      text:'Backlog increasing'
    };
  }

  return {
    arrow:'▲',
    value:Math.max(1, down - up),
    text:'Overall improving'
  };
}

function getQueueIcon(name) {
  if (/surface|finish/i.test(name)) return '◷';
  if (/re-cal/i.test(name)) return '⚒';
  if (/frame/i.test(name)) return '▢';
  if (/hold/i.test(name)) return '▣';
  return '•';
}

function toEastern(ts) {
  if (!ts) return '--';
  try {
    const d = new Date(ts);
    if (isNaN(d)) return ts;
    return d.toLocaleString('en-US', {
      timeZone: 'America/New_York',
      month: 'short', day: 'numeric', year: 'numeric',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
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
  const n = parseInt(v, 10);
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

  const current = parseInt(String(el.textContent).replace(/,/g, ''), 10) || 0;

  if (current === target) {
    el.textContent = target.toLocaleString();
    return;
  }

  // Flash the parent KPI card border
  const card = el.closest('.kpi-card');
  if (card) {
    card.classList.remove('flash-update');
    void card.offsetWidth; // reflow to restart animation
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
      // num-pop flash at end
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
/* ══════════════════════════════════════════════════
   MOTION LAYER — Particles, Ticker, Live Effects
   ══════════════════════════════════════════════════ */

function initParticles() {
  const count = 18;
  const colors = ['rgba(25,200,255,', 'rgba(180,76,255,', 'rgba(255,159,28,', 'rgba(33,229,124,'];

  for (let i = 0; i < count; i++) {
    const p = document.createElement('div');
    p.className = 'data-particle';

    const size   = Math.random() * 3 + 1.5;
    const left   = Math.random() * 100;
    const delay  = Math.random() * 18;
    const dur    = 14 + Math.random() * 18;
    const drift  = (Math.random() - 0.5) * 120;
    const color  = colors[Math.floor(Math.random() * colors.length)];
    const opacity = 0.15 + Math.random() * 0.3;

    p.style.cssText = `
      width:${size}px;height:${size}px;
      left:${left}%;bottom:-10px;
      background:${color}${opacity});
      box-shadow:0 0 ${size*3}px ${color}0.5);
      --drift:${drift}px;
      animation-duration:${dur}s;
      animation-delay:${delay}s;
    `;
    document.body.appendChild(p);
  }
}

function updateTicker(wip, completion) {
  const sf  = (wip.find(r => r.picking === 'SF Scan & Verify')  || {}).total ?? '--';
  const fsv = (wip.find(r => r.picking === 'FSV Scan & Verify') || {}).total ?? '--';
  const brk = (wip.find(r => r.picking === 'Breakage to Picking') || {}).total ?? '--';
  const tot = sum(wip, 'total');
  const done= completion.filter(e=>e.direction==='completed').reduce((s,e)=>s+e.completed,0);
  const lines = wip.filter(r=>r.total>0).length;

  ['tk-sf','tk-sf2'].forEach(id => setText(id, sf));
  ['tk-fsv','tk-fsv2'].forEach(id => setText(id, fsv));
  ['tk-brk','tk-brk2'].forEach(id => setText(id, brk));
  ['tk-total','tk-total2'].forEach(id => setText(id, tot));
  ['tk-done','tk-done2'].forEach(id => setText(id, done));
  ['tk-lines','tk-lines2'].forEach(id => setText(id, lines));
}

// Hook ticker updates into renderAll
const _origRenderAll = renderAll;
// eslint-disable-next-line no-global-assign
window._motionRenderAll = function() {
  // called after renderAll finishes via monkey-patch below
  updateTicker(STATE.wip, STATE.completion);
};

// Particle and row bar width trigger — run once on load
document.addEventListener('DOMContentLoaded', () => {
  initParticles();

  // Trigger row-bar widths after first render (CSS transition from 0)
  setTimeout(() => {
    document.querySelectorAll('.row-bar i, .progress-mini i, .q-card-bar i').forEach(el => {
      const w = el.style.width;
      el.style.width = '0';
      requestAnimationFrame(() => { el.style.width = w; });
    });
  }, 300);
});

// Patch renderAll to also update ticker after every data refresh
(function patchRenderAll() {
  const orig = window.renderAll || renderAll;
  if (typeof orig !== 'function') return;

  // We can't reassign const, so we hook into fetchAll completing instead
  // updateTicker is called directly from the global renderAll override below
})();