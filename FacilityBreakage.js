/* ═══════════════════════════════════════════════════════════
   ZENNI FACILITY BREAKAGE — app.js
   Requires: Chart.js 4.x loaded before this file
═══════════════════════════════════════════════════════════ */

'use strict';

/* ── CONFIG ─────────────────────────────────────────────── */
const API_URL = 'https://script.google.com/macros/s/AKfycbylW0fs7zWLncknhz7peJcCm9eyWCXTGCLtk-xtMIzjarot5PcCpmCP4Gy85WqT3f17/exec';

const GOALS = {
  'Lab Total':      5.00,
  'AR':             0.85,
  'Finish':         1.70,
  'Surface':        2.80,
  'Frame Breakage': 1.00,
};

const DEPT_COLORS = {
  AR:        '#a78bfa',
  Finish:    '#34d399',
  Surface:   '#fb923c',
  LMS:       '#60a5fa',
  Inventory: '#f472b6',
};

const CHART_DEFAULTS = {
  gridColor:   'rgba(255,255,255,0.05)',
  textColor:   '#9aafc4',
  tooltipBg:   '#131a24',
  tooltipBorder:'rgba(255,255,255,0.1)',
  fontMono:    'DM Mono, monospace',
};

/* ── CHART REGISTRY ─────────────────────────────────────── */
const Charts = {};
function destroyChart(id) {
  if (Charts[id]) { Charts[id].destroy(); delete Charts[id]; }
}

/* ── STATE ──────────────────────────────────────────────── */
let State = {
  meta:    {},
  summary: {},
  depts:   {},
  history: [],
};

/* ── CHART.JS GLOBAL DEFAULTS ───────────────────────────── */
Chart.defaults.color          = CHART_DEFAULTS.textColor;
Chart.defaults.font.family    = CHART_DEFAULTS.fontMono;
Chart.defaults.font.size      = 11;
Chart.defaults.plugins.legend.display = false;
Chart.defaults.plugins.tooltip.backgroundColor = CHART_DEFAULTS.tooltipBg;
Chart.defaults.plugins.tooltip.borderColor     = CHART_DEFAULTS.tooltipBorder;
Chart.defaults.plugins.tooltip.borderWidth     = 1;
Chart.defaults.plugins.tooltip.titleColor      = CHART_DEFAULTS.textColor;
Chart.defaults.plugins.tooltip.bodyColor       = '#e8edf2';
Chart.defaults.plugins.tooltip.padding         = 10;
Chart.defaults.scale.grid.color                = CHART_DEFAULTS.gridColor;
Chart.defaults.scale.ticks.color               = CHART_DEFAULTS.textColor;

/* ═══════════════════════════════════════════════════════════
   API LAYER
═══════════════════════════════════════════════════════════ */
const API = {
  async fetch(action, extra = '') {
    const res  = await fetch(`${API_URL}?action=${action}${extra}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || `API error: ${action}`);
    return json;
  },

  async all()      { return this.fetch('all'); },
  async history()  { return this.fetch('snapshotHistory'); },

  async snapshot(name) {
    const res  = await fetch(`${API_URL}?action=takeSnapshot&snappedBy=${encodeURIComponent(name)}`, { method: 'POST' });
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    return json;
  },
};

/* ═══════════════════════════════════════════════════════════
   UTILITY HELPERS
═══════════════════════════════════════════════════════════ */
const U = {
  fmt:    n  => Number(n || 0).toLocaleString(),
  pct:    d  => (d * 100).toFixed(2) + '%',
  pctRaw: v  => v.toFixed(2) + '%',

  statusColor: (val, goal) => val <= goal ? 'var(--green)' : 'var(--red)',

  statusBadge(val, goal) {
    const ok = val <= goal;
    return `<div class="status-badge ${ok ? 'green' : 'red'}">${ok ? 'ON GOAL' : 'OVER GOAL'}</div>`;
  },

  deltaHtml(today, prev) {
    if (prev === null || prev === undefined || prev === 0) return '';
    const diff   = today - prev;
    const diffPct = prev !== 0 ? ((diff / prev) * 100).toFixed(1) : '0.0';
    // For breakage: going up is bad (red), going down is good (green)
    if (diff > 0) return `
      <div class="kpi-delta up">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>
        +${diffPct}% vs prev
      </div>`;
    if (diff < 0) return `
      <div class="kpi-delta down">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>
        ${diffPct}% vs prev
      </div>`;
    return `<div class="kpi-delta flat">— no change</div>`;
  },

  emptyState: msg => `<div class="empty-state">${msg || 'No data available'}</div>`,

  parseHistoryPct(val) {
    if (val === '' || val === undefined || val === null) return null;
    const s = val.toString().replace('%', '').trim();
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  },

  shortDate(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    if (isNaN(d)) return dateStr;
    return (d.getMonth() + 1) + '/' + d.getDate();
  },
};

/* ═══════════════════════════════════════════════════════════
   CHART FACTORY — reusable builders
═══════════════════════════════════════════════════════════ */
const ChartFactory = {

  bar(id, labels, datasets, opts = {}) {
    destroyChart(id);
    const ctx = document.getElementById(id);
    if (!ctx) return;
    Charts[id] = new Chart(ctx, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: opts.legend || false } },
        scales: {
          x: {
            grid: { display: false },
            ticks: { maxRotation: 35, font: { size: 10 } },
          },
          y: {
            beginAtZero: true,
            ticks: { font: { size: 10 } },
          },
          ...(opts.scales || {}),
        },
        ...opts.extra,
      },
    });
  },

  line(id, labels, datasets, opts = {}) {
    destroyChart(id);
    const ctx = document.getElementById(id);
    if (!ctx) return;
    Charts[id] = new Chart(ctx, {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: {
            display: opts.legend || false,
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#9aafc4' },
          },
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 10 } } },
          y: {
            beginAtZero: true,
            ticks: {
              font: { size: 10 },
              callback: v => opts.isPct ? v.toFixed(2) + '%' : v,
            },
          },
        },
        ...opts.extra,
      },
    });
  },

  doughnut(id, labels, data, colors) {
    destroyChart(id);
    const ctx = document.getElementById(id);
    if (!ctx) return;
    Charts[id] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{ data, backgroundColor: colors, borderWidth: 0, hoverOffset: 8 }],
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '62%',
        plugins: {
          legend: {
            display: true, position: 'bottom',
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, color: '#9aafc4', font: { size: 11 } },
          },
        },
      },
    });
  },

  lineDataset(label, data, color, opts = {}) {
    return {
      label,
      data,
      borderColor:           color,
      backgroundColor:       color + '20',
      pointBackgroundColor:  color,
      pointBorderColor:      '#080b0f',
      pointBorderWidth:      2,
      pointRadius:           4,
      pointHoverRadius:      6,
      borderWidth:           2,
      tension:               0.4,
      fill:                  opts.fill || false,
      ...opts,
    };
  },
};

/* ═══════════════════════════════════════════════════════════
   RENDERERS
═══════════════════════════════════════════════════════════ */

/* ── META BAR ───────────────────────────────────────────── */
function renderMeta(meta) {
  document.getElementById('metaOrders').textContent  = U.fmt(meta.orderCount);
  document.getElementById('metaLenses').textContent  = U.fmt(meta.lensCount);
  document.getElementById('reportDate').textContent  = meta.reportDate  || '—';
  document.getElementById('lastRefresh').textContent = meta.lastRefresh || '—';
  document.getElementById('headerSub').textContent   = `Report: ${meta.reportDate || '—'}  ·  ${U.fmt(meta.lensCount)} lenses`;
}

/* ── SUMMARY TAB ────────────────────────────────────────── */
function renderSummary(summary, meta, history) {
  const lt        = summary.labTotal;
  const labPct    = lt.labLensPct   * 100;
  const framePct  = lt.labFramePct  * 100;

  // get previous day's snapshot for delta
  const prev = history.length >= 2 ? history[history.length - 2] : null;
  const prevLabPct   = prev ? U.parseHistoryPct(prev['Lab Lens %'])   : null;
  const prevFramePct = prev ? U.parseHistoryPct(prev['Lab Frame %'])  : null;

  /* Lab KPIs */
  document.getElementById('labKpis').innerHTML = [
    { label: 'Lab Lenses Broken', value: U.fmt(lt.labLensesBroken), sub: `of ${U.fmt(lt.lensCount)} lenses`, accent: 'var(--cyan)',  delta: '' },
    { label: 'Lab Lens %',   value: U.pctRaw(labPct),   sub: 'Goal ≤5.00%', accent: U.statusColor(labPct,  5.00), badge: U.statusBadge(labPct, 5.00),  delta: U.deltaHtml(labPct, prevLabPct) },
    { label: 'Frames Broken', value: U.fmt(lt.framesBroken), sub: `of ${U.fmt(lt.orderCount)} orders`, accent: 'var(--amber)', delta: '' },
    { label: 'Frame Brk %',  value: U.pctRaw(framePct), sub: 'Goal ≤1.00%', accent: U.statusColor(framePct, 1.00), badge: U.statusBadge(framePct, 1.00), delta: U.deltaHtml(framePct, prevFramePct) },
    { label: 'Lens Count',   value: U.fmt(lt.lensCount),   sub: 'Total lenses today',  accent: 'var(--lms)',  delta: '' },
    { label: 'Order Count',  value: U.fmt(lt.orderCount),  sub: 'Mailroom orders',     accent: 'var(--mail)', delta: '' },
  ].map(k => `
    <div class="kpi-card" style="--accent:${k.accent}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
      ${k.badge  || ''}
      ${k.delta  || ''}
    </div>`).join('');

  /* Goal bars */
  const deptMap = {};
  summary.departments.forEach(d => { deptMap[d.department] = d; });

  const goalRows = [
    { label: 'Lab Total',      pct: labPct,                              target: 5.00, color: 'var(--cyan)' },
    { label: 'AR',             pct: (deptMap.AR?.lensBrkPct    || 0) * 100, target: 0.85, color: 'var(--ar)'   },
    { label: 'Finish',         pct: (deptMap.Finish?.lensBrkPct  || 0) * 100, target: 1.70, color: 'var(--fin)'  },
    { label: 'Surface',        pct: (deptMap.Surface?.lensBrkPct || 0) * 100, target: 2.80, color: 'var(--srf)'  },
    { label: 'Frame Breakage', pct: framePct,                            target: 1.00, color: 'var(--mail)' },
  ];
  const maxPct = Math.max(...goalRows.map(r => Math.max(r.pct, r.target))) * 1.25 || 8;

  document.getElementById('goalBars').innerHTML = goalRows.map(r => {
    const fillW   = Math.min((r.pct    / maxPct) * 100, 100);
    const markerW = Math.min((r.target / maxPct) * 100, 100);
    const ok      = r.pct <= r.target;
    const fill    = ok ? 'var(--green)' : 'var(--red)';
    return `
      <div class="goal-row">
        <div class="goal-dept">${r.label}</div>
        <div class="goal-track">
          <div class="goal-fill" style="width:${fillW}%;background:${fill}"></div>
          <div class="goal-marker" style="left:${markerW}%"></div>
        </div>
        <div class="goal-pct" style="color:${fill}">${r.pct.toFixed(2)}%</div>
        <div class="goal-target">Goal ${r.target}%</div>
      </div>`;
  }).join('');

  /* Dept cards */
  document.getElementById('deptCards').innerHTML = summary.departments.map(d => {
    const color = DEPT_COLORS[d.department] || 'var(--cyan)';
    return `
      <div class="dept-card" data-dept="${d.department}"
           style="border-top:2px solid ${color}35"
           onclick="App.switchTab('${d.department.toLowerCase()}')">
        <div class="dept-name">${d.department}</div>
        <div class="dept-broken" style="color:${color}">${U.fmt(d.lensesBroken)}</div>
        <div class="dept-pct"   style="color:${color}">${U.pct(d.lensBrkPct)}</div>
        ${d.framesBroken > 0 ? `<div style="font-family:var(--font-mono);font-size:10px;color:var(--muted)">${U.fmt(d.framesBroken)} frames</div>` : ''}
        <div class="dept-top">↑ ${d.topReason || '—'}</div>
      </div>`;
  }).join('');

  /* Top reasons panels */
  const deptKeys = ['AR', 'Finish', 'Surface', 'LMS'];
  const deptClrs = ['var(--ar)', 'var(--fin)', 'var(--srf)', 'var(--lms)'];
  document.getElementById('topReasonsGrid').innerHTML = deptKeys.map((dept, i) => {
    const reasons = summary.topReasons[dept] || [];
    if (!reasons.length) return '';
    const maxB = reasons[0]?.lensesBroken || 1;
    return `
      <div class="panel">
        <div class="panel-header">
          <div class="panel-title" style="color:${deptClrs[i]}">${dept} — Top Reasons</div>
        </div>
        <div class="panel-body">
          <table class="data-table">
            <thead><tr><th>#</th><th>Reason</th><th style="text-align:right">Broken</th><th style="text-align:right">%</th></tr></thead>
            <tbody>
              ${reasons.map((r, ri) => `
                <tr>
                  <td class="rank ${ri === 0 ? 'r1' : ri === 1 ? 'r2' : 'r3'}">${r.rank}</td>
                  <td>
                    <div class="reason-name">${r.reason}</div>
                    <div class="mini-bar">
                      <div class="mini-bar-fill" style="width:${(r.lensesBroken/maxB)*100}%;background:${deptClrs[i]}"></div>
                    </div>
                  </td>
                  <td class="count" style="color:${deptClrs[i]}">${r.lensesBroken}</td>
                  <td class="pct-cell">${((r.lensesBroken||0) / (State.meta.lensCount||1) * 100).toFixed(2)}%</td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }).filter(Boolean).join('');
}

/* ── DAILY SUMMARY TAB ──────────────────────────────────── */
function renderDaily(summary, history) {
  const lt      = summary.labTotal;
  const labPct  = lt.labLensPct  * 100;
  const framePct = lt.labFramePct * 100;

  const deptMap = {};
  summary.departments.forEach(d => { deptMap[d.department] = d; });

  // Get yesterday from history
  const prev = history.length >= 1 ? history[history.length - 1] : null;
  const prevLabPct    = prev ? U.parseHistoryPct(prev['Lab Lens %'])    : null;
  const prevFramePct  = prev ? U.parseHistoryPct(prev['Lab Frame %'])   : null;
  const prevARPct     = prev ? U.parseHistoryPct(prev['AR %'])          : null;
  const prevFinPct    = prev ? U.parseHistoryPct(prev['Fin %'])         : null;
  const prevSrfPct    = prev ? U.parseHistoryPct(prev['Srf %'])         : null;

  document.getElementById('dailyDate').textContent = prev
    ? `Today vs ${prev['Date'] || 'Previous'}`
    : 'Today (no prior snapshot)';

  const arPct  = (deptMap.AR?.lensBrkPct    || 0) * 100;
  const finPct = (deptMap.Finish?.lensBrkPct  || 0) * 100;
  const srfPct = (deptMap.Surface?.lensBrkPct || 0) * 100;

  /* Daily KPIs with deltas */
  document.getElementById('dailyKpis').innerHTML = [
    { label: 'Lab Lens %',     value: U.pctRaw(labPct),   accent: U.statusColor(labPct,  5.00), delta: U.deltaHtml(labPct,  prevLabPct),   badge: U.statusBadge(labPct,  5.00) },
    { label: 'AR %',           value: U.pctRaw(arPct),    accent: U.statusColor(arPct,   0.85), delta: U.deltaHtml(arPct,   prevARPct),    badge: U.statusBadge(arPct,   0.85) },
    { label: 'Finish %',       value: U.pctRaw(finPct),   accent: U.statusColor(finPct,  1.70), delta: U.deltaHtml(finPct,  prevFinPct),   badge: U.statusBadge(finPct,  1.70) },
    { label: 'Surface %',      value: U.pctRaw(srfPct),   accent: U.statusColor(srfPct,  2.80), delta: U.deltaHtml(srfPct,  prevSrfPct),   badge: U.statusBadge(srfPct,  2.80) },
    { label: 'Frame Brk %',    value: U.pctRaw(framePct), accent: U.statusColor(framePct, 1.00), delta: U.deltaHtml(framePct, prevFramePct), badge: U.statusBadge(framePct, 1.00) },
    { label: 'Total Broken',   value: U.fmt(lt.labLensesBroken), accent: 'var(--cyan)', delta: '', badge: '' },
  ].map(k => `
    <div class="kpi-card" style="--accent:${k.accent}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
      ${k.badge}${k.delta}
    </div>`).join('');

  /* Dept comparison bar chart */
  const depts  = ['AR', 'Finish', 'Surface', 'LMS'];
  const colors = ['#a78bfa', '#34d399', '#fb923c', '#60a5fa'];

  const todayVals = depts.map(d => {
    const dept = summary.departments.find(x => x.department === d);
    return dept ? parseFloat((dept.lensBrkPct * 100).toFixed(3)) : 0;
  });
  const prevVals = depts.map((d, i) => {
    const keys = { AR: 'AR %', Finish: 'Fin %', Surface: 'Srf %', LMS: null };
    if (!keys[d] || !prev) return 0;
    return U.parseHistoryPct(prev[keys[d]]) || 0;
  });

  ChartFactory.bar('dailyDeptChart',
    depts,
    [
      {
        label: 'Today',
        data: todayVals,
        backgroundColor: colors.map(c => c + 'cc'),
        borderRadius: 5, borderSkipped: false,
      },
      {
        label: 'Previous Snapshot',
        data: prevVals,
        backgroundColor: colors.map(() => 'rgba(255,255,255,0.08)'),
        borderRadius: 5, borderSkipped: false,
      },
    ],
    {
      legend: true,
      extra: {
        plugins: {
          legend: {
            display: true,
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, color: '#9aafc4' },
          },
          tooltip: {
            callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.raw.toFixed(2)}%` },
          },
        },
        scales: {
          y: { ticks: { callback: v => v.toFixed(2) + '%' } },
        },
      },
    }
  );

  /* Top reasons today (cross-dept) */
  const allReasons = [];
  const byReason   = summary.topReasons;
  ['AR', 'Finish', 'Surface', 'LMS'].forEach(dept => {
    (byReason[dept] || []).forEach(r => {
      allReasons.push({ ...r, dept, color: DEPT_COLORS[dept] });
    });
  });
  allReasons.sort((a, b) => b.lensesBroken - a.lensesBroken);
  const topAll = allReasons.slice(0, 10);
  const maxB   = topAll[0]?.lensesBroken || 1;

  document.getElementById('dailyTopReasons').innerHTML = topAll.length
    ? `<table class="data-table">
        <thead><tr><th>Reason</th><th>Dept</th><th style="text-align:right">Broken</th></tr></thead>
        <tbody>
          ${topAll.map((r, i) => `
            <tr>
              <td>
                <div class="reason-name">${r.reason}</div>
                <div class="mini-bar">
                  <div class="mini-bar-fill" style="width:${(r.lensesBroken/maxB)*100}%;background:${r.color}"></div>
                </div>
              </td>
              <td><span class="op-reason" style="color:${r.color}">${r.dept}</span></td>
              <td class="count" style="color:${r.color};text-align:right">${r.lensesBroken}</td>
            </tr>`).join('')}
        </tbody>
      </table>`
    : U.emptyState('No reason data available');

  /* Goal status cards */
  const goalData = [
    { label: 'Lab Total',      pct: labPct,  target: 5.00 },
    { label: 'AR',             pct: arPct,   target: 0.85 },
    { label: 'Finish',         pct: finPct,  target: 1.70 },
    { label: 'Surface',        pct: srfPct,  target: 2.80 },
    { label: 'Frame Breakage', pct: framePct, target: 1.00 },
  ];

  document.getElementById('goalStatusGrid').innerHTML = goalData.map(g => {
    const ok = g.pct <= g.target;
    return `
      <div class="goal-status-card ${ok ? 'on-goal' : 'over-goal'}">
        <div class="gs-status">${ok ? '✓' : '✗'}</div>
        <div class="gs-label">${g.label}</div>
        <div class="gs-value">${g.pct.toFixed(2)}%</div>
        <div class="gs-target">Goal: ≤${g.target}%</div>
      </div>`;
  }).join('');
}

/* ── WEEKLY SUMMARY TAB ─────────────────────────────────── */
function renderWeekly(history) {
  if (!history || history.length === 0) {
    ['weeklyKpis', 'weeklyReasons'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = U.emptyState('No snapshot history yet. Take daily snapshots to see weekly trends.');
    });
    return;
  }

  // Use last 7 snapshots
  const recent = history.slice(-7);
  const labels = recent.map(r => U.shortDate(r['Date']));

  document.getElementById('weeklyRange').textContent =
    `${recent.length} snapshots${recent.length >= 2 ? ` · ${U.shortDate(recent[0]['Date'])} – ${U.shortDate(recent[recent.length-1]['Date'])}` : ''}`;

  /* Extract series */
  const labSeries  = recent.map(r => U.parseHistoryPct(r['Lab Lens %'])  || 0);
  const arSeries   = recent.map(r => U.parseHistoryPct(r['AR %'])        || 0);
  const finSeries  = recent.map(r => U.parseHistoryPct(r['Fin %'])       || 0);
  const srfSeries  = recent.map(r => U.parseHistoryPct(r['Srf %'])       || 0);
  const frameSeries= recent.map(r => U.parseHistoryPct(r['Lab Frame %']) || 0);
  const volSeries  = recent.map(r => parseFloat(r['Lab Lenses'] || 0));

  /* Weekly averages for KPIs */
  const avg = arr => arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length) : 0;
  const trend = arr => {
    if (arr.length < 2) return 0;
    return arr[arr.length - 1] - arr[arr.length - 2];
  };

  const avgLab  = avg(labSeries);
  const avgAR   = avg(arSeries);
  const avgFin  = avg(finSeries);
  const avgSrf  = avg(srfSeries);

  document.getElementById('weeklyKpis').innerHTML = [
    { label: 'Avg Lab Lens %',  value: avgLab.toFixed(2)  + '%', accent: U.statusColor(avgLab, 5.00),  delta: U.deltaHtml(labSeries.at(-1), labSeries.at(-2)) },
    { label: 'Avg AR %',        value: avgAR.toFixed(2)   + '%', accent: U.statusColor(avgAR,  0.85),  delta: U.deltaHtml(arSeries.at(-1),  arSeries.at(-2))  },
    { label: 'Avg Finish %',    value: avgFin.toFixed(2)  + '%', accent: U.statusColor(avgFin, 1.70),  delta: U.deltaHtml(finSeries.at(-1), finSeries.at(-2)) },
    { label: 'Avg Surface %',   value: avgSrf.toFixed(2)  + '%', accent: U.statusColor(avgSrf, 2.80),  delta: U.deltaHtml(srfSeries.at(-1), srfSeries.at(-2)) },
    { label: 'Snapshots',       value: recent.length,             accent: 'var(--muted)',                delta: '' },
  ].map(k => `
    <div class="kpi-card" style="--accent:${k.accent}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
      ${k.delta}
    </div>`).join('');

  /* Weekly trend chart — dept % lines */
  const legendEl = document.getElementById('weeklyLegend');
  if (legendEl) {
    const items = [
      { label: 'Lab Total', color: '#38bdf8' },
      { label: 'AR',        color: '#a78bfa' },
      { label: 'Finish',    color: '#34d399' },
      { label: 'Surface',   color: '#fb923c' },
    ];
    legendEl.innerHTML = items.map(i => `
      <div class="legend-item">
        <div class="legend-dot" style="background:${i.color}"></div>
        ${i.label}
      </div>`).join('');
  }

  ChartFactory.line('weeklyTrendChart', labels, [
    ChartFactory.lineDataset('Lab Total', labSeries, '#38bdf8', { fill: true }),
    ChartFactory.lineDataset('AR',        arSeries,  '#a78bfa'),
    ChartFactory.lineDataset('Finish',    finSeries, '#34d399'),
    ChartFactory.lineDataset('Surface',   srfSeries, '#fb923c'),
  ], {
    isPct: true, legend: true,
    extra: {
      plugins: {
        legend: {
          display: true,
          position: 'bottom',
          labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#9aafc4' },
        },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${ctx.raw.toFixed(2)}%` } },
      },
      scales: {
        y: {
          ticks: { callback: v => v.toFixed(2) + '%' },
          // Draw goal lines via plugin if needed
        },
      },
    },
  });

  /* Volume bar chart */
  ChartFactory.bar('weeklyVolumeChart', labels, [{
    label: 'Lenses Broken',
    data: volSeries,
    backgroundColor: volSeries.map((v, i) => {
      // Gradient by index
      const opacity = 0.4 + (i / volSeries.length) * 0.5;
      return `rgba(56,189,248,${opacity.toFixed(2)})`;
    }),
    borderRadius: 5, borderSkipped: false,
  }], {
    extra: {
      plugins: { tooltip: { callbacks: { label: ctx => ` ${ctx.raw} lenses broken` } } },
    },
  });

  /* Frame breakage trend */
  ChartFactory.line('weeklyFrameChart', labels, [
    ChartFactory.lineDataset('Frame Brk %', frameSeries, '#facc15', { fill: true }),
  ], {
    isPct: true,
    extra: {
      plugins: {
        tooltip: { callbacks: { label: ctx => ` Frame Brk: ${ctx.raw.toFixed(2)}%` } },
        annotation: {}, // could add goal line plugin if desired
      },
      scales: {
        y: { ticks: { callback: v => v.toFixed(2) + '%' } },
      },
    },
  });

  /* Dept contribution pie */
  const latestAR  = arSeries.at(-1)  || 0;
  const latestFin = finSeries.at(-1) || 0;
  const latestSrf = srfSeries.at(-1) || 0;
  ChartFactory.doughnut(
    'weeklyDeptPieChart',
    ['AR', 'Finish', 'Surface'],
    [latestAR, latestFin, latestSrf],
    ['#a78bfa', '#34d399', '#fb923c']
  );

  /* Recurring top reasons from AR Top Reason, Fin Top Reason, Srf Top Reason columns */
  const reasonCount = {};
  recent.forEach(r => {
    ['AR Top Reason', 'Fin Top Reason', 'Srf Top Reason'].forEach(col => {
      const val = r[col];
      if (val && val.trim()) {
        reasonCount[val] = (reasonCount[val] || 0) + 1;
      }
    });
  });
  const sortedReasons = Object.entries(reasonCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);
  const maxRCount = sortedReasons[0]?.[1] || 1;

  document.getElementById('weeklyReasons').innerHTML = sortedReasons.length
    ? sortedReasons.map(([reason, count]) => {
        const dept = reason.startsWith('A-') ? 'AR'
                   : reason.startsWith('F-') ? 'Finish'
                   : reason.startsWith('S-') ? 'Surface' : 'Lab';
        const color = DEPT_COLORS[dept] || 'var(--cyan)';
        return `
          <div class="weekly-reason-row">
            <div class="wr-name">${reason}</div>
            <div class="wr-dept" style="color:${color}">${dept}</div>
            <div class="wr-bar"><div class="wr-fill" style="width:${(count/maxRCount)*100}%;background:${color}"></div></div>
            <div class="wr-count" style="color:${color}">${count}</div>
          </div>`;
      }).join('')
    : U.emptyState('Take more snapshots to see recurring reasons');
}

/* ── DEPARTMENT DETAIL TABS ─────────────────────────────── */
function renderDeptTab(tabId, data, deptName, color, chartId) {
  if (!data) {
    ['Kpis', 'ReasonsTable', 'Operators'].forEach(suffix => {
      const el = document.querySelector(`#tab-${tabId} [id$="${suffix}"]`);
      if (el) el.innerHTML = U.emptyState('No data available');
    });
    return;
  }

  const { totals, reasons, operators } = data;

  // Calculate % directly from counts — avoids sheet formula inconsistency
  const lensCount  = State.meta.lensCount  || 1;
  const orderCount = State.meta.orderCount || 1;
  const lensPct    = totals.lensesBroken > 0 ? (totals.lensesBroken / lensCount * 100) : 0;
  const framePctDisp = totals.framesBroken > 0 ? (totals.framesBroken / orderCount * 100) : 0;

  /* Badge */
  const badgeEl = document.querySelector(`#tab-${tabId} .section-badge`);
  if (badgeEl) badgeEl.textContent = `${U.fmt(totals.lensesBroken)} broken · ${lensPct.toFixed(2)}%`;

  /* KPIs */
  const kpiEl = document.querySelector(`#tab-${tabId} [id$="Kpis"]`);
  if (kpiEl) {
    const goal = { ar: 0.85, finish: 1.70, surface: 2.80 }[tabId];
    kpiEl.innerHTML = [
      { label: 'Lenses Broken',  value: U.fmt(totals.lensesBroken), accent: color, sub: '' },
      { label: 'Lens Brk %',     value: lensPct.toFixed(2) + '%', accent: U.statusColor(lensPct, goal || 999), sub: goal ? `Goal ≤${goal}%` : '', badge: goal ? U.statusBadge(lensPct, goal) : '' },
      { label: 'Frames Broken',  value: U.fmt(totals.framesBroken), accent: 'var(--amber)', sub: totals.framesBroken > 0 ? framePctDisp.toFixed(2) + '%' : '—' },
      { label: 'Unique Reasons', value: reasons.filter(r => r.lensesBroken > 0 || r.framesBroken > 0).length, accent: 'var(--muted)', sub: '' },
    ].map(k => `
      <div class="kpi-card" style="--accent:${k.accent}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
        ${k.sub   ? `<div class="kpi-sub">${k.sub}</div>` : ''}
        ${k.badge || ''}
      </div>`).join('');
  }

  /* Bar chart — top reasons */
  if (chartId) {
    const top = reasons.filter(r => r.lensesBroken > 0).slice(0, 8);
    if (top.length) {
      ChartFactory.bar(chartId, top.map(r => r.reason), [{
        data: top.map(r => r.lensesBroken),
        backgroundColor: top.map((_, i) => i === 0 ? color + 'dd' : color + '55'),
        borderRadius: 4, borderSkipped: false,
      }], {
        extra: {
          plugins: { tooltip: { callbacks: { label: ctx => ` ${ctx.raw} lenses broken` } } },
          scales: { x: { ticks: { maxRotation: 35, font: { size: 10 } } } },
        },
      });
    }
  }

  /* Reasons table */
  const reasonsEl = document.querySelector(`#tab-${tabId} [id$="ReasonsTable"]`);
  if (reasonsEl) {
    const active = reasons.filter(r => r.lensesBroken > 0 || r.framesBroken > 0);
    if (active.length) {
      const maxB = active[0].lensesBroken || active[0].framesBroken || 1;
      reasonsEl.innerHTML = `
        <table class="data-table">
          <thead><tr><th>#</th><th>Reason</th><th style="text-align:right">Broken</th><th style="text-align:right">%</th></tr></thead>
          <tbody>
            ${active.map((r, i) => `
              <tr>
                <td class="rank ${i===0?'r1':i===1?'r2':i===2?'r3':''}">${r.rank}</td>
                <td>
                  <div class="reason-name">${r.reason}</div>
                  <div class="mini-bar">
                    <div class="mini-bar-fill" style="width:${((r.lensesBroken||r.framesBroken)/maxB)*100}%;background:${color}"></div>
                  </div>
                </td>
                <td class="count" style="color:${color};text-align:right">${r.lensesBroken || r.framesBroken}</td>
                <td class="pct-cell">${((r.lensesBroken||r.framesBroken||0) / (State.meta.lensCount||1) * 100).toFixed(2)}%</td>
              </tr>`).join('')}
          </tbody>
        </table>`;
    } else {
      reasonsEl.innerHTML = U.emptyState('No breakage data for today');
    }
  }

  /* Operators table */
  const opsEl = document.querySelector(`#tab-${tabId} [id$="Operators"]`);
  if (opsEl) {
    const isFinish = tabId === 'finish';
    if (operators.length) {
      opsEl.innerHTML = `
        <table class="op-table">
          <thead><tr>
            <th>Operator</th><th>Reason</th>
            ${isFinish ? '<th style="text-align:right">Lenses</th><th style="text-align:right">Frames</th>' : '<th style="text-align:right">Total</th>'}
          </tr></thead>
          <tbody>
            ${operators.map(o => `
              <tr>
                <td class="op-name">${o.operator || '—'}</td>
                <td><span class="op-reason">${o.reason || '—'}</span></td>
                ${isFinish
                  ? `<td class="op-total" style="color:${color};text-align:right">${o.lensTotal||0}</td>
                     <td class="op-total" style="color:var(--amber);text-align:right">${o.frameTotal||0}</td>`
                  : `<td class="op-total" style="color:${color};text-align:right">${o.total||0}</td>`}
              </tr>`).join('')}
          </tbody>
        </table>`;
    } else {
      opsEl.innerHTML = U.emptyState('No operator data for today');
    }
  }
}

/* ── MAILROOM ───────────────────────────────────────────── */
function renderMailroom(meta) {
  const orders = meta.orderCount || 0;

  document.getElementById('mailCount').textContent = U.fmt(orders);
  document.getElementById('mailSub').textContent   = `Report Date: ${meta.reportDate || '—'}  ·  Last Refresh: ${meta.lastRefresh || '—'}`;
}

/* ── HISTORY ────────────────────────────────────────────── */
function renderHistory(rows) {
  document.getElementById('historyCount').textContent = `${rows.length} record${rows.length !== 1 ? 's' : ''}`;
  if (!rows.length) {
    document.getElementById('historyTableWrap').innerHTML =
      U.emptyState('No snapshots yet. Click "Snapshot" to save today\'s data.');
    return;
  }

  const cols = ['Date','Lab Lenses','Lab Lens %','Lab Frames','Lab Frame %',
                 'AR Lenses','AR %','Finish Lenses','Fin %',
                 'Surface Lenses','Srf %','LMS Lenses','Mailroom Orders','Snapped By'];

  document.getElementById('historyTableWrap').innerHTML = `
    <table class="history-table">
      <thead><tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr></thead>
      <tbody>
        ${rows.slice().reverse().map(r => `
          <tr>${cols.map(c => {
            const val = r[c] !== undefined ? r[c] : '—';
            return `<td>${val}</td>`;
          }).join('')}</tr>`).join('')}
      </tbody>
    </table>`;
}


/* ── DAILY SUMMARY ──────────────────────────────────────── */
function renderDaily(summary, depts, meta) {
  const lt       = summary.labTotal;
  const lensCount  = meta.lensCount  || 1;
  const orderCount = meta.orderCount || 1;
  const labPct   = lt.labLensPct * 100;
  const framePct = lt.labFramePct * 100;

  // Date badge
  const today = new Date();
  const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  document.getElementById('dailyDateBadge').textContent =
    dayNames[today.getDay()] + ' · ' + (meta.reportDate || today.toLocaleDateString());

  // ── Facility KPIs ──
  const deptMap = {};
  summary.departments.forEach(d => { deptMap[d.department] = d; });

  document.getElementById('dailyFacilityKpis').innerHTML = [
    { label: 'Lab Lenses Broken', value: lt.labLensesBroken, accent: 'var(--cyan)', sub: `of ${lt.lensCount.toLocaleString()} lenses` },
    { label: 'Lab Lens %',  value: labPct.toFixed(2) + '%',  accent: U.statusColor(labPct, 5.00),  sub: 'Goal ≤5.00%', badge: U.statusBadge(labPct, 5.00) },
    { label: 'Frames Broken', value: lt.framesBroken,        accent: 'var(--amber)', sub: `of ${orderCount.toLocaleString()} orders` },
    { label: 'Frame Brk %', value: framePct.toFixed(2) + '%', accent: U.statusColor(framePct, 1.00), sub: 'Goal ≤1.00%', badge: U.statusBadge(framePct, 1.00) },
  ].map(k => `
    <div class="kpi-card" style="--accent:${k.accent}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="color:${k.accent}">${typeof k.value === 'number' && !k.value.toString().includes('%') ? k.value.toLocaleString() : k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
      ${k.badge || ''}
    </div>`).join('');

  // ── Goal Status ──
  const arPct  = deptMap.AR      ? (deptMap.AR.lensesBroken  / lensCount * 100) : 0;
  const finPct = deptMap.Finish  ? (deptMap.Finish.lensesBroken  / lensCount * 100) : 0;
  const srfPct = deptMap.Surface ? (deptMap.Surface.lensesBroken / lensCount * 100) : 0;

  const goalCards = [
    { label: 'Lab Total',      pct: labPct,   target: 5.00 },
    { label: 'AR',             pct: arPct,    target: 0.85 },
    { label: 'Finish',         pct: finPct,   target: 1.70 },
    { label: 'Surface',        pct: srfPct,   target: 2.80 },
    { label: 'Frame Breakage', pct: framePct, target: 1.00 },
  ];

  document.getElementById('dailyGoalStatus').innerHTML = goalCards.map(g => {
    const ok = g.pct <= g.target;
    return `
      <div class="goal-status-card ${ok ? 'on-goal' : 'over-goal'}">
        <div class="gs-icon">${ok ? '✓' : '✗'}</div>
        <div class="gs-label">${g.label}</div>
        <div class="gs-value">${g.pct.toFixed(2)}%</div>
        <div class="gs-target">Goal ≤${g.target}%</div>
      </div>`;
  }).join('');

  // ── Dept Breakdown Table ──
  const deptConfigs = [
    { key: 'AR',        color: '#a78bfa', goal: 0.85, data: depts.ar },
    { key: 'Finish',    color: '#34d399', goal: 1.70, data: depts.finish },
    { key: 'Surface',   color: '#fb923c', goal: 2.80, data: depts.surface },
    { key: 'LMS',       color: '#60a5fa', goal: null, data: depts.lms },
    { key: 'Inventory', color: '#f472b6', goal: null, data: depts.inventory },
  ];

  document.getElementById('dailyDeptBreakdown').innerHTML = `
    <div class="panel">
      <div class="panel-body" style="padding:0">
        <table class="dept-breakdown-table">
          <thead><tr>
            <th>Dept</th>
            <th style="text-align:right">Broken</th>
            <th style="text-align:right">%</th>
            <th style="text-align:center">Status</th>
            <th>Top Reasons</th>
            <th>Top Operators</th>
          </tr></thead>
          <tbody>
            ${deptConfigs.map(dc => {
              const dept = summary.departments.find(d => d.department === dc.key);
              if (!dept) return '';
              const pct = dept.lensesBroken / lensCount * 100;
              const ok  = dc.goal ? pct <= dc.goal : true;
              const topReasons  = (dc.data?.reasons  || []).filter(r => r.lensesBroken > 0).slice(0, 3);
              const topOps      = (dc.data?.operators || [])
                .reduce((acc, o) => {
                  const key = o.operator;
                  if (!key || key === 'NONE') return acc;
                  acc[key] = (acc[key] || 0) + (o.total || o.lensTotal || 0);
                  return acc;
                }, {});
              const sortedOps = Object.entries(topOps).sort((a,b) => b[1]-a[1]).slice(0, 3);

              return `
                <tr>
                  <td><div class="dbt-dept" style="color:${dc.color}">${dc.key}</div></td>
                  <td><div class="dbt-count" style="color:${dc.color}">${dept.lensesBroken}</div></td>
                  <td><div class="dbt-pct" style="color:${U.statusColor(pct, dc.goal||999)}">${pct.toFixed(2)}%</div></td>
                  <td class="dbt-status">${ok ? '<span style="color:var(--green);font-size:16px">✓</span>' : '<span style="color:var(--red);font-size:16px">✗</span>'}</td>
                  <td>
                    <div class="dbt-reasons">
                      ${topReasons.length ? topReasons.map((r,i) => `
                        <div class="dbt-reason-item">
                          <span class="dbt-reason-rank">${i+1}</span>
                          <span class="dbt-reason-name">${r.reason}</span>
                          <span class="dbt-reason-count" style="color:${dc.color}">${r.lensesBroken}</span>
                        </div>`).join('') : '<span style="color:var(--muted)">—</span>'}
                    </div>
                  </td>
                  <td>
                    <div class="dbt-operators">
                      ${sortedOps.length ? sortedOps.map(([name, count]) => `
                        <div class="dbt-op-item">
                          <span class="dbt-op-name">${name}</span>
                          <span class="dbt-op-count" style="color:${dc.color}">${count}</span>
                        </div>`).join('') : '<span style="color:var(--muted)">—</span>'}
                    </div>
                  </td>
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>`;

  // ── Improvement Opportunities ──
  const improvements = [];

  goalCards.forEach(g => {
    if (g.pct > g.target) {
      const over = (g.pct - g.target).toFixed(2);
      const deptKey = g.label === 'Frame Breakage' ? 'Frames' : g.label === 'Lab Total' ? 'Lab' : g.label;
      improvements.push({
        type: 'critical',
        dept: deptKey,
        title: `Over goal by ${over}%  (${g.pct.toFixed(2)}% vs ≤${g.target}%)`,
        desc: `Reduce breakage by ${Math.ceil((g.pct - g.target) / g.pct * 100)}% to reach target`,
      });
    }
  });

  // Top offending reasons across depts
  const allReasons = [];
  deptConfigs.forEach(dc => {
    (dc.data?.reasons || []).filter(r => r.lensesBroken > 0).forEach(r => {
      allReasons.push({ reason: r.reason, dept: dc.key, count: r.lensesBroken, color: dc.color });
    });
  });
  allReasons.sort((a,b) => b.count - a.count);
  allReasons.slice(0, 3).forEach(r => {
    improvements.push({
      type: 'warning',
      dept: r.dept,
      title: `${r.reason} — ${r.count} lenses`,
      desc: `Review training and process controls for this failure mode`,
    });
  });

  // All goals met?
  const allMet = goalCards.every(g => g.pct <= g.target);
  if (allMet) {
    improvements.push({ type: 'ok', dept: 'Lab', title: 'All goals met today', desc: 'Facility operating within all breakage targets' });
  }

  // Build compact table rows
  const impRows = improvements.map(imp => {
    const deptMatch = imp.title.match(/^(Lab Total|AR|Finish|Surface|Frame Breakage|High volume:.+?\((\w+)\))/);
    const dept = imp.dept || (imp.title.includes('Surface') ? 'Surface'
               : imp.title.includes('Finish') ? 'Finish'
               : imp.title.includes('AR') ? 'AR'
               : imp.title.includes('Frame') ? 'Frames'
               : imp.title.includes('Lab') ? 'Lab' : '—');
    const deptColor = { AR:'#a78bfa', Finish:'#34d399', Surface:'#fb923c', Frames:'#f5a623', Lab:'#38bdf8' }[dept] || 'var(--muted)';
    const priority = imp.type === 'critical' ? 'high' : imp.type === 'warning' ? 'medium' : 'low';
    const priorityLabel = priority === 'high' ? 'HIGH' : priority === 'medium' ? 'MEDIUM' : 'OK';
    return { dept, deptColor, issue: imp.title, action: imp.desc, priority, priorityLabel };
  });

  document.getElementById('dailyImprovements').innerHTML = `
    <div class="panel">
      <div class="panel-body" style="padding:0">
        <table class="imp-table">
          <thead><tr>
            <th>Dept</th>
            <th>Issue</th>
            <th>Recommended Action</th>
            <th>Priority</th>
          </tr></thead>
          <tbody>
            ${impRows.map(r => `
              <tr>
                <td><span class="imp-dept" style="color:${r.deptColor}">${r.dept}</span></td>
                <td><span class="imp-issue">${r.issue}</span></td>
                <td><span class="imp-action">${r.action}</span></td>
                <td><span class="imp-priority ${r.priority}">${r.priorityLabel}</span></td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}


/* ── WEEKLY SUMMARY (Sun–Sat) ───────────────────────────── */
function renderWeekly(history, meta) {
  const lensCount  = meta.lensCount  || 1;
  const orderCount = meta.orderCount || 1;

  // Get current Sun–Sat week boundaries
  const now     = new Date();
  const dayOfWk = now.getDay(); // 0=Sun
  const weekStart = new Date(now); weekStart.setDate(now.getDate() - dayOfWk); weekStart.setHours(0,0,0,0);
  const weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6); weekEnd.setHours(23,59,59,999);

  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const fmt = d => `${d.getMonth()+1}/${d.getDate()}`;
  document.getElementById('weeklyRangeBadge').textContent =
    `Week: ${fmt(weekStart)} – ${fmt(weekEnd)}`;

  // Filter history to this Sun–Sat week
  const weekSnaps = history.filter(r => {
    if (!r['Date']) return false;
    const d = new Date(r['Date']);
    return d >= weekStart && d <= weekEnd;
  });

  // Build day slots Sun–Sat
  const daySlots = Array.from({length: 7}, (_, i) => {
    const d = new Date(weekStart);
    d.setDate(weekStart.getDate() + i);
    const snap = weekSnaps.find(r => {
      const rd = new Date(r['Date']);
      return rd.getDay() === i;
    });
    return { day: dayNames[i], date: fmt(d), snap, isFuture: d > now };
  });

  const hasData = weekSnaps.length > 0;

  // ── Week KPI Overview ──
  if (!hasData) {
    document.getElementById('weeklyFacilityKpis').innerHTML = `
      <div style="grid-column:1/-1;padding:20px;text-align:center;font-family:var(--font-mono);font-size:12px;color:var(--muted)">
        No snapshots taken this week yet. Take daily snapshots to populate weekly data.
      </div>`;
  } else {
    const labPcts   = weekSnaps.map(r => U.parseHistoryPct(r['Lab Lens %'])   || 0);
    const framePcts = weekSnaps.map(r => U.parseHistoryPct(r['Lab Frame %'])  || 0);
    const arPcts    = weekSnaps.map(r => U.parseHistoryPct(r['AR %'])         || 0);
    const finPcts   = weekSnaps.map(r => U.parseHistoryPct(r['Fin %'])        || 0);
    const srfPcts   = weekSnaps.map(r => U.parseHistoryPct(r['Srf %'])        || 0);
    const avg = arr => arr.length ? arr.reduce((a,b) => a+b, 0) / arr.length : 0;
    const avgLab = avg(labPcts), avgFin = avg(finPcts), avgSrf = avg(srfPcts), avgAR = avg(arPcts);

    document.getElementById('weeklyFacilityKpis').innerHTML = [
      { label: 'Snapshots This Week', value: weekSnaps.length,       accent: 'var(--muted)',    sub: `${weekSnaps.length} of 7 days` },
      { label: 'Avg Lab Lens %',      value: avgLab.toFixed(2) + '%', accent: U.statusColor(avgLab, 5.00),  sub: 'Goal ≤5.00%', badge: U.statusBadge(avgLab, 5.00) },
      { label: 'Avg AR %',            value: avgAR.toFixed(2)  + '%', accent: U.statusColor(avgAR,  0.85),  sub: 'Goal ≤0.85%', badge: U.statusBadge(avgAR, 0.85)  },
      { label: 'Avg Finish %',        value: avgFin.toFixed(2) + '%', accent: U.statusColor(avgFin, 1.70),  sub: 'Goal ≤1.70%', badge: U.statusBadge(avgFin, 1.70) },
      { label: 'Avg Surface %',       value: avgSrf.toFixed(2) + '%', accent: U.statusColor(avgSrf, 2.80),  sub: 'Goal ≤2.80%', badge: U.statusBadge(avgSrf, 2.80) },
    ].map(k => `
      <div class="kpi-card" style="--accent:${k.accent}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
        <div class="kpi-sub">${k.sub}</div>
        ${k.badge || ''}
      </div>`).join('');

    // ── Weekly Trend Chart ──
    const chartLabels = daySlots.map(d => d.day + ' ' + d.date);
    const getVal = (slot, key) => slot.snap ? (U.parseHistoryPct(slot.snap[key]) || 0) : null;

    destroyChart('weeklyTrendChart');
    const ctx = document.getElementById('weeklyTrendChart');
    if (ctx) {
      Charts['weeklyTrendChart'] = new Chart(ctx, {
        type: 'line',
        data: {
          labels: chartLabels,
          datasets: [
            { label: 'Lab Total', data: daySlots.map(d => getVal(d,'Lab Lens %')),  borderColor: '#38bdf8', backgroundColor: '#38bdf820', pointBackgroundColor: '#38bdf8', pointBorderColor: '#080b0f', pointBorderWidth: 2, pointRadius: 5, borderWidth: 2, tension: 0.3, spanGaps: true },
            { label: 'AR',        data: daySlots.map(d => getVal(d,'AR %')),         borderColor: '#a78bfa', backgroundColor: 'transparent', pointBackgroundColor: '#a78bfa', pointBorderColor: '#080b0f', pointBorderWidth: 2, pointRadius: 4, borderWidth: 2, tension: 0.3, spanGaps: true },
            { label: 'Finish',    data: daySlots.map(d => getVal(d,'Fin %')),        borderColor: '#34d399', backgroundColor: 'transparent', pointBackgroundColor: '#34d399', pointBorderColor: '#080b0f', pointBorderWidth: 2, pointRadius: 4, borderWidth: 2, tension: 0.3, spanGaps: true },
            { label: 'Surface',   data: daySlots.map(d => getVal(d,'Srf %')),        borderColor: '#fb923c', backgroundColor: 'transparent', pointBackgroundColor: '#fb923c', pointBorderColor: '#080b0f', pointBorderWidth: 2, pointRadius: 4, borderWidth: 2, tension: 0.3, spanGaps: true },
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { display: true, position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#9aafc4' } },
            tooltip: { callbacks: { label: ctx => ctx.raw !== null ? ` ${ctx.dataset.label}: ${ctx.raw.toFixed(2)}%` : ' No data' } },
          },
          scales: {
            x: { grid: { display: false }, ticks: { color: '#9aafc4', font: { size: 10 } } },
            y: { beginAtZero: true, ticks: { callback: v => v.toFixed(2) + '%', color: '#9aafc4', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,0.05)' } },
          },
        },
      });
    }

    // ── Weekly Dept Breakdown (per day) ──
    const deptCols = [
      { key: 'AR %',   label: 'AR',      color: '#a78bfa', goal: 0.85 },
      { key: 'Fin %',  label: 'Finish',  color: '#34d399', goal: 1.70 },
      { key: 'Srf %',  label: 'Surface', color: '#fb923c', goal: 2.80 },
    ];

    document.getElementById('weeklyDeptBreakdown').innerHTML = `
      <div class="panel">
        <div class="panel-body" style="padding:0">
          <table class="dept-breakdown-table">
            <thead><tr>
              <th>Day</th>
              <th style="text-align:right">Lab %</th>
              ${deptCols.map(d => `<th style="text-align:right;color:${d.color}">${d.label}</th>`).join('')}
              <th>AR Top Reason</th>
              <th>Fin Top Reason</th>
              <th>Srf Top Reason</th>
            </tr></thead>
            <tbody>
              ${daySlots.map(slot => {
                const s = slot.snap;
                if (slot.isFuture && !s) return `
                  <tr style="opacity:0.3">
                    <td><div class="dbt-dept" style="color:var(--muted)">${slot.day} ${slot.date}</div></td>
                    <td colspan="6" style="color:var(--muted);font-family:var(--font-mono);font-size:11px">—</td>
                  </tr>`;
                if (!s) return `
                  <tr style="opacity:0.5">
                    <td><div class="dbt-dept" style="color:var(--muted)">${slot.day} ${slot.date}</div></td>
                    <td colspan="6" style="color:var(--muted);font-family:var(--font-mono);font-size:11px">No snapshot</td>
                  </tr>`;
                const labP = U.parseHistoryPct(s['Lab Lens %']) || 0;
                const labOk = labP <= 5.00;
                return `
                  <tr>
                    <td><div class="dbt-dept" style="color:var(--text)">${slot.day} ${slot.date}</div></td>
                    <td class="dbt-pct" style="color:${labOk ? 'var(--green)' : 'var(--red)'}">${labP.toFixed(2)}%</td>
                    ${deptCols.map(d => {
                      const v = U.parseHistoryPct(s[d.key]) || 0;
                      return `<td class="dbt-pct" style="color:${v <= d.goal ? 'var(--green)' : 'var(--red)'}">${v.toFixed(2)}%</td>`;
                    }).join('')}
                    <td class="dbt-reasons"><span style="color:var(--ar)">${s['AR Top Reason'] || '—'}</span></td>
                    <td class="dbt-reasons"><span style="color:var(--fin)">${s['Fin Top Reason'] || '—'}</span></td>
                    <td class="dbt-reasons"><span style="color:var(--srf)">${s['Srf Top Reason'] || '—'}</span></td>
                  </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>`;

    // ── Top Recurring Reasons This Week ──
    const reasonCount = {};
    weekSnaps.forEach(r => {
      ['AR Top Reason','Fin Top Reason','Srf Top Reason'].forEach(col => {
        const val = r[col];
        if (val && val.trim()) {
          if (!reasonCount[val]) reasonCount[val] = { count: 0, dept: col.split(' ')[0] };
          reasonCount[val].count++;
        }
      });
    });
    const sortedReasons = Object.entries(reasonCount).sort((a,b) => b[1].count - a[1].count).slice(0, 8);
    const maxRC = sortedReasons[0]?.[1].count || 1;
    const deptColorMap = { AR: '#a78bfa', Fin: '#34d399', Srf: '#fb923c' };

    document.getElementById('weeklyTopReasons').innerHTML = sortedReasons.length
      ? sortedReasons.map(([reason, {count, dept}]) => {
          const color = deptColorMap[dept] || 'var(--cyan)';
          return `
            <div class="weekly-reason-row">
              <div class="wr-name">${reason}</div>
              <div class="wr-dept" style="color:${color}">${dept}</div>
              <div class="wr-bar"><div class="wr-fill" style="width:${(count/maxRC)*100}%;background:${color}"></div></div>
              <div class="wr-count" style="color:${color}">${count}d</div>
            </div>`;
        }).join('')
      : '<div style="color:var(--muted);font-family:var(--font-mono);font-size:12px;text-align:center;padding:20px">Take snapshots daily to see recurring reasons</div>';

    // ── Weekly Improvements ──
    const weekImprovements = [];
    const avgLabPct = avg(labPcts);
    const overGoalDays = daySlots.filter(d => d.snap && (U.parseHistoryPct(d.snap['Lab Lens %'])||0) > 5.00);
    const overARDays   = daySlots.filter(d => d.snap && (U.parseHistoryPct(d.snap['AR %'])||0)  > 0.85);
    const overFinDays  = daySlots.filter(d => d.snap && (U.parseHistoryPct(d.snap['Fin %'])||0) > 1.70);
    const overSrfDays  = daySlots.filter(d => d.snap && (U.parseHistoryPct(d.snap['Srf %'])||0) > 2.80);

    if (overGoalDays.length > 0)
      weekImprovements.push({ type: 'critical', icon: '⚠', title: `Lab over 5% goal on ${overGoalDays.length} day(s)`, desc: overGoalDays.map(d => d.day).join(', ') + ' — review process controls' });
    if (overSrfDays.length > 0)
      weekImprovements.push({ type: 'critical', icon: '⚠', title: `Surface over 2.80% goal on ${overSrfDays.length} day(s)`, desc: overSrfDays.map(d => d.day).join(', ') + ' — S-Power and S-HC Pit most common drivers' });
    if (overARDays.length > 0)
      weekImprovements.push({ type: 'warning', icon: '↑', title: `AR over 0.85% goal on ${overARDays.length} day(s)`, desc: overARDays.map(d => d.day).join(', ') + ' — review A-Off color and A-Scratch frequency' });
    if (overFinDays.length > 0)
      weekImprovements.push({ type: 'warning', icon: '↑', title: `Finish over 1.70% goal on ${overFinDays.length} day(s)`, desc: overFinDays.map(d => d.day).join(', ') + ' — check F-Slip equipment calibration' });
    if (sortedReasons[0])
      weekImprovements.push({ type: 'warning', icon: '↑', title: `"${sortedReasons[0][0]}" appeared ${sortedReasons[0][1].count} day(s) as top reason`, desc: 'Most persistent breakage reason this week — prioritize training focus' });
    if (weekImprovements.length === 0)
      weekImprovements.push({ type: 'ok', icon: '✓', title: 'All goals met across all snapshot days this week', desc: 'Facility is consistently within all breakage targets' });

    const weekImpRows = weekImprovements.map(imp => {
      const dept = imp.title.includes('Surface') ? 'Surface'
                 : imp.title.includes('Finish')  ? 'Finish'
                 : imp.title.includes('AR')       ? 'AR'
                 : imp.title.includes('Lab')      ? 'Lab' : 'Lab';
      const deptColor = { AR:'#a78bfa', Finish:'#34d399', Surface:'#fb923c', Lab:'#38bdf8' }[dept] || 'var(--muted)';
      const priority = imp.type === 'critical' ? 'high' : imp.type === 'warning' ? 'medium' : 'low';
      return { dept, deptColor, issue: imp.title, action: imp.desc, priority, priorityLabel: priority === 'high' ? 'HIGH' : priority === 'medium' ? 'MEDIUM' : 'OK' };
    });

    document.getElementById('weeklyImprovements').innerHTML = `
      <table class="imp-table">
        <thead><tr>
          <th>Dept</th><th>Issue</th><th>Action</th><th>Priority</th>
        </tr></thead>
        <tbody>
          ${weekImpRows.map(r => `
            <tr>
              <td><span class="imp-dept" style="color:${r.deptColor}">${r.dept}</span></td>
              <td><span class="imp-issue">${r.issue}</span></td>
              <td><span class="imp-action">${r.action}</span></td>
              <td><span class="imp-priority ${r.priority}">${r.priorityLabel}</span></td>
            </tr>`).join('')}
        </tbody>
      </table>`;
  }
}

/* ═══════════════════════════════════════════════════════════
   MAIN APP CONTROLLER
═══════════════════════════════════════════════════════════ */
const App = {

  async loadAll() {
    const btn = document.getElementById('refreshBtn');
    btn.classList.add('loading');
    try {
      const [allRes, histRes] = await Promise.all([API.all(), API.history()]);

      State.meta    = allRes.meta;
      State.summary = allRes.data.summary;
      State.depts   = allRes.data;
      State.history = histRes.data;

      renderMeta(allRes.meta);
      renderSummary(allRes.data.summary, allRes.meta, histRes.data);
      renderDeptTab('ar',        allRes.data.ar,        'AR',        '#a78bfa', 'arChart');
      renderDeptTab('finish',    allRes.data.finish,    'Finish',    '#34d399', 'finChart');
      renderDeptTab('surface',   allRes.data.surface,   'Surface',   '#fb923c', 'srfChart');
      renderDeptTab('lms',       allRes.data.lms,       'LMS',       '#60a5fa', 'lmsChart');
      renderDeptTab('inventory', allRes.data.inventory, 'Inventory', '#f472b6', null);

      renderMailroom(allRes.meta);
      renderHistory(histRes.data);
      renderDaily(allRes.data.summary, allRes.data, allRes.meta);
      renderWeekly(histRes.data, allRes.meta);

      document.getElementById('loadingOverlay').classList.add('hidden');
      App.showToast('Data refreshed successfully', 'success');

    } catch (err) {
      console.error(err);
      App.showToast('Error loading data: ' + err.message, 'error');
      document.getElementById('loadingOverlay').classList.add('hidden');
    }
    btn.classList.remove('loading');
  },

  switchTab(id) {
    const tab = document.querySelector(`.nav-tab[data-tab="${id}"]`);
    if (tab) {
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
      document.getElementById('tab-' + id)?.classList.add('active');
      tab.classList.add('active');
    }
  },

  async triggerSnapshot() {
    const name = prompt('Snapshot taken by (your name):', 'Manager');
    if (!name || !name.trim()) return;
    try {
      await API.snapshot(name.trim());
      App.showToast(`Snapshot saved by ${name.trim()}`, 'success');
      // Refresh history
      const histRes = await API.history();
      State.history = histRes.data;
      renderHistory(histRes.data);
      renderWeekly(histRes.data, State.meta);
    } catch (err) {
      App.showToast('Snapshot failed: ' + err.message, 'error');
    }
  },

  showToast(msg, type = '') {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.className   = `toast ${type} show`;
    setTimeout(() => t.classList.remove('show'), 3500);
  },
};

/* ── NAV TAB CLICK WIRING ───────────────────────────────── */
document.querySelectorAll('.nav-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    const id = btn.dataset.tab;
    document.getElementById('tab-' + id)?.classList.add('active');
    btn.classList.add('active');
  });
});

/* ── BOOT ───────────────────────────────────────────────── */
App.loadAll();