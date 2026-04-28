// FacilityBreakage.js — v20260427_1947
/* ═══════════════════════════════════════════════════════════
   ZENNI FACILITY BREAKAGE — app.js
   Requires: Chart.js 4.x loaded before this file
═══════════════════════════════════════════════════════════ */

'use strict';

/* ── CONFIG ─────────────────────────────────────────────── */
const API_URL = 'https://script.google.com/macros/s/AKfycbylW0fs7zWLncknhz7peJcCm9eyWCXTGCLtk-xtMIzjarot5PcCpmCP4Gy85WqT3f17/exec';

const GOALS = {
  'Lab Total':      5.00,
  'AR':             0.50,
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
  Breakage:  '#ef4444',
};

const CHART_DEFAULTS = {
  gridColor:   'rgba(255,255,255,0.05)',
  textColor:   '#ffffff',
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
  async get(action, params = {}) {
    const qs   = Object.entries(params).map(([k,v]) => `&${k}=${encodeURIComponent(v)}`).join('');
    const res  = await fetch(`${API_URL}?action=${action}${qs}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error || `API error: ${action}`);
    return json;
  },

  // Live: action=all (no date)
  // History: action=all&date=MM/DD/YYYY
  async all(date)    { return date ? this.get('all', { date }) : this.get('all'); },
  async status()     { return this.get('status'); },
  async historyDates(){ return this.get('history'); },
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
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, color: '#ffffff' },
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
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, color: '#ffffff', font: { size: 11 } },
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
  const rd = document.getElementById('reportDate');
  const lr = document.getElementById('lastRefresh');
  if (rd) rd.textContent = meta.reportDate  || '—';
  if (lr) lr.textContent = meta.lastRefresh || '—';
  const sub = document.getElementById('headerSub');
  if (sub) {
    const isHistory = App._isHistoryMode;
    sub.textContent = isHistory
      ? `History: ${meta.reportDate || '—'}`
      : meta.lensCount > 0
        ? `Report: ${meta.reportDate || '—'}  ·  ${U.fmt(meta.lensCount)} lenses`
        : 'Waiting for RawData upload…';
  }
}

/* ── SUMMARY TAB ────────────────────────────────────────── */
function renderSummary(summary, meta, history) {
  // Guard — skip if Overview tab elements not in DOM
  if (!document.getElementById('labKpis')) return;
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
    { label: 'Order Count',  value: U.fmt(lt.orderCount), sub: 'Total orders today', accent: 'var(--muted)', delta: '' },
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
    { label: 'AR',             pct: (deptMap.AR?.lensBrkPct    || 0) * 100, target: 0.50, color: 'var(--ar)'   },
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
        <div class="dept-top">↑ ${(!d.topReason || d.topReason.startsWith('#')) ? '—' : d.topReason}</div>
      </div>`;
  }).join('');

  /* Top reasons panels */
  const deptKeys = ['AR', 'Finish', 'Surface', 'LMS', 'Inventory', 'Breakage'];
  const deptClrs = ['var(--ar)', 'var(--fin)', 'var(--srf)', 'var(--lms)', 'var(--inv)', 'var(--brk)'];
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
    { label: 'AR %',           value: U.pctRaw(arPct),    accent: U.statusColor(arPct,   0.50), delta: U.deltaHtml(arPct,   prevARPct),    badge: U.statusBadge(arPct,   0.50) },
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
            labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, color: '#ffffff' },
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
    { label: 'AR',             pct: arPct,   target: 0.50 },
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
function renderWeekly(data) {
  if (!data || !data.days) {
    document.getElementById('weeklyFacilityKpis').innerHTML = U.emptyState('No weekly data available');
    return;
  }

  const { days, weekStart, weekEnd, availableWeeks } = data;

  document.getElementById('weeklyRangeBadge').textContent =
    `${weekStart || '—'} – ${weekEnd || '—'}`;

  const daysWithData = days.filter(d => d.hasData);

  // ── Week KPI averages ──
  const avg = arr => arr.length ? arr.reduce((a,b) => a+b,0)/arr.length : 0;
  const avgLab = avg(daysWithData.map(d => d.labLensPct));
  const avgAR  = avg(daysWithData.map(d => d.arPct));
  const avgFin = avg(daysWithData.map(d => d.finPct));
  const avgSrf = avg(daysWithData.map(d => d.srfPct));
  const totalBroken = daysWithData.reduce((s,d) => s + (d.labLenses||0), 0);

  document.getElementById('weeklyFacilityKpis').innerHTML = daysWithData.length === 0
    ? `<div style="grid-column:1/-1">${U.emptyState('No snapshots found for this week')}</div>`
    : [
        { label: 'Days with Data',   value: daysWithData.length + ' of 7', accent: 'var(--muted)' },
        { label: 'Total Lenses Broken', value: U.fmt(totalBroken),        accent: 'var(--cyan)' },
        { label: 'Avg Lab %',    value: avgLab.toFixed(2)+'%', accent: U.statusColor(avgLab, 5.00), badge: U.statusBadge(avgLab, 5.00) },
        { label: 'Avg AR %',     value: avgAR.toFixed(2) +'%', accent: U.statusColor(avgAR,  0.50), badge: U.statusBadge(avgAR,  0.50) },
        { label: 'Avg Finish %', value: avgFin.toFixed(2)+'%', accent: U.statusColor(avgFin, 1.70), badge: U.statusBadge(avgFin, 1.70) },
        { label: 'Avg Surface %',value: avgSrf.toFixed(2)+'%', accent: U.statusColor(avgSrf, 2.80), badge: U.statusBadge(avgSrf, 2.80) },
      ].map(k => `
        <div class="kpi-card" style="--accent:${k.accent}">
          <div class="kpi-label">${k.label}</div>
          <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
          ${k.badge || ''}
        </div>`).join('');

  // ── Weekly Trend Chart ──
  destroyChart('weeklyTrendChart');
  const ctx = document.getElementById('weeklyTrendChart');
  if (ctx && daysWithData.length > 0) {
    const labels = days.map(d => d.day + ' ' + d.date.slice(0,5));
    const getVal = (day, key) => day.hasData ? day[key] : null;
    Charts['weeklyTrendChart'] = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [
          { label:'Lab Total', data: days.map(d => getVal(d,'labLensPct')),  borderColor:'#38bdf8', backgroundColor:'#38bdf820', pointBackgroundColor:'#38bdf8', pointBorderColor:'#080b0f', pointBorderWidth:2, pointRadius:5, borderWidth:2, tension:0.3, spanGaps:true },
          { label:'AR',        data: days.map(d => getVal(d,'arPct')),       borderColor:'#a78bfa', backgroundColor:'transparent', pointBackgroundColor:'#a78bfa', pointBorderColor:'#080b0f', pointBorderWidth:2, pointRadius:4, borderWidth:2, tension:0.3, spanGaps:true },
          { label:'Finish',    data: days.map(d => getVal(d,'finPct')),      borderColor:'#34d399', backgroundColor:'transparent', pointBackgroundColor:'#34d399', pointBorderColor:'#080b0f', pointBorderWidth:2, pointRadius:4, borderWidth:2, tension:0.3, spanGaps:true },
          { label:'Surface',   data: days.map(d => getVal(d,'srfPct')),      borderColor:'#fb923c', backgroundColor:'transparent', pointBackgroundColor:'#fb923c', pointBorderColor:'#080b0f', pointBorderWidth:2, pointRadius:4, borderWidth:2, tension:0.3, spanGaps:true },
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        interaction:{ mode:'index', intersect:false },
        plugins:{
          legend:{ display:true, position:'bottom', labels:{ usePointStyle:true, pointStyle:'circle', padding:16, color:'#9aafc4' }},
          tooltip:{ callbacks:{ label: ctx => ctx.raw !== null ? ` ${ctx.dataset.label}: ${ctx.raw.toFixed(2)}%` : ' No data' }},
        },
        scales:{
          x:{ grid:{ display:false }, ticks:{ color:'#9aafc4', font:{ size:10 }}},
          y:{ beginAtZero:true, ticks:{ callback: v => v.toFixed(2)+'%', color:'#9aafc4', font:{ size:10 }}, grid:{ color:'rgba(255,255,255,0.05)' }},
        },
      },
    });
  }

  // ── Day-by-day table ──
  const DEPTS = [
    { key:'arPct',  label:'AR',      color:'#a78bfa', goal:0.50 },
    { key:'finPct', label:'Finish',  color:'#34d399', goal:1.70 },
    { key:'srfPct', label:'Surface', color:'#fb923c', goal:2.80 },
  ];

  const deptBreakdownEl = document.getElementById('weeklyDeptBreakdown');
  if (deptBreakdownEl) {
    deptBreakdownEl.innerHTML = `
      <div class="panel">
        <div class="panel-body" style="padding:0">
          <table class="dept-breakdown-table">
            <thead><tr>
              <th>Day</th>
              <th style="text-align:right">Lab %</th>
              ${DEPTS.map(d => `<th style="text-align:right;color:${d.color}">${d.label}</th>`).join('')}
              <th>Top Reason</th>
            </tr></thead>
            <tbody>
              ${days.map(day => {
                if (day.isFuture && !day.hasData) return `
                  <tr style="opacity:0.3">
                    <td><span style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${day.day} ${day.date.slice(0,5)}</span></td>
                    <td colspan="5" style="color:var(--muted);font-family:var(--font-mono);font-size:11px">—</td>
                  </tr>`;
                if (!day.hasData) return `
                  <tr style="opacity:0.5">
                    <td><span style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${day.day} ${day.date.slice(0,5)}</span></td>
                    <td colspan="5" style="color:var(--muted);font-family:var(--font-mono);font-size:11px">No snapshot</td>
                  </tr>`;
                const labOk = day.labLensPct <= 5.00;
                const topReason = day.topReasons && day.topReasons.Surface && day.topReasons.Surface[0]
                  ? day.topReasons.Surface[0].reason
                  : (day.topReasons && day.topReasons.Finish && day.topReasons.Finish[0] ? day.topReasons.Finish[0].reason : '—');
                return `
                  <tr>
                    <td style="font-family:var(--font-mono);font-size:11px;font-weight:600">${day.day} ${day.date.slice(0,5)}</td>
                    <td style="text-align:right;font-family:var(--font-mono);font-size:12px;color:${labOk?'var(--green)':'var(--red)'}">${day.labLensPct.toFixed(2)}%</td>
                    ${DEPTS.map(d => {
                      const v = day[d.key] || 0;
                      return `<td style="text-align:right;font-family:var(--font-mono);font-size:12px;color:${v<=d.goal?'var(--green)':'var(--red)'}">${v.toFixed(2)}%</td>`;
                    }).join('')}
                    <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${topReason}</td>
                  </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>`;
  }

  // ── Top Recurring Reasons ──
  const reasonCount = {};
  daysWithData.forEach(day => {
    ['AR','Finish','Surface','LMS'].forEach(dept => {
      const top = (day.topReasons[dept] || [])[0];
      if (top && top.reason) {
        if (!reasonCount[top.reason]) reasonCount[top.reason] = { count:0, dept, color: DEPT_COLORS[dept]||'var(--cyan)' };
        reasonCount[top.reason].count++;
      }
    });
  });
  const sortedReasons = Object.entries(reasonCount).sort((a,b) => b[1].count - a[1].count).slice(0,8);
  const maxRC = sortedReasons[0]?.[1].count || 1;

  const weeklyReasonsEl = document.getElementById('weeklyTopReasons');
  if (weeklyReasonsEl) {
    weeklyReasonsEl.innerHTML = sortedReasons.length
      ? sortedReasons.map(([reason, {count, dept, color}]) => `
          <div class="weekly-reason-row">
            <div class="wr-name">${reason}</div>
            <div class="wr-dept" style="color:${color}">${dept}</div>
            <div class="wr-bar"><div class="wr-fill" style="width:${(count/maxRC)*100}%;background:${color}"></div></div>
            <div class="wr-count" style="color:${color}">${count}d</div>
          </div>`)
        .join('')
      : U.emptyState('Take daily snapshots to see recurring reasons');
  }

  // ── Weekly Improvements ──
  const impEl = document.getElementById('weeklyImprovements');
  if (impEl && daysWithData.length > 0) {
    const overLab = daysWithData.filter(d => d.labLensPct > 5.00);
    const overSrf = daysWithData.filter(d => d.srfPct > 2.80);
    const overFin = daysWithData.filter(d => d.finPct  > 1.70);
    const overAR  = daysWithData.filter(d => d.arPct   > 0.50);
    const imps = [];
    if (overLab.length) imps.push({ dept:'Lab',     issue:`Over 5% goal on ${overLab.length} day(s)`,    action:overLab.map(d=>d.day).join(', ')+' — review all depts',       priority:'high' });
    if (overSrf.length) imps.push({ dept:'Surface',  issue:`Over 2.80% goal on ${overSrf.length} day(s)`, action:overSrf.map(d=>d.day).join(', ')+' — review S-Power/S-HC Pit', priority:'high' });
    if (overFin.length) imps.push({ dept:'Finish',   issue:`Over 1.70% goal on ${overFin.length} day(s)`, action:overFin.map(d=>d.day).join(', ')+' — check F-Slip equipment',  priority:'medium' });
    if (overAR.length)  imps.push({ dept:'AR',       issue:`Over 0.50% goal on ${overAR.length} day(s)`,  action:overAR.map(d=>d.day).join(', ')+'  — review A-Off color',      priority:'medium' });
    if (!imps.length)   imps.push({ dept:'Lab',      issue:'All goals met this week',                      action:'Facility within all breakage targets',                         priority:'low' });

    impEl.innerHTML = `
      <table class="imp-table">
        <thead><tr><th>Dept</th><th>Issue</th><th>Action</th><th>Priority</th></tr></thead>
        <tbody>
          ${imps.map(r => {
            const color = {Lab:'#38bdf8',Surface:'#fb923c',Finish:'#34d399',AR:'#a78bfa'}[r.dept]||'var(--muted)';
            const pc    = r.priority==='high'?'high':r.priority==='medium'?'medium':'low';
            const pl    = r.priority==='high'?'HIGH':r.priority==='medium'?'MEDIUM':'OK';
            return `<tr>
              <td><span class="imp-dept" style="color:${color}">${r.dept}</span></td>
              <td><span class="imp-issue">${r.issue}</span></td>
              <td><span class="imp-action">${r.action}</span></td>
              <td><span class="imp-priority ${pc}">${pl}</span></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>`;
  }
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
    const goal = { ar: 0.50, finish: 1.70, surface: 2.80 }[tabId];
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

  /* Breakage by Operator/Reason — toggle between grouped-by-operator and grouped-by-reason */
  const opsEl = document.querySelector(`#tab-${tabId} [id$="Operators"]`);
  if (opsEl) {
    const isFinish  = tabId === 'finish';
    const isSurface = tabId === 'surface';
    const colLabel  = isSurface ? 'Machine' : 'Operator';

    if (operators.length) {
      // operators now have: { operator, total/lensTotal/frameTotal, reason (top), reasons: [{reason, lenses, frames}] }
      const sorted = [...operators].sort((a, b) => {
        const at = a.total ?? ((a.lensTotal||0) + (a.frameTotal||0));
        const bt = b.total ?? ((b.lensTotal||0) + (b.frameTotal||0));
        return bt - at;
      });

      const renderByOperator = () => {
        const uid = tabId;
        return `
          <div class="op-collapse-list">
            ${sorted.map((op, idx) => {
              const name      = op.operator || 'NONE';
              const isNone    = name === 'NONE';
              const total     = op.total ?? ((op.lensTotal||0) + (op.frameTotal||0));
              const maxTotal  = (sorted[0]?.total ?? ((sorted[0]?.lensTotal||0)+(sorted[0]?.frameTotal||0))) || 1;
              const detailId  = `op-detail-${uid}-${idx}`;
              const headerId  = `op-header-${uid}-${idx}`;
              const reasonsList = (op.reasons || []).sort((a,b)=>(b.lenses+b.frames)-(a.lenses+a.frames));
              return `
                <div class="op-group-header ${isNone?'op-none':''}" id="${headerId}" onclick="toggleOperator('${detailId}','${headerId}')">
                  <svg class="op-header-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                  <div class="op-header-name" style="color:${isNone?'var(--red)':color}">${isNone?'⚠ Unassigned':name}</div>
                  <div class="op-header-meta">
                    <span>${reasonsList.length} reason${reasonsList.length!==1?'s':''}</span>
                    ${isFinish && (op.frameTotal||0) > 0 ? `<span style="color:var(--amber)">${op.frameTotal}f</span>` : ''}
                  </div>
                  <div class="op-header-total" style="color:${isNone?'var(--red)':color}">${total}</div>
                </div>
                <div class="op-reasons-detail" id="${detailId}">
                  <div class="op-reasons-inner">
                    ${reasonsList.length ? reasonsList.map(r => `
                      <div class="op-reason-row">
                        <span class="op-reason-tag">${r.reason}</span>
                        <span class="op-reason-bar-wrap">
                          <span class="op-reason-bar" style="width:${total>0?Math.round(((r.lenses+r.frames)/total)*100):0}%;background:${isNone?'var(--red)':color}60"></span>
                        </span>
                        <span class="op-reason-num" style="color:${isNone?'var(--red)':color}">${r.lenses}</span>
                        ${isFinish && r.frames > 0 ? `<span class="op-reason-frames">+${r.frames}f</span>` : ''}
                      </div>`).join('')
                    : '<div class="op-reason-row"><span class="op-reason-tag" style="color:var(--muted)">No reason data</span></div>'}
                  </div>
                </div>`;
            }).join('')}
          </div>`;
      };

      const renderByReason = () => {
        // Group by reason across all operators
        const byReason = {};
        operators.forEach(op => {
          const total = op.total ?? ((op.lensTotal||0)+(op.frameTotal||0));
          if (total === 0 && !(op.reasons||[]).length) return;
          (op.reasons||[]).forEach(r => {
            if (!byReason[r.reason]) byReason[r.reason] = { reason: r.reason, lenses: 0, frames: 0, ops: [] };
            byReason[r.reason].lenses += r.lenses;
            byReason[r.reason].frames += r.frames;
            byReason[r.reason].ops.push({ name: op.operator||'NONE', lenses: r.lenses, frames: r.frames });
          });
        });
        const sortedR = Object.values(byReason).sort((a,b)=>(b.lenses+b.frames)-(a.lenses+a.frames));
        const maxR    = sortedR[0] ? sortedR[0].lenses + sortedR[0].frames : 1;
        const uid     = tabId + 'r';
        return `
          <div class="op-collapse-list">
            ${sortedR.map((r, idx) => {
              const detailId = `op-detail-${uid}-${idx}`;
              const headerId = `op-header-${uid}-${idx}`;
              const total    = r.lenses + r.frames;
              return `
                <div class="op-group-header" id="${headerId}" onclick="toggleOperator('${detailId}','${headerId}')">
                  <svg class="op-header-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                  <div class="op-header-name" style="color:${color}">${r.reason}</div>
                  <div class="op-header-meta"><span>${r.ops.length} ${colLabel.toLowerCase()}${r.ops.length!==1?'s':''}</span></div>
                  <div class="op-header-total" style="color:${color}">${r.lenses}</div>
                </div>
                <div class="op-reasons-detail" id="${detailId}">
                  <div class="op-reasons-inner">
                    ${r.ops.sort((a,b)=>(b.lenses+b.frames)-(a.lenses+a.frames)).map(op => `
                      <div class="op-reason-row">
                        <span class="op-reason-tag">${op.name==='NONE'?'⚠ Unassigned':op.name}</span>
                        <span class="op-reason-bar-wrap">
                          <span class="op-reason-bar" style="width:${Math.round((op.lenses/total)*100)}%;background:${color}60"></span>
                        </span>
                        <span class="op-reason-num" style="color:${color}">${op.lenses}</span>
                      </div>`).join('')}
                  </div>
                </div>`;
            }).join('')}
          </div>`;
      };

      // Toggle state per tab
      if (!window._opFilter) window._opFilter = {};
      if (!window._opFilter[tabId]) window._opFilter[tabId] = 'operator';

      const renderToggle = () => {
        const mode = window._opFilter[tabId];
        const listWrapperId = `op-list-wrap-${tabId}`;
        opsEl.innerHTML = `
          <div class="op-filter-toggle">
            <button class="op-filter-btn ${mode==='operator'?'active':''}"
              onclick="window._opFilter['${tabId}']='operator';document.getElementById('${listWrapperId}').innerHTML=window._opRenderFn['${tabId}']();document.querySelector('#tab-${tabId} [id\$=Operators] .op-filter-btn:first-child').classList.add('active');document.querySelector('#tab-${tabId} [id\$=Operators] .op-filter-btn:last-child').classList.remove('active')">
              By ${colLabel}
            </button>
            <button class="op-filter-btn ${mode==='reason'?'active':''}"
              onclick="window._opFilter['${tabId}']='reason';document.getElementById('${listWrapperId}').innerHTML=window._opRenderFn['${tabId}']();document.querySelector('#tab-${tabId} [id\$=Operators] .op-filter-btn:first-child').classList.remove('active');document.querySelector('#tab-${tabId} [id\$=Operators] .op-filter-btn:last-child').classList.add('active')">
              By Reason
            </button>
          </div>
          <div id="${listWrapperId}">${mode==='operator' ? renderByOperator() : renderByReason()}</div>`;
      };

      if (!window._opRenderFn) window._opRenderFn = {};
      window._opRenderFn[tabId] = () => window._opFilter[tabId]==='operator' ? renderByOperator() : renderByReason();

      renderToggle();
    } else {
      opsEl.innerHTML = U.emptyState(`No ${colLabel.toLowerCase()} data for today`);
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
  // Daily tab removed — skip if elements don't exist
  if (!document.getElementById('dailyFacilityKpis') && !document.getElementById('dailySidebarItems')) return;
  const lt         = summary.labTotal;
  const lensCount  = meta.lensCount  || 1;
  const orderCount = meta.orderCount || 1;
  const labPct     = lt.labLensPct * 100;
  const framePct   = lt.labFramePct * 100;

  // Date badge
  const today    = new Date();
  const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const dateBadge = document.getElementById('dailyDateBadge');
  if (dateBadge) dateBadge.textContent = dayNames[today.getDay()] + ' · ' + (meta.reportDate || today.toLocaleDateString());

  // ── Facility KPIs strip ──
  const dFKpis = document.getElementById('dailyFacilityKpis');
  if (dFKpis) dFKpis.innerHTML = [
    { label: 'Lab Lenses Broken', value: U.fmt(lt.labLensesBroken), accent: 'var(--cyan)',  sub: `of ${U.fmt(lensCount)} lenses` },
    { label: 'Lab Lens %',        value: labPct.toFixed(2) + '%',   accent: U.statusColor(labPct, 5.00),  sub: 'Goal ≤5.00%', badge: U.statusBadge(labPct, 5.00) },
    { label: 'Frames Broken',     value: U.fmt(lt.framesBroken),    accent: 'var(--amber)', sub: `of ${U.fmt(orderCount)} orders` },
    { label: 'Frame Brk %',       value: framePct.toFixed(2) + '%', accent: U.statusColor(framePct, 1.00), sub: 'Goal ≤1.00%', badge: U.statusBadge(framePct, 1.00) },
    { label: 'Total Orders',      value: U.fmt(orderCount),         accent: 'var(--mail)',  sub: 'Mailroom shipped' },
    { label: 'Report Date',       value: meta.reportDate || '—',    accent: 'var(--muted)', sub: 'Data as of' },
  ].map(k => `
    <div class="kpi-card" style="--accent:${k.accent}">
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
      ${k.badge || ''}
    </div>`).join('');

  // ── Surface machine IDs ──
  const SURFACE_MACHINES = ['ORB','54R','FLEX','ODT','OTL','NONE'];
  const isSurfaceMachine = name => SURFACE_MACHINES.some(m => name && name.toUpperCase().startsWith(m));
  const isMeiMachine     = name => name && name.toUpperCase().startsWith('MEI');

  // ── Build dept data list sorted high → low ──
  const deptMap = {};
  summary.departments.forEach(d => { deptMap[d.department] = d; });

  const deptConfigs = [
    { key: 'AR',        label: 'AR',        color: '#a78bfa', goal: 0.50, data: depts.ar,        isMachine: false },
    { key: 'Finish',    label: 'Finish',    color: '#34d399', goal: 1.70, data: depts.finish,     isMachine: false },
    { key: 'Surface',   label: 'Surface',   color: '#fb923c', goal: 2.80, data: depts.surface,    isMachine: true  },
    { key: 'LMS',       label: 'LMS',       color: '#60a5fa', goal: null, data: depts.lms,        isMachine: false },
    { key: 'Inventory', label: 'Inventory', color: '#f472b6', goal: null, data: depts.inventory,  isMachine: false },
  ];

  // Build full sorted dept list with computed pcts
  const deptList = deptConfigs.map(dc => {
    const dept = summary.departments.find(d => d.department === dc.key);
    const broken = dept ? dept.lensesBroken : 0;
    const pct    = broken / lensCount * 100;
    return { ...dc, broken, pct, dept };
  }).sort((a, b) => b.broken - a.broken);

  // Build full Lab total
  const labEntry = {
    key: 'LAB', label: 'Lab Total', color: '#38bdf8',
    broken: lt.labLensesBroken, pct: labPct,
    goal: 5.00, isLab: true,
  };

  const allEntries = [labEntry, ...deptList];

  // ── Build sidebar ──
  const dSide = document.getElementById('dailySidebarItems');
  if (dSide) dSide.innerHTML = allEntries.map((entry, idx) => {
    const ok     = entry.goal ? entry.pct <= entry.goal : true;
    const status = entry.goal ? (ok ? 'ON GOAL' : 'OVER GOAL') : '—';
    const sc     = entry.goal ? (ok ? 'var(--green)' : 'var(--red)') : 'var(--muted)';
    return `
      <div class="daily-sb-item" id="sb-item-${entry.key}"
           style="--item-color:${entry.color}"
           onclick="selectDailyDept('${entry.key}')">
        <div class="daily-sb-dot"></div>
        <div class="daily-sb-info">
          <div class="daily-sb-name ${entry.isLab ? 'lab-total' : ''}">${entry.label}</div>
          <div class="daily-sb-pct" style="color:${sc}">${entry.pct.toFixed(2)}%</div>
          <div class="daily-sb-count">${U.fmt(entry.broken)} lenses</div>
        </div>
        <div class="daily-sb-badge" style="color:${sc};border-color:${sc}30;background:${sc}12">${status}</div>
      </div>`;
  }).join('');

  // ── Store data for click handler ──
  window._dailyData = { allEntries, deptList, labEntry, deptConfigs, summary, depts, meta, lensCount, orderCount, framePct };

  // Auto-select first entry (Lab Total)
  selectDailyDept('LAB');
}

/* ── DAILY DEPT DETAIL RENDERER ─────────────────────────── */
function selectDailyDept(key) {
  if (!window._dailyData) return;
  const { allEntries, deptList, labEntry, deptConfigs, summary, depts, meta, lensCount, orderCount, framePct } = window._dailyData;

  // Update sidebar active state
  document.querySelectorAll('.daily-sb-item').forEach(el => el.classList.remove('active'));
  const activeItem = document.getElementById('sb-item-' + key);
  if (activeItem) activeItem.classList.add('active');

  const pane = document.getElementById('dailyDetailPane');

  // ── LAB TOTAL view ──
  if (key === 'LAB') {
    const labPct = labEntry.pct;
    const ok     = labPct <= 5.00;

    // All reasons across all depts sorted high → low
    const allReasons = [];
    deptConfigs.forEach(dc => {
      (dc.data?.reasons || []).filter(r => r.lensesBroken > 0).forEach(r => {
        allReasons.push({ reason: r.reason, dept: dc.key, count: r.lensesBroken, color: dc.color });
      });
    });
    allReasons.sort((a, b) => b.count - a.count);

    // All depts sorted high → low for summary
    const deptsSorted = [...deptList].sort((a, b) => b.broken - a.broken);
    const overGoal    = deptsSorted.filter(d => d.goal && d.pct > d.goal);
    const onGoal      = deptsSorted.filter(d => !d.goal || d.pct <= d.goal);

    // Unassigned check across all depts
    const unassignedDepts = [];
    deptConfigs.forEach(dc => {
      const noneCount = (dc.data?.operators || [])
        .filter(o => !o.operator || o.operator === 'NONE')
        .reduce((sum, o) => sum + (o.lensTotal || o.total || 0), 0);
      if (noneCount > 0) unassignedDepts.push({ dept: dc.key, count: noneCount });
    });

    // Write professional summary
    const topReason = allReasons[0];
    const summaryText = `As of ${meta.reportDate || 'today'}, the facility recorded ${U.fmt(labEntry.broken)} lens breakages representing a ${labPct.toFixed(2)}% breakage rate across ${U.fmt(lensCount)} lenses processed. The lab is currently ${ok ? 'within' : 'exceeding'} the 5.00% facility goal${!ok ? ` by ${(labPct - 5.00).toFixed(2)} percentage points` : ''}.
${overGoal.length > 0 ? `${overGoal.map(d => d.label).join(' and ')} ${overGoal.length === 1 ? 'is' : 'are'} operating above goal, requiring immediate attention. ` : 'All departments are operating within their respective goals. '}${topReason ? `The leading breakage cause facility-wide is ${topReason.reason} (${topReason.dept}) with ${topReason.count} lenses. ` : ''}${unassignedDepts.length > 0 ? `Unassigned operator entries were detected in ${unassignedDepts.map(u => u.dept).join(', ')} — totaling ${unassignedDepts.reduce((s, u) => s + u.count, 0)} lenses without operator attribution. Data integrity review is recommended.` : 'All breakage records have been attributed to an operator or machine.'}`;

    pane.innerHTML = `
      <div class="dd-header">
        <div>
          <div class="dd-title" style="color:var(--cyan)">FULL LAB SUMMARY</div>
          <div style="font-family:var(--font-mono);font-size:11px;color:var(--muted);margin-top:4px">${meta.reportDate || '—'} · ${U.fmt(lensCount)} lenses · ${U.fmt(orderCount)} orders</div>
        </div>
        <div class="dd-meta">
          <div class="dd-pct" style="color:${ok ? 'var(--green)' : 'var(--red)'}">${labPct.toFixed(2)}%</div>
          <div class="dd-goal" style="color:${ok ? 'var(--green)' : 'var(--red)'}">Lab Goal ≤5.00% — ${ok ? 'On Goal' : 'Over Goal'}</div>
        </div>
      </div>

      ${unassignedDepts.length > 0 ? `
      <div class="dd-unassigned-alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        ${unassignedDepts.reduce((s,u) => s + u.count, 0)} lenses unassigned (NONE operator) across: ${unassignedDepts.map(u => `${u.dept} — ${u.count}`).join(', ')} · Data integrity review required
      </div>` : ''}

      <div class="dd-cards">
        <div class="dd-card">
          <div class="dd-card-label">All Departments — Highest to Lowest</div>
          ${deptsSorted.map(d => `
            <div class="dd-item">
              <div class="dd-item-name" style="color:${d.color};font-weight:600">${d.label}</div>
              <div class="dd-item-bar"><div class="dd-item-fill" style="width:${deptsSorted[0].broken > 0 ? (d.broken/deptsSorted[0].broken*100) : 0}%;background:${d.color}70"></div></div>
              <div class="dd-item-num" style="color:${d.color}">${d.broken}</div>
            </div>`).join('')}
        </div>
        <div class="dd-card">
          <div class="dd-card-label">Top Reasons Facility-Wide</div>
          ${allReasons.slice(0, 8).map(r => `
            <div class="dd-item">
              <div class="dd-item-name">${r.reason} <span style="font-size:9px;color:${r.color};font-family:var(--font-mono)">${r.dept}</span></div>
              <div class="dd-item-bar"><div class="dd-item-fill" style="width:${allReasons[0].count > 0 ? (r.count/allReasons[0].count*100) : 0}%;background:${r.color}70"></div></div>
              <div class="dd-item-num" style="color:${r.color}">${r.count}</div>
            </div>`).join('')}
        </div>
      </div>

      <div class="dd-summary" style="--summary-color:var(--cyan)">
        <div class="dd-summary-label">Executive Summary — ${meta.reportDate || '—'}</div>
        <div class="dd-summary-text">${summaryText}</div>
      </div>`;
    return;
  }

  // ── DEPARTMENT view ──
  const entry  = allEntries.find(e => e.key === key);
  const dc     = deptConfigs.find(d => d.key === key);
  if (!entry || !dc) return;

  const ok         = dc.goal ? entry.pct <= dc.goal : true;
  const color      = entry.color;
  const statusText = dc.goal ? (ok ? `On Goal (${dc.goal}%)` : `Over Goal — ${(entry.pct - dc.goal).toFixed(2)}pp above ${dc.goal}%`) : 'No goal set';
  const statusColor= dc.goal ? (ok ? 'var(--green)' : 'var(--red)') : 'var(--muted)';
  const isMachine  = dc.isMachine;
  const colLabel   = isMachine ? 'Machine' : 'Operator';

  // Reasons sorted high → low, filter out zero
  const reasons = (dc.data?.reasons || [])
    .filter(r => r.lensesBroken > 0 || r.framesBroken > 0)
    .sort((a, b) => (b.lensesBroken || b.framesBroken) - (a.lensesBroken || a.framesBroken));

  // Operators/machines — group by name, sort high → low
  const opMap = {};
  (dc.data?.operators || []).forEach(o => {
    const name    = o.operator || 'NONE';
    const isNone  = !o.operator || o.operator === 'NONE';
    const lenses  = (key === 'Finish') ? (o.lensTotal || 0) : (o.total || 0);
    const frames  = (key === 'Finish') ? (o.frameTotal || 0) : 0;
    if (!opMap[name]) opMap[name] = { name, lenses: 0, frames: 0, isNone };
    opMap[name].lenses += lenses;
    opMap[name].frames += frames;
  });
  const opList = Object.values(opMap)
    .sort((a, b) => b.lenses - a.lenses);

  const maxReasons = reasons[0] ? (reasons[0].lensesBroken || reasons[0].framesBroken) : 1;
  const maxOps     = opList[0]  ? opList[0].lenses : 1;
  const noneTotal  = opList.filter(o => o.isNone).reduce((s, o) => s + o.lenses, 0);
  const topOp      = opList.filter(o => !o.isNone)[0];
  const topReason  = reasons[0];

  // Write professional summary
  const dept = entry.dept;
  const deptFrames = dc?.data?.totals?.framesBroken || 0;
  const framesNote = key === 'Finish' && deptFrames > 0
    ? ` Additionally, ${deptFrames} frames were broken (F-Frame Breakage) at a ${dc.data?.totals ? (dc.data.totals.labFramePct < 0.01 ? (dc.data.totals.labFramePct * 10000).toFixed(2) : (dc.data.totals.labFramePct * 100).toFixed(2)) : '0.00'}% frame breakage rate.` : '';
  const noneNote = noneTotal > 0
    ? ` ${noneTotal} lenses were logged without a valid ${colLabel.toLowerCase()} assignment (recorded as NONE). This represents a data quality concern that should be investigated by the supervisor.` : '';
  const summaryText = `${entry.label} department recorded ${U.fmt(entry.broken)} lens breakages on ${meta.reportDate || 'this date'}, representing a ${entry.pct.toFixed(2)}% breakage rate. ${dc.goal ? `The department is ${ok ? `within` : `exceeding`} its ${dc.goal}% goal${!ok ? `, exceeding the target by ${(entry.pct - dc.goal).toFixed(2)} percentage points` : ''}.` : ''} ${topReason ? `The primary driver is ${topReason.reason} with ${topReason.lensesBroken || topReason.framesBroken} lenses, accounting for ${entry.broken > 0 ? Math.round(((topReason.lensesBroken || topReason.framesBroken)/entry.broken)*100) : 0}% of department breakage.` : ''} ${topOp ? `The highest contributing ${colLabel.toLowerCase()} is ${topOp.name} with ${topOp.lenses} lenses broken.` : ''}${framesNote}${noneNote}`;

  pane.innerHTML = `
    <div class="dd-header">
      <div>
        <div class="dd-title" style="color:${color}">${entry.label.toUpperCase()} DEPARTMENT</div>
        <div style="font-family:var(--font-mono);font-size:11px;color:var(--muted);margin-top:4px">${meta.reportDate || '—'} · ${U.fmt(lensCount)} facility lenses</div>
      </div>
      <div class="dd-meta">
        <div class="dd-pct" style="color:${statusColor}">${entry.pct.toFixed(2)}%</div>
        <div class="dd-goal" style="color:${statusColor}">${statusText}</div>
        <div style="font-family:var(--font-mono);font-size:11px;color:var(--muted);margin-top:3px">${U.fmt(entry.broken)} lenses broken</div>
      </div>
    </div>

    ${noneTotal > 0 ? `
    <div class="dd-unassigned-alert">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      ${noneTotal} lenses recorded with no ${colLabel.toLowerCase()} assignment (NONE) — supervisor review required
    </div>` : ''}

    <div class="dd-cards">
      <div class="dd-card">
        <div class="dd-card-label">Breakage by Reason — Highest to Lowest</div>
        ${reasons.map(r => {
          const cnt = r.lensesBroken || r.framesBroken;
          return `
          <div class="dd-item">
            <div class="dd-item-name">${r.reason}</div>
            <div class="dd-item-bar"><div class="dd-item-fill" style="width:${Math.round((cnt/maxReasons)*100)}%;background:${color}70"></div></div>
            <div class="dd-item-num" style="color:${color}">${cnt}</div>
          </div>`;
        }).join('')}
      </div>
      <div class="dd-card">
        <div class="dd-card-label">${colLabel} Breakdown — Highest to Lowest</div>
        ${opList.map(op => `
          <div class="dd-item">
            <div class="dd-item-name ${op.isNone ? 'unassigned' : ''}">${op.isNone ? 'Unassigned' : op.name}</div>
            <div class="dd-item-bar"><div class="dd-item-fill" style="width:${Math.round((op.lenses/maxOps)*100)}%;background:${op.isNone ? 'var(--red)' : color}70"></div></div>
            <div class="dd-item-num ${op.isNone ? 'unassigned' : ''}" style="${op.isNone ? '' : `color:${color}`}">${op.lenses}${op.frames > 0 ? `<span style="font-size:10px;color:var(--amber)"> +${op.frames}f</span>` : ''}</div>
          </div>`).join('')}
      </div>
    </div>

    <div class="dd-summary" style="--summary-color:${color}">
      <div class="dd-summary-label">Department Summary — ${meta.reportDate || '—'}</div>
      <div class="dd-summary-text">${summaryText}</div>
    </div>`;
}



/* ── WEEKLY SUMMARY (Sun–Sat) ───────────────────────────── */



/* ── OPERATOR ROW TOGGLE ────────────────────────────────── */
function toggleOperator(detailId, headerId) {
  const detail = document.getElementById(detailId);
  const header = document.getElementById(headerId);
  if (!detail || !header) return;
  const isOpen = detail.classList.contains('open');
  if (isOpen) {
    detail.classList.remove('open');
    header.classList.remove('open');
  } else {
    detail.classList.add('open');
    header.classList.add('open');
  }
}


/* ── NO DATA STATE ──────────────────────────────────────── */
function renderNoData(meta) {
  // Update header meta with dashes
  document.getElementById('reportDate').textContent  = '—';
  document.getElementById('lastRefresh').textContent = meta?.lastRefresh || '—';
  document.getElementById('headerSub').textContent   = 'Waiting for RawData upload…';

  const noDataHTML = `
    <div style="
      display:flex; flex-direction:column; align-items:center; justify-content:center;
      min-height:400px; gap:16px; text-align:center;
      background:var(--bg2); border:1px solid var(--border);
      border-radius:var(--radius-md); padding:40px;
    ">
      <svg viewBox="0 0 24 24" fill="none" stroke="var(--muted)" stroke-width="1.5"
           width="48" height="48">
        <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/>
        <polyline points="13 2 13 9 20 9"/>
        <line x1="9" y1="13" x2="15" y2="13"/>
      </svg>
      <div style="font-family:var(--font-display);font-size:22px;letter-spacing:2px;color:var(--text)">
        NO DATA AVAILABLE
      </div>
      <div style="font-family:var(--font-mono);font-size:12px;color:var(--muted);max-width:380px;line-height:1.7">
        RawData is currently empty.<br>
        Data will populate automatically once today's breakage data is uploaded.
      </div>
      <div style="display:flex;gap:10px;margin-top:8px">
        <button onclick="App.loadAll()" style="
          font-family:var(--font-mono);font-size:11px;
          background:var(--bg3);border:1px solid var(--border2);
          border-radius:var(--radius-sm);color:var(--cyan);
          padding:8px 16px;cursor:pointer;letter-spacing:0.5px;
        ">↻ Check Again</button>
        <button onclick="App.switchTab('history')" style="
          font-family:var(--font-mono);font-size:11px;
          background:var(--bg3);border:1px solid var(--border2);
          border-radius:var(--radius-sm);color:var(--muted);
          padding:8px 16px;cursor:pointer;letter-spacing:0.5px;
        ">View History</button>
      </div>
    </div>`;

  // Show no-data state in all main content areas
  const targets = ['labKpis','goalBars','deptCards','topReasonsGrid',
                   'dailyFacilityKpis','dailySidebarItems','dailyDetailPane',
                   'arKpis','finKpis','srfKpis','lmsKpis','invKpis'];
  targets.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = id === 'labKpis' ? noDataHTML : '';
  });

  // Hide goal bars section
  const goalWrap = document.querySelector('.goal-bar-wrap');
  if (goalWrap) goalWrap.style.display = 'none';
  const deptGrid = document.getElementById('deptCards');
  if (deptGrid) deptGrid.innerHTML = '';
  const topGrid = document.getElementById('topReasonsGrid');
  if (topGrid) topGrid.innerHTML = '';
}


/* ── TRANSITION OVERLAY ─────────────────────────────────────── */
function showTransition(mode, subText) {
  const overlay = document.getElementById('transitionOverlay');
  const label   = document.getElementById('transitionLabel');
  const sub     = document.getElementById('transitionSub');
  const svg     = document.getElementById('transitionIconSvg');
  if (!overlay) return;

  overlay.className = `transition-overlay show ${mode === 'live' ? 'live-mode' : 'history-mode'}`;
  label.textContent = mode === 'live' ? 'SWITCHING TO LIVE' : 'LOADING HISTORY';
  sub.textContent   = subText || (mode === 'live' ? 'Reading RawData…' : 'Reading Snapshot_History…');

  // Icon paths
  if (mode === 'live') {
    svg.setAttribute('stroke', '#22d47e');
    svg.innerHTML = '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/>';
  } else {
    svg.setAttribute('stroke', '#f5a623');
    svg.innerHTML = '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>';
  }
}

function hideTransition() {
  const overlay = document.getElementById('transitionOverlay');
  if (overlay) overlay.classList.remove('show');
}


/* ═══════════════════════════════════════════════════════════
   SUMMARY TAB (Daily + Weekly combined)
═══════════════════════════════════════════════════════════ */













/* ── GENERATE REPORT ──────────────────────────────────────── */





/* ═══════════════════════════════════════════════════════════
   SUMMARY TAB
═══════════════════════════════════════════════════════════ */
const SUM = {
  mode: 'daily',
  dates: [],
  weeks: [],
  selectedDate: null,
  selectedWeek: null,
  data: null,
  weekData: null,
  selectedDept: 'LAB',
};

async function initSummaryTab() {
  try {
    // Load available dates
    const res = await API.get('history');
    SUM.dates = res.data || [];

    // Add today if live data available
    if (State.meta?.lensCount > 0 && State.meta?.reportDate) {
      const today = State.meta.reportDate;
      if (!SUM.dates.includes(today)) SUM.dates.unshift(today);
    }

    // Load available weeks
    try {
      const wRes = await fetch(`${API_URL}?action=weekly`);
      const wJson = await wRes.json();
      if (wJson.success) SUM.weeks = wJson.data?.availableWeeks || [];
    } catch(e) {}

    // Populate date selects
    const ds = document.getElementById('sumDateSelect');
    if (ds) {
      ds.innerHTML = '<option value="">Select date…</option>' +
        SUM.dates.map(d => {
          const isLive = d === State.meta?.reportDate && State.meta?.lensCount > 0;
          return `<option value="${d}">${d}${isLive ? ' ● Live' : ''}</option>`;
        }).join('');
    }

    const ws = document.getElementById('sumWeekSelect');
    if (ws) {
      ws.innerHTML = '<option value="">Select week…</option>' +
        SUM.weeks.map(w => `<option value="${w}">Week of ${w}</option>`).join('');
    }

    // Auto-select first date
    if (SUM.dates.length > 0) {
      const ds2 = document.getElementById('sumDateSelect');
      if (ds2) ds2.value = SUM.dates[0];
      await loadSumDate(SUM.dates[0]);
    }
  } catch(e) { console.error('initSummaryTab:', e); }
}

function setSummaryMode(mode) {
  SUM.mode = mode;
  document.getElementById('sumDailyView').style.display = 'block';
}

async function onSumDateChange(date) {
  if (!date) return;
  showTransition('history', 'Loading ' + date + '…');
  await loadSumDate(date);
}

async function onSumWeekChange(week) {
  if (!week) return;
  showTransition('history', 'Loading week of ' + week + '…');
  await loadSumWeek(week);
}

/* ═══════════════════════════════════════════════════════════
   COMPARE TAB — multi-date (up to 7)
═══════════════════════════════════════════════════════════ */
const CMP = {
  dates:    [],   // all available dates
  selected: [],   // checked dates (max 7)
  dataMap:  {},   // date → API data
};

async function initCompareTab() {
  // Reuse SUM.dates if already loaded, otherwise fetch
  if (SUM.dates.length > 0) {
    CMP.dates = SUM.dates;
  } else {
    try {
      const res = await API.get('history');
      CMP.dates = res.data || [];
      if (State.meta?.lensCount > 0 && State.meta?.reportDate) {
        if (!CMP.dates.includes(State.meta.reportDate)) CMP.dates.unshift(State.meta.reportDate);
      }
    } catch(e) { console.error('initCompareTab:', e); return; }
  }
  renderComparePicker();
}

function renderComparePicker() {
  const grid = document.getElementById('cmpDateGrid');
  if (!grid) return;
  grid.innerHTML = CMP.dates.map(d => {
    const isLive    = d === State.meta?.reportDate && State.meta?.lensCount > 0;
    const isChecked = CMP.selected.includes(d);
    return `
      <label class="cmp-date-chip ${isChecked ? 'checked' : ''}" id="cmp-chip-${d.replace(/\//g,'-')}">
        <input type="checkbox" value="${d}" ${isChecked ? 'checked' : ''}
               onchange="onCmpDateToggle(this)"
               style="display:none">
        ${d}${isLive ? ' <span style="color:var(--green);font-size:9px">●</span>' : ''}
      </label>`;
  }).join('');
  updateCmpSelCount();
}

function onCmpDateToggle(cb) {
  const d = cb.value;
  if (cb.checked) {
    if (CMP.selected.length >= 7) {
      cb.checked = false;
      App.showToast('Maximum 7 dates', 'error');
      return;
    }
    if (!CMP.selected.includes(d)) CMP.selected.push(d);
  } else {
    CMP.selected = CMP.selected.filter(x => x !== d);
  }
  // Update chip style
  const chip = document.getElementById('cmp-chip-' + d.replace(/\//g,'-'));
  if (chip) chip.classList.toggle('checked', cb.checked);
  updateCmpSelCount();
}

function updateCmpSelCount() {
  const n   = CMP.selected.length;
  const el  = document.getElementById('cmpSelCount');
  const btn = document.getElementById('cmpRunBtn');
  if (el)  el.textContent = `(${n} selected)`;
  if (btn) {
    const ready = n >= 2;
    btn.style.opacity       = ready ? '1'    : '0.4';
    btn.style.pointerEvents = ready ? 'auto' : 'none';
  }
}

async function runCompare() {
  if (CMP.selected.length < 2) return;
  const resultsEl = document.getElementById('cmpResults');
  resultsEl.innerHTML = '<div class="empty-state" style="padding:40px">Loading data…</div>';

  const fetchDate = async (date) => {
    if (CMP.dataMap[date]) return CMP.dataMap[date];
    const isLive = date === State.meta?.reportDate && State.meta?.lensCount > 0;
    const d = isLive ? State.depts : (await API.get('all', { date })).data;
    CMP.dataMap[date] = d;
    return d;
  };

  try {
    showTransition('history', 'Loading comparison data…');
    const dataArr = await Promise.all(CMP.selected.map(fetchDate));
    hideTransition();
    renderCompareResults(CMP.selected, dataArr, resultsEl);
  } catch(e) {
    hideTransition();
    resultsEl.innerHTML = `<div class="empty-state" style="padding:40px">Error: ${e.message}</div>`;
  }
}

function renderCompareResults(dates, dataArr, el) {
  const n = dates.length;

  // Helper: delta between first date and each subsequent
  const baseData = dataArr[0];
  const lcArr    = dataArr.map(d => d.lensCount  || 1);
  const ocArr    = dataArr.map(d => d.orderCount || 1);

  const pctArr   = dataArr.map((d,i) => (d.summary?.labTotal?.labLensPct  || 0) * 100);
  const fpArr    = dataArr.map((d,i) => (d.summary?.labTotal?.labFramePct || 0) * 100);
  const brkArr   = dataArr.map(d    => d.summary?.labTotal?.labLensesBroken || 0);
  const frmArr   = dataArr.map(d    => d.summary?.labTotal?.framesBroken    || 0);

  const arrowHtml = (base, val, lowerBetter = true) => {
    const diff = val - base;
    if (Math.abs(diff) < 0.001 && typeof diff === 'number') return '';
    if (diff === 0) return '';
    const good  = lowerBetter ? diff < 0 : diff > 0;
    const color = good ? 'var(--green)' : 'var(--red)';
    const sym   = diff > 0 ? '▲' : '▼';
    const abs   = typeof val === 'number' && val % 1 !== 0 ? Math.abs(diff).toFixed(2) : Math.abs(diff).toLocaleString();
    return `<span style="color:${color};font-size:10px;margin-left:4px">${sym}${abs}</span>`;
  };

  // ── Lab totals header strip ──
  const kpiCols = dates.map((d, i) => `
    <div class="cmp-col-kpi ${i === 0 ? 'cmp-col-base' : ''}">
      <div class="cmp-col-date">${d}${i===0 ? ' <span style="font-size:9px;color:var(--muted)">(base)</span>' : ''}</div>
      <div class="cmp-col-stat" style="color:var(--cyan)">${brkArr[i].toLocaleString()}
        ${i > 0 ? arrowHtml(brkArr[0], brkArr[i]) : ''}
      </div>
      <div class="cmp-col-sub">${pctArr[i].toFixed(2)}%
        ${i > 0 ? arrowHtml(pctArr[0], pctArr[i]) : ''}
        <span style="color:${pctArr[i]<=5?'var(--green)':'var(--red)'}"> ${pctArr[i]<=5?'✓':'✗'}</span>
      </div>
      <div class="cmp-col-sub2">${lcArr[i].toLocaleString()} lenses · ${ocArr[i].toLocaleString()} orders</div>
    </div>`).join('');

  // ── Dept table ──
  const DEPTS     = ['AR','Finish','Surface','LMS','Inventory','Breakage'];
  const DEPT_KEYS = { AR:'ar', Finish:'finish', Surface:'surface', LMS:'lms', Inventory:'inventory', Breakage:'breakage' };
  const deptRows  = DEPTS.map(name => {
    const key  = DEPT_KEYS[name];
    const col  = DEPT_COLORS[name] || 'var(--cyan)';
    const vals = dataArr.map((d,i) => {
      const t   = d[key]?.totals || {};
      const pct = t.lensesBroken > 0 ? (t.lensesBroken / lcArr[i] * 100) : 0;
      return { brk: t.lensesBroken || 0, pct };
    });
    const cells = vals.map((v, i) => `
      <td class="cmp-num">${v.brk}${i > 0 ? arrowHtml(vals[0].brk, v.brk) : ''}</td>
      <td class="cmp-pct">${v.pct.toFixed(2)}%${i > 0 ? arrowHtml(vals[0].pct, v.pct) : ''}</td>`).join('');
    return `<tr><td style="color:${col};font-weight:500">${name}</td>${cells}</tr>`;
  }).join('');

  const deptHeadCols = dates.map(d => `<th colspan="2" class="cmp-date-head">${d}</th>`).join('');

  // ── Top reasons table (union, sorted by sum across all dates) ──
  const reasonTotals = {};
  dataArr.forEach(d => {
    DEPTS.forEach(name => {
      (d[DEPT_KEYS[name]]?.reasons || []).forEach(r => {
        reasonTotals[r.reason] = (reasonTotals[r.reason] || 0) + r.lensesBroken;
      });
    });
  });
  const topReasons = Object.keys(reasonTotals).sort((a,b) => reasonTotals[b] - reasonTotals[a]).slice(0, 12);

  const reasonRows = topReasons.map(reason => {
    const vals = dataArr.map(d => {
      let sum = 0;
      DEPTS.forEach(name => {
        const r = (d[DEPT_KEYS[name]]?.reasons || []).find(x => x.reason === reason);
        if (r) sum += r.lensesBroken;
      });
      return sum;
    });
    const cells = vals.map((v,i) => `<td class="cmp-num">${v || '—'}${i>0&&v!==vals[0]?arrowHtml(vals[0],v):''}</td>`).join('');
    return `<tr><td>${reason}</td>${cells}</tr>`;
  }).join('');

  const reasonHeadCols = dates.map(d => `<th class="cmp-num cmp-date-head">${d}</th>`).join('');

  el.innerHTML = `
    <div class="cmp-wrap">

      <!-- Lab totals strip -->
      <div class="panel">
        <div class="panel-header"><div class="panel-title">LAB TOTALS — LENSES BROKEN / LENS %</div></div>
        <div class="panel-body">
          <div class="cmp-kpi-cols">${kpiCols}</div>
        </div>
      </div>

      <div class="cmp-tables-row">

        <!-- Dept breakdown -->
        <div class="panel" style="flex:1.4">
          <div class="panel-header"><div class="panel-title">BY DEPARTMENT</div></div>
          <div class="panel-body">
            <table class="data-table">
              <thead>
                <tr><th>Dept</th>${deptHeadCols}</tr>
                <tr><th></th>${dates.map(() => '<th class="cmp-num">Broken</th><th class="cmp-pct">%</th>').join('')}</tr>
              </thead>
              <tbody>${deptRows}</tbody>
            </table>
          </div>
        </div>

        <!-- Top reasons -->
        <div class="panel" style="flex:1">
          <div class="panel-header"><div class="panel-title">TOP REASONS (facility-wide)</div></div>
          <div class="panel-body">
            <table class="data-table">
              <thead><tr><th>Reason</th>${reasonHeadCols}</tr></thead>
              <tbody>${reasonRows}</tbody>
            </table>
          </div>
        </div>

      </div>
    </div>`;
}

async function loadSumDate(date) {
  SUM.selectedDate = date;
  const isLive = date === State.meta?.reportDate && State.meta?.lensCount > 0;

  let d;
  if (isLive) {
    d = State.depts;
    d.lensCount  = State.meta.lensCount;
    d.orderCount = State.meta.orderCount;
    d.reportDate = State.meta.reportDate;
  } else {
    try {
      const res = await API.get('all', { date });
      d = res.data;
    } catch(e) { App.showToast('Error: ' + e.message, 'error'); hideTransition(); return; }
  }

  SUM.data = d;
  SUM.selectedDept = 'LAB';
  renderSumKpis(d);
  renderSumSidebar(d);
  renderSumMain(d, 'LAB');
  hideTransition();
}

async function loadSumWeek(weekStart) {
  SUM.selectedWeek = weekStart;
  try {
    const res  = await fetch(`${API_URL}?action=weekly&weekStart=${encodeURIComponent(weekStart)}`);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    SUM.weekData = json.data;
    renderSumWeekly(json.data);
    hideTransition();
  } catch(e) { App.showToast('Weekly error: ' + e.message, 'error'); hideTransition(); }
}

function renderSumKpis(d) {
  const lt       = d.summary?.labTotal || {};
  const lc       = d.lensCount  || 1;
  const oc       = d.orderCount || 1;
  const labPct   = (lt.labLensPct  || 0) * 100;
  const framePct = (lt.labFramePct || 0) * 100;
  const dayName  = d.reportDate ? new Date(d.reportDate).toLocaleDateString('en-US',{weekday:'long'}) : '';

  const el = document.getElementById('sumKpiStrip');
  if (!el) return;
  el.innerHTML = `
    <div style="grid-column:1/-1;display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px">
      <div style="font-family:var(--font-display);font-size:15px;letter-spacing:2px">DAILY SUMMARY</div>
      <div style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${dayName.toUpperCase()} · ${d.reportDate||'—'}</div>
    </div>
    ${[
      { label:'Lab Lenses Broken', value:U.fmt(lt.labLensesBroken||0), sub:`of ${U.fmt(lc)} lenses`,    accent:'var(--cyan)' },
      { label:'Lab Lens %',        value:labPct.toFixed(2)+'%',        sub:'Goal ≤5.00%', accent:U.statusColor(labPct,5.00), badge:U.statusBadge(labPct,5.00) },
      { label:'Frames Broken',     value:U.fmt(lt.framesBroken||0),    sub:`of ${U.fmt(oc)} orders`,    accent:'var(--amber)' },
      { label:'Frame Brk %',       value:framePct.toFixed(2)+'%',      sub:'Goal ≤1.00%', accent:U.statusColor(framePct,1.00), badge:U.statusBadge(framePct,1.00) },
      { label:'Total Orders',      value:U.fmt(oc),                    sub:'Mailroom shipped', accent:'var(--mail)' },
      { label:'Report Date',       value:d.reportDate||'—',            sub:'Date as of', accent:'var(--muted)' },
    ].map(k => `
      <div class="kpi-card" style="--accent:${k.accent}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
        <div class="kpi-sub">${k.sub}</div>
        ${k.badge||''}
      </div>`).join('')}`;
}

function renderSumSidebar(d) {
  const el = document.getElementById('sumDeptSidebarItems');
  if (!el) return;
  const lc = d.lensCount || 1;
  const lt = d.summary?.labTotal || {};
  const labPct = (lt.labLensPct||0)*100;

  // Lab Total + depts sorted high→low
  const depts = [...(d.summary?.departments||[])].sort((a,b)=>b.lensesBroken-a.lensesBroken);
  const allItems = [
    { key:'LAB', label:'Lab Total', color:'var(--cyan)', pct:labPct, broken:lt.labLensesBroken||0, goal:5.00 },
    ...depts.map(dep => ({
      key:    dep.department,
      label:  dep.department,
      color:  DEPT_COLORS[dep.department]||'var(--cyan)',
      pct:    (dep.lensBrkPct||0)*100,
      broken: dep.lensesBroken,
      goal:   {AR:0.50,Finish:1.70,Surface:2.80}[dep.department]||null,
    }))
  ];

  el.innerHTML = allItems.map(item => {
    const ok  = item.goal ? item.pct <= item.goal : true;
    const sc  = item.goal ? (ok?'var(--green)':'var(--red)') : 'var(--muted)';
    const bad = item.goal ? (ok?'ON GOAL':'OVER') : '—';
    return `
      <div class="sum-dept-item ${item.key===SUM.selectedDept?'active':''}"
           style="--item-color:${item.color}"
           onclick="selectSumDept('${item.key}')">
        <div class="sum-dept-dot"></div>
        <div class="sum-dept-info">
          <div class="sum-dept-name">${item.label}</div>
          <div class="sum-dept-pct" style="color:${sc}">${item.pct.toFixed(2)}%</div>
          <div class="sum-dept-count">${U.fmt(item.broken)} lenses</div>
        </div>
        <div class="sum-dept-badge" style="color:${sc};border:1px solid ${sc}30;background:${sc}10">${bad}</div>
      </div>`;
  }).join('');
}

function selectSumDept(key) {
  SUM.selectedDept = key;
  // Update active state
  document.querySelectorAll('.sum-dept-item').forEach(el => {
    el.classList.toggle('active', el.onclick?.toString().includes(`'${key}'`));
  });
  renderSumSidebar(SUM.data); // re-render to update active
  renderSumMain(SUM.data, key);
}

function renderSumMain(d, deptKey) {
  const el = document.getElementById('sumMain');
  if (!el) return;
  const lc = d.lensCount || 1;
  const oc = d.orderCount || 1;

  if (deptKey === 'LAB') {
    // Full lab summary view (matches image 2 right panel)
    const lt      = d.summary?.labTotal || {};
    const labPct  = (lt.labLensPct||0)*100;
    const ok      = labPct <= 5.00;
    const depts   = [...(d.summary?.departments||[])].sort((a,b)=>b.lensesBroken-a.lensesBroken);
    const maxD    = depts[0]?.lensesBroken || 1;

    // Collect all reasons facility-wide sorted high→low
    const allReasons = [];
    const deptKeys = ['AR','Finish','Surface','LMS','Inventory'];
    deptKeys.forEach(dep => {
      (d.summary?.topReasons?.[dep]||[]).forEach(r => {
        allReasons.push({ reason:r.reason, dept:dep, count:r.lensesBroken, color:DEPT_COLORS[dep]||'var(--cyan)' });
      });
    });
    allReasons.sort((a,b)=>b.count-a.count);
    const maxR = allReasons[0]?.count||1;

    // Unassigned check
    const noneTotal = deptKeys.reduce((sum, dep) => {
      const opData = d[dep.toLowerCase()]?.operators||[];
      return sum + opData.filter(o=>!o.operator||o.operator==='NONE')
        .reduce((s,o)=>s+(o.total||(o.lensTotal||0)),0);
    }, 0);

    const noneDepts = deptKeys.filter(dep => {
      const opData = d[dep.toLowerCase()]?.operators||[];
      return opData.some(o=>!o.operator||o.operator==='NONE');
    });

    // Exec summary text
    const topReason = allReasons[0];
    const overGoal  = depts.filter(dep=>dep.lensBrkPct*100>({AR:0.50,Finish:1.70,Surface:2.80}[dep.department]||999));
    const execText  = `As of ${d.reportDate||'today'}, the facility recorded ${U.fmt(lt.labLensesBroken||0)} lens breakages representing a ${labPct.toFixed(2)}% breakage rate across ${U.fmt(lc)} lenses processed. The lab is currently ${ok?'within':'exceeding'} the 5.00% facility goal.${overGoal.length?' '+overGoal.map(dep=>`${dep.department} is over goal at ${(dep.lensBrkPct*100).toFixed(2)}%.`).join(' '):'  All departments are operating within their respective goals.'} ${topReason?`The leading breakage cause facility-wide is ${topReason.reason} (${topReason.dept}) with ${topReason.count} lenses.`:''} ${noneTotal>0?`Unassigned operator entries were detected in ${noneDepts.join(', ')} — totaling ${noneTotal} lenses without operator attribution. Data integrity review is recommended.`:'All breakage records have been attributed to an operator or machine.'}`;

    el.innerHTML = `
      <div class="sum-main-title" style="color:var(--cyan)">FULL LAB SUMMARY</div>
      <div class="sum-main-subtitle">${d.reportDate||'—'} · ${U.fmt(lc)} lenses · ${U.fmt(oc)} orders</div>

      ${noneTotal>0?`
      <div class="sum-none-alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        ${noneTotal} lenses unassigned (NONE operator) across: ${noneDepts.join(', ')} · Data integrity review required
      </div>`:''}

      <div class="sum-two-col">
        <div class="sum-card">
          <div class="sum-card-label">All Departments — Highest to Lowest</div>
          ${depts.map(dep => {
            const color = DEPT_COLORS[dep.department]||'var(--cyan)';
            return `<div class="sum-item">
              <div class="sum-item-name dept-label" style="color:${color}">${dep.department}</div>
              <div class="sum-item-bar"><div class="sum-item-fill" style="width:${Math.round((dep.lensesBroken/maxD)*100)}%;background:${color}70"></div></div>
              <div class="sum-item-num" style="color:${color}">${dep.lensesBroken}</div>
            </div>`;
          }).join('')}
        </div>
        <div class="sum-card">
          <div class="sum-card-label">Top Reasons Facility-Wide</div>
          ${allReasons.slice(0,8).map(r => `
            <div class="sum-item">
              <div class="sum-item-name">${r.reason} <span style="font-size:9px;color:${r.color};font-family:var(--font-mono)">${r.dept}</span></div>
              <div class="sum-item-bar"><div class="sum-item-fill" style="width:${Math.round((r.count/maxR)*100)}%;background:${r.color}70"></div></div>
              <div class="sum-item-num" style="color:${r.color}">${r.count}</div>
            </div>`).join('')}
        </div>
      </div>

      <div class="sum-exec" style="--exec-color:var(--cyan)">
        <div class="sum-exec-label">Executive Summary — ${d.reportDate||'—'}</div>
        <div class="sum-exec-text">${execText}</div>
      </div>`;

  } else {
    // Single dept view
    const depKey  = deptKey.toLowerCase();
    const depData = d[depKey] || {};
    const depInfo = d.summary?.departments?.find(x=>x.department===deptKey)||{};
    const color   = DEPT_COLORS[deptKey]||'var(--cyan)';
    const pct     = (depInfo.lensBrkPct||0)*100;
    const goal    = {AR:0.50,Finish:1.70,Surface:2.80}[deptKey]||null;
    const ok      = goal?pct<=goal:true;
    const sc      = goal?(ok?'var(--green)':'var(--red)'):'var(--muted)';
    const isSurface = deptKey === 'surface';
    const colLabel  = isSurface ? 'Machine' : 'Operator';

    const reasons  = (depData.reasons||[]).filter(r=>r.lensesBroken>0||r.framesBroken>0);
    const ops      = depData.operators||[];
    const maxR     = reasons[0]?.lensesBroken||reasons[0]?.framesBroken||1;
    const maxO     = ops[0] ? (ops[0].total||(ops[0].lensTotal||0)) : 1;

    const noneOps = ops.filter(o=>!o.operator||o.operator==='NONE');
    const noneTotal = noneOps.reduce((s,o)=>s+(o.total||(o.lensTotal||0)),0);

    el.innerHTML = `
      <div class="sum-main-title" style="color:${color}">${deptKey.toUpperCase()} DEPARTMENT</div>
      <div class="sum-main-subtitle">${d.reportDate||'—'} · ${U.fmt(depInfo.lensesBroken||0)} lenses broken · ${goal?`Goal ≤${goal}% — `:''}${goal?(ok?'<span style="color:var(--green)">On Goal</span>':'<span style="color:var(--red)">Over Goal</span>'):'No goal set'}</div>

      ${noneTotal>0?`
      <div class="sum-none-alert">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        ${noneTotal} lenses with no ${colLabel.toLowerCase()} assigned (NONE) · supervisor review required
      </div>`:''}

      <div class="sum-two-col">
        <div class="sum-card">
          <div class="sum-card-label">Breakage by Reason — Highest to Lowest</div>
          ${reasons.slice(0,8).map(r=>{
            const cnt=r.lensesBroken||r.framesBroken;
            return `<div class="sum-item">
              <div class="sum-item-name">${r.reason}</div>
              <div class="sum-item-bar"><div class="sum-item-fill" style="width:${Math.round((cnt/maxR)*100)}%;background:${color}70"></div></div>
              <div class="sum-item-num" style="color:${color}">${cnt}</div>
            </div>`;
          }).join('')||'<div style="color:var(--muted);font-family:var(--font-mono);font-size:11px">No breakage data</div>'}
        </div>
        <div class="sum-card">
          <div class="sum-card-label">${colLabel} Breakdown — Highest to Lowest</div>
          ${ops.slice(0,10).map(op=>{
            const isNone=!op.operator||op.operator==='NONE';
            const total=op.total||(op.lensTotal||0);
            return `<div class="sum-item">
              <div class="sum-item-name${isNone?' unassigned':''}">${isNone?'Unassigned':op.operator}</div>
              <div class="sum-item-bar"><div class="sum-item-fill" style="width:${Math.round((total/maxO)*100)}%;background:${isNone?'var(--red)':color}70"></div></div>
              <div class="sum-item-num${isNone?' unassigned':''}" style="${isNone?'':'color:'+color}">${total}</div>
            </div>`;
          }).join('')||'<div style="color:var(--muted);font-family:var(--font-mono);font-size:11px">No operator data</div>'}
        </div>
      </div>

      <div class="sum-exec" style="--exec-color:${color}">
        <div class="sum-exec-label">Department Summary — ${d.reportDate||'—'}</div>
        <div class="sum-exec-text">${deptKey} department recorded ${U.fmt(depInfo.lensesBroken||0)} lens breakages on ${d.reportDate||'this date'}, representing a ${pct.toFixed(2)}% breakage rate.${goal?` The department is ${ok?'within':'exceeding'} its ${goal}% goal${!ok?`, exceeding the target by ${(pct-goal).toFixed(2)} percentage points`:''}. `:' '}${reasons[0]?`The primary driver is ${reasons[0].reason} with ${reasons[0].lensesBroken||reasons[0].framesBroken} lenses.`:''}${noneTotal>0?` ${noneTotal} lenses were logged without a valid ${colLabel.toLowerCase()} (NONE). Data quality review recommended.`:''}</div>
      </div>`;
  }
}

function renderSumWeekly(data) {
  const { days, weekStart, weekEnd } = data;
  const daysWithData = days.filter(d=>d.hasData);
  const avg = arr => arr.length ? (arr.reduce((a,b)=>a+b,0)/arr.length) : 0;

  const el = document.getElementById('sumWeekKpiStrip');
  if (el) {
    const avgLab = avg(daysWithData.map(d=>d.labLensPct||0));
    const total  = daysWithData.reduce((s,d)=>s+(d.labLenses||0),0);
    el.innerHTML = `
      <div style="grid-column:1/-1;display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px">
        <div style="font-family:var(--font-display);font-size:15px;letter-spacing:2px">WEEKLY SUMMARY</div>
        <div style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${weekStart} – ${weekEnd}</div>
      </div>
      ${[
        {label:'Days Captured',  value:daysWithData.length+'/7', accent:'var(--muted)'},
        {label:'Total Broken',   value:U.fmt(total),             accent:'var(--cyan)'},
        {label:'Avg Lab %',      value:avgLab.toFixed(2)+'%',    accent:U.statusColor(avgLab,5.00), badge:U.statusBadge(avgLab,5.00)},
        {label:'Avg Surface %',  value:avg(daysWithData.map(d=>d.srfPct||0)).toFixed(2)+'%', accent:U.statusColor(avg(daysWithData.map(d=>d.srfPct||0)),2.80)},
        {label:'Avg Finish %',   value:avg(daysWithData.map(d=>d.finPct||0)).toFixed(2)+'%', accent:U.statusColor(avg(daysWithData.map(d=>d.finPct||0)),1.70)},
        {label:'Avg AR %',       value:avg(daysWithData.map(d=>d.arPct||0)).toFixed(2)+'%',  accent:U.statusColor(avg(daysWithData.map(d=>d.arPct||0)),0.50)},
      ].map(k=>`<div class="kpi-card" style="--accent:${k.accent}">
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value" style="color:${k.accent}">${k.value}</div>
        ${k.badge||''}
      </div>`).join('')}`;
  }

  const wd = document.getElementById('sumWeekDetail');
  if (!wd) return;
  const DCOLS = [{key:'arPct',c:'var(--ar)',g:0.50,l:'AR'},{key:'finPct',c:'var(--fin)',g:1.70,l:'Fin'},{key:'srfPct',c:'var(--srf)',g:2.80,l:'Srf'}];
  wd.innerHTML = `
    <div class="panel">
      <div class="panel-header"><div class="panel-title">Day-by-Day Breakdown</div></div>
      <div class="panel-body" style="padding:0">
        <table class="dept-breakdown-table">
          <thead><tr>
            <th>Day</th><th style="text-align:right">Lab %</th>
            ${DCOLS.map(c=>`<th style="text-align:right;color:${c.c}">${c.l}</th>`).join('')}
            <th>Top Reason</th>
          </tr></thead>
          <tbody>
            ${days.map(day=>{
              if(!day.hasData) return `<tr style="opacity:${day.isFuture?0.2:0.5}">
                <td style="font-family:var(--font-mono);font-size:11px">${day.day} ${day.date.slice(0,5)}</td>
                <td colspan="5" style="color:var(--muted);font-family:var(--font-mono);font-size:11px">${day.isFuture?'—':'No snapshot'}</td>
              </tr>`;
              const labOk=day.labLensPct<=5.00;
              const topR=day.topReasons?.Surface?.[0]?.reason||day.topReasons?.Finish?.[0]?.reason||'—';
              return `<tr>
                <td style="font-family:var(--font-mono);font-size:11px;font-weight:600">${day.day} ${day.date.slice(0,5)}</td>
                <td style="text-align:right;font-family:var(--font-mono);font-size:12px;color:${labOk?'var(--green)':'var(--red)'}">${day.labLensPct.toFixed(2)}%</td>
                ${DCOLS.map(c=>{const v=day[c.key]||0;return`<td style="text-align:right;font-family:var(--font-mono);font-size:12px;color:${v<=c.g?'var(--green)':'var(--red)'}">${v.toFixed(2)}%</td>`;}).join('')}
                <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${topR}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}


/* ═══════════════════════════════════════════════════════════
   GENERATE SUMMARY REPORT — per selected dept or full lab
═══════════════════════════════════════════════════════════ */
function generateSummaryReport() {
  const panel  = document.getElementById('sumReportPanel');
  const btn    = document.getElementById('sumReportBtn');
  const d      = SUM.data;
  const deptKey = SUM.selectedDept || 'LAB';

  if (!d) { App.showToast('Select a date first', 'error'); return; }

  // Toggle off if already showing
  if (panel.style.display !== 'none') {
    panel.style.display = 'none';
    btn.textContent = '📄 Generate Report';
    return;
  }

  btn.textContent = 'Building…';

  const lensCount  = d.lensCount  || 1;
  const orderCount = d.orderCount || 1;
  const date       = d.reportDate || SUM.selectedDate || '—';
  const lt         = d.summary?.labTotal || {};
  const labPct     = (lt.labLensPct  || 0) * 100;
  const framePct   = (lt.labFramePct || 0) * 100;
  const DEPT_GOALS = { AR:0.50, Finish:1.70, Surface:2.80 };
  const colLabel   = key => key === 'Surface' ? 'Machine' : 'Operator';
  const fmt        = n => (n||0).toLocaleString();
  const pf         = n => (n||0).toFixed(2) + '%';

  // Helper: status color style string
  const sc  = (ok) => ok ? 'color:var(--green)' : 'color:var(--red)';
  const scG = (v, goal) => !goal ? 'color:var(--muted)' : sc(v <= goal);

  // ── FULL LAB REPORT ──────────────────────────────────────
  if (deptKey === 'LAB') {
    const depts = [...(d.summary?.departments || [])].sort((a,b) => b.lensesBroken - a.lensesBroken);
    const overGoal = depts.filter(dep => DEPT_GOALS[dep.department] && dep.lensBrkPct*100 > DEPT_GOALS[dep.department]);
    const onGoal   = depts.filter(dep => !DEPT_GOALS[dep.department] || dep.lensBrkPct*100 <= DEPT_GOALS[dep.department]);

    // All reasons across all depts sorted high → low
    const allReasons = [];
    ['AR','Finish','Surface','LMS','Inventory'].forEach(dep => {
      const depData = d[dep.toLowerCase()];
      (depData?.reasons || []).forEach(r => {
        allReasons.push({ reason: r.reason, dept: dep, count: r.lensesBroken||0,
          color: DEPT_COLORS[dep]||'var(--cyan)',
          pctLab:  ((r.lensesBroken||0)/lensCount*100).toFixed(2),
          pctDept: depts.find(x=>x.department===dep)?.lensesBroken > 0
            ? ((r.lensesBroken||0)/depts.find(x=>x.department===dep).lensesBroken*100).toFixed(0) : '0',
        });
      });
    });
    allReasons.sort((a,b) => b.count - a.count);

    // All operators across all depts sorted high → low (no NONE)
    const allOps = [];
    ['AR','Finish','Surface','LMS','Inventory'].forEach(dep => {
      const depData = d[dep.toLowerCase()];
      (depData?.operators || []).filter(o => o.operator && o.operator !== 'NONE').forEach(o => {
        const total = o.total ?? ((o.lensTotal||0) + (o.frameTotal||0));
        if (total > 0) allOps.push({ name: o.operator, dept: dep, color: DEPT_COLORS[dep]||'var(--cyan)',
          lenses: total, frames: o.frameTotal||0,
          topReason: (o.reasons||[])[0]?.reason || '—', reasonCount: (o.reasons||[]).length });
      });
    });
    allOps.sort((a,b) => b.lenses - a.lenses);

    // Unassigned per dept
    const unassigned = [];
    ['AR','Finish','Surface','LMS','Inventory'].forEach(dep => {
      const depData = d[dep.toLowerCase()];
      const noneOps = (depData?.operators || []).filter(o => !o.operator || o.operator === 'NONE');
      const count   = noneOps.reduce((s,o) => s + (o.total||(o.lensTotal||0)), 0);
      const reasons = noneOps.flatMap(o => o.reasons||[]).sort((a,b)=>(b.lenses||0)-(a.lenses||0));
      if (count > 0) unassigned.push({ dept: dep, color: DEPT_COLORS[dep]||'var(--cyan)',
        count, reasons, deptTotal: depts.find(x=>x.department===dep)?.lensesBroken||0 });
    });
    const totalUnassigned = unassigned.reduce((s,u) => s+u.count, 0);

    // Frame breakage detail
    const frameOps = [];
    const finData = d.finish;
    (finData?.operators || []).filter(o => (o.frameTotal||0) > 0).forEach(o => {
      frameOps.push({ name: o.operator||'Unassigned', frames: o.frameTotal, none: !o.operator||o.operator==='NONE' });
    });

    const labOk   = labPct <= 5.00;
    const frameOk = framePct <= 1.00;

    panel.style.setProperty('--report-color', 'var(--cyan)');
    panel.innerHTML = `
      <div class="rpt-h1" style="color:var(--cyan)">FULL LAB BREAKAGE REPORT</div>
      <div class="rpt-meta">${date} &nbsp;·&nbsp; ${fmt(lensCount)} lenses processed &nbsp;·&nbsp; ${fmt(orderCount)} orders &nbsp;·&nbsp; Generated ${new Date().toLocaleTimeString()}</div>

      <!-- 1. Facility Overview -->
      <div class="rpt-h2">1. Facility Overview</div>
      <p class="rpt-p">On <strong>${date}</strong>, the facility processed <strong style="color:var(--cyan)">${fmt(lensCount)} lenses</strong> across <strong>${fmt(orderCount)} orders</strong>.
      A total of <strong style="${sc(labOk)}">${fmt(lt.labLensesBroken||0)} lenses</strong> were broken, representing a <strong style="${sc(labOk)}">${labPct.toFixed(2)}%</strong> lab-wide breakage rate.
      The facility is <strong style="${sc(labOk)}">${labOk ? 'within' : 'exceeding'}</strong> the 5.00% facility goal${labOk ? ', with <strong>' + (5.00-labPct).toFixed(2) + 'pp of margin remaining</strong>.' : ', exceeding by <strong>' + (labPct-5.00).toFixed(2) + 'pp</strong>.'}
      Frame breakage: <strong style="${sc(frameOk)}">${fmt(lt.framesBroken||0)} frames (${framePct.toFixed(2)}%)</strong> — ${frameOk ? 'within' : 'exceeding'} the 1.00% goal.</p>

      <!-- 2. Goal Status Table -->
      <div class="rpt-h2">2. Department Goal Status</div>
      ${overGoal.length > 0
        ? `<div class="rpt-flag err">⚠ ${overGoal.length} department${overGoal.length>1?'s':''} over goal: ${overGoal.map(d=>d.department).join(', ')}</div>`
        : `<div class="rpt-flag ok">✓ All departments with goals are operating within their targets.</div>`}
      <table class="rpt-table">
        <thead><tr><th>Department</th><th>Broken</th><th>Rate</th><th>Goal</th><th>Status</th><th>Margin</th><th>Top Reason</th></tr></thead>
        <tbody>
          ${depts.filter(dep => DEPT_GOALS[dep.department]).map(dep => {
            const goal   = DEPT_GOALS[dep.department];
            const pct    = (dep.lensBrkPct||0)*100;
            const ok     = pct <= goal;
            const margin = (goal - pct).toFixed(2);
            const topR   = (d[dep.department.toLowerCase()]?.reasons||[])[0]?.reason || '—';
            return `<tr>
              <td style="font-weight:600;color:${DEPT_COLORS[dep.department]||'var(--cyan)'}">${dep.department}</td>
              <td style="font-family:var(--font-mono)">${dep.lensesBroken}</td>
              <td style="font-family:var(--font-mono);${sc(ok)}">${pct.toFixed(2)}%</td>
              <td style="font-family:var(--font-mono);color:var(--muted)">≤${goal}%</td>
              <td style="font-family:var(--font-mono);${sc(ok)}">${ok?'✓ On Goal':'✗ Over Goal'}</td>
              <td style="font-family:var(--font-mono);${sc(ok)}">${ok?'+'+margin:'−'+Math.abs(margin)}pp</td>
              <td style="font-size:11px;color:var(--muted)">${topR}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>

      <!-- 3. Dept Deep Dive -->
      <div class="rpt-h2">3. Department Deep Dive — Highest to Lowest</div>
      ${depts.map(dep => {
        const depKey  = dep.department.toLowerCase();
        const depData = d[depKey] || {};
        const goal    = DEPT_GOALS[dep.department];
        const pct     = (dep.lensBrkPct||0)*100;
        const ok      = !goal || pct <= goal;
        const color   = DEPT_COLORS[dep.department]||'var(--cyan)';
        const reasons = (depData.reasons||[]).slice(0,6);
        const ops     = (depData.operators||[]).filter(o=>o.operator&&o.operator!=='NONE').slice(0,6);
        const noneC   = (depData.operators||[]).filter(o=>!o.operator||o.operator==='NONE').reduce((s,o)=>s+(o.total||(o.lensTotal||0)),0);
        const maxR    = reasons[0] ? (reasons[0].lensesBroken||reasons[0].framesBroken||1) : 1;
        return `<div class="rpt-dept-card" style="--dc:${color}">
          <div class="rpt-dept-head">
            <span style="font-family:var(--font-display);font-size:16px;font-weight:700;color:${color}">${dep.department}</span>
            <span style="font-family:var(--font-mono);font-size:12px;${sc(ok)}">${pct.toFixed(2)}%</span>
            ${goal ? `<span style="font-family:var(--font-mono);font-size:10px;color:var(--muted)">Goal ≤${goal}%</span>` : ''}
            <span style="font-family:var(--font-display);font-size:20px;font-weight:700;color:${color};margin-left:auto">${dep.lensesBroken} lenses</span>
          </div>
          ${reasons.length ? `<div style="font-family:var(--font-mono);font-size:9px;color:var(--muted);margin-bottom:6px">TOP REASONS</div>
          <div class="rpt-reasons-grid">
            ${reasons.map(r => {
              const cnt = r.lensesBroken||r.framesBroken||0;
              return `<div class="rpt-reason-chip">
                <div style="width:4px;height:4px;border-radius:50%;background:${color};flex-shrink:0"></div>
                <span style="flex:1;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.reason}</span>
                <span style="font-family:var(--font-display);font-size:13px;font-weight:700;color:${color}">${cnt}</span>
                <span style="font-family:var(--font-mono);font-size:9px;color:var(--muted)">${dep.lensesBroken>0?(cnt/dep.lensesBroken*100).toFixed(0):0}%</span>
              </div>`;
            }).join('')}
          </div>` : ''}
          ${ops.length ? `<div style="font-family:var(--font-mono);font-size:9px;color:var(--muted);margin-bottom:6px;margin-top:8px">TOP ${dep.department==='Surface'?'MACHINES':'OPERATORS'}</div>
          <div class="rpt-op-chips">
            ${ops.map(op => {
              const total = op.total??(op.lensTotal||0);
              return `<div class="rpt-op-chip" style="border:1px solid ${color}30;background:${color}10;color:${color}">${op.operator} — ${total}${(op.frameTotal||0)>0?' (+'+op.frameTotal+'f)':''}</div>`;
            }).join('')}
          </div>` : ''}
          ${noneC > 0 ? `<div style="font-family:var(--font-mono);font-size:10px;color:var(--red);margin-top:8px;padding:5px 10px;background:rgba(240,79,90,0.08);border-radius:6px;border:1px solid rgba(240,79,90,0.2)">⚠ ${noneC} lenses unassigned (NONE) — supervisor review required</div>` : ''}
        </div>`;
      }).join('')}

      <!-- 4. Top Reasons Facility-Wide -->
      <div class="rpt-h2">4. Top Breakage Reasons — Facility-Wide</div>
      <table class="rpt-table">
        <thead><tr><th>#</th><th>Reason</th><th>Dept</th><th>Count</th><th>% of Lab</th><th>% of Dept</th></tr></thead>
        <tbody>
          ${allReasons.slice(0,12).map((r,i) => `<tr>
            <td style="font-family:var(--font-mono);color:var(--muted)">${i+1}</td>
            <td style="font-weight:500">${r.reason}</td>
            <td style="color:${r.color};font-family:var(--font-mono);font-size:11px">${r.dept}</td>
            <td style="font-family:var(--font-display);font-weight:700;color:${r.color}">${r.count}</td>
            <td style="font-family:var(--font-mono);font-size:11px">${r.pctLab}%</td>
            <td style="font-family:var(--font-mono);font-size:11px">${r.pctDept}%</td>
          </tr>`).join('')}
        </tbody>
      </table>

      <!-- 5. Top Operators -->
      ${allOps.length > 0 ? `<div class="rpt-h2">5. Top Contributing Operators — All Departments</div>
      <table class="rpt-table">
        <thead><tr><th>Name</th><th>Dept</th><th>Lenses</th><th>Frames</th><th>Top Reason</th><th>Reasons</th></tr></thead>
        <tbody>
          ${allOps.slice(0,10).map(op => `<tr>
            <td style="font-weight:500;color:${op.color}">${op.name}</td>
            <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${op.dept}</td>
            <td style="font-family:var(--font-display);font-weight:700;color:${op.color}">${op.lenses}</td>
            <td style="font-family:var(--font-mono);font-size:11px;color:${op.frames>0?'var(--amber)':'var(--muted)'}">${op.frames>0?op.frames:'—'}</td>
            <td style="font-size:11px;color:var(--muted)">${op.topReason}</td>
            <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${op.reasonCount}</td>
          </tr>`).join('')}
        </tbody>
      </table>` : ''}

      <!-- 6. Data Integrity -->
      ${totalUnassigned > 0 ? `<div class="rpt-h2">6. Data Integrity — Unassigned Records</div>
      <p class="rpt-p" style="color:var(--red)"><strong>${totalUnassigned} lenses</strong> were recorded without a valid operator assignment (NONE). These cannot be attributed for performance tracking. Supervisor review required for: ${unassigned.map(u=>u.dept).join(', ')}.</p>
      <table class="rpt-table">
        <thead><tr><th>Dept</th><th>Unassigned</th><th>% of Dept</th><th>Top Unassigned Reasons</th></tr></thead>
        <tbody>
          ${unassigned.map(u => `<tr>
            <td style="font-weight:600;color:${u.color}">${u.dept}</td>
            <td style="font-family:var(--font-display);font-weight:700;color:var(--red)">${u.count}</td>
            <td style="font-family:var(--font-mono);font-size:11px;color:var(--red)">${u.deptTotal>0?(u.count/u.deptTotal*100).toFixed(0):0}% of dept</td>
            <td style="font-size:11px;color:var(--muted)">${u.reasons.slice(0,3).map(r=>r.reason||r.r||'—').join(', ')}</td>
          </tr>`).join('')}
        </tbody>
      </table>` : ''}

      <!-- 7. Frame Breakage -->
      ${(lt.framesBroken||0) > 0 ? `<div class="rpt-h2">${totalUnassigned>0?'7':'6'}. Frame Breakage Detail</div>
      <p class="rpt-p">${lt.framesBroken} frames broken at ${framePct.toFixed(2)}% (${frameOk?'within':'exceeding'} the 1.00% goal).
      ${frameOps.length>0?'Frame breakage attributed to: '+frameOps.map(o=>o.name+' ('+o.frames+' frame'+(o.frames>1?'s':'')+')').join(', '):''}</p>` : ''}

      <!-- 8. Key Observations -->
      <div class="rpt-h2">${totalUnassigned>0&&(lt.framesBroken||0)>0?'8':totalUnassigned>0||(lt.framesBroken||0)>0?'7':'6'}. Key Observations & Recommendations</div>
      ${overGoal.map(dep => {
        const goal = DEPT_GOALS[dep.department];
        const pct  = (dep.lensBrkPct||0)*100;
        const depData = d[dep.department.toLowerCase()]||{};
        const topR = (depData.reasons||[]).slice(0,2).map(r=>r.reason).join(' and ');
        return `<div class="rpt-flag err">✗ ${dep.department} is ${(pct-goal).toFixed(2)}pp over goal — review ${topR||'top reasons'} with lead supervisor</div>`;
      }).join('')}
      ${depts.filter(dep => DEPT_GOALS[dep.department] && (dep.lensBrkPct*100 <= DEPT_GOALS[dep.department]) && (DEPT_GOALS[dep.department] - dep.lensBrkPct*100) < 0.30).map(dep =>
        `<div class="rpt-flag warn">⚡ ${dep.department} is within ${(DEPT_GOALS[dep.department]-dep.lensBrkPct*100).toFixed(2)}pp of goal — monitor closely</div>`
      ).join('')}
      ${allReasons[0] ? `<div class="rpt-flag err">🔎 ${allReasons[0].reason} (${allReasons[0].dept}) is the #1 facility driver — ${allReasons[0].count} lenses (${allReasons[0].pctLab}% of all breakage)</div>` : ''}
      ${totalUnassigned > 0 ? `<div class="rpt-flag err">⚠ ${totalUnassigned} unassigned lenses — operator scan compliance requires immediate supervisor attention</div>` : ''}
      ${allOps[0] ? `<div class="rpt-flag warn">👤 Highest contributor: ${allOps[0].name} (${allOps[0].dept}) — ${allOps[0].lenses} lenses, top reason: ${allOps[0].topReason}</div>` : ''}
      ${overGoal.length===0 ? `<div class="rpt-flag ok">✓ All departments within goal — strong operational day</div>` : ''}
      ${totalUnassigned===0 ? `<div class="rpt-flag ok">✓ All breakage records have full operator attribution</div>` : ''}

      <div class="rpt-footer">
        <span>Generated ${new Date().toLocaleString()} · Zenni Facility Breakage Dashboard</span>
        <button onclick="document.getElementById('sumReportPanel').style.display='none';document.getElementById('sumReportBtn').textContent='📄 Generate Report'"
          style="font-family:var(--font-mono);font-size:11px;padding:5px 14px;background:var(--bg3);border:1px solid var(--border2);border-radius:6px;color:var(--muted);cursor:pointer">✕ Close</button>
      </div>`;

  } else {
    // ── DEPARTMENT-SPECIFIC REPORT ──────────────────────────
    const depKey  = deptKey.toLowerCase();
    const depData = d[depKey] || {};
    const depInfo = d.summary?.departments?.find(x => x.department === deptKey) || {};
    const color   = DEPT_COLORS[deptKey] || 'var(--cyan)';
    const goal    = DEPT_GOALS[deptKey] || null;
    const pct     = (depInfo.lensBrkPct || 0) * 100;
    const ok      = !goal || pct <= goal;
    const isSurface = deptKey === 'Surface';
    const cLabel  = colLabel(deptKey);

    const reasons  = [...(depData.reasons || [])].sort((a,b) => (b.lensesBroken||0)-(a.lensesBroken||0));
    const operators= [...(depData.operators || [])];
    const namedOps = operators.filter(o => o.operator && o.operator !== 'NONE');
    const noneOps  = operators.filter(o => !o.operator || o.operator === 'NONE');
    const noneCount= noneOps.reduce((s,o) => s+(o.total||(o.lensTotal||0)), 0);
    const totalOps = operators.length;
    const maxR     = reasons[0] ? (reasons[0].lensesBroken||1) : 1;
    const topOp    = namedOps[0];
    const framesBroken = depData.totals?.framesBroken || depInfo.framesBroken || 0;
    const framePctDept = orderCount > 0 ? (framesBroken/orderCount*100) : 0;

    // Concentration analysis — top reason % of dept
    const topReasonPct = reasons[0] && depInfo.lensesBroken > 0
      ? (((reasons[0].lensesBroken||0)/depInfo.lensesBroken)*100).toFixed(0) : 0;

    // Operator spread — how many ops contributed
    const opSpread = namedOps.filter(o=>(o.total||(o.lensTotal||0))>0).length;

    panel.style.setProperty('--report-color', color);
    panel.innerHTML = `
      <div class="rpt-h1" style="color:${color}">${deptKey.toUpperCase()} DEPARTMENT REPORT</div>
      <div class="rpt-meta">${date} &nbsp;·&nbsp; ${fmt(lensCount)} facility lenses &nbsp;·&nbsp; ${fmt(depInfo.lensesBroken||0)} dept lenses broken &nbsp;·&nbsp; Generated ${new Date().toLocaleTimeString()}</div>

      <!-- 1. Dept Overview -->
      <div class="rpt-h2">1. Department Overview</div>
      <p class="rpt-p">${deptKey} department recorded <strong style="color:${color}">${fmt(depInfo.lensesBroken||0)} lenses</strong> broken on ${date}, representing a <strong style="${sc(ok)}">${pct.toFixed(2)}%</strong> breakage rate against ${fmt(lensCount)} facility lenses.
      ${goal ? 'The department is <strong style="'+sc(ok)+'">'+( ok?'within':'exceeding')+' its '+goal+'% goal</strong>'+(ok?', with <strong>'+(goal-pct).toFixed(2)+'pp of margin remaining</strong>.':' by <strong>'+(pct-goal).toFixed(2)+'pp</strong>.') : 'No breakage goal is set for this department.'}
      ${framesBroken > 0 ? 'Additionally, <strong style="color:var(--amber)">'+framesBroken+' frames</strong> were broken ('+framePctDept.toFixed(2)+'% frame breakage rate).' : ''}</p>

      <!-- 2. Goal & Position -->
      ${goal ? `<div class="rpt-h2">2. Goal Analysis</div>
      <div class="rpt-flag ${ok?'ok':'err'}">${ok?'✓':'✗'} ${deptKey} at ${pct.toFixed(2)}% vs goal ≤${goal}% — ${ok?'On Goal · '+( goal-pct).toFixed(2)+'pp remaining':'Over Goal · '+(pct-goal).toFixed(2)+'pp above target'}</div>` : ''}

      <!-- 3. Reason Breakdown -->
      <div class="rpt-h2">${goal?'3':'2'}. Breakage Reason Analysis</div>
      ${reasons.length === 0 ? '<p class="rpt-p" style="color:var(--muted)">No reason data available.</p>' : `
      <p class="rpt-p">
        ${reasons.length} distinct reason${reasons.length>1?'s were':' was'} recorded.
        The primary driver is <strong style="color:${color}">${reasons[0].reason}</strong> with ${reasons[0].lensesBroken||0} lenses, accounting for <strong>${topReasonPct}%</strong> of all ${deptKey} breakage.
        ${reasons.length > 1 ? 'The top 3 reasons account for <strong>'+(reasons.slice(0,3).reduce((s,r)=>s+(r.lensesBroken||0),0))+' lenses</strong> ('+(depInfo.lensesBroken>0?(reasons.slice(0,3).reduce((s,r)=>s+(r.lensesBroken||0),0)/depInfo.lensesBroken*100).toFixed(0):0)+'% of dept breakage).' : ''}
      </p>
      <table class="rpt-table">
        <thead><tr><th>#</th><th>Reason</th><th>Lenses</th><th>% of Dept</th><th>% of Lab</th></tr></thead>
        <tbody>
          ${reasons.map((r,i) => {
            const cnt = r.lensesBroken||r.framesBroken||0;
            return `<tr>
              <td style="font-family:var(--font-mono);color:var(--muted)">${i+1}</td>
              <td style="font-weight:${i===0?'600':'400'}">${r.reason}</td>
              <td style="font-family:var(--font-display);font-weight:700;color:${color}">${cnt}</td>
              <td style="font-family:var(--font-mono);font-size:11px">${depInfo.lensesBroken>0?(cnt/depInfo.lensesBroken*100).toFixed(1):0}%</td>
              <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${(cnt/lensCount*100).toFixed(2)}%</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>`}

      <!-- 4. Operator/Machine Breakdown -->
      <div class="rpt-h2">${goal?'4':'3'}. ${cLabel} Breakdown</div>
      ${namedOps.length === 0 ? '<p class="rpt-p" style="color:var(--muted)">No '+cLabel.toLowerCase()+' data available.</p>' : `
      <p class="rpt-p">
        ${opSpread} ${cLabel.toLowerCase()}${opSpread>1?'s':''} contributed to ${deptKey} breakage.
        ${topOp ? 'The highest contributor is <strong style="color:'+color+'">'+(topOp.operator)+'</strong> with <strong>'+(topOp.total||(topOp.lensTotal||0))+'</strong> lenses'+(topOp.reasons&&topOp.reasons[0]?', primarily from '+topOp.reasons[0].reason:'')+'.' : ''}
      </p>
      <table class="rpt-table">
        <thead><tr><th>${cLabel}</th><th>Lenses</th>${deptKey==='Finish'?'<th>Frames</th>':''}<th>% of Dept</th><th>Top Reason</th><th>Reasons</th></tr></thead>
        <tbody>
          ${namedOps.map(op => {
            const total   = op.total ?? (op.lensTotal||0);
            const topR    = (op.reasons||[])[0]?.reason || '—';
            const rCount  = (op.reasons||[]).length;
            return `<tr>
              <td style="font-weight:500;color:${color}">${op.operator}</td>
              <td style="font-family:var(--font-display);font-weight:700;color:${color}">${total}</td>
              ${deptKey==='Finish'?'<td style="font-family:var(--font-mono);font-size:11px;color:'+(op.frameTotal>0?'var(--amber)':'var(--muted)')+'">'+((op.frameTotal||0)>0?op.frameTotal:'—')+'</td>':''}
              <td style="font-family:var(--font-mono);font-size:11px">${depInfo.lensesBroken>0?(total/depInfo.lensesBroken*100).toFixed(1):0}%</td>
              <td style="font-size:11px;color:var(--muted)">${topR}</td>
              <td style="font-family:var(--font-mono);font-size:11px;color:var(--muted)">${rCount}</td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>`}

      <!-- 5. Unassigned -->
      ${noneCount > 0 ? `<div class="rpt-h2">${goal?'5':'4'}. Data Integrity — Unassigned Records</div>
      <div class="rpt-flag err">⚠ ${noneCount} lenses recorded without a valid ${cLabel.toLowerCase()} (NONE) — ${depInfo.lensesBroken>0?(noneCount/depInfo.lensesBroken*100).toFixed(0):0}% of dept breakage lacks attribution</div>
      <p class="rpt-p" style="color:var(--red)">The unassigned records were associated with: ${noneOps.flatMap(o=>o.reasons||[]).sort((a,b)=>(b.lenses||0)-(a.lenses||0)).slice(0,5).map(r=>r.reason||'—').join(', ')}. Supervisor review and corrective scanning protocol recommended.</p>` : ''}

      <!-- 6. Key Observations -->
      <div class="rpt-h2">${goal ? (noneCount>0?'6':'5') : (noneCount>0?'5':'4')}. Key Observations & Recommendations</div>
      ${!ok ? `<div class="rpt-flag err">✗ ${deptKey} is ${(pct-goal).toFixed(2)}pp over the ${goal}% goal — escalate to department supervisor</div>` : ''}
      ${ok && goal && (goal-pct) < 0.30 ? `<div class="rpt-flag warn">⚡ ${deptKey} is within ${(goal-pct).toFixed(2)}pp of goal — elevated monitoring recommended</div>` : ''}
      ${reasons[0] ? `<div class="rpt-flag ${!ok?'err':'warn'}">🔎 ${reasons[0].reason} is the primary driver at ${reasons[0].lensesBroken||0} lenses (${topReasonPct}% of dept) — focus root cause review here</div>` : ''}
      ${topOp ? `<div class="rpt-flag warn">👤 ${topOp.operator} is the highest contributing ${cLabel.toLowerCase()} with ${topOp.total||(topOp.lensTotal||0)} lenses — review for targeted coaching or machine calibration</div>` : ''}
      ${noneCount > 0 ? `<div class="rpt-flag err">⚠ ${noneCount} unassigned lenses (${depInfo.lensesBroken>0?(noneCount/depInfo.lensesBroken*100).toFixed(0):0}%) — operator scan compliance action required</div>` : ''}
      ${ok && noneCount===0 ? `<div class="rpt-flag ok">✓ ${deptKey} is within goal with full operator attribution — no immediate action required</div>` : ''}

      <div class="rpt-footer">
        <span>Generated ${new Date().toLocaleString()} · Zenni Facility Breakage Dashboard</span>
        <button onclick="document.getElementById('sumReportPanel').style.display='none';document.getElementById('sumReportBtn').textContent='📄 Generate Report'"
          style="font-family:var(--font-mono);font-size:11px;padding:5px 14px;background:var(--bg3);border:1px solid var(--border2);border-radius:6px;color:var(--muted);cursor:pointer">✕ Close</button>
      </div>`;
  }

  panel.style.display = 'block';
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="13" height="13"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> Generate Report';
}

/* ═══════════════════════════════════════════════════════════
   MAIN APP CONTROLLER
═══════════════════════════════════════════════════════════ */
const App = {

  // Current mode tracking
  _isHistoryMode: false,
  _currentDate: null,
  _currentWeekStart: null,   // null = current week

  // ── Set Live mode UI ──
  setLiveMode() {
    App._isHistoryMode = false;
    const badge = document.getElementById('modeBadge');
    const label = document.getElementById('modeLabel');
    const banner = document.getElementById('historyBanner');
    const picker = document.getElementById('datePicker');
    if (badge)  { badge.classList.remove('history-mode'); }
    if (label)  { label.textContent = 'LIVE'; }
    if (banner) { banner.classList.remove('show'); }
    if (picker) { picker.classList.remove('history'); picker.value = ''; }
    document.querySelector('.mode-dot')?.classList.add('live-dot-anim');
  },

  // ── Set History mode UI ──
  setHistoryMode(dateStr) {
    App._isHistoryMode = true;
    App._currentDate   = dateStr;
    const badge  = document.getElementById('modeBadge');
    const label  = document.getElementById('modeLabel');
    const banner = document.getElementById('historyBanner');
    const bannerDate = document.getElementById('historyBannerDate');
    const picker = document.getElementById('datePicker');
    if (badge)      { badge.classList.add('history-mode'); }
    if (label)      { label.textContent = 'HISTORY'; }
    if (banner)     { banner.classList.add('show'); }
    if (bannerDate) { bannerDate.textContent = dateStr; }
    if (picker)     { picker.classList.add('history'); picker.value = dateStr; }
    document.querySelector('.mode-dot')?.classList.remove('live-dot-anim');
  },

  // ── Load today (live mode) ──
  async loadToday() {
    App.setLiveMode();
    showTransition('live', 'Switching to live data…');
    await App.loadAll();
  },

  // ── Week navigation ──
  async weekNav(direction) {
    // direction: -1=prev, 0=this week, 1=next
    let current = App._currentWeekStart
      ? new Date(App._currentWeekStart)
      : (() => { const d = new Date(); d.setDate(d.getDate() - d.getDay()); return d; })();

    if (direction === 0) {
      App._currentWeekStart = null;
    } else {
      current.setDate(current.getDate() + direction * 7);
      const mm = String(current.getMonth()+1).padStart(2,'0');
      const dd = String(current.getDate()).padStart(2,'0');
      const yy = current.getFullYear();
      App._currentWeekStart = `${mm}/${dd}/${yy}`;
    }
    await App.loadWeekly();
  },

  // ── Load weekly data ──
  async loadWeekly() {
    try {
      const qs  = App._currentWeekStart ? `&weekStart=${encodeURIComponent(App._currentWeekStart)}` : '';
      const res = await fetch(`${API_URL}?action=weekly${qs}`);
      const json = await res.json();
      if (!json.success) throw new Error(json.error);
      renderWeekly(json.data);
    } catch(err) {
      App.showToast('Weekly error: ' + err.message, 'error');
    }
  },

  // ── Load by selected date ──
  async loadByDate(dateStr) {
    if (!dateStr) { await App.loadToday(); return; }

    // Check if selected date is today
    const today = new Date().toISOString().split('T')[0];
    if (dateStr === today) { await App.loadToday(); return; }

    // Load from Snapshot_History
    App.setHistoryMode(dateStr);
    await App.loadFromHistory(dateStr);
  },

  // ── Load from Snapshot_History for a specific date ──
  async loadFromHistory(dateStr) {
    const btn = document.getElementById('refreshBtn');
    btn.classList.add('loading');
    showTransition('history', 'Loading ' + dateStr + '…');
    try {
      const histRes = await API.all(dateStr);
      const d       = histRes.data;

      if (!d.lensCount || d.lensCount === 0) {
        App.showToast('No snapshot found for ' + dateStr, 'error');
        App.setLiveMode();
        btn.classList.remove('loading');
        return;
      }

      const meta = {
        lensCount:   d.lensCount,
        orderCount:  d.orderCount,
        reportDate:  d.reportDate || dateStr,
        lastRefresh: d.reportDate || dateStr,
      };

      State.meta    = meta;
      State.summary = d.summary;
      State.depts   = d;

      renderMeta(meta);
      renderSummary(d.summary, meta, []);
      renderDeptTab('ar',        d.ar,        'AR',        '#a78bfa', 'arChart');
      renderDeptTab('finish',    d.finish,    'Finish',    '#34d399', 'finChart');
      renderDeptTab('surface',   d.surface,   'Surface',   '#fb923c', 'srfChart');
      renderDeptTab('lms',       d.lms,       'LMS',       '#60a5fa', 'lmsChart');
      renderDeptTab('inventory', d.inventory, 'Inventory', '#f472b6', null);
      renderDeptTab('breakage',  d.breakage,  'Breakage',  '#ef4444', null);
      renderMailroom(meta);

      document.getElementById('loadingOverlay').classList.add('hidden');
      hideTransition();
      App.showToast('History loaded: ' + (d.reportDate || dateStr), 'success');

    } catch (err) {
      console.error(err);
      hideTransition();
      App.showToast('History error: ' + err.message, 'error');
      App.setLiveMode();
    }
    btn.classList.remove('loading');
  },

  async loadAll() {
    const btn = document.getElementById('refreshBtn');
    btn.classList.add('loading');
    showTransition('live', 'Reading RawData…');
    try {
      const res = await API.all(); // live — no date param
      const d   = res.data;

      // Empty RawData check
      if (!d.lensCount || d.lensCount === 0) {
        renderNoData(d);
        document.getElementById('loadingOverlay').classList.add('hidden');
        App.showToast('No live data — RawData is empty', 'error');
        btn.classList.remove('loading');
        return;
      }

      const meta = {
        lensCount:   d.lensCount,
        orderCount:  d.orderCount,
        reportDate:  d.reportDate,
        lastRefresh: d.lastRefresh || new Date().toLocaleTimeString(),
      };

      State.meta    = meta;
      State.summary = d.summary;
      State.depts   = d;
      State.history = [];

      App.setLiveMode();
      renderMeta(meta);
      renderSummary(d.summary, meta, []);
      renderDeptTab('ar',        d.ar,        'AR',        '#a78bfa', 'arChart');
      renderDeptTab('finish',    d.finish,    'Finish',    '#34d399', 'finChart');
      renderDeptTab('surface',   d.surface,   'Surface',   '#fb923c', 'srfChart');
      renderDeptTab('lms',       d.lms,       'LMS',       '#60a5fa', 'lmsChart');
      renderDeptTab('inventory', d.inventory, 'Inventory', '#f472b6', null);
      renderDeptTab('breakage',  d.breakage,  'Breakage',  '#ef4444', null);
      renderMailroom(meta);
      initSummaryTab(); // pre-load summary dates

      document.getElementById('loadingOverlay').classList.add('hidden');
      hideTransition();
      App.showToast('Live data loaded', 'success');

    } catch (err) {
      console.error(err);
      hideTransition();
      App.showToast('Error: ' + err.message, 'error');
      document.getElementById('loadingOverlay').classList.add('hidden');
    }
    btn.classList.remove('loading');
  },

  switchTab(id) {
    const tab = document.querySelector(`.nav-tab[data-tab="${id}"]`);
    if (tab) {
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active', 'collapsing'));
      document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active', 'collapsed'));
      document.getElementById('tab-' + id)?.classList.add('active');
      tab.classList.add('active');
      // Init summary tab when switched to
      if (id === 'summary' && SUM.dates.length === 0) initSummaryTab();
      if (id === 'compare' && CMP.dates.length === 0) initCompareTab();
    }
  },

  async triggerSnapshot() {
    const name = prompt('Snapshot taken by (your name):', 'Manager');
    if (!name || !name.trim()) return;
    try {
      await API.snapshot(name.trim());
      App.showToast(`Snapshot saved by ${name.trim()}`, 'success');
      // Refresh history

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

/* ── NAV TAB CLICK WIRING — collapse toggle ─────────────── */
document.querySelectorAll('.nav-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    const id          = btn.dataset.tab;
    const panel       = document.getElementById('tab-' + id);
    const isActive    = btn.classList.contains('active');
    const isCollapsed = btn.classList.contains('collapsed');

    if (isActive && !isCollapsed) {
      // Currently open — collapse it
      panel.classList.add('collapsing');
      setTimeout(() => {
        panel.classList.remove('active', 'collapsing');
        btn.classList.add('collapsed');
      }, 240);
      return;
    }

    if (isActive && isCollapsed) {
      // Currently collapsed — re-expand
      btn.classList.remove('collapsed');
      panel.classList.add('active');
      return;
    }

    // Switch to a new tab — close everything first
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active', 'collapsing'));
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active', 'collapsed'));
    panel?.classList.add('active');
    btn.classList.add('active');
  });
});

/* ── BOOT ───────────────────────────────────────────────── */
App.loadAll();