/* ============================================================
   Surface Flow — Power Breakage v5
   power.js

   COLOR MAP (mirrors CSS :root):
   CLR.peach  #d28c64  S-Power · Source→AR41 · event tag · KPI1
   CLR.peri   #e8a040  S-Axis  · AR41→Brk   · Rx(−)     · KPI2
   CLR.teal   #d4c0a8  machines· Brk→Proc   · Rx(+)     · KPI3
   CLR.series [...]    material/chart multi-series
============================================================ */

const API = 'https://script.google.com/macros/s/AKfycbzpPE9mQznnF6oaIYW897SgqRVpsiMqZsTYj9HyMw0JADU31-6jN7slD_wJQs6f2njn/exec';

const CLR = {
  peach:  '#d28c64',
  peri:   '#e8a040',
  teal:   '#d4c0a8',
  series: ['#d28c64','#e8a040','#d4c0a8','#c8967a','#f0b870','#e8d0b8','#b87850','#d4a060'],
  gc: 'rgba(255,255,255,0.05)',
  tc: 'rgba(255,255,255,0.65)',
};

const charts = {};

// Active working sets (may be date-filtered)
let _brk=[], _reason=[], _research=[], _anom=[];

// Full loaded sets (always the complete live server response)
let _allBrk=[], _allReason=[], _allResearch=[], _allAnom=[];

/* ============================================================
   FETCH HELPER
   Passes optional startDate / endDate (YYYY-MM-DD strings) to
   Apps Script. When supplied, script reads BREAKAGE_HISTORY
   instead of live sheets.
============================================================ */
function fetchTab(tab, startDate, endDate) {
  let url = `${API}?tab=${tab}`;
  if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
  if (endDate)   url += `&endDate=${encodeURIComponent(endDate)}`;
  return fetch(url)
    .then(r => r.json())
    .then(d => {
      if (d.status !== 'ok') throw new Error(d.message);
      return d.data;
    });
}

/* ============================================================
   FORMAT HELPERS
============================================================ */
function fmtMin(m) {
  m = parseFloat(m);
  if (isNaN(m) || m <= 0) return '--';
  if (m < 60) return `${Math.round(m)}m`;
  const d  = Math.floor(m / 1440);
  const h  = Math.floor((m % 1440) / 60);
  const mn = Math.round(m % 60);
  return [d > 0 ? `${d}d` : '', h > 0 ? `${h}h` : '', mn > 0 ? `${mn}m` : ''].filter(Boolean).join(' ');
}

function fmtDate(s) {
  if (!s) return '--';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return `${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${String(d.getMinutes()).padStart(2,'0')}`;
}

function diffMin(a, b) {
  const da = new Date(a), db = new Date(b);
  if (isNaN(da) || isNaN(db)) return 0;
  return Math.max(0, (db - da) / 60000);
}

function avgArr(arr) {
  const n = arr.map(v => parseFloat(v)).filter(v => !isNaN(v) && v !== null);
  return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0;
}

function reasonColor(r) {
  if (!r) return CLR.teal;
  const s = r.toLowerCase();
  if (s.includes('power')) return CLR.peach;
  if (s.includes('axis'))  return CLR.peri;
  return CLR.teal;
}

function reasonPillClass(r) {
  if (!r) return 'pill-other';
  const s = r.toLowerCase();
  if (s.includes('power')) return 'pill-peach';
  if (s.includes('axis'))  return 'pill-peri';
  return 'pill-other';
}

/* ============================================================
   CHART HELPER
============================================================ */
function mkChart(id, type, labels, datasets, extraOpts = {}) {
  if (charts[id]) charts[id].destroy();
  function merge(t, s) {
    const o = Object.assign({}, t);
    for (const k in s) o[k] = (s[k] && typeof s[k] === 'object' && !Array.isArray(s[k])) ? merge(t[k] || {}, s[k]) : s[k];
    return o;
  }
  const base = {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { color: CLR.gc }, ticks: { color: CLR.tc, font: { size: 11, family: "'IBM Plex Mono'" } } },
      y: { grid: { color: CLR.gc }, ticks: { color: CLR.tc, font: { size: 11, family: "'IBM Plex Mono'" } } },
    },
  };
  charts[id] = new Chart(document.getElementById(id), { type, data: { labels, datasets }, options: merge(base, extraOpts) });
}

/* ============================================================
   NAVIGATION
============================================================ */
function sw(id, btn) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('t-' + id).classList.add('active');
  btn.classList.add('active');
}

function swInner(id, btn) {
  document.querySelectorAll('.inner-section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.inner-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('is-' + id).classList.add('active');
  btn.classList.add('active');
  setTimeout(() => { if (charts[id + 'C']) charts[id + 'C'].resize(); }, 50);
}

/* ============================================================
   BAR LIST BUILDER
============================================================ */
function buildBarList(elId, entries) {
  const el = document.getElementById(elId);
  if (!el) return;
  const max = entries[0]?.[1] || 1;
  el.innerHTML = entries.map(([label, val, display, color]) => `
    <div class="bar-row">
      <div class="bar-label">${label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.round(val / max * 100)}%;background:${color || CLR.teal}"></div></div>
      <div class="bar-val">${display || val}</div>
    </div>`).join('');
}

/* ============================================================
   OVERVIEW
============================================================ */
function buildOverview() {
  const totalJobs = _brk.length;
  const totalLens = _brk.reduce((s, r) => s + (parseInt(r.LensesBroken) || 0), 0);
  const ar41Avg   = avgArr(_brk.map(r => r.AR41_to_Breakage_Min));
  const brkAvg    = avgArr(_brk.map(r => r.Breakage_to_Processed_Min));

  document.getElementById('m-jobs').textContent     = totalJobs;
  document.getElementById('m-jobs-sub').textContent  = `${totalLens} lenses total`;
  document.getElementById('m-lens').textContent     = totalLens;
  document.getElementById('m-lens-sub').textContent  = `across ${totalJobs} jobs`;
  document.getElementById('m-ar41brk').textContent  = fmtMin(ar41Avg);
  document.getElementById('m-brkpro').textContent   = fmtMin(brkAvg);

  // Banner
  const worst = [..._brk].sort((a, b) => (parseFloat(b.AR41_to_Breakage_Min) || 0) - (parseFloat(a.AR41_to_Breakage_Min) || 0))[0];
  if (worst) document.getElementById('bannerSub').textContent =
    `${worst.BrkSourceMachine || worst.AR41Machine} — avg AR41→Brk ${fmtMin(ar41Avg)} · ${totalJobs} jobs`;
  const now = new Date();
  document.getElementById('bannerTime').textContent =
    `${now.getMonth()+1}/${now.getDate()} ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')}`;

  // Reason timing
  const rValid  = _reason.filter(r => r.JobCount > 0);
  const maxAR41 = Math.max(...rValid.map(r => parseFloat(r.AvgAR41_to_Breakage_Min) || 0), 1);
  const maxBrk  = Math.max(...rValid.map(r => parseFloat(r.AvgBreakage_to_Processed_Min) || 0), 1);
  const timingEl = document.getElementById('reasonTimingList');
  if (timingEl) {
    timingEl.innerHTML = `<div class="reason-timing">${rValid.map(r => {
      const ar41v = parseFloat(r.AvgAR41_to_Breakage_Min) || 0;
      const brkv  = parseFloat(r.AvgBreakage_to_Processed_Min) || 0;
      const color = reasonColor(r.BrkReason);
      return `
        <div class="rt-block">
          <div class="rt-name" style="color:${color}">${r.BrkReason}
            <span class="rt-jobs">${r.JobCount} jobs · ${r.TotalLensesBroken || '--'} lenses</span>
          </div>
          <div class="rt-row">
            <div class="rt-row-label">AR41 → Brk</div>
            <div class="rt-track"><div class="rt-fill" style="width:${Math.round(ar41v / maxAR41 * 100)}%;background:${color}"></div></div>
            <div class="rt-val">${fmtMin(ar41v)}</div>
          </div>
          <div class="rt-row">
            <div class="rt-row-label">Brk → Proc</div>
            <div class="rt-track"><div class="rt-fill" style="width:${Math.round(brkv / maxBrk * 100)}%;background:${color};opacity:0.6"></div></div>
            <div class="rt-val">${fmtMin(brkv)}</div>
          </div>
          ${rValid.indexOf(r) < rValid.length - 1 ? '<div class="rt-divider"></div>' : ''}
        </div>`;
    }).join('')}</div>`;
  }

  // BrkSource machine bars
  const srcAgg = {}, srcRx = {};
  _brk.forEach(r => {
    const s = r.BrkSourceMachine || 'Unknown';
    srcAgg[s] = (srcAgg[s] || 0) + 1;
    if (!srcRx[s]) srcRx[s] = [];
    srcRx[s].push(r);
  });
  const srcE  = Object.entries(srcAgg).sort((a, b) => b[1] - a[1]);
  const srcEl = document.getElementById('srcBarList');
  if (srcEl) {
    const max = srcE[0]?.[1] || 1;
    srcEl.innerHTML = srcE.map(([label, val]) => {
      const jobs = srcRx[label] || [];
      const rc   = {};
      const tl   = jobs.reduce((s, j) => s + (parseInt(j.LensesBroken) || 0), 0);
      jobs.forEach(j => { const r = j.BrkReason || '?'; rc[r] = (rc[r] || 0) + 1; });
      const tipLines = [...Object.entries(rc).map(([r, c]) => `${r}: ${c}`), `Lenses: ${tl}`].join(' · ');
      return `<div class="bar-row" title="${tipLines}">
        <div class="bar-label">${label}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.round(val / max * 100)}%;background:${CLR.teal}"></div></div>
        <div class="bar-val">${val}</div>
      </div>`;
    }).join('');
  }

  // Material bars
  const matAgg = {};
  _brk.forEach(r => { const m = r.Material || 'Unknown'; matAgg[m] = (matAgg[m] || 0) + (parseInt(r.LensesBroken) || 0); });
  const matE = Object.entries(matAgg).sort((a, b) => b[1] - a[1]);
  buildBarList('matList', matE.map(([label, val], i) => [label, val, val, CLR.series[i % CLR.series.length]]));

  // Reason donut
  const rsnAgg = {};
  _brk.forEach(r => { const k = r.BrkReason || 'Unknown'; rsnAgg[k] = (rsnAgg[k] || 0) + 1; });
  const rsnE = Object.entries(rsnAgg).sort((a, b) => b[1] - a[1]);
  document.getElementById('reasonLeg').innerHTML = rsnE.map(([r, c]) =>
    `<span class="leg-item"><span class="leg-sq" style="background:${reasonColor(r)}"></span>${r} — ${c}</span>`
  ).join('');
  mkChart('reasonPieC', 'doughnut', rsnE.map(e => e[0]), [{
    data: rsnE.map(e => e[1]),
    backgroundColor: rsnE.map(e => reasonColor(e[0])),
    borderWidth: 2, borderColor: '#0d1015',
  }], { plugins: { legend: { display: false } }, scales: { x: { display: false }, y: { display: false } } });
}

/* ============================================================
   FLOW FILTERS
============================================================ */
function populateFlowFilters() {
  [
    ['filterReason', _brk.map(r => r.BrkReason)],
    ['filterSource', _brk.map(r => r.BrkSourceMachine)],
    ['filterAR41',   _brk.map(r => r.AR41Machine)],
  ].forEach(([id, vals]) => {
    const el = document.getElementById(id);
    while (el.options.length > 1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option'); o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}

/* ── Tooltip ── */
const tip = document.getElementById('ganttTip');
function showTip(e, html) { if(!tip) return; tip.innerHTML = html; tip.style.display = 'block'; moveTip(e); }
function moveTip(e) { if(!tip) return; tip.style.left = (e.clientX + 14) + 'px'; tip.style.top = (e.clientY + 14) + 'px'; }
function hideTip() { if(!tip) return; tip.style.display = 'none'; }
document.addEventListener('mousemove', e => { if (tip && tip.style.display === 'block') moveTip(e); });

/* ============================================================
   RENDER FLOW
============================================================ */
function renderFlow() {
  const fR   = document.getElementById('filterReason').value;
  const fS   = document.getElementById('filterSource').value;
  const fA   = document.getElementById('filterAR41').value;
  const from = parseDateInput('flowDateFrom');
  const to   = parseDateInput('flowDateTo');

  let data = [..._brk];
  if (fR) data = data.filter(r => r.BrkReason === fR);
  if (fS) data = data.filter(r => r.BrkSourceMachine === fS);
  if (fA) data = data.filter(r => r.AR41Machine === fA);
  if (from || to) data = data.filter(r => inDateRange(r.AR41ScanTime || r.BrkTableScanTime, from, to));

  document.getElementById('flowCount').textContent = data.length;
  if (!data.length) {
    document.getElementById('flowList').innerHTML = '<div class="empty">No jobs match filters</div>';
    return;
  }

  const resMap = {};
  _research.forEach(r => { resMap[String(r.RxNumber)] = r; });

  const rows = data.map(r => {
    const res   = resMap[String(r.RxNumber)];
    const src   = (r.BrkSourceMachine || '').toUpperCase();
    let srcTime = null;
    if (res) srcTime = src.includes('ORB') ? res.ORBScanTime : res.OTBScanTime;
    const d1    = srcTime ? diffMin(srcTime, r.AR41ScanTime) : 0;
    const d2    = parseFloat(r.AR41_to_Breakage_Min) || 0;
    const d3    = parseFloat(r.Breakage_to_Processed_Min) || 0;
    const total = d1 + d2 + d3 || 1;
    const p1    = Math.max(d1 > 0 ? 1 : 0, Math.round(d1 / total * 100));
    const p2    = Math.max(1, Math.round(d2 / total * 100));
    const p3    = Math.max(1, 100 - p1 - p2);
    const t1    = `<b>Source → AR41</b><br>${r.BrkSourceMachine || '?'}: ${fmtDate(srcTime)}<br>AR41: ${fmtDate(r.AR41ScanTime)}<br>Duration: ${fmtMin(d1)}`;
    const t2    = `<b>AR41 → Breakage</b><br>AR41: ${fmtDate(r.AR41ScanTime)}<br>Brk Table: ${fmtDate(r.BrkTableScanTime)}<br>Duration: ${fmtMin(d2)}`;
    const t3    = `<b>Breakage → Processed</b><br>Brk Table: ${fmtDate(r.BrkTableScanTime)}<br>Processed: ${fmtDate(r.BreakageProcessedTime)}<br>Duration: ${fmtMin(d3)}`;
    return `
      <div class="gantt-row">
        <div class="gantt-rx" title="${r.RxNumber}">${r.RxNumber}</div>
        <div class="gantt-bar">
          ${d1 > 0 ? `<div class="gantt-seg seg-1" style="width:${p1}%;min-width:2px" data-tip="${t1}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p1 > 10 ? fmtMin(d1) : ''}</div>` : ''}
          <div class="gantt-seg seg-2" style="width:${p2}%;min-width:4px" data-tip="${t2}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p2 > 10 ? fmtMin(d2) : ''}</div>
          <div class="gantt-seg seg-3" style="width:${p3}%;min-width:4px" data-tip="${t3}" onmouseenter="showTip(event,this.dataset.tip)" onmouseleave="hideTip()">${p3 > 10 ? fmtMin(d3) : ''}</div>
        </div>
        <span class="reason-pill ${reasonPillClass(r.BrkReason)}">${r.BrkReason || '--'}</span>
        <div class="gantt-total">${fmtMin(total)}</div>
      </div>`;
  }).join('');
  document.getElementById('flowList').innerHTML = rows;
}

/* ============================================================
   RESEARCH FILTERS
============================================================ */
function populateResearchFilters() {
  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const orbVals = isHistorical
    ? _research.map(r => r.BrkSourceMachine || r.ORBMachine)
    : _research.map(r => r.ORBMachine);

  // Enrich Material from _brk so the Material dropdown is populated
  const brkMatMap = {};
  _brk.forEach(r => { if (r.RxNumber && r.Material && r.Material !== 'Unknown') brkMatMap[String(r.RxNumber)] = r.Material; });
  const matVals = _research.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r.Material;
    return brkMatMap[String(r.RxNumber)] || '';
  });

  [
    ['resORB',      orbVals],
    ['resReason',   _research.map(r => r.BrkReason)],
    ['resMaterial', matVals],
  ].forEach(([id, vals]) => {
    const el = document.getElementById(id);
    while (el.options.length > 1) el.remove(1);
    [...new Set(vals.filter(Boolean))].sort().forEach(v => {
      const o = document.createElement('option'); o.value = v; o.textContent = v; el.appendChild(o);
    });
  });
}

/* ── Sph / Cyl bucket helpers ── */
function sphBuckets(vals) {
  const b = { '≤-6': 0, '-6/-4': 0, '-4/-2': 0, '-2/0': 0, '0/+2': 0, '+2/+4': 0, '≥+4': 0 };
  vals.forEach(v => {
    if (isNaN(v)) return;
    if (v <= -6)      b['≤-6']++;
    else if (v <= -4) b['-6/-4']++;
    else if (v <= -2) b['-4/-2']++;
    else if (v < 0)   b['-2/0']++;
    else if (v < 2)   b['0/+2']++;
    else if (v < 4)   b['+2/+4']++;
    else              b['≥+4']++;
  });
  return b;
}

function cylBuckets(vals) {
  const b = { '≤-3': 0, '-3/-2': 0, '-2/-1': 0, '-1/0': 0, '0': 0, '0/+1': 0, '≥+1': 0 };
  vals.forEach(v => {
    if (isNaN(v)) return;
    if (v <= -3)      b['≤-3']++;
    else if (v <= -2) b['-3/-2']++;
    else if (v <= -1) b['-2/-1']++;
    else if (v < 0)   b['-1/0']++;
    else if (v === 0) b['0']++;
    else if (v < 1)   b['0/+1']++;
    else              b['≥+1']++;
  });
  return b;
}

function sphColors(keys) {
  return keys.map(k => (k.startsWith('+') || k === '0/+2' || k === '≥+4') ? '#d4c030' : '#4d94d4');
}

function cylColors(keys) {
  const neg = new Set(['≤-3', '-3/-2', '-2/-1', '-1/0']);
  const pos = new Set(['0/+1', '≥+1']);
  return keys.map(k => {
    if (k === '0')   return 'rgba(255,255,255,0.2)';
    if (pos.has(k))  return '#4d94d4';
    if (neg.has(k))  return '#d4c030';
    return 'rgba(255,255,255,0.2)';
  });
}

/* ============================================================
   RENDER RESEARCH
============================================================ */
function renderResearch() {
  const fO   = document.getElementById('resORB').value;
  const fR   = document.getElementById('resReason').value;
  const fM   = document.getElementById('resMaterial').value;

  // In historical mode _research === _brk so global date already applied.
  // Only apply the Research sub-date-filters in live mode.
  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const from = isHistorical ? null : parseDateInput('resDateFrom');
  const to   = isHistorical ? null : parseDateInput('resDateTo');

  let data = [..._research];
  if (fO) data = data.filter(r => (r.BrkSourceMachine || r.ORBMachine) === fO);
  if (fR) data = data.filter(r => r.BrkReason === fR);
  if (from || to) data = data.filter(r => inDateRange(r.ORBScanTime || r.OTBScanTime, from, to));

  // Build Material lookup from _brk — BreakageSummary always has Material populated
  const brkMaterialMap = {};
  _brk.forEach(r => {
    if (r.RxNumber && r.Material && r.Material !== 'Unknown')
      brkMaterialMap[String(r.RxNumber)] = r.Material;
  });

  // Enrich rows where Material is missing, then apply material filter
  data = data.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r;
    const mat = brkMaterialMap[String(r.RxNumber)];
    return mat ? { ...r, Material: mat } : r;
  });
  if (fM) data = data.filter(r => r.Material === fM);

  const totalLens = data.reduce((s, r) => s + (parseInt(r.LensesBroken) || 0), 0);

  document.getElementById('res-jobs').textContent  = data.length;
  document.getElementById('res-lens').textContent  = totalLens;

  // Sph/Cyl charts — show "no Rx data" note in historical mode
  const noRxData = isHistorical && data.every(r => !r.R_Sph && !r.L_Sph);

  ['rsphC','lsphC','rcylC','lcylC'].forEach(id => {
    if (charts[id]) { charts[id].destroy(); delete charts[id]; }
  });

  // Show/hide historical note and handle Sph/Cyl charts
  const rxNote = document.getElementById('rxHistoricalNote');
  if (rxNote) rxNote.style.display = noRxData ? 'flex' : 'none';

  if (noRxData) {
    // Destroy any existing charts and show empty placeholder message
    ['rsphC','lsphC','rcylC','lcylC'].forEach(id => {
      if (charts[id]) { charts[id].destroy(); delete charts[id]; }
      const wrap = document.querySelector(`[data-canvas="${id}"]`);
      if (wrap) wrap.innerHTML = `<div class="empty" style="height:100%;display:flex;align-items:center;justify-content:center;font-size:11px;opacity:0.4">No prescription data in history</div>`;
    });
  } else {
    // Restore canvas elements if they were replaced by the empty message
    ['rsphC','lsphC','rcylC','lcylC'].forEach(id => {
      const wrap = document.querySelector(`[data-canvas="${id}"]`);
      if (wrap && !document.getElementById(id)) {
        wrap.innerHTML = '';
        const c = document.createElement('canvas');
        c.id = id; wrap.appendChild(c);
      }
    });

    const rsph = sphBuckets(data.map(r => parseFloat(r.R_Sph)));
    mkChart('rsphC','bar',Object.keys(rsph),[{data:Object.values(rsph),backgroundColor:sphColors(Object.keys(rsph)),borderWidth:0,borderRadius:4}],{});

    const lsph = sphBuckets(data.map(r => parseFloat(r.L_Sph)));
    mkChart('lsphC','bar',Object.keys(lsph),[{data:Object.values(lsph),backgroundColor:sphColors(Object.keys(lsph)),borderWidth:0,borderRadius:4}],{});

    const rcyl = cylBuckets(data.map(r => parseFloat(r.R_Cyl)));
    mkChart('rcylC','bar',Object.keys(rcyl),[{data:Object.values(rcyl),backgroundColor:cylColors(Object.keys(rcyl)),borderWidth:0,borderRadius:4}],{});

    const lcyl = cylBuckets(data.map(r => parseFloat(r.L_Cyl)));
    mkChart('lcylC','bar',Object.keys(lcyl),[{data:Object.values(lcyl),backgroundColor:cylColors(Object.keys(lcyl)),borderWidth:0,borderRadius:4}],{});
  }

  function pp(v) {
    const n = parseFloat(v);
    if (isNaN(n) || v === '' || v === null || v === undefined) return `<span class="px-zero">—</span>`;
    if (n === 0)  return `<span class="px-zero">0</span>`;
    return n > 0 ? `<span class="px-pos">+${n}</span>` : `<span class="px-neg">${n}</span>`;
  }

  // In historical mode show AR41 machine instead of ORB (ORBMachine not in history)
  const orbCol = isHistorical ? 'AR41' : 'ORB';

  document.getElementById('resTable').innerHTML = `
    <table class="res-table">
      <thead><tr>
        <th>RX</th><th>${orbCol}</th><th>BrkSrc</th><th>Reason</th>
        <th>Material</th>
        ${isHistorical ? '' : '<th>Option</th><th>Curve</th><th>R Sph</th><th>L Sph</th><th>R Cyl</th><th>L Cyl</th><th>Add</th>'}
        <th>Broken</th><th>AR41→Brk</th><th>Brk→Proc</th>
      </tr></thead>
      <tbody>${data.map(r => `
        <tr>
          <td>${r.RxNumber || '--'}</td>
          <td>${isHistorical ? (r.AR41Machine || '--') : (r.ORBMachine || '--')}</td>
          <td>${r.BrkSourceMachine || '--'}</td>
          <td><span class="reason-pill ${reasonPillClass(r.BrkReason)}">${r.BrkReason || '--'}</span></td>
          <td>${r.Material || '--'}</td>
          ${isHistorical ? '' : `
          <td>${r.LensOption || '--'}</td><td>${r.BaseCurve || '--'}</td>
          <td>${pp(r.R_Sph)}</td><td>${pp(r.L_Sph)}</td>
          <td>${pp(r.R_Cyl)}</td><td>${pp(r.L_Cyl)}</td>
          <td>${r.R_AddPower || '—'}</td>`}
          <td>${r.LensesBroken || '1'}</td>
          <td style="color:var(--peri)">${fmtMin(r.AR41_to_Breakage_Min)}</td>
          <td style="color:var(--teal)">${fmtMin(r.Breakage_to_Processed_Min)}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

/* ============================================================
   ALERTS
============================================================ */
function buildAlerts() {
  const total  = _anom.length;
  const sPower = _anom.filter(r => (r.BrkReason || '').toLowerCase().includes('power')).length;
  const sAxis  = _anom.filter(r => (r.BrkReason || '').toLowerCase().includes('axis')).length;

  document.getElementById('al-total').textContent  = total;
  document.getElementById('al-spower').textContent = sPower;
  document.getElementById('al-saxis').textContent  = sAxis;

  const sorted = [..._anom]
    .map(r => ({ ...r, _v: parseFloat(r.AR41_to_Breakage_Min) || 0 }))
    .sort((a, b) => b._v - a._v);
  const maxV = sorted[0]?._v || 1;

  document.getElementById('alertList').innerHTML = sorted.map(r => {
    const color = reasonColor(r.BrkReason);
    const pct   = Math.max(2, Math.round(r._v / maxV * 100));
    return `
      <div class="alert-bar-row">
        <div>
          <div class="alert-bar-rx">${r.RxNumber}</div>
          <div class="alert-bar-desc">${r.AR41Machine || '--'} · ${r.BrkSourceMachine || '--'} · ${r.BrkReason || '--'}</div>
        </div>
        <div class="alert-bar-track"><div class="alert-bar-fill" style="width:${pct}%;background:${color}"></div></div>
        <div class="alert-bar-val" style="color:${color}">${fmtMin(r._v)}</div>
      </div>`;
  }).join('') || '<div class="empty">No alerts</div>';

  const worst   = sorted[0];
  const avgAR41 = avgArr(_anom.map(r => r.AR41_to_Breakage_Min));
  const avgBrk  = avgArr(_anom.map(r => r.Breakage_to_Processed_Min));

  document.getElementById('quickStats').innerHTML = [
    ['Total delays', total,                          ''],
    ['S-Power',      sPower,                         CLR.peach],
    ['S-Axis',       sAxis,                          CLR.peri],
    ['Avg AR41→Brk', fmtMin(avgAR41),                CLR.teal],
    ['Avg Brk→Proc', fmtMin(avgBrk),                 ''],
    ['Worst job',    worst ? fmtMin(worst._v) : '--', CLR.peach],
  ].map(([label, val, color]) => `
    <div class="stat-row">
      <span class="stat-label">${label}</span>
      <span class="stat-val" style="color:${color || 'inherit'}">${val}</span>
    </div>`).join('');
}

/* ============================================================
   POWER ANALYSIS SPLASH — OVERLAY DRIVERS
============================================================ */
const PA_STEPS = ['ls-brk','ls-reason','ls-research','ls-anom'];
const PA_LABELS = {
  'ls-brk':      'Loading breakage records…',
  'ls-reason':   'Aggregating reason data…',
  'ls-research': 'Pulling research prescriptions…',
  'ls-anom':     'Scanning for anomalies…',
};

function setLoadStep(id, state) {
  // state: 'active' | 'done' | ''
  PA_STEPS.forEach(s => {
    const el = document.getElementById(s);
    if (!el) return;
    el.classList.remove('active','done');
  });

  // Mark all steps before this one as done, this one as state
  let found = false;
  PA_STEPS.forEach(s => {
    const el = document.getElementById(s);
    if (!el) return;
    if (s === id) { found = true; if (state) el.classList.add(state); return; }
    if (!found) el.classList.add('done');
  });

  // Update progress bar
  const idx = PA_STEPS.indexOf(id);
  const pct = state === 'done'
    ? 100
    : state === 'active'
      ? Math.round((idx / PA_STEPS.length) * 100)
      : 0;
  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = pct + '%';

  // Update label
  const lbl = document.getElementById('paProgressLabel');
  if (lbl && state === 'active') lbl.textContent = PA_LABELS[id] || '';
  if (lbl && state === 'done')   lbl.textContent = 'Almost ready…';
}

function showOverlay() {
  const o = document.getElementById('loadingOverlay');
  if (o) o.classList.remove('hidden');
  // Reset all steps
  PA_STEPS.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active','done');
  });
  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = '0%';
  const lbl = document.getElementById('paProgressLabel');
  if (lbl) lbl.textContent = 'Connecting to data source…';
}

function hideOverlay() {
  const bar = document.getElementById('paProgressBar');
  if (bar) bar.style.width = '100%';
  const lbl = document.getElementById('paProgressLabel');
  if (lbl) lbl.textContent = 'Ready';
  // Mark all done
  PA_STEPS.forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.classList.remove('active'); el.classList.add('done'); }
  });
  setTimeout(() => {
    const o = document.getElementById('loadingOverlay');
    if (o) o.classList.add('hidden');
  }, 400);
}

/* ============================================================
   DATE FILTER HELPERS  (client-side date objects for sub-filters)
============================================================ */
function parseDateInput(id) {
  const v = document.getElementById(id)?.value;
  return v ? new Date(v + 'T00:00:00') : null;
}

function inDateRange(dateStr, from, to) {
  if (!from && !to) return true;
  const d = new Date(dateStr);
  if (isNaN(d)) return true;
  if (from && d < from) return false;
  if (to   && d > new Date(to.getTime() + 86400000 - 1)) return false;
  return true;
}

/* ============================================================
   FIELD NAME HELPERS
   Maps both live sheet field names (camelCase, from BreakageSummary/
   PowerResearch) AND BREAKAGE_HISTORY raw column names (with spaces,
   exactly as seen in the sheet headers).

   BREAKAGE_HISTORY columns confirmed from sheet:
   SnapshotDate | RX Number | OTB Blocker Scan Point | OTB Blocker Scan Time
   ORB Scan Point | ORB Scan Time | AR41 Scan Point | AR41 Scan Time
   Operator Scan | Scan to Breakage Table | Brk Scan Point | Brk Reason
   Breakage Processed Time | Lenses Broken | [Material etc from RawData]

   Live BreakageSummary columns:
   RxNumber | AR41Machine | AR41ScanTime | BrkTableScanTime
   BrkSourceMachine | BrkReason | BreakageProcessedTime | LensesBroken
   Material | AR41_to_Breakage_Min | Breakage_to_Processed_Min | Status

   Live PowerResearch columns:
   RxNumber | OTBMachine | OTBScanTime | ORBMachine | ORBScanTime
   AR41Machine | AR41ScanTime | BrkSourceMachine | BrkReason | Material
   LensOption | BaseCurve | R_Sph | L_Sph | R_Cyl | L_Cyl | R_AddPower ...
============================================================ */
function fv(row, ...keys) {
  for (const k of keys) {
    const v = row[k];
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return '';
}

// Calculate minutes between two date strings
function calcMin(fromStr, toStr) {
  if (!fromStr || !toStr) return 0;
  const a = new Date(fromStr), b = new Date(toStr);
  if (isNaN(a) || isNaN(b)) return 0;
  return Math.max(0, (b - a) / 60000);
}

function normalizeRow(r) {
  // ── Core time fields ──────────────────────────────────────────
  // Live sheet uses camelCase; BREAKAGE_HISTORY uses spaced names
  const ar41ScanTime          = fv(r, 'AR41ScanTime',          'AR41 Scan Time');
  const brkTableScanTime      = fv(r, 'BrkTableScanTime',      'Scan to Breakage Table');
  const breakageProcessedTime = fv(r, 'BreakageProcessedTime', 'Breakage Processed Time');
  const orbScanTime           = fv(r, 'ORBScanTime',           'ORB Scan Time');
  const otbScanTime           = fv(r, 'OTBScanTime',           'OTB Blocker Scan Time');

  // ── Timing mins — use pre-calculated if present, else compute ──
  const ar41ToBreakage = parseFloat(
    fv(r, 'AR41_to_Breakage_Min', 'AR41 to Breakage Min')
  ) || calcMin(ar41ScanTime, brkTableScanTime);

  const brkToProcessed = parseFloat(
    fv(r, 'Breakage_to_Processed_Min', 'Breakage to Processed Min')
  ) || calcMin(brkTableScanTime, breakageProcessedTime);

  return {
    // ── Identity ──────────────────────────────────────────────
    RxNumber:     fv(r, 'RxNumber', 'RX Number'),

    // ── Machines ──────────────────────────────────────────────
    // BREAKAGE_HISTORY: AR41 Scan Point = AR41Machine
    //                   Brk Scan Point  = BrkSourceMachine (the ORB that produced)
    //                   ORB Scan Point  = ORBMachine
    //                   OTB Blocker Scan Point = OTBMachine
    AR41Machine:      fv(r, 'AR41Machine',      'AR41 Scan Point'),
    BrkSourceMachine: fv(r, 'BrkSourceMachine', 'Brk Scan Point'),
    ORBMachine:       fv(r, 'ORBMachine',        'ORB Scan Point'),
    OTBMachine:       fv(r, 'OTBMachine',        'OTB Blocker Scan Point'),

    // ── Reason / Material ──────────────────────────────────────
    // Live sheets use 'Material'; BREAKAGE_HISTORY uses 'Lens Material'
    BrkReason:    fv(r, 'BrkReason',    'Brk Reason'),
    Material:     fv(r, 'Material', 'Lens Material', 'material') || 'Unknown',
    LensesBroken: parseFloat(fv(r, 'LensesBroken', 'Lenses Broken')) || 0,

    // ── Times ──────────────────────────────────────────────────
    AR41ScanTime:          ar41ScanTime,
    BrkTableScanTime:      brkTableScanTime,
    BreakageProcessedTime: breakageProcessedTime,
    ORBScanTime:           orbScanTime,
    OTBScanTime:           otbScanTime,

    // ── Calculated timing ─────────────────────────────────────
    AR41_to_Breakage_Min:      ar41ToBreakage,
    Breakage_to_Processed_Min: brkToProcessed,

    // ── Research / Rx fields ──────────────────────────────────
    // BREAKAGE_HISTORY: 'Lens Material'=Material, 'Lens Option'=LensOption, 'Base Curve'=BaseCurve
    LensOption: fv(r, 'LensOption', 'Lens Option'),
    BaseCurve:  fv(r, 'BaseCurve',  'Base Curve'),
    R_Sph:      fv(r, 'R_Sph', 'R Sph'),
    L_Sph:      fv(r, 'L_Sph', 'L Sph'),
    R_Cyl:      fv(r, 'R_Cyl', 'R Cyl'),
    L_Cyl:      fv(r, 'L_Cyl', 'L Cyl'),
    R_AddPower: fv(r, 'R_AddPower', 'R AddPower', 'R Add Power'),
  };
}

/* ============================================================
   APPLY DATE FILTER  (single date → sends same value as start+end)
   Re-fetches from Apps Script — reads BREAKAGE_HISTORY server-side
============================================================ */
async function applyDateFilter() {
  const dateVal = document.getElementById('dateSingle')?.value;  // "YYYY-MM-DD"
  if (!dateVal) return;

  showOverlay();
  document.getElementById('liveStatus').textContent   = 'Loading...';
  document.getElementById('liveDot').style.animation  = 'none';
  document.getElementById('liveDot').style.background = CLR.peri;

  // Show Historical badge
  const badge = document.getElementById('dateModeBadge');
  const label = document.getElementById('dateModeLabel');
  if (badge) badge.style.display = 'flex';
  if (label) label.textContent   = dateVal;

  try {
    setLoadStep('ls-brk', 'active');
    const brk = await fetchTab('breakageSummary', dateVal, dateVal);
    setLoadStep('ls-brk', 'done'); setLoadStep('ls-reason', 'active');

    const reason = await fetchTab('reasonSummary', dateVal, dateVal);
    setLoadStep('ls-reason', 'done'); setLoadStep('ls-research', 'active');

    const research = await fetchTab('powerResearch', dateVal, dateVal);
    setLoadStep('ls-research', 'done'); setLoadStep('ls-anom', 'active');

    const anom = await fetchTab('anomalies', dateVal, dateVal);
    setLoadStep('ls-anom', 'done');

    // Normalize all rows so field names are consistent regardless of source
    _brk      = (Array.isArray(brk)      ? brk      : []).map(normalizeRow);
    _reason   = Array.isArray(reason)   ? reason   : [];  // already aggregated by Apps Script
    _anom     = (Array.isArray(anom)     ? anom     : []).map(normalizeRow);

    // BREAKAGE_HISTORY has no Rx prescription fields (R_Sph, L_Sph, etc.)
    // For historical Research tab, reuse _brk rows — they have machine/reason/material
    // Rx charts will show empty (no prescription data in history) but table will work
    _research = _brk;

    document.getElementById('dateFilterCount').textContent = _brk.length;
    buildOverview();
    populateFlowFilters();     renderFlow();
    populateResearchFilters(); renderResearch();
    buildAlerts();
    buildWeekOptions();

    document.getElementById('liveStatus').textContent   = 'Historical';
    document.getElementById('liveDot').style.background = CLR.peri;
    document.getElementById('liveDot').style.animation  = 'none';

  } catch (err) {
    console.error('Date filter fetch error:', err);
    document.getElementById('liveStatus').textContent   = 'Error';
    document.getElementById('liveDot').style.background = '#f87171';
  } finally {
    hideOverlay();
  }
}

/* ============================================================
   CLEAR DATE FILTER — back to live
============================================================ */
async function clearDateFilter() {
  const inp = document.getElementById('dateSingle');
  if (inp) inp.value = '';
  const badge = document.getElementById('dateModeBadge');
  if (badge) badge.style.display = 'none';
  await loadAll();
}

/* ============================================================
   EXPORT HELPERS
============================================================ */
function getFilteredResearchData() {
  const fO   = document.getElementById('resORB').value;
  const fR   = document.getElementById('resReason').value;
  const fM   = document.getElementById('resMaterial').value;
  const isHistorical = (document.getElementById('dateSingle')?.value || '') !== '';
  const from = isHistorical ? null : parseDateInput('resDateFrom');
  const to   = isHistorical ? null : parseDateInput('resDateTo');

  let data = [..._research];
  if (fO) data = data.filter(r => (r.BrkSourceMachine || r.ORBMachine) === fO);
  if (fR) data = data.filter(r => r.BrkReason === fR);
  if (from || to) data = data.filter(r => inDateRange(r.ORBScanTime || r.OTBScanTime, from, to));

  // Enrich Material from _brk (BreakageSummary has Material; PowerResearch may not)
  const brkMatMap = {};
  _brk.forEach(r => { if (r.RxNumber && r.Material && r.Material !== 'Unknown') brkMatMap[String(r.RxNumber)] = r.Material; });
  data = data.map(r => {
    if (r.Material && r.Material !== 'Unknown') return r;
    const mat = brkMatMap[String(r.RxNumber)];
    return mat ? { ...r, Material: mat } : r;
  });

  // Apply material filter after enrichment so it works correctly
  if (fM) data = data.filter(r => r.Material === fM);
  return data;
}

const RES_COLS = [
  ['RxNumber', 'RX'],   ['ORBMachine', 'ORB'],    ['BrkSourceMachine', 'BrkSrc'],
  ['BrkReason', 'Reason'], ['Material', 'Material'], ['LensOption', 'Option'],
  ['BaseCurve', 'Curve'],  ['R_Sph', 'R Sph'],      ['L_Sph', 'L Sph'],
  ['R_Cyl', 'R Cyl'],      ['L_Cyl', 'L Cyl'],      ['R_AddPower', 'Add'],
  ['LensesBroken', 'Broken'],
];

function exportResearchCSV() {
  const data = getFilteredResearchData();
  if (!data.length) { alert('No data to export.'); return; }
  const header = RES_COLS.map(c => c[1]).join(',');
  const rows   = data.map(r =>
    RES_COLS.map(([key]) => {
      const v = r[key] ?? '';
      const s = String(v);
      return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')
  );
  const csv  = [header, ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `power_research_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportResearchXLSX() {
  const data = getFilteredResearchData();
  if (!data.length) { alert('No data to export.'); return; }

  const esc         = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const numericCols = new Set(['R_Sph', 'L_Sph', 'R_Cyl', 'L_Cyl', 'R_AddPower', 'LensesBroken']);

  const headerRow = '<Row>' + RES_COLS.map(([, label]) =>
    `<Cell><Data ss:Type="String">${esc(label)}</Data></Cell>`).join('') + '</Row>';

  const dataRows = data.map(r =>
    '<Row>' + RES_COLS.map(([key]) => {
      const v = r[key] ?? '';
      const n = parseFloat(v);
      if (numericCols.has(key) && !isNaN(n)) return `<Cell><Data ss:Type="Number">${n}</Data></Cell>`;
      return `<Cell><Data ss:Type="String">${esc(v)}</Data></Cell>`;
    }).join('') + '</Row>'
  ).join('');

  const xml = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Worksheet ss:Name="Power Research">
  <Table>${headerRow}${dataRows}</Table>
 </Worksheet>
</Workbook>`;

  const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `power_research_${new Date().toISOString().slice(0, 10)}.xls`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ============================================================
   LOAD ALL  (live — no date params sent to Apps Script)
============================================================ */
async function loadAll() {
  showOverlay();
  document.getElementById('liveStatus').textContent   = 'Loading...';
  document.getElementById('liveDot').style.background = CLR.peri;

  try {
    setLoadStep('ls-brk', 'active');
    const brk = await fetchTab('breakageSummary');
    setLoadStep('ls-brk', 'done'); setLoadStep('ls-reason', 'active');

    const reason = await fetchTab('reasonSummary');
    setLoadStep('ls-reason', 'done'); setLoadStep('ls-research', 'active');

    const research = await fetchTab('powerResearch');
    setLoadStep('ls-research', 'done'); setLoadStep('ls-anom', 'active');

    const anom = await fetchTab('anomalies');
    setLoadStep('ls-anom', 'done');

    _allBrk      = (Array.isArray(brk)      ? brk      : []).map(normalizeRow);
    _allReason   =  Array.isArray(reason)   ? reason   : [];
    _allResearch = (Array.isArray(research)  ? research : []).map(normalizeRow);
    _allAnom     = (Array.isArray(anom)      ? anom     : []).map(normalizeRow);

    // If single date filter is already set, re-fetch with that date
    const dateVal = document.getElementById('dateSingle')?.value;

    if (dateVal) {
      await applyDateFilter();
    } else {
      _brk      = [..._allBrk];
      _reason   = [..._allReason];
      _research = [..._allResearch];
      _anom     = [..._allAnom];

      document.getElementById('dateFilterCount').textContent = _brk.length;
      buildOverview();
      populateFlowFilters();    renderFlow();
      populateResearchFilters(); renderResearch();
      buildAlerts();
      buildWeekOptions();

      document.getElementById('liveStatus').textContent   = 'Live';
      document.getElementById('liveDot').style.background = '#d4c0a8';
      document.getElementById('liveDot').style.animation  = 'pulse 1.5s infinite';
    }

  } catch (err) {
    console.error(err);
    document.getElementById('liveStatus').textContent   = 'Error';
    document.getElementById('liveDot').style.background = '#f87171';
    document.getElementById('liveDot').style.animation  = 'none';
  } finally {
    setTimeout(hideOverlay, 300);
  }
}

/* ============================================================
   KEEP-WARM PING
   Fires a cheap request to Apps Script on page load so the
   V8 runtime is already warm when loadAll() fetches real data.
   Also fires every 4 minutes in sync with the server-side trigger.
============================================================ */
function pingWarm() {
  fetch(`${API}?tab=keepwarm`).catch(() => {}); // silent — we don't need the response
}

pingWarm();
setInterval(pingWarm, 4 * 60 * 1000);
setInterval(loadAll, 5 * 60 * 1000);
loadAll();

/* ============================================================
   WEEKLY SUMMARY
============================================================ */

let _sumWeeks            = [];
let _sumCurrentNarrative = '';
let _sumCurrentStats     = null;

function getMondayOf(date) {
  const d   = new Date(date);
  const day = d.getDay();
  const diff = (day === 0) ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getSundayOf(monday) {
  const d = new Date(monday);
  d.setDate(d.getDate() + 6);
  d.setHours(23, 59, 59, 999);
  return d;
}

function fmtWeekLabel(monday, sunday) {
  const mo = m => ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][m];
  return `${mo(monday.getMonth())} ${monday.getDate()} – ${mo(sunday.getMonth())} ${sunday.getDate()}, ${sunday.getFullYear()}`;
}

function fmtShortDate(d) {
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}

function buildWeekOptions() {
  const sel = document.getElementById('weekSelect');
  if (!sel) return;

  // Generate last 12 weeks (Mon–Sun) going back from current week
  // This ensures past weeks with BREAKAGE_HISTORY data are always available
  const today = new Date();
  const thisMonday = getMondayOf(today);

  _sumWeeks = [];
  for (let i = 0; i < 12; i++) {
    const mon = new Date(thisMonday);
    mon.setDate(mon.getDate() - (i * 7));
    mon.setHours(0,0,0,0);
    const sun = getSundayOf(mon);
    _sumWeeks.push({ label: fmtWeekLabel(mon, sun), from: mon, to: sun });
  }

  while (sel.options.length) sel.remove(0);
  _sumWeeks.forEach((w, i) => {
    const o = document.createElement('option');
    o.value = i; o.textContent = w.label; sel.appendChild(o);
  });

  // Default to most recent complete week (index 1 = last full Mon–Sun)
  // Index 0 is current week (may be partial), index 1 is last completed week
  const todayDay = today.getDay(); // 0=Sun
  sel.value = todayDay === 0 ? 0 : 1; // if today IS Sunday, current week just ended
  onWeekChange(/*silent=*/true);
}

async function onWeekChange(silent) {
  const sel  = document.getElementById('weekSelect');
  const idx  = parseInt(sel.value);
  const week = _sumWeeks[idx];
  if (!week) return;

  const rangeEl = document.getElementById('sumWeekRange');
  if (rangeEl) rangeEl.textContent = `${fmtShortDate(week.from)} → ${fmtShortDate(week.to)}`;

  // Determine if this week is the current live week or a past week
  const todayStr = new Date().toISOString().slice(0,10);
  const weekEnd  = week.to.toISOString().slice(0,10);
  const isPast   = weekEnd < todayStr;

  // Format dates for API
  const startStr = week.from.toISOString().slice(0,10);
  const endStr   = week.to.toISOString().slice(0,10);

  let wBrk = [];

  if (isPast) {
    // Fetch from BREAKAGE_HISTORY for this date range
    const genBtn = document.getElementById('sumGenBtn');
    if (genBtn) genBtn.disabled = true;

    // Show loading state in stat cards
    ['ss-jobs','ss-lens','ss-power','ss-axis','ss-ar41','ss-proc']
      .forEach(id => { const el = document.getElementById(id); if(el) el.textContent = '…'; });

    try {
      const raw = await fetchTab('breakageSummary', startStr, endStr);
      wBrk = (Array.isArray(raw) ? raw : []).map(normalizeRow);
    } catch(e) {
      console.error('Summary week fetch error:', e);
      wBrk = [];
    }
    if (genBtn) genBtn.disabled = false;
  } else {
    // Current week — filter from live data already loaded
    wBrk = _allBrk.filter(r => {
      const ds = r.AR41ScanTime || r.BrkTableScanTime;
      if (!ds) return false;
      const d = new Date(ds);
      return !isNaN(d) && d >= week.from && d <= week.to;
    });
  }

  const totalJobs = wBrk.length;
  const totalLens = wBrk.reduce((s, r) => s + (parseInt(r.LensesBroken) || 0), 0);
  const sPower    = wBrk.filter(r => (r.BrkReason || '').toLowerCase().includes('power')).length;
  const sAxis     = wBrk.filter(r => (r.BrkReason || '').toLowerCase().includes('axis')).length;
  const avgAR41   = avgArr(wBrk.map(r => r.AR41_to_Breakage_Min));
  const avgProc   = avgArr(wBrk.map(r => r.Breakage_to_Processed_Min));

  _sumCurrentStats = { week, wBrk, wRes: [], totalJobs, totalLens, sPower, sAxis, avgAR41, avgProc };

  document.getElementById('ss-jobs').textContent  = totalJobs  || '--';
  document.getElementById('ss-lens').textContent  = totalLens  || '--';
  document.getElementById('ss-power').textContent = sPower     || '--';
  document.getElementById('ss-axis').textContent  = sAxis      || '--';
  document.getElementById('ss-ar41').textContent  = fmtMin(avgAR41);
  document.getElementById('ss-proc').textContent  = fmtMin(avgProc);

  const machAgg = {};
  wBrk.forEach(r => { const m = r.BrkSourceMachine || 'Unknown'; machAgg[m] = (machAgg[m] || 0) + 1; });
  buildBarList('sumMachineList', Object.entries(machAgg).sort((a, b) => b[1] - a[1]).map(([l, v], i) => [l, v, v, CLR.series[i % CLR.series.length]]));

  const matAgg = {};
  wBrk.forEach(r => { const m = r.Material || 'Unknown'; matAgg[m] = (matAgg[m] || 0) + (parseInt(r.LensesBroken) || 0); });
  buildBarList('sumMaterialList', Object.entries(matAgg).sort((a, b) => b[1] - a[1]).map(([l, v], i) => [l, v, v, CLR.series[i % CLR.series.length]]));

  const grid = document.getElementById('sumBreakdownGrid');
  if (grid) grid.style.display = totalJobs > 0 ? '' : 'none';

  if (!silent) {
    _sumCurrentNarrative = '';
    document.getElementById('sumEmpty').style.display     = '';
    document.getElementById('sumTyping').style.display    = 'none';
    document.getElementById('sumNarrative').style.display = 'none';
    document.getElementById('sumNarrative').innerHTML     = '';
    const exportBtn = document.getElementById('sumExportBtn');
    if (exportBtn) exportBtn.disabled = true;
  }
}

function buildSummaryPayload(stats) {
  const { week, wBrk, totalJobs, totalLens, sPower, sAxis, avgAR41, avgProc } = stats;

  const reasonAgg = {};
  wBrk.forEach(r => { const k = r.BrkReason || 'Unknown'; reasonAgg[k] = (reasonAgg[k] || 0) + 1; });

  const machAgg = {};
  wBrk.forEach(r => { const m = r.BrkSourceMachine || 'Unknown'; machAgg[m] = (machAgg[m] || 0) + 1; });

  const matAgg = {};
  wBrk.forEach(r => { const m = r.Material || 'Unknown'; matAgg[m] = (matAgg[m] || 0) + (parseInt(r.LensesBroken) || 0); });

  const worst = [...wBrk]
    .map(r => ({ rx: r.RxNumber, machine: r.BrkSourceMachine, reason: r.BrkReason, min: parseFloat(r.AR41_to_Breakage_Min) || 0 }))
    .sort((a, b) => b.min - a.min)
    .slice(0, 5);

  return { weekLabel: week.label, totalJobs, totalLens, sPower, sAxis, avgAR41_min: Math.round(avgAR41), avgProc_min: Math.round(avgProc), reasonBreakdown: reasonAgg, machineBreakdown: machAgg, materialBreakdown: matAgg, worstJobs: worst };
}

function generateSummary() {
  if (!_sumCurrentStats) return;

  const genBtn = document.getElementById('sumGenBtn');
  genBtn.disabled = true;
  document.getElementById('sumEmpty').style.display     = 'none';
  document.getElementById('sumNarrative').style.display = 'none';
  document.getElementById('sumTyping').style.display    = 'flex';

  // Small timeout so the typing indicator renders before we do the work
  setTimeout(() => {
    try {
      const html = buildDataReport(_sumCurrentStats);
      _sumCurrentNarrative = html;

      document.getElementById('sumTyping').style.display    = 'none';
      const narEl = document.getElementById('sumNarrative');
      narEl.innerHTML     = html;
      narEl.style.display = '';

      const exportBtn = document.getElementById('sumExportBtn');
      if (exportBtn) exportBtn.disabled = false;

      const hint = document.getElementById('sumGenHint');
      const now  = new Date();
      if (hint) hint.textContent = `Generated ${now.getMonth()+1}/${now.getDate()} at ${now.getHours()}:${String(now.getMinutes()).padStart(2,'0')} · Auto-generated from breakage data`;

    } catch(err) {
      console.error('Summary build error:', err);
      document.getElementById('sumTyping').style.display    = 'none';
      document.getElementById('sumNarrative').innerHTML     = '<p style="color:var(--peach)">Failed to build report.</p>';
      document.getElementById('sumNarrative').style.display = '';
    } finally {
      genBtn.disabled = false;
    }
  }, 200);
}

/* ── Pure data-driven report builder — no API needed ── */
function buildDataReport(stats) {
  const { week, wBrk, totalJobs, totalLens, sPower, sAxis, avgAR41, avgProc } = stats;

  if (totalJobs === 0) {
    return `<h2>Week at a Glance</h2><p>No breakage events recorded for the week of <strong>${week.label}</strong>. Clean week — no action required.</p>`;
  }

  // ── Aggregations ──────────────────────────────────────────
  const reasonAgg = {};
  wBrk.forEach(r => { const k = r.BrkReason || 'Unknown'; reasonAgg[k] = (reasonAgg[k] || 0) + 1; });
  const reasonList = Object.entries(reasonAgg).sort((a,b) => b[1]-a[1]);
  const topReason  = reasonList[0];

  const machAgg = {};
  wBrk.forEach(r => { const m = r.BrkSourceMachine || 'Unknown'; machAgg[m] = (machAgg[m] || 0) + 1; });
  const machList = Object.entries(machAgg).sort((a,b) => b[1]-a[1]);
  const topMach  = machList[0];

  const matAgg = {};
  wBrk.forEach(r => { const m = r.Material || 'Unknown'; matAgg[m] = (matAgg[m] || 0) + (parseInt(r.LensesBroken)||0); });
  const matList = Object.entries(matAgg).sort((a,b) => b[1]-a[1]);
  const topMat  = matList[0];

  const worst = [...wBrk]
    .map(r => ({ rx: r.RxNumber, machine: r.BrkSourceMachine, reason: r.BrkReason, min: parseFloat(r.AR41_to_Breakage_Min)||0 }))
    .sort((a,b) => b.min - a.min)
    .slice(0, 5);

  const avgAR41Fmt = fmtMin(avgAR41);
  const avgProcFmt = fmtMin(avgProc);

  // ── Threshold flags ────────────────────────────────────────
  const ar41Flag  = avgAR41 > 480;   // > 8 hours is notable
  const procFlag  = avgProc > 120;   // > 2 hours is notable
  const highVol   = totalJobs >= 20;
  const lowVol    = totalJobs <= 3;
  const powerPct  = totalJobs > 0 ? Math.round(sPower / totalJobs * 100) : 0;
  const axisPct   = totalJobs > 0 ? Math.round(sAxis  / totalJobs * 100) : 0;

  // ── Section 1: Week at a Glance ───────────────────────────
  let glance = `<h2>Week at a Glance</h2>`;
  glance += `<p>The week of <strong>${week.label}</strong> recorded <strong>${totalJobs} breakage job${totalJobs!==1?'s':''}</strong> with a total of <strong>${totalLens} lens${totalLens!==1?'es':''} broken</strong>. `;

  if (lowVol) {
    glance += `This was a light week with minimal breakage activity. `;
  } else if (highVol) {
    glance += `This was a high-volume breakage week that warrants close review. `;
  } else {
    glance += `Volume was within normal range. `;
  }

  if (ar41Flag) {
    glance += `Average time from AR41 scan to Breakage entry was <strong>${avgAR41Fmt}</strong> — notably long and worth investigating for process delays.`;
  } else {
    glance += `Average AR41→Breakage time was <strong>${avgAR41Fmt}</strong> and processing time averaged <strong>${avgProcFmt}</strong>.`;
  }
  glance += `</p>`;

  // ── Section 2: Breakage Reasons ───────────────────────────
  let reasons = `<h2>Breakage Reasons</h2>`;
  if (reasonList.length === 1) {
    reasons += `<p>All <strong>${totalJobs} jobs</strong> were attributed to <strong>${topReason[0]}</strong> this week.</p>`;
  } else {
    reasons += `<p><strong>${topReason[0]}</strong> was the leading cause with <strong>${topReason[1]} job${topReason[1]!==1?'s':''}</strong> (${Math.round(topReason[1]/totalJobs*100)}% of total). `;
    if (reasonList.length > 1) {
      const second = reasonList[1];
      reasons += `<strong>${second[0]}</strong> accounted for <strong>${second[1]} job${second[1]!==1?'s':''}</strong> (${Math.round(second[1]/totalJobs*100)}%).`;
    }
    reasons += `</p>`;
  }

  reasons += `<ul>`;
  reasonList.forEach(([reason, count]) => {
    const pct = Math.round(count/totalJobs*100);
    reasons += `<li><strong>${reason}</strong> — ${count} job${count!==1?'s':''} · ${count > 1 ? parseInt(wBrk.filter(r=>r.BrkReason===reason).reduce((s,r)=>s+(parseInt(r.LensesBroken)||0),0)) + ' lenses broken' : '1 lens broken'} (${pct}%)</li>`;
  });
  reasons += `</ul>`;

  if (sPower === 0) reasons += `<div class="sum-callout">✓ No S-Power events this week — positive result.</div>`;
  if (sAxis  === 0) reasons += `<div class="sum-callout">✓ No S-Axis events this week — positive result.</div>`;

  // ── Section 3: Machines & Materials ───────────────────────
  let machines = `<h2>Machines &amp; Materials</h2>`;
  machines += `<p><strong>${topMach[0]}</strong> had the most breakage events this week with <strong>${topMach[1]} job${topMach[1]!==1?'s':''}</strong>`;
  if (machList.length > 1) {
    machines += `, followed by <strong>${machList[1][0]}</strong> (${machList[1][1]} job${machList[1][1]!==1?'s':''})`;
  }
  machines += `.</p>`;

  machines += `<p>By material, <strong>${topMat[0]}</strong> accounted for the most broken lenses (<strong>${topMat[1]}</strong>)`;
  if (matList.length > 1) {
    machines += `, followed by ${matList.slice(1,3).map(([m,c])=>`<strong>${m}</strong> (${c})`).join(' and ')}`;
  }
  machines += `.</p>`;

  if (machList.length > 1) {
    machines += `<ul>`;
    machList.forEach(([m, c]) => {
      machines += `<li><strong>${m}</strong> — ${c} job${c!==1?'s':''}</li>`;
    });
    machines += `</ul>`;
  }

  // ── Section 4: Observations & Recommendations ─────────────
  let obs = `<h2>Key Observations &amp; Recommendations</h2><ul>`;

  // Volume flag
  if (highVol) {
    obs += `<li><strong>High breakage volume</strong> — ${totalJobs} jobs this week is elevated. Review if any process changes coincide with this period.</li>`;
  } else if (lowVol) {
    obs += `<li><strong>Low breakage volume</strong> — only ${totalJobs} job${totalJobs!==1?'s':''} this week. No major concerns.</li>`;
  } else {
    obs += `<li><strong>Normal volume</strong> — ${totalJobs} jobs is within expected range.</li>`;
  }

  // AR41 delay flag
  if (ar41Flag) {
    obs += `<li><strong>AR41→Breakage time is high</strong> — averaging ${avgAR41Fmt}. Investigate queuing or staffing delays between AR41 and the breakage table.</li>`;
  } else {
    obs += `<li><strong>AR41→Breakage timing is acceptable</strong> — averaging ${avgAR41Fmt}.</li>`;
  }

  // Processing delay flag
  if (procFlag) {
    obs += `<li><strong>Breakage→Processed time is elevated</strong> — averaging ${avgProcFmt}. Check if breakage processing is being completed promptly.</li>`;
  }

  // Top machine flag
  if (machList.length > 0 && topMach[1] > 1) {
    obs += `<li><strong>Monitor ${topMach[0]}</strong> — led all machines with ${topMach[1]} breakage events. Consider a check of machine calibration or process parameters.</li>`;
  }

  // Worst delay jobs
  if (worst.length > 0 && worst[0].min > 60) {
    obs += `<li><strong>Longest delay job</strong> — RX <strong>${worst[0].rx}</strong> on ${worst[0].machine||'?'} had an AR41→Breakage time of <strong>${fmtMin(worst[0].min)}</strong>. `;
    obs += `Review this job for root cause.</li>`;
  }

  // S-Power dominant
  if (powerPct >= 70) {
    obs += `<li><strong>S-Power dominated</strong> at ${powerPct}% of all breakage. Focus troubleshooting on power-related process steps.</li>`;
  }

  obs += `</ul>`;

  return glance + reasons + machines + obs;
}

function exportSummaryTXT() {
  if (!_sumCurrentNarrative || !_sumCurrentStats) return;
  const { totalJobs, totalLens, sPower, sAxis, avgAR41, avgProc, week } = _sumCurrentStats;

  const tmp     = document.createElement('div');
  tmp.innerHTML = _sumCurrentNarrative;
  const plain   = tmp.innerText || tmp.textContent || '';

  const header = [
    `SURFACE FLOW — WEEKLY BREAKAGE REPORT`,
    `Week: ${week.label}`,
    `Generated: ${new Date().toLocaleString()}`,
    ``,
    `SUMMARY STATISTICS`,
    `  Jobs:             ${totalJobs}`,
    `  Lenses broken:    ${totalLens}`,
    `  S-Power events:   ${sPower}`,
    `  S-Axis events:    ${sAxis}`,
    `  Avg AR41 to Brk:  ${fmtMin(avgAR41)}`,
    `  Avg Brk to Proc:  ${fmtMin(avgProc)}`,
    ``,
    `─────────────────────────────────────────`,
    ``,
  ].join('\n');

  const blob = new Blob([header + plain], { type: 'text/plain;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `breakage_summary_${week.from.toISOString().slice(0, 10)}.txt`;
  a.click();
  URL.revokeObjectURL(a.href);
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('summaryTabBtn')?.addEventListener('click', () => {
    buildWeekOptions();
  });
});