/* ============================================================
   INCOMING JOBS — OPTION B: SPLIT PANEL  (UI IMPROVED)
   ============================================================ */

const API_URL = "https://script.google.com/macros/s/AKfycbyI2YqO9wXZ4-v34OXqqxD-yCYe3Dnly1-d_cf9mYtkcVkZoXeWhgQ8u6WbgY-lvIQQjg/exec";

let chart           = null;
let compChart       = null;
let trendChart      = null;
let sparklineCharts = {};
let lastData        = null;
let lastYestData    = null;
let currentHour     = null;
let currentDate     = null;
let prevDeptTotals  = {};
let selectedQueues  = []; // [{ dept, queue, lineColor }] — multi-select for hourly comparison

/* 8-colour palette for comparison lines — guaranteed distinct */
const MULTI_PALETTE = ["#38bdf8","#facc15","#22d3ee","#fb923c","#a78bfa","#4ade80","#f472b6","#2dd4bf"];

/* ===== DEPT CONFIG ===== */
const DEPT_COLORS = {
  "Finish":     "#22d3ee",
  "Speciality": "#a78bfa",
  "Specialty":  "#a78bfa",
  "Surface":    "#38bdf8",
  "Frame Only": "#fb923c"
};

/* ===== QUEUE COLORS ===== */
const QUEUE_COLORS = {
  "Surface Rush Delivery":"#00E5FF","Rush Delivery Queue":"#00E5FF","In Rush Delivery Queue":"#00E5FF",
  "Surface China Rush Delivery":"#F97316","China Rush Queue":"#F97316","In China Rush Queue":"#F97316","In China Rush Delivery Queue":"#F97316",
  "Surface Standard":"#3B82F6","Standard Queue":"#3B82F6","In Standard Queue":"#3B82F6","Fin Standard":"#93C5FD",
  "Overnight Queue":"#FACC15","In Overnight Queue":"#FACC15","Surface Overnight Delivery":"#FACC15","Fin Overnight Delivery":"#A3A3A3",
  "Fin Rush Delivery":"#86EFAC",
  "Beast Queue":"#A855F7","In Beast Queue":"#A855F7","Beast Surface Queue":"#D946EF","Beast Finish Queue":"#14B8A6",
  "Frame Only Queue":"#22C55E","In Frame Only Queue":"#22C55E","Frame Only Test Queue":"#15803D",
  "Test Jobs Queue":"#2DD4BF","Echo Queue":"#FB923C","All Queued Jobs":"#10B981"
};

const HOUR_ORDER = [
  "12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM",
  "6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM",
  "12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM",
  "6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"
];

/* ===== DATE HELPERS ===== */
function getLocalDate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function getYesterdayDate() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

/* ===== CACHE ===== */
const DATA_CACHE = {};

function cacheKey(date) { return date || getLocalDate(); }

function getCached(date) {
  const key   = cacheKey(date);
  const entry = DATA_CACHE[key];
  if (!entry) return null;
  const age     = Date.now() - entry.ts;
  const isToday = key === getLocalDate();
  if (!isToday && age < 86_400_000) return entry;
  if (isToday  && age < 5_000)      return entry;
  return null;
}

function setCache(date, data, total) {
  DATA_CACHE[cacheKey(date)] = { data, total, ts: Date.now() };
}

/* ===== SKELETON ===== */
function showSkeleton() {
  const grid = document.getElementById("kpiGrid");
  if (grid && !grid.querySelector(".kpi-skeleton")) {
    grid.innerHTML = "";
    for (let i = 0; i < 4; i++) {
      const el = document.createElement("div");
      el.className = "kpi-card kpi-skeleton";
      el.innerHTML = `<div class="skel skel-t"></div><div class="skel skel-v"></div><div class="skel skel-s"></div>`;
      grid.appendChild(el);
    }
  }
  const c = document.getElementById("container");
  if (c && !c.querySelector(".queue-loading")) {
    c.innerHTML = `<div class="queue-loading">
      <div class="skel skel-dh"></div><div class="skel skel-qr"></div>
      <div class="skel skel-qr"></div><div class="skel skel-qr"></div>
      <div class="skel skel-dh" style="margin-top:16px"></div>
      <div class="skel skel-qr"></div><div class="skel skel-qr"></div>
    </div>`;
  }
}

/* ===== INIT ===== */
document.addEventListener("DOMContentLoaded", () => {
  const today = getLocalDate();
  document.getElementById("dateFilter").value = today;
  currentDate = today;
  startLiveClock();
  showSkeleton();
  loadData();
  setInterval(() => {
    if (currentDate === getLocalDate()) loadData();
  }, 60_000);
});

/* ===== LIVE CLOCK ===== */
function startLiveClock() {
  const el = document.getElementById("lastUpdated");
  if (!el) return;
  const tick = () => {
    el.textContent = new Date().toLocaleTimeString([], { hour:"2-digit", minute:"2-digit", second:"2-digit" });
  };
  tick();
  setInterval(tick, 1000);
}

/* ===== TAB SWITCH ===== */
function switchTab(tab) {
  const inc = document.getElementById("incomingTab");
  const tre = document.getElementById("trendTab");
  const bI  = document.getElementById("tabIncomingBtn");
  const bT  = document.getElementById("tabTrendBtn");
  if (!inc || !tre) return;
  bI?.classList.remove("active"); bT?.classList.remove("active");
  if (tab === "incoming") {
    inc.style.display = "block"; tre.style.display = "none"; bI?.classList.add("active");
  } else {
    inc.style.display = "none"; tre.style.display = "block"; bT?.classList.add("active");
    if (!trendChart) loadTrend(14);
  }
}

/* ===== COLOR HELPERS ===== */
function getQueueColor(n) {
  if (QUEUE_COLORS[n]) return QUEUE_COLORS[n];
  const q = String(n||"").toLowerCase();
  if (q.includes("error")||q.includes("waiting")) return "#f87171";
  if (q.includes("hold"))        return "#fb7185";
  if (q.includes("surface rush"))return "#00E5FF";
  if (q.includes("china rush"))  return "#f87171";
  if (q.includes("standard"))    return "#3B82F6";
  if (q.includes("overnight"))   return "#FACC15";
  if (q.includes("beast"))       return "#A855F7";
  if (q.includes("frame only"))  return "#22C55E";
  if (q.includes("test"))        return "#2DD4BF";
  if (q.includes("echo"))        return "#FB923C";
  return "#5ab4ff";
}

function getDeptColor(d) { return DEPT_COLORS[d] || "#38bdf8"; }

function toBarFill(c) {
  const n=parseInt(c.replace("#",""),16);
  const r=Math.min(255,(n>>16)+40),g=Math.min(255,((n>>8)&0xff)+40),b=Math.min(255,(n&0xff)+40);
  return `linear-gradient(90deg,${c}cc,rgb(${r},${g},${b}))`;
}

/* ===== DEPT ABBREVIATION ===== */
function deptAbbr(dept) {
  const map = { "Finish":"FIN","Surface":"SUR","Speciality":"SPE","Specialty":"SPE","Frame Only":"FRM" };
  return map[dept] || dept.slice(0,3).toUpperCase();
}

/* ===== STATUS ===== */
function updateRefreshStatus(state="auto") {
  const el = document.getElementById("refreshStatus");
  if (!el) return;
  const today = getLocalDate();
  let mode = state === "updating" ? "updating" : (!currentDate || currentDate === today ? "live" : "history");
  el.className = `status-pill ${mode}`;
  el.textContent = { live:"Live", history:"History", updating:"Updating…" }[mode];
}

/* ===== DATE FILTER ===== */
function applyDateFilter() {
  const val = document.getElementById("dateFilter").value;
  if (!val) return;
  currentDate = val;
  currentHour = null;
  updateRefreshStatus();
  loadData();
}

function resetDateFilter() {
  currentDate = getLocalDate();
  currentHour = null;
  document.getElementById("dateFilter").value = currentDate;
  updateRefreshStatus();
  loadData();
}

/* ===== ERROR BANNER — uses CSS class instead of inline style ===== */
function showErrorBanner(msg) {
  let b = document.getElementById("errorBanner");
  if (!b) {
    b = document.createElement("div");
    b.id = "errorBanner";
    document.body.appendChild(b);
  }
  b.className = 'error-toast';
  b.innerHTML = `<span class="error-toast-icon">⚠</span><span>${msg}</span>`;
  b.style.opacity = "1";
  clearTimeout(b._t);
  b._t = setTimeout(() => { b.style.opacity = "0"; }, 5000);
}

/* ===== FETCH ===== */
async function apiFetch(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

/* ===== LOAD DATA (TODAY + YESTERDAY IN PARALLEL) ===== */
async function loadData() {
  if (window._isLoading) return;
  window._isLoading = true;
  updateRefreshStatus("updating");

  const todayDate = currentDate || getLocalDate();
  const yesterdayDate = (() => {
    const [y, m, d] = todayDate.split("-").map(Number);
    const dt = new Date(y, m - 1, d - 1);
    return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
  })();

  const cachedToday = getCached(todayDate);
  const cachedYest  = getCached(yesterdayDate);

  if (cachedToday && cachedYest) {
    renderAll(cachedToday.data, cachedToday.total, cachedYest.data);
    window._isLoading = false;
    updateRefreshStatus();
    return;
  }

  try {
    const [todayJson, yesterdayJson] = await Promise.all([
      cachedToday ? Promise.resolve({ success:true, data:cachedToday.data, total:cachedToday.total })
                  : apiFetch(`${API_URL}?date=${todayDate}`),
      cachedYest  ? Promise.resolve({ success:true, data:cachedYest.data,  total:cachedYest.total  })
                  : apiFetch(`${API_URL}?date=${yesterdayDate}`)
    ]);

    if (!todayJson.success) throw new Error("API error (today)");

    setCache(todayDate,    todayJson.data,    todayJson.total);
    setCache(yesterdayDate, yesterdayJson.success ? yesterdayJson.data : {}, yesterdayJson.success ? yesterdayJson.total : 0);

    lastYestData = yesterdayJson.success ? yesterdayJson.data : {};
    renderAll(todayJson.data, todayJson.total, lastYestData);

  } catch (err) {
    console.error("❌ loadData:", err);
    const stale = DATA_CACHE[cacheKey(todayDate)];
    if (stale) { renderAll(stale.data, stale.total, lastYestData || {}); showErrorBanner("Using cached data — server unreachable"); }
    else showErrorBanner("Failed to load data. Will retry automatically.");
  } finally {
    window._isLoading = false;
    updateRefreshStatus();
  }
}

function renderAll(data, total, yesterdayData) {
  const el = document.getElementById("totalJobs");
  if (el) el.textContent = total.toLocaleString();
  renderKPI(data, total, yesterdayData);
  currentHour ? showHourBreakdown(currentHour, data) : renderQueueBars(data, total);
  buildHourlyChart(data);
  setChartInsight(data);
  buildComparisonChart(data, yesterdayData || {});
  lastData = data;
}

/* ===== KPI CARDS — with % arc ring SVG ===== */
function renderKPI(data, grandTotal, yesterdayData) {
  const grid = document.getElementById("kpiGrid");
  if (!grid) return;
  Object.values(sparklineCharts).forEach(c => c?.destroy());
  sparklineCharts = {};
  grid.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    const todayVal = data[dept]._total ?? 0;
    const yesterdayVal = yesterdayData?.[dept]?._total ?? null;
    const pct = grandTotal > 0 ? ((todayVal / grandTotal)*100).toFixed(1) : "0.0";
    const isAlert = prevDeptTotals[dept] != null && todayVal > prevDeptTotals[dept] * 1.25;
    const deptColor = getDeptColor(dept);

    let chgHtml = "";
    if (yesterdayVal !== null && yesterdayVal > 0) {
      const chg = ((todayVal - yesterdayVal) / yesterdayVal * 100).toFixed(1);
      const cls = chg > 0 ? "up" : chg < 0 ? "down" : "flat";
      const icon = chg > 0 ? "↑" : chg < 0 ? "↓" : "—";
      const yesterdayDate = (() => {
        const base = currentDate || getLocalDate();
        const [y, m, d] = base.split("-").map(Number);
        const dt = new Date(y, m - 1, d - 1);
        return dt.toLocaleDateString([], { month:"short", day:"numeric" });
      })();
      chgHtml = `<div class="kpi-chg ${cls}">${icon} ${Math.abs(chg)}% vs ${yesterdayDate}</div>`;
    }

    const circ = 81.7;
    const arcLen = grandTotal > 0 ? ((todayVal / grandTotal) * circ).toFixed(1) : "0.0";
    const arcGap = (circ - parseFloat(arcLen)).toFixed(1);

    const card = document.createElement("div");
    card.className = "kpi-card" + (isAlert ? " alert" : "");
    card.dataset.dept = dept;
    card.innerHTML = `
      <div class="kpi-lbl">${dept}</div>
      ${isAlert ? `<div class="kpi-alert-badge">↑ Spike</div>` : ""}
      <div class="kpi-num" id="kv-${dept.replace(/\s/g,"-")}">0</div>
      <div class="kpi-pct">${pct}% of total</div>
      ${chgHtml}
      <canvas class="kpi-sparkline" id="sp-${dept.replace(/\s/g,"-")}"></canvas>
      <svg class="kpi-ring" viewBox="0 0 34 34">
        <circle fill="none" cx="17" cy="17" r="13" stroke="rgba(90,160,255,.08)" stroke-width="3"/>
        <circle fill="none" cx="17" cy="17" r="13" stroke="${deptColor}" stroke-width="3"
          stroke-dasharray="${arcLen} ${arcGap}" stroke-linecap="round"
          style="transform-origin:50% 50%;transform:rotate(-90deg)"/>
      </svg>
    `;
    grid.appendChild(card);

    animateCount(document.getElementById(`kv-${dept.replace(/\s/g,"-")}`), todayVal);
    drawSparkline(`sp-${dept.replace(/\s/g,"-")}`, buildDeptHourly(data[dept]), deptColor);
    prevDeptTotals[dept] = todayVal;
  });
}

/* ===== COUNT ANIMATION ===== */
function animateCount(el, target, dur=700) {
  if (!el) return;
  const start = performance.now();
  const step = now => {
    const p = Math.min((now - start) / dur, 1);
    const e = 1 - Math.pow(1-p, 3);
    el.textContent = Math.round(target * e).toLocaleString();
    if (p < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}

/* ===== DEPT HOURLY ARRAY ===== */
function buildDeptHourly(deptData) {
  const m = {};
  HOUR_ORDER.forEach(h => { m[h] = 0; });
  Object.keys(deptData).forEach(q => {
    if (q === "_total") return;
    Object.keys(deptData[q].hours || {}).forEach(h => {
      deptData[q].hours[h].forEach(e => { m[h] = (m[h]||0) + e.value; });
    });
  });
  return HOUR_ORDER.map(h => m[h]);
}

/* ===== SPARKLINE ===== */
function drawSparkline(id, data, color) {
  const canvas = document.getElementById(id);
  if (!canvas || !data.some(v => v > 0)) return;
  canvas.width  = canvas.parentElement?.offsetWidth || 200;
  canvas.height = 38;
  const ctx = canvas.getContext("2d");
  const g = ctx.createLinearGradient(0,0,0,38);
  g.addColorStop(0, color+"55"); g.addColorStop(1, color+"00");
  sparklineCharts[id] = new Chart(ctx, {
    type:"line",
    data:{ labels:HOUR_ORDER, datasets:[{ data, borderColor:color, backgroundColor:g, borderWidth:1.5, tension:.4, fill:true, pointRadius:0 }] },
    options:{ responsive:false, animation:false, layout:{padding:0},
      plugins:{legend:{display:false},tooltip:{enabled:false}},
      scales:{x:{display:false},y:{display:false,min:0}} }
  });
}

/* ===== TOGGLE DEPT (collapsible sections) ===== */
function toggleDept(id) {
  const rows = document.getElementById('drows-' + id);
  const chev = document.getElementById('dchev-' + id);
  if (!rows) return;
  const isOpen = !rows.classList.contains('collapsed');
  rows.classList.toggle('collapsed', isOpen);
  if (chev) chev.classList.toggle('open', !isOpen);
  rows.style.maxHeight = isOpen ? '0px' : rows.scrollHeight + 'px';
}

/* ===== QUEUE BARS — collapsible depts + dept tag + multi-select ===== */
function renderQueueBars(data, grandTotal) {
  const container = document.getElementById("container");
  if (!container) return;
  container.innerHTML = "";
  const label = document.getElementById("queueHourLabel");
  if (label) label.textContent = "";

  let deptIndex = 0;
  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    const i = deptIndex++;
    const deptTotal = data[dept]._total ?? 0;
    const color = getDeptColor(dept);
    const abbr  = deptAbbr(dept);
    const safeId = dept.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');

    const block = document.createElement("div");
    block.className = "queue-dept";

    const deptRowsDiv = document.createElement("div");
    deptRowsDiv.className = "dept-rows" + (i === 0 ? "" : " collapsed");
    deptRowsDiv.id = `drows-${safeId}`;
    deptRowsDiv.style.maxHeight = i === 0 ? "400px" : "0px";

    block.innerHTML = `
      <div class="dept-header" style="color:${color}" onclick="toggleDept('${safeId}')">
        <span>${dept}</span>
        <div class="dept-meta">
          <span class="dept-count">${deptTotal.toLocaleString()}</span>
          <span class="dept-chevron ${i === 0 ? 'open' : ''}" id="dchev-${safeId}">›</span>
        </div>
      </div>`;

    Object.keys(data[dept]).forEach(q => {
      if (q === "_total") return;
      const val    = data[dept][q].total;
      const barPct = deptTotal > 0 ? Math.max((val / deptTotal) * 100, 2) : 0;
      const qPct   = deptTotal > 0 ? ((val / deptTotal) * 100).toFixed(0) : 0;
      const qcolor = getQueueColor(q);

      // Is this (dept, queue) currently selected? Get its line colour
      const selIdx   = selectedQueues.findIndex(s => s.dept === dept && s.queue === q);
      const isActive = selIdx !== -1;
      const lineColor = isActive ? selectedQueues[selIdx].lineColor : null;

      const row = document.createElement("div");
      row.className = "queue-row" + (isActive ? " active" : "");
      row.dataset.dept  = dept;
      row.dataset.queue = q;
      // Use the selection's palette colour as a left-border indicator
      if (isActive) row.style.boxShadow = `inset 3px 0 0 ${lineColor}`;

      row.innerHTML = `<div class="q-label">${q}</div><div class="q-bar"><div class="q-fill"></div></div><div class="q-val">${val}</div><div class="q-pct">${qPct}%</div><span class="q-dept-tag" style="color:${color}">${abbr}</span>`;
      row.addEventListener('click', () => selectQueue(dept, q, qcolor));
      deptRowsDiv.appendChild(row);

      const fill = row.querySelector(".q-fill");
      const qlbl = row.querySelector(".q-label");
      qlbl.style.color      = qcolor;
      fill.style.background = toBarFill(qcolor);
      fill.style.boxShadow  = `0 0 8px ${qcolor}88`;
      requestAnimationFrame(() => requestAnimationFrame(() => { fill.style.width = barPct + "%"; }));
    });

    block.appendChild(deptRowsDiv);
    container.appendChild(block);
  });
}

/* ===== HOUR BREAKDOWN — collapsible dept headers ===== */
function showHourBreakdown(hour, dataOverride=null) {
  const data = dataOverride || lastData;
  if (!data) return;
  currentHour = hour;
  const container = document.getElementById("container");
  const label = document.getElementById("queueHourLabel");
  if (label) label.textContent = hour;
  container.innerHTML = `<div class="dept-header" style="color:var(--accent)"><span>Hour breakdown</span><span class="reset-link" onclick="resetView()">← All</span></div>`;

  let deptIndex = 0;
  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    const rows = [];
    Object.keys(data[dept]).forEach(q => {
      if (q === "_total") return;
      const entries = data[dept][q].hours?.[hour] || [];
      const total = entries.reduce((s,x)=>s+x.value,0);
      if (total > 0) rows.push({ q, total });
    });
    rows.sort((a,b)=>b.total-a.total);
    if (!rows.length) return;

    const i = deptIndex++;
    const deptTotal = rows.reduce((s,r)=>s+r.total,0);
    const color = getDeptColor(dept);
    const safeId = `hr-${dept.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}`;

    const block = document.createElement("div");
    block.className = "queue-dept";

    const deptRowsDiv = document.createElement("div");
    deptRowsDiv.className = "dept-rows" + (i === 0 ? "" : " collapsed");
    deptRowsDiv.id = `drows-${safeId}`;
    deptRowsDiv.style.maxHeight = i === 0 ? "400px" : "0px";

    block.innerHTML = `
      <div class="dept-header" style="color:${color}" onclick="toggleDept('${safeId}')">
        <span>${dept}</span>
        <div class="dept-meta">
          <span class="dept-count">${deptTotal.toLocaleString()}</span>
          <span class="dept-chevron ${i === 0 ? 'open' : ''}" id="dchev-${safeId}">›</span>
        </div>
      </div>`;

    rows.forEach(r => {
      const qcolor = getQueueColor(r.q);
      const pct  = Math.max((r.total / deptTotal) * 100, 4);
      const qPct = deptTotal > 0 ? ((r.total / deptTotal) * 100).toFixed(0) : 0;
      const row  = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `<div class="q-label">${r.q}</div><div class="q-bar"><div class="q-fill"></div></div><div class="q-val">${r.total}</div><div class="q-pct">${qPct}%</div><span class="q-dept-tag" style="color:${color}">${deptAbbr(dept)}</span>`;
      deptRowsDiv.appendChild(row);
      const fill = row.querySelector(".q-fill");
      row.querySelector(".q-label").style.color = qcolor;
      fill.style.background = toBarFill(qcolor);
      fill.style.boxShadow  = `0 0 8px ${qcolor}88`;
      fill.style.width = pct + "%";
    });

    block.appendChild(deptRowsDiv);
    container.appendChild(block);
  });
}

function resetView() {
  currentHour = null;
  if (lastData) renderQueueBars(lastData, lastData._total || 0);
}

/* ============================================================
   MULTI-QUEUE SELECTION — click rows to compare on hourly chart
   ============================================================ */

function selectQueue(dept, queueName, qcolor) {
  const existingIdx = selectedQueues.findIndex(s => s.dept === dept && s.queue === queueName);

  if (existingIdx !== -1) {
    // Deselect — remove and re-assign palette slots so colours stay consistent
    selectedQueues.splice(existingIdx, 1);
    selectedQueues.forEach((s, i) => { s.lineColor = MULTI_PALETTE[i % MULTI_PALETTE.length]; });
  } else {
    // Select — assign next available palette colour
    const lineColor = MULTI_PALETTE[selectedQueues.length % MULTI_PALETTE.length];
    selectedQueues.push({ dept, queue: queueName, lineColor });
  }

  refreshRowActiveStates();
  if (lastData) { buildHourlyChart(lastData); setChartInsight(lastData); }
  updateHourlyFilterChips();
  updateHintText();
}

function clearQueueFilter() {
  selectedQueues = [];
  refreshRowActiveStates();
  if (lastData) { buildHourlyChart(lastData); setChartInsight(lastData); }
  updateHourlyFilterChips();
  updateHintText();
}

/* Re-apply active class + line-colour border to all currently rendered rows */
function refreshRowActiveStates() {
  document.querySelectorAll('.queue-row[data-dept][data-queue]').forEach(r => {
    const idx = selectedQueues.findIndex(s => s.dept === r.dataset.dept && s.queue === r.dataset.queue);
    if (idx !== -1) {
      r.classList.add('active');
      r.style.boxShadow = `inset 3px 0 0 ${selectedQueues[idx].lineColor}`;
    } else {
      r.classList.remove('active');
      r.style.boxShadow = '';
    }
  });
}

/* Render chips — one per selected queue + "Clear all" when 2+ */
function updateHourlyFilterChips() {
  const container = document.getElementById('hourlyFilterChips');
  if (!container) return;
  container.innerHTML = '';

  if (selectedQueues.length === 0) {
    container.style.display = 'none';
    return;
  }
  container.style.display = 'flex';

  selectedQueues.forEach(sel => {
    const chip = document.createElement('button');
    chip.className = 'hourly-filter-badge';
    chip.innerHTML = `<span class="hourly-filter-dot" style="background:${sel.lineColor}"></span><span>${sel.queue} · ${deptAbbr(sel.dept)}</span><span style="opacity:.45;margin-left:4px">×</span>`;
    chip.onclick = () => selectQueue(sel.dept, sel.queue, null);
    container.appendChild(chip);
  });

  if (selectedQueues.length > 1) {
    const clearBtn = document.createElement('button');
    clearBtn.className = 'hourly-filter-badge';
    clearBtn.style.cssText = 'background:rgba(248,113,113,.08);border-color:rgba(248,113,113,.20);color:var(--red)';
    clearBtn.innerHTML = '× Clear all';
    clearBtn.onclick = clearQueueFilter;
    container.appendChild(clearBtn);
  }
}

function updateHintText() {
  const hint = document.querySelector('.chart-click-hint');
  if (!hint) return;
  if (selectedQueues.length === 0) {
    hint.textContent = "↑ Click any bar to drill into that hour's breakdown";
  } else if (selectedQueues.length === 1) {
    hint.textContent = `↑ ${selectedQueues[0].queue} · ${deptAbbr(selectedQueues[0].dept)} — click any bar to drill into that hour`;
  } else {
    hint.textContent = `↑ ${selectedQueues.length} queues compared — click any bar to drill into that hour`;
  }
}

/* Kept for backward compat (no-op in multi-select mode) */
function updateHourlyFilterBadge() {}

/* ===== COMPARISON CHART (Today vs Yesterday) ===== */
function buildComparisonChart(todayData, yesterdayData) {
  const canvas = document.getElementById("comparisonChart");
  if (!canvas) return;

  const titleEl = document.getElementById("compareTitle");
  if (titleEl) {
    const base = currentDate || getLocalDate();
    const [ty, tm, td] = base.split("-").map(Number);
    const tDate = new Date(ty, tm - 1, td);
    const yDate = new Date(ty, tm - 1, td - 1);
    const fmt = d => d.toLocaleDateString([], { month:"short", day:"numeric" });
    titleEl.textContent = `${fmt(tDate)} vs ${fmt(yDate)}`;
  }

  const legToday = document.getElementById("legToday");
  const legYest  = document.getElementById("legYest");
  if (legToday && legYest) {
    const base2 = currentDate || getLocalDate();
    const [ly, lm, ld] = base2.split("-").map(Number);
    const tDate2 = new Date(ly, lm - 1, ld);
    const yDate2 = new Date(ly, lm - 1, ld - 1);
    const fmt = d => d.toLocaleDateString([], { month:"short", day:"numeric" });
    legToday.textContent = fmt(tDate2);
    legYest.textContent  = fmt(yDate2);
  }

  const depts = Object.keys(todayData).filter(d => d !== "_total");
  const todayVals = depts.map(d => todayData[d]?._total ?? 0);
  const yesterVals = depts.map(d => yesterdayData[d]?._total ?? 0);

  if (compChart) {
    const tLbl = (()=>{ const b=currentDate||getLocalDate(); const [y,m,d]=b.split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString([],{month:"short",day:"numeric"}); })();
    const yLbl = (()=>{ const b=currentDate||getLocalDate(); const [y,m,d]=b.split("-").map(Number); return new Date(y,m-1,d-1).toLocaleDateString([],{month:"short",day:"numeric"}); })();
    compChart.data.labels           = depts;
    compChart.data.datasets[0].data = todayVals;
    compChart.data.datasets[0].label = tLbl;
    compChart.data.datasets[1].data = yesterVals;
    compChart.data.datasets[1].label = yLbl;
    compChart.update("none");
    return;
  }

  const ctx = canvas.getContext("2d");
  compChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: depts,
      datasets: [
        {
          label: (()=>{ const b=currentDate||getLocalDate(); const [y,m,d]=b.split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString([],{month:"short",day:"numeric"}); })(),
          data: todayVals, backgroundColor: "#38bdf8bb", borderRadius: 5, borderSkipped: false
        },
        {
          label: (()=>{ const b=currentDate||getLocalDate(); const [y,m,d]=b.split("-").map(Number); return new Date(y,m-1,d-1).toLocaleDateString([],{month:"short",day:"numeric"}); })(),
          data: yesterVals, backgroundColor: "rgba(56,189,248,0.22)", borderRadius: 5, borderSkipped: false
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      animation: { duration:600, easing:"easeOutQuart" },
      plugins: {
        legend: { display:false },
        tooltip: {
          backgroundColor:"rgba(6,13,26,.95)", borderColor:"rgba(56,189,248,.25)", borderWidth:1,
          titleColor:"#9ed8ff", bodyColor:"#e8f2ff",
          titleFont:{ family:"'Space Mono',monospace", size:11 },
          bodyFont: { family:"'DM Sans',sans-serif", size:12 }, padding:10,
          callbacks: { label: item => ` ${item.dataset.label}: ${item.raw.toLocaleString()}` }
        }
      },
      scales: {
        x: { grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12}} },
        y: { grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

/* ============================================================
   HOURLY CHART — bar (no selection) OR multi-line (1+ selected)
   ============================================================ */
function buildHourlyChart(data) {
  const canvas = document.getElementById("hourlyChart");
  if (!canvas) return;

  /* ---- MODE A: all queues combined, bar chart ---- */
  if (selectedQueues.length === 0) {

    // If chart was previously a line chart, destroy it
    if (chart && chart.config.type === 'line') { chart.destroy(); chart = null; }

    const hourMap = {};
    HOUR_ORDER.forEach(h => { hourMap[h] = 0; });
    Object.keys(data).forEach(dept => {
      if (dept === "_total") return;
      Object.keys(data[dept]).forEach(q => {
        if (q === "_total") return;
        Object.keys(data[dept][q].hours || {}).forEach(h => {
          data[dept][q].hours[h].forEach(e => { hourMap[h] = (hourMap[h]||0) + e.value; });
        });
      });
    });

    const labels = HOUR_ORDER.filter(h => hourMap[h] > 0);
    const values = labels.map(h => hourMap[h]);
    const maxVal = Math.max(...values, 1);
    const avg    = values.reduce((a,b)=>a+b,0) / (values.length||1);

    const colors = values.map(v => {
      const r = v / maxVal;
      if (r >= 1) return "rgba(120,255,190,.92)";
      if (r < .2) return "rgba(90,150,255,.28)";
      return "rgba(56,189,248,.72)";
    });

    if (chart) {
      chart.data.labels = labels;
      chart.data.datasets[0].data = values;
      chart.data.datasets[0].backgroundColor = colors;
      chart.data.datasets[1].data = values.map(()=>Math.round(avg));
      chart.update("none");
      return;
    }

    const ctx = canvas.getContext("2d");
    chart = new Chart(ctx, {
      type:"bar",
      data:{
        labels,
        datasets:[
          { type:"bar", data:values, backgroundColor:colors, borderRadius:4, borderSkipped:false, order:2 },
          { type:"line", label:"Avg", data:values.map(()=>Math.round(avg)),
            borderColor:"rgba(250,204,21,.55)", borderWidth:1.5, borderDash:[6,4],
            pointRadius:0, fill:false, tension:0, order:1 }
        ]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        animation:{ duration:600, easing:"easeOutQuart" },
        onClick:(evt,els,ci)=>{
          const pts=ci.getElementsAtEventForMode(evt,"index",{intersect:false},true);
          if(!pts.length) return;
          showHourBreakdown(ci.data.labels[pts[0].index]);
        },
        plugins:{
          legend:{ display:false },
          tooltip:{
            backgroundColor:"rgba(6,13,26,.95)", borderColor:"rgba(56,189,248,.25)", borderWidth:1,
            titleColor:"#9ed8ff", bodyColor:"#e8f2ff",
            titleFont:{family:"'Space Mono',monospace",size:11},
            bodyFont:{family:"'DM Sans',sans-serif",size:12}, padding:10,
            filter:item=>item.datasetIndex===0,
            callbacks:{ title:i=>i[0].label, label:i=>` ${i.raw.toLocaleString()} jobs` }
          }
        },
        scales:{
          x:{ grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},maxRotation:45} },
          y:{ grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},callback:v=>v.toLocaleString()} }
        }
      }
    });

  /* ---- MODE B: multi-queue comparison, line chart ---- */
  } else {

    // Destroy bar chart if switching from Mode A
    if (chart && chart.config.type === 'bar') { chart.destroy(); chart = null; }

    // Union of all hours that have any data across selected queues
    const allHourMap = {};
    HOUR_ORDER.forEach(h => { allHourMap[h] = 0; });
    selectedQueues.forEach(sel => {
      if (!data[sel.dept]?.[sel.queue]) return;
      Object.keys(data[sel.dept][sel.queue].hours || {}).forEach(h => {
        data[sel.dept][sel.queue].hours[h].forEach(e => { allHourMap[h] = (allHourMap[h]||0) + e.value; });
      });
    });
    const labels = HOUR_ORDER.filter(h => allHourMap[h] > 0);

    // Helper: get per-hour values for one (dept, queue) selection
    const getVals = sel => labels.map(h => {
      if (!data[sel.dept]?.[sel.queue]) return 0;
      return (data[sel.dept][sel.queue].hours?.[h] || []).reduce((s, e) => s + e.value, 0);
    });

    // Update existing line chart in-place if dataset count matches (smooth refresh)
    if (chart && chart.config.type === 'line' && chart.data.datasets.length === selectedQueues.length) {
      chart.data.labels = labels;
      selectedQueues.forEach((sel, i) => {
        chart.data.datasets[i].data = getVals(sel);
        chart.data.datasets[i].borderColor = sel.lineColor;
        chart.data.datasets[i].backgroundColor = sel.lineColor + '18';
      });
      chart.update("none");
      return;
    }

    if (chart) { chart.destroy(); chart = null; }

    const datasets = selectedQueues.map(sel => ({
      label: `${sel.queue} · ${deptAbbr(sel.dept)}`,
      data: getVals(sel),
      borderColor: sel.lineColor,
      backgroundColor: sel.lineColor + '18',
      borderWidth: 2,
      tension: 0.4,
      pointRadius: 3,
      pointHoverRadius: 6,
      fill: false
    }));

    const ctx = canvas.getContext("2d");
    chart = new Chart(ctx, {
      type: 'line',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        animation: { duration: 400, easing: "easeOutQuart" },
        onClick: (evt, els, ci) => {
          const pts = ci.getElementsAtEventForMode(evt, "index", {intersect:false}, true);
          if (!pts.length) return;
          showHourBreakdown(ci.data.labels[pts[0].index]);
        },
        plugins: {
          legend: {
            display: true,
            labels: { color:"#6a8ab0", font:{family:"'DM Sans',sans-serif",size:11}, boxWidth:12, padding:12, usePointStyle:true }
          },
          tooltip: {
            mode: 'index', intersect: false,
            backgroundColor:"rgba(6,13,26,.95)", borderColor:"rgba(56,189,248,.25)", borderWidth:1,
            titleColor:"#9ed8ff", bodyColor:"#e8f2ff",
            titleFont:{family:"'Space Mono',monospace",size:11},
            bodyFont:{family:"'DM Sans',sans-serif",size:12}, padding:10,
            callbacks: { label: item => ` ${item.dataset.label}: ${item.raw.toLocaleString()}` }
          }
        },
        scales: {
          x: { grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},maxRotation:45} },
          y: { grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},callback:v=>v.toLocaleString()} }
        }
      }
    });
  }
}

/* ===== CHART INSIGHT ===== */
function setChartInsight(data) {
  const el = document.getElementById("chartInsight");
  if (!el) return;

  // Multi-queue: show count
  if (selectedQueues.length > 1) {
    el.textContent = `${selectedQueues.length} queues`;
    return;
  }

  // No selection or single selection — compute peak hour
  const fDept  = selectedQueues[0]?.dept  ?? null;
  const fQueue = selectedQueues[0]?.queue ?? null;

  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });
  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    if (fDept && dept !== fDept) return;
    Object.keys(data[dept]).forEach(q => {
      if (q === "_total") return;
      if (fQueue && q !== fQueue) return;
      Object.keys(data[dept][q].hours||{}).forEach(h => {
        data[dept][q].hours[h].forEach(e => { hourMap[h]=(hourMap[h]||0)+e.value; });
      });
    });
  });
  let peak=null,val=0;
  Object.keys(hourMap).forEach(h => { if(hourMap[h]>val){val=hourMap[h];peak=h;} });
  el.textContent = peak ? `Peak: ${peak} (${val.toLocaleString()})` : "";
}

/* ===== TREND ===== */
let trendView = "dept";   // "dept" | "queue"
let trendDays = 14;
let trendData = null;     // cached last API response

const TREND_DEPT_CFG = [
  { key:"Finish",     label:"Finish",   color:"#818cf8", alpha:"rgba(129,140,248,.75)" },
  { key:"Surface",    label:"Surface",  color:"#22d3ee", alpha:"rgba(34,211,238,.75)"  },
  { key:"Specialty",  label:"Specialty",color:"#f472b6", alpha:"rgba(244,114,182,.75)" },
  { key:"Frame Only", label:"Frame",    color:"#fb923c", alpha:"rgba(251,146,60,.75)"  },
];

// ── view switch ──────────────────────────────────
function setTrendView(view) {
  trendView = view;
  const dTab = document.getElementById("trendTabDept");
  const qTab = document.getElementById("trendTabQueue");
  if (dTab) dTab.className = "trend-view-tab" + (view==="dept" ? " active-dept"  : "");
  if (qTab) qTab.className = "trend-view-tab" + (view==="queue"? " active-queue" : "");
  const dZone = document.getElementById("trendDeptZone");
  const qZone = document.getElementById("trendQueueZone");
  if (dZone) dZone.style.display = view==="dept"  ? "block" : "none";
  if (qZone) qZone.style.display = view==="queue" ? "block" : "none";
  if (trendData) {
    if (view==="dept") buildTrendDeptChart(trendData);
    else               buildTrendHeatmap(trendData);
    _updateTrendInsightBar(trendData);
  } else {
    loadTrend(trendDays);
  }
}

// ── fetch ─────────────────────────────────────────
async function loadTrend(days=14) {
  trendDays = days;
  document.querySelectorAll(".trend-btn").forEach(b => {
    b.classList.toggle("active", b.textContent.trim()===`${days}D`);
  });
  try {
    const json = await apiFetch(`${API_URL}?mode=trend&days=${days}&queues=1`);
    if (!json.success) return;
    trendData = json.trend;
    if (trendView==="dept") buildTrendDeptChart(json.trend);
    else                    buildTrendHeatmap(json.trend);
    setTrendInsight(json.trend);
    _updateTrendInsightBar(json.trend);
  } catch(err) { console.error("Trend error:", err); }
}

// ── All Dept: grouped bar + total line + avg ──────
function buildTrendDeptChart(trend) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;
  if (trendChart) { trendChart.destroy(); trendChart = null; }

  const labels = trend.map(d => d.date);
  const totals = trend.map(d => d.total);
  const avgV   = Math.round(totals.reduce((a,b)=>a+b,0) / totals.length);
  const n      = labels.length;

  const datasets = [
    // dept grouped bars
    ...TREND_DEPT_CFG.map(d => ({
      type:"bar", label:d.label,
      data: trend.map(r => r.departments?.[d.key] || 0),
      backgroundColor:d.alpha, borderColor:d.color,
      borderWidth:1, borderRadius:3, order:2
    })),
    // total line
    {
      type:"line", label:"Total", data:totals,
      borderColor:"#f1f5f9", backgroundColor:"transparent",
      borderWidth:2.5, tension:.4,
      pointRadius: n > 20 ? 0 : 4,
      pointBackgroundColor:"#f1f5f9",
      pointBorderColor:"#0f172a", pointBorderWidth:2,
      order:1
    },
    // avg dashed
    {
      type:"line", label:"Avg", data:Array(n).fill(avgV),
      borderColor:"#facc15", backgroundColor:"transparent",
      borderWidth:1.5, borderDash:[7,4], pointRadius:0, order:3
    }
  ];

  trendChart = new Chart(canvas.getContext("2d"), {
    type:"bar",
    data:{ labels, datasets },
    options:{
      responsive:true, maintainAspectRatio:false, animation:{duration:400},
      plugins:{
        legend:{ display:false },
        tooltip:{
          mode:"index", intersect:false,
          backgroundColor:"rgba(6,13,26,.95)",
          borderColor:"rgba(56,189,248,.25)", borderWidth:1,
          titleColor:"#9ed8ff", bodyColor:"#e8f2ff", padding:10,
          callbacks:{
            footer(items) {
              const tot = items
                .filter(i => !["Total","Avg"].includes(i.dataset.label))
                .reduce((s,i) => s + (i.raw||0), 0);
              return `Total: ${tot.toLocaleString()}`;
            }
          }
        }
      },
      scales:{
        x:{ grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:11},maxTicksLimit:10,maxRotation:45} },
        y:{ grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:11},callback:v=>v.toLocaleString()} }
      }
    }
  });

  const legRow = document.getElementById("trendLegRow");
  if (legRow) {
    legRow.innerHTML =
      TREND_DEPT_CFG.map(d =>
        `<span class="trend-leg-item"><i class="trend-leg-sq" style="background:${d.alpha};border:1px solid ${d.color}"></i>${d.label}</span>`
      ).join("")
      + `<span class="trend-leg-item"><i class="trend-leg-dot" style="background:#f1f5f9"></i>Total</span>`
      + `<span class="trend-leg-item" style="border-left:2px dashed #facc15;padding-left:6px;color:#facc15">Avg ${avgV.toLocaleString()}</span>`;
  }
}

// ── By Queue: heatmap grid ────────────────────────
function buildTrendHeatmap(trend) {
  const grid = document.getElementById("trendHmGrid");
  if (!grid) return;

  const DEPT_MAP = {};
  TREND_DEPT_CFG.forEach(d => { DEPT_MAP[d.key] = d; });
  const DEPT_ORDER = TREND_DEPT_CFG.map(d => d.key);

  // Detect if API returned per-queue data inside d.queues
  const hasQueues = trend.some(d => d.queues && Object.keys(d.queues).length > 0);

  let rows = [];
  if (hasQueues) {
    // Build queue list per dept from union of all days
    const qSetByDept = {};
    DEPT_ORDER.forEach(dk => { qSetByDept[dk] = new Set(); });
    trend.forEach(d => {
      if (!d.queues) return;
      Object.entries(d.queues).forEach(([dk, qMap]) => {
        if (qSetByDept[dk]) Object.keys(qMap).forEach(q => qSetByDept[dk].add(q));
      });
    });
    DEPT_ORDER.forEach(dk => {
      Array.from(qSetByDept[dk]).forEach(qName => {
        rows.push({
          label:  qName,
          dept:   dk,
          color:  DEPT_MAP[dk].color,
          isQueue: true,
          vals:   trend.map(d => d.queues?.[dk]?.[qName] || 0)
        });
      });
    });
  } else {
    // Fallback: dept-level rows (same data as chart, different visual)
    DEPT_ORDER.forEach(dk => {
      rows.push({
        label:  DEPT_MAP[dk].label,
        dept:   dk,
        color:  DEPT_MAP[dk].color,
        isQueue: false,
        vals:   trend.map(d => d.departments?.[dk] || 0)
      });
    });
  }

  const labels = trend.map(d => d.date);
  const n = labels.length;
  const allVals = rows.flatMap(r => r.vals);
  const minV = Math.min(...allVals), maxV = Math.max(...allVals);

  grid.innerHTML = "";
  grid.style.gridTemplateColumns = `100px repeat(${n}, 1fr)`;

  // date header row
  const blankHdr = document.createElement("div");
  blankHdr.className = "trend-hm-header"; grid.appendChild(blankHdr);
  labels.forEach(l => {
    const el = document.createElement("div");
    el.className = "trend-hm-header"; el.textContent = l; grid.appendChild(el);
  });

  if (hasQueues) {
    DEPT_ORDER.forEach((dk, di) => {
      const dRows = rows.filter(r => r.dept === dk);
      if (!dRows.length) return;
      const cfg = DEPT_MAP[dk];
      // dept header row
      const dLbl = document.createElement("div");
      dLbl.className = "trend-hm-dept-lbl";
      dLbl.style.color = cfg.color;
      dLbl.textContent = cfg.label;
      grid.appendChild(dLbl);
      for (let c = 0; c < n; c++) {
        const f = document.createElement("div");
        f.style.cssText = "height:26px;background:rgba(255,255,255,.025);border-radius:4px";
        grid.appendChild(f);
      }
      // queue rows
      dRows.forEach(row => _hmRow(grid, row, n, minV, maxV));
      // gap between depts
      if (di < DEPT_ORDER.length - 1) {
        const gap = document.createElement("div");
        gap.className = "trend-hm-gap"; gap.style.gridColumn = "1 / -1";
        grid.appendChild(gap);
      }
    });
  } else {
    rows.forEach(row => _hmRow(grid, row, n, minV, maxV));
  }

  // legend
  const hmLeg = document.getElementById("trendHmLeg");
  if (hmLeg) {
    hmLeg.innerHTML =
      `<span class="trend-leg-item" style="color:#475569;font-size:10px">Volume →</span>`
      + [.12,.3,.5,.7,.9].map(a =>
          `<span class="trend-leg-item"><i class="trend-leg-sq" style="background:#818cf8;opacity:${a}"></i></span>`
        ).join("")
      + `<span class="trend-leg-item" style="color:#475569;font-size:10px">Low → High &nbsp;·&nbsp; Hover for exact count</span>`
      + `&emsp;`
      + TREND_DEPT_CFG.map(d =>
          `<span class="trend-leg-item"><i class="trend-leg-dot" style="background:${d.color}"></i>${d.label}</span>`
        ).join("");
  }
}

function _hmRow(grid, row, n, minV, maxV) {
  const lbl = document.createElement("div");
  lbl.className = "trend-hm-lbl"; lbl.textContent = row.label; lbl.title = row.label;
  grid.appendChild(lbl);
  row.vals.forEach(v => {
    // use per-row min/max so every row is self-scaled (no row disappears)
    const rowMax = Math.max(...row.vals);
    const rowMin = Math.min(...row.vals);
    const t = (v - rowMin) / (rowMax - rowMin || 1);
    const cell = document.createElement("div");
    cell.className = "trend-hm-cell";
    cell.style.background = row.color;
    // min opacity 0.22 so even low-volume rows are always visible
    cell.style.opacity = v === 0 ? 0.07 : 0.22 + t * 0.72;
    cell.title = `${row.label}: ${v.toLocaleString()}`;
    // always show value ("—" for zero)
    cell.textContent = v === 0 ? "—" : v >= 1000 ? (v/1000).toFixed(1)+"k" : v.toLocaleString();
    grid.appendChild(cell);
  });
}

// ── insight strip ─────────────────────────────────
function _updateTrendInsightBar(trend) {
  const bar = document.getElementById("trendInsightBar");
  if (!bar || !trend.length) return;
  const totals  = trend.map(d => d.total);
  const peakIdx = totals.indexOf(Math.max(...totals));
  const avgV    = Math.round(totals.reduce((a,b)=>a+b,0) / totals.length);
  const total   = totals.reduce((a,b)=>a+b,0);
  bar.innerHTML =
    `<div class="trend-insight-item">Peak <span class="trend-insight-val">${trend[peakIdx].date}</span> &mdash; <span class="trend-insight-val">${totals[peakIdx].toLocaleString()}</span> jobs</div>`
    + `<div class="trend-insight-item">Daily avg <span class="trend-insight-val">${avgV.toLocaleString()}</span></div>`
    + `<div class="trend-insight-item">Period total <span class="trend-insight-val">${total.toLocaleString()}</span> jobs</div>`;
}

// ── insight pill (top-right) ──────────────────────
function setTrendInsight(trend) {
  const el = document.getElementById("trendInsight");
  if (!el || trend.length < 2) return;
  const peak   = trend.reduce((b,d) => d.total > b.total ? d : b, trend[0]);
  const latest = trend[trend.length-1];
  const prev   = trend[trend.length-2];
  const chg    = prev.total > 0 ? ((latest.total - prev.total) / prev.total * 100).toFixed(1) : "0.0";
  el.textContent = `Peak ${peak.date}  (${peak.total.toLocaleString()})  ${Number(chg)>=0?"↑":"↓"}  ${Math.abs(chg)}%`;
}

/* ===== NAVIGATION ===== */
function goBack() { window.location.href="index.html"; }

/* ============================================================
   COMPARE TAB
   ============================================================ */

const COMP_COLORS = ["#38bdf8","#a78bfa","#22d3ee","#fb923c","#4ade80"];
let compMultiChart = null;

/* ===== SWITCH TAB — extend original ===== */
const _origSwitchTab = switchTab;
switchTab = function(tab) {
  const compTab = document.getElementById("compareTab");
  const compBtn = document.getElementById("tabCompareBtn");
  if (compTab) compTab.style.display = "none";
  if (compBtn) compBtn.classList.remove("active");

  if (tab === "compare") {
    document.getElementById("incomingTab").style.display = "none";
    document.getElementById("trendTab").style.display    = "none";
    compTab.style.display = "block";
    compBtn.classList.add("active");
    document.getElementById("tabIncomingBtn")?.classList.remove("active");
    document.getElementById("tabTrendBtn")?.classList.remove("active");
    initCompareDefaults();
  } else {
    _origSwitchTab(tab);
  }
};

function initCompareDefaults() {
  const inputs = document.querySelectorAll(".comp-date-input");
  const today  = getLocalDate();
  inputs.forEach((inp, i) => {
    if (!inp.value) {
      const [y,m,d] = today.split("-").map(Number);
      const dt = new Date(y, m-1, d - i);
      inp.value = `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
    }
  });
}

function resetComparison() {
  document.querySelectorAll(".comp-date-input").forEach(inp => inp.value = "");
  document.getElementById("compKpiGrid").innerHTML  = "";
  document.getElementById("compDeptGrid").innerHTML = "";
  document.getElementById("compChartCard").style.display = "none";
  if (compMultiChart) { compMultiChart.destroy(); compMultiChart = null; }
  initCompareDefaults();
}

async function runComparison() {
  const inputs = document.querySelectorAll(".comp-date-input");
  const dates  = [...inputs].map(i => i.value).filter(Boolean);

  if (dates.length < 2) {
    alert("Please select at least 2 dates to compare.");
    return;
  }

  const kpiGrid   = document.getElementById("compKpiGrid");
  const deptGrid  = document.getElementById("compDeptGrid");
  const chartCard = document.getElementById("compChartCard");
  kpiGrid.innerHTML  = `<div class="comp-empty"><div class="comp-empty-icon">⟳</div>Loading ${dates.length} dates…</div>`;
  deptGrid.innerHTML = "";
  chartCard.style.display = "none";

  try {
    const results = await Promise.all(dates.map(async (date, i) => {
      const cached = getCached(date);
      if (cached) return { date, color: COMP_COLORS[i], data: cached.data, total: cached.total };
      const json = await apiFetch(`${API_URL}?date=${date}`);
      if (json.success) setCache(date, json.data, json.total);
      return { date, color: COMP_COLORS[i], data: json.success ? json.data : {}, total: json.success ? json.total : 0 };
    }));

    renderCompKPI(results);
    renderCompChart(results);
    renderCompDeptGrid(results);

  } catch(err) {
    console.error("Compare error:", err);
    kpiGrid.innerHTML = `<div class="comp-empty"><div class="comp-empty-icon">⚠</div>Failed to load data. Check your connection.</div>`;
  }
}

function fmtDate(dateStr) {
  const [y,m,d] = dateStr.split("-").map(Number);
  return new Date(y,m-1,d).toLocaleDateString([],{month:"short",day:"numeric"});
}

function renderCompKPI(results) {
  const grid = document.getElementById("compKpiGrid");
  grid.innerHTML = "";

  const depts = [...new Set(results.flatMap(r => Object.keys(r.data).filter(k => k !== "_total")))];

  depts.forEach(dept => {
    const deptColor = getDeptColor(dept);
    const card = document.createElement("div");
    card.className = "comp-kpi-card";
    card.style.setProperty("--ckc", deptColor);

    const rows = results.map((r, i) => {
      const val = r.data[dept]?._total ?? 0;
      return { date: r.date, color: r.color, val, idx: i };
    });

    const base = rows[0].val;

    let rowsHtml = rows.map((r, i) => {
      const chg = i > 0 && base > 0 ? ((r.val - base) / base * 100).toFixed(1) : null;
      const chgClass = chg === null ? "" : chg > 0 ? "up" : chg < 0 ? "down" : "flat";
      const chgIcon  = chg === null ? "" : chg > 0 ? "↑" : chg < 0 ? "↓" : "—";
      const chgText  = chg !== null ? `<span class="comp-kpi-chg ${chgClass}">${chgIcon}${Math.abs(chg)}%</span>` : "";
      return `
        <div class="comp-kpi-row">
          <div class="comp-kpi-date">
            <div class="comp-kpi-date-dot" style="background:${r.color}"></div>
            ${fmtDate(r.date)}
          </div>
          <div style="display:flex;align-items:center;gap:6px">
            <div class="comp-kpi-num">${r.val.toLocaleString()}</div>
            ${chgText}
          </div>
        </div>`;
    }).join("");

    card.innerHTML = `<div class="comp-kpi-dept">${dept}</div>${rowsHtml}`;
    grid.appendChild(card);
  });
}

function renderCompChart(results) {
  const canvas    = document.getElementById("compMultiChart");
  const chartCard = document.getElementById("compChartCard");
  const insight   = document.getElementById("compInsight");
  if (!canvas) return;

  chartCard.style.display = "block";

  const depts = [...new Set(results.flatMap(r => Object.keys(r.data).filter(k => k !== "_total")))];

  const peak = results.reduce((b, r) => r.total > b.total ? r : b, results[0]);
  if (insight) insight.textContent = `Peak: ${fmtDate(peak.date)} (${peak.total.toLocaleString()})`;

  const datasets = results.map(r => ({
    label: fmtDate(r.date),
    data: depts.map(d => r.data[d]?._total ?? 0),
    backgroundColor: r.color + "bb",
    borderRadius: 4,
    borderSkipped: false
  }));

  if (compMultiChart) {
    compMultiChart.data.labels   = depts;
    compMultiChart.data.datasets = datasets;
    compMultiChart.update("none");
    return;
  }

  const ctx = canvas.getContext("2d");
  compMultiChart = new Chart(ctx, {
    type: "bar",
    data: { labels: depts, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      animation: { duration:600, easing:"easeOutQuart" },
      plugins: {
        legend: { display: true, labels: { color:"#6a8ab0", font:{family:"'DM Sans',sans-serif",size:12}, boxWidth:12, padding:14, usePointStyle:true } },
        tooltip: {
          backgroundColor:"rgba(6,13,26,.95)", borderColor:"rgba(56,189,248,.25)", borderWidth:1,
          titleColor:"#9ed8ff", bodyColor:"#e8f2ff",
          titleFont:{family:"'Space Mono',monospace",size:11},
          bodyFont:{family:"'DM Sans',sans-serif",size:12}, padding:10,
          callbacks: { label: item => ` ${item.dataset.label}: ${item.raw.toLocaleString()}` }
        }
      },
      scales: {
        x: { grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12}} },
        y: { grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#c0d0e0",font:{family:"'DM Sans',sans-serif",size:12},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

function renderCompDeptGrid(results) {
  const grid = document.getElementById("compDeptGrid");
  grid.innerHTML = "";

  const depts = [...new Set(results.flatMap(r => Object.keys(r.data).filter(k => k !== "_total")))];

  depts.forEach(dept => {
    const deptColor = getDeptColor(dept);

    const allQueues = [...new Set(
      results.flatMap(r => Object.keys(r.data[dept] || {}).filter(k => k !== "_total"))
    )];

    if (!allQueues.length) return;

    const card = document.createElement("div");
    card.className = "comp-dept-card";
    card.style.setProperty("--dkc", deptColor);

    let qRowsHtml = allQueues.map(q => {
      const vals   = results.map(r => r.data[dept]?.[q]?.total ?? 0);
      const maxVal = Math.max(...vals, 1);
      const qcolor = getQueueColor(q);

      const barsHtml = vals.map((v, i) => {
        const pct = Math.max((v / maxVal) * 100, 2);
        return `<div class="comp-mini-bar" style="width:${pct}%;background:${results[i].color};opacity:0.75"></div>`;
      }).join("");

      const valsHtml = vals.map(v =>
        `<div class="comp-queue-val">${v > 0 ? v.toLocaleString() : "—"}</div>`
      ).join("");

      return `
        <div class="comp-queue-row">
          <div class="comp-queue-name" style="color:${qcolor}" title="${q}">${q}</div>
          <div class="comp-bars-col">${barsHtml}</div>
          <div class="comp-queue-vals">${valsHtml}</div>
        </div>`;
    }).join("");

    card.innerHTML = `
      <div class="comp-dept-strip" style="background:linear-gradient(90deg,${deptColor}20 0%,transparent 100%);">
        <div class="comp-dept-title" style="color:${deptColor}">${dept}</div>
        <div class="comp-dept-totals">
          ${results.map(r => `<div class="comp-total-chip"><div class="chip-dot" style="background:${r.color}"></div>${(r.data[dept]?._total ?? 0).toLocaleString()}</div>`).join('')}
        </div>
      </div>
      <div class="comp-dept-body">${qRowsHtml}</div>
    `;
    grid.appendChild(card);
  });
}
