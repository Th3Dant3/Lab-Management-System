/* ============================================================
   INCOMING JOBS — OPTION B: SPLIT PANEL
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
  currentDate = val; // use raw value directly — no Date() conversion to avoid UTC timezone shift
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

/* ===== ERROR BANNER ===== */
function showErrorBanner(msg) {
  let b = document.getElementById("errorBanner");
  if (!b) {
    b = document.createElement("div");
    b.id = "errorBanner";
    b.style.cssText = "position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:rgba(248,113,113,.12);border:1px solid rgba(248,113,113,.35);color:#fca5a5;padding:9px 20px;border-radius:10px;font-family:'Space Mono',monospace;font-size:12px;z-index:9999;backdrop-filter:blur(10px);transition:opacity .5s;pointer-events:none";
    document.body.appendChild(b);
  }
  b.textContent = `⚠  ${msg}`;
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
    // Parse date parts directly to avoid UTC timezone shift
    const [y, m, d] = todayDate.split("-").map(Number);
    const dt = new Date(y, m - 1, d - 1); // local time, subtract 1 day
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

/* ===== KPI CARDS ===== */
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

/* ===== QUEUE BARS ===== */
function renderQueueBars(data, grandTotal) {
  const container = document.getElementById("container");
  if (!container) return;
  container.innerHTML = "";
  const label = document.getElementById("queueHourLabel");
  if (label) label.textContent = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    const deptTotal = data[dept]._total ?? 0;
    const color = getDeptColor(dept);
    const block = document.createElement("div");
    block.className = "queue-dept";
    block.innerHTML = `<div class="dept-header" style="color:${color}"><span>${dept}</span><span class="dept-count">${deptTotal.toLocaleString()}</span></div>`;

    Object.keys(data[dept]).forEach(q => {
      if (q === "_total") return;
      const val   = data[dept][q].total;
      const pct   = deptTotal > 0 ? Math.max((val/deptTotal)*100, 2) : 0;
      const qcolor = getQueueColor(q);
      const row  = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `<div class="q-label">${q}</div><div class="q-bar"><div class="q-fill"></div></div><div class="q-val">${val}</div>`;
      block.appendChild(row);
      const fill  = row.querySelector(".q-fill");
      const qlbl  = row.querySelector(".q-label");
      qlbl.style.color     = qcolor;
      fill.style.background = toBarFill(qcolor);
      fill.style.boxShadow  = `0 0 8px ${qcolor}88`;
      requestAnimationFrame(() => requestAnimationFrame(() => { fill.style.width = pct + "%"; }));
    });
    container.appendChild(block);
  });
}

/* ===== HOUR BREAKDOWN ===== */
function showHourBreakdown(hour, dataOverride=null) {
  const data = dataOverride || lastData;
  if (!data) return;
  currentHour = hour;
  const container = document.getElementById("container");
  const label = document.getElementById("queueHourLabel");
  if (label) label.textContent = hour;
  container.innerHTML = `<div class="dept-header" style="color:var(--accent)"><span>Hour breakdown</span><span class="reset-link" onclick="resetView()">← All</span></div>`;

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
    const deptTotal = rows.reduce((s,r)=>s+r.total,0);
    const color = getDeptColor(dept);
    const block = document.createElement("div");
    block.className = "queue-dept";
    block.innerHTML = `<div class="dept-header" style="color:${color}"><span>${dept}</span><span class="dept-count">${deptTotal.toLocaleString()}</span></div>`;
    rows.forEach(r => {
      const qcolor = getQueueColor(r.q);
      const pct = Math.max((r.total/deptTotal)*100, 4);
      const row = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `<div class="q-label">${r.q}</div><div class="q-bar"><div class="q-fill"></div></div><div class="q-val">${r.total}</div>`;
      block.appendChild(row);
      const fill = row.querySelector(".q-fill");
      row.querySelector(".q-label").style.color = qcolor;
      fill.style.background = toBarFill(qcolor);
      fill.style.boxShadow  = `0 0 8px ${qcolor}88`;
      fill.style.width = pct + "%";
    });
    container.appendChild(block);
  });
}

function resetView() {
  currentHour = null;
  if (lastData) renderQueueBars(lastData, lastData._total || 0);
}

/* ===== COMPARISON CHART (Today vs Yesterday) ===== */
function buildComparisonChart(todayData, yesterdayData) {
  const canvas = document.getElementById("comparisonChart");
  if (!canvas) return;

  // Update title with real dates
  const titleEl = document.getElementById("compareTitle");
  if (titleEl) {
    const base = currentDate || getLocalDate();
    const [ty, tm, td] = base.split("-").map(Number);
    const tDate = new Date(ty, tm - 1, td);
    const yDate = new Date(ty, tm - 1, td - 1);
    const fmt = d => d.toLocaleDateString([], { month:"short", day:"numeric" });
    titleEl.textContent = `${fmt(tDate)} vs ${fmt(yDate)}`;
  }

  // Update legend labels with real dates
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
    compChart.data.labels                      = depts;
    compChart.data.datasets[0].data            = todayVals;
    compChart.data.datasets[0].label           = tLbl;
    compChart.data.datasets[1].data            = yesterVals;
    compChart.data.datasets[1].label           = yLbl;
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
          data: todayVals,
          backgroundColor: "#38bdf8bb",
          borderRadius: 5,
          borderSkipped: false
        },
        {
          label: (()=>{ const b=currentDate||getLocalDate(); const [y,m,d]=b.split("-").map(Number); return new Date(y,m-1,d-1).toLocaleDateString([],{month:"short",day:"numeric"}); })(),
          data: yesterVals,
          backgroundColor: "rgba(56,189,248,0.22)",
          borderRadius: 5,
          borderSkipped: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration:600, easing:"easeOutQuart" },
      plugins: {
        legend: { display:false },
        tooltip: {
          backgroundColor:"rgba(6,13,26,.95)",
          borderColor:"rgba(56,189,248,.25)",
          borderWidth:1,
          titleColor:"#9ed8ff",
          bodyColor:"#e8f2ff",
          titleFont:{ family:"'Space Mono',monospace", size:11 },
          bodyFont: { family:"'DM Sans',sans-serif",   size:12 },
          padding:10,
          callbacks: {
            label: item => ` ${item.dataset.label}: ${item.raw.toLocaleString()}`
          }
        }
      },
      scales: {
        x: { grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10}} },
        y: { grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

/* ===== HOURLY CHART ===== */
function buildHourlyChart(data) {
  const canvas = document.getElementById("hourlyChart");
  if (!canvas) return;

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
    if (r >= 1)  return "rgba(120,255,190,.92)";
    if (r < .2)  return "rgba(90,150,255,.28)";
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
        { type:"bar",  data:values, backgroundColor:colors, borderRadius:4, borderSkipped:false, order:2 },
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
          bodyFont:{family:"'DM Sans',sans-serif",size:12},
          padding:10,
          filter:item=>item.datasetIndex===0,
          callbacks:{ title:i=>i[0].label, label:i=>` ${i.raw.toLocaleString()} jobs` }
        }
      },
      scales:{
        x:{ grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},maxRotation:45} },
        y:{ grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

/* ===== CHART INSIGHT ===== */
function setChartInsight(data) {
  const el = document.getElementById("chartInsight");
  if (!el) return;
  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });
  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    Object.keys(data[dept]).forEach(q => {
      if (q === "_total") return;
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
async function loadTrend(days=14) {
  document.querySelectorAll(".trend-btn").forEach(b => {
    b.classList.toggle("active", b.textContent.trim()===`${days}D`);
  });
  try {
    const json = await apiFetch(`${API_URL}?mode=trend&days=${days}`);
    if (!json.success) return;
    buildTrendChart(json.trend);
    setTrendInsight(json.trend);
  } catch(err) { console.error("Trend error:", err); }
}

function buildTrendChart(trend) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;
  if (trendChart) { trendChart.destroy(); trendChart=null; }
  const labels   = trend.map(d=>d.date);
  const totals   = trend.map(d=>d.total);
  const finish   = trend.map(d=>d.departments?.["Finish"]    ||0);
  const surface  = trend.map(d=>d.departments?.["Surface"]   ||0);
  const specialty= trend.map(d=>d.departments?.["Specialty"] ||0);
  const frame    = trend.map(d=>d.departments?.["Frame Only"]||0);
  const ctx = canvas.getContext("2d");
  const grad = ctx.createLinearGradient(0,0,0,300);
  grad.addColorStop(0,"rgba(56,189,248,.20)"); grad.addColorStop(1,"rgba(56,189,248,0)");
  trendChart = new Chart(ctx, {
    type:"line",
    data:{ labels, datasets:[
      { label:"Total",    data:totals,   borderColor:"#38bdf8", backgroundColor:grad, borderWidth:2.5, tension:.4, fill:true, pointRadius:3, pointHoverRadius:6 },
      { label:"Finish",   data:finish,   borderColor:"#22d3ee", borderWidth:1.5, tension:.4, borderDash:[5,5], pointRadius:2, fill:false },
      { label:"Surface",  data:surface,  borderColor:"#38bdf8", borderWidth:1.5, tension:.4, borderDash:[5,5], pointRadius:2, fill:false, borderDashOffset:4 },
      { label:"Specialty",data:specialty,borderColor:"#a78bfa", borderWidth:1.5, tension:.4, borderDash:[5,5], pointRadius:2, fill:false },
      { label:"Frame",    data:frame,    borderColor:"#fb923c", borderWidth:1.5, tension:.4, borderDash:[5,5], pointRadius:2, fill:false }
    ]},
    options:{
      responsive:true, maintainAspectRatio:false, animation:{duration:500},
      plugins:{
        legend:{ display:true, labels:{color:"#6a8ab0",font:{family:"'DM Sans',sans-serif",size:12},boxWidth:12,padding:16,usePointStyle:true} },
        tooltip:{ backgroundColor:"rgba(6,13,26,.95)",borderColor:"rgba(56,189,248,.25)",borderWidth:1,titleColor:"#9ed8ff",bodyColor:"#e8f2ff",padding:10 }
      },
      scales:{
        x:{ grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},maxRotation:45} },
        y:{ grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

function setTrendInsight(trend) {
  const el = document.getElementById("trendInsight");
  if (!el||trend.length<2) return;
  const peak   = trend.reduce((b,d)=>d.total>b.total?d:b,trend[0]);
  const latest = trend[trend.length-1];
  const prev   = trend[trend.length-2];
  const chg    = prev.total>0?((latest.total-prev.total)/prev.total*100).toFixed(1):"0.0";
  el.textContent = `Peak ${peak.date} (${peak.total.toLocaleString()})  ${chg>0?"↑":"↓"} ${Math.abs(chg)}%`;
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

/* Pre-fill compare inputs with recent dates */
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

/* ===== RUN COMPARISON ===== */
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
    // Fetch all dates in parallel, use cache when available
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

/* ===== FORMAT DATE LABEL ===== */
function fmtDate(dateStr) {
  const [y,m,d] = dateStr.split("-").map(Number);
  return new Date(y,m-1,d).toLocaleDateString([],{month:"short",day:"numeric"});
}

/* ===== RENDER KPI COMPARISON ===== */
function renderCompKPI(results) {
  const grid = document.getElementById("compKpiGrid");
  grid.innerHTML = "";

  // Get all dept names
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

    const maxVal = Math.max(...rows.map(r => r.val), 1);
    const base   = rows[0].val; // first selected date as baseline

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

/* ===== RENDER MULTI-DATE CHART ===== */
function renderCompChart(results) {
  const canvas    = document.getElementById("compMultiChart");
  const chartCard = document.getElementById("compChartCard");
  const insight   = document.getElementById("compInsight");
  if (!canvas) return;

  chartCard.style.display = "block";

  const depts = [...new Set(results.flatMap(r => Object.keys(r.data).filter(k => k !== "_total")))];

  // Find peak date
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
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration:600, easing:"easeOutQuart" },
      plugins: {
        legend: {
          display: true,
          labels: { color:"#6a8ab0", font:{family:"'DM Sans',sans-serif",size:12}, boxWidth:12, padding:14, usePointStyle:true }
        },
        tooltip: {
          backgroundColor:"rgba(6,13,26,.95)", borderColor:"rgba(56,189,248,.25)", borderWidth:1,
          titleColor:"#9ed8ff", bodyColor:"#e8f2ff",
          titleFont:{family:"'Space Mono',monospace",size:11},
          bodyFont:{family:"'DM Sans',sans-serif",size:12}, padding:10,
          callbacks: { label: item => ` ${item.dataset.label}: ${item.raw.toLocaleString()}` }
        }
      },
      scales: {
        x: { grid:{color:"rgba(90,160,255,.05)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10}} },
        y: { grid:{color:"rgba(90,160,255,.07)"}, ticks:{color:"#4a6a90",font:{family:"'Space Mono',monospace",size:10},callback:v=>v.toLocaleString()} }
      }
    }
  });
}

/* ===== RENDER PER-DEPT BREAKDOWN ===== */
function renderCompDeptGrid(results) {
  const grid = document.getElementById("compDeptGrid");
  grid.innerHTML = "";

  const depts = [...new Set(results.flatMap(r => Object.keys(r.data).filter(k => k !== "_total")))];

  depts.forEach(dept => {
    const deptColor = getDeptColor(dept);

    // Get all queues across all dates
    const allQueues = [...new Set(
      results.flatMap(r => Object.keys(r.data[dept] || {}).filter(k => k !== "_total"))
    )];

    if (!allQueues.length) return;

    const card = document.createElement("div");
    card.className = "comp-dept-card";
    card.style.setProperty("--dkc", deptColor);

    const deptTotals = results.map(r => r.data[dept]?._total ?? 0);
    const maxDeptTotal = Math.max(...deptTotals, 1);

    let qRowsHtml = allQueues.map(q => {
      const vals = results.map(r => r.data[dept]?.[q]?.total ?? 0);
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

    // Dept totals row
    const deptTotalsHtml = results.map((r, i) => `
      <div class="comp-kpi-row" style="padding:3px 0">
        <div class="comp-kpi-date">
          <div class="comp-kpi-date-dot" style="background:${r.color}"></div>
          ${fmtDate(r.date)}
        </div>
        <div class="comp-kpi-num">${(r.data[dept]?._total ?? 0).toLocaleString()}</div>
      </div>`).join("");

    card.innerHTML = `
      <div class="comp-dept-title" style="color:${deptColor}">${dept}</div>
      ${deptTotalsHtml}
      <div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--border)">${qRowsHtml}</div>
    `;
    grid.appendChild(card);
  });
}