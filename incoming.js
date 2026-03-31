/* ============================================================
   INCOMING JOBS DASHBOARD — VISUAL UPGRADE v3
   ============================================================ */

const API_URL = "https://script.google.com/macros/s/AKfycbyI2YqO9wXZ4-v34OXqqxD-yCYe3Dnly1-d_cf9mYtkcVkZoXeWhgQ8u6WbgY-lvIQQjg/exec";

let chart        = null;
let trendChart   = null;
let lastData     = null;
let currentHour  = null;
let currentDate  = null;
let isRenderingChart = false;
let sparklineCharts  = {};

// Track previous totals for alert detection
let prevDeptTotals = {};

/* ================= DEPT CONFIG ================= */

const DEPT_COLORS = {
  "Finish":     "#22c55e",
  "Speciality": "#a855f7",
  "Specialty":  "#a855f7",
  "Surface":    "#00e5ff",
  "Frame Only": "#f97316"
};

/* ================= QUEUE COLORS ================= */

const QUEUE_COLORS = {
  "Surface Rush Delivery":          "#00E5FF",
  "Rush Delivery Queue":            "#00E5FF",
  "In Rush Delivery Queue":         "#00E5FF",
  "Surface China Rush Delivery":    "#F97316",
  "China Rush Queue":               "#F97316",
  "In China Rush Queue":            "#F97316",
  "In China Rush Delivery Queue":   "#F97316",
  "Surface Standard":               "#3B82F6",
  "Standard Queue":                 "#3B82F6",
  "In Standard Queue":              "#3B82F6",
  "Fin Standard":                   "#93C5FD",
  "Overnight Queue":                "#FACC15",
  "In Overnight Queue":             "#FACC15",
  "Surface Overnight Delivery":     "#FACC15",
  "Fin Overnight Delivery":         "#A3A3A3",
  "Fin Rush Delivery":              "#86EFAC",
  "Beast Queue":                    "#A855F7",
  "In Beast Queue":                 "#A855F7",
  "Beast Surface Queue":            "#D946EF",
  "Beast Finish Queue":             "#14B8A6",
  "Frame Only Queue":               "#22C55E",
  "In Frame Only Queue":            "#22C55E",
  "Frame Only Test Queue":          "#15803D",
  "Test Jobs Queue":                "#2DD4BF",
  "Echo Queue":                     "#FB923C",
  "All Queued Jobs":                "#10B981"
};

const HOUR_ORDER = [
  "12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM",
  "6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM",
  "12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM",
  "6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"
];

/* ================= CACHE ================= */

const DATA_CACHE   = {};
const CACHE_TTL_MS = 55_000;

function cacheKey() {
  return currentDate || new Date().toISOString().split("T")[0];
}

function getCached() {
  const entry = DATA_CACHE[cacheKey()];
  if (!entry) return null;
  const age     = Date.now() - entry.ts;
  const isToday = cacheKey() === new Date().toISOString().split("T")[0];
  if (!isToday && age < 24 * 60 * 60 * 1000) return entry; // historical: cache 24h
  if (isToday  && age < 5_000) return entry;                // today: debounce 5s
  return null;
}

function setCache(data, total) {
  DATA_CACHE[cacheKey()] = { data, total, ts: Date.now() };
}

/* ================= SKELETON LOADER ================= */

function showSkeleton() {
  const grid = document.getElementById("kpiGrid");
  if (!grid || grid.querySelector(".kpi-skeleton")) return;
  grid.innerHTML = "";
  for (let i = 0; i < 4; i++) {
    const el = document.createElement("div");
    el.className = "kpi-card kpi-skeleton";
    el.innerHTML = `
      <div class="skel skel-title"></div>
      <div class="skel skel-value"></div>
      <div class="skel skel-sub"></div>
    `;
    grid.appendChild(el);
  }
}

/* ================= INIT ================= */

document.addEventListener("DOMContentLoaded", () => {
  const today = new Date().toISOString().split("T")[0];
  document.getElementById("dateFilter").value = today;
  currentDate = today;

  startLiveClock();
  showSkeleton();
  loadData();

  // Auto-refresh every 60s on today
  setInterval(() => {
    const isToday = !currentDate || currentDate === new Date().toISOString().split("T")[0];
    if (isToday) loadData();
  }, 60_000);
});

/* ================= LIVE CLOCK ================= */

function startLiveClock() {
  const el = document.getElementById("lastUpdated");
  if (!el) return;

  function tick() {
    el.textContent = new Date().toLocaleTimeString([], {
      hour: "2-digit", minute: "2-digit", second: "2-digit"
    });
  }

  tick();
  setInterval(tick, 1000);
}

/* ================= TAB SWITCH ================= */

function switchTab(tab) {
  const incomingTab = document.getElementById("incomingTab");
  const trendTab    = document.getElementById("trendTab");
  const btnIncoming = document.getElementById("tabIncomingBtn");
  const btnTrend    = document.getElementById("tabTrendBtn");

  if (!incomingTab || !trendTab) return;

  btnIncoming?.classList.remove("active");
  btnTrend?.classList.remove("active");

  if (tab === "incoming") {
    incomingTab.style.display = "block";
    trendTab.style.display    = "none";
    btnIncoming?.classList.add("active");
  } else {
    incomingTab.style.display = "none";
    trendTab.style.display    = "block";
    btnTrend?.classList.add("active");
    if (!trendChart) loadTrend(14);
  }
}

/* ================= COLOR HELPERS ================= */

function getQueueColor(queueName) {
  if (QUEUE_COLORS[queueName]) return QUEUE_COLORS[queueName];
  const q = String(queueName || "").toLowerCase();
  if (q.includes("error") || q.includes("waiting")) return "#f87171";
  if (q.includes("hold"))         return "#fb7185";
  if (q.includes("surface rush")) return "#00E5FF";
  if (q.includes("china rush"))   return "#f87171";
  if (q.includes("standard"))     return "#3B82F6";
  if (q.includes("overnight"))    return "#FACC15";
  if (q.includes("beast"))        return "#A855F7";
  if (q.includes("frame only"))   return "#22C55E";
  if (q.includes("test"))         return "#2DD4BF";
  if (q.includes("echo"))         return "#FB923C";
  return "#6fb3ff";
}

function getDeptColor(dept) {
  return DEPT_COLORS[dept] || "#5ab4ff";
}

function brightenColor(hex, amt = 40) {
  const n = parseInt(hex.replace("#",""), 16);
  const r = Math.min(255, (n >> 16) + amt);
  const g = Math.min(255, ((n >> 8) & 0xff) + amt);
  const b = Math.min(255, (n & 0xff) + amt);
  return `rgb(${r},${g},${b})`;
}

function toBarFill(color) {
  return `linear-gradient(90deg, ${color}cc, ${brightenColor(color, 40)})`;
}

function toGlow(color, strength = 10) {
  return `0 0 ${strength}px ${color}99`;
}

/* ================= DATA LOAD ================= */

// Timeout wrapper — aborts fetch if server takes too long
async function fetchWithTimeout(url, timeoutMs = 12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, { signal: controller.signal });
    clearTimeout(timer);
    return res;
  } catch (err) {
    clearTimeout(timer);
    throw err;
  }
}

// Retry wrapper — retries up to N times with delay between attempts
async function fetchWithRetry(url, retries = 2, delayMs = 1500) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const res = await fetchWithTimeout(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res;
    } catch (err) {
      const isLast = attempt === retries;
      if (isLast) throw err;
      console.warn(`⚠️ Attempt ${attempt + 1} failed, retrying in ${delayMs}ms…`);
      await new Promise(r => setTimeout(r, delayMs));
    }
  }
}

// Error banner — floats at bottom, auto-hides after 5s
function showErrorBanner(message) {
  let banner = document.getElementById("errorBanner");
  if (!banner) {
    banner = document.createElement("div");
    banner.id = "errorBanner";
    banner.style.cssText = [
      "position:fixed","bottom:24px","left:50%","transform:translateX(-50%)",
      "background:rgba(248,113,113,0.12)","border:1px solid rgba(248,113,113,0.38)",
      "color:#fca5a5","padding:10px 22px","border-radius:10px",
      "font-family:'Space Mono',monospace","font-size:12px","letter-spacing:0.04em",
      "z-index:9999","backdrop-filter:blur(10px)",
      "box-shadow:0 0 16px rgba(248,113,113,0.28)",
      "transition:opacity 0.5s ease","pointer-events:none"
    ].join(";");
    document.body.appendChild(banner);
  }
  banner.textContent = `⚠  ${message}`;
  banner.style.opacity = "1";
  clearTimeout(banner._t);
  banner._t = setTimeout(() => { banner.style.opacity = "0"; }, 5000);
}

// Central render helper — used by loadData for both fresh + cached paths
function renderAll(data, total) {
  const totalEl = document.getElementById("totalJobs");
  if (totalEl) totalEl.textContent = total.toLocaleString();

  renderKPI(data, total);

  if (currentHour) {
    showHourBreakdown(currentHour, data);
  } else {
    renderQueueBars(data, total);
  }

  buildChart(data);
  setChartInsight(data);
  lastData = data;
}

async function loadData() {
  if (window._isLoading) return;
  window._isLoading = true;
  updateRefreshStatus("updating");

  // Serve cache instantly — no spinner, no flash
  const cached = getCached();
  if (cached) {
    renderAll(cached.data, cached.total);
    window._isLoading = false;
    updateRefreshStatus();
    return;
  }

  try {
    const url  = currentDate ? `${API_URL}?date=${currentDate}` : API_URL;
    const res  = await fetchWithRetry(url);
    const json = await res.json();

    if (!json.success) throw new Error("API returned success=false");

    setCache(json.data, json.total);
    renderAll(json.data, json.total);

  } catch (err) {
    console.error("❌ loadData:", err);

    // Fallback: show stale data rather than a blank screen
    const stale = DATA_CACHE[cacheKey()];
    if (stale) {
      console.warn("📦 Showing stale cache due to error");
      renderAll(stale.data, stale.total);
      showErrorBanner("Using cached data — server unreachable");
    } else {
      showErrorBanner("Failed to load data. Will retry automatically.");
    }
  } finally {
    window._isLoading = false;
    updateRefreshStatus();
  }
}

/* ================= STATUS ================= */

function updateRefreshStatus(state = "auto") {
  const el = document.getElementById("refreshStatus");
  if (!el) return;

  const today = new Date().toISOString().split("T")[0];
  let mode;
  if (state === "updating") {
    mode = "updating";
  } else if (!currentDate || currentDate === today) {
    mode = "live";
  } else {
    mode = "history";
  }

  el.className = `status-badge ${mode}`;
  el.textContent = { live: "Live", history: "History", updating: "Updating…" }[mode];
}

/* ================= DATE FILTER ================= */

function applyDateFilter() {
  const val = document.getElementById("dateFilter").value;
  if (!val) return;
  currentDate = new Date(val).toISOString().split("T")[0];
  currentHour = null;
  updateRefreshStatus();
  loadData();
}

function resetDateFilter() {
  currentDate = null;
  currentHour = null;
  document.getElementById("dateFilter").value =
    new Date().toISOString().split("T")[0];
  updateRefreshStatus();
  loadData();
}

/* ================= KPI CARDS ================= */

function renderKPI(data, grandTotal) {
  const grid = document.getElementById("kpiGrid");
  if (!grid) return;

  // Destroy old sparklines
  Object.values(sparklineCharts).forEach(c => c?.destroy());
  sparklineCharts = {};

  grid.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const value    = data[dept]._total ?? 0;
    const pct      = grandTotal > 0 ? ((value / grandTotal) * 100).toFixed(1) : "0.0";
    const prevVal  = prevDeptTotals[dept] ?? null;
    const isAlert  = prevVal !== null && value > prevVal * 1.25; // 25% spike = alert
    const deptColor = getDeptColor(dept);

    const card = document.createElement("div");
    card.className = "kpi-card" + (isAlert ? " alert" : "");
    card.dataset.dept = dept;

    card.innerHTML = `
      <div class="kpi-title">${dept}</div>
      ${isAlert ? `<div class="kpi-alert-badge">↑ Spike</div>` : ""}
      <div class="kpi-value" id="kpi-val-${dept.replace(/\s/g,"-")}">0</div>
      <div class="kpi-sub">${pct}% of total</div>
      <canvas class="kpi-sparkline" id="spark-${dept.replace(/\s/g,"-")}"></canvas>
    `;

    grid.appendChild(card);

    // Animate number roll-up
    animateCount(
      document.getElementById(`kpi-val-${dept.replace(/\s/g,"-")}`),
      value
    );

    // Draw sparkline from hourly data
    drawSparkline(
      `spark-${dept.replace(/\s/g,"-")}`,
      buildDeptHourlyArray(data[dept]),
      deptColor
    );

    prevDeptTotals[dept] = value;
  });
}

/* ================= COUNT ANIMATION ================= */

function animateCount(el, target, duration = 700) {
  if (!el) return;
  const start = performance.now();
  const from  = 0;

  function step(now) {
    const p = Math.min((now - start) / duration, 1);
    const ease = 1 - Math.pow(1 - p, 3); // cubic ease-out
    el.textContent = Math.round(from + (target - from) * ease).toLocaleString();
    if (p < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ================= BUILD HOURLY ARRAY FOR DEPT ================= */

function buildDeptHourlyArray(deptData) {
  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });

  Object.keys(deptData).forEach(queue => {
    if (queue === "_total") return;
    Object.keys(deptData[queue].hours || {}).forEach(hour => {
      deptData[queue].hours[hour].forEach(e => {
        hourMap[hour] = (hourMap[hour] || 0) + e.value;
      });
    });
  });

  return HOUR_ORDER.map(h => hourMap[h]);
}

/* ================= SPARKLINE ================= */

function drawSparkline(canvasId, data, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const nonZero = data.filter(v => v > 0);
  if (!nonZero.length) return;

  // Force canvas size before Chart.js touches it
  canvas.width  = canvas.parentElement?.offsetWidth || 200;
  canvas.height = 40;

  const ctx = canvas.getContext("2d");

  const grad = ctx.createLinearGradient(0, 0, 0, 40);
  grad.addColorStop(0, color + "55");
  grad.addColorStop(1, color + "00");

  const sparkChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: HOUR_ORDER,
      datasets: [{
        data,
        borderColor: color,
        backgroundColor: grad,
        borderWidth: 1.5,
        tension: 0.4,
        fill: true,
        pointRadius: 0,
        pointHoverRadius: 0
      }]
    },
    options: {
      responsive: false,
      animation: false,
      layout: { padding: 0 },
      plugins: { legend: { display: false }, tooltip: { enabled: false } },
      scales: {
        x: { display: false },
        y: { display: false, min: 0 }
      }
    }
  });

  sparklineCharts[canvasId] = sparkChart;
}

/* ================= QUEUE BARS ================= */

function renderQueueBars(data, grandTotal) {
  const container = document.getElementById("container");
  if (!container) return;
  container.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const deptTotal  = data[dept]._total ?? 0;
    const deptColor  = getDeptColor(dept);
    const deptBlock  = document.createElement("div");
    deptBlock.className = "queue-dept";
    deptBlock.innerHTML = `
      <div class="dept-header" style="color:${deptColor}">
        <span>${dept}</span>
        <span>${deptTotal.toLocaleString()}</span>
      </div>
    `;

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      const q     = data[dept][queue];
      const pct   = deptTotal > 0 ? Math.max((q.total / deptTotal) * 100, 2) : 0;
      const color = getQueueColor(queue);

      const row = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `
        <div class="label">${queue}</div>
        <div class="bar"><div class="fill"></div></div>
        <div class="value">${q.total}</div>
      `;
      deptBlock.appendChild(row);

      const fill  = row.querySelector(".fill");
      const label = row.querySelector(".label");

      label.style.color     = color;
      fill.style.background = toBarFill(color);
      fill.style.boxShadow  = toGlow(color);

      requestAnimationFrame(() => {
        requestAnimationFrame(() => { fill.style.width = pct + "%"; });
      });
    });

    container.appendChild(deptBlock);
  });
}

/* ================= HOUR BREAKDOWN ================= */

function showHourBreakdown(hour, dataOverride = null) {
  const data = dataOverride || lastData;
  if (!data) return;

  currentHour = hour;
  const container = document.getElementById("container");
  container.innerHTML = `
    <div class="dept-header" style="color:var(--accent)">
      <span>Hour: ${hour}</span>
      <span class="reset-link" onclick="resetView()">← All Hours</span>
    </div>
  `;

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const rows = [];
    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;
      const entries = data[dept][queue].hours?.[hour] || [];
      const total   = entries.reduce((s, x) => s + x.value, 0);
      if (total > 0) rows.push({ queue, total });
    });

    rows.sort((a, b) => b.total - a.total);
    if (!rows.length) return;

    const deptTotal = rows.reduce((s, r) => s + r.total, 0);
    const deptColor = getDeptColor(dept);

    const block = document.createElement("div");
    block.className = "queue-dept";
    block.innerHTML = `
      <div class="dept-header" style="color:${deptColor}">
        <span>${dept}</span>
        <span>${deptTotal.toLocaleString()}</span>
      </div>
    `;

    rows.forEach(r => {
      const color = getQueueColor(r.queue);
      const pct   = Math.max((r.total / deptTotal) * 100, 4);

      const row = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `
        <div class="label">${r.queue}</div>
        <div class="bar"><div class="fill"></div></div>
        <div class="value">${r.total}</div>
      `;
      block.appendChild(row);

      const fill  = row.querySelector(".fill");
      const label = row.querySelector(".label");
      label.style.color     = color;
      fill.style.background = toBarFill(color);
      fill.style.boxShadow  = toGlow(color);
      fill.style.width      = pct + "%";
    });

    container.appendChild(block);
  });
}

/* ================= RESET VIEW ================= */

function resetView() {
  currentHour = null;
  if (lastData) renderQueueBars(lastData, lastData._total || 0);
}

/* ================= HOURLY CHART ================= */

function buildChart(data) {
  if (isRenderingChart) return;
  isRenderingChart = true;

  try {
    const canvas = document.getElementById("hourlyChart");
    if (!canvas) return;

    const hourMap = {};
    HOUR_ORDER.forEach(h => { hourMap[h] = 0; });

    Object.keys(data).forEach(dept => {
      if (dept === "_total") return;
      Object.keys(data[dept]).forEach(queue => {
        if (queue === "_total") return;
        Object.keys(data[dept][queue].hours || {}).forEach(hour => {
          data[dept][queue].hours[hour].forEach(e => {
            hourMap[hour] = (hourMap[hour] || 0) + e.value;
          });
        });
      });
    });

    const labels = HOUR_ORDER.filter(h => hourMap[h] > 0);
    const values = labels.map(h => hourMap[h]);
    const maxVal = Math.max(...values, 1);
    const avg    = values.reduce((a, b) => a + b, 0) / (values.length || 1);

    // Bar colors
    const colors = values.map(v => {
      const ratio = v / maxVal;
      if (ratio >= 1)   return "rgba(120,255,190,0.92)";
      if (ratio < 0.2)  return "rgba(90,150,255,0.28)";
      return "rgba(90,182,255,0.70)";
    });

    // Average line dataset
    const avgData = values.map(() => Math.round(avg));

    if (chart) {
      chart.data.labels                      = labels;
      chart.data.datasets[0].data            = values;
      chart.data.datasets[0].backgroundColor = colors;
      chart.data.datasets[1].data            = avgData;
      chart.update("none");
    } else {
      const ctx = canvas.getContext("2d");

      chart = new Chart(ctx, {
        type: "bar",
        data: {
          labels,
          datasets: [
            {
              type: "bar",
              data: values,
              backgroundColor: colors,
              borderRadius: 5,
              borderSkipped: false,
              order: 2
            },
            {
              type: "line",
              label: "Avg",
              data: avgData,
              borderColor: "rgba(250,204,21,0.6)",
              borderWidth: 1.5,
              borderDash: [6,4],
              pointRadius: 0,
              fill: false,
              tension: 0,
              order: 1
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: { duration: 600, easing: "easeOutQuart" },
          onClick: (evt, elements, chartInst) => {
            const pts = chartInst.getElementsAtEventForMode(evt, "index", { intersect: false }, true);
            if (!pts.length) return;
            showHourBreakdown(chartInst.data.labels[pts[0].index]);
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: "rgba(5,13,26,0.94)",
              borderColor: "rgba(90,180,255,0.28)",
              borderWidth: 1,
              titleColor: "#9ed8ff",
              bodyColor: "#e8f2ff",
              titleFont: { family: "'Space Mono', monospace", size: 11 },
              bodyFont:  { family: "'DM Sans', sans-serif",   size: 13 },
              padding: 10,
              filter: item => item.datasetIndex === 0,
              callbacks: {
                title: items => items[0].label,
                label: item  => ` ${item.raw.toLocaleString()} jobs`
              }
            }
          },
          scales: {
            x: {
              grid: { color: "rgba(90,160,255,0.05)" },
              ticks: {
                color: "#4a6a90",
                font: { family: "'Space Mono', monospace", size: 10 },
                maxRotation: 45
              }
            },
            y: {
              grid: { color: "rgba(90,160,255,0.07)" },
              ticks: {
                color: "#4a6a90",
                font: { family: "'Space Mono', monospace", size: 10 },
                callback: v => v.toLocaleString()
              }
            }
          }
        }
      });
    }

  } catch (err) {
    console.error("Chart error:", err);
  } finally {
    setTimeout(() => { isRenderingChart = false; }, 80);
  }
}

/* ================= CHART INSIGHT ================= */

function setChartInsight(data) {
  const el = document.getElementById("chartInsight");
  if (!el) return;

  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;
    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;
      Object.keys(data[dept][queue].hours || {}).forEach(hour => {
        data[dept][queue].hours[hour].forEach(e => {
          hourMap[hour] = (hourMap[hour] || 0) + e.value;
        });
      });
    });
  });

  let peak = null, val = 0;
  Object.keys(hourMap).forEach(h => {
    if (hourMap[h] > val) { val = hourMap[h]; peak = h; }
  });

  el.textContent = peak ? `Peak: ${peak} (${val.toLocaleString()})` : "";
}

/* ================= TREND ================= */

async function loadTrend(days = 14) {
  document.querySelectorAll(".trend-btn").forEach(b => {
    b.classList.toggle("active", b.textContent.trim() === `${days}D`);
  });

  try {
    const res  = await fetch(`${API_URL}?mode=trend&days=${days}`);
    const json = await res.json();
    if (!json.success) return;
    buildTrendChart(json.trend);
    setTrendInsight(json.trend);
  } catch (err) {
    console.error("Trend error:", err);
  }
}

/* ================= TREND CHART ================= */

function buildTrendChart(trend) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;

  if (trendChart) { trendChart.destroy(); trendChart = null; }

  const labels    = trend.map(d => d.date);
  const totals    = trend.map(d => d.total);
  const finish    = trend.map(d => d.departments?.["Finish"]     || 0);
  const surface   = trend.map(d => d.departments?.["Surface"]    || 0);
  const specialty = trend.map(d => d.departments?.["Specialty"]  || 0);
  const frame     = trend.map(d => d.departments?.["Frame Only"] || 0);

  const ctx = canvas.getContext("2d");

  // Gradient fill under Total line
  const grad = ctx.createLinearGradient(0, 0, 0, 300);
  grad.addColorStop(0, "rgba(90,180,255,0.22)");
  grad.addColorStop(1, "rgba(90,180,255,0.00)");

  trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: totals,
          borderColor: "#5ab4ff",
          backgroundColor: grad,
          borderWidth: 2.5,
          tension: 0.4,
          fill: true,
          pointBackgroundColor: "#5ab4ff",
          pointRadius: 3,
          pointHoverRadius: 6
        },
        {
          label: "Finish",
          data: finish,
          borderColor: "#22c55e",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2,
          fill: false
        },
        {
          label: "Surface",
          data: surface,
          borderColor: "#00e5ff",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2,
          fill: false
        },
        {
          label: "Specialty",
          data: specialty,
          borderColor: "#a855f7",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2,
          fill: false
        },
        {
          label: "Frame",
          data: frame,
          borderColor: "#f97316",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 500 },
      plugins: {
        legend: {
          display: true,
          labels: {
            color: "#6a8ab0",
            font: { family: "'DM Sans', sans-serif", size: 12 },
            boxWidth: 12,
            padding: 16,
            usePointStyle: true
          }
        },
        tooltip: {
          backgroundColor: "rgba(5,13,26,0.94)",
          borderColor: "rgba(90,180,255,0.28)",
          borderWidth: 1,
          titleColor: "#9ed8ff",
          bodyColor: "#e8f2ff",
          titleFont: { family: "'Space Mono', monospace", size: 11 },
          bodyFont:  { family: "'DM Sans', sans-serif",   size: 12 },
          padding: 10
        }
      },
      scales: {
        x: {
          grid: { color: "rgba(90,160,255,0.05)" },
          ticks: {
            color: "#4a6a90",
            font: { family: "'Space Mono', monospace", size: 10 },
            maxRotation: 45
          }
        },
        y: {
          grid: { color: "rgba(90,160,255,0.07)" },
          ticks: {
            color: "#4a6a90",
            font: { family: "'Space Mono', monospace", size: 10 },
            callback: v => v.toLocaleString()
          }
        }
      }
    }
  });
}

/* ================= TREND INSIGHT ================= */

function setTrendInsight(trend) {
  const el = document.getElementById("trendInsight");
  if (!el || trend.length < 2) return;

  const peak   = trend.reduce((best, d) => d.total > best.total ? d : best, trend[0]);
  const latest = trend[trend.length - 1];
  const prev   = trend[trend.length - 2];
  const change = prev.total > 0
    ? ((latest.total - prev.total) / prev.total * 100).toFixed(1)
    : "0.0";
  const icon = change > 0 ? "↑" : "↓";

  el.textContent = `Peak ${peak.date} (${peak.total.toLocaleString()})  ${icon} ${Math.abs(change)}%`;
}

/* ================= NAVIGATION ================= */

function goBack() {
  window.location.href = "index.html";
}