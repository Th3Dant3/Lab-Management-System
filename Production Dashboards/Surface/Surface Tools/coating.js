const API_URL = "https://script.google.com/macros/s/AKfycbxGEYOoJviGPzBSIX_Kh5X5ZAOAHFSyw9AO6ID-buDr1A82ERaPsauF3cSvSINr8VvT/exec";

console.log("Coating JS loaded");

/* =====================================================
   GLOBAL VARIABLES
===================================================== */

let trendChart   = null;
let reasonChart  = null;
let machineChart = null;
let flowChart    = null;

let dashboardProcessed = null;
let dashboardMachine   = null;
let dashboardData      = null;

let trendsLoaded    = false;
let currentMode     = "processed";
let currentFlowMode = "average";

let currentDate        = null;
let refreshInterval    = 5 * 60;
let refreshCountdown   = refreshInterval;
let refreshTimerHandle = null;

/* =====================================================
   ANIMATION UTILITIES
===================================================== */

// Animate a numeric KPI value counting up
function animateValue(el, targetVal, suffix, colorClass) {
  if (!el) return;
  const isFloat  = String(targetVal).includes(".");
  const start    = 0;
  const duration = 700;
  const startTs  = performance.now();

  function step(ts) {
    const progress = Math.min((ts - startTs) / duration, 1);
    const ease     = 1 - Math.pow(1 - progress, 3); // easeOutCubic
    const current  = start + (targetVal - start) * ease;
    el.textContent = (isFloat ? current.toFixed(2) : Math.round(current)) + (suffix || "");
    if (progress < 1) requestAnimationFrame(step);
    else el.textContent = (isFloat ? targetVal.toFixed(2) : targetVal) + (suffix || "");
  }
  requestAnimationFrame(step);

  if (colorClass) el.className = "kpi-value " + colorClass;
  el.style.animation = "none";
  requestAnimationFrame(() => {
    el.style.animation = "countUp 0.4s cubic-bezier(.22,1,.36,1) both";
  });
}

// Stagger-animate all KPI cards
function animateKpiCards() {
  document.querySelectorAll(".kpi-card").forEach((card, i) => {
    card.style.animation = "none";
    card.style.opacity   = "0";
    requestAnimationFrame(() => {
      setTimeout(() => {
        card.style.animation = `scaleIn 0.38s cubic-bezier(.22,1,.36,1) both`;
        card.style.opacity   = "";
      }, i * 45);
    });
  });
}

// Stagger-animate table rows
function animateTableRows() {
  document.querySelectorAll("#hourlyBody tr").forEach((tr, i) => {
    tr.style.animation = "none";
    tr.style.opacity   = "0";
    requestAnimationFrame(() => {
      setTimeout(() => {
        tr.style.animation = `rowSlideIn 0.3s cubic-bezier(.22,1,.36,1) both`;
        tr.style.opacity   = "";
      }, i * 40);
    });
  });
}

// Show a toast message at bottom
let _toastTimer = null;
function showToast(msg, color) {
  let toast = document.getElementById("historyToast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "historyToast";
    toast.className = "history-mode-toast";
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.style.background = color === "live"
    ? "rgba(32,160,145,0.12)"   : "rgba(251,191,36,0.12)";
  toast.style.borderColor  = color === "live"
    ? "rgba(32,160,145,0.35)"   : "rgba(251,191,36,0.35)";
  toast.style.color = color === "live" ? "#20a091" : "#fbbf24";

  clearTimeout(_toastTimer);
  toast.classList.remove("show");
  requestAnimationFrame(() => {
    requestAnimationFrame(() => { toast.classList.add("show"); });
  });
  _toastTimer = setTimeout(() => toast.classList.remove("show"), 3000);
}

// Flash the container when switching modes
function flashContainer(mode) {
  const container = document.querySelector(".container");
  if (!container) return;
  const cls = mode === "history" ? "history-flash" : "live-flash";
  container.classList.remove("history-flash", "live-flash");
  requestAnimationFrame(() => {
    requestAnimationFrame(() => container.classList.add(cls));
  });
  setTimeout(() => container.classList.remove(cls), 600);
}

/* =====================================================
   NEON NOIR CHART CONSTANTS
===================================================== */

const CHART_FONT = "'Inter', sans-serif";
const CHART_MONO = "'JetBrains Mono', monospace";
const CHART_ORB  = "'Orbitron', monospace";

const GLASS_TOOLTIP = {
  backgroundColor : "rgba(4, 6, 12, 0.97)",
  borderColor     : "rgba(0, 255, 200, 0.2)",
  borderWidth     : 1,
  titleColor      : "#00ffc8",
  bodyColor       : "rgba(0, 255, 200, 0.75)",
  footerColor     : "rgba(0, 255, 200, 0.4)",
  titleFont       : { family: CHART_FONT, size: 15, weight: "700" },
  bodyFont        : { family: CHART_MONO, size: 13 },
  padding         : 14,
  cornerRadius    : 4,
  displayColors   : true,
  boxWidth        : 10,
  boxHeight       : 10,
  boxPadding      : 3,
};

const GLASS_LEGEND = {
  labels: {
    color        : "#ffffff",
    font         : { family: CHART_FONT, size: 14, weight: "600" },
    usePointStyle: true,
    pointStyle   : "circle",
    padding      : 28,
    boxWidth     : 14,
    boxHeight    : 14,
  },
  position: "top",
};

const GLASS_GRID = {
  color     : "rgba(0, 255, 200, 0.04)",
  drawBorder: false,
  drawTicks : false,
};

// Glass color palette — rich but not harsh
const GC = {
  purple : "#9b7bff",
  red    : "#ff4757",
  teal   : "#00ffc8",
  orange : "#ff9f43",
  cyan   : "#00d4ff",
  green  : "#00ff88",
  yellow : "#ffd32a",
  blue   : "#00d4ff",
  violet : "#c084fc",
  pink   : "#f472b6",
};

const BREAKAGE_COLOR_MAP = {
  "S-HC Contamination"       : "#ff4d4d",
  "S-HC Pit"                 : "#9b72ff",
  "S-HC Run"                 : "#ff8c00",
  "S-HC Wagon Wheel"         : "#38bdf8",
  "S-HC Suction Cup Marks"   : "#00e5cc",
  "S-HC HC Suction Cup Marks": "#00e5cc",
};

/* =====================================================
   CHART HELPERS
===================================================== */

function formatMachineLabel(machine) {
  if (!machine) return "";
  return machine.replace(/^AR41-/, "44R1-");
}

function getFlowColor(minutes) {
  if (minutes <= 15) return GC.green;
  if (minutes <= 30) return GC.yellow;
  return GC.red;
}

// Glassy layered fill — rich at top, melts to nothing
function glassFill(ctx, color, h) {
  const height = h || 380;
  const g = ctx.createLinearGradient(0, 0, 0, height);
  g.addColorStop(0,    color + "3a");
  g.addColorStop(0.4,  color + "18");
  g.addColorStop(0.85, color + "06");
  g.addColorStop(1,    color + "00");
  return g;
}

// Stacked bar glass fill — vertical, slightly richer
function glassBarFill(ctx, color, h) {
  const height = h || 380;
  const g = ctx.createLinearGradient(0, 0, 0, height);
  g.addColorStop(0,   color + "55");
  g.addColorStop(0.5, color + "22");
  g.addColorStop(1,   color + "06");
  return g;
}

function axisStyle(tickColor, label) {
  return {
    ticks : { color: "#ffffff", font: { family: CHART_FONT, size: 13 }, maxRotation: 0 },
    grid  : GLASS_GRID,
    border: { display: false },
    title : label
      ? { display: true, text: label, color: "#ffffff", font: { family: CHART_FONT, size: 14, weight: "700" }, padding: { bottom: 8 } }
      : { display: false },
  };
}

/* =====================================================
   CHART PLUGINS
===================================================== */

// Soft per-dataset glow
const GLOW_PLUGIN = {
  id: "glassGlow",
  beforeDatasetDraw(chart, args) {
    const ds  = chart.data.datasets[args.index];
    if (!ds) return;
    const col = typeof ds.borderColor === "string" ? ds.borderColor : GC.blue;
    chart.ctx.save();
    chart.ctx.shadowColor  = col + "88";
    chart.ctx.shadowBlur   = chart.getDatasetMeta(args.index)?.type === "bar" ? 16 : 10;
  },
  afterDatasetDraw(chart) { chart.ctx.restore(); },
};

// Floating value labels on bars
const VALUE_LABEL_PLUGIN = {
  id: "glassValueLabels",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden || meta.type !== "bar") return;
      const col = typeof ds.borderColor === "string" ? ds.borderColor
        : (Array.isArray(ds.borderColor) ? ds.borderColor[0] : "#fff");
      meta.data.forEach((bar, idx) => {
        const raw = ds.data[idx];
        if (raw === null || raw === undefined || raw === 0) return;
        const lbl = ds._valueLabels
          ? (ds._valueLabels[idx] ?? "")
          : (typeof raw === "number" ? (Number.isInteger(raw) ? raw : raw.toFixed(1)) : raw);
        if (!lbl && lbl !== 0) return;
        ctx.save();
        ctx.font        = `500 10px ${CHART_MONO}`;
        ctx.fillStyle   = col;
        ctx.shadowColor = col + "99";
        ctx.shadowBlur  = 6;
        const isH = chart.options.indexAxis === "y";
        if (isH) {
          ctx.textAlign    = "left";
          ctx.textBaseline = "middle";
          ctx.fillText(lbl, bar.x + 8, bar.y);
        } else {
          ctx.textAlign    = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(lbl, bar.x, bar.y - 6);
        }
        ctx.restore();
      });
    });
  },
};

// Glowing ring + label on line chart peak
const PEAK_PLUGIN = {
  id: "glassPeak",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, i) => {
      if (!ds._showPeak) return;
      const meta = chart.getDatasetMeta(i);
      const vals = ds.data.map(d => (typeof d === "object" ? d?.y : d) || 0);
      const pk   = Math.max(...vals);
      if (pk <= 0) return;
      const col = typeof ds.borderColor === "string" ? ds.borderColor : "#fff";
      meta.data.forEach((pt, idx) => {
        if (vals[idx] !== pk) return;
        ctx.save();
        // outer glow ring
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, 9, 0, Math.PI * 2);
        ctx.strokeStyle = col + "33";
        ctx.lineWidth   = 6;
        ctx.stroke();
        // inner dot
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, 4, 0, Math.PI * 2);
        ctx.fillStyle   = "#050810";
        ctx.strokeStyle = col;
        ctx.lineWidth   = 1.5;
        ctx.shadowColor = col;
        ctx.shadowBlur  = 12;
        ctx.fill();
        ctx.stroke();
        // value label
        ctx.font        = `600 11px ${CHART_FONT}`;
        ctx.fillStyle   = "#ffffff";
        ctx.shadowColor = col;
        ctx.shadowBlur  = 10;
        ctx.textAlign   = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(String(pk), pt.x, pt.y - 13);
        ctx.restore();
      });
    });
  },
};

// Bubble count label plugin for flow chart breakage dots
const BUBBLE_LABEL_PLUGIN = {
  id: "bubbleLabel",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((ds, i) => {
      if (!ds._bubbleLabel) return;
      const meta = chart.getDatasetMeta(i);
      meta.data.forEach((pt, idx) => {
        const raw = ds.data[idx];
        if (!raw || raw.count < 4) return;
        ctx.save();
        ctx.font        = `600 10px ${CHART_MONO}`;
        ctx.fillStyle   = "#fff";
        ctx.shadowColor = GC.pink;
        ctx.shadowBlur  = 8;
        ctx.textAlign   = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(String(raw.count), pt.x, pt.y);
        ctx.restore();
      });
    });
  },
};

/* =====================================================
   STATUS BADGE
===================================================== */

function updateStatusBadge(state) {
  const el = document.getElementById("refreshStatus");
  if (!el) return;
  el.classList.remove("live", "updating", "history");
  el.classList.add(state);
  el.textContent = { live: "Live", updating: "Updating...", history: "History" }[state] || state;
}

function updateReportDateDisplay(data) {
  const el = document.getElementById("activeReportDate");
  if (!el) return;
  el.innerHTML = currentDate === null
    ? "Viewing Report Date: <strong style='color:#4ade80'>LIVE</strong>"
    : "Viewing Report Date: <strong>" + currentDate + "</strong>";
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = (value === null || value === undefined || value === "") ? "-" : value;
}

/* =====================================================
   TAB SWITCH
===================================================== */

function showTab(tabId, button) {
  document.querySelectorAll(".tab-content").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  const tabEl = document.getElementById(tabId);
  if (tabEl)  tabEl.classList.add("active");
  if (button) button.classList.add("active");

  if (tabId === "trends") {
    setTimeout(() => {
      buildTrendCharts();
      if (dashboardProcessed) buildFlowChart(dashboardProcessed);
    }, 100);
  }
  if (tabId === "reasons") {
    if (reasonChart) { reasonChart.destroy(); reasonChart = null; }
    setTimeout(() => buildReasonChart(dashboardData), 80);
  }
  if (tabId === "machines") {
    if (machineChart) { machineChart.destroy(); machineChart = null; }
    setTimeout(() => buildMachineChart(dashboardMachine || dashboardData), 80);
  }
  if (tabId === "compare") {
    const inp = document.getElementById("compareDateInputs");
    if (inp && inp.children.length === 0) compareInit();
  }
  if (tabId === "daily") {
    showSummaryLoader("daily");
    setTimeout(() => {
      buildDailySummary(dashboardData);
      hideSummaryLoader("daily");
    }, 80);
  }
  if (tabId === "weekly") {
    const we = document.getElementById("weekEndDate");
    if (we && !we.value) {
      const today = new Date();
      we.value = today.toISOString().slice(0,10);
    }
    showSummaryLoader("weekly");
    setTimeout(() => {
      buildWeeklySummary().finally(() => hideSummaryLoader("weekly"));
    }, 80);
  }
}

/* =====================================================
   LOAD DASHBOARD
===================================================== */

async function loadDashboard() {
  updateStatusBadge("updating");
  const startTime = performance.now();
  const dateParam = currentDate ? `&date=${encodeURIComponent(currentDate)}` : "";

  const [processedRes, machineRes] = await Promise.all([
    fetch(`${API_URL}?mode=processed${dateParam}`),
    fetch(`${API_URL}?mode=machine${dateParam}`)
  ]);

  dashboardProcessed = await processedRes.json();
  dashboardMachine   = await machineRes.json();
  dashboardData      = dashboardProcessed;

  const summary = dashboardData.summary || {};

  buildInsights(dashboardProcessed);
  updateReportDateDisplay(dashboardProcessed);
  checkSpikeAlert(dashboardData.hourly || []);

  // Push live stats to loading screen preview while it's still visible
  if (typeof window._updateLoadingStats === "function") {
    window._updateLoadingStats(summary);
  }

  // ── Animate KPI values ──
  const brokenVal = summary.totalBreakLenses || 0;
  const pctVal    = parseFloat(summary.breakPercent) || 0;
  const avgHrs    = parseFloat(summary.avgBreakTimeHours) || 0;

  animateValue(document.getElementById("totalJobs"),   summary.totalJobs ?? 0, "", "teal");
  animateValue(document.getElementById("totalLenses"), summary.totalLenses ?? 0, "", "");
  animateValue(document.getElementById("totalBroken"), brokenVal, "",
    brokenVal === 0 ? "green" : brokenVal <= 20 ? "yellow" : "red");

  const pctEl = document.getElementById("rxBreakage");
  if (pctEl) {
    animateValue(pctEl, pctVal, "%  (" + brokenVal + ")",
      pctVal < 2 ? "green" : pctVal < 4 ? "yellow" : "red");
  }

  const avgEl = document.getElementById("avgTime");
  if (avgEl) {
    animateValue(avgEl, avgHrs, avgHrs > 0 ? " hrs" : "",
      avgHrs === 0 ? "green" : avgHrs < 8 ? "yellow" : "red");
  }

  setText("peakHour", summary.peakHour ?? "--");

  if (summary.flowHealth) {
    animateValue(document.getElementById("flowHealthy"),   summary.flowHealth.healthy   || 0, "", "green");
    animateValue(document.getElementById("flowWatch"),     summary.flowHealth.watch     || 0, "", "yellow");
    animateValue(document.getElementById("flowDelayed"),   summary.flowHealth.delayed   || 0, "", "red");
    animateValue(document.getElementById("flowOvernight"), summary.flowHealth.overnight || 0, "", "blue");
  }
  if (summary.aging) {
    animateValue(document.getElementById("sameDayBreakage"),   summary.aging.sameDay || 0, "", "green");
    animateValue(document.getElementById("yesterdayBreakage"), summary.aging.oneDay  || 0, "", "yellow");
    animateValue(document.getElementById("twoPlusBreakage"),   summary.aging.twoPlus || 0, "", "red");
  }

  // Stagger KPI cards in
  animateKpiCards();

  // Update flow stat pills if they exist
  updateFlowStatPills(summary);

  buildHourlyTable(dashboardData.hourly || []);
  // Stagger table rows in after a brief pause
  setTimeout(animateTableRows, 80);

  if (trendChart)   trendChart.destroy();
  if (reasonChart)  reasonChart.destroy();
  if (machineChart) machineChart.destroy();
  if (flowChart)    flowChart.destroy();
  trendChart = reasonChart = machineChart = flowChart = null;

  const activeTab = document.querySelector(".tab.active");
  if (activeTab) {
    const tabId = activeTab.getAttribute("onclick") || "";
    if (tabId.includes("trends"))        buildTrendCharts();
    else if (tabId.includes("reasons"))  buildReasonChart(dashboardData);
    else if (tabId.includes("machines")) buildMachineChart(dashboardMachine || dashboardData);
  }
  if (document.getElementById("flowChart")) buildFlowChart(dashboardProcessed);

  updateStatusBadge(currentDate === null ? "live" : "history");
  // Refresh Daily Summary if it's visible, and invalidate weekly cache for today
  buildDailySummary(dashboardData);
  const todayKey = (() => { const n=new Date(); return `${n.getMonth()+1}/${n.getDate()}/${n.getFullYear()}`; })();
  if (currentDate === null && weeklyCache) weeklyCache[todayKey] = dashboardData;
  console.log("Load time (ms):", Math.round(performance.now() - startTime));
}

function updateFlowStatPills(summary) {
  const pills = {
    "flowStatHealthy"  : { val: summary.flowHealth?.healthy   || 0, color: "#00e676" },
    "flowStatWatch"    : { val: summary.flowHealth?.watch     || 0, color: "#ffd600" },
    "flowStatDelayed"  : { val: summary.flowHealth?.delayed   || 0, color: "#ff1744" },
    "flowStatBreakage" : { val: summary.totalBreakLenses      || 0, color: "#ff4081" },
  };
  Object.entries(pills).forEach(([id, { val, color }]) => {
    const el = document.getElementById(id);
    if (el) { el.textContent = val; el.style.color = color; }
  });
}

// Show history/live loading overlay
function showHistoryLoader(mode, dateLabel) {
  const overlay = document.getElementById("historyLoadOverlay");
  const title   = document.getElementById("hloTitle");
  const date    = document.getElementById("hloDate");
  const bar     = document.getElementById("hloBarFill");
  const status  = document.getElementById("hloStatus");
  if (!overlay) return;

  overlay.classList.remove("mode-live", "mode-history");
  overlay.classList.add(mode === "live" ? "mode-live" : "mode-history");

  if (title)  title.textContent  = mode === "live" ? "Returning to Live" : "Loading History";
  if (date)   date.textContent   = mode === "live" ? "Today" : (dateLabel || "—");
  if (bar)    bar.style.width    = "0%";
  if (status) status.textContent = "Fetching scan records...";

  overlay.classList.add("active");

  // Animate bar in steps
  const steps = [
    { pct: 30, msg: "Fetching scan records...",     delay: 0   },
    { pct: 60, msg: "Loading machine data...",       delay: 350 },
    { pct: 85, msg: "Building dashboard...",         delay: 700 },
  ];
  steps.forEach(s => {
    setTimeout(() => {
      if (bar)    bar.style.width    = s.pct + "%";
      if (status) status.textContent = s.msg;
    }, s.delay);
  });
}

function hideHistoryLoader() {
  const overlay = document.getElementById("historyLoadOverlay");
  const bar     = document.getElementById("hloBarFill");
  const status  = document.getElementById("hloStatus");
  if (!overlay) return;
  if (bar)    bar.style.width    = "100%";
  if (status) status.textContent = "Ready.";
  setTimeout(() => overlay.classList.remove("active"), 300);
}

/* =====================================================
   DATE FILTER
===================================================== */

function applyDateFilter() {
  const dateInput = document.getElementById("historyDate").value;
  if (!dateInput) return;
  const parts = dateInput.split("-");
  currentDate = parseInt(parts[1], 10) + "/" + parseInt(parts[2], 10) + "/" + parts[0];
  showHistoryLoader("history", currentDate);
  loadDashboard().finally(() => hideHistoryLoader());
  startRefreshCountdown();
}

function resetToToday() {
  currentDate = null;
  const dateEl = document.getElementById("historyDate");
  if (dateEl) dateEl.value = "";
  showHistoryLoader("live", "Today");
  loadDashboard().finally(() => hideHistoryLoader());
  startRefreshCountdown();
}

document.addEventListener("DOMContentLoaded", () => {
  const liveBtn = document.getElementById("liveBtn");
  if (liveBtn) {
    liveBtn.addEventListener("click", () => {
      currentDate = null;
      const d = document.getElementById("historyDate");
      if (d) d.value = "";
      showHistoryLoader("live", "Today");
      loadDashboard().finally(() => hideHistoryLoader());
      startRefreshCountdown();
    });
  }
  startRefreshCountdown();
  compareInit();

  /* ── CUSTOM CURSOR ── */
  const cursor = document.getElementById("customCursor");
  const trail  = document.getElementById("cursorTrail");
  let trailX = 0, trailY = 0, cursorX = 0, cursorY = 0;

  document.addEventListener("mousemove", e => {
    cursorX = e.clientX; cursorY = e.clientY;
    if (cursor) { cursor.style.left = cursorX + "px"; cursor.style.top = cursorY + "px"; }
  });

  function animateTrail() {
    trailX += (cursorX - trailX) * 0.14;
    trailY += (cursorY - trailY) * 0.14;
    if (trail) { trail.style.left = trailX + "px"; trail.style.top = trailY + "px"; }
    requestAnimationFrame(animateTrail);
  }
  animateTrail();

  document.addEventListener("mousedown", () => {
    if (cursor) { cursor.style.transform = "translate(-50%,-50%) scale(0.7)"; }
    if (trail)  { trail.style.transform  = "translate(-50%,-50%) scale(0.7)"; }
  });
  document.addEventListener("mouseup", () => {
    if (cursor) { cursor.style.transform = "translate(-50%,-50%) scale(1)"; }
    if (trail)  { trail.style.transform  = "translate(-50%,-50%) scale(1)"; }
  });

  /* ── LOADING SCREEN ── */
  const loadScreen = document.getElementById("loadingScreen");
  const barFill    = document.getElementById("loadingBarFill");
  const statusEl   = document.getElementById("loadingStatus");

  const lsSteps = [
    { pct: 15,  msg: "Connecting to data source...",  step: 0 },
    { pct: 40,  msg: "Fetching processed scan data...",step: 1 },
    { pct: 65,  msg: "Loading machine metrics...",     step: 2 },
    { pct: 85,  msg: "Building dashboard...",          step: 3 },
  ];

  function setLsStep(idx) {
    for (let i = 0; i < 4; i++) {
      const el = document.getElementById("lsStep" + i);
      if (!el) continue;
      el.className = "ls-step" + (i < idx ? " done" : i === idx ? " active" : "");
    }
  }

  let stepIdx = 0;
  const stepInterval = setInterval(() => {
    if (stepIdx < lsSteps.length) {
      const s = lsSteps[stepIdx];
      if (barFill) barFill.style.width = s.pct + "%";
      if (statusEl) statusEl.textContent = s.msg;
      setLsStep(s.step);
      stepIdx++;
    } else {
      clearInterval(stepInterval);
    }
  }, 420);

  // stepInterval and setLsStep defined inside DOMContentLoaded, but we need refs outside
  // The actual dismiss/stats functions are now defined at module level below
});

// ── Module-level loading screen functions ──
// Must be defined HERE (not inside DOMContentLoaded) so loadDashboard() can call them
// even if the API responds before DOMContentLoaded fires.

window._updateLoadingStats = function(summary) {
  const statsStrip = document.querySelector(".ls-stats");
  if (statsStrip) {
    statsStrip.style.animation = "none";
    statsStrip.style.opacity   = "1";
  }
  const brokenVal = summary.totalBreakLenses || 0;
  const pctVal    = parseFloat(summary.breakPercent) || 0;

  const lsJobs = document.getElementById("lsJobs");
  const lsLen  = document.getElementById("lsLenses");
  const lsBrk  = document.getElementById("lsBroken");
  const lsPct  = document.getElementById("lsPct");
  const lsPk   = document.getElementById("lsPeak");

  if (lsJobs) { lsJobs.textContent = (summary.totalJobs ?? 0).toLocaleString(); lsJobs.classList.add("ready"); }
  if (lsLen)  { lsLen.textContent  = (summary.totalLenses ?? 0).toLocaleString(); lsLen.classList.add("ready"); }
  if (lsBrk)  {
    lsBrk.textContent = brokenVal;
    lsBrk.classList.add("ready");
    lsBrk.classList.remove("green","yellow","red");
    lsBrk.classList.add(brokenVal === 0 ? "green" : brokenVal <= 20 ? "yellow" : "red");
  }
  if (lsPct)  {
    lsPct.textContent = pctVal.toFixed(2) + "%";
    lsPct.classList.add("ready");
    lsPct.classList.remove("green","yellow","red");
    lsPct.classList.add(pctVal < 2 ? "green" : pctVal < 4 ? "yellow" : "red");
  }
  if (lsPk) { lsPk.textContent = summary.peakHour ?? "—"; lsPk.classList.add("ready"); }
};

window._dismissLoadingScreen = function() {
  // Mark all steps done
  for (let i = 0; i < 4; i++) {
    const el = document.getElementById("lsStep" + i);
    if (el) el.className = "ls-step done";
  }
  const barFill  = document.getElementById("loadingBarFill");
  const statusEl = document.getElementById("loadingStatus");
  if (barFill)  barFill.style.width = "100%";
  if (statusEl) statusEl.textContent = "Dashboard ready ✓";
  const loadScreen = document.getElementById("loadingScreen");
  setTimeout(() => {
    if (loadScreen) {
      loadScreen.classList.add("fade-out");
      setTimeout(() => { loadScreen.style.display = "none"; }, 720);
    }
  }, 900);
};

loadDashboard().then(() => {
  if (typeof window._dismissLoadingScreen === "function") window._dismissLoadingScreen();
}).catch(() => {
  if (typeof window._dismissLoadingScreen === "function") window._dismissLoadingScreen();
});
loadPreviousDay();

/* =====================================================
   TREND CHART SYSTEM
===================================================== */

function buildTrendCharts() {
  trendsLoaded = true;
  buildTrendChart(dashboardProcessed);
  setTimeout(() => buildFlowChart(dashboardProcessed), 100);
}

function switchTrend(mode) {
  currentMode = mode;
  if (mode === "processed") buildTrendChart(dashboardProcessed);
  else if (mode === "machine") buildTrendChart(dashboardMachine);
  document.querySelectorAll(".mode-btn").forEach(btn => btn.classList.remove("active"));
  const btn = document.getElementById(mode + "Btn");
  if (btn) btn.classList.add("active");
}

/* =====================================================
   TREND CHART — GLASSMORPHISM SMOOTH CURVES
===================================================== */

function buildTrendChart(data) {
  const canvas = document.getElementById("trendChart");
  if (!canvas || !data || !data.hourly) return;
  const ctx  = canvas.getContext("2d");
  if (trendChart) trendChart.destroy();

  const sorted = [...data.hourly].sort((a, b) =>
    new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour)
  );
  const hours   = sorted.map(h => h.hour);
  const chartH  = canvas.clientHeight || 420;
  const datasets = [];
  const allReasons = new Set();

  sorted.forEach(hour =>
    Object.values(hour.machines || {}).forEach(m =>
      Object.keys(m.reasons || {}).forEach(r => allReasons.add(r))
    )
  );

  // Sort reasons by total so top reason gets peak annotation
  const fallback = [GC.red, GC.purple, GC.orange, GC.teal, GC.blue, GC.green, GC.pink];
  let ci = 0;

  const reasonTotalsMap = {};
  allReasons.forEach(reason => {
    reasonTotalsMap[reason] = sorted.map(hour => {
      let t = 0;
      Object.values(hour.machines || {}).forEach(m => { t += m.reasons?.[reason]?.total || 0; });
      return t;
    });
  });

  const sortedReasons = [...allReasons].sort((a, b) =>
    reasonTotalsMap[b].reduce((s, v) => s+v, 0) - reasonTotalsMap[a].reduce((s, v) => s+v, 0)
  );

  sortedReasons.forEach((reason, rank) => {
    const color  = BREAKAGE_COLOR_MAP[reason] || fallback[ci % fallback.length];
    const totals = reasonTotalsMap[reason];
    const peak   = Math.max(...totals, 0);
    const isTop  = rank === 0;

    datasets.push({
      label              : reason,
      type               : "bar",
      data               : totals,
      backgroundColor    : color + "cc",
      borderColor        : color,
      borderWidth        : 1,
      borderRadius       : { topLeft: 4, topRight: 4 },
      borderSkipped      : false,
      barPercentage      : 0.75,
      categoryPercentage : 0.85,
      yAxisID            : "yBroken",
      order              : 2,
    });
    ci++;
  });

  // Coating jobs — amber dashed reference line
  datasets.push({
    label            : "Coating Jobs",
    type             : "line",
    data             : sorted.map(h => h.coatingJobs || 0),
    borderColor      : GC.yellow + "88",
    backgroundColor  : "transparent",
    borderWidth      : 1.5,
    borderDash       : [5, 4],
    tension          : 0.3,
    fill             : false,
    pointRadius      : 0,
    pointHoverRadius : 4,
    yAxisID          : "yJobs",
    order            : 1,
  });

  trendChart = new Chart(ctx, {
    type: "line",
    data: { labels: hours, datasets },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      interaction        : { mode: "index", intersect: false },
      animation          : { duration: 500, easing: "easeOutQuart" },
      plugins: {
        legend : GLASS_LEGEND,
        tooltip: {
          ...GLASS_TOOLTIP,
          callbacks: {
            title    : items => `  ${items[0].label}`,
            label    : ctx  => {
              if (ctx.dataset.label === "Coating Jobs") return `  Jobs: ${ctx.raw}`;
              return ctx.raw > 0 ? `  ${ctx.dataset.label}: ${ctx.raw}` : null;
            },
            afterBody: items => {
              const total = items
                .filter(i => i.dataset.label !== "Coating Jobs")
                .reduce((s, i) => s + (i.raw || 0), 0);
              return total > 0 ? ["", `  Total broken: ${total}`] : [];
            },
          },
        },
      },
      scales: {
        x: { ...axisStyle("rgba(140,175,220,0.4)"), grid: { ...GLASS_GRID } },
        yBroken: {
          type: "linear", position: "left", beginAtZero: true,
          ...axisStyle("rgba(248,113,113,0.6)", "Breakage"),
          ticks: { color: "#ff8888", font: { family: CHART_FONT, size: 13 }, stepSize: 1 },
          grid: { ...GLASS_GRID, color: "rgba(248,113,113,0.05)" },
        },
        yJobs: {
          type: "linear", position: "right", beginAtZero: true,
          ...axisStyle("rgba(251,191,36,0.35)", "Jobs"),
          ticks: { color: "#fbbf24", font: { family: CHART_FONT, size: 13 } },
          grid: { drawOnChartArea: false },
        },
      },
    },
    plugins: [GLOW_PLUGIN, PEAK_PLUGIN],
  });
}

/* =====================================================
   FLOW CHART — STACKED BARS + BUBBLE BREAKAGE EVENTS
===================================================== */

function switchFlowMode(mode) {
  currentFlowMode = mode;
  document.querySelectorAll("#avgFlowBtn, #machineFlowBtn, #individualFlowBtn")
    .forEach(btn => btn.classList.remove("active"));
  const btnMap = { average: "avgFlowBtn", machine: "machineFlowBtn", individual: "individualFlowBtn" };
  const b = document.getElementById(btnMap[mode]);
  if (b) b.classList.add("active");
  buildFlowChart(dashboardProcessed);
}

function buildFlowChart(data) {
  const canvas = document.getElementById("flowChart");
  if (!canvas || !data || !data.hourly) return;
  const ctx = canvas.getContext("2d");
  if (flowChart) flowChart.destroy();

  const sorted = [...data.hourly].sort((a, b) =>
    new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour)
  );
  const filtered = sorted.filter(h => {
    if (!h.hour) return false;
    let n = parseInt(h.hour.split(":")[0], 10);
    if (h.hour.includes("PM") && n !== 12) n += 12;
    if (h.hour.includes("AM") && n === 12) n = 0;
    return n >= 6 && n <= 20;
  });
  const hours  = filtered.map(h => h.hour);
  const chartH = canvas.clientHeight || 420;

  // Cap at 360m (6h) — anything over is "Overnight" and excluded from avg calculations
  // This prevents single outlier jobs (e.g. 857m) from blowing the broken avg line off the chart
  function sanitize(v) {
    if (v === null || v === undefined) return null;
    const n = Number(v);
    return n > 360 ? null : n;
  }

  // ── MACHINE MODE — smooth curves per machine ──
  if (currentFlowMode === "machine") {
    const machineSet = new Set();
    filtered.forEach(h => Object.keys(h.machines || {}).forEach(m => machineSet.add(m)));
    const palette = [GC.cyan, GC.red, GC.orange, GC.yellow, GC.green, GC.purple, GC.blue, GC.pink];

    const datasets = Array.from(machineSet).map((machine, idx) => {
      const color = palette[idx % palette.length];
      return {
        label              : formatMachineLabel(machine),
        data               : filtered.map(h => ({
          x    : h.hour,
          y    : sanitize(h.machines?.[machine]?.avgFlowAll),
          count: h.machines?.[machine]?.flowAllCount || 0,
        })),
        borderColor        : color,
        backgroundColor    : glassFill(ctx, color, chartH),
        borderWidth        : 2,
        tension            : 0.42,
        fill               : false,
        pointRadius        : 3,
        pointHoverRadius   : 7,
        pointBackgroundColor: color,
        spanGaps           : true,
      };
    });

    flowChart = new Chart(ctx, {
      type   : "line",
      data   : { labels: hours, datasets },
      options: getMachineFlowOptions(),
      plugins: [GLOW_PLUGIN],
    });
    return;
  }

  // ── INDIVIDUAL SCATTER ──
  if (currentFlowMode === "individual") {
    const points = [];
    filtered.forEach(hourObj =>
      Object.entries(hourObj.machines || {}).forEach(([machine, mData]) => {
        const fv = sanitize(mData.avgFlowAll);
        if (fv === null) return;
        points.push({ x: hourObj.hour, y: fv, machine, rx: null });
      })
    );
    flowChart = new Chart(ctx, {
      type: "scatter",
      data: {
        datasets: [{
          label              : "Machine Flow",
          data               : points,
          pointRadius        : 9,
          pointHoverRadius   : 12,
          backgroundColor    : points.map(p => getFlowColor(p.y) + "bb"),
          borderColor        : points.map(p => getFlowColor(p.y)),
          borderWidth        : 1.5,
        }],
      },
      options: getMachineFlowOptions(),
      plugins: [GLOW_PLUGIN],
    });
    return;
  }

  // ── AVERAGE MODE — multi-line: one curve per bucket + threshold bands + breakage dots ──

  const bucketDefs = [
    { label: "Healthy ≤15m",    color: "#4ade80", filter: p => p.flow > 0  && p.flow <= 15  },
    { label: "Watch 16–30m",    color: "#fbbf24", filter: p => p.flow > 15 && p.flow <= 30  },
    { label: "Delayed 31–360m", color: "#f87171", filter: p => p.flow > 30 && p.flow <= 360 },
    { label: "Overnight >6h",   color: "#38bdf8", filter: p => p.flow > 360                 },
  ];

  // Smart Y cap — ignore overnight outliers for axis calculation
  const allFlowVals = [];
  filtered.forEach(h => (h.flowPoints || []).forEach(p => {
    if (p.flow > 0 && p.flow <= 360) allFlowVals.push(p.flow);
  }));
  const p95 = allFlowVals.length
    ? allFlowVals.sort((a,b)=>a-b)[Math.floor(allFlowVals.length * 0.95)]
    : 60;
  const yAxisMax = Math.max(60, Math.ceil(p95 * 1.25 / 10) * 10);

  // Avg flow time per bucket per hour — also store job counts
  const bucketDatasets = bucketDefs.map(b => {
    const isOvernight = b.label.includes("Overnight");
    const vals = filtered.map(h => {
      const pts = (h.flowPoints || []).filter(b.filter);
      if (!pts.length) return null;
      const avg = Math.round(pts.reduce((s, p) => s + p.flow, 0) / pts.length);
      return isOvernight ? Math.min(avg, yAxisMax) : avg;
    });
    // Store count per hour index for tooltip use
    const counts = filtered.map(h => (h.flowPoints || []).filter(b.filter).length);
    const fill = ctx.createLinearGradient(0, 0, 0, chartH);
    fill.addColorStop(0,   b.color + "22");
    fill.addColorStop(1,   b.color + "00");
    return {
      label              : b.label,
      data               : vals,
      _counts            : counts,
      borderColor        : b.color,
      backgroundColor    : fill,
      borderWidth        : 2.5,
      tension            : 0.42,
      fill               : true,
      pointRadius        : vals.map(v => v === null ? 0 : 4),
      pointHoverRadius   : 10,
      pointHitRadius     : 16,
      pointBackgroundColor: b.color,
      pointBorderColor   : "rgba(5,8,16,0.7)",
      pointBorderWidth   : 1.5,
      spanGaps           : false,
      _bucketColor       : b.color,
      _isOvernight       : isOvernight,
    };
  });

  // Broken jobs avg flow — dashed violet line
  // Recalculate from flowPoints (broken flag, ≤360m) to exclude overnight outliers like 857m
  // Falls back to h.avgFlowBroken if flowPoints don't have a .broken flag
  const brokenAvgVals = filtered.map(h => {
    const brokenPts = (h.flowPoints || []).filter(p => p.broken && p.flow > 0 && p.flow <= 360);
    if (brokenPts.length) {
      return Math.round(brokenPts.reduce((s, p) => s + p.flow, 0) / brokenPts.length);
    }
    return sanitize(h.avgFlowBroken); // fallback to API value (already capped at 360 by sanitize)
  });
  const brokenCounts = filtered.map(h => h.totalBroken || 0);
  bucketDatasets.push({
    label              : "Broken Jobs Avg",
    data               : brokenAvgVals,
    _counts            : brokenCounts,
    borderColor        : "rgba(192,166,255,0.75)",
    backgroundColor    : "transparent",
    borderWidth        : 1.5,
    borderDash         : [5, 4],
    tension            : 0.42,
    fill               : false,
    pointRadius        : brokenAvgVals.map(v => v === null ? 0 : 2.5),
    pointHoverRadius   : 6,
    pointBackgroundColor: "#c084fc",
    pointBorderColor   : "rgba(5,8,16,0.7)",
    pointBorderWidth   : 1,
    spanGaps           : true,
  });

  // Breakage events — placed on chart via custom plugin
  // Y position = avg flow of BROKEN jobs excluding overnight outliers (≤360m)
  const breakageDots = filtered.map((h, i) => {
    const count = h.totalBroken || 0;
    if (!count) return null;
    // Recalculate avg from flowPoints for broken jobs only, capping overnight
    const brokenPts = (h.flowPoints || []).filter(p => p.broken && p.flow > 0 && p.flow <= 360);
    const dotY = brokenPts.length
      ? Math.round(brokenPts.reduce((s, p) => s + p.flow, 0) / brokenPts.length)
      : sanitize(h.avgFlowBroken) || 10;
    return { x: hours[i], y: dotY || 10, count, hour: h.hour };
  }).filter(Boolean);

  // ── THRESHOLD BAND PLUGIN ──
  const BAND_PLUGIN = {
    id: "thresholdBands",
    beforeDraw(chart) {
      const { ctx: c, chartArea: a, scales: { y } } = chart;
      if (!a || !y) return;
      const toY = v => y.getPixelForValue(v);
      const maxY = y.max;
      c.save();

      c.fillStyle = "rgba(74,222,128,0.05)";
      c.fillRect(a.left, toY(15), a.width, toY(0) - toY(15));

      c.fillStyle = "rgba(251,191,36,0.05)";
      c.fillRect(a.left, toY(30), a.width, toY(15) - toY(30));

      c.fillStyle = "rgba(248,113,113,0.05)";
      c.fillRect(a.left, toY(maxY), a.width, toY(30) - toY(maxY));

      [[15, "rgba(74,222,128,0.4)", "≤15m"], [30, "rgba(251,191,36,0.4)", "≤30m"]].forEach(([val, col, lbl]) => {
        const yPx = toY(val);
        c.strokeStyle = col;
        c.lineWidth   = 1;
        c.setLineDash([4, 4]);
        c.beginPath();
        c.moveTo(a.left, yPx);
        c.lineTo(a.right, yPx);
        c.stroke();
        c.setLineDash([]);
        c.fillStyle    = col;
        c.font         = `500 9px ${CHART_MONO}`;
        c.textAlign    = "right";
        c.textBaseline = "bottom";
        c.fillText(lbl, a.right - 4, yPx - 3);
      });

      c.restore();
    }
  };

  // ── BREAKAGE DOT PLUGIN ──
  const BREAKAGE_DOT_PLUGIN = {
    id: "breakageDots",
    afterDatasetsDraw(chart) {
      const { ctx: c, scales: { x, y } } = chart;
      if (!x || !y) return;
      breakageDots.forEach(pt => {
        const xPx = x.getPixelForValue(pt.x);
        const yPx = y.getPixelForValue(pt.y);
        if (xPx == null || yPx == null) return;
        c.save();
        // outer pulse ring
        c.beginPath();
        c.arc(xPx, yPx, 13, 0, Math.PI * 2);
        c.strokeStyle = "rgba(255,64,129,0.2)";
        c.lineWidth   = 6;
        c.stroke();
        // inner dot
        c.beginPath();
        c.arc(xPx, yPx, 7, 0, Math.PI * 2);
        c.fillStyle   = "rgba(255,64,129,0.28)";
        c.strokeStyle = "#ff4081";
        c.lineWidth   = 2;
        c.shadowColor = "#ff4081";
        c.shadowBlur  = 14;
        c.fill();
        c.stroke();
        c.shadowBlur  = 0;
        // count label
        c.font         = `600 10px ${CHART_MONO}`;
        c.fillStyle    = "#ffffff";
        c.textAlign    = "center";
        c.textBaseline = "middle";
        c.fillText(String(pt.count), xPx, yPx);
        c.restore();
      });
    }
  };

  flowChart = new Chart(ctx, {
    type: "line",
    data: { labels: hours, datasets: bucketDatasets },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      animation          : { duration: 600, easing: "easeOutQuart" },
      interaction        : { mode: "nearest", intersect: true },

      plugins: {
        legend: {
          display  : true,
          position : "top",
          labels: {
            color        : "#ffffff",
            font         : { family: CHART_FONT, size: 14, weight: "600" },
            usePointStyle: true,
            pointStyle   : "circle",
            padding      : 24,
            boxWidth     : 12,
            boxHeight    : 12,
            // Strike-through text on hidden datasets
            generateLabels(chart) {
              return chart.data.datasets.map((ds, i) => {
                const meta   = chart.getDatasetMeta(i);
                const hidden = meta.hidden;
                return {
                  text          : ds.label,
                  fillStyle     : ds.borderColor,
                  strokeStyle   : ds.borderColor,
                  pointStyle    : "circle",
                  hidden,
                  lineDash      : ds.borderDash || [],
                  datasetIndex  : i,
                  fontColor     : hidden ? "rgba(255,255,255,0.18)" : "rgba(255,255,255,0.55)",
                  lineWidth     : 1,
                };
              });
            },
          },
          onClick(e, legendItem, legend) {
            const index = legendItem.datasetIndex;
            const ci    = legend.chart;
            const meta  = ci.getDatasetMeta(index);
            meta.hidden = !meta.hidden;
            ci.update();
          },
        },
        tooltip: {
          ...GLASS_TOOLTIP,
          titleFont : { family: CHART_FONT, size: 15, weight: "700" },
          bodyFont  : { family: CHART_MONO, size: 13 },
          callbacks: {
            title: items => {
              if (!items.length) return "";
              const hourIdx = items[0].dataIndex;
              const h       = filtered[hourIdx];
              const jobs    = h?.coatingJobs || 0;
              // Show the line name + hour + total jobs that hour
              const lineName = items[0].dataset.label || "";
              return [`  ${items[0].label}`, `  ${lineName}   ·   ${jobs} jobs ran`];
            },
            label: ctx => {
              const v       = ctx.raw;
              const hourIdx = ctx.dataIndex;
              const count   = ctx.dataset._counts?.[hourIdx] ?? 0;
              if (v === null || v === undefined) return null;

              if (ctx.dataset.label === "Broken Jobs Avg") {
                if (!count) return `  Avg flow of broken jobs: ${v}m`;
                return `  Avg flow: ${v}m  ·  ${count} lenses broken`;
              }

              // Bucket line — show this line's job count + avg flow
              const h   = filtered[hourIdx];
              const pct = count > 0 && (h?.coatingJobs || 0) > 0
                ? ` (${((count / (h.coatingJobs)) * 100).toFixed(0)}% of jobs)`
                : "";
              return count > 0
                ? `  ${count} jobs this hour · avg ${v}m${pct}`
                : `  No jobs in this range this hour`;
            },
            afterBody: items => {
              if (!items.length) return [];
              const hour  = items[0].label;
              const brkPt = breakageDots.find(d => d.hour === hour);
              // Only show breakage line if hovering Broken Jobs Avg or if there's breakage
              if (!brkPt) return [];
              return ["", `  🔴 ${brkPt.count} lenses broken this hour`];
            },
          },
        },
      },

      scales: {
        x: {
          ...axisStyle("rgba(140,175,220,0.45)"),
          grid: { ...GLASS_GRID },
        },
        y: {
          beginAtZero : true,
          max         : yAxisMax,
          ...axisStyle("rgba(140,175,220,0.4)", "Flow Time"),
          ticks: {
            color: "rgba(255,255,255,0.75)",
            font    : { family: CHART_MONO, size: 10 },
            callback: v => v >= 60 ? (v/60).toFixed(1) + "h" : v + "m",
          },
          grid: { ...GLASS_GRID },
        },
      },
    },
    plugins: [BAND_PLUGIN, GLOW_PLUGIN, BREAKAGE_DOT_PLUGIN],
  });
}

function getMachineFlowOptions() {
  return {
    responsive         : true,
    maintainAspectRatio: false,
    animation          : { duration: 500, easing: "easeOutQuart" },
    interaction        : { mode: "nearest", intersect: true },
    onClick(evt, elements, chart) {
      if (!elements.length) return;
      const pt = chart.data.datasets[elements[0].datasetIndex].data[elements[0].index];
      if (!pt) return;
      showFlowDetails({ rx: pt.rx || "Machine Flow", machine: pt.machine || "Unknown", reason: pt.reason || "N/A", flow: pt.y, x: pt.x });
    },
    plugins: {
      legend : GLASS_LEGEND,
      tooltip: {
        ...GLASS_TOOLTIP,
        callbacks: {
          title: items => `  ${items[0].label || items[0].raw?.x || ""}`,
          label: ctx => {
            const raw = ctx.raw, lines = [];
            if (raw?.machine) lines.push(`  Machine: ${raw.machine}`);
            const y = raw?.y ?? raw;
            if (typeof y === "number") {
              const h = Math.floor(y / 60), m = Math.round(y % 60);
              lines.push(`  Flow: ${h > 0 ? h + "h " : ""}${m}m`);
            }
            if (raw?.count > 0) lines.push(`  Jobs: ${raw.count}`);
            return lines;
          },
        },
      },
    },
    scales: {
      x: { type: "category", ...axisStyle("rgba(140,175,220,0.4)") },
      y: {
        beginAtZero: true,
        ...axisStyle("rgba(56,189,248,0.5)", "Minutes"),
        ticks: {
          color   : "rgba(56,189,248,0.5)",
          font    : { family: CHART_MONO, size: 10 },
          callback: v => v >= 60 ? (v / 60).toFixed(1) + "h" : v + "m",
        },
        grid: { ...GLASS_GRID, color: "rgba(56,189,248,0.05)" },
      },
    },
  };
}

/* =====================================================
   FLOW MODAL DETAILS
===================================================== */

function showFlowDetails(point) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;
  const tm = Number(point.flow) || 0;
  const h  = Math.floor(tm / 60), m = Math.round(tm % 60);
  const fmt = h > 0 ? (m > 0 ? `${h}h ${m}m` : `${h}h`) : `${m}m`;
  modalBody.innerHTML = `
    <h2>${point.rx !== "Machine Flow" ? "RX " + point.rx : "Machine Flow"}</h2>
    <div style="margin-top:14px;padding:18px;background:rgba(255,255,255,0.03);border:1px solid rgba(120,160,255,0.12);border-radius:10px;font-family:${CHART_FONT};font-size:13px;">
      <div style="margin-bottom:10px"><span style="color:rgba(100,140,180,0.6);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;">Machine</span><br><strong style="color:#c8dff5;">${point.machine}</strong></div>
      <div style="margin-bottom:10px"><span style="color:rgba(100,140,180,0.6);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;">Breakage Reason</span><br><strong style="color:#c8dff5;">${point.reason || "None"}</strong></div>
      <div style="margin-bottom:10px"><span style="color:rgba(100,140,180,0.6);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;">Flow Time</span><br><strong style="color:${GC.cyan};font-size:22px;">${fmt}</strong></div>
      <div><span style="color:rgba(100,140,180,0.6);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;">Hour</span><br><strong style="color:#c8dff5;">${point.x}</strong></div>
    </div>`;
  modal.classList.add("active");
}

function showFlowHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;
  let html = "";
  (hourData.flowPoints || []).forEach(p => {
    const col = getFlowColor(p.flow);
    html += `<div style="margin-bottom:10px;padding:12px;background:rgba(255,255,255,0.03);border-left:3px solid ${col};border-radius:6px;font-size:13px;line-height:1.8;font-family:${CHART_FONT};"><strong style="color:${col}">RX ${p.rx}</strong><br>${p.machine} · ${p.reason} · <strong style="color:${col}">${p.flow}m</strong></div>`;
  });
  modalBody.innerHTML = `<h2>${hourData.hour}</h2>${html || "<p style='color:rgba(100,140,180,0.5)'>No flow data</p>"}`;
  modal.classList.add("active");
}

/* =====================================================
   HOURLY TABLE
===================================================== */

function buildHourlyTable(hourly) {
  const tbody = document.getElementById("hourlyBody");
  if (!tbody || !Array.isArray(hourly)) return;
  tbody.innerHTML = "";

  hourly.forEach(row => {
    const tr     = document.createElement("tr");
    const broken = row.totalBroken || 0;

    let brkClass = "brk-zero";
    if (broken >= 10)     brkClass = "brk-high";
    else if (broken >= 4) brkClass = "brk-mid";
    else if (broken > 0)  brkClass = "brk-low";

    let primaryDriverHTML  = "<span style='color:rgba(255,255,255,0.3)'>—</span>";
    let topAccessPointHTML = "<span style='color:rgba(255,255,255,0.3)'>—</span>";

    if (broken > 0 && row.machines) {
      let topMachine = null, topMachineTotal = 0;
      Object.entries(row.machines).forEach(([machine, stats]) => {
        if ((stats.total || 0) > topMachineTotal) { topMachine = machine; topMachineTotal = stats.total || 0; }
      });
      if (topMachine && row.machines[topMachine]) {
        let topReason = null, topReasonTotal = 0;
        Object.entries(row.machines[topMachine].reasons || {}).forEach(([reason, rStats]) => {
          if ((rStats.total || 0) > topReasonTotal) { topReason = reason; topReasonTotal = rStats.total || 0; }
        });
        primaryDriverHTML = `<div class="primary-driver"><div class="driver-machine">${formatMachineLabel(topMachine)}</div><div class="driver-reason">${topReason || ""}</div></div>`;
      }

      const brkMachines = Object.entries(row.machines)
        .filter(([, s]) => (s.total || 0) > 0)
        .sort(([, a], [, b]) => (b.total || 0) - (a.total || 0))
        .map(([m, s]) => `<span style="color:#38bdf8;font-family:'JetBrains Mono',monospace;font-size:11px;">${formatMachineLabel(m)}</span><span style="color:rgba(255,255,255,0.4);font-size:11px;"> ×${s.total}</span>`)
        .join("<br>");
      if (brkMachines) topAccessPointHTML = brkMachines;
    }

    // ── Flow Breakdown ──
    const flowPts   = row.flowPoints || [];
    const totalFlow = flowPts.length;
    let flowHTML    = `<span style="color:rgba(255,255,255,0.3)">—</span>`;

    if (totalFlow > 0) {
      const buckets = [
        { label: "H",  color: "#4ade80", pts: flowPts.filter(p => p.flow > 0  && p.flow <= 15)  },
        { label: "W",  color: "#fbbf24", pts: flowPts.filter(p => p.flow > 15 && p.flow <= 30)  },
        { label: "D",  color: "#f87171", pts: flowPts.filter(p => p.flow > 30 && p.flow <= 360) },
        { label: "ON", color: "#38bdf8", pts: flowPts.filter(p => p.flow > 360)                 },
      ];

      const pills = buckets
        .filter(b => b.pts.length > 0)
        .map(b => {
          const avg = Math.round(b.pts.reduce((s, p) => s + p.flow, 0) / b.pts.length);
          return `<span class="flow-pill" style="--pill-color:${b.color}">
            <span class="fp-label">${b.label}</span>
            <span class="fp-count">${b.pts.length}</span>
            <span class="fp-avg">${avg}m avg</span>
          </span>`;
        })
        .join("");

      flowHTML = `<div class="flow-breakdown-cell">
        <div class="flow-total-jobs">${totalFlow} flow records</div>
        <div class="flow-pills">${pills}</div>
      </div>`;
    }

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td class="${brkClass}">${broken > 0 ? broken : '<span style="color:rgba(255,255,255,0.3)">0</span>'}</td>
      <td>${row.coatingJobs || 0}</td>
      <td>${flowHTML}</td>
      <td>${topAccessPointHTML}</td>
      <td>${primaryDriverHTML}</td>`;
    tr.addEventListener("click", () => showHourDetails(row));
    tbody.appendChild(tr);
  });
}

/* =====================================================
   HOUR DETAIL MODAL
===================================================== */

function showHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody || !hourData) return;
  let machineHTML = "";

  Object.entries(hourData.machines || {}).forEach(([machine, stats]) => {
    const dm = formatMachineLabel(machine);
    let reasonHTML = "";
    Object.entries(stats.reasons || {}).forEach(([reason, rStats]) => {
      const c = BREAKAGE_COLOR_MAP[reason] || "#60a5fa";
      reasonHTML += `<div style="margin-top:10px;padding:10px 14px;background:rgba(255,255,255,0.04);border-left:3px solid ${c};border-radius:4px;">
        <strong style="color:${c};font-size:14px;">${reason}</strong>
        <span style="color:#ffffff;font-size:14px;margin-left:8px;">— ${rStats.total || 0}</span>
        <div style="color:rgba(255,255,255,0.55);font-size:12px;margin-top:4px;">Same: ${rStats.sameDay||0} &nbsp;·&nbsp; Prev: ${rStats.oneDay||0} &nbsp;·&nbsp; 2+: ${rStats.twoPlus||0}</div>
      </div>`;
    });
    machineHTML += `<div style="margin-bottom:14px;padding:16px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);border-radius:10px;line-height:1.8;">
      <div style="color:#ffffff;font-size:16px;font-weight:700;margin-bottom:8px;">${dm}</div>
      <div style="font-size:13px;margin-bottom:4px;">
        <span style="color:rgba(255,255,255,0.6);">JOBS:</span> <strong style="color:#ffffff;font-size:14px;">${stats.jobs||0}</strong>
        &nbsp;&nbsp;
        <span style="color:rgba(255,255,255,0.6);">BROKEN:</span> <strong style="color:${GC.red};font-size:14px;">${stats.total||0}</strong>
      </div>
      <div style="color:rgba(255,255,255,0.55);font-size:12px;">Same: ${stats.sameDay||0} &nbsp;·&nbsp; Prev: ${stats.oneDay||0} &nbsp;·&nbsp; 2+: ${stats.twoPlus||0}</div>
      ${reasonHTML}
    </div>`;
  });

  modalBody.innerHTML = `
    <h2 style="font-size:20px;color:#ffffff;margin-bottom:6px;">${hourData.hour}</h2>
    <div style="font-size:15px;margin-bottom:16px;">
      <span style="color:rgba(255,255,255,0.6);">Total Lenses Broken:</span>
      <strong style="color:${GC.red};font-size:22px;margin-left:8px;">${hourData.totalBroken||0}</strong>
    </div>
    ${machineHTML || "<p style='color:rgba(255,255,255,0.4);font-size:14px;'>No data</p>"}`;
  modal.classList.add("active");
}

function closeModal() {
  const modal = document.getElementById("chartModal");
  if (modal) modal.classList.remove("active");
}

/* =====================================================
   REASON CHART — GLASSMORPHISM HORIZONTAL BARS
===================================================== */

function buildReasonChart(data) {
  if (!data || !data.topReasons) return;
  const canvas = document.getElementById("reasonChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (reasonChart) reasonChart.destroy();

  const entries = Object.entries(data.topReasons)
    .map(([reason, stats]) => ({ reason, total: stats.total || 0 }))
    .filter(e => e.total > 0)
    .sort((a, b) => b.total - a.total);
  if (!entries.length) return;

  const maxVal     = entries[0].total;
  const grandTotal = entries.reduce((s, e) => s + e.total, 0);
  const fallback   = [GC.red, GC.orange, GC.yellow, GC.teal, GC.purple, GC.blue];
  const colors     = entries.map((e, i) => BREAKAGE_COLOR_MAP[e.reason] || fallback[i % fallback.length]);
  const valueLabels = entries.map(e => `${e.total}  ·  ${((e.total / grandTotal) * 100).toFixed(1)}%`);

  // Solid vivid colors — no fading
  const bgs = colors.map(c => c + "ee");

  reasonChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels  : entries.map(e => e.reason),
      datasets: [{
        label           : "Breakage Count",
        data            : entries.map(e => e.total),
        _valueLabels    : valueLabels,
        backgroundColor : bgs,
        borderColor     : colors.map(c => c + "cc"),
        borderWidth     : 0,
        borderRadius    : 6,
        borderSkipped   : false,
        barPercentage      : 0.65,
        categoryPercentage : 0.88,
      }],
    },
    options: {
      indexAxis          : "y",
      responsive         : true,
      maintainAspectRatio: false,
      animation          : { duration: 500, easing: "easeOutQuart" },
      layout             : { padding: { right: 150, top: 6, bottom: 6 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...GLASS_TOOLTIP,
          callbacks: {
            title    : items => `  ${entries[items[0].dataIndex]?.reason}`,
            label    : ctx  => { const e = entries[ctx.dataIndex]; return [`  Count: ${e.total}`, `  Share: ${((e.total/grandTotal)*100).toFixed(1)}%`]; },
            afterLabel: ctx => ctx.dataIndex === 0 ? `  ⚠  Top contributor` : "",
          },
        },
      },
      scales: {
        x: {
          beginAtZero: true,
          max        : Math.ceil(maxVal * 1.06),
          ticks: { color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "600" }, stepSize: Math.max(1, Math.ceil(maxVal / 8)) },
          grid  : GLASS_GRID,
          border: { display: false },
        },
        y: {
          ticks : { color: "#ffffff", font: { family: CHART_FONT, size: 14, weight: "600" }, padding: 12 },
          grid  : { display: false },
          border: { display: false },
        },
      },
    },
    plugins: [GLOW_PLUGIN, VALUE_LABEL_PLUGIN],
  });
}

/* =====================================================
   MACHINE CHART — GLASSMORPHISM SEVERITY BARS
===================================================== */

function buildMachineChart(data) {
  if (!data || !data.machineTotals) return;
  const canvas = document.getElementById("machineChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (machineChart) machineChart.destroy();

  // Merge coater machineTotals + accessPointTotals (FLEX/tray positions) when present
  const combined = { ...(data.machineTotals || {}) };
  if (data.accessPointTotals) {
    Object.entries(data.accessPointTotals).forEach(([ap, stats]) => {
      if (!combined[ap]) {
        combined[ap] = { jobs: 0, breakLenses: Number(stats.breakLenses || stats.total || 0) };
      } else {
        combined[ap].breakLenses = (combined[ap].breakLenses || 0) + Number(stats.breakLenses || stats.total || 0);
      }
    });
  }

  const entries = Object.entries(combined)
    .map(([machine, stats]) => {
      const jobs  = Number(stats.jobs || 0);
      const broken = Number(stats.breakLenses || 0);
      const total  = jobs * 2;
      const percent = total > 0 ? (broken / total) * 100 : (broken > 0 ? 100 : 0);
      return { machine, jobs, broken, percent };
    })
    .sort((a, b) => {
      if (a.jobs === 0 && b.jobs === 0) return b.broken - a.broken;
      if (a.jobs === 0) return 1;
      if (b.jobs === 0) return -1;
      return b.percent - a.percent;
    });
  if (!entries.length) return;

  const labels     = entries.map(e => formatMachineLabel(e.machine));
  const maxPercent = Math.max(...entries.map(e => e.percent), 0.1);
  const maxJobs    = Math.max(...entries.map(e => e.jobs), 1);

  function mColor(e) {
    if (e.jobs === 0 && e.broken > 0) return "#f472b6";
    if (e.percent >= 6)  return "#ef4444";
    if (e.percent >= 3)  return "#f97316";
    if (e.percent >= 1)  return "#eab308";
    return "#22c55e";
  }
  const brkColors = entries.map(mColor);

  // Solid vivid fills — no fading
  const brkBg = brkColors.map(c => c);
  const jobBg = "#3b82f6";

  machineChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label              : "Breakage %",
          data               : entries.map(e => parseFloat(e.percent.toFixed(2))),
          _valueLabels       : entries.map(e => e.jobs === 0 && e.broken > 0 ? `${e.broken} brk` : `${e.percent.toFixed(1)}%`),
          backgroundColor    : brkBg,
          borderColor        : brkColors.map(c => c),
          borderWidth        : 1,
          borderRadius       : { topLeft: 4, topRight: 4 },
          borderSkipped      : false,
          barPercentage      : 0.55,
          categoryPercentage : 0.75,
          yAxisID            : "y1",
          order              : 1,
        },
        {
          label              : "Total Jobs",
          data               : entries.map(e => e.jobs),
          _valueLabels       : entries.map(e => e.jobs > 0 ? String(e.jobs) : ""),
          backgroundColor    : jobBg,
          borderColor        : "#60a5fa",
          borderWidth        : 1,
          borderRadius       : { topLeft: 4, topRight: 4 },
          borderSkipped      : false,
          barPercentage      : 0.55,
          categoryPercentage : 0.75,
          yAxisID            : "y2",
          order              : 2,
        },
      ],
    },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      animation          : { duration: 500, easing: "easeOutQuart" },
      interaction        : { mode: "index", intersect: false },
      plugins: {
        legend: {
          display : true,
          position: "top",
          labels: {
            color        : "#ffffff",
            font         : { family: CHART_FONT, size: 15, weight: "600" },
            usePointStyle: true,
            pointStyle   : "circle",
            padding      : 28,
            boxWidth     : 14,
            boxHeight    : 14,
            generateLabels: () => [
              { text: "Breakage %", fillStyle: "#ef4444", strokeStyle: "#ef4444", fontColor: "#ffffff", pointStyle: "circle", hidden: false, datasetIndex: 0 },
              { text: "Total Jobs",  fillStyle: "#3b82f6", strokeStyle: "#3b82f6", fontColor: "#ffffff", pointStyle: "circle", hidden: false, datasetIndex: 1 },
            ],
          },
        },
        tooltip: {
          ...GLASS_TOOLTIP,
          callbacks: {
            title    : items => `  ${items[0].label}`,
            label    : ctx  => {
              const e = entries[ctx.dataIndex];
              if (ctx.datasetIndex === 0) {
                if (e.jobs === 0 && e.broken > 0) return [`  ⚠ ${e.broken} lenses — no job data`];
                return [`  Breakage: ${e.percent.toFixed(2)}%`, `  Broken: ${e.broken} / ${e.jobs * 2}`];
              }
              return [`  Jobs: ${e.jobs}`];
            },
            afterBody: items => ["", `  Rank #${items[0].dataIndex + 1} by breakage`],
          },
        },
      },
      scales: {
        x: {
          ticks : { color: "rgba(160,200,240,0.6)", font: { family: CHART_FONT, size: 12, weight: "500" }, padding: 6 },
          grid  : { display: false },
          border: { display: false },
        },
        y1: {
          position   : "left",
          beginAtZero: true,
          suggestedMax: Math.ceil(maxPercent * 1.4),
          ticks: { color: "rgba(248,113,113,0.6)", font: { family: CHART_MONO, size: 10 }, callback: v => v + "%" },
          grid  : { ...GLASS_GRID, color: "rgba(248,113,113,0.05)" },
          border: { display: false },
          title : { display: true, text: "Breakage %", color: "#ff8888", font: { family: CHART_FONT, size: 14, weight: "700" } },
        },
        y2: {
          position   : "right",
          beginAtZero: true,
          suggestedMax: Math.ceil(maxJobs * 1.25),
          ticks: { color: "#93c5fd", font: { family: CHART_FONT, size: 13 } },
          grid  : { drawOnChartArea: false },
          border: { display: false },
          title : { display: true, text: "Jobs", color: "#93c5fd", font: { family: CHART_FONT, size: 14, weight: "700" } },
        },
      },
    },
    plugins: [GLOW_PLUGIN, VALUE_LABEL_PLUGIN],
  });
}

/* =====================================================
   AI INSIGHTS
===================================================== */

function buildInsights(data) {
  const el = document.getElementById("insightBox");
  if (!el || !data) return;

  const summary       = data.summary       || {};
  const hourly        = data.hourly        || [];
  const machineTotals = data.machineTotals || {};
  const topReasons    = data.topReasons    || {};
  const insights      = [];

  let peakHour = "-", peakValue = 0;
  hourly.forEach(h => { const v = h.totalBroken||0; if (v > peakValue) { peakValue = v; peakHour = h.hour; } });
  if (peakValue > 0) insights.push({ cls: "insight-peak", text: `🔥 Peak: <b>${peakHour}</b> (${peakValue} lenses)` });

  let topReason = "-", topReasonCount = 0;
  Object.entries(topReasons).forEach(([r, s]) => { if ((s.total||0) > topReasonCount) { topReason = r; topReasonCount = s.total||0; } });
  if (topReasonCount > 0) insights.push({ cls: "insight-warn", text: `⚠️ Top Issue: <b>${topReason}</b> (${topReasonCount})` });

  let worstMachine = "-", worstPercent = 0;
  Object.entries(machineTotals).forEach(([m, s]) => {
    const jobs = s.jobs||0, broken = s.breakLenses||0;
    if (jobs > 0) { const pct = (broken/(jobs*2))*100; if (pct > worstPercent) { worstPercent = pct; worstMachine = m; } }
  });
  if (worstPercent > 0) insights.push({ cls: "insight-machine", text: `🛠️ Worst Machine: <b>${formatMachineLabel(worstMachine)}</b> ${worstPercent.toFixed(1)}%` });

  if (summary.flowHealth) {
    const { delayed = 0, healthy = 0 } = summary.flowHealth;
    if (delayed > healthy)  insights.push({ cls: "insight-warn", text: `🚨 Flow Risk: Delayed > Healthy` });
    else if (delayed > 0)   insights.push({ cls: "insight-flow", text: `⚡ Flow Warning: ${delayed} delayed` });
    else                    insights.push({ cls: "insight-ok",   text: `✅ Flow Stable` });
  }

  if (hourly.length >= 3) {
    const last = hourly.slice(-3).map(h => h.totalBroken||0);
    if (last[2] > last[1] && last[1] > last[0]) insights.push({ cls: "insight-trend", text: `📈 Breakage Rising (last 3 hrs)` });
    else if (last[2] < last[1] && last[1] < last[0]) insights.push({ cls: "insight-trend", text: `📉 Breakage Falling (last 3 hrs)` });
  }

  if (summary.totalBreakLenses === 0) insights.push({ cls: "insight-ok", text: `🟢 Zero Breakage` });

  el.innerHTML = insights.length
    ? insights.map(i => `<span class="insight-item ${i.cls}">${i.text}</span>`).join("")
    : "<span style='color:var(--text-dim);font-size:13px;'>No insights available</span>";
}

/* =====================================================
   SUMMARY TAB LOADERS
===================================================== */

function showSummaryLoader(tab) {
  const el = document.getElementById(tab + "Loader");
  if (!el) return;
  el.style.display = "flex";
  // Update text based on tab
  const textEl = el.querySelector(".sl-text");
  if (textEl) textEl.textContent = tab === "weekly" ? "Fetching week data..." : "Building Daily Summary...";
}

function hideSummaryLoader(tab) {
  const el = document.getElementById(tab + "Loader");
  if (el) el.style.display = "none";
}

/* =====================================================
/* =====================================================
   DAILY SUMMARY
===================================================== */

function buildDailySummary(data) {
  if (!data) return;
  const summary       = data.summary       || {};
  const hourly        = data.hourly        || [];
  const machineTotals = data.machineTotals || {};
  const topReasons    = data.topReasons    || {};

  // ── Date badge ──
  const dateBadge = document.getElementById("dailySummaryDate");
  if (dateBadge) {
    dateBadge.textContent = currentDate
      ? "📅 " + currentDate
      : "📅 Today · Live";
  }

  // ── Scorecard ──
  const brokenVal = summary.totalBreakLenses || 0;
  const pctVal    = parseFloat(summary.breakPercent) || 0;
  const avgHrs    = parseFloat(summary.avgBreakTimeHours) || 0;
  const allFlowPts = hourly.flatMap(h => (h.flowPoints||[]).filter(p => p.flow > 0 && p.flow <= 360));
  const avgFlowMins = allFlowPts.length
    ? Math.round(allFlowPts.reduce((s,p)=>s+p.flow,0)/allFlowPts.length)
    : null;

  function setSc(id, val, cls) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = val;
    if (cls) el.className = "sc-val " + cls;
  }
  setSc("scJobs",   summary.totalJobs ?? "—", "teal");
  setSc("scLenses", (summary.totalLenses||0).toLocaleString(), "");
  setSc("scBroken", brokenVal, brokenVal === 0 ? "green" : brokenVal <= 20 ? "yellow" : "red");
  setSc("scPct",    pctVal.toFixed(2) + "%", pctVal < 2 ? "green" : pctVal < 4 ? "yellow" : "red");
  setSc("scPeak",   summary.peakHour ?? "—", "yellow");
  setSc("scFlow",   avgFlowMins !== null ? avgFlowMins + "m avg" : "—", avgFlowMins === null ? "" : avgFlowMins <= 15 ? "green" : avgFlowMins <= 30 ? "yellow" : "red");

  // ── Flow Health Bar ──
  const fhEl = document.getElementById("dailyFlowHealth");
  if (fhEl && summary.flowHealth) {
    const fh = summary.flowHealth;
    const total = (fh.healthy||0) + (fh.watch||0) + (fh.delayed||0) + (fh.overnight||0) || 1;
    const pH = ((fh.healthy||0)/total*100).toFixed(1);
    const pW = ((fh.watch||0)/total*100).toFixed(1);
    const pD = ((fh.delayed||0)/total*100).toFixed(1);
    const pO = ((fh.overnight||0)/total*100).toFixed(1);
    fhEl.innerHTML = `
      <div class="fh-bar-wrap">
        <div class="fh-segment" style="flex:${pH};background:#4ade80;"></div>
        <div class="fh-segment" style="flex:${pW};background:#fbbf24;"></div>
        <div class="fh-segment" style="flex:${pD};background:#f87171;"></div>
        <div class="fh-segment" style="flex:${pO};background:#38bdf8;"></div>
      </div>
      <div class="fh-labels">
        <span class="fh-label"><span class="fh-dot" style="background:#4ade80"></span>Healthy <span class="fh-count" style="color:#4ade80">${fh.healthy||0}</span></span>
        <span class="fh-label"><span class="fh-dot" style="background:#fbbf24"></span>Watch <span class="fh-count" style="color:#fbbf24">${fh.watch||0}</span></span>
        <span class="fh-label"><span class="fh-dot" style="background:#f87171"></span>Delayed <span class="fh-count" style="color:#f87171">${fh.delayed||0}</span></span>
        <span class="fh-label"><span class="fh-dot" style="background:#38bdf8"></span>Overnight <span class="fh-count" style="color:#38bdf8">${fh.overnight||0}</span></span>
        <span class="fh-label" style="margin-left:auto;color:rgba(255,255,255,0.35);font-size:11px;">${total} total flow records</span>
      </div>`;
  }

  // ── Top Machines ──
  const machEl = document.getElementById("dailyMachinesPanel");
  if (machEl) {
    const machEntries = Object.entries(machineTotals)
      .map(([m,s]) => ({ m, broken: s.breakLenses||0, jobs: s.jobs||0 }))
      .filter(e => e.broken > 0)
      .sort((a,b) => b.broken - a.broken)
      .slice(0, 8);
    const maxBrk = machEntries[0]?.broken || 1;
    machEl.innerHTML = machEntries.length
      ? machEntries.map((e,i) => {
          const col = e.broken >= 10 ? "#f87171" : e.broken >= 4 ? "#fbbf24" : "#4ade80";
          return `<div class="summary-row">
            <span class="summary-row-rank">${i+1}</span>
            <span class="summary-row-name" style="color:${col}">${formatMachineLabel(e.m)}</span>
            <div class="summary-row-bar-wrap"><div class="summary-row-bar" style="width:${(e.broken/maxBrk*100).toFixed(1)}%;background:${col};"></div></div>
            <span class="summary-row-val" style="color:${col}">${e.broken}</span>
          </div>`;
        }).join("")
      : `<div style="padding:16px;color:rgba(255,255,255,0.3);font-size:13px;">No breakage recorded</div>`;
  }

  // ── Top Reasons ──
  const resEl = document.getElementById("dailyReasonsPanel");
  if (resEl) {
    const resEntries = Object.entries(topReasons)
      .map(([r,s]) => ({ r, total: s.total||0 }))
      .filter(e => e.total > 0)
      .sort((a,b) => b.total - a.total)
      .slice(0, 8);
    const maxRes = resEntries[0]?.total || 1;
    resEl.innerHTML = resEntries.length
      ? resEntries.map((e,i) => {
          const col = BREAKAGE_COLOR_MAP[e.r] || GC.blue;
          return `<div class="summary-row">
            <span class="summary-row-rank">${i+1}</span>
            <span class="summary-row-name">${e.r}</span>
            <div class="summary-row-bar-wrap"><div class="summary-row-bar" style="width:${(e.total/maxRes*100).toFixed(1)}%;background:${col};"></div></div>
            <span class="summary-row-val" style="color:${col}">${e.total}</span>
          </div>`;
        }).join("")
      : `<div style="padding:16px;color:rgba(255,255,255,0.3);font-size:13px;">No breakage recorded</div>`;
  }

  // ── Hourly Timeline ──
  const tlEl = document.getElementById("dailyTimeline");
  if (tlEl) {
    const sorted = [...hourly].sort((a,b) => new Date("1/1/2000 "+a.hour) - new Date("1/1/2000 "+b.hour));
    const maxJobs = Math.max(...sorted.map(h => h.coatingJobs||0), 1);
    tlEl.innerHTML = sorted.map(h => {
      const brk   = h.totalBroken || 0;
      const jobs  = h.coatingJobs || 0;
      const cls   = jobs === 0 ? "tl-empty" : brk === 0 ? "tl-ok" : brk <= 3 ? "tl-warn" : "tl-danger";
      const brkTxt = brk > 0 ? `<div class="tl-brk">🔴 ${brk} broken</div>` : `<div class="tl-brk" style="color:rgba(74,222,128,0.5)">✓ clean</div>`;
      return `<div class="tl-cell ${cls}">
        <div class="tl-hour">${h.hour}</div>
        <div class="tl-jobs">${jobs}</div>
        <div style="font-size:9px;color:rgba(255,255,255,0.3);margin-bottom:2px;">jobs</div>
        ${brkTxt}
      </div>`;
    }).join("") || `<div style="color:rgba(255,255,255,0.3);padding:16px;">No hourly data</div>`;
  }

  // ── Narrative ──
  const narEl = document.getElementById("dailyNarrative");
  if (narEl) {
    const totalJobs    = summary.totalJobs || 0;
    const totalLenses  = summary.totalLenses || 0;
    const flowHealth   = summary.flowHealth || {};
    const healthyPct   = totalJobs > 0 ? ((flowHealth.healthy||0)/totalJobs*100).toFixed(0) : 0;
    const delayedCount = flowHealth.delayed || 0;
    const overnightCount = flowHealth.overnight || 0;

    // Find peak hour
    let peakHour = "—", peakBrk = 0;
    hourly.forEach(h => { if ((h.totalBroken||0) > peakBrk) { peakBrk = h.totalBroken||0; peakHour = h.hour; } });

    // Find busiest hour
    let busiestHour = "—", busiestJobs = 0;
    hourly.forEach(h => { if ((h.coatingJobs||0) > busiestJobs) { busiestJobs = h.coatingJobs||0; busiestHour = h.hour; } });

    // Top machine
    let topMach = null, topMachBrk = 0;
    Object.entries(machineTotals).forEach(([m,s]) => { if ((s.breakLenses||0) > topMachBrk) { topMachBrk = s.breakLenses||0; topMach = m; } });

    const statusWord = pctVal === 0 ? `<span class="nar-good">zero breakage</span>`
      : pctVal < 2 ? `<span class="nar-good">healthy breakage rate of ${pctVal.toFixed(2)}%</span>`
      : pctVal < 4 ? `<span class="nar-warn">elevated breakage rate of ${pctVal.toFixed(2)}%</span>`
      : `<span class="nar-danger">high breakage rate of ${pctVal.toFixed(2)}%</span>`;

    narEl.innerHTML = `
      <strong>Summary:</strong> ${totalJobs.toLocaleString()} coating jobs processed across ${totalLenses.toLocaleString()} lenses with ${statusWord}.
      ${brokenVal > 0 ? `<strong>${brokenVal} lenses were broken</strong>, peaking at <strong>${peakHour}</strong> (${peakBrk} lenses).` : ""}
      The busiest hour was <strong>${busiestHour}</strong> with <strong>${busiestJobs} jobs</strong>.
      ${topMach && topMachBrk > 0 ? `Top breakage machine was <strong style="color:#38bdf8">${formatMachineLabel(topMach)}</strong> with ${topMachBrk} lenses.` : ""}
      Flow health: <span class="nar-good">${flowHealth.healthy||0} healthy</span> ·
      <span class="nar-warn">${flowHealth.watch||0} watch</span> ·
      ${delayedCount > 0 ? `<span class="nar-danger">${delayedCount} delayed</span>` : `<span class="nar-good">0 delayed</span>`}
      ${overnightCount > 0 ? ` · <span style="color:#38bdf8">${overnightCount} overnight carryover</span>.` : "."}
      ${avgFlowMins !== null ? `Average detaper-to-coater flow time was <strong>${avgFlowMins} minutes</strong>.` : ""}`;
  }
}

/* =====================================================
   WEEKLY SUMMARY
===================================================== */

let weeklyTrendChart = null;
const weeklyCache    = {};  // "M/D/YYYY" → data

async function buildWeeklySummary() {
  const weEl = document.getElementById("weekEndDate");
  if (!weEl || !weEl.value) return;

  const endDate  = new Date(weEl.value + "T00:00:00");
  const rangeEl  = document.getElementById("weekRangeLabel");

  // Build array of 7 days ending on selected date
  const days = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(endDate);
    d.setDate(d.getDate() - i);
    days.push(d);
  }

  const startLabel = days[0].toLocaleDateString("en-US", { month:"short", day:"numeric" });
  const endLabel   = days[6].toLocaleDateString("en-US", { month:"short", day:"numeric", year:"numeric" });
  if (rangeEl) rangeEl.textContent = startLabel + " – " + endLabel;

  // Convert to API date strings
  const apiDates = days.map(d => `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`);
  const todayApi = (() => { const n = new Date(); return `${n.getMonth()+1}/${n.getDate()}/${n.getFullYear()}`; })();

  // Show loading state
  document.getElementById("weeklyDayGrid").innerHTML    = `<div style="color:rgba(255,255,255,0.35);font-size:13px;padding:12px;">Loading week data...</div>`;
  document.getElementById("weeklyMachinesPanel").innerHTML = "";
  document.getElementById("weeklyReasonsPanel").innerHTML  = "";

  // Fetch missing days
  await Promise.all(apiDates.map(async apiDate => {
    if (weeklyCache[apiDate]) return;
    if (apiDate === todayApi && dashboardData) { weeklyCache[apiDate] = dashboardData; return; }
    try {
      const res = await fetch(`${API_URL}?mode=processed&date=${encodeURIComponent(apiDate)}`);
      weeklyCache[apiDate] = await res.json();
    } catch(e) { weeklyCache[apiDate] = null; }
  }));

  // Aggregate weekly totals
  let wJobs = 0, wLenses = 0, wBroken = 0, wFlowSum = 0, wFlowCount = 0;
  let bestDay = null, bestPct = Infinity, worstDay = null, worstPct = -1;
  const machWeekly = {};
  const resWeekly  = {};
  let daysWithData = 0;
  const dailyBreakage = [];

  apiDates.forEach((apiDate, idx) => {
    const d    = days[idx];
    const data = weeklyCache[apiDate];
    if (!data?.summary) { dailyBreakage.push(null); return; }
    daysWithData++;
    const s = data.summary;
    wJobs   += s.totalJobs   || 0;
    wLenses += s.totalLenses || 0;
    wBroken += s.totalBreakLenses || 0;

    // Collect flow for weekly avg
    (data.hourly || []).forEach(h => {
      (h.flowPoints||[]).forEach(p => {
        if (p.flow > 0 && p.flow <= 360) { wFlowSum += p.flow; wFlowCount++; }
      });
    });

    const pct = parseFloat(s.breakPercent) || 0;
    dailyBreakage.push({ date: d, apiDate, pct, broken: s.totalBreakLenses||0, jobs: s.totalJobs||0, lenses: s.totalLenses||0, summary: s });
    if (pct < bestPct)  { bestPct  = pct;  bestDay  = idx; }
    if (pct > worstPct) { worstPct = pct;  worstDay = idx; }

    // Accumulate machine totals
    Object.entries(data.machineTotals || {}).forEach(([m,ms]) => {
      if (!machWeekly[m]) machWeekly[m] = { broken: 0, jobs: 0 };
      machWeekly[m].broken += ms.breakLenses || 0;
      machWeekly[m].jobs   += ms.jobs || 0;
    });

    // Accumulate reason totals
    Object.entries(data.topReasons || {}).forEach(([r,rs]) => {
      if (!resWeekly[r]) resWeekly[r] = 0;
      resWeekly[r] += rs.total || 0;
    });
  });

  const wAvgPct  = wLenses > 0 ? ((wBroken / wLenses) * 100).toFixed(2) : "0.00";
  const wAvgFlow = wFlowCount > 0 ? Math.round(wFlowSum / wFlowCount) : null;

  // ── Weekly scorecard ──
  function setWc(id, val, cls) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = val;
    if (cls) el.className = "sc-val " + cls;
  }
  setWc("wcDays",   daysWithData, "teal");
  setWc("wcJobs",   wJobs.toLocaleString(), "");
  setWc("wcBroken", wBroken, wBroken === 0 ? "green" : wBroken <= 50 ? "yellow" : "red");
  setWc("wcPct",    wAvgPct + "%", parseFloat(wAvgPct) < 2 ? "green" : parseFloat(wAvgPct) < 4 ? "yellow" : "red");
  setWc("wcFlow",   wAvgFlow !== null ? wAvgFlow + "m" : "—", wAvgFlow===null?"":wAvgFlow<=15?"green":wAvgFlow<=30?"yellow":"red");
  if (bestDay !== null && dailyBreakage[bestDay]) {
    const bd = dailyBreakage[bestDay];
    setWc("wcBest", bd.date.toLocaleDateString("en-US",{month:"short",day:"numeric"}) + "\n" + bd.pct.toFixed(2) + "%", "green");
  }
  if (worstDay !== null && dailyBreakage[worstDay]) {
    const wd2 = dailyBreakage[worstDay];
    setWc("wcWorst", wd2.date.toLocaleDateString("en-US",{month:"short",day:"numeric"}) + "\n" + wd2.pct.toFixed(2) + "%", "red");
  }

  // ── Day-by-day grid ──
  const gridEl = document.getElementById("weeklyDayGrid");
  if (gridEl) {
    const maxBrk = Math.max(...dailyBreakage.filter(Boolean).map(d => d.broken), 1);
    const DOW = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
    gridEl.innerHTML = dailyBreakage.map((d, i) => {
      if (!d) {
        return `<div class="wd-card wd-nodata">
          <div class="wd-dow">${DOW[days[i].getDay()]}</div>
          <div class="wd-date">${days[i].toLocaleDateString("en-US",{month:"short",day:"numeric"})}</div>
          <div style="font-size:12px;color:rgba(255,255,255,0.25);margin-top:8px;">No data</div>
        </div>`;
      }
      const isBest  = i === bestDay;
      const isWorst = i === worstDay && d.broken > 0;
      const pctCol  = d.pct < 2 ? "#4ade80" : d.pct < 4 ? "#fbbf24" : "#f87171";
      const fillPct = (d.broken / maxBrk * 100).toFixed(1);
      return `<div class="wd-card${isBest?" wd-best":isWorst?" wd-worst":""}">
        ${isBest  ? `<div class="wd-badge best">Best</div>`  : ""}
        ${isWorst ? `<div class="wd-badge worst">Worst</div>` : ""}
        <div class="wd-dow">${DOW[d.date.getDay()]}</div>
        <div class="wd-date">${d.date.toLocaleDateString("en-US",{month:"short",day:"numeric"})}</div>
        <div class="wd-stat-row" style="margin-top:8px;">
          <div><div class="wd-stat-label">Jobs</div><div class="wd-stat-val">${d.jobs}</div></div>
          <div><div class="wd-stat-label">Broken</div><div class="wd-stat-val" style="color:${d.broken>0?pctCol:"#4ade80"}">${d.broken}</div></div>
          <div><div class="wd-stat-label">Brkg%</div><div class="wd-stat-val" style="color:${pctCol}">${d.pct.toFixed(2)}%</div></div>
        </div>
        <div class="wd-pct-bar"><div class="wd-pct-fill" style="width:${fillPct}%;background:${pctCol};"></div></div>
      </div>`;
    }).join("");
  }

  // ── Weekly machines leaderboard ──
  const wMachEl = document.getElementById("weeklyMachinesPanel");
  if (wMachEl) {
    const entries = Object.entries(machWeekly)
      .map(([m,s]) => ({ m, broken: s.broken }))
      .filter(e => e.broken > 0)
      .sort((a,b) => b.broken - a.broken)
      .slice(0,8);
    const maxB = entries[0]?.broken || 1;
    wMachEl.innerHTML = entries.length
      ? entries.map((e,i) => {
          const col = e.broken >= 20 ? "#f87171" : e.broken >= 8 ? "#fbbf24" : "#4ade80";
          return `<div class="summary-row">
            <span class="summary-row-rank">${i+1}</span>
            <span class="summary-row-name" style="color:${col}">${formatMachineLabel(e.m)}</span>
            <div class="summary-row-bar-wrap"><div class="summary-row-bar" style="width:${(e.broken/maxB*100).toFixed(1)}%;background:${col};"></div></div>
            <span class="summary-row-val" style="color:${col}">${e.broken}</span>
          </div>`;
        }).join("")
      : `<div style="padding:16px;color:rgba(255,255,255,0.3);font-size:13px;">No breakage this week</div>`;
  }

  // ── Weekly reasons leaderboard ──
  const wResEl = document.getElementById("weeklyReasonsPanel");
  if (wResEl) {
    const entries = Object.entries(resWeekly)
      .filter(([,v]) => v > 0)
      .sort(([,a],[,b]) => b - a)
      .slice(0,8);
    const maxR = entries[0]?.[1] || 1;
    wResEl.innerHTML = entries.length
      ? entries.map(([r,v],i) => {
          const col = BREAKAGE_COLOR_MAP[r] || GC.blue;
          return `<div class="summary-row">
            <span class="summary-row-rank">${i+1}</span>
            <span class="summary-row-name">${r}</span>
            <div class="summary-row-bar-wrap"><div class="summary-row-bar" style="width:${(v/maxR*100).toFixed(1)}%;background:${col};"></div></div>
            <span class="summary-row-val" style="color:${col}">${v}</span>
          </div>`;
        }).join("")
      : `<div style="padding:16px;color:rgba(255,255,255,0.3);font-size:13px;">No breakage this week</div>`;
  }

  // ── Weekly sparkline trend chart ──
  const tCanvas = document.getElementById("weeklyTrendChart");
  if (tCanvas) {
    if (weeklyTrendChart) { weeklyTrendChart.destroy(); weeklyTrendChart = null; }
    const labels  = days.map(d => d.toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}));
    const brkData = dailyBreakage.map(d => d ? d.broken : null);
    const pctData = dailyBreakage.map(d => d ? parseFloat(d.pct) : null);
    const jobData = dailyBreakage.map(d => d ? d.jobs : null);
    const ctx2    = tCanvas.getContext("2d");
    const chartH  = tCanvas.clientHeight || 200;

    weeklyTrendChart = new Chart(ctx2, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label           : "Lenses Broken",
            type            : "bar",
            data            : brkData,
            backgroundColor : brkData.map(v => v === null ? "transparent" : v === 0 ? "rgba(74,222,128,0.5)" : v <= 5 ? "rgba(251,191,36,0.6)" : "rgba(248,113,113,0.65)"),
            borderColor     : brkData.map(v => v === null ? "transparent" : v === 0 ? "#4ade80" : v <= 5 ? "#fbbf24" : "#f87171"),
            borderWidth     : 1,
            borderRadius    : { topLeft: 5, topRight: 5 },
            yAxisID         : "yBrk",
            order           : 2,
          },
          {
            label           : "Coating Jobs",
            type            : "line",
            data            : jobData,
            borderColor     : "rgba(251,191,36,0.45)",
            backgroundColor : "transparent",
            borderWidth     : 1.5,
            borderDash      : [4,4],
            tension         : 0.35,
            pointRadius     : 3,
            pointBackgroundColor: "#fbbf24",
            fill            : false,
            yAxisID         : "yJobs",
            order           : 1,
          },
        ],
      },
      options: {
        responsive          : true,
        maintainAspectRatio : false,
        animation           : { duration: 600, easing: "easeOutQuart" },
        interaction         : { mode: "index", intersect: false },
        plugins: {
          legend : GLASS_LEGEND,
          tooltip: {
            ...GLASS_TOOLTIP,
            callbacks: {
              title : items => `  ${items[0].label}`,
              label : ctx  => {
                if (ctx.dataset.label === "Coating Jobs") return `  Jobs: ${ctx.raw ?? "—"}`;
                const v = ctx.raw;
                if (v === null) return null;
                return `  Broken: ${v}`;
              },
            },
          },
        },
        scales: {
          x: { ...axisStyle(), ticks: { color: "#ffffff", font: { family: CHART_FONT, size: 12 }, maxRotation: 0 }, grid: GLASS_GRID },
          yBrk: {
            type: "linear", position: "left", beginAtZero: true,
            ticks: { color: "#f87171", font: { family: CHART_FONT, size: 12 }, stepSize: 1 },
            grid: { ...GLASS_GRID, color: "rgba(248,113,113,0.05)" },
            title: { display: true, text: "Broken", color: "#f87171", font: { family: CHART_FONT, size: 12, weight:"700" } },
          },
          yJobs: {
            type: "linear", position: "right", beginAtZero: true,
            ticks: { color: "#fbbf24", font: { family: CHART_FONT, size: 12 } },
            grid: { drawOnChartArea: false },
            title: { display: true, text: "Jobs", color: "#fbbf24", font: { family: CHART_FONT, size: 12, weight:"700" } },
          },
        },
      },
      plugins: [GLOW_PLUGIN],
    });
  }
}
/* =====================================================
   REBUILD ALL CHARTS
===================================================== */

function rebuildAllCharts() {
  [trendChart, reasonChart, machineChart, flowChart].forEach(c => { if (c) c.destroy(); });
  trendChart = reasonChart = machineChart = flowChart = null;
  const activeTab = document.querySelector(".tab.active");
  if (!activeTab) return;
  const tabText = activeTab.textContent.toLowerCase();
  if (tabText.includes("trend"))        buildTrendCharts();
  else if (tabText.includes("reason"))  buildReasonChart(dashboardData);
  else if (tabText.includes("machine")) buildMachineChart(dashboardMachine || dashboardData);
  if (document.getElementById("flowChart")) buildFlowChart(dashboardProcessed);
}

/* =====================================================
   REFRESH COUNTDOWN
===================================================== */

function startRefreshCountdown() {
  if (refreshTimerHandle) clearInterval(refreshTimerHandle);
  const timerEl = document.getElementById("refreshTimer");
  if (!timerEl) return;
  refreshCountdown = refreshInterval;
  refreshTimerHandle = setInterval(() => {
    if (currentDate !== null) { timerEl.textContent = "History Mode"; return; }
    const m = Math.floor(refreshCountdown / 60), s = refreshCountdown % 60;
    timerEl.textContent = `Refresh in ${m}:${s.toString().padStart(2, "0")}`;
    refreshCountdown--;
    if (refreshCountdown < 0) refreshCountdown = refreshInterval;
  }, 1000);
}

setInterval(() => {
  if (currentDate === null) { loadDashboard(); startRefreshCountdown(); }
}, refreshInterval * 1000);

/* =====================================================
   NAVIGATION
===================================================== */

function goBack() { window.location.href = "index.html"; }

/* =====================================================
   MULTI-DATE COMPARISON
===================================================== */

let compareSlots      = [];
let compareDataCache  = {};   // { "M/D/YYYY": full API response }
let compareMetric     = "broken";
let compareTrendChart = null;
const MAX_COMPARE     = 7;

// Palette for each date line
const COMPARE_PALETTE = [
  "#60a5fa","#4ade80","#f87171","#fbbf24","#a78bfa","#2dd4bf","#fb923c"
];

function compareInit() {
  const today = new Date();
  const yest  = new Date(today);
  yest.setDate(yest.getDate() - 1);
  compareSlots = [];
  compareAddSlot(formatDateInput(yest));
  compareAddSlot(formatDateInput(today));
  renderCompareDateInputs();
}

function formatDateInput(dateObj) {
  // Returns YYYY-MM-DD for <input type="date">
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth()+1).padStart(2,"0");
  const d = String(dateObj.getDate()).padStart(2,"0");
  return `${y}-${m}-${d}`;
}

function inputToApiDate(inputVal) {
  // Converts YYYY-MM-DD → M/D/YYYY for API
  if (!inputVal) return null;
  const [y, m, d] = inputVal.split("-");
  return `${parseInt(m)}/${parseInt(d)}/${y}`;
}

function apiDateToLabel(apiDate) {
  // Converts M/D/YYYY → short label like "Apr 4"
  if (!apiDate) return "--";
  const parts = apiDate.split("/");
  if (parts.length < 3) return apiDate;
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${months[parseInt(parts[0])-1]} ${parseInt(parts[1])}`;
}

function compareDateToApi(inputVal) {
  return inputToApiDate(inputVal);
}

function compareAddSlot(prefillValue) {
  if (compareSlots.length >= MAX_COMPARE) return;
  compareSlots.push(prefillValue || "");
  renderCompareDateInputs();
  const addBtn = document.getElementById("compareAddBtn");
  if (addBtn) addBtn.disabled = compareSlots.length >= MAX_COMPARE;
}

function compareRemoveSlot(idx) {
  compareSlots.splice(idx, 1);
  renderCompareDateInputs();
  const addBtn = document.getElementById("compareAddBtn");
  if (addBtn) addBtn.disabled = compareSlots.length >= MAX_COMPARE;
}

function renderCompareDateInputs() {
  const container = document.getElementById("compareDateInputs");
  if (!container) return;
  container.innerHTML = compareSlots.map((val, i) => `
    <div class="compare-date-slot" id="compareSlot${i}">
      <input type="date" value="${val}" onchange="compareSlots[${i}]=this.value" />
      <button class="compare-date-remove" onclick="compareRemoveSlot(${i})" title="Remove">×</button>
    </div>
  `).join("");
}

function setCompareMetric(metric, btn) {
  compareMetric = metric;
  document.querySelectorAll(".compare-metric-btn").forEach(b => b.classList.remove("active"));
  if (btn) btn.classList.add("active");
  // Rebuild chart with new metric if data already loaded
  if (Object.keys(compareDataCache).length) buildCompareChart();
}

async function compareRun() {
  const tableEl    = document.getElementById("compareTable");
  const chartWrap  = document.getElementById("compareTrendWrap");
  if (!tableEl) return;

  const filledSlots = compareSlots.filter(s => s && s.trim());
  if (!filledSlots.length) {
    tableEl.innerHTML = `<div class="compare-empty">Add at least one date to compare.</div>`;
    return;
  }

  tableEl.innerHTML = `<div class="compare-loading"><div class="compare-spinner"></div>Loading ${filledSlots.length} date${filledSlots.length > 1 ? "s" : ""}...</div>`;

  const todayApiDate = (() => {
    const n = new Date();
    return `${n.getMonth()+1}/${n.getDate()}/${n.getFullYear()}`;
  })();

  const apiDates = filledSlots.map(inputToApiDate);

  // Fetch each date (use cache to avoid re-fetching)
  await Promise.all(apiDates.map(async (apiDate) => {
    if (!apiDate || compareDataCache[apiDate]) return;
    // Check if it's today and we already have live data
    if (apiDate === todayApiDate && dashboardData) {
      compareDataCache[apiDate] = dashboardData;
      return;
    }
    try {
      const res  = await fetch(`${API_URL}?mode=processed&date=${encodeURIComponent(apiDate)}`);
      compareDataCache[apiDate] = await res.json();
    } catch(e) {
      compareDataCache[apiDate] = null;
    }
  }));

  const validDates = apiDates.filter(d => d && compareDataCache[d]?.summary);
  if (!validDates.length) {
    tableEl.innerHTML = `<div class="compare-empty">No data found. Make sure the dates exist in your history sheets.</div>`;
    return;
  }

  // Build trend chart
  buildCompareChart(validDates, todayApiDate);

  // Build summary table
  buildCompareTable(validDates, todayApiDate, tableEl);
}

function buildCompareChart(validDates, todayApiDate) {
  const canvas = document.getElementById("compareTrendChart");
  if (!canvas) return;
  if (compareTrendChart) { compareTrendChart.destroy(); compareTrendChart = null; }

  // Use passed dates or all cached dates
  const dates = validDates || Object.keys(compareDataCache).filter(d => compareDataCache[d]?.summary);
  if (!dates.length) return;

  const today = todayApiDate || (() => {
    const n = new Date(); return `${n.getMonth()+1}/${n.getDate()}/${n.getFullYear()}`;
  })();

  // Collect all unique hours across all dates, sorted
  const allHoursSet = new Set();
  dates.forEach(d => {
    (compareDataCache[d]?.hourly || []).forEach(h => allHoursSet.add(h.hour));
  });
  const allHours = [...allHoursSet].sort((a,b) => new Date("1/1/2000 " + a) - new Date("1/1/2000 " + b));

  // Metric extractor per hour
  function getHourMetric(hourObj) {
    if (!hourObj) return null;
    if (compareMetric === "broken") return hourObj.totalBroken || 0;
    if (compareMetric === "jobs")   return hourObj.coatingJobs || 0;
    if (compareMetric === "flow")   return hourObj.avgFlowAll  || null;
    if (compareMetric === "pct") {
      const jobs = hourObj.coatingJobs || 0;
      const brk  = hourObj.totalBroken || 0;
      return jobs > 0 ? parseFloat(((brk / (jobs*2))*100).toFixed(2)) : 0;
    }
    return null;
  }

  const metricLabels = {
    broken: "Lenses Broken",
    pct   : "Breakage %",
    flow  : "Avg Flow Time (min)",
    jobs  : "Coating Jobs",
  };

  const ctx = canvas.getContext("2d");
  const datasets = dates.map((apiDate, i) => {
    const color   = COMPARE_PALETTE[i % COMPARE_PALETTE.length];
    const hourly  = compareDataCache[apiDate]?.hourly || [];
    const hourMap = {};
    hourly.forEach(h => { hourMap[h.hour] = h; });

    const isToday = apiDate === today;
    return {
      label          : apiDateToLabel(apiDate) + (isToday ? " ★" : ""),
      data           : allHours.map(h => getHourMetric(hourMap[h])),
      borderColor    : color,
      backgroundColor: color + "18",
      borderWidth    : isToday ? 2.5 : 1.8,
      borderDash     : isToday ? [] : [],
      tension        : 0.42,
      fill           : false,
      pointRadius    : 3,
      pointHoverRadius: 7,
      pointBackgroundColor: color,
      pointBorderColor: "rgba(8,10,15,0.7)",
      pointBorderWidth: 1.5,
      spanGaps       : true,
    };
  });

  compareTrendChart = new Chart(ctx, {
    type: "line",
    data: { labels: allHours, datasets },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      animation          : { duration: 400, easing: "easeOutQuart" },
      interaction        : { mode: "index", intersect: false },
      plugins: {
        legend: {
          display : true,
          position: "top",
          labels: {
            color        : "#ffffff",
            font         : { family: CHART_FONT, size: 14, weight: "600" },
            usePointStyle: true,
            pointStyle   : "circle",
            padding      : 24,
          },
          onClick(e, item, legend) {
            const meta = legend.chart.getDatasetMeta(item.datasetIndex);
            meta.hidden = !meta.hidden;
            legend.chart.update();
          },
        },
        tooltip: {
          ...GLASS_TOOLTIP,
          callbacks: {
            title : items => `  ${items[0].label}`,
            label : ctx => {
              const v = ctx.raw;
              if (v === null || v === undefined) return null;
              const suffix = compareMetric === "flow" ? "m" : compareMetric === "pct" ? "%" : "";
              return `  ${ctx.dataset.label}: ${v}${suffix}`;
            },
          },
        },
      },
      scales: {
        x: {
          ticks : { color: "rgba(255,255,255,0.75)", font: { family: CHART_MONO, size: 10 }, maxRotation: 0 },
          grid  : { color: "rgba(255,255,255,0.04)" },
          border: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: {
            color   : "rgba(255,255,255,0.75)",
            font    : { family: CHART_MONO, size: 10 },
            callback: v => compareMetric === "pct" ? v + "%" : compareMetric === "flow" ? v + "m" : v,
          },
          grid  : { color: "rgba(255,255,255,0.04)" },
          border: { display: false },
          title: { display: true, text: metricLabels[compareMetric], color: "#ffffff", font: { family: CHART_FONT, size: 14, weight: "700" },
          },
        },
      },
    },
  });
}

function buildCompareTable(validDates, todayApiDate, tableEl) {
  const metrics = [
    { key: "totalJobs",          label: "Total Coating Jobs",  lowerIsBetter: false },
    { key: "totalLenses",        label: "Total Lenses",        lowerIsBetter: false },
    { key: "totalBreakLenses",   label: "Lenses Broken",       lowerIsBetter: true  },
    { key: "breakPercent",       label: "Breakage %",          lowerIsBetter: true,  suffix: "%" },
    { key: "avgDetaper",         label: "Avg Flow Time",       lowerIsBetter: true,  suffix: "m" },
    { key: "peakHour",           label: "Peak Hour",           lowerIsBetter: false, isText: true },
    { key: "aging.sameDay",      label: "Same Day Breakage",   lowerIsBetter: true  },
    { key: "aging.oneDay",       label: "Prev Day Breakage",   lowerIsBetter: true  },
    { key: "aging.twoPlus",      label: "2+ Day Breakage",     lowerIsBetter: true  },
    { key: "flowHealth.healthy", label: "Flow Healthy Jobs",   lowerIsBetter: false },
    { key: "flowHealth.watch",   label: "Flow Watch Jobs",     lowerIsBetter: true  },
    { key: "flowHealth.delayed", label: "Flow Delayed Jobs",   lowerIsBetter: true  },
  ];

  function getVal(summary, key) {
    if (!summary) return null;
    if (key.includes(".")) {
      const [a, b] = key.split(".");
      return summary[a]?.[b] ?? null;
    }
    return summary[key] ?? null;
  }

  let html = `<table>
    <thead><tr>
      <th style="min-width:170px;">Metric</th>
      ${validDates.map((d, i) => {
        const isToday = d === todayApiDate;
        const color   = COMPARE_PALETTE[i % COMPARE_PALETTE.length];
        return `<th class="date-col ${isToday ? "compare-today-col" : ""}"
          style="border-top:2px solid ${color};">
          ${apiDateToLabel(d)}${isToday ? " ★" : ""}
        </th>`;
      }).join("")}
    </tr></thead>
    <tbody>`;

  metrics.forEach(m => {
    const summaries = validDates.map(d => compareDataCache[d]?.summary || null);
    const vals      = summaries.map(s => getVal(s, m.key));
    const nums      = vals.filter(v => v !== null && !isNaN(Number(v))).map(Number);

    let best = null, worst = null;
    if (!m.isText && nums.length > 1) {
      best  = m.lowerIsBetter ? Math.min(...nums) : Math.max(...nums);
      worst = m.lowerIsBetter ? Math.max(...nums) : Math.min(...nums);
    }

    html += `<tr><td class="metric-label">${m.label}</td>`;
    vals.forEach((raw, i) => {
      if (raw === null || raw === undefined) {
        html += `<td style="color:rgba(255,255,255,0.15)">—</td>`;
        return;
      }
      const num     = Number(raw);
      const display = m.isText ? raw : (isNaN(num) ? raw : (m.suffix ? parseFloat(num).toFixed(m.suffix === "%" ? 2 : 0) + m.suffix : num.toLocaleString()));
      const isBest  = !m.isText && !isNaN(num) && nums.length > 1 && num === best;
      const isWorst = !m.isText && !isNaN(num) && nums.length > 1 && num === worst;
      const isToday = validDates[i] === todayApiDate;
      const cls     = [isBest ? "best" : isWorst ? "worst" : "", isToday ? "compare-today-col" : ""].join(" ").trim();
      html += `<td class="${cls}">${display}</td>`;
    });
    html += `</tr>`;
  });

  html += `</tbody></table>
    <div style="padding:8px 14px 6px;font-size:10px;color:var(--dim);">
      <span style="color:var(--green);font-weight:700;">Green</span> = best &nbsp;·&nbsp;
      <span style="color:var(--red);font-weight:700;">Red</span> = worst &nbsp;·&nbsp;
      ★ = today's live data
    </div>`;

  tableEl.innerHTML = html;
}

function compareClear() {
  compareSlots     = [];
  compareDataCache = {};
  if (compareTrendChart) { compareTrendChart.destroy(); compareTrendChart = null; }
  compareAddSlot();
  compareAddSlot();
  const tableEl = document.getElementById("compareTable");
  if (tableEl) tableEl.innerHTML = `<div class="compare-empty">Select dates above and click Compare.</div>`;
}

// ── Legacy: keep today's data for spike alert ──
let previousDayData = null;

async function loadPreviousDay() {
  try {
    const now  = new Date();
    const yest = new Date(now);
    yest.setDate(yest.getDate() - 1);
    const m = yest.getMonth() + 1, d = yest.getDate(), y = yest.getFullYear();
    const dateStr = `${m}/${d}/${y}`;
    const res = await fetch(`${API_URL}?mode=processed&date=${encodeURIComponent(dateStr)}`);
    previousDayData = await res.json();
  } catch(e) { previousDayData = null; }
}

/* =====================================================
   NEON NOIR — SPIKE ALERT BANNER
===================================================== */

const SPIKE_THRESHOLD = 10;

function checkSpikeAlert(hourly) {
  const banner = document.getElementById("alertBanner");
  const text   = document.getElementById("alertText");
  if (!banner || !text) return;

  let spikeHour = null, spikeCount = 0, spikeMachine = "--";
  (hourly || []).forEach(h => {
    if ((h.totalBroken || 0) > spikeCount) {
      spikeCount   = h.totalBroken;
      spikeHour    = h.hour;
      // find top machine
      let topM = null, topMc = 0;
      Object.entries(h.machines || {}).forEach(([m, s]) => {
        if ((s.total || 0) > topMc) { topMc = s.total; topM = m; }
      });
      if (topM) spikeMachine = formatMachineLabel(topM);
    }
  });

  if (spikeCount >= SPIKE_THRESHOLD) {
    text.innerHTML = `<b>BREAKAGE SPIKE DETECTED</b> — ${spikeHour} hit ${spikeCount} lenses · ${spikeMachine} primary driver`;
    banner.style.display = "flex";
  } else {
    banner.style.display = "none";
  }
}

/* =====================================================
   NEON NOIR — EXPORT CSV
===================================================== */

function exportCSV() {
  if (!dashboardData) { alert("No data loaded yet."); return; }
  const hourly = dashboardData.hourly || [];
  const rows   = [["Hour","Total Broken","Coating Jobs","Machines","Primary Machine","Primary Reason"]];

  hourly.forEach(h => {
    const machines = Object.entries(h.machines || {})
      .filter(([,s]) => (s.total||0) > 0)
      .sort(([,a],[,b]) => (b.total||0)-(a.total||0));
    const machineStr = machines.map(([m,s]) => `${formatMachineLabel(m)}x${s.total}`).join(" | ");
    let topM = "", topR = "";
    if (machines.length) {
      topM = formatMachineLabel(machines[0][0]);
      const reasons = Object.entries(machines[0][1].reasons||{}).sort(([,a],[,b])=>(b.total||0)-(a.total||0));
      if (reasons.length) topR = reasons[0][0];
    }
    rows.push([h.hour, h.totalBroken||0, h.coatingJobs||0, machineStr, topM, topR]);
  });

  const csv  = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  const date = currentDate || new Date().toLocaleDateString("en-US").replace(/\//g,"-");
  a.href = url; a.download = `coating-flow-${date}.csv`; a.click();
  URL.revokeObjectURL(url);
}

/* =====================================================
   NEON NOIR — EXPORT PDF (print)
===================================================== */

function exportPDF() {
  // Collect all visible chart canvases
  const canvases = [
    { id: "trendChart",      label: "Breakage Trend"              },
    { id: "flowChart",       label: "Detaper → Coater Flow"       },
    { id: "reasonChart",     label: "Top Breakage Reasons"        },
    { id: "machineChart",    label: "Machine Performance"         },
    { id: "compareTrendChart", label: "Date Comparison Trend"     },
  ];

  const today = new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  const reportDate = currentDate || today;
  const summary = dashboardData?.summary || {};

  // Build HTML for print window
  let chartSections = "";
  canvases.forEach(({ id, label }) => {
    const canvas = document.getElementById(id);
    if (!canvas || canvas.width === 0) return;
    try {
      const img = canvas.toDataURL("image/png", 1.0);
      chartSections += `
        <div class="chart-section">
          <div class="chart-label">${label}</div>
          <img src="${img}" style="width:100%;border-radius:6px;border:1px solid #2a2d35;" />
        </div>`;
    } catch(e) { /* skip cross-origin issues */ }
  });

  const kpis = [
    { label: "Total Jobs",       value: summary.totalJobs       || 0 },
    { label: "Total Lenses",     value: summary.totalLenses     || 0 },
    { label: "Lenses Broken",    value: summary.totalBreakLenses|| 0 },
    { label: "Breakage %",       value: (summary.breakPercent   || 0) + "%" },
    { label: "Peak Hour",        value: summary.peakHour        || "--" },
    { label: "Avg Flow Time",    value: summary.avgDetaper ? summary.avgDetaper + "m" : "--" },
  ];

  const kpiHTML = kpis.map(k => `
    <div class="kpi">
      <div class="kpi-l">${k.label}</div>
      <div class="kpi-v">${k.value}</div>
    </div>`).join("");

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Coating Flow Report — ${reportDate}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background: #080a0f;
      color: #e8f0ff;
      font-family: 'Segoe UI', 'Inter', sans-serif;
      padding: 32px;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      border-bottom: 2px solid #1e2535;
      padding-bottom: 16px;
      margin-bottom: 24px;
    }
    .header-title { font-size: 22px; font-weight: 700; color: #ffffff; letter-spacing: -0.3px; }
    .header-sub   { font-size: 12px; color: rgba(255,255,255,0.4); margin-top: 4px; letter-spacing: 1px; text-transform: uppercase; }
    .header-date  { font-size: 13px; color: rgba(255,255,255,0.6); text-align: right; }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(6, 1fr);
      gap: 10px;
      margin-bottom: 28px;
    }
    .kpi {
      background: #111520;
      border: 1px solid #1e2535;
      border-radius: 8px;
      padding: 14px 12px;
    }
    .kpi-l { font-size: 9px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: rgba(255,255,255,0.45); margin-bottom: 6px; }
    .kpi-v { font-size: 22px; font-weight: 700; color: #ffffff; }
    .chart-section { margin-bottom: 28px; page-break-inside: avoid; }
    .chart-label {
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 1.5px;
      text-transform: uppercase;
      color: rgba(255,255,255,0.5);
      margin-bottom: 8px;
      padding-left: 2px;
      border-left: 3px solid rgba(255,255,255,0.25);
      padding-left: 10px;
    }
    .footer {
      margin-top: 32px;
      padding-top: 14px;
      border-top: 1px solid #1e2535;
      font-size: 11px;
      color: rgba(255,255,255,0.3);
      display: flex;
      justify-content: space-between;
    }
    @media print {
      body { background: #080a0f !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <div class="header">
    <div>
      <div class="header-title">Coating Flow Tracker</div>
      <div class="header-sub">Production · Optical · Zenni Lab</div>
    </div>
    <div class="header-date">
      Report Date: <strong style="color:#ffffff">${reportDate}</strong><br>
      Generated: ${new Date().toLocaleTimeString()}
    </div>
  </div>

  <div class="kpi-grid">${kpiHTML}</div>

  ${chartSections || "<p style='color:rgba(255,255,255,0.3);font-size:13px;'>Open each chart tab first to capture charts in the export.</p>"}

  <div class="footer">
    <span>Coating Flow Tracker — Auto-generated Report</span>
    <span>${reportDate}</span>
  </div>
</body>
</html>`;

  const win = window.open("", "_blank", "width=1100,height=800");
  if (!win) { alert("Please allow pop-ups to export PDF."); return; }
  win.document.write(html);
  win.document.close();
  win.onload = () => {
    setTimeout(() => {
      win.print();
    }, 500);
  };
}