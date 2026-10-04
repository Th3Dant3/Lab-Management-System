const API_URL = "https://script.google.com/macros/s/AKfycbxGEYOoJviGPzBSIX_Kh5X5ZAOAHFSyw9AO6ID-buDr1A82ERaPsauF3cSvSINr8VvT/exec";

console.log("Coating JS loaded");


/* =====================================================
   COATING PERFORMANCE AUDIT — LOGGING ONLY
   No API URL, cache behavior, refresh cadence, loader timing,
   dashboard calculations, history logic, or rendering logic changes.
===================================================== */
const coatingPerfState = {
  loadSeq: 0,
  previousDaySeq: 0,
  initialLoadId: null,
  initialDataReadyAt: null,
  initialDismissRequestedAt: null,
  initialFadeStartedAt: null,
  initialReadyLogged: false
};

function coatingPerfNow_() {
  return performance.now();
}

function coatingPerfLog_(message, details = {}) {
  console.log(`[CoatingPerf] ${message}`, details);
}

function coatingPerfStart_(label, details = {}) {
  const startedAt = coatingPerfNow_();
  coatingPerfLog_(`▶ ${label}`, details);
  return startedAt;
}

function coatingPerfEnd_(label, startedAt, details = {}) {
  const elapsed = coatingPerfNow_() - startedAt;
  coatingPerfLog_(`✓ ${label}: ${elapsed.toFixed(1)} ms`, details);
  return elapsed;
}

coatingPerfLog_(
  `JS executing at ${coatingPerfNow_().toFixed(1)} ms after navigation start`,
  {}
);

window.addEventListener("load", () => {
  coatingPerfLog_(
    `window.load at ${coatingPerfNow_().toFixed(1)} ms after navigation start`,
    {}
  );
});

/* =====================================================
   GLOBAL VARIABLES
===================================================== */

let trendChart   = null;
let reasonChart  = null;
let machineChart = null;
let flowChart    = null;

let dashboardProcessed = null;
let dashboardMachine   = null;
let coatingMachineFailed_ = false;   // true when the last Machine API call failed
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
  bodyColor       : "#ffffff",            // bright text on black (was 75% teal)
  footerColor     : "#ffffff",
  titleFont       : { family: CHART_FONT, size: 15, weight: "700" },
  bodyFont        : { family: CHART_MONO, size: 14, weight: "600" },
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
  purple : "#b39bff",
  red    : "#ff6b6b",
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
  "S-HC Contamination"       : "#ff6b6b",
  "S-HC Pit"                 : "#b39bff",
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
      we.value = coatingTodayISO_();
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

/* Apps Script sometimes answers a slow request with 404 from script.googleusercontent.com
   (the redirected result is gone by the time the browser asks for it). Retry once, after a
   short pause, only for 404 / 5xx / network errors. One retry max so we never pile up load. */
async function coatingFetchRetry_(url, label) {
  try {
    const r = await fetch(url);
    if (r.status !== 404 && r.status < 500) return r;
    coatingPerfLog_(`${label} got HTTP ${r.status}; retrying once in 1.5 s`, {});
  } catch (e) {
    coatingPerfLog_(`${label} network error; retrying once in 1.5 s`, { error: String(e) });
  }
  await new Promise(res => setTimeout(res, 1500));
  return fetch(url);
}

async function loadDashboard() {
  const loadId = ++coatingPerfState.loadSeq;
  const isInitialLoad = coatingPerfState.initialLoadId === null;
  if (isInitialLoad) coatingPerfState.initialLoadId = loadId;

  const loadStartedAt = coatingPerfStart_(
    `${isInitialLoad ? "INITIAL " : ""}Coating dashboard load #${loadId}`,
    {
      currentDate,
      refreshSeconds: refreshInterval,
      requests: ["processed", "machine"]
    }
  );

  updateStatusBadge("updating");
  const startTime = performance.now();
  const dateParam = currentDate ? `&date=${encodeURIComponent(currentDate)}` : "";

  let processedResponseHeadersMs = 0;
  let machineResponseHeadersMs = 0;
  let processedBodyMs = 0;
  let machineBodyMs = 0;
  let processedParseMs = 0;
  let machineParseMs = 0;
  let renderMs = 0;
  let processedChars = 0;
  let machineChars = 0;

  const processedStartedAt = coatingPerfStart_(
    `Processed API #${loadId}`,
    { mode: "processed", currentDate }
  );
  const machineStartedAt = coatingPerfStart_(
    `Machine API #${loadId}`,
    { mode: "machine", currentDate }
  );

  const processedPromise = (async () => {
    const headersStartedAt = coatingPerfNow_();
    const response = await coatingFetchRetry_(`${API_URL}?mode=processed${dateParam}`, `Processed API #${loadId}`);
    processedResponseHeadersMs = coatingPerfNow_() - headersStartedAt;

    coatingPerfLog_(
      `Processed API #${loadId} response headers: ${processedResponseHeadersMs.toFixed(1)} ms`,
      { httpStatus: response.status, ok: response.ok }
    );

    const bodyStartedAt = coatingPerfNow_();
    const responseText = await response.text();
    processedBodyMs = coatingPerfNow_() - bodyStartedAt;
    processedChars = responseText.length;

    coatingPerfLog_(
      `Processed API #${loadId} body read: ${processedBodyMs.toFixed(1)} ms`,
      { responseChars: processedChars }
    );

    const parseStartedAt = coatingPerfNow_();
    const payload = JSON.parse(responseText);
    processedParseMs = coatingPerfNow_() - parseStartedAt;

    coatingPerfLog_(
      `Processed API #${loadId} JSON parse: ${processedParseMs.toFixed(1)} ms`,
      {
        hasSummary: !!payload?.summary,
        hourlyRows: Array.isArray(payload?.hourly) ? payload.hourly.length : 0
      }
    );

    coatingPerfEnd_(`Processed API #${loadId}`, processedStartedAt, {
      responseHeadersMs: Number(processedResponseHeadersMs.toFixed(1)),
      bodyReadMs: Number(processedBodyMs.toFixed(1)),
      jsonParseMs: Number(processedParseMs.toFixed(1)),
      responseChars: processedChars,
      httpStatus: response.status
    });

    return payload;
  })();

  const machinePromise = (async () => {
    const headersStartedAt = coatingPerfNow_();
    const response = await coatingFetchRetry_(`${API_URL}?mode=machine${dateParam}`, `Machine API #${loadId}`);
    machineResponseHeadersMs = coatingPerfNow_() - headersStartedAt;

    coatingPerfLog_(
      `Machine API #${loadId} response headers: ${machineResponseHeadersMs.toFixed(1)} ms`,
      { httpStatus: response.status, ok: response.ok }
    );

    const bodyStartedAt = coatingPerfNow_();
    const responseText = await response.text();
    machineBodyMs = coatingPerfNow_() - bodyStartedAt;
    machineChars = responseText.length;

    coatingPerfLog_(
      `Machine API #${loadId} body read: ${machineBodyMs.toFixed(1)} ms`,
      { responseChars: machineChars }
    );

    if (!response.ok) throw new Error(`Machine API HTTP ${response.status}`);
    const parseStartedAt = coatingPerfNow_();
    const payload = JSON.parse(responseText);
    machineParseMs = coatingPerfNow_() - parseStartedAt;

    coatingPerfLog_(
      `Machine API #${loadId} JSON parse: ${machineParseMs.toFixed(1)} ms`,
      {
        hasSummary: !!payload?.summary,
        hourlyRows: Array.isArray(payload?.hourly) ? payload.hourly.length : 0
      }
    );

    coatingPerfEnd_(`Machine API #${loadId}`, machineStartedAt, {
      responseHeadersMs: Number(machineResponseHeadersMs.toFixed(1)),
      bodyReadMs: Number(machineBodyMs.toFixed(1)),
      jsonParseMs: Number(machineParseMs.toFixed(1)),
      responseChars: machineChars,
      httpStatus: response.status
    });

    return payload;
  })();

  // Processed is required. Machine is optional: if it fails, the page still renders
  // from processed data and the Machine views say so (was: one failure blanked the whole page).
  const [procRes, machRes] = await Promise.allSettled([processedPromise, machinePromise]);
  if (procRes.status === "rejected") throw procRes.reason;
  dashboardProcessed = procRes.value;
  if (machRes.status === "fulfilled") {
    dashboardMachine = machRes.value;
    coatingMachineFailed_ = false;
  } else {
    dashboardMachine = null;
    coatingMachineFailed_ = true;
    coatingPerfLog_(`Machine API #${loadId} failed; rendering from processed data`, { error: String(machRes.reason) });
  }
  dashboardData = dashboardProcessed;

  const renderStartedAt = coatingPerfStart_(
    `Coating build/render #${loadId}`,
    {}
  );

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
  {
    // Machine Analysis tab: bay + reason matrix + hour heatmap, all from the machine payload
    const machData = (dashboardMachine && dashboardMachine.machineTotals) ? dashboardMachine : dashboardProcessed;
    buildCoaterBay(machData, "coaterBay");
    buildMachinePanels(machData);
    const mNote = document.getElementById("machineSourceNote");
    if (mNote) mNote.textContent = coatingMachineFailed_
      ? "Machine report didn't load on the last refresh. Showing coater numbers from the processed report instead."
      : "";
    // Overview board: summary + hourly from processed, coater totals from machine payload
    buildOverviewFlow(dashboardProcessed, machData);
    syncDateInput_();
  }
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

  renderMs = coatingPerfEnd_(
    `Coating build/render #${loadId}`,
    renderStartedAt,
    {
      processedHourlyRows: Array.isArray(dashboardProcessed?.hourly)
        ? dashboardProcessed.hourly.length
        : 0,
      machineHourlyRows: Array.isArray(dashboardMachine?.hourly)
        ? dashboardMachine.hourly.length
        : 0
    }
  );

  const totalMs = coatingPerfEnd_(
    `${isInitialLoad ? "INITIAL " : ""}Coating dashboard load #${loadId}`,
    loadStartedAt,
    {
      processedHeadersMs: Number(processedResponseHeadersMs.toFixed(1)),
      machineHeadersMs: Number(machineResponseHeadersMs.toFixed(1)),
      processedBodyMs: Number(processedBodyMs.toFixed(1)),
      machineBodyMs: Number(machineBodyMs.toFixed(1)),
      processedParseMs: Number(processedParseMs.toFixed(1)),
      machineParseMs: Number(machineParseMs.toFixed(1)),
      renderMs: Number(renderMs.toFixed(1)),
      success: true
    }
  );

  if (isInitialLoad) {
    coatingPerfState.initialDataReadyAt = coatingPerfNow_();
    coatingPerfLog_(
      `Data/render ready for initial Coating load #${loadId}; loader intentionally waits 250 ms before fade and 720 ms before display:none`,
      {
        apiAndRenderMs: Number(totalMs.toFixed(1)),
        preFadeHoldMs: 250,
        fadeDisplayNoneDelayMs: 720
      }
    );
  }

  console.log("Load time (ms):", Math.round(performance.now() - startTime));
}

function updateFlowStatPills(summary) {
  const pills = {
    "flowStatHealthy"  : { val: summary.flowHealth?.healthy   || 0, color: "#00e676" },
    "flowStatWatch"    : { val: summary.flowHealth?.watch     || 0, color: "#ffd600" },
    "flowStatDelayed"  : { val: summary.flowHealth?.delayed   || 0, color: "#ff6b6b" },
    "flowStatBreakage" : { val: summary.totalBreakLenses      || 0, color: "#c4b5fd" },
  };
  Object.entries(pills).forEach(([id, { val, color }]) => {
    const el = document.getElementById(id);
    if (el) { el.textContent = val; el.style.color = color; }
  });
  // share of all jobs under each count
  const fh = summary.flowHealth || {};
  const tot = (fh.healthy || 0) + (fh.watch || 0) + (fh.delayed || 0) + (fh.overnight || 0);
  const pct = n => tot ? `${(n / tot * 100).toFixed(1)}% of ${tot} jobs` : "";
  const sub = (id, t) => { const e = document.getElementById(id); if (e) e.textContent = t; };
  sub("flowStatHealthySub", pct(fh.healthy || 0));
  sub("flowStatWatchSub",   pct(fh.watch || 0));
  sub("flowStatDelayedSub", pct(fh.delayed || 0) + ((fh.overnight || 0) ? ` \u00b7 ${fh.overnight} overnight` : ""));
  sub("flowStatBreakageSub", "all breakage today");
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
   HEADER CLOCK + DATE INPUT (America/New_York)
===================================================== */
function coatingTodayISO_() {
  // en-CA formats as YYYY-MM-DD, which is what <input type="date"> uses
  return new Date().toLocaleDateString("en-CA", { timeZone: "America/New_York" });
}
function syncDateInput_() {
  const el = document.getElementById("historyDate");
  if (!el) return;
  const today = coatingTodayISO_();
  el.max = today;                                  // no future dates
  if (currentDate === null) { el.value = today; el.classList.add("is-live"); }
  else el.classList.remove("is-live");
}
function tickHeaderClock_() {
  const now = new Date();
  const d = now.toLocaleDateString("en-US", { timeZone: "America/New_York", weekday: "short", month: "short", day: "numeric", year: "numeric" });
  const t = now.toLocaleTimeString("en-US", { timeZone: "America/New_York", hour: "numeric", minute: "2-digit" });
  const dEl = document.getElementById("nowDate"), tEl = document.getElementById("nowTime");
  if (dEl) dEl.textContent = d;
  if (tEl) tEl.innerHTML = `${t}<small>ET</small>`;
}
document.addEventListener("DOMContentLoaded", () => {
  tickHeaderClock_();
  setInterval(tickHeaderClock_, 15000);
  syncDateInput_();
  const el = document.getElementById("historyDate");
  if (el) el.addEventListener("click", () => { try { el.showPicker && el.showPicker(); } catch (e) {} });
});

/* =====================================================
   DATE FILTER
===================================================== */

function applyDateFilter() {
  const dateInput = document.getElementById("historyDate").value;
  if (!dateInput) return;
  // Picking today's date means live mode, not a frozen history snapshot of today
  if (dateInput === coatingTodayISO_()) { if (currentDate !== null) resetToToday(); return; }
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
  const coatingBootStartedAt = coatingPerfStart_(
    "DOMContentLoaded → Coating boot",
    {}
  );
  coatingPerfLog_(
    `DOMContentLoaded fired at ${coatingPerfNow_().toFixed(1)} ms after navigation start`,
    {}
  );

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

  coatingPerfEnd_(
    "DOMContentLoaded → Coating boot",
    coatingBootStartedAt,
    {
      refreshSeconds: refreshInterval,
      loaderStepIntervalMs: 420
    }
  );
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

  coatingPerfState.initialDismissRequestedAt = coatingPerfNow_();
  coatingPerfLog_(
    "Initial loader dismissal requested; intentional 250 ms hold begins",
    {}
  );

  setTimeout(() => {
    coatingPerfState.initialFadeStartedAt = coatingPerfNow_();

    coatingPerfLog_(
      "Initial loader fade-out started",
      {
        holdMs: coatingPerfState.initialDismissRequestedAt == null
          ? null
          : Number(
              (coatingPerfState.initialFadeStartedAt -
               coatingPerfState.initialDismissRequestedAt).toFixed(1)
            )
      }
    );

    if (loadScreen) {
      loadScreen.classList.add("fade-out");

      setTimeout(() => {
        loadScreen.style.display = "none";

        if (!coatingPerfState.initialReadyLogged) {
          coatingPerfState.initialReadyLogged = true;
          const readyAt = coatingPerfNow_();

          coatingPerfLog_("✅ INITIAL COATING DASHBOARD READY", {
            timestamp: new Date().toISOString(),
            fetchId: coatingPerfState.initialLoadId,
            dataRenderReadyMs:
              coatingPerfState.initialDataReadyAt == null
                ? null
                : coatingPerfState.initialDataReadyAt.toFixed(1),
            preFadeHoldMs: 250,
            fadeDisplayNoneDelayMs: 720,
            totalReadyMs: readyAt.toFixed(1)
          });

          console.log(
            `[CoatingPerf] MAIN COATING DASHBOARD TIME TO READY: ${readyAt.toFixed(1)} ms (${(readyAt / 1000).toFixed(2)} sec)`
          );
        }
      }, 720);
    }
  }, 250);
};

loadDashboard().then(() => {
  if (typeof window._dismissLoadingScreen === "function") window._dismissLoadingScreen();

  /*
    COATING STAGE 1:
    Previous-day data is useful for comparison, but it must not compete
    with the two initial live dashboard requests. Start it only after
    processed + machine have both completed.
  */
  coatingPerfLog_(
    "Main live dashboard complete; starting deferred previous-day request",
    {}
  );
  // loadPreviousDay();  // disabled: its data was never used and the request took ~24 s
}).catch(() => {
  if (typeof window._dismissLoadingScreen === "function") window._dismissLoadingScreen();

  // Preserve previous-day availability even if the main load has a problem.
  coatingPerfLog_(
    "Main live dashboard failed; starting deferred previous-day request after main request settled",
    {}
  );
  // loadPreviousDay();  // disabled: its data was never used and the request took ~24 s
});

/* =====================================================
   TREND CHART SYSTEM
===================================================== */

function buildTrendCharts() {
  trendsLoaded = true;
  buildTrendChart(currentMode === "machine" ? dashboardMachine : dashboardProcessed);
  setTimeout(() => buildFlowChart(dashboardProcessed), 100);
}

function switchTrend(mode) {
  currentMode = mode;
  if (mode === "processed") buildTrendChart(dashboardProcessed);
  else if (mode === "machine") buildTrendChart(dashboardMachine);
  // only the two trend buttons (was clearing every .mode-btn on the page, incl. the flow-chart buttons)
  document.querySelectorAll("#processedBtn, #machineBtn").forEach(btn => btn.classList.remove("active"));
  const btn = document.getElementById(mode + "Btn");
  if (btn) btn.classList.add("active");
}

/* =====================================================
   TREND CHART — GLASSMORPHISM SMOOTH CURVES
===================================================== */

/* =====================================================
   BREAKAGE TREND — bars = breakage, line = workflow (jobs coated)
   One chart, two scales whose grid lines line up (1 lens = N jobs).
   - Bars stacked by AGE (same day / previous day / 2+ / unknown) or by REASON
   - All bars are breakage found TODAY; age = the day of the machine scan (today / previous day / 2+ days)
   - Line: straight segments; the hour still in progress is dashed
   - Flow row (W / D per hour) and a one-line summary under the chart
   Keeps canvas id "trendChart" so the PDF export still captures it.
===================================================== */
let trendColorBy = "age";

const TREND_AGE_SERIES = [
  { key: "s", label: "Machine scan today",        color: "#ff6b6b" },
  { key: "o", label: "Machine scan previous day", color: "#fbbf24" },
  { key: "p", label: "Machine scan 2+ days ago",  color: "#f472b6" },
  { key: "u", label: "Age unknown",  color: "#e5e7eb" },
];

function setTrendColorBy(mode) {
  trendColorBy = mode;
  document.querySelectorAll(".trend-colorby-btn").forEach(b => b.classList.toggle("active", b.dataset.by === mode));
  buildTrendChart(currentMode === "machine" ? dashboardMachine : dashboardProcessed);
}

const trendPatternCache_ = {};
function trendPattern_(color, kind) {
  // kind: "s" solid, "o" diagonal stripes, "p" dots
  if (kind === "s") return color;
  const key = color + kind;
  if (trendPatternCache_[key]) return trendPatternCache_[key];
  const c = document.createElement("canvas"); c.width = c.height = 8;
  const x = c.getContext("2d");
  x.fillStyle = color; x.globalAlpha = 0.35; x.fillRect(0, 0, 8, 8); x.globalAlpha = 1;
  if (kind === "o") {
    x.strokeStyle = color; x.lineWidth = 3;
    x.beginPath(); x.moveTo(-2, 10); x.lineTo(10, -2); x.moveTo(-6, 6); x.lineTo(6, -6); x.moveTo(2, 14); x.lineTo(14, 2); x.stroke();
  } else {
    x.fillStyle = color; x.beginPath(); x.arc(4, 4, 2.2, 0, Math.PI * 2); x.fill();
  }
  const pat = x.createPattern(c, "repeat") || color;
  trendPatternCache_[key] = pat;
  return pat;
}

function trendHourKey_(label) {
  const h = ovHour24_(label);
  return h === null ? null : Math.floor(h);
}

function trendRowBreak_(row) {
  const b = { t: 0, s: 0, o: 0, p: 0, reasons: {}, ra: {} };
  let mSum = 0;
  Object.values(row.machines || {}).forEach(m => {
    mSum  += Number(m.total || 0);
    b.s   += Number(m.sameDay || 0);
    b.o   += Number(m.oneDay  || 0);
    b.p   += Number(m.twoPlus || 0);
    Object.entries(m.reasons || {}).forEach(([r, rs]) => {
      b.reasons[r] = (b.reasons[r] || 0) + Number(rs.total || 0);
      const ra = b.ra[r] || (b.ra[r] = { t: 0, s: 0, o: 0, p: 0 });
      ra.t += Number(rs.total || 0); ra.s += Number(rs.sameDay || 0); ra.o += Number(rs.oneDay || 0); ra.p += Number(rs.twoPlus || 0);
    });
  });
  b.t = (row.totalBroken !== undefined && row.totalBroken !== null) ? Number(row.totalBroken) : mSum;
  return b;
}

function trendRowFlow_(row) {
  const pts = row.flowPoints || [];
  return {
    w: pts.filter(p => p.flow > 15 && p.flow <= 30).length,
    d: pts.filter(p => p.flow > 30 && p.flow <= 360).length,
    n: pts.length,
  };
}

function trendNiceStep_(raw) {
  const nice = [5, 10, 15, 20, 25, 30, 40, 50, 60, 75, 100, 150, 200, 250, 500];
  return nice.find(n => n >= raw) || Math.ceil(raw / 100) * 100;
}

function buildTrendChart(data) {
  const canvas = document.getElementById("trendChart");
  if (currentMode === "machine" && !data && coatingMachineFailed_) {
    const hint = document.getElementById("trendModeHint");
    if (hint) hint.innerHTML = `<span class="trend-warn">Machine Scan data didn't load on the last refresh. Switch to Breakage Processed, or wait for the next refresh.</span>`;
    if (trendChart) { trendChart.destroy(); trendChart = null; }
    ["trendLegend", "trendFlowRow", "trendSummary"].forEach(id => { const e = document.getElementById(id); if (e) { e.innerHTML = ""; e._last = ""; } });
    return;
  }
  if (!canvas || !data || !data.hourly) return;
  const ctx = canvas.getContext("2d");
  if (trendChart) { trendChart.destroy(); trendChart = null; }

  const isMachine = currentMode === "machine";
  const isLive    = currentDate === null;
  const esc       = coaterBayEsc_;

  // ── Hour map for this mode ──
  const rowsByH = {};
  data.hourly.forEach(r => { const k = trendHourKey_(r.hour); if (k !== null) rowsByH[k] = r; });

  // Jobs + flow always come from the processed report (coating jobs per hour)
  const procByH = {};
  ((dashboardProcessed && dashboardProcessed.hourly) || []).forEach(r => { const k = trendHourKey_(r.hour); if (k !== null) procByH[k] = r; });


  const hourSet = new Set([...Object.keys(rowsByH), ...Object.keys(procByH)].map(Number));
  const hoursN = [...hourSet].sort((a, b) => a - b);
  const labels = hoursN.map(h => ovHourName_(h));

  const nowNY = ovNowNY_();
  const nowHr = nowNY.getHours();
  const brk   = hoursN.map(h => rowsByH[h] ? trendRowBreak_(rowsByH[h]) : { t: 0, s: 0, o: 0, p: 0, reasons: {}, ra: {} });
  const jobs  = hoursN.map(h => {
    if (isLive && h > nowHr) return null;              // future hours: no line
    return procByH[h] ? Number(procByH[h].coatingJobs || 0) : (rowsByH[h] || !isLive ? 0 : null);
  });
  const partialIdx = isLive ? hoursN.indexOf(nowHr) : -1;

  // ── Bar datasets ──
  const datasets = [];
  if (trendColorBy === "age") {
    TREND_AGE_SERIES.forEach(sr => {
      const vals = brk.map(b => sr.key === "u" ? Math.max(0, b.t - (b.s + b.o + b.p)) : b[sr.key]);
      if (!vals.some(v => v > 0)) return;
      datasets.push({ label: sr.label, type: "bar", data: vals, backgroundColor: sr.color, borderColor: sr.color,
        borderWidth: 0, borderRadius: 3, stack: "brk", yAxisID: "yBroken", order: 2,
        barPercentage: 0.62, categoryPercentage: 0.8 });
    });
  } else if (trendColorBy === "both") {
    // Color = reason, fill = age (solid same day, stripes previous day, dots 2+, outline = age unknown)
    const rTot = {};
    brk.forEach(b => Object.entries(b.reasons).forEach(([r, n]) => { rTot[r] = (rTot[r] || 0) + n; }));
    const fallback = [GC.red, GC.purple, GC.orange, GC.teal, GC.blue, GC.green, GC.pink];
    const reasonsSorted = Object.keys(rTot).filter(r => rTot[r] > 0).sort((a, b) => rTot[b] - rTot[a]);
    reasonsSorted.forEach((r, ri) => {
      const col = BREAKAGE_COLOR_MAP[r] || fallback[ri % fallback.length];
      TREND_AGE_SERIES.forEach(sr => {
        const vals = brk.map(b => {
          const ra = b.ra[r]; if (!ra) return 0;
          return sr.key === "u" ? Math.max(0, ra.t - (ra.s + ra.o + ra.p)) : ra[sr.key];
        });
        if (!vals.some(v => v > 0)) return;
        const isUnk = sr.key === "u";
        datasets.push({ label: `${r} \u00b7 ${sr.label}`, type: "bar", data: vals,
          backgroundColor: isUnk ? "rgba(0,0,0,0)" : trendPattern_(col, sr.key),
          borderColor: isUnk ? col : "#061210", borderWidth: isUnk ? 2 : 1, borderRadius: 3,
          _reason: r, _reasonColor: col, _age: sr.key,
          stack: "brk", yAxisID: "yBroken", order: 2, barPercentage: 0.62, categoryPercentage: 0.8 });
      });
    });
    const unr = brk.map(b => Math.max(0, b.t - Object.values(b.reasons).reduce((s, n) => s + n, 0)));
    if (unr.some(v => v > 0)) datasets.push({ label: "No reason", type: "bar", data: unr, backgroundColor: "#e5e7eb",
      borderWidth: 0, borderRadius: 3, stack: "brk", yAxisID: "yBroken", order: 2, barPercentage: 0.62, categoryPercentage: 0.8 });
  } else {
    const rTot = {};
    brk.forEach(b => Object.entries(b.reasons).forEach(([r, n]) => { rTot[r] = (rTot[r] || 0) + n; }));
    const fallback = [GC.red, GC.purple, GC.orange, GC.teal, GC.blue, GC.green, GC.pink];
    Object.keys(rTot).filter(r => rTot[r] > 0).sort((a, b) => rTot[b] - rTot[a]).forEach((r, i) => {
      const col = BREAKAGE_COLOR_MAP[r] || fallback[i % fallback.length];
      datasets.push({ label: r, type: "bar", data: brk.map(b => b.reasons[r] || 0), backgroundColor: col, borderColor: col,
        borderWidth: 0, borderRadius: 3, stack: "brk", yAxisID: "yBroken", order: 2,
        barPercentage: 0.62, categoryPercentage: 0.8 });
    });
    // breakage with no reason recorded still has to show
    const unr = brk.map(b => Math.max(0, b.t - Object.values(b.reasons).reduce((s, n) => s + n, 0)));
    if (unr.some(v => v > 0)) datasets.push({ label: "No reason", type: "bar", data: unr, backgroundColor: "#e5e7eb",
      borderWidth: 0, borderRadius: 3, stack: "brk", yAxisID: "yBroken", order: 2, barPercentage: 0.62, categoryPercentage: 0.8 });
  }

  // ── Line: workflow ──
  datasets.push({
    label: "Jobs coated", type: "line", data: jobs, yAxisID: "yJobs", order: 0,
    borderColor: "#3ddbc4", backgroundColor: "rgba(61,219,196,0.10)", fill: "origin",
    borderWidth: 3.5, tension: 0, spanGaps: false,
    pointRadius: jobs.map((v, i) => v === null ? 0 : 6), pointHoverRadius: 8,
    pointBackgroundColor: jobs.map((v, i) => i === partialIdx ? "#061210" : "#3ddbc4"),
    pointBorderColor: "#3ddbc4", pointBorderWidth: 2.5,
    segment: { borderDash: c => (partialIdx >= 0 && c.p1DataIndex === partialIdx) ? [7, 6] : undefined },
  });

  // ── Scales whose grid lines line up ──
  const stackMax = Math.max(0, ...brk.map(b => b.t));
  const bMax  = Math.max(5, Math.ceil(stackMax * 1.15));
  const jMaxV = Math.max(1, ...jobs.map(v => v || 0));
  const jStep = trendNiceStep_((jMaxV * 1.12) / bMax);

  // ── Plugin: line labels ──
  const lineLabelPlugin = {
    id: "trendLineLabels",
    afterDatasetsDraw(chart) {
      const li = chart.data.datasets.length - 1;
      const lm = chart.getDatasetMeta(li);
      if (!lm || lm.hidden) return;
      const c = chart.ctx, yB = chart.scales.yBroken;
      c.save();
      c.font = `700 14px ${CHART_FONT}`; c.textAlign = "center"; c.textBaseline = "middle";
      lm.data.forEach((pt, i) => {
        const v = jobs[i];
        if (v === null || v === undefined) return;
        const barTop = yB.getPixelForValue(brk[i].t), barBot = yB.getPixelForValue(0);
        const onBar  = brk[i].t > 0 && pt.y > barTop - 30 && pt.y < barBot + 4;
        const lx = onBar ? pt.x + 36 : pt.x, ly = onBar ? pt.y : pt.y - 20;
        const txt = String(v), w = c.measureText(txt).width + 12;
        c.fillStyle = "rgba(6,18,16,0.88)"; c.fillRect(lx - w / 2, ly - 10, w, 20);
        c.fillStyle = "#3ddbc4"; c.fillText(txt, lx, ly);
      });
      c.restore();
    },
    afterRender(chart) { trendPlaceFlowRow_(chart, hoursN); },
  };

  trendChart = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      animation: { duration: 450, easing: "easeOutQuart" },
      layout: { padding: { top: 26 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...GLASS_TOOLTIP,
          // Hide empty bar segments, but always keep the jobs line so the tooltip is never empty
          filter: item => item.raw !== null && item.raw !== undefined && !(item.dataset.type === "bar" && item.raw === 0),
          callbacks: {
            title: items => items.length ? `  ${items[0].label}` : "",
            label: c => `  ${c.dataset.label}: ${c.raw}`,
            afterBody: items => {
              if (!items.length) return [];
              const i = items[0].dataIndex, out = [];
              if (i === partialIdx) out.push("  Hour still in progress");
              return out;
            },
          },
        },
      },
      scales: {
        x: { stacked: true, grid: { display: false }, border: { display: false },
             ticks: { color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "600" } } },
        yBroken: {
          position: "left", stacked: true, min: 0, max: bMax,
          ticks: { stepSize: 1, color: "#ff8a8a", font: { family: CHART_FONT, size: 13, weight: "700" } },
          grid: { color: "rgba(255,255,255,0.07)" }, border: { display: false },
          title: { display: true, text: "Broken lenses", color: "#ff8a8a", font: { family: CHART_FONT, size: 14, weight: "700" } },
        },
        yJobs: {
          position: "right", stacked: false, min: 0, max: bMax * jStep,
          ticks: { stepSize: jStep, color: "#3ddbc4", font: { family: CHART_FONT, size: 13, weight: "700" } },
          grid: { drawOnChartArea: false }, border: { display: false },
          title: { display: true, text: "Jobs coated", color: "#3ddbc4", font: { family: CHART_FONT, size: 14, weight: "700" } },
        },
      },
    },
    plugins: [lineLabelPlugin],
  });

  // ── Legend ──
  const leg = document.getElementById("trendLegend");
  if (leg) {
    const items = [`<span class="tl-key" style="--c:#3ddbc4">Jobs coated</span>`];
    if (trendColorBy === "both") {
      const seenR = {}, seenA = {};
      datasets.filter(d => d._reason).forEach(d => { seenR[d._reason] = d._reasonColor; seenA[d._age] = true; });
      items.push(`<b class="tl-group">Color = reason:</b>`);
      Object.entries(seenR).forEach(([r, c]) => items.push(`<span class="tl-key" style="--c:${c}">${esc(r)}</span>`));
      items.push(`<b class="tl-group">Fill = machine scan day:</b>`);
      if (seenA.s) items.push(`<span class="tl-key tl-age-s">Solid = today</span>`);
      if (seenA.o) items.push(`<span class="tl-key tl-age-o">Striped = previous day</span>`);
      if (seenA.p) items.push(`<span class="tl-key tl-age-p">Dotted = 2+ days ago</span>`);
      if (seenA.u) items.push(`<span class="tl-key tl-age-u">Outline = age unknown</span>`);
      datasets.filter(d => d.type === "bar" && !d._reason).forEach(d => items.push(`<span class="tl-key" style="--c:${d.backgroundColor}">${esc(d.label)}</span>`));
    } else {
      datasets.filter(d => d.type === "bar").forEach(d => items.push(`<span class="tl-key" style="--c:${d.backgroundColor}">${esc(d.label)}</span>`));
    }
    items.push(`<span class="tl-note">W = on watch \u00b7 D = delayed</span>`);
    leg.innerHTML = items.join("");
  }

  // ── Mode hint ──
  const hint = document.getElementById("trendModeHint");
  if (hint) hint.textContent = isMachine
    ? "Machine Scan: today's breakage placed at the hour of the machine scan. Lenses machine-scanned yesterday that broke today show at yesterday's scan hour."
    : "Breakage Processed: today's breakage by the hour it was processed.";

  // ── Flow row data (placed after render) ──
  trendFlowCells_ = hoursN.map((h, i) => {
    if (jobs[i] === null) return null;
    const r = procByH[h];
    if (!r || !Number(r.coatingJobs || 0)) return { none: true };
    return trendRowFlow_(r);
  });

  // ── Summary ──
  const sum = document.getElementById("trendSummary");
  if (sum) {
    const T = brk.reduce((a, b) => ({ t: a.t + b.t, s: a.s + b.s, o: a.o + b.o, p: a.p + b.p }), { t: 0, s: 0, o: 0, p: 0 });
    let peak = -1; brk.forEach((b, i) => { if (b.t > 0 && (peak < 0 || b.t > brk[peak].t)) peak = i; });
    const parts = [
      `<span><b>${T.t}</b>broken ${isLive ? "so far" : "this day"}</span>`,
      `<span><b style="color:#ff6b6b">${T.s}</b>machine scan today</span>`,
      `<span><b style="color:#fbbf24">${T.o}</b>machine scan previous day</span>`,
    ];
    if (T.p) parts.push(`<span><b style="color:#f472b6">${T.p}</b>machine scan 2+ days ago</span>`);
    if (peak >= 0) parts.push(`<span><b style="color:#fbbf24">${labels[peak]}</b>peak hour</span>`);
    sum.innerHTML = parts.join("");
  }
}

let trendFlowCells_ = [];
function trendPlaceFlowRow_(chart, hoursN) {
  const row = document.getElementById("trendFlowRow");
  if (!row || !chart.scales || !chart.scales.x) return;
  const xS = chart.scales.x;
  const cells = trendFlowCells_.map((c, i) => {
    if (!c) return "";
    const x = xS.getPixelForValue(i);
    let inner;
    if (c.none) inner = `<span class="tf-none">\u2014</span>`;
    else if (!c.w && !c.d) inner = `<span class="tf-ok">all healthy</span>`;
    else inner = (c.w ? `<span class="tf-chip tf-w">W ${c.w}</span>` : "") + (c.d ? `<span class="tf-chip tf-d">D ${c.d}</span>` : "");
    return `<div class="tf-cell" style="left:${x.toFixed(0)}px">${inner}</div>`;
  }).join("");
  const html = `<span class="tf-label" style="left:${(xS.left - 12).toFixed(0)}px">Flow</span>` + cells;
  if (row._last !== html) { row.innerHTML = html; row._last = html; }
}


/* =====================================================
   FLOW CHART — STACKED BARS + BUBBLE BREAKAGE EVENTS
===================================================== */

function switchFlowMode(mode) {
  currentFlowMode = mode;
  document.querySelectorAll("#avgFlowBtn, #machineFlowBtn")
    .forEach(btn => btn.classList.remove("active"));
  const btnMap = { average: "avgFlowBtn", machine: "machineFlowBtn" };
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
  // Overall-only pieces (legend, jobs row) hide in By Machine / Individual RX; their Chart.js legend hint shows instead
  if (currentFlowMode !== "average" && currentFlowMode !== "machine") currentFlowMode = "average";   // Individual RX removed
  document.querySelectorAll(".flow-other-only").forEach(e => e.style.display = "none");
  if (currentFlowMode !== "average") buildFlowCompareCards_(data);

  // Cap at 360m (6h) — anything over is "Overnight" and excluded from avg calculations
  // This prevents single outlier jobs (e.g. 857m) from blowing the broken avg line off the chart
  function sanitize(v) {
    if (v === null || v === undefined) return null;
    const n = Number(v);
    return n > 360 ? null : n;
  }

  // ── BY MACHINE — OUTPUT: jobs coated per coater each hour ──
  //   Same style as Overall (straight lines, flow particles), thicker lines.
  //   0 is a REAL value here: the coater coated no jobs that hour (down, starved, or not started).
  if (currentFlowMode === "machine") {
    const isLive = currentDate === null;
    const nowHr  = ovNowNY_().getHours();
    const partialIdx = isLive ? filtered.map(h => { const k = ovHour24_(h.hour); return k === null ? null : Math.floor(k); }).indexOf(nowHr) : -1;
    const reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

    const machineSet = new Set();
    filtered.forEach(h => Object.entries(h.machines || {}).forEach(([m, s]) => { if (Number(s.jobs || 0) > 0) machineSet.add(m); }));
    const machines = [...machineSet].sort((x, y) => formatMachineLabel(x).localeCompare(formatMachineLabel(y), undefined, { numeric: true }));
    // Coater colors avoid green / yellow / red, which already mean Healthy / Watch / Delayed
    const palette = [
      { c: "#38bdf8", t: "#c9ecfd" }, { c: "#a78bfa", t: "#e4dcfe" }, { c: "#f472b6", t: "#fcd3e8" },
      { c: "#e5e7eb", t: "#ffffff" }, { c: "#2dd4bf", t: "#c4f5ec" }, { c: "#fb923c", t: "#fed7b5" },
    ];

    const datasets = machines.map((m, idx) => {
      const col = palette[idx % palette.length];
      const jobsArr = filtered.map(h => Number((h.machines && h.machines[m] && h.machines[m].jobs) || 0));
      const flowArr = filtered.map(h => {
        const s = h.machines && h.machines[m];
        return (s && Number(s.flowAllCount || 0) > 0) ? sanitize(s.avgFlowAll) : null;
      });
      return {
        label: formatMachineLabel(m), data: jobsArr, _real: jobsArr, _flow: flowArr, _tint: col.t, _kind: "machine",
        _total: jobsArr.reduce((s, n) => s + n, 0),
        borderColor: col.c, backgroundColor: col.c, borderWidth: 5, tension: 0, spanGaps: false, fill: false,
        pointRadius: 6, pointHoverRadius: 9,
        pointBackgroundColor: jobsArr.map((v, i) => i === partialIdx ? "#061210" : col.c),
        pointBorderColor: col.c, pointBorderWidth: 3,
        segment: { borderDash: () => [] },
        hidden: flowHidden_.has("m:" + m),
      };
    });

    const maxV = Math.max(1, ...datasets.flatMap(d => d._real));
    const step = maxV <= 25 ? 5 : 10;
    const yMax = Math.ceil((maxV * 1.2) / step) * step;
    const firstRender = !window._flowMachineEntranceDone;
    window._flowMachineEntranceDone = true;

    flowChart = new Chart(ctx, {
      type: "line",
      data: { labels: hours.map(h => ovHourName_(Math.floor(ovHour24_(h) ?? 0))), datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        animation: firstRender && !reduceMotion
          ? { duration: 600, easing: "easeOutQuart", delay: c => (c.type === "data" && c.mode === "default") ? c.dataIndex * 90 : 0 }
          : false,
        layout: { padding: { top: 16 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...GLASS_TOOLTIP,
            titleColor: "#3ddbc4", bodyColor: "#ffffff", footerColor: "#ffffff",
            titleFont: { family: CHART_FONT, size: 15, weight: "700" },
            bodyFont: { family: CHART_FONT, size: 14, weight: "600" },
            itemSort: (a2, b2) => b2.raw - a2.raw,
            callbacks: {
              title: items => items.length ? `  ${items[0].label}` : "",
              label: c => {
                const f = c.dataset._flow[c.dataIndex];
                return `  ${c.dataset.label}: ${c.raw} ${c.raw === 1 ? "job" : "jobs"}${f !== null ? `  \u00b7  flow ${f.toFixed(1)} min` : ""}`;
              },
              afterBody: items => (items.length && items[0].dataIndex === partialIdx) ? ["  Hour still in progress"] : [],
            },
          },
        },
        scales: {
          x: { grid: { display: false }, border: { display: false },
               ticks: { color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "600" } } },
          y: { min: 0, max: yMax, border: { display: false },
               ticks: { stepSize: step, color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "700" } },
               grid: { color: "rgba(255,255,255,0.07)" },
               title: { display: true, text: "Jobs coated per hour", color: "#ffffff", font: { family: CHART_FONT, size: 14, weight: "700" } } },
        },
      },
      plugins: [flowDotPlugin_(reduceMotion), { id: "flowJobsRowM", afterRender(ch) { flowPlaceJobsRow_(ch, filtered); } }],
    });

    const leg = document.getElementById("flowLegend");
    if (leg) {
      leg.innerHTML = datasets.map((d, i) => {
        const m = machines[i], key = "m:" + m, off = flowHidden_.has(key);
        return `<button type="button" class="tl-key tl-toggle${off ? " is-off" : ""}" style="--c:${d.borderColor}"
          data-key="${coaterBayEsc_(key)}" data-idx="${i}" aria-pressed="${!off}" title="Click to ${off ? "show" : "hide"}">${coaterBayEsc_(d.label)} \u00b7 ${d._total} jobs</button>`;
      }).join("") + `<span class="tl-note">Click a coater to hide or show it \u00b7 0 = that coater coated no jobs that hour</span>`;
    }
    flowStartDotLoop_(reduceMotion);
    return;
  }

  // ── OVERALL MODE — five lines: avg flow per group + broken jobs (no bars) ──
  //   Each point = average flow time of the jobs in that group that hour.
  //   Lines break where an hour had no jobs in that group (spanGaps: false, no fake zeros).
  //   Labels = how many jobs the point is based on. Values above the axis are pinned to
  //   the top with their real value in the label.
  const isLive = currentDate === null;
  const nowHr  = ovNowNY_().getHours();
  const hKeys  = filtered.map(h => { const k = ovHour24_(h.hour); return k === null ? null : Math.floor(k); });
  const partialIdx = isLive ? hKeys.indexOf(nowHr) : -1;

  const groups = [
    { key: "healthy",   label: "Healthy \u00b7 15 min or less", color: "#4ade80", tint: "#c8f7da", test: f => f > 0  && f <= 15,  countKey: "flowHealthy"   },
    { key: "watch",     label: "Watch \u00b7 16 to 30 min",     color: "#fbbf24", tint: "#fde9a8", test: f => f > 15 && f <= 30,  countKey: "flowWatch"     },
    { key: "delayed",   label: "Delayed \u00b7 31 min to 6 h",  color: "#ff6b6b", tint: "#ffc2c2", test: f => f > 30 && f <= 360, countKey: "flowDelayed"   },
    { key: "overnight", label: "Overnight \u00b7 over 6 h",     color: "#60a5fa", tint: "#c7e0ff", test: f => f > 360,             countKey: "flowOvernight" },
  ];

  const series = groups.map(g => {
    const real = [], counts = [];
    filtered.forEach(h => {
      const pts = (h.flowPoints || []).filter(p => g.test(Number(p.flow)));
      real.push(pts.length ? pts.reduce((s, p) => s + Number(p.flow), 0) / pts.length : null);
      counts.push(Number(h[g.countKey] || pts.length || 0));
    });
    return { ...g, real, counts };
  });
  const brokenReal   = filtered.map(h => (Number(h.flowBrokenCount || 0) > 0) ? Number(h.avgFlowBroken || 0) : null);
  const brokenCounts = filtered.map(h => Number(h.flowBrokenCount || 0));

  // Axis: 50 min, or 60 if something sits between 45 and 60. Anything higher is pinned to the top.
  const allVals = series.filter(s => s.key !== "overnight").flatMap(s => s.real).concat(brokenReal).filter(v => v !== null);
  const yMax = allVals.some(v => v > 45) ? 60 : 50;
  const pin  = v => v === null ? null : Math.min(v, yMax);

  const fmtMin = v => v >= 60 ? `${(v / 60).toFixed(1)} h` : `${Math.round(v)} min`;

  // Hours with ZERO jobs in a group run along a thin, faint lane at the bottom so every
  // line keeps flowing all day. Each group has its own lane height so they don't overlap.
  // The lane is drawn thinner and dimmer on purpose: it means "no jobs", not "0 minutes".
  const LANE = { healthy: 0.4, watch: 0.9, delayed: 1.4, overnight: 1.9, broken: 2.4 };
  const laneData = (real, key) => real.map(v => v === null ? LANE[key] : pin(v));
  const segStyle = (color, dashReal) => ({
    borderColor: c => (c.p0.skip || c.p1.skip) ? undefined
      : ((isLaneIdx(c, 0) || isLaneIdx(c, 1)) ? color + "73" : color),
    borderWidth: c => (isLaneIdx(c, 0) || isLaneIdx(c, 1)) ? 2.5 : 4.5,
    borderDash: c => {
      if (isLaneIdx(c, 0) || isLaneIdx(c, 1)) return [];
      return [];               // solid lines everywhere; the moving flow dots show direction
    },
  });
  // Chart.js segment context gives datasetIndex + p0/p1 data indexes
  function isLaneIdx(c, end) {
    const ds = c.chart && c.chart.data.datasets[c.datasetIndex];
    if (!ds || !ds._real) return false;
    const v = ds._real[end === 0 ? c.p0DataIndex : c.p1DataIndex];
    return v === null || v === undefined;
  }

  const datasets = series.map(s => ({
    label: s.label, data: laneData(s.real, s.key), _real: s.real, _counts: s.counts, _kind: s.key, _tint: s.tint,
    borderColor: s.color, backgroundColor: s.color, borderWidth: 4.5, tension: 0, spanGaps: false, fill: false,
    pointRadius: s.real.map(v => v === null ? 0 : 6.5), pointHoverRadius: s.real.map(v => v === null ? 0 : 9),
    pointBackgroundColor: s.real.map((v, i) => i === partialIdx ? "#061210" : s.color),
    pointBorderColor: s.color, pointBorderWidth: 2.5,
    segment: segStyle(s.color, undefined),
    hidden: flowHidden_.has(s.key),
  }));
  datasets.push({
    label: "Broken jobs", data: laneData(brokenReal, "broken"), _real: brokenReal, _counts: brokenCounts, _kind: "broken", _tint: "#ece6ff",
    borderColor: "#c4b5fd", backgroundColor: "#c4b5fd", borderWidth: 4.5, tension: 0, spanGaps: false, fill: false,
    pointRadius: brokenReal.map(v => v === null ? 0 : 6.5), pointHoverRadius: brokenReal.map(v => v === null ? 0 : 9),
    pointBackgroundColor: brokenReal.map((v, i) => i === partialIdx ? "#061210" : "#c4b5fd"),
    pointBorderColor: "#c4b5fd", pointBorderWidth: 2.5,
    segment: segStyle("#c4b5fd", []),
    hidden: flowHidden_.has("broken"),
  });

  const bandPlugin = flowBandPlugin_(yMax);
  const labelPlugin = {
    id: "flowPointLabels",
    afterDatasetsDraw(chart) {
      const c = chart.ctx;
      c.save(); c.font = `700 12px ${CHART_FONT}`; c.textAlign = "center"; c.textBaseline = "middle";
      chart.data.datasets.forEach((ds, di) => {
        if (ds._kind === "healthy") return;                     // healthy has most jobs; label would clutter
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        meta.data.forEach((pt, i) => {
          const real = ds._real[i];
          if (real === null || real === undefined) return;
          const n = ds._counts[i];
          let txt = ds._kind === "broken" ? `${n} broken` : `${n} ${n === 1 ? "job" : "jobs"}`;
          if (real > yMax) txt += ` \u00b7 ${fmtMin(real)}`;
          const dy = ds._kind === "broken" ? 18 : -15;
          c.fillStyle = ds.borderColor; c.fillText(txt, pt.x, pt.y + dy);
        });
      });
      c.restore();
    },
    afterRender(chart) { flowPlaceJobsRow_(chart, filtered); },
  };

  // Flow dots: light dots in each line's own tint travel left to right along every segment
  const reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const flowDotPlugin = flowDotPlugin_(reduceMotion);
  const firstRender = !window._flowEntranceDone;
  window._flowEntranceDone = true;

  flowChart = new Chart(ctx, {
    type: "line",
    data: { labels: hours.map(h => ovHourName_(Math.floor(ovHour24_(h) ?? 0))), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      // Entrance only on the first load: points rise hour by hour, left to right. Refreshes don't replay it.
      animation: firstRender && !reduceMotion
        ? { duration: 600, easing: "easeOutQuart", delay: c => (c.type === "data" && c.mode === "default") ? c.dataIndex * 90 : 0 }
        : false,
      layout: { padding: { top: 22, right: 34 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...GLASS_TOOLTIP,
          filter: item => item.dataset._real[item.dataIndex] !== null && item.dataset._real[item.dataIndex] !== undefined,
          callbacks: {
            title: items => items.length ? `  ${items[0].label}` : "",
            label: c => {
              const real = c.dataset._real[c.dataIndex], n = c.dataset._counts[c.dataIndex];
              return `  ${c.dataset.label}: ${fmtMin(real)} avg (${n} ${c.dataset._kind === "broken" ? "broken" : (n === 1 ? "job" : "jobs")})`;
            },
            afterBody: items => (items.length && items[0].dataIndex === partialIdx) ? ["  Hour still in progress"] : [],
          },
        },
      },
      scales: {
        x: { grid: { display: false }, border: { display: false },
             ticks: { color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "600" } } },
        y: { min: 0, max: yMax, border: { display: false },
             ticks: { stepSize: 10, color: "#ffffff", font: { family: CHART_FONT, size: 13, weight: "700" }, callback: v => v + "m" },
             grid: { color: "rgba(255,255,255,0.07)" },
             title: { display: true, text: "Average flow time", color: "#ffffff", font: { family: CHART_FONT, size: 14, weight: "700" } } },
      },
    },
    plugins: [bandPlugin, labelPlugin, flowDotPlugin],
  });
  flowStartDotLoop_(reduceMotion);

  // Legend (HTML) + comparison + breakage-rate cards
  const leg = document.getElementById("flowLegend");
  if (leg) {
    const btn = (key, idx, color, text, extraCls = "") => {
      const off = flowHidden_.has(key);
      return `<button type="button" class="tl-key tl-toggle${extraCls}${off ? " is-off" : ""}" style="--c:${color}"
        data-key="${key}" data-idx="${idx}" aria-pressed="${!off}" title="Click to ${off ? "show" : "hide"}">${text}</button>`;
    };
    leg.innerHTML = series.map((s, i) => {
      const none = s.real.every(v => v === null);
      return btn(s.key, i, s.color, s.label + (none ? " (none)" : ""));
    }).join("") + btn("broken", series.length, "#c4b5fd", "Broken jobs") +
      `<span class="tl-note">Click a name to hide or show it \u00b7 Faint lines along the bottom = no jobs in that group that hour</span>`;
    if (!leg._wired) {
      leg._wired = true;
      leg.addEventListener("click", e => {
        const b = e.target.closest(".tl-toggle");
        if (!b || !flowChart) return;
        const key = b.dataset.key, idx = Number(b.dataset.idx);
        const nowHidden = !flowHidden_.has(key);
        if (nowHidden) flowHidden_.add(key); else flowHidden_.delete(key);
        flowChart.setDatasetVisibility(idx, !nowHidden);
        flowChart.update("none");
        b.classList.toggle("is-off", nowHidden);
        b.setAttribute("aria-pressed", String(!nowHidden));
        b.title = `Click to ${nowHidden ? "show" : "hide"}`;
      });
    }
  }
  buildFlowCompareCards_(data);
}

/* Shared by Overall and By Machine: 15 / 30 min zones and guide lines */
function flowBandPlugin_(yMax) {
  return {
    id: "flowBands",
    beforeDatasetsDraw(chart) {
      const { ctx: c, chartArea: a, scales: { y } } = chart;
      const band = (lo, hi, col) => { c.fillStyle = col; c.fillRect(a.left, y.getPixelForValue(hi), a.right - a.left, y.getPixelForValue(lo) - y.getPixelForValue(hi)); };
      c.save();
      const top = v => Math.min(v, yMax);
      band(0, top(15), "rgba(74,222,128,0.05)");
      if (yMax > 15) band(15, top(30), "rgba(251,191,36,0.06)");
      if (yMax > 30) band(30, yMax, "rgba(255,107,107,0.05)");
      c.setLineDash([6, 5]); c.lineWidth = 1.2; c.font = `700 12px ${CHART_FONT}`; c.textAlign = "right";
      [[15, "#4ade80", "15 min"], [30, "#fbbf24", "30 min"]].filter(([v]) => v < yMax).forEach(([v, col, t]) => {
        const py = y.getPixelForValue(v);
        c.strokeStyle = col; c.globalAlpha = 0.6; c.beginPath(); c.moveTo(a.left, py); c.lineTo(a.right, py); c.stroke();
        c.globalAlpha = 1; c.fillStyle = col; c.fillText(t, a.right - 6, py - 6);
      });
      c.restore();
    },
  };
}

/* Shared by Overall and By Machine: round glowing particles moving along every line */
function flowDotPlugin_(reduceMotion) {
  return {
    id: "flowDots",
    afterDatasetsDraw(chart) {
      if (reduceMotion) return;
      const c = chart.ctx, a = chart.chartArea;
      // Round glowing particles (not dashes): near-zero dash + round cap = a circle
      const GAP = 20, offset = -((performance.now() / 1400) * GAP * 2) % GAP;
      c.save();
      c.beginPath(); c.rect(a.left, a.top - 6, a.right - a.left, a.bottom - a.top + 12); c.clip();
      c.lineCap = "round"; c.setLineDash([0.01, GAP]); c.lineDashOffset = offset;
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        for (let i = 0; i < meta.data.length - 1; i++) {
          const p0 = meta.data[i], p1 = meta.data[i + 1];
          if (!p0 || !p1 || p0.skip || p1.skip) continue;
          const lane = ds._real[i] === null || ds._real[i + 1] === null;
          c.strokeStyle = ds._tint; c.globalAlpha = lane ? 0.5 : 1;
          c.lineWidth = lane ? 3 : Math.max(5, (ds.borderWidth || 3) + 2);
          c.shadowColor = ds._tint; c.shadowBlur = lane ? 0 : 8;
          c.beginPath(); c.moveTo(p0.x, p0.y); c.lineTo(p1.x, p1.y); c.stroke();
        }
      });
      c.restore();
    },
  };
}

/* Lines the user turned off (by group key); kept across the 5-minute refreshes */
const flowHidden_ = new Set();

/* Full-screen view of the flow chart card (good for the floor TV) */
function toggleFlowFullscreen() {
  const card = document.getElementById("flowCard");
  if (!card) return;
  if (document.fullscreenElement) document.exitFullscreen();
  else if (card.requestFullscreen) card.requestFullscreen();
}
document.addEventListener("fullscreenchange", () => {
  const b = document.getElementById("flowFullBtn");
  if (b) b.textContent = document.fullscreenElement ? "Exit full screen" : "Full screen";
  if (flowChart) setTimeout(() => flowChart.resize(), 60);
});

/* Keeps the flow dots moving: redraw about 30 times a second, only while the chart is visible */
let flowDotLoop_ = null;
function flowStartDotLoop_(reduceMotion) {
  if (flowDotLoop_) cancelAnimationFrame(flowDotLoop_);
  flowDotLoop_ = null;
  if (reduceMotion) return;
  let last = 0;
  const tick = t => {
    flowDotLoop_ = requestAnimationFrame(tick);
    if (t - last < 33) return;
    last = t;
    const cv = document.getElementById("flowChart");
    if (!flowChart || document.hidden || !cv || cv.offsetParent === null) return;
    flowChart.draw();
  };
  flowDotLoop_ = requestAnimationFrame(tick);
}

/* Jobs-per-hour row under the flow chart */
function flowPlaceJobsRow_(chart, rows) {
  const row = document.getElementById("flowJobsRow");
  if (!row || !chart.scales || !chart.scales.x) return;
  const xS = chart.scales.x;
  const html = `<span class="tf-label" style="left:${(xS.left - 12).toFixed(0)}px">Jobs</span>` +
    rows.map((h, i) => `<div class="tf-cell" style="left:${xS.getPixelForValue(i).toFixed(0)}px"><span class="fj-n">${Number(h.coatingJobs || 0)}</span></div>`).join("");
  if (row._last !== html) { row.innerHTML = html; row._last = html; }
}

/* "Do broken jobs flow slower?" + "Breakage rate by flow group" */
function buildFlowCompareCards_(data) {
  const cmp = document.getElementById("flowCompare");
  const rate = document.getElementById("flowRate");
  const hourly = (data && data.hourly) || [];
  if (cmp) {
    let bSum = 0, bN = 0;
    hourly.forEach(h => { const n = Number(h.flowBrokenCount || 0); if (n > 0) { bSum += Number(h.avgFlowBroken || 0) * n; bN += n; } });
    const allAvg = Number((data.summary && data.summary.avgDetaper) || 0);
    if (!bN) {
      cmp.innerHTML = `<p class="fc-note">No broken jobs with flow data yet.</p>`;
    } else {
      const bAvg = bSum / bN, diff = bAvg - allAvg;
      const word = Math.abs(diff) < 0.5 ? "about the same as" : (diff > 0 ? `${diff.toFixed(1)} min longer than` : `${(-diff).toFixed(1)} min shorter than`);
      cmp.innerHTML = `
        <div class="fc-cmp">
          <div>Broken jobs<b style="color:#c4b5fd">${bAvg.toFixed(1)} min</b></div>
          <div>All jobs<b style="color:#3ddbc4">${allAvg.toFixed(1)} min</b></div>
          <div class="fc-say">Broken jobs took <b>${word}</b> all jobs on average${bN < 30 ? ` <span class="fc-small">(${bN} broken jobs, small sample)</span>` : ` (${bN} broken jobs)`}</div>
        </div>`;
    }
  }
  if (rate) {
    const fb = data && data.flowBreakage;
    if (!fb) {
      rate.innerHTML = `<p class="fc-note">Shows after the Coating API update (breakage by flow group).</p>`;
      return;
    }
    const defs = [["healthy", "Healthy", "#4ade80"], ["watch", "Watch", "#fbbf24"], ["delayed", "Delayed", "#ff6b6b"], ["overnight", "Overnight", "#60a5fa"]];
    const rows = defs.filter(([k]) => fb[k] && (fb[k].jobs > 0 || fb[k].brokenLenses > 0));
    const maxRate = Math.max(0.01, ...rows.map(([k]) => Number(fb[k].rate || 0)));
    rate.innerHTML = rows.map(([k, label, col]) => {
      const g = fb[k], r = Number(g.rate || 0);
      return `<div class="fr-row"><span>${label}</span>
        <div class="fr-track"><div style="width:${(r / maxRate * 100).toFixed(1)}%;background:${col}"></div></div>
        <b>${r.toFixed(2)}%</b><span class="fr-sub">${g.brokenLenses} of ${g.lenses} lenses</span></div>`;
    }).join("") + (fb.noFlowData ? `<p class="fc-note">${fb.noFlowData} same-day broken lenses had no flow time and are not included.</p>` : "");
  }
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
          color   : "rgb(56,189,248)",
          font    : { family: CHART_MONO, size: 12 },
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
      <div style="margin-bottom:10px"><span style="color:rgb(168,200,234);font-size:12px;text-transform:uppercase;letter-spacing:1.2px;">Machine</span><br><strong style="color:#c8dff5;">${point.machine}</strong></div>
      <div style="margin-bottom:10px"><span style="color:rgb(168,200,234);font-size:12px;text-transform:uppercase;letter-spacing:1.2px;">Breakage Reason</span><br><strong style="color:#c8dff5;">${point.reason || "None"}</strong></div>
      <div style="margin-bottom:10px"><span style="color:rgb(168,200,234);font-size:12px;text-transform:uppercase;letter-spacing:1.2px;">Flow Time</span><br><strong style="color:${GC.cyan};font-size:22px;">${fmt}</strong></div>
      <div><span style="color:rgb(168,200,234);font-size:12px;text-transform:uppercase;letter-spacing:1.2px;">Hour</span><br><strong style="color:#c8dff5;">${point.x}</strong></div>
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
  modalBody.innerHTML = `<h2>${hourData.hour}</h2>${html || "<p style='color:rgb(168,200,234)'>No flow data</p>"}`;
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

    let primaryDriverHTML  = "<span style='color:#ffffff'>—</span>";
    let topAccessPointHTML = "<span style='color:#ffffff'>—</span>";

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
        .map(([m, s]) => `<span style="color:#38bdf8;font-family:'JetBrains Mono',monospace;font-size:13px;">${formatMachineLabel(m)}</span><span style="color:#ffffff;font-size:13px;"> ×${s.total}</span>`)
        .join("<br>");
      if (brkMachines) topAccessPointHTML = brkMachines;
    }

    // ── Flow Breakdown ──
    const flowPts   = row.flowPoints || [];
    const totalFlow = flowPts.length;
    let flowHTML    = `<span style="color:#ffffff">—</span>`;

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
        <div class="flow-total-jobs">${totalFlow} flow ${totalFlow === 1 ? "job" : "jobs"}</div>
        <div class="flow-pills">${pills}</div>
      </div>`;
    }

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td class="${brkClass}">${broken > 0 ? broken : '<span style="color:#ffffff">0</span>'}</td>
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
        <div style="color:#ffffff;font-size:13px;margin-top:4px;">Same: ${rStats.sameDay||0} &nbsp;·&nbsp; Prev: ${rStats.oneDay||0} &nbsp;·&nbsp; 2+: ${rStats.twoPlus||0}</div>
      </div>`;
    });
    machineHTML += `<div style="margin-bottom:14px;padding:16px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1);border-radius:10px;line-height:1.8;">
      <div style="color:#ffffff;font-size:16px;font-weight:700;margin-bottom:8px;">${dm}</div>
      <div style="font-size:13px;margin-bottom:4px;">
        <span style="color:#ffffff;">JOBS:</span> <strong style="color:#ffffff;font-size:14px;">${stats.jobs||0}</strong>
        &nbsp;&nbsp;
        <span style="color:#ffffff;">BROKEN:</span> <strong style="color:${GC.red};font-size:14px;">${stats.total||0}</strong>
      </div>
      <div style="color:#ffffff;font-size:13px;">Same: ${stats.sameDay||0} &nbsp;·&nbsp; Prev: ${stats.oneDay||0} &nbsp;·&nbsp; 2+: ${stats.twoPlus||0}</div>
      ${reasonHTML}
    </div>`;
  });

  modalBody.innerHTML = `
    <h2 style="font-size:20px;color:#ffffff;margin-bottom:6px;">${hourData.hour}</h2>
    <div style="font-size:15px;margin-bottom:16px;">
      <span style="color:#ffffff;">Total Lenses Broken:</span>
      <strong style="color:${GC.red};font-size:22px;margin-left:8px;">${hourData.totalBroken||0}</strong>
    </div>
    ${machineHTML || "<p style='color:#ffffff;font-size:14px;'>No data</p>"}`;
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

/* =====================================================
   COATER BAY — one card per coater, rendered from machineTotals.
   Visual only: reads data already loaded by the page, no API calls.
   - Uses machineTotals ONLY (coaters). accessPointTotals are not
     coaters and are intentionally excluded here.
   - Lenses = jobs x 2, same basis as Total Lenses and the machine chart.
   - A coater with 0 jobs never shows a percentage (no fake 100%).
   - Cards keep a stable order (by label) so they don't jump on refresh.
===================================================== */
const COATER_BAY_THRESHOLDS = { watch: 1, high: 3, crit: 6 }; // same as Machine Analysis colors

function coaterBayEsc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c]));
}

function coaterBayStatus_(jobs, broken, pct) {
  if (jobs === 0 && broken === 0) return { cls: "st-idle",  label: "Idle" };
  if (jobs === 0)                 return { cls: "st-nojob", label: "Check data" };
  if (pct >= COATER_BAY_THRESHOLDS.crit)  return { cls: "st-crit",  label: "Critical" };
  if (pct >= COATER_BAY_THRESHOLDS.high)  return { cls: "st-high",  label: "High" };
  if (pct >= COATER_BAY_THRESHOLDS.watch) return { cls: "st-watch", label: "Watch" };
  return { cls: "st-ok", label: "Normal" };
}

function coaterBayArt_(uid) {
  // Stylized spin coater: steel housing, chamber window with lens, fluid
  // reservoir, on a lit platform. Accent strokes pick up the card status color.
  return `
<svg viewBox="0 0 200 150" role="img" aria-hidden="true">
  <defs>
    <linearGradient id="ccSteel${uid}" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#e3e9f0"/><stop offset="0.55" stop-color="#a9b6c4"/><stop offset="1" stop-color="#6c7a8a"/>
    </linearGradient>
    <linearGradient id="ccSide${uid}" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0" stop-color="#7b8998"/><stop offset="1" stop-color="#4a5664"/>
    </linearGradient>
    <radialGradient id="ccWin${uid}" cx="0.5" cy="0.45" r="0.6">
      <stop offset="0" stop-color="#1c2f3d"/><stop offset="1" stop-color="#050b10"/>
    </radialGradient>
  </defs>

  <!-- platform -->
  <ellipse cx="100" cy="132" rx="82" ry="13" fill="#03080b"/>
  <ellipse cx="100" cy="132" rx="82" ry="13" fill="none" class="glow" stroke-width="1" opacity="0.35"/>
  <ellipse cx="100" cy="128" rx="66" ry="10" fill="#08131a"/>
  <ellipse cx="100" cy="128" rx="66" ry="10" fill="none" class="glow" stroke-width="2.2" opacity="0.9"/>
  <ellipse cx="100" cy="128" rx="52" ry="7" fill="none" class="glow" stroke-width="1" opacity="0.45"/>
  <ellipse cx="100" cy="129" rx="60" ry="6" class="glowf" opacity="0.18"/>

  <!-- reservoir (right) -->
  <rect x="146" y="58" width="18" height="62" rx="5" fill="url(#ccSide${uid})"/>
  <rect x="149.5" y="74" width="11" height="40" rx="3" fill="#050b10"/>
  <rect x="149.5" y="88" width="11" height="26" rx="3" class="glowf" opacity="0.75"/>
  <path d="M146 70 H138" stroke="#8b98a6" stroke-width="3" stroke-linecap="round"/>

  <!-- housing -->
  <rect x="56" y="34" width="86" height="88" rx="9" fill="url(#ccSteel${uid})"/>
  <rect x="56" y="34" width="86" height="12" rx="6" fill="#f2f5f8" opacity="0.55"/>
  <rect x="136" y="38" width="6" height="80" rx="3" fill="#000" opacity="0.18"/>
  <path d="M68 40 H96 M68 43 H96" stroke="#5d6a78" stroke-width="1.2" opacity="0.7"/>

  <!-- chamber window -->
  <circle cx="99" cy="80" r="29" fill="#2b3642"/>
  <circle cx="99" cy="80" r="25" fill="url(#ccWin${uid})"/>
  <circle cx="99" cy="80" r="25" fill="none" class="glow" stroke-width="1.6" opacity="0.85"/>
  <circle cx="99" cy="80" r="18" class="glowf" opacity="0.12"/>
  <g class="lens">
    <ellipse cx="99" cy="80" rx="14" ry="14" fill="none" class="glow" stroke-width="1.2" opacity="0.9"/>
    <path d="M89 74 A12 12 0 0 1 104 69" fill="none" stroke="#ffffff" stroke-width="1.6" stroke-linecap="round" opacity="0.7"/>
    <circle cx="99" cy="80" r="2.2" class="glowf"/>
  </g>

  <!-- control strip -->
  <rect x="64" y="112" width="70" height="5" rx="2.5" fill="#3b4652"/>
  <circle cx="70" cy="114.5" r="1.8" class="glowf"/>
  <circle cx="76" cy="114.5" r="1.8" fill="#9aa7b4"/>
  <rect x="104" y="113" width="24" height="3" rx="1.5" class="glowf" opacity="0.6"/>
</svg>`;
}

function buildCoaterBay(data, targetId = "coaterBay") {
  const bay = document.getElementById(targetId);
  if (!bay) return;

  const totals = (data && data.machineTotals) || {};
  const entries = Object.entries(totals)
    .map(([machine, s]) => {
      const jobs   = Number((s && s.jobs) || 0);
      const broken = Number((s && s.breakLenses) || 0);
      const lenses = jobs * 2;
      const pct    = lenses > 0 ? (broken / lenses) * 100 : null;
      return { machine, label: formatMachineLabel(machine), jobs, broken, lenses, pct };
    })
    .sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true }));

  if (!entries.length) {
    bay.innerHTML = `<div class="bay-empty">No coater data for this date. Check that the machine report imported for the selected day.</div>`;
    return;
  }

  bay.innerHTML = entries.map((e, i) => {
    const st = coaterBayStatus_(e.jobs, e.broken, e.pct ?? 0);
    const running = e.jobs > 0 ? " is-running" : "";
    const name = coaterBayEsc_(e.label);

    let hero;
    if (e.pct !== null) {
      hero = `<div class="cc-pct">${e.pct.toFixed(2)}<small>%</small></div>`;
    } else if (e.broken > 0) {
      hero = `<div class="cc-pct is-text">No jobs recorded</div>`;
    } else {
      hero = `<div class="cc-pct is-text">No activity</div>`;
    }

    const brokenText = e.lenses > 0
      ? `<span><b>${e.broken}</b> of ${e.lenses} lenses</span>`
      : `<span><b>${e.broken}</b> broken</span>`;

    const aria = e.pct !== null
      ? `${name}: ${e.pct.toFixed(2)} percent breakage, ${e.broken} of ${e.lenses} lenses broken across ${e.jobs} jobs, status ${st.label}`
      : `${name}: ${e.jobs} jobs, ${e.broken} lenses broken, status ${st.label}`;

    return `
<article class="coater-card ${st.cls}${running}" aria-label="${aria}">
  <div class="cc-top"><span class="cc-name">${name}</span><span class="cc-status">${st.label}</span></div>
  <div class="cc-art">${coaterBayArt_(targetId + i)}</div>
  ${hero}
  <div class="cc-meta"><span><b>${e.jobs}</b> jobs</span>${brokenText}</div>
</article>`;
  }).join("");
}

/* =====================================================
   OVERVIEW FLOW (helpers + renderer)
   Reads only data the page already loads. No API calls.
   - Same-day breakage % = summary.aging.sameDay / summary.totalLenses
   - All breakage %      = summary.breakPercent (backend value, unchanged)
   - Projection is an ESTIMATE (current pace x time left), live mode only,
     shown only after the first full hour of the shift.
   - Shift windows from the project handoff; times use America/New_York.
===================================================== */
const OV_SHIFT = {
  weekday: { start: 7,   end: 17.5, label: "7:00 AM \u2013 5:30 PM" },
  weekend: { start: 6.5, end: 18.5, label: "6:30 AM \u2013 6:30 PM" },
};
const OV_BREAK_THRESHOLDS = { watch: 2, bad: 4 }; // same as the existing Coating Brkg % card

function ovNowNY_() {
  return new Date(new Date().toLocaleString("en-US", { timeZone: "America/New_York" }));
}
function ovHour24_(label) {
  const m = /(\d{1,2}):(\d{2})\s*(AM|PM)/i.exec(String(label || ""));
  if (!m) return null;
  let h = Number(m[1]) % 12;
  if (m[3].toUpperCase() === "PM") h += 12;
  return h + Number(m[2]) / 60;
}
function ovShortHour_(h) {
  const hh = Math.floor(h), ap = hh >= 12 ? "p" : "a", h12 = hh % 12 || 12;
  return `${h12}${ap}`;
}
function ovHourName_(h) {
  const hh = Math.floor(h), ap = hh >= 12 ? "PM" : "AM", h12 = hh % 12 || 12;
  return `${h12} ${ap}`;
}
function ovFmtDur_(hours) {
  const m = Math.max(0, Math.round(hours * 60));
  return `${Math.floor(m / 60)}h ${String(m % 60).padStart(2, "0")}m`;
}
function ovSet_(id, html) { const el = document.getElementById(id); if (el) el.innerHTML = html; }

function buildOverviewFlow(processed, machData) {
  const stage = document.getElementById("ovFlowStage");
  if (!stage) return;
  const esc     = coaterBayEsc_;
  const summary = (processed && processed.summary) || {};
  const isLive  = currentDate === null;
  const nowNY   = ovNowNY_();
  const nowH    = nowNY.getHours() + nowNY.getMinutes() / 60;

  // ── Shift window ──
  let dayRef = nowNY;
  if (!isLive) {
    const p = String(currentDate).split("/").map(Number);
    if (p.length === 3) dayRef = new Date(p[2], p[0] - 1, p[1]);
  }
  const dow   = dayRef.getDay();
  // Same rule as the Coating API getShiftFromDate(): Fri, Sat, Sun = weekend shift
  const shift = (dow === 5 || dow === 6 || dow === 0) ? OV_SHIFT.weekend : OV_SHIFT.weekday;
  const [shiftStartLabel, shiftEndLabel] = shift.label.split(" \u2013 ");
  const len     = shift.end - shift.start;
  const elapsed = isLive ? Math.min(Math.max(nowH - shift.start, 0), len) : len;
  const inShift = isLive && nowH >= shift.start && nowH < shift.end;
  const pct     = (elapsed / len) * 100;

  ovSet_("ovShiftWindow", esc(shift.label));
  const fill = document.getElementById("ovShiftFill");
  const nowMark = document.getElementById("ovShiftNow");
  if (fill) fill.style.width = pct.toFixed(1) + "%";
  if (nowMark) { nowMark.style.display = inShift ? "block" : "none"; nowMark.style.left = pct.toFixed(1) + "%"; }
  if (!isLive)                 ovSet_("ovShiftState", "Past day");
  else if (nowH < shift.start) ovSet_("ovShiftState", `Starts at ${esc(shiftStartLabel)}`);
  else if (!inShift)           ovSet_("ovShiftState", "Shift complete");
  else                         ovSet_("ovShiftState", `<b>${Math.round(pct)}%</b> \u00b7 ${ovFmtDur_(len - elapsed)} left`);

  // ── Numbers ──
  const jobs      = Number(summary.totalJobs || 0);
  const lenses    = Number(summary.totalLenses || 0);
  const brokenAll = Number(summary.totalBreakLenses || 0);
  const allPct    = parseFloat(summary.breakPercent) || 0;
  const fh        = summary.flowHealth || {};
  const healthy   = Number(fh.healthy || 0), watch = Number(fh.watch || 0);
  const delayed   = Number(fh.delayed || 0), overnight = Number(fh.overnight || 0);
  const aging     = summary.aging || null;

  ovSet_("ovCoatedLabel", isLive ? "Coated today" : "Coated this day");
  ovSet_("ovCoated", String(jobs));
  const pace = (inShift && elapsed >= 1) ? jobs / elapsed : null;
  ovSet_("ovCoatedSub", `jobs \u00b7 ${lenses} lenses${pace !== null ? ` \u00b7 ${Math.round(pace)}/hr` : ""}`);

  const groups = [
    { id: "ovHealthy",   n: healthy,   cls: "ov-node-ok",    color: "#4ade80" },
    { id: "ovWatchN",    n: watch,     cls: "ov-node-watch", color: "#fbbf24" },
    { id: "ovDelayedN",  n: delayed,   cls: "ov-node-bad",   color: "#ff6b6b" },
    { id: "ovOvernightN",n: overnight, cls: "ov-node-info",  color: "#60a5fa" },
  ];
  groups.forEach(g => {
    ovSet_(g.id, String(g.n));
    const node = document.getElementById(g.id)?.closest(".ov-node");
    if (node) node.classList.toggle("is-zero", g.n === 0);
  });

  ovSet_("ovBroken", String(brokenAll));
  if (aging) {
    const same = Number(aging.sameDay || 0), prev = Number(aging.oneDay || 0), old = Number(aging.twoPlus || 0);
    ovSet_("ovBrokenSub", `<span class="age-same">${same} same day</span> \u00b7 <span class="age-prev">${prev} previous day</span>${old ? ` \u00b7 <span class="age-old">${old} from 2+ days</span>` : ""}`);
  } else {
    ovSet_("ovBrokenSub", "age split not available");
  }

  // ── Diagram paths (viewBox 1200 x 240, matches node % positions) ──
  const svg = document.getElementById("ovFlowSvg");
  if (svg) {
    const ftot = healthy + watch + delayed + overnight;
    const w = n => (ftot > 0 && n > 0) ? Math.max(3, (n / ftot) * 36) : 0;
    const srcX = 266, srcY = 118, dstX = 934;
    const targetY = { ovHealthy: 28, ovWatchN: 88, ovDelayedN: 148, ovOvernightN: 208 };
    let base = "", pulses = "";
    let offset = -16;
    groups.forEach(g => {
      const sw = w(g.n);
      if (!sw) return;
      const y0 = (srcY + offset + sw / 2).toFixed(1);
      offset += sw;
      const ty = targetY[g.id];
      const d = `M${srcX} ${y0} C ${srcX + 260} ${y0}, ${dstX - 260} ${ty}, ${dstX} ${ty}`;
      base   += `<path d="${d}" stroke="${g.color}" stroke-opacity=".55" stroke-width="${sw.toFixed(1)}"/>`;
      pulses += `<path class="ov-pulse" d="${d}" stroke-opacity=".5" stroke-width="${Math.min(Math.max(sw * 0.3, 1.5), 5).toFixed(1)}"/>`;
    });
    if (brokenAll > 0) {
      const d = "M130 172 C 130 212, 200 212, 356 212";
      base   += `<path d="${d}" stroke="#ff6b6b" stroke-opacity=".6" stroke-width="3" stroke-linecap="round"/>`;
      pulses += `<path class="ov-pulse ov-pulse-break" d="${d}" stroke-opacity=".9" stroke-width="2.5"/>`;
    }
    svg.innerHTML = base + pulses;
  }

  // ── Integrity checks ──
  const warns = [];
  const ftot = healthy + watch + delayed + overnight;
  if (ftot > 0 && ftot !== jobs) warns.push(`Flow groups add up to ${ftot}, but ${jobs} jobs were coated.`);
  if (aging) {
    const at = Number(aging.sameDay || 0) + Number(aging.oneDay || 0) + Number(aging.twoPlus || 0);
    if (at !== brokenAll) warns.push(`The age split adds up to ${at}, but ${brokenAll} lenses are broken in total.`);
  }
  ovSet_("ovFlowWarn", warns.length ? "Check data: " + warns.map(esc).join(" ") : "");

  // ── Facts ──
  const sameEl = document.getElementById("ovSameDayPct");
  if (aging && lenses > 0) {
    const same = Number(aging.sameDay || 0);
    const sp = (same / lenses) * 100;
    ovSet_("ovSameDayPct", sp.toFixed(2) + "%");
    if (sameEl) sameEl.style.color = sp < OV_BREAK_THRESHOLDS.watch ? "var(--ov-ok)" : sp < OV_BREAK_THRESHOLDS.bad ? "var(--ov-watch)" : "var(--ov-bad)";
    ovSet_("ovSameDaySub", `${same} of ${lenses} \u00b7 all: ${allPct.toFixed(2)}%`);
  } else {
    ovSet_("ovSameDayPct", lenses > 0 ? allPct.toFixed(2) + "%" : "--");
    if (sameEl) sameEl.style.color = "";
    ovSet_("ovSameDaySub", lenses > 0 ? "all breakage (same-day split not available)" : "No lenses coated yet");
  }

  const mt = (machData && machData.machineTotals) || {};
  let topM = null, topMB = 0, mSum = 0;
  Object.entries(mt).forEach(([m, s]) => {
    const b = Number((s && s.breakLenses) || 0); mSum += b;
    if (b > topMB) { topMB = b; topM = m; }
  });
  const topEl = document.getElementById("ovTopCoater");
  ovSet_("ovTopCoater", topM ? esc(formatMachineLabel(topM)) : "None");
  if (topEl) topEl.style.color = topM ? "var(--ov-high)" : "var(--ov-ok)";
  ovSet_("ovTopCoaterSub", topM ? `${topMB} of ${mSum} lenses` : "No coater breakage");

  let topR = null, topRN = 0;
  Object.entries((processed && processed.topReasons) || {}).forEach(([r, s]) => {
    const n = Number((s && s.total) || 0); if (n > topRN) { topRN = n; topR = r; }
  });
  let peak = null;
  ((processed && processed.hourly) || []).forEach(r => {
    const b = Number(r.totalBroken || 0), h = ovHour24_(r.hour);
    if (h !== null && b > 0 && (!peak || b > peak.b)) peak = { h, b };
  });
  ovSet_("ovTopReason", topR ? esc(topR) : "None");
  ovSet_("ovTopReasonSub", topR ? `${topRN} lenses${peak ? ` \u00b7 peak ${ovHourName_(peak.h)}` : ""}` : "No reasons recorded");

  if (inShift && elapsed >= 1) {
    const projected = Math.round(jobs + (jobs / elapsed) * (len - elapsed));
    ovSet_("ovProjLabel", `Projected by ${esc(shiftEndLabel)}`);
    ovSet_("ovProjection", `~${projected}`);
    ovSet_("ovProjSub", "estimate at current pace");
  } else if (inShift) {
    ovSet_("ovProjLabel", `Projected by ${esc(shiftEndLabel)}`);
    ovSet_("ovProjection", "--");
    ovSet_("ovProjSub", "starts after the first hour");
  } else {
    ovSet_("ovProjLabel", "Jobs per shift hour");
    ovSet_("ovProjection", len > 0 ? String(Math.round(jobs / len)) : "--");
    ovSet_("ovProjSub", isLive && nowH < shift.start ? "shift not started" : `${jobs} jobs over ${len} hrs`);
  }
}

/* =====================================================
   MACHINE ANALYSIS PANELS
   Built only from fields the payload already has:
     machineTotals[m].breakLenses
     hourly[].hour, hourly[].machines[m].total, .reasons[r].total
   No per-machine hourly JOBS exist in the data, so these panels show
   broken-lens COUNTS only, never a per-hour percentage.
===================================================== */
function buildMachinePanels(data) {
  const reasonEl = document.getElementById("machineReasonMatrix");
  const hourEl   = document.getElementById("machineHourHeat");
  const recEl    = document.getElementById("machineReconcile");
  const ageEl    = document.getElementById("machineAgeNote");
  if (!reasonEl || !hourEl) return;

  const hourly = Array.isArray(data && data.hourly) ? [...data.hourly] : [];
  hourly.sort((a, b) => new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour));

  // Each bucket: { t: total, s: same day, o: previous day, p: 2+ days }
  const blank = () => ({ t: 0, s: 0, o: 0, p: 0 });
  const add = (bk, st) => {
    bk.t += Number((st && st.total) || 0);
    bk.s += Number((st && st.sameDay) || 0);
    bk.o += Number((st && st.oneDay) || 0);
    bk.p += Number((st && st.twoPlus) || 0);
  };
  const byMachine = {};   // m -> { all, reasons:{r:bucket}, hours:{h:bucket} }
  const reasonTotals = {};
  hourly.forEach(row => {
    Object.entries(row.machines || {}).forEach(([m, s]) => {
      if (!byMachine[m]) byMachine[m] = { all: blank(), reasons: {}, hours: {} };
      add(byMachine[m].all, s);
      if (!byMachine[m].hours[row.hour]) byMachine[m].hours[row.hour] = blank();
      add(byMachine[m].hours[row.hour], s);
      Object.entries((s && s.reasons) || {}).forEach(([r, rs]) => {
        if (!byMachine[m].reasons[r]) byMachine[m].reasons[r] = blank();
        add(byMachine[m].reasons[r], rs);
        reasonTotals[r] = (reasonTotals[r] || 0) + Number((rs && rs.total) || 0);
      });
    });
  });

  const machines = Object.keys(byMachine)
    .filter(m => byMachine[m].all.t > 0)
    .sort((a, b) => formatMachineLabel(a).localeCompare(formatMachineLabel(b), undefined, { numeric: true }));

  // Age coverage: how many broken lenses carry an age split
  let ageKnown = 0, ageTotal = 0;
  machines.forEach(m => { const x = byMachine[m].all; ageTotal += x.t; ageKnown += Math.min(x.t, x.s + x.o + x.p); });

  if (!machines.length) {
    const msg = `<div class="mx-empty">No broken lenses recorded by coater for this date.</div>`;
    reasonEl.innerHTML = msg; hourEl.innerHTML = msg;
  } else {
    const esc = coaterBayEsc_;
    // Numbers colored by age; cell brightness = total
    const ageBody = bk => {
      const unk = Math.max(0, bk.t - (bk.s + bk.o + bk.p));
      const parts = [[bk.s, "age-same", "same day"], [bk.o, "age-prev", "previous day"], [bk.p, "age-old", "2+ days"], [unk, "age-unk", "age unknown"]]
        .filter(p => p[0] > 0);
      return {
        title: parts.map(p => `${p[0]} ${p[2]}`).join(", "),
        html : parts.map(p => `<span class="${p[1]}">${p[0]}</span>`).join(`<span class="mx-plus">+</span>`)
      };
    };
    const cell = (bk, max) => {
      if (!bk || bk.t <= 0) return `<td class="mx-cell is-zero">0</td>`;
      const b = ageBody(bk);
      return `<td class="mx-cell" style="--a:${(bk.t / max).toFixed(3)}" title="${b.title}">${b.html}</td>`;
    };
    const totalCell = bk => { const b = ageBody(bk); return `<td class="mx-total" title="${b.title}"><b class="mx-total-n">${bk.t}</b><span class="mx-split">${b.html}</span></td>`; };

    // ── Reason matrix: top 6 reasons, rest grouped as "Other" (still counted) ──
    const reasons = Object.keys(reasonTotals).filter(r => reasonTotals[r] > 0)
      .sort((a, b) => reasonTotals[b] - reasonTotals[a]);
    const shown = reasons.slice(0, 6);
    const hasOther = reasons.length > shown.length;
    const rowVals = machines.map(m => {
      const vals = shown.map(r => byMachine[m].reasons[r] || blank());
      if (hasOther) {
        const o = blank();
        reasons.slice(6).forEach(r => { const x = byMachine[m].reasons[r]; if (x) { o.t += x.t; o.s += x.s; o.o += x.o; o.p += x.p; } });
        vals.push(o);
      }
      return vals;
    });
    const rMax = Math.max(1, ...rowVals.flat().map(v => v.t));
    const dot = r => {
      const col = (typeof BREAKAGE_COLOR_MAP !== "undefined" && BREAKAGE_COLOR_MAP[r]) || "#60a5fa";
      return `<span class="mx-reason-dot" style="background:${col}"></span>`;
    };
    reasonEl.innerHTML = reasons.length ? `
<table class="mx-table">
  <thead><tr><th class="mx-rowhead">Coater</th>${shown.map(r => `<th>${dot(r)}${esc(r)}</th>`).join("")}${hasOther ? "<th>Other</th>" : ""}<th class="mx-total">Total</th></tr></thead>
  <tbody>${machines.map((m, i) => `<tr><td class="mx-rowhead">${esc(formatMachineLabel(m))}</td>${rowVals[i].map(v => cell(v, rMax)).join("")}${totalCell(byMachine[m].all)}</tr>`).join("")}</tbody>
</table>` : `<div class="mx-empty">Breakage was recorded without a reason.</div>`;

    // ── Hour heatmap: every hour present in the report ──
    const hours = hourly.map(r => r.hour);
    const hMax = Math.max(1, ...machines.flatMap(m => hours.map(h => (byMachine[m].hours[h] || blank()).t)));
    const shortHour = h => String(h).replace(":00", "").replace(" AM", "a").replace(" PM", "p");
    hourEl.innerHTML = `
<table class="mx-table mx-hour">
  <thead><tr><th class="mx-rowhead">Coater</th>${hours.map(h => `<th title="${esc(h)}">${esc(shortHour(h))}</th>`).join("")}<th class="mx-total">Total</th></tr></thead>
  <tbody>${machines.map(m => `<tr><td class="mx-rowhead">${esc(formatMachineLabel(m))}</td>${hours.map(h => cell(byMachine[m].hours[h], hMax)).join("")}${totalCell(byMachine[m].all)}</tr>`).join("")}</tbody>
</table>`;
  }

  // ── Age coverage note ──
  if (ageEl) {
    if (!ageTotal || ageKnown === ageTotal) ageEl.textContent = "";
    else if (ageKnown === 0) ageEl.textContent = "This report has no same-day / previous-day split, so all numbers show as age unknown.";
    else ageEl.textContent = `${ageTotal - ageKnown} of ${ageTotal} broken lenses have no age split and show as age unknown.`;
  }

  // ── Reconciliation: hourly detail vs machine totals ──
  if (recEl) {
    const totals = (data && data.machineTotals) || {};
    const totalFromTotals = Object.values(totals).reduce((s, v) => s + Number((v && v.breakLenses) || 0), 0);
    const totalFromHourly = Object.values(byMachine).reduce((s, v) => s + v.all.t, 0);
    if (!hourly.length && !totalFromTotals) {
      recEl.textContent = "";
      recEl.classList.remove("is-warn");
    } else if (totalFromHourly === totalFromTotals) {
      recEl.textContent = `Check passed: hourly detail and coater totals both show ${totalFromTotals} broken lenses.`;
      recEl.classList.remove("is-warn");
    } else {
      recEl.textContent = `Data mismatch: hourly detail shows ${totalFromHourly} broken lenses, coater totals show ${totalFromTotals}. Check the machine report import before trusting these panels.`;
      recEl.classList.add("is-warn");
    }
  }
}

function buildMachineChart(data) {
  if (!data || !data.machineTotals) return;
  const canvas = document.getElementById("machineChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (machineChart) machineChart.destroy();

  // Merge coater machineTotals + accessPointTotals (FLEX/tray positions) when present
  const combined = { ...(data.machineTotals || {}) };
  if (data.accessPointTotals) {
    // Access points like "54R 2B" are the SAME coater as machineTotals "54R-02-B".
    // Adding them again double-counted breakage and drew fake 100% bars.
    // Match on a normalized name; only keep access points that are not a coater.
    const normKey = s => String(s).toUpperCase().replace(/[^A-Z0-9]/g, "").replace(/(\D)0+(\d)/g, "$1$2");
    const coaterByNorm = {};
    Object.keys(combined).forEach(k => { coaterByNorm[normKey(k)] = k; });
    Object.entries(data.accessPointTotals).forEach(([ap, stats]) => {
      const apBroken = Number(stats.breakLenses || stats.total || 0);
      const match = combined[ap] ? ap : coaterByNorm[normKey(ap)];
      if (match) {
        const coaterBroken = Number(combined[match].breakLenses || 0);
        if (apBroken !== coaterBroken) {
          console.warn(`[Machine chart] ${ap} reports ${apBroken} broken, coater ${match} reports ${coaterBroken}. Using coater value.`);
        }
        return; // already counted under the coater
      }
      combined[ap] = { jobs: 0, breakLenses: apBroken };
    });
  }

  const entries = Object.entries(combined)
    .map(([machine, stats]) => {
      const jobs  = Number(stats.jobs || 0);
      const broken = Number(stats.breakLenses || 0);
      const total  = jobs * 2;
      // No jobs = no valid percentage. Bar stays at 0; label/tooltip show the broken count.
      const percent = total > 0 ? (broken / total) * 100 : 0;
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
          ticks : { color: "rgb(160,200,240)", font: { family: CHART_FONT, size: 12, weight: "500" }, padding: 6 },
          grid  : { display: false },
          border: { display: false },
        },
        y1: {
          position   : "left",
          beginAtZero: true,
          suggestedMax: Math.ceil(maxPercent * 1.4),
          ticks: { color: "rgb(248,113,113)", font: { family: CHART_MONO, size: 12 }, callback: v => v + "%" },
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

/* =====================================================
   DAILY SUMMARY — end-of-shift report for the selected date
   - Leads with the day's main issue (flow or breakage), found by written rules
   - Coater report card, notable events (each shows its rule), flow groups,
     breakage by reason, shift hour by hour
   - Compares with the PREVIOUS DAY (one extra API call, cached; past days are
     cached 6 h on the API side too)
   - Previous / next day buttons switch the whole dashboard to that date
===================================================== */
const DAILY_RULES = {
  slowShare      : 15,   // % of jobs over 15 min flow → "flow is slow"
  coaterShare    : 50,   // % of breakage from one coater → concentration
  minBrokenForMix: 4,    // don't call concentration on tiny totals
  worstHourMin   : 3,    // broken lenses in one hour to call it out
  topTwoReasons  : 75,   // % of breakage from top two reasons
  delayedCluster : 3,    // delayed jobs in one hour
  smallSampleJobs: 20,   // coater status needs at least this many jobs
};
const dailyPrevInFlight_ = {};

function dailyDateKey_(offsetDays) {
  // selected date (or today in New York) shifted by N days → "M/D/YYYY"
  let base;
  if (currentDate) { const p = String(currentDate).split("/").map(Number); base = new Date(p[2], p[0] - 1, p[1]); }
  else base = ovNowNY_();
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate() + offsetDays);
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

function dailyStepDay(delta) {
  const key = dailyDateKey_(delta);
  const [m, d, y] = key.split("/").map(Number);
  const iso = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  if (iso >= coatingTodayISO_()) { resetToToday(); return; }          // today or later = live
  const el = document.getElementById("historyDate");
  if (el) { el.value = iso; applyDateFilter(); }
}

function dailyLoadPrevDay_() {
  const key = dailyDateKey_(-1);
  if (weeklyCache[key] !== undefined || dailyPrevInFlight_[key]) return;
  dailyPrevInFlight_[key] = true;
  coatingFetchRetry_(`${API_URL}?mode=processed&date=${encodeURIComponent(key)}`, "Daily previous-day")
    .then(r => r.ok ? r.json() : null)
    .then(j => { weeklyCache[key] = j; })
    .catch(() => { weeklyCache[key] = null; })
    .finally(() => {
      delete dailyPrevInFlight_[key];
      const tab = document.getElementById("daily");
      if (tab && tab.classList.contains("active")) buildDailySummary(dashboardData);
    });
}

function dailySameDay_(summary) {
  const lenses = Number(summary.totalLenses || 0);
  const a = summary.aging || null;
  if (!a || !lenses) return null;
  const s = Number(a.sameDay || 0);
  return { n: s, prev: Number(a.oneDay || 0), old: Number(a.twoPlus || 0), pct: (s / lenses) * 100 };
}

function buildDailySummary(data) {
  if (!data) return;
  const esc     = coaterBayEsc_;
  const summary = data.summary || {};
  const hourly  = Array.isArray(data.hourly) ? data.hourly : [];
  const isLive  = currentDate === null;
  const nowNY   = ovNowNY_();
  const nowH    = nowNY.getHours() + nowNY.getMinutes() / 60;

  // ── Date + shift ──
  const key = dailyDateKey_(0);
  const [mm, dd, yy] = key.split("/").map(Number);
  const dayObj = new Date(yy, mm - 1, dd);
  const dow = dayObj.getDay();
  const shift = (dow === 5 || dow === 6 || dow === 0) ? OV_SHIFT.weekend : OV_SHIFT.weekday;
  const shiftName = (dow === 5 || dow === 6 || dow === 0) ? "Weekend shift" : "Weekday shift";
  const dateLabel = dayObj.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric", year: "numeric" });
  const shiftLen = shift.end - shift.start;
  const shiftPct = isLive ? Math.max(0, Math.min(100, ((nowH - shift.start) / shiftLen) * 100)) : 100;
  ovSet_("dailyDateLabel", esc(dateLabel));
  ovSet_("dailyShiftLabel", `${shiftName} ${esc(shift.label)}`);
  ovSet_("dailyStatusLabel", isLive
    ? `Live \u00b7 updated ${nowNY.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" })}`
    : "Full day");
  const nextBtn = document.getElementById("dailyNextBtn");
  if (nextBtn) nextBtn.disabled = isLive;

  // ── Numbers ──
  const jobs    = Number(summary.totalJobs || 0);
  const lenses  = Number(summary.totalLenses || 0);
  const broken  = Number(summary.totalBreakLenses || 0);
  const allPct  = parseFloat(summary.breakPercent) || 0;
  const avgFlow = Number(summary.avgDetaper || 0);
  const fh      = summary.flowHealth || {};
  const H = Number(fh.healthy || 0), W = Number(fh.watch || 0), D = Number(fh.delayed || 0), O = Number(fh.overnight || 0);
  const flowTot = H + W + D + O;
  const slow    = W + D + O;
  const slowPct = flowTot ? (slow / flowTot) * 100 : 0;
  const same    = dailySameDay_(summary);

  // Previous day (for comparison)
  const prevKey = dailyDateKey_(-1);
  const prev = weeklyCache[prevKey];
  if (prev === undefined) dailyLoadPrevDay_();
  const ps = (prev && prev.summary) || null;
  const prevName = new Date(yy, mm - 1, dd - 1).toLocaleDateString("en-US", { weekday: "short" });
  const cmp = (now, then, unit, better) => {
    if (then === null || then === undefined || isNaN(then)) return prev === undefined ? "loading previous day\u2026" : "no previous-day data";
    const d = now - then;
    if (Math.abs(d) < (unit === "%" ? 0.05 : 0.5)) return `same as ${prevName}`;
    const worse = better === "lower" ? d > 0 : d < 0;
    const txt = unit === "%" ? `${Math.abs(d).toFixed(Math.abs(d) >= 10 ? 0 : 1)} pts` : `${Math.abs(d).toFixed(unit === "min" ? 1 : 0)}${unit === "min" ? " min" : ""}`;
    return `<span class="${worse ? "dl-worse" : "dl-better"}">${d > 0 ? "+" : "\u2212"}${txt}</span> vs ${prevName}`;
  };
  const prevSame = ps ? dailySameDay_(ps) : null;
  // Live: compare jobs with the previous day UP TO THE SAME HOUR (a partial day vs a full day is misleading)
  let prevJobsCmp = ps ? Number(ps.totalJobs || 0) : null;
  if (ps && isLive && Array.isArray(prev.hourly)) {
    const cut = Math.floor(nowH);
    prevJobsCmp = prev.hourly.reduce((s, h) => { const k = ovHour24_(h.hour); return s + (k !== null && Math.floor(k) <= cut ? Number(h.coatingJobs || 0) : 0); }, 0);
  }
  const prevFH = ps ? (ps.flowHealth || {}) : {};
  const prevFlowTot = ps ? ["healthy", "watch", "delayed", "overnight"].reduce((s, k) => s + Number(prevFH[k] || 0), 0) : 0;
  const prevSlowPct = ps && prevFlowTot ? ((Number(prevFH.watch || 0) + Number(prevFH.delayed || 0) + Number(prevFH.overnight || 0)) / prevFlowTot) * 100 : null;

  // ── Coaters (from hourly detail) ──
  const coat = {};
  hourly.forEach(h => {
    Object.entries(h.machines || {}).forEach(([m, s]) => {
      const c = coat[m] || (coat[m] = { jobs: 0, broken: 0, fSum: 0, fN: 0 });
      c.jobs += Number(s.jobs || 0);
      c.broken += Number(s.total || 0);
    });
    (h.flowPoints || []).forEach(p => {
      const f = Number(p.flow || 0);
      if (f > 0 && coat[p.machine]) { coat[p.machine].fSum += f; coat[p.machine].fN++; }
    });
  });
  const coaters = Object.keys(coat).filter(m => coat[m].jobs > 0 || coat[m].broken > 0)
    .sort((a, b) => coat[b].broken - coat[a].broken || formatMachineLabel(a).localeCompare(formatMachineLabel(b), undefined, { numeric: true }));
  const status = c => {
    if (c.jobs < DAILY_RULES.smallSampleJobs) return { t: "Small sample", col: "#ffffff" };
    const r = (c.broken / (c.jobs * 2)) * 100;
    if (r >= COATER_BAY_THRESHOLDS.crit)  return { t: "Critical", col: "#ff6b6b" };
    if (r >= COATER_BAY_THRESHOLDS.high)  return { t: "High",     col: "#fb923c" };
    if (r >= COATER_BAY_THRESHOLDS.watch) return { t: "Watch",    col: "#fbbf24" };
    return { t: "Normal", col: "#4ade80" };
  };

  // ── Reasons ──
  const reasons = Object.entries(data.topReasons || {})
    .map(([r, s]) => ({ r, t: Number(s.total || 0), s: Number(s.sameDay || 0), o: Number(s.oneDay || 0) + Number(s.twoPlus || 0) }))
    .filter(x => x.t > 0).sort((a, b) => b.t - a.t);

  // ── Hours ──
  const rowsByH = {};
  hourly.forEach(h => { const k = ovHour24_(h.hour); if (k !== null) rowsByH[Math.floor(k)] = h; });
  const hourSet = new Set(Object.keys(rowsByH).map(Number));
  for (let h = Math.floor(shift.start); h < Math.ceil(shift.end); h++) hourSet.add(h);
  const hourList = [...hourSet].sort((a, b) => a - b);

  // ── Notable events (each with its rule) ──
  const events = [];
  if (flowTot && slowPct > DAILY_RULES.slowShare)
    events.push({ col: "#ff6b6b", w: 3, t: `Flow is slow: ${slowPct.toFixed(0)}% of jobs took over 15 min`, rule: `more than ${DAILY_RULES.slowShare}% of the day's jobs on watch or slower` });
  if (same && same.pct >= OV_BREAK_THRESHOLDS.watch)
    events.push({ col: same.pct >= OV_BREAK_THRESHOLDS.bad ? "#ff6b6b" : "#fbbf24", w: 4, t: `Same-day breakage is ${same.pct.toFixed(2)}%`, rule: `${OV_BREAK_THRESHOLDS.watch}% or higher` });
  const cBroken = coaters.reduce((s, m) => s + coat[m].broken, 0);
  if (coaters.length && cBroken >= DAILY_RULES.minBrokenForMix) {
    const top = coaters[0], share = (coat[top].broken / cBroken) * 100;
    if (share > DAILY_RULES.coaterShare)
      events.push({ col: "#fb923c", w: 2, t: `${formatMachineLabel(top)} caused ${coat[top].broken} of ${cBroken} broken lenses (${share.toFixed(0)}%)`, rule: `one coater has more than ${DAILY_RULES.coaterShare}% of the day's breakage` });
  }
  let worst = null;
  hourList.forEach(h => { const r = rowsByH[h]; const b = r ? Number(r.totalBroken || 0) : 0; if (b >= DAILY_RULES.worstHourMin && (!worst || b > worst.b)) worst = { h, b }; });
  if (worst) events.push({ col: "#fbbf24", w: 1, t: `${ovHourName_(worst.h)} was the worst hour: ${worst.b} lenses broken`, rule: `hour with the most broken lenses (at least ${DAILY_RULES.worstHourMin})` });
  hourList.forEach(h => { const r = rowsByH[h]; const d2 = r ? Number(r.flowDelayed || 0) : 0;
    if (d2 >= DAILY_RULES.delayedCluster) events.push({ col: "#ff6b6b", w: 1.5, t: `${ovHourName_(h)}: ${d2} delayed jobs (over 30 min)`, rule: `${DAILY_RULES.delayedCluster} or more delayed jobs in one hour` }); });
  if (reasons.length >= 2 && broken >= DAILY_RULES.minBrokenForMix) {
    const two = reasons[0].t + reasons[1].t, share = (two / broken) * 100;
    if (share > DAILY_RULES.topTwoReasons)
      events.push({ col: "#c4b5fd", w: 1, t: `${reasons[0].r} and ${reasons[1].r}: ${two} of ${broken} broken lenses`, rule: `top two reasons cover more than ${DAILY_RULES.topTwoReasons}% of breakage` });
  }
  events.sort((a, b) => b.w - a.w);

  // ── Verdict ──
  let vClass = "ok", vHtml;
  const topCoater = coaters.length && coat[coaters[0]].broken > 0 ? formatMachineLabel(coaters[0]) : null;
  const topReason = reasons.length ? reasons[0].r : null;
  const brkTail = broken
    ? `Breakage is ${broken} lenses (${allPct.toFixed(2)}%)${topCoater ? `; <b>${esc(topCoater)}</b> caused ${coat[coaters[0]].broken} of them` : ""}${topReason ? `, mostly ${esc(topReason)}` : ""}.`
    : "No broken lenses.";
  if (!jobs) {
    vClass = "none"; vHtml = isLive ? "No jobs coated yet today." : "No coating jobs recorded for this date.";
  } else if (same && same.pct >= OV_BREAK_THRESHOLDS.watch) {
    vClass = same.pct >= OV_BREAK_THRESHOLDS.bad ? "bad" : "warn";
    vHtml = `<b>Breakage is the problem${isLive ? " today" : ""}:</b> same-day breakage is <b>${same.pct.toFixed(2)}%</b> (${same.n} of ${lenses} lenses). ${brkTail}`;
  } else if (slowPct > DAILY_RULES.slowShare) {
    vClass = "warn";
    vHtml = `<b>Flow is the problem${isLive ? " today" : ""}:</b> jobs average <b>${avgFlow} min</b> detaper to coater, with ${W} on watch and ${D + O} delayed (${slowPct.toFixed(0)}% of all jobs). ${brkTail}`;
  } else {
    vHtml = `<b>A normal ${isLive ? "day so far" : "day"}:</b> ${jobs} jobs coated, average flow ${avgFlow} min. ${brkTail}`;
  }
  const vEl = document.getElementById("dailyVerdict");
  if (vEl) vEl.className = "dl-verdict is-" + vClass;
  ovSet_("dailyVerdictKicker", isLive ? "THE DAY SO FAR" : "THE DAY");
  ovSet_("dailyVerdictText", vHtml);

  // ── KPI cards ──
  ovSet_("dailyKpis", [
    `<div class="dl-kpi">Jobs coated<b style="color:#3ddbc4">${jobs}</b><small>${lenses} lenses${isLive ? ` \u00b7 ${shiftPct.toFixed(0)}% of shift done` : ""} \u00b7 ${ps ? cmp(jobs, prevJobsCmp, "", "higher") + (isLive ? " at this time" : "") : cmp(jobs, null)}</small></div>`,
    `<div class="dl-kpi">Broken lenses<b style="color:${broken ? "#ff6b6b" : "#4ade80"}">${broken}</b><small>${allPct.toFixed(2)}% of ${lenses} lenses</small></div>`,
    same
      ? `<div class="dl-kpi">Same-day breakage<b style="color:${same.pct < OV_BREAK_THRESHOLDS.watch ? "#4ade80" : same.pct < OV_BREAK_THRESHOLDS.bad ? "#fbbf24" : "#ff6b6b"}">${same.pct.toFixed(2)}%</b><small>${same.n} same day \u00b7 ${same.prev + same.old} earlier \u00b7 ${cmp(same.pct, prevSame ? prevSame.pct : null, "%", "lower")}</small></div>`
      : `<div class="dl-kpi">Same-day breakage<b>--</b><small>split not available</small></div>`,
    `<div class="dl-kpi">Average flow time<b style="color:${avgFlow <= 15 ? "#4ade80" : avgFlow <= 30 ? "#fbbf24" : "#ff6b6b"}">${avgFlow} min</b><small>${cmp(avgFlow, ps ? Number(ps.avgDetaper || 0) : null, "min", "lower")}</small></div>`,
    `<div class="dl-kpi">Slow jobs (over 15 min)<b style="color:${slowPct > DAILY_RULES.slowShare ? "#ff6b6b" : "#ffffff"}">${slow}</b><small>${slowPct.toFixed(0)}% of jobs \u00b7 ${cmp(slowPct, prevSlowPct, "%", "lower").replace(" pts", " pts")}</small></div>`,
  ].join(""));

  // ── Coater report card ──
  ovSet_("dailyCoaters", coaters.length ? `
    <table class="dl-table"><thead><tr><th>Coater</th><th>Jobs</th><th>Broken</th><th>Rate</th><th>Avg flow</th><th>Status</th></tr></thead><tbody>
    ${coaters.map(m => {
      const c = coat[m], st = status(c), rate = c.jobs ? (c.broken / (c.jobs * 2)) * 100 : null;
      return `<tr><td class="dl-m">${esc(formatMachineLabel(m))}</td><td>${c.jobs}</td>
        <td style="font-weight:800;color:${c.broken ? st.col : "#ffffff"}">${c.broken}</td>
        <td>${rate === null ? "--" : rate.toFixed(2) + "%"}</td>
        <td>${c.fN ? (c.fSum / c.fN).toFixed(1) + " min" : "--"}</td>
        <td><span class="dl-pill" style="color:${st.col}">${st.t}</span></td></tr>`;
    }).join("")}</tbody></table>
    <p class="dl-foot">Rate = broken lenses \u00f7 (jobs \u00d7 2), all breakage found this day. Status uses the Machine Analysis limits (${COATER_BAY_THRESHOLDS.watch}% / ${COATER_BAY_THRESHOLDS.high}% / ${COATER_BAY_THRESHOLDS.crit}%); under ${DAILY_RULES.smallSampleJobs} jobs = small sample.</p>`
    : `<p class="dl-empty">No coater data for this date.</p>`);

  // ── Events ──
  ovSet_("dailyEvents", events.length
    ? events.map(e => `<div class="dl-ev" style="--c:${e.col}"><i></i><div><b>${esc(e.t)}</b><div>Rule: ${esc(e.rule)}</div></div></div>`).join("")
    : `<p class="dl-empty">Nothing unusual found by the rules.</p>`);

  // ── Flow groups ──
  const fb = data.flowBreakage || null;
  const grp = [["Healthy", H, "#4ade80", "healthy"], ["Watch", W, "#fbbf24", "watch"], ["Delayed", D, "#ff6b6b", "delayed"], ["Overnight", O, "#60a5fa", "overnight"]];
  ovSet_("dailyFlow", flowTot ? `
    <div class="dl-split">${grp.filter(g => g[1] > 0).map(g => `<div style="flex:${g[1]};background:${g[2]}">${g[1] / flowTot >= 0.06 ? g[1] : ""}</div>`).join("")}</div>
    <div class="dl-lg">${grp.map(g => `<span style="--c:${g[2]}">${g[0]} <b>${g[1]}</b> (${(g[1] / flowTot * 100).toFixed(0)}%)${fb && fb[g[3]] && fb[g[3]].lenses ? ` \u00b7 breaks ${Number(fb[g[3]].rate || 0).toFixed(2)}%` : ""}</span>`).join("")}</div>`
    : `<p class="dl-empty">No flow data.</p>`);

  // ── Reasons ──
  const rMax = Math.max(1, ...reasons.map(x => x.t));
  ovSet_("dailyReasons", reasons.length ? reasons.slice(0, 6).map(x => {
    const col = BREAKAGE_COLOR_MAP[x.r] || "#60a5fa";
    return `<div class="dl-rr"><span>${esc(x.r)}</span><div class="dl-t"><div style="width:${(x.t / rMax * 100).toFixed(0)}%;background:${col}"></div></div><b>${x.t}</b>
      <small><span class="age-same">${x.s} same day</span>${x.o ? ` \u00b7 <span class="age-prev">${x.o} earlier</span>` : ""}</small></div>`;
  }).join("") : `<p class="dl-empty">No broken lenses.</p>`);

  // ── Shift hour by hour ──
  const maxJ = Math.max(1, ...hourList.map(h => rowsByH[h] ? Number(rowsByH[h].coatingJobs || 0) : 0));
  ovSet_("dailyHours", hourList.map(h => {
    const r = rowsByH[h], future = isLive && h > Math.floor(nowH), now = isLive && h === Math.floor(nowH);
    if (!r || future) return `<div class="dl-hr${future ? " is-fut" : ""}"><div class="dl-hh">${ovHourName_(h)}</div><div class="dl-bar"></div><div class="dl-hj">${future ? "\u2014" : "0"}</div><div class="dl-hb">&nbsp;</div></div>`;
    const j2 = Number(r.coatingJobs || 0), b2 = Number(r.totalBroken || 0);
    const bc = b2 >= 5 ? "bad" : b2 > 0 ? "warn" : "ok";
    return `<div class="dl-hr${now ? " is-now" : ""}" title="${ovHourName_(h)}: ${j2} jobs, ${b2} broken"><div class="dl-hh">${ovHourName_(h)}</div>
      <div class="dl-bar"><div class="${now ? "is-part" : ""}" style="height:${(j2 / maxJ * 100).toFixed(0)}%"></div></div>
      <div class="dl-hj">${j2} <span>jobs</span></div><div class="dl-hb ${bc}">${b2 ? b2 + " broken" : "\u2713 0 broken"}</div></div>`;
  }).join(""));

  // ── Reconciliation ──
  const hJobs = hourly.reduce((s, h) => s + Number(h.coatingJobs || 0), 0);
  const hBrk  = hourly.reduce((s, h) => s + Number(h.totalBroken || 0), 0);
  const notes = [];
  if (hJobs !== jobs) notes.push(`hourly jobs add up to ${hJobs}, total says ${jobs}`);
  if (hBrk !== broken) notes.push(`hourly breakage adds up to ${hBrk}, total says ${broken}`);
  if (flowTot && flowTot !== jobs) notes.push(`flow groups add up to ${flowTot}, jobs are ${jobs}`);
  const chk = document.getElementById("dailyCheck");
  if (chk) {
    chk.className = "dl-check" + (notes.length ? " is-warn" : "");
    chk.textContent = notes.length ? "Check data: " + notes.join("; ") + "." : `Check passed: hours, coaters, and flow groups all add up (${jobs} jobs, ${broken} broken lenses).`;
  }
}


/* =====================================================
   WEEKLY SUMMARY
===================================================== */

/* =====================================================
   WEEKLY REVIEW — for managers and Quality
   Reads top to bottom in three levels:
     1. The week in 30 seconds (what happened + where to look first)
     2. What changed (6 mini trend tiles + 7 day cards)
     3. For Quality: reasons (with daily lines), coaters by RATE, day × hour heatmap
   - Uses the 7 days ending on the "Week ending" date (same as before)
   - Today is "in progress": it never counts as best/worst and stays out of
     the trend tiles until the shift ends
   - No extra API calls beyond the 7 days this tab already loaded
===================================================== */
let weeklyTrendChart = null;          // kept: other code may still reference it
const weeklyCache    = {};            // "M/D/YYYY" -> processed payload
const WEEK_RULES = {
  risingStreak  : 3,    // days in a row with a higher rate than the day before
  flowRisePct   : 50,   // % rise in average flow (first -> last complete day)
  reasonShare   : 40,   // a reason this share of the week's breakage gets called out
  shareGrowPts  : 10,   // share growth (pts) to say "and its share grew"
  hotHour       : 10,   // broken lenses in one hour to call it out
  coaterRateHigh: 3,    // % coater rate to call out
  smallSample   : 20,   // coater jobs below this = small sample
};

function weeklyIsoToApi_(iso) { const [y, m, d] = iso.split("-").map(Number); return `${m}/${d}/${y}`; }
function weeklyApiToIso_(api) { const [m, d, y] = api.split("/").map(Number); return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`; }

function weeklyStep(deltaDays) {
  const el = document.getElementById("weekEndDate");
  if (!el) return;
  const today = coatingTodayISO_();
  let iso = el.value || today;
  if (deltaDays === 0) iso = today;
  else if (deltaDays === null) iso = (el.value && el.value <= today) ? el.value : today;   // picked in the date box
  else {
    const [y, m, d] = iso.split("-").map(Number);
    const dt = new Date(y, m - 1, d + deltaDays);
    iso = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
    if (iso > today) iso = today;
  }
  el.value = iso;
  showSummaryLoader("weekly");
  buildWeeklySummary().finally(() => hideSummaryLoader("weekly"));
}

function weeklyOpenDay_(api) {
  const iso = weeklyApiToIso_(api);
  if (iso >= coatingTodayISO_()) resetToToday();
  else { const el = document.getElementById("historyDate"); if (el) { el.value = iso; applyDateFilter(); } }
  const btn = document.querySelector(`.tab[onclick*="'daily'"]`);
  showTab("daily", btn);
}

function weeklySpark_(vals, color, opts = {}) {
  // vals: numbers for complete days, null for days without data / in progress
  const W = opts.w || 300, H = opts.h || 92, labels = opts.labels !== false;
  const X0 = 12, X1 = W - 12, T = 10, B = H - (labels ? 22 : 6);
  const real = vals.filter(v => v !== null);
  const X = i => X0 + (X1 - X0) * i / Math.max(1, vals.length - 1);
  let s = `<svg viewBox="0 0 ${W} ${H}" class="wk-spark" aria-hidden="true">`;
  if (real.length) {
    let lo = Math.min(...real), hi = Math.max(...real);
    if (hi === lo) { hi = lo + 1; lo = Math.max(0, lo - 1); }
    const Y = v => B - (v - lo) / (hi - lo) * (B - T);
    const runs = []; let cur = [];
    vals.forEach((v, i) => { if (v === null) { if (cur.length) runs.push(cur); cur = []; } else cur.push(i); });
    if (cur.length) runs.push(cur);
    runs.forEach(run => {
      const pts = run.map(i => `${X(i).toFixed(0)},${Y(vals[i]).toFixed(0)}`).join(" ");
      if (run.length > 1) {
        s += `<polygon points="${X(run[0]).toFixed(0)},${B} ${pts} ${X(run[run.length - 1]).toFixed(0)},${B}" fill="${color}" fill-opacity=".1"/>`;
        s += `<polyline points="${pts}" fill="none" stroke="${color}" stroke-width="4" stroke-linejoin="round"/>`;
      }
      run.forEach((i, k) => s += `<circle cx="${X(i).toFixed(0)}" cy="${Y(vals[i]).toFixed(0)}" r="${(k === 0 || k === run.length - 1) ? 5 : 3.5}" fill="${color}"/>`);
    });
  }
  (opts.pending || []).forEach(i => s += `<rect x="${(X(i) - 6).toFixed(0)}" y="${T}" width="12" height="${B - T}" rx="6" fill="none" stroke="rgba(255,255,255,.3)" stroke-dasharray="3 3"/>`);
  if (labels) (opts.letters || []).forEach((l, i) => s += `<text x="${X(i).toFixed(0)}" y="${H - 5}" fill="#fff" font-size="12" text-anchor="middle">${l}</text>`);
  return s + "</svg>";
}

async function buildWeeklySummary() {
  const weEl = document.getElementById("weekEndDate");
  if (!weEl) return;
  if (!weEl.value) weEl.value = coatingTodayISO_();
  const esc = coaterBayEsc_;
  const todayIso = coatingTodayISO_();
  const todayApi = weeklyIsoToApi_(todayIso);

  // 7 days ending on the selected date
  const [ey, em, ed] = weEl.value.split("-").map(Number);
  const days = [];
  for (let i = 6; i >= 0; i--) days.push(new Date(ey, em - 1, ed - i));
  const apiDates = days.map(d => `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`);
  const fmtShort = d => d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
  ovSet_("weekRangeLabel", `${fmtShort(days[0])} \u2013 ${days[6].toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`);
  const nextBtn = document.getElementById("weekNextBtn");
  if (nextBtn) nextBtn.disabled = weEl.value >= todayIso;

  ovSet_("weekVerdictText", "Loading the week\u2026");

  // Fetch missing days (today = the live data already on the page)
  await Promise.all(apiDates.map(async api => {
    if (api === todayApi && currentDate === null && dashboardData) { weeklyCache[api] = dashboardData; return; }
    if (weeklyCache[api] !== undefined && api !== todayApi) return;
    if (weeklyIsoToApi_(todayIso) && weeklyApiToIso_(api) > todayIso) { weeklyCache[api] = null; return; }
    try {
      const r = await coatingFetchRetry_(`${API_URL}?mode=processed&date=${encodeURIComponent(api)}`, `Weekly ${api}`);
      weeklyCache[api] = r.ok ? await r.json() : null;
    } catch (e) { weeklyCache[api] = null; }
  }));

  // ── Per-day numbers ──
  const D = days.map((dt, i) => {
    const api = apiDates[i], data = weeklyCache[api];
    const s = (data && data.summary) || null;
    const live = api === todayApi;
    const jobs = s ? Number(s.totalJobs || 0) : 0;
    const lenses = s ? Number(s.totalLenses || jobs * 2) : 0;
    const broken = s ? Number(s.totalBreakLenses || 0) : 0;
    const fh = (s && s.flowHealth) || {};
    const fTot = ["healthy", "watch", "delayed", "overnight"].reduce((t, k) => t + Number(fh[k] || 0), 0);
    const slow = Number(fh.watch || 0) + Number(fh.delayed || 0) + Number(fh.overnight || 0);
    const same = s && s.aging && lenses ? (Number(s.aging.sameDay || 0) / lenses) * 100 : null;
    return {
      dt, api, data, live, has: !!(s && jobs > 0),
      complete: !!(s && jobs > 0) && !live,
      dayName: dt.toLocaleDateString("en-US", { weekday: "short" }),
      letter: dt.toLocaleDateString("en-US", { weekday: "narrow" }),
      label: fmtShort(dt),
      jobs, lenses, broken,
      rate: lenses ? (broken / lenses) * 100 : null,
      same, flow: s ? Number(s.avgDetaper || 0) : null,
      slowPct: fTot ? (slow / fTot) * 100 : null,
      reasons: (data && data.topReasons) || {},
      hourly: (data && Array.isArray(data.hourly)) ? data.hourly : [],
      mt: (data && data.machineTotals) || {},
    };
  });
  const comp = D.filter(d => d.complete);
  const withData = D.filter(d => d.has);
  const wJobs = withData.reduce((t, d) => t + d.jobs, 0);
  const wLenses = withData.reduce((t, d) => t + d.lenses, 0);
  const wBroken = withData.reduce((t, d) => t + d.broken, 0);
  const pendingIdx = D.map((d, i) => d.live ? i : -1).filter(i => i >= 0);

  // Reasons across the week
  const rTot = {};
  withData.forEach(d => Object.entries(d.reasons).forEach(([r, x]) => { rTot[r] = (rTot[r] || 0) + Number((x && x.total) || 0); }));
  const reasons = Object.keys(rTot).filter(r => rTot[r] > 0).sort((a, b) => rTot[b] - rTot[a]);
  const topReason = reasons[0] || null;
  const shareOf = (d, r) => d.broken ? (Number((d.reasons[r] && d.reasons[r].total) || 0) / d.broken) * 100 : null;

  // ── Verdict ──
  const first = comp[0], last = comp[comp.length - 1];
  let streak = 0, best = 0, streakStart = null, streakEnd = null, runStart = null;
  for (let i = 1; i < comp.length; i++) {
    if (comp[i].rate > comp[i - 1].rate) { if (!streak) runStart = comp[i - 1]; streak++; if (streak > best) { best = streak; streakStart = runStart; streakEnd = comp[i]; } }
    else streak = 0;
  }
  const flowRise = first && last && first.flow ? ((last.flow - first.flow) / first.flow) * 100 : 0;
  const topShare = topReason && wBroken ? (rTot[topReason] / wBroken) * 100 : 0;
  let vClass = "ok", vText;
  const bestDay = comp.length ? comp.reduce((a, b) => (b.rate < a.rate ? b : a)) : null;
  const worstDay = comp.length ? comp.reduce((a, b) => (b.rate > a.rate ? b : a)) : null;
  if (!withData.length) { vClass = "none"; vText = "No coating data for this week."; }
  else if (best >= WEEK_RULES.risingStreak) {
    vClass = "bad";
    const mult = streakStart.rate > 0 ? streakEnd.rate / streakStart.rate : null;
    vText = `<b>Breakage rose ${best} days in a row, ${streakStart.dayName} to ${streakEnd.dayName}</b> (${streakStart.rate.toFixed(2)}% \u2192 ${streakEnd.rate.toFixed(2)}%${mult && mult >= 1.5 ? `, about ${mult.toFixed(0)}\u00d7` : ""})`;
    if (flowRise >= WEEK_RULES.flowRisePct) vText += `, while flow time rose from ${first.flow} to ${last.flow} min`;
    vText += ".";
    if (topShare >= WEEK_RULES.reasonShare) vText += ` ${esc(topReason)} is ${topShare.toFixed(0)}% of all breakage.`;
  } else if (worstDay && worstDay.rate >= OV_BREAK_THRESHOLDS.bad) {
    vClass = "warn";
    vText = `<b>${worstDay.dayName} ${worstDay.label} was the problem day</b> at ${worstDay.rate.toFixed(2)}% breakage; the week overall was ${(wBroken / wLenses * 100).toFixed(2)}%.`;
  } else {
    vText = `<b>A steady week:</b> ${wJobs.toLocaleString()} jobs, ${(wLenses ? wBroken / wLenses * 100 : 0).toFixed(2)}% breakage${bestDay ? `, best day ${bestDay.dayName} (${bestDay.rate.toFixed(2)}%)` : ""}.`;
  }
  const vEl = document.getElementById("weekVerdict");
  if (vEl) vEl.className = "wk-verdict is-" + vClass;
  ovSet_("weekVerdictText", vText);

  // ── Coaters (week) ──
  const coat = {};
  withData.forEach(d => Object.entries(d.mt).forEach(([m, x]) => {
    const c = coat[m] || (coat[m] = { jobs: 0, broken: 0, worst: null });
    const b = Number((x && x.breakLenses) || 0);
    c.jobs += Number((x && x.jobs) || 0); c.broken += b;
    if (b > 0 && (!c.worst || b > c.worst.b)) c.worst = { b, day: d.dayName };
  }));
  const coaters = Object.keys(coat).filter(m => coat[m].jobs > 0 || coat[m].broken > 0).map(m => ({
    m, ...coat[m], rate: coat[m].jobs ? (coat[m].broken / (coat[m].jobs * 2)) * 100 : null,
  })).sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));

  // ── Heatmap (day × hour) ──
  const hourSet = new Set();
  withData.forEach(d => d.hourly.forEach(h => { const k = ovHour24_(h.hour); if (k !== null) hourSet.add(Math.floor(k)); }));
  for (let h = 7; h <= 17; h++) hourSet.add(h);
  const hours = [...hourSet].sort((a, b) => a - b);
  const cellVal = (d, h) => { const r = d.hourly.find(x => { const k = ovHour24_(x.hour); return k !== null && Math.floor(k) === h; }); return r ? Number(r.totalBroken || 0) : 0; };
  const cellTop = (d, h) => {
    const r = d.hourly.find(x => { const k = ovHour24_(x.hour); return k !== null && Math.floor(k) === h; });
    if (!r) return null; let top = null;
    Object.entries(r.machines || {}).forEach(([m, s]) => { const t = Number(s.total || 0); if (t > 0 && (!top || t > top.t)) top = { m, t }; });
    return top;
  };
  let hot = null;
  withData.forEach(d => hours.forEach(h => { const v = cellVal(d, h); if (v >= WEEK_RULES.hotHour && (!hot || v > hot.v)) hot = { d, h, v }; }));

  // ── Where to look first (max 3) ──
  const look = [];
  if (topReason && topShare >= WEEK_RULES.reasonShare) {
    const s0 = first ? shareOf(first, topReason) : null, s1 = last ? shareOf(last, topReason) : null;
    const grew = s0 !== null && s1 !== null && s1 - s0 >= WEEK_RULES.shareGrowPts;
    look.push(`<b>${esc(topReason)}</b>: ${rTot[topReason]} of ${wBroken} broken lenses${grew ? ", and its share grew during the week" : ""}.`);
  }
  if (hot) {
    const top = cellTop(hot.d, hot.h);
    look.push(`<b>${hot.d.dayName} ${ovHourName_(hot.h)}</b>: ${hot.v} lenses broken in one hour${top ? `, mostly ${esc(formatMachineLabel(top.m))}` : ""}.`);
  }
  if (first && last && (flowRise >= WEEK_RULES.flowRisePct || (last.slowPct || 0) > DAILY_RULES.slowShare))
    look.push(`<b>Flow</b>: average ${first.flow} \u2192 ${last.flow} min; slow jobs ${(first.slowPct || 0).toFixed(0)}% \u2192 ${(last.slowPct || 0).toFixed(0)}%.`);
  const hiCoater = coaters.find(c => c.jobs >= WEEK_RULES.smallSample && c.rate !== null && c.rate >= WEEK_RULES.coaterRateHigh);
  if (hiCoater) look.push(`<b>${esc(formatMachineLabel(hiCoater.m))}</b>: ${hiCoater.rate.toFixed(2)}% breakage rate this week (${hiCoater.broken} lenses).`);
  ovSet_("weekLook", look.length ? `<ol>${look.slice(0, 3).map(x => `<li>${x}</li>`).join("")}</ol>`
    : `<p class="wk-empty">Nothing stands out by the rules this week.</p>`);

  // ── Trend tiles (complete days only; today shown as a dotted slot) ──
  const series = (fn) => D.map(d => d.complete ? fn(d) : null);
  const letters = D.map(d => d.letter);
  const change = (vals, fmt, upIsBad) => {
    const real = vals.filter(v => v !== null);
    if (real.length < 2) return { txt: "not enough days yet", cls: "" };
    const a = real[0], b = real[real.length - 1];
    const up = b > a, same = Math.abs(b - a) < 1e-9;
    return { txt: `${same ? "" : up ? "\u25b2 " : "\u25bc "}${fmt(a)} \u2192 ${fmt(b)}`, cls: same ? "" : ((up === upIsBad) ? "wk-bad" : "wk-good") };
  };
  const pct = v => v.toFixed(2) + "%", pct0 = v => v.toFixed(0) + "%", min = v => Math.round(v) + " min", num = v => Math.round(v).toLocaleString();
  const lastLabel = last ? last.dayName : "";
  const tiles = [
    { n: "Breakage rate", v: series(d => d.rate), c: "#ff6b6b", f: pct, bad: true },
    { n: "Same-day breakage", v: series(d => d.same), c: "#ff6b6b", f: pct, bad: true },
    { n: "Average flow time", v: series(d => d.flow), c: "#fbbf24", f: min, bad: true },
    { n: "Slow jobs (over 15 min)", v: series(d => d.slowPct), c: "#fbbf24", f: pct0, bad: true },
    { n: "Jobs coated", v: series(d => d.jobs), c: "#3ddbc4", f: num, bad: false },
    { n: topReason ? `${topReason} share` : "Top reason share", v: series(d => topReason ? shareOf(d, topReason) : null), c: "#c4b5fd", f: pct0, bad: true },
  ];
  ovSet_("weekTiles", tiles.map(t => {
    const real = t.v.filter(v => v !== null), ch = change(t.v, t.f, t.bad);
    return `<div class="wk-tile"><div class="wk-th">${esc(t.n)}</div>
      <div class="wk-tv" style="color:${t.c}">${real.length ? t.f(real[real.length - 1]) : "--"} <span>${esc(lastLabel)}</span></div>
      <div class="wk-chg ${ch.cls}">${ch.txt}</div>
      ${weeklySpark_(t.v, t.c, { letters, pending: pendingIdx })}</div>`;
  }).join(""));

  // ── Day cards ──
  const rc = r => r < OV_BREAK_THRESHOLDS.watch ? "#4ade80" : r < OV_BREAK_THRESHOLDS.bad ? "#fbbf24" : "#ff6b6b";
  ovSet_("weekDays", D.map((d, i) => {
    if (!d.has) return `<div class="wk-dc is-empty"><div class="wk-dh"><b>${d.dayName}</b> ${d.label}</div><div class="wk-dr">--</div><div class="wk-dm">no data</div></div>`;
    const tag = d.live ? `<span class="wk-t live">IN PROGRESS</span>`
      : (comp.length > 1 && d === bestDay ? `<span class="wk-t best">BEST</span>` : (comp.length > 1 && d === worstDay ? `<span class="wk-t worst">WORST</span>` : ""));
    return `<button type="button" class="wk-dc" style="--c:${rc(d.rate)}" onclick="weeklyOpenDay_('${d.api}')" title="Open ${d.dayName} in the Daily Summary">
      <div class="wk-dh"><b>${d.dayName}</b> ${d.label} ${tag}</div><div class="wk-dr">${d.rate.toFixed(2)}%</div>
      <div class="wk-dm">${d.jobs.toLocaleString()} jobs \u00b7 ${d.broken} broken</div></button>`;
  }).join(""));

  // ── Reasons with daily lines ──
  const shown = reasons.slice(0, 3);
  const other = reasons.slice(3);
  const reasonRows = shown.map(r => ({ r, t: rTot[r], col: BREAKAGE_COLOR_MAP[r] || "#60a5fa", v: D.map(d => d.has ? Number((d.reasons[r] && d.reasons[r].total) || 0) : null) }));
  if (other.length) reasonRows.push({ r: `Other (${other.length} reason${other.length > 1 ? "s" : ""})`, t: other.reduce((s, r) => s + rTot[r], 0), col: "#e5e7eb",
    v: D.map(d => d.has ? other.reduce((s, r) => s + Number((d.reasons[r] && d.reasons[r].total) || 0), 0) : null) });
  ovSet_("weekReasons", reasonRows.length ? reasonRows.map(x => `
    <div class="wk-rr"><span>${esc(x.r)}</span><b>${x.t}</b><em>${(x.t / wBroken * 100).toFixed(0)}%</em>
    <div class="wk-mini">${weeklySpark_(x.v, x.col, { w: 220, h: 40, labels: false, pending: pendingIdx })}</div></div>`).join("")
    : `<p class="wk-empty">No broken lenses this week.</p>`);

  // ── Coaters by rate ──
  ovSet_("weekCoaters", coaters.length ? `<table class="dl-table"><thead><tr><th>Coater</th><th>Jobs</th><th>Broken</th><th>Rate</th><th>Worst day</th></tr></thead><tbody>
    ${coaters.map(c => {
      const small = c.jobs < WEEK_RULES.smallSample;
      const col = c.rate === null ? "#ffffff" : c.rate >= COATER_BAY_THRESHOLDS.high ? "#ff6b6b" : c.rate >= COATER_BAY_THRESHOLDS.watch ? "#fbbf24" : "#4ade80";
      return `<tr><td class="dl-m">${esc(formatMachineLabel(c.m))}</td><td>${c.jobs.toLocaleString()}</td>
        <td style="font-weight:800;color:${col}">${c.broken}</td>
        <td>${c.rate === null ? "--" : c.rate.toFixed(2) + "%"}${small ? " <small>(small sample)</small>" : ""}</td>
        <td>${c.worst ? `${c.worst.day} \u00b7 ${c.worst.b}` : "--"}</td></tr>`;
    }).join("")}</tbody></table>` : `<p class="wk-empty">No coater data.</p>`);

  // ── Heatmap ──
  const hMax = Math.max(1, ...withData.flatMap(d => hours.map(h => cellVal(d, h))));
  const nowHr = ovNowNY_().getHours();
  ovSet_("weekHeat", `<table class="wk-heat"><thead><tr><th></th>${hours.map(h => `<th>${ovShortHour_(h)}</th>`).join("")}<th>Total</th><th>Rate</th></tr></thead><tbody>
    ${D.map(d => {
      const rowHead = `<th class="wk-rh"><b>${d.dayName}</b> ${d.label}${d.live ? ' <em>live</em>' : ""}</th>`;
      if (!d.has) return `<tr>${rowHead}${hours.map(() => `<td class="wk-c is-fut"></td>`).join("")}<td class="wk-tot">--</td><td class="wk-tot">--</td></tr>`;
      return `<tr>${rowHead}${hours.map(h => {
        if (d.live && h > nowHr) return `<td class="wk-c is-fut"></td>`;
        const v = cellVal(d, h);
        if (!v) return `<td class="wk-c is-zero">0</td>`;
        const a = 0.12 + 0.88 * (v / hMax);
        return `<td class="wk-c" style="background:rgba(255,107,107,${a.toFixed(2)});${a > 0.55 ? "color:#061210;" : ""}" title="${d.dayName} ${ovHourName_(h)}: ${v} broken">${v}</td>`;
      }).join("")}<td class="wk-tot">${d.broken}</td><td class="wk-tot" style="color:${rc(d.rate)}">${d.rate.toFixed(2)}%</td></tr>`;
    }).join("")}</tbody></table>`);

  // ── Data check ──
  const heatTot = withData.reduce((t, d) => t + d.hourly.reduce((s, h) => s + Number(h.totalBroken || 0), 0), 0);
  const coatTot = coaters.reduce((t, c) => t + c.broken, 0);
  const notes = [];
  if (heatTot !== wBroken) notes.push(`hourly breakage adds up to ${heatTot}, day totals say ${wBroken}`);
  if (coatTot !== wBroken) notes.push(`coater breakage adds up to ${coatTot}, day totals say ${wBroken}`);
  const missing = D.filter(d => !d.has && !d.live && weeklyApiToIso_(d.api) <= todayIso).length;
  if (missing) notes.push(`${missing} day${missing > 1 ? "s" : ""} with no data`);
  const chk = document.getElementById("weekCheck");
  if (chk) {
    chk.className = "dl-check" + (notes.length ? " is-warn" : "");
    chk.textContent = notes.length ? "Check data: " + notes.join("; ") + "."
      : `Check passed: ${wJobs.toLocaleString()} jobs and ${wBroken} broken lenses add up across days, hours, and coaters.`;
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
  if (currentDate === null) {
    refreshUiStart_();
    loadDashboard().then(refreshUiDone_).catch(refreshUiFailed_);
    startRefreshCountdown();
  }
}, refreshInterval * 1000);

/* =====================================================
   AUTO-REFRESH INDICATOR
   Non-blocking: a top progress bar while fetching, then a short toast.
   "New data" only when numbers actually changed; changed values pulse.
   A failed refresh stays visible so nobody trusts stale numbers.
===================================================== */
let refreshLastSnap_ = null;
let refreshLastOkAt_ = null;
let refreshToastTimer_ = null;

function refreshSnap_() {
  const s = (dashboardProcessed && dashboardProcessed.summary) || {};
  const f = s.flowHealth || {};
  return {
    jobs: Number(s.totalJobs || 0), broken: Number(s.totalBreakLenses || 0),
    healthy: Number(f.healthy || 0), watch: Number(f.watch || 0),
    delayed: Number(f.delayed || 0), overnight: Number(f.overnight || 0)
  };
}
function refreshClock_(d) {
  return d.toLocaleTimeString("en-US", { timeZone: "America/New_York", hour: "numeric", minute: "2-digit" });
}
function refreshToast_(cls, html, ms) {
  const el = document.getElementById("refreshToast");
  if (!el) return;
  clearTimeout(refreshToastTimer_);
  el.className = "refresh-toast is-show " + cls;
  el.innerHTML = html;
  if (ms) refreshToastTimer_ = setTimeout(() => el.classList.remove("is-show"), ms);
}
function refreshUiStart_() {
  if (!refreshLastSnap_) refreshLastSnap_ = refreshSnap_();
  const bar = document.getElementById("refreshBar");
  if (bar) bar.classList.add("is-active");
  refreshToast_("is-busy", `<span class="rt-spin" aria-hidden="true"></span>Updating data\u2026`, 0);
}
function refreshUiDone_() {
  const bar = document.getElementById("refreshBar");
  if (bar) bar.classList.remove("is-active");
  const now = new Date(); refreshLastOkAt_ = now;
  const prev = refreshLastSnap_, cur = refreshSnap_();
  refreshLastSnap_ = cur;
  const diffs = [];
  const fmt = (n, label) => `${n > 0 ? "+" : ""}${n} ${label}`;
  if (prev) {
    if (cur.jobs    !== prev.jobs)    diffs.push(fmt(cur.jobs - prev.jobs, "jobs"));
    if (cur.broken  !== prev.broken)  diffs.push(fmt(cur.broken - prev.broken, "broken"));
    if (cur.delayed !== prev.delayed) diffs.push(fmt(cur.delayed - prev.delayed, "delayed"));
  }
  if (coatingMachineFailed_) {
    refreshToast_("is-fail", `<b>Partly updated</b> \u00b7 ${refreshClock_(now)} \u00b7 machine report didn't load`, 8000);
  } else if (diffs.length) {
    refreshToast_("is-new", `<b>New data</b> \u00b7 ${refreshClock_(now)} \u00b7 ${diffs.join(" \u00b7 ")}`, 6000);
    // pulse only the numbers that changed
    const map = { jobs: "ovCoated", broken: "ovBroken", healthy: "ovHealthy", watch: "ovWatchN", delayed: "ovDelayedN", overnight: "ovOvernightN" };
    Object.keys(map).forEach(k => {
      if (!prev || cur[k] === prev[k]) return;
      const el = document.getElementById(map[k]);
      if (!el) return;
      el.classList.remove("rt-changed"); void el.offsetWidth; el.classList.add("rt-changed");
    });
  } else {
    refreshToast_("is-same", `Up to date \u00b7 ${refreshClock_(now)} \u00b7 no new jobs`, 2500);
  }
}
function refreshUiFailed_() {
  const bar = document.getElementById("refreshBar");
  if (bar) bar.classList.remove("is-active");
  const since = refreshLastOkAt_ ? ` \u00b7 showing data from ${refreshClock_(refreshLastOkAt_)}` : "";
  refreshToast_("is-fail", `<b>Update failed</b>${since} \u00b7 retrying at next refresh`, 0);
}

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
          ticks : { color: "#ffffff", font: { family: CHART_MONO, size: 12 }, maxRotation: 0 },
          grid  : { color: "rgba(255,255,255,0.04)" },
          border: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: {
            color   : "#ffffff",
            font    : { family: CHART_MONO, size: 12 },
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
        html += `<td style="color:#ffffff">—</td>`;
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
    <div style="padding:8px 14px 6px;font-size:12px;color:var(--dim);">
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
  const prevId = ++coatingPerfState.previousDaySeq;
  const totalStartedAt = coatingPerfStart_(
    `Previous-day API load #${prevId}`,
    {}
  );

  try {
    const now  = new Date();
    const yest = new Date(now);
    yest.setDate(yest.getDate() - 1);
    const m = yest.getMonth() + 1, d = yest.getDate(), y = yest.getFullYear();
    const dateStr = `${m}/${d}/${y}`;

    const headersStartedAt = coatingPerfNow_();
    const res = await fetch(`${API_URL}?mode=processed&date=${encodeURIComponent(dateStr)}`);
    const headersMs = coatingPerfNow_() - headersStartedAt;

    coatingPerfLog_(
      `Previous-day API #${prevId} response headers: ${headersMs.toFixed(1)} ms`,
      { httpStatus: res.status, ok: res.ok, date: dateStr }
    );

    const bodyStartedAt = coatingPerfNow_();
    const responseText = await res.text();
    const bodyMs = coatingPerfNow_() - bodyStartedAt;

    const parseStartedAt = coatingPerfNow_();
    previousDayData = JSON.parse(responseText);
    const parseMs = coatingPerfNow_() - parseStartedAt;

    coatingPerfEnd_(`Previous-day API load #${prevId}`, totalStartedAt, {
      responseHeadersMs: Number(headersMs.toFixed(1)),
      bodyReadMs: Number(bodyMs.toFixed(1)),
      jsonParseMs: Number(parseMs.toFixed(1)),
      responseChars: responseText.length,
      success: true
    });
  } catch(e) {
    previousDayData = null;
    coatingPerfEnd_(`Previous-day API load #${prevId}`, totalStartedAt, {
      success: false,
      message: e?.message || String(e)
    });
  }
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
    .header-sub   { font-size: 13px; color: #ffffff; margin-top: 4px; letter-spacing: 1px; text-transform: uppercase; }
    .header-date  { font-size: 13px; color: #ffffff; text-align: right; }
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
    .kpi-l { font-size: 9px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: #ffffff; margin-bottom: 6px; }
    .kpi-v { font-size: 22px; font-weight: 700; color: #ffffff; }
    .chart-section { margin-bottom: 28px; page-break-inside: avoid; }
    .chart-label {
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 1.5px;
      text-transform: uppercase;
      color: #ffffff;
      margin-bottom: 8px;
      padding-left: 2px;
      border-left: 3px solid rgba(255,255,255,0.25);
      padding-left: 10px;
    }
    .footer {
      margin-top: 32px;
      padding-top: 14px;
      border-top: 1px solid #1e2535;
      font-size: 13px;
      color: #ffffff;
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

  ${chartSections || "<p style='color:#ffffff;font-size:13px;'>Open each chart tab first to capture charts in the export.</p>"}

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