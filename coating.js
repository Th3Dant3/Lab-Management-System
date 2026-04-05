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
  bodyColor       : "rgba(0, 255, 200, 0.65)",
  footerColor     : "rgba(0, 255, 200, 0.3)",
  titleFont       : { family: CHART_ORB,  size: 9, weight: "600" },
  bodyFont        : { family: CHART_MONO, size: 11 },
  padding         : 13,
  cornerRadius    : 4,
  displayColors   : true,
  boxWidth        : 8,
  boxHeight       : 8,
  boxPadding      : 3,
};

const GLASS_LEGEND = {
  labels: {
    color        : "rgba(0, 255, 200, 0.35)",
    font         : { family: CHART_ORB, size: 8 },
    usePointStyle: true,
    pointStyle   : "circle",
    padding      : 20,
    boxWidth     : 8,
    boxHeight    : 8,
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
  "S-HC Contamination"       : "#f87171",
  "S-HC Pit"                 : "#a78bfa",
  "S-HC Run"                 : "#fb923c",
  "S-HC Wagon Wheel"         : "#60a5fa",
  "S-HC Suction Cup Marks"   : "#2dd4bf",
  "S-HC HC Suction Cup Marks": "#2dd4bf",
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
    ticks : { color: tickColor, font: { family: CHART_MONO, size: 10 }, maxRotation: 0 },
    grid  : GLASS_GRID,
    border: { display: false },
    title : label
      ? { display: true, text: label, color: tickColor, font: { family: CHART_MONO, size: 10 }, padding: { bottom: 4 } }
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
  // Always rebuild these — they may have been destroyed on data reload
  if (tabId === "reasons") {
    if (reasonChart) { reasonChart.destroy(); reasonChart = null; }
    setTimeout(() => buildReasonChart(dashboardData), 80);
  }
  if (tabId === "machines") {
    if (machineChart) { machineChart.destroy(); machineChart = null; }
    setTimeout(() => buildMachineChart(dashboardMachine || dashboardData), 80);
  }
  if (tabId === "compare") {
    // Init the comparison picker if not already done
    const inp = document.getElementById("compareDateInputs");
    if (inp && inp.children.length === 0) compareInit();
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

  setText("totalJobs",   summary.totalJobs         ?? 0);
  setText("totalLenses", summary.totalLenses        ?? 0);
  setText("totalBroken", summary.totalBreakLenses   ?? 0);
  setText("rxBreakage",
    summary.breakPercent !== undefined
      ? `${Number(summary.breakPercent).toFixed(2)}% (${summary.totalBreakLenses || 0})`
      : "0%"
  );
  setText("avgTime",  summary.avgBreakTimeHours ? summary.avgBreakTimeHours + " hrs" : "-");
  setText("peakHour", summary.peakHour ?? "-");

  if (summary.flowHealth) {
    setText("flowHealthy",   summary.flowHealth.healthy   || 0);
    setText("flowWatch",     summary.flowHealth.watch     || 0);
    setText("flowDelayed",   summary.flowHealth.delayed   || 0);
    setText("flowOvernight", summary.flowHealth.overnight || 0);
  }
  if (summary.aging) {
    setText("sameDayBreakage",   summary.aging.sameDay || 0);
    setText("yesterdayBreakage", summary.aging.oneDay  || 0);
    setText("twoPlusBreakage",   summary.aging.twoPlus || 0);
  }

  // Update flow stat pills if they exist
  updateFlowStatPills(summary);

  buildHourlyTable(dashboardData.hourly || []);

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

/* =====================================================
   DATE FILTER
===================================================== */

function applyDateFilter() {
  const dateInput = document.getElementById("historyDate").value;
  if (!dateInput) return;
  const parts = dateInput.split("-");
  currentDate = parseInt(parts[1], 10) + "/" + parseInt(parts[2], 10) + "/" + parts[0];
  loadDashboard();
  startRefreshCountdown();
}

function resetToToday() {
  currentDate = null;
  const dateEl = document.getElementById("historyDate");
  if (dateEl) dateEl.value = "";
  loadDashboard();
  startRefreshCountdown();
}

document.addEventListener("DOMContentLoaded", () => {
  const liveBtn = document.getElementById("liveBtn");
  if (liveBtn) {
    liveBtn.addEventListener("click", () => {
      currentDate = null;
      const d = document.getElementById("historyDate");
      if (d) d.value = "";
      loadDashboard();
      startRefreshCountdown();
    });
  }
  startRefreshCountdown();
  compareInit();
});

loadDashboard();
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
      data               : totals,
      borderColor        : color,
      backgroundColor    : glassFill(ctx, color, chartH),
      borderWidth        : isTop ? 2 : 1.5,
      tension            : 0.42,
      fill               : isTop,
      pointRadius        : ctx => { const v = totals[ctx.dataIndex] || 0; return v === peak && v > 0 ? 5 : (v > 0 ? 2.5 : 0); },
      pointHoverRadius   : 7,
      pointBackgroundColor: ctx => { const v = totals[ctx.dataIndex] || 0; return v === peak && v > 0 ? "#fff" : color; },
      pointBorderColor   : color,
      pointBorderWidth   : 1.5,
      spanGaps           : false,
      yAxisID            : "yBroken",
      order              : 2,
      _showPeak          : isTop,
    });
    ci++;
  });

  // Coating jobs — amber dashed reference line
  datasets.push({
    label            : "Coating Jobs",
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
          ticks: { color: "rgba(248,113,113,0.6)", font: { family: CHART_MONO, size: 10 }, stepSize: 1 },
          grid: { ...GLASS_GRID, color: "rgba(248,113,113,0.05)" },
        },
        yJobs: {
          type: "linear", position: "right", beginAtZero: true,
          ...axisStyle("rgba(251,191,36,0.35)", "Jobs"),
          ticks: { color: "rgba(251,191,36,0.35)", font: { family: CHART_MONO, size: 10 } },
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

  function sanitize(v) {
    if (v === null || v === undefined) return null;
    const n = Number(v);
    return n > 500 ? null : n;
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

  // Avg flow time per bucket per hour
  const bucketDatasets = bucketDefs.map(b => {
    const isOvernight = b.label.includes("Overnight");
    const vals = filtered.map(h => {
      const pts = (h.flowPoints || []).filter(b.filter);
      if (!pts.length) return null;
      const avg = Math.round(pts.reduce((s, p) => s + p.flow, 0) / pts.length);
      // Cap overnight on main axis so it doesn't blow the scale
      return isOvernight ? Math.min(avg, yAxisMax) : avg;
    });
    const fill = ctx.createLinearGradient(0, 0, 0, chartH);
    fill.addColorStop(0,   b.color + "22");
    fill.addColorStop(1,   b.color + "00");
    return {
      label              : b.label,
      data               : vals,
      borderColor        : b.color,
      backgroundColor    : fill,
      borderWidth        : 2,
      tension            : 0.42,
      fill               : true,
      pointRadius        : vals.map(v => v === null ? 0 : 3),
      pointHoverRadius   : 7,
      pointBackgroundColor: b.color,
      pointBorderColor   : "rgba(5,8,16,0.7)",
      pointBorderWidth   : 1.5,
      spanGaps           : false,
      _bucketColor       : b.color,
      _isOvernight       : isOvernight,
    };
  });

  // Broken jobs avg flow — dashed violet line
  const brokenAvgVals = filtered.map(h => sanitize(h.avgFlowBroken));
  bucketDatasets.push({
    label              : "Broken Jobs Avg",
    data               : brokenAvgVals,
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
  const breakageDots = filtered.map((h, i) => {
    const count = h.totalBroken || 0;
    if (!count) return null;
    return { x: hours[i], y: sanitize(h.avgFlowBroken) || 10, count, hour: h.hour };
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
      interaction        : { mode: "index", intersect: false },

      plugins: {
        legend: {
          display  : true,
          position : "top",
          labels: {
            color        : "rgba(255,255,255,0.4)",
            font         : { family: CHART_FONT, size: 11, weight: "500" },
            usePointStyle: true,
            pointStyle   : "circle",
            padding      : 20,
            boxWidth     : 9,
            boxHeight    : 9,
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
          callbacks: {
            title: items => `  ${items[0].label}`,
            label: ctx => {
              const v = ctx.raw;
              if (v === null || v === undefined) return null;
              if (ctx.dataset.label === "Broken Jobs Avg") return `  Broken avg: ${v}m`;
              // Overnight values are capped for display — show label
              if (ctx.dataset._isOvernight && v >= yAxisMax) {
                return `  ${ctx.dataset.label}: >${Math.floor(yAxisMax/60)}h (capped on axis)`;
              }
              if (v >= 60) {
                const h = Math.floor(v/60), m = v%60;
                return `  ${ctx.dataset.label}: ${h}h${m>0?" "+m+"m":""}`;
              }
              return `  ${ctx.dataset.label}: ${v}m avg`;
            },
            afterBody: items => {
              const hour  = items[0]?.label;
              const brkPt = breakageDots.find(d => d.hour === hour);
              return brkPt ? ["", `  Breakage: ${brkPt.count} lenses this hour`] : [];
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
            color   : "rgba(140,175,220,0.45)",
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

    let primaryDriverHTML = "<span style='color:var(--text-dim)'>—</span>";
    let topAccessPointHTML = "<span style='color:var(--text-dim)'>—</span>";

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

      // Access point: list all machines that had breakage this hour
      const brkMachines = Object.entries(row.machines)
        .filter(([, s]) => (s.total || 0) > 0)
        .sort(([, a], [, b]) => (b.total || 0) - (a.total || 0))
        .map(([m, s]) => `<span style="color:#38bdf8;font-family:'JetBrains Mono',monospace;font-size:11px;">${formatMachineLabel(m)}</span><span style="color:var(--text-dim);font-size:11px;"> ×${s.total}</span>`)
        .join("<br>");
      if (brkMachines) topAccessPointHTML = brkMachines;
    }

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td class="${brkClass}">${broken > 0 ? broken : '<span style="color:var(--text-dim)">0</span>'}</td>
      <td>${row.coatingJobs || 0}</td>
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
      const c = BREAKAGE_COLOR_MAP[reason] || "#8aabcc";
      reasonHTML += `<div style="margin-top:8px;padding:8px 12px;background:rgba(255,255,255,0.02);border-left:2px solid ${c};border-radius:3px;font-size:12px;"><strong style="color:${c}">${reason}</strong> — <strong>${rStats.total || 0}</strong><br><span style="color:rgba(100,140,180,0.6);font-size:11px;">Same: ${rStats.sameDay||0} · Prev: ${rStats.oneDay||0} · 2+: ${rStats.twoPlus||0}</span></div>`;
    });
    machineHTML += `<div style="margin-bottom:12px;padding:14px;background:rgba(255,255,255,0.03);border:1px solid rgba(120,160,255,0.08);border-radius:8px;font-size:13px;line-height:1.9;font-family:${CHART_FONT};"><strong style="color:rgba(160,200,255,0.9);font-size:14px;">${dm}</strong><br><span style="color:rgba(100,140,180,0.6);font-size:11px;">JOBS: <strong style="color:#c8dff5;">${stats.jobs||0}</strong> &nbsp;·&nbsp; BROKEN: <strong style="color:${GC.red};">${stats.total||0}</strong></span><br><span style="color:rgba(100,140,180,0.5);font-size:11px;">Same: ${stats.sameDay||0} · Prev: ${stats.oneDay||0} · 2+: ${stats.twoPlus||0}</span>${reasonHTML}</div>`;
  });

  modalBody.innerHTML = `<h2>${hourData.hour}</h2><div style="color:rgba(140,175,220,0.7);font-size:13px;margin-bottom:14px;">Total Lenses Broken: <strong style="color:${GC.red};font-size:18px;">${hourData.totalBroken||0}</strong></div>${machineHTML || "<p style='color:rgba(100,140,180,0.4)'>No data</p>"}`;
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

  const bgs = colors.map(c => {
    const W = canvas.clientWidth * 0.78 || 700;
    const g = ctx.createLinearGradient(0, 0, W, 0);
    g.addColorStop(0,   c + "ff");
    g.addColorStop(0.7, c + "dd");
    g.addColorStop(1,   c + "88");
    return g;
  });

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
          ticks: { color: "rgba(140,175,220,0.3)", font: { family: CHART_MONO, size: 10 }, stepSize: Math.max(1, Math.ceil(maxVal / 8)) },
          grid  : GLASS_GRID,
          border: { display: false },
        },
        y: {
          ticks : { color: "rgba(160,200,240,0.65)", font: { family: CHART_FONT, size: 12, weight: "500" }, padding: 10 },
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
          ...GLASS_LEGEND,
          labels: {
            ...GLASS_LEGEND.labels,
            generateLabels: () => [
              { text: "Breakage %", fillStyle: "#ef4444", strokeStyle: "#ef4444", pointStyle: "circle", hidden: false, datasetIndex: 0 },
              { text: "Total Jobs", fillStyle: "#3b82f6", strokeStyle: "#3b82f6", pointStyle: "circle", hidden: false, datasetIndex: 1 },
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
          title : { display: true, text: "Breakage %", color: "rgba(248,113,113,0.5)", font: { family: CHART_MONO, size: 10 } },
        },
        y2: {
          position   : "right",
          beginAtZero: true,
          suggestedMax: Math.ceil(maxJobs * 1.25),
          ticks: { color: "rgba(96,165,250,0.4)", font: { family: CHART_MONO, size: 10 } },
          grid  : { drawOnChartArea: false },
          border: { display: false },
          title : { display: true, text: "Jobs", color: "rgba(96,165,250,0.35)", font: { family: CHART_MONO, size: 10 } },
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
            color        : "rgba(255,255,255,0.4)",
            font         : { family: CHART_FONT, size: 11 },
            usePointStyle: true,
            pointStyle   : "circle",
            padding      : 18,
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
          ticks : { color: "rgba(255,255,255,0.3)", font: { family: CHART_MONO, size: 10 }, maxRotation: 0 },
          grid  : { color: "rgba(255,255,255,0.04)" },
          border: { display: false },
        },
        y: {
          beginAtZero: true,
          ticks: {
            color   : "rgba(255,255,255,0.3)",
            font    : { family: CHART_MONO, size: 10 },
            callback: v => compareMetric === "pct" ? v + "%" : compareMetric === "flow" ? v + "m" : v,
          },
          grid  : { color: "rgba(255,255,255,0.04)" },
          border: { display: false },
          title : {
            display: true,
            text   : metricLabels[compareMetric],
            color  : "rgba(255,255,255,0.25)",
            font   : { family: CHART_MONO, size: 10 },
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
  window.print();
}