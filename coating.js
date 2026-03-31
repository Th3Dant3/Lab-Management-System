const API_URL = "https://script.google.com/macros/s/AKfycbxGEYOoJviGPzBSIX_Kh5X5ZAOAHFSyw9AO6ID-buDr1A82ERaPsauF3cSvSINr8VvT/exec";

console.log("Coating JS loaded");

/* =====================================================
   GLOBAL VARIABLES
===================================================== */

let trendChart    = null;
let reasonChart   = null;
let machineChart  = null;
let flowChart     = null;

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
   SHARED CHART CONSTANTS
===================================================== */

const CHART_FONT = "'DM Sans', 'Segoe UI', sans-serif";

const BASE_TOOLTIP = {
  backgroundColor : "rgba(6, 11, 20, 0.96)",
  borderColor     : "rgba(0, 212, 255, 0.3)",
  borderWidth     : 1,
  titleColor      : "#00d4ff",
  bodyColor       : "#c8dff5",
  footerColor     : "#5a7a9a",
  titleFont       : { family: CHART_FONT, size: 13, weight: "600" },
  bodyFont        : { family: CHART_FONT, size: 12 },
  padding         : 14,
  cornerRadius    : 8,
  displayColors   : true,
  boxWidth        : 10,
  boxHeight       : 10,
  boxPadding      : 4,
};

const BASE_LEGEND = {
  labels: {
    color        : "#8aabcc",
    font         : { family: CHART_FONT, size: 12 },
    usePointStyle: true,
    pointStyle   : "circle",
    padding      : 20,
    boxWidth     : 8,
    boxHeight    : 8,
  },
  position: "top",
};

const BASE_GRID   = { color: "rgba(255,255,255,0.04)" };
const BASE_BORDER = { color: "rgba(255,255,255,0.06)" };

/* =====================================================
   GLOBAL COLOR MAP
===================================================== */

const BREAKAGE_COLOR_MAP = {
  "S-HC Contamination"       : "#FF4D4D",
  "S-HC Pit"                 : "#9b7bff",
  "S-HC Run"                 : "#ff9f43",
  "S-HC Wagon Wheel"         : "#4da3ff",
  "S-HC Suction Cup Marks"   : "#00d4ff",
  "S-HC HC Suction Cup Marks": "#00d4ff",
};

/* =====================================================
   CHART HELPERS
===================================================== */

function formatMachineLabel(machine) {
  if (!machine) return "";
  return machine.replace(/^AR41-/, "44R1-");
}

function slimGradient(ctx, color, chartHeight) {
  const h = chartHeight || 360;
  const g = ctx.createLinearGradient(0, 0, 0, h);
  g.addColorStop(0,   color + "28");
  g.addColorStop(0.6, color + "0a");
  g.addColorStop(1,   color + "00");
  return g;
}

function axisStyle(color, label) {
  return {
    ticks : { color, font: { family: CHART_FONT, size: 11 }, maxRotation: 0 },
    grid  : BASE_GRID,
    border: BASE_BORDER,
    title : label
      ? { display: true, text: label, color, font: { family: CHART_FONT, size: 11 }, padding: { bottom: 6 } }
      : { display: false },
  };
}

function getFlowColor(minutes) {
  if (minutes <= 15) return "#00ff88";
  if (minutes <= 30) return "#ffcc00";
  return "#ff3b3b";
}

/* =====================================================
   CHART PLUGINS
===================================================== */

const GLOW_PLUGIN = {
  id: "glowV2",
  beforeDatasetDraw(chart, args) {
    const ds = chart.data.datasets[args.index];
    if (!ds || ds.type === "bar") return;
    const color = typeof ds.borderColor === "string" ? ds.borderColor : null;
    if (!color) return;
    chart.ctx.save();
    chart.ctx.shadowColor = color + "88";
    chart.ctx.shadowBlur  = 12;
  },
  afterDatasetDraw(chart) { chart.ctx.restore(); },
};

const VALUE_LABEL_PLUGIN = {
  id: "valueLabels",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden || meta.type !== "bar") return;
      meta.data.forEach((bar, index) => {
        const raw = dataset.data[index];
        if (raw === null || raw === undefined) return;
        const label = dataset._valueLabels
          ? (dataset._valueLabels[index] ?? "")
          : (typeof raw === "number" ? (Number.isInteger(raw) ? raw : raw.toFixed(1)) : raw);
        if (label === "" || label === null || label === undefined) return;
        ctx.save();
        ctx.font        = `600 11px ${CHART_FONT}`;
        ctx.fillStyle   = "#e8f4ff";
        ctx.shadowColor = "rgba(0,0,0,0.5)";
        ctx.shadowBlur  = 4;
        if (chart.options.indexAxis === "y") {
          ctx.textAlign    = "left";
          ctx.textBaseline = "middle";
          ctx.fillText(label, bar.x + 6, bar.y);
        } else {
          ctx.textAlign    = "center";
          ctx.textBaseline = "bottom";
          ctx.fillText(label, bar.x, bar.y - 6);
        }
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

/* =====================================================
   REPORT DATE DISPLAY
===================================================== */

function updateReportDateDisplay(data) {
  const el = document.getElementById("activeReportDate");
  if (!el) return;
  el.innerHTML = currentDate === null
    ? "Viewing Report Date: <strong style='color:#32ff7e'>LIVE</strong>"
    : "Viewing Report Date: <strong>" + currentDate + "</strong>";
}

/* =====================================================
   SAFE TEXT SETTER
===================================================== */

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
  if (tabId === "reasons"  && !reasonChart)  buildReasonChart(dashboardData);
  if (tabId === "machines" && !machineChart) buildMachineChart(dashboardMachine || dashboardData);
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
    setText("flowHealthy",   summary.flowHealth.healthy  || 0);
    setText("flowWatch",     summary.flowHealth.watch    || 0);
    setText("flowDelayed",   summary.flowHealth.delayed  || 0);
    setText("flowOvernight", summary.flowHealth.overnight|| 0);
  }
  if (summary.aging) {
    setText("sameDayBreakage",   summary.aging.sameDay || 0);
    setText("yesterdayBreakage", summary.aging.oneDay  || 0);
    setText("twoPlusBreakage",   summary.aging.twoPlus || 0);
  }

  buildHourlyTable(dashboardData.hourly || []);

  if (trendChart)   trendChart.destroy();
  if (reasonChart)  reasonChart.destroy();
  if (machineChart) machineChart.destroy();
  if (flowChart)    flowChart.destroy();
  trendChart = reasonChart = machineChart = flowChart = null;

  const activeTab = document.querySelector(".tab.active");
  if (activeTab) {
    const tabId = activeTab.getAttribute("onclick") || "";
    if (tabId.includes("trends"))   buildTrendCharts();
    else if (tabId.includes("reasons"))  buildReasonChart(dashboardData);
    else if (tabId.includes("machines")) buildMachineChart(dashboardMachine || dashboardData);
  }
  if (document.getElementById("flowChart")) buildFlowChart(dashboardProcessed);

  updateStatusBadge(currentDate === null ? "live" : "history");
  console.log("Load time (ms):", Math.round(performance.now() - startTime));
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
      const dateInput = document.getElementById("historyDate");
      if (dateInput) dateInput.value = "";
      loadDashboard();
      startRefreshCountdown();
    });
  }
  startRefreshCountdown();
});

loadDashboard();

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
  const activeBtn = document.getElementById(mode + "Btn");
  if (activeBtn) activeBtn.classList.add("active");
}

/* =====================================================
   TREND CHART
===================================================== */

function buildTrendChart(data) {
  const canvas = document.getElementById("trendChart");
  if (!canvas || !data || !data.hourly) return;
  const ctx = canvas.getContext("2d");
  if (trendChart) trendChart.destroy();

  const sortedHourly = [...data.hourly].sort((a, b) =>
    new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour)
  );
  const hours      = sortedHourly.map(h => h.hour);
  const datasets   = [];
  const allReasons = new Set();
  const chartH     = canvas.clientHeight || 420;

  sortedHourly.forEach(hour =>
    Object.values(hour.machines || {}).forEach(machine =>
      Object.keys(machine.reasons || {}).forEach(r => allReasons.add(r))
    )
  );

  const fallbackColors = ["#FF4D4D","#9b7bff","#ff9f43","#00d4ff","#4da3ff","#00ff88","#ff6bcb"];
  let ci = 0;

  allReasons.forEach(reason => {
    const color = BREAKAGE_COLOR_MAP[reason] || fallbackColors[ci % fallbackColors.length];
    const reasonTotals = sortedHourly.map(hour => {
      let total = 0;
      Object.values(hour.machines || {}).forEach(m => { total += m.reasons?.[reason]?.total || 0; });
      return total;
    });
    const peakValue = Math.max(...reasonTotals, 0);

    datasets.push({
      label            : reason,
      data             : reasonTotals,
      borderColor      : color,
      backgroundColor  : slimGradient(ctx, color, chartH),
      borderWidth      : 2.5,
      tension          : 0.42,
      fill             : true,
      pointRadius      : ctx => { const v = reasonTotals[ctx.dataIndex] || 0; return v === peakValue && v > 0 ? 6 : (v > 0 ? 3.5 : 2); },
      pointHoverRadius : 8,
      pointBackgroundColor: ctx => { const v = reasonTotals[ctx.dataIndex] || 0; return v === peakValue && v > 0 ? "#fff" : color; },
      pointBorderColor : color,
      pointBorderWidth : 1.5,
      spanGaps         : true,
      yAxisID          : "yBroken",
      order            : 2,
    });
    ci++;
  });

  datasets.push({
    label            : "Coating Jobs",
    data             : sortedHourly.map(h => h.coatingJobs || 0),
    borderColor      : "rgba(160, 200, 255, 0.4)",
    backgroundColor  : "transparent",
    borderDash       : [5, 4],
    borderWidth      : 1.5,
    tension          : 0.35,
    fill             : false,
    pointRadius      : 0,
    pointHoverRadius : 4,
    yAxisID          : "yJobs",
    order            : 1,
  });

  trendChart = new Chart(ctx, {
    type   : "line",
    data   : { labels: hours, datasets },
    options: {
      responsive         : true,
      maintainAspectRatio: false,
      interaction        : { mode: "index", intersect: false },
      animation          : { duration: 400, easing: "easeOutQuart" },
      plugins: {
        legend : BASE_LEGEND,
        tooltip: {
          ...BASE_TOOLTIP,
          callbacks: {
            title: items => `Hour: ${items[0].label}`,
            label: ctx => {
              if (ctx.dataset.label === "Coating Jobs") return `  Coating Jobs: ${ctx.raw}`;
              const v = ctx.raw;
              return `  ${ctx.dataset.label}: ${v} lens${v !== 1 ? "es" : ""}`;
            },
            afterBody: items => {
              const total = items.filter(i => i.dataset.label !== "Coating Jobs").reduce((s, i) => s + (i.raw || 0), 0);
              return total > 0 ? ["", `  Total broken this hour: ${total}`] : [];
            },
          },
        },
      },
      scales: {
        x      : { ...axisStyle("rgba(255,255,255,0.4)"), grid: { ...BASE_GRID, tickLength: 0 } },
        yBroken: { type: "linear", position: "left",  beginAtZero: true, ...axisStyle("#ff6b6b", "Breakage Count"), ticks: { color: "#ff6b6b", font: { family: CHART_FONT, size: 11 }, stepSize: 1 } },
        yJobs  : { type: "linear", position: "right", beginAtZero: true, ...axisStyle("rgba(160,200,255,0.35)", "Coating Jobs"), ticks: { color: "rgba(160,200,255,0.35)", font: { family: CHART_FONT, size: 11 } }, grid: { drawOnChartArea: false } },
      },
    },
    plugins: [GLOW_PLUGIN],
  });
}

/* =====================================================
   FLOW CHART SYSTEM
===================================================== */

function switchFlowMode(mode) {
  currentFlowMode = mode;
  document.querySelectorAll("#avgFlowBtn, #machineFlowBtn, #individualFlowBtn").forEach(btn => btn.classList.remove("active"));
  const btnMap = { average: "avgFlowBtn", machine: "machineFlowBtn", individual: "individualFlowBtn" };
  const activeBtn = document.getElementById(btnMap[mode]);
  if (activeBtn) activeBtn.classList.add("active");
  buildFlowChart(dashboardProcessed);
}

function buildFlowChart(data) {
  const canvas = document.getElementById("flowChart");
  if (!canvas || !data || !data.hourly) return;
  const ctx = canvas.getContext("2d");
  if (flowChart) flowChart.destroy();

  const sortedHourly = [...data.hourly].sort((a, b) =>
    new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour)
  );
  const filteredHours = sortedHourly.filter(h => {
    if (!h.hour) return false;
    let hourNum = parseInt(h.hour.split(":")[0], 10);
    const isPM = h.hour.includes("PM"), isAM = h.hour.includes("AM");
    if (isPM && hourNum !== 12) hourNum += 12;
    if (isAM && hourNum === 12) hourNum = 0;
    return hourNum >= 6 && hourNum <= 20;
  });
  const hours  = filteredHours.map(h => h.hour);
  const chartH = canvas.clientHeight || 420;

  function sanitize(value) {
    if (value === null || value === undefined) return null;
    const num = Number(value);
    return num > 500 ? null : num;
  }

  // MACHINE MODE
  if (currentFlowMode === "machine") {
    const machineSet = new Set();
    filteredHours.forEach(h => Object.keys(h.machines || {}).forEach(m => machineSet.add(m)));
    const machinePalette = ["#4da3ff","#ff4d4d","#ff9f43","#ffd32a","#32ff7e","#9b7bff","#00d4ff","#ff6bcb"];
    const datasets = Array.from(machineSet).map((machine, idx) => {
      const color = machinePalette[idx % machinePalette.length];
      return {
        label: formatMachineLabel(machine),
        data : filteredHours.map(h => ({ x: h.hour, y: sanitize(h.machines?.[machine]?.avgFlowAll), count: h.machines?.[machine]?.flowAllCount || 0 })),
        borderColor: color, backgroundColor: "transparent", borderWidth: 2.5, tension: 0.35,
        fill: false, pointRadius: 4, pointHoverRadius: 7, pointBackgroundColor: color, spanGaps: true,
      };
    });
    flowChart = new Chart(ctx, { type: "line", data: { labels: hours, datasets }, options: getFlowOptions("Machine Average Flow (Detaper → Coater)"), plugins: [GLOW_PLUGIN] });
    return;
  }

  // AVERAGE MODE
  if (currentFlowMode === "average") {
    const buckets = [
      { label: "Healthy  (≤15 min)",    color: "#00ff88", filter: p => p.flow <= 15 },
      { label: "Watch  (16–30 min)",    color: "#ffd32a", filter: p => p.flow > 15 && p.flow <= 30 },
      { label: "Delayed  (31–360 min)", color: "#ff4d4d", filter: p => p.flow > 30 && p.flow <= 360 },
      { label: "Overnight  (>6 hrs)",   color: "#00a8ff", filter: p => p.flow > 360 },
    ];
    const datasets = buckets.map(b => ({
      label: b.label, borderColor: b.color, backgroundColor: slimGradient(ctx, b.color, chartH),
      borderWidth: 2.5, tension: 0.4, fill: true, pointRadius: 3.5, pointHoverRadius: 7,
      pointBackgroundColor: b.color, spanGaps: true,
      data: filteredHours.map(h => {
        const pts = (h.flowPoints || []).filter(b.filter);
        return { x: h.hour, y: pts.length ? Math.round(pts.reduce((a, p) => a + p.flow, 0) / pts.length) : 0, jobs: pts.length };
      }),
    }));
    datasets.push({
      label: "Broken Jobs Avg", borderColor: "#9b7bff", backgroundColor: "transparent",
      borderWidth: 2, borderDash: [5, 4], tension: 0.4, fill: false, pointRadius: 3,
      pointHoverRadius: 6, pointBackgroundColor: "#9b7bff", spanGaps: true,
      data: filteredHours.map(h => ({ x: h.hour, y: sanitize(h.avgFlowBroken), count: h.flowBrokenCount || 0 })),
    });
    flowChart = new Chart(ctx, { type: "line", data: { labels: hours, datasets }, options: getFlowOptions("Detaper → Coater Flow Analysis"), plugins: [GLOW_PLUGIN] });
    return;
  }

  // INDIVIDUAL SCATTER
  const points = [];
  filteredHours.forEach(hourObj =>
    Object.entries(hourObj.machines || {}).forEach(([machine, mData]) => {
      const fv = sanitize(mData.avgFlowAll);
      if (fv === null) return;
      points.push({ x: hourObj.hour, y: fv, machine, rx: null });
    })
  );
  flowChart = new Chart(ctx, {
    type: "scatter",
    data: { datasets: [{ label: "Machine Flow", data: points, pointRadius: 9, pointHoverRadius: 12, backgroundColor: points.map(p => getFlowColor(p.y) + "cc"), borderColor: points.map(p => getFlowColor(p.y)), borderWidth: 1.5 }] },
    options: getFlowOptions("Machine Flow Scatter (Detaper → Coater)"),
    plugins: [GLOW_PLUGIN],
  });
}

function getFlowOptions(titleText) {
  return {
    responsive: true, maintainAspectRatio: false,
    animation : { duration: 400, easing: "easeOutQuart" },
    interaction: { mode: "nearest", intersect: true },
    onClick(evt, elements, chart) {
      if (!elements.length) return;
      const point = chart.data.datasets[elements[0].datasetIndex].data[elements[0].index];
      if (!point) return;
      showFlowDetails({ rx: point.rx || "Machine Flow", machine: point.machine || "Unknown", reason: point.reason || "N/A", flow: point.y, x: point.x });
    },
    plugins: {
      legend : BASE_LEGEND,
      tooltip: {
        ...BASE_TOOLTIP,
        callbacks: {
          title: items => items[0].label || items[0].raw?.x || "",
          label: ctx => {
            const raw = ctx.raw, lines = [];
            if (raw?.machine) lines.push(`  Machine: ${raw.machine}`);
            if (raw?.rx)      lines.push(`  RX: ${raw.rx}`);
            const y = raw?.y ?? raw;
            if (typeof y === "number") { const hrs = Math.floor(y/60), mins = Math.round(y%60); lines.push(`  Flow Time: ${hrs > 0 ? hrs+"h " : ""}${mins}m`); }
            if (raw?.jobs  != null) lines.push(`  Jobs: ${raw.jobs}`);
            if (raw?.count != null && raw.count > 0) lines.push(`  Count: ${raw.count}`);
            if (!lines.length) lines.push(`  ${ctx.dataset.label}: ${typeof raw === "number" ? raw : raw?.y} min`);
            return lines;
          },
        },
      },
    },
    scales: {
      x: { type: "category", ...axisStyle("rgba(255,255,255,0.4)") },
      y: { beginAtZero: true, ...axisStyle("#4d9fff", "Minutes"), ticks: { color: "#4d9fff", font: { family: CHART_FONT, size: 11 }, callback: v => v >= 60 ? (v/60).toFixed(1)+"h" : v+"m" } },
    },
  };
}

/* =====================================================
   FLOW MODAL DETAILS
===================================================== */

function showFlowDetails(point) {
  const modal = document.getElementById("chartModal"), modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;
  const tm = Number(point.flow) || 0, hrs = Math.floor(tm/60), mins = Math.round(tm%60);
  const formatted = hrs > 0 ? (mins > 0 ? `${hrs}h ${mins}m` : `${hrs}h`) : `${mins}m`;
  modalBody.innerHTML = `
    <h2>RX ${point.rx}</h2>
    <div style="margin-top:15px;padding:20px;background:#0f1a2e;border:1px solid rgba(0,212,255,0.15);border-radius:10px;line-height:2;">
      <p style="margin:0"><strong>Machine:</strong> ${point.machine}</p>
      <p style="margin:0"><strong>Breakage Reason:</strong> ${point.reason || "None"}</p>
      <p style="margin:0"><strong>Flow Time:</strong> ${formatted}</p>
      <p style="margin:0"><strong>Hour:</strong> ${point.x}</p>
    </div>`;
  modal.classList.add("active");
}

function showFlowHourDetails(hourData) {
  const modal = document.getElementById("chartModal"), modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;
  let html = "";
  (hourData.flowPoints || []).forEach(p => {
    const color = getFlowColor(p.flow);
    html += `<div style="margin-bottom:10px;padding:12px;background:#0f1a2e;border-left:4px solid ${color};border-radius:6px;font-size:13px;line-height:1.8;"><strong>RX ${p.rx}</strong><br>Machine: ${p.machine}<br>Reason: ${p.reason}<br>Flow: ${p.flow}m</div>`;
  });
  modalBody.innerHTML = `<h2>${hourData.hour}</h2>${html || "<p style='color:#5a7a9a'>No flow data</p>"}`;
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
    const tr = document.createElement("tr");
    const broken = row.totalBroken || 0;
    let brkClass = "brk-zero";
    if (broken >= 10) brkClass = "brk-high";
    else if (broken >= 4) brkClass = "brk-mid";
    else if (broken > 0)  brkClass = "brk-low";

    let primaryDriverHTML = "<span style='color:var(--text-dim)'>—</span>";
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
    }

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td class="${brkClass}">${broken > 0 ? broken : '<span style="color:var(--text-dim)">0</span>'}</td>
      <td>${row.coatingJobs || 0}</td>
      <td>${primaryDriverHTML}</td>`;
    tr.addEventListener("click", () => showHourDetails(row));
    tbody.appendChild(tr);
  });
}

/* =====================================================
   HOUR DETAIL MODAL
===================================================== */

function showHourDetails(hourData) {
  const modal = document.getElementById("chartModal"), modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody || !hourData) return;
  let machineHTML = "";

  Object.entries(hourData.machines || {}).forEach(([machine, stats]) => {
    const displayMachine = formatMachineLabel(machine);
    let reasonHTML = "";
    Object.entries(stats.reasons || {}).forEach(([reason, rStats]) => {
      reasonHTML += `<div style="margin-top:8px;padding-left:12px;font-size:13px;">• <strong>${reason}</strong> — ${rStats.total||0}<br><span style="font-size:11px;color:#5a7a9a;">Same Day: ${rStats.sameDay||0} &nbsp;|&nbsp; Prev Day: ${rStats.oneDay||0} &nbsp;|&nbsp; 2+ Days: ${rStats.twoPlus||0}</span></div>`;
    });
    machineHTML += `<div style="margin-bottom:14px;padding:14px;background:#0f1a2e;border:1px solid rgba(0,212,255,0.1);border-left:3px solid #4d9fff;border-radius:8px;font-size:13px;line-height:1.9;"><strong style="color:#c8dff5;font-size:14px;">${displayMachine}</strong><br>Jobs Ran: ${stats.jobs||0}<br>Total Brkg: ${stats.total||0}<br>Same Day: ${stats.sameDay||0} &nbsp;|&nbsp; Prev Day: ${stats.oneDay||0} &nbsp;|&nbsp; 2+ Days: ${stats.twoPlus||0}${reasonHTML}</div>`;
  });

  modalBody.innerHTML = `<h2>${hourData.hour}</h2><h3 style="color:#8aabcc;font-weight:500;margin-top:0;">Total Lenses Broken: ${hourData.totalBroken||0}</h3>${machineHTML || "<p style='color:#5a7a9a'>No data</p>"}`;
  modal.classList.add("active");
}

function closeModal() {
  const modal = document.getElementById("chartModal");
  if (modal) modal.classList.remove("active");
}

/* =====================================================
   REASON CHART
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

  const maxVal = entries[0].total;
  const grandTotal = entries.reduce((s, e) => s + e.total, 0);
  const colors = entries.map((e, i) => BREAKAGE_COLOR_MAP[e.reason] || ["#ff4d4d","#ff9f43","#ffd32a","#00d4ff","#9b7bff","#4da3ff"][i % 6]);
  const valueLabels = entries.map(e => `${e.total}  (${((e.total/grandTotal)*100).toFixed(1)}%)`);

  reasonChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels  : entries.map(e => e.reason),
      datasets: [{
        label: "Breakage Count", data: entries.map(e => e.total), _valueLabels: valueLabels,
        backgroundColor: colors.map(c => { const g = ctx.createLinearGradient(0,0,canvas.clientWidth*0.85||900,0); g.addColorStop(0,c+"ff"); g.addColorStop(1,c+"55"); return g; }),
        borderColor: colors.map(c => c+"cc"), borderWidth: 1, borderRadius: 6, borderSkipped: false,
        barPercentage: 0.72, categoryPercentage: 0.88, hoverBorderColor: "#fff", hoverBorderWidth: 1.5,
      }],
    },
    options: {
      indexAxis: "y", responsive: true, maintainAspectRatio: false,
      animation: { duration: 500, easing: "easeOutQuart" },
      layout: { padding: { right: 120 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...BASE_TOOLTIP,
          callbacks: {
            title    : items => entries[items[0].dataIndex]?.reason,
            label    : ctx  => { const e=entries[ctx.dataIndex]; return [`  Breakage Count: ${e.total}`, `  Share of Total: ${((e.total/grandTotal)*100).toFixed(1)}%`]; },
            afterLabel: ctx => ctx.dataIndex === 0 ? `  ⚠️ Highest contributor` : "",
          },
        },
      },
      scales: {
        x: { beginAtZero: true, max: Math.ceil(maxVal*1.08), ...axisStyle("rgba(255,255,255,0.3)"), ticks: { color: "rgba(255,255,255,0.3)", font: { family: CHART_FONT, size: 11 }, stepSize: Math.max(1,Math.ceil(maxVal/8)) } },
        y: { ticks: { color: "#c8dff5", font: { family: CHART_FONT, size: 12, weight: "500" }, padding: 8 }, grid: { display: false }, border: { display: false } },
      },
    },
    plugins: [VALUE_LABEL_PLUGIN, GLOW_PLUGIN],
  });
}

/* =====================================================
   MACHINE CHART
===================================================== */

function buildMachineChart(data) {
  if (!data || !data.machineTotals) return;
  const canvas = document.getElementById("machineChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  if (machineChart) machineChart.destroy();

  const entries = Object.entries(data.machineTotals)
    .map(([machine, stats]) => {
      const jobs = Number(stats.jobs||0), broken = Number(stats.breakLenses||0), totalLenses = jobs*2;
      const percent = totalLenses > 0 ? (broken/totalLenses)*100 : (broken > 0 ? 100 : 0);
      return { machine, jobs, broken, percent };
    })
    .sort((a, b) => { if(a.jobs===0&&b.jobs===0) return b.broken-a.broken; if(a.jobs===0) return 1; if(b.jobs===0) return -1; return b.percent-a.percent; });
  if (!entries.length) return;

  const labels     = entries.map(e => formatMachineLabel(e.machine));
  const maxPercent = Math.max(...entries.map(e => e.percent), 0.1);
  const maxJobs    = Math.max(...entries.map(e => e.jobs), 1);
  const jobColor   = "rgba(77,159,255,0.55)";
  const jobBorder  = "rgba(77,159,255,0.9)";

  function machineColor(e) {
    if (e.jobs===0&&e.broken>0) return "#ff6bcb";
    if (e.percent>=6) return "#ff4d4d";
    if (e.percent>=3) return "#ff9f43";
    if (e.percent>=1) return "#ffd32a";
    return "#32ff7e";
  }
  const brkColors = entries.map(machineColor);

  machineChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Breakage %", data: entries.map(e => parseFloat(e.percent.toFixed(2))),
          _valueLabels: entries.map(e => e.jobs===0&&e.broken>0 ? `${e.broken} brk` : `${e.percent.toFixed(1)}%`),
          backgroundColor: brkColors.map(c => { const g=ctx.createLinearGradient(0,0,0,300); g.addColorStop(0,c+"ee"); g.addColorStop(1,c+"55"); return g; }),
          borderColor: brkColors, borderWidth: 1.5, borderRadius: { topLeft:5, topRight:5 }, borderSkipped: false,
          barPercentage: 0.45, categoryPercentage: 0.8, yAxisID: "y1", order: 1,
        },
        {
          label: "Total Jobs", data: entries.map(e => e.jobs),
          _valueLabels: entries.map(e => e.jobs > 0 ? String(e.jobs) : ""),
          backgroundColor: jobColor, borderColor: jobBorder, borderWidth: 1,
          borderRadius: { topLeft:5, topRight:5 }, borderSkipped: false,
          barPercentage: 0.45, categoryPercentage: 0.8, yAxisID: "y2", order: 2,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      animation : { duration: 450, easing: "easeOutQuart" },
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          ...BASE_LEGEND,
          labels: { ...BASE_LEGEND.labels, generateLabels: () => [
            { text: "Breakage %", fillStyle: "#ff6b6b", strokeStyle: "#ff6b6b", pointStyle: "circle", hidden: false, datasetIndex: 0 },
            { text: "Total Jobs", fillStyle: jobBorder,  strokeStyle: jobBorder,  pointStyle: "circle", hidden: false, datasetIndex: 1 },
          ]},
        },
        tooltip: {
          ...BASE_TOOLTIP,
          callbacks: {
            title: items => `Machine: ${items[0].label}`,
            label: ctx => {
              const e = entries[ctx.dataIndex];
              if (ctx.datasetIndex === 0) {
                if (e.jobs===0&&e.broken>0) return [`  ⚠️ Breakage: ${e.broken} lenses — no job data`];
                return [`  Breakage %: ${e.percent.toFixed(2)}%`, `  Lenses broken: ${e.broken} / ${e.jobs*2}`];
              }
              return [`  Total Jobs: ${e.jobs}`];
            },
            afterBody: items => ["", `  Rank #${items[0].dataIndex+1} by breakage rate`],
          },
        },
      },
      scales: {
        x : { ...axisStyle("#c8dff5"), ticks: { color: "#c8dff5", font: { family: CHART_FONT, size: 12, weight: "600" }, padding: 6 }, grid: { display: false } },
        y1: { position: "left",  beginAtZero: true, suggestedMax: Math.ceil(maxPercent*1.35), ...axisStyle("#ff6b6b","Breakage %"), ticks: { color: "#ff6b6b", font: { family: CHART_FONT, size: 11 }, callback: v => v+"%" }, grid: BASE_GRID },
        y2: { position: "right", beginAtZero: true, suggestedMax: Math.ceil(maxJobs*1.2),    ...axisStyle(jobBorder,"Total Jobs"),  ticks: { color: "rgba(77,159,255,0.7)", font: { family: CHART_FONT, size: 11 } }, grid: { drawOnChartArea: false } },
      },
    },
    plugins: [VALUE_LABEL_PLUGIN, GLOW_PLUGIN],
  });
}

/* =====================================================
   AI INSIGHTS
===================================================== */

function buildInsights(data) {
  const el = document.getElementById("insightBox");
  if (!el || !data) return;
  const summary = data.summary||{}, hourly = data.hourly||[], machineTotals = data.machineTotals||{}, topReasons = data.topReasons||{};
  const insights = [];

  let peakHour="-", peakValue=0;
  hourly.forEach(h => { const val=h.totalBroken||0; if(val>peakValue){peakValue=val;peakHour=h.hour;} });
  if (peakValue > 0) insights.push({ cls:"insight-peak", text:`🔥 Peak: <b>${peakHour}</b> (${peakValue} lenses)` });

  let topReason="-", topReasonCount=0;
  Object.entries(topReasons).forEach(([reason,stats]) => { if((stats.total||0)>topReasonCount){topReason=reason;topReasonCount=stats.total||0;} });
  if (topReasonCount > 0) insights.push({ cls:"insight-warn", text:`⚠️ Top Issue: <b>${topReason}</b> (${topReasonCount})` });

  let worstMachine="-", worstPercent=0;
  Object.entries(machineTotals).forEach(([machine,stats]) => {
    const jobs=stats.jobs||0, broken=stats.breakLenses||0;
    if(jobs>0){const pct=(broken/(jobs*2))*100; if(pct>worstPercent){worstPercent=pct;worstMachine=machine;}}
  });
  if (worstPercent > 0) insights.push({ cls:"insight-machine", text:`🛠️ Worst Machine: <b>${formatMachineLabel(worstMachine)}</b> ${worstPercent.toFixed(1)}%` });

  if (summary.flowHealth) {
    const {delayed=0,healthy=0} = summary.flowHealth;
    if (delayed>healthy) insights.push({cls:"insight-warn", text:`🚨 Flow Risk: Delayed > Healthy`});
    else if (delayed>0)  insights.push({cls:"insight-flow", text:`⚡ Flow Warning: ${delayed} delayed`});
    else                 insights.push({cls:"insight-ok",   text:`✅ Flow Stable`});
  }

  if (hourly.length >= 3) {
    const last = hourly.slice(-3).map(h => h.totalBroken||0);
    if (last[2]>last[1]&&last[1]>last[0]) insights.push({cls:"insight-trend",text:`📈 Breakage Rising (last 3 hrs)`});
    else if (last[2]<last[1]&&last[1]<last[0]) insights.push({cls:"insight-trend",text:`📉 Breakage Falling (last 3 hrs)`});
  }

  if (summary.totalBreakLenses === 0) insights.push({cls:"insight-ok",text:`🟢 Zero Breakage`});

  el.innerHTML = insights.length
    ? insights.map(i => `<span class="insight-item ${i.cls}">${i.text}</span>`).join("")
    : "<span style='color:var(--text-dim);font-size:13px;'>No insights available</span>";
}

/* =====================================================
   REBUILD ALL CHARTS
===================================================== */

function rebuildAllCharts() {
  [trendChart,reasonChart,machineChart,flowChart].forEach(c => { if(c) c.destroy(); });
  trendChart = reasonChart = machineChart = flowChart = null;
  const activeTab = document.querySelector(".tab.active");
  if (!activeTab) return;
  const tabText = activeTab.textContent.toLowerCase();
  if (tabText.includes("trend"))        buildTrendCharts();
  else if (tabText.includes("reason"))  buildReasonChart(dashboardData);
  else if (tabText.includes("machine")) buildMachineChart(dashboardMachine||dashboardData);
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
    const m = Math.floor(refreshCountdown/60), s = refreshCountdown%60;
    timerEl.textContent = `Refresh in ${m}:${s.toString().padStart(2,"0")}`;
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