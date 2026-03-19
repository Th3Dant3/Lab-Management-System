const API_URL = "https://script.google.com/macros/s/AKfycbxGEYOoJviGPzBSIX_Kh5X5ZAOAHFSyw9AO6ID-buDr1A82ERaPsauF3cSvSINr8VvT/exec";

console.log("Coating JS loaded");
/* =====================================================
   DASHBOARD VARIABLES Global
   ===================================================== */

let trendChart = null;
let reasonChart = null;
let machineChart = null;
let flowChart = null;              

let dashboardProcessed = null;
let dashboardMachine = null;
let dashboardData = null;

let trendsLoaded = false;
let currentMode = "processed";
let currentFlowMode = "average";

let currentDate = null;
let refreshInterval = 5 * 60; // 5 minutes in seconds
let refreshCountdown = refreshInterval;
let refreshTimerHandle = null;

// =====================================================
// GLOBAL CHART STYLE HELPERS
// =====================================================

function createGradient(ctx, color) {
  const gradient = ctx.createLinearGradient(0, 0, 0, 300);
  gradient.addColorStop(0, color + "66");
  gradient.addColorStop(1, color + "00");
  return gradient;
}

const GLOW_PLUGIN = {
  id: "glowEffect",
  beforeDatasetsDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.shadowColor = "rgba(0, 212, 255, 0.35)";
    ctx.shadowBlur = 10;
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
  },
  afterDatasetsDraw(chart) {
    chart.ctx.restore();
  }
};

// =====================================================
// GLOBAL BREAKAGE COLOR MAP (FIXED COLORS)
// =====================================================

const BREAKAGE_COLOR_MAP = {
  "S-HC Contamination": "#FF4D4D",       // Red
  "S-HC Pit": "#9b7bff",                 // Purple
  "S-HC Run": "#ff9f43",                 // Orange
  "S-HC Wagon Wheel": "#4da3ff",         // Blue
  "S-HC Suction Cup Marks": "#00d4ff",   // Cyan
  "S-HC HC Suction Cup Marks": "#00d4ff",
  // Add future reasons here to lock colors permanently
};

function formatMachineLabel(machine) {
  if (!machine) return "";
  return machine.replace(/^AR41-/, "44R1-");
}

/* =====================================================
   REPORT DATE DISPLAY (RESTORED - SAFE)
===================================================== */
function updateReportDateDisplay(data) {
  const el = document.getElementById("activeReportDate");
  if (!el) return;

  if (currentDate === null) {
    el.innerHTML = "Viewing Report Date: <strong style='color:#32ff7e'>LIVE</strong>";
  } else {
    el.innerHTML = "Viewing Report Date: <strong>" + currentDate + "</strong>";
  }
}

/* =====================================================
   SAFE TEXT SETTER (RESTORED)
===================================================== */
function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;

  if (value === null || value === undefined || value === "") {
    el.textContent = "-";
    return;
  }

  el.textContent = value;
}

/* =====================================================
   TAB SWITCH
===================================================== */
function showTab(tabId, button) {
  document.querySelectorAll(".tab-content")
    .forEach(t => t.classList.remove("active"));

  document.querySelectorAll(".tab")
    .forEach(t => t.classList.remove("active"));

  const tabEl = document.getElementById(tabId);
  if (tabEl) tabEl.classList.add("active");
  if (button) button.classList.add("active");

  if (tabId === "trends") {
    setTimeout(() => {
      buildTrendCharts();

      // ENSURE FLOW ALWAYS LOADS
      if (dashboardProcessed) {
        buildFlowChart(dashboardProcessed);
      }
    }, 100);
  }

  if (tabId === "reasons" && !reasonChart) {
    buildReasonChart(dashboardData);
  }

  if (tabId === "machines" && !machineChart) {
    buildMachineChart(dashboardData);
  }
}

/* =====================================================
   LOAD DASHBOARD
===================================================== */
async function loadDashboard() {
  const startTime = performance.now();

  const dateParam = currentDate
    ? `&date=${encodeURIComponent(currentDate)}`
    : "";

  const [processedRes, machineRes] = await Promise.all([
    fetch(`${API_URL}?mode=processed${dateParam}`),
    fetch(`${API_URL}?mode=machine${dateParam}`)
  ]);

  dashboardProcessed = await processedRes.json();
  dashboardMachine = await machineRes.json();

  dashboardData = dashboardProcessed;
  buildInsights(dashboardProcessed);

  updateReportDateDisplay(dashboardProcessed);

  const summary = dashboardData.summary || {};

  /* ===============================
     SUMMARY CARDS (FIXED)
  ================================ */

  setText("totalJobs", summary.totalJobs ?? summary.coatingJobs ?? 0);
  setText("totalBreakage", summary.totalBreakRX ?? 0);
  setText("totalLenses", summary.totalBreakLenses ?? 0);

  setText(
    "rxBreakage",
    summary.breakPercent !== undefined
      ? summary.breakPercent + "%"
      : "0%"
  );

  setText(
    "avgTime",
    summary.avgBreakTimeHours
      ? summary.avgBreakTimeHours + " hrs"
      : "-"
  );

  setText("peakHour", summary.peakHour ?? "-");

  /* ===============================
     FLOW HEALTH SUMMARY
  ================================ */

  if (summary.flowHealth) {
    setText("flowHealthy", summary.flowHealth.healthy || 0);
    setText("flowWatch", summary.flowHealth.watch || 0);
    setText("flowDelayed", summary.flowHealth.delayed || 0);
    setText("flowOvernight", summary.flowHealth.overnight || 0);
  }

  /* ===============================
     BREAKAGE AGING SUMMARY
  ================================ */

  if (summary.aging) {
    setText("sameDayBreakage", summary.aging.sameDay || 0);
    setText("yesterdayBreakage", summary.aging.oneDay || 0);
    setText("twoPlusBreakage", summary.aging.twoPlus || 0);
  }

  /* ===============================
     HOURLY TABLE
  ================================ */

  buildHourlyTable(dashboardData.hourly || []);

  /* ===============================
     FORCE CHART REFRESH
  ================================ */

  if (trendChart) {
    trendChart.destroy();
    trendChart = null;
  }

  if (reasonChart) {
    reasonChart.destroy();
    reasonChart = null;
  }

  if (machineChart) {
    machineChart.destroy();
    machineChart = null;
  }

  if (flowChart) {
    flowChart.destroy();
    flowChart = null;
  }

  const activeTab = document.querySelector(".tab.active");

  if (activeTab) {
    const tabId = activeTab.getAttribute("onclick") || "";

    if (tabId.includes("trends")) {
      buildTrendCharts();
    } else if (tabId.includes("reasons")) {
      buildReasonChart(dashboardData);
    } else if (tabId.includes("machines")) {
      buildMachineChart(dashboardData);
    }
  }

  if (document.getElementById("flowChart")) {
    buildFlowChart(dashboardProcessed);
  }

  const endTime = performance.now();
  console.log("Load time (ms):", Math.round(endTime - startTime));
}

// =====================================================
// DATE FILTER
// =====================================================
function applyDateFilter() {
  const dateInput = document.getElementById("historyDate").value;

  if (!dateInput) {
    currentDate = null;
  } else {
    const d = new Date(dateInput);
    currentDate = (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
  }

  loadDashboard();
  startRefreshCountdown();
}

function resetToToday() {
  currentDate = null;
  const dateEl = document.getElementById("historyDate");
  if (dateEl) dateEl.value = "";
  loadDashboard();
}

// LIVE MODE BUTTON
document.addEventListener("DOMContentLoaded", () => {
  const liveBtn = document.getElementById("liveBtn");

  if (liveBtn) {
    liveBtn.addEventListener("click", () => {
      currentDate = null;

      const dateInput = document.getElementById("historyDate");
      if (dateInput) dateInput.value = "";

      loadDashboard();
      startRefreshCountdown();

      console.log("Switched to LIVE MODE");
    });
  }

  startRefreshCountdown();
});

// Keep this at the bottom
loadDashboard();

/* =====================================================
   TREND CHART SYSTEM
===================================================== */

function buildTrendCharts() {
  trendsLoaded = true;

  buildTrendChart(dashboardProcessed);

  setTimeout(() => {
    buildFlowChart(dashboardProcessed);
  }, 100);
}

function switchTrend(mode) {
  currentMode = mode;

  let data;

  if (mode === "processed") {
    data = dashboardProcessed;
    buildTrendChart(data);
  } else if (mode === "machine") {
    data = dashboardMachine;
    buildTrendChart(data);
  } else if (mode === "flow") {
    data = dashboardProcessed;
    buildDetaperFlowTrend(data);
  }

  document.querySelectorAll(".mode-btn")
    .forEach(btn => btn.classList.remove("active"));

  const activeBtn = document.getElementById(mode + "Btn");
  if (activeBtn) activeBtn.classList.add("active");
}

/* =====================================================
   FLOW CHART (DETAPER → COATER)
===================================================== */

function buildDetaperFlowTrend(data) {
  const canvas = document.getElementById("trendChart");
  if (!canvas || !data || !data.hourly) return;

  const ctx = canvas.getContext("2d");

  if (trendChart) trendChart.destroy();

  const hours = data.hourly.map(h => h.hour);

  const todayData = [];
  const yesterdayData = [];
  const twoPlusData = [];

  data.hourly.forEach(h => {
    todayData.push(h.avgSameDay || 0);
    yesterdayData.push(h.avgOneDay || 0);
    twoPlusData.push(h.avgTwoPlus || 0);
  });

  trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: hours,
      datasets: [
        {
          label: "Today (Flow)",
          data: todayData,
          borderColor: "#32ff7e",
          backgroundColor: createGradient(ctx, "#32ff7e"),
          borderWidth: 3,
          tension: 0.4,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 6
        },
        {
          label: "Yesterday (Flow)",
          data: yesterdayData,
          borderColor: "#ffd32a",
          backgroundColor: createGradient(ctx, "#ffd32a"),
          borderWidth: 3,
          tension: 0.4,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 6
        },
        {
          label: "2+ Days (Flow)",
          data: twoPlusData,
          borderColor: "#ff3f34",
          backgroundColor: createGradient(ctx, "#ff3f34"),
          borderWidth: 3,
          tension: 0.4,
          fill: true,
          pointRadius: 3,
          pointHoverRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },

      plugins: {
        legend: { labels: { color: "#E6F1FF" } },
        title: {
          display: true,
          text: "Process Flow Time (Detaper → Coater)",
          color: "#E6F1FF",
          font: { size: 16 }
        },
        tooltip: {
          bodyFont: { size: 14 },
          titleFont: { size: 15 },
          padding: 12,
          cornerRadius: 6,
          backgroundColor: "rgba(15, 25, 45, 0.95)",
          borderColor: "#4da3ff",
          borderWidth: 1
        }
      },

      scales: {
        x: {
          ticks: { color: "rgba(255,255,255,0.6)" },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        y: {
          beginAtZero: true,
          ticks: { color: "#4da3ff" },
          grid: { color: "rgba(255,255,255,0.05)" },
          title: {
            display: true,
            text: "Minutes",
            color: "#4da3ff"
          }
        }
      }
    },
    plugins: [GLOW_PLUGIN]
  });
}

function buildTrendChart(data) {
  const canvas = document.getElementById("trendChart");
  if (!canvas || !data || !data.hourly) return;

  const ctx = canvas.getContext("2d");

  if (trendChart) trendChart.destroy();

  const sortedHourly = [...data.hourly].sort((a, b) => {
    return new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour);
  });

  const hours = sortedHourly.map(h => h.hour);

  const datasets = [];
  const allReasons = new Set();

  // Collect unique reasons from machines only
  sortedHourly.forEach(hour => {
    Object.values(hour.machines || {}).forEach(machine => {
      Object.keys(machine.reasons || {}).forEach(reason => {
        allReasons.add(reason);
      });
    });
  });

  const colors = [
    "#FF4D4D",
    "#9b7bff",
    "#ff9f43",
    "#00d4ff",
    "#4da3ff",
    "#00ff88",
    "#ff6bcb"
  ];

  let colorIndex = 0;

  // Peak detection by total per reason series max
  allReasons.forEach(reason => {
    const color = BREAKAGE_COLOR_MAP[reason] || colors[colorIndex % colors.length];

    const reasonTotals = sortedHourly.map(hour => {
      let total = 0;

      Object.values(hour.machines || {}).forEach(machine => {
        total += machine.reasons?.[reason]?.total || 0;
      });

      return total;
    });

    const peakValue = Math.max(...reasonTotals, 0);

    datasets.push({
      label: reason,
      data: reasonTotals,
      borderColor: color,
      backgroundColor: createGradient(ctx, color),
      borderWidth: 3,
      tension: 0.4,
      fill: true,
      pointRadius: pointCtx => {
        const value = reasonTotals[pointCtx.dataIndex] || 0;
        return value === peakValue && value > 0 ? 5 : 3;
      },
      pointHoverRadius: 7,
      pointBackgroundColor: pointCtx => {
        const value = reasonTotals[pointCtx.dataIndex] || 0;
        return value === peakValue && value > 0 ? "#ffffff" : color;
      },
      pointBorderColor: color,
      pointBorderWidth: 2,
      spanGaps: true,
      yAxisID: "yBroken"
    });

    colorIndex++;
  });

  // Add Coating Jobs (right axis)
  datasets.push({
    label: "Total Coating Jobs",
    data: sortedHourly.map(h => h.coatingJobs || 0),
    borderColor: "#00ff88",
    backgroundColor: "transparent",
    borderDash: [6, 6],
    borderWidth: 2,
    tension: 0.35,
    fill: false,
    pointRadius: 2,
    pointHoverRadius: 5,
    yAxisID: "yJobs"
  });

  trendChart = new Chart(ctx, {
    type: "line",
    data: { labels: hours, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },

      plugins: {
        legend: { labels: { color: "#E6F1FF" } },

        title: {
          display: true,
          text: currentMode === "machine"
            ? "Machine Scan Trend"
            : "Breakage Processed by Reason",
          color: "#E6F1FF",
          font: { size: 16 }
        },

        tooltip: {
          bodyFont: { size: 14 },
          titleFont: { size: 15 },
          padding: 12,
          cornerRadius: 6,
          backgroundColor: "rgba(15, 25, 45, 0.95)",
          borderColor: "#4da3ff",
          borderWidth: 1,
          callbacks: {
            label: function(context) {
              const value = context.raw;
              if (context.dataset.label === "Total Coating Jobs") {
                return `${context.dataset.label}: ${value}`;
              }
              return `${context.dataset.label}: ${value} lenses`;
            }
          }
        }
      },

      scales: {
        x: {
          ticks: { color: "rgba(255,255,255,0.6)" },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        yBroken: {
          type: "linear",
          position: "left",
          beginAtZero: true,
          ticks: { color: "#ff4d4d" },
          grid: { color: "rgba(255,255,255,0.05)" },
          title: {
            display: true,
            text: "Breakage Count",
            color: "#ff4d4d"
          }
        },
        yJobs: {
          type: "linear",
          position: "right",
          beginAtZero: true,
          ticks: { color: "#00ff88" },
          grid: { drawOnChartArea: false },
          title: {
            display: true,
            text: "Coating Jobs",
            color: "#00ff88"
          }
        }
      }
    },
    plugins: [GLOW_PLUGIN]
  });
}

function getFlowColor(minutes) {
  if (minutes <= 15) return "#00ff88";
  if (minutes <= 30) return "#ffcc00";
  return "#ff3b3b";
}

/* =====================================================
   FLOW CHART SYSTEM (Detaper → Coater)
===================================================== */

function switchFlowMode(mode) {
  currentFlowMode = mode;

  document.querySelectorAll("#avgFlowBtn, #machineFlowBtn, #individualFlowBtn")
    .forEach(btn => btn.classList.remove("active"));

  const btnMap = {
    average: "avgFlowBtn",
    machine: "machineFlowBtn",
    individual: "individualFlowBtn"
  };

  const activeBtn = document.getElementById(btnMap[mode]);
  if (activeBtn) activeBtn.classList.add("active");

  buildFlowChart(dashboardProcessed);
}

function buildFlowChart(data) {
  const canvas = document.getElementById("flowChart");
  if (!canvas || !data || !data.hourly) return;

  const ctx = canvas.getContext("2d");

  if (flowChart) flowChart.destroy();

  const sortedHourly = [...data.hourly].sort((a, b) => {
    return new Date("1/1/2000 " + a.hour) - new Date("1/1/2000 " + b.hour);
  });

  const filteredHours = sortedHourly.filter(h => {
    if (!h.hour) return false;

    const hourParts = h.hour.split(":");
    let hourNum = parseInt(hourParts[0], 10);

    const isPM = h.hour.includes("PM");
    const isAM = h.hour.includes("AM");

    if (isPM && hourNum !== 12) hourNum += 12;
    if (isAM && hourNum === 12) hourNum = 0;

    return hourNum >= 6 && hourNum <= 20;
  });

  const hours = filteredHours.map(h => h.hour);

  function sanitize(value) {
    if (value === null || value === undefined) return null;
    const num = Number(value);
    if (num > 500) return null;
    return num;
  }

  // MACHINE MODE
  if (currentFlowMode === "machine") {
    const machineSet = new Set();

    filteredHours.forEach(h => {
      Object.keys(h.machines || {}).forEach(m => machineSet.add(m));
    });

    const machines = Array.from(machineSet);

    const machinePalette = [
      "#4da3ff",
      "#ff4d4d",
      "#ff9f43",
      "#ffd32a",
      "#32ff7e",
      "#9b7bff",
      "#00d4ff",
      "#ff6bcb"
    ];

    const datasets = machines.map((machine, idx) => {
      const color = machinePalette[idx % machinePalette.length];

      const dataPoints = filteredHours.map(h => ({
        x: h.hour,
        y: sanitize(h.machines?.[machine]?.avgFlowAll),
        count: h.machines?.[machine]?.flowAllCount || 0
      }));

      return {
        label: formatMachineLabel(machine),
        data: dataPoints,
        borderColor: color,
        backgroundColor: createGradient(ctx, color),
        borderWidth: 3,
        tension: 0.35,
        fill: false,
        pointRadius: 3,
        pointHoverRadius: 6,
        spanGaps: true
      };
    });

    flowChart = new Chart(ctx, {
      type: "line",
      data: { labels: hours, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },

        plugins: {
          legend: { labels: { color: "#E6F1FF" } },
          title: {
            display: true,
            text: "Machine Average Flow (Detaper → Coater)",
            color: "#E6F1FF"
          },
          tooltip: {
            backgroundColor: "rgba(15, 25, 45, 0.95)",
            borderColor: "#4da3ff",
            borderWidth: 1,
            callbacks: {
              label: function(context) {
                const raw = context.raw;
                const value = raw?.y ?? raw;
                const count = raw?.count ?? 0;
                return `${context.dataset.label}: ${value} mins (${count} jobs)`;
              }
            }
          }
        },

        scales: {
          x: {
            ticks: { color: "rgba(255,255,255,0.6)" },
            grid: { color: "rgba(255,255,255,0.05)" }
          },
          y: {
            beginAtZero: true,
            suggestedMax: 60,
            ticks: { color: "#4da3ff" },
            grid: { color: "rgba(255,255,255,0.05)" },
            title: {
              display: true,
              text: "Minutes",
              color: "#4da3ff"
            }
          }
        }
      },
      plugins: [GLOW_PLUGIN]
    });

    return;
  }

  // AVERAGE MODE
  if (currentFlowMode === "average") {
    const healthyFlow = filteredHours.map(h => {
      const healthy = (h.flowPoints || []).filter(p => p.flow <= 15);

      const minutes = healthy.length
        ? Math.round(healthy.reduce((a, b) => a + b.flow, 0) / healthy.length)
        : 0;

      return {
        x: h.hour,
        y: minutes,
        jobs: healthy.length
      };
    });

    const watchFlow = filteredHours.map(h => {
      const watch = (h.flowPoints || [])
        .filter(p => p.flow > 15 && p.flow <= 30);

      const minutes = watch.length
        ? Math.round(watch.reduce((a, b) => a + b.flow, 0) / watch.length)
        : 0;

      return {
        x: h.hour,
        y: minutes,
        jobs: watch.length
      };
    });

    const delayedFlow = filteredHours.map(h => {
      const delayed = (h.flowPoints || [])
        .filter(p => p.flow > 30 && p.flow <= 360);

      const minutes = delayed.length
        ? Math.round(delayed.reduce((a, b) => a + b.flow, 0) / delayed.length)
        : 0;

      return {
        x: h.hour,
        y: minutes,
        jobs: delayed.length
      };
    });

    const overnightFlow = filteredHours.map(h => {
      const overnight = (h.flowPoints || [])
        .filter(p => p.flow > 360);

      const minutes = overnight.length
        ? Math.round(overnight.reduce((a, b) => a + b.flow, 0) / overnight.length)
        : 0;

      return {
        x: h.hour,
        y: minutes,
        jobs: overnight.length
      };
    });

    const brokenFlow = filteredHours.map(h => ({
      x: h.hour,
      y: sanitize(h.avgFlowBroken),
      count: h.flowBrokenCount || 0
    }));

    flowChart = new Chart(ctx, {
      type: "line",
      data: {
        labels: hours,
        datasets: [
          {
            label: "Workflow Healthy",
            data: healthyFlow,
            borderColor: "#32ff7e",
            backgroundColor: createGradient(ctx, "#32ff7e"),
            borderWidth: 3,
            tension: 0.4,
            fill: true,
            spanGaps: true,
            pointRadius: 3,
            pointHoverRadius: 6
          },
          {
            label: "Workflow Watch",
            data: watchFlow,
            borderColor: "#ffd32a",
            backgroundColor: createGradient(ctx, "#ffd32a"),
            borderWidth: 3,
            tension: 0.4,
            fill: true,
            spanGaps: true,
            pointRadius: 3,
            pointHoverRadius: 6
          },
          {
            label: "Workflow Delayed",
            data: delayedFlow,
            borderColor: "#ff3f34",
            backgroundColor: createGradient(ctx, "#ff3f34"),
            borderWidth: 3,
            tension: 0.4,
            fill: true,
            spanGaps: true,
            pointRadius: 3,
            pointHoverRadius: 6
          },
          {
            label: "Overnight Carryover",
            data: overnightFlow,
            borderColor: "#00a8ff",
            backgroundColor: createGradient(ctx, "#00a8ff"),
            borderWidth: 3,
            tension: 0.4,
            fill: true,
            spanGaps: true,
            pointRadius: 3,
            pointHoverRadius: 6
          },
          {
            label: "Broken Jobs Avg Flow",
            data: brokenFlow,
            borderColor: "#9b7bff",
            backgroundColor: createGradient(ctx, "#9b7bff"),
            borderWidth: 3,
            tension: 0.4,
            fill: true,
            spanGaps: true,
            pointRadius: 3,
            pointHoverRadius: 6
          }
        ]
      },
      options: getFlowOptions("Detaper → Coater Flow Analysis"),
      plugins: [GLOW_PLUGIN]
    });

    return;
  }

  // INDIVIDUAL MODE
  const points = [];

  filteredHours.forEach(hourObj => {
    Object.entries(hourObj.machines || {}).forEach(([machine, mData]) => {
      const flowValue = sanitize(mData.avgFlowAll);
      if (flowValue === null) return;

      points.push({
        x: hourObj.hour,
        y: flowValue,
        machine: machine,
        rx: null
      });
    });
  });

  flowChart = new Chart(ctx, {
    type: "scatter",
    data: {
      datasets: [{
        label: "Machine Flow",
        data: points,
        pointRadius: 8,
        pointHoverRadius: 11,
        backgroundColor: points.map(p => getFlowColor(p.y)),
        borderColor: points.map(p => getFlowColor(p.y))
      }]
    },
    options: getFlowOptions("Machine Flow Scatter (Detaper → Coater)"),
    plugins: [GLOW_PLUGIN]
  });
}

// =====================================================
// FLOW CHART OPTIONS
// =====================================================

function getFlowOptions(titleText) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "nearest", intersect: true },

    onClick: function(evt, elements, chart) {
      if (!elements.length) return;

      const element = elements[0];
      const datasetIndex = element.datasetIndex;
      const index = element.index;

      const dataset = chart.data.datasets[datasetIndex];
      const point = dataset.data[index];

      if (!point) return;

      showFlowDetails({
        rx: point.rx || "Machine Flow",
        machine: point.machine || "Unknown",
        reason: point.reason || "N/A",
        flow: point.y,
        x: point.x
      });
    },

    plugins: {
      legend: {
        labels: { color: "#E6F1FF" }
      },

      title: {
        display: true,
        text: titleText,
        color: "#E6F1FF",
        font: { size: 16 }
      },

      tooltip: {
        bodyFont: { size: 14 },
        titleFont: { size: 15 },
        padding: 12,
        cornerRadius: 6,
        backgroundColor: "rgba(15, 25, 45, 0.95)",
        borderColor: "#4da3ff",
        borderWidth: 1,

        callbacks: {
          label: function(context) {
            const raw = context.raw;
            if (!raw) return "";

            if (raw.rx) {
              const lines = [];
              lines.push(`RX: ${raw.rx}`);

              if (raw.machine) lines.push(`Machine: ${raw.machine}`);
              if (raw.reason) lines.push(`Breakage: ${raw.reason}`);

              lines.push(`Flow Time: ${raw.y} mins`);

              if (raw.x) lines.push(`Hour: ${raw.x}`);

              return lines;
            }

            if (raw.machine) {
              const lines = [];
              lines.push(`Machine: ${raw.machine}`);
              lines.push(`Flow Time: ${raw.y} mins`);

              if (raw.x) lines.push(`Hour: ${raw.x}`);

              return lines;
            }

            const value = raw?.y ?? raw;
            const jobs = raw?.jobs ?? raw?.count ?? null;

            if (jobs !== null) {
              return `${context.dataset.label}: ${value} mins (${jobs} jobs)`;
            }

            return `${context.dataset.label}: ${value} mins`;
          }
        }
      }
    },

    scales: {
      x: {
        type: "category",
        ticks: { color: "rgba(255,255,255,0.6)" },
        grid: { color: "rgba(255,255,255,0.05)" }
      },

      y: {
        beginAtZero: true,
        suggestedMax: 60,
        ticks: { color: "#4da3ff" },
        grid: { color: "rgba(255,255,255,0.05)" },
        title: {
          display: true,
          text: "Minutes",
          color: "#4da3ff"
        }
      }
    }
  };
}

// =============================================
// FLOW MODAL DETAIL (INDIVIDUAL)
// =============================================
function showFlowDetails(point) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;

  const totalMinutes = Number(point.flow) || 0;

  const hours = Math.floor(totalMinutes / 60);
  const minutes = Math.round(totalMinutes % 60);

  let formatted;

  if (hours > 0) {
    formatted = minutes > 0
      ? `${hours}h ${minutes}m`
      : `${hours}h`;
  } else {
    formatted = `${minutes}m`;
  }

  modalBody.innerHTML = `
    <h2>RX ${point.rx}</h2>
    <div style="margin-top:15px; padding:20px; background:#243b63; border-radius:10px;">
      <p><strong>Machine:</strong> ${point.machine}</p>
      <p><strong>Breakage Reason:</strong> ${point.reason || "None"}</p>
      <p><strong>Flow Time:</strong> ${formatted}</p>
      <p><strong>Hour:</strong> ${point.x}</p>
    </div>
  `;

  modal.classList.add("active");
}

// =============================================
// FLOW MODAL DETAIL (HOUR CLICK)
// =============================================
function showFlowHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody) return;

  let html = "";

  (hourData.flowPoints || []).forEach(p => {
    const color = getFlowColor(p.flow);

    html += `
      <div style="margin-bottom:10px; padding:12px; background:#1e2f4d; border-left:5px solid ${color}; border-radius:6px;">
        <strong>RX ${p.rx}</strong><br>
        Machine: ${p.machine}<br>
        Reason: ${p.reason}<br>
        Flow: ${p.flow}m
      </div>
    `;
  });

  modalBody.innerHTML = `
    <h2>${hourData.hour}</h2>
    ${html || "<p>No flow data</p>"}
  `;

  modal.classList.add("active");
}

/* =====================================================
   HOURLY TABLE
===================================================== */
function buildHourlyTable(hourly) {
  const tbody = document.getElementById("hourlyBody");
  if (!tbody) return;
  if (!Array.isArray(hourly)) return;

  tbody.innerHTML = "";

  hourly.forEach(row => {
    const tr = document.createElement("tr");

    let primaryDriverHTML = "-";

    if ((row.totalBroken || 0) > 0 && row.machines) {
      let topMachine = null;
      let topMachineTotal = 0;

      Object.entries(row.machines).forEach(([machine, stats]) => {
        if ((stats.total || 0) > topMachineTotal) {
          topMachine = machine;
          topMachineTotal = stats.total || 0;
        }
      });

      if (topMachine && row.machines[topMachine]) {
        let topReason = null;
        let topReasonTotal = 0;

        const reasons = row.machines[topMachine].reasons || {};

        Object.entries(reasons).forEach(([reason, rStats]) => {
          if ((rStats.total || 0) > topReasonTotal) {
            topReason = reason;
            topReasonTotal = rStats.total || 0;
          }
        });

        primaryDriverHTML = `
          <div class="primary-driver">
            <div class="driver-machine">${formatMachineLabel(topMachine)}</div>
            <div class="driver-reason">${topReason || ""}</div>
          </div>
        `;
      }
    }

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td>${row.totalBroken || 0}</td>
      <td>${row.coatingJobs || 0}</td>
      <td>${primaryDriverHTML}</td>
    `;

    tr.addEventListener("click", () => {
      showHourDetails(row);
    });

    tbody.appendChild(tr);
  });
}

/* =====================================================
   MODAL DETAILS (RESTORED)
===================================================== */
function showHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");
  if (!modal || !modalBody || !hourData) return;

  let machineHTML = "";

  Object.entries(hourData.machines || {}).forEach(([machine, stats]) => {
    const displayMachine = formatMachineLabel(machine);
    let reasonHTML = "";

    const reasonsObj = stats.reasons || {};

    Object.entries(reasonsObj).forEach(([reason, rStats]) => {
      reasonHTML += `
        <div style="margin-top:8px; padding-left:12px;">
          • <strong>${reason}</strong> — ${rStats.total || 0}
          <br>
          <span style="font-size:12px; opacity:.7;">
            Same Day Brkg: ${rStats.sameDay || 0} |
            Previous Day Brkg: ${rStats.oneDay || 0} |
            2+ Days Brkg: ${rStats.twoPlus || 0}
          </span>
        </div>
      `;
    });

    machineHTML += `
      <div style="margin-bottom:16px; padding:14px; background:#243b63; border-radius:8px;">
        <strong>${displayMachine}</strong><br>
        Jobs Ran: ${stats.jobs || 0}<br>
        Total Brkg: ${stats.total || 0}<br>
        Same Day Brkg: ${stats.sameDay || 0}<br>
        Previous Day Brkg: ${stats.oneDay || 0}<br>
        2+ Days Brkg: ${stats.twoPlus || 0}
        ${reasonHTML}
      </div>
    `;
  });

  modalBody.innerHTML = `
    <h2>${hourData.hour}</h2>
    <h3>Total Lenses Broken: ${hourData.totalBroken || 0}</h3>
    ${machineHTML || "<p>No data</p>"}
  `;

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

  if (!data.topReasons) return;

  const ctx = document.getElementById("reasonChart").getContext("2d");

  if (reasonChart) reasonChart.destroy();

  const entries = Object.entries(data.topReasons)
    .map(([reason, stats]) => ({
      reason,
      total: stats.total || 0
    }))
    .sort((a, b) => b.total - a.total);

  const maxIndex = 0; // 🔥 highest bar

  reasonChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: entries.map(e => e.reason),
      datasets: [{
        label: "Breakage Count",
        data: entries.map(e => e.total),

        backgroundColor: (ctx) => {
          const i = ctx.dataIndex;
          const baseColor =
            BREAKAGE_COLOR_MAP[entries[i].reason] || "#4da3ff";

          return createGradient(ctx.chart.ctx, baseColor);
        },

        borderRadius: 10,
        borderSkipped: false,
        borderWidth: (ctx) => ctx.dataIndex === maxIndex ? 2 : 0,
        borderColor: "#ffffff"
      }]
    },

    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "rgba(15,25,45,0.95)"
        }
      },

      scales: {
        x: {
          ticks: { color: "#9db4d4" },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        y: {
          ticks: { color: "#E6F1FF" }
        }
      }
    },

    plugins: [GLOW_PLUGIN]
  });
}

/* =====================================================
   MACHINE CHART
===================================================== */
function buildMachineChart(data) {

  if (!data.machineTotals) return;

  const ctx = document.getElementById("machineChart").getContext("2d");

  if (machineChart) machineChart.destroy();

  const entries = Object.entries(data.machineTotals)
    .map(([machine, stats]) => {

      const jobs = stats.jobs || 0;
      const broken = stats.breakLenses || 0;
      const percent = jobs > 0 ? (broken / jobs) * 100 : 0;

      return {
        machine,
        percent,
        jobs
      };
    })
    .sort((a, b) => b.percent - a.percent);

  const worstIndex = 0;

  machineChart = new Chart(ctx, {
    type: "bar",

    data: {
      labels: entries.map(e => e.machine),

      datasets: [

        // 🔥 BREAKAGE %
        {
          label: "Breakage %",
          data: entries.map(e => e.percent),

          backgroundColor: (ctx) => {
            const i = ctx.dataIndex;
            const val = entries[i].percent;

            if (i === worstIndex) return "#ff3f34"; // 🔥 worst
            if (val > 5) return "#ff9f43";
            return "#ffd32a";
          },

          borderRadius: 8,
          yAxisID: "y1"
        },

        // 🔵 JOBS
        {
          label: "Total Jobs",
          data: entries.map(e => e.jobs),
          backgroundColor: createGradient(ctx, "#4da3ff"),
          borderRadius: 8,
          yAxisID: "y2"
        }

      ]
    },

    options: {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          labels: { color: "#E6F1FF" }
        },
        tooltip: {
          backgroundColor: "rgba(15,25,45,0.95)"
        }
      },

      scales: {

        y1: {
          position: "left",
          title: {
            display: true,
            text: "Breakage %",
            color: "#ff3f34"
          },
          ticks: { color: "#ff3f34" }
        },

        y2: {
          position: "right",
          title: {
            display: true,
            text: "Jobs",
            color: "#4da3ff"
          },
          ticks: { color: "#4da3ff" },
          grid: { drawOnChartArea: false }
        },

        x: {
          ticks: { color: "#E6F1FF" }
        }

      }
    },

    plugins: [GLOW_PLUGIN]
  });
}
// AUTO REFRESH EVERY 5 MINUTES (LIVE MODE ONLY)
setInterval(() => {
  if (currentDate === null) {
    console.log("Auto refreshing (Live Mode)");
    loadDashboard();
    startRefreshCountdown();
  }
}, refreshInterval * 1000);

// DATE FILTER
const dateInput = document.getElementById("historyDate");

if (dateInput) {
  dateInput.addEventListener("change", function() {
    if (!this.value) {
      currentDate = null;
    } else {
      const parts = this.value.split("-");

      currentDate =
        parseInt(parts[1], 10) + "/" +
        parseInt(parts[2], 10) + "/" +
        parts[0];
    }

    console.log("Date changed to:", currentDate);
    loadDashboard();
    startRefreshCountdown();
  });
}

/* =====================================================
   🧠 BUILD INSIGHTS ENGINE
===================================================== */
function buildInsights(data) {

  const el = document.getElementById("insightBox");
  if (!el || !data) return;

  const summary = data.summary || {};
  const hourly = data.hourly || [];
  const machineTotals = data.machineTotals || {};
  const topReasons = data.topReasons || {};

  let insights = [];

  /* ===============================
     1. PEAK HOUR
  ================================ */
  let peakHour = "-";
  let peakValue = 0;

  hourly.forEach(h => {
    const val = h.totalBroken || 0;
    if (val > peakValue) {
      peakValue = val;
      peakHour = h.hour;
    }
  });

  if (peakValue > 0) {
    insights.push(`🔥 Peak Breakage: <b>${peakHour}</b> (${peakValue} lenses)`);
  }

  /* ===============================
     2. TOP BREAKAGE REASON
  ================================ */
  let topReason = "-";
  let topReasonCount = 0;

  Object.entries(topReasons).forEach(([reason, stats]) => {
    if ((stats.total || 0) > topReasonCount) {
      topReason = reason;
      topReasonCount = stats.total || 0;
    }
  });

  if (topReasonCount > 0) {
    insights.push(`⚠️ Top Issue: <b>${topReason}</b> (${topReasonCount})`);
  }

  /* ===============================
     3. WORST MACHINE
  ================================ */
  let worstMachine = "-";
  let worstPercent = 0;

  Object.entries(machineTotals).forEach(([machine, stats]) => {
    const jobs = stats.jobs || 0;
    const broken = stats.breakLenses || 0;

    if (jobs > 0) {
      const percent = (broken / jobs) * 100;

      if (percent > worstPercent) {
        worstPercent = percent;
        worstMachine = machine;
      }
    }
  });

  if (worstPercent > 0) {
    insights.push(
      `🛠️ Highest Breakage Machine: <b>${formatMachineLabel(worstMachine)}</b> (${worstPercent.toFixed(1)}%)`
    );
  }

  /* ===============================
     4. FLOW HEALTH RISK
  ================================ */
  if (summary.flowHealth) {
    const delayed = summary.flowHealth.delayed || 0;
    const healthy = summary.flowHealth.healthy || 0;

    if (delayed > healthy) {
      insights.push(`🚨 Flow Risk: Delayed jobs exceed healthy flow`);
    } else if (delayed > 0) {
      insights.push(`⚡ Flow Warning: ${delayed} delayed jobs detected`);
    } else {
      insights.push(`✅ Flow Stable: No delays detected`);
    }
  }

  /* ===============================
     5. TREND DETECTION (LAST HOURS)
  ================================ */
  if (hourly.length >= 3) {
    const last = hourly.slice(-3).map(h => h.totalBroken || 0);

    if (last[2] > last[1] && last[1] > last[0]) {
      insights.push(`📈 Increasing Breakage Trend detected (last 3 hrs)`);
    }

    if (last[2] < last[1] && last[1] < last[0]) {
      insights.push(`📉 Breakage decreasing trend (last 3 hrs)`);
    }
  }

  /* ===============================
     6. ZERO BREAKAGE ALERT
  ================================ */
  if (summary.totalBreakLenses === 0) {
    insights.push(`🟢 No breakage recorded (perfect run)`);
  }

  /* ===============================
     OUTPUT
  ================================ */

  if (insights.length === 0) {
    el.innerHTML = "<p>No insights available</p>";
    return;
  }

  el.innerHTML = insights
    .map(i => `<div style="margin-bottom:8px;">${i}</div>`)
    .join("");
}

// =====================================================
// FORCE REBUILD ALL CHARTS
// =====================================================
function rebuildAllCharts() {
  if (trendChart) {
    trendChart.destroy();
    trendChart = null;
  }

  if (reasonChart) {
    reasonChart.destroy();
    reasonChart = null;
  }

  if (machineChart) {
    machineChart.destroy();
    machineChart = null;
  }

  if (flowChart) {
    flowChart.destroy();
    flowChart = null;
  }

  const activeTab = document.querySelector(".tab.active");
  if (!activeTab) return;

  const tabText = activeTab.textContent.toLowerCase();

  if (tabText.includes("trend")) {
    buildTrendCharts();
  } else if (tabText.includes("reason")) {
    buildReasonChart(dashboardData);
  } else if (tabText.includes("machine")) {
    buildMachineChart(dashboardData);
  }

  if (document.getElementById("flowChart")) {
    buildFlowChart(dashboardProcessed);
  }
}

function startRefreshCountdown() {
  if (refreshTimerHandle) {
    clearInterval(refreshTimerHandle);
  }

  const timerEl = document.getElementById("refreshTimer");
  if (!timerEl) return;

  refreshCountdown = refreshInterval;

  refreshTimerHandle = setInterval(() => {
    if (currentDate !== null) {
      timerEl.textContent = "History Mode";
      return;
    }

    const minutes = Math.floor(refreshCountdown / 60);
    const seconds = refreshCountdown % 60;

    timerEl.textContent =
      `Refresh in ${minutes}:${seconds.toString().padStart(2, "0")}`;

    refreshCountdown--;

    if (refreshCountdown < 0) {
      refreshCountdown = refreshInterval;
    }
  }, 1000);
}

// =====================================================
// NAVIGATION
// =====================================================
function goBack() {
  window.location.href = "index.html";
}