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
let isLiveMode = true;

/* =====================================================
   GLOBAL BREAKAGE COLOR MAP (FIXED COLORS)
   ===================================================== */

const BREAKAGE_COLOR_MAP = {
  "S-HC Contamination": "#FF4D4D",
  "S-HC Pit": "#9b7bff",
  "S-HC Run": "#ff9f43",
  "S-HC Wagon Wheel": "#4da3ff",
  "S-HC Suction Cup Marks": "#00d4ff",
  "S-HC HC Suction Cup Marks": "#00d4ff",
  // Add future reasons here to lock colors permanently
};

function formatMachineLabel(machine) {
  if (!machine) return "";
  return machine.replace(/^AR41-/, "44R1-");
}

/* =====================================================
   REPORT DATE DISPLAY
   ===================================================== */
function updateReportDateDisplay(data) {
  const el = document.getElementById("activeReportDate");
  if (!el) return;

  if (isLiveMode || currentDate === null) {
    el.innerHTML = "Viewing Report Date: <strong>Today (Live)</strong>";
  } else {
    el.innerHTML = "Viewing Report Date: <strong>" + currentDate + "</strong>";
  }
}

/* =====================================================
   SAFE TEXT SETTER
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

  document.getElementById(tabId).classList.add("active");
  button.classList.add("active");

  if (tabId === "trends" && !trendChart) {
    buildTrendCharts();
    buildFlowChart(dashboardProcessed);
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

  try {
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

    updateReportDateDisplay(dashboardProcessed);

    const summary = dashboardData.summary || {};

    /* ===============================
       SUMMARY CARDS
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

    /* ===============================
       AUTO BUILD CHARTS
       no click needed anymore
    ================================ */
    trendsLoaded = true;
    buildTrendChart(dashboardProcessed);
    buildReasonChart(dashboardProcessed);
    buildMachineChart(dashboardProcessed);
    buildFlowChart(dashboardProcessed);

    const endTime = performance.now();
    console.log("Load time (ms):", Math.round(endTime - startTime));
  } catch (err) {
    console.error("loadDashboard error:", err);
  }
}

/* =====================================================
   DATE FILTER
   ===================================================== */
function applyDateFilter() {
  const dateInput = document.getElementById("historyDate").value;

  if (!dateInput) {
    currentDate = null;
    isLiveMode = true;
  } else {
    const d = new Date(dateInput);
    currentDate = (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
    isLiveMode = false;
  }

  loadDashboard();
}

function resetToToday() {
  currentDate = null;
  isLiveMode = true;

  const dateEl = document.getElementById("historyDate");
  if (dateEl) dateEl.value = "";

  loadDashboard();
}

// Keep this at the bottom
loadDashboard();

/* =====================================================
   TREND CHART SYSTEM
   ===================================================== */

function buildTrendCharts() {
  trendsLoaded = true;
  buildTrendChart(currentMode === "machine" ? dashboardMachine : dashboardProcessed);
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
          borderWidth: 3,
          tension: 0.35,
          fill: false
        },
        {
          label: "Yesterday (Flow)",
          data: yesterdayData,
          borderColor: "#ffd32a",
          borderWidth: 3,
          tension: 0.35,
          fill: false
        },
        {
          label: "2+ Days (Flow)",
          data: twoPlusData,
          borderColor: "#ff3f34",
          borderWidth: 3,
          tension: 0.35,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          labels: {
            color: "#E6F1FF",
            font: { size: 14 }
          }
        },
        title: {
          display: true,
          text: "Process Flow Time (Detaper → Coater)",
          color: "#E6F1FF",
          font: { size: 18 }
        }
      },

      scales: {
        x: {
          ticks: {
            color: "rgba(255,255,255,0.75)",
            font: { size: 13 }
          },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        y: {
          beginAtZero: true,
          ticks: {
            color: "#4da3ff",
            font: { size: 13 }
          },
          grid: { color: "rgba(255,255,255,0.05)" },
          title: {
            display: true,
            text: "Minutes",
            color: "#4da3ff",
            font: { size: 14 }
          }
        }
      }
    }
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

  allReasons.forEach(reason => {
    const reasonTotals = sortedHourly.map(hour => {
      let total = 0;

      Object.values(hour.machines || {}).forEach(machine => {
        total += machine.reasons?.[reason]?.total || 0;
      });

      return total;
    });

    datasets.push({
      label: reason,
      data: reasonTotals,
      borderColor: colors[colorIndex % colors.length],
      backgroundColor: colors[colorIndex % colors.length],
      borderWidth: 3,
      tension: 0.35,
      fill: false,
      pointRadius: 4,
      yAxisID: "yBroken"
    });

    colorIndex++;
  });

  datasets.push({
    label: "Total Coating Jobs",
    data: sortedHourly.map(h => h.coatingJobs || 0),
    borderColor: "#00ff88",
    borderDash: [6, 6],
    borderWidth: 2,
    tension: 0.35,
    fill: false,
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
        legend: {
          labels: {
            color: "#E6F1FF",
            font: { size: 14 }
          }
        },
        title: {
          display: true,
          text: "Breakage Processed by Reason",
          color: "#E6F1FF",
          font: { size: 18 }
        }
      },
      scales: {
        x: {
          ticks: {
            color: "rgba(255,255,255,0.75)",
            font: { size: 13 }
          },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        yBroken: {
          type: "linear",
          position: "left",
          beginAtZero: true,
          ticks: {
            color: "#ff4d4d",
            font: { size: 13 }
          },
          title: {
            display: true,
            text: "Breakage Count",
            color: "#ff4d4d",
            font: { size: 14 }
          }
        },
        yJobs: {
          type: "linear",
          position: "right",
          beginAtZero: true,
          ticks: {
            color: "#00ff88",
            font: { size: 13 }
          },
          grid: { drawOnChartArea: false },
          title: {
            display: true,
            text: "Coating Jobs",
            color: "#00ff88",
            font: { size: 14 }
          }
        }
      }
    }
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

  /* =========================
     MACHINE MODE
  ========================= */
  if (currentFlowMode === "machine") {
    const machineSet = new Set();

    filteredHours.forEach(h => {
      Object.keys(h.machines || {}).forEach(m => machineSet.add(m));
    });

    const machines = Array.from(machineSet);

    const datasets = machines.map(machine => {
      const dataPoints = filteredHours.map(h => ({
        x: h.hour,
        y: sanitize(h.machines?.[machine]?.avgFlowAll),
        count: h.machines?.[machine]?.flowAllCount || 0
      }));

      return {
        label: formatMachineLabel(machine),
        data: dataPoints,
        borderWidth: 3,
        tension: 0.35,
        fill: false,
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
          legend: {
            labels: {
              color: "#E6F1FF",
              font: { size: 14 }
            }
          },
          title: {
            display: true,
            text: "Machine Average Flow (Detaper → Coater)",
            color: "#E6F1FF",
            font: { size: 18 }
          },
          tooltip: {
            callbacks: {
              label: function (context) {
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
            ticks: {
              color: "rgba(255,255,255,0.75)",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" }
          },
          y: {
            beginAtZero: true,
            suggestedMax: 60,
            ticks: {
              color: "#4da3ff",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" },
            title: {
              display: true,
              text: "Minutes",
              color: "#4da3ff",
              font: { size: 14 }
            }
          }
        }
      }
    });

    return;
  }

  /* =========================
     AVERAGE MODE
  ========================= */
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

    const postDelay = filteredHours.map(h => ({
      x: h.hour,
      y: sanitize(h.avgPostCoatDelay),
      count: h.postCoatCount || 0
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
            backgroundColor: "rgba(50,255,126,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          },
          {
            label: "Workflow Watch",
            data: watchFlow,
            borderColor: "#ffd32a",
            backgroundColor: "rgba(255,211,42,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          },
          {
            label: "Workflow Delayed",
            data: delayedFlow,
            borderColor: "#ff3f34",
            backgroundColor: "rgba(255,63,52,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          },
          {
            label: "Overnight Carryover",
            data: overnightFlow,
            borderColor: "#00a8ff",
            backgroundColor: "rgba(0,168,255,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          },
          {
            label: "Broken Jobs Avg Flow",
            data: brokenFlow,
            borderColor: "#9b7bff",
            backgroundColor: "rgba(155,123,255,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          },
          {
            label: "Coater → Break Delay",
            data: postDelay,
            borderColor: "#ff6b6b",
            backgroundColor: "rgba(255,107,107,0.15)",
            borderWidth: 3,
            tension: 0.35,
            fill: true,
            spanGaps: true
          }
        ]
      },
      options: getFlowOptions("Detaper → Coater Flow Analysis")
    });

    return;
  }

  /* =========================
     INDIVIDUAL MODE
  ========================= */
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
    options: getFlowOptions("Machine Flow Scatter (Detaper → Coater)")
  });
}

/* =====================================================
   FLOW CHART OPTIONS
   ===================================================== */

function getFlowOptions(titleText) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "nearest", intersect: true },

    onClick: function (evt, elements, chart) {
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
        labels: {
          color: "#E6F1FF",
          font: { size: 14 }
        }
      },

      title: {
        display: true,
        text: titleText,
        color: "#E6F1FF",
        font: { size: 18 }
      },

      tooltip: {
        callbacks: {
          label: function (context) {
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
        ticks: {
          color: "rgba(255,255,255,0.75)",
          font: { size: 13 }
        },
        grid: { color: "rgba(255,255,255,0.05)" }
      },

      y: {
        beginAtZero: true,
        suggestedMax: 60,
        ticks: {
          color: "#4da3ff",
          font: { size: 13 }
        },
        grid: { color: "rgba(255,255,255,0.05)" },
        title: {
          display: true,
          text: "Minutes",
          color: "#4da3ff",
          font: { size: 14 }
        }
      }
    }
  };
}

/* =============================================
   FLOW MODAL DETAIL (INDIVIDUAL)
   ============================================= */
function showFlowDetails(point) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");

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

/* =============================================
   FLOW MODAL DETAIL (HOUR CLICK)
   ============================================= */
function showFlowHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");

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
   MODAL DETAILS
   ===================================================== */
function showHourDetails(hourData) {
  const modal = document.getElementById("chartModal");
  const modalBody = document.getElementById("modalBody");

  if (!hourData) return;

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
  if (!data || !data.topReasons) return;

  if (reasonChart) reasonChart.destroy();

  const entries = Object.entries(data.topReasons)
    .map(([reason, stats]) => ({
      reason,
      total: stats.total || 0
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 6);

  const totalBreakage = entries.reduce((sum, r) => sum + r.total, 0);

  reasonChart = new Chart(
    document.getElementById("reasonChart"),
    {
      type: "bar",
      data: {
        labels: entries.map(e => e.reason),
        datasets: [{
          label: "Total Breakage",
          data: entries.map(e => e.total),
          backgroundColor: entries.map(e =>
            BREAKAGE_COLOR_MAP[e.reason] || "#4da3ff"
          ),
          borderRadius: 8
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,

        plugins: {
          legend: { display: false },
          title: {
            display: true,
            text: "Top Breakage Reasons",
            color: "#E6F1FF",
            font: { size: 18 }
          },
          tooltip: {
            callbacks: {
              label: function (context) {
                const value = context.raw;
                const percent = totalBreakage
                  ? ((value / totalBreakage) * 100).toFixed(1)
                  : 0;

                return `${value} lenses (${percent}%)`;
              }
            }
          }
        },

        scales: {
          x: {
            ticks: {
              color: "rgba(255,255,255,0.75)",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" }
          },
          y: {
            ticks: {
              color: "#E6F1FF",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" }
          }
        }
      }
    }
  );
}

/* =====================================================
   MACHINE CHART
   ===================================================== */
function buildMachineChart(data) {
  if (!data || !data.machineTotals) return;

  if (machineChart) machineChart.destroy();

  const entries = Object.entries(data.machineTotals)
    .map(([machine, stats]) => {
      const jobs = stats.jobs || 0;
      const broken = stats.breakLenses || 0;
      const percent = jobs > 0 ? (broken / jobs) * 100 : 0;

      return {
        machine,
        displayMachine: formatMachineLabel(machine),
        jobs,
        broken,
        percent
      };
    })
    .sort((a, b) => b.percent - a.percent);

  function getMachineColor(p) {
    if (p > 12) return "#ff3b3b";
    if (p > 6) return "#ffcc00";
    return "#32ff7e";
  }

  machineChart = new Chart(
    document.getElementById("machineChart"),
    {
      type: "bar",

      data: {
        labels: entries.map(e => e.displayMachine),
        datasets: [
          {
            label: "Breakage %",
            data: entries.map(e => e.percent),
            backgroundColor: entries.map(e => getMachineColor(e.percent)),
            borderRadius: 10
          },
          {
            label: "Total Jobs",
            data: entries.map(e => e.jobs),
            backgroundColor: "#5c8df6",
            borderRadius: 10
          }
        ]
      },

      options: {
        responsive: true,
        maintainAspectRatio: false,

        plugins: {
          legend: {
            labels: {
              color: "#E6F1FF",
              font: { size: 14 }
            }
          },
          title: {
            display: true,
            text: "Machine Analysis",
            color: "#E6F1FF",
            font: { size: 18 }
          },
          tooltip: {
            callbacks: {
              label: function (context) {
                const i = context.dataIndex;
                const e = entries[i];

                if (context.dataset.label === "Breakage %") {
                  return `${e.percent.toFixed(2)}% (${e.broken} broken)`;
                }

                return `Total Jobs: ${e.jobs}`;
              },

              afterBody: function (context) {
                const i = context[0].dataIndex;
                const e = entries[i];
                const machine = e.machine;

                const lines = [];

                const totalBroken = entries.reduce((s, m) => s + m.broken, 0);

                if (totalBroken > 0) {
                  const workflow = ((e.broken / totalBroken) * 100).toFixed(1);
                  lines.push("");
                  lines.push(`Workflow Impact: ${workflow}%`);
                }

                if (data.machineReasons && data.machineReasons[machine]) {
                  const reasons = data.machineReasons[machine];

                  const total = Object.values(reasons)
                    .reduce((a, b) => a + b, 0);

                  lines.push("");
                  lines.push("Top Reasons");

                  Object.entries(reasons)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 3)
                    .forEach(([r, count]) => {
                      const pct = total > 0 ? ((count / total) * 100).toFixed(1) : "0.0";
                      lines.push(`${r} – ${count} (${pct}%)`);
                    });
                }

                return lines;
              }
            }
          }
        },

        scales: {
          x: {
            ticks: {
              color: "rgba(255,255,255,0.75)",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" }
          },

          y: {
            beginAtZero: true,
            ticks: {
              color: "#E6F1FF",
              font: { size: 13 }
            },
            grid: { color: "rgba(255,255,255,0.05)" },
            title: {
              display: true,
              text: "Breakage %",
              color: "#E6F1FF",
              font: { size: 14 }
            }
          }
        }
      }
    }
  );
}

/* =====================================================
   AUTO REFRESH EVERY 5 MINUTES (ONLY LIVE MODE)
   ===================================================== */
setInterval(() => {
  if (isLiveMode || currentDate === null) {
    console.log("Auto refreshing dashboard (Live mode)...");
    loadDashboard();
  } else {
    console.log("Auto refresh skipped (History mode active)");
  }
}, 3 * 60 * 1000);

/* =====================================================
   FILTER LISTENERS (DATE + LIVE)
   ===================================================== */
document.addEventListener("DOMContentLoaded", () => {
  const dateInput = document.getElementById("historyDate");

  if (dateInput) {
    dateInput.addEventListener("change", function () {
      if (!this.value) {
        currentDate = null;
        isLiveMode = true;
      } else {
        const parts = this.value.split("-");
        currentDate =
          parseInt(parts[1], 10) + "/" +
          parseInt(parts[2], 10) + "/" +
          parts[0];

        isLiveMode = false;
      }

      console.log("Date changed to:", currentDate, "Live Mode:", isLiveMode);
      loadDashboard();
    });
  }

  const liveBtn = document.getElementById("liveBtn");

  if (liveBtn) {
    liveBtn.addEventListener("click", function () {
      currentDate = null;
      isLiveMode = true;

      const dateEl = document.getElementById("historyDate");
      if (dateEl) dateEl.value = "";

      console.log("Switched back to Live Mode");
      loadDashboard();
    });
  }
});

/* =====================================================
   FORCE REBUILD ALL CHARTS
   ===================================================== */
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

  buildTrendChart(currentMode === "machine" ? dashboardMachine : dashboardProcessed);
  buildReasonChart(dashboardData);
  buildMachineChart(dashboardData);
  buildFlowChart(dashboardProcessed);
}

/* =====================================================
   NAVIGATION
   ===================================================== */
function goBack() {
  window.location.href = "index.html";
}