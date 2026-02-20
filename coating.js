const API_URL = "https://script.google.com/macros/s/AKfycbw0fRQkxOZsFH_V_zFerZ1ZnnHmRUcxLO9BZ9CmmZZv0N8itstkr45EHZD74q6C1PBB/exec";


/* =====================================================
   DASHBOARD VARIABLES
   ===================================================== */

let trendChart = null;
let reasonChart = null;
let machineChart = null;
let flowChart = null;              // 👈 ADD THIS

let dashboardProcessed = null;
let dashboardMachine = null;
let dashboardData = null;

let trendsLoaded = false;
let currentMode = "processed";
let currentFlowMode = "average";



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

  if (tabId === "trends" && !trendsLoaded) {
    buildTrendCharts();
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

  const [processedRes, machineRes] = await Promise.all([
    fetch(`${API_URL}?mode=processed`),
    fetch(`${API_URL}?mode=machine`)
  ]);

  dashboardProcessed = await processedRes.json();
  dashboardMachine = await machineRes.json();

  dashboardData = dashboardProcessed;

  const summary = dashboardData.summary || {};

  // ===== SUMMARY MAPPING (NEW STRUCTURE) =====
  setText("reportDate", summary.reportDate);
  setText("avgTime", summary.avgTime);

const today = summary.todayBreakage || 0;
const yesterday = summary.yesterdayBreakage || 0;
const late = summary.lateBreakage || 0;

const totalBreakage = today + yesterday + late;

// Set TOTAL breakage card
setText("totalBreakage", totalBreakage);



  // If you add these to HTML later they are ready:
  setText("totalJobs", summary.totalJobs);
  setText("rxBreakage", summary.rxBreakage || "0%");
  setText("peakHour", summary.peakHour);

  setText("totalLenses", summary.totalLenses);
  setText("todayLenses", summary.todayLenses);
  setText("yesterdayLenses", summary.yesterdayLenses);
  setText("lateLenses", summary.lateLenses);

  buildHourlyTable(dashboardData.hourly || []);

  const endTime = performance.now();
  console.log("Load time (ms):", Math.round(endTime - startTime));
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;

  // If it's not a number (like "3.56%"), just set it normally
  if (isNaN(value)) {
    el.textContent = value ?? "-";
    return;
  }

  const finalValue = Number(value) || 0;
  const duration = 600;
  const startTime = performance.now();

  function animate(currentTime) {
    const progress = Math.min((currentTime - startTime) / duration, 1);
    const currentValue = Math.floor(progress * finalValue);
    el.textContent = currentValue;

    if (progress < 1) {
      requestAnimationFrame(animate);
    } else {
      el.textContent = finalValue;
      el.classList.add("pulse");
      setTimeout(() => el.classList.remove("pulse"), 600);
    }
  }

  requestAnimationFrame(animate);

}



loadDashboard();

/* =====================================================
   TREND CHART SYSTEM
===================================================== */

function buildTrendCharts() {
  trendsLoaded = true;
  buildTrendChart(dashboardProcessed);
}

function switchTrend(mode) {

  currentMode = mode;

  let data;

  if (mode === "processed") {
    data = dashboardProcessed;
    buildTrendChart(data);
  } 
  else if (mode === "machine") {
    data = dashboardMachine;
    buildTrendChart(data);
  } 
 else if (mode === "flow") {
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

  trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: hours,
      datasets: [{
        label: "Avg Detaper → Coater (mins)",
        data: data.hourly.map(h => h.avgDetaperToCoater || 0),
        borderColor: data.hourly.map(h => h.flowColor),
pointBackgroundColor: data.hourly.map(h => h.flowColor),
pointBorderColor: data.hourly.map(h => h.flowColor),

        borderWidth: 3,
        tension: 0.35,
        fill: false,
        pointRadius: 4
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,

      plugins: {
        legend: {
          labels: { color: "#E6F1FF" }
        },
        title: {
          display: true,
          text: "Process Flow Time (Detaper → Coater)",
          color: "#E6F1FF",
          font: { size: 16 }
        }
      },

      scales: {
        x: {
          ticks: { color: "rgba(255,255,255,0.6)" },
          grid: { color: "rgba(255,255,255,0.05)" }
        },
        y: {
          beginAtZero: true,
          ticks: { color: "#5ad1a3" },
          grid: { color: "rgba(255,255,255,0.05)" },
          title: {
            display: true,
            text: "Average Hours",
            color: "#5ad1a3"
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

  const hours = data.hourly.map(h => h.hour);
  const datasets = [];

  // Collect all reasons dynamically
  const allReasons = new Set();
  data.hourly.forEach(h => {
    Object.keys(h.reasons || {}).forEach(r => allReasons.add(r));
  });

  // Stable color map
  const reasonColorMap = {
    "S-HC Pit": "#FF2C2C",           // Modern Red
    "S-HC Run": "#9b7bff",           // Soft Purple
    "S-HC Contamination": "#ff9f43",  // Professional Orange
	"S-HC HC Suction Cup Marks": "#FF2C2C"         // Modern Red
  };

  const fallbackColor = "#3A86FF";

  allReasons.forEach(reason => {

    datasets.push({
      label: reason,
      data: data.hourly.map(h => h.reasons?.[reason]?.total || 0),
      borderColor: reasonColorMap[reason] || fallbackColor,
      borderWidth: 3,
      tension: 0.35,
      fill: false,
      pointRadius: 4,
      yAxisID: "y"
    });

  });

  // Add coating jobs line
  datasets.push({
    label: "Total Coating Jobs",
    data: data.hourly.map(h => h.coatingJobs || 0),
    borderColor: "#00ff88",
    borderDash: [6, 6],
    borderWidth: 2,
    tension: 0.35,
    fill: false,
    yAxisID: "y1"
  });

  trendChart = new Chart(ctx, {
    type: "line",
    data: { labels: hours, datasets },
    options: {

      responsive: true,
      maintainAspectRatio: false,

      animation: {
        duration: 1000,
        easing: "easeOutQuart"
      },

      interaction: {
        mode: "nearest",
        intersect: true
      },

      plugins: {
        legend: {
          labels: { color: "#E6F1FF" }
        },
        title: {
          display: true,
          text: currentMode === "processed"
            ? "Breakage Processed Time"
            : "Machine Scan Time",
          color: "#E6F1FF",
          font: { size: 16 }
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
            text: "Breakage Count",
            color: "#4da3ff"
          }
        },
        y1: {
          position: "right",
          grid: { drawOnChartArea: false },
          ticks: { color: "#00ff88" },
          title: {
            display: true,
            text: "Total Coating Jobs",
            color: "#00ff88"
          }
        }
      },

      onClick: (event) => {

        const points = trendChart.getElementsAtEventForMode(
          event,
          "nearest",
          { intersect: true },
          true
        );

        if (!points.length) return;

        const index = points[0].index;
        showHourDetails(data.hourly[index]);
      }

    }
  });

}

function getFlowColor(minutes) {
  if (minutes <= 15) return "#00ff88";   // Green
  if (minutes <= 30) return "#ffcc00";   // Yellow
  return "#ff3b3b";                      // Red
}


/* =====================================================
   FLOW CHART SYSTEM (Detaper → Coater)
===================================================== */

function switchFlowMode(mode) {

  currentFlowMode = mode;

  document.querySelectorAll("#avgFlowBtn, #individualFlowBtn")
    .forEach(btn => btn.classList.remove("active"));

  const activeBtn = document.getElementById(
    mode === "average" ? "avgFlowBtn" : "individualFlowBtn"
  );

  if (activeBtn) activeBtn.classList.add("active");

  buildFlowChart(dashboardProcessed);
}


function buildFlowChart(data) {

  const canvas = document.getElementById("flowChart");
  if (!canvas || !data || !data.hourly) return;

  const ctx = canvas.getContext("2d");

  if (flowChart) flowChart.destroy();

  // 🔥 FILTER 6AM–8PM
  const filteredHours = data.hourly.filter(h => {
  const date = new Date(`1970-01-01 ${h.hour}`);
  const hourNum = date.getHours();
  return hourNum >= 6 && hourNum <= 20;
});


  const hours = filteredHours.map(h => h.hour);

 // ===============================
// AVERAGE MODE (Bucket Colored)
// ===============================
if (currentFlowMode === "average") {

  const todayData = [];
  const yesterdayData = [];
  const twoPlusData = [];

  filteredHours.forEach(h => {

    const avg = h.avgDetaperToCoater || 0;

    if (h.bucketBreakdown.sameDay > 0) {
      todayData.push(avg);
      yesterdayData.push(null);
      twoPlusData.push(null);
    }
    else if (h.bucketBreakdown.oneDay > 0) {
      todayData.push(null);
      yesterdayData.push(avg);
      twoPlusData.push(null);
    }
    else if (h.bucketBreakdown.twoPlus > 0) {
      todayData.push(null);
      yesterdayData.push(null);
      twoPlusData.push(avg);
    }
    else {
      todayData.push(null);
      yesterdayData.push(null);
      twoPlusData.push(null);
    }
  });

  flowChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: hours,
      datasets: [

        {
          label: "Today (Flow)",
          data: todayData,
          borderColor: "#32ff7e",
          backgroundColor: "rgba(50,255,126,0.15)",
          borderWidth: 3,
          tension: 0.35,
          fill: true,
          spanGaps: true
        },

        {
          label: "Yesterday (Flow)",
          data: yesterdayData,
          borderColor: "#ffd32a",
          backgroundColor: "rgba(255,211,42,0.15)",
          borderWidth: 3,
          tension: 0.35,
          fill: true,
          spanGaps: true
        },

        {
          label: "2+ Days (Flow)",
          data: twoPlusData,
          borderColor: "#ff3f34",
          backgroundColor: "rgba(255,63,52,0.15)",
          borderWidth: 3,
          tension: 0.35,
          fill: true,
          spanGaps: true
        }

      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: "#E6F1FF" } },
        title: {
          display: true,
          text: "Average Detaper → Coater Time",
          color: "#E6F1FF"
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { color: "#4da3ff" },
          title: {
            display: true,
            text: "Minutes",
            color: "#4da3ff"
          }
        },
        x: {
          ticks: { color: "rgba(255,255,255,0.6)" }
        }
      }
    }
  });

  return;
}


  // ===============================
  // INDIVIDUAL MODE (Severity Colors)
  // ===============================

  const datasets = [];

  filteredHours.forEach(hourObj => {

    (hourObj.flowPoints || []).forEach(p => {

      datasets.push({
  label: "",   // remove RX label
        data: [{
          x: hourObj.hour,
          y: p.flow,
          rx: p.rx,
          machine: p.machine,
          reason: p.reason,
          flow: p.flow
        }],
        backgroundColor: getFlowColor(p.flow),
        borderColor: getFlowColor(p.flow),
        pointRadius: 8,
        pointHoverRadius: 11
      });

    });

  });

  flowChart = new Chart(ctx, {
  type: "scatter",
  data: {
  labels: hours,
  datasets
},

    options: {
      responsive: true,
      maintainAspectRatio: false,

      interaction: {
        mode: "nearest",
        intersect: true
      },

plugins: {

  title: {
    display: true,
    text: "Flow Severity (Average Time per RX)",
    color: "#E6F1FF",
    font: { size: 16, weight: "600" },
    padding: { bottom: 10 }
  },

  legend: {
  display: true,
  position: "top",
  align: "center",
  labels: {
    color: "#FFFFFF",
    font: {
      size: 13,
      weight: "600"
    }, 
    generateLabels: function(chart) {
      return [
        {
          text: "0–15m",
          fillStyle: "#00ff88",
          strokeStyle: "#00ff88",
          lineWidth: 0,
          hidden: false,
          index: 0
        },
        {
          text: "16–30m",
          fillStyle: "#f1c40f",
          strokeStyle: "#f1c40f",
          lineWidth: 0,
          hidden: false,
          index: 1
        },
        {
          text: "31m+",
          fillStyle: "#ff3b3b",
          strokeStyle: "#ff3b3b",
          lineWidth: 0,
          hidden: false,
          index: 2
        }
      ];
    },
    usePointStyle: true,
    pointStyle: "circle",
    color: "#E6F1FF",
    padding: 20
  }
},

  tooltip: {
    callbacks: {
      title: function(context) {
        return context[0].label;
      },
      label: function(context) {
        const p = context.raw;
        const totalMinutes = Number(p.flow) || 0;

        const hrs = Math.floor(totalMinutes / 60);
        const mins = Math.round(totalMinutes % 60);

        let formatted = hrs > 0
          ? (mins > 0 ? `${hrs}h ${mins}m` : `${hrs}h`)
          : `${mins}m`;

        return [
          `RX: ${p.rx}`,
          `Machine: ${p.machine}`,
          `Reason: ${p.reason}`,
          `Flow: ${formatted}`
        ];
      }
    }
  }
      },

      scales: {
        x: {
  type: "category",
  ticks: { color: "rgba(255,255,255,0.6)" }
},

        y: {
          beginAtZero: true,
          ticks: { color: "#4da3ff" },
          title: {
            display: true,
            text: "Minutes",
            color: "#4da3ff"
          }
        }
      },

      onClick: (event, elements, chart) => {

  const points = chart.getElementsAtEventForMode(
    event,
    "nearest",
    { intersect: true },
    true
  );

  // Dot clicked
  if (points.length) {
    const datasetIndex = points[0].datasetIndex;
    const point = chart.data.datasets[datasetIndex].data[0];
    showFlowDetails(point);
    return;
  }

  // Detect hour by index
  const xScale = chart.scales.x;
  const clickedIndex = xScale.getValueForPixel(event.x);

  if (clickedIndex == null) return;

  const hourLabel = chart.data.labels[clickedIndex];

  const hourData = filteredHours.find(h => h.hour === hourLabel);

  if (hourData) {
    showFlowHourDetails(hourData);
  }
}

    }
  });
}


// =============================================
// FLOW MODAL DETAIL (INDIVIDUAL)
// =============================================
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
      <p><strong>Breakage Reason:</strong> ${point.reason}</p>
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

    tr.innerHTML = `
      <td>${row.hour}</td>
      <td>${row.totalBroken || 0}</td>
      <td>${row.coatingJobs || 0}</td>
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

  if (!hourData) return;

  let machineHTML = "";

  Object.entries(hourData.machines || {}).forEach(([machine, stats]) => {

    let reasonHTML = "";

    const reasonsObj = stats.reasons || {};

    Object.entries(reasonsObj).forEach(([reason, rStats]) => {

      reasonHTML += `
        <div style="margin-top:8px; padding-left:12px;">
          • <strong>${reason}</strong> — ${rStats.total || 0}
          <br>
          <span style="font-size:12px; opacity:.7;">
            Today: ${rStats.sameDay || 0} |
            Yesterday: ${rStats.oneDay || 0} |
            2+ Late: ${rStats.twoPlus || 0}
          </span>
        </div>
      `;
    });

    machineHTML += `
      <div style="margin-bottom:16px; padding:14px; background:#243b63; border-radius:8px;">
        <strong>${machine}</strong><br>
        Total: ${stats.total || 0}<br>
        Today: ${stats.sameDay || 0}<br>
        Yesterday: ${stats.oneDay || 0}<br>
        2+ Late: ${stats.twoPlus || 0}
        ${reasonHTML}
      </div>
    `;
  });

  modalBody.innerHTML = `
    <h2>${hourData.hour}</h2>
    <h3>Total Broken: ${hourData.totalBroken || 0}</h3>
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
    .sort((a, b) => b[1] - a[1]);

  reasonChart = new Chart(
    document.getElementById("reasonChart"),
    {
      type: "bar",
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{
          label: "Total Breakage",
          data: entries.map(e => e[1]),
          backgroundColor: entries.map((e, i) =>
            ["#4da3ff", "#9b7bff", "#ff9f43", "#00ff88"][i % 4]
          ),
          borderRadius: 8
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } }
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
    .sort((a, b) => b[1] - a[1]);

  machineChart = new Chart(
    document.getElementById("machineChart"),
    {
      type: "bar",
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{
          label: "Total Breakage",
          data: entries.map(e => e[1]),
          backgroundColor: entries.map((e, i) =>
            ["#4da3ff", "#9b7bff", "#ff9f43", "#00ff88", "#5ad1a3"][i % 5]
          ),
          borderRadius: 10
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    }
  );
}

// 🔄 AUTO REFRESH EVERY 5 MINUTES
setInterval(() => {
  console.log("Auto refreshing dashboard...");
  loadDashboard();
}, 5 * 60 * 1000);

function goBack() {
  window.location.href = "index.html";

}
