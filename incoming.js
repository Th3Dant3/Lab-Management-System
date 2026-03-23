const API_URL = "https://script.google.com/macros/s/AKfycbyI2YqO9wXZ4-v34OXqqxD-yCYe3Dnly1-d_cf9mYtkcVkZoXeWhgQ8u6WbgY-lvIQQjg/exec";

let chart = null;
let lastData = null;
let currentHour = null;
let currentDate = null;
let trendChart = null;
let isRenderingChart = false;

/* ================= COLOR MAP ================= */

const status = document.getElementById("refreshStatus");

const QUEUE_COLORS = {

  /* 🔥 RUSH */
  "Surface Rush Delivery": "#00E5FF",
  "Rush Delivery Queue": "#00E5FF",
  "In Rush Delivery Queue": "#00E5FF",

  /* 🔴 CHINA */
 "Surface China Rush Delivery": "#F97316",
 "China Rush Queue": "#F97316",
 "In China Rush Queue": "#F97316",
 "In China Rush Delivery Queue": "#F97316",

  /* 🔵 STANDARD */
  "Surface Standard": "#3B82F6",
  "Standard Queue": "#3B82F6",
  "In Standard Queue": "#3B82F6",
  "Fin Standard": "#93C5FD",

  /* 🟡 OVERNIGHT */
  "Overnight Queue": "#FACC15",
  "In Overnight Queue": "#FACC15",
  "Surface Overnight Delivery": "#FACC15",
  "Fin Overnight Delivery": "#A3A3A3",

  /* 🟢 FINISH */
  "Fin Rush Delivery": "#86EFAC",

  /* 🟣 BEAST */
  "Beast Queue": "#A855F7",
  "In Beast Queue": "#A855F7",
  "Beast Surface Queue": "#D946EF",
  "Beast Finish Queue": "#14B8A6",

  /* 🟢 FRAME */
  "Frame Only Queue": "#22C55E",
  "In Frame Only Queue": "#22C55E",
  "Frame Only Test Queue": "#15803D",

  /* ⚙️ OTHER */
  "Test Jobs Queue": "#2DD4BF",
  "Echo Queue": "#FB923C",

  /* 🟢 TOTAL */
  "All Queued Jobs": "#10B981"
};

const HOUR_ORDER = [
  "12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM",
  "6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM",
  "12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM",
  "6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"
];

/* ================= INIT ================= */

document.addEventListener("DOMContentLoaded", () => {

  loadData();
  startAutoRefresh();

  const today = new Date().toISOString().split("T")[0];
  document.getElementById("dateFilter").value = today;
  currentDate = today;

});

/* ================= SMART REFRESH ================= */

function startAutoRefresh() {

  const interval = 10 * 60 * 1000;

  setInterval(() => {

    if (document.visibilityState !== "visible") return;

    const today = new Date().toISOString().split("T")[0];

    if (!currentDate || currentDate === today) {

      console.log("🔄 Auto Refresh (Today Only)");
      loadData();

    } else {

      console.log("⏸ Viewing past date — no auto refresh");

    }

    // ✅ ALWAYS UPDATE STATUS HERE
    updateRefreshStatus();

  }, interval);

  document.addEventListener("visibilitychange", () => {

    if (document.visibilityState === "visible") {
      console.log("👀 User returned → refresh");
      loadData();
      updateRefreshStatus(); // ✅ ADD THIS
    }

  });
}

/* ================= TAB SWITCH ================= */

function switchTab(tab) {

  const incoming = document.getElementById("incomingTab");
  const trend = document.getElementById("trendTab");

  const btnIncoming = document.getElementById("tabIncomingBtn");
  const btnTrend = document.getElementById("tabTrendBtn");

  if (!incoming || !trend) return;

  if (btnIncoming) btnIncoming.classList.remove("active");
  if (btnTrend) btnTrend.classList.remove("active");

  if (tab === "incoming") {

    incoming.style.display = "block";
    trend.style.display = "none";

    if (btnIncoming) btnIncoming.classList.add("active");

  } else {

    incoming.style.display = "none";
    trend.style.display = "block";

    if (btnTrend) btnTrend.classList.add("active");

    if (!trendChart) {
      console.log("📊 Loading Trend...");
      loadTrend(14);
    }
  }
}

/* ================= HELPERS ================= */

function getQueueColor(queueName) {
  if (QUEUE_COLORS[queueName]) return QUEUE_COLORS[queueName];

  const q = String(queueName || "").toLowerCase();

  if (q.includes("error") || q.includes("waiting")) return "#FF0000";
  if (q.includes("hold")) return "#FB7185";
  if (q.includes("surface rush")) return "#00E5FF";
  if (q.includes("china rush")) return "#FF3B3B";
  if (q.includes("standard")) return "#3B82F6";
  if (q.includes("overnight")) return "#FACC15";
  if (q.includes("beast")) return "#A855F7";
  if (q.includes("frame only")) return "#22C55E";
  if (q.includes("test")) return "#2DD4BF";
  if (q.includes("echo")) return "#FB923C";

  return "#6fb3ff";
}

function brightenColor(hex, percent = 30) {
  const num = parseInt(hex.replace("#",""),16);
  const r = Math.min(255, (num >> 16) + percent);
  const g = Math.min(255, ((num >> 8) & 0x00FF) + percent);
  const b = Math.min(255, (num & 0x0000FF) + percent);
  return `rgb(${r},${g},${b})`;
}

function toBarFill(color) {
  return `linear-gradient(90deg, ${color}, ${brightenColor(color, 40)})`;
}

/* ================= LOAD ================= */

async function loadData() {

  // 🔒 PREVENT OVERLAP
  if (window.isLoadingData) {
    console.log("⛔ Skipping load — already running");
    return;
  }

  window.isLoadingData = true;
  updateRefreshStatus("updating");

  try {

    let url = API_URL;

    if (currentDate) {
      url += "?date=" + currentDate;
    }

    const res = await fetch(url);
    const json = await res.json();

    if (!json.success) return;

    document.getElementById("totalJobs").textContent = "Total: " + json.total;
    document.getElementById("lastUpdated").textContent =
      "Updated: " + new Date().toLocaleTimeString();

    renderKPI(json.data, json.total);

    if (currentHour) {
      showHourBreakdown(currentHour, json.data);
    } else {
      renderQueueBars(json.data, json.total);
    }

    buildChart(json.data);
    setChartInsight(json.data);

    lastData = json.data;
	// ✅ ADD THIS
updateRefreshStatus();

  } catch (err) {
    console.error(err);
  }

  // 🔓 RELEASE LOCK (ALWAYS RUNS)
  window.isLoadingData = false;
  updateRefreshStatus();
}

/* ================= UpdateStatus ================= */

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

  el.classList.remove("live", "history", "updating");

  if (mode === "live") {
    el.textContent = "Live";
    el.classList.add("live");
  }

  if (mode === "history") {
    el.textContent = "History";
    el.classList.add("history");
  }

  if (mode === "updating") {
    el.textContent = "Updating...";
    el.classList.add("updating");
  }
}


/* ================= Date Filter  ================= */

function applyDateFilter() {

  const input = document.getElementById("dateFilter");
  const value = input.value;

  if (!value) return;

  currentDate = new Date(value)
    .toISOString()
    .split("T")[0];

  currentHour = null;

  console.log("📅 Filter Applied:", currentDate);

  updateRefreshStatus();   // 🔥 ADD THIS
loadData();
}

function resetDateFilter() {

  currentDate = null;
  currentHour = null;

  document.getElementById("dateFilter").value = "";

  console.log("🔄 Reset to Today");

  updateRefreshStatus();   // 🔥 ADD THIS
loadData();
}

/* ================= KPI ================= */

function renderKPI(data, grandTotal) {
  const grid = document.getElementById("kpiGrid");
  grid.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const value = data[dept]._total;
    const pct = ((value / grandTotal) * 100).toFixed(1);

    const card = document.createElement("div");
    card.className = "kpi-card";
    card.innerHTML = `
      <div class="kpi-title">${dept}</div>
      <div class="kpi-value">${value}</div>
      <div class="kpi-sub">${pct}% of total</div>
    `;
    grid.appendChild(card);
  });
}

/* ================= QUEUE ================= */

function renderQueueBars(data, grandTotal) {
  const container = document.getElementById("container");
  container.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const deptTotal = data[dept]._total;

    const deptBlock = document.createElement("div");
    deptBlock.className = "queue-dept";

    deptBlock.innerHTML = `
      <div class="dept-header">
        ${dept} <span>${deptTotal}</span>
      </div>
    `;

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      const q = data[dept][queue];
      const pct = ((q.total / deptTotal) * 100).toFixed(1);
      const color = getQueueColor(queue);

      const row = document.createElement("div");
      row.className = "queue-row";

      row.innerHTML = `
        <div class="label">${queue}</div>
        <div class="bar"><div class="fill"></div></div>
        <div class="value">${q.total}</div>
      `;

      deptBlock.appendChild(row);

      const fill = row.querySelector(".fill");
      const label = row.querySelector(".label");

      label.style.color = color;
      fill.style.background = toBarFill(color);
      fill.style.boxShadow = `0 0 10px ${color}`;

      requestAnimationFrame(() => {
        fill.style.width = pct + "%";
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
    <div class="dept-header">
      Hour: ${hour}
      <span class="reset-link" onclick="resetView()">Reset</span>
    </div>
  `;

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    let deptTotal = 0;
    const rows = [];

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      const entries = data[dept][queue].hours[hour] || [];
      const total = entries.reduce((s,x)=>s+x.value,0);

      if (total > 0) {
        deptTotal += total;
        rows.push({queue,total});
      }
    });

    if (!deptTotal) return;

    const block = document.createElement("div");
    block.className = "queue-dept";

    block.innerHTML = `
      <div class="dept-header">
        ${dept} <span>${deptTotal}</span>
      </div>
    `;

    rows.forEach(r => {
      const color = getQueueColor(r.queue);
      const pct = Math.max((r.total / deptTotal) * 100, 4);

      const row = document.createElement("div");
      row.className = "queue-row";

      row.innerHTML = `
        <div class="label">${r.queue}</div>
        <div class="bar"><div class="fill"></div></div>
        <div class="value">${r.total}</div>
      `;

      const fill = row.querySelector(".fill");
      const label = row.querySelector(".label");

      label.style.color = color;
      fill.style.background = toBarFill(color);
      fill.style.boxShadow = `0 0 10px ${color}`;
      fill.style.width = pct + "%";

      block.appendChild(row);
    });

    container.appendChild(block);
  });
}

/* ================= RESET ================= */

function resetView() {
  currentHour = null;
  if (lastData) renderQueueBars(lastData, lastData._total || 0);
}

/* ================= CHART ================= */

function buildChart(data) {

  if (isRenderingChart) {
    console.log("⛔ Chart update skipped (busy)");
    return;
  }

  isRenderingChart = true;

  try {

    const canvas = document.getElementById("hourlyChart");
    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    const hourMap = {};
    HOUR_ORDER.forEach(h => hourMap[h] = 0);

    Object.keys(data).forEach(dept => {
      if (dept === "_total") return;

      Object.keys(data[dept]).forEach(queue => {
        if (queue === "_total") return;

        Object.keys(data[dept][queue].hours).forEach(hour => {
          data[dept][queue].hours[hour].forEach(entry => {
            hourMap[hour] += entry.value;
          });
        });
      });
    });

    const labels = HOUR_ORDER.filter(h => hourMap[h] > 0);
    const values = labels.map(h => hourMap[h]);
    const max = Math.max(...values);

    const colors = values.map(v => {
      if (v === max) return "rgba(120,255,180,0.9)";
      if (v < max * 0.2) return "rgba(120,160,255,0.3)";
      return "rgba(90,182,255,0.7)";
    });

    /* =========================
       🔥 UPDATE INSTEAD OF DESTROY
    ========================= */

    if (chart) {

      chart.data.labels = labels;
      chart.data.datasets[0].data = values;
      chart.data.datasets[0].backgroundColor = colors;

      chart.update("none"); // smooth, no flicker

    } else {

      chart = new Chart(ctx, {
        type: "bar",
        data: {
          labels,
          datasets: [{
            data: values,
            backgroundColor: colors,
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          onClick: (evt, elements, chartInstance) => {
            const points = chartInstance.getElementsAtEventForMode(
              evt, "index", { intersect: false }, true
            );
            if (!points.length) return;
            const hour = chartInstance.data.labels[points[0].index];
            showHourBreakdown(hour);
          },
          plugins: { legend: { display: false } }
        }
      });

    }

  } catch (err) {
    console.error("Chart error:", err);
  }

  setTimeout(() => {
    isRenderingChart = false;
  }, 50);
}

/* ================= TREND ================= */

async function loadTrend(days = 14) {

  try {

    const res = await fetch(API_URL + "?mode=trend&days=" + days);
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

  const ctx = document.getElementById("trendChart");
  if (!ctx) return;

  // ✅ SAME PATTERN AS HOURLY
  if (trendChart) trendChart.destroy();

  const labels = trend.map(d => d.date);
  const totals = trend.map(d => d.total);

  const finish = trend.map(d => d.departments["Finish"] || 0);
  const surface = trend.map(d => d.departments["Surface"] || 0);
  const specialty = trend.map(d => d.departments["Specialty"] || 0);
  const frame = trend.map(d => d.departments["Frame Only"] || 0);

  trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: totals,
          borderWidth: 3,
          tension: 0.4
        },
        {
          label: "Finish",
          data: finish,
          borderDash: [6,6]
        },
        {
          label: "Surface",
          data: surface,
          borderDash: [6,6]
        },
        {
          label: "Specialty",
          data: specialty,
          borderDash: [6,6]
        },
        {
          label: "Frame",
          data: frame,
          borderDash: [6,6]
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

/* ================= TREND INSIGHT ================= */

function setTrendInsight(trend) {

  const el = document.getElementById("trendInsight");
  if (!el || trend.length < 2) return;

  let peak = trend[0];

  trend.forEach(d => {
    if (d.total > peak.total) peak = d;
  });

  const latest = trend[trend.length - 1];
  const prev = trend[trend.length - 2];

  let change = ((latest.total - prev.total) / prev.total * 100).toFixed(1);

  const trendIcon = change > 0 ? "📈" : "📉";

  el.textContent = `Peak ${peak.date} (${peak.total}) | ${trendIcon} ${change}%`;
}

/* ================= INSIGHT ================= */

function setChartInsight(data) {
  const insight = document.getElementById("chartInsight");
  if (!insight) return;

  const hourMap = {};
  HOUR_ORDER.forEach(h => hourMap[h] = 0);

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      Object.keys(data[dept][queue].hours).forEach(hour => {
        data[dept][queue].hours[hour].forEach(entry => {
          hourMap[hour] += entry.value;
        });
      });
    });
  });

  let peak = null;
  let val = 0;

  Object.keys(hourMap).forEach(h => {
    if (hourMap[h] > val) {
      val = hourMap[h];
      peak = h;
    }
  });

  insight.textContent = peak ? `Peak: ${peak} (${val})` : "";
}

/* ================= NAV ================= */

function goBack() {
  window.location.href = "index.html";
}