const API_URL = "https://script.google.com/macros/s/AKfycbyI2YqO9wXZ4-v34OXqqxD-yCYe3Dnly1-d_cf9mYtkcVkZoXeWhgQ8u6WbgY-lvIQQjg/exec";

let chart = null;
let lastData = null;
let currentHour = null;
let currentDate = null;
let trendChart = null;

let refreshTimerHandle = null;
let isRefreshing = false;
let isTrendRefreshing = false;

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

  const today = new Date().toISOString().split("T")[0];
  const dateInput = document.getElementById("dateFilter");

  if (dateInput) dateInput.value = today;
  currentDate = today;

  updateRefreshStatus();

  loadData(); // ✅ no await

  startAutoRefresh();
});

/* ================= MODE HELPERS ================= */

function getTodayString() {
  return new Date().toISOString().split("T")[0];
}

function isTodayMode() {
  return !currentDate || currentDate === getTodayString();
}

function updateRefreshStatus() {
  const statusEl = document.getElementById("refreshStatus");
  if (!statusEl) return;

  if (document.visibilityState !== "visible") {
    statusEl.textContent = "Paused";
    statusEl.style.color = "#94a3b8";
    return;
  }

  if (isTodayMode()) {
    statusEl.textContent = "Live";
    statusEl.style.color = "#22c55e";
  } else {
    statusEl.textContent = "History";
    statusEl.style.color = "#facc15";
  }
}

/* ================= SMART REFRESH ================= */

function startAutoRefresh() {
  const interval = 10 * 60 * 1000;

  if (refreshTimerHandle) {
    clearInterval(refreshTimerHandle);
    refreshTimerHandle = null;
  }

  refreshTimerHandle = setInterval(async () => {
    if (document.visibilityState !== "visible") {
      updateRefreshStatus();
      return;
    }

    if (!isTodayMode()) {
      console.log("⏸ Viewing past date — no auto refresh");
      updateRefreshStatus();
      return;
    }

    console.log("🔄 Auto Refresh (Today Only)");
    updateRefreshStatus();

    await loadData();
   
  }, interval);

  document.addEventListener("visibilitychange", async () => {
    updateRefreshStatus();

    if (document.visibilityState !== "visible") return;
    if (!isTodayMode()) return;
    if (isRefreshing || isTrendRefreshing) return;

    console.log("👀 User returned → safe refresh");
    await loadData();
    
  });
}

/* ================= TrendTab Swtich ================= */

function switchTab(tab) {

  const incoming = document.getElementById("incomingTab");
  const trend = document.getElementById("trendTab");

  const btnIncoming = document.getElementById("tabIncomingBtn");
  const btnTrend = document.getElementById("tabTrendBtn");

  // reset buttons
  btnIncoming.classList.remove("active");
  btnTrend.classList.remove("active");

  if (tab === "incoming") {
    incoming.style.display = "block";
    trend.style.display = "none";
    btnIncoming.classList.add("active");
  } else {
    incoming.style.display = "none";
    trend.style.display = "block";
    btnTrend.classList.add("active");

    // 🔥 ONLY LOAD TREND WHEN TAB IS OPENED
    if (!trendChart) {
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
  if (q.includes("china rush")) return "#F97316";
  if (q.includes("standard")) return "#3B82F6";
  if (q.includes("overnight")) return "#FACC15";
  if (q.includes("beast")) return "#A855F7";
  if (q.includes("frame only")) return "#22C55E";
  if (q.includes("test")) return "#2DD4BF";
  if (q.includes("echo")) return "#FB923C";

  return "#6fb3ff";
}

function brightenColor(hex, percent = 30) {
  const num = parseInt(hex.replace("#", ""), 16);
  const r = Math.min(255, (num >> 16) + percent);
  const g = Math.min(255, ((num >> 8) & 0x00FF) + percent);
  const b = Math.min(255, (num & 0x0000FF) + percent);
  return `rgb(${r},${g},${b})`;
}

function toBarFill(color) {
  return `linear-gradient(90deg, ${color}, ${brightenColor(color, 40)})`;
}


console.log("🚀 loadData called");

/* ================= LOAD ================= */

async function loadData() {
  if (isRefreshing) return;
  isRefreshing = true;

  try {
    let url = API_URL;

    if (currentDate) {

  const d = new Date(currentDate);

  const formatted =
    (d.getMonth() + 1).toString().padStart(2, "0") + "/" +
    d.getDate().toString().padStart(2, "0") + "/" +
    d.getFullYear();

  url += "?date=" + encodeURIComponent(formatted);
}

    const res = await fetch(url, { cache: "no-store" });

const text = await res.text();
console.log("🔥 RAW RESPONSE:", text);

let json;
try {
  json = JSON.parse(text);
} catch (e) {
  console.error("❌ JSON PARSE FAILED", e);
  return;
}

console.log("✅ PARSED JSON:", json);

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
  } catch (err) {
    console.error("loadData error:", err);
  } finally {
    isRefreshing = false;
    updateRefreshStatus();
  }
}



/* ================= DATE FILTER ================= */

async function applyDateFilter() {
  const input = document.getElementById("dateFilter");
  const value = input.value;
  if (!value) return;

  currentDate = new Date(value).toISOString().split("T")[0];
  currentHour = null;

  console.log("📅 Filter Applied:", currentDate);

  updateRefreshStatus();
  await loadData();
}

async function resetDateFilter() {
  const today = getTodayString();

  currentDate = today;
  currentHour = null;

  const input = document.getElementById("dateFilter");
  if (input) input.value = today;

  console.log("🔄 Reset to Today");

  updateRefreshStatus();
  await loadData();
  
}

/* ================= KPI ================= */

function renderKPI(data, grandTotal) {
  const grid = document.getElementById("kpiGrid");
  if (!grid) return;

  grid.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const value = data[dept]._total || 0;
    const pct = grandTotal > 0 ? ((value / grandTotal) * 100).toFixed(1) : "0.0";

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
  if (!container) return;

  container.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const deptTotal = data[dept]._total || 0;

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
      const queueTotal = q.total || 0;
      const pct = deptTotal > 0 ? ((queueTotal / deptTotal) * 100).toFixed(1) : 0;
      const color = getQueueColor(queue);

      const row = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML = `
        <div class="label">${queue}</div>
        <div class="bar"><div class="fill"></div></div>
        <div class="value">${queueTotal}</div>
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
  if (!container) return;

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

      const entries = (data[dept][queue].hours && data[dept][queue].hours[hour]) || [];
      const total = entries.reduce((s, x) => s + (x.value || 0), 0);

      if (total > 0) {
        deptTotal += total;
        rows.push({ queue, total });
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
  const canvas = document.getElementById("hourlyChart");
  if (!canvas) return;

  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      const queueHours = data[dept][queue].hours || {};
      Object.keys(queueHours).forEach(hour => {
        queueHours[hour].forEach(entry => {
          hourMap[hour] += entry.value || 0;
        });
      });
    });
  });

  const labels = HOUR_ORDER.filter(h => hourMap[h] > 0);
  const values = labels.map(h => hourMap[h]);
  const max = values.length ? Math.max(...values) : 0;

  const colors = values.map(v => {
    if (v === max && max > 0) return "rgba(120,255,180,0.9)";
    if (max > 0 && v < max * 0.2) return "rgba(120,160,255,0.3)";
    return "rgba(90,182,255,0.7)";
  });

  const chartData = {
    labels,
    datasets: [{
      data: values,
      backgroundColor: colors,
      borderRadius: 6
    }]
  };

  if (chart) {
    chart.data = chartData;
    chart.update("none");
    return;
  }

  chart = new Chart(canvas, {
    type: "bar",
    data: chartData,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      onClick: (evt, elements, chartInstance) => {
        const points = chartInstance.getElementsAtEventForMode(
          evt,
          "index",
          { intersect: false },
          true
        );
        if (!points.length) return;
        const hour = chartInstance.data.labels[points[0].index];
        showHourBreakdown(hour);
      },
      plugins: {
        legend: { display: false }
      }
    }
  });
}

/* ================= TREND ================= */

async function loadTrend(days = 14) {
  if (isTrendRefreshing) return;
  isTrendRefreshing = true;

  try {
    const res = await fetch(API_URL + "?mode=trend&days=" + days, { cache: "no-store" });
    const json = await res.json();

    if (!json.success) return;

    buildTrendChart(json.trend);
    setTrendInsight(json.trend);
  } catch (err) {
    console.error("Trend error:", err);
  } finally {
    isTrendRefreshing = false;
  }
}

/* ================= TREND CHART ================= */

function buildTrendChart(trend) {
  const canvas = document.getElementById("trendChart");
  if (!canvas) return;

  const labels = trend.map(d => d.date);
  const totals = trend.map(d => d.total);
  const finish = trend.map(d => (d.departments && d.departments["Finish"]) || 0);
  const surface = trend.map(d => (d.departments && d.departments["Surface"]) || 0);
  const specialty = trend.map(d => (d.departments && d.departments["Specialty"]) || 0);
  const frame = trend.map(d => (d.departments && d.departments["Frame Only"]) || 0);

  const trendData = {
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
        borderDash: [6, 6],
        tension: 0.4
      },
      {
        label: "Surface",
        data: surface,
        borderDash: [6, 6],
        tension: 0.4
      },
      {
        label: "Specialty",
        data: specialty,
        borderDash: [6, 6],
        tension: 0.4
      },
      {
        label: "Frame",
        data: frame,
        borderDash: [6, 6],
        tension: 0.4
      }
    ]
  };

  if (trendChart) {
    trendChart.data = trendData;
    trendChart.update("none");
    return;
  }

  trendChart = new Chart(canvas, {
    type: "line",
    data: trendData,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false
    }
  });
}

/* ================= TREND INSIGHT ================= */

function setTrendInsight(trend) {
  const el = document.getElementById("trendInsight");
  if (!el || !trend || trend.length < 2) return;

  let peak = trend[0];
  trend.forEach(d => {
    if (d.total > peak.total) peak = d;
  });

  const latest = trend[trend.length - 1];
  const prev = trend[trend.length - 2];

  let change = 0;
  if (prev.total > 0) {
    change = ((latest.total - prev.total) / prev.total * 100).toFixed(1);
  } else {
    change = latest.total > 0 ? "100.0" : "0.0";
  }

  const trendIcon = Number(change) > 0 ? "📈" : "📉";
  el.textContent = `Peak ${peak.date} (${peak.total}) | ${trendIcon} ${change}%`;
}

/* ================= INSIGHT ================= */

function setChartInsight(data) {
  const insight = document.getElementById("chartInsight");
  if (!insight) return;

  const hourMap = {};
  HOUR_ORDER.forEach(h => { hourMap[h] = 0; });

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    Object.keys(data[dept]).forEach(queue => {
      if (queue === "_total") return;

      const queueHours = data[dept][queue].hours || {};
      Object.keys(queueHours).forEach(hour => {
        queueHours[hour].forEach(entry => {
          hourMap[hour] += entry.value || 0;
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