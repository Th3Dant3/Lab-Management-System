/* ============================================================
   INCOMING JOBS DASHBOARD — UPGRADED JS
   ============================================================ */

const API_URL = "https://script.google.com/macros/s/AKfycbyI2YqO9wXZ4-v34OXqqxD-yCYe3Dnly1-d_cf9mYtkcVkZoXeWhgQ8u6WbgY-lvIQQjg/exec";

let chart      = null;
let trendChart = null;
let lastData   = null;
let currentHour  = null;
let currentDate  = null;
let isRenderingChart = false;
let refreshTimer = null;

/* ================= CONSTANTS ================= */

const QUEUE_COLORS = {
  /* RUSH */
  "Surface Rush Delivery":    "#00E5FF",
  "Rush Delivery Queue":      "#00E5FF",
  "In Rush Delivery Queue":   "#00E5FF",
  /* CHINA */
  "Surface China Rush Delivery":    "#F97316",
  "China Rush Queue":               "#F97316",
  "In China Rush Queue":            "#F97316",
  "In China Rush Delivery Queue":   "#F97316",
  /* STANDARD */
  "Surface Standard":     "#3B82F6",
  "Standard Queue":       "#3B82F6",
  "In Standard Queue":    "#3B82F6",
  "Fin Standard":         "#93C5FD",
  /* OVERNIGHT */
  "Overnight Queue":              "#FACC15",
  "In Overnight Queue":           "#FACC15",
  "Surface Overnight Delivery":   "#FACC15",
  "Fin Overnight Delivery":       "#A3A3A3",
  /* FINISH */
  "Fin Rush Delivery":  "#86EFAC",
  /* BEAST */
  "Beast Queue":          "#A855F7",
  "In Beast Queue":       "#A855F7",
  "Beast Surface Queue":  "#D946EF",
  "Beast Finish Queue":   "#14B8A6",
  /* FRAME */
  "Frame Only Queue":       "#22C55E",
  "In Frame Only Queue":    "#22C55E",
  "Frame Only Test Queue":  "#15803D",
  /* OTHER */
  "Test Jobs Queue":  "#2DD4BF",
  "Echo Queue":       "#FB923C",
  /* TOTAL */
  "All Queued Jobs":  "#10B981"
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
  document.getElementById("dateFilter").value = today;
  currentDate = today;

  loadData();

  // Auto-refresh every 60s when viewing today
  refreshTimer = setInterval(() => {
    const isToday = !currentDate || currentDate === new Date().toISOString().split("T")[0];
    if (isToday) loadData();
  }, 60_000);
});

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
  if (q.includes("hold"))          return "#fb7185";
  if (q.includes("surface rush"))  return "#00E5FF";
  if (q.includes("china rush"))    return "#f87171";
  if (q.includes("standard"))      return "#3B82F6";
  if (q.includes("overnight"))     return "#FACC15";
  if (q.includes("beast"))         return "#A855F7";
  if (q.includes("frame only"))    return "#22C55E";
  if (q.includes("test"))          return "#2DD4BF";
  if (q.includes("echo"))          return "#FB923C";
  return "#6fb3ff";
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

async function loadData() {
  if (window._isLoading) return;
  window._isLoading = true;
  updateRefreshStatus("updating");

  try {
    const url = currentDate ? `${API_URL}?date=${currentDate}` : API_URL;
    const res  = await fetch(url);
    const json = await res.json();

    if (!json.success) throw new Error("API returned success=false");

    // Update header meta
    const totalEl   = document.getElementById("totalJobs");
    const updatedEl = document.getElementById("lastUpdated");
    if (totalEl)   totalEl.textContent   = json.total.toLocaleString();
    if (updatedEl) updatedEl.textContent = new Date().toLocaleTimeString([], { hour:'2-digit', minute:'2-digit', second:'2-digit' });

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
    console.error("❌ loadData error:", err);
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

  el.textContent = {
    live:     "Live",
    history:  "History",
    updating: "Updating…"
  }[mode];
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
  grid.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const value = data[dept]._total ?? 0;
    const pct   = grandTotal > 0 ? ((value / grandTotal) * 100).toFixed(1) : "0.0";

    const card = document.createElement("div");
    card.className = "kpi-card";
    card.innerHTML = `
      <div class="kpi-title">${dept}</div>
      <div class="kpi-value">${value.toLocaleString()}</div>
      <div class="kpi-sub">${pct}% of total</div>
    `;
    grid.appendChild(card);
  });
}

/* ================= QUEUE BARS ================= */

function renderQueueBars(data, grandTotal) {
  const container = document.getElementById("container");
  if (!container) return;
  container.innerHTML = "";

  Object.keys(data).forEach(dept => {
    if (dept === "_total") return;

    const deptTotal = data[dept]._total ?? 0;
    const deptBlock = document.createElement("div");
    deptBlock.className = "queue-dept";
    deptBlock.innerHTML = `
      <div class="dept-header">
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

      label.style.color  = color;
      fill.style.background = toBarFill(color);
      fill.style.boxShadow  = toGlow(color);

      // Animate width after paint
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
    <div class="dept-header">
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
    const block = document.createElement("div");
    block.className = "queue-dept";
    block.innerHTML = `
      <div class="dept-header">
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
          data[dept][queue].hours[hour].forEach(entry => {
            hourMap[hour] = (hourMap[hour] || 0) + entry.value;
          });
        });
      });
    });

    const labels = HOUR_ORDER.filter(h => hourMap[h] > 0);
    const values = labels.map(h => hourMap[h]);
    const maxVal = Math.max(...values, 1);

    // Color bars: peak = bright, low = dim, rest = mid
    const colors = values.map(v => {
      const ratio = v / maxVal;
      if (ratio >= 1)    return "rgba(120,255,190,0.92)";
      if (ratio < 0.2)   return "rgba(90,150,255,0.28)";
      return "rgba(90,182,255,0.70)";
    });

    if (chart) {
      chart.data.labels                      = labels;
      chart.data.datasets[0].data            = values;
      chart.data.datasets[0].backgroundColor = colors;
      chart.update("none");
    } else {
      const ctx = canvas.getContext("2d");
      chart = new Chart(ctx, {
        type: "bar",
        data: {
          labels,
          datasets: [{
            data: values,
            backgroundColor: colors,
            borderRadius: 5,
            borderSkipped: false,
          }]
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
              backgroundColor: "rgba(5,13,26,0.92)",
              borderColor: "rgba(90,180,255,0.3)",
              borderWidth: 1,
              titleColor: "#9ed8ff",
              bodyColor: "#e8f2ff",
              titleFont: { family: "'Space Mono', monospace", size: 11 },
              bodyFont:  { family: "'DM Sans', sans-serif",   size: 13 },
              padding: 10,
              callbacks: {
                title: items => items[0].label,
                label: item  => ` ${item.raw.toLocaleString()} jobs`
              }
            }
          },
          scales: {
            x: {
              grid: { color: "rgba(90,160,255,0.06)" },
              ticks: {
                color: "#4a6a90",
                font: { family: "'Space Mono', monospace", size: 10 },
                maxRotation: 45
              }
            },
            y: {
              grid: { color: "rgba(90,160,255,0.08)" },
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
  // Highlight active button
  document.querySelectorAll(".trend-btn").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".trend-btn").forEach(b => {
    if (b.textContent.trim() === `${days}D`) b.classList.add("active");
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

  const labels   = trend.map(d => d.date);
  const totals   = trend.map(d => d.total);
  const finish   = trend.map(d => d.departments?.["Finish"]     || 0);
  const surface  = trend.map(d => d.departments?.["Surface"]    || 0);
  const specialty= trend.map(d => d.departments?.["Specialty"]  || 0);
  const frame    = trend.map(d => d.departments?.["Frame Only"] || 0);

  const ctx = canvas.getContext("2d");
  trendChart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Total",
          data: totals,
          borderColor: "#5ab4ff",
          backgroundColor: "rgba(90,180,255,0.08)",
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
          borderColor: "#86efac",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2
        },
        {
          label: "Surface",
          data: surface,
          borderColor: "#00e5ff",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2
        },
        {
          label: "Specialty",
          data: specialty,
          borderColor: "#a855f7",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2
        },
        {
          label: "Frame",
          data: frame,
          borderColor: "#22c55e",
          borderWidth: 1.5,
          tension: 0.4,
          borderDash: [5,5],
          pointRadius: 2
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
            padding: 16
          }
        },
        tooltip: {
          backgroundColor: "rgba(5,13,26,0.92)",
          borderColor: "rgba(90,180,255,0.3)",
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
          grid: { color: "rgba(90,160,255,0.06)" },
          ticks: {
            color: "#4a6a90",
            font: { family: "'Space Mono', monospace", size: 10 },
            maxRotation: 45
          }
        },
        y: {
          grid: { color: "rgba(90,160,255,0.08)" },
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