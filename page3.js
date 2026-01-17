/* =====================================================
   INVESTIGATION HOLD TRACKER — PAGE 3
   ===================================================== */

const API_URL =
  "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

/* INIT */
document.addEventListener("DOMContentLoaded", loadPage3);

/* =====================================================
   MAIN LOAD
   ===================================================== */

async function loadPage3() {
  try {
    const res = await fetch(API_URL, { cache: "no-store" });
    const data = await res.json();

    if (data.error) {
      console.error("API Error:", data.error);
      return;
    }

    updateKPIs(data);
    renderReasons(data.byReason);
    renderSentBack(data.bySentBack);

  } catch (err) {
    console.error("Failed to load investigation data", err);
  }
}

/* =====================================================
   KPI UPDATE
   ===================================================== */

function updateKPIs(data) {
  const active = Number(data.active || 0);
  const evaluated = Number(data.evaluated || 0);
  const total = Number(data.total || 0);

  const coverage =
    total > 0 ? ((evaluated / total) * 100).toFixed(1) : "0.0";

  document.getElementById("activeHolds").textContent = active;
  document.getElementById("evaluatedCount").textContent = evaluated;
  document.getElementById("coveragePct").textContent = `${coverage}%`;

  document.getElementById("lastUpdated").textContent =
    data.updatedAt
      ? "Last updated: " + new Date(data.updatedAt).toLocaleString()
      : "—";
}

/* =====================================================
   REASONS (LEFT PANEL)
   ===================================================== */

function renderReasons(byReason) {
  const container = document.getElementById("reasonList");
  container.innerHTML = "";

  const entries = Object.entries(byReason || {});
  if (!entries.length) {
    container.innerHTML =
      `<div class="empty">No reason data available</div>`;
    return;
  }

  const total = entries.reduce((s, [, c]) => s + Number(c), 0);

  entries
    .sort((a, b) => b[1] - a[1])
    .forEach(([reason, count]) => {
      const pct =
        total > 0 ? ((count / total) * 100).toFixed(1) : "0.0";

      const row = document.createElement("div");
      row.className = "reason-row";
      row.innerHTML = `
        <div class="reason-title">
          <span>${reason}</span>
          <span>${count} (${pct}%)</span>
        </div>
        <div class="reason-bar">
          <span style="width:${pct}%"></span>
        </div>
      `;

      container.appendChild(row);
    });
}

/* =====================================================
   SENT BACK TO (RIGHT PANEL)
   ===================================================== */

function renderSentBack(bySentBack) {
  const tbody = document.getElementById("sentBackTable");
  tbody.innerHTML = "";

  const entries = Object.entries(bySentBack || {});
  if (!entries.length) {
    tbody.innerHTML =
      `<tr><td colspan="4" class="empty">No investigative data available</td></tr>`;
    return;
  }

  const total = entries.reduce((s, [, v]) => s + Number(v.count || 0), 0);

  entries
    .sort((a, b) => b[1].count - a[1].count)
    .forEach(([dept, info]) => {
      const pct =
        total > 0
          ? ((info.count / total) * 100).toFixed(1)
          : "0.0";

      const oldest =
        info.lastAddedDate
          ? new Date(info.lastAddedDate).toLocaleDateString()
          : "—";

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${dept}</td>
        <td>${info.count}</td>
        <td>${pct}%</td>
        <td>${oldest}</td>
      `;

      tbody.appendChild(tr);
    });
}