const API_URL =
  "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

/* ELEMENTS */
const elActive = document.getElementById("activeHolds");
const elEval = document.getElementById("evaluatedCount");
const elCoverage = document.getElementById("coveragePct");
const elUpdated = document.getElementById("lastUpdated");

const reasonWrap = document.getElementById("reasonTable").querySelector("tbody");
const sentWrap = document.getElementById("sentBackTable").querySelector("tbody");

/* LOAD */
async function loadTracker() {
  try {
    const res = await fetch(API_URL, { cache: "no-store" });
    const data = await res.json();

    if (data.error) {
      console.error(data.error);
      return;
    }

    renderStats(data);
    renderReasons(data.byReason || {});
    renderSentBack(data.bySentBack || {});
    renderUpdated(data.updatedAt);

  } catch (err) {
    console.error("Tracker load failed", err);
  }
}

/* STATS */
function renderStats(data) {
  elActive.textContent = data.active ?? 0;
  elEval.textContent = data.evaluated ?? 0;

  const coverage =
    data.total > 0 ? ((data.evaluated / data.total) * 100).toFixed(1) : "0.0";

  elCoverage.textContent = `${coverage}%`;
}

/* LAST UPDATED */
function renderUpdated(ts) {
  if (!ts) {
    elUpdated.textContent = "Last updated —";
    return;
  }

  const d = new Date(ts);
  elUpdated.textContent = `Last updated: ${d.toLocaleString()}`;
}

/* REASONS */
function renderReasons(reasons) {
  reasonWrap.innerHTML = "";

  const entries = Object.entries(reasons);

  if (!entries.length) {
    reasonWrap.innerHTML =
      `<tr><td colspan="3" class="empty">No reason data available</td></tr>`;
    return;
  }

  const total = entries.reduce((s, [, v]) => s + v, 0);

  entries
    .sort((a, b) => b[1] - a[1])
    .forEach(([reason, count]) => {
      const pct = ((count / total) * 100).toFixed(1);

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${reason}</td>
        <td>${count}</td>
        <td>${pct}%</td>
      `;
      reasonWrap.appendChild(tr);
    });
}

/* SENT BACK TO */
function renderSentBack(buckets) {
  sentWrap.innerHTML = "";

  const entries = Object.entries(buckets);

  if (!entries.length) {
    sentWrap.innerHTML =
      `<tr><td colspan="4" class="empty">No investigative data available</td></tr>`;
    return;
  }

  const total = entries.reduce((s, [, v]) => s + v.count, 0);

  entries
    .sort((a, b) => b[1].count - a[1].count)
    .forEach(([dept, info]) => {
      const pct = ((info.count / total) * 100).toFixed(1);

      const oldest = info.lastAddedDate
        ? new Date(info.lastAddedDate).toLocaleDateString()
        : "—";

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${dept}</td>
        <td>${info.count}</td>
        <td>${pct}%</td>
        <td>${oldest}</td>
      `;
      sentWrap.appendChild(tr);
    });
}

/* INIT */
loadTracker();