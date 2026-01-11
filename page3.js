const API_URL =
  "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

const REFRESH_MS = 15000;

let lastActive = null;

async function loadHoldData() {
  const activeEl = document.getElementById("activeCount");
  const evalEl   = document.getElementById("evaluatedCount");
  const pctEl    = document.getElementById("completionPct");
  const tableEl  = document.getElementById("holdTable");
  const banner   = document.getElementById("stateBanner");
  const header   = document.getElementById("stateHeader");
  const timeEl   = document.getElementById("lastUpdated");

  try {
    const res = await fetch(API_URL, { cache:"no-store" });
    const data = await res.json();

    const active = Number(data.active || 0);
    const evaluated = Number(data.evaluated || 0);
    const total = active + evaluated;

    /* ===== COMPLETION (TRUTHFUL) ===== */
    let pct;
    if (total === 0) {
      pct = 100;
    } else if (active > 0) {
      pct = Math.min(99, Math.floor((evaluated / total) * 100));
    } else {
      pct = 100;
    }

    /* ===== STATE LOGIC ===== */
    let state, stateText, stateClass, textClass;

    if (active === 0) {
      state = "ALL CLEAR";
      stateText = "All investigative holds evaluated";
      stateClass = "state-green";
      textClass = "state-text-green";
    } else if (lastActive !== null && active > lastActive) {
      state = "ACTION REQUIRED";
      stateText = "Backlog increasing — intervene now";
      stateClass = "state-red";
      textClass = "state-text-red";
    } else {
      state = "ATTENTION";
      stateText = "Active holds pending review";
      stateClass = "state-yellow";
      textClass = "state-text-yellow";
    }

    /* ===== APPLY STATE ===== */
    header.className = `header ${stateClass}`;
    banner.innerHTML = `${state}
      <span class="badge ${textClass}">${active} Active</span>`;

    if (state === "ATTENTION") {
      header.classList.add("attention-pulse");
    } else {
      header.classList.remove("attention-pulse");
    }

    /* ===== COUNTERS ===== */
    activeEl.textContent = active;
    evalEl.textContent = evaluated;
    pctEl.textContent = `${pct}%`;

    if (lastActive !== null && active !== lastActive) {
      activeEl.style.transform = "scale(1.1)";
      setTimeout(() => activeEl.style.transform = "scale(1)", 300);
    }

    lastActive = active;

    /* ===== TIME ===== */
    if (data.updatedAt) {
      timeEl.textContent =
        "Last updated: " +
        new Date(data.updatedAt).toLocaleString();
    }

    /* ===== BREAKDOWN ===== */
    tableEl.innerHTML = "";

    const entries = Object.entries(data.bySentBack || {})
      .sort((a,b) => b[1] - a[1]);

    if (entries.length === 0) {
      tableEl.innerHTML =
        `<tr><td colspan="2">No evaluated holds</td></tr>`;
      return;
    }

    entries.forEach(([dest, count], idx) => {
      const tr = document.createElement("tr");

      if (idx === 0 && active > 0) {
        tr.classList.add("breakdown-hot");
      }

      tr.style.cursor = "pointer";
      tr.innerHTML = `
        <td>${dest}</td>
        <td>${count}</td>
      `;

      tr.onclick = () =>
        alert(`Drill-down: ${dest}\n(Next: RX modal)`);

      tableEl.appendChild(tr);
    });

  } catch (err) {
    console.error(err);
    activeEl.textContent = "0";
    evalEl.textContent = "0";
    pctEl.textContent = "—";
    tableEl.innerHTML =
      `<tr><td colspan="2">Data unavailable</td></tr>`;
  }
}

loadHoldData();
setInterval(loadHoldData, REFRESH_MS);