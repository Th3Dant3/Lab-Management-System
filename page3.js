const API_URL =
  "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

async function loadInvestigation() {
  try {
    const res = await fetch(API_URL, { cache: "no-store" });
    const data = await res.json();

    if (data.error) {
      console.error(data.error);
      return;
    }

    // ======================
    // KPIs
    // ======================
    document.getElementById("activeHolds").textContent = data.active ?? 0;
    document.getElementById("evaluatedCount").textContent = data.evaluated ?? 0;

    const coverage =
      data.total > 0
        ? ((data.evaluated / data.total) * 100).toFixed(1)
        : "0.0";

    document.getElementById("coveragePct").textContent = `${coverage}%`;

    document.getElementById("lastUpdated").textContent =
      "Last updated: " + new Date(data.updatedAt).toLocaleString();

    // ======================
    // REASONS
    // ======================
    const reasonBody = document.querySelector("#reasonTable tbody");
    reasonBody.innerHTML = "";

    const reasons = data.byReason || {};
    const totalEvaluated = data.evaluated || 0;

    if (Object.keys(reasons).length === 0) {
      reasonBody.innerHTML =
        `<tr><td colspan="3" class="empty">No reason data available</td></tr>`;
    } else {
      Object.entries(reasons)
        .sort((a, b) => b[1] - a[1])
        .forEach(([reason, count]) => {
          const pct =
            totalEvaluated > 0
              ? ((count / totalEvaluated) * 100).toFixed(1)
              : "0.0";

          reasonBody.insertAdjacentHTML(
            "beforeend",
            `<tr>
              <td>${reason}</td>
              <td>${count}</td>
              <td>${pct}%</td>
            </tr>`
          );
        });
    }

    // ======================
    // SENT BACK TO
    // ======================
    const sentBody = document.querySelector("#sentBackTable tbody");
    sentBody.innerHTML = "";

    const sentBack = data.bySentBack || {};

    if (Object.keys(sentBack).length === 0) {
      sentBody.innerHTML =
        `<tr><td colspan="4" class="empty">No investigative data available</td></tr>`;
    } else {
      Object.entries(sentBack)
        .sort((a, b) => b[1].count - a[1].count)
        .forEach(([dept, obj]) => {
          const pct =
            totalEvaluated > 0
              ? ((obj.count / totalEvaluated) * 100).toFixed(1)
              : "0.0";

          const oldest = obj.lastAddedDate
            ? new Date(obj.lastAddedDate).toLocaleDateString()
            : "—";

          sentBody.insertAdjacentHTML(
            "beforeend",
            `<tr>
              <td>${dept}</td>
              <td>${obj.count}</td>
              <td>${pct}%</td>
              <td>${oldest}</td>
            </tr>`
          );
        });
    }

  } catch (err) {
    console.error("Investigation load failed", err);
  }
}

loadInvestigation();