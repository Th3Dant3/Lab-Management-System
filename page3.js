const API_URL = "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

const activeEl   = document.getElementById("activeVal");
const evalEl     = document.getElementById("evalVal");
const compEl     = document.getElementById("compVal");
const updatedEl  = document.getElementById("updated");
const tbody      = document.getElementById("breakdownBody");
const detail     = document.getElementById("detail");

fetch(API_URL)
  .then(r => r.json())
  .then(data => {

    const active     = Number(data.active || 0);
    const evaluated  = Number(data.evaluated || 0);
    const total      = Number(data.totalJobs || (active + evaluated));

    const completion = total
      ? ((evaluated / total) * 100).toFixed(2)
      : "0.00";

    activeEl.textContent  = active;
    evalEl.textContent    = evaluated;
    compEl.textContent    = `${completion}%`;
    updatedEl.textContent = `Last updated: ${new Date(data.updatedAt).toLocaleString()}`;

    tbody.innerHTML = "";

    /* ===============================
       SORT HIGH → LOW BY COUNT
       =============================== */
    const sortedBuckets = Object.entries(data.bySentBack || {})
      .sort((a, b) => b[1] - a[1]); // 🔥 KEY LINE

    sortedBuckets.forEach(([key, count]) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${key}</td>
        <td><strong>${count}</strong></td>
      `;
      tr.onclick = () => showDetail(key, count, evaluated);
      tbody.appendChild(tr);
    });

  })
  .catch(err => {
    console.error(err);
    activeEl.textContent = "ERR";
    tbody.innerHTML = `<tr><td colspan="2" style="color:red">Failed to load data</td></tr>`;
  });

function showDetail(bucket, count, totalEvaluated) {
  const pct = totalEvaluated
    ? ((count / totalEvaluated) * 100).toFixed(1)
    : "0.0";

  detail.style.display = "block";
  detail.innerHTML = `
    <h3>${bucket}</h3>
    <p>Jobs in this bucket: <strong>${count}</strong></p>
    <p>${pct}% of all evaluated jobs</p>
  `;
}
