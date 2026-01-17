/* =====================================================
   TRUE CURVE / POWER ISSUES — JS
   ===================================================== */

const API_URL =
  "https://script.google.com/macros/s/AKfycbzuUpwsa2VTGjxeSP3wVV-x2Z4yMkGH_eYS86VpFH0CzC2Ii5S5U2ag-S5d89RZ8oHJ/exec";

let selectedYear = "2026";
let currentView = "prod";
let groupMode = "none";

/* INIT */
document.addEventListener("DOMContentLoaded", () => {
  wireTabs();
  loadData();
});

/* =====================================================
   TAB WIRING
   ===================================================== */

function wireTabs() {
  document.querySelectorAll(".tab").forEach(tab => {
    tab.addEventListener("click", () => {
      const group = tab.dataset.group;
      const view = tab.dataset.view;
      const year = tab.dataset.year;

      if (group) groupMode = group;
      if (view) currentView = view;
      if (year) selectedYear = year;

      document
        .querySelectorAll(`.tab[data-${group ? "group" : view ? "view" : "year"}]`)
        .forEach(t => t.classList.remove("active"));

      tab.classList.add("active");
      loadData();
    });
  });
}

/* =====================================================
   LOAD DATA
   ===================================================== */

async function loadData() {
  const tbody = document.getElementById("tc-body");
  tbody.innerHTML = `<tr><td colspan="6" class="empty">Loading…</td></tr>`;

  try {
    const res = await fetch(`${API_URL}?year=${selectedYear}`, { cache: "no-store" });
    const json = await res.json();

    if (json.error) {
      console.error(json.error);
      tbody.innerHTML =
        `<tr><td colspan="6" class="empty">API error</td></tr>`;
      return;
    }

    let rows = json.rows || [];

    // FILTER PROD / TEST
    rows = rows.filter(r =>
      currentView === "prod"
        ? r.status !== "TEST" && r.status !== "Cancel"
        : r.status === "TEST" || r.status === "Cancel"
    );

    renderTable(rows);

  } catch (err) {
    console.error(err);
    tbody.innerHTML =
      `<tr><td colspan="6" class="empty">Failed to load data</td></tr>`;
  }
}

/* =====================================================
   RENDER TABLE
   ===================================================== */

function renderTable(rows) {
  const tbody = document.getElementById("tc-body");
  tbody.innerHTML = "";

  if (!rows.length) {
    tbody.innerHTML =
      `<tr><td colspan="6" class="empty">No data available</td></tr>`;
    return;
  }

  // FLAT LIST
  if (groupMode === "none") {
    rows
      .sort((a, b) => (b.breakage || 0) - (a.breakage || 0))
      .forEach(r => tbody.appendChild(makeRow(r)));
    return;
  }

  // GROUP BY MATERIAL
  const groups = {};
  rows.forEach(r => {
    const key = (r.material || "Unknown").trim();
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  Object.entries(groups)
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([label, items]) => {
      let open = false;

      const header = document.createElement("tr");
      header.className = "group-row";
      header.innerHTML =
        `<td colspan="6">▶ ${label} (${items.length})</td>`;

      tbody.appendChild(header);

      const children = items
        .sort((a, b) => (b.breakage || 0) - (a.breakage || 0))
        .map(r => {
          const tr = makeRow(r);
          tr.style.display = "none";
          tbody.appendChild(tr);
          return tr;
        });

      header.onclick = () => {
        open = !open;
        children.forEach(tr => tr.style.display = open ? "" : "none");
        header.firstChild.textContent =
          `${open ? "▼" : "▶"} ${label} (${items.length})`;
      };
    });
}

/* =====================================================
   ROW BUILDER
   ===================================================== */

function makeRow(r) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${r.rx}</td>
    <td>${r.sku || ""}</td>
    <td>${r.material || ""}</td>
    <td>${r.status || ""}</td>
    <td>${r.breakage ?? ""}</td>
    <td>${[r.adj1, r.adj2, r.adj3, r.adj4].filter(Boolean).join(" | ")}</td>
  `;
  return tr;
}