const API_URL =
  "https://script.google.com/macros/s/AKfycbzuUpwsa2VTGjxeSP3wVV-x2Z4yMkGH_eYS86VpFH0CzC2Ii5S5U2ag-S5d89RZ8oHJ/exec";

let selectedYear = "2026";
let currentView = "prod";
let groupMode = "none";

/* YEAR */
document.querySelectorAll(".year-btn").forEach(btn => {
  btn.onclick = () => {
    selectedYear = btn.dataset.year;
    document.querySelectorAll(".year-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    load();
  };
});

/* VIEW */
document.querySelectorAll("[data-view]").forEach(tab => {
  tab.onclick = () => {
    currentView = tab.dataset.view;
    document.querySelectorAll("[data-view]").forEach(t => t.classList.remove("active"));
    tab.classList.add("active");
    load();
  };
});

/* GROUP MODE */
document.querySelectorAll("[data-group]").forEach(tab => {
  tab.onclick = () => {
    groupMode = tab.dataset.group;
    document.querySelectorAll("[data-group]").forEach(t => t.classList.remove("active"));
    tab.classList.add("active");
    load();
  };
});

/* LOAD DATA */
async function load(){
  const res = await fetch(`${API_URL}?year=${selectedYear}`, { cache:"no-store" });
  const json = await res.json();
  if(json.error){
    console.error(json.error);
    return;
  }

  const rows = json.rows.filter(r =>
    currentView === "prod"
      ? r.status !== "TEST" && r.status !== "Cancel"
      : r.status === "TEST" || r.status === "Cancel"
  );

  renderMetrics(rows);
  renderTable(rows);
}

/* METRICS */
function renderMetrics(rows){
  document.getElementById("m-total").textContent = rows.length;
  document.getElementById("m-active").textContent = rows.filter(r => r.status === "Active").length;
  document.getElementById("m-adjusted").textContent =
    rows.filter(r => r.adj1 || r.adj2 || r.adj3 || r.adj4).length;
  document.getElementById("m-escalated").textContent = rows.filter(r => r.status === "Escalated").length;
  document.getElementById("m-test").textContent = rows.filter(r => r.status === "TEST").length;
  document.getElementById("m-cancel").textContent = rows.filter(r => r.status === "Cancel").length;
}

/* TABLE */
function renderTable(rows){
  const body = document.getElementById("table-body");
  body.innerHTML = "";

  if(!rows.length){
    body.innerHTML = `<tr><td colspan="6" class="empty">No data available</td></tr>`;
    return;
  }

  if(groupMode === "none"){
    rows.forEach(r => body.appendChild(makeRow(r)));
    return;
  }

  const key = groupMode === "sku" ? "sku" : "material";
  const groups = {};

  rows.forEach(r => {
    const g = (r[key] || "Unknown").trim();
    if(!groups[g]) groups[g] = [];
    groups[g].push(r);
  });

  const sortedGroups = Object.entries(groups)
    .sort((a,b) => b[1].length - a[1].length);

  sortedGroups.forEach(([label, items]) => {
    let open = false;

    const header = document.createElement("tr");
    header.className = "group-row";
    header.innerHTML = `<td colspan="6">▶ ${label} (${items.length})</td>`;
    body.appendChild(header);

    const children = items.map(r => {
      const tr = makeRow(r);
      tr.classList.add("hidden");
      body.appendChild(tr);
      return tr;
    });

    header.onclick = () => {
      open = !open;
      children.forEach(tr => tr.classList.toggle("hidden", !open));
      header.firstChild.textContent =
        `${open ? "▼" : "▶"} ${label} (${items.length})`;
    };
  });
}

function makeRow(r){
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${r.rx}</td>
    <td>${r.sku}</td>
    <td>${r.material}</td>
    <td>${r.status}</td>
    <td>${r.breakage ?? ""}</td>
    <td>${[r.adj1,r.adj2,r.adj3,r.adj4].filter(Boolean).join(" | ")}</td>
  `;
  return tr;
}

/* INIT */
load();