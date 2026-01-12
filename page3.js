const ENDPOINT = "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";


let dashboardData = null;

function animateValue(el, start, end){
  let cur = start;
  const step = Math.max(1, Math.floor((end - start) / 30));
  const timer = setInterval(() => {
    cur += step;
    if(cur >= end){
      cur = end;
      clearInterval(timer);
    }
    el.textContent = cur;
  }, 16);
}

async function loadDashboard(){
  try{
    const res = await fetch(ENDPOINT);
    const data = await res.json();
    dashboardData = data;

    const active = data.active || 0;
    const evaluated = data.evaluated || 0;
    const total = active + evaluated;
    const completion = total ? Math.round((evaluated / total) * 100) : 0;

    animateValue(document.getElementById("active"),0,active);
    animateValue(document.getElementById("evaluated"),0,evaluated);
    document.getElementById("completion").textContent = `${completion}%`;

    document.getElementById("updated").textContent =
      `Last updated: ${new Date(data.updatedAt).toLocaleString()}`;

    document.getElementById("source").textContent =
      `Source: Tracker Count sheet • Rows scanned: ${data.rows || "?"}`;

    renderBreakdown(data.bySentBack || {});
  }
  catch(err){
    console.error(err);
    document.getElementById("breakdownBody").innerHTML =
      `<tr><td colspan="2">Failed to load data</td></tr>`;
  }
}

function renderBreakdown(breakdown){
  const body = document.getElementById("breakdownBody");
  body.innerHTML = "";

  const entries = Object.entries(breakdown);
  if(!entries.length){
    body.innerHTML = `<tr><td colspan="2">No data</td></tr>`;
    return;
  }

  entries.sort((a,b)=>b[1]-a[1]).forEach(([bucket,count])=>{
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${bucket}</td><td>${count}</td>`;

    tr.addEventListener("click",()=>{
      document
        .querySelectorAll("#breakdownBody tr")
        .forEach(r=>r.classList.remove("active"));

      tr.classList.add("active");
      showDetail(bucket,count);
    });

    body.appendChild(tr);
  });
}

function showDetail(bucket,count){
  const panel = document.getElementById("detailPanel");
  panel.style.display = "block";

  const total = (dashboardData.active || 0) + (dashboardData.evaluated || 0);
  const pct = total ? ((count / total) * 100).toFixed(1) : "0";

  document.getElementById("detailTitle").textContent = bucket;
  document.getElementById("detailCount").textContent =
    `Jobs routed here: ${count}`;
  document.getElementById("detailPercent").textContent =
    `Impact share: ${pct}% of workload`;

  const last =
    dashboardData.lastChangeByBucket?.[bucket] ||
    dashboardData.updatedAt ||
    null;

  document.getElementById("detailUpdated").textContent =
    `Last signal received: ${last ? new Date(last).toLocaleString() : "Unknown"}`;
}

loadDashboard();
setInterval(loadDashboard, 30000);
