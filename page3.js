document.addEventListener("DOMContentLoaded", () => {

  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  const REFRESH_MS = 30000;

  const set = (id, v) => {
    const el = document.getElementById(id);
    if (el) el.textContent = v;
  };

  function load() {
    fetch(API_URL)
      .then(r => r.json())
      .then(d => render(d))
      .catch(() => fail());
  }

  function render(d) {
    set("activeHolds", d.active);
    set("evaluatedCount", d.evaluated);
    set("coveragePct", (d.coverage * 100).toFixed(1) + "%");
    set("lastUpdated", d.lastUpdated);

    renderBubbles("reasonBubbles", d.reasons, d.total);
    renderBubbles("sentBackBubbles", d.sentBack, d.total, true);
  }

  function renderBubbles(targetId, data, total, nested = false) {
    const wrap = document.getElementById(targetId);
    if (!wrap || !data) return;

    wrap.innerHTML = "";

    const entries = Object.entries(data)
      .map(([k, v]) => nested ? [k, v.count] : [k, v])
      .sort((a, b) => b[1] - a[1]);

    entries.forEach(([label, count]) => {
      const pct = (count / total) * 100;

      wrap.innerHTML += `
        <div class="pill">
          <div class="pill-left">
            <div class="pill-title">${label}</div>
            <div class="pill-bar">
              <div class="pill-fill" style="width:${pct}%"></div>
            </div>
            <div class="pill-pct">${pct.toFixed(1)}%</div>
          </div>
          <div class="pill-count">${count}</div>
        </div>
      `;
    });
  }

  function fail() {
    ["activeHolds","evaluatedCount","coveragePct","lastUpdated"]
      .forEach(id => set(id, "ERR"));
  }

  // TAB SWITCH
  document.getElementById("tabReasons").onclick = () => switchTab("reasons");
  document.getElementById("tabSentBack").onclick = () => switchTab("sent");

  function switchTab(tab) {
    document.getElementById("reasonsView").style.display =
      tab === "reasons" ? "block" : "none";
    document.getElementById("sentBackView").style.display =
      tab === "sent" ? "block" : "none";

    document.getElementById("tabReasons").classList.toggle("active", tab === "reasons");
    document.getElementById("tabSentBack").classList.toggle("active", tab === "sent");
  }

  load();
  setInterval(load, REFRESH_MS);
});