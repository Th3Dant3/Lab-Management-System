document.addEventListener("DOMContentLoaded", () => {

  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  const els = {
    dateFilter: document.getElementById("dateFilter"),
    btnAll: document.getElementById("btnAll"),

    tabReasons: document.getElementById("tabReasons"),
    tabSentBack: document.getElementById("tabSentBack"),
    tabDelay: document.getElementById("tabDelay"),

    reasonsView: document.getElementById("reasonsView"),
    sentBackView: document.getElementById("sentBackView"),
    delayView: document.getElementById("delayView"),

    reasonBubbles: document.getElementById("reasonBubbles"),
    sentBackBubbles: document.getElementById("sentBackBubbles"),

    delayTableBody: document.getElementById("delayTableBody"),
    delayByDept: document.getElementById("delayByDept"),

    bucket12: document.getElementById("bucket-1-2"),
    bucket35: document.getElementById("bucket-3-5"),
    bucket5p: document.getElementById("bucket-5plus")
  };

  let activeTab = "reasons";

  /***********************
   HELPERS
  ***********************/
  const safeNumber = v => Number.isFinite(Number(v)) ? Number(v) : 0;

  const formatInt = v => safeNumber(v).toLocaleString("en-US");

  const formatPct = v => `${safeNumber(v).toFixed(0)}%`;

  const formatDec = (v, d = 2) => safeNumber(v).toFixed(d);

  const escapeHtml = str =>
    String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");

  const getFilterValue = () =>
    els.dateFilter?.value?.trim() || "all";

  function set(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  }

  /***********************
   LOAD
  ***********************/
  async function load() {
    const url = `${API_URL}?date=${encodeURIComponent(getFilterValue())}`;

    try {
      const res = await fetch(url, { cache: "no-store" });
      const data = await res.json();
      render(data);
    } catch (e) {
      console.error("LOAD ERROR:", e);
    }
  }

  /***********************
   RENDER
  ***********************/
  function render(d) {

    const total = safeNumber(d.total);

    set("activeHolds", formatInt(d.active));
    set("evaluatedCount", formatInt(d.evaluated));
    set("coveragePct", formatPct(d.coverage));
    set("lastUpdated", d.lastUpdated || "--");

    set(
      "avgInvestigationTime",
      `${formatDec(d.avgInvestigationDays)} Days (${formatDec(d.avgInvestigationHours)} Hours)`
    );

    set("receivedToday", formatInt(d.receivedOnSelectedDate));
    set("evaluatedToday", formatInt(d.evaluatedOnSelectedDate));
    set("avgArrivalDelayDays", `${formatDec(d.avgArrivalDelayDays)} Days`);
    set("oldestInvestigation", d.oldestInvestigation || "--");
    set("latestInvestigation", d.latestInvestigation || "--");

    renderTop10("reasonBubbles", d.reasonsTop10, total);
    renderTop10("sentBackBubbles", d.sentBackTop10, total);

    renderDelaySummary(d);
    renderDelayTable(d.delayRows);
  }

  /***********************
   TOP 10
  ***********************/
  function renderTop10(target, rows, total) {
    const wrap = document.getElementById(target);
    if (!wrap) return;

    if (!rows?.length) {
      wrap.innerHTML = `<div class="empty-state">No data</div>`;
      return;
    }

    wrap.innerHTML = rows.map(item => {
      const pct = total ? (item.count / total) * 100 : 0;

      return `
        <div class="pill">
          <div class="pill-left">
            <div class="pill-title">${escapeHtml(item.name)}</div>
            <div class="pill-bar">
              <div class="pill-fill" style="width:${pct.toFixed(1)}%"></div>
            </div>
            <div class="pill-pct">${pct.toFixed(1)}%</div>
          </div>
          <div class="pill-count">${formatInt(item.count)}</div>
        </div>
      `;
    }).join("");
  }

  /***********************
   DELAY SUMMARY (NO API NEEDED)
  ***********************/
  function renderDelaySummary(d) {

    const rows = d.delayRows || [];

    let b12 = 0, b35 = 0, b5p = 0;

    rows.forEach(r => {
      const days = safeNumber(r.daysToArrive);

      if (days >= 1 && days <= 2) b12++;
      else if (days > 2 && days <= 5) b35++;
      else if (days > 5) b5p++;
    });

    if (els.bucket12) els.bucket12.textContent = b12;
    if (els.bucket35) els.bucket35.textContent = b35;
    if (els.bucket5p) els.bucket5p.textContent = b5p;

    // GROUP
    const map = {};

    rows.forEach(r => {
      const key = r.sentBack || "Unknown";
      const days = safeNumber(r.daysToArrive);

      if (!map[key]) {
        map[key] = { count: 0, total: 0, max: 0 };
      }

      map[key].count++;
      map[key].total += days;
      if (days > map[key].max) map[key].max = days;
    });

    const grouped = Object.keys(map).map(k => ({
      name: k,
      count: map[k].count,
      avgDays: (map[k].total / map[k].count).toFixed(1),
      maxDays: map[k].max.toFixed(1)
    }));

    grouped.sort((a, b) => b.maxDays - a.maxDays);

    const wrap = els.delayByDept;
    if (!wrap) return;

    if (!grouped.length) {
      wrap.innerHTML = `<div class="empty-state">No delay data</div>`;
      return;
    }

    wrap.innerHTML = grouped.map(r => `
      <div class="pill">
        <div class="pill-left">
          <div class="pill-title">${escapeHtml(r.name)}</div>
          <div class="pill-bar">
            <div class="pill-fill" style="width:${Math.min(r.count * 10, 100)}%"></div>
          </div>
          <div class="pill-pct">
            Avg: ${r.avgDays} days • Max: ${r.maxDays}
          </div>
        </div>
        <div class="pill-count">${formatInt(r.count)}</div>
      </div>
    `).join("");
  }

  /***********************
   TABLE
  ***********************/
  function renderDelayTable(rows) {

    const body = els.delayTableBody;
    if (!body) return;

    if (!rows?.length) {
      body.innerHTML = `<tr><td colspan="5">No delay records</td></tr>`;
      return;
    }

    body.innerHTML = rows.map(r => {

      const d = safeNumber(r.daysToArrive);

      const cls =
        d >= 3 ? "delay-high" :
        d >= 1 ? "delay-medium" :
        "delay-low";

      return `
        <tr>
          <td>${escapeHtml(r.rx)}</td>
          <td>${escapeHtml(r.sentBack)}</td>
          <td>${escapeHtml(r.scanDate)}</td>
          <td>${escapeHtml(r.arrived)}</td>
          <td>
            <span class="delay-badge ${cls}">
              ${d}
            </span>
          </td>
        </tr>
      `;
    }).join("");
  }

  /***********************
   TABS (FIXED)
  ***********************/
  function switchTab(tab) {

    activeTab = tab;

    if (els.reasonsView)
      els.reasonsView.style.display = tab === "reasons" ? "block" : "none";

    if (els.sentBackView)
      els.sentBackView.style.display = tab === "sent" ? "block" : "none";

    if (els.delayView)
      els.delayView.style.display = tab === "delay" ? "block" : "none";

    els.tabReasons?.classList.toggle("active", tab === "reasons");
    els.tabSentBack?.classList.toggle("active", tab === "sent");
    els.tabDelay?.classList.toggle("active", tab === "delay");
  }

  /***********************
   EVENTS
  ***********************/
  els.dateFilter?.addEventListener("change", load);

  els.btnAll?.addEventListener("click", () => {
    els.dateFilter.value = "";
    load();
  });

  els.tabReasons?.addEventListener("click", () => switchTab("reasons"));
  els.tabSentBack?.addEventListener("click", () => switchTab("sent"));
  els.tabDelay?.addEventListener("click", () => switchTab("delay"));

  /***********************
   INIT
  ***********************/
  switchTab(activeTab);
  load();

});