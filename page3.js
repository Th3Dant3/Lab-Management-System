document.addEventListener("DOMContentLoaded", () => {
  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  const els = {
    dateFilter: document.getElementById("dateFilter"),
    btnAll: document.getElementById("btnAll"),
    tabReasons: document.getElementById("tabReasons"),
    tabSentBack: document.getElementById("tabSentBack"),
    reasonsView: document.getElementById("reasonsView"),
    sentBackView: document.getElementById("sentBackView"),
    reasonBubbles: document.getElementById("reasonBubbles"),
    sentBackBubbles: document.getElementById("sentBackBubbles"),
    delayTableBody: document.getElementById("delayTableBody"),

    // NEW
    delayByDept: document.getElementById("delayByDept"),
    bucket12: document.getElementById("bucket-1-2"),
    bucket35: document.getElementById("bucket-3-5"),
    bucket5p: document.getElementById("bucket-5plus")
  };

  let activeTab = "reasons";

  /***********************
   HELPERS
  ***********************/
  function set(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function safeNumber(value) {
    const n = Number(value);
    return Number.isFinite(n) ? n : 0;
  }

  function formatInteger(value) {
    return safeNumber(value).toLocaleString("en-US");
  }

  function formatPercent(value) {
    return `${safeNumber(value).toFixed(0)}%`;
  }

  function formatDecimal(value, digits = 2) {
    return safeNumber(value).toFixed(digits);
  }

  function formatDelayValue(value) {
    const n = safeNumber(value);
    return Number.isInteger(n) ? String(n) : n.toFixed(2);
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function getFilterValue() {
    return els.dateFilter?.value?.trim() || "all";
  }

  function setLoadingState(loading) {
    document.body.classList.toggle("is-loading", loading);
  }

  function markAllButton() {
    if (!els.btnAll) return;
    els.btnAll.classList.toggle("active", getFilterValue() === "all");
  }

  function showMetricSkeleton() {
    [
      "activeHolds",
      "evaluatedCount",
      "coveragePct",
      "lastUpdated",
      "avgInvestigationTime",
      "receivedToday",
      "evaluatedToday",
      "avgArrivalDelayDays",
      "oldestInvestigation",
      "latestInvestigation"
    ].forEach(id => set(id, "Loading..."));
  }

  /***********************
   LOAD
  ***********************/
  async function load() {
    const url = `${API_URL}?date=${encodeURIComponent(getFilterValue())}`;

    markAllButton();
    setLoadingState(true);
    showMetricSkeleton();

    try {
      const res = await fetch(url, { cache: "no-store" });
      const data = await res.json();
      render(data);
    } catch (e) {
      console.error(e);
      fail();
    } finally {
      setLoadingState(false);
    }
  }

  /***********************
   RENDER
  ***********************/
  function render(d) {
    const total = safeNumber(d.total);

    set("activeHolds", formatInteger(d.active));
    set("evaluatedCount", formatInteger(d.evaluated));
    set("coveragePct", formatPercent(d.coverage));
    set("lastUpdated", d.lastUpdated || "--");

    set(
      "avgInvestigationTime",
      `${formatDecimal(d.avgInvestigationDays)} Days (${formatDecimal(d.avgInvestigationHours)} Hours)`
    );

    set("receivedToday", formatInteger(d.receivedOnSelectedDate));
    set("evaluatedToday", formatInteger(d.evaluatedOnSelectedDate));
    set("avgArrivalDelayDays", `${formatDecimal(d.avgArrivalDelayDays)} Days`);
    set("oldestInvestigation", d.oldestInvestigation || "--");
    set("latestInvestigation", d.latestInvestigation || "--");

    renderTop10Bubbles("reasonBubbles", d.reasonsTop10, total);
    renderTop10Bubbles("sentBackBubbles", d.sentBackTop10, total);

    renderDelaySummary(d); // 🔥 NEW
    renderDelayTable(d.delayRows);
  }

  /***********************
   DELAY SUMMARY (NEW)
  ***********************/
  function renderDelaySummary(d) {

    // BUCKETS
    if (els.bucket12) els.bucket12.textContent = d.delayBuckets?.["1-2 Days"] || 0;
    if (els.bucket35) els.bucket35.textContent = d.delayBuckets?.["3-5 Days"] || 0;
    if (els.bucket5p) els.bucket5p.textContent = d.delayBuckets?.["5+ Days"] || 0;

    // GROUPED
    const wrap = els.delayByDept;
    if (!wrap) return;

    wrap.innerHTML = "";

    const rows = d.delayByDept || [];

    if (!rows.length) {
      wrap.innerHTML = `<div class="empty-state">No delay data</div>`;
      return;
    }

    rows.forEach(r => {
      const el = document.createElement("div");
      el.className = "pill";

      el.innerHTML = `
        <div class="pill-left">
          <div class="pill-title">${escapeHtml(r.name)}</div>

          <div class="pill-bar">
            <div class="pill-fill" style="width:${Math.min(r.count * 10, 100)}%"></div>
          </div>

          <div class="pill-pct">
            Avg: ${r.avgDays} days • Max: ${r.maxDays}
          </div>
        </div>

        <div class="pill-count">${formatInteger(r.count)}</div>
      `;

      wrap.appendChild(el);
    });
  }

  /***********************
   TOP 10
  ***********************/
  function renderTop10Bubbles(target, rows, total) {
    const wrap = document.getElementById(target);
    if (!wrap) return;

    if (!rows?.length) {
      wrap.innerHTML = `<div class="empty-state">No data</div>`;
      return;
    }

    wrap.innerHTML = rows.map(item => {
      const pct = total > 0 ? (item.count / total) * 100 : 0;

      return `
        <div class="pill">
          <div class="pill-left">
            <div class="pill-title">${escapeHtml(item.name)}</div>
            <div class="pill-bar">
              <div class="pill-fill" style="width:${pct.toFixed(1)}%"></div>
            </div>
            <div class="pill-pct">${pct.toFixed(1)}%</div>
          </div>
          <div class="pill-count">${formatInteger(item.count)}</div>
        </div>
      `;
    }).join("");
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

    body.innerHTML = rows.map(row => {
      const d = safeNumber(row.daysToArrive);

      const cls =
        d >= 3 ? "delay-high" :
        d >= 1 ? "delay-medium" :
        "delay-low";

      return `
        <tr>
          <td>${escapeHtml(row.rx)}</td>
          <td>${escapeHtml(row.sentBack)}</td>
          <td>${escapeHtml(row.scanDate)}</td>
          <td>${escapeHtml(row.arrived)}</td>
          <td>
            <span class="delay-badge ${cls}">
              ${formatDelayValue(d)}
            </span>
          </td>
        </tr>
      `;
    }).join("");
  }

  /***********************
   FAIL
  ***********************/
  function fail() {
    ["activeHolds","evaluatedCount","coveragePct","lastUpdated"].forEach(id => set(id, "ERR"));
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

  function switchTab(tab) {
    activeTab = tab;
    els.reasonsView.style.display = tab === "reasons" ? "block" : "none";
    els.sentBackView.style.display = tab === "sent" ? "block" : "none";
  }

  /***********************
   INIT
  ***********************/
  switchTab(activeTab);
  load();
});