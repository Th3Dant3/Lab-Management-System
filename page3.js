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
    delayTableBody: document.getElementById("delayTableBody")
  };

  let activeTab = "reasons";
  let isLoading = false;

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
    const raw = els.dateFilter?.value?.trim();
    return raw ? raw : "all";
  }

  function setLoadingState(loading) {
    isLoading = loading;
    document.body.classList.toggle("is-loading", loading);
  }

  function markAllButton() {
    if (!els.btnAll) return;

    const value = getFilterValue();
    els.btnAll.classList.toggle("active", value === "all");
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
    const date = getFilterValue();
    const url = `${API_URL}?date=${encodeURIComponent(date)}`;

    console.log("API Request:", url);

    markAllButton();
    setLoadingState(true);
    showMetricSkeleton();

    try {
      const response = await fetch(url, {
        method: "GET",
        cache: "no-store"
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const data = await response.json();
      render(data);
    } catch (error) {
      console.error("Dashboard load failed:", error);
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
    const evaluated = safeNumber(d.evaluated);
    const active = safeNumber(d.active);
    const coverage = safeNumber(d.coverage);

    set("activeHolds", formatInteger(active));
    set("evaluatedCount", formatInteger(evaluated));
    set("coveragePct", formatPercent(coverage));
    set("lastUpdated", d.lastUpdated || "--");

    set(
      "avgInvestigationTime",
      `${formatDecimal(d.avgInvestigationDays, 2)} Days (${formatDecimal(d.avgInvestigationHours, 2)} Hours)`
    );

    set("receivedToday", formatInteger(d.receivedOnSelectedDate));
    set("evaluatedToday", formatInteger(d.evaluatedOnSelectedDate));
    set("avgArrivalDelayDays", `${formatDecimal(d.avgArrivalDelayDays, 2)} Days`);
    set("oldestInvestigation", d.oldestInvestigation || "--");
    set("latestInvestigation", d.latestInvestigation || "--");

    renderTop10Bubbles(
      "reasonBubbles",
      d.reasonsTop10 || [],
      total,
      "No root cause data available"
    );

    renderTop10Bubbles(
      "sentBackBubbles",
      d.sentBackTop10 || [],
      total,
      "No returned-to data available"
    );

    renderDelayTable(d.delayRows || []);
  }

  /***********************
   RENDER TOP 10 BREAKDOWN
  ***********************/
  function renderTop10Bubbles(target, rows, total, emptyMessage = "No data available") {
    const wrap = document.getElementById(target);
    if (!wrap) return;

    wrap.innerHTML = "";

    if (!Array.isArray(rows) || !rows.length) {
      wrap.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
      return;
    }

    const cleanRows = rows
      .map(item => ({
        name: item?.name || "",
        count: safeNumber(item?.count)
      }))
      .filter(item => item.name && item.count > 0)
      .slice(0, 10);

    if (!cleanRows.length) {
      wrap.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
      return;
    }

    const safeTotal = total > 0
      ? total
      : cleanRows.reduce((sum, item) => sum + item.count, 0);

    wrap.innerHTML = cleanRows.map(item => {
      const pct = safeTotal > 0 ? (item.count / safeTotal) * 100 : 0;

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
   DELAY TABLE
  ***********************/
  function renderDelayTable(rows) {
    const body = els.delayTableBody;
    if (!body) return;

    if (!Array.isArray(rows) || !rows.length) {
      body.innerHTML = `
        <tr>
          <td colspan="5">No delay records available</td>
        </tr>
      `;
      return;
    }

    body.innerHTML = rows
      .map(row => {
        const delay = safeNumber(row.daysToArrive);

        const delayClass =
          delay >= 3
            ? "delay-high"
            : delay >= 1
            ? "delay-medium"
            : "delay-low";

        return `
          <tr>
            <td class="delay-rx">${escapeHtml(row.rx || "")}</td>
            <td class="delay-returned">${escapeHtml(row.sentBack || "")}</td>
            <td class="delay-date">${escapeHtml(row.scanDate || "")}</td>
            <td class="delay-date">${escapeHtml(row.arrived || "")}</td>
            <td class="delay-days">
              <span class="delay-badge ${delayClass}">
                ${escapeHtml(formatDelayValue(delay))}
              </span>
            </td>
          </tr>
        `;
      })
      .join("");
  }

  /***********************
   FAIL
  ***********************/
  function fail() {
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
    ].forEach(id => set(id, "ERR"));

    if (els.reasonBubbles) {
      els.reasonBubbles.innerHTML = `<div class="empty-state">Unable to load root cause data</div>`;
    }

    if (els.sentBackBubbles) {
      els.sentBackBubbles.innerHTML = `<div class="empty-state">Unable to load returned-to data</div>`;
    }

    if (els.delayTableBody) {
      els.delayTableBody.innerHTML = `
        <tr>
          <td colspan="5">Unable to load delay records</td>
        </tr>
      `;
    }
  }

  /***********************
   TABS
  ***********************/
  function switchTab(tab) {
    activeTab = tab;

    if (els.reasonsView) {
      els.reasonsView.style.display = tab === "reasons" ? "block" : "none";
    }

    if (els.sentBackView) {
      els.sentBackView.style.display = tab === "sent" ? "block" : "none";
    }

    els.tabReasons?.classList.toggle("active", tab === "reasons");
    els.tabSentBack?.classList.toggle("active", tab === "sent");

    els.tabReasons?.setAttribute("aria-selected", tab === "reasons" ? "true" : "false");
    els.tabSentBack?.setAttribute("aria-selected", tab === "sent" ? "true" : "false");
  }

  els.tabReasons?.addEventListener("click", () => switchTab("reasons"));
  els.tabSentBack?.addEventListener("click", () => switchTab("sent"));

  /***********************
   FILTER EVENTS
  ***********************/
  els.dateFilter?.addEventListener("change", () => {
    markAllButton();
    load();
  });

  els.btnAll?.addEventListener("click", () => {
    if (els.dateFilter) {
      els.dateFilter.value = "";
    }
    markAllButton();
    load();
  });

  /***********************
   INIT
  ***********************/
  switchTab(activeTab);
  markAllButton();
  load();
});