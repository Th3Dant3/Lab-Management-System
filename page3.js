document.addEventListener("DOMContentLoaded", () => {
  const API_URL =
    "https://script.google.com/macros/s/AKfycbzb9LNa7_5dfr7lfFf_MCkHVamM3T5Sw7iByx58WKgWCGvvl6ysZZyIsEBWppuCL3A/exec";

  const REFRESH_MS = 30000;

  const els = {
    dateFilter: document.getElementById("dateFilter"),
    btnToday: document.getElementById("btnToday"),
    btnYesterday: document.getElementById("btnYesterday"),
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

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function toLocalInputDate(date) {
    const yr = date.getFullYear();
    const mo = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${yr}-${mo}-${day}`;
  }

  function getFilterValue() {
    const raw = els.dateFilter?.value?.trim();
    return raw ? raw : "all";
  }

  function setLoadingState(loading) {
    isLoading = loading;
    document.body.classList.toggle("is-loading", loading);
  }

  function markQuickButton(mode) {
    [els.btnToday, els.btnYesterday, els.btnAll].forEach(btn => {
      if (btn) btn.classList.remove("active");
    });

    if (mode === "today") els.btnToday?.classList.add("active");
    if (mode === "yesterday") els.btnYesterday?.classList.add("active");
    if (mode === "all") els.btnAll?.classList.add("active");
  }

  function updateQuickButtonsFromDate() {
    const value = getFilterValue();

    if (value === "all") {
      markQuickButton("all");
      return;
    }

    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);

    const todayStr = toLocalInputDate(today);
    const yesterdayStr = toLocalInputDate(yesterday);

    [els.btnToday, els.btnYesterday, els.btnAll].forEach(btn => {
      if (btn) btn.classList.remove("active");
    });

    if (value === todayStr) {
      markQuickButton("today");
    } else if (value === yesterdayStr) {
      markQuickButton("yesterday");
    }
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

    updateQuickButtonsFromDate();
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

    set("receivedToday", formatInteger(d.receivedToday));
    set("evaluatedToday", formatInteger(d.evaluatedToday));
    set("avgArrivalDelayDays", `${formatDecimal(d.avgArrivalDelayDays, 2)} Days`);
    set("oldestInvestigation", d.oldestInvestigation || "--");
    set("latestInvestigation", d.latestInvestigation || "--");

    renderBubbles("reasonBubbles", d.reasons, total, false, "No root cause data available");
    renderBubbles("sentBackBubbles", d.sentBack, total, true, "No returned-to data available");
    renderDelayTable(d.delayRows || []);
  }

  /***********************
   RENDER BUBBLES
  ***********************/
  function renderBubbles(target, data, total, nested = false, emptyMessage = "No data available") {
    const wrap = document.getElementById(target);
    if (!wrap) return;

    wrap.innerHTML = "";

    if (!data || typeof data !== "object" || Object.keys(data).length === 0) {
      wrap.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
      return;
    }

    const entries = Object.entries(data)
      .map(([key, value]) => [key, nested ? safeNumber(value?.count) : safeNumber(value)])
      .filter(([, count]) => count > 0)
      .sort((a, b) => b[1] - a[1]);

    if (!entries.length) {
      wrap.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
      return;
    }

    const safeTotal = total > 0 ? total : entries.reduce((sum, [, count]) => sum + count, 0);

    wrap.innerHTML = entries.map(([label, count]) => {
      const pct = safeTotal > 0 ? (count / safeTotal) * 100 : 0;

      return `
        <div class="pill">
          <div class="pill-left">
            <div class="pill-title">${escapeHtml(label)}</div>
            <div class="pill-bar">
              <div class="pill-fill" style="width:${pct.toFixed(1)}%"></div>
            </div>
            <div class="pill-pct">${pct.toFixed(1)}%</div>
          </div>
          <div class="pill-count">${formatInteger(count)}</div>
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

    if (!rows || !rows.length) {
      body.innerHTML = `
        <tr>
          <td colspan="5">No delay records available</td>
        </tr>
      `;
      return;
    }

    body.innerHTML = rows
      .slice(0, 50)
      .map(row => {
        const delayClass =
          Number(row.daysToArrive) >= 3
            ? "delay-high"
            : Number(row.daysToArrive) >= 1
            ? "delay-medium"
            : "delay-low";

        return `
          <tr>
            <td>${escapeHtml(row.rx)}</td>
            <td>${escapeHtml(row.sentBack || "")}</td>
            <td>${escapeHtml(row.scanDate || "")}</td>
            <td>${escapeHtml(row.arrived || "")}</td>
            <td><span class="delay-badge ${delayClass}">${escapeHtml(row.daysToArrive)}</span></td>
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
    updateQuickButtonsFromDate();
    load();
  });

  els.btnToday?.addEventListener("click", () => {
    els.dateFilter.value = toLocalInputDate(new Date());
    markQuickButton("today");
    load();
  });

  els.btnYesterday?.addEventListener("click", () => {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    els.dateFilter.value = toLocalInputDate(d);
    markQuickButton("yesterday");
    load();
  });

  els.btnAll?.addEventListener("click", () => {
    els.dateFilter.value = "";
    markQuickButton("all");
    load();
  });

  /***********************
   INIT
  ***********************/
  switchTab(activeTab);
  updateQuickButtonsFromDate();
  load();

  /***********************
   AUTO REFRESH
  ***********************/
  setInterval(() => {
    if (!isLoading) load();
  }, REFRESH_MS);
});