(() => {
  "use strict";

  const CONFIG = {
    API_URL:
      "https://script.google.com/macros/s/AKfycbwjj177i1WtRDjZTiBJyeQefQfTgrvH92i__g4cStQz9MoKH1HbH5zL5eeSe73SFyEJbA/exec?action=facilityWipDashboard",
    AUTO_REFRESH_MS: 5 * 60 * 1000,
    COUNTDOWN_INTERVAL_MS: 1000,
    FETCH_TIMEOUT_MS: 30000,
    CACHE_KEY: "facilityWip:lastGood",
    CACHE_MAX_AGE_MS: 3 * 60 * 60 * 1000,
    VISIBLE_STATIONS: 3,
    COUNT_UP_MS: 1000,

    /*
     * Bottleneck detection with hysteresis.
     * A lane alerts when its share of Physical WIP crosses
     * ALERT_ON and stays alerted until it drops below
     * ALERT_OFF. This prevents strobing when a department
     * oscillates around a single threshold.
     */
    ALERT_ON_SHARE: 45,
    ALERT_OFF_SHARE: 38
  };

  const REQUIRED_PAYLOAD_KEYS = [
    "incomingWip",
    "physicalDetail",
    "incomingSnapshot",
    "physicalSnapshot"
  ];

  const PROCESS_LANES = ["Inventory", "Surface", "AR", "Finish"];

  const SUPPORT_ORDER = ["Breakage", "Customer Service", "LMS"];

  const DISPLAY_DEPARTMENT_RULES = [
    {
      sourceDepartment: "Inventory",
      subDepartment: "FSV Scan & Verify",
      displayDepartment: "Finish"
    },
    {
      sourceDepartment: "Inventory",
      subDepartment: "SF Scan & Verify",
      displayDepartment: "Surface"
    },
    {
      sourceDepartment: "Surface",
      subDepartment: "Surface Inspection",
      displayDepartment: "AR"
    }
  ];

  /*
   * Station cards follow the real production process.
   * Departments without a configured sequence fall back to
   * WIP-descending sort.
   *
   * IMPORTANT: strings must match the API payload EXACTLY
   * (hyphens included). normalizeKey collapses spaces and
   * case but NOT hyphens — "DERING" will not match "DE-RING".
   */
  const DEPARTMENT_STATION_ORDER = {
    Surface: [
      "SF Scan & Verify",
      "Surface Unbox",
      "Blocking Line B",
      "Cooling Storage",
      "Generating Line B",
      "Polishing Line B",
      "Engraving Line B",
      "Detaping Line B",
      "Coating Line B",
      "Surface Lead Desk"
    ],
    AR: [
      "Surface Inspection",
      "AR-IN",
      "Basket",
      "Oven",
      "Sectoring",
      "DeRing"
    ],
    Finish: [
      "FSV Scan & Verify",
      "Finish Unbox",
      "AR-OUT",
      "MEI Line A",
      "MEI Line B",
      "MEI Line C",
      "MEI Easy Fit",
      "Mounting",
      "Drill Station",
      "Handstone",
      "Final Inspection",
      "Finish Lead Desk",
      "Foco Vision"
    ]
  };

  const STATION_DISPLAY_NAMES = {
    Surface: {
      "SF Scan & Verify": "Scan & Verify (In Totes)",
      "Surface Unbox": "SF Unbox",
      "Blocking Line B": "Auto Blocker",
      "Generating Line B": "Generating",
      "Polishing Line B": "Polishing Line",
      "Engraving Line B": "Engraving",
      "Detaping Line B": "Detaping",
      "Coating Line B": "Coating Line"
    },
    Finish: {
      "FSV Scan & Verify": "FSV Scan & Verify",
      "Finish Unbox": "Finish Unbox",
      "AR-OUT": "AR-OUT",
      "MEI Line A": "MEI Line A",
      "MEI Line B": "MEI Line B",
      "MEI Line C": "MEI Line C",
      "MEI Easy Fit": "MEI Easy Fit",
      "Drill Station": "Drill",
      "Final Inspection": "Final Inspection"
    }
  };

  const STATION_FLOW_DESCRIPTIONS = {
    Surface: {
      "SF Scan & Verify": "In totes • waiting to be unboxed",
      "Surface Unbox": "Moving to taping and Auto Blocker",
      "Blocking Line B": "Moving to Cooling Storage or Reject Line",
      "Cooling Storage": "Waiting phase • moving to Generating",
      "Generating Line B": "Moving to Polishing or Reject Line",
      "Polishing Line B": "Moving to Engraving or Reject Line",
      "Engraving Line B": "Moving to Detaping",
      "Detaping Line B": "Moving to Coating Line or Reject Line",
      "Coating Line B": "Moving to Surface Inspection",
      "Surface Lead Desk": "Exception handling and production support"
    },
    AR: {
      "Surface Inspection": "Moving to AR-IN or Reject Line",
      "AR-IN": "Cookie sheet • moving to Basket",
      "Basket": "Moving through T-40 to inside AR",
      "Oven": "Degas time",
      "Sectoring": "Building process • chamber or flip lens",
      "DeRing": "Cookie sheet • moving to AR-OUT"
    },
    Finish: {
      "FSV Scan & Verify": "In totes • moving to Finish Unbox",
      "Finish Unbox": "Moving to MEI",
      "AR-OUT": "Moving to MEI",
      "MEI Line A": "Moving to Mounting or Reject",
      "MEI Line B": "Moving to Mounting or Reject",
      "MEI Line C": "Moving to Mounting or Reject",
      "MEI Easy Fit": "Moving to Mounting or Reject",
      "Mounting": "Moving to Drill, Handstone, or Final Inspection",
      "Drill Station": "Moving to Final Inspection or Breakage",
      "Handstone": "Moving to Final Inspection or Breakage",
      "Final Inspection": "Moving to Breakage",
      "Finish Lead Desk": "Exception handling and production support",
      "Foco Vision": "Finish support process"
    }
  };

  const HIDDEN_STATIONS = [
    { department: "Finish", subDepartment: "Shipping" },
    { department: "Surface", subDepartment: "Generating Line A" }
  ];

  const state = {
    data: null,
    stationRows: [],
    expandedDepartments: new Set(),
    alertedDepartments: new Set(),
    previousValues: {},
    nextRefreshAt: Date.now() + CONFIG.AUTO_REFRESH_MS,
    countdownTimer: null,
    refreshTimer: null,
    isLoading: false,
    isCachedView: false
  };

  const elements = {
    facilityWip: document.getElementById("facilityWip"),
    incomingWip: document.getElementById("incomingWip"),
    facilityUpdatedAt: document.getElementById("facilityUpdatedAt"),
    incomingQueueCount: document.getElementById("incomingQueueCount"),
    gapMinutes: document.getElementById("gapMinutes"),
    snapshotTimes: document.getElementById("snapshotTimes"),
    nextRefresh: document.getElementById("nextRefresh"),
    bottleneckNote: document.getElementById("bottleneckNote"),
    kpiFacility: document.getElementById("kpiFacility"),

    connectionStatus: document.getElementById("connectionStatus"),
    connectionStatusText: document.getElementById("connectionStatusText"),
    refreshButton: document.getElementById("refreshButton"),
    refreshIcon: document.getElementById("refreshIcon"),
    expandAllButton: document.getElementById("expandAllButton"),

    alertBanner: document.getElementById("alertBanner"),
    alertTitle: document.getElementById("alertTitle"),
    alertMessage: document.getElementById("alertMessage"),

    laneGrid: document.getElementById("laneGrid"),
    supportStrip: document.getElementById("supportStrip"),
    incomingPanel: document.getElementById("incomingPanel"),
    incomingQueueList: document.getElementById("incomingQueueList"),
    bottleneckList: document.getElementById("bottleneckList"),

    loadingOverlay: document.getElementById("loadingOverlay"),

    laneTemplate: document.getElementById("laneTemplate"),
    stationRowTemplate: document.getElementById("stationRowTemplate"),
    supportCardTemplate: document.getElementById("supportCardTemplate"),
    queueNodeTemplate: document.getElementById("queueNodeTemplate")
  };

  const reducedMotion = window.matchMedia(
    "(prefers-reduced-motion: reduce)"
  ).matches;

  /* ---------- formatting + validation ---------- */

  function formatNumber(value) {
    return new Intl.NumberFormat("en-US").format(Number(value) || 0);
  }

  function safeNumber(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }

    const parsed = Number(value);

    return Number.isFinite(parsed) ? parsed : null;
  }

  function parseDate(value) {
    if (!value) return null;

    const parsed = new Date(value);

    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDateTime(value) {
    const date = parseDate(value);

    if (!date) return "Unavailable";

    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit"
    }).format(date);
  }

  function normalizeKey(value) {
    return String(value || "")
      .trim()
      .replace(/\s+/g, " ")
      .toUpperCase();
  }

  function setText(element, value) {
    if (element) {
      element.textContent = value;
    }
  }

  function validatePayload(payload) {
    const missingKeys = REQUIRED_PAYLOAD_KEYS.filter(
      (key) => !(key in payload)
    );

    if (missingKeys.length > 0) {
      throw new Error(
        `API schema mismatch. Missing: ${missingKeys.join(", ")}. ` +
        "Backend may have been redeployed with changed field names."
      );
    }

    if (!Array.isArray(payload.physicalDetail)) {
      throw new Error(
        "physicalDetail is not an array. Refusing to render."
      );
    }
  }

  /* ---------- last-known-good cache ---------- */

  function saveLastGood(payload) {
    try {
      localStorage.setItem(
        CONFIG.CACHE_KEY,
        JSON.stringify({
          savedAt: Date.now(),
          payload
        })
      );
    } catch (error) {
      /* Quota or private mode. Cache is a bonus, never a blocker. */
    }
  }

  function loadLastGood() {
    try {
      const raw = localStorage.getItem(CONFIG.CACHE_KEY);

      if (!raw) return null;

      const cached = JSON.parse(raw);

      if (
        !cached ||
        !cached.payload ||
        Date.now() - cached.savedAt > CONFIG.CACHE_MAX_AGE_MS
      ) {
        localStorage.removeItem(CONFIG.CACHE_KEY);
        return null;
      }

      return cached;
    } catch (error) {
      return null;
    }
  }

  /* ---------- UI chrome ---------- */

  function setLoading(isLoading, initial = false) {
    state.isLoading = isLoading;

    elements.refreshButton.disabled = isLoading;
    elements.expandAllButton.disabled = isLoading;

    elements.refreshIcon.classList.toggle(
      "primary-button__icon--spinning",
      isLoading
    );

    if (initial) {
      elements.loadingOverlay.classList.toggle(
        "loading-overlay--hidden",
        !isLoading
      );

      elements.loadingOverlay.setAttribute(
        "aria-hidden",
        String(!isLoading)
      );
    }
  }

  function setConnectionStatus(type, text) {
    elements.connectionStatus.classList.remove(
      "status-pill--connected",
      "status-pill--warning",
      "status-pill--error",
      "status-pill--loading"
    );

    elements.connectionStatus.classList.add(`status-pill--${type}`);

    setText(elements.connectionStatusText, text);
  }

  function showAlert(title, message, type = "warning") {
    setText(elements.alertTitle, title);
    setText(elements.alertMessage, message);

    elements.alertBanner.classList.remove(
      "alert-banner--hidden",
      "alert-banner--error"
    );

    if (type === "error") {
      elements.alertBanner.classList.add("alert-banner--error");
    }
  }

  function hideAlert() {
    elements.alertBanner.classList.add("alert-banner--hidden");
  }

  /* ---------- count-up (fires only on changed values) ---------- */

  function animateValue(element, key, endValue) {
    const previous = state.previousValues[key];

    state.previousValues[key] = endValue;

    if (
      reducedMotion ||
      previous === undefined ||
      previous === endValue
    ) {
      setText(element, formatNumber(endValue));
      return;
    }

    const startValue = previous;
    const startTime = performance.now();

    function step(now) {
      const progress = Math.min(
        (now - startTime) / CONFIG.COUNT_UP_MS,
        1
      );

      const eased = 1 - Math.pow(1 - progress, 3);

      const current = Math.round(
        startValue + (endValue - startValue) * eased
      );

      setText(element, formatNumber(current));

      if (progress < 1) {
        requestAnimationFrame(step);
      }
    }

    requestAnimationFrame(step);
  }

  /* ---------- networking ---------- */

  async function fetchWithTimeout(url, timeoutMs) {
    const controller = new AbortController();

    const timeoutId = window.setTimeout(
      () => controller.abort(),
      timeoutMs
    );

    try {
      const response = await fetch(url, {
        method: "GET",
        cache: "no-store",
        signal: controller.signal
      });

      if (!response.ok) {
        throw new Error(
          `HTTP ${response.status} ${response.statusText}`
        );
      }

      return await response.json();
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  function buildRequestUrl() {
    const url = new URL(CONFIG.API_URL);

    url.searchParams.set("pageTs", String(Date.now()));

    return url.toString();
  }

  /* ---------- data shaping (unchanged business rules) ---------- */

  function getDisplayDepartment(sourceDepartment, subDepartment) {
    const sourceKey = normalizeKey(sourceDepartment);
    const stationKey = normalizeKey(subDepartment);

    const matchedRule = DISPLAY_DEPARTMENT_RULES.find((rule) => {
      return (
        normalizeKey(rule.sourceDepartment) === sourceKey &&
        normalizeKey(rule.subDepartment) === stationKey
      );
    });

    return matchedRule
      ? matchedRule.displayDepartment
      : String(sourceDepartment || "Unassigned").trim();
  }

  function applyDisplayOwnershipRules(rows) {
    return rows.map((row) => {
      const sourceDepartment = String(
        row.department || "Unassigned"
      ).trim();

      const subDepartment = String(row.subDepartment || "").trim();

      const displayDepartment = getDisplayDepartment(
        sourceDepartment,
        subDepartment
      );

      return {
        ...row,
        sourceDepartment,
        department: displayDepartment,
        displayDepartment,
        ownershipAdjusted:
          normalizeKey(sourceDepartment) !==
          normalizeKey(displayDepartment)
      };
    });
  }

  function isHiddenStation(row) {
    const departmentKey = normalizeKey(
      row.department || row.displayDepartment || ""
    );

    const stationKey = normalizeKey(row.subDepartment || "");

    return HIDDEN_STATIONS.some((item) => {
      return (
        normalizeKey(item.department) === departmentKey &&
        normalizeKey(item.subDepartment) === stationKey
      );
    });
  }

  function getStationDisplayName(department, subDepartment) {
    const names = STATION_DISPLAY_NAMES[department];

    if (
      names &&
      Object.prototype.hasOwnProperty.call(names, subDepartment)
    ) {
      return names[subDepartment];
    }

    return subDepartment;
  }

  function getStationFlowDescription(department, subDepartment) {
    const descriptions = STATION_FLOW_DESCRIPTIONS[department];

    if (
      descriptions &&
      Object.prototype.hasOwnProperty.call(
        descriptions,
        subDepartment
      )
    ) {
      return descriptions[subDepartment];
    }

    return "";
  }

  function sortDepartmentStations(department, stations) {
    const configuredOrder = DEPARTMENT_STATION_ORDER[department];

    stations.sort((a, b) => {
      if (Array.isArray(configuredOrder)) {
        const findOrder = (name) => {
          const key = normalizeKey(name);

          const index = configuredOrder.findIndex(
            (stationName) => normalizeKey(stationName) === key
          );

          return index === -1 ? 999 : index;
        };

        const orderDiff =
          findOrder(a.subDepartment) - findOrder(b.subDepartment);

        if (orderDiff !== 0) return orderDiff;

        return String(a.subDepartment || "").localeCompare(
          String(b.subDepartment || "")
        );
      }

      const valueDifference =
        Number(b.currentWip || 0) - Number(a.currentWip || 0);

      if (valueDifference !== 0) return valueDifference;

      return String(a.subDepartment || "").localeCompare(
        String(b.subDepartment || "")
      );
    });
  }

  function buildDepartmentGroups(rows) {
    const groups = {};

    rows.forEach((row) => {
      const department = String(
        row.department || "Unassigned"
      ).trim();

      if (!groups[department]) {
        groups[department] = {
          department,
          currentWip: 0,
          stations: []
        };
      }

      groups[department].currentWip +=
        Number(row.currentWip) || 0;

      if (!isHiddenStation(row)) {
        groups[department].stations.push(row);
      }
    });

    Object.values(groups).forEach((group) => {
      sortDepartmentStations(group.department, group.stations);
    });

    return groups;
  }

  /* ---------- bottleneck detection with hysteresis ---------- */

  function updateAlerts(groups, physicalTotal) {
    if (!physicalTotal) {
      state.alertedDepartments.clear();
      return;
    }

    PROCESS_LANES.forEach((department) => {
      const group = groups[department];

      const share = group
        ? (group.currentWip / physicalTotal) * 100
        : 0;

      const isAlerted = state.alertedDepartments.has(department);

      if (!isAlerted && share >= CONFIG.ALERT_ON_SHARE) {
        state.alertedDepartments.add(department);
      } else if (isAlerted && share < CONFIG.ALERT_OFF_SHARE) {
        state.alertedDepartments.delete(department);
      }
    });

    const alerted = [...state.alertedDepartments];

    setText(
      elements.bottleneckNote,
      alerted.length > 0
        ? `Bottleneck: ${alerted.join(", ")}`
        : "No bottleneck detected"
    );

    elements.bottleneckNote.classList.toggle(
      "kpi__meta--alert",
      alerted.length > 0
    );
  }

  /* ---------- rendering ---------- */

  function renderDashboard(data, options = {}) {
    const isCached = options.cached === true;

    state.isCachedView = isCached;

    document.body.classList.toggle("is-cached", isCached);

    const physicalWip = safeNumber(
      data.currentPhysicalWip ?? data.physicalWip
    );

    const facilityWip = safeNumber(
      data.currentFacilityWip ??
      data.currentPhysicalWip ??
      data.physicalWip ??
      data.facilityWip
    );

    const incomingWip = safeNumber(data.incomingWip);

    if (facilityWip === null) {
      setText(elements.facilityWip, "Unavailable");
    } else {
      animateValue(elements.facilityWip, "facility", facilityWip);
    }

    if (incomingWip === null) {
      setText(elements.incomingWip, "Unavailable");
    } else {
      animateValue(elements.incomingWip, "incoming", incomingWip);
    }

    setText(
      elements.facilityUpdatedAt,
      `Updated ${formatDateTime(data.hubUpdatedAt)}`
    );

    setText(
      elements.incomingQueueCount,
      `${Number(data.incomingRowCount) || 0} queues • not in total`
    );

    const gapMinutes = safeNumber(
      data.synchronizationGapMinutes ?? data.gapMinutes
    );

    setText(
      elements.gapMinutes,
      gapMinutes === null ? "Unavailable" : `${gapMinutes} min`
    );

    setText(
      elements.snapshotTimes,
      `IN ${formatDateTime(data.incomingSnapshot)} • PHY ${formatDateTime(
        data.currentPhysicalSnapshot || data.physicalSnapshot
      )}`
    );

    const groups = buildDepartmentGroups(state.stationRows);

    updateAlerts(groups, physicalWip || 0);

    renderLanes(groups, physicalWip || 0);
    renderSupport(groups, physicalWip || 0);
    renderBottlenecks(groups, physicalWip || 0);

    renderIncomingQueues(
      Array.isArray(data.incomingBreakdown)
        ? data.incomingBreakdown
        : []
    );
  }

  function renderLanes(groups, physicalTotal) {
    elements.laneGrid.innerHTML = "";

    const laneGroups = PROCESS_LANES.map((department) => {
      return (
        groups[department] || {
          department,
          currentWip: 0,
          stations: []
        }
      );
    });

    const maxValue = Math.max(
      ...laneGroups.map((group) => group.currentWip),
      1
    );

    laneGroups.forEach((group, index) => {
      if (index > 0) {
        const arrow = document.createElement("div");
        arrow.className = "lane-arrow";
        arrow.setAttribute("aria-hidden", "true");
        elements.laneGrid.appendChild(arrow);
      }

      const node = elements.laneTemplate.content.cloneNode(true);
      const article = node.querySelector(".lane");
      const head = node.querySelector(".lane__head");
      const stationsBox = node.querySelector(".lane__stations");
      const moreButton = node.querySelector(".lane__more");

      const isExpanded = state.expandedDepartments.has(
        group.department
      );

      const isAlerted = state.alertedDepartments.has(
        group.department
      );

      article.classList.toggle("lane--alert", isAlerted);
      article.dataset.department = normalizeKey(group.department);
      article.style.animationDelay = `${index * 70}ms`;

      head.setAttribute("aria-expanded", String(isExpanded));

      setText(node.querySelector(".lane__name"), group.department);

      setText(
        node.querySelector(".lane__value"),
        formatNumber(group.currentWip)
      );

      const width = (group.currentWip / maxValue) * 100;

      node.querySelector(".lane__fill").style.width =
        `${Math.max(width, 2)}%`;

      const share =
        physicalTotal > 0
          ? (group.currentWip / physicalTotal) * 100
          : 0;

      setText(
        node.querySelector(".lane__share"),
        `${share.toFixed(1)}% of Physical • ${group.stations.length} stations`
      );

      const visible = isExpanded
        ? group.stations
        : group.stations.slice(0, CONFIG.VISIBLE_STATIONS);

      visible.forEach((station) => {
        const rowNode =
          elements.stationRowTemplate.content.cloneNode(true);

        setText(
          rowNode.querySelector(".station-row__name"),
          getStationDisplayName(
            group.department,
            station.subDepartment
          ) || "Unknown Station"
        );

        const description =
          getStationFlowDescription(
            group.department,
            station.subDepartment
          ) ||
          (station.ownershipAdjusted
            ? `Moved from ${station.sourceDepartment}`
            : "");

        const descElement =
          rowNode.querySelector(".station-row__desc");

        if (description) {
          setText(descElement, description);
        } else {
          descElement.remove();
        }

        setText(
          rowNode.querySelector(".station-row__value"),
          formatNumber(station.currentWip)
        );

        stationsBox.appendChild(rowNode);
      });

      const hiddenCount =
        group.stations.length - CONFIG.VISIBLE_STATIONS;

      if (hiddenCount > 0) {
        moreButton.hidden = false;

        setText(
          moreButton,
          isExpanded
            ? "Show Less"
            : `Show +${hiddenCount} More`
        );
      }

      const toggle = () => {
        if (state.expandedDepartments.has(group.department)) {
          state.expandedDepartments.delete(group.department);
        } else {
          state.expandedDepartments.add(group.department);
        }

        rerenderFromState();
      };

      head.addEventListener("click", toggle);
      moreButton.addEventListener("click", toggle);

      elements.laneGrid.appendChild(node);
    });

    updateExpandAllButtonLabel();
  }


  function openSupportDetails(department, group) {
    const existing = document.getElementById("supportDetailsModal");

    if (existing) {
      existing.remove();
    }

    const stations =
      group && Array.isArray(group.stations)
        ? group.stations
        : [];

    const total =
      group && Number.isFinite(Number(group.currentWip))
        ? Number(group.currentWip)
        : 0;

    const overlay = document.createElement("div");
    overlay.id = "supportDetailsModal";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.setAttribute("aria-label", `${department} WIP details`);

    overlay.style.cssText = `
      position: fixed;
      inset: 0;
      z-index: 99999;
      display: grid;
      place-items: center;
      padding: 24px;
      background: rgba(2, 8, 16, 0.82);
      backdrop-filter: blur(10px);
    `;

    const panel = document.createElement("section");

    panel.style.cssText = `
      width: min(680px, 100%);
      max-height: 82vh;
      overflow: auto;
      border: 1px solid rgba(46, 215, 208, 0.7);
      border-radius: 12px;
      color: #f4f8fb;
      background:
        linear-gradient(160deg, rgba(12, 35, 58, 0.99), rgba(5, 17, 31, 0.99));
      box-shadow:
        0 0 0 1px rgba(62, 163, 255, 0.18),
        0 0 40px rgba(46, 215, 208, 0.22);
      font-family: Rajdhani, Inter, sans-serif;
    `;

    const header = document.createElement("header");

    header.style.cssText = `
      position: sticky;
      top: 0;
      z-index: 2;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 18px 20px;
      border-bottom: 1px solid rgba(46, 215, 208, 0.24);
      background: rgba(5, 17, 31, 0.97);
    `;

    const titleWrap = document.createElement("div");

    const eyebrow = document.createElement("div");
    eyebrow.textContent = "SUPPORT & OPERATIONS";
    eyebrow.style.cssText = `
      margin-bottom: 4px;
      color: #2ed7d0;
      font-size: 0.68rem;
      font-weight: 700;
      letter-spacing: 0.16em;
    `;

    const title = document.createElement("h2");
    title.textContent = department;
    title.style.cssText = `
      margin: 0;
      font-family: Orbitron, sans-serif;
      font-size: 1.35rem;
      letter-spacing: 0.05em;
    `;

    titleWrap.append(eyebrow, title);

    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.textContent = "×";
    closeButton.setAttribute("aria-label", "Close details");
    closeButton.style.cssText = `
      width: 42px;
      height: 42px;
      border: 1px solid rgba(46, 215, 208, 0.4);
      border-radius: 8px;
      color: #dffffd;
      background: rgba(46, 215, 208, 0.08);
      cursor: pointer;
      font-size: 1.6rem;
      line-height: 1;
    `;

    header.append(titleWrap, closeButton);

    const body = document.createElement("div");
    body.style.cssText = "padding: 18px 20px 22px;";

    const totalCard = document.createElement("div");
    totalCard.style.cssText = `
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 16px;
      padding: 16px;
      border: 1px solid rgba(62, 163, 255, 0.25);
      border-radius: 9px;
      background: rgba(62, 163, 255, 0.06);
    `;

    const totalLabel = document.createElement("span");
    totalLabel.textContent = "CURRENT WIP";
    totalLabel.style.cssText = `
      color: #8fa7bc;
      font-size: 0.75rem;
      font-weight: 700;
      letter-spacing: 0.12em;
    `;

    const totalValue = document.createElement("strong");
    totalValue.textContent = formatNumber(total);
    totalValue.style.cssText = `
      color: #dffffd;
      font-family: Orbitron, sans-serif;
      font-size: 2rem;
    `;

    totalCard.append(totalLabel, totalValue);
    body.appendChild(totalCard);

    if (stations.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = `No ${department} sub-department details are available.`;
      empty.style.cssText = `
        padding: 22px;
        border: 1px dashed rgba(143, 167, 188, 0.3);
        border-radius: 8px;
        color: #8fa7bc;
        text-align: center;
      `;
      body.appendChild(empty);
    } else {
      stations.forEach((station) => {
        const row = document.createElement("div");
        row.style.cssText = `
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          padding: 13px 4px;
          border-bottom: 1px solid rgba(143, 167, 188, 0.14);
        `;

        const name = document.createElement("span");
        name.textContent =
          getStationDisplayName(department, station.subDepartment) ||
          station.subDepartment ||
          "Unknown";

        name.style.cssText = `
          color: #bdd0df;
          font-size: 1rem;
          font-weight: 600;
        `;

        const value = document.createElement("strong");
        value.textContent = formatNumber(station.currentWip);
        value.style.cssText = `
          color: #dffffd;
          font-family: Orbitron, sans-serif;
          font-size: 1rem;
        `;

        row.append(name, value);
        body.appendChild(row);
      });
    }

    panel.append(header, body);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    const closeModal = () => {
      overlay.remove();
      document.removeEventListener("keydown", handleEscape);
    };

    const handleEscape = (event) => {
      if (event.key === "Escape") {
        closeModal();
      }
    };

    closeButton.addEventListener("click", closeModal);

    overlay.addEventListener("click", (event) => {
      if (event.target === overlay) {
        closeModal();
      }
    });

    document.addEventListener("keydown", handleEscape);
    closeButton.focus();
  }

  function renderSupport(groups) {
    elements.supportStrip.innerHTML = "";

    SUPPORT_ORDER.forEach((department) => {
      const group = groups[department];

      const node =
        elements.supportCardTemplate.content.cloneNode(true);

      const card = node.querySelector(".support-card");

      card.classList.add(
        `support-card--${normalizeKey(department)
          .toLowerCase()
          .replace(/\s+/g, "-")}`
      );
      card.dataset.department = normalizeKey(department);

      setText(node.querySelector(".support-card__name"), department);

      setText(
        node.querySelector(".support-card__value"),
        group ? formatNumber(group.currentWip) : "0"
      );

      card.addEventListener("click", () => {
        const latestGroups = buildDepartmentGroups(
          state.stationRows || []
        );

        openSupportDetails(
          department,
          latestGroups[department] || group || null
        );
      });

      elements.supportStrip.appendChild(node);
    });

    /* Incoming queues toggle lives in the support strip. */
    const node =
      elements.supportCardTemplate.content.cloneNode(true);

    const card = node.querySelector(".support-card");

    card.classList.add("support-card--incoming");
    card.dataset.department = "INCOMING QUEUES";

    setText(
      node.querySelector(".support-card__name"),
      "Incoming Queues"
    );

    setText(
      node.querySelector(".support-card__value"),
      elements.incomingPanel.hidden ? "Show" : "Hide"
    );

    card.addEventListener("click", () => {
      elements.incomingPanel.hidden =
        !elements.incomingPanel.hidden;

      rerenderFromState();
    });

    elements.supportStrip.appendChild(node);
  }


  function renderBottlenecks(groups, physicalTotal) {
    if (!elements.bottleneckList) return;

    const candidates = [];

    Object.values(groups).forEach((group) => {
      if (!PROCESS_LANES.includes(group.department)) return;

      group.stations.forEach((station) => {
        const value = Number(station.currentWip) || 0;

        if (value <= 0) return;

        candidates.push({
          department: group.department,
          station: getStationDisplayName(
            group.department,
            station.subDepartment
          ) || station.subDepartment || "Unknown Station",
          value
        });
      });
    });

    candidates.sort((a, b) => b.value - a.value);

    const top = candidates.slice(0, 5);
    elements.bottleneckList.innerHTML = "";

    if (top.length === 0) {
      elements.bottleneckList.innerHTML =
        '<div class="empty-state">No active bottleneck data.</div>';
      return;
    }

    const maxValue = Math.max(...top.map((item) => item.value), 1);

    top.forEach((item, index) => {
      const ratio = item.value / maxValue;
      const severity =
        index === 0 || ratio >= 0.9
          ? "critical"
          : ratio >= 0.65
            ? "high"
            : "medium";

      const share =
        physicalTotal > 0
          ? (item.value / physicalTotal) * 100
          : 0;

      const card = document.createElement("article");
      card.className = `bottleneck-card bottleneck-card--${severity}`;

      card.innerHTML = `
        <div class="bottleneck-card__rank">${index + 1}</div>
        <div class="bottleneck-card__body">
          <strong class="bottleneck-card__name"></strong>
          <span class="bottleneck-card__department"></span>
          <span class="bottleneck-card__severity"></span>
          <span class="bottleneck-card__share"></span>
        </div>
        <div class="bottleneck-card__gauge" style="--gauge:${Math.max(
          12,
          Math.round(ratio * 100)
        )}%">
          <strong></strong>
          <span>WIP</span>
        </div>
      `;

      setText(
        card.querySelector(".bottleneck-card__name"),
        item.station
      );
      setText(
        card.querySelector(".bottleneck-card__department"),
        `${item.department} Department`
      );
      setText(
        card.querySelector(".bottleneck-card__severity"),
        severity.toUpperCase()
      );
      setText(
        card.querySelector(".bottleneck-card__share"),
        `${share.toFixed(1)}% of Physical WIP`
      );
      setText(
        card.querySelector(".bottleneck-card__gauge strong"),
        formatNumber(item.value)
      );

      elements.bottleneckList.appendChild(card);
    });
  }

  function renderIncomingQueues(rows) {
    elements.incomingQueueList.innerHTML = "";

    if (rows.length === 0) {
      elements.incomingQueueList.innerHTML =
        '<div class="empty-state">No Incoming queue data is available.</div>';
      return;
    }

    rows
      .slice()
      .sort((a, b) => {
        return (
          Number(b.currentWip || 0) - Number(a.currentWip || 0)
        );
      })
      .forEach((row, index) => {
        const node =
          elements.queueNodeTemplate.content.cloneNode(true);

        const article = node.querySelector(".queue-node");

        article.style.animationDelay = `${index * 55}ms`;

        setText(
          node.querySelector(".queue-node__department"),
          row.department || "Unassigned"
        );

        setText(
          node.querySelector(".queue-node__name"),
          row.queue || "Unknown Queue"
        );

        setText(
          node.querySelector(".queue-node__value"),
          formatNumber(row.currentWip)
        );

        elements.incomingQueueList.appendChild(node);
      });
  }

  function rerenderFromState() {
    if (!state.data) return;

    renderDashboard(state.data, {
      cached: state.isCachedView
    });
  }

  function updateExpandAllButtonLabel() {
    const allExpanded = PROCESS_LANES.every((department) =>
      state.expandedDepartments.has(department)
    );

    setText(
      elements.expandAllButton,
      allExpanded ? "Collapse All" : "Expand All"
    );
  }

  function toggleAllDepartments() {
    const allExpanded = PROCESS_LANES.every((department) =>
      state.expandedDepartments.has(department)
    );

    if (allExpanded) {
      state.expandedDepartments.clear();
    } else {
      PROCESS_LANES.forEach((department) => {
        state.expandedDepartments.add(department);
      });
    }

    rerenderFromState();
  }

  /* ---------- load cycle ---------- */

  async function loadFacilityWip(options = {}) {
    if (state.isLoading) return;

    const initial = options.initial === true;

    setLoading(true, initial);
    setConnectionStatus("loading", "Refreshing");

    try {
      const payload = await fetchWithTimeout(
        buildRequestUrl(),
        CONFIG.FETCH_TIMEOUT_MS
      );

      if (
        !payload ||
        payload.success !== true ||
        String(payload.status || "").toLowerCase() !== "success"
      ) {
        throw new Error(
          payload && payload.message
            ? payload.message
            : "Facility WIP API returned an invalid response."
        );
      }

      validatePayload(payload);
      saveLastGood(payload);

      state.data = payload;

      state.stationRows = applyDisplayOwnershipRules(
        payload.physicalDetail
      );

      renderDashboard(payload, { cached: false });

      setConnectionStatus("connected", "Live Physical");
      hideAlert();
    } catch (error) {
      console.error("Facility WIP load failed:", error);

      /*
       * Cold start with no live data: fall back to the
       * last-known-good cache so wall monitors never sit
       * on an empty loading screen during an outage.
       */
      if (!state.data) {
        const cached = loadLastGood();

        if (cached) {
          const ageMinutes = Math.round(
            (Date.now() - cached.savedAt) / 60000
          );

          state.data = cached.payload;

          state.stationRows = applyDisplayOwnershipRules(
            Array.isArray(cached.payload.physicalDetail)
              ? cached.payload.physicalDetail
              : []
          );

          renderDashboard(cached.payload, { cached: true });

          setConnectionStatus("warning", "Cached Data");

          showAlert(
            "Showing cached data",
            `Live refresh failed. Displaying values from ${ageMinutes} minutes ago. Retrying every 5 minutes.`,
            "warning"
          );

          setLoading(false, initial);
          return;
        }
      }

      setConnectionStatus("error", "Connection Error");

      showAlert(
        "Facility WIP could not refresh",
        error.name === "AbortError"
          ? "The request timed out. The last successful values remain on screen."
          : `${error.message} The last successful values remain on screen.`,
        "error"
      );
    } finally {
      setLoading(false, initial);
    }
  }

  /* ---------- refresh scheduling (single clock) ---------- */

  function scheduleNextRefresh() {
    window.clearTimeout(state.refreshTimer);

    state.nextRefreshAt = Date.now() + CONFIG.AUTO_REFRESH_MS;

    state.refreshTimer = window.setTimeout(() => {
      loadFacilityWip();
      scheduleNextRefresh();
    }, CONFIG.AUTO_REFRESH_MS);
  }

  function updateCountdown() {
    const remainingMs = Math.max(
      state.nextRefreshAt - Date.now(),
      0
    );

    const totalSeconds = Math.ceil(remainingMs / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;

    setText(
      elements.nextRefresh,
      `${minutes}:${String(seconds).padStart(2, "0")}`
    );
  }

  function bindEvents() {
    elements.refreshButton.addEventListener("click", () => {
      scheduleNextRefresh();
      loadFacilityWip();
    });

    elements.expandAllButton.addEventListener(
      "click",
      toggleAllDepartments
    );

    document.addEventListener("visibilitychange", () => {
      if (
        document.visibilityState === "visible" &&
        Date.now() >= state.nextRefreshAt
      ) {
        scheduleNextRefresh();
        loadFacilityWip();
      }
    });
  }

  async function initialize() {
    bindEvents();

    state.countdownTimer = window.setInterval(
      updateCountdown,
      CONFIG.COUNTDOWN_INTERVAL_MS
    );

    scheduleNextRefresh();

    await loadFacilityWip({ initial: true });
  }

  initialize();
})();