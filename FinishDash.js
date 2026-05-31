const FINISH_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?area=Finish";

const FINISH_OPERATOR_API_URL =
  "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec?action=operatorActivity&area=Finish";

/* Personal-profile storage.
   This is NOT global. It saves operator roster and Certified/TQ/Training assignments
   under the logged-in dashboard username in localStorage.
   Director / Manager / Supervisor / LMS can edit. Team Lead and others are view-only. */
const FINISH_PERSONAL_OPERATOR_ROSTER_BASE_KEY = "finishPersonalOperatorRoster_v1";
const FINISH_PERSONAL_OPERATOR_ASSIGNMENTS_BASE_KEY = "finishPersonalOperatorAssignments_v1";
const FINISH_PERSONAL_PROFILE_META_BASE_KEY = "finishPersonalOperatorProfileMeta_v1";
const FINISH_CAPACITY_CONFIG_KEY = "finishCapacityConfig";


/* Login-sheet fallback overrides for testing/personal profile resolution.
   This is still local JS only. It does not save globally.
   Add more users here only if their login page is not saving Role/Username correctly. */
const FINISH_LOGIN_PROFILE_OVERRIDES = {
  /* LMS */
  "BLOPEZ": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },
  "BRIAN LOPEZ CABRERA": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },
  "LOPEZ CABRERA": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },

  /* DIRECTOR */
  "RTATE": { username: "RTATE", role: "SR Director", subRole: "Zenni USA Ohio", firstName: "Rob", lastName: "Tate" },
  "ROB TATE": { username: "RTATE", role: "SR Director", subRole: "Zenni USA Ohio", firstName: "Rob", lastName: "Tate" },
  "TATE": { username: "RTATE", role: "SR Director", subRole: "Zenni USA Ohio", firstName: "Rob", lastName: "Tate" },

  /* MANAGERS */
  "MLITTLE": { username: "MLITTLE", role: "Manager", subRole: "Engineering", firstName: "Matt", lastName: "Little" },
  "MATT LITTLE": { username: "MLITTLE", role: "Manager", subRole: "Engineering", firstName: "Matt", lastName: "Little" },
  "LITTLE": { username: "MLITTLE", role: "Manager", subRole: "Engineering", firstName: "Matt", lastName: "Little" },

  "BKARR": { username: "BKARR", role: "Manager", subRole: "Production", firstName: "Bobby", lastName: "Karr" },
  "BOBBY KARR": { username: "BKARR", role: "Manager", subRole: "Production", firstName: "Bobby", lastName: "Karr" },
  "KARR": { username: "BKARR", role: "Manager", subRole: "Production", firstName: "Bobby", lastName: "Karr" },

  "BBLAKE": { username: "BBLAKE", role: "Manager", subRole: "Manufacturing Operation Program", firstName: "Beth", lastName: "Blake" },
  "BETH BLAKE": { username: "BBLAKE", role: "Manager", subRole: "Manufacturing Operation Program", firstName: "Beth", lastName: "Blake" },
  "BLAKE": { username: "BBLAKE", role: "Manager", subRole: "Manufacturing Operation Program", firstName: "Beth", lastName: "Blake" },

  "SANDERSON": { username: "SANDERSON", role: "SR Manager", subRole: "DC & Inventory", firstName: "Scott", lastName: "Anderson" },
  "SCOTT ANDERSON": { username: "SANDERSON", role: "SR Manager", subRole: "DC & Inventory", firstName: "Scott", lastName: "Anderson" },
  "ANDERSON": { username: "SANDERSON", role: "SR Manager", subRole: "DC & Inventory", firstName: "Scott", lastName: "Anderson" },

  /* SUPERVISORS */
  "AIVANOVSKI": { username: "AIVANOVSKI", role: "Supervisor", subRole: "DC & Inventory", firstName: "Aleks", lastName: "Ivanovski" },
  "ALEKS IVANOVSKI": { username: "AIVANOVSKI", role: "Supervisor", subRole: "DC & Inventory", firstName: "Aleks", lastName: "Ivanovski" },
  "IVANOVSKI": { username: "AIVANOVSKI", role: "Supervisor", subRole: "DC & Inventory", firstName: "Aleks", lastName: "Ivanovski" },

  "CBAYLIS": { username: "CBAYLIS", role: "Supervisor", subRole: "DC & Inventory", firstName: "Crystal", lastName: "Baylis" },
  "CRYSTAL BAYLIS": { username: "CBAYLIS", role: "Supervisor", subRole: "DC & Inventory", firstName: "Crystal", lastName: "Baylis" },
  "BAYLIS": { username: "CBAYLIS", role: "Supervisor", subRole: "DC & Inventory", firstName: "Crystal", lastName: "Baylis" },

  "YKEEBEE": { username: "YKEEBEE", role: "Supervisor", subRole: "DC & Inventory", firstName: "", lastName: "" },

  "KMANACK": { username: "KMANACK", role: "Supervisor", subRole: "Production", firstName: "Kim", lastName: "Manack" },
  "KIM MANACK": { username: "KMANACK", role: "Supervisor", subRole: "Production", firstName: "Kim", lastName: "Manack" },
  "MANACK": { username: "KMANACK", role: "Supervisor", subRole: "Production", firstName: "Kim", lastName: "Manack" },

  "BHECK": { username: "BHECK", role: "Supervisor", subRole: "Production", firstName: "Brittany", lastName: "Heckman" },
  "BRITTANY HECKMAN": { username: "BHECK", role: "Supervisor", subRole: "Production", firstName: "Brittany", lastName: "Heckman" },
  "HECKMAN": { username: "BHECK", role: "Supervisor", subRole: "Production", firstName: "Brittany", lastName: "Heckman" },

  "BHONICKER": { username: "BHONICKER", role: "Supervisor", subRole: "Production", firstName: "Brian", lastName: "Honicker" },
  "BRIAN HONICKER": { username: "BHONICKER", role: "Supervisor", subRole: "Production", firstName: "Brian", lastName: "Honicker" },
  "HONICKER": { username: "BHONICKER", role: "Supervisor", subRole: "Production", firstName: "Brian", lastName: "Honicker" },

  "NPOSTON": { username: "NPOSTON", role: "Supervisor", subRole: "Production", firstName: "Nash", lastName: "Poston" },
  "NASH POSTON": { username: "NPOSTON", role: "Supervisor", subRole: "Production", firstName: "Nash", lastName: "Poston" },
  "POSTON": { username: "NPOSTON", role: "Supervisor", subRole: "Production", firstName: "Nash", lastName: "Poston" },

  /* TRAINING */
  "BDADE": { username: "BDADE", role: "Training", subRole: "Coordinator", firstName: "Brad", lastName: "Dade" },
  "BRAD DADE": { username: "BDADE", role: "Training", subRole: "Coordinator", firstName: "Brad", lastName: "Dade" },
  "DADE": { username: "BDADE", role: "Training", subRole: "Coordinator", firstName: "Brad", lastName: "Dade" },

  "PTOWNSEND": { username: "PTOWNSEND", role: "Training", subRole: "Coordinator", firstName: "Patsy", lastName: "Townsend" },
  "PATSY TOWNSEND": { username: "PTOWNSEND", role: "Training", subRole: "Coordinator", firstName: "Patsy", lastName: "Townsend" },
  "TOWNSEND": { username: "PTOWNSEND", role: "Training", subRole: "Coordinator", firstName: "Patsy", lastName: "Townsend" }
};

function cleanLoginLookupKey(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function applyFinishLoginOverride(profile) {
  if (!profile || typeof profile !== "object") return profile || {};

  const possibleKeys = [
    profile.username,
    profile.Username,
    profile.userName,
    profile.UserName,
    profile.email,
    profile.Email,
    `${profile.firstName || profile.FirstName || profile["First Name"] || ""} ${profile.lastName || profile.LastName || profile["Last Name"] || ""}`,
    profile.lastName,
    profile.LastName,
    profile["Last Name"]
  ];

  for (const value of possibleKeys) {
    const key = cleanLoginLookupKey(value);
    if (key && FINISH_LOGIN_PROFILE_OVERRIDES[key]) {
      return {
        ...profile,
        ...FINISH_LOGIN_PROFILE_OVERRIDES[key]
      };
    }
  }

  return profile;
}

const FINISH_ASSIGNMENT_ALLOWED_EDIT_ROLE_CHECK = role => {
  const value = String(role || "").trim().toLowerCase();
  return (
    value === "lms" ||
    value.includes("director") ||
    value.includes("manager") ||
    value.includes("supervisor") ||
    value.includes("training")
  );
};

function safeParseLoginValue(raw, key = "") {
  if (raw == null || raw === "") return null;

  if (typeof raw === "object") return raw;

  const text = String(raw).trim();
  if (!text) return null;

  try {
    const parsed = JSON.parse(text);
    return parsed;
  } catch {
    // Some login scripts save only the username or role as plain text.
    const cleanKey = String(key || "").toLowerCase();
    if (cleanKey.includes("role")) return { role: text };
    if (cleanKey.includes("user") || cleanKey.includes("login") || cleanKey.includes("name")) {
      return { username: text };
    }
    return null;
  }
}

function readJsonStorageValue(key) {
  const rawSession = sessionStorage.getItem(key);
  const rawLocal = localStorage.getItem(key);

  return safeParseLoginValue(rawSession, key) || safeParseLoginValue(rawLocal, key);
}

function pickLoginField(user, keys) {
  if (!user || typeof user !== "object") return "";

  for (const key of keys) {
    if (user[key] != null && String(user[key]).trim() !== "") {
      return String(user[key]).trim();
    }
  }

  // Support nested objects like { user: { username, role } } or { profile: {...} }.
  const nestedObjects = [user.user, user.profile, user.account, user.auth, user.data, user.currentUser];
  for (const nested of nestedObjects) {
    if (!nested || typeof nested !== "object") continue;
    for (const key of keys) {
      if (nested[key] != null && String(nested[key]).trim() !== "") {
        return String(nested[key]).trim();
      }
    }
  }

  return "";
}

function normalizeLoginUser(candidate) {
  if (!candidate || typeof candidate !== "object") return null;

  const username = pickLoginField(candidate, [
    "username",
    "Username",
    "userName",
    "UserName",
    "login",
    "Login",
    "email",
    "Email",
    "user",
    "User"
  ]);

  const role = pickLoginField(candidate, [
    "role",
    "Role",
    "userRole",
    "UserRole",
    "accessRole",
    "AccessRole"
  ]);

  const subRole = pickLoginField(candidate, [
    "subRole",
    "SubRole",
    "Sub-Role",
    "sub-role",
    "departmentRole",
    "DepartmentRole"
  ]);

  const firstName = pickLoginField(candidate, [
    "firstName",
    "FirstName",
    "First Name",
    "firstname"
  ]);

  const lastName = pickLoginField(candidate, [
    "lastName",
    "LastName",
    "Last Name",
    "lastname"
  ]);

  if (!username && !role && !firstName && !lastName) return null;

  return applyFinishLoginOverride({
    ...candidate,
    username,
    role,
    subRole,
    firstName,
    lastName
  });
}

function readSeparateLoginFields() {
  const getAny = keys => {
    for (const key of keys) {
      const value = sessionStorage.getItem(key) || localStorage.getItem(key);
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  const username = getAny([
    "lmsUsername",
    "LMS_USERNAME",
    "currentUsername",
    "CurrentUsername",
    "username",
    "Username",
    "userName",
    "UserName",
    "loggedInUsername",
    "loggedInUserName"
  ]);

  const role = getAny([
    "lmsRole",
    "LMS_ROLE",
    "currentUserRole",
    "CurrentUserRole",
    "role",
    "Role",
    "userRole",
    "UserRole",
    "loggedInRole"
  ]);

  const subRole = getAny([
    "lmsSubRole",
    "LMS_SUB_ROLE",
    "currentUserSubRole",
    "subRole",
    "SubRole",
    "Sub-Role"
  ]);

  const firstName = getAny([
    "firstName",
    "FirstName",
    "First Name",
    "lmsFirstName"
  ]);

  const lastName = getAny([
    "lastName",
    "LastName",
    "Last Name",
    "lmsLastName"
  ]);

  if (!username && !role) return null;

  return applyFinishLoginOverride({ username, role, subRole, firstName, lastName });
}

function readLoginUserFromDom() {
  const body = document.body || {};
  const dataset = body.dataset || {};

  const username =
    dataset.username ||
    dataset.userName ||
    dataset.lmsUsername ||
    document.querySelector("[data-current-username]")?.getAttribute("data-current-username") ||
    document.querySelector("#currentUsername")?.textContent ||
    "";

  const role =
    dataset.role ||
    dataset.userRole ||
    dataset.lmsRole ||
    document.querySelector("[data-current-role]")?.getAttribute("data-current-role") ||
    document.querySelector("#currentUserRole")?.textContent ||
    "";

  if (!username && !role) return null;
  return applyFinishLoginOverride({ username: String(username).trim(), role: String(role).trim() });
}

function findStoredLoginUser() {
  // 1) If your login page exposes a global user object, use it first.
  const globals = [
    window.lmsCurrentUser,
    window.currentUser,
    window.loggedInUser,
    window.dashboardUser,
    window.LMS_USER,
    window.lmsUser,
    window.userProfile
  ];

  for (const item of globals) {
    const normalized = normalizeLoginUser(item);
    if (normalized) return normalized;
  }

  // 2) Common local/session storage object keys.
  const likelyKeys = [
    "lmsCurrentUser",
    "currentUser",
    "loggedInUser",
    "dashboardUser",
    "LMS_USER",
    "lmsUser",
    "authUser",
    "userProfile",
    "loginUser",
    "loginProfile",
    "sessionUser",
    "activeUser",
    "LMS_CURRENT_USER",
    "LMSCurrentUser"
  ];

  for (const key of likelyKeys) {
    const normalized = normalizeLoginUser(readJsonStorageValue(key));
    if (normalized) return normalized;
  }

  // 3) Separate saved fields like localStorage.Username and localStorage.Role.
  const separate = normalizeLoginUser(readSeparateLoginFields());
  if (separate) return separate;

  // 4) DOM fallback if the login page stamped the page/body with user info.
  const domUser = normalizeLoginUser(readLoginUserFromDom());
  if (domUser) return domUser;

  // 5) Last-resort scan: useful if your login script used a custom key name.
  const stores = [sessionStorage, localStorage];
  for (const store of stores) {
    for (let i = 0; i < store.length; i++) {
      const key = store.key(i);
      if (!key) continue;

      const normalized = normalizeLoginUser(readJsonStorageValue(key));
      if (normalized) return normalized;
    }
  }

  return {};
}

function getCurrentLoggedInUser() {
  const user = findStoredLoginUser();

  // Browser-console test helper. This does not save globally; it only helps you test this browser.
  // Example:
  // window.setFinishLoginProfileForTesting("BLOPEZ", "LMS", "Brian", "Lopez Cabrera");
  window.setFinishLoginProfileForTesting = function(username, role, firstName = "", lastName = "", subRole = "") {
    const profile = { username, role, firstName, lastName, subRole };
    localStorage.setItem("lmsCurrentUser", JSON.stringify(profile));
    console.log("Finish login profile saved for testing:", profile);
    return profile;
  };

  window.debugFinishLoginProfile = function() {
    const profile = findStoredLoginUser();
    console.log("Finish detected login profile:", profile);
    console.log("Can edit Finish operator assignments:", FINISH_ASSIGNMENT_ALLOWED_EDIT_ROLE_CHECK(profile.role || profile.Role));
    return profile;
  };

  window.debugFinishLoginStorage = function() {
    const rows = [];
    [sessionStorage, localStorage].forEach((store, storeIndex) => {
      const storeName = storeIndex === 0 ? "sessionStorage" : "localStorage";
      for (let i = 0; i < store.length; i++) {
        const key = store.key(i);
        rows.push({ store: storeName, key, value: store.getItem(key) });
      }
    });
    console.table(rows);
    return rows;
  };

  return user;
}

function getCurrentUsername() {
  const user = getCurrentLoggedInUser();
  const username = String(
    user.username ||
    user.Username ||
    user.userName ||
    user.UserName ||
    user.email ||
    user.Email ||
    "default"
  ).trim();

  return (username || "default")
    .toUpperCase()
    .replace(/[^A-Z0-9._@-]/g, "_");
}

function getCurrentUserRole() {
  const user = getCurrentLoggedInUser();
  return String(user.role || user.Role || user.userRole || user.UserRole || "").trim();
}

function getCurrentUserDisplayName() {
  const user = getCurrentLoggedInUser();
  const first = user.firstName || user.FirstName || user["First Name"] || "";
  const last = user.lastName || user.LastName || user["Last Name"] || "";
  const full = `${first} ${last}`.trim();
  return full || getCurrentUsername();
}

function canEditFinishOperatorAssignments() {
  return FINISH_ASSIGNMENT_ALLOWED_EDIT_ROLE_CHECK(getCurrentUserRole());
}

function isCurrentUserLmsOnly() {
  const role = String(getCurrentUserRole() || "").trim().toLowerCase();
  return role === "lms";
}

function applyLmsControlTabVisibility() {
  const isLms = isCurrentUserLmsOnly();

  const lmsTabBtn = document.querySelector('.tab-btn[data-tab="lms-control"]');
  const lmsTabContent = document.querySelector('.tab-content[data-content="lms-control"]');

  if (!isLms) {
    if (lmsTabBtn) lmsTabBtn.style.display = "none";
    if (lmsTabContent) lmsTabContent.style.display = "none";

    const activeTab = document.querySelector(".tab-btn.active");
    const activeContent = document.querySelector(".tab-content.active");

    if (activeTab?.dataset.tab === "lms-control") {
      activeTab.classList.remove("active");
      activeContent?.classList.remove("active");

      const morningBtn = document.querySelector('.tab-btn[data-tab="personal"]');
      const morningContent = document.querySelector('.tab-content[data-content="personal"]');
      const floorBtn = document.querySelector('.tab-btn[data-tab="floor"]');
      const floorContent = document.querySelector('.tab-content[data-content="floor"]');

      if (morningBtn && morningContent) {
        morningBtn.classList.add("active");
        morningContent.classList.add("active");
      } else {
        floorBtn?.classList.add("active");
        floorContent?.classList.add("active");
      }
    }
  } else {
    if (lmsTabBtn) lmsTabBtn.style.display = "";
    if (lmsTabContent) lmsTabContent.style.display = "";
  }
}

function getFinishPersonalOperatorRosterKey() {
  return `${FINISH_PERSONAL_OPERATOR_ROSTER_BASE_KEY}_${getCurrentUsername()}`;
}

function getFinishPersonalOperatorAssignmentsKey() {
  return `${FINISH_PERSONAL_OPERATOR_ASSIGNMENTS_BASE_KEY}_${getCurrentUsername()}`;
}

function getFinishPersonalProfileMetaKey() {
  return `${FINISH_PERSONAL_PROFILE_META_BASE_KEY}_${getCurrentUsername()}`;
}

function showFinishAssignmentToast(message) {
  let toast = document.getElementById("finishAssignmentToast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "finishAssignmentToast";
    toast.className = "finish-assignment-toast";
    document.body.appendChild(toast);
  }

  toast.textContent = message;
  toast.classList.add("show");
  window.clearTimeout(showFinishAssignmentToast._timer);
  showFinishAssignmentToast._timer = window.setTimeout(() => {
    toast.classList.remove("show");
  }, 2800);
}

let lastSyncTime = null;
let finishDashboardInitialLoadComplete = false;

const DEFAULT_CONFIG = {
  unboxRate: 150,
  unboxCount: 1,

  meiARate: 50,
  meiACount: 0,
  meiBRate: 50,
  meiBCount: 5,
  meiCRate: 50,
  meiCCount: 5,

  mountRate: 25,
  mountCount: 10,

 mountTraineeW1: 0,
 mountTraineeW2: 0,
 mountTraineeW3: 0,
 mountTraineeW4: 0,
 mountTraineeW5: 0,
 mountTraineeW6: 0,
 mountTraineeW7: 0,
 mountTraineeW8: 0, 

  drillRate: 6,
  drillCount: 1,
 finalRate: 75,
finalCount: 3,

finalTraineeW1: 0,
finalTraineeW2: 0,
finalTraineeW3: 0,
finalTraineeW4: 0,
finalTraineeW5: 0,

  /* Operator-based capacity assignments.
     role: ignore | core | tq | training
     trainingWeek applies only to Mounting and Final Inspection. */
  operatorAssignments: {}

};

const SHIFT_RULES = {
  weekday: {
    shiftStart: "7:00 AM",
    shiftEnd: "5:30 PM",
    breaks: [
      { name: "Morning Break", start: "9:30 AM", end: "9:45 AM" },
      { name: "Lunch Group 1", start: "11:30 AM", end: "12:00 PM" },
      { name: "Lunch Group 2", start: "12:00 PM", end: "12:30 PM" },
      { name: "Afternoon Break", start: "3:00 PM", end: "3:15 PM" }
    ]
  },

  weekend: {
    shiftStart: "6:30 AM",
    shiftEnd: "6:30 PM",
    breaks: [
      { name: "Morning Break", start: "9:30 AM", end: "9:50 AM" },
      { name: "Lunch Group 1", start: "11:30 AM", end: "12:00 PM" },
      { name: "Lunch Group 2", start: "12:00 PM", end: "12:30 PM" },
      { name: "Afternoon Break", start: "3:30 PM", end: "3:50 PM" }
    ]
  }
};

document.addEventListener("DOMContentLoaded", () => {
  initTabs();
  applyLmsControlTabVisibility();
  initClock();
  initConfigPanel();
  initRefreshButton();
  initHoverEffects();
  initOperatorCommandUI();
  initFinishThreeLineCell();
  initMorningSetupControls();

  loadDashboard({ showLoader: true });
  loadFinishOperatorActivity();

  setInterval(() => loadDashboard({ showLoader: false }), 5 * 60 * 1000);
  setInterval(loadFinishOperatorActivity, 5 * 60 * 1000);
  setInterval(updateSyncAge, 1000);
});

/* ===============================
   MAIN LOAD
================================ */

async function loadDashboard(options = {}) {
  const shouldShowLoader = options.showLoader !== false && !finishDashboardInitialLoadComplete;

  if (shouldShowLoader) {
    showLoading(true);
  }

  try {
    const response = await fetch(FINISH_API_URL, {
      method: "GET",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`API failed with status ${response.status}`);
    }

    const data = await response.json();

    if (data.finishDashboard?.allStations) {
      renderEnrichedDashboard(data.finishDashboard);
    } else {
      renderLegacyDashboard(data);
    }

    lastSyncTime = Date.now();
    setText("syncAge", "just now");
  } catch (error) {
    console.error("Finish Dashboard Error:", error);
    renderError(error);
  } finally {
    finishDashboardInitialLoadComplete = true;

    if (shouldShowLoader) {
      showLoading(false);
    }
  }
}

/* ===============================
   ENRICHED DASHBOARD
================================ */

function renderEnrichedDashboard(finishDashboard) {
  const config = loadConfig();

  const floorStations = (finishDashboard.allStations || []).map(station =>
    applyConfigToStation({ ...station }, config)
  );

  const healthStations = (finishDashboard.stationHealth || []).map(station =>
    applyConfigToStation({ ...station }, config)
  );

  const summary =
    finishDashboard.summary ||
    buildSummaryFromStations(floorStations);

  renderKpis(summary);
  renderStationNumbers(floorStations);
  renderFlowCore(summary, floorStations);

  renderStatusPanel(healthStations);
  renderStationDetail(healthStations);
  renderManagementSummary(summary, healthStations);
  renderHourlyBreakdown(healthStations);

  applyStationStateClasses(floorStations);

  setText("summaryUpdated", formatTime(new Date()));
}

function renderKpis(summary) {
  setText("kpiIncomingWip", summary.incomingWip || 0);
  setText("kpiMeiFeedWip", summary.meiFeedWip || 0);
  setText("kpiMeiBankWip", summary.edgingWip || 0);
  setText("kpiMountDrillWip", summary.mountingDrillWip || 0);
  setText("kpiFinalWip", summary.finalInspectionWip || 0);

  setText("summaryTotalWip", summary.totalWip || 0);
  setText("summaryLargestValue", summary.largestWipTotal || 0);
  setText("summaryLargestName", summary.largestWipStation || "--");
}


function renderStationNumbers(stations) {
  const get = name => stations.find(s => s.flowStep === name || s.displayName === name) || {};

  const fsv = get("FSV Scan & Verify");
  const frameOnly = get("Frame Only Scan & Verify");
  const unbox = get("Finish Unbox");
  const arout = get("AR-OUT");

  const meiB = get("MEI Line B");
  const meiC = get("MEI Line C");
  const easy = get("MEI Easy Fit");

  const mounting = get("Mounting");
  const drill = get("Drill");
  const bigs = get("Bigs");
  const sharps = get("Sharps");
  const final = get("Final Inspection");

  setText("fsvWip", fsv.currentWip || 0);
  setText("frameOnlyWip", frameOnly.currentWip || 0);

  setText("unboxWip", unbox.currentWip || 0);
  setText("unboxCnt", unbox.activityToday || 0);

  setText("aroutWip", arout.currentWip || 0);

  setText("meiBWip", meiB.currentWip || 0);
  setText("meiBCnt", meiB.activityToday || 0);

  setText("easyWip", easy.currentWip || 0);
  setText("easyCnt", easy.activityToday || 0);

  setText("meiCWip", meiC.currentWip || 0);
  setText("meiCCnt", meiC.activityToday || 0);

  setText("mountingWip", mounting.currentWip || 0);
  setText("mountingCnt", mounting.activityToday || 0);
  setText("mountingWipCenter", mounting.currentWip || 0);
  setText("mountingCntCenter", mounting.activityToday || 0);

  setText("drillWip", drill.currentWip || 0);
  setText("drillCnt", drill.activityToday || 0);
  setText("drillWipCenter", drill.currentWip || 0);
  setText("drillCntCenter", drill.activityToday || 0);

  setText("bigsWip", bigs.currentWip || 0);
  setText("bigsCnt", bigs.activityToday || 0);

  setText("sharpsWip", sharps.currentWip || 0);
  setText("sharpsCnt", sharps.activityToday || 0);

  setText("finalWip", final.currentWip || 0);
  setText("finalCnt", final.activityToday || 0);
}

function renderFlowCore(summary, stations) {
  const totalWip = Number(summary.totalWip || 0);

  const get = name => (stations || []).find(s => s.flowStep === name || s.displayName === name) || {};
  const finishWip =
    Number(get("Finish Unbox").currentWip || 0) +
    Number(get("AR-OUT").currentWip || 0) +
    Number(get("MEI Line B").currentWip || 0) +
    Number(get("MEI Line C").currentWip || 0) +
    Number(get("MEI Easy Fit").currentWip || 0) +
    Number(get("Mounting").currentWip || 0) +
    Number(get("Drill").currentWip || 0) +
    Number(get("Bigs").currentWip || 0) +
    Number(get("Sharps").currentWip || 0) +
    Number(get("Final Inspection").currentWip || 0);

  setText("totalWipValue", totalWip);
  setText("finishWipValue", finishWip);
}

/* ===============================
   LEGACY FALLBACK
================================ */

function renderLegacyDashboard(data) {
  const rows =
    data.productionFlow ||
    data.finishFlow ||
    data.areaSummary ||
    data.stations ||
    [];

  const stations = normalizeLegacyStations(rows);
  const summary = buildSummaryFromStations(stations);

  renderKpis(summary);
  renderStationNumbers(stations);
  renderFlowCore(summary, stations);
  renderStatusPanel(stations);
  renderStationDetail(stations);
  renderHourlyBreakdown(stations);
  applyStationStateClasses(stations);
}

function normalizeLegacyStations(rows) {
  if (!Array.isArray(rows)) return [];

  return rows.map(row => {
    const name =
      row.flowStep ||
      row.displayName ||
      row.station ||
      row.Station ||
      row.name ||
      "";

    const currentWip =
      row.currentWip ??
      row.wip ??
      row.WIP ??
      row.CurrentWip ??
      0;

    const activityToday =
      row.activityToday ??
      row.countToday ??
      row.cnt ??
      row.CNT ??
      row.totalToday ??
      0;

    const lastHourActivity =
      row.lastHourActivity ??
      row.lastHour ??
      row.LastHour ??
      0;

    const expectedNormalPerHour =
      row.expectedNormalPerHour ??
      row.capacityPerHour ??
      row.ratePerHour ??
      0;

    return {
      flowStep: normalizeStationName(name),
      displayName: normalizeStationName(name),
      currentWip: Number(currentWip) || 0,
      activityToday: Number(activityToday) || 0,
      lastHourActivity: Number(lastHourActivity) || 0,
      expectedNormalPerHour: Number(expectedNormalPerHour) || 0,
      metricMode: row.metricMode || "STANDARD",
      hourly: row.hourly || row.Hours || {}
    };
  }).map(station => {
    station.utilizationPct = calculateUtilization(station);
    station.status = calculateStatus(station);
    station.statusLabel = getStatusLabel(station.status);
    station.statusMessage = getStatusMessage(station);
    return station;
  });
}

function normalizeStationName(name) {
  const value = String(name || "").trim().toLowerCase();

  if (value.includes("frame only")) return "Frame Only Scan & Verify";
  if (value.includes("fsv")) return "FSV Scan & Verify";
  if (value.includes("unbox")) return "Finish Unbox";
  if (value.includes("ar-out") || value.includes("ar out")) return "AR-OUT";
  if (value.includes("line b")) return "MEI Line B";
  if (value.includes("line c")) return "MEI Line C";
  if (value.includes("easy")) return "MEI Easy Fit";
  if (value.includes("mount")) return "Mounting";
  if (value.includes("drill")) return "Drill";
  if (value.includes("big")) return "Bigs";
  if (value.includes("sharp")) return "Sharps";
  if (value.includes("final")) return "Final Inspection";

  return name || "--";
}

function buildSummaryFromStations(stations) {
  const get = name => stations.find(s => s.flowStep === name) || {};

  const incomingWip =
    Number(get("FSV Scan & Verify").currentWip || 0) +
    Number(get("Frame Only Scan & Verify").currentWip || 0);

  const meiFeedWip =
    Number(get("Finish Unbox").currentWip || 0) +
    Number(get("AR-OUT").currentWip || 0);

  const edgingWip =
    Number(get("MEI Line B").currentWip || 0) +
    Number(get("MEI Line C").currentWip || 0) +
    Number(get("MEI Easy Fit").currentWip || 0);

  const mountingDrillWip =
    Number(get("Mounting").currentWip || 0) +
    Number(get("Drill").currentWip || 0);

  const finalInspectionWip =
    Number(get("Final Inspection").currentWip || 0);

  const sideStationWip =
    Number(get("Bigs").currentWip || 0) +
    Number(get("Sharps").currentWip || 0);

  const totalWip = stations.reduce((sum, s) => sum + Number(s.currentWip || 0), 0);
  const totalActivityToday = stations.reduce((sum, s) => sum + Number(s.activityToday || 0), 0);

  const largest = stations.reduce((best, s) => {
    return Number(s.currentWip || 0) > Number(best.currentWip || 0) ? s : best;
  }, { displayName: "--", currentWip: 0 });

  return {
    incomingWip,
    meiFeedWip,
    edgingWip,
    mountingDrillWip,
    finalInspectionWip,
    sideStationWip,
    totalWip,
    totalActivityToday,
    largestWipTotal: largest.currentWip || 0,
    largestWipStation: largest.displayName || largest.flowStep || "--"
  };
}


function renderManagementSummary(summary, stations) {
  const safeStations = Array.isArray(stations) ? stations : [];

  const riskStations = safeStations.filter(s =>
    ["CRITICAL", "WARNING", "STARVED", "OVERLOAD"].includes(String(s.status || "").toUpperCase())
  );

  const criticalStations = safeStations.filter(s =>
    String(s.status || "").toUpperCase() === "CRITICAL"
  );

  const bottleneck = safeStations.reduce((best, station) => {
    return Number(station.currentWip || 0) > Number(best.currentWip || 0) ? station : best;
  }, { displayName: "--", currentWip: 0 });

  const totalWip = Number(summary.totalWip || 0);
  const bottleneckName = bottleneck.displayName || bottleneck.flowStep || summary.largestWipStation || "--";
  const bottleneckWip = Number(bottleneck.currentWip || summary.largestWipTotal || 0);

  setText("summaryRiskCount", riskStations.length);

  if (criticalStations.length > 0) {
    setText("mgmtSummaryTitle", "Immediate attention required");
    setText(
      "mgmtSummaryText",
      `${criticalStations.length} station(s) are currently below expected performance. The largest WIP pressure is ${bottleneckName} with ${formatNumber(bottleneckWip)} jobs.`
    );
    setText(
      "mgmtRecommendation",
      `Prioritize support around ${bottleneckName}. Review staffing, machine availability, and upstream feed before WIP continues building.`
    );
    setText(
      "mgmtNextAction",
      `Check ${criticalStations.map(s => s.displayName || s.flowStep).join(", ")} and confirm whether the issue is labor, equipment, or scan timing.`
    );
    return;
  }

  if (riskStations.length > 0) {
    setText("mgmtSummaryTitle", "Operation stable but needs monitoring");
    setText(
      "mgmtSummaryText",
      `${riskStations.length} station(s) are trending below target. Total WIP is ${formatNumber(totalWip)}, with the highest load at ${bottleneckName}.`
    );
    setText(
      "mgmtRecommendation",
      `Monitor ${bottleneckName} and confirm the next hourly update improves before shifting more work downstream.`
    );
    setText(
      "mgmtNextAction",
      "Use the Hourly Breakdown tab to confirm whether the slowdown is isolated to one hour or becoming a trend."
    );
    return;
  }

  setText("mgmtSummaryTitle", "Operation running within control");
  setText(
    "mgmtSummaryText",
    `Finish is currently operating within expected range. Total WIP is ${formatNumber(totalWip)}, and the largest active WIP point is ${bottleneckName}.`
  );
  setText(
    "mgmtRecommendation",
    "Maintain current staffing plan and continue watching the bottleneck station during the next refresh cycle."
  );
  setText(
    "mgmtNextAction",
    "No immediate escalation needed. Continue monitoring WIP movement and hourly pace."
  );
}

/* ===============================
   STATUS SYSTEM
================================ */

function applyConfigToStation(station, config) {
  const cfg = config || loadConfig();
  const flowStep = normalizeFinishOperatorStation(station.flowStep || station.displayName || station.name);
  station.flowStep = flowStep;

  switch (flowStep) {
    case "Finish Unbox": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.unboxCount || 0);
      station.expectedNormalPerHour = Number(cfg.unboxRate || 0) * count;
      break;
    }

    case "MEI Line A": {
      const count = Number(cfg.meiACount || 0);
      station.expectedNormalPerHour = Number(cfg.meiARate || 0) * count;
      break;
    }

    case "MEI Line B": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiBCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiBRate || 0) * count;
      break;
    }

    case "MEI Line C": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiCCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiCRate || 0) * count;
      break;
    }

    case "MEI Easy Fit": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.meiBCount || 0);
      station.expectedNormalPerHour = Number(cfg.meiBRate || 0) * count;
      break;
    }

    case "Mounting": {
      const assignedCoreCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const coreCount = assignedCoreCount > 0 ? assignedCoreCount : Number(cfg.mountCount || 0);
      const coreTotal = Number(cfg.mountRate || 0) * coreCount;
      const assignedTrainingTotal = getConfiguredTrainingOperatorTotal(flowStep, cfg);
      const numericTrainingTotal = getNumericMountTrainingTotal(cfg);
      station.expectedNormalPerHour = coreTotal + (assignedTrainingTotal > 0 ? assignedTrainingTotal : numericTrainingTotal);
      break;
    }

    case "Drill": {
      const assignedCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const count = assignedCount > 0 ? assignedCount : Number(cfg.drillCount || 0);
      station.expectedNormalPerHour = Number(cfg.drillRate || 0) * count;
      break;
    }

    case "Final Inspection": {
      const assignedCoreCount = getConfiguredCoreOperatorCount(flowStep, cfg);
      const coreCount = assignedCoreCount > 0 ? assignedCoreCount : Number(cfg.finalCount || 0);
      const coreTotal = Number(cfg.finalRate || 0) * coreCount;
      const assignedTrainingTotal = getConfiguredTrainingOperatorTotal(flowStep, cfg);
      const numericTrainingTotal = getNumericFinalTrainingTotal(cfg);
      station.expectedNormalPerHour = coreTotal + (assignedTrainingTotal > 0 ? assignedTrainingTotal : numericTrainingTotal);
      break;
    }

    case "Bigs":
    case "Sharps":
      station.expectedNormalPerHour = 0;
      break;

    default:
      station.expectedNormalPerHour = Number(station.expectedNormalPerHour || 0);
  }

  station.currentWip = Number(station.currentWip || 0);
  station.activityToday = Number(station.activityToday || 0);
  station.lastHourActivity = Number(station.lastHourActivity || 0);

  station.utilizationPct = calculateUtilization(station);
  station.status = calculateStatus(station);
  station.statusLabel = getStatusLabel(station.status);
  station.statusMessage = getStatusMessage(station);

  return station;
}


function calculateUtilization(station) {
  const capacity = Number(station.expectedNormalPerHour || 0);
  const lastHour = Number(station.lastHourActivity || 0);

  if (!capacity || capacity <= 0) return 0;

  return Math.round((lastHour / capacity) * 100);
}

function calculateStatus(station) {
  const wip = Number(station.currentWip || 0);
  const util = Number(station.utilizationPct || 0);
  const capacity = Number(station.expectedNormalPerHour || 0);

  if (station.metricMode === "WIP_ONLY") return "NORMAL";

  if (wip > 0 && util === 0 && capacity > 0) return "STARVED";
  if (util < 30 && capacity > 0) return "CRITICAL";
  if (util < 70 && capacity > 0) return "WARNING";
  if (util > 150) return "OVERLOAD";

  return "NORMAL";
}

function getStatusLabel(status) {
  switch (status) {
    case "CRITICAL": return "Critical";
    case "WARNING": return "Warning";
    case "STARVED": return "Starved";
    case "OVERLOAD": return "Overload";
    case "NORMAL": return "Normal";
    default: return "Normal";
  }
}

function getStatusMessage(station) {
  const name = station.displayName || station.flowStep || "Station";

  switch (station.status) {
    case "CRITICAL":
      return `${name} is running below expected threshold. Review staffing, feed, or station availability.`;

    case "WARNING":
      return `${name} is below target pace. Monitor before it becomes a bottleneck.`;

    case "STARVED":
      return `${name} has WIP but no recent throughput signal. Validate scan activity or station movement.`;

    case "OVERLOAD":
      return `${name} is above normal utilization. Watch for downstream buildup.`;

    default:
      return `${name} is operating within normal range.`;
  }
}

function renderStatusPanel(stations) {
  const container = document.getElementById("statusPanelContainer");
  if (!container) return;

  const groups = {
    CRITICAL: stations.filter(s => s.status === "CRITICAL"),
    WARNING: stations.filter(s => s.status === "WARNING"),
    STARVED: stations.filter(s => s.status === "STARVED"),
    OVERLOAD: stations.filter(s => s.status === "OVERLOAD"),
    NORMAL: stations.filter(s => s.status === "NORMAL")
  };

  container.innerHTML = `
    ${buildStatusGroup("Critical Alerts", groups.CRITICAL, "critical")}
    ${buildStatusGroup("Warnings", groups.WARNING, "warning")}
    ${buildStatusGroup("Starved Stations", groups.STARVED, "starved")}
    ${buildStatusGroup("Overload Watch", groups.OVERLOAD, "overload")}
    ${buildStatusGroup("Normal Operations", groups.NORMAL, "normal")}
  `;
}

function buildStatusGroup(title, stations, className) {
  if (!stations.length) return "";

  return `
    <div class="status-group ${className}">
      <div class="status-group-title">${title}</div>
      <div class="status-grid">
        ${stations.map(buildStatusCard).join("")}
      </div>
    </div>
  `;
}

function buildStatusCard(station) {
  const statusClass = String(station.status || "NORMAL").toLowerCase();
  const util = Number(station.utilizationPct || 0);
  const capped = Math.max(0, Math.min(util, 100));

  return `
    <article class="status-card status-${statusClass}">
      <div class="status-card-top">
        <h3>${escapeHtml(station.displayName || station.flowStep || "--")}</h3>
        <span>${escapeHtml(station.statusLabel || "Normal")}</span>
      </div>

      <div class="status-metrics">
        <div>
          <span>WIP</span>
          <strong>${formatNumber(station.currentWip || 0)}</strong>
        </div>
        <div>
          <span>Today</span>
          <strong>${formatNumber(station.activityToday || 0)}</strong>
        </div>
        <div>
          <span>Last Hr</span>
          <strong>${formatNumber(station.lastHourActivity || 0)}</strong>
        </div>
      </div>

      <div class="util-bar">
        <i style="width:${capped}%"></i>
      </div>

      <div class="util-copy">
        <span>${util}% utilization</span>
        <span>${formatNumber(station.expectedNormalPerHour || 0)}/hr target</span>
      </div>

      <p>${escapeHtml(station.statusMessage || "")}</p>
    </article>
  `;
}

function renderStationDetail(stations) {
  const body = document.getElementById("stationDetailBody");
  if (!body) return;

  if (!stations.length) {
    body.innerHTML = `<tr><td colspan="6">No station data available.</td></tr>`;
    return;
  }

  body.innerHTML = stations.map(station => `
    <tr>
      <td>${escapeHtml(station.displayName || station.flowStep || "--")}</td>
      <td>${formatNumber(station.currentWip || 0)}</td>
      <td>${formatNumber(station.activityToday || 0)}</td>
      <td>${formatNumber(station.lastHourActivity || 0)}</td>
      <td>${formatNumber(station.expectedNormalPerHour || 0)}</td>
      <td><span class="table-status ${String(station.status || "NORMAL").toLowerCase()}">${escapeHtml(station.statusLabel || "Normal")}</span></td>
    </tr>
  `).join("");
}

/* ===============================
   HOURLY BREAKDOWN
================================ */

function renderHourlyBreakdown(stations) {
  const container = document.getElementById("hourlyBreakdownGrid");
  if (!container) return;

  const safeStations = Array.isArray(stations) ? stations : [];

  const hourOrder = [
    "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM",
    "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM",
    "4:00 PM", "5:00 PM", "6:00 PM"
  ];

  if (!safeStations.length) {
    container.innerHTML = `
      <article class="hourly-station-card">
        <div class="hourly-station-header">
          <h3>No hourly data available</h3>
          <span>--</span>
        </div>
      </article>
    `;
    return;
  }

  container.innerHTML = safeStations.map((station, index) => {
    const baseTarget = Number(station.expectedNormalPerHour || 0);
    const stationName = station.displayName || station.flowStep || "--";
    const isNoTarget = station.flowStep === "Bigs" || station.flowStep === "Sharps" || baseTarget <= 0;

    const cells = hourOrder.map(hour => {
      const count = Number(station.hourly?.[hour] || 0);
      const productiveMinutes = getProductiveMinutesForHour(hour);
      const adjustedTarget = isNoTarget ? 0 : Math.round(baseTarget * (productiveMinutes / 60));
      const pct = adjustedTarget > 0 ? Math.round((count / adjustedTarget) * 100) : 0;

      let statusClass = "hour-red";

      if (isNoTarget || productiveMinutes <= 0) {
        statusClass = "hour-neutral";
      } else if (pct >= 100) {
        statusClass = "hour-green";
      } else if (pct >= 90) {
        statusClass = "hour-amber";
      }

      return `
        <div class="hour-cell ${statusClass}" title="${escapeHtml(stationName)} · ${escapeHtml(hour)}">
          <span>${escapeHtml(hour)}</span>
          <strong>${formatNumber(count)}</strong>
          <small>${isNoTarget ? "No Target" : `${pct}%`}</small>
          <em>${isNoTarget ? "Count Only" : `${formatNumber(adjustedTarget)} target`}</em>
        </div>
      `;
    }).join("");

    return `
      <details class="hourly-accordion" ${index === 0 ? "open" : ""}>
        <summary class="hourly-accordion-header">
          <div>
            <h3>${escapeHtml(stationName)}</h3>
            <small>
              Today: ${formatNumber(station.activityToday || station.hourlyTotal || 0)}
              · Last Hr: ${formatNumber(station.lastHourActivity || 0)}
            </small>
          </div>

          <div class="hourly-accordion-right">
            <span>${isNoTarget ? "No hourly target" : `${formatNumber(baseTarget)} / hr target`}</span>
            <b>Open / Close</b>
          </div>
        </summary>

        <div class="hourly-cells">
          ${cells}
        </div>
      </details>
    `;
  }).join("");
}

function getCurrentShiftRule() {
  const day = new Date().getDay();
  const isWeekend = day === 0 || day === 6;
  return isWeekend ? SHIFT_RULES.weekend : SHIFT_RULES.weekday;
}

function getProductiveMinutesForHour(hourLabel) {
  const rule = getCurrentShiftRule();

  const hourStart = timeToMinutes(hourLabel);
  const hourEnd = hourStart + 60;

  const shiftStart = timeToMinutes(rule.shiftStart);
  const shiftEnd = timeToMinutes(rule.shiftEnd);

  let productiveStart = Math.max(hourStart, shiftStart);
  let productiveEnd = Math.min(hourEnd, shiftEnd);

  if (productiveEnd <= productiveStart) return 0;

  let productiveMinutes = productiveEnd - productiveStart;

  rule.breaks.forEach(breakItem => {
    const breakStart = timeToMinutes(breakItem.start);
    const breakEnd = timeToMinutes(breakItem.end);

    const overlapStart = Math.max(productiveStart, breakStart);
    const overlapEnd = Math.min(productiveEnd, breakEnd);

    if (overlapEnd > overlapStart) {
      productiveMinutes -= overlapEnd - overlapStart;
    }
  });

  return Math.max(0, productiveMinutes);
}

function timeToMinutes(label) {
  const text = String(label || "").trim().toUpperCase();
  const match = text.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/);

  if (!match) return 0;

  let hour = Number(match[1]);
  const minute = Number(match[2] || 0);
  const suffix = match[3];

  if (suffix === "PM" && hour !== 12) hour += 12;
  if (suffix === "AM" && hour === 12) hour = 0;

  return hour * 60 + minute;
}

/* ===============================
   FLOOR STATE CLASSES
================================ */

function applyStationStateClasses(stations) {
  const map = {
    "FSV Scan & Verify": ".station-fsv",
    "Finish Unbox": ".station-unbox",
    "AR-OUT": ".station-arout",
    "MEI Line B": ".mei-b",
    "MEI Easy Fit": ".easy-fit",
    "MEI Line C": ".mei-c",
    "Mounting": ".process-card:nth-child(1), .rail-card.orange-card:nth-of-type(1)",
    "Drill": ".process-card:nth-child(2), .rail-card.orange-card:nth-of-type(2)",
    "Bigs": ".purple-card:nth-of-type(3)",
    "Sharps": ".purple-card:nth-of-type(4)",
    "Final Inspection": ".final-card"
  };

  document
    .querySelectorAll(".state-normal, .state-warning, .state-critical, .state-starved, .state-overload")
    .forEach(el => {
      el.classList.remove(
        "state-normal",
        "state-warning",
        "state-critical",
        "state-starved",
        "state-overload"
      );
    });

  stations.forEach(station => {
    const selector = map[station.flowStep];
    if (!selector) return;

    document.querySelectorAll(selector).forEach(el => {
      el.classList.add(`state-${String(station.status || "NORMAL").toLowerCase()}`);
    });
  });
}

/* ===============================
   CONFIG
================================ */

function loadFinishPersonalOperatorAssignments() {
  try {
    return JSON.parse(localStorage.getItem(getFinishPersonalOperatorAssignmentsKey()) || "{}");
  } catch (error) {
    console.warn("Could not read personal Finish operator assignments:", error);
    return {};
  }
}

function saveFinishPersonalOperatorAssignments(assignments) {
  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role can see operator setup but cannot edit it.");
    return false;
  }

  try {
    localStorage.setItem(getFinishPersonalOperatorAssignmentsKey(), JSON.stringify(assignments || {}));
    localStorage.setItem(getFinishPersonalProfileMetaKey(), JSON.stringify({
      savedAt: new Date().toISOString(),
      storage: "localStorage",
      scope: "personal-login-profile",
      username: getCurrentUsername(),
      role: getCurrentUserRole()
    }));
    return true;
  } catch (error) {
    console.warn("Could not save personal Finish operator assignments:", error);
    return false;
  }
}

function loadConfig() {
  try {
    const saved = localStorage.getItem(FINISH_CAPACITY_CONFIG_KEY);
    const baseConfig = saved ? { ...DEFAULT_CONFIG, ...JSON.parse(saved) } : { ...DEFAULT_CONFIG };
    return {
      ...baseConfig,
      operatorAssignments: loadFinishPersonalOperatorAssignments()
    };
  } catch {
    return {
      ...DEFAULT_CONFIG,
      operatorAssignments: loadFinishPersonalOperatorAssignments()
    };
  }
}

function saveConfig(config) {
  const safeConfig = { ...(config || {}) };
  delete safeConfig.operatorAssignments;
  localStorage.setItem(FINISH_CAPACITY_CONFIG_KEY, JSON.stringify(safeConfig));
}

function initConfigPanel() {
  const config = loadConfig();

  setInputValue("cfgUnboxRate", config.unboxRate);
  setInputValue("cfgUnboxCount", config.unboxCount);

  setInputValue("cfgMeiARate", config.meiARate);
  setInputValue("cfgMeiACount", config.meiACount);
  setInputValue("cfgMeiBRate", config.meiBRate);
  setInputValue("cfgMeiBCount", config.meiBCount);
  setInputValue("cfgMeiCRate", config.meiCRate);
  setInputValue("cfgMeiCCount", config.meiCCount);

  setInputValue("cfgMountRate", config.mountRate);
  setInputValue("cfgMountCount", config.mountCount);
  setInputValue("cfgMountTraineeW1", config.mountTraineeW1);
 setInputValue("cfgMountTraineeW2", config.mountTraineeW2);
 setInputValue("cfgMountTraineeW3", config.mountTraineeW3);
 setInputValue("cfgMountTraineeW4", config.mountTraineeW4);
 setInputValue("cfgMountTraineeW5", config.mountTraineeW5);
 setInputValue("cfgMountTraineeW6", config.mountTraineeW6);
 setInputValue("cfgMountTraineeW7", config.mountTraineeW7);
 setInputValue("cfgMountTraineeW8", config.mountTraineeW8);

  setInputValue("cfgDrillRate", config.drillRate);
  setInputValue("cfgDrillCount", config.drillCount);
  setInputValue("cfgFinalRate", config.finalRate);
  setInputValue("cfgFinalCount", config.finalCount);
  setInputValue("cfgFinalTraineeW1", config.finalTraineeW1);
setInputValue("cfgFinalTraineeW2", config.finalTraineeW2);
setInputValue("cfgFinalTraineeW3", config.finalTraineeW3);
setInputValue("cfgFinalTraineeW4", config.finalTraineeW4);
setInputValue("cfgFinalTraineeW5", config.finalTraineeW5);

  ensureOperatorAssignmentConfigPanel();
  renderOperatorAssignmentConfigPanel();
  updateConfigTotals();

  [
    "cfgUnboxRate",
    "cfgUnboxCount",
    "cfgMeiARate",
    "cfgMeiACount",
    "cfgMeiBRate",
    "cfgMeiBCount",
    "cfgMeiCRate",
    "cfgMeiCCount",
    "cfgMountRate",
    "cfgMountTraineeW1",
    "cfgMountTraineeW2",
    "cfgMountTraineeW3",
    "cfgMountTraineeW4",
    "cfgMountTraineeW5",
    "cfgMountTraineeW6",
    "cfgMountTraineeW7",
    "cfgMountTraineeW8",
    "cfgMountCount",
    "cfgDrillRate",
    "cfgDrillCount",
    "cfgFinalRate",
"cfgFinalCount",
"cfgFinalTraineeW1",
"cfgFinalTraineeW2",
"cfgFinalTraineeW3",
"cfgFinalTraineeW4",
"cfgFinalTraineeW5"

  ].forEach(id => {
    document.getElementById(id)?.addEventListener("input", updateConfigTotals);
  });

  document.getElementById("configSaveBtn")?.addEventListener("click", () => {
    const newConfig = {
      unboxRate: getInputValue("cfgUnboxRate"),
      unboxCount: getInputValue("cfgUnboxCount"),

      meiARate: getInputValue("cfgMeiARate"),
      meiACount: getInputValue("cfgMeiACount"),
      meiBRate: getInputValue("cfgMeiBRate"),
      meiBCount: getInputValue("cfgMeiBCount"),
      meiCRate: getInputValue("cfgMeiCRate"),
      meiCCount: getInputValue("cfgMeiCCount"),

      mountRate: getInputValue("cfgMountRate"),
      mountCount: getInputValue("cfgMountCount"),
      mountTraineeW1: getInputValue("cfgMountTraineeW1"),
mountTraineeW2: getInputValue("cfgMountTraineeW2"),
mountTraineeW3: getInputValue("cfgMountTraineeW3"),
mountTraineeW4: getInputValue("cfgMountTraineeW4"),
mountTraineeW5: getInputValue("cfgMountTraineeW5"),
mountTraineeW6: getInputValue("cfgMountTraineeW6"),
mountTraineeW7: getInputValue("cfgMountTraineeW7"),
mountTraineeW8: getInputValue("cfgMountTraineeW8"),
      drillRate: getInputValue("cfgDrillRate"),
      drillCount: getInputValue("cfgDrillCount"),
      finalRate: getInputValue("cfgFinalRate"),
finalCount: getInputValue("cfgFinalCount"),

finalTraineeW1: getInputValue("cfgFinalTraineeW1"),
finalTraineeW2: getInputValue("cfgFinalTraineeW2"),
finalTraineeW3: getInputValue("cfgFinalTraineeW3"),
finalTraineeW4: getInputValue("cfgFinalTraineeW4"),
finalTraineeW5: getInputValue("cfgFinalTraineeW5"),
      operatorAssignments: loadConfig().operatorAssignments || {}

    };

    saveConfig(newConfig);
    flashButton("configSaveBtn", "Saved");
    loadDashboard({ showLoader: false });
  });

  document.getElementById("configResetBtn")?.addEventListener("click", () => {
    localStorage.removeItem(FINISH_CAPACITY_CONFIG_KEY);
    initConfigPanel();
    loadDashboard({ showLoader: false });
  });
}

function updateConfigTotals() {
  const saved = loadConfig();
  const config = {
    ...saved,
    unboxRate: getInputValue("cfgUnboxRate"),
    unboxCount: getInputValue("cfgUnboxCount"),
    meiARate: getInputValue("cfgMeiARate"),
    meiACount: getInputValue("cfgMeiACount"),
    meiBRate: getInputValue("cfgMeiBRate"),
    meiBCount: getInputValue("cfgMeiBCount"),
    meiCRate: getInputValue("cfgMeiCRate"),
    meiCCount: getInputValue("cfgMeiCCount"),
    mountRate: getInputValue("cfgMountRate"),
    mountCount: getInputValue("cfgMountCount"),
    mountTraineeW1: getInputValue("cfgMountTraineeW1"),
    mountTraineeW2: getInputValue("cfgMountTraineeW2"),
    mountTraineeW3: getInputValue("cfgMountTraineeW3"),
    mountTraineeW4: getInputValue("cfgMountTraineeW4"),
    mountTraineeW5: getInputValue("cfgMountTraineeW5"),
    mountTraineeW6: getInputValue("cfgMountTraineeW6"),
    mountTraineeW7: getInputValue("cfgMountTraineeW7"),
    mountTraineeW8: getInputValue("cfgMountTraineeW8"),
    drillRate: getInputValue("cfgDrillRate"),
    drillCount: getInputValue("cfgDrillCount"),
    finalRate: getInputValue("cfgFinalRate"),
    finalCount: getInputValue("cfgFinalCount"),
    finalTraineeW1: getInputValue("cfgFinalTraineeW1"),
    finalTraineeW2: getInputValue("cfgFinalTraineeW2"),
    finalTraineeW3: getInputValue("cfgFinalTraineeW3"),
    finalTraineeW4: getInputValue("cfgFinalTraineeW4"),
    finalTraineeW5: getInputValue("cfgFinalTraineeW5"),
    operatorAssignments: saved.operatorAssignments || {}
  };

  const unboxAssigned = getConfiguredCoreOperatorCount("Finish Unbox", config);
  const unboxCount = unboxAssigned > 0 ? unboxAssigned : config.unboxCount;
  const unboxTotal = Number(config.unboxRate || 0) * Number(unboxCount || 0);

  const meiATotal = Number(config.meiARate || 0) * Number(config.meiACount || 0);
  const meiBAssigned = getConfiguredCoreOperatorCount("MEI Line B", config);
  const meiBAssignedCount = meiBAssigned > 0 ? meiBAssigned : config.meiBCount;
  const meiBTotal = Number(config.meiBRate || 0) * Number(meiBAssignedCount || 0);

  const meiCAssigned = getConfiguredCoreOperatorCount("MEI Line C", config);
  const meiCAssignedCount = meiCAssigned > 0 ? meiCAssigned : config.meiCCount;
  const meiCTotal = Number(config.meiCRate || 0) * Number(meiCAssignedCount || 0);

  const mountAssignedCore = getConfiguredCoreOperatorCount("Mounting", config);
  const mountCount = mountAssignedCore > 0 ? mountAssignedCore : config.mountCount;
  const mountTotal = Number(config.mountRate || 0) * Number(mountCount || 0);
  const assignedMountTrainingTotal = getConfiguredTrainingOperatorTotal("Mounting", config);
  const trainingMountTotal = assignedMountTrainingTotal > 0
    ? assignedMountTrainingTotal
    : getNumericMountTrainingTotal(config);

  const drillAssigned = getConfiguredCoreOperatorCount("Drill", config);
  const drillCount = drillAssigned > 0 ? drillAssigned : config.drillCount;
  const drillTotal = Number(config.drillRate || 0) * Number(drillCount || 0);

  const finalAssignedCore = getConfiguredCoreOperatorCount("Final Inspection", config);
  const finalCount = finalAssignedCore > 0 ? finalAssignedCore : config.finalCount;
  const finalTotal = Number(config.finalRate || 0) * Number(finalCount || 0);
  const assignedFinalTrainingTotal = getConfiguredTrainingOperatorTotal("Final Inspection", config);
  const finalTrainingTotal = assignedFinalTrainingTotal > 0
    ? assignedFinalTrainingTotal
    : getNumericFinalTrainingTotal(config);

  setText("cfgUnboxTotal", unboxTotal);
  setText("cfgMeiATotal", meiATotal);
  setText("cfgMeiBTotal", meiBTotal);
  setText("cfgMeiCTotal", meiCTotal);
  setText("cfgMountTotal", mountTotal);
  setText("cfgMountTrainingTotal", trainingMountTotal);
  setText("cfgMountAdjustedTotal", mountTotal + trainingMountTotal);
  setText("cfgDrillTotal", drillTotal);
  setText("cfgFinalTotal", finalTotal);
  setText("cfgFinalTrainingTotal", finalTrainingTotal);
  setText("cfgFinalAdjustedTotal", finalTotal + finalTrainingTotal);

  setText("cfgAssignedMountCore", mountAssignedCore);
  setText("cfgAssignedMountTraining", getConfiguredTrainingOperatorCount("Mounting", config));
  setText("cfgAssignedFinalCore", finalAssignedCore);
  setText("cfgAssignedFinalTraining", getConfiguredTrainingOperatorCount("Final Inspection", config));
}



/* ===============================
   OPERATOR-BASED CAPACITY ASSIGNMENTS
   UI is injected into the existing Configuration tab.
   No backend change required. Saved in localStorage with finishCapacityConfig.
================================ */

const FINISH_ASSIGNMENT_STATIONS = [
  "Finish Unbox",
  "MEI Line B",
  "MEI Line C",
  "MEI Easy Fit",
  "Mounting",
  "Drill",
  "Final Inspection"
];

const MOUNTING_TRAINING_RATES = {
  1: 3,
  2: 6,
  3: 9,
  4: 12,
  5: 15,
  6: 18,
  7: 21,
  8: 25
};

const FINAL_TRAINING_RATES = {
  1: 15,
  2: 30,
  3: 45,
  4: 60,
  5: 75
};

function getNumericMountTrainingTotal(config = loadConfig()) {
  return (
    (Number(config.mountTraineeW1 || 0) * 3) +
    (Number(config.mountTraineeW2 || 0) * 6) +
    (Number(config.mountTraineeW3 || 0) * 9) +
    (Number(config.mountTraineeW4 || 0) * 12) +
    (Number(config.mountTraineeW5 || 0) * 15) +
    (Number(config.mountTraineeW6 || 0) * 18) +
    (Number(config.mountTraineeW7 || 0) * 21) +
    (Number(config.mountTraineeW8 || 0) * 25)
  );
}

function getNumericFinalTrainingTotal(config = loadConfig()) {
  return (
    (Number(config.finalTraineeW1 || 0) * 15) +
    (Number(config.finalTraineeW2 || 0) * 30) +
    (Number(config.finalTraineeW3 || 0) * 45) +
    (Number(config.finalTraineeW4 || 0) * 60) +
    (Number(config.finalTraineeW5 || 0) * 75)
  );
}

function normalizeAssignmentOperatorName(name) {
  return String(name || "").trim();
}

function makeAssignmentKey(stationName, operatorName) {
  return `${normalizeFinishOperatorStation(stationName)}||${normalizeAssignmentOperatorName(operatorName)}`;
}

function getOperatorAssignment(operatorName, stationName, config = loadConfig()) {
  const op = normalizeAssignmentOperatorName(operatorName);
  if (!op) return null;
  const assignments = config.operatorAssignments || {};
  return assignments[makeAssignmentKey(stationName, op)] || null;
}

function getTrainingRateForStationWeek(stationName, week) {
  const name = normalizeFinishOperatorStation(stationName);
  const safeWeek = Number(week || 0);

  if (name === "Mounting") return Number(MOUNTING_TRAINING_RATES[safeWeek] || 0);
  if (name === "Final Inspection") return Number(FINAL_TRAINING_RATES[safeWeek] || 0);

  return 0;
}

function getConfiguredCoreOperatorCount(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.filter(item =>
    normalizeFinishOperatorStation(item.station) === name &&
    ["core", "tq", "certified"].includes(String(item.role || "").toLowerCase())
  ).length;
}

function getConfiguredTrainingOperatorCount(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.filter(item =>
    normalizeFinishOperatorStation(item.station) === name &&
    String(item.role || "").toLowerCase() === "training"
  ).length;
}

function getConfiguredTrainingOperatorTotal(stationName, config = loadConfig()) {
  const name = normalizeFinishOperatorStation(stationName);
  const assignments = Object.values(config.operatorAssignments || {});

  return assignments.reduce((sum, item) => {
    if (normalizeFinishOperatorStation(item.station) !== name) return sum;
    if (String(item.role || "").toLowerCase() !== "training") return sum;
    return sum + getTrainingRateForStationWeek(name, item.trainingWeek);
  }, 0);
}

function getAssignmentRoleLabel(operatorName, stationName, config = loadConfig()) {
  const assignment = getOperatorAssignment(operatorName, stationName, config);
  if (!assignment || assignment.role === "ignore") return "Unassigned";
  if (assignment.role === "training") return `Training W${Number(assignment.trainingWeek || 1)}`;
  if (assignment.role === "tq") return "TQ";
  return "Certified";
}

function getAssignmentTargetLabel(operatorName, stationName, config = loadConfig()) {
  const base = getFinishOperatorBaseRate(stationName, operatorName);
  const role = getAssignmentRoleLabel(operatorName, stationName, config);
  return `${role} · ${numberFmt(base)}/hr`;
}

function loadFinishPersonalOperatorRoster() {
  try {
    return JSON.parse(localStorage.getItem(getFinishPersonalOperatorRosterKey()) || "{}");
  } catch (error) {
    console.warn("Could not read personal Finish operator roster:", error);
    return {};
  }
}

function saveFinishPersonalOperatorRoster(roster) {
  try {
    localStorage.setItem(getFinishPersonalOperatorRosterKey(), JSON.stringify(roster || {}));
    localStorage.setItem(getFinishPersonalProfileMetaKey(), JSON.stringify({
      savedAt: new Date().toISOString(),
      storage: "localStorage",
      scope: "personal-browser-profile"
    }));
  } catch (error) {
    console.warn("Could not save personal Finish operator roster:", error);
  }
}

function upsertFinishPersonalRosterOperator(stationName, operatorName, total = 0, accessPoints = []) {
  const station = normalizeFinishOperatorStation(stationName);
  const operator = normalizeAssignmentOperatorName(operatorName);

  if (!operator || operator === "Unassigned / No Operator") return;
  if (!isFinishOperatorStationAllowed(station)) return;

  const roster = loadFinishPersonalOperatorRoster();
  const key = makeAssignmentKey(station, operator);
  const now = new Date().toISOString();
  const existing = roster[key] || {};

  const nextAccessPoints = Array.from(new Set([
    ...((existing.accessPoints || []).filter(Boolean)),
    ...((Array.isArray(accessPoints) ? accessPoints : [accessPoints]).filter(Boolean))
  ]));

  roster[key] = {
    station,
    operator,
    total: Number(total || 0),
    accessPoints: nextAccessPoints,
    firstSeen: existing.firstSeen || now,
    lastSeen: now
  };

  saveFinishPersonalOperatorRoster(roster);
}

function syncFinishPersonalRosterFromStations(stations) {
  Object.values(stations || {}).forEach(station => {
    (station.operatorList || []).forEach(operator => {
      upsertFinishPersonalRosterOperator(
        station.name,
        operator.name,
        operator.total,
        operator.accessPoints || []
      );
    });
  });
}

function getFinishOperatorRoster() {
  const stations = finishOperatorState?.stations || {};
  const savedRoster = loadFinishPersonalOperatorRoster();
  const roster = [];
  const seen = new Set();

  // 1) Start with saved personal-profile roster so previously seen operators stay available.
  Object.values(savedRoster || {}).forEach(item => {
    const stationName = normalizeFinishOperatorStation(item.station);
    const operatorName = normalizeAssignmentOperatorName(item.operator);
    if (!stationName || !operatorName || !isFinishOperatorStationAllowed(stationName)) return;
    if (!FINISH_ASSIGNMENT_STATIONS.includes(stationName)) return;

    const key = makeAssignmentKey(stationName, operatorName);
    if (seen.has(key)) return;
    seen.add(key);

    roster.push({
      station: stationName,
      operator: operatorName,
      total: Number(item.total || 0),
      accessPoints: item.accessPoints || [],
      firstSeen: item.firstSeen || "",
      lastSeen: item.lastSeen || "",
      savedProfile: true
    });
  });

  // 2) Overlay current live roster totals, so active operators show today's live count.
  FINISH_ASSIGNMENT_STATIONS.forEach(stationName => {
    const station = stations[stationName];
    (station?.operatorList || []).forEach(operator => {
      const key = makeAssignmentKey(stationName, operator.name);
      const liveItem = {
        station: stationName,
        operator: operator.name,
        total: Number(operator.total || 0),
        accessPoints: operator.accessPoints || [],
        firstSeen: "",
        lastSeen: new Date().toISOString(),
        savedProfile: true,
        liveNow: true
      };

      if (seen.has(key)) {
        const idx = roster.findIndex(item => makeAssignmentKey(item.station, item.operator) === key);
        if (idx >= 0) roster[idx] = { ...roster[idx], ...liveItem, firstSeen: roster[idx].firstSeen || liveItem.firstSeen };
        return;
      }

      seen.add(key);
      roster.push(liveItem);
    });
  });

  return roster.sort((a, b) => {
    const s = FINISH_ASSIGNMENT_STATIONS.indexOf(a.station) - FINISH_ASSIGNMENT_STATIONS.indexOf(b.station);
    if (s !== 0) return s;
    if (Boolean(b.liveNow) !== Boolean(a.liveNow)) return Number(Boolean(b.liveNow)) - Number(Boolean(a.liveNow));
    return Number(b.total || 0) - Number(a.total || 0);
  });
}

function ensureOperatorAssignmentConfigPanel() {
  if (document.getElementById("operatorAssignmentConfigPanel")) return;

  injectOperatorAssignmentStyles();

  const configTab = document.querySelector('[data-content="config"]');
  const saveRow = document.querySelector(".config-actions") || document.getElementById("configSaveBtn")?.parentElement;
  const targetParent = saveRow?.parentElement || configTab;
  if (!targetParent) return;

  const panel = document.createElement("details");
  panel.id = "operatorAssignmentConfigPanel";
  panel.className = "config-group operator-assignment-config-group";
  panel.open = true;
  panel.innerHTML = `
    <summary>
      <div>
        <h3>Operator Capacity Assignment</h3>
        <p>Select who counts as Certified/TQ core capacity and who is on Training Metric. Saved to this browser profile only.</p>
      </div>
      <span>Open / Close</span>
    </summary>

    <div class="operator-assignment-shell">
      <div class="assignment-warning-card">
        <strong>Personal Login Profile</strong>        
      </div>

      <div id="finishAssignmentPermissionNotice" class="assignment-permission-notice"></div>

      <div class="assignment-summary-grid">
        <article><span>Mount Core/TQ</span><strong id="cfgAssignedMountCore">0</strong></article>
        <article><span>Mount Training</span><strong id="cfgAssignedMountTraining">0</strong></article>
        <article><span>Final Core/TQ</span><strong id="cfgAssignedFinalCore">0</strong></article>
        <article><span>Final Training</span><strong id="cfgAssignedFinalTraining">0</strong></article>
      </div>

      <div class="assignment-toolbar">
        <label>
          Station Filter
          <select id="assignmentStationFilter">
            <option value="all">All Finish stations</option>
            ${FINISH_ASSIGNMENT_STATIONS.map(station => `<option value="${escapeHtml(station)}">${escapeHtml(station)}</option>`).join("")}
          </select>
        </label>
        <button id="assignmentClearBtn" type="button">Clear Assignments</button>
        <button id="assignmentClearRosterBtn" type="button">Clear Saved Roster</button>
      </div>

      <div id="operatorAssignmentRoster" class="assignment-roster"></div>
    </div>
  `;

  if (saveRow && saveRow.parentElement) {
    saveRow.parentElement.insertBefore(panel, saveRow);
  } else {
    targetParent.appendChild(panel);
  }

  document.getElementById("assignmentStationFilter")?.addEventListener("change", renderOperatorAssignmentConfigPanel);
  document.getElementById("assignmentClearBtn")?.addEventListener("click", () => {
    if (!canEditFinishOperatorAssignments()) {
      showFinishAssignmentToast("View only. Your role can see operator setup but cannot edit it.");
      return;
    }
    saveFinishPersonalOperatorAssignments({});
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    loadDashboard();
    if (finishOperatorState?.selectedStation) {
      const station = finishOperatorState.stations?.[finishOperatorState.selectedStation];
      if (station) renderDrawerOperators(station);
    }
  });

  document.getElementById("assignmentClearRosterBtn")?.addEventListener("click", () => {
    if (!canEditFinishOperatorAssignments()) {
      showFinishAssignmentToast("View only. Your role can see operator setup but cannot edit it.");
      return;
    }
    localStorage.removeItem(getFinishPersonalOperatorRosterKey());
    localStorage.removeItem(getFinishPersonalProfileMetaKey());
    syncFinishPersonalRosterFromStations(finishOperatorState.stations || {});
    renderOperatorAssignmentConfigPanel();
  });
}

function renderOperatorAssignmentConfigPanel() {
  ensureOperatorAssignmentConfigPanel();

  const rosterTarget = document.getElementById("operatorAssignmentRoster");
  if (!rosterTarget) return;

  const config = loadConfig();
  const filter = document.getElementById("assignmentStationFilter")?.value || "all";
  const roster = getFinishOperatorRoster()
    .filter(item => filter === "all" || item.station === filter);

  if (!roster.length) {
    rosterTarget.innerHTML = `
      <article class="assignment-empty">
        No live Finish operator roster loaded yet. Click Refresh Operators on the Operator tab or wait for the API refresh.
      </article>`;
    updateConfigTotals();
    applyFinishAssignmentPermissionLock();
    return;
  }

  rosterTarget.innerHTML = roster.map(item => {
    const assignment = getOperatorAssignment(item.operator, item.station, config) || {
      role: "ignore",
      trainingWeek: 1
    };
    const canTrain = item.station === "Mounting" || item.station === "Final Inspection";
    const role = String(assignment.role || "ignore");
    const trainingRate = getTrainingRateForStationWeek(item.station, assignment.trainingWeek || 1);
    const baseRate = getFinishOperatorBaseRate(item.station, item.operator);
    const effective = role === "training" ? trainingRate : (["core", "tq", "certified"].includes(role) ? baseRate : 0);

    return `
      <article class="assignment-row" data-assignment-row="${escapeHtml(makeAssignmentKey(item.station, item.operator))}">
        <div class="assignment-operator">
          <strong>${escapeHtml(item.operator)}</strong>
          <span>${escapeHtml(item.station)} · Today ${numberFmt(item.total)}${item.liveNow ? " · Live" : " · Saved"}</span>
        </div>

        <label>
          Role
          <select data-assignment-role data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}">
            <option value="ignore" ${role === "ignore" ? "selected" : ""}>Do not count</option>
            <option value="core" ${role === "core" || role === "certified" ? "selected" : ""}>Certified / Core Capacity</option>
            <option value="tq" ${role === "tq" ? "selected" : ""}>TQ / Core Capacity</option>
            ${canTrain ? `<option value="training" ${role === "training" ? "selected" : ""}>Training Metric</option>` : ""}
          </select>
        </label>

        <label class="assignment-week ${role === "training" && canTrain ? "active" : "disabled"}">
          Week
          <select data-assignment-week data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}" ${role === "training" && canTrain ? "" : "disabled"}>
            ${buildTrainingWeekOptions(item.station, assignment.trainingWeek || 1)}
          </select>
        </label>

        ${buildLineSlotAssignmentControls(item.station, item.operator, assignment)}

        <div class="assignment-target">
          <span>Target</span>
          <strong>${effective > 0 ? numberFmt(effective) + "/hr" : "—"}</strong>
        </div>
      </article>`;
  }).join("");

  rosterTarget.querySelectorAll("[data-assignment-role]").forEach(select => {
    select.addEventListener("change", event => {
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        event.target.value,
        Number(getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station)?.trainingWeek || 1)
      );
    });
  });

  rosterTarget.querySelectorAll("[data-assignment-week]").forEach(select => {
    select.addEventListener("change", event => {
      const current = getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station) || { role: "training" };
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        current.role || "training",
        Number(event.target.value || 1),
        current.line || "",
        current.slot || ""
      );
    });
  });

  rosterTarget.querySelectorAll("[data-assignment-line], [data-assignment-slot]").forEach(select => {
    select.addEventListener("change", event => {
      const station = event.target.dataset.station;
      const operator = event.target.dataset.operator;
      const row = event.target.closest(".assignment-row");
      const current = getOperatorAssignment(operator, station) || { role: "ignore", trainingWeek: 1 };
      const line = row?.querySelector("[data-assignment-line]")?.value || current.line || "";
      const slot = row?.querySelector("[data-assignment-slot]")?.value || current.slot || "";
      saveOperatorAssignment(
        station,
        operator,
        current.role || "ignore",
        Number(current.trainingWeek || 1),
        line,
        slot
      );
    });
  });

  updateConfigTotals();
  applyFinishAssignmentPermissionLock();
  renderFinishThreeLineCell();
}


async function saveConfigAssignedIndividualJph(rowEl) {
  if (!rowEl) return;

  const operatorName = rowEl.dataset.assignedRosterRow || "";
  const source = (window.__finishConfigAssignedRosterRows || []).find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "") === normalizeAssignmentOperatorName(operatorName) &&
    normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(rowEl.dataset.area || rowEl.querySelector(".assignment-operator span")?.textContent?.split(" · ")?.[0] || "")
  );

  const areaLinePositionText = rowEl.querySelector(".assignment-operator span")?.textContent || "";
  const parts = areaLinePositionText.split(" · ").map(x => String(x || "").trim());

  const defaultArea = source?.defaultArea || parts[0] || "";
  const defaultLine = source?.defaultLine || parts[1] || "";
  const defaultPosition = source?.defaultPosition || parts[2] || "";
  const individualJph = rowEl.querySelector("[data-config-jph-input]")?.value || "";

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getFinishConfigOwnerUsername(),
      shiftType: getFinishConfigSelectedShiftType(),
      operatorName,
      activeStatus: source?.activeStatus || "Active",
      defaultLine,
      defaultArea,
      defaultPosition,
      roleType: source?.roleType || "Unassigned",
      trainingWeek: source?.roleType === "Training" ? (source?.trainingWeek || "") : "",
      individualJph,
      originalDefaultArea: defaultArea,
      originalDefaultLine: defaultLine,
      originalDefaultPosition: defaultPosition
    });

    showFinishAssignmentToast(`Saved JPH for ${operatorName}`);
    await renderOperatorAssignmentConfigPanel();
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

function applyFinishAssignmentPermissionLock() {
  const canEdit = canEditFinishOperatorAssignments();
  const role = getCurrentUserRole() || "Unknown role";
  const user = getCurrentUserDisplayName();

  document
    .querySelectorAll("#operatorAssignmentConfigPanel select, #operatorAssignmentConfigPanel input, #operatorAssignmentConfigPanel button")
    .forEach(el => {
      // Station filter should stay enabled so Team Leads can still view/filter.
      if (el.id === "assignmentStationFilter") {
        el.disabled = false;
        return;
      }

      el.disabled = !canEdit;
      el.classList.toggle("locked", !canEdit);
      el.title = canEdit ? "" : "View only for your role";
    });

  const notice = document.getElementById("finishAssignmentPermissionNotice");
  if (notice) {
    notice.classList.toggle("can-edit", canEdit);
    notice.classList.toggle("view-only", !canEdit);
    notice.innerHTML = canEdit
      ? `<strong>Edit access active</strong><span>${escapeHtml(user)} · ${escapeHtml(role)} · personal setup saves under ${escapeHtml(getCurrentUsername())}</span>`
      : `<strong>View only</strong><span>${escapeHtml(user)} · ${escapeHtml(role)} · Director, Manager, Supervisor, and LMS can edit. Team Lead and other roles can only view.</span>`;
  }
}

function buildTrainingWeekOptions(stationName, selectedWeek) {
  const rates = normalizeFinishOperatorStation(stationName) === "Final Inspection"
    ? FINAL_TRAINING_RATES
    : MOUNTING_TRAINING_RATES;

  return Object.keys(rates).map(week => `
    <option value="${week}" ${Number(selectedWeek) === Number(week) ? "selected" : ""}>
      Week ${week} · ${rates[week]}/hr
    </option>
  `).join("");
}

function saveOperatorAssignment(stationName, operatorName, role, trainingWeek = 1, line = "", slot = "") {
  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role can see operator setup but cannot edit it.");
    renderOperatorAssignmentConfigPanel();
    return;
  }

  const config = loadConfig();
  const station = normalizeFinishOperatorStation(stationName);
  const operator = normalizeAssignmentOperatorName(operatorName);
  if (!operator) return;

  const assignments = { ...(config.operatorAssignments || {}) };
  const key = makeAssignmentKey(station, operator);

  if (!role || role === "ignore") {
    delete assignments[key];
  } else {
    assignments[key] = {
      station,
      operator,
      role,
      trainingWeek: Number(trainingWeek || 1),
      line: line || getOperatorAssignment(operator, station)?.line || "",
      slot: slot || getOperatorAssignment(operator, station)?.slot || "",
      updatedAt: new Date().toISOString(),
      updatedBy: getCurrentUsername(),
      updatedByRole: getCurrentUserRole()
    };
  }

  const saved = saveFinishPersonalOperatorAssignments(assignments);
  if (!saved) return;

  renderOperatorAssignmentConfigPanel();
  updateConfigTotals();
  loadDashboard();
  renderFinishThreeLineCell();

  if (finishOperatorState?.selectedStation) {
    const activeStation = finishOperatorState.stations?.[finishOperatorState.selectedStation];
    if (activeStation) renderDrawerOperators(activeStation);
  }
}

function injectOperatorAssignmentStyles() {
  if (document.getElementById("operatorAssignmentStyles")) return;

  const style = document.createElement("style");
  style.id = "operatorAssignmentStyles";
  style.textContent = `
    .operator-assignment-shell {
      display: grid;
      gap: 16px;
    }

    .assignment-warning-card {
      display: grid;
      gap: 5px;
      padding: 16px;
      border: 1px solid rgba(251, 191, 36, .28);
      border-radius: 16px;
      background: rgba(251, 191, 36, .07);
      color: #eef8ff;
    }

    .assignment-warning-card strong {
      color: #fbbf24;
      letter-spacing: .08em;
      text-transform: uppercase;
      font-size: 12px;
    }

    .assignment-warning-card span {
      color: #bfdaf2;
      font-size: 13px;
      line-height: 1.45;
    }

    .assignment-summary-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }

    .assignment-summary-grid article {
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .16);
      border-radius: 16px;
      background: rgba(2, 12, 24, .64);
    }

    .assignment-summary-grid span,
    .assignment-toolbar label,
    .assignment-row label,
    .assignment-target span {
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 800;
      letter-spacing: .1em;
      text-transform: uppercase;
    }

    .assignment-summary-grid strong {
      display: block;
      margin-top: 8px;
      color: #67e8f9;
      font-size: 28px;
      line-height: 1;
    }

    .assignment-toolbar {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 14px;
      flex-wrap: wrap;
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .14);
      border-radius: 16px;
      background: rgba(3, 12, 24, .56);
    }

    .assignment-toolbar label,
    .assignment-row label {
      display: grid;
      gap: 8px;
    }

    .assignment-toolbar select,
    .assignment-row select {
      height: 42px;
      min-width: 150px;
      border-radius: 12px;
      border: 1px solid rgba(0, 217, 255, .28);
      background: #061225;
      color: #eef8ff;
      padding: 0 12px;
      font-weight: 800;
      outline: none;
    }

    .assignment-toolbar button {
      height: 42px;
      padding: 0 16px;
      border-radius: 12px;
      border: 1px solid rgba(251, 113, 133, .35);
      background: rgba(251, 113, 133, .08);
      color: #fb7185;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
      cursor: pointer;
    }

    .assignment-roster {
      display: grid;
      gap: 10px;
    }

    .assignment-row {
      display: grid;
      grid-template-columns: minmax(220px, 1.15fr) minmax(190px, .72fr) minmax(160px, .62fr) minmax(145px, .55fr) minmax(160px, .62fr) 96px;
      align-items: center;
      gap: 12px;
      padding: 14px;
      border: 1px solid rgba(100, 221, 255, .14);
      border-radius: 16px;
      background: linear-gradient(180deg, rgba(10, 28, 52, .72), rgba(2, 10, 22, .88));
    }

    .assignment-operator strong {
      display: block;
      color: #eef8ff;
      font-size: 14px;
      font-weight: 900;
    }

    .assignment-operator span {
      display: block;
      margin-top: 5px;
      color: #8fa6c3;
      font-size: 12px;
    }

    .assignment-week.disabled {
      opacity: .38;
    }

    .assignment-target {
      display: grid;
      justify-items: end;
      gap: 6px;
    }

    .assignment-target strong {
      color: #00f5a0;
      font-size: 18px;
    }

    .assignment-empty {
      padding: 18px;
      border: 1px dashed rgba(100, 221, 255, .25);
      border-radius: 16px;
      color: #8fa6c3;
      background: rgba(2, 12, 24, .52);
    }

    .assignment-permission-notice {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .16);
      background: rgba(2, 12, 24, .62);
    }

    .assignment-permission-notice strong {
      color: #eef8ff;
      font-size: 13px;
      font-weight: 900;
      text-transform: uppercase;
      letter-spacing: .08em;
    }

    .assignment-permission-notice span {
      color: #8fa6c3;
      font-size: 12px;
      text-align: right;
    }

    .assignment-permission-notice.can-edit {
      border-color: rgba(0, 245, 160, .28);
      background: rgba(0, 245, 160, .07);
    }

    .assignment-permission-notice.can-edit strong {
      color: #00f5a0;
    }

    .assignment-permission-notice.view-only {
      border-color: rgba(251, 191, 36, .28);
      background: rgba(251, 191, 36, .07);
    }

    .assignment-permission-notice.view-only strong {
      color: #fbbf24;
    }

    #operatorAssignmentConfigPanel select.locked,
    #operatorAssignmentConfigPanel input.locked,
    #operatorAssignmentConfigPanel button.locked {
      opacity: .48;
      cursor: not-allowed;
      filter: grayscale(.35);
    }

    .finish-assignment-toast {
      position: fixed;
      right: 22px;
      bottom: 22px;
      z-index: 10050;
      max-width: 360px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(251, 191, 36, .35);
      background: rgba(3, 12, 24, .96);
      color: #eef8ff;
      box-shadow: 0 18px 42px rgba(0, 0, 0, .42);
      transform: translateY(14px);
      opacity: 0;
      pointer-events: none;
      transition: .24s ease;
      font-size: 13px;
      font-weight: 800;
    }

    .finish-assignment-toast.show {
      transform: translateY(0);
      opacity: 1;
    }

    /* ===== professional + animated polish (assigned roster) ===== */

    /* KPI summary cards */
    .assignment-summary-grid article {
      position: relative;
      overflow: hidden;
      transition: transform .2s ease, border-color .2s ease, box-shadow .2s ease;
      animation: cfgCardIn .5s cubic-bezier(.22, 1, .36, 1) both;
    }
    .assignment-summary-grid article::after {
      content: "";
      position: absolute;
      left: 14px; right: 14px; top: 0;
      height: 2px;
      background: linear-gradient(90deg, #00f5a0, #00d9ff);
      opacity: .55;
    }
    .assignment-summary-grid article:nth-child(2)::after { background: linear-gradient(90deg, #00d9ff, #38bdf8); }
    .assignment-summary-grid article:nth-child(3)::after { background: linear-gradient(90deg, #b06cff, #00d9ff); }
    .assignment-summary-grid article:nth-child(4)::after { background: linear-gradient(90deg, #ff9f1a, #00f5a0); }
    .assignment-summary-grid article:hover {
      transform: translateY(-2px);
      border-color: rgba(0, 217, 255, .4);
      box-shadow: 0 12px 26px rgba(0, 0, 0, .38), 0 0 22px rgba(0, 217, 255, .08);
    }
    .assignment-summary-grid strong { text-shadow: 0 0 16px rgba(0, 217, 255, .35); }
    .assignment-summary-grid article:nth-child(1) { animation-delay: .02s; }
    .assignment-summary-grid article:nth-child(2) { animation-delay: .07s; }
    .assignment-summary-grid article:nth-child(3) { animation-delay: .12s; }
    .assignment-summary-grid article:nth-child(4) { animation-delay: .17s; }

    /* roster row: fix to 4 columns, add hover + accent + staggered entrance */
    .assignment-row {
      grid-template-columns: minmax(240px, 1.7fr) minmax(150px, .85fr) minmax(140px, .75fr) minmax(210px, 1fr);
      position: relative;
      overflow: hidden;
      transition: transform .2s ease, border-color .2s ease, box-shadow .2s ease;
      animation: cfgRowIn .45s cubic-bezier(.22, 1, .36, 1) both;
    }
    .assignment-row::before {
      content: "";
      position: absolute;
      left: 0; top: 0; bottom: 0;
      width: 3px;
      background: linear-gradient(180deg, #00f5a0, #00d9ff);
      opacity: 0;
      transition: opacity .2s ease;
    }
    .assignment-row:hover {
      transform: translateY(-2px);
      border-color: rgba(0, 217, 255, .38);
      box-shadow: 0 14px 30px rgba(0, 0, 0, .42), 0 0 22px rgba(0, 217, 255, .08);
    }
    .assignment-row:hover::before { opacity: .9; }
    .assignment-row:nth-child(1) { animation-delay: .03s; }
    .assignment-row:nth-child(2) { animation-delay: .06s; }
    .assignment-row:nth-child(3) { animation-delay: .09s; }
    .assignment-row:nth-child(4) { animation-delay: .12s; }
    .assignment-row:nth-child(5) { animation-delay: .15s; }
    .assignment-row:nth-child(6) { animation-delay: .18s; }
    .assignment-row:nth-child(7) { animation-delay: .21s; }
    .assignment-row:nth-child(8) { animation-delay: .24s; }
    .assignment-row:nth-child(9) { animation-delay: .27s; }
    .assignment-row:nth-child(n+10) { animation-delay: .3s; }

    @keyframes cfgRowIn {
      from { opacity: 0; transform: translateY(10px); }
      to   { opacity: 1; transform: translateY(0); }
    }
    @keyframes cfgCardIn {
      from { opacity: 0; transform: translateY(8px) scale(.98); }
      to   { opacity: 1; transform: translateY(0) scale(1); }
    }

    /* field values */
    .assignment-field { display: grid; gap: 7px; align-content: start; }
    .assignment-output-value { color: #8fa6c3; font-size: 20px; font-weight: 900; }
    .assignment-output-value.has-output { color: #00f5a0; text-shadow: 0 0 14px rgba(0, 245, 160, .35); }

    /* role badge */
    .assignment-role-badge {
      display: inline-flex;
      align-items: center;
      align-self: start;
      padding: 5px 11px;
      border-radius: 999px;
      border: 1px solid rgba(100, 221, 255, .3);
      background: rgba(0, 217, 255, .08);
      color: #aef3ff;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .06em;
      text-transform: uppercase;
      white-space: nowrap;
    }
    .assignment-role-badge[data-role*="certified"],
    .assignment-role-badge[data-role*="core"],
    .assignment-role-badge[data-role*="tq"] {
      border-color: rgba(0, 245, 160, .4);
      background: rgba(0, 245, 160, .1);
      color: #7dffce;
    }
    .assignment-role-badge[data-role*="training"] {
      border-color: rgba(176, 108, 255, .4);
      background: rgba(176, 108, 255, .1);
      color: #d4b6ff;
    }
    .assignment-role-badge[data-role*="unassigned"],
    .assignment-role-badge[data-role*="ignore"] {
      border-color: rgba(143, 166, 195, .28);
      background: rgba(143, 166, 195, .08);
      color: #aabfda;
    }

    /* JPH cell: label on top, input + save side by side */
    .assignment-target {
      grid-template-columns: 1fr auto;
      grid-template-areas: "lbl lbl" "inp save";
      justify-items: stretch;
      align-items: center;
      gap: 7px 8px;
    }
    .assignment-target > span { grid-area: lbl; justify-self: start; }
    .assignment-target > strong { grid-area: inp; justify-self: end; }

    .assignment-jph-input {
      grid-area: inp;
      width: 100%;
      min-width: 0;
      height: 40px;
      border-radius: 11px;
      border: 1px solid rgba(0, 217, 255, .28);
      background: #061225;
      color: #eef8ff;
      padding: 0 12px;
      font-family: "JetBrains Mono", monospace;
      font-size: 14px;
      font-weight: 800;
      outline: none;
      -moz-appearance: textfield;
      transition: border-color .18s ease, box-shadow .18s ease, background .18s ease;
    }
    .assignment-jph-input::-webkit-outer-spin-button,
    .assignment-jph-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
    .assignment-jph-input::placeholder { color: #5d7290; font-weight: 700; }
    .assignment-jph-input:focus {
      border-color: rgba(0, 217, 255, .75);
      background: #07182e;
      box-shadow: 0 0 0 3px rgba(0, 217, 255, .16), 0 0 16px rgba(0, 217, 255, .22);
    }

    .assignment-jph-save {
      grid-area: save;
      height: 40px;
      padding: 0 18px;
      border-radius: 11px;
      border: 1px solid rgba(0, 245, 160, .45);
      background: linear-gradient(180deg, rgba(0, 245, 160, .18), rgba(0, 217, 255, .12));
      color: #d7fff0;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .1em;
      text-transform: uppercase;
      cursor: pointer;
      transition: transform .14s ease, box-shadow .18s ease, background .18s ease, border-color .18s ease;
    }
    .assignment-jph-save:hover {
      transform: translateY(-1px);
      border-color: rgba(0, 245, 160, .85);
      background: linear-gradient(180deg, rgba(0, 245, 160, .28), rgba(0, 217, 255, .18));
      box-shadow: 0 6px 18px rgba(0, 245, 160, .22), 0 0 14px rgba(0, 217, 255, .16);
    }
    .assignment-jph-save:active { transform: translateY(0) scale(.96); }
    .assignment-jph-save:focus-visible { outline: 2px solid rgba(0, 217, 255, .6); outline-offset: 2px; }

    /* custom select arrow + focus */
    .assignment-toolbar select,
    .assignment-row select {
      appearance: none;
      -webkit-appearance: none;
      background-image:
        linear-gradient(45deg, transparent 50%, #67e8f9 50%),
        linear-gradient(135deg, #67e8f9 50%, transparent 50%);
      background-position: calc(100% - 18px) 19px, calc(100% - 13px) 19px;
      background-size: 5px 5px, 5px 5px;
      background-repeat: no-repeat;
      padding-right: 34px;
      transition: border-color .18s ease, box-shadow .18s ease;
    }
    .assignment-toolbar select:focus,
    .assignment-row select:focus {
      border-color: rgba(0, 217, 255, .7);
      box-shadow: 0 0 0 3px rgba(0, 217, 255, .14);
    }

    /* refresh = primary action (not danger) */
    #configAssignedRefreshBtn {
      border-color: rgba(0, 217, 255, .45);
      background: linear-gradient(180deg, rgba(0, 217, 255, .16), rgba(0, 245, 160, .1));
      color: #aef3ff;
      transition: transform .14s ease, box-shadow .18s ease, border-color .18s ease;
    }
    #configAssignedRefreshBtn:hover {
      transform: translateY(-1px);
      border-color: rgba(0, 217, 255, .85);
      box-shadow: 0 6px 18px rgba(0, 217, 255, .22);
    }
    #configAssignedRefreshBtn:active { transform: translateY(0) scale(.97); }

    @media (prefers-reduced-motion: reduce) {
      .assignment-row,
      .assignment-summary-grid article { animation: none !important; }
      .assignment-row,
      .assignment-summary-grid article,
      .assignment-jph-save,
      #configAssignedRefreshBtn { transition: none !important; }
    }

    @media (max-width: 1100px) {
      .assignment-summary-grid,
      .assignment-row {
        grid-template-columns: 1fr;
      }

      .assignment-target {
        justify-items: start;
      }
    }
  `;

  document.head.appendChild(style);
}


/* ===============================
   TABS
================================ */

function initTabs() {
  const buttons = document.querySelectorAll(".tab-btn");
  const contents = document.querySelectorAll(".tab-content");

  buttons.forEach(button => {
    button.addEventListener("click", () => {
      const tab = button.dataset.tab;

      buttons.forEach(btn => btn.classList.remove("active"));
      contents.forEach(content => content.classList.remove("active"));

      button.classList.add("active");

      const target = document.querySelector(`[data-content="${tab}"]`);
      if (target) target.classList.add("active");

      if (tab === "personal" && typeof renderFinishThreeLineCell === "function") {
        renderFinishThreeLineCell();
    renderMorningSetupRosterSummary();
      }
    });
  });
}

/* ===============================
   UI EFFECTS
================================ */

function initClock() {
  updateClock();
  setInterval(updateClock, 1000);
}

function updateClock() {
  setText("clock", new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  }));
}

function updateSyncAge() {
  if (!lastSyncTime) {
    setText("syncAge", "--");
    return;
  }

  const seconds = Math.floor((Date.now() - lastSyncTime) / 1000);

  if (seconds < 60) {
    setText("syncAge", `${seconds}s ago`);
  } else {
    setText("syncAge", `${Math.floor(seconds / 60)}m ago`);
  }
}

function initRefreshButton() {
  document.getElementById("refreshBtn")?.addEventListener("click", () => {
    loadDashboard({ showLoader: false });
  });
}

function initHoverEffects() {
  document.querySelectorAll(".station-card, .machine-card, .process-card, .rail-card, .kpi-card").forEach(card => {
    card.addEventListener("mousemove", event => {
      const rect = card.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;

      card.style.setProperty("--mx", `${x}px`);
      card.style.setProperty("--my", `${y}px`);
    });
  });
}

let loadingInterval = null;
let loadingProgress = 0;

function startLoadingAnimation() {
  stopLoadingAnimation();

  const statuses = [
    "Connecting to Finish dashboard API...",
    "Reading live station signals...",
    "Calculating WIP and throughput metrics...",
    "Building hourly production performance...",
    "Rendering command center interface..."
  ];

  const stepIds = ["loadStep1", "loadStep2", "loadStep3", "loadStep4"];
  let statusIndex = 0;
  loadingProgress = 0;

  updateLoadingVisuals(loadingProgress, statuses[0], 0, stepIds);

  loadingInterval = setInterval(() => {
    if (loadingProgress < 92) {
      loadingProgress += Math.floor(Math.random() * 8) + 4;
      if (loadingProgress > 92) loadingProgress = 92;
    }

    const stepIndex =
      loadingProgress < 25 ? 0 :
      loadingProgress < 50 ? 1 :
      loadingProgress < 75 ? 2 : 3;

    statusIndex = stepIndex;

    updateLoadingVisuals(
      loadingProgress,
      statuses[statusIndex] || statuses[statuses.length - 1],
      stepIndex,
      stepIds
    );
  }, 350);
}

function finishLoadingAnimation() {
  updateLoadingVisuals(
    100,
    "Dashboard ready.",
    3,
    ["loadStep1", "loadStep2", "loadStep3", "loadStep4"]
  );

  stopLoadingAnimation();
}

function stopLoadingAnimation() {
  if (loadingInterval) {
    clearInterval(loadingInterval);
    loadingInterval = null;
  }
}

function updateLoadingVisuals(
  percent,
  text,
  activeStep,
  stepIds = ["loadStep1", "loadStep2", "loadStep3", "loadStep4"]
) {
  const fill = document.getElementById("loadingProgressFill");
  const percentEl = document.getElementById("loadingPercent");
  const statusText = document.getElementById("loadingStatusText");
  const footerStatus = document.getElementById("loadingFooterStatus");

  if (fill) fill.style.width = `${percent}%`;
  if (percentEl) percentEl.textContent = `${percent}%`;
  if (statusText) statusText.textContent = text;
  if (footerStatus) footerStatus.textContent = text;

  stepIds.forEach((id, index) => {
    const el = document.getElementById(id);
    if (!el) return;

    el.classList.remove("active", "done");

    if (index < activeStep) {
      el.classList.add("done");
    } else if (index === activeStep) {
      el.classList.add("active");
    }
  });
}

function showLoading(show) {
  const loader = document.getElementById("loadingScreen");
  if (!loader) return;

  if (show) {
    loader.classList.remove("hidden");
    startLoadingAnimation();
    return;
  }

  finishLoadingAnimation();

  setTimeout(() => {
    loader.classList.add("hidden");
  }, 350);
}

function renderError(error) {
  console.error(error);

  setText("syncAge", "error");

  const container = document.getElementById("statusPanelContainer");
  if (container) {
    container.innerHTML = `
      <article class="status-card status-critical">
        <div class="status-card-top">
          <h3>Dashboard API Error</h3>
          <span>Critical</span>
        </div>
        <p>${escapeHtml(error.message || "Unable to load Finish Dashboard data.")}</p>
      </article>
    `;
  }
}

function flashButton(id, label) {
  const btn = document.getElementById(id);
  if (!btn) return;

  const oldText = btn.textContent;
  btn.textContent = label;
  btn.classList.add("saved");

  setTimeout(() => {
    btn.textContent = oldText;
    btn.classList.remove("saved");
  }, 1500);
}

/* ===============================
   HELPERS
================================ */

function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;

  if (typeof value === "number") {
    el.textContent = formatNumber(value);
  } else {
    el.textContent = value ?? "0";
  }
}

function getInputValue(id) {
  const el = document.getElementById(id);
  return el ? Number(el.value || 0) : 0;
}

function setInputValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function formatTime(date) {
  return new Date(date).toLocaleString([], {
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/* =====================================================
   FINISH OPERATOR COMMAND UI
   Live API: ?action=operatorActivity&area=Finish
===================================================== */

const FINISH_OPERATOR_STATION_ORDER = [
  "Finish Unbox",
  "MEI Line B",
  "MEI Line C",
  "MEI Easy Fit",
  "Mounting",
  "Drill",
  "Bigs",
  "Sharps",
  "Final Inspection"
];

const FINISH_OPERATOR_STATION_META = {
  "Finish Unbox": { zone: "MEI Feed", accent: "green", icon: "▣" },
  "MEI Line B": { zone: "Edging Bank", accent: "cyan", icon: "B" },
  "MEI Line C": { zone: "Edging Bank", accent: "purple", icon: "C" },
  "MEI Easy Fit": { zone: "Edging Bank", accent: "blue", icon: "EF" },
  "Mounting": { zone: "Mount / Assemble", accent: "orange", icon: "M" },
  "Drill": { zone: "Specialty", accent: "amber", icon: "D" },
  "Bigs": { zone: "Side Station", accent: "purple", icon: "BG" },
  "Sharps": { zone: "Side Station", accent: "purple", icon: "SH" },
  "Final Inspection": { zone: "Final QA", accent: "cyan", icon: "FI" }
};

function isFinishOperatorStationAllowed(stationName) {
  const name = normalizeFinishOperatorStation(stationName);

  return ![
    "FSV Scan & Verify",
    "FSV / Frame Only",
    "Frame Only",
    "Frame Only Scan & Verify",
    "FSV Scan & Verify / Frame Only",
    "AR-OUT"
  ].includes(name);
}

let finishOperatorState = {
  rows: [],
  stations: {},
  selectedStation: null,
  selectedOperator: "all",
  uiStatus: {}
};

function initOperatorCommandUI() {
  document.getElementById("operatorRefreshBtn")?.addEventListener("click", loadFinishOperatorActivity);
  hideOperatorApiConnectedPill();
  ensureExpandedOperatorViewer();

  document.querySelectorAll("[data-close-operator-drawer]").forEach(el => {
    el.addEventListener("click", closeOperatorDrawer);
  });

  document.querySelectorAll("[data-status-action]").forEach(btn => {
    btn.addEventListener("click", () => {
      const station = finishOperatorState.selectedStation;
      if (!station) return;

      const action = btn.dataset.statusAction || "clear";
      if (action === "clear") {
        delete finishOperatorState.uiStatus[station];
      } else {
        finishOperatorState.uiStatus[station] = action;
      }

      renderOperatorStationCards(finishOperatorState.stations);
      openOperatorDrawer(station);
    });
  });

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeOperatorDrawer();
  });
}


function hideOperatorApiConnectedPill() {
  const apiStatus = document.getElementById("operatorApiStatus");
  if (!apiStatus) return;

  const apiPill = apiStatus.closest(".operator-api-pill, .hud-pill, .operator-status-pill");
  if (apiPill) {
    apiPill.style.display = "none";
  } else {
    apiStatus.style.display = "none";
  }
}

function getFinishOperatorBaseRate(stationName, operatorName = "") {
  const config = loadConfig();
  const name = normalizeFinishOperatorStation(stationName);
  const assignment = getOperatorAssignment(operatorName, name, config);

  if (assignment && assignment.role === "training") {
    const trainingRate = getTrainingRateForStationWeek(name, assignment.trainingWeek);
    if (trainingRate > 0) return trainingRate;
  }

  switch (name) {
    case "Finish Unbox":
      return Number(config.unboxRate || 0);
    case "MEI Line B":
      return Number(config.meiBRate || 0);
    case "MEI Line C":
      return Number(config.meiCRate || 0);
    case "MEI Easy Fit":
      return Number(config.meiBRate || 0);
    case "Mounting":
      return Number(config.mountRate || 0);
    case "Drill":
      return Number(config.drillRate || 0);
    case "Final Inspection":
      return Number(config.finalRate || 0);
    case "Bigs":
    case "Sharps":
      return 0;
    default:
      return 0;
  }
}

function getFinishOperatorStationTarget(stationName) {
  const station = applyConfigToStation({ flowStep: normalizeFinishOperatorStation(stationName) }, loadConfig());
  return Number(station.expectedNormalPerHour || 0);
}

function getOperatorHourTarget(stationName, hour, mode = "operator", operatorName = "") {
  const baseRate = mode === "station"
    ? getFinishOperatorStationTarget(stationName)
    : getFinishOperatorBaseRate(stationName, operatorName);

  if (baseRate <= 0) return 0;

  const productiveMinutes = getProductiveMinutesForHour(hour);
  if (productiveMinutes <= 0) return 0;

  return Math.round(baseRate * (productiveMinutes / 60));
}

function getOperatorPerformanceStatus(count, target) {
  const safeTarget = Number(target || 0);
  if (safeTarget <= 0) {
    return { pct: 0, className: "neutral", color: "#38bdf8", label: "No Target" };
  }

  const pct = Math.round((Number(count || 0) / safeTarget) * 100);

  if (pct >= 100) {
    return { pct, className: "green", color: "#00f5a0", label: `${pct}%` };
  }

  if (pct >= 90) {
    return { pct, className: "amber", color: "#fbbf24", label: `${pct}%` };
  }

  return { pct, className: "red", color: "#fb7185", label: `${pct}%` };
}

function getLatestOperatorPerformance(operator, stationName) {
  const hours = operator?.hourly || {};
  const orderedHours = getOperatorHourOrder(hours);
  let latestHour = "";
  let latestValue = 0;
  let latestTarget = 0;

  orderedHours.forEach(hour => {
    const target = getOperatorHourTarget(stationName, hour, "operator", operator?.name || "");
    const value = Number(hours[hour] || 0);

    if (target > 0 && value > 0) {
      latestHour = hour;
      latestValue = value;
      latestTarget = target;
    }
  });

  if (!latestHour) {
    const now = new Date();
    const currentHour = now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }).replace(/:\d{2}/, ":00");
    latestHour = orderedHours.includes(currentHour) ? currentHour : orderedHours[orderedHours.length - 1] || "";
    latestValue = Number(hours[latestHour] || 0);
    latestTarget = getOperatorHourTarget(stationName, latestHour, "operator", operator?.name || "");
  }

  const status = getOperatorPerformanceStatus(latestValue, latestTarget);
  return { hour: latestHour, value: latestValue, target: latestTarget, ...status };
}

function renderOperatorFilter(station, operators) {
  const target = document.getElementById("drawerOperatorList");
  if (!target) return;

  const existing = document.getElementById("drawerOperatorFilterWrap");
  if (existing) existing.remove();

  const wrap = document.createElement("div");
  wrap.id = "drawerOperatorFilterWrap";
  wrap.className = "operator-filter-wrap";
  wrap.style.cssText = "margin:0 0 16px;padding:14px 16px;border:1px solid rgba(100,221,255,.18);border-radius:16px;background:rgba(3,12,24,.72);display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;";

  const options = operators.map(operator =>
    `<option value="${escapeHtml(operator.name)}" ${finishOperatorState.selectedOperator === operator.name ? "selected" : ""}>${escapeHtml(operator.name)}</option>`
  ).join("");

  wrap.innerHTML = `
    <div>
      <strong style="display:block;font-size:13px;letter-spacing:.08em;text-transform:uppercase;color:#eef8ff;">Operator Filter</strong>
      <span style="display:block;margin-top:4px;font-size:11px;color:#8fa6c3;">Select one operator or view the full station crew.</span>
    </div>
    <select id="drawerOperatorFilter" style="min-width:260px;height:42px;border-radius:12px;border:1px solid rgba(0,217,255,.35);background:#061225;color:#eef8ff;padding:0 12px;font-weight:800;outline:none;">
      <option value="all" ${finishOperatorState.selectedOperator === "all" ? "selected" : ""}>All operators</option>
      ${options}
    </select>
  `;

  target.before(wrap);

  document.getElementById("drawerOperatorFilter")?.addEventListener("change", event => {
    finishOperatorState.selectedOperator = event.target.value || "all";
    renderDrawerOperators(station);
  });
}

async function loadFinishOperatorActivity() {
  const apiStatus = document.getElementById("operatorApiStatus");
  if (apiStatus) apiStatus.textContent = "Loading";

  try {
    const response = await fetch(FINISH_OPERATOR_API_URL, {
      method: "GET",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`Operator API failed with status ${response.status}`);
    }

    const payload = await response.json();

    if (payload.status === "error" || payload.success === false) {
      throw new Error(payload.message || "Operator API returned an error.");
    }

    const rows = (Array.isArray(payload.operatorActivity)
      ? payload.operatorActivity
      : []
    ).filter(row => isFinishOperatorStationAllowed(
      row.FlowStation || row.flowStation || row.station || row.AccessPoint
    ));

    finishOperatorState.rows = rows;
    finishOperatorState.stations = buildFinishOperatorStations(rows);
    syncFinishPersonalRosterFromStations(finishOperatorState.stations);

    renderOperatorSummary(payload.summary || buildOperatorSummaryFromStations(finishOperatorState.stations));
    renderOperatorStationCards(finishOperatorState.stations);
    renderOperatorAssignmentConfigPanel();
    renderFinishThreeLineCell();

    if (apiStatus) apiStatus.textContent = "Connected";
  } catch (error) {
    console.error("Finish Operator Activity Error:", error);
    if (apiStatus) apiStatus.textContent = "API Error";
    renderOperatorError(error);
  }
}

function buildFinishOperatorStations(rows) {
  const stations = {};

  rows.forEach(row => {
    const rawStation = row.FlowStation || row.flowStation || row.station || row.AccessPoint || "Unmapped";
    const stationName = normalizeFinishOperatorStation(rawStation);

    if (!stations[stationName]) {
      stations[stationName] = {
        name: stationName,
        total: 0,
        operators: {},
        hourly: {},
        accessPoints: new Set()
      };
    }

    const station = stations[stationName];
    const operator = String(row.Operator || row.operator || "Unassigned / No Operator").trim();
    const total = Number(row.Total ?? row.total ?? row.HourlyTotal ?? row.hourlyTotal ?? 0) || 0;
    const hours = row.Hours || row.hours || {};
    const accessPoint = String(row.AccessPoint || row.accessPoint || "").trim();

    station.total += total;
    if (accessPoint) station.accessPoints.add(accessPoint);

    if (!station.operators[operator]) {
      station.operators[operator] = {
        name: operator,
        total: 0,
        hourly: {},
        accessPoints: new Set()
      };
    }

    station.operators[operator].total += total;
    if (accessPoint) station.operators[operator].accessPoints.add(accessPoint);

    getOperatorHourOrder(hours).forEach(hour => {
      const value = Number(hours[hour] || 0) || 0;
      station.hourly[hour] = (station.hourly[hour] || 0) + value;
      station.operators[operator].hourly[hour] = (station.operators[operator].hourly[hour] || 0) + value;
    });
  });

  Object.values(stations).forEach(station => {
    station.accessPoints = Array.from(station.accessPoints);
    station.operatorList = Object.values(station.operators)
      .map(operator => ({
        ...operator,
        accessPoints: Array.from(operator.accessPoints)
      }))
      .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));
  });

  return stations;
}

function normalizeFinishOperatorStation(value) {
  const text = String(value || "").trim();
  const key = text.toLowerCase();

  if (key.includes("unbox")) return "Finish Unbox";
  if (key.includes("line b")) return "MEI Line B";
  if (key.includes("line c")) return "MEI Line C";
  if (key.includes("easy")) return "MEI Easy Fit";
  if (key.includes("mount")) return "Mounting";
  if (key.includes("drill")) return "Drill";
  if (key.includes("big")) return "Bigs";
  if (key.includes("sharp")) return "Sharps";
  if (key.includes("final")) return "Final Inspection";
  if (key.includes("fsv") || key.includes("frame only")) return "FSV Scan & Verify";
  if (key.includes("ar-out") || key.includes("ar out")) return "AR-OUT";

  return text || "Unmapped";
}

function renderOperatorSummary(summary) {
  setText("operatorTotalJobs", summary.totalJobs || 0);
  setText("operatorTotalOperators", summary.totalOperators || 0);
  setText("operatorTopStationTotal", summary.topStationTotal || 0);
  setText("operatorTopStation", summary.topStation || "--");
  setText("operatorPeakHourTotal", summary.peakHourTotal || 0);
  setText("operatorPeakHour", summary.peakHour || "--");
}

function buildOperatorSummaryFromStations(stations) {
  const allStations = Object.values(stations || {});
  const operatorSet = {};
  let totalJobs = 0;
  let topStation = "";
  let topStationTotal = 0;
  let peakHour = "";
  let peakHourTotal = 0;
  const hourTotals = {};

  allStations.forEach(station => {
    totalJobs += Number(station.total || 0);

    if (Number(station.total || 0) > topStationTotal) {
      topStationTotal = Number(station.total || 0);
      topStation = station.name;
    }

    (station.operatorList || []).forEach(operator => {
      operatorSet[operator.name] = true;
    });

    Object.keys(station.hourly || {}).forEach(hour => {
      hourTotals[hour] = (hourTotals[hour] || 0) + Number(station.hourly[hour] || 0);
    });
  });

  Object.keys(hourTotals).forEach(hour => {
    if (hourTotals[hour] > peakHourTotal) {
      peakHourTotal = hourTotals[hour];
      peakHour = hour;
    }
  });

  return {
    totalJobs,
    totalOperators: Object.keys(operatorSet).length,
    topStation,
    topStationTotal,
    peakHour,
    peakHourTotal
  };
}

function renderOperatorStationCards(stations) {
  const grid = document.getElementById("operatorStationGrid");
  if (!grid) return;

  const ordered = getOrderedFinishOperatorStations(stations);

  if (!ordered.length) {
    grid.innerHTML = `<article class="operator-empty-card">No Finish operator rows returned yet. Confirm RAW_ACTIVITY_CURRENT has Area = Finish rows.</article>`;
    return;
  }

  grid.innerHTML = ordered.map(station => {
    const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish", accent: "cyan", icon: "•" };
    const topOperator = (station.operatorList || [])[0];
    const peak = getPeakHour(station.hourly);
    const uiStatus = finishOperatorState.uiStatus[station.name] || "online";

    return `
      <article class="operator-station-card ${escapeHtml(meta.accent)} ui-${escapeHtml(uiStatus)}" data-operator-station="${escapeHtml(station.name)}">
        <div class="operator-card-top">
          <div class="operator-station-icon">${escapeHtml(meta.icon)}</div>
          <div>
            <span>${escapeHtml(meta.zone)}</span>
            <h3>${escapeHtml(station.name)}</h3>
          </div>
          <strong class="operator-ui-status">${escapeHtml(formatUiStatus(uiStatus))}</strong>
        </div>

        <div class="operator-card-main-metric">
          <span>Output Today</span>
          <strong>${numberFmt(station.total)}</strong>
        </div>

        <div class="operator-card-metrics">
          <div><span>Operators</span><strong>${numberFmt((station.operatorList || []).length)}</strong></div>
          <div><span>Peak</span><strong>${escapeHtml(peak.hour || "--")}</strong></div>
          <div><span>Peak CNT</span><strong>${numberFmt(peak.value)}</strong></div>
        </div>

        <div class="operator-card-footer">
          <span>Top: ${escapeHtml(topOperator?.name || "--")}</span>
          <button type="button">Open Detail</button>
        </div>
      </article>`;
  }).join("");

  grid.querySelectorAll("[data-operator-station]").forEach(card => {
    card.addEventListener("click", () => openOperatorDrawer(card.dataset.operatorStation));
  });
}

function getOrderedFinishOperatorStations(stations) {
  const stationMap = stations || {};
  const ordered = [];

  FINISH_OPERATOR_STATION_ORDER.forEach(name => {
    if (stationMap[name]) ordered.push(stationMap[name]);
  });

  Object.keys(stationMap)
    .filter(name =>
      !FINISH_OPERATOR_STATION_ORDER.includes(name) &&
      isFinishOperatorStationAllowed(name)
    )
    .sort()
    .forEach(name => ordered.push(stationMap[name]));

  return ordered;
}

function openOperatorDrawer(stationName) {
  const station = finishOperatorState.stations?.[stationName];
  const drawer = document.getElementById("operatorDrawer");
  if (!station || !drawer) return;

  finishOperatorState.selectedStation = stationName;

  const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish" };
  const peak = getPeakHour(station.hourly);

  setText("operatorDrawerTitle", station.name);
  setText("operatorDrawerSub", `${meta.zone} • ${station.accessPoints?.join(" + ") || "Mapped station"}`);
  setText("drawerStationTotal", station.total || 0);
  setText("drawerOperatorCount", (station.operatorList || []).length);
  setText("drawerPeakHour", peak.hour ? `${peak.hour} (${peak.value})` : "--");

  if (!finishOperatorState.selectedOperator) finishOperatorState.selectedOperator = "all";
  const operatorNames = (station.operatorList || []).map(operator => operator.name);
  if (finishOperatorState.selectedOperator !== "all" && !operatorNames.includes(finishOperatorState.selectedOperator)) {
    finishOperatorState.selectedOperator = "all";
  }

  renderDrawerStationHourly(station);
  renderDrawerOperators(station);
  updateDrawerStatusButtons(finishOperatorState.uiStatus[station.name] || "online");

  drawer.classList.add("active");
  drawer.setAttribute("aria-hidden", "false");
  document.body.classList.add("operator-drawer-open");
}

function closeOperatorDrawer() {
  const drawer = document.getElementById("operatorDrawer");
  if (!drawer) return;

  drawer.classList.remove("active");
  drawer.setAttribute("aria-hidden", "true");
  document.body.classList.remove("operator-drawer-open");
}

function renderDrawerStationHourly(station) {
  /*
    Station hourly output is intentionally hidden in the Operator Command drawer.
    That full station-by-hour view already exists in Hourly Production Performance.
    The drawer stays focused on individual operator hourly output only.
  */
  const target = document.getElementById("drawerStationHourly");
  if (!target) return;

  const block = target.closest(".operator-drawer-block");
  if (block) block.style.display = "none";

  target.innerHTML = "";
}

function renderDrawerOperators(station) {
  const target = document.getElementById("drawerOperatorList");
  if (!target) return;

  const operators = station.operatorList || [];

  if (!operators.length) {
    const existing = document.getElementById("drawerOperatorFilterWrap");
    if (existing) existing.remove();
    target.innerHTML = `<article class="operator-empty-card">No individual operators returned for this station.</article>`;
    return;
  }

  renderOperatorFilter(station, operators);

  const filteredOperators = finishOperatorState.selectedOperator === "all"
    ? operators
    : operators.filter(operator => operator.name === finishOperatorState.selectedOperator);

  target.innerHTML = filteredOperators.map(operator => {
    const hours = operator.hourly || {};
    const currentPerf = getLatestOperatorPerformance(operator, station.name);

    return `
      <article class="operator-person-card perf-${escapeHtml(currentPerf.className)} operator-person-clickable" data-expanded-operator="${escapeHtml(operator.name)}" style="border-color:${currentPerf.color};box-shadow:0 0 18px ${currentPerf.color}1f;cursor:pointer;">
        <div class="operator-person-head">
          <div>
            <h4>${escapeHtml(operator.name)}</h4>
            <span>${escapeHtml(operator.accessPoints?.join(" + ") || station.name)} · ${escapeHtml(getAssignmentTargetLabel(operator.name, station.name))}</span>
            <small style="display:inline-flex;margin-top:8px;padding:5px 9px;border-radius:999px;border:1px solid ${currentPerf.color};color:${currentPerf.color};font-weight:900;letter-spacing:.06em;text-transform:uppercase;">
              Current JPH ${escapeHtml(currentPerf.label)} · ${escapeHtml(shortHour(currentPerf.hour)) || "--"}: ${numberFmt(currentPerf.value)} / ${numberFmt(currentPerf.target)}
            </small>
          </div>
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:8px;">
            <strong style="color:${currentPerf.color};">${numberFmt(operator.total)}</strong>
            <button type="button" class="operator-expand-btn" data-expanded-operator-btn="${escapeHtml(operator.name)}">Expand</button>
          </div>
        </div>

        <div class="operator-person-hours">
          ${getOperatorHourOrder(hours).map(hour => {
            const value = Number(hours[hour] || 0);
            const targetCount = getOperatorHourTarget(station.name, hour, "operator", operator.name);
            const status = getOperatorPerformanceStatus(value, targetCount);
            const pct = targetCount > 0
              ? Math.min(100, Math.round((value / targetCount) * 100))
              : 0;

            return `
              <div class="operator-mini-hour perf-${escapeHtml(status.className)}" title="${escapeHtml(operator.name)} · ${escapeHtml(hour)} · ${escapeHtml(status.label)}">
                <span>${escapeHtml(shortHour(hour))}</span>
                <b style="color:${status.color};">${numberFmt(value)}</b>
                <small style="color:${status.color};font-size:10px;font-weight:900;">${escapeHtml(status.label)}</small>
                <i style="height:${Math.max(4, pct)}%;background:${status.color};box-shadow:0 0 14px ${status.color};"></i>
              </div>`;
          }).join("")}
        </div>
      </article>`;
  }).join("");

  target.querySelectorAll("[data-expanded-operator]").forEach(card => {
    card.addEventListener("click", event => {
      event.stopPropagation();
      const operatorName = card.dataset.expandedOperator;
      openExpandedOperatorViewer(station.name, operatorName);
    });
  });

  target.querySelectorAll("[data-expanded-operator-btn]").forEach(btn => {
    btn.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      openExpandedOperatorViewer(station.name, btn.dataset.expandedOperatorBtn);
    });
  });
}

function ensureExpandedOperatorViewer() {
  if (document.getElementById("expandedOperatorViewer")) return;

  const style = document.createElement("style");
  style.id = "expandedOperatorViewerStyles";
  style.textContent = `
    .operator-expand-btn {
      height: 30px;
      padding: 0 12px;
      border-radius: 999px;
      border: 1px solid rgba(0, 217, 255, 0.42);
      background: rgba(0, 217, 255, 0.08);
      color: #67e8f9;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
      letter-spacing: .08em;
      text-transform: uppercase;
      cursor: pointer;
    }

    .operator-person-clickable {
      transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
    }

    .operator-person-clickable:hover {
      transform: translateY(-2px);
      box-shadow: 0 0 28px rgba(0, 217, 255, .18), inset 0 1px 0 rgba(255,255,255,.08) !important;
    }

    .expanded-operator-backdrop {
      position: fixed;
      inset: 0;
      z-index: 9998;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 26px;
      background: rgba(0, 4, 10, .72);
      backdrop-filter: blur(12px);
    }

    .expanded-operator-backdrop.active {
      display: flex;
    }

    .expanded-operator-shell {
      width: min(1220px, 96vw);
      max-height: 92vh;
      overflow: hidden;
      border: 1px solid rgba(100, 221, 255, .24);
      border-radius: 26px;
      background:
        radial-gradient(circle at 8% 0%, rgba(0, 217, 255, .14), transparent 32%),
        radial-gradient(circle at 92% 8%, rgba(0, 245, 160, .10), transparent 34%),
        linear-gradient(180deg, rgba(10, 28, 52, .98), rgba(2, 7, 17, .99));
      box-shadow: 0 34px 90px rgba(0,0,0,.58), 0 0 38px rgba(0,217,255,.14);
    }

    .expanded-operator-header {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 18px;
      align-items: start;
      padding: 24px 26px 20px;
      border-bottom: 1px solid rgba(100, 221, 255, .18);
      background: linear-gradient(90deg, rgba(0, 217, 255, .08), transparent);
    }

    .expanded-operator-kicker {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: #5eead4;
      font-family: "JetBrains Mono", monospace;
      font-size: 11px;
      font-weight: 900;
      letter-spacing: .16em;
      text-transform: uppercase;
    }

    .expanded-operator-kicker::before {
      content: "";
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #00f5a0;
      box-shadow: 0 0 14px #00f5a0;
    }

    .expanded-operator-title {
      margin: 10px 0 0;
      font-size: clamp(30px, 4vw, 54px);
      line-height: .95;
      font-weight: 900;
      letter-spacing: -.04em;
    }

    .expanded-operator-sub {
      margin-top: 8px;
      color: #9fb3cb;
      font-size: 14px;
      font-weight: 700;
    }

    .expanded-operator-close {
      width: 46px;
      height: 46px;
      display: grid;
      place-items: center;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .22);
      background: rgba(3, 12, 24, .8);
      color: #eef8ff;
      font-size: 24px;
      cursor: pointer;
    }

    .expanded-operator-body {
      max-height: calc(92vh - 132px);
      overflow: auto;
      padding: 22px 26px 26px;
    }

    .expanded-operator-scoreboard {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }

    .expanded-score-card {
      min-height: 112px;
      padding: 18px;
      border-radius: 20px;
      border: 1px solid rgba(100, 221, 255, .18);
      background: rgba(3, 12, 24, .72);
      box-shadow: inset 0 1px 0 rgba(255,255,255,.06);
    }

    .expanded-score-card span {
      display: block;
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
      letter-spacing: .14em;
      text-transform: uppercase;
    }

    .expanded-score-card strong {
      display: block;
      margin-top: 10px;
      color: #eef8ff;
      font-size: 34px;
      line-height: 1;
      font-weight: 900;
    }

    .expanded-score-card small {
      display: block;
      margin-top: 8px;
      color: #9fb3cb;
      font-weight: 800;
    }

    .expanded-operator-timeline {
      display: grid;
      grid-template-columns: repeat(15, minmax(64px, 1fr));
      gap: 10px;
      align-items: end;
      min-height: 360px;
      padding: 22px;
      border-radius: 24px;
      border: 1px solid rgba(100, 221, 255, .18);
      background:
        linear-gradient(180deg, rgba(255,255,255,.045), transparent),
        rgba(3, 12, 24, .72);
      overflow-x: auto;
    }

    .expanded-hour-tower {
      min-width: 64px;
      height: 300px;
      display: grid;
      grid-template-rows: auto 1fr auto;
      gap: 8px;
      text-align: center;
    }

    .expanded-hour-top span,
    .expanded-hour-foot span {
      display: block;
      color: #8fa6c3;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 900;
    }

    .expanded-hour-top strong {
      display: block;
      margin-top: 5px;
      font-size: 22px;
      font-weight: 900;
    }

    .expanded-hour-bar-wrap {
      position: relative;
      align-self: stretch;
      min-height: 190px;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .10);
      background: rgba(255,255,255,.035);
      overflow: hidden;
    }

    .expanded-hour-target-line {
      position: absolute;
      left: 8px;
      right: 8px;
      bottom: var(--target-line, 70%);
      height: 2px;
      background: rgba(255,255,255,.72);
      box-shadow: 0 0 12px rgba(255,255,255,.38);
      z-index: 2;
    }

    .expanded-hour-fill {
      position: absolute;
      left: 9px;
      right: 9px;
      bottom: 8px;
      height: var(--bar-height, 4%);
      min-height: 6px;
      border-radius: 12px 12px 4px 4px;
      background: var(--bar-color, #38bdf8);
      box-shadow: 0 0 18px var(--bar-color, #38bdf8);
    }

    .expanded-hour-foot strong {
      display: block;
      margin-top: 4px;
      font-size: 12px;
      font-weight: 900;
    }

    .expanded-operator-note {
      margin-top: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(100, 221, 255, .14);
      background: rgba(255,255,255,.035);
      color: #bfdaf2;
      font-size: 13px;
      line-height: 1.5;
    }

    @media (max-width: 900px) {
      .expanded-operator-scoreboard { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .expanded-operator-header { padding: 20px; }
      .expanded-operator-body { padding: 18px; }
    }
  `;
  document.head.appendChild(style);

  const viewer = document.createElement("section");
  viewer.id = "expandedOperatorViewer";
  viewer.className = "expanded-operator-backdrop";
  viewer.setAttribute("aria-hidden", "true");
  viewer.innerHTML = `
    <div class="expanded-operator-shell" role="dialog" aria-modal="true" aria-label="Expanded operator performance">
      <header class="expanded-operator-header">
        <div>
          <div class="expanded-operator-kicker">Operator Drilldown</div>
          <h2 class="expanded-operator-title" id="expandedOperatorName">--</h2>
          <div class="expanded-operator-sub" id="expandedOperatorSub">--</div>
        </div>
        <button type="button" class="expanded-operator-close" data-close-expanded-operator aria-label="Close expanded operator view">×</button>
      </header>
      <div class="expanded-operator-body" id="expandedOperatorBody"></div>
    </div>
  `;

  document.body.appendChild(viewer);

  viewer.addEventListener("click", event => {
    if (event.target === viewer || event.target.closest("[data-close-expanded-operator]")) {
      closeExpandedOperatorViewer();
    }
  });

  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeExpandedOperatorViewer();
  });
}

function openExpandedOperatorViewer(stationName, operatorName) {
  ensureExpandedOperatorViewer();

  const station = finishOperatorState.stations?.[stationName];
  const operator = (station?.operatorList || []).find(item => item.name === operatorName);
  const viewer = document.getElementById("expandedOperatorViewer");
  const body = document.getElementById("expandedOperatorBody");

  if (!station || !operator || !viewer || !body) return;

  const currentPerf = getLatestOperatorPerformance(operator, station.name);
  const hours = operator.hourly || {};
  const orderedHours = getOperatorHourOrder(hours);
  const meta = FINISH_OPERATOR_STATION_META[station.name] || { zone: "Finish" };
  const best = getPeakHour(hours);
  const activeHours = orderedHours.filter(hour => Number(hours[hour] || 0) > 0).length;
  const avgPerActiveHour = activeHours > 0 ? Math.round(Number(operator.total || 0) / activeHours) : 0;

  setText("expandedOperatorName", operator.name);
  setText("expandedOperatorSub", `${station.name} • ${meta.zone} • ${getAssignmentTargetLabel(operator.name, station.name)} • ${operator.accessPoints?.join(" + ") || "Mapped station"}`);

  body.innerHTML = `
    <section class="expanded-operator-scoreboard">
      <article class="expanded-score-card">
        <span>Total Output</span>
        <strong>${numberFmt(operator.total)}</strong>
        <small>Jobs scanned today</small>
      </article>
      <article class="expanded-score-card">
        <span>Current JPH</span>
        <strong style="color:${currentPerf.color};">${escapeHtml(currentPerf.label)}</strong>
        <small>${escapeHtml(shortHour(currentPerf.hour)) || "--"}: ${numberFmt(currentPerf.value)} / ${numberFmt(currentPerf.target)}</small>
      </article>
      <article class="expanded-score-card">
        <span>Peak Hour</span>
        <strong>${escapeHtml(shortHour(best.hour) || "--")}</strong>
        <small>${numberFmt(best.value)} jobs</small>
      </article>
      <article class="expanded-score-card">
        <span>Active Hour Avg</span>
        <strong>${numberFmt(avgPerActiveHour)}</strong>
        <small>${numberFmt(activeHours)} active hours</small>
      </article>
    </section>

    <section class="expanded-operator-timeline">
      ${orderedHours.map(hour => {
        const value = Number(hours[hour] || 0);
        const target = getOperatorHourTarget(station.name, hour, "operator", operator.name);
        const status = getOperatorPerformanceStatus(value, target);
        const pctRaw = target > 0 ? Math.round((value / target) * 100) : 0;
        const barPct = target > 0 ? Math.min(100, Math.max(4, pctRaw)) : (value > 0 ? 18 : 4);
        const targetLine = target > 0 ? 70 : 4;

        return `
          <article class="expanded-hour-tower perf-${escapeHtml(status.className)}">
            <div class="expanded-hour-top">
              <span>${escapeHtml(shortHour(hour))}</span>
              <strong style="color:${status.color};">${numberFmt(value)}</strong>
            </div>
            <div class="expanded-hour-bar-wrap" style="--target-line:${targetLine}%;">
              <div class="expanded-hour-target-line"></div>
              <div class="expanded-hour-fill" style="--bar-height:${barPct}%;--bar-color:${status.color};"></div>
            </div>
            <div class="expanded-hour-foot">
              <span>${target > 0 ? numberFmt(target) + " target" : "No target"}</span>
              <strong style="color:${status.color};">${escapeHtml(status.label)}</strong>
            </div>
          </article>`;
      }).join("")}
    </section>

    <div class="expanded-operator-note">
      Green means the operator met or beat the configured JPH target. Amber means 90% to 99%. Red means below 90%. The white marker represents the target line for that hour after shift breaks are applied.
    </div>
  `;

  viewer.classList.add("active");
  viewer.setAttribute("aria-hidden", "false");
}

function closeExpandedOperatorViewer() {
  const viewer = document.getElementById("expandedOperatorViewer");
  if (!viewer) return;

  viewer.classList.remove("active");
  viewer.setAttribute("aria-hidden", "true");
}

function updateDrawerStatusButtons(activeStatus) {
  document.querySelectorAll("[data-status-action]").forEach(btn => {
    const action = btn.dataset.statusAction || "clear";
    btn.classList.toggle("active", action === activeStatus || (action === "clear" && !activeStatus));
  });
}

function getPeakHour(hours) {
  let bestHour = "";
  let bestValue = 0;

  Object.keys(hours || {}).forEach(hour => {
    const value = Number(hours[hour] || 0);
    if (value > bestValue) {
      bestValue = value;
      bestHour = hour;
    }
  });

  return { hour: bestHour, value: bestValue };
}

function getOperatorHourOrder(hours) {
  const preferred = [
    "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM", "11:00 AM",
    "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", "5:00 PM",
    "6:00 PM", "7:00 PM", "8:00 PM"
  ];

  const keys = Object.keys(hours || {});
  const ordered = preferred.filter(hour => keys.includes(hour));
  const leftovers = keys.filter(hour => !preferred.includes(hour)).sort();

  return [...ordered, ...leftovers];
}

function shortHour(hour) {
  return String(hour || "")
    .replace(":00", "")
    .replace(" AM", "A")
    .replace(" PM", "P");
}

function formatUiStatus(status) {
  if (status === "down") return "Down";
  if (status === "issue") return "Issue";
  return "Online";
}

function numberFmt(value) {
  return Number(value || 0).toLocaleString();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}



/* =========================================================
   FINISH 3-LINE PHYSICAL CELL
   Line A / B / C with odd side left and even side right.
   Operator output still comes from the live operator API.
   Line + slot placement comes from personal profile assignment.
========================================================= */
const FINISH_LINE_NAMES = ["Line A", "Line B", "Line C"];
const FINISH_LINE_LAYOUT = {
  mountingLeft: ["M09", "M07", "M05", "M03", "M01"],
  mountingRight: ["M10", "M08", "M06", "M04", "M02"],
  finalLeft: ["FI-03", "FI-01"],
  finalRight: ["FI-04", "FI-02"]
};

function initFinishThreeLineCell() {
  installFinishLineSlotDelegatedClickHandler();
  renderFinishThreeLineCell();
}

/*
  Hard click fix for Morning Set Up slots.
  Some visual layers can sit above the station cards depending on screen width.
  This delegated handler detects the slot by coordinates, so M01-M10 and FI-01-FI-04
  open even if a converter/connector layer receives the actual click target.
*/
function installFinishLineSlotDelegatedClickHandler() {
  if (window.__finishLineSlotDelegatedClickInstalled) return;
  window.__finishLineSlotDelegatedClickInstalled = true;

  document.addEventListener("click", event => {
    const target = event.target;

    // Do not steal clicks from line controls, modal controls, tabs, or form fields.
    if (target?.closest?.(
      "button, a, input, select, textarea, [data-line-toggle], [data-line-status], .line-status-buttons, .morning-slot-modal"
    )) {
      return;
    }

    const lineCell = target?.closest?.(".finish-line-cell");
    if (!lineCell) return;

    let slotEl = target.closest?.("[data-line-slot]");

    if (!slotEl) {
      slotEl = getFinishLineSlotAtPoint(event.clientX, event.clientY, lineCell);
    }

    if (!slotEl) return;

    event.preventDefault();
    event.stopPropagation();
    openMorningLineSlotAssignment(
      slotEl.dataset.line,
      slotEl.dataset.station,
      slotEl.dataset.slot,
      slotEl.dataset.operator || ""
    );
  }, true);
}

function getFinishLineSlotAtPoint(clientX, clientY, lineCell) {
  const slots = Array.from(lineCell.querySelectorAll("[data-line-slot]"));

  // First pass: exact rectangle hit.
  const directHit = slots.find(slot => {
    const rect = slot.getBoundingClientRect();
    return (
      clientX >= rect.left &&
      clientX <= rect.right &&
      clientY >= rect.top &&
      clientY <= rect.bottom
    );
  });

  if (directHit) return directHit;

  // Second pass: small tolerance for tight/overlapped visuals.
  const tolerance = 8;
  return slots.find(slot => {
    const rect = slot.getBoundingClientRect();
    return (
      clientX >= rect.left - tolerance &&
      clientX <= rect.right + tolerance &&
      clientY >= rect.top - tolerance &&
      clientY <= rect.bottom + tolerance
    );
  }) || null;
}

function isLineSlotStation(stationName) {
  const station = normalizeFinishOperatorStation(stationName);
  return station === "Mounting" || station === "Final Inspection";
}

function getFinishLineSlotStatusKey() {
  return `finishLineSlotStatus_${getCurrentUsername()}`;
}

function loadFinishLineSlotStatus() {
  try {
    return JSON.parse(localStorage.getItem(getFinishLineSlotStatusKey()) || "{}");
  } catch {
    return {};
  }
}

function saveFinishLineSlotStatus(statusMap) {
  localStorage.setItem(getFinishLineSlotStatusKey(), JSON.stringify(statusMap || {}));
}

function makeLineSlotKey(line, station, slot) {
  return `${line}||${normalizeFinishOperatorStation(station)}||${slot}`;
}

function buildLineOptions(selectedLine = "") {
  return `<option value="">Select line</option>` + FINISH_LINE_NAMES.map(line =>
    `<option value="${escapeHtml(line)}" ${line === selectedLine ? "selected" : ""}>${escapeHtml(line)}</option>`
  ).join("");
}

function getSlotsForLineStation(stationName) {
  const station = normalizeFinishOperatorStation(stationName);
  if (station === "Mounting") {
    return [
      ...FINISH_LINE_LAYOUT.mountingLeft,
      ...FINISH_LINE_LAYOUT.mountingRight
    ].sort((a, b) => Number(a.replace(/\D/g, "")) - Number(b.replace(/\D/g, "")));
  }
  if (station === "Final Inspection") {
    return ["FI-01", "FI-02", "FI-03", "FI-04"];
  }
  if (station === "Drill") {
    return ["D07", "D09"];
  }
  return [];
}

function buildSlotOptions(stationName, selectedSlot = "") {
  const slots = getSlotsForLineStation(stationName);
  return `<option value="">Select slot</option>` + slots.map(slot =>
    `<option value="${escapeHtml(slot)}" ${slot === selectedSlot ? "selected" : ""}>${escapeHtml(slot)}</option>`
  ).join("");
}

function buildLineSlotAssignmentControls(stationName, operatorName, assignment = {}) {
  if (!isLineSlotStation(stationName)) {
    return `
      <label class="assignment-line disabled">
        Line
        <select disabled><option>Not used</option></select>
      </label>
      <label class="assignment-slot disabled">
        Position
        <select disabled><option>Not used</option></select>
      </label>`;
  }

  return `
    <label class="assignment-line">
      Line
      <select data-assignment-line data-station="${escapeHtml(stationName)}" data-operator="${escapeHtml(operatorName)}">
        ${buildLineOptions(assignment.line || "")}
      </select>
    </label>
    <label class="assignment-slot">
      Position
      <select data-assignment-slot data-station="${escapeHtml(stationName)}" data-operator="${escapeHtml(operatorName)}">
        ${buildSlotOptions(stationName, assignment.slot || "")}
      </select>
    </label>`;
}

function getAssignedLineSlots(config = loadConfig(), rosterRows = null) {
  const sourceRows = Array.isArray(rosterRows)
    ? rosterRows
    : (finishRosterApiState.morningRoster || []);

  const assignments = Object.values(loadFinishPersonalOperatorAssignments(sourceRows) || {});
  const map = {};

  assignments.forEach(assignment => {
    const line = assignment.line || "";
    const station = normalizeFinishOperatorStation(assignment.station || "");
    const slot = assignment.slot || "";
    if (!line || !station || !slot) return;
    if (!isLineSlotStation(station)) return;

    const key = makeLineSlotKey(line, station, slot);
    const live = getLiveOperatorForStation(assignment.operator, station);
    map[key] = {
      ...assignment,
      station,
      line,
      slot,
      total: live?.total ?? 0,
      hourly: live?.hourly || {},
      liveNow: !!live
    };
  });

  return map;
}

function getLiveOperatorForStation(operatorName, stationName) {
  const station = finishOperatorState?.stations?.[normalizeFinishOperatorStation(stationName)];
  if (!station) return null;
  const target = normalizeAssignmentOperatorName(operatorName);
  return (station.operatorList || []).find(op => normalizeAssignmentOperatorName(op.name) === target) || null;
}

function renderFinishThreeLineCell() {
  const grid = document.getElementById("finishThreeLineGrid");
  if (!grid) return;

  const config = loadConfig();
  const assignmentMap = getAssignedLineSlots(config, finishRosterApiState.morningRoster || []);
  const statusMap = loadFinishLineSlotStatus();

  renderMorningSetupRosterSummary();
  grid.innerHTML = FINISH_LINE_NAMES.map(line => renderOneFinishLine(line, assignmentMap, statusMap)).join("");
  bindFinishLineCellEvents();
}


function getLineStationCapacityStats(lineAssignments, stationName) {
  const station = normalizeFinishOperatorStation(stationName);
  const assigned = (lineAssignments || []).filter(a => normalizeFinishOperatorStation(a.station) === station);
  const assignedCount = assigned.length;
  const shiftType = getMorningSetupShiftType();
  const shiftHours = getMorningSetupShiftHours();

  let totalJph = 0;

  assigned.forEach(item => {
    const role = String(item.role || "").toLowerCase();
    let rate = Number(item.individualJph || 0) || 0;

    if (!rate && role === "training") {
      rate = Number(getTrainingRateForStationWeek(station, item.trainingWeek || 1)) || 0;
    } else if (!rate && ["core", "certified", "tq"].includes(role)) {
      rate = Number(getFinishOperatorBaseRate(station, item.operator)) || 0;
    }

    totalJph += rate;
  });

  const shiftCapacity = Math.round(totalJph * shiftHours);
  const avgJph = assignedCount ? Math.round(totalJph / assignedCount) : 0;

  return {
    assignedCount,
    shiftType,
    shiftHours,
    totalJph,
    avgJph,
    shiftCapacity,
    className: assignedCount ? "capacity" : "neutral",
    color: assignedCount ? "#00f5a0" : "#38bdf8",
    label: assignedCount ? `${shiftHours}h shift capacity` : "No assigned operators"
  };
}

function renderLineCapacityMetric(label, stats) {
  const mainText = stats.assignedCount
    ? `${numberFmt(stats.shiftCapacity)} jobs`
    : "No Target";

  const smallText = stats.assignedCount
    ? `${numberFmt(stats.totalJph)} JPH · AVG ${numberFmt(stats.avgJph)} · ${numberFmt(stats.assignedCount)} op`
    : "Assign active roster";

  return `
    <div class="line-jph-metric perf-${escapeHtml(stats.className)}" style="--metric-color:${escapeHtml(stats.color)};">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(mainText)}</strong>
      <small>${escapeHtml(smallText)}</small>
    </div>`;
}

function renderMorningSetupRosterSummary() {
  const host = document.getElementById("morningSetupRosterSummary");
  if (!host) return;

  const rows = (finishRosterApiState.morningRoster || [])
    .filter(row => String(row.activeStatus || "").toLowerCase() === "active");

  const shiftType = getMorningSetupShiftType();
  const hours = getMorningSetupShiftHours();
  const mount = rows.filter(row => normalizeFinishOperatorStation(row.defaultArea) === "Mounting").length;
  const final = rows.filter(row => normalizeFinishOperatorStation(row.defaultArea) === "Final Inspection").length;

  host.innerHTML = `
    <div><span>Profile</span><strong>${escapeHtml(getCurrentUsername())}</strong></div>
    <div><span>Shift</span><strong>${escapeHtml(shiftType)}</strong></div>
    <div><span>Hours</span><strong>${hours}</strong></div>
    <div><span>Active</span><strong>${numberFmt(rows.length)}</strong></div>
    <div><span>Mounting</span><strong>${numberFmt(mount)}</strong></div>
    <div><span>Final</span><strong>${numberFmt(final)}</strong></div>`;
}
function renderOneFinishLine(line, assignmentMap, statusMap) {
  const lineAssignments = Object.values(assignmentMap).filter(a => a.line === line);
  const mountingCount = lineAssignments.filter(a => a.station === "Mounting").length;
  const finalCount = lineAssignments.filter(a => a.station === "Final Inspection").length;
  const liveTotal = lineAssignments.reduce((sum, a) => sum + Number(a.total || 0), 0);
  const mountingStats = getLineStationCapacityStats(lineAssignments, "Mounting");
  const finalStats = getLineStationCapacityStats(lineAssignments, "Final Inspection");
  const isCollapsed = finishRosterUiState.collapsedLines.has(line);

  return `
    <article class="finish-line-cell ${isCollapsed ? "is-collapsed" : "is-expanded"}" data-finish-line="${escapeHtml(line)}">
      <header class="finish-line-topbar">
        <div class="finish-line-title">
          <span>Finish Production Cell</span>
          <h3>${escapeHtml(line)}</h3>
        </div>
        <div class="finish-line-metrics">
          <div><span>Mounting</span><strong>${numberFmt(mountingCount)}/10</strong></div>
          <div><span>Final</span><strong>${numberFmt(finalCount)}/4</strong></div>
          ${renderLineCapacityMetric("Mount Cap", mountingStats)}
          ${renderLineCapacityMetric("Final Cap", finalStats)}
          <div><span>Live Output</span><strong>${numberFmt(liveTotal)}</strong></div>
          <button class="finish-line-collapse-btn" type="button" data-line-toggle="${escapeHtml(line)}">
            ${isCollapsed ? "Open" : "Close"}
          </button>
        </div>
      </header>

      <div class="finish-line-body">
        <div class="line-side left">
          <div class="line-side-label"><span>Odd Side</span><span>Left</span></div>
          ${FINISH_LINE_LAYOUT.mountingLeft.map(slot => renderLineSlot(line, "Mounting", slot, assignmentMap, statusMap, "left")).join("")}
          <div class="line-final-divider"></div>
          ${FINISH_LINE_LAYOUT.finalLeft.map(slot => renderLineSlot(line, "Final Inspection", slot, assignmentMap, statusMap, "left")).join("")}
        </div>

        <div class="line-converter-core" aria-label="${escapeHtml(line)} converter">
          <div class="line-converter-arrows">
            <span>↓</span><span>↓</span><span>↓</span><span>↓</span><span>↓</span>
          </div>
        </div>

        <div class="line-side right">
          <div class="line-side-label"><span>Even Side</span><span>Right</span></div>
          ${FINISH_LINE_LAYOUT.mountingRight.map(slot => renderLineSlot(line, "Mounting", slot, assignmentMap, statusMap, "right")).join("")}
          <div class="line-final-divider"></div>
          ${FINISH_LINE_LAYOUT.finalRight.map(slot => renderLineSlot(line, "Final Inspection", slot, assignmentMap, statusMap, "right")).join("")}
        </div>

        <div class="line-status-buttons">
          <button type="button" data-line-status="online" data-line="${escapeHtml(line)}">Online</button>
          <button type="button" data-line-status="issue" data-line="${escapeHtml(line)}">Issue</button>
          <button type="button" data-line-status="down" data-line="${escapeHtml(line)}">Down</button>
          <button type="button" data-line-status="clear" data-line="${escapeHtml(line)}">Clear</button>
        </div>
      </div>
    </article>`;
}

function renderLineSlot(line, station, slot, assignmentMap, statusMap, side) {
  const key = makeLineSlotKey(line, station, slot);
  const assignment = assignmentMap[key];
  const status = statusMap[key] || "online";
  const isFinal = normalizeFinishOperatorStation(station) === "Final Inspection";
  const operator = assignment?.operator || "Click to assign";
  const role = assignment ? getAssignmentRoleLabel(assignment.operator, station, loadConfig()) : "Select from your active roster";
  const total = Number(assignment?.total || 0);
  const liveText = assignment?.liveNow ? "Live" : (assignment ? "Saved" : "Open");
  const direction = side === "left" ? "→" : "←";

  return `
    <article class="finish-slot ${isFinal ? "final" : "mounting"} ${assignment ? "assigned" : "empty"} slot-status-${escapeHtml(status)}" data-line-slot="${escapeHtml(key)}" data-line="${escapeHtml(line)}" data-station="${escapeHtml(station)}" data-slot="${escapeHtml(slot)}" data-operator="${escapeHtml(assignment?.operator || "")}">
      <span class="slot-status-dot"></span>
      <div class="slot-code">${escapeHtml(slot)}</div>
      <div class="slot-copy">
        <strong>${escapeHtml(operator)}</strong>
        <span>${escapeHtml(role)} · ${escapeHtml(liveText)}</span>
      </div>
      <div class="slot-output">
        <strong>${numberFmt(total)}</strong>
        <span>${escapeHtml(direction)}</span>
      </div>
    </article>`;
}

function bindFinishLineCellEvents() {
  document.querySelectorAll("[data-line-toggle]").forEach(button => {
    button.addEventListener("click", event => {
      event.stopPropagation();
      const line = button.dataset.lineToggle;
      if (!line) return;
      if (finishRosterUiState.collapsedLines.has(line)) {
        finishRosterUiState.collapsedLines.delete(line);
      } else {
        finishRosterUiState.collapsedLines.add(line);
      }
      renderFinishThreeLineCell();
    });
  });

  document.querySelectorAll("[data-line-slot]").forEach(slot => {
    slot.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      openMorningLineSlotAssignment(slot.dataset.line, slot.dataset.station, slot.dataset.slot, slot.dataset.operator || "");
    });
  });

  document.querySelectorAll("[data-line-status]").forEach(button => {
    button.addEventListener("click", event => {
      event.stopPropagation();
      const line = button.dataset.line;
      const mode = button.dataset.lineStatus;
      const statusMap = loadFinishLineSlotStatus();
      const slots = document.querySelectorAll(`[data-finish-line="${CSS.escape(line)}"] [data-line-slot]`);
      slots.forEach(slot => {
        const key = slot.dataset.lineSlot;
        if (mode === "clear") delete statusMap[key];
        else statusMap[key] = mode;
      });
      saveFinishLineSlotStatus(statusMap);
      renderFinishThreeLineCell();
    });
  });
}


function getMorningActiveRosterRows() {
  return (finishRosterApiState.morningRoster || [])
    .filter(row => String(row.activeStatus || "").toLowerCase() === "active")
    .filter(row => normalizeAssignmentOperatorName(row.operatorName || ""));
}

function findMorningRosterRow(operatorName) {
  const target = normalizeAssignmentOperatorName(operatorName).toLowerCase();
  return getMorningActiveRosterRows().find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === target
  ) || null;
}

function ensureMorningSlotModal() {
  let modal = document.getElementById("morningSlotAssignModal");
  if (modal) return modal;

  modal = document.createElement("div");
  modal.id = "morningSlotAssignModal";
  modal.className = "morning-slot-modal";
  modal.innerHTML = `
    <div class="morning-slot-backdrop" data-morning-slot-close></div>
    <section class="morning-slot-card" role="dialog" aria-modal="true" aria-label="Morning slot assignment">
      <header>
        <div>
          <span>Morning Set Up</span>
          <h3 id="morningSlotModalTitle">Assign Position</h3>
          <p id="morningSlotModalSub">Select from your active roster.</p>
        </div>
        <button type="button" data-morning-slot-close>×</button>
      </header>
      <div id="morningSlotModalBody" class="morning-slot-body"></div>
    </section>`;

  document.body.appendChild(modal);
  modal.querySelectorAll("[data-morning-slot-close]").forEach(btn => {
    btn.addEventListener("click", closeMorningLineSlotAssignment);
  });
  return modal;
}

function openMorningLineSlotAssignment(line, station, slot, currentOperator = "") {
  const modal = ensureMorningSlotModal();
  const title = modal.querySelector("#morningSlotModalTitle");
  const sub = modal.querySelector("#morningSlotModalSub");
  const body = modal.querySelector("#morningSlotModalBody");
  const rows = getMorningActiveRosterRows();
  const canEdit = canEditFinishOperatorAssignments();

  if (title) title.textContent = `${line} · ${slot}`;
  if (sub) {
    sub.textContent = `${normalizeFinishOperatorStation(station)} · ${getMorningSetupShiftType()} · ${getCurrentUsername()}`;
  }

  if (!rows.length) {
    body.innerHTML = `<article class="morning-slot-empty">No active associates are assigned to your ${escapeHtml(getMorningSetupShiftType())} roster yet. LMS must add them first.</article>`;
  } else {
    body.innerHTML = `
      <div class="morning-slot-current">
        <span>Current</span>
        <strong>${escapeHtml(currentOperator || "Open position")}</strong>
      </div>
      <div class="morning-slot-list">
        ${rows.map(row => {
          const name = normalizeAssignmentOperatorName(row.operatorName || "");
          const isCurrent = name === currentOperator;
          const placement = [row.defaultLine, row.defaultArea, row.defaultPosition].filter(Boolean).join(" / ") || "No position";
          const role = row.roleType === "Training" ? `${row.roleType} ${row.trainingWeek || ""}` : (row.roleType || "Unassigned");
          return `
            <button type="button" class="morning-slot-choice ${isCurrent ? "is-current" : ""}" data-morning-slot-operator="${escapeHtml(name)}" ${canEdit ? "" : "disabled"}>
              <strong>${escapeHtml(name)}</strong>
              <span>${escapeHtml(placement)} · ${escapeHtml(role)}</span>
            </button>`;
        }).join("")}
      </div>
      <div class="morning-slot-actions">
        ${currentOperator ? `<button type="button" class="morning-slot-secondary" data-morning-slot-view="${escapeHtml(currentOperator)}">View Output</button>` : ""}
        ${currentOperator && canEdit ? `<button type="button" class="morning-slot-danger" data-morning-slot-clear="${escapeHtml(currentOperator)}">Clear Position</button>` : ""}
      </div>`;

    body.querySelectorAll("[data-morning-slot-operator]").forEach(button => {
      button.addEventListener("click", () => assignMorningOperatorToSlot(button.dataset.morningSlotOperator, line, station, slot));
    });

    body.querySelector("[data-morning-slot-view]")?.addEventListener("click", event => {
      const op = event.currentTarget.dataset.morningSlotView;
      closeMorningLineSlotAssignment();
      if (finishOperatorState?.stations?.[normalizeFinishOperatorStation(station)]) {
        openExpandedOperatorViewer(normalizeFinishOperatorStation(station), op);
      }
    });

    body.querySelector("[data-morning-slot-clear]")?.addEventListener("click", event => {
      clearMorningOperatorPosition(event.currentTarget.dataset.morningSlotClear, station);
    });
  }

  modal.dataset.line = line;
  modal.dataset.station = station;
  modal.dataset.slot = slot;
  modal.classList.add("open");
}

function closeMorningLineSlotAssignment() {
  document.getElementById("morningSlotAssignModal")?.classList.remove("open");
}

async function assignMorningOperatorToSlot(operatorName, line, station, slot) {
  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role cannot change Morning Set Up positions.");
    return;
  }

  const existing = findMorningRosterRow(operatorName);
  if (!existing) {
    showFinishAssignmentToast("This associate is not active on your selected shift roster.");
    return;
  }

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getMorningSetupOwnerUsername(),
      shiftType: getMorningSetupShiftType(),
      operatorName,
      activeStatus: "Active",
      defaultLine: line,
      defaultArea: normalizeFinishOperatorStation(station),
      defaultPosition: slot,
      roleType: existing.roleType || "Unassigned",
      trainingWeek: existing.roleType === "Training" ? (existing.trainingWeek || "") : "",
      individualJph: existing.individualJph || "",
      originalDefaultArea: existing.defaultArea || "",
      originalDefaultLine: existing.defaultLine || "",
      originalDefaultPosition: existing.defaultPosition || ""
    });

    showFinishAssignmentToast(`Assigned ${operatorName} to ${line} / ${slot}`);
    closeMorningLineSlotAssignment();
    await loadMorningSetupRoster({ silent: true });
    await loadFinishRosterForSelectedProfile({ silent: true }).catch(() => {});
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

async function clearMorningOperatorPosition(operatorName, station) {
  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role cannot clear Morning Set Up positions.");
    return;
  }

  const existing = findMorningRosterRow(operatorName);
  if (!existing) return;

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getMorningSetupOwnerUsername(),
      shiftType: getMorningSetupShiftType(),
      operatorName,
      activeStatus: "Active",
      defaultLine: "",
      defaultArea: normalizeFinishOperatorStation(station),
      defaultPosition: "",
      roleType: existing.roleType || "Unassigned",
      trainingWeek: existing.roleType === "Training" ? (existing.trainingWeek || "") : ""
    });

    showFinishAssignmentToast(`Cleared position for ${operatorName}`);
    closeMorningLineSlotAssignment();
    await loadMorningSetupRoster({ silent: true });
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

function initMorningSetupControls() {
  const select = document.getElementById("morningSetupShiftSelect");
  if (!select) return;

  select.value = getMorningSetupShiftType();
  select.addEventListener("change", async event => {
    finishRosterApiState.morningShiftType = String(event.target.value || "Weekday");
    await loadMorningSetupRoster({ silent: false });
  });
}

function renderOperatorError(error) {
  const grid = document.getElementById("operatorStationGrid");
  if (!grid) return;

  grid.innerHTML = `
    <article class="operator-empty-card operator-error-card">
      Finish Operator API could not load.<br />
      <small>${escapeHtml(error.message || error)}</small>
    </article>`;
}


/* =========================================================
   FINISH ROSTER API FRONTEND OVERRIDE
   Source of truth: FINISH_USER_PROFILES, FINISH_OPERATOR_MASTER,
   FINISH_OPERATOR_ROSTER_CONTROL, FINISH_ROSTER_AUDIT_LOG.
   localStorage remains only for dashboard capacity config and UI-only line status.
========================================================= */
const FINISH_ROSTER_API_BASE_URL = "https://script.google.com/macros/s/AKfycbxJR3xCmLA-CW8WamTDuW3704meywwulltVe7i4-wmS7ulZN2YpnMrxwawbcVjcfLJ93Q/exec";

const finishRosterApiState = {
  profiles: [],
  operators: [],
  roster: [],
  audit: [],
  ownerUsername: "",
  shiftType: "Weekday",

  // Morning Set Up is intentionally separate from LMS Control Center.
  // It always belongs to the logged-in profile and selected Morning shift.
  morningOwnerUsername: "",
  morningShiftType: "Weekday",
  morningRoster: [],

  loading: false,
  ready: false,
  lastError: ""
};

const finishRosterUiState = {
  expandedRosterRows: new Set(),
  collapsedLines: new Set()
};

function buildFinishRosterApiUrl(action, params = {}) {
  const url = new URL(FINISH_ROSTER_API_BASE_URL);
  url.searchParams.set("action", action);
  Object.entries(params || {}).forEach(([key, value]) => {
    if (value !== undefined && value !== null && String(value) !== "") {
      url.searchParams.set(key, String(value));
    }
  });
  return url.toString();
}

async function fetchFinishRosterApi(action, params = {}) {
  const response = await fetch(buildFinishRosterApiUrl(action, params), {
    method: "GET",
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error(`Finish roster API failed: ${response.status}`);
  }

  const payload = await response.json();

  if (payload.status === "error" || payload.success === false) {
    throw new Error(payload.message || payload.error || `Finish roster API action failed: ${action}`);
  }

  return payload;
}

function getFinishRosterCurrentProfile() {
  const current = getCurrentUsername();
  return (finishRosterApiState.profiles || []).find(profile =>
    String(profile.username || "").toUpperCase() === current
  ) || null;
}

function canEditFinishOperatorAssignments() {
  const profile = getFinishRosterCurrentProfile();

  if (profile) {
    return !!profile.canEdit && !!profile.isActive && FINISH_ASSIGNMENT_ALLOWED_EDIT_ROLE_CHECK(profile.role);
  }

  return FINISH_ASSIGNMENT_ALLOWED_EDIT_ROLE_CHECK(getCurrentUserRole());
}

function getRosterOwnerUsername() {
  return finishRosterApiState.ownerUsername || getCurrentUsername();
}

function getRosterShiftType() {
  return finishRosterApiState.shiftType || "Weekday";
}

function getMorningSetupOwnerUsername() {
  return getCurrentUsername();
}

function getMorningSetupShiftType() {
  const select = document.getElementById("morningSetupShiftSelect");
  const value = String(select?.value || finishRosterApiState.morningShiftType || "Weekday").trim().toLowerCase();
  return value === "weekend" ? "Weekend" : "Weekday";
}

function getMorningSetupShiftHours() {
  return getMorningSetupShiftType() === "Weekend" ? 10.5 : 9.5;
}

async function loadMorningSetupRoster(options = {}) {
  const ownerUsername = getMorningSetupOwnerUsername();
  const shiftType = getMorningSetupShiftType();

  finishRosterApiState.morningOwnerUsername = ownerUsername;
  finishRosterApiState.morningShiftType = shiftType;

  try {
    const rosterPayload = await fetchFinishRosterApi("getFinishRosterControl", {
      ownerUsername,
      shiftType
    });

    finishRosterApiState.morningRoster = (rosterPayload.data || [])
      .filter(row => String(row.ownerUsername || "").toUpperCase() === ownerUsername)
      .filter(row => String(row.shiftType || "").toLowerCase() === shiftType.toLowerCase())
      .filter(row => String(row.activeStatus || "").toLowerCase() !== "removed");

    if (!options.silent) {
      renderFinishThreeLineCell();
      renderMorningSetupRosterSummary();
    }

    return finishRosterApiState.morningRoster;
  } catch (error) {
    console.error("Morning Set Up roster load failed:", error);
    finishRosterApiState.morningRoster = [];
    if (!options.silent) {
      showFinishAssignmentToast(error.message || String(error));
      renderFinishThreeLineCell();
      renderMorningSetupRosterSummary();
    }
    return [];
  }
}

function convertApiRoleToLocal(roleType) {
  const role = String(roleType || "").toLowerCase();
  if (role === "certified") return "core";
  if (role === "tq") return "tq";
  if (role === "training") return "training";
  return "ignore";
}

function convertLocalRoleToApi(role) {
  const value = String(role || "").toLowerCase();
  if (value === "core" || value === "certified") return "Certified";
  if (value === "tq") return "TQ";
  if (value === "training") return "Training";
  return "Unassigned";
}

function convertApiWeekToNumber(week) {
  const match = String(week || "").match(/(\d+)/);
  return match ? Number(match[1]) : 1;
}

function convertNumberToApiWeek(week) {
  const n = Number(week || 0);
  return n > 0 ? `W${n}` : "";
}

function loadFinishPersonalOperatorAssignments(rosterRows) {
  const assignments = {};
  const sourceRows = Array.isArray(rosterRows) ? rosterRows : (finishRosterApiState.roster || []);

  sourceRows
    .filter(row => String(row.activeStatus || "").toLowerCase() !== "removed")
    .forEach(row => {
      const station = normalizeFinishOperatorStation(row.defaultArea || row.station || "");
      const operator = normalizeAssignmentOperatorName(row.operatorName || row.operator || "");
      const line = row.defaultLine || "";
      const slot = row.defaultPosition || "";
      if (!operator || !station) return;

      const role = convertApiRoleToLocal(row.roleType);
      if (role === "ignore" && !line && !slot) return;

      // Include line/slot in the key so the same associate can have multiple roles/positions.
      const key = `${makeAssignmentKey(station, operator)}||${line}||${slot}`;

      assignments[key] = {
        station,
        operator,
        role,
        trainingWeek: convertApiWeekToNumber(row.trainingWeek),
        individualJph: Number(row.individualJph || row.IndividualJPH || 0) || 0,
        line,
        slot,
        activeStatus: row.activeStatus || "Active",
        updatedAt: row.updatedAt || "",
        updatedBy: row.updatedBy || ""
      };
    });

  return assignments;
}

function saveFinishPersonalOperatorAssignments() {
  // Backend save is handled per-row by saveOperatorAssignment().
  return true;
}

function getFinishOperatorRoster() {
  const rosterByKey = {};

  (finishRosterApiState.operators || []).forEach(op => {
    const operatorName = normalizeAssignmentOperatorName(op.operatorName || op.OperatorName || "");
    if (!operatorName) return;

    const stationName = normalizeFinishOperatorStation(op.lastFlowStation || op.LastFlowStation || op.lastAccessPoint || op.LastAccessPoint || "");
    if (!stationName || stationName === "FSV Scan & Verify") return;

    const key = makeAssignmentKey(stationName, operatorName);
    rosterByKey[key] = {
      station: stationName,
      operator: operatorName,
      total: Number(op.lastTotal || op.LastTotal || 0) || 0,
      liveNow: true,
      accessPoints: [op.lastAccessPoint || op.LastAccessPoint || ""].filter(Boolean)
    };
  });

  // Include saved roster rows even if that operator has no live activity today.
  (finishRosterApiState.roster || []).forEach(row => {
    if (String(row.activeStatus || "").toLowerCase() === "removed") return;

    const stationName = normalizeFinishOperatorStation(row.defaultArea || "");
    const operatorName = normalizeAssignmentOperatorName(row.operatorName || "");
    if (!stationName || !operatorName) return;

    const key = makeAssignmentKey(stationName, operatorName);
    if (!rosterByKey[key]) {
      rosterByKey[key] = {
        station: stationName,
        operator: operatorName,
        total: 0,
        liveNow: false,
        accessPoints: []
      };
    }
  });

  return Object.values(rosterByKey).sort((a, b) => {
    const stationCompare = a.station.localeCompare(b.station);
    if (stationCompare !== 0) return stationCompare;
    return a.operator.localeCompare(b.operator);
  });
}

async function loadFinishRosterBackend(options = {}) {
  const silent = !!options.silent;
  finishRosterApiState.loading = true;
  finishRosterApiState.lastError = "";
  setText("finishRosterApiStatus", "Loading...");

  try {
    if (!finishRosterApiState.ownerUsername) {
      finishRosterApiState.ownerUsername = getCurrentUsername();
    }

    const [profilesPayload, operatorsPayload] = await Promise.all([
      fetchFinishRosterApi("getFinishUserProfiles"),
      fetchFinishRosterApi("getFinishOperatorMaster")
    ]);

    finishRosterApiState.profiles = profilesPayload.data || [];
    finishRosterApiState.operators = operatorsPayload.data || [];

    const currentProfile = getFinishRosterCurrentProfile();
    if (!currentProfile && finishRosterApiState.ownerUsername === "DEFAULT") {
      finishRosterApiState.ownerUsername = (finishRosterApiState.profiles[0]?.username || getCurrentUsername()).toUpperCase();
    }

    await loadFinishRosterForSelectedProfile({ silent: true });
    await loadMorningSetupRoster({ silent: true });
    finishRosterApiState.ready = true;
    setText("finishRosterApiStatus", "Connected");
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  } catch (error) {
    finishRosterApiState.lastError = error.message || String(error);
    console.error("Finish roster backend error:", error);
    setText("finishRosterApiStatus", "Error");
    renderFinishRosterApiPanel();
    if (!silent) showFinishAssignmentToast(finishRosterApiState.lastError);
  } finally {
    finishRosterApiState.loading = false;
  }
}

async function loadFinishRosterForSelectedProfile(options = {}) {
  const ownerUsername = getRosterOwnerUsername();
  const shiftType = getRosterShiftType();

  const [rosterPayload, auditPayload] = await Promise.all([
    fetchFinishRosterApi("getFinishRosterControl", { ownerUsername, shiftType }),
    fetchFinishRosterApi("getFinishRosterAuditLog", { targetUsername: ownerUsername, limit: 75 })
  ]);

  finishRosterApiState.roster = rosterPayload.data || [];
  finishRosterApiState.audit = auditPayload.data || [];

  if (!options.silent) {
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  }
}

function initFinishRosterApiPanel() {
  const ownerSelect = document.getElementById("finishRosterOwnerSelect");
  const shiftSelect = document.getElementById("finishRosterShiftSelect");
  const areaFilter = document.getElementById("finishRosterAreaFilter");
  const refreshBtn = document.getElementById("finishRosterRefreshBtn");

  ownerSelect?.addEventListener("change", async event => {
    finishRosterApiState.ownerUsername = String(event.target.value || getCurrentUsername()).toUpperCase();
    await loadFinishRosterForSelectedProfile();
  });

  shiftSelect?.addEventListener("change", async event => {
    finishRosterApiState.shiftType = String(event.target.value || "Weekday");
    await loadFinishRosterForSelectedProfile();
  });

  areaFilter?.addEventListener("change", renderFinishRosterApiPanel);
  refreshBtn?.addEventListener("click", () => loadFinishRosterBackend());
}

function renderFinishRosterApiPanel() {
  const panel = document.getElementById("finishRosterApiPanel");
  if (!panel) return;

  const ownerSelect = document.getElementById("finishRosterOwnerSelect");
  const shiftSelect = document.getElementById("finishRosterShiftSelect");
  const areaFilter = document.getElementById("finishRosterAreaFilter");
  const list = document.getElementById("finishRosterEditorList");
  const banner = document.getElementById("finishRosterPermissionBanner");

  const currentUsername = getCurrentUsername();
  const canEdit = canEditFinishOperatorAssignments();
  const currentProfile = getFinishRosterCurrentProfile();
  const selectedOwner = getRosterOwnerUsername();
  const selectedShift = getRosterShiftType();

  if (ownerSelect) {
    const profiles = finishRosterApiState.profiles.length
      ? finishRosterApiState.profiles
      : [{ username: currentUsername, fullName: getCurrentUserDisplayName(), role: getCurrentUserRole(), shiftGroup: "Weekday" }];

    ownerSelect.innerHTML = profiles.map(profile => {
      const username = String(profile.username || "").toUpperCase();
      const label = `${profile.fullName || username} · ${profile.role || "Role"} · ${profile.shiftGroup || "Shift"}`;
      return `<option value="${escapeHtml(username)}" ${username === selectedOwner ? "selected" : ""}>${escapeHtml(label)}</option>`;
    }).join("");
    ownerSelect.disabled = !canEdit;
  }

  if (shiftSelect) shiftSelect.value = selectedShift;

  if (banner) {
    banner.classList.toggle("can-edit", canEdit);
    banner.innerHTML = canEdit
      ? `<strong>Edit access active</strong><span>${escapeHtml(getCurrentUserDisplayName())} · ${escapeHtml(currentProfile?.role || getCurrentUserRole())} can save/edit/delete roster rows.</span>`
      : `<strong>View only</strong><span>${escapeHtml(getCurrentUserDisplayName())} · your role can view this roster, but the API blocks edit/delete.</span>`;
  }

  const filter = areaFilter?.value || "all";
  const rows = buildFinishRosterEditorRows(filter);

  setText("finishRosterOperatorCount", finishRosterApiState.operators.length || 0);
  setText("finishRosterActiveCount", (finishRosterApiState.roster || []).filter(r => String(r.activeStatus).toLowerCase() === "active").length);
  setText("finishRosterMountingCount", (finishRosterApiState.roster || []).filter(r => String(r.activeStatus).toLowerCase() === "active" && r.defaultArea === "Mounting").length);
  setText("finishRosterFinalCount", (finishRosterApiState.roster || []).filter(r => String(r.activeStatus).toLowerCase() === "active" && r.defaultArea === "Final Inspection").length);

  if (list) {
    if (finishRosterApiState.lastError) {
      list.innerHTML = `<article class="roster-api-empty">${escapeHtml(finishRosterApiState.lastError)}</article>`;
    } else if (!rows.length) {
      list.innerHTML = `<article class="roster-api-empty">No floor operators available yet. Run syncFinishOperatorMasterFromRawActivity() or wait for RAW_ACTIVITY_CURRENT to populate.</article>`;
    } else {
      list.innerHTML = rows.map(row => renderFinishRosterEditorRow(row, canEdit)).join("");
      bindFinishRosterEditorEvents();
    }
  }

  renderFinishRosterAuditLog();
}

function makeFinishRosterAssignmentKey(row) {
  return [
    normalizeAssignmentOperatorName(row.operatorName || ""),
    String(row.shiftType || getRosterShiftType() || "Weekday"),
    normalizeFinishOperatorStation(row.defaultArea || ""),
    String(row.defaultLine || ""),
    String(row.defaultPosition || "")
  ].join("||").toLowerCase();
}

function buildFinishRosterEditorRows(filter = "all") {
  const selectedShift = getRosterShiftType();
  const rosterRows = Array.isArray(finishRosterApiState.roster) ? finishRosterApiState.roster : [];
  const masterOps = Array.isArray(finishRosterApiState.operators) ? finishRosterApiState.operators : [];

  const liveByName = {};
  masterOps.forEach(op => {
    const operatorName = normalizeAssignmentOperatorName(op.operatorName || "");
    if (!operatorName) return;

    liveByName[operatorName.toLowerCase()] = {
      liveStation: normalizeFinishOperatorStation(op.lastFlowStation || op.lastAccessPoint || ""),
      lastTotal: Number(op.lastTotal || 0) || 0
    };
  });

  const rows = [];

  // 1) Saved assignments first. This is the source of truth.
  rosterRows.forEach(saved => {
    const operatorName = normalizeAssignmentOperatorName(saved.operatorName || "");
    if (!operatorName) return;

    const live = liveByName[operatorName.toLowerCase()] || {};

    rows.push({
      operatorName,
      shiftType: saved.shiftType || selectedShift,
      liveStation: live.liveStation || normalizeFinishOperatorStation(saved.defaultArea || ""),
      lastTotal: Number(live.lastTotal || 0) || 0,
      activeStatus: saved.activeStatus || "Active",
      defaultLine: saved.defaultLine || "",
      defaultArea: normalizeFinishOperatorStation(saved.defaultArea || ""),
      defaultPosition: saved.defaultPosition || "",
      roleType: saved.roleType || "Unassigned",
      trainingWeek: saved.trainingWeek || "",
      individualJph: saved.individualJph ?? saved.individualJPH ?? "",
      isSaved: true
    });
  });

  // 2) Add unsaved floor operators so LMS can assign new people.
  const savedNames = new Set(
    rows.map(row => normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase())
  );

  masterOps.forEach(op => {
    const operatorName = normalizeAssignmentOperatorName(op.operatorName || "");
    if (!operatorName) return;
    if (operatorName.toLowerCase().includes("unassigned")) return;
    if (savedNames.has(operatorName.toLowerCase())) return;

    const liveStation = normalizeFinishOperatorStation(op.lastFlowStation || op.lastAccessPoint || "");
    if (liveStation === "FSV Scan & Verify") return;

    rows.push({
      operatorName,
      shiftType: selectedShift,
      liveStation,
      lastTotal: Number(op.lastTotal || 0) || 0,
      activeStatus: "Active",
      defaultLine: "",
      defaultArea: isLineSlotStation(liveStation) ? liveStation : "Mounting",
      defaultPosition: "",
      roleType: "Unassigned",
      trainingWeek: "",
      individualJph: "",
      isSaved: false
    });
  });

  return rows
    .filter(row => filter === "all" || row.defaultArea === filter || row.liveStation === filter)
    .sort((a, b) => {
      const activeA = a.activeStatus === "Active" ? 0 : 1;
      const activeB = b.activeStatus === "Active" ? 0 : 1;
      if (activeA !== activeB) return activeA - activeB;

      const savedA = a.isSaved ? 0 : 1;
      const savedB = b.isSaved ? 0 : 1;
      if (savedA !== savedB) return savedA - savedB;

      const nameCompare = a.operatorName.localeCompare(b.operatorName);
      if (nameCompare !== 0) return nameCompare;

      const areaCompare = String(a.defaultArea || "").localeCompare(String(b.defaultArea || ""));
      if (areaCompare !== 0) return areaCompare;

      return String(a.defaultPosition || "").localeCompare(String(b.defaultPosition || ""));
    });
}

function renderFinishRosterEditorRow(row, canEdit) {
  const disabled = canEdit ? "" : "disabled";
  const statusClass = row.activeStatus === "Archived" ? "is-archived" : "is-active";
  const rowKey = makeFinishRosterAssignmentKey(row);
  const expanded = finishRosterUiState.expandedRosterRows.has(rowKey);
  const shiftType = row.shiftType || getRosterShiftType();

  const target = row.roleType === "Training"
    ? `${row.trainingWeek || "W?"}`
    : row.roleType && row.roleType !== "Unassigned"
      ? row.roleType
      : "—";

  return `
    <article
      class="finish-roster-row ${statusClass} ${expanded ? "is-expanded" : "is-collapsed"}"
      data-roster-operator="${escapeHtml(row.operatorName)}"
      data-roster-key="${escapeHtml(rowKey)}"
      data-roster-shift="${escapeHtml(shiftType)}"
      data-roster-area="${escapeHtml(row.defaultArea || "")}"
      data-roster-line="${escapeHtml(row.defaultLine || "")}"
      data-roster-position="${escapeHtml(row.defaultPosition || "")}"
    >
      <div class="finish-roster-row-head">
        <div class="finish-roster-operator">
          <strong>${escapeHtml(row.operatorName)}</strong>
          <span>
            Live: ${escapeHtml(row.liveStation || "—")}
            · Today ${numberFmt(row.lastTotal)}
            · ${row.isSaved ? "Saved" : "Not saved"}
          </span>
        </div>

        <div class="finish-roster-row-tags">
          <span>${escapeHtml(row.activeStatus || "Active")}</span>
          <span>${escapeHtml(shiftType)}</span>
          <span>${escapeHtml(row.defaultArea || "No area")}</span>
          <span>${escapeHtml(row.defaultLine || "No line")}</span>
          <span>${escapeHtml(row.defaultPosition || "No position")}</span>
          <span>${escapeHtml(target)}</span>
        </div>

        <button class="finish-roster-row-toggle" type="button" data-roster-toggle>
          ${expanded ? "Close" : "Open"}
        </button>
      </div>

      <div class="finish-roster-row-controls">
        <label><span>Status</span>
          <select data-roster-field="activeStatus" ${disabled}>
            ${["Active", "Archived"].map(v => `<option value="${v}" ${row.activeStatus === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label><span>Shift</span>
          <select data-roster-field="shiftType" ${disabled}>
            ${["Weekday", "Weekend"].map(v => `<option value="${v}" ${shiftType === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label><span>Area</span>
          <select data-roster-field="defaultArea" ${disabled}>
            ${["Mounting", "Final Inspection", "Drill"].map(v => `<option value="${v}" ${row.defaultArea === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label><span>Line</span>
          <select data-roster-field="defaultLine" ${disabled}>
            ${["", "Line A", "Line B", "Line C"].map(v => `<option value="${v}" ${row.defaultLine === v ? "selected" : ""}>${v || "Select"}</option>`).join("")}
          </select>
        </label>

        <label><span>Position</span>
          <select data-roster-field="defaultPosition" ${disabled}>
            ${buildRosterPositionOptions(row.defaultArea, row.defaultPosition)}
          </select>
        </label>

        <label><span>Role</span>
          <select data-roster-field="roleType" ${disabled}>
            ${["Unassigned", "Certified", "TQ", "Training"].map(v => `<option value="${v}" ${row.roleType === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label><span>Week</span>
          <select data-roster-field="trainingWeek" ${disabled || row.roleType !== "Training" ? "disabled" : ""}>
            ${["", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"].map(v => `<option value="${v}" ${row.trainingWeek === v ? "selected" : ""}>${v || "—"}</option>`).join("")}
          </select>
        </label>

        <label><span>Individual JPH</span>
          <input
            data-roster-field="individualJph"
            type="number"
            min="0"
            step="0.1"
            placeholder="Default"
            value="${escapeHtml(row.individualJph || "")}"
            ${disabled}
          />
        </label>

        <div class="finish-roster-actions">
          <button class="finish-roster-action-btn archive" type="button" data-roster-archive ${disabled}>Archive</button>
          <button class="finish-roster-action-btn delete" type="button" data-roster-delete ${disabled}>Delete</button>
        </div>
      </div>
    </article>`;
}

function buildRosterPositionOptions(area, selected) {
  const normalizedArea = normalizeFinishOperatorStation(area);

  const slots =
    normalizedArea === "Final Inspection"
      ? ["", "FI-01", "FI-02", "FI-03", "FI-04"]
      : normalizedArea === "Mounting"
        ? ["", "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10"]
        : normalizedArea === "Drill"
          ? ["", "D07", "D09"]
          : [""];

  return slots
    .map(slot => `<option value="${escapeHtml(slot)}" ${slot === selected ? "selected" : ""}>${escapeHtml(slot || "Select")}</option>`)
    .join("");
}

function bindFinishRosterEditorEvents() {
  document.querySelectorAll("[data-roster-toggle]").forEach(button => {
    button.addEventListener("click", event => {
      const row = event.target.closest(".finish-roster-row");
      const rowKey = row?.dataset.rosterKey || "";
      if (!rowKey) return;

      if (finishRosterUiState.expandedRosterRows.has(rowKey)) {
        finishRosterUiState.expandedRosterRows.delete(rowKey);
        row.classList.remove("is-expanded");
        row.classList.add("is-collapsed");
        button.textContent = "Open";
      } else {
        finishRosterUiState.expandedRosterRows.add(rowKey);
        row.classList.remove("is-collapsed");
        row.classList.add("is-expanded");
        button.textContent = "Close";
      }
    });
  });

  document.querySelectorAll("[data-roster-field]").forEach(input => {
    const eventName = input.tagName === "INPUT" ? "change" : "change";

    input.addEventListener(eventName, event => {
      const row = event.target.closest(".finish-roster-row");
      if (!row) return;

      if (event.target.dataset.rosterField === "defaultArea") {
        const positionSelect = row.querySelector('[data-roster-field="defaultPosition"]');
        if (positionSelect) {
          positionSelect.innerHTML = buildRosterPositionOptions(event.target.value, "");
        }

        const lineSelect = row.querySelector('[data-roster-field="defaultLine"]');
        if (lineSelect && event.target.value === "Drill") {
          lineSelect.value = "Line C";
        }
      }

      if (event.target.dataset.rosterField === "roleType") {
        const weekSelect = row.querySelector('[data-roster-field="trainingWeek"]');
        if (weekSelect) {
          weekSelect.disabled = event.target.value !== "Training";
          if (event.target.value !== "Training") weekSelect.value = "";
        }
      }

      saveFinishRosterRowFromDom(row);
    });
  });

  document.querySelectorAll("[data-roster-archive]").forEach(button => {
    button.addEventListener("click", event =>
      updateFinishRosterRowStatus(event.target.closest(".finish-roster-row"), "archiveFinishRosterOperator")
    );
  });

  document.querySelectorAll("[data-roster-delete]").forEach(button => {
    button.addEventListener("click", event => {
      const row = event.target.closest(".finish-roster-row");
      const operator = row?.dataset.rosterOperator || "";
      const shift = row?.querySelector('[data-roster-field="shiftType"]')?.value || getRosterShiftType();
      if (!operator) return;

      if (confirm(`Remove ${operator} from this ${shift} roster row? This marks the row as Removed and keeps audit history.`)) {
        updateFinishRosterRowStatus(row, "deleteFinishRosterOperator");
      }
    });
  });
}

async function saveFinishRosterRowFromDom(rowEl) {
  if (!rowEl) return;

  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role cannot edit Finish roster rows.");
    renderFinishRosterApiPanel();
    return;
  }

  const operatorName = rowEl.dataset.rosterOperator;
  const getField = field => rowEl.querySelector(`[data-roster-field="${field}"]`)?.value || "";

  const saveShift = getField("shiftType") || rowEl.dataset.rosterShift || getRosterShiftType();
  const defaultArea = getField("defaultArea");
  const defaultLine = getField("defaultLine");
  const defaultPosition = getField("defaultPosition");

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getRosterOwnerUsername(),
      shiftType: saveShift,
      operatorName,
      activeStatus: getField("activeStatus") || "Active",
      defaultLine,
      defaultArea,
      defaultPosition,
      roleType: getField("roleType"),
      trainingWeek: getField("roleType") === "Training" ? getField("trainingWeek") : "",
      individualJph: getField("individualJph")
    });

    showFinishAssignmentToast(`Saved ${operatorName} / ${saveShift}`);

    if (saveShift !== getRosterShiftType()) {
      finishRosterApiState.shiftType = saveShift;
      const shiftSelect = document.getElementById("finishRosterShiftSelect");
      if (shiftSelect) shiftSelect.value = saveShift;
    }

    await loadFinishRosterForSelectedProfile({ silent: true });
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
    await loadFinishRosterForSelectedProfile({ silent: true }).catch(() => {});
    renderFinishRosterApiPanel();
  }
}

async function updateFinishRosterRowStatus(rowEl, action) {
  if (!rowEl) return;

  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role cannot delete/archive Finish roster rows.");
    return;
  }

  const operatorName = rowEl.dataset.rosterOperator;
  const getField = field => rowEl.querySelector(`[data-roster-field="${field}"]`)?.value || "";

  try {
    await fetchFinishRosterApi(action, {
      updatedBy: getCurrentUsername(),
      ownerUsername: getRosterOwnerUsername(),
      shiftType: getField("shiftType") || rowEl.dataset.rosterShift || getRosterShiftType(),
      operatorName,
      defaultArea: getField("defaultArea") || rowEl.dataset.rosterArea || "",
      defaultLine: getField("defaultLine") || rowEl.dataset.rosterLine || "",
      defaultPosition: getField("defaultPosition") || rowEl.dataset.rosterPosition || ""
    });

    showFinishAssignmentToast(`${action.includes("delete") ? "Deleted" : "Archived"} ${operatorName}`);
    await loadFinishRosterForSelectedProfile({ silent: true });
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

function renderFinishRosterAuditLog() {
  const target = document.getElementById("finishRosterAuditLog");
  if (!target) return;

  const logs = finishRosterApiState.audit || [];
  if (!logs.length) {
    target.innerHTML = `<article class="roster-api-empty">No audit changes for this profile yet.</article>`;
    return;
  }

  target.innerHTML = logs.map(row => `
    <article class="roster-audit-row">
      <span>${escapeHtml(row.timestamp || "")}</span>
      <strong>${escapeHtml(row.action || "")}</strong>
      <span>${escapeHtml(row.operatorName || "")}</span>
      <span>${escapeHtml(row.fieldChanged || "")} · ${escapeHtml(row.oldValue || "—")} → ${escapeHtml(row.newValue || "—")}</span>
      <em>${escapeHtml(row.updatedBy || "")}</em>
    </article>
  `).join("");
}

function renderOperatorAssignmentConfigPanel() {
  ensureOperatorAssignmentConfigPanel();
  renderFinishRosterApiPanel();

  const rosterTarget = document.getElementById("operatorAssignmentRoster");
  if (!rosterTarget) return;

  const config = loadConfig();
  const filter = document.getElementById("assignmentStationFilter")?.value || "all";
  const roster = getFinishOperatorRoster()
    .filter(item => filter === "all" || item.station === filter);

  if (!roster.length) {
    rosterTarget.innerHTML = `
      <article class="assignment-empty">
        No Finish operator master loaded yet. Use the Personal tab roster control or run syncFinishOperatorMasterFromRawActivity().
      </article>`;
    updateConfigTotals();
    applyFinishAssignmentPermissionLock();
    return;
  }

  rosterTarget.innerHTML = roster.map(item => {
    const assignment = getOperatorAssignment(item.operator, item.station, config) || {
      role: "ignore",
      trainingWeek: 1
    };
    const canTrain = item.station === "Mounting" || item.station === "Final Inspection";
    const role = String(assignment.role || "ignore");
    const trainingRate = getTrainingRateForStationWeek(item.station, assignment.trainingWeek || 1);
    const baseRate = getFinishOperatorBaseRate(item.station, item.operator);
    const effective = role === "training" ? trainingRate : (["core", "tq", "certified"].includes(role) ? baseRate : 0);

    return `
      <article class="assignment-row" data-assignment-row="${escapeHtml(makeAssignmentKey(item.station, item.operator))}">
        <div class="assignment-operator">
          <strong>${escapeHtml(item.operator)}</strong>
          <span>${escapeHtml(item.station)} · Today ${numberFmt(item.total)}${item.liveNow ? " · Live" : " · Saved API"}</span>
        </div>

        <label>
          Role
          <select data-assignment-role data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}">
            <option value="ignore" ${role === "ignore" ? "selected" : ""}>Do not count</option>
            <option value="core" ${role === "core" || role === "certified" ? "selected" : ""}>Certified / Core Capacity</option>
            <option value="tq" ${role === "tq" ? "selected" : ""}>TQ / Core Capacity</option>
            ${canTrain ? `<option value="training" ${role === "training" ? "selected" : ""}>Training Metric</option>` : ""}
          </select>
        </label>

        <label class="assignment-week ${role === "training" && canTrain ? "active" : "disabled"}">
          Week
          <select data-assignment-week data-station="${escapeHtml(item.station)}" data-operator="${escapeHtml(item.operator)}" ${role === "training" && canTrain ? "" : "disabled"}>
            ${buildTrainingWeekOptions(item.station, assignment.trainingWeek || 1)}
          </select>
        </label>

        ${buildLineSlotAssignmentControls(item.station, item.operator, assignment)}

        <div class="assignment-target">
          <span>Target</span>
          <strong>${effective > 0 ? numberFmt(effective) + "/hr" : "—"}</strong>
        </div>
      </article>`;
  }).join("");

  rosterTarget.querySelectorAll("[data-assignment-role]").forEach(select => {
    select.addEventListener("change", event => {
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        event.target.value,
        Number(getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station)?.trainingWeek || 1)
      );
    });
  });

  rosterTarget.querySelectorAll("[data-assignment-week]").forEach(select => {
    select.addEventListener("change", event => {
      const current = getOperatorAssignment(event.target.dataset.operator, event.target.dataset.station) || { role: "training" };
      saveOperatorAssignment(
        event.target.dataset.station,
        event.target.dataset.operator,
        current.role || "training",
        Number(event.target.value || 1),
        current.line || "",
        current.slot || ""
      );
    });
  });

  rosterTarget.querySelectorAll("[data-assignment-line], [data-assignment-slot]").forEach(select => {
    select.addEventListener("change", event => {
      const station = event.target.dataset.station;
      const operator = event.target.dataset.operator;
      const row = event.target.closest(".assignment-row");
      const current = getOperatorAssignment(operator, station) || { role: "ignore", trainingWeek: 1 };
      const line = row?.querySelector("[data-assignment-line]")?.value || current.line || "";
      const slot = row?.querySelector("[data-assignment-slot]")?.value || current.slot || "";
      saveOperatorAssignment(
        station,
        operator,
        current.role || "ignore",
        Number(current.trainingWeek || 1),
        line,
        slot
      );
    });
  });

  updateConfigTotals();
  applyFinishAssignmentPermissionLock();
  renderFinishThreeLineCell();
}

async function saveOperatorAssignment(stationName, operatorName, role, trainingWeek = 1, line = "", slot = "") {
  if (!canEditFinishOperatorAssignments()) {
    showFinishAssignmentToast("View only. Your role can see operator setup but cannot edit it.");
    renderOperatorAssignmentConfigPanel();
    return;
  }

  const station = normalizeFinishOperatorStation(stationName);
  const operator = normalizeAssignmentOperatorName(operatorName);
  if (!operator) return;

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getRosterOwnerUsername(),
      shiftType: getRosterShiftType(),
      operatorName: operator,
      activeStatus: role && role !== "ignore" ? "Active" : "Archived",
      defaultLine: line || getOperatorAssignment(operator, station)?.line || "",
      defaultArea: station,
      defaultPosition: slot || getOperatorAssignment(operator, station)?.slot || "",
      roleType: convertLocalRoleToApi(role),
      trainingWeek: role === "training" ? convertNumberToApiWeek(trainingWeek) : "",
      individualJph: getOperatorAssignment(operator, station)?.individualJph || ""
    });

    await loadFinishRosterForSelectedProfile({ silent: true });
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    loadDashboard({ showLoader: false });
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();

    if (finishOperatorState?.selectedStation) {
      const activeStation = finishOperatorState.stations?.[finishOperatorState.selectedStation];
      if (activeStation) renderDrawerOperators(activeStation);
    }
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
    await loadFinishRosterForSelectedProfile({ silent: true }).catch(() => {});
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
  }
}


async function saveConfigAssignedIndividualJph(rowEl) {
  if (!rowEl) return;

  const operatorName = rowEl.dataset.assignedRosterRow || "";
  const source = (window.__finishConfigAssignedRosterRows || []).find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "") === normalizeAssignmentOperatorName(operatorName) &&
    normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(rowEl.dataset.area || rowEl.querySelector(".assignment-operator span")?.textContent?.split(" · ")?.[0] || "")
  );

  const areaLinePositionText = rowEl.querySelector(".assignment-operator span")?.textContent || "";
  const parts = areaLinePositionText.split(" · ").map(x => String(x || "").trim());

  const defaultArea = source?.defaultArea || parts[0] || "";
  const defaultLine = source?.defaultLine || parts[1] || "";
  const defaultPosition = source?.defaultPosition || parts[2] || "";
  const individualJph = rowEl.querySelector("[data-config-jph-input]")?.value || "";

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getFinishConfigOwnerUsername(),
      shiftType: getFinishConfigSelectedShiftType(),
      operatorName,
      activeStatus: source?.activeStatus || "Active",
      defaultLine,
      defaultArea,
      defaultPosition,
      roleType: source?.roleType || "Unassigned",
      trainingWeek: source?.roleType === "Training" ? (source?.trainingWeek || "") : "",
      individualJph,
      originalDefaultArea: defaultArea,
      originalDefaultLine: defaultLine,
      originalDefaultPosition: defaultPosition
    });

    showFinishAssignmentToast(`Saved JPH for ${operatorName}`);
    await renderOperatorAssignmentConfigPanel();
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

function applyFinishAssignmentPermissionLock() {
  const canEdit = canEditFinishOperatorAssignments();
  const role = getFinishRosterCurrentProfile()?.role || getCurrentUserRole() || "Unknown role";
  const user = getCurrentUserDisplayName();

  document
    .querySelectorAll("#operatorAssignmentConfigPanel select, #operatorAssignmentConfigPanel input, #operatorAssignmentConfigPanel button")
    .forEach(el => {
      if (el.id === "assignmentStationFilter") {
        el.disabled = false;
        return;
      }
      el.disabled = !canEdit;
      el.classList.toggle("locked", !canEdit);
      el.title = canEdit ? "" : "View only for your role";
    });

  const notice = document.getElementById("finishAssignmentPermissionNotice");
  if (notice) {
    notice.classList.toggle("can-edit", canEdit);
    notice.classList.toggle("view-only", !canEdit);
    notice.innerHTML = canEdit
      ? `<strong>Edit access active</strong><span>${escapeHtml(user)} · ${escapeHtml(role)} · saving to Finish roster API for ${escapeHtml(getRosterOwnerUsername())} / ${escapeHtml(getRosterShiftType())}</span>`
      : `<strong>View only</strong><span>${escapeHtml(user)} · ${escapeHtml(role)} · edit/delete is blocked by the API.</span>`;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initFinishRosterApiPanel();
  loadFinishRosterBackend({ silent: true });
});


/*******************************************************
 * FINAL CONFIGURATION FIX — ASSIGNED ROSTER ONLY
 * Date: 2026-05-31
 *
 * Fixes requested:
 * 1) Weekend shift capacity hours = 10.5, not 11.
 * 2) Configuration shows a Weekday / Weekend dropdown.
 * 3) Configuration populates ONLY associates assigned to the logged-in profile.
 * 4) Associate list source = FINISH_OPERATOR_ROSTER_CONTROL.
 * 5) Operator Activity is used ONLY for live output totals.
 * 6) Drill JPH = 6.
 *******************************************************/

// Drill is 6 JPH. Keep this before the page initializes config totals.
try {
  if (typeof DEFAULT_CONFIG === "object") {
    DEFAULT_CONFIG.drillRate = 6;
  }
} catch (err) {
  console.warn("Could not update DEFAULT_CONFIG.drillRate:", err);
}

// Preserve the original config loader, but normalize old saved Drill=5 to Drill=6.
const __finishOriginalLoadConfig_AssignedRosterOnly = typeof loadConfig === "function" ? loadConfig : null;
if (__finishOriginalLoadConfig_AssignedRosterOnly) {
  loadConfig = function loadConfig() {
    const cfg = __finishOriginalLoadConfig_AssignedRosterOnly();

    // If this browser saved the previous default 5, migrate it to 6.
    if (Number(cfg.drillRate || 0) === 5 || Number(cfg.drillRate || 0) <= 0) {
      cfg.drillRate = 6;
    }

    // When Configuration has loaded an assigned roster, use it for capacity assignment math.
    if (Array.isArray(window.__finishConfigAssignedRosterRows)) {
      cfg.operatorAssignments = loadFinishPersonalOperatorAssignments(window.__finishConfigAssignedRosterRows);
    }

    return cfg;
  };
}

function getMorningSetupShiftHours() {
  return getMorningSetupShiftType() === "Weekend" ? 10.5 : 9.5;
}

function getFinishConfigOwnerUsername() {
  return String(getCurrentUsername() || "")
    .trim()
    .toUpperCase();
}

function getFinishConfigDefaultShiftType() {
  const owner = getFinishConfigOwnerUsername();
  const profile = (finishRosterApiState.profiles || []).find(item =>
    String(item.username || "").trim().toUpperCase() === owner
  );

  const shift = String(profile?.shiftGroup || "").trim().toLowerCase();
  return shift === "weekend" ? "Weekend" : "Weekday";
}

function getFinishConfigSelectedShiftType() {
  const select = document.getElementById("configAssignedShiftSelect");
  const value = String(select?.value || window.__finishConfigSelectedShiftType || getFinishConfigDefaultShiftType())
    .trim()
    .toLowerCase();

  return value === "weekend" ? "Weekend" : "Weekday";
}

function getFinishConfigAreaFilter() {
  return String(document.getElementById("configAssignedAreaFilter")?.value || "all").trim();
}

function getFinishConfigLiveOperatorTotal(operatorName, stationName) {
  const live = typeof getLiveOperatorForStation === "function"
    ? getLiveOperatorForStation(operatorName, stationName)
    : null;

  return Number(live?.total || 0);
}

function getFinishConfigLiveOperatorLabel(operatorName, stationName) {
  const live = typeof getLiveOperatorForStation === "function"
    ? getLiveOperatorForStation(operatorName, stationName)
    : null;

  if (!live) return "Assigned · No scans yet";
  return `Assigned · Live ${numberFmt(Number(live.total || 0))}`;
}

function getFinishConfigRoleTarget(row, config = loadConfig()) {
  const explicitJph = Number(row.individualJph || row.IndividualJPH || 0) || 0;
  if (explicitJph > 0) return explicitJph;

  const area = normalizeFinishOperatorStation(row.defaultArea || "");
  const role = String(row.roleType || "").trim().toLowerCase();
  const week = convertApiWeekToNumber(row.trainingWeek || "");

  if (role === "training") {
    if (area === "Mounting") return getTrainingRateForStationWeek("Mounting", week);
    if (area === "Final Inspection") return getTrainingRateForStationWeek("Final Inspection", week);
    return 0;
  }

  if (["certified", "tq", "core"].includes(role)) {
    if (area === "Mounting") return Number(config.mountRate || 0);
    if (area === "Final Inspection") return Number(config.finalRate || 0);
    if (area === "Drill") return Number(config.drillRate || 6);
  }

  return 0;
}

function ensureOperatorAssignmentConfigPanel() {
  let panel = document.getElementById("operatorAssignmentConfigPanel");

  if (panel) {
    return panel;
  }

  if (typeof injectOperatorAssignmentStyles === "function") {
    injectOperatorAssignmentStyles();
  }

  const configTab = document.querySelector('[data-content="config"]');
  const saveRow = document.querySelector(".config-actions") || document.getElementById("configSaveBtn")?.parentElement;
  const targetParent = saveRow?.parentElement || configTab;
  if (!targetParent) return null;

  panel = document.createElement("details");
  panel.id = "operatorAssignmentConfigPanel";
  panel.className = "config-group operator-assignment-config-group";
  panel.open = true;
  panel.innerHTML = `
    <summary>
      <div>
        <h3>My Active Associates</h3>
        <p>Shows associates assigned to your profile. Use Weekday / Weekend to view your assigned shift roster.</p>
      </div>
      <span>Open / Close</span>
    </summary>

    <div class="operator-assignment-shell">
      <div class="assignment-warning-card">
        <strong>Assigned Roster View</strong>
        <span id="configAssignedRosterStatus">Roster source: FINISH_OPERATOR_ROSTER_CONTROL. Output source: Operator Activity.</span>
      </div>

      <div id="finishAssignmentPermissionNotice" class="assignment-permission-notice can-edit"></div>

      <div class="assignment-summary-grid">
        <article><span>Shift</span><strong id="configAssignedShiftSummary">--</strong></article>
        <article><span>Active Associates</span><strong id="configAssignedActiveCount">0</strong></article>
        <article><span>Mounting</span><strong id="configAssignedMountingCount">0</strong></article>
        <article><span>Final / Drill</span><strong id="configAssignedFinalDrillCount">0 / 0</strong></article>
      </div>

      <div class="assignment-toolbar">
        <label>
          Shift
          <select id="configAssignedShiftSelect">
            <option value="Weekday">Weekday</option>
            <option value="Weekend">Weekend</option>
          </select>
        </label>

        <label>
          Area Filter
          <select id="configAssignedAreaFilter">
            <option value="all">All assigned areas</option>
            <option value="Mounting">Mounting</option>
            <option value="Final Inspection">Final Inspection</option>
            <option value="Drill">Drill</option>
          </select>
        </label>

        <button id="configAssignedRefreshBtn" type="button">Refresh Assigned Roster</button>
      </div>

      <div id="operatorAssignmentRoster" class="assignment-roster"></div>
    </div>
  `;

  if (saveRow && saveRow.parentElement) {
    saveRow.parentElement.insertBefore(panel, saveRow);
  } else {
    targetParent.appendChild(panel);
  }

  const shiftSelect = document.getElementById("configAssignedShiftSelect");
  if (shiftSelect) {
    shiftSelect.value = window.__finishConfigSelectedShiftType || getFinishConfigDefaultShiftType();
    shiftSelect.addEventListener("change", () => {
      window.__finishConfigSelectedShiftType = getFinishConfigSelectedShiftType();
      renderOperatorAssignmentConfigPanel();
    });
  }

  document.getElementById("configAssignedAreaFilter")?.addEventListener("change", renderOperatorAssignmentConfigPanel);
  document.getElementById("configAssignedRefreshBtn")?.addEventListener("click", renderOperatorAssignmentConfigPanel);

  return panel;
}

async function loadFinishConfigAssignedRosterRows() {
  const ownerUsername = getFinishConfigOwnerUsername();
  const shiftType = getFinishConfigSelectedShiftType();

  const payload = await fetchFinishRosterApi("getFinishRosterControl", {
    ownerUsername,
    shiftType,
    t: Date.now()
  });

  const rows = Array.isArray(payload.data) ? payload.data : [];

  return rows
    .filter(row => String(row.ownerUsername || "").trim().toUpperCase() === ownerUsername)
    .filter(row => String(row.shiftType || "").trim().toLowerCase() === shiftType.toLowerCase())
    .filter(row => String(row.activeStatus || "").trim().toLowerCase() === "active")
    .sort((a, b) => {
      const areaCompare = String(a.defaultArea || "").localeCompare(String(b.defaultArea || ""));
      if (areaCompare !== 0) return areaCompare;

      const lineCompare = String(a.defaultLine || "").localeCompare(String(b.defaultLine || ""));
      if (lineCompare !== 0) return lineCompare;

      return String(a.defaultPosition || "").localeCompare(String(b.defaultPosition || ""));
    });
}

async function renderOperatorAssignmentConfigPanel() {
  const panel = ensureOperatorAssignmentConfigPanel();
  const rosterTarget = document.getElementById("operatorAssignmentRoster");
  if (!panel || !rosterTarget) return;

  const ownerUsername = getFinishConfigOwnerUsername();
  const shiftType = getFinishConfigSelectedShiftType();
  const areaFilter = getFinishConfigAreaFilter();

  const shiftSelect = document.getElementById("configAssignedShiftSelect");
  if (shiftSelect && shiftSelect.value !== shiftType) {
    shiftSelect.value = shiftType;
  }

  rosterTarget.innerHTML = `<article class="assignment-empty">Loading assigned ${escapeHtml(shiftType)} roster...</article>`;

  try {
    const assignedRows = await loadFinishConfigAssignedRosterRows();
    window.__finishConfigAssignedRosterRows = assignedRows;

    const visibleRows = assignedRows.filter(row =>
      areaFilter === "all" || normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(areaFilter)
    );

    const mountCount = assignedRows.filter(row => normalizeFinishOperatorStation(row.defaultArea) === "Mounting").length;
    const finalCount = assignedRows.filter(row => normalizeFinishOperatorStation(row.defaultArea) === "Final Inspection").length;
    const drillCount = assignedRows.filter(row => normalizeFinishOperatorStation(row.defaultArea) === "Drill").length;

    setText("configAssignedShiftSummary", shiftType);
    setText("configAssignedActiveCount", assignedRows.length);
    setText("configAssignedMountingCount", mountCount);
    setText("configAssignedFinalDrillCount", `${finalCount} / ${drillCount}`);

    const status = document.getElementById("configAssignedRosterStatus");
    if (status) {
      status.textContent = `${getCurrentUserDisplayName()} · ${ownerUsername} · ${shiftType} assigned roster only`;
    }

    const notice = document.getElementById("finishAssignmentPermissionNotice");
    if (notice) {
      notice.classList.add("can-edit");
      notice.classList.remove("view-only");
      notice.innerHTML = `
        <strong>Assigned roster only</strong>
        <span>Associate list comes from saved roster assignment. Live totals come from Operator Activity.</span>
      `;
    }

    if (!visibleRows.length) {
      rosterTarget.innerHTML = `
        <article class="assignment-empty">
          No active associates assigned to ${escapeHtml(ownerUsername)} / ${escapeHtml(shiftType)}${areaFilter !== "all" ? " / " + escapeHtml(areaFilter) : ""}.
        </article>
      `;
      updateConfigTotals();
      return;
    }

    const config = loadConfig();

    rosterTarget.innerHTML = visibleRows.map(row => {
      const operatorName = row.operatorName || "";
      const area = normalizeFinishOperatorStation(row.defaultArea || "");
      const line = row.defaultLine || "No line";
      const position = row.defaultPosition || "No position";
      const role = row.roleType || "Unassigned";
      const week = row.trainingWeek ? ` · ${row.trainingWeek}` : "";
      const liveTotal = getFinishConfigLiveOperatorTotal(operatorName, area);
      const target = getFinishConfigRoleTarget(row, config);

      return `
        <article class="assignment-row" data-assigned-roster-row="${escapeHtml(operatorName)}" data-area="${escapeHtml(area)}" data-line="${escapeHtml(line)}" data-position="${escapeHtml(position)}">
          <div class="assignment-operator">
            <strong>${escapeHtml(operatorName)}</strong>
            <span>${escapeHtml(area)} · ${escapeHtml(line)} · ${escapeHtml(position)} · ${escapeHtml(getFinishConfigLiveOperatorLabel(operatorName, area))}</span>
          </div>

          <div class="assignment-field assignment-field--role">
            <label>Role</label>
            <strong class="assignment-role-badge" data-role="${escapeHtml(String(role).toLowerCase())}">${escapeHtml(role)}${escapeHtml(week)}</strong>
          </div>

          <div class="assignment-field assignment-field--output">
            <label>Today Output</label>
            <strong class="assignment-output-value ${liveTotal > 0 ? "has-output" : ""}">${numberFmt(liveTotal)}</strong>
          </div>

          <div class="assignment-target">
            <span>JPH</span>
            <input class="assignment-jph-input" data-config-jph-input type="number" min="0" step="0.1" value="${escapeHtml(row.individualJph || "")}" placeholder="${target > 0 ? numberFmt(target) : "Default"}" />
            <button class="assignment-jph-save" type="button" data-config-jph-save>Save</button>
          </div>
        </article>
      `;
    }).join("");

    rosterTarget.querySelectorAll("[data-config-jph-save]").forEach(button => {
      button.addEventListener("click", event => saveConfigAssignedIndividualJph(event.target.closest(".assignment-row")));
    });

    updateConfigTotals();
    renderFinishThreeLineCell();
  } catch (error) {
    console.error("Assigned roster render failed:", error);
    rosterTarget.innerHTML = `
      <article class="assignment-empty">
        Could not load assigned roster: ${escapeHtml(error.message || String(error))}
      </article>
    `;
  }
}


async function saveConfigAssignedIndividualJph(rowEl) {
  if (!rowEl) return;

  const operatorName = rowEl.dataset.assignedRosterRow || "";
  const source = (window.__finishConfigAssignedRosterRows || []).find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "") === normalizeAssignmentOperatorName(operatorName) &&
    normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(rowEl.dataset.area || rowEl.querySelector(".assignment-operator span")?.textContent?.split(" · ")?.[0] || "")
  );

  const areaLinePositionText = rowEl.querySelector(".assignment-operator span")?.textContent || "";
  const parts = areaLinePositionText.split(" · ").map(x => String(x || "").trim());

  const defaultArea = source?.defaultArea || parts[0] || "";
  const defaultLine = source?.defaultLine || parts[1] || "";
  const defaultPosition = source?.defaultPosition || parts[2] || "";
  const individualJph = rowEl.querySelector("[data-config-jph-input]")?.value || "";

  try {
    await fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getFinishConfigOwnerUsername(),
      shiftType: getFinishConfigSelectedShiftType(),
      operatorName,
      activeStatus: source?.activeStatus || "Active",
      defaultLine,
      defaultArea,
      defaultPosition,
      roleType: source?.roleType || "Unassigned",
      trainingWeek: source?.roleType === "Training" ? (source?.trainingWeek || "") : "",
      individualJph,
      originalDefaultArea: defaultArea,
      originalDefaultLine: defaultLine,
      originalDefaultPosition: defaultPosition
    });

    showFinishAssignmentToast(`Saved JPH for ${operatorName}`);
    await renderOperatorAssignmentConfigPanel();
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
  }
}

function applyFinishAssignmentPermissionLock() {
  // Configuration My Active Associates is a view-only assigned roster list.
  // LMS assignment controls live in LMS Control Center.
  const shift = document.getElementById("configAssignedShiftSelect");
  const area = document.getElementById("configAssignedAreaFilter");
  const refresh = document.getElementById("configAssignedRefreshBtn");
  if (shift) shift.disabled = false;
  if (area) area.disabled = false;
  if (refresh) refresh.disabled = false;
}

// Re-render Configuration when the tab is opened, using selected Weekday/Weekend.
document.addEventListener("click", event => {
  const tab = event.target.closest?.('.tab-btn[data-tab="config"]');
  if (!tab) return;

  window.setTimeout(() => {
    renderOperatorAssignmentConfigPanel();
  }, 75);
});

// If the page already initialized before this block loaded, refresh the panel safely.
window.refreshMyActiveAssociatesFromAssignedRoster = function () {
  return renderOperatorAssignmentConfigPanel();
};