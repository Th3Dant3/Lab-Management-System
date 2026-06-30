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


/*******************************************************
 * FINISH WEBPAGE HIDE CONTROLS
 * JS-only visual hide. Sheet/API data stays untouched.
 *******************************************************/
const FINISH_WEBPAGE_HIDE_OPERATORS = new Set([
  "CYNTHIA FORBES",
  "BRIAN HONICKER",
  "CALEB DAY",
  "MARIA MITCHELL",
  "OHIOUSER3",
  "OHIOUSER5",
  "SARAH MCCARTNEY",
  "ESTENFANIA MONTENEGRO",
  "CARLY WOOD",
  "PRINCESS HENRY"
]);

const FINISH_LMS_HIDE_AREAS = new Set([
  "UTILITY",
  "MEI LINE B",
  "MEI LINE C",
  "MEI EASY FIT",
  "MEI"
]);

function normalizeFinishHiddenOperatorName_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function shouldHideFinishOperatorOnWebpage(name) {
  const clean = normalizeFinishHiddenOperatorName_(name);

  return (
    FINISH_WEBPAGE_HIDE_OPERATORS.has(clean) ||
    clean === "MEI" ||
    clean.startsWith("MEI ") ||
    clean.startsWith("MEI0") ||
    clean.startsWith("MEI-")
  );
}

function getFinishOperatorNameFromAnyRow_(row) {
  if (!row || typeof row !== "object") return "";

  return (
    row.OperatorName ||
    row.operatorName ||
    row.Operator ||
    row.operator ||
    row.Associate ||
    row.associate ||
    row.Name ||
    row.name ||
    row.EmployeeName ||
    row.employeeName ||
    ""
  );
}

function filterHiddenFinishOperators(rows) {
  if (!Array.isArray(rows)) return [];

  return rows.filter(row => {
    const name = getFinishOperatorNameFromAnyRow_(row);
    return !shouldHideFinishOperatorOnWebpage(name);
  });
}

function normalizeFinishLmsArea_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function shouldHideFinishLmsArea_(value) {
  return FINISH_LMS_HIDE_AREAS.has(normalizeFinishLmsArea_(value));
}

function getFinishAreaFromAnyRow_(row) {
  if (!row || typeof row !== "object") return "";

  return (
    row.defaultArea ||
    row.DefaultArea ||
    row.Area ||
    row.area ||
    row.DefaultRole ||
    row.defaultRole ||
    row.Role ||
    row.role ||
    row.FlowStation ||
    row.flowStation ||
    row.Station ||
    row.station ||
    row.lastFlowStation ||
    row.LastFlowStation ||
    row.lastAccessPoint ||
    row.LastAccessPoint ||
    row.accessPoint ||
    row.AccessPoint ||
    ""
  );
}

function filterFinishLmsVisibleRows_(rows) {
  if (!Array.isArray(rows)) return [];

  return rows.filter(row => {
    const area = getFinishAreaFromAnyRow_(row);
    const operatorName = getFinishOperatorNameFromAnyRow_(row);

    return (
      !shouldHideFinishLmsArea_(area) &&
      !shouldHideFinishLmsArea_(operatorName) &&
      !shouldHideFinishOperatorOnWebpage(operatorName)
    );
  });
}

function filterFinishWebpageRows_(rows) {
  return filterFinishLmsVisibleRows_(filterHiddenFinishOperators(rows));
}


function normalizeFinishAssignmentStatusForUi_(row) {
  if (!row || typeof row !== "object") return "Unassigned";

  const activeStatus = String(row.activeStatus || row.ActiveStatus || "").trim();
  const roleType = String(row.roleType || row.RoleType || row.role || row.Role || "").trim();
  const defaultArea = String(row.defaultArea || row.DefaultArea || row.area || row.Area || "").trim();
  const defaultPosition = String(row.defaultPosition || row.DefaultPosition || row.position || row.Position || "").trim();
  const defaultLine = String(row.defaultLine || row.DefaultLine || row.line || row.Line || "").trim();
  const shiftType = String(row.shiftType || row.ShiftType || row.shift || row.Shift || "").trim();

  const removed = activeStatus.toLowerCase() === "removed" || activeStatus.toLowerCase() === "inactive";
  if (removed) return "Unassigned";

  const hasAssignment =
    row.isSaved === true ||
    row.saved === true ||
    !!row.assignmentKey ||
    !!row.AssignmentKey ||
    !!shiftType ||
    !!defaultArea ||
    !!defaultPosition ||
    !!defaultLine ||
    (roleType && roleType.toLowerCase() !== "unassigned") ||
    (activeStatus && activeStatus.toLowerCase() === "active");

  return hasAssignment ? "Active" : "Unassigned";
}

function normalizeFinishRoleStatusForUi_(row) {
  if (!row || typeof row !== "object") return "Unassigned";

  const roleType = String(row.roleType || row.RoleType || row.role || row.Role || "").trim();
  const activeStatus = String(row.activeStatus || row.ActiveStatus || "").trim();

  if (roleType) {
    const upper = roleType.toUpperCase();
    if (upper === "TQ") return "Certified";
    if (upper === "TRAINING") return "Training";
    if (upper === "CERTIFIED") return "Certified";
    if (upper !== "UNASSIGNED") return roleType;
  }

  if (normalizeFinishAssignmentStatusForUi_(row) === "Active") {
    return activeStatus && activeStatus.toLowerCase() !== "unassigned" ? activeStatus : "Certified";
  }

  return "Unassigned";
}

function isFinishRosterRowAssignedForUi_(row) {
  return normalizeFinishAssignmentStatusForUi_(row) === "Active";
}

function normalizeFinishRosterShiftDropdown_() {
  const shiftSelect = document.getElementById("finishRosterShiftSelect");
  if (!shiftSelect) return;

  Array.from(shiftSelect.options || []).forEach(option => {
    const text = String(option.textContent || "").trim().toLowerCase();
    const value = String(option.value || "").trim().toLowerCase();

    if (text === "all floor operators" || value === "all floor operators") {
      option.textContent = "All Shifts";
      option.value = "all";
    }
  });

  if (!Array.from(shiftSelect.options || []).some(option => String(option.value || "").toLowerCase() === "all")) {
    const option = document.createElement("option");
    option.value = "all";
    option.textContent = "All Shifts";
    shiftSelect.appendChild(option);
  }
}

function cleanFinishLmsDropdownOptions_() {
  const selectors = [
    "#finishRosterAreaFilter",
    "#finishRosterDetailArea",
    "[data-content='lms-control'] select",
    ".lms-control-center select",
    "#lmsControlCenter select"
  ];

  const selects = document.querySelectorAll(selectors.join(","));

  selects.forEach(select => {
    Array.from(select.options || []).forEach(option => {
      const text = option.textContent || option.value || "";
      const value = option.value || text;

      if (shouldHideFinishLmsArea_(text) || shouldHideFinishLmsArea_(value)) {
        option.remove();
      }
    });

    const selected = select.selectedOptions && select.selectedOptions[0];
    if (selected && (shouldHideFinishLmsArea_(selected.textContent) || shouldHideFinishLmsArea_(selected.value))) {
      select.selectedIndex = 0;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });
}

function initFinishLmsHideAreaObserver_() {
  normalizeFinishRosterShiftDropdown_();
  applyFinishEditingProfileVisibility();
  cleanFinishLmsDropdownOptions_();

  const target = document.querySelector("[data-content='lms-control']") || document.body;
  if (!target || target.__finishLmsHideAreaObserverAttached) return;

  target.__finishLmsHideAreaObserverAttached = true;

  const observer = new MutationObserver(() => {
    cleanFinishLmsDropdownOptions_();
  });

  observer.observe(target, {
    childList: true,
    subtree: true
  });
}



/* Login-sheet fallback overrides for testing/personal profile resolution.
   This is still local JS only. It does not save globally.
   Add more users here only if their login page is not saving Role/Username correctly. */
const FINISH_LOGIN_PROFILE_OVERRIDES = {
  /* LMS */
  "BLOPEZ": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },
  "BRIAN LOPEZ CABRERA": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },
  "LOPEZ CABRERA": { username: "BLOPEZ", role: "LMS", subRole: "Admin", firstName: "Brian", lastName: "Lopez Cabrera" },

  "JBOOMERSHINE": { username: "JBOOMERSHINE", role: "LMS", subRole: "Admin", firstName: "", lastName: "" },
  "BOOMERSHINE": { username: "JBOOMERSHINE", role: "LMS", subRole: "Admin", firstName: "", lastName: "Boomershine" },

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

/* =========================================================
   FINISH TAB VISIBILITY — USERNAME BASED
   Do NOT use role for these tabs. Role is too broad.
========================================================= */
const FINISH_CONFIG_DOWNTIME_VISIBLE_USERS = new Set([
  "BLOPEZ",
  "MLITTLE",
  "JBOOMERSHINE",
  "BKARR",
  "RTATE",
  "KMANACK",
  "BHECK",
  "BHONICKER",
  "NPOSTON",
  "BDADE"
]);

const FINISH_LMS_CONTROL_VISIBLE_USERS = new Set([
  "BLOPEZ",
  "JBOOMERSHINE"
]);

/* Morning Set Up is shared globally for the floor.
   Only these usernames can update positions. Everyone else can view. */
const FINISH_MORNING_SETUP_EDIT_USERS = new Set([
  "KMANACK",
  "BHONICKER",
  "NPOSTON",
  "BHECK",
  "BLOPEZ"
]);

const FINISH_SHARED_MORNING_SETUP_OWNER = "BLOPEZ";


const FINISH_GLOBAL_ROSTER_OWNER = "BLOPEZ";
const FINISH_PROFILE_SELECTOR_VISIBLE_USERS = new Set([
  "BLOPEZ",
  "JBOOMERSHINE"
]);

function canSeeFinishEditingProfileControl() {
  return FINISH_PROFILE_SELECTOR_VISIBLE_USERS.has(getFinishVisibilityUsername());
}

function getFinishGlobalRosterOwnerUsername() {
  return FINISH_GLOBAL_ROSTER_OWNER;
}

function applyFinishEditingProfileVisibility() {
  const canSee = canSeeFinishEditingProfileControl();
  const ownerSelect = document.getElementById("finishRosterOwnerSelect");
  if (!ownerSelect) return;

  const wrappers = [
    ownerSelect.closest(".setup-filter"),
    ownerSelect.closest(".filter-field"),
    ownerSelect.closest("label"),
    ownerSelect.parentElement
  ].filter(Boolean);

  const target = wrappers[0] || ownerSelect;

  target.style.display = canSee ? "" : "none";
  target.hidden = !canSee;
  target.setAttribute("aria-hidden", canSee ? "false" : "true");

  ownerSelect.disabled = !canSee || !canEditFinishOperatorAssignments();
}


function normalizeFinishVisibilityUsername(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function getFinishVisibilityUsername() {
  return normalizeFinishVisibilityUsername(getCurrentUsername());
}

function canViewFinishConfigAndDowntimeTabs() {
  return FINISH_CONFIG_DOWNTIME_VISIBLE_USERS.has(getFinishVisibilityUsername());
}

function canViewFinishLmsControlTab() {
  const username = getFinishVisibilityUsername();
  const role = String(getCurrentUserRole() || "").trim().toLowerCase();

  return (
    FINISH_LMS_CONTROL_VISIBLE_USERS.has(username) ||
    role === "lms" ||
    role.includes("director") ||
    role.includes("manager") ||
    role.includes("supervisor") ||
    role.includes("training")
  );
}

function canEditFinishMorningSetup() {
  return FINISH_MORNING_SETUP_EDIT_USERS.has(getFinishVisibilityUsername());
}

function findFinishTabButtons(tabKeys = [], labelMatches = []) {
  const keySet = new Set(tabKeys.map(v => String(v || "").toLowerCase()));
  const labels = labelMatches.map(v => String(v || "").toLowerCase());

  return Array.from(document.querySelectorAll(".tab-btn")).filter(btn => {
    const tab = String(btn.dataset.tab || "").toLowerCase();
    const text = String(btn.textContent || "").replace(/\s+/g, " ").trim().toLowerCase();
    return keySet.has(tab) || labels.some(label => text === label);
  });
}

function findFinishTabContents(tabKeys = [], labelMatches = []) {
  // IMPORTANT: content matching must be exact by data-content only.
  // Do not scan textContent here; Morning Set Up contains helper text/buttons
  // and broad text matching can hide the wrong panel, leaving a blank screen.
  const keySet = new Set(tabKeys.map(v => String(v || "").toLowerCase()));

  return Array.from(document.querySelectorAll(".tab-content")).filter(content => {
    const key = String(content.dataset.content || "").toLowerCase();
    return keySet.has(key);
  });
}

function forceMorningSetupTabVisible() {
  const morningBtn = document.querySelector('.tab-btn[data-tab="personal"]');
  const morningContent = document.querySelector('.tab-content[data-content="personal"]');

  [morningBtn, morningContent].forEach(el => {
    if (!el) return;
    el.hidden = false;
    el.style.display = "";
    el.removeAttribute("aria-hidden");
    el.dataset.restrictedTabHidden = "false";
  });
}

function setFinishRestrictedTabVisibility({ tabKeys = [], labelMatches = [], allowed = false }) {
  const buttons = findFinishTabButtons(tabKeys, labelMatches);
  const contents = findFinishTabContents(tabKeys, labelMatches);

  buttons.forEach(btn => {
    btn.style.display = allowed ? "" : "none";
    btn.hidden = !allowed;
    btn.setAttribute("aria-hidden", allowed ? "false" : "true");
    btn.dataset.restrictedTabHidden = allowed ? "false" : "true";
  });

  contents.forEach(content => {
    content.style.display = allowed ? "" : "none";
    content.hidden = !allowed;
    content.setAttribute("aria-hidden", allowed ? "false" : "true");
    content.dataset.restrictedTabHidden = allowed ? "false" : "true";
  });
}

function activateFinishFallbackTab() {
  const fallbackOrder = ["personal", "floor", "summary", "flow", "operator", "hourly"];

  for (const tabName of fallbackOrder) {
    const btn = document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
    const content = document.querySelector(`.tab-content[data-content="${tabName}"]`);
    if (!btn || !content || btn.hidden || content.hidden || btn.style.display === "none" || content.style.display === "none") continue;

    document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
    btn.classList.add("active");
    content.classList.add("active");

    if (tabName === "personal" && typeof renderFinishThreeLineCell === "function") {
      renderFinishThreeLineCell();
      if (typeof renderMorningSetupRosterSummary === "function") renderMorningSetupRosterSummary();
      setTimeout(applyFinishMorningSlotStatusDots_, 0);
    }
    return;
  }
}

function moveOffHiddenFinishTabIfNeeded() {
  const activeTab = document.querySelector(".tab-btn.active");
  const activeContent = document.querySelector(".tab-content.active");

  const activeTabHidden = activeTab && (activeTab.hidden || activeTab.style.display === "none" || activeTab.dataset.restrictedTabHidden === "true");
  const activeContentHidden = activeContent && (activeContent.hidden || activeContent.style.display === "none" || activeContent.dataset.restrictedTabHidden === "true");

  if (activeTabHidden || activeContentHidden) {
    activeTab?.classList.remove("active");
    activeContent?.classList.remove("active");
    activateFinishFallbackTab();
  }
}


/*******************************************************
 * HARD FIX — FINISH SETUP TAB VISIBILITY
 * Supervisors/Managers/LMS/Training must see Finish Setup.
 * Editing Profile selector remains BLOPEZ/JBOOMERSHINE only.
 *******************************************************/
function forceFinishSetupTabVisibilityForAllowedRoles_() {
  const allowed = canViewFinishLmsControlTab();
  const tab = document.querySelector('.tab-btn[data-tab="lms-control"]');
  const content = document.querySelector('.tab-content[data-content="lms-control"]');

  if (tab) {
    tab.textContent = "Finish Setup";
    tab.hidden = !allowed;
    tab.style.display = allowed ? "" : "none";
    tab.setAttribute("aria-hidden", allowed ? "false" : "true");
    tab.dataset.restrictedTabHidden = allowed ? "false" : "true";
  }

  if (content) {
    content.hidden = !allowed;
    content.style.display = allowed ? "" : "none";
    content.setAttribute("aria-hidden", allowed ? "false" : "true");
    content.dataset.restrictedTabHidden = allowed ? "false" : "true";
  }
}


function applyFinishRestrictedTabVisibility() {
  const canViewConfigDowntime = canViewFinishConfigAndDowntimeTabs();
  const canViewLmsControl = canViewFinishLmsControlTab();

  setFinishRestrictedTabVisibility({
    tabKeys: ["config", "configuration"],
    labelMatches: ["configuration"],
    allowed: canViewConfigDowntime
  });

  setFinishRestrictedTabVisibility({
    tabKeys: ["downtime", "associate-downtime", "associate-downtime-log", "scan-downtime", "inactivity", "inactivity-log"],
    labelMatches: ["associate downtime", "downtime log", "associate downtime log"],
    allowed: canViewConfigDowntime
  });

  setFinishRestrictedTabVisibility({
    tabKeys: ["lms-control", "lms", "lms-control-center"],
    labelMatches: ["lms control", "lms control center"],
    allowed: canViewLmsControl
  });

  // Morning Setup must always be visible. Edit/save is restricted separately.
  forceMorningSetupTabVisible();

  moveOffHiddenFinishTabIfNeeded();

  const activeTab = document.querySelector(".tab-btn.active");
  if (activeTab?.dataset?.tab === "personal") {
    const personalContent = document.querySelector('.tab-content[data-content="personal"]');
    personalContent?.classList.add("active");
    if (typeof renderFinishThreeLineCell === "function") {
      setTimeout(() => {
        renderFinishThreeLineCell();
        if (typeof renderMorningSetupRosterSummary === "function") renderMorningSetupRosterSummary();
      }, 0);
    }
  }

  forceFinishSetupTabVisibilityForAllowedRoles_();
}

// Keep the old function name so existing startup code still works.
function applyLmsControlTabVisibility() {
  applyFinishRestrictedTabVisibility();
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

function getFinishCapacityConfigKey() {
  return `${FINISH_CAPACITY_CONFIG_KEY}_${getCurrentUsername()}`;
}

function setConfigSaveState(message) {
  const el = document.getElementById("configSaveState");
  if (el) el.textContent = message || "Ready";
}

function updateConfigProfileDisplay() {
  setText("configActiveProfileName", getCurrentUserDisplayName());
  setConfigSaveState("Ready");
}

function showFinishProfileSavedToast(message = "Change saved") {
  const profile = getCurrentUserDisplayName();
  showFinishAssignmentToast(`${message} for ${profile}`);
}

function initConfigSideTabs() {
  const tabs = document.querySelectorAll(".config-side-tab[data-config-pane]");
  const panes = document.querySelectorAll(".config-pane[data-config-panel]");
  if (!tabs.length || !panes.length) return;

  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      const key = tab.dataset.configPane;
      tabs.forEach(btn => btn.classList.toggle("active", btn === tab));
      panes.forEach(pane => pane.classList.toggle("active", pane.dataset.configPanel === key));

      if (key === "associates") {
        renderOperatorAssignmentConfigPanel();
      }
    });
  });
}

function initConfigSteppers() {
  document.querySelectorAll("[data-step-target]").forEach(button => {
    button.addEventListener("click", () => {
      const input = document.getElementById(button.dataset.stepTarget || "");
      if (!input) return;

      const step = Number(button.dataset.step || 1);
      const current = Number(input.value || 0);
      const min = input.min !== "" ? Number(input.min) : -Infinity;
      input.value = Math.max(min, current + step);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      setConfigSaveState("Unsaved changes");
    });
  });
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

  // Re-apply after any late panels/tabs finish rendering.
  setTimeout(applyFinishRestrictedTabVisibility, 500);
  setTimeout(applyFinishRestrictedTabVisibility, 2000);
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
    const saved = localStorage.getItem(getFinishCapacityConfigKey());
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
  localStorage.setItem(getFinishCapacityConfigKey(), JSON.stringify({ ...safeConfig, savedAt: new Date().toISOString(), savedBy: getCurrentUsername() }));
}

function initConfigPanel() {
  updateConfigProfileDisplay();
  initConfigSideTabs();
  initConfigSteppers();
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
    document.getElementById(id)?.addEventListener("input", () => {
      updateConfigTotals();
      setConfigSaveState("Unsaved changes");
    });
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
    setConfigSaveState("Saved just now");
    showFinishProfileSavedToast("Change saved");
    loadDashboard({ showLoader: false });
  });

  document.getElementById("configResetBtn")?.addEventListener("click", () => {
    localStorage.removeItem(getFinishCapacityConfigKey());
    initConfigPanel();
    setConfigSaveState("Defaults restored");
    showFinishProfileSavedToast("Defaults restored");
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


/*******************************************************
 * FINISH OPERATOR COMMAND — FLOATER DISPLAY
 * If an operator works multiple Finish areas today, display role as Floater.
 * JPH target DOES NOT change. Target still comes from the current station/area.
 *******************************************************/
function getFinishOperatorMultiStationInfo_(operatorName) {
  const targetName = normalizeAssignmentOperatorName(operatorName || "").toLowerCase();
  const stations = [];
  let total = 0;

  if (!targetName) {
    return { isFloater: false, count: 0, stations: [], total: 0 };
  }

  Object.values(finishOperatorState?.stations || {}).forEach(station => {
    const stationName = normalizeFinishOperatorStation(station?.name || "");
    if (!stationName || !isFinishOperatorStationAllowed(stationName)) return;
    if (shouldHideFinishLmsArea_(stationName)) return;

    const matched = (station.operatorList || []).find(operator =>
      normalizeAssignmentOperatorName(operator?.name || "").toLowerCase() === targetName
    );

    if (!matched) return;

    const matchedTotal = Number(matched.total || 0) || 0;
    if (matchedTotal <= 0) return;

    stations.push(stationName);
    total += matchedTotal;
  });

  return {
    isFloater: stations.length > 1,
    count: stations.length,
    stations,
    total
  };
}

function getFinishOperatorDisplayRoleLabel_(operatorName, stationName) {
  const multi = getFinishOperatorMultiStationInfo_(operatorName);

  if (multi.isFloater) {
    return "Floater";
  }

  return getAssignmentRoleLabel(operatorName, stationName, loadConfig());
}

function getFinishOperatorAreaDisplayLabel_(operatorName, stationName) {
  const multi = getFinishOperatorMultiStationInfo_(operatorName);

  if (multi.isFloater) {
    return `Floater · ${multi.count} areas`;
  }

  return normalizeFinishOperatorStation(stationName || "") || "Finish";
}


function getAssignmentTargetLabel(operatorName, stationName, config = loadConfig()) {
  // IMPORTANT:
  // Multi-area operators display as Floater, but the base/JPH target remains station-specific.
  const base = getFinishOperatorBaseRate(stationName, operatorName);
  const role = getFinishOperatorDisplayRoleLabel_(operatorName, stationName);
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
      applyFinishRestrictedTabVisibility();

      if (button.hidden || button.style.display === "none" || button.dataset.restrictedTabHidden === "true") {
        moveOffHiddenFinishTabIfNeeded();
        return;
      }

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


/*******************************************************
 * FINISH OPERATOR PERFORMANCE — AR STYLE DASHBOARD
 * Reworks the Finish Hourly Operator Performance tab into the AR-style layout:
 * KPI strip, process flow cards, productivity overview, leaderboard, and loadout cards.
 *******************************************************/
const FINISH_AR_STYLE_STATION_ORDER = [
  "Finish Unbox",
  "Mounting",
  "Drill",
  "Bigs",
  "Sharps",
  "Final Inspection"
];

function getFinishArStyleStationConfig_(stationName) {
  const name = normalizeFinishOperatorStation(stationName);
  const config = loadConfig();

  const map = {
    "Finish Unbox": {
      label: "FINISH UNBOX",
      zone: "Feed",
      capacity: Number(config.unboxRate || 0) * Math.max(1, Number(config.unboxCount || 1)),
      assigned: 1,
      statusType: "watch"
    },
    "Mounting": {
      label: "MOUNTING",
      zone: "Mount / Assemble",
      capacity: Number(config.mountRate || 0) * Math.max(1, Number(config.mountCount || 1)),
      assigned: Math.max(1, Number(config.mountCount || 1)),
      statusType: "watch"
    },
    "Drill": {
      label: "DRILL",
      zone: "Specialty",
      capacity: Number(config.drillRate || 0) * Math.max(1, Number(config.drillCount || 1)),
      assigned: Math.max(1, Number(config.drillCount || 1)),
      statusType: "constraint"
    },
    "Bigs": {
      label: "BIGS",
      zone: "Specialty",
      capacity: 0,
      assigned: 0,
      statusType: "constraint"
    },
    "Sharps": {
      label: "SHARPS",
      zone: "Specialty",
      capacity: 0,
      assigned: 0,
      statusType: "constraint"
    },
    "Final Inspection": {
      label: "FINAL INSPECTION",
      zone: "QA",
      capacity: Number(config.finalRate || 0) * Math.max(1, Number(config.finalCount || 1)),
      assigned: Math.max(1, Number(config.finalCount || 1)),
      statusType: "watch"
    }
  };

  return map[name] || {
    label: name.toUpperCase(),
    zone: "Finish",
    capacity: getFinishOperatorStationTarget(name),
    assigned: 0,
    statusType: "watch"
  };
}

function getFinishArVisibleStations_(stations) {
  const stationMap = stations || {};
  const ordered = [];

  FINISH_AR_STYLE_STATION_ORDER.forEach(name => {
    if (stationMap[name]) ordered.push(stationMap[name]);
  });

  Object.values(stationMap || {})
    .filter(station => {
      const name = normalizeFinishOperatorStation(station.name || "");
      if (FINISH_AR_STYLE_STATION_ORDER.includes(name)) return false;
      if (!isFinishOperatorStationAllowed(name)) return false;
      if (shouldHideFinishLmsArea_(name)) return false;
      return true;
    })
    .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
    .forEach(station => ordered.push(station));

  return ordered;
}

function getFinishArStationStatus_(station) {
  const cfg = getFinishArStyleStationConfig_(station.name);
  const output = Number(station.total || 0);
  const capacity = Number(cfg.capacity || 0);
  const pct = capacity > 0 ? Math.round((output / capacity) * 100) : 0;

  if (capacity <= 0) {
    return { label: "CONSTRAINT", className: "constraint", pct: 0, gap: "No JPH" };
  }

  if (pct >= 100) {
    return { label: "ON TRACK", className: "ontrack", pct, gap: output - capacity };
  }

  if (pct >= 75) {
    return { label: "WATCH", className: "watch", pct, gap: output - capacity };
  }

  return { label: "BEHIND", className: "behind", pct, gap: output - capacity };
}

function buildFinishArOperatorLeaderboard_(stations) {
  const byName = {};

  Object.values(stations || {}).forEach(station => {
    const stationName = normalizeFinishOperatorStation(station.name || "");
    if (!isFinishOperatorStationAllowed(stationName)) return;
    if (shouldHideFinishLmsArea_(stationName)) return;

    (station.operatorList || []).forEach(operator => {
      const name = normalizeAssignmentOperatorName(operator.name || "");
      if (!name || shouldHideFinishOperatorOnWebpage(name)) return;

      if (!byName[name]) {
        byName[name] = {
          name,
          total: 0,
          stations: new Set(),
          bestHour: "",
          bestHourValue: 0,
          primaryStation: stationName
        };
      }

      byName[name].total += Number(operator.total || 0) || 0;
      byName[name].stations.add(stationName);

      Object.keys(operator.hourly || {}).forEach(hour => {
        const value = Number(operator.hourly[hour] || 0) || 0;
        if (value > byName[name].bestHourValue) {
          byName[name].bestHourValue = value;
          byName[name].bestHour = hour;
          byName[name].primaryStation = stationName;
        }
      });
    });
  });

  return Object.values(byName)
    .map(row => ({
      ...row,
      stations: Array.from(row.stations),
      roleLabel: row.stations.size > 1 ? "Floater" : getAssignmentRoleLabel(row.name, row.primaryStation, loadConfig())
    }))
    .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));
}


function installFinishArFloaterAreaListStyles_() {
  if (document.getElementById("finishArFloaterAreaListStyles")) return;

  const style = document.createElement("style");
  style.id = "finishArFloaterAreaListStyles";
  style.textContent = `
    .finish-ar-floater-area-list {
      margin-top: 10px;
      padding: 10px;
      border: 1px solid rgba(95,216,255,.26);
      border-radius: 12px;
      background: rgba(95,216,255,.055);
    }

    .finish-ar-floater-area-list > strong {
      display: block;
      color: #5fd8ff;
      font-family: "JetBrains Mono", monospace;
      font-size: 10px;
      font-weight: 950;
      letter-spacing: .10em;
      text-transform: uppercase;
      margin-bottom: 8px;
    }

    .finish-ar-floater-area-list > div {
      display: flex;
      flex-wrap: wrap;
      gap: 7px;
    }

    .finish-ar-floater-area-pill {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      padding: 6px 9px;
      border-radius: 999px;
      border: 1px solid rgba(95,216,255,.38);
      background: rgba(2,12,24,.72);
      color: #dff8ff;
      font-size: 11px;
      font-weight: 900;
    }

    .finish-ar-floater-area-pill b {
      color: #ffd166;
      font-size: 12px;
    }

    .finish-ar-floater-area-pill small {
      color: #8fa6c3;
      font-size: 10px;
      font-weight: 800;
    }
  `;
  document.head.appendChild(style);
}


function renderFinishArStyleOperatorPerformance_(stations, summary = {}) {
  installFinishArFloaterAreaListStyles_();

  const root =
    document.getElementById("operatorCommandPanel") ||
    document.querySelector('[data-content="operators"] .operator-command-panel') ||
    document.querySelector('[data-content="operators"]');

  if (!root) return false;

  const visibleStations = getFinishArVisibleStations_(stations);
  const leaderboard = buildFinishArOperatorLeaderboard_(stations);
  const totalOutput = visibleStations.reduce((sum, station) => sum + (Number(station.total || 0) || 0), 0);
  const activeOperators = leaderboard.length;
  const capacityToday = visibleStations.reduce((sum, station) => {
    const cfg = getFinishArStyleStationConfig_(station.name);
    return sum + (Number(cfg.capacity || 0) || 0);
  }, 0);
  const outputPct = capacityToday > 0 ? Math.round((totalOutput / capacityToday) * 100) : 0;
  const topStation = visibleStations.slice().sort((a, b) => Number(b.total || 0) - Number(a.total || 0))[0];
  const peakHour = summary.peakHour || buildOperatorSummaryFromStations(stations).peakHour || "--";
  const topOperator = leaderboard[0];
  const pressureStation = visibleStations
    .map(station => ({ station, status: getFinishArStationStatus_(station) }))
    .sort((a, b) => Number(a.status.gap || 0) - Number(b.status.gap || 0))[0];

  root.classList.add("finish-ar-style-command");

  root.innerHTML = `
    <section class="finish-ar-hero">
      <div>
        <div class="finish-ar-eyebrow">FINISH PRODUCTIVITY COMMAND CENTER</div>
        <h2>Finish Operator Performance</h2>
        <p>Professional associate productivity, station flow, capacity health, and hourly performance by role target.</p>
      </div>
      <div class="finish-ar-actions">
        <button id="operatorRefreshBtn" class="finish-ar-refresh" type="button">Refresh Metrics</button>
      </div>
    </section>

    <section class="finish-ar-kpi-grid">
      ${renderFinishArKpiCard_("Total Finish Output", numberFmt(totalOutput), "All visible Finish station activity today")}
      ${renderFinishArKpiCard_("Active Operators", numberFmt(activeOperators), "Operators with Finish activity")}
      ${renderFinishArKpiCard_("Capacity Today", numberFmt(capacityToday), "Configured station JPH capacity")}
      ${renderFinishArKpiCard_("Output vs Capacity", capacityToday > 0 ? outputPct + "%" : "N/A", capacityToday > 0 ? "Gap: " + numberFmt(totalOutput - capacityToday) : "No capacity target")}
      ${renderFinishArKpiCard_("Top Station", topStation?.name || "--", numberFmt(topStation?.total || 0) + " output today")}
      ${renderFinishArKpiCard_("Peak Hour", peakHour || "--", "Peak combined Finish demand")}
    </section>

    <section class="finish-ar-section">
      <header class="finish-ar-section-head">
        <div>
          <h3>Finish Process Flow</h3>
          <p>Station health, capacity pressure, and labor status from Finish Unbox through Final Inspection.</p>
        </div>
        <div class="finish-ar-legend">
          <span><i class="ontrack"></i>On Track</span>
          <span><i class="watch"></i>Watch</span>
          <span><i class="behind"></i>Behind</span>
          <span><i class="constraint"></i>Constraint</span>
        </div>
      </header>
      <div class="finish-ar-flow-grid">
        ${visibleStations.map(renderFinishArProcessCard_).join("")}
      </div>
    </section>

    <section class="finish-ar-two-col">
      <article class="finish-ar-panel">
        <header class="finish-ar-panel-head">
          <h3>Operator Productivity Overview</h3>
          <span>${numberFmt(activeOperators)} operators with output</span>
        </header>
        <div class="finish-ar-overview-body">
          <div class="finish-ar-ring">
            <strong>${capacityToday > 0 ? outputPct : 0}%</strong>
            <span>Avg target</span>
          </div>
          <div class="finish-ar-overview-lines">
            ${renderFinishArOverviewLine_("Full Target", leaderboard.filter(op => getFinishOperatorMultiStationInfo_(op.name).isFloater || true).slice(0, 0).length, "green")}
            ${renderFinishArOverviewLine_("Top Output", topOperator ? topOperator.name : "--", "amber")}
            ${renderFinishArOverviewLine_("Highest Pressure", pressureStation?.station?.name || "--", "red")}
          </div>
        </div>
        <div class="finish-ar-mini-panels">
          <div><span>Best Performer</span><strong>${escapeHtml(topOperator?.name || "--")}</strong><small>${numberFmt(topOperator?.total || 0)} total output</small></div>
          <div><span>Highest Pressure</span><strong>${escapeHtml(pressureStation?.station?.name || "--")}</strong><small>Gap: ${numberFmt(pressureStation?.status?.gap || 0)}</small></div>
        </div>
      </article>

      <article class="finish-ar-panel">
        <header class="finish-ar-panel-head">
          <h3>Operator Performance Leaderboard</h3>
          <button type="button" class="finish-ar-small-btn">View All</button>
        </header>
        <div class="finish-ar-leaderboard">
          ${leaderboard.slice(0, 6).map((op, index) => renderFinishArLeaderboardRow_(op, index)).join("")}
        </div>
        <div class="finish-ar-signal-grid">
          <div><span>Station Focus</span><strong>${escapeHtml(pressureStation?.station?.name || "--")}</strong></div>
          <div><span>Labor Signal</span><strong>${escapeHtml(topOperator ? topOperator.name + " leads output" : "--")}</strong></div>
          <div><span>Peak Demand</span><strong>${escapeHtml(peakHour || "--")}</strong></div>
        </div>
      </article>
    </section>

    <section class="finish-ar-section">
      <header class="finish-ar-section-head">
        <div>
          <h3>Operator Loadout Details</h3>
          <p>Click to open associate productivity cards. Multi-area operators show as Floater while JPH remains by area.</p>
        </div>
        <span class="finish-ar-sort">Sort by output</span>
      </header>
      <div class="finish-ar-loadout-grid">
        ${leaderboard.map(renderFinishArLoadoutCard_).join("")}
      </div>
    </section>
  `;

  root.querySelector("#operatorRefreshBtn")?.addEventListener("click", loadFinishOperatorActivity);

  root.querySelectorAll("[data-finish-ar-station]").forEach(card => {
    card.addEventListener("click", () => openOperatorDrawer(card.dataset.finishArStation));
  });

  root.querySelectorAll("[data-finish-ar-operator]").forEach(card => {
    card.addEventListener("click", () => {
      const stationName = card.dataset.finishArStation || card.dataset.primaryStation || "Mounting";
      const operatorName = card.dataset.finishArOperator;
      openExpandedOperatorViewer(stationName, operatorName);
    });
  });

  return true;
}

function renderFinishArKpiCard_(label, value, sub) {
  return `
    <article class="finish-ar-kpi-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
      <small>${escapeHtml(sub || "")}</small>
    </article>`;
}

function renderFinishArProcessCard_(station) {
  const cfg = getFinishArStyleStationConfig_(station.name);
  const status = getFinishArStationStatus_(station);
  const output = Number(station.total || 0) || 0;
  const capacity = Number(cfg.capacity || 0) || 0;
  const bar = capacity > 0 ? Math.min(100, Math.round((output / capacity) * 100)) : 100;

  return `
    <article class="finish-ar-process-card ${escapeHtml(status.className)}" data-finish-ar-station="${escapeHtml(station.name)}">
      <header>
        <h4>${escapeHtml(cfg.label)}</h4>
        <span>${escapeHtml(status.label)}</span>
      </header>
      <div class="finish-ar-process-metrics">
        <div><span>Output</span><strong>${numberFmt(output)}</strong></div>
        <div><span>Capacity</span><strong>${capacity > 0 ? numberFmt(capacity) : "N/A"}</strong></div>
      </div>
      <div class="finish-ar-bar"><i style="width:${bar}%;"></i></div>
      <footer>
        <div><span>Pace</span><strong>${capacity > 0 ? status.pct + "%" : "Process"}</strong></div>
        <div><span>Gap</span><strong>${capacity > 0 ? numberFmt(status.gap) : "No JPH"}</strong></div>
        <div><span>Assigned</span><strong>${numberFmt((station.operatorList || []).length)}</strong></div>
      </footer>
    </article>`;
}

function renderFinishArOverviewLine_(label, value, tone) {
  return `
    <div class="finish-ar-overview-line ${escapeHtml(tone)}">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(String(value))}</strong>
    </div>`;
}

function renderFinishArLeaderboardRow_(operator, index) {
  const primaryStation = operator.primaryStation || operator.stations?.[0] || "Mounting";
  const isFloater = operator.stations.length > 1;
  const target = getFinishOperatorBaseRate(primaryStation, operator.name);
  const targetPct = target > 0 ? Math.round((operator.bestHourValue / target) * 100) : 0;
  const status = targetPct >= 100 ? "Elite" : targetPct >= 90 ? "Watch" : "Needs Support";

  return `
    <button class="finish-ar-leader-row" type="button" data-finish-ar-operator="${escapeHtml(operator.name)}" data-primary-station="${escapeHtml(primaryStation)}">
      <b>${index + 1}</b>
      <span>
        <strong>${escapeHtml(operator.name)}</strong>
        <small>${escapeHtml(isFloater ? "Floater · " + operator.stations.join(" / ") : primaryStation)} · ${escapeHtml(operator.roleLabel)}</small>
      </span>
      <em>${numberFmt(operator.total)}<small>current output</small></em>
      <em>${target > 0 ? targetPct + "%" : "N/A"}<small>target %</small></em>
      <em>${escapeHtml(shortHour(operator.bestHour) || "--")}<small>peak hour</small></em>
      <i>${escapeHtml(status)}</i>
    </button>`;
}


function renderFinishArFloaterAreaList_(operator) {
  const stations = Array.isArray(operator?.stations) ? operator.stations : [];
  if (stations.length <= 1) return "";

  const items = stations.map(stationName => {
    const station = finishOperatorState.stations?.[stationName];
    const stationOperator = (station?.operatorList || []).find(item => item.name === operator.name);
    const output = Number(stationOperator?.total || 0) || 0;
    const target = getFinishOperatorBaseRate(stationName, operator.name);

    return `
      <span class="finish-ar-floater-area-pill">
        ${escapeHtml(stationName)}
        <b>${numberFmt(output)}</b>
        <small>${target > 0 ? numberFmt(target) + "/hr" : "No JPH"}</small>
      </span>`;
  }).join("");

  return `
    <div class="finish-ar-floater-area-list">
      <strong>Areas worked today</strong>
      <div>${items}</div>
    </div>`;
}


function renderFinishArLoadoutCard_(operator) {
  const primaryStation = operator.primaryStation || operator.stations?.[0] || "Mounting";
  const isFloater = operator.stations.length > 1;
  const initials = getFinishRosterCardInitials(operator.name);
  const station = finishOperatorState.stations?.[primaryStation];
  const stationOperator = (station?.operatorList || []).find(item => item.name === operator.name);
  const hours = stationOperator?.hourly || {};
  const best = getPeakHour(hours);
  const target = getFinishOperatorBaseRate(primaryStation, operator.name);
  const targetPct = target > 0 ? Math.round((best.value / target) * 100) : 0;

  return `
    <article class="finish-ar-loadout-card ${targetPct >= 100 ? "good" : targetPct >= 90 ? "watch" : "behind"}" data-finish-ar-operator="${escapeHtml(operator.name)}" data-finish-ar-station="${escapeHtml(primaryStation)}">
      <header>
        <span>${escapeHtml(initials)}</span>
        <div>
          <h4>${escapeHtml(operator.name)}</h4>
          <small>${escapeHtml(isFloater ? "Floater" : primaryStation)} · ${escapeHtml(operator.roleLabel)}</small>
        </div>
        <strong>${numberFmt(operator.total)}<small>output</small></strong>
      </header>
      ${isFloater ? `<div class="finish-ar-floater-strip">↔ Floater · ${numberFmt(operator.stations.length)} areas</div>` : ""}
      ${renderFinishArFloaterAreaList_(operator)}
      <div class="finish-ar-loadout-metrics">
        <div><span>Target</span><strong>${target > 0 ? targetPct + "%" : "N/A"}</strong></div>
        <div><span>Peak Hour</span><strong>${escapeHtml(shortHour(best.hour) || "--")}</strong></div>
        <div><span>Status</span><strong>${targetPct >= 100 ? "Elite" : targetPct >= 90 ? "Watch" : "Needs Support"}</strong></div>
      </div>
      <div class="finish-ar-hours-mini">
        ${getOperatorHourOrder(hours).map(hour => {
          const value = Number(hours[hour] || 0) || 0;
          const hourTarget = getOperatorHourTarget(primaryStation, hour, "operator", operator.name);
          const status = getOperatorPerformanceStatus(value, hourTarget);
          return `<i class="${escapeHtml(status.className)}"><span>${escapeHtml(shortHour(hour))}</span><b>${numberFmt(value)}</b></i>`;
        }).join("")}
      </div>
    </article>`;
}


function renderOperatorSummary(summary) {
  if (renderFinishArStyleOperatorPerformance_(finishOperatorState.stations || {}, summary || {})) {
    return;
  }

  setText("operatorTotalJobs", summary.totalJobs || 0);
  setText("operatorTotalOperators", summary.totalOperators || 0);
  setText("operatorTopStation", summary.topStation || "--");
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
    const multiInfo = getFinishOperatorMultiStationInfo_(operator.name);
    const floaterBadge = multiInfo.isFloater
      ? `<small style="display:inline-flex;margin-top:8px;margin-right:6px;padding:5px 9px;border-radius:999px;border:1px solid rgba(251,191,36,.65);color:#fbbf24;font-weight:900;letter-spacing:.06em;text-transform:uppercase;">Floater · ${numberFmt(multiInfo.count)} areas</small>`
      : "";

    return `
      <article class="operator-person-card perf-${escapeHtml(currentPerf.className)} operator-person-clickable" data-expanded-operator="${escapeHtml(operator.name)}" style="border-color:${currentPerf.color};box-shadow:0 0 18px ${currentPerf.color}1f;cursor:pointer;">
        <div class="operator-person-head">
          <div>
            <h4>${escapeHtml(operator.name)}</h4>
            <span>${escapeHtml(getFinishOperatorAreaDisplayLabel_(operator.name, station.name))} · ${escapeHtml(getAssignmentTargetLabel(operator.name, station.name))}</span>
            ${floaterBadge}
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


/*******************************************************
 * EXPANDED OPERATOR VIEW — FLOATER AREA BREAKDOWN
 * Shows every area worked in the expanded modal.
 *******************************************************/
function getFinishExpandedOperatorAreaEntries_(operatorName) {
  const cleanName = normalizeAssignmentOperatorName(operatorName || "").toLowerCase();
  const entries = [];

  if (!cleanName) return entries;

  Object.values(finishOperatorState?.stations || {}).forEach(station => {
    const stationName = normalizeFinishOperatorStation(station?.name || "");
    if (!stationName || !isFinishOperatorStationAllowed(stationName)) return;
    if (shouldHideFinishLmsArea_(stationName)) return;

    const operator = (station.operatorList || []).find(item =>
      normalizeAssignmentOperatorName(item?.name || "").toLowerCase() === cleanName
    );

    if (!operator) return;

    const output = Number(operator.total || 0) || 0;
    if (output <= 0) return;

    entries.push({
      station,
      stationName,
      operator,
      output,
      hours: operator.hourly || {},
      accessPoints: operator.accessPoints || []
    });
  });

  return entries.sort((a, b) => Number(b.output || 0) - Number(a.output || 0));
}

function getFinishExpandedCombinedHours_(entries) {
  const combined = {};

  (entries || []).forEach(entry => {
    Object.keys(entry.hours || {}).forEach(hour => {
      combined[hour] = (combined[hour] || 0) + (Number(entry.hours[hour] || 0) || 0);
    });
  });

  return combined;
}

function getFinishExpandedBestAreaHour_(entries) {
  let best = { stationName: "", hour: "", value: 0 };

  (entries || []).forEach(entry => {
    Object.keys(entry.hours || {}).forEach(hour => {
      const value = Number(entry.hours[hour] || 0) || 0;
      if (value > best.value) {
        best = { stationName: entry.stationName, hour, value };
      }
    });
  });

  return best;
}

function getFinishExpandedLastActiveHour_(entries) {
  const combined = getFinishExpandedCombinedHours_(entries);
  let last = "";

  getOperatorHourOrder(combined).forEach(hour => {
    if ((Number(combined[hour] || 0) || 0) > 0) last = hour;
  });

  return last;
}

function renderFinishExpandedAreaCard_(entry) {
  const stationName = entry.stationName;
  const operator = entry.operator;
  const hours = entry.hours || {};
  const orderedHours = getOperatorHourOrder(hours);
  const targetPerHour = getFinishOperatorBaseRate(stationName, operator.name);
  const role = getAssignmentRoleLabel(operator.name, stationName, loadConfig());
  const best = getPeakHour(hours);
  let lastActive = "";

  orderedHours.forEach(hour => {
    if ((Number(hours[hour] || 0) || 0) > 0) lastActive = hour;
  });

  return `
    <article class="expanded-floater-area-card">
      <header>
        <div>
          <h3>${escapeHtml(stationName)}</h3>
          <p>${escapeHtml(role)} · ${targetPerHour > 0 ? numberFmt(targetPerHour) + "/hr target" : "No JPH target"}</p>
          <small>Best hour: ${escapeHtml(shortHour(best.hour) || "--")} (${numberFmt(best.value)}) · Last active: ${escapeHtml(shortHour(lastActive) || "--")}</small>
        </div>
        <strong>${numberFmt(entry.output)}<span>output</span></strong>
      </header>

      <div class="expanded-floater-hour-grid">
        ${orderedHours.map(hour => {
          const value = Number(hours[hour] || 0) || 0;
          const target = getOperatorHourTarget(stationName, hour, "operator", operator.name);
          const status = getOperatorPerformanceStatus(value, target);
          const pctRaw = target > 0 ? Math.round((value / target) * 100) : 0;
          const barPct = target > 0 ? Math.min(100, Math.max(4, pctRaw)) : (value > 0 ? 18 : 4);

          return `
            <div class="expanded-floater-hour perf-${escapeHtml(status.className)}">
              <span>${escapeHtml(shortHour(hour))}</span>
              <strong style="color:${status.color};">${numberFmt(value)}</strong>
              <i style="height:${barPct}%;background:${status.color};box-shadow:0 0 14px ${status.color};"></i>
              <small>${target > 0 ? numberFmt(target) + " target" : "No target"}</small>
              <b style="color:${status.color};">${escapeHtml(status.label)}</b>
            </div>`;
        }).join("")}
      </div>
    </article>`;
}

function installFinishExpandedFloaterStyles_() {
  if (document.getElementById("finishExpandedFloaterStyles")) return;

  const style = document.createElement("style");
  style.id = "finishExpandedFloaterStyles";
  style.textContent = `
    .expanded-floater-area-stack {
      display: grid;
      gap: 16px;
      margin-top: 16px;
    }

    .expanded-floater-area-card {
      padding: 16px;
      border: 1px solid rgba(95,216,255,.22);
      border-radius: 18px;
      background: rgba(2,12,24,.62);
      box-shadow: inset 0 1px 0 rgba(255,255,255,.04);
    }

    .expanded-floater-area-card header {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
      padding-bottom: 12px;
      margin-bottom: 14px;
      border-bottom: 1px solid rgba(95,216,255,.16);
    }

    .expanded-floater-area-card h3 {
      margin: 0;
      color: #eef8ff;
      font-size: 21px;
      font-weight: 950;
    }

    .expanded-floater-area-card p {
      margin: 6px 0 0;
      color: #5fd8ff;
      font-size: 13px;
      font-weight: 900;
    }

    .expanded-floater-area-card small {
      color: #8fa6c3;
      font-size: 11px;
      font-weight: 800;
    }

    .expanded-floater-area-card header strong {
      color: #ffd166;
      font-size: 30px;
      font-weight: 950;
      text-align: right;
    }

    .expanded-floater-area-card header strong span {
      display: block;
      color: #8fa6c3;
      font-size: 10px;
      letter-spacing: .12em;
      text-transform: uppercase;
    }

    .expanded-floater-hour-grid {
      display: grid;
      grid-template-columns: repeat(15, minmax(54px, 1fr));
      gap: 8px;
      overflow-x: auto;
      padding-bottom: 2px;
    }

    .expanded-floater-hour {
      min-height: 128px;
      position: relative;
      padding: 8px 6px;
      border: 1px solid rgba(255,255,255,.08);
      border-radius: 12px;
      background: rgba(255,255,255,.045);
      display: grid;
      grid-template-rows: auto auto 1fr auto auto;
      gap: 5px;
      text-align: center;
      overflow: hidden;
    }

    .expanded-floater-hour span {
      color: #8fa6c3;
      font-size: 10px;
      font-weight: 900;
    }

    .expanded-floater-hour strong {
      font-size: 16px;
      font-weight: 950;
    }

    .expanded-floater-hour i {
      align-self: end;
      width: 74%;
      justify-self: center;
      min-height: 4px;
      border-radius: 9px 9px 3px 3px;
      opacity: .92;
    }

    .expanded-floater-hour small {
      color: #8fa6c3;
      font-size: 9px;
      font-weight: 800;
      line-height: 1.1;
    }

    .expanded-floater-hour b {
      font-size: 10px;
      font-weight: 950;
    }

    .expanded-floater-summary-note {
      margin-top: 14px;
      padding: 14px 16px;
      border: 1px solid rgba(95,216,255,.20);
      border-radius: 14px;
      background: rgba(95,216,255,.055);
      color: #bfdaf2;
      font-size: 13px;
      line-height: 1.5;
    }
  `;
  document.head.appendChild(style);
}


function openExpandedOperatorViewer(stationName, operatorName) {
  ensureExpandedOperatorViewer();
  installFinishExpandedFloaterStyles_();

  const station = finishOperatorState.stations?.[stationName];
  const operator = (station?.operatorList || []).find(item => item.name === operatorName);
  const viewer = document.getElementById("expandedOperatorViewer");
  const body = document.getElementById("expandedOperatorBody");

  if (!station || !operator || !viewer || !body) return;

  const entries = getFinishExpandedOperatorAreaEntries_(operator.name);
  const safeEntries = entries.length ? entries : [{
    station,
    stationName: station.name,
    operator,
    output: Number(operator.total || 0) || 0,
    hours: operator.hourly || {},
    accessPoints: operator.accessPoints || []
  }];

  const isFloater = safeEntries.length > 1;
  const totalOutput = safeEntries.reduce((sum, entry) => sum + (Number(entry.output || 0) || 0), 0);
  const combinedHours = getFinishExpandedCombinedHours_(safeEntries);
  const orderedCombinedHours = getOperatorHourOrder(combinedHours);
  const activeHours = orderedCombinedHours.filter(hour => Number(combinedHours[hour] || 0) > 0).length;
  const avgPerActiveHour = activeHours > 0 ? Math.round(totalOutput / activeHours) : 0;
  const best = getFinishExpandedBestAreaHour_(safeEntries);
  const lastActive = getFinishExpandedLastActiveHour_(safeEntries);
  const roleLabel = isFloater ? `Floater · ${safeEntries.length} areas` : getAssignmentRoleLabel(operator.name, station.name, loadConfig());

  setText("expandedOperatorName", operator.name);
  setText(
    "expandedOperatorSub",
    `${roleLabel} • JPH by area • Current station: ${station.name}`
  );

  body.innerHTML = `
    <section class="expanded-operator-scoreboard">
      <article class="expanded-score-card">
        <span>Role</span>
        <strong>${escapeHtml(isFloater ? "Floater" : getAssignmentRoleLabel(operator.name, station.name, loadConfig()))}</strong>
        <small>${isFloater ? numberFmt(safeEntries.length) + " areas worked today" : escapeHtml(station.name)}</small>
      </article>
      <article class="expanded-score-card">
        <span>Final JPH</span>
        <strong>By area</strong>
        <small>Floater does not change the station target</small>
      </article>
      <article class="expanded-score-card">
        <span>Total Output</span>
        <strong>${numberFmt(totalOutput)}</strong>
        <small>Across visible Finish areas</small>
      </article>
      <article class="expanded-score-card">
        <span>Best Hour</span>
        <strong>${escapeHtml(best.stationName || "--")} · ${escapeHtml(shortHour(best.hour) || "--")}</strong>
        <small>${numberFmt(best.value)} jobs</small>
      </article>
      <article class="expanded-score-card">
        <span>Last Active</span>
        <strong>${escapeHtml(shortHour(lastActive) || "--")}</strong>
        <small>${numberFmt(activeHours)} active hours · ${numberFmt(avgPerActiveHour)} avg/hr</small>
      </article>
    </section>

    <section class="expanded-floater-area-stack">
      ${safeEntries.map(renderFinishExpandedAreaCard_).join("")}
    </section>

    <div class="expanded-floater-summary-note">
      Floater is display-only when an operator works multiple Finish areas. Each area keeps its own configured JPH target, so the hourly colors and target lines remain accurate by station.
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


/*******************************************************
 * MORNING SET UP — SLOT STATUS DOTS
 * Green = activity this hour / active recently
 * Yellow = watch
 * Red = no activity / inactive
 *******************************************************/
function getFinishMorningSlotStatus_(assignment) {
  if (!assignment || !assignment.operator) {
    return { key: "empty", label: "Open", dot: "empty" };
  }

  const total = Number(assignment.total || 0) || 0;
  const thisHour =
    Number(
      assignment.thisHour ??
      assignment.currentHour ??
      assignment.hourOutput ??
      assignment.currentHourOutput ??
      0
    ) || 0;

  const statusText = String(
    assignment.status ||
    assignment.activityStatus ||
    assignment.gapStatus ||
    assignment.liveStatus ||
    ""
  ).trim().toLowerCase();

  if (
    statusText.includes("inactive") ||
    statusText.includes("no activity") ||
    statusText.includes("no scan") ||
    statusText.includes("down")
  ) {
    return { key: "inactive", label: "No activity", dot: "red" };
  }

  if (
    statusText.includes("watch") ||
    statusText.includes("warning") ||
    statusText.includes("idle") ||
    statusText.includes("gap")
  ) {
    return { key: "watch", label: "Watch", dot: "yellow" };
  }

  if (thisHour > 0 || total > 0 || assignment.liveNow) {
    return { key: "active", label: "Active", dot: "green" };
  }

  return { key: "inactive", label: "No activity", dot: "red" };
}

function applyFinishMorningSlotStatusDots_() {
  const slots = document.querySelectorAll(".finish-slot[data-operator], .finish-slot.assigned, .finish-slot.empty");

  slots.forEach(slot => {
    const operator = String(slot.dataset.operator || "").trim();
    const station = normalizeFinishOperatorStation(slot.dataset.station || "");
    const line = String(slot.dataset.line || "").trim();
    const position = String(slot.dataset.slot || "").trim();

    let assignment = null;

    if (operator) {
      assignment = findFinishMorningSlotAssignment_(operator, station, line, position);
    }

    const status = getFinishMorningSlotStatus_(assignment);

    slot.classList.remove(
      "morning-slot-dot-active",
      "morning-slot-dot-watch",
      "morning-slot-dot-inactive",
      "morning-slot-dot-empty"
    );

    slot.classList.add(`morning-slot-dot-${status.dot}`);

    // Use the existing visual dot if the card already has one.
    // Do NOT append a second dot near the output number.
    const dot =
      slot.querySelector(".slot-status-dot") ||
      slot.querySelector(".morning-slot-live-dot");

    if (dot) {
      dot.classList.add("morning-slot-live-dot");
      dot.setAttribute("title", status.label);
      dot.setAttribute("aria-label", status.label);
    }
  });
}

function findFinishMorningSlotAssignment_(operatorName, stationName, line, position) {
  const cleanOperator = normalizeAssignmentOperatorName(operatorName || "").toLowerCase();
  const cleanStation = normalizeFinishOperatorStation(stationName || "");
  const cleanLine = String(line || "").trim().toLowerCase();
  const cleanPosition = String(position || "").trim().toLowerCase();

  const roster = [
    ...(Array.isArray(finishRosterApiState?.morningRoster) ? finishRosterApiState.morningRoster : []),
    ...(Array.isArray(window.__finishConfigAssignedRosterRows) ? window.__finishConfigAssignedRosterRows : [])
  ];

  const saved = roster.find(row => {
    const rowName = normalizeAssignmentOperatorName(row.operatorName || row.OperatorName || "").toLowerCase();
    const rowStation = normalizeFinishOperatorStation(row.defaultArea || row.DefaultArea || row.area || row.Area || "");
    const rowLine = String(row.defaultLine || row.DefaultLine || row.line || row.Line || "").trim().toLowerCase();
    const rowPosition = String(row.defaultPosition || row.DefaultPosition || row.position || row.Position || "").trim().toLowerCase();

    return (
      rowName === cleanOperator &&
      (!cleanStation || rowStation === cleanStation) &&
      (!cleanLine || rowLine === cleanLine) &&
      (!cleanPosition || rowPosition === cleanPosition)
    );
  });

  const liveStation = finishOperatorState?.stations?.[cleanStation];
  const liveOperator = (liveStation?.operatorList || []).find(item =>
    normalizeAssignmentOperatorName(item.name || "").toLowerCase() === cleanOperator
  );

  return {
    ...(saved || {}),
    operator: operatorName,
    total: Number(liveOperator?.total || saved?.total || 0) || 0,
    liveNow: !!liveOperator,
    thisHour: getFinishLiveOperatorCurrentHourOutput_(liveOperator)
  };
}

function getFinishLiveOperatorCurrentHourOutput_(operator) {
  if (!operator || !operator.hourly) return 0;

  const hours = operator.hourly || {};
  const ordered = getOperatorHourOrder(hours);
  let latestValue = 0;

  ordered.forEach(hour => {
    const value = Number(hours[hour] || 0) || 0;
    if (value > 0) latestValue = value;
  });

  return latestValue;
}

function installFinishMorningSlotStatusDotStyles_() {
  if (document.getElementById("finishMorningSlotStatusDotStyles")) return;

  const style = document.createElement("style");
  style.id = "finishMorningSlotStatusDotStyles";
  style.textContent = `
    .finish-slot {
      position: relative;
    }

    /* Use the existing slot dot. Keep it away from the output number. */
    .finish-slot .slot-status-dot,
    .finish-slot .morning-slot-live-dot {
      position: absolute !important;
      top: 9px !important;
      right: 9px !important;
      width: 8px !important;
      height: 8px !important;
      min-width: 8px !important;
      min-height: 8px !important;
      border-radius: 50% !important;
      z-index: 3 !important;
      pointer-events: none !important;
    }

    /* Give the output number breathing room so it does not touch the status dot. */
    .finish-slot [class*="output"],
    .finish-slot [class*="total"],
    .finish-slot .slot-total,
    .finish-slot .slot-output,
    .finish-slot .slot-count {
      padding-right: 18px !important;
    }

    .morning-slot-dot-green .slot-status-dot,
    .morning-slot-dot-green .morning-slot-live-dot {
      background: #4ade80 !important;
      box-shadow: 0 0 0 3px rgba(74,222,128,.14), 0 0 16px rgba(74,222,128,.85) !important;
    }

    .morning-slot-dot-yellow .slot-status-dot,
    .morning-slot-dot-yellow .morning-slot-live-dot {
      background: #facc15 !important;
      box-shadow: 0 0 0 3px rgba(250,204,21,.14), 0 0 16px rgba(250,204,21,.85) !important;
    }

    .morning-slot-dot-red .slot-status-dot,
    .morning-slot-dot-red .morning-slot-live-dot {
      background: #fb7185 !important;
      box-shadow: 0 0 0 3px rgba(251,113,133,.14), 0 0 16px rgba(251,113,133,.85) !important;
    }

    .morning-slot-dot-empty .slot-status-dot,
    .morning-slot-dot-empty .morning-slot-live-dot {
      background: #64748b !important;
      box-shadow: 0 0 0 3px rgba(100,116,139,.14), 0 0 10px rgba(100,116,139,.45) !important;
    }

    .finish-slot.morning-slot-dot-green {
      border-color: rgba(74,222,128,.74) !important;
    }

    .finish-slot.morning-slot-dot-yellow {
      border-color: rgba(250,204,21,.74) !important;
    }

    .finish-slot.morning-slot-dot-red {
      border-color: rgba(251,113,133,.74) !important;
    }
  `;
  document.head.appendChild(style);
}


function initFinishThreeLineCell() {
  installFinishMorningSlotStatusDotStyles_();
  installFinishLineSlotDelegatedClickHandler();
  renderFinishThreeLineCell();
  setTimeout(applyFinishMorningSlotStatusDots_, 0);
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


function getSelectedFinishMorningLine() {
  const selected = finishRosterUiState.selectedMorningLine || FINISH_LINE_NAMES[0];
  return FINISH_LINE_NAMES.includes(selected) ? selected : FINISH_LINE_NAMES[0];
}

function renderFinishLineSelector(assignmentMap = {}) {
  return `
    <aside class="finish-line-selector" aria-label="Finish line selector">
      <span class="line-selector-kicker">Production Cell</span>
      ${FINISH_LINE_NAMES.map(line => {
        const lineAssignments = Object.values(assignmentMap || {}).filter(a => a.line === line);
        const mount = lineAssignments.filter(a => a.station === "Mounting").length;
        const final = lineAssignments.filter(a => a.station === "Final Inspection").length;
        const liveTotal = lineAssignments.reduce((sum, a) => sum + Number(a.total || 0), 0);
        const active = getSelectedFinishMorningLine() === line;
        return `
          <button type="button" class="finish-line-selector-btn ${active ? "active" : ""}" data-finish-line-select="${escapeHtml(line)}">
            <span>${escapeHtml(line)}</span>
            <strong>${numberFmt(mount)}/10 Mount · ${numberFmt(final)}/4 Final</strong>
            <em>${numberFmt(liveTotal)} live output</em>
          </button>`;
      }).join("")}
    </aside>`;
}

function renderFinishThreeLineCell() {
  const grid = document.getElementById("finishThreeLineGrid");
  if (!grid) return;

  const config = loadConfig();
  const assignmentMap = getAssignedLineSlots(config, finishRosterApiState.morningRoster || []);
  const statusMap = loadFinishLineSlotStatus();
  const selectedLine = getSelectedFinishMorningLine();

  renderMorningSetupRosterSummary();
  grid.innerHTML = `
    ${renderFinishLineSelector(assignmentMap)}
    <div class="finish-line-stage">
      ${renderOneFinishLine(selectedLine, assignmentMap, statusMap)}
    </div>`;
  bindFinishLineCellEvents();

  setTimeout(applyFinishMorningSlotStatusDots_, 0);
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
  const hasAssigned = !!stats.assignedCount;
  const mainText = hasAssigned ? `${numberFmt(stats.shiftCapacity)}` : "No Target";
  const unitText = hasAssigned ? "jobs" : "";
  const jphText = hasAssigned ? `${numberFmt(stats.totalJph)} JPH` : "Assign roster";
  const avgText = hasAssigned ? `AVG ${numberFmt(stats.avgJph)}` : "No avg";
  const opText = hasAssigned ? `${numberFmt(stats.assignedCount)} OP` : "";

  return `
    <div class="line-jph-metric perf-${escapeHtml(stats.className)}" style="--metric-color:${escapeHtml(stats.color)};">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(mainText)}${unitText ? `<em>${escapeHtml(unitText)}</em>` : ""}</strong>
      <small>
        <b>${escapeHtml(jphText)}</b>
        <b>${escapeHtml(avgText)}</b>
        ${opText ? `<b>${escapeHtml(opText)}</b>` : ""}
      </small>
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
  const isCollapsed = false;

  return `
    <article class="finish-line-cell is-expanded" data-finish-line="${escapeHtml(line)}">
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
  document.querySelectorAll("[data-finish-line-select]").forEach(button => {
    button.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      const line = button.dataset.finishLineSelect;
      if (!FINISH_LINE_NAMES.includes(line)) return;
      finishRosterUiState.selectedMorningLine = line;
      renderFinishThreeLineCell();
    });
  });


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


function getRawMorningActiveRosterRows() {
  return (finishRosterApiState.morningRoster || [])
    .filter(row => String(row.activeStatus || "").toLowerCase() === "active")
    .filter(row => normalizeAssignmentOperatorName(row.operatorName || ""));
}

function getMorningActiveRosterRows() {
  return dedupeMorningRosterRowsByOperator(getRawMorningActiveRosterRows());
}

function getMorningRosterRowScore(row, targetStation = "", currentOperator = "") {
  let score = 0;
  const rowStation = normalizeFinishOperatorStation(row.defaultArea || "");
  const target = normalizeFinishOperatorStation(targetStation || "");
  const rowName = normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase();
  const current = normalizeAssignmentOperatorName(currentOperator || "").toLowerCase();

  if (current && rowName === current) score += 10000;
  if (target && rowStation === target) score += 1000;
  if (row.defaultLine) score += 120;
  if (row.defaultPosition) score += 120;
  if (String(row.roleType || "").toLowerCase() !== "unassigned") score += 40;
  if (Number(row.individualJph || row.IndividualJPH || 0) > 0) score += 20;
  if (String(row.activeStatus || "").toLowerCase() === "active") score += 10;
  return score;
}

function dedupeMorningRosterRowsByOperator(rows = [], targetStation = "", currentOperator = "") {
  const best = new Map();

  (Array.isArray(rows) ? rows : []).forEach(row => {
    const operatorName = normalizeAssignmentOperatorName(row.operatorName || "");
    if (!operatorName) return;

    const key = operatorName.toLowerCase();
    const current = best.get(key);

    if (!current || getMorningRosterRowScore(row, targetStation, currentOperator) >= getMorningRosterRowScore(current, targetStation, currentOperator)) {
      best.set(key, row);
    }
  });

  return Array.from(best.values());
}

function getMorningAssignableRosterRows(targetStation = "", currentOperator = "") {
  return dedupeMorningRosterRowsByOperator(getRawMorningActiveRosterRows(), targetStation, currentOperator);
}

function findMorningRosterRow(operatorName) {
  const target = normalizeAssignmentOperatorName(operatorName).toLowerCase();
  const rows = getRawMorningActiveRosterRows().filter(row =>
    normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === target
  );
  return dedupeMorningRosterRowsByOperator(rows)[0] || null;
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


function getMorningSlotRoleLabel(row) {
  const role = row.roleType === "Training"
    ? `${row.roleType} ${row.trainingWeek || ""}`.trim()
    : (row.roleType || "Unassigned");
  return role;
}

function getMorningSlotPlacementLabel(row) {
  return [row.defaultArea, row.defaultLine, row.defaultPosition].filter(Boolean).join(" · ") || "No assigned position";
}

function getMorningSlotJphLabel(row) {
  const explicit = Number(row.individualJph || row.IndividualJPH || 0) || 0;
  if (explicit > 0) return `${numberFmt(explicit)} JPH`;

  const target = getFinishConfigRoleTarget({
    defaultArea: row.defaultArea || "",
    roleType: row.roleType || "",
    trainingWeek: row.trainingWeek || "",
    individualJph: ""
  }, loadConfig());

  return target > 0 ? `${numberFmt(target)} JPH` : "Default JPH";
}

function sortMorningSlotRowsForTarget(rows, station, currentOperator = "") {
  const targetStation = normalizeFinishOperatorStation(station);
  const currentName = normalizeAssignmentOperatorName(currentOperator || "").toLowerCase();

  return [...rows].sort((a, b) => {
    const aName = normalizeAssignmentOperatorName(a.operatorName || "").toLowerCase();
    const bName = normalizeAssignmentOperatorName(b.operatorName || "").toLowerCase();

    const score = row => {
      let value = 0;
      const rowName = normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase();
      const rowStation = normalizeFinishOperatorStation(row.defaultArea || "");
      if (rowName === currentName) value -= 1000;
      if (rowStation === targetStation) value -= 200;
      if (!row.defaultLine && !row.defaultPosition) value -= 40;
      if (String(row.roleType || "").toLowerCase() === "training") value += 10;
      return value;
    };

    const scoreDiff = score(a) - score(b);
    if (scoreDiff !== 0) return scoreDiff;
    return aName.localeCompare(bName);
  });
}

function applyMorningSlotChoiceFilter(modal) {
  const search = String(modal.querySelector("[data-morning-slot-search]")?.value || "").toLowerCase().trim();
  const mode = modal.querySelector("[data-morning-slot-filter].active")?.dataset.morningSlotFilter || "best";
  const targetStation = normalizeFinishOperatorStation(modal.dataset.station || "");

  let visibleCount = 0;
  modal.querySelectorAll("[data-morning-slot-operator]").forEach(choice => {
    const haystack = String(choice.dataset.search || "").toLowerCase();
    const rowStation = normalizeFinishOperatorStation(choice.dataset.station || "");
    const hasPosition = String(choice.dataset.hasPosition || "") === "true";
    const matchesSearch = !search || haystack.includes(search);
    const matchesMode =
      mode === "all" ||
      (mode === "best" && rowStation === targetStation) ||
      (mode === "open" && !hasPosition);

    const show = matchesSearch && matchesMode;
    choice.hidden = !show;
    if (show) visibleCount++;
  });

  const empty = modal.querySelector("[data-morning-slot-filter-empty]");
  if (empty) {
    empty.hidden = visibleCount > 0;
  }
}

function bindMorningSlotModalFilters(modal) {
  modal.querySelector("[data-morning-slot-search]")?.addEventListener("input", () => applyMorningSlotChoiceFilter(modal));

  modal.querySelectorAll("[data-morning-slot-filter]").forEach(button => {
    button.addEventListener("click", () => {
      modal.querySelectorAll("[data-morning-slot-filter]").forEach(btn => btn.classList.remove("active"));
      button.classList.add("active");
      applyMorningSlotChoiceFilter(modal);
    });
  });

  applyMorningSlotChoiceFilter(modal);
}

function openMorningLineSlotAssignment(line, station, slot, currentOperator = "") {
  const modal = ensureMorningSlotModal();
  const title = modal.querySelector("#morningSlotModalTitle");
  const sub = modal.querySelector("#morningSlotModalSub");
  const body = modal.querySelector("#morningSlotModalBody");
  const rows = sortMorningSlotRowsForTarget(getMorningAssignableRosterRows(station, currentOperator), station, currentOperator);
  const canEdit = canEditFinishMorningSetup();
  const stationName = normalizeFinishOperatorStation(station);
  const currentRow = currentOperator ? findMorningRosterRow(currentOperator) : null;

  modal.dataset.line = line;
  modal.dataset.station = stationName;
  modal.dataset.slot = slot;

  if (title) title.textContent = `${line} · ${slot}`;
  if (sub) {
    sub.textContent = `${stationName} · ${getMorningSetupShiftType()} · ${getCurrentUsername()}`;
  }

  if (!rows.length) {
    body.innerHTML = `<article class="morning-slot-empty">No active associates are assigned to your ${escapeHtml(getMorningSetupShiftType())} roster yet. LMS must add them first.</article>`;
  } else {
    body.innerHTML = `
      <div class="morning-slot-command">
        <div class="morning-slot-current">
          <span>Current Assignment</span>
          <strong>${escapeHtml(currentOperator || "Open position")}</strong>
          ${currentRow ? `<small>${escapeHtml(getMorningSlotRoleLabel(currentRow))} · ${escapeHtml(getMorningSlotJphLabel(currentRow))}</small>` : `<small>Ready for assignment</small>`}
        </div>
        <div class="morning-slot-target">
          <span>Target Slot</span>
          <strong>${escapeHtml(line)} / ${escapeHtml(slot)}</strong>
          <small>${escapeHtml(stationName)}</small>
        </div>
      </div>

      <div class="morning-slot-toolbar">
        <input type="search" data-morning-slot-search placeholder="Search associate, role, line, or position..." />
        <div class="morning-slot-filter-pills">
          <button type="button" class="active" data-morning-slot-filter="best">Best match</button>
          <button type="button" data-morning-slot-filter="open">Open roster</button>
          <button type="button" data-morning-slot-filter="all">All active</button>
        </div>
      </div>

      <div class="morning-slot-list">
        ${rows.map(row => {
          const name = normalizeAssignmentOperatorName(row.operatorName || "");
          const isCurrent = name === currentOperator;
          const rowStation = normalizeFinishOperatorStation(row.defaultArea || "");
          const placement = getMorningSlotPlacementLabel(row);
          const role = getMorningSlotRoleLabel(row);
          const jph = getMorningSlotJphLabel(row);
          const liveTotal = getFinishConfigLiveOperatorTotal(name, rowStation || stationName);
          const hasPosition = !!(row.defaultLine || row.defaultPosition);
          const searchText = [name, rowStation, row.defaultLine, row.defaultPosition, role, jph].join(" ");
          return `
            <button
              type="button"
              class="morning-slot-choice ${isCurrent ? "is-current" : ""} ${rowStation === stationName ? "is-match" : ""}"
              data-morning-slot-operator="${escapeHtml(name)}"
              data-station="${escapeHtml(rowStation)}"
              data-has-position="${hasPosition ? "true" : "false"}"
              data-search="${escapeHtml(searchText)}"
              ${canEdit ? "" : "disabled"}
            >
              <div>
                <strong>${escapeHtml(name)}</strong>
                <span>${escapeHtml(placement)} · ${escapeHtml(role)}</span>
              </div>
              <aside>
                <small>${escapeHtml(jph)}</small>
                <em>${numberFmt(liveTotal)}</em>
              </aside>
            </button>`;
        }).join("")}
        <article class="morning-slot-empty" data-morning-slot-filter-empty hidden>No associates match this filter.</article>
      </div>

      <div class="morning-slot-actions">
        ${currentOperator ? `<button type="button" class="morning-slot-secondary" data-morning-slot-view="${escapeHtml(currentOperator)}">View Output</button>` : ""}
        ${currentOperator && canEdit ? `<button type="button" class="morning-slot-danger" data-morning-slot-clear="${escapeHtml(currentOperator)}">Clear Current Assignment</button>` : ""}
      </div>`;

    body.querySelectorAll("[data-morning-slot-operator]").forEach(button => {
      button.addEventListener("click", () => assignMorningOperatorToSlot(button.dataset.morningSlotOperator, line, stationName, slot, currentOperator));
    });

    body.querySelector("[data-morning-slot-view]")?.addEventListener("click", event => {
      const op = event.currentTarget.dataset.morningSlotView;
      closeMorningLineSlotAssignment();
      if (finishOperatorState?.stations?.[stationName]) {
        openExpandedOperatorViewer(stationName, op);
      }
    });

    body.querySelector("[data-morning-slot-clear]")?.addEventListener("click", event => {
      clearMorningOperatorPosition(event.currentTarget.dataset.morningSlotClear, stationName, line, slot);
    });

    bindMorningSlotModalFilters(modal);
  }

  modal.classList.add("open");
}

function closeMorningLineSlotAssignment() {
  document.getElementById("morningSlotAssignModal")?.classList.remove("open");
}


function patchMorningRosterRowInMemory(operatorName, updates = {}, matcher = null) {
  const target = normalizeAssignmentOperatorName(operatorName || "").toLowerCase();
  let changed = false;

  finishRosterApiState.morningRoster = (finishRosterApiState.morningRoster || []).map(row => {
    const isOperator = normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === target;
    const passesMatcher = typeof matcher === "function" ? matcher(row) : true;
    if (!isOperator || !passesMatcher) return row;
    changed = true;
    return { ...row, ...updates };
  });

  finishRosterApiState.roster = (finishRosterApiState.roster || []).map(row => {
    const isOperator = normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === target;
    const passesMatcher = typeof matcher === "function" ? matcher(row) : true;
    if (!isOperator || !passesMatcher) return row;
    return { ...row, ...updates };
  });

  return changed;
}

function refreshMorningSetupUiFast() {
  closeMorningLineSlotAssignment();
  renderFinishThreeLineCell();
  updateConfigTotals();
  if (typeof renderOperatorAssignmentConfigPanel === "function") {
    renderOperatorAssignmentConfigPanel();
  }
}

async function assignMorningOperatorToSlot(operatorName, line, station, slot, currentOperatorToReplace = "") {
  if (!canEditFinishMorningSetup()) {
    showFinishAssignmentToast("View only. You can see Morning Set Up, but only approved users can update it.");
    return;
  }

  const existing = findMorningRosterRow(operatorName);
  if (!existing) {
    showFinishAssignmentToast("This associate is not active on your selected shift roster.");
    return;
  }

  const stationName = normalizeFinishOperatorStation(station);
  const previousOperator = normalizeAssignmentOperatorName(currentOperatorToReplace || "");
  const nextOperator = normalizeAssignmentOperatorName(operatorName || "");
  const previousRows = previousOperator && previousOperator !== nextOperator
    ? findMorningRosterRowsForSlot(previousOperator, stationName, line, slot)
    : [];

  const slotMatcher = row =>
    normalizeFinishOperatorStation(row.defaultArea || "") === stationName &&
    String(row.defaultLine || "").trim() === String(line || "").trim() &&
    String(row.defaultPosition || "").trim() === String(slot || "").trim();

  // Optimistic UI: clear/redraw immediately instead of waiting for Google Apps Script round trips.
  if (previousRows.length) {
    patchMorningRosterRowInMemory(previousOperator, {
      defaultLine: "",
      defaultArea: stationName,
      defaultPosition: ""
    }, slotMatcher);
  }

  patchMorningRosterRowInMemory(operatorName, {
    defaultLine: line,
    defaultArea: stationName,
    defaultPosition: slot,
    roleType: existing.roleType || "Unassigned",
    trainingWeek: existing.roleType === "Training" ? (existing.trainingWeek || "") : "",
    individualJph: existing.individualJph || ""
  });

  showFinishAssignmentToast(`Assigning ${operatorName} to ${line} / ${slot}...`);
  refreshMorningSetupUiFast();

  try {
    const requests = [];

    previousRows.forEach(previous => {
      requests.push(fetchFinishRosterApi("saveFinishRosterControl", {
        updatedBy: getCurrentUsername(),
        ownerUsername: getMorningSetupOwnerUsername(),
        shiftType: getMorningSetupShiftType(),
        operatorName: previousOperator,
        activeStatus: "Active",
        defaultLine: "",
        defaultArea: stationName,
        defaultPosition: "",
        roleType: previous.roleType || "Unassigned",
        trainingWeek: previous.roleType === "Training" ? (previous.trainingWeek || "") : "",
        individualJph: previous.individualJph || "",
        originalDefaultArea: previous.defaultArea || stationName,
        originalDefaultLine: previous.defaultLine || line,
        originalDefaultPosition: previous.defaultPosition || slot
      }));
    });

    requests.push(fetchFinishRosterApi("saveFinishRosterControl", {
      updatedBy: getCurrentUsername(),
      ownerUsername: getMorningSetupOwnerUsername(),
      shiftType: getMorningSetupShiftType(),
      operatorName,
      activeStatus: "Active",
      defaultLine: line,
      defaultArea: stationName,
      defaultPosition: slot,
      roleType: existing.roleType || "Unassigned",
      trainingWeek: existing.roleType === "Training" ? (existing.trainingWeek || "") : "",
      individualJph: existing.individualJph || "",
      originalDefaultArea: existing.defaultArea || "",
      originalDefaultLine: existing.defaultLine || "",
      originalDefaultPosition: existing.defaultPosition || ""
    }));

    await Promise.all(requests);

    showFinishAssignmentToast(`Assigned ${operatorName} to ${line} / ${slot}`);
    loadMorningSetupRoster({ silent: true }).then(() => renderFinishThreeLineCell()).catch(() => {});
    loadFinishRosterForSelectedProfile({ silent: true }).catch(() => {});
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();
  }
}

async function clearMorningOperatorPosition(operatorName, station, line = "", slot = "") {
  if (!canEditFinishMorningSetup()) {
    showFinishAssignmentToast("View only. You can see Morning Set Up, but only approved users can update it.");
    return;
  }

  const stationName = normalizeFinishOperatorStation(station || "");
  const matches = findMorningRosterRowsForSlot(operatorName, stationName, line, slot);

  if (!matches.length) {
    showFinishAssignmentToast("No active roster row found for that associate/slot.");
    return;
  }

  const slotMatcher = row =>
    normalizeFinishOperatorStation(row.defaultArea || "") === stationName &&
    String(row.defaultLine || "").trim() === String(line || "").trim() &&
    String(row.defaultPosition || "").trim() === String(slot || "").trim();

  // Optimistic UI: remove from station instantly, then save to Google Sheet in the background.
  patchMorningRosterRowInMemory(operatorName, {
    defaultLine: "",
    defaultArea: stationName,
    defaultPosition: ""
  }, slotMatcher);

  showFinishAssignmentToast(`Clearing ${operatorName} from ${line || stationName} ${slot || ""}...`);
  refreshMorningSetupUiFast();

  try {
    await Promise.all(matches.map(existing =>
      fetchFinishRosterApi("saveFinishRosterControl", {
        updatedBy: getCurrentUsername(),
        ownerUsername: getMorningSetupOwnerUsername(),
        shiftType: getMorningSetupShiftType(),
        operatorName,
        activeStatus: "Active",
        defaultLine: "",
        defaultArea: stationName || normalizeFinishOperatorStation(existing.defaultArea || ""),
        defaultPosition: "",
        roleType: existing.roleType || "Unassigned",
        trainingWeek: existing.roleType === "Training" ? (existing.trainingWeek || "") : "",
        individualJph: existing.individualJph || "",
        originalDefaultArea: existing.defaultArea || stationName,
        originalDefaultLine: existing.defaultLine || line || "",
        originalDefaultPosition: existing.defaultPosition || slot || ""
      })
    ));

    showFinishAssignmentToast(`Cleared position for ${operatorName}`);
    loadMorningSetupRoster({ silent: true }).then(() => renderFinishThreeLineCell()).catch(() => {});
    loadFinishRosterForSelectedProfile({ silent: true }).catch(() => {});
  } catch (error) {
    console.error(error);
    showFinishAssignmentToast(error.message || String(error));
    await loadMorningSetupRoster({ silent: true }).catch(() => {});
    renderFinishThreeLineCell();
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

const FINISH_ROSTER_ASSIGNABLE_AREAS = [
  "Finish Unbox",
  "Mounting",
  "Drill",
  "Bigs",
  "Sharps",
  "Final Inspection",
  "Floater"
];

function isFinishRosterAssignableArea(value) {
  const station = normalizeFinishOperatorStation(value);
  return FINISH_ROSTER_ASSIGNABLE_AREAS.includes(station);
}

function buildFinishRosterAreaOptions(selectedArea, includeAll = false) {
  const selected = normalizeFinishOperatorStation(selectedArea || "");
  const visibleAreas = FINISH_ROSTER_ASSIGNABLE_AREAS.filter(area => !shouldHideFinishLmsArea_(area));
  const options = includeAll ? ['all', ...visibleAreas] : visibleAreas;
  return options.map(area => {
    const value = area === 'all' ? 'all' : area;
    const label = area === 'all' ? 'All Floor Operators' : area;
    const isSelected = area === 'all' ? String(selectedArea || '').toLowerCase() === 'all' : selected === area;
    return `<option value="${escapeHtml(value)}" ${isSelected ? 'selected' : ''}>${escapeHtml(label)}</option>`;
  }).join('');
}

const finishRosterApiState = {
  profiles: [],
  operators: [],
  roster: [],
  audit: [],
  ownerUsername: "",
  shiftType: "Weekday",

  // Morning Set Up is intentionally separate from LMS Control Center.
  // It is shared globally by shift so every user sees the same floor setup.
  morningOwnerUsername: "",
  morningShiftType: "Weekday",
  morningRoster: [],

  loading: false,
  ready: false,
  lastError: ""
};

const finishRosterUiState = {
  expandedRosterRows: new Set(),
  collapsedLines: new Set(),
  selectedRosterKey: "",
  draftAssignment: null
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
  // Shared Finish Setup roster.
  // Supervisors/Managers/LMS can edit, but all saves go to BLOPEZ so everyone sees the same roster.
  return getFinishGlobalRosterOwnerUsername();
}

function getRosterShiftType() {
  const raw =
    document.getElementById("finishRosterShiftSelect")?.value ||
    finishRosterApiState.shiftType ||
    "Weekday";

  const value = String(raw || "").trim().toLowerCase();

  if (value === "all" || value === "all shifts" || value === "all floor operators") {
    return "all";
  }

  return value === "weekend" ? "Weekend" : "Weekday";
}

function getMorningSetupOwnerUsername() {
  // Morning Set Up reads the same shared roster built in Finish Setup.
  return getFinishGlobalRosterOwnerUsername();
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

    finishRosterApiState.morningRoster = filterFinishWebpageRows_(rosterPayload.data || [])
      .filter(row => String(row.ownerUsername || "").toUpperCase() === ownerUsername)
      .filter(row => String(row.shiftType || "").toLowerCase() === shiftType.toLowerCase())
      .filter(row => String(row.activeStatus || "").toLowerCase() !== "removed");

    console.log("[Finish Morning Setup] Loaded shared roster", {
      ownerUsername,
      shiftType,
      rows: finishRosterApiState.morningRoster.length
    });

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


function normalizeFinishRosterOperatorList(rows) {
  const byName = new Map();
  (Array.isArray(rows) ? rows : []).forEach(row => {
    const operatorName = normalizeAssignmentOperatorName(
      row.operatorName || row.OperatorName || row.operator || row.Operator || row.name || row.Name || ''
    );
    if (!operatorName || operatorName.toLowerCase().includes('unassigned')) return;

    const existing = byName.get(operatorName.toLowerCase()) || {
      operatorName,
      area: 'Finish',
      lastFlowStation: '',
      lastAccessPoint: '',
      lastTotal: 0,
      isActive: true,
      source: 'Finish roster fallback'
    };

    const station = normalizeFinishOperatorStation(
      row.lastFlowStation || row.LastFlowStation || row.flowStation || row.FlowStation || row.accessPoint || row.AccessPoint || ''
    );

    const total = Number(row.lastTotal || row.LastTotal || row.total || row.Total || row.hourlyTotal || row.HourlyTotal || 0) || 0;

    existing.lastTotal += total;
    if (station && station !== 'Unmapped') existing.lastFlowStation = station;
    existing.lastAccessPoint = row.lastAccessPoint || row.LastAccessPoint || row.accessPoint || row.AccessPoint || existing.lastAccessPoint || '';
    existing.updatedAt = row.updatedAt || row.UpdatedAt || row.generatedAt || existing.updatedAt || '';

    byName.set(operatorName.toLowerCase(), existing);
  });

  return Array.from(byName.values()).sort((a, b) => a.operatorName.localeCompare(b.operatorName));
}

async function loadFinishOperatorMasterWithFallback() {
  const masterPayload = await fetchFinishRosterApi('getFinishOperatorMaster', {
    activeOnly: 'false',
    debug: 'true',
    t: Date.now()
  });

  let operators = normalizeFinishRosterOperatorList(masterPayload.data || masterPayload.operators || []);

  if (operators.length) return filterFinishWebpageRows_(operators);

  // Fallback: use today's live Finish operator activity so the dropdown never goes blank.
  const activityPayload = await fetchFinishRosterApi('operatorActivity', {
    area: 'Finish',
    debug: 'true',
    t: Date.now()
  });

  operators = normalizeFinishRosterOperatorList(
    activityPayload.operatorActivity || activityPayload.data || activityPayload.rows || []
  );

  return filterFinishWebpageRows_(operators);
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

    const [profilesPayload, operators] = await Promise.all([
      fetchFinishRosterApi("getFinishUserProfiles", { debug: "true", t: Date.now() }),
      loadFinishOperatorMasterWithFallback()
    ]);

    finishRosterApiState.profiles = profilesPayload.data || [];
    finishRosterApiState.operators = filterFinishWebpageRows_(operators);

    if (!finishRosterApiState.profiles.length) {
      finishRosterApiState.profiles = [{
        username: getCurrentUsername(),
        fullName: getCurrentUserDisplayName(),
        role: getCurrentUserRole() || "LMS",
        shiftGroup: "Weekday",
        canEdit: canEditFinishOperatorAssignments(),
        isActive: true
      }];
    }

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

  let rosterRows = [];
  let auditPayload = { data: [] };

  if (String(shiftType).toLowerCase() === "all") {
    const [weekdayPayload, weekendPayload, auditResult] = await Promise.all([
      fetchFinishRosterApi("getFinishRosterControl", { ownerUsername, shiftType: "Weekday" }),
      fetchFinishRosterApi("getFinishRosterControl", { ownerUsername, shiftType: "Weekend" }),
      fetchFinishRosterApi("getFinishRosterAuditLog", { targetUsername: ownerUsername, limit: 75 })
    ]);

    rosterRows = [
      ...(weekdayPayload.data || []),
      ...(weekendPayload.data || [])
    ];

    auditPayload = auditResult || { data: [] };
  } else {
    const [rosterPayload, auditResult] = await Promise.all([
      fetchFinishRosterApi("getFinishRosterControl", { ownerUsername, shiftType }),
      fetchFinishRosterApi("getFinishRosterAuditLog", { targetUsername: ownerUsername, limit: 75 })
    ]);

    rosterRows = rosterPayload.data || [];
    auditPayload = auditResult || { data: [] };
  }

  finishRosterApiState.roster = filterFinishWebpageRows_(rosterRows);
  finishRosterApiState.audit = filterHiddenFinishOperators(auditPayload.data || []);

  if (!options.silent) {
    renderFinishRosterApiPanel();
    renderOperatorAssignmentConfigPanel();
    updateConfigTotals();
    renderFinishThreeLineCell();
  }
}


/*******************************************************
 * RENAME LMS CONTROL CENTER → FINISH SETUP
 *******************************************************/
function applyFinishSetupLabelRename_() {
  const renameExact = (selector, replacement) => {
    document.querySelectorAll(selector).forEach(el => {
      const current = String(el.textContent || "").trim().replace(/\s+/g, " ");
      if (current === "LMS Control Center" || current === "LMS CONTROL CENTER") {
        el.textContent = replacement;
      }
    });
  };

  renameExact('.tab-btn[data-tab="lms-control"]', "Finish Setup");
  renameExact('[data-content="lms-control"] .lms-control-title', "Finish Setup");
  renameExact('[data-content="lms-control"] h2', "Finish Setup");
  renameExact('[data-content="lms-control"] h3', "Finish Setup");

  document.querySelectorAll('[data-content="lms-control"], .lms-control-center, #finishRosterApiPanel').forEach(root => {
    root.querySelectorAll("*").forEach(el => {
      if (!el.childNodes || el.childNodes.length !== 1 || el.childNodes[0].nodeType !== Node.TEXT_NODE) return;

      const text = String(el.textContent || "").trim();
      if (text === "LMS Control Center" || text === "LMS CONTROL CENTER") {
        el.textContent = "Finish Setup";
      }
      if (text === "Finish Roster Control") {
        el.textContent = "Finish Setup";
      }
    });
  });

  const tab = document.querySelector('.tab-btn[data-tab="lms-control"]');
  if (tab) tab.textContent = "Finish Setup";
}


function initFinishRosterApiPanel() {
  const ownerSelect = document.getElementById("finishRosterOwnerSelect");
  const shiftSelect = document.getElementById("finishRosterShiftSelect");
  const roleFilter = document.getElementById("finishRosterRoleFilter");
  const areaFilter = document.getElementById("finishRosterAreaFilter");
  const searchInput = document.getElementById("finishRosterSearchInput");
  const refreshBtn = document.getElementById("finishRosterRefreshBtn");
  const addBtn = document.getElementById("finishRosterAddOperatorBtn");

  ownerSelect?.addEventListener("change", async event => {
    // Profile selector is visible only to BLOPEZ/JBOOMERSHINE.
    // Roster save/read stays global so every Supervisor/Manager/LMS sees the same shift setup.
    finishRosterApiState.ownerUsername = getFinishGlobalRosterOwnerUsername();
    event.target.value = finishRosterApiState.ownerUsername;
    finishRosterUiState.selectedRosterKey = "";
    finishRosterUiState.draftAssignment = null;
    await loadFinishRosterForSelectedProfile();
  });

  shiftSelect?.addEventListener("change", async event => {
    const value = String(event.target.value || "Weekday").trim().toLowerCase();
    finishRosterApiState.shiftType = value === "all" || value === "all shifts" || value === "all floor operators"
      ? "all"
      : (value === "weekend" ? "Weekend" : "Weekday");
    finishRosterUiState.selectedRosterKey = "";
    finishRosterUiState.draftAssignment = null;
    await loadFinishRosterForSelectedProfile();
  });

  if (areaFilter) {
    areaFilter.innerHTML = buildFinishRosterAreaOptions(areaFilter.value || 'all', true);
  }

  roleFilter?.addEventListener("change", renderFinishRosterApiPanel);
  areaFilter?.addEventListener("change", renderFinishRosterApiPanel);
  searchInput?.addEventListener("input", renderFinishRosterApiPanel);
  refreshBtn?.addEventListener("click", () => loadFinishRosterBackend());
  addBtn?.addEventListener("click", () => openNewFinishRosterAssignment());
}



/*******************************************************
 * ROLE FILTER FIX — NO TQ
 * Role filter must only show:
 * All Floor Operators, Unassigned, Certified, Training
 *******************************************************/
function buildFinishRosterRoleFilterOptions_(selectedValue = "all") {
  const selected = String(selectedValue || "all").trim();
  const options = [
    { value: "all", label: "All Floor Operators" },
    { value: "Unassigned", label: "Unassigned" },
    { value: "Certified", label: "Certified" },
    { value: "Training", label: "Training" }
  ];

  return options.map(option => {
    const isSelected = selected.toLowerCase() === option.value.toLowerCase();
    return `<option value="${escapeHtml(option.value)}" ${isSelected ? "selected" : ""}>${escapeHtml(option.label)}</option>`;
  }).join("");
}

function fixFinishRosterRoleFilterNoTq_() {
  const roleFilter = document.getElementById("finishRosterRoleFilter");
  if (!roleFilter) return;

  const current = String(roleFilter.value || "all").trim();
  roleFilter.innerHTML = buildFinishRosterRoleFilterOptions_(
    current.toUpperCase() === "TQ" ? "Certified" : current
  );
}


function renderFinishRosterApiPanel() {
  applyFinishSetupLabelRename_();

  const panel = document.getElementById("finishRosterApiPanel");
  if (!panel) return;

  const ownerSelect = document.getElementById("finishRosterOwnerSelect");
  const shiftSelect = document.getElementById("finishRosterShiftSelect");
  const roleFilter = document.getElementById("finishRosterRoleFilter");
  const areaFilter = document.getElementById("finishRosterAreaFilter");
  const searchInput = document.getElementById("finishRosterSearchInput");
  const tableBody = document.getElementById("finishRosterTableBody");
  const banner = document.getElementById("finishRosterPermissionBanner");

  const currentUsername = getCurrentUsername();
  const canEdit = canEditFinishOperatorAssignments();
  const currentProfile = getFinishRosterCurrentProfile();
  const selectedOwner = getRosterOwnerUsername();
  const selectedShift = getRosterShiftType();

  if (ownerSelect) {
    const canSeeProfileSelector = canSeeFinishEditingProfileControl();
    const globalOwner = getFinishGlobalRosterOwnerUsername();

    ownerSelect.innerHTML = `
      <option value="${escapeHtml(globalOwner)}" selected>
        ${escapeHtml("Shared Finish Roster · LMS Controlled")}
      </option>`;

    ownerSelect.value = globalOwner;
    finishRosterApiState.ownerUsername = globalOwner;
    ownerSelect.disabled = !canSeeProfileSelector || !canEdit;

    applyFinishEditingProfileVisibility();
  }

  if (shiftSelect) {
    const existing = Array.from(shiftSelect.options || []).map(option => String(option.value || "").toLowerCase());
    if (!existing.includes("all")) {
      const option = document.createElement("option");
      option.value = "all";
      option.textContent = "All Shifts";
      shiftSelect.appendChild(option);
    }

    Array.from(shiftSelect.options || []).forEach(option => {
      const text = String(option.textContent || "").trim().toLowerCase();
      if (text === "all floor operators") {
        option.textContent = "All Shifts";
        option.value = "all";
      }
    });

    shiftSelect.value = String(selectedShift || "").toLowerCase() === "all" ? "all" : selectedShift;
  }
  if (banner) {
    banner.classList.toggle("can-edit", canEdit);
    banner.innerHTML = canEdit
      ? `<strong>Edit access active</strong><span>${escapeHtml(getCurrentUserDisplayName())} · ${escapeHtml(currentProfile?.role || getCurrentUserRole())} can save/edit/delete roster rows.</span>`
      : `<strong>View only</strong><span>${escapeHtml(getCurrentUserDisplayName())} · your role can view this roster, but the API blocks edit/delete.</span>`;
  }

  if (areaFilter && !areaFilter.dataset.finishAreasLoaded) {
    areaFilter.innerHTML = buildFinishRosterAreaOptions(areaFilter.value || 'all', true);
    areaFilter.dataset.finishAreasLoaded = 'true';
  }

  fixFinishRosterRoleFilterNoTq_();

  const filterArea = areaFilter?.value || "all";
  const filterRole = roleFilter?.value || "all";
  const searchText = String(searchInput?.value || "").trim().toLowerCase();
  const rows = buildFinishRosterEditorRows({ filterArea, filterRole, searchText });
  const savedRows = (finishRosterApiState.roster || []).filter(r => String(r.activeStatus || "").toLowerCase() !== "removed");
  const activeRows = savedRows.filter(r => String(r.activeStatus || "").toLowerCase() === "active");
  const mountingCount = activeRows.filter(r => normalizeFinishOperatorStation(r.defaultArea || "") === "Mounting").length;
  const finalCount = activeRows.filter(r => normalizeFinishOperatorStation(r.defaultArea || "") === "Final Inspection").length;
  const exceptionCount = rows.filter(r => !r.isSaved || !r.shiftType || !r.defaultArea || !r.defaultPosition || String(r.roleType || "").toLowerCase() === "unassigned").length;

  [["finishRosterOperatorCount", finishRosterApiState.operators.length || 0], ["finishRosterActiveCount", activeRows.length], ["finishRosterMountingCount", mountingCount], ["finishRosterFinalCount", finalCount], ["finishRosterExceptionCount", exceptionCount], ["finishRosterOperatorCountTop", finishRosterApiState.operators.length || 0], ["finishRosterActiveCountTop", activeRows.length], ["finishRosterMountingCountTop", mountingCount], ["finishRosterFinalCountTop", finalCount], ["finishRosterExceptionCountTop", exceptionCount]].forEach(([id, value]) => setText(id, value));
  setText("finishRosterCountMeta", `${rows.length} results`);
  setText("finishRosterApiSyncNote", `Last synced: ${formatTime(new Date())}`);

  const stationLoadout = document.getElementById("finishRosterStationLoadout");

  if (stationLoadout) {
    renderFinishRosterStationLoadout(rows, stationLoadout);
  }

  if (tableBody) {
    if (finishRosterApiState.lastError) {
      tableBody.innerHTML = `<tr><td colspan="10" class="roster-api-empty">${escapeHtml(finishRosterApiState.lastError)}</td></tr>`;
    } else if (!rows.length) {
      tableBody.innerHTML = `<tr><td colspan="10" class="roster-api-empty">No operators match the current filter selection.</td></tr>`;
    } else {
      tableBody.innerHTML = rows.map(renderFinishRosterTableRow).join("");
    }
  }

  bindFinishRosterTableEvents(rows);

  if (!finishRosterUiState.selectedRosterKey && rows.length && !finishRosterUiState.draftAssignment) {
    finishRosterUiState.selectedRosterKey = makeFinishRosterAssignmentKey(rows[0]);
  }

  renderFinishRosterDetailPanel(rows, currentProfile, canEdit);
  renderFinishRosterAuditLog();
  fixFinishRosterRoleFilterNoTq_();
}


function getFinishRosterStationTheme(station) {
  const clean = normalizeFinishOperatorStation(station || "");
  if (clean.includes("MEI")) return "cyan";
  if (clean === "Mounting") return "orange";
  if (clean === "Final Inspection") return "green";
  if (clean === "Drill") return "purple";
  if (clean === "Bigs" || clean === "Sharps") return "pink";
  if (clean === "Floater" || clean === "Utility") return "amber";
  return "blue";
}

function getFinishRosterCardInitials(name) {
  return String(name || "")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map(part => part[0])
    .join("")
    .toUpperCase() || "OP";
}

function getFinishRosterStationJph(rows) {
  return (rows || []).reduce((sum, row) => {
    const custom = Number(String(row.individualJph || "").replace(/,/g, ""));
    if (Number.isFinite(custom) && custom > 0) return sum + custom;

    const finalJph = Number(String(row.finalJph || row.FinalJPH || "").replace(/,/g, ""));
    if (Number.isFinite(finalJph) && finalJph > 0) return sum + finalJph;

    const station = normalizeFinishOperatorStation(row.defaultArea || row.liveStation || "");
    if (station === "Mounting") return sum + 25;
    if (station === "Final Inspection") return sum + 75;
    if (station === "Drill") return sum + 65;
    if (station === "Finish Unbox") return sum + 125;
    if (station.includes("MEI")) return sum + 50;
    if (station === "Bigs" || station === "Sharps") return sum + 15;
    return sum;
  }, 0);
}

function renderFinishRosterStationLoadout(rows, target) {
  if (!target) return;

  if (finishRosterApiState.lastError) {
    target.innerHTML = `<article class="roster-api-empty">${escapeHtml(finishRosterApiState.lastError)}</article>`;
    return;
  }

  const stationOrder = FINISH_ROSTER_ASSIGNABLE_AREAS.filter(station => !shouldHideFinishLmsArea_(station));
  const grouped = {};
  stationOrder.forEach(station => { grouped[station] = []; });

  (rows || []).forEach(row => {
    const station = normalizeFinishOperatorStation(row.defaultArea || row.liveStation || "Unassigned");
    if (shouldHideFinishLmsArea_(station)) return;
    const key = stationOrder.includes(station) ? station : "Floater";
    grouped[key].push(row);
  });

  Object.keys(grouped).forEach(key => {
    grouped[key] = (grouped[key] || []).filter(row => {
      const name = row.operatorName || "";
      const station = row.defaultArea || row.liveStation || key || "";
      return (
        !shouldHideFinishOperatorOnWebpage(name) &&
        !shouldHideFinishLmsArea_(name) &&
        !shouldHideFinishLmsArea_(station)
      );
    });
  });

  const visibleStations = stationOrder.filter(station => grouped[station] && grouped[station].length);
  const stationsToRender = visibleStations.length ? visibleStations : stationOrder.slice(0, 6);

  if (!rows || !rows.length) {
    target.innerHTML = `<article class="finish-roster-empty-loadout"><h3>No operators match this setup</h3><p>Clear filters or add an operator to begin the Finish roster setup.</p></article>`;
    return;
  }

  target.innerHTML = stationsToRender.map(station => {
    const stationRows = (grouped[station] || []).sort((a, b) => String(a.operatorName || "").localeCompare(String(b.operatorName || "")));
    const theme = getFinishRosterStationTheme(station);
    const activeCount = stationRows.filter(row => isFinishRosterRowAssignedForUi_(row)).length;
    const jph = getFinishRosterStationJph(stationRows);
    const cards = stationRows.length
      ? stationRows.map(renderFinishRosterStationAssociateCard).join("")
      : `<article class="finish-roster-empty-station">No associates assigned</article>`;

    return `
      <section class="finish-roster-station-panel ${escapeHtml(theme)}">
        <header class="finish-roster-station-header">
          <div>
            <span class="finish-roster-station-dot"></span>
            <h4>${escapeHtml(station)}</h4>
          </div>
          <div class="finish-roster-station-metrics">
            <strong>${numberFmt(activeCount)}</strong><span>body</span>
            <strong>${numberFmt(jph)}</strong><span>lane jph</span>
          </div>
        </header>
        <div class="finish-roster-associate-grid">
          ${cards}
        </div>
      </section>`;
  }).join("");
}

function renderFinishRosterStationAssociateCard(row) {
  const key = makeFinishRosterAssignmentKey(row);
  const selected = finishRosterUiState.selectedRosterKey === key && !finishRosterUiState.draftAssignment;
  const initials = getFinishRosterCardInitials(row.operatorName);
  const isAssigned = isFinishRosterRowAssignedForUi_(row);
  const status = isAssigned ? normalizeFinishAssignmentStatusForUi_(row) : "Unassigned";
  const role = isAssigned ? normalizeFinishRoleStatusForUi_(row) : "Unassigned";
  const position = [row.defaultLine, row.defaultPosition].filter(Boolean).join(" · ");
  const jph = row.individualJph || row.finalJph || row.FinalJPH || "Default";

  return `
    <button type="button" class="finish-roster-person-card ${selected ? "is-selected" : ""} ${isAssigned ? "" : "is-unassigned"}" data-roster-select="${escapeHtml(key)}">
      <span class="finish-roster-person-avatar">${escapeHtml(initials)}</span>
      <span class="finish-roster-person-main">
        <strong>${escapeHtml(row.operatorName || "Unnamed")}</strong>
        <small>${escapeHtml(row.liveStation || row.defaultArea || "No station")} ${position ? "· " + escapeHtml(position) : ""}</small>
      </span>
      <span class="finish-roster-person-tags">
        <i>${escapeHtml(status)}</i>
        <i>${escapeHtml(role)}</i>
        <i>${escapeHtml(String(jph))} JPH</i>
      </span>
    </button>`;
}

function renderFinishRosterTableRow(row) {
  const selected = finishRosterUiState.selectedRosterKey === makeFinishRosterAssignmentKey(row) && !finishRosterUiState.draftAssignment;
  const initials = String(row.operatorName || "").split(/\s+/).filter(Boolean).slice(0,2).map(part => part[0]).join("").toUpperCase() || "OP";
  const isAssigned = isFinishRosterRowAssignedForUi_(row);
  const status = isAssigned ? normalizeFinishAssignmentStatusForUi_(row) : "Unassigned";
  const shiftDisplay = isAssigned ? row.shiftType : "Not Assigned";
  const statusClass = status.toLowerCase() === "active" ? "active" : "archived";

  return `
    <tr class="${selected ? 'is-selected' : ''}" data-roster-select="${escapeHtml(makeFinishRosterAssignmentKey(row))}">
      <td>
        <div class="finish-roster-associate">
          <span class="finish-roster-associate-badge">${escapeHtml(initials)}</span>
          <div>
            <strong>${escapeHtml(row.operatorName)}</strong>
            <span>${escapeHtml(row.liveStation || '—')} · Today ${numberFmt(row.lastTotal)}${isAssigned ? '' : ' · Not assigned'}</span>
          </div>
        </div>
      </td>
      <td><span class="finish-roster-pill ${isAssigned ? statusClass : 'unassigned'}">${escapeHtml(isAssigned ? status : 'Unassigned')}</span></td>
      <td><span class="${isAssigned ? '' : 'finish-roster-shift-unassigned'}">${escapeHtml(shiftDisplay)}</span></td>
      <td>${escapeHtml(row.defaultArea || '—')}</td>
      <td>${escapeHtml(row.defaultLine || '—')}</td>
      <td>${escapeHtml(row.defaultPosition || '—')}</td>
      <td>${escapeHtml(row.roleType || 'Unassigned')}</td>
      <td>${escapeHtml(row.trainingWeek || '—')}</td>
      <td>${escapeHtml(row.individualJph || 'Default')}</td>
      <td><button type="button" class="finish-roster-row-view" data-roster-view="${escapeHtml(makeFinishRosterAssignmentKey(row))}">View</button></td>
    </tr>`;
}

function bindFinishRosterTableEvents(rows) {
  const map = {};
  rows.forEach(row => { map[makeFinishRosterAssignmentKey(row)] = row; });
  document.querySelectorAll('[data-roster-select], [data-roster-view]').forEach(el => {
    el.addEventListener('click', event => {
      const key = event.currentTarget.dataset.rosterSelect || event.currentTarget.dataset.rosterView || '';
      if (!key || !map[key]) return;
      finishRosterUiState.selectedRosterKey = key;
      finishRosterUiState.draftAssignment = null;
      renderFinishRosterApiPanel();
    });
  });
}

function renderFinishRosterDetailPanel(rows, currentProfile, canEdit) {
  const target = document.getElementById('finishRosterDetailPanel');
  if (!target) return;

  const selected = finishRosterUiState.draftAssignment ? finishRosterUiState.draftAssignment : (rows || []).find(row => makeFinishRosterAssignmentKey(row) === finishRosterUiState.selectedRosterKey);

  if (!selected) {
    target.innerHTML = `<div class="finish-roster-detail-empty"><h3>Assignment Console</h3><p>Select an associate card or use Add Operator to create a new Finish setup row.</p></div>`;
    return;
  }

  const isDraft = !!selected.__draft;
  const operatorOptions = buildFinishRosterOperatorOptions(selected.operatorName || '');
  const disabled = canEdit ? '' : 'disabled';
  const statusTag = String((selected.isSaved || selected.__draft) ? (selected.activeStatus || 'Active') : 'Unassigned');

  target.innerHTML = `
    <div class="finish-roster-detail-header">
      <div>
        <h3>${escapeHtml(selected.operatorName || 'New Assignment')}</h3>
        <p>${isDraft ? 'Draft assignment' : `Live: ${escapeHtml(selected.liveStation || '—')}`} · ${escapeHtml(selected.shiftType || 'Not Assigned')}</p>
      </div>
      <span class="finish-roster-detail-tag">${escapeHtml(statusTag)}</span>
    </div>

    ${(!selected.isSaved || !selected.shiftType || isDraft) ? `<div class="finish-roster-detail-note">Shift is not auto-assigned. Select Weekday or Weekend, then save. After it is saved, Supervisor, Training, Manager, Director, or LMS profiles can manage it in Configuration.</div>` : ""}

    <div class="finish-roster-detail-meta">
      <div class="finish-roster-detail-grid">
        <label class="full"><span>Operator</span>${isDraft ? `<select id="finishRosterDetailOperator" ${disabled}>${operatorOptions}</select>` : `<input id="finishRosterDetailOperator" value="${escapeHtml(selected.operatorName || '')}" disabled />`}</label>
        <label><span>Status</span><select id="finishRosterDetailStatus" ${disabled}>${['Active','Archived'].map(v => `<option value="${v}" ${statusTag === v ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
        <label><span>Shift</span><select id="finishRosterDetailShift" ${disabled}>${['','Weekday','Weekend'].map(v => `<option value="${v}" ${String(selected.shiftType || '') === v ? 'selected' : ''}>${v || 'Select Shift'}</option>`).join('')}</select></label>
        <label><span>Area</span><select id="finishRosterDetailArea" ${disabled}>${buildFinishRosterAreaOptions(selected.defaultArea || '', false)}</select></label>
        <label><span>Line</span><select id="finishRosterDetailLine" ${disabled}>${['','Line A','Line B','Line C'].map(v => `<option value="${v}" ${String(selected.defaultLine || '') === v ? 'selected' : ''}>${v || 'Select'}</option>`).join('')}</select></label>
        <label><span>Position</span><select id="finishRosterDetailPosition" ${disabled}>${buildRosterPositionOptions(selected.defaultArea || '', selected.defaultPosition || '')}</select></label>
        <label><span>Role</span><select id="finishRosterDetailRole" ${disabled}>${['Unassigned','Certified','Training'].map(v => `<option value="${v}" ${String(selected.roleType || 'Unassigned') === v ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
        <label><span>Week</span><select id="finishRosterDetailWeek" ${disabled}>${['','W1','W2','W3','W4','W5','W6','W7','W8'].map(v => `<option value="${v}" ${String(selected.trainingWeek || '') === v ? 'selected' : ''}>${v || '—'}</option>`).join('')}</select></label>
        <label><span>Individual JPH</span><input id="finishRosterDetailJph" type="number" min="0" step="0.1" value="${escapeHtml(selected.individualJph || '')}" placeholder="Default" ${disabled} /></label>
      </div>
    </div>

    <div class="finish-roster-detail-summary">
      <article><span>Editing Profile</span><strong>${escapeHtml(currentProfile?.fullName || getRosterOwnerUsername())}</strong></article>
      <article><span>Target Username</span><strong>${escapeHtml(getRosterOwnerUsername())}</strong></article>
      <article><span>Last Updated</span><strong>${escapeHtml(selected.updatedAt || '—')}</strong></article>
      <article><span>Updated By</span><strong>${escapeHtml(selected.updatedBy || '—')}</strong></article>
    </div>

    <div class="finish-roster-detail-actions">
      <button id="finishRosterDetailSave" class="finish-roster-detail-btn primary" type="button" ${disabled}>${isDraft ? 'Create Assignment' : 'Save Assignment'}</button>
      <button id="finishRosterDetailArchive" class="finish-roster-detail-btn archive" type="button" ${disabled}>Archive Operator</button>
      <button id="finishRosterDetailDelete" class="finish-roster-detail-btn delete" type="button" ${disabled}>Delete Operator</button>
    </div>`;

  bindFinishRosterDetailEvents(selected, isDraft, canEdit);
}

function buildFinishRosterOperatorOptions(selectedName) {
  const options = (finishRosterApiState.operators || [])
    .map(op => normalizeAssignmentOperatorName(op.operatorName || ''))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  const list = options.length ? options : [selectedName || ''];
  return list.map(name => `<option value="${escapeHtml(name)}" ${String(selectedName || '') === name ? 'selected' : ''}>${escapeHtml(name)}</option>`).join('');
}

function openNewFinishRosterAssignment() {
  const canEdit = canEditFinishOperatorAssignments();
  if (!canEdit) {
    showFinishAssignmentToast('View only. Your role cannot create new roster rows.');
    return;
  }
  const firstAvailable = (finishRosterApiState.operators || []).map(op => normalizeAssignmentOperatorName(op.operatorName || '')).find(Boolean) || '';
  finishRosterUiState.selectedRosterKey = '';
  finishRosterUiState.draftAssignment = { __draft: true, operatorName: firstAvailable, shiftType: '', activeStatus: 'Active', defaultArea: 'Mounting', defaultLine: '', defaultPosition: '', roleType: 'Unassigned', trainingWeek: '', individualJph: '', updatedBy: getCurrentUsername(), updatedAt: '' };
  renderFinishRosterApiPanel();
}

function bindFinishRosterDetailEvents(selected, isDraft, canEdit) {
  const areaSelect = document.getElementById('finishRosterDetailArea');
  const positionSelect = document.getElementById('finishRosterDetailPosition');
  const roleSelect = document.getElementById('finishRosterDetailRole');
  const weekSelect = document.getElementById('finishRosterDetailWeek');
  const lineSelect = document.getElementById('finishRosterDetailLine');
  const saveBtn = document.getElementById('finishRosterDetailSave');
  const archiveBtn = document.getElementById('finishRosterDetailArchive');
  const deleteBtn = document.getElementById('finishRosterDetailDelete');
  areaSelect?.addEventListener('change', event => { if (positionSelect) positionSelect.innerHTML = buildRosterPositionOptions(event.target.value, ''); if (lineSelect && event.target.value === 'Drill') lineSelect.value = 'Line C'; });
  roleSelect?.addEventListener('change', event => { if (!weekSelect) return; const enabled = event.target.value === 'Training'; weekSelect.disabled = !enabled || !canEdit; if (!enabled) weekSelect.value = ''; });
  saveBtn?.addEventListener('click', () => saveFinishRosterDetail(selected, isDraft));
  archiveBtn?.addEventListener('click', () => updateFinishRosterDetailStatus(selected, 'archiveFinishRosterOperator'));
  deleteBtn?.addEventListener('click', () => updateFinishRosterDetailStatus(selected, 'deleteFinishRosterOperator'));
}


function getRosterRowDuplicateScore(row) {
  let score = 0;
  if (String(row.activeStatus || "").toLowerCase() === "active") score += 100;
  if (row.defaultArea) score += 10;
  if (row.defaultLine) score += 20;
  if (row.defaultPosition) score += 20;
  if (row.roleType && String(row.roleType).toLowerCase() !== "unassigned") score += 10;
  if (Number(row.individualJph || row.IndividualJPH || 0) > 0) score += 5;
  return score;
}

function makeRosterDuplicateUiKey(row) {
  return [
    String(row.ownerUsername || row.OwnerUsername || getRosterOwnerUsername() || "").trim().toUpperCase(),
    String(row.shiftType || row.ShiftType || getRosterShiftType() || "").trim().toLowerCase(),
    normalizeAssignmentOperatorName(row.operatorName || row.OperatorName || "").toLowerCase(),
    normalizeFinishOperatorStation(row.defaultArea || row.DefaultArea || "")
  ].join("||");
}

function dedupeFinishRosterRowsForUi(rows = []) {
  const best = new Map();

  (Array.isArray(rows) ? rows : []).forEach(row => {
    const key = makeRosterDuplicateUiKey(row);
    if (!key || key.includes("||||")) return;

    const current = best.get(key);
    if (!current || getRosterRowDuplicateScore(row) >= getRosterRowDuplicateScore(current)) {
      best.set(key, row);
    }
  });

  return Array.from(best.values());
}

function findMorningRosterRowsForSlot(operatorName, station, line = "", slot = "") {
  const targetName = normalizeAssignmentOperatorName(operatorName || "").toLowerCase();
  const targetArea = normalizeFinishOperatorStation(station || "");
  const targetLine = String(line || "").trim();
  const targetSlot = String(slot || "").trim();

  const rows = getRawMorningActiveRosterRows().filter(row =>
    normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === targetName &&
    (!targetArea || normalizeFinishOperatorStation(row.defaultArea || "") === targetArea)
  );

  if (!rows.length) return [];

  const exact = rows.filter(row =>
    String(row.defaultLine || "").trim() === targetLine &&
    String(row.defaultPosition || "").trim() === targetSlot
  );

  if (exact.length) return exact;

  const blank = rows.filter(row =>
    !String(row.defaultLine || "").trim() &&
    !String(row.defaultPosition || "").trim()
  );

  return blank.length ? blank : rows;
}

function findMorningRosterRowForSlot(operatorName, station, line = "", slot = "") {
  return findMorningRosterRowsForSlot(operatorName, station, line, slot)[0] || null;
}

function findExistingRosterRowForUpsert(payload = {}) {
  const targetName = normalizeAssignmentOperatorName(payload.operatorName || "").toLowerCase();
  const targetShift = String(payload.shiftType || "").trim().toLowerCase();
  const targetArea = normalizeFinishOperatorStation(payload.defaultArea || "");
  const targetLine = String(payload.defaultLine || "").trim();
  const targetPosition = String(payload.defaultPosition || "").trim();

  const candidates = (finishRosterApiState.roster || []).filter(row =>
    String(row.activeStatus || "").toLowerCase() !== "removed" &&
    normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === targetName &&
    String(row.shiftType || "").trim().toLowerCase() === targetShift
  );

  if (!candidates.length) return null;

  return candidates.find(row =>
    normalizeFinishOperatorStation(row.defaultArea || "") === targetArea &&
    String(row.defaultLine || "").trim() === targetLine &&
    String(row.defaultPosition || "").trim() === targetPosition
  ) || candidates.find(row =>
    normalizeFinishOperatorStation(row.defaultArea || "") === targetArea &&
    !String(row.defaultLine || "").trim() &&
    !String(row.defaultPosition || "").trim()
  ) || candidates.find(row =>
    !String(row.defaultArea || "").trim() &&
    !String(row.defaultLine || "").trim() &&
    !String(row.defaultPosition || "").trim()
  ) || null;
}

function applyOriginalRosterLookupToPayload(payload, sourceRow) {
  if (!payload || !sourceRow) return payload;
  payload.originalDefaultArea = sourceRow.defaultArea || "";
  payload.originalDefaultLine = sourceRow.defaultLine || "";
  payload.originalDefaultPosition = sourceRow.defaultPosition || "";
  return payload;
}

async function saveFinishRosterDetail(selected, isDraft) {
  if (!canEditFinishOperatorAssignments()) { showFinishAssignmentToast('View only. Your role cannot edit Finish roster rows.'); return; }
  const operatorName = String(document.getElementById('finishRosterDetailOperator')?.value || selected.operatorName || '').trim();
  if (!operatorName) { showFinishAssignmentToast('Pick an operator before saving.'); return; }
  const selectedShiftValue = document.getElementById('finishRosterDetailShift')?.value || '';
  if (!selectedShiftValue) { showFinishAssignmentToast('Select Weekday or Weekend before saving this operator.'); return; }
  const selectedRoleRaw = document.getElementById('finishRosterDetailRole')?.value || 'Unassigned';
  const selectedRole = String(selectedRoleRaw).toUpperCase() === 'TQ' ? 'Certified' : selectedRoleRaw;
  const payload = { updatedBy: getCurrentUsername(), ownerUsername: getRosterOwnerUsername(), shiftType: selectedShiftValue, operatorName, activeStatus: document.getElementById('finishRosterDetailStatus')?.value || 'Active', defaultArea: document.getElementById('finishRosterDetailArea')?.value || '', defaultLine: document.getElementById('finishRosterDetailLine')?.value || '', defaultPosition: document.getElementById('finishRosterDetailPosition')?.value || '', roleType: selectedRole, trainingWeek: selectedRole === 'Training' ? (document.getElementById('finishRosterDetailWeek')?.value || '') : '', individualJph: document.getElementById('finishRosterDetailJph')?.value || '' };
  if (!isDraft) {
    payload.originalDefaultArea = selected.defaultArea || '';
    payload.originalDefaultLine = selected.defaultLine || '';
    payload.originalDefaultPosition = selected.defaultPosition || '';
  } else {
    const existingForUpsert = findExistingRosterRowForUpsert(payload);
    if (existingForUpsert) applyOriginalRosterLookupToPayload(payload, existingForUpsert);
  }

  try {
    await fetchFinishRosterApi('saveFinishRosterControl', payload);
    finishRosterUiState.draftAssignment = null;
    showFinishAssignmentToast(`Saved ${operatorName} / ${payload.shiftType}`);
    await loadFinishRosterForSelectedProfile({ silent: true });
    const rows = buildFinishRosterEditorRows({ filterArea: document.getElementById('finishRosterAreaFilter')?.value || 'all', filterRole: document.getElementById('finishRosterRoleFilter')?.value || 'all', searchText: String(document.getElementById('finishRosterSearchInput')?.value || '').trim().toLowerCase() });
    const match = rows.find(row => normalizeAssignmentOperatorName(row.operatorName || '') === normalizeAssignmentOperatorName(operatorName) && String(row.defaultArea || '') === String(payload.defaultArea || '') && String(row.defaultLine || '') === String(payload.defaultLine || '') && String(row.defaultPosition || '') === String(payload.defaultPosition || ''));
    finishRosterUiState.selectedRosterKey = match ? makeFinishRosterAssignmentKey(match) : '';
    renderFinishRosterApiPanel();
  } catch (error) { console.error(error); showFinishAssignmentToast(error.message || String(error)); }
}

async function updateFinishRosterDetailStatus(selected, action) {
  if (!selected || selected.__draft || !selected.shiftType) { showFinishAssignmentToast('Save the assignment with a shift first before archive/delete.'); return; }
  if (!canEditFinishOperatorAssignments()) { showFinishAssignmentToast('View only. Your role cannot delete/archive Finish roster rows.'); return; }
  try {
    await fetchFinishRosterApi(action, { updatedBy: getCurrentUsername(), ownerUsername: getRosterOwnerUsername(), shiftType: selected.shiftType || getRosterShiftType(), operatorName: selected.operatorName, defaultArea: selected.defaultArea || '', defaultLine: selected.defaultLine || '', defaultPosition: selected.defaultPosition || '' });
    showFinishAssignmentToast(`${action.includes('delete') ? 'Deleted' : 'Archived'} ${selected.operatorName}`);
    finishRosterUiState.selectedRosterKey = ''; finishRosterUiState.draftAssignment = null;
    await loadFinishRosterForSelectedProfile({ silent: true });
    renderFinishRosterApiPanel();
  } catch (error) { console.error(error); showFinishAssignmentToast(error.message || String(error)); }
}

function makeFinishRosterAssignmentKey(row) {
  return [
    normalizeAssignmentOperatorName(row.operatorName || ""),
    String(row.shiftType || row.rosterBucket || "Not Assigned"),
    normalizeFinishOperatorStation(row.defaultArea || ""),
    String(row.defaultLine || ""),
    String(row.defaultPosition || "")
  ].join("||").toLowerCase();
}

function buildFinishRosterEditorRows(options = {}) {
  const { filterArea = "all", filterRole = "all", searchText = "" } = options || {};
  const selectedShift = getRosterShiftType();
  const rosterRows = dedupeFinishRosterRowsForUi(filterFinishWebpageRows_(Array.isArray(finishRosterApiState.roster) ? finishRosterApiState.roster : []));
  const masterOps = filterFinishWebpageRows_(Array.isArray(finishRosterApiState.operators) ? finishRosterApiState.operators : []);

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

  rosterRows.forEach(saved => {
    const operatorName = normalizeAssignmentOperatorName(saved.operatorName || "");
    if (!operatorName) return;
    const live = liveByName[operatorName.toLowerCase()] || {};
    rows.push({
      operatorName,
      shiftType: saved.shiftType || selectedShift,
      liveStation: live.liveStation || normalizeFinishOperatorStation(saved.defaultArea || ""),
      lastTotal: Number(live.lastTotal || 0) || 0,
      activeStatus: normalizeFinishAssignmentStatusForUi_(saved),
      defaultLine: saved.defaultLine || "",
      defaultArea: normalizeFinishOperatorStation(saved.defaultArea || ""),
      defaultPosition: saved.defaultPosition || "",
      roleType: normalizeFinishRoleStatusForUi_(saved),
      trainingWeek: saved.trainingWeek || "",
      individualJph: saved.individualJph ?? saved.individualJPH ?? "",
      isSaved: true
    });
  });

  const savedKeys = new Set(rows.map(row => `${normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase()}||${String(row.defaultArea || "").toLowerCase()}||${String(row.defaultLine || "").toLowerCase()}||${String(row.defaultPosition || "").toLowerCase()}`));

  masterOps.forEach(op => {
    const operatorName = normalizeAssignmentOperatorName(op.operatorName || "");
    if (!operatorName) return;
    if (operatorName.toLowerCase().includes("unassigned")) return;

    const liveStation = normalizeFinishOperatorStation(op.lastFlowStation || op.lastAccessPoint || "");
    if (liveStation === "FSV Scan & Verify") return;

    const key = `${operatorName.toLowerCase()}||||`;
    const alreadySaved = rows.some(row => normalizeAssignmentOperatorName(row.operatorName || "").toLowerCase() === operatorName.toLowerCase());
    if (alreadySaved) return;

    // Strict shift filter:
    // Weekday/Weekend shows only associates already saved to that selected shift.
    // All Shifts is the only view that includes first-time / not-assigned live operators.
    if (String(selectedShift || "").toLowerCase() !== "all") return;

    rows.push({
      operatorName,
      shiftType: "",
      rosterBucket: "Not Assigned",
      liveStation,
      lastTotal: Number(op.lastTotal || 0) || 0,
      activeStatus: "Unassigned",
      defaultLine: liveStation === "Drill" ? "Line C" : "",
      defaultArea: isFinishRosterAssignableArea(liveStation) ? liveStation : "",
      defaultPosition: "",
      roleType: "Unassigned",
      trainingWeek: "",
      individualJph: "",
      isSaved: false
    });
  });

  return rows
    .filter(row => {
      if (String(selectedShift || "").toLowerCase() === "all") return true;
      return (
        !!row.shiftType &&
        String(row.shiftType || "").toLowerCase() === String(selectedShift || "").toLowerCase()
      );
    })
    .filter(row => filterArea === "all" || row.defaultArea === filterArea || row.liveStation === filterArea)
    .filter(row => filterRole === "all" || (filterRole === "Unassigned" ? (!row.isSaved || String(row.roleType || "Unassigned") === "Unassigned") : String(row.roleType || "Unassigned") === filterRole))
    .filter(row => {
      if (!searchText) return true;
      const haystack = [row.operatorName, row.defaultArea, row.defaultLine, row.defaultPosition, row.roleType, row.liveStation]
        .join(" ")
        .toLowerCase();
      return haystack.includes(searchText);
    })
    .sort((a, b) => {
      const activeA = a.activeStatus === "Active" ? 0 : 1;
      const activeB = b.activeStatus === "Active" ? 0 : 1;
      if (activeA !== activeB) return activeA - activeB;
      const savedA = a.isSaved ? 0 : 1;
      const savedB = b.isSaved ? 0 : 1;
      if (savedA !== savedB) return savedA - savedB;
      return a.operatorName.localeCompare(b.operatorName);
    });
}

function renderFinishRosterEditorRow(row, canEdit) {
  const disabled = canEdit ? "" : "disabled";
  const statusClass = row.activeStatus === "Archived" ? "is-archived" : "is-active";
  const rowKey = makeFinishRosterAssignmentKey(row);
  const expanded = finishRosterUiState.expandedRosterRows.has(rowKey);
  const shiftType = row.shiftType || getRosterShiftType();
  const roleLabel = row.roleType === "Training" ? `${row.roleType} ${row.trainingWeek || ""}`.trim() : (row.roleType || "Unassigned");

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
          <span>${escapeHtml(row.defaultArea || row.liveStation || "Unassigned")} · ${escapeHtml(row.defaultLine || "No Line")} · ${escapeHtml(row.defaultPosition || "No Position")}</span>
          <span>${escapeHtml(roleLabel)} · ${escapeHtml(shiftType)} · Today ${numberFmt(row.lastTotal)}</span>
        </div>

        <div class="finish-roster-row-tags">
          <span>${escapeHtml(row.activeStatus || "Active")}</span>
          <span>${escapeHtml(shiftType)}</span>
          <span>${escapeHtml(row.defaultArea || "No area")}</span>
          <span>${escapeHtml(row.defaultLine || "No line")}</span>
          <span>${escapeHtml(row.defaultPosition || "No position")}</span>
          <span>${row.isSaved ? "Saved" : "Not saved"}</span>
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
            ${buildFinishRosterAreaOptions(row.defaultArea || '', false)}
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
            ${["Unassigned", "Certified", "Training"].map(v => `<option value="${v}" ${row.roleType === v ? "selected" : ""}>${v}</option>`).join("")}
          </select>
        </label>

        <label><span>Week</span>
          <select data-roster-field="trainingWeek" ${disabled || row.roleType !== "Training" ? "disabled" : ""}>
            ${["", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8"].map(v => `<option value="${v}" ${row.trainingWeek === v ? "selected" : ""}>${v || "—"}</option>`).join("")}
          </select>
        </label>

        <label><span>Individual JPH</span>
          <input data-roster-field="individualJph" type="number" min="0" step="0.1" placeholder="Default" value="${escapeHtml(row.individualJph || "")}" ${disabled} />
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

    if (String(getRosterShiftType()).toLowerCase() !== "all" && saveShift !== getRosterShiftType()) {
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

    showFinishProfileSavedToast("Roster change saved");
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
  applyFinishSetupLabelRename_();
  forceFinishSetupTabVisibilityForAllowedRoles_();
  initFinishRosterApiPanel();
  initFinishLmsHideAreaObserver_();
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

  if (typeof injectOperatorAssignmentStyles === "function") {
    injectOperatorAssignmentStyles();
  }

  const configTab = document.querySelector('[data-content="config"]');
  const saveRow = document.querySelector(".config-actions") || document.getElementById("configSaveBtn")?.parentElement;
  const targetParent = saveRow?.parentElement || configTab;
  if (!targetParent && !panel) return null;

  if (!panel) {
    panel = document.createElement("div");
    panel.id = "operatorAssignmentConfigPanel";
    panel.className = "operator-assignment-config-panel";

    if (saveRow && saveRow.parentElement) {
      saveRow.parentElement.insertBefore(panel, saveRow);
    } else {
      targetParent.appendChild(panel);
    }
  }

  // The side-tab HTML already includes an empty #operatorAssignmentConfigPanel.
  // If we return early, the roster list never gets created. Always hydrate it.
  if (!panel.querySelector("#operatorAssignmentRoster")) {
    panel.innerHTML = `
      <div class="operator-assignment-shell config-associates-shell">
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

        <div id="operatorAssignmentRoster" class="assignment-roster">
          <article class="assignment-empty">Waiting for assigned roster...</article>
        </div>
      </div>
    `;
  }

  const shiftSelect = document.getElementById("configAssignedShiftSelect");
  if (shiftSelect && !shiftSelect.dataset.boundConfigAssignedShift) {
    shiftSelect.dataset.boundConfigAssignedShift = "1";
    shiftSelect.value = window.__finishConfigSelectedShiftType || getFinishConfigDefaultShiftType();
    shiftSelect.addEventListener("change", () => {
      window.__finishConfigSelectedShiftType = getFinishConfigSelectedShiftType();
      renderOperatorAssignmentConfigPanel();
    });
  }

  const areaFilter = document.getElementById("configAssignedAreaFilter");
  if (areaFilter && !areaFilter.dataset.boundConfigAssignedArea) {
    areaFilter.dataset.boundConfigAssignedArea = "1";
    areaFilter.addEventListener("change", renderOperatorAssignmentConfigPanel);
  }

  const refreshBtn = document.getElementById("configAssignedRefreshBtn");
  if (refreshBtn && !refreshBtn.dataset.boundConfigAssignedRefresh) {
    refreshBtn.dataset.boundConfigAssignedRefresh = "1";
    refreshBtn.addEventListener("click", renderOperatorAssignmentConfigPanel);
  }

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

  return dedupeFinishRosterRowsForUi(rows
    .filter(row => String(row.ownerUsername || "").trim().toUpperCase() === ownerUsername)
    .filter(row => String(row.shiftType || "").trim().toLowerCase() === shiftType.toLowerCase())
    .filter(row => String(row.activeStatus || "").trim().toLowerCase() === "active"))
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
      const canEdit = canEditFinishOperatorAssignments();
      notice.classList.toggle("can-edit", canEdit);
      notice.classList.toggle("view-only", !canEdit);
      notice.innerHTML = canEdit
        ? `<strong>Role + training control active</strong><span>Supervisor, Training, Manager, Director, and LMS can set Certified / Training role, training week, and JPH right here.</span>`
        : `<strong>Assigned roster only</strong><span>Associate list comes from saved roster assignment. Live totals come from Operator Activity.</span>`;
    }

    if (!visibleRows.length) {
      rosterTarget.innerHTML = `
        <article class="assignment-empty">
          No active associates assigned to ${escapeHtml(ownerUsername)} / ${escapeHtml(shiftType)}${areaFilter !== "all" ? " / " + escapeHtml(areaFilter) : ""}.
        </article>
      `;
      updateConfigTotals();
      applyFinishAssignmentPermissionLock();
      return;
    }

    const config = loadConfig();

    rosterTarget.innerHTML = visibleRows.map(row => {
      const operatorName = row.operatorName || "";
      const area = normalizeFinishOperatorStation(row.defaultArea || "");
      const line = row.defaultLine || "No line";
      const position = row.defaultPosition || "No position";
      const role = row.roleType || "Unassigned";
      const liveTotal = getFinishConfigLiveOperatorTotal(operatorName, area);
      const target = getFinishConfigRoleTarget(row, config);
      const canTrain = finishConfigAreaSupportsTraining(area);
      const weekValue = String(convertApiWeekToNumber(row.trainingWeek || "") || 1);

      return `
        <article class="assignment-row" data-assigned-roster-row="${escapeHtml(operatorName)}" data-area="${escapeHtml(area)}" data-line="${escapeHtml(line)}" data-position="${escapeHtml(position)}">
          <div class="assignment-operator">
            <strong>${escapeHtml(operatorName)}</strong>
            <span>${escapeHtml(area)} · ${escapeHtml(line)} · ${escapeHtml(position)} · ${escapeHtml(getFinishConfigLiveOperatorLabel(operatorName, area))}</span>
          </div>

          <label class="assignment-field assignment-field--role">
            <span>Role</span>
            <select data-config-role-select>
              ${buildConfigAssignedRoleOptions(area, role)}
            </select>
          </label>

          <label class="assignment-field assignment-field--week ${role === 'Training' && canTrain ? 'active' : 'disabled'}">
            <span>Week</span>
            <select data-config-week-select ${role === 'Training' && canTrain ? '' : 'disabled'}>
              ${buildConfigAssignedWeekOptions(area, role, weekValue)}
            </select>
          </label>

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

    rosterTarget.querySelectorAll(".assignment-row").forEach(rowEl => {
      syncConfigAssignedRowInputs(rowEl, { force: false });
    });

    rosterTarget.querySelectorAll("[data-config-role-select]").forEach(select => {
      select.addEventListener("change", event => {
        const rowEl = event.target.closest(".assignment-row");
        syncConfigAssignedRowInputs(rowEl, { force: true });
      });
    });

    rosterTarget.querySelectorAll("[data-config-week-select]").forEach(select => {
      select.addEventListener("change", event => {
        const rowEl = event.target.closest(".assignment-row");
        syncConfigAssignedRowInputs(rowEl, { force: true });
      });
    });

    rosterTarget.querySelectorAll("[data-config-jph-save]").forEach(button => {
      button.addEventListener("click", event => saveConfigAssignedIndividualJph(event.target.closest(".assignment-row")));
    });

    updateConfigTotals();
    renderFinishThreeLineCell();
    applyFinishAssignmentPermissionLock();
  } catch (error) {
    console.error("Assigned roster render failed:", error);
    rosterTarget.innerHTML = `
      <article class="assignment-empty">
        Could not load assigned roster: ${escapeHtml(error.message || String(error))}
      </article>
    `;
  }
}

function finishConfigAreaSupportsTraining(area) {
  const station = normalizeFinishOperatorStation(area || '');
  return station === 'Mounting' || station === 'Final Inspection';
}

function buildConfigAssignedRoleOptions(area, selectedRole) {
  const role = String(selectedRole || 'Unassigned');
  const options = [
    { value: 'Certified', label: 'Certified' },
    { value: 'TQ', label: 'TQ' }
  ];

  if (finishConfigAreaSupportsTraining(area)) {
    options.push({ value: 'Training', label: 'Training' });
  }

  options.push({ value: 'Unassigned', label: 'Unassigned' });

  return options.map(opt => `<option value="${escapeHtml(opt.value)}" ${String(role).toLowerCase() === String(opt.value).toLowerCase() ? 'selected' : ''}>${escapeHtml(opt.label)}</option>`).join('');
}

function buildConfigAssignedWeekOptions(area, roleType, selectedWeek) {
  if (String(roleType || '').toLowerCase() !== 'training' || !finishConfigAreaSupportsTraining(area)) {
    return '<option value="1">Week 1</option>';
  }
  return buildTrainingWeekOptions(area, Number(selectedWeek || 1));
}

function getConfigAssignedSuggestedTarget(area, roleType, weekValue) {
  return getFinishConfigRoleTarget({
    defaultArea: area,
    roleType,
    trainingWeek: String(roleType || '').toLowerCase() === 'training' ? `W${Number(weekValue || 1)}` : '',
    individualJph: ''
  }, loadConfig());
}

function syncConfigAssignedRowInputs(rowEl, { force = false } = {}) {
  if (!rowEl) return;
  const area = rowEl.dataset.area || '';
  const roleSelect = rowEl.querySelector('[data-config-role-select]');
  const weekSelect = rowEl.querySelector('[data-config-week-select]');
  const weekWrap = weekSelect?.closest('.assignment-field') || weekSelect?.closest('label');
  const jphInput = rowEl.querySelector('[data-config-jph-input]');
  if (!roleSelect || !jphInput) return;

  const roleType = roleSelect.value || 'Unassigned';
  const canTrain = finishConfigAreaSupportsTraining(area) && String(roleType).toLowerCase() === 'training';

  if (weekSelect) {
    const selectedWeek = weekSelect.value || '1';
    weekSelect.innerHTML = buildConfigAssignedWeekOptions(area, roleType, selectedWeek);
    weekSelect.disabled = !canTrain;
    if (!canTrain) {
      weekSelect.value = '1';
    } else if (!weekSelect.value) {
      weekSelect.value = '1';
    }
  }
  if (weekWrap) weekWrap.classList.toggle('disabled', !canTrain);
  if (weekWrap) weekWrap.classList.toggle('active', canTrain);

  const suggested = Number(getConfigAssignedSuggestedTarget(area, roleType, weekSelect?.value || '1')) || 0;
  jphInput.placeholder = suggested > 0 ? numberFmt(suggested) : 'Default';
  if (force) {
    jphInput.value = suggested > 0 ? String(suggested) : '';
  }
}

async function saveConfigAssignedIndividualJph(rowEl) {
  if (!rowEl) return;

  const operatorName = rowEl.dataset.assignedRosterRow || "";
  const source = (window.__finishConfigAssignedRosterRows || []).find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "") === normalizeAssignmentOperatorName(operatorName) &&
    normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(rowEl.dataset.area || rowEl.querySelector(".assignment-operator span")?.textContent?.split(" · ")?.[0] || "") &&
    String(row.defaultLine || "No line") === String(rowEl.dataset.line || '') &&
    String(row.defaultPosition || "No position") === String(rowEl.dataset.position || '')
  ) || (window.__finishConfigAssignedRosterRows || []).find(row =>
    normalizeAssignmentOperatorName(row.operatorName || "") === normalizeAssignmentOperatorName(operatorName) &&
    normalizeFinishOperatorStation(row.defaultArea || "") === normalizeFinishOperatorStation(rowEl.dataset.area || rowEl.querySelector(".assignment-operator span")?.textContent?.split(" · ")?.[0] || "")
  );

  const areaLinePositionText = rowEl.querySelector(".assignment-operator span")?.textContent || "";
  const parts = areaLinePositionText.split(" · ").map(x => String(x || "").trim());

  const defaultArea = source?.defaultArea || parts[0] || "";
  const defaultLine = source?.defaultLine || rowEl.dataset.line || parts[1] || "";
  const defaultPosition = source?.defaultPosition || rowEl.dataset.position || parts[2] || "";
  const roleType = rowEl.querySelector('[data-config-role-select]')?.value || source?.roleType || 'Unassigned';
  const weekNumber = rowEl.querySelector('[data-config-week-select]')?.value || '1';
  const trainingWeek = String(roleType).toLowerCase() === 'training' && finishConfigAreaSupportsTraining(defaultArea) ? `W${Number(weekNumber || 1)}` : '';
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
      roleType,
      trainingWeek,
      individualJph,
      originalDefaultArea: defaultArea,
      originalDefaultLine: defaultLine,
      originalDefaultPosition: defaultPosition
    });

    showFinishAssignmentToast(`Saved ${operatorName} / ${roleType}${trainingWeek ? ' ' + trainingWeek : ''}`);
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
  const shift = document.getElementById("configAssignedShiftSelect");
  const area = document.getElementById("configAssignedAreaFilter");
  const refresh = document.getElementById("configAssignedRefreshBtn");
  if (shift) shift.disabled = false;
  if (area) area.disabled = false;
  if (refresh) refresh.disabled = false;

  document.querySelectorAll('#operatorAssignmentConfigPanel [data-config-role-select], #operatorAssignmentConfigPanel [data-config-week-select], #operatorAssignmentConfigPanel [data-config-jph-input], #operatorAssignmentConfigPanel [data-config-jph-save]').forEach(el => {
    if (!el) return;
    if (el.hasAttribute('data-config-week-select')) {
      const row = el.closest('.assignment-row');
      const roleType = row?.querySelector('[data-config-role-select]')?.value || 'Unassigned';
      const trainingEnabled = String(roleType).toLowerCase() === 'training';
      el.disabled = !canEdit || !trainingEnabled;
    } else {
      el.disabled = !canEdit;
    }
    el.classList.toggle('locked', !canEdit);
  });
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

/*******************************************************
 * FINISH MORNING SET UP — INACTIVITY FRONTEND ADD-ON
 * Clean version: removes duplicated handlers, uses safe API calls,
 * never crashes on blank rows, and does not wipe the panel on empty response.
 *******************************************************/

const finishMorningInactivityState = {
  rows: [],
  logRows: [],
  dailyRows: [],
  latestRows: [],
  automationEnabled: false,
  thresholdSnapshots: 3,
  loading: false,
  lastSync: ""
};

function normalizeFinishInactivityStation(value) {
  const text = String(value || "").trim().toLowerCase();
  if (text.includes("final")) return "Final Inspection";
  if (text.includes("drill")) return "Drill";
  if (text.includes("mount")) return "Mounting";
  return String(value || "").trim();
}

function getFinishMorningInactivityThreshold() {
  const value = Number(document.getElementById("finishInactivityThreshold")?.value || finishMorningInactivityState.thresholdSnapshots || 3) || 3;
  return Math.max(1, Math.min(12, value));
}

function getFinishMorningInactivityRows() {
  return Array.isArray(finishMorningInactivityState.rows) ? finishMorningInactivityState.rows : [];
}

function getFinishInactivityOwnerUsernameSafe() {
  if (typeof getMorningSetupOwnerUsername === "function") return getMorningSetupOwnerUsername();
  if (typeof getCurrentUsername === "function") return getCurrentUsername();
  return "BLOPEZ";
}

function getFinishInactivityShiftTypeSafe() {
  if (typeof getMorningSetupShiftType === "function") return getMorningSetupShiftType();
  const select = document.getElementById("morningShiftSelect") || document.getElementById("finishMorningShiftSelect");
  const raw = select?.value || select?.selectedOptions?.[0]?.textContent || "Weekday";
  return String(raw).toLowerCase().includes("weekend") ? "Weekend" : "Weekday";
}

async function fetchFinishInactivityApiSafe(action, params = {}) {
  const url = new URL(FINISH_API_URL);
  url.searchParams.set("action", String(action || "").toLowerCase());

  Object.entries(params || {}).forEach(([key, value]) => {
    if (value !== undefined && value !== null && String(value) !== "") {
      url.searchParams.set(key, value);
    }
  });

  url.searchParams.set("cacheBust", Date.now());

  const response = await fetch(url.toString(), { cache: "no-store" });
  if (!response.ok) throw new Error(`Finish inactivity API failed: ${response.status}`);

  const payload = await response.json();
  if (!payload || payload.success === false || payload.status === "error") {
    throw new Error(payload?.message || payload?.error || "Finish inactivity API returned an error.");
  }

  return payload;
}

function getFinishInactivityValue(row, ...keys) {
  if (!row || typeof row !== "object") return "";
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") return row[key];
  }
  return "";
}

function normalizeFinishInactivityStatus(row) {
  const status = String(getFinishInactivityValue(row, "CurrentStatus", "Status", "status") || "").trim().toUpperCase();
  if (status === "OPEN" || status === "INACTIVE") return "OPEN";
  if (status === "WATCH") return "WATCH";
  if (status === "RESOLVED") return "RESOLVED";
  if (status === "ACTIVE") return "ACTIVE";
  return "WATCH";
}

function finishInactivityExtractLatestRows(payload) {
  return (
    Array.isArray(payload?.latestRows) ? payload.latestRows :
    Array.isArray(payload?.liveRows) ? payload.liveRows :
    Array.isArray(payload?.currentRows) ? payload.currentRows :
    []
  ).filter(row => row && typeof row === "object");
}

function finishInactivityExtractLogRows(payload) {
  return (
    Array.isArray(payload?.logRows) ? payload.logRows :
    Array.isArray(payload?.rows) ? payload.rows :
    Array.isArray(payload?.data) ? payload.data :
    []
  ).filter(row => row && typeof row === "object");
}

function finishInactivityExtractDailyRows(payload) {
  return (
    Array.isArray(payload?.dailyRows) ? payload.dailyRows :
    []
  ).filter(row => row && typeof row === "object");
}

function finishInactivityBuildRowsFromPayload(payload) {
  const dailyRows = finishInactivityExtractDailyRows(payload);
  const logRows = finishInactivityExtractLogRows(payload);
  const latestRows = finishInactivityExtractLatestRows(payload);

  /*
    Use latestRows first when the API provides them.
    latestRows are the live/current associate rows with RAW_ACTIVITY_CURRENT totals and hourly output.
    logRows remain available for session history, but should not be the primary source for current totals.
  */
  const primaryRows = latestRows.length ? latestRows : logRows;
  const latestByKey = new Map();

  primaryRows.forEach((row, index) => {
    if (!row || typeof row !== "object") return;

    const key = finishInactivityRowIdentityKey_(row);
    if (!key) return;

    const currentScore = finishInactivityRowFreshnessScore_(row, index);
    const existing = latestByKey.get(key);

    if (!existing || currentScore >= existing.__score) {
      latestByKey.set(key, { ...row, __score: currentScore });
    }
  });

  const dailyByKey = new Map();
  dailyRows.forEach((row, index) => {
    if (!row || typeof row !== "object") return;
    const key = finishInactivityRowIdentityKey_(row);
    if (!key) return;
    dailyByKey.set(key, { ...row, __score: finishInactivityRowFreshnessScore_(row, index) });
  });

  dailyByKey.forEach((dailyRow, key) => {
    const existing = latestByKey.get(key);

    if (existing) {
      latestByKey.set(key, {
        ...existing,
        DailyStatus: dailyRow.CurrentStatus || dailyRow.Status || "",
        CurrentStatus: dailyRow.CurrentStatus || existing.CurrentStatus || existing.Status || "",
        TotalInactiveMinutes: dailyRow.TotalInactiveMinutes || existing.TotalInactiveMinutes || existing.InactiveMinutes || "",
        TotalInactiveHours: dailyRow.TotalInactiveHours || existing.TotalInactiveHours || "",
        CurrentInactiveMinutes: dailyRow.CurrentInactiveMinutes || existing.CurrentInactiveMinutes || existing.InactiveMinutes || "",
        InactiveSessions: dailyRow.InactiveSessions || existing.InactiveSessions || "",
        LongestInactiveMinutes: dailyRow.LongestInactiveMinutes || existing.LongestInactiveMinutes || "",
        FirstInactiveAt: dailyRow.FirstInactiveAt || existing.FirstInactiveAt || existing.InactiveStartAt || "",
        LastInactiveAt: dailyRow.LastInactiveAt || existing.LastInactiveAt || "",
        LastOutputTotal: dailyRow.LastOutputTotal || existing.LastOutputTotal || existing.CurrentTotal || "",
        LastUpdated: dailyRow.LastUpdated || existing.LastUpdated || "",
        __score: Math.max(existing.__score || 0, dailyRow.__score || 0)
      });
    } else {
      latestByKey.set(key, { ...dailyRow, __score: dailyRow.__score || 0 });
    }
  });

  const rows = Array.from(latestByKey.values()).map(row => {
    const copy = { ...row };
    delete copy.__score;
    return copy;
  });

  return rows.length ? rows : dailyRows;
}

function finishInactivityRowIdentityKey_(row) {
  if (!row || typeof row !== "object") return "";

  const explicitKey = getFinishInactivityValue(row, "SnapshotKey", "snapshotKey");
  if (explicitKey) return String(explicitKey).trim().toUpperCase();

  const owner = getFinishInactivityValue(row, "OwnerUsername", "ownerUsername");
  const shift = getFinishInactivityValue(row, "ShiftType", "shiftType");
  const name = getFinishInactivityValue(row, "OperatorName", "operatorName");
  const station = getFinishInactivityValue(row, "StationGroup", "stationGroup");
  const line = getFinishInactivityValue(row, "DefaultLine", "defaultLine");
  const position = getFinishInactivityValue(row, "DefaultPosition", "defaultPosition");

  return [owner, shift, name, station, line, position]
    .map(value => String(value || "").trim().toUpperCase())
    .join("||");
}

function finishInactivityRowFreshnessScore_(row, fallbackIndex = 0) {
  if (!row || typeof row !== "object") return fallbackIndex;

  const possibleDates = [
    getFinishInactivityValue(row, "LastUpdated", "lastUpdated"),
    getFinishInactivityValue(row, "LastCheckedAt", "lastCheckedAt"),
    getFinishInactivityValue(row, "SnapshotAt", "snapshotAt"),
    getFinishInactivityValue(row, "LastOutputAt", "lastOutputAt"),
    getFinishInactivityValue(row, "InactiveStartAt", "inactiveStartAt")
  ];

  for (const value of possibleDates) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) return date.getTime() + fallbackIndex;
  }

  return fallbackIndex;
}

function finishInactivityPayloadHasRows(payload) {
  return finishInactivityBuildRowsFromPayload(payload).length > 0;
}

function finishInactivityApplyPayload(payload, options = {}) {
  const allowClear = !!options.allowClear;
  const updateLastSyncOnly = !!options.updateLastSyncOnly;
  const rows = finishInactivityBuildRowsFromPayload(payload);
  const hasRows = rows.length > 0;
  const alreadyHasRows = Array.isArray(finishMorningInactivityState.rows) && finishMorningInactivityState.rows.length > 0;

  if (!hasRows && alreadyHasRows && !allowClear) {
    console.warn("Finish inactivity refresh returned zero rows. Keeping current display.", payload);
    if (payload?.lastSync) {
      finishMorningInactivityState.lastSync = payload.lastSync;
      setText("finishInactivityLastSync", formatFinishInactivityTime(payload.lastSync));
    }
    return false;
  }

  if (updateLastSyncOnly && !hasRows) {
    if (payload?.lastSync) {
      finishMorningInactivityState.lastSync = payload.lastSync;
      setText("finishInactivityLastSync", formatFinishInactivityTime(payload.lastSync));
    }
    return false;
  }

  finishMorningInactivityState.latestRows = finishInactivityExtractLatestRows(payload);
  finishMorningInactivityState.logRows = finishInactivityExtractLogRows(payload);
  finishMorningInactivityState.dailyRows = finishInactivityExtractDailyRows(payload);
  finishMorningInactivityState.rows = rows;
  finishMorningInactivityState.lastSync = payload?.lastSync || payload?.generatedAt || new Date().toISOString();
  renderFinishMorningInactivityPanel();
  return true;
}

function ensureFinishMorningInactivityPanel() {
  if (document.getElementById("finishMorningInactivityPanel")) return;

  const anchor = document.getElementById("finishInactivityTabMount");
  if (!anchor) return;

  const panel = document.createElement("section");
  panel.id = "finishMorningInactivityPanel";
  panel.className = "finish-inactivity-panel associate-inactivity-panel";
  panel.innerHTML = `
    <header class="finish-inactivity-header associate-inactivity-header">
      <div>
        <span class="finish-inactivity-eyebrow">Associate Activity Watch</span>
        <h3>Inactivity Log</h3>
        <p>Tracks seated Mounting, Final Inspection, and Drill associates. Click any associate to open their KPI profile.</p>
      </div>
      <div class="finish-inactivity-tools finish-inactivity-actions">
        <label>
          X Snapshots
          <input id="finishInactivityThreshold" type="number" min="1" max="12" step="1" value="3" readonly />
        </label>
        <button id="finishInactivityAutoToggleBtn" class="finish-inactivity-auto-toggle is-off" type="button">Auto Snapshot: OFF</button>
        <button id="finishRunInactivitySnapshotBtn" type="button">Run Snapshot Now</button>
        <button id="finishRefreshInactivityLogBtn" type="button">Refresh Log</button>
      </div>
    </header>

    <div class="finish-inactivity-scoreboard associate-inactivity-scoreboard">
      <article><span>Tracked</span><strong id="finishInactivityTrackedCount">0</strong><small>Seated assignments</small></article>
      <article class="watch"><span>Idle Watch</span><strong id="finishInactivityWatchCount">0</strong><small>No output below threshold</small></article>
      <article class="danger"><span>Inactive</span><strong id="finishInactivityInactiveCount">0</strong><small>Threshold reached</small></article>
      <article><span>Last Sync</span><strong id="finishInactivityLastSync">--</strong><small>Frontend refresh</small></article>
    </div>

    <div class="associate-inactivity-layout">
      <div class="associate-inactivity-list-wrap">
        <div class="associate-inactivity-list-head">
          <div>
            <span>Live Associate List</span>
            <strong>Click a row to open profile KPIs</strong>
          </div>
        </div>
        <div id="finishInactivityList" class="finish-inactivity-list associate-inactivity-list">
          <article class="finish-inactivity-empty">No inactivity snapshots loaded yet. Run a snapshot to start tracking.</article>
        </div>
      </div>

      <aside id="finishInactivityProfilePanel" class="associate-inactivity-profile-panel">
        <div class="associate-profile-empty">
          <span>ASSOCIATE PROFILE</span>
          <h3>Select an associate</h3>
          <p>Open a profile to review active, idle, and inactive time building through the day.</p>
        </div>
      </aside>
    </div>
  `;

  anchor.innerHTML = "";
  anchor.appendChild(panel);
  wireFinishInactivityButtons();
}

function wireFinishInactivityButtons() {
  const autoBtn = document.getElementById("finishInactivityAutoToggleBtn");
  const runBtn = document.getElementById("finishRunInactivitySnapshotBtn");
  const refreshBtn = document.getElementById("finishRefreshInactivityLogBtn");

  if (autoBtn && !autoBtn.dataset.wired) {
    autoBtn.dataset.wired = "true";
    autoBtn.addEventListener("click", toggleFinishInactivityAutomationFromPage);
  }

  if (runBtn && !runBtn.dataset.wired) {
    runBtn.dataset.wired = "true";
    runBtn.addEventListener("click", runFinishMorningInactivitySnapshot);
  }

  if (refreshBtn && !refreshBtn.dataset.wired) {
    refreshBtn.dataset.wired = "true";
    refreshBtn.addEventListener("click", loadFinishMorningInactivityLog);
  }
}

async function refreshFinishInactivityAutomationStatus() {
  ensureFinishMorningInactivityPanel();
  const btn = document.getElementById("finishInactivityAutoToggleBtn");

  try {
    const payload = await fetchFinishInactivityApiSafe("getfinishinactivityautomationstatus", {});
    finishMorningInactivityState.automationEnabled = !!payload.enabled;

    if (btn) {
      btn.dataset.enabled = payload.enabled ? "true" : "false";
      btn.textContent = payload.enabled ? "Auto Snapshot: ON" : "Auto Snapshot: OFF";
      btn.classList.toggle("is-on", !!payload.enabled);
      btn.classList.toggle("is-off", !payload.enabled);
    }
  } catch (error) {
    console.warn("Finish inactivity automation status could not load:", error);
    if (btn) {
      btn.dataset.enabled = "false";
      btn.textContent = "Auto Snapshot: UNKNOWN";
      btn.classList.add("is-off");
      btn.classList.remove("is-on");
    }
  }
}

async function toggleFinishInactivityAutomationFromPage() {
  const btn = document.getElementById("finishInactivityAutoToggleBtn");
  const enabled = btn?.dataset.enabled === "true";

  if (btn) {
    btn.disabled = true;
    btn.textContent = enabled ? "Turning Off..." : "Turning On...";
  }

  try {
    await fetchFinishInactivityApiSafe(enabled ? "disablefinishinactivityautomation" : "enablefinishinactivityautomation", {});
    await refreshFinishInactivityAutomationStatus();
  } catch (error) {
    console.error("Finish inactivity automation toggle failed:", error);
    showFinishAssignmentToast?.(error.message || String(error));
  } finally {
    if (btn) btn.disabled = false;
  }
}

function formatFinishInactivityTime(value) {
  if (!value) return "--";
  const text = String(value);
  if (/^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(text)) return text.toUpperCase();
  const date = new Date(value);
  if (isNaN(date.getTime())) return text;
  return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}


function finishInactivityNumberValue(row, ...keys) {
  const raw = getFinishInactivityValue(row, ...keys);
  const num = Number(raw);
  return Number.isFinite(num) ? num : 0;
}

function finishInactivityMinutesSince(value) {
  if (!value) return 0;

  const text = String(value || "").trim();

  // Google Sheets display values sometimes return only "6/4/2026".
  // A date-only value cannot be used for live minute math without creating huge false totals.
  const hasTimeSignal = /\d{1,2}:\d{2}/.test(text) || /T\d{2}:\d{2}/.test(text);
  if (!hasTimeSignal) return 0;

  const date = new Date(text);
  if (isNaN(date.getTime())) return 0;

  return Math.max(0, Math.round((Date.now() - date.getTime()) / 60000));
}

function getFinishInactiveMinutesForRow(row) {
  const status = normalizeFinishInactivityStatus(row);

  const currentInactive = finishInactivityNumberValue(
    row,
    "CurrentInactiveMinutes",
    "currentInactiveMinutes"
  );

  const totalInactive = finishInactivityNumberValue(
    row,
    "TotalInactiveMinutes",
    "totalInactiveMinutes"
  );

  const eventInactive = finishInactivityNumberValue(
    row,
    "InactiveMinutes",
    "inactiveMinutes"
  );

  const openSessionMinutes = finishInactivityMinutesSince(
    getFinishInactivityValue(
      row,
      "InactiveStartAt",
      "inactiveStartAt",
      "FirstInactiveAt",
      "firstInactiveAt"
    )
  );

  // OPEN = current session is still running, so show the largest live/stored value.
  if (status === "OPEN") {
    return Math.max(currentInactive, totalInactive, eventInactive, openSessionMinutes, 0);
  }

  // RESOLVED = current inactive should be 0, but total inactive for the closed session/day must stay visible.
  if (status === "RESOLVED") {
    return Math.max(totalInactive, eventInactive, currentInactive, 0);
  }

  // WATCH = not threshold-inactive yet; keep this as idle/watch time, not inactive time.
  if (status === "WATCH") {
    return 0;
  }

  return Math.max(totalInactive, eventInactive, currentInactive, 0);
}

function getFinishIdleMinutesForRow(row) {
  const status = normalizeFinishInactivityStatus(row);
  if (status !== "WATCH") return 0;

  return Math.max(
    finishInactivityNumberValue(row, "InactiveMinutes", "inactiveMinutes"),
    finishInactivityNumberValue(row, "CurrentInactiveMinutes", "currentInactiveMinutes"),
    0
  );
}

function getFinishInactivityElapsedProductiveMinutes() {
  const shiftType = String(getFinishInactivityShiftTypeSafe() || "Weekday").toLowerCase();
  const rule = shiftType.includes("weekend") ? SHIFT_RULES.weekend : SHIFT_RULES.weekday;
  if (!rule) return 0;

  const now = new Date();
  const startMinutes = timeToMinutes(rule.shiftStart || "7:00 AM");
  const endMinutes = timeToMinutes(rule.shiftEnd || "5:30 PM");
  const currentMinutes = now.getHours() * 60 + now.getMinutes();
  const productiveEnd = Math.max(startMinutes, Math.min(currentMinutes, endMinutes));
  if (productiveEnd <= startMinutes) return 0;

  let elapsed = productiveEnd - startMinutes;
  (rule.breaks || []).forEach(breakItem => {
    const breakStart = timeToMinutes(breakItem.start);
    const breakEnd = timeToMinutes(breakItem.end);
    const overlapStart = Math.max(startMinutes, breakStart);
    const overlapEnd = Math.min(productiveEnd, breakEnd);
    if (overlapEnd > overlapStart) elapsed -= (overlapEnd - overlapStart);
  });

  return Math.max(0, elapsed);
}

function buildFinishAssociateTimeKpis(row) {
  const status = normalizeFinishInactivityStatus(row);
  const elapsed = getFinishInactivityElapsedProductiveMinutes();
  // Inactive total must come from the same session history displayed in the profile.
  // Do NOT use daily summary TotalInactiveMinutes here because it can include merged/older duplicate totals.
  const inactive = getFinishAssociateInactiveSessionTotalMinutes(row);
  const idle = getFinishIdleMinutesForRow(row);

  // Do NOT calculate active as elapsed - inactive anymore.
  // That made associates with strong output look like 0 active once inactive time built up.
  // Active now means "hours/buckets where work was produced" from RAW_ACTIVITY_CURRENT hourly output.
  const hourlyOutput = getFinishHourlyOutputForAssociate(row);
  let activeFromOutput = calculateFinishProducedWindowMinutes(hourlyOutput);

  // Fallback: if API has not returned hourly buckets yet but the associate has output,
  // show at least one produced window instead of misleading 0m.
  // This is still labeled Output Hours, not true active minutes.
  const currentTotal = finishInactivityNumberValue(row, "CurrentTotal", "currentTotal", "LastOutputTotal", "lastOutputTotal");
  const lastOutputAt = getFinishInactivityValue(row, "LastOutputAt", "lastOutputAt", "LastCheckedAt", "SnapshotAt");
  if (activeFromOutput <= 0 && currentTotal > 0) {
    activeFromOutput = lastOutputAt ? 60 : 0;
  }

  const active = Math.max(0, activeFromOutput);

  const sessionHistoryCount = getFinishAssociateInactiveSessions(row).length;

  return {
    elapsed,
    active,
    idle,
    inactive,
    hourlyOutput,
    // Use the real visible session-history count, not the daily summary counter.
    // The daily summary can include older/merged rows or duplicate owner/profile totals.
    sessions: sessionHistoryCount || (inactive > 0 ? 1 : 0),
    status
  };
}

function formatFinishMinutesLabel(minutes) {
  const value = Math.max(0, Number(minutes || 0) || 0);
  if (value < 60) return `${Math.round(value)}m`;
  const h = Math.floor(value / 60);
  const m = Math.round(value % 60);
  return m ? `${h}h ${m}m` : `${h}h`;
}


const FINISH_INACTIVITY_HOUR_ORDER = [
  "6:00 AM", "7:00 AM", "8:00 AM", "9:00 AM", "10:00 AM",
  "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM",
  "4:00 PM", "5:00 PM", "6:00 PM", "7:00 PM", "8:00 PM"
];

function parseFinishHourlyOutputValue(value) {
  if (!value) return {};
  if (typeof value === "object") return value;
  try {
    const parsed = JSON.parse(String(value || ""));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    return {};
  }
}

function getFinishHourlyOutputForAssociate(row) {
  const direct = parseFinishHourlyOutputValue(
    getFinishInactivityValue(row, "HourlyOutput", "LiveHourlyOutput", "hourlyOutput", "liveHourlyOutput", "Hours", "hours")
  );

  if (Object.keys(direct).length) return direct;

  const targetKey = finishInactivityRowIdentityKey_(row);
  const sources = [];
  if (Array.isArray(finishMorningInactivityState.latestRows)) sources.push(...finishMorningInactivityState.latestRows);
  if (Array.isArray(finishMorningInactivityState.logRows)) sources.push(...finishMorningInactivityState.logRows);
  if (Array.isArray(finishMorningInactivityState.rows)) sources.push(...finishMorningInactivityState.rows);

  for (let i = sources.length - 1; i >= 0; i--) {
    const item = sources[i];
    if (!item || typeof item !== "object") continue;
    if (finishInactivityRowIdentityKey_(item) !== targetKey) continue;
    const hours = parseFinishHourlyOutputValue(
      getFinishInactivityValue(item, "HourlyOutput", "LiveHourlyOutput", "hourlyOutput", "liveHourlyOutput", "Hours", "hours")
    );
    if (Object.keys(hours).length) return hours;
  }

  return {};
}

function finishHourLabelToStartMinutes(label) {
  const text = String(label || "").trim().toUpperCase();
  const match = text.match(/^(\d{1,2})(?::\d{2})?\s*(AM|PM)$/);
  if (!match) return null;
  let hour = Number(match[1]);
  const suffix = match[2];
  if (suffix === "PM" && hour !== 12) hour += 12;
  if (suffix === "AM" && hour === 12) hour = 0;
  return hour * 60;
}

function calculateFinishProducedWindowMinutes(hourlyOutput) {
  const hours = hourlyOutput || {};
  const now = new Date();
  const nowMinutes = now.getHours() * 60 + now.getMinutes();
  let minutes = 0;

  FINISH_INACTIVITY_HOUR_ORDER.forEach(hour => {
    const output = Number(hours[hour] || 0) || 0;
    if (output <= 0) return;
    const start = finishHourLabelToStartMinutes(hour);
    if (start == null) return;
    const end = start + 60;
    if (nowMinutes >= end) minutes += 60;
    else if (nowMinutes > start) minutes += Math.max(0, nowMinutes - start);
  });

  return minutes;
}

function buildFinishHourlyActivityVsInactiveRows(row) {
  const hourlyOutput = getFinishHourlyOutputForAssociate(row);
  const sessions = getFinishAssociateInactiveSessions(row);
  const now = new Date();

  return FINISH_INACTIVITY_HOUR_ORDER.map(hour => {
    const output = Number(hourlyOutput[hour] || 0) || 0;
    const hourStartMinutes = finishHourLabelToStartMinutes(hour);
    if (hourStartMinutes == null) {
      return { hour, output, inactiveMinutes: 0, status: output > 0 ? "PRODUCTIVE" : "NO DATA" };
    }

    const base = new Date();
    base.setHours(0, 0, 0, 0);
    const hourStart = new Date(base.getTime() + hourStartMinutes * 60000);
    const hourEnd = new Date(hourStart.getTime() + 60 * 60000);

    let inactiveMinutes = 0;
    sessions.forEach(session => {
      const start = finishInactivityParseDateSafe(session.inactiveStart);
      const end = finishInactivityParseDateSafe(session.inactiveEnd || session.lastChecked) || now;
      if (!start || !end) return;
      const overlapStart = Math.max(start.getTime(), hourStart.getTime());
      const overlapEnd = Math.min(end.getTime(), hourEnd.getTime(), now.getTime());
      if (overlapEnd > overlapStart) inactiveMinutes += Math.round((overlapEnd - overlapStart) / 60000);
    });

    let status = "NO OUTPUT";
    if (output > 0 && inactiveMinutes > 0) status = "MIXED";
    else if (output > 0) status = "PRODUCED";
    else if (inactiveMinutes > 0) status = "INACTIVE";

    return { hour, output, inactiveMinutes, status };
  }).filter(item => item.output > 0 || item.inactiveMinutes > 0);
}

function renderFinishHourlyActivityVsInactive(row) {
  const rows = buildFinishHourlyActivityVsInactiveRows(row);

  if (!rows.length) {
    return `
      <section class="associate-hourly-activity">
        <div class="associate-hourly-head">
          <span>Hourly Output vs Inactive Time</span>
          <strong>No hourly output or inactive session detail yet</strong>
        </div>
      </section>
    `;
  }

  return `
    <section class="associate-hourly-activity">
      <div class="associate-hourly-head">
        <span>Hourly Output vs Inactive Time</span>
        <strong>Produced jobs compared with inactive minutes</strong>
      </div>
      <div class="associate-hourly-grid">
        <div class="associate-hourly-header">Hour</div>
        <div class="associate-hourly-header">Produced</div>
        <div class="associate-hourly-header">Inactive</div>
        <div class="associate-hourly-header">Status</div>
        ${rows.map(item => `
          <div>${escapeHtml(item.hour.replace(":00", ""))}</div>
          <div><strong>${escapeHtml(item.output)}</strong> jobs</div>
          <div><strong>${escapeHtml(formatFinishMinutesLabel(item.inactiveMinutes))}</strong></div>
          <div><span class="associate-hourly-status status-${escapeHtml(String(item.status).toLowerCase().replace(/\s+/g, "-"))}">${escapeHtml(item.status)}</span></div>
        `).join("")}
      </div>
    </section>
  `;
}


function finishInactivityParseDateSafe(value) {
  if (!value) return null;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function finishInactivityMinutesBetweenSafe(startValue, endValue) {
  const start = finishInactivityParseDateSafe(startValue);
  const end = finishInactivityParseDateSafe(endValue) || new Date();
  if (!start || !end) return 0;
  return Math.max(0, Math.round((end.getTime() - start.getTime()) / 60000));
}

function getFinishAssociateInactiveSessions(row) {
  if (!row || typeof row !== "object") return [];

  const targetKey = finishInactivityRowIdentityKey_(row);
  if (!targetKey) return [];

  const sourceRows = Array.isArray(finishMorningInactivityState.logRows)
    ? finishMorningInactivityState.logRows
    : [];

  const sessionMap = new Map();

  function getSessionEndValue(item) {
    return getFinishInactivityValue(
      item,
      "InactiveEndAt",
      "ResolvedAt",
      "inactiveEndAt",
      "resolvedAt"
    );
  }

  function getSessionLastCheckedValue(item) {
    return getFinishInactivityValue(
      item,
      "LastCheckedAt",
      "LastUpdated",
      "SnapshotAt",
      "lastCheckedAt",
      "lastUpdated",
      "snapshotAt"
    );
  }

  function getSessionDateMs(value) {
    const d = finishInactivityParseDateSafe(value);
    return d ? d.getTime() : 0;
  }

  function isRealSessionRow(item) {
    if (!item || typeof item !== "object") return false;

    const sessionId = String(getFinishInactivityValue(item, "SessionID", "sessionId") || "").trim();
    const inactiveStart = getFinishInactivityValue(
      item,
      "InactiveStartAt",
      "inactiveStartAt"
    );

    /*
      Strict rule:
      Session History must only use real inactive-session rows.
      Normal snapshot rows can be ACTIVE/WATCH/INACTIVE and can carry InactiveStartAt,
      but they do not represent a completed/open session unless SessionID exists.
      This prevents mixed/overlapping fake sessions.
    */
    return !!sessionId && !!inactiveStart;
  }

  function buildSessionFromGroup(rows, groupIndex) {
    const cleanRows = rows.filter(Boolean);
    if (!cleanRows.length) return null;

    const first = cleanRows[0];
    const sessionId = String(getFinishInactivityValue(first, "SessionID", "sessionId") || "").trim();
    if (!sessionId) return null;

    const starts = cleanRows
      .map(r => getFinishInactivityValue(r, "InactiveStartAt", "inactiveStartAt"))
      .filter(Boolean)
      .sort((a, b) => getSessionDateMs(a) - getSessionDateMs(b));

    const inactiveStart = starts[0] || "";
    if (!inactiveStart) return null;

    const resolvedRows = cleanRows.filter(r => {
      const st = normalizeFinishInactivityStatus(r);
      return st === "RESOLVED" || !!getSessionEndValue(r);
    });

    const status = resolvedRows.length ? "RESOLVED" : normalizeFinishInactivityStatus(cleanRows[cleanRows.length - 1]) || "OPEN";

    const latestRow = cleanRows
      .slice()
      .sort((a, b) => finishInactivityRowFreshnessScore_(b, 0) - finishInactivityRowFreshnessScore_(a, 0))[0];

    const bestEndRow = (resolvedRows.length ? resolvedRows : cleanRows)
      .slice()
      .sort((a, b) => {
        const aEnd = getSessionDateMs(getSessionEndValue(a) || getSessionLastCheckedValue(a));
        const bEnd = getSessionDateMs(getSessionEndValue(b) || getSessionLastCheckedValue(b));
        return bEnd - aEnd;
      })[0];

    const inactiveEnd = status === "RESOLVED"
      ? (getSessionEndValue(bestEndRow) || getSessionLastCheckedValue(bestEndRow) || "")
      : "";

    const lastChecked = getSessionLastCheckedValue(bestEndRow) || getSessionLastCheckedValue(latestRow) || "";
    const endForDuration = inactiveEnd || lastChecked || new Date();
    const timestampDuration = finishInactivityMinutesBetweenSafe(inactiveStart, endForDuration);

    const fallbackDuration = Math.max.apply(null, cleanRows.map(r => {
      const n = finishInactivityNumberValue(r, "InactiveMinutes", "inactiveMinutes");
      return Number.isFinite(n) ? n : 0;
    }).concat([0]));

    const durationMinutes = timestampDuration > 0 ? timestampDuration : fallbackDuration;

    const startTotals = cleanRows
      .map(r => getFinishInactivityValue(r, "StartTotal", "startTotal", "PreviousTotal", "previousTotal"))
      .filter(v => v !== "" && v !== null && v !== undefined);

    const endTotals = cleanRows
      .map(r => getFinishInactivityValue(r, "EndTotal", "endTotal", "CurrentTotal", "currentTotal", "LastOutputTotal", "lastOutputTotal"))
      .filter(v => v !== "" && v !== null && v !== undefined);

    return {
      sessionId,
      status,
      inactiveStart,
      inactiveEnd,
      startTotal: startTotals.length ? startTotals[0] : "--",
      endTotal: endTotals.length ? endTotals[endTotals.length - 1] : "--",
      lastChecked,
      durationMinutes: Math.max(0, Math.round(Number(durationMinutes) || 0)),
      notes: getFinishInactivityValue(bestEndRow, "Notes", "notes") || getFinishInactivityValue(latestRow, "Notes", "notes") || "",
      score: finishInactivityRowFreshnessScore_(latestRow, groupIndex)
    };
  }

  sourceRows.forEach((item, index) => {
    if (!isRealSessionRow(item)) return;
    if (finishInactivityRowIdentityKey_(item) !== targetKey) return;

    const sessionId = String(getFinishInactivityValue(item, "SessionID", "sessionId") || "").trim();
    if (!sessionMap.has(sessionId)) sessionMap.set(sessionId, []);
    sessionMap.get(sessionId).push(item);
  });

  // Include the clicked row only if it is also a real session row.
  if (isRealSessionRow(row) && finishInactivityRowIdentityKey_(row) === targetKey) {
    const selectedSessionId = String(getFinishInactivityValue(row, "SessionID", "sessionId") || "").trim();
    if (!sessionMap.has(selectedSessionId)) sessionMap.set(selectedSessionId, []);
    sessionMap.get(selectedSessionId).push(row);
  }

  let sessions = Array.from(sessionMap.values())
    .map((rows, index) => buildSessionFromGroup(rows, index))
    .filter(session => session && session.inactiveStart)
    .sort((a, b) => getSessionDateMs(a.inactiveStart) - getSessionDateMs(b.inactiveStart));

  /*
    Repair display-only overlap:
    Old versions left stale OPEN sessions that span across later resolved sessions.
    Even after backend repair, historical rows can remain. A real associate/station
    cannot have overlapping inactive sessions, so remove stale spanning records from
    the visible session history.
  */
  sessions = sessions.filter((session, index, all) => {
    const sStart = getSessionDateMs(session.inactiveStart);
    const sEnd = getSessionDateMs(session.inactiveEnd || session.lastChecked) || sStart;
    if (!sStart || !sEnd) return true;

    const containedLaterSessions = all.filter((other, otherIndex) => {
      if (otherIndex === index) return false;
      const oStart = getSessionDateMs(other.inactiveStart);
      const oEnd = getSessionDateMs(other.inactiveEnd || other.lastChecked) || oStart;
      return oStart > sStart && oEnd <= sEnd;
    });

    const hasLaterResolvedAfterStart = all.some((other, otherIndex) => {
      if (otherIndex === index) return false;
      const oStart = getSessionDateMs(other.inactiveStart);
      const oEnd = getSessionDateMs(other.inactiveEnd || other.lastChecked) || oStart;
      return String(other.status).toUpperCase() === "RESOLVED" && oStart > sStart && oEnd > sStart;
    });

    const isOpenLike = ["OPEN", "INACTIVE", "WATCH"].includes(String(session.status || "").toUpperCase());
    const duration = Number(session.durationMinutes) || 0;

    // Drop old OPEN rows if later resolved sessions exist after they started.
    if (isOpenLike && hasLaterResolvedAfterStart) return false;

    // Drop stale long sessions that swallow multiple smaller real sessions.
    if (containedLaterSessions.length >= 2 && duration > 60) return false;

    return true;
  });

  // Final chronological order: Session 1 = earliest, Session N = latest.
  return sessions.sort((a, b) => getSessionDateMs(a.inactiveStart) - getSessionDateMs(b.inactiveStart));
}
function getFinishAssociateInactiveSessionTotalMinutes(row) {
  const sessions = getFinishAssociateInactiveSessions(row);
  return sessions.reduce((sum, session) => {
    return sum + (Number(session.durationMinutes) || 0);
  }, 0);
}

function renderFinishAssociateSessionHistory(row) {
  const sessions = getFinishAssociateInactiveSessions(row);

  if (!sessions.length) {
    return `
      <section class="associate-session-history">
        <div class="associate-session-history-head">
          <span>Session History</span>
          <strong>No inactive sessions recorded yet</strong>
        </div>
        <article class="associate-session-empty">Once this associate goes inactive and returns active, each session will show here.</article>
      </section>
    `;
  }

  return `
    <section class="associate-session-history">
      <div class="associate-session-history-head">
        <span>Session History</span>
        <strong>${sessions.length} inactive session${sessions.length === 1 ? "" : "s"} today</strong>
      </div>
      <div class="associate-session-list">
        ${sessions.map((session, index) => {
          const status = String(session.status || "OPEN").toUpperCase();
          const isOpen = status === "OPEN" || status === "INACTIVE";
          const start = formatFinishInactivityTime(session.inactiveStart) || "--";
          const end = isOpen ? "Still Open" : (formatFinishInactivityTime(session.inactiveEnd || session.lastChecked) || "--");
          const duration = formatFinishMinutesLabel(session.durationMinutes || 0);
          const output = `${session.startTotal || "--"} → ${session.endTotal || "--"}`;

          return `
            <article class="associate-session-card status-${escapeHtml(String(status).toLowerCase())}">
              <div class="associate-session-title">
                <b>Session ${index + 1}</b>
                <span>${escapeHtml(status)}</span>
              </div>
              <div class="associate-session-grid">
                <div><small>Start</small><strong>${escapeHtml(start)}</strong></div>
                <div><small>End</small><strong>${escapeHtml(end)}</strong></div>
                <div><small>Duration</small><strong>${escapeHtml(duration)}</strong></div>
                <div><small>Output</small><strong>${escapeHtml(output)}</strong></div>
              </div>
              ${session.notes ? `<p>${escapeHtml(session.notes)}</p>` : ""}
            </article>
          `;
        }).join("")}
      </div>
    </section>
  `;
}

function openFinishInactivityAssociateProfile(row) {
  const panel = document.getElementById("finishInactivityProfilePanel");
  if (!panel || !row) return;

  const kpi = buildFinishAssociateTimeKpis(row);
  const operatorName = getFinishInactivityValue(row, "OperatorName", "operatorName") || "Unknown";
  const station = normalizeFinishInactivityStation(getFinishInactivityValue(row, "StationGroup", "stationGroup") || getFinishInactivityValue(row, "RoleType", "roleType"));
  const roleType = getFinishInactivityValue(row, "RoleType", "roleType") || station || "Role";
  const defaultLine = getFinishInactivityValue(row, "DefaultLine", "defaultLine") || "--";
  const defaultPosition = getFinishInactivityValue(row, "DefaultPosition", "defaultPosition") || "--";
  const current = getFinishInactivityValue(row, "CurrentTotal", "currentTotal", "LastOutputTotal", "lastOutputTotal") || "0";
  const previous = getFinishInactivityValue(row, "PreviousTotal", "previousTotal", "StartTotal", "startTotal") || "--";
  const noChange = getFinishInactivityValue(row, "NoChangeSnapshots", "noChangeSnapshots") || "--";
  const firstInactive = getFinishInactivityValue(row, "FirstInactiveAt", "InactiveStartAt", "inactiveStartAt") || "--";
  const lastInactive = getFinishInactivityValue(row, "LastInactiveAt", "LastCheckedAt", "LastUpdated", "lastUpdated") || "--";
  const lastOutputAt = getFinishInactivityValue(row, "LastOutputAt", "lastOutputAt") || "--";
  const minutesSinceLastOutput = getFinishInactivityValue(row, "MinutesSinceLastOutput", "minutesSinceLastOutput") || "--";

  const elapsed = Math.max(1, kpi.elapsed || 1);
  const activePct = Math.min(100, Math.round((kpi.active / elapsed) * 100));
  const idlePct = Math.min(100, Math.round((kpi.idle / elapsed) * 100));
  const inactivePct = Math.min(100, Math.round((kpi.inactive / elapsed) * 100));

  panel.innerHTML = `
    <div class="associate-profile-card status-${escapeHtml(String(kpi.status).toLowerCase())}">
      <div class="associate-profile-top">
        <span>ASSOCIATE PROFILE</span>
        <h3>${escapeHtml(operatorName)}</h3>
        <p>${escapeHtml(station)} · ${escapeHtml(defaultLine)} · ${escapeHtml(defaultPosition)} · ${escapeHtml(roleType)}</p>
        <b>${escapeHtml(kpi.status)}</b>
      </div>

      <div class="associate-profile-kpis">
        <article class="active"><span>Output Hours</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.active))}</strong><small>Hours with production</small></article>
        <article class="idle"><span>Idle</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.idle))}</strong><small>No output below threshold</small></article>
        <article class="inactive"><span>Inactive</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.inactive))}</strong><small>Threshold reached</small></article>
      </div>

      <div class="associate-profile-bar" aria-label="Active idle inactive time split">
        <i class="active" style="width:${activePct}%"></i>
        <i class="idle" style="width:${idlePct}%"></i>
        <i class="inactive" style="width:${inactivePct}%"></i>
      </div>

      <div class="associate-profile-metrics">
        <div><span>Current Output</span><strong>${escapeHtml(current)}</strong></div>
        <div><span>Previous Output</span><strong>${escapeHtml(previous)}</strong></div>
        <div><span>No Change Checks</span><strong>${escapeHtml(noChange)}</strong></div>
        <div><span>Inactive Sessions</span><strong>${escapeHtml(kpi.sessions)}</strong></div>
        <div><span>Total Inactive</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.inactive))}</strong></div>
        <div><span>Last Output</span><strong>${escapeHtml(formatFinishInactivityTime(lastOutputAt))}</strong></div>
        <div><span>Minutes Since Output</span><strong>${escapeHtml(minutesSinceLastOutput === "--" ? "--" : formatFinishMinutesLabel(minutesSinceLastOutput))}</strong></div>
        <div><span>First Inactive</span><strong>${escapeHtml(formatFinishInactivityTime(firstInactive))}</strong></div>
        <div><span>Last Inactive</span><strong>${escapeHtml(formatFinishInactivityTime(lastInactive))}</strong></div>
      </div>

      ${renderFinishHourlyActivityVsInactive(row)}

      ${renderFinishAssociateSessionHistory(row)}

      <p class="associate-profile-note">Each session shows when the associate crossed the inactive threshold and when output movement resolved it. Output Hours counts hourly buckets where production exists; inactive total is the sum of the visible session-history durations from the inactivity log. True active minutes require individual scan timestamps.</p>
    </div>
  `;
}

function renderFinishMorningInactivityPanel() {
  ensureFinishMorningInactivityPanel();

  const rows = getFinishMorningInactivityRows().filter(row => row && typeof row === "object");

  const watchCount = rows.filter(row => normalizeFinishInactivityStatus(row) === "WATCH").length;
  const inactiveCount = rows.filter(row => normalizeFinishInactivityStatus(row) === "OPEN").length;
  const trackedCount = rows.length;

  setText("finishInactivityTrackedCount", Math.max(0, trackedCount));
  setText("finishInactivityWatchCount", Math.max(0, watchCount));
  setText("finishInactivityInactiveCount", Math.max(0, inactiveCount));
  setText("finishInactivityLastSync", formatFinishInactivityTime(finishMorningInactivityState.lastSync));

  const list = document.getElementById("finishInactivityList");
  if (!list) return;

  if (!rows.length) {
    list.innerHTML = `<article class="finish-inactivity-empty">No inactivity snapshots loaded yet. Click Refresh Log or Run Snapshot Now.</article>`;
    return;
  }

  const rowsToShow = rows
    .slice()
    .sort((a, b) => {
      const rank = { OPEN: 1, INACTIVE: 1, WATCH: 2, RESOLVED: 3, ACTIVE: 4 };
      const aRank = rank[normalizeFinishInactivityStatus(a)] || 9;
      const bRank = rank[normalizeFinishInactivityStatus(b)] || 9;
      if (aRank !== bRank) return aRank - bRank;
      const aMin = getFinishAssociateInactiveSessionTotalMinutes(a) || getFinishIdleMinutesForRow(a);
      const bMin = getFinishAssociateInactiveSessionTotalMinutes(b) || getFinishIdleMinutesForRow(b);
      return bMin - aMin;
    })
    .slice(0, 80);

  list.innerHTML = rowsToShow.map((row, index) => {
    const status = normalizeFinishInactivityStatus(row);
    const operatorName = getFinishInactivityValue(row, "OperatorName", "operatorName") || "Unknown";
    const station = normalizeFinishInactivityStation(getFinishInactivityValue(row, "StationGroup", "stationGroup") || getFinishInactivityValue(row, "RoleType", "roleType"));
    const roleType = getFinishInactivityValue(row, "RoleType", "roleType") || station || "Role";
    const defaultLine = getFinishInactivityValue(row, "DefaultLine", "defaultLine");
    const defaultPosition = getFinishInactivityValue(row, "DefaultPosition", "defaultPosition");
    const position = [defaultLine, defaultPosition].filter(Boolean).join(" · ") || "No position";
    const current = getFinishInactivityValue(row, "CurrentTotal", "currentTotal", "LastOutputTotal", "lastOutputTotal") || "0";
    const previous = getFinishInactivityValue(row, "PreviousTotal", "previousTotal", "StartTotal", "startTotal") || "--";
    const noChange = getFinishInactivityValue(row, "NoChangeSnapshots", "noChangeSnapshots") || "--";
    const kpi = buildFinishAssociateTimeKpis(row);
    const sessions = getFinishAssociateInactiveSessions(row).length || "--";

    return `
      <article class="finish-inactivity-row associate-inactivity-row status-${escapeHtml(String(status).toLowerCase())}" data-finish-inactivity-index="${index}">
        <div class="finish-inactivity-main">
          <strong>${escapeHtml(operatorName)}</strong>
          <span>${escapeHtml(station)} · ${escapeHtml(position)} · ${escapeHtml(roleType)}</span>
        </div>
        <div class="finish-inactivity-metrics">
          <div><span>Current</span><strong>${escapeHtml(current)}</strong></div>
          <div><span>Previous</span><strong>${escapeHtml(previous)}</strong></div>
          <div><span>No Change</span><strong>${escapeHtml(noChange)}</strong></div>
          <div><span>Output Hours</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.active))}</strong></div>
          <div><span>Idle</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.idle))}</strong></div>
          <div><span>Inactive</span><strong>${escapeHtml(formatFinishMinutesLabel(kpi.inactive))}</strong></div>
          <div><span>Sessions</span><strong>${escapeHtml(sessions)}</strong></div>
          <b class="finish-inactivity-status">${escapeHtml(status)}</b>
        </div>
      </article>`;
  }).join("");

  list.querySelectorAll("[data-finish-inactivity-index]").forEach(item => {
    item.addEventListener("click", () => {
      const row = rowsToShow[Number(item.dataset.finishInactivityIndex || 0)];
      if (row) openFinishInactivityAssociateProfile(row);
    });
  });

  const currentProfile = document.getElementById("finishInactivityProfilePanel");
  if (currentProfile && currentProfile.querySelector(".associate-profile-empty") && rowsToShow[0]) {
    openFinishInactivityAssociateProfile(rowsToShow[0]);
  }
}

async function loadFinishMorningInactivityLog() {
  const btn = document.getElementById("finishRefreshInactivityLogBtn");
  const ownerUsername = getFinishInactivityOwnerUsernameSafe();
  const shiftType = getFinishInactivityShiftTypeSafe();

  if (btn) {
    btn.disabled = true;
    btn.textContent = "Refreshing...";
  }

  try {
    const payload = await fetchFinishInactivityApiSafe("getfinishinactivitylog", {
      ownerUsername,
      shiftType,
      todayOnly: true,
      limit: 250
    });

    finishMorningInactivityState.summary = {
      tracked: Number(payload.tracked || payload.count || 0) || 0,
      watch: Number(payload.watch || 0) || 0,
      inactive: Number(payload.inactive || payload.open || 0) || 0,
      open: Number(payload.open || payload.inactive || 0) || 0
    };

    finishInactivityApplyPayload(payload, { allowClear: false });
  } catch (error) {
    console.warn("Finish inactivity log could not load. Keeping current display.", error);
    showFinishAssignmentToast?.(error.message || String(error));
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Refresh Log";
    }
  }
}

async function runFinishMorningInactivitySnapshot() {
  if (finishMorningInactivityState.loading) return;

  const btn = document.getElementById("finishRunInactivitySnapshotBtn");
  const thresholdSnapshots = getFinishMorningInactivityThreshold();
  const ownerUsername = getFinishInactivityOwnerUsernameSafe();
  const shiftType = getFinishInactivityShiftTypeSafe();

  finishMorningInactivityState.loading = true;
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Running...";
  }

  try {
    await fetchFinishInactivityApiSafe("runfinishinactivitysnapshotnow", {
      ownerUsername,
      shiftType,
      thresholdSnapshots,
      appendAll: true
    });

    await loadFinishMorningInactivityLog();

    const inactiveCount = Number(finishMorningInactivityState.summary?.inactive || finishMorningInactivityState.summary?.open || 0) || 0;
    showFinishAssignmentToast?.(`Inactivity snapshot complete · ${inactiveCount} inactive`);
  } catch (error) {
    console.error("Finish inactivity snapshot failed. Refreshing saved log instead.", error);
    showFinishAssignmentToast?.("Snapshot failed, refreshing saved inactivity log.");
    await loadFinishMorningInactivityLog();
  } finally {
    finishMorningInactivityState.loading = false;
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Run Snapshot Now";
    }
  }
}

function startFinishInactivityFrontendAutoRefresh() {
  if (window.finishInactivityAutoRefreshTimerFinal) {
    clearInterval(window.finishInactivityAutoRefreshTimerFinal);
  }

  loadFinishMorningInactivityLog();

  window.finishInactivityAutoRefreshTimerFinal = setInterval(() => {
    loadFinishMorningInactivityLog();
  }, 120000);
}

function resetFinishInactivityButtons() {
  const runBtn = document.getElementById("finishRunInactivitySnapshotBtn");
  const refreshBtn = document.getElementById("finishRefreshInactivityLogBtn");

  if (runBtn) {
    runBtn.disabled = false;
    runBtn.textContent = "Run Snapshot Now";
  }

  if (refreshBtn) {
    refreshBtn.disabled = false;
    refreshBtn.textContent = "Refresh Log";
  }
}

window.runFinishMorningInactivitySnapshot = runFinishMorningInactivitySnapshot;
window.loadFinishMorningInactivityLog = loadFinishMorningInactivityLog;
window.startFinishInactivityFrontendAutoRefresh = startFinishInactivityFrontendAutoRefresh;
window.addEventListener("error", resetFinishInactivityButtons);
window.addEventListener("unhandledrejection", resetFinishInactivityButtons);

setTimeout(() => {
  ensureFinishMorningInactivityPanel();
  wireFinishInactivityButtons();
  refreshFinishInactivityAutomationStatus();
  loadFinishMorningInactivityLog();
  startFinishInactivityFrontendAutoRefresh();
}, 1800);

/* =========================================================
   FINAL OVERRIDE — ASSOCIATE DOWNTIME LOG ADVANCED UI
   Purpose:
   - Uses the separate Finish Scan Activity Tracker API.
   - Keeps current Finish Dashboard HTML; rebuilds only the downtime tab.
   - Adds cleaner list, filters, selected associate profile, gap history,
     multi-scan grouping, and a compact mini graph.
========================================================= */
(function installFinishScanDowntimeAdvancedUI() {
  const API_URL = "https://script.google.com/macros/s/AKfycbx7TAoSeBugvJWbYyRMzNpjgktge9NuTF3oqFYisqUt5RWvYJCGPiRaW21w44Ze8jw_/exec";

  const state = window.finishScanDowntimeState = window.finishScanDowntimeState || {};
  Object.assign(state, {
    loading: false,
    activeAccess: state.activeAccess || "Mounting",
    statusFilter: state.statusFilter || "ALL",
    sortBy: state.sortBy || "currentGap",
    search: state.search || "",
    selectedKey: state.selectedKey || "",
    selectedDetailTab: state.selectedDetailTab || "overview",
    payload: state.payload || null,
    operators: state.operators || [],
    gapLog: state.gapLog || [],
    multiScanEvents: state.multiScanEvents || [],
    downtimeDaily: state.downtimeDaily || [],
    lastSync: state.lastSync || ""
  });

  function esc(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function clean(value) {
    return String(value || "").trim();
  }

  function cleanKey(value) {
    return clean(value).toUpperCase().replace(/\s+/g, " ");
  }

  function cleanAccess(value) {
    const text = clean(value).toLowerCase();

    // Drill comes in through the Mounting side of the scan report,
    // but it must still be recognized as its own access point so we can
    // combine it cleanly with Mounting when needed.
    if (text.includes("mount") && text.includes("drill")) return "Mounting / Drill";
    if (text.includes("final")) return "Final Inspection";
    if (text.includes("drill")) return "Drill";
    if (text.includes("mount")) return "Mounting";

    return clean(value) || "Unknown";
  }

  function operatorName(row) {
    return clean(row.OperatorName || row.operatorName || row.Operator || row.operator || "Unknown Operator");
  }

  function rowAccess(row) {
    return cleanAccess(row.AccessPoint || row.accessPoint || row.ScanStage || row.scanStage || "");
  }

  function rowKey(row) {
    if (row && row.__rowKey) return String(row.__rowKey);
    return `${cleanKey(operatorName(row))}||${cleanKey(rowAccess(row))}`;
  }

  function baseRowKey(row) {
    return `${cleanKey(operatorName(row))}||${cleanKey(rowAccess(row))}`;
  }

  function rowSourceKeys(row) {
    if (row && Array.isArray(row.__sourceKeys) && row.__sourceKeys.length) {
      return row.__sourceKeys.map(String);
    }
    return [baseRowKey(row)];
  }

  function isMountingFamilyAccess(value) {
    const access = cleanAccess(value);
    return access === "Mounting" || access === "Drill" || access === "Mounting / Drill";
  }

  function rowMatchesAccess(row, access) {
    const selected = cleanAccess(access);
    const rowAp = rowAccess(row);

    // The Mounting tab is intentionally Mounting + Drill.
    // If an associate scans both, they should show once as Mounting / Drill.
    if (selected === "Mounting" || selected === "Mounting / Drill") {
      return isMountingFamilyAccess(rowAp);
    }

    return rowAp === selected;
  }

  function statusFromGapMinutes(minutes) {
    const n = Number(minutes) || 0;
    if (n >= 10) return "INACTIVE";
    if (n >= 5) return "WATCH";
    return "ACTIVE";
  }

  function combineMountingDrillRows(rows) {
    const groups = new Map();

    (rows || []).forEach(row => {
      const name = operatorName(row);
      const key = cleanKey(name);
      const access = rowAccess(row);
      const existing = groups.get(key);

      if (!existing) {
        groups.set(key, {
          ...row,
          __rowKey: `${key}||MOUNTING_DRILL`,
          __sourceKeys: [baseRowKey(row)],
          __accessList: [access],
          AccessPoint: access,
          accessPoint: access
        });
        return;
      }

      const currentKeys = new Set(existing.__sourceKeys || []);
      currentKeys.add(baseRowKey(row));
      existing.__sourceKeys = Array.from(currentKeys);

      const accessSet = new Set(existing.__accessList || []);
      accessSet.add(access);
      existing.__accessList = Array.from(accessSet);

      existing.AccessPoint = existing.__accessList.length > 1 ? "Mounting / Drill" : existing.__accessList[0];
      existing.accessPoint = existing.AccessPoint;

      existing.TotalScans = metric(existing, "TotalScans", "totalScans") + metric(row, "TotalScans", "totalScans");
      existing.totalScans = existing.TotalScans;

      existing.ScansThisHour = metric(existing, "ScansThisHour", "scansThisHour", "ThisHour", "thisHour") + metric(row, "ScansThisHour", "scansThisHour", "ThisHour", "thisHour");
      existing.scansThisHour = existing.ScansThisHour;

      existing.MultiScanSeconds = metric(existing, "MultiScanSeconds", "multiScanSeconds") + metric(row, "MultiScanSeconds", "multiScanSeconds");
      existing.multiScanSeconds = existing.MultiScanSeconds;

      existing.LongestGapMinutes = Math.max(metric(existing, "LongestGapMinutes", "longestGapMinutes"), metric(row, "LongestGapMinutes", "longestGapMinutes"));
      existing.longestGapMinutes = existing.LongestGapMinutes;

      const existingAvg = metric(existing, "AverageGapMinutes", "averageGapMinutes");
      const rowAvg = metric(row, "AverageGapMinutes", "averageGapMinutes");
      existing.AverageGapMinutes = Math.round(((existingAvg + rowAvg) / 2) * 100) / 100;
      existing.averageGapMinutes = existing.AverageGapMinutes;

      // Current status is based on the most recent scan across Mounting + Drill.
      // This prevents someone from looking inactive on Drill after they moved to Mounting.
      const existingLast = dateMs(existing.LastScanAt || existing.lastScanAt);
      const rowLast = dateMs(row.LastScanAt || row.lastScanAt);
      if (rowLast > existingLast) {
        existing.LastScanAt = row.LastScanAt || row.lastScanAt;
        existing.lastScanAt = existing.LastScanAt;
        existing.CurrentGapMinutes = metric(row, "CurrentGapMinutes", "currentGapMinutes");
        existing.currentGapMinutes = existing.CurrentGapMinutes;
        existing.CurrentStatus = normalizeStatus(row.CurrentStatus || row.currentStatus || row.Status || row.status);
        existing.currentStatus = existing.CurrentStatus;
      }
    });

    return Array.from(groups.values()).map(row => {
      if ((row.__accessList || []).length > 1) {
        row.AccessPoint = "Mounting / Drill";
        row.accessPoint = "Mounting / Drill";
      }
      const gap = metric(row, "CurrentGapMinutes", "currentGapMinutes");
      row.CurrentStatus = statusFromGapMinutes(gap);
      row.currentStatus = row.CurrentStatus;
      return row;
    });
  }

  function normalizeStatus(value) {
    const status = clean(value).toUpperCase();
    if (status === "INACTIVE" || status === "OPEN") return "INACTIVE";
    if (status === "WATCH" || status === "IDLE") return "WATCH";
    if (status === "RESOLVED") return "ACTIVE";
    return status || "ACTIVE";
  }

  function dateMs(value) {
    if (!value) return 0;
    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value.getTime();
    const date = new Date(value);
    return isNaN(date.getTime()) ? 0 : date.getTime();
  }

  function toDate(value) {
    const ms = dateMs(value);
    return ms ? new Date(ms) : null;
  }

  function formatTime(value) {
    const date = toDate(value);
    if (!date) return "--";
    return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  function formatDateShort(value) {
    const date = toDate(value);
    if (!date) return "--";
    return date.toLocaleDateString([], { month: "numeric", day: "numeric" });
  }

  function formatDateTime(value) {
    const date = toDate(value);
    if (!date) return "--";
    return `${date.toLocaleDateString([], { month: "numeric", day: "numeric" })} ${date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
  }

  function formatMinutes(value) {
    const total = Math.max(0, Math.round(Number(value) || 0));
    const h = Math.floor(total / 60);
    const m = total % 60;
    if (h && m) return `${h}h ${m}m`;
    if (h) return `${h}h`;
    return `${m}m`;
  }

  function numberFmt(value) {
    return Number(value || 0).toLocaleString();
  }

  function setTextSafe(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function metric(row, ...keys) {
    for (const key of keys) {
      const value = row?.[key];
      if (value !== undefined && value !== null && value !== "") return Number(value) || 0;
    }
    return 0;
  }

  async function fetchFinishScanActivityApi(action = "getfinishscanactivity", params = {}) {
    const url = new URL(API_URL);
    url.searchParams.set("action", action);
    Object.entries(params || {}).forEach(([key, value]) => {
      if (value === undefined || value === null || value === "") return;
      url.searchParams.set(key, String(value));
    });

    const response = await fetch(url.toString(), { method: "GET", cache: "no-store" });
    if (!response.ok) throw new Error(`Finish Scan Tracker API failed: ${response.status}`);

    const payload = await response.json();
    const status = clean(payload.status || payload.Status).toLowerCase();
    const success = payload.success === true || payload.Success === true || status === "success" || status === "ok" || (!payload.error && !String(payload.message || "").toLowerCase().includes("error"));
    if (!success && (payload.message || payload.error)) throw new Error(payload.message || payload.error);
    return payload;
  }

  function getPayloadArray(payload, keys) {
    for (const key of keys) {
      if (Array.isArray(payload?.[key])) return payload[key];
      if (Array.isArray(payload?.data?.[key])) return payload.data[key];
      if (Array.isArray(payload?.payload?.[key])) return payload.payload[key];
    }
    return [];
  }

  function applyPayload(payload) {
    state.payload = payload || {};
    state.operators = getPayloadArray(payload, ["operators", "operatorTimeline", "timeline", "rows", "data"]);
    state.gapLog = getPayloadArray(payload, ["gapLog", "gaps", "gapRows"]);
    state.multiScanEvents = getPayloadArray(payload, ["multiScanEvents", "multiScan", "multiRows"]);
    state.downtimeDaily = getPayloadArray(payload, ["downtimeDaily", "daily", "downtimeRows"]);
    state.lastSync = payload?.generatedAt || payload?.lastUpdated || new Date().toISOString();
    renderPanel();
  }

  function dailyForRow(row) {
    const keys = new Set(rowSourceKeys(row));
    const matches = (state.downtimeDaily || []).filter(d => keys.has(baseRowKey(d)));

    if (!matches.length) return {};
    if (matches.length === 1) return matches[0];

    const firstScan = matches
      .map(d => d.FirstScanAt || d.firstScanAt)
      .filter(Boolean)
      .sort((a, b) => dateMs(a) - dateMs(b))[0] || "";

    const lastScan = matches
      .map(d => d.LastScanAt || d.lastScanAt)
      .filter(Boolean)
      .sort((a, b) => dateMs(b) - dateMs(a))[0] || "";

    return {
      ...matches[0],
      AccessPoint: "Mounting / Drill",
      accessPoint: "Mounting / Drill",
      FirstScanAt: firstScan,
      LastScanAt: lastScan,
      TotalScans: matches.reduce((sum, d) => sum + metric(d, "TotalScans", "totalScans"), 0),
      TotalDowntimeMinutes: matches.reduce((sum, d) => sum + metric(d, "TotalDowntimeMinutes", "totalDowntimeMinutes", "InactiveMinutes", "inactiveMinutes"), 0),
      InactiveMinutes: matches.reduce((sum, d) => sum + metric(d, "InactiveMinutes", "inactiveMinutes"), 0),
      WatchMinutes: matches.reduce((sum, d) => sum + metric(d, "WatchMinutes", "watchMinutes"), 0),
      InactiveSessions: matches.reduce((sum, d) => sum + metric(d, "InactiveSessions", "inactiveSessions"), 0),
      MultiScanEvents: matches.reduce((sum, d) => sum + metric(d, "MultiScanEvents", "multiScanEvents"), 0),
      LongestGapMinutes: Math.max(...matches.map(d => metric(d, "LongestGapMinutes", "longestGapMinutes")), 0),
      CurrentGapMinutes: Math.min(...matches.map(d => metric(d, "CurrentGapMinutes", "currentGapMinutes")).filter(n => n >= 0))
    };
  }

  function rowsForAccess(access, options = {}) {
    const key = cleanAccess(access);
    let rows = (state.operators || []).filter(row => rowMatchesAccess(row, key));

    if (key === "Mounting" || key === "Mounting / Drill") {
      rows = combineMountingDrillRows(rows);
    }

    const status = cleanKey(options.statusFilter ?? state.statusFilter ?? "ALL");
    if (status && status !== "ALL") {
      rows = rows.filter(row => normalizeStatus(row.CurrentStatus || row.currentStatus || row.Status || row.status) === status);
    }

    const search = clean(options.search ?? state.search).toLowerCase();
    if (search) {
      rows = rows.filter(row => `${operatorName(row)} ${rowAccess(row)}`.toLowerCase().includes(search));
    }

    const sortBy = options.sortBy || state.sortBy || "currentGap";
    rows.sort((a, b) => {
      const statusOrder = { INACTIVE: 0, WATCH: 1, ACTIVE: 2 };
      if (sortBy === "status") {
        const sa = statusOrder[normalizeStatus(a.CurrentStatus || a.currentStatus)] ?? 9;
        const sb = statusOrder[normalizeStatus(b.CurrentStatus || b.currentStatus)] ?? 9;
        if (sa !== sb) return sa - sb;
      }
      if (sortBy === "longestGap") return metric(b, "LongestGapMinutes", "longestGapMinutes") - metric(a, "LongestGapMinutes", "longestGapMinutes");
      if (sortBy === "multiScan") return metric(b, "MultiScanSeconds", "multiScanSeconds") - metric(a, "MultiScanSeconds", "multiScanSeconds");
      if (sortBy === "totalScans") return metric(b, "TotalScans", "totalScans") - metric(a, "TotalScans", "totalScans");
      if (sortBy === "name") return operatorName(a).localeCompare(operatorName(b));
      const currentGapDiff = metric(b, "CurrentGapMinutes", "currentGapMinutes") - metric(a, "CurrentGapMinutes", "currentGapMinutes");
      if (currentGapDiff !== 0) return currentGapDiff;
      const sa = statusOrder[normalizeStatus(a.CurrentStatus || a.currentStatus)] ?? 9;
      const sb = statusOrder[normalizeStatus(b.CurrentStatus || b.currentStatus)] ?? 9;
      if (sa !== sb) return sa - sb;
      return operatorName(a).localeCompare(operatorName(b));
    });

    return rows;
  }

  function allRowsForAccess(access) {
    const key = cleanAccess(access);
    let rows = (state.operators || []).filter(row => rowMatchesAccess(row, key));
    if (key === "Mounting" || key === "Mounting / Drill") rows = combineMountingDrillRows(rows);
    return rows;
  }

  function accessSummary(access) {
    const rows = allRowsForAccess(access);
    return {
      tracked: rows.length,
      watch: rows.filter(row => normalizeStatus(row.CurrentStatus || row.currentStatus) === "WATCH").length,
      inactive: rows.filter(row => normalizeStatus(row.CurrentStatus || row.currentStatus) === "INACTIVE").length,
      totalScans: rows.reduce((sum, row) => sum + metric(row, "TotalScans", "totalScans"), 0),
      multiScan: rows.reduce((sum, row) => sum + metric(row, "MultiScanSeconds", "multiScanSeconds"), 0),
      downtime: rows.reduce((sum, row) => sum + metric(dailyForRow(row), "TotalDowntimeMinutes", "totalDowntimeMinutes", "InactiveMinutes", "inactiveMinutes"), 0)
    };
  }

  function isRealDisplayGap(g) {
    const minutes = metric(g, "GapMinutes", "gapMinutes");
    const status = normalizeStatus(g.Status || g.status);

    // Do not show 0m / normal ACTIVE rows in the Gap History.
    // The scan report can create same-minute/duplicate timestamp rows; those are not downtime gaps.
    if (minutes < 5) return false;
    if (status === "ACTIVE" && minutes < 5) return false;

    return true;
  }

  function operatorGapRows(row) {
    const keys = new Set(rowSourceKeys(row));
    return (state.gapLog || [])
      .filter(g => keys.has(baseRowKey(g)))
      .filter(isRealDisplayGap)
      .sort((a, b) => dateMs(b.GapStartAt || b.PreviousScanTime) - dateMs(a.GapStartAt || a.PreviousScanTime));
  }

  function operatorMultiRows(row) {
    const keys = new Set(rowSourceKeys(row));
    return (state.multiScanEvents || [])
      .filter(g => keys.has(baseRowKey(g)))
      .sort((a, b) => dateMs(b.ScanTime || b.scanTime) - dateMs(a.ScanTime || a.scanTime));
  }

  function getSelectedRow(rows) {
    if (!rows.length) return null;
    if (state.selectedKey) {
      const found = rows.find(row => rowKey(row) === state.selectedKey);
      if (found) return found;
    }
    state.selectedKey = rowKey(rows[0]);
    return rows[0];
  }

  function cleanupFinishDowntimeSnapshotRuleCards() {
    const mount = document.getElementById("finishInactivityTabMount");
    if (!mount) return;

    // Remove the old top hero/banner completely. The useful content starts at Associate Downtime Log.
    mount.querySelectorAll(".scan-advanced-hero").forEach(el => el.remove());

    // The snapshot rule cards belong to the old ME/snapshot tracker, not the scan-gap downtime view.
    mount.querySelectorAll(".scan-advanced-rule-card").forEach(el => el.remove());

    // Extra protection: remove any remaining top banner/card that says Snapshot Rule / 3 Checks.
    mount.querySelectorAll("header, section, article, div").forEach(el => {
      const text = String(el.textContent || "").replace(/\s+/g, " ").trim().toUpperCase();
      if (!text) return;
      const hasSnapshotRule = text.includes("SNAPSHOT RULE") || text.includes("3 CHECKS");
      const isSafePanel = el.id === "finishScanDowntimePanel" || el.closest("#finishScanDowntimePanel .scan-advanced-panel");
      if (hasSnapshotRule && !isSafePanel) el.remove();
    });

    const advancedPanel = mount.querySelector("#finishScanDowntimePanel");

    // Remove duplicated old downtime hero blocks if the legacy inactivity renderer runs before/after this UI.
    Array.from(mount.children).forEach(child => {
      if (child === advancedPanel) return;
      const text = String(child.textContent || "").replace(/\s+/g, " ").trim().toUpperCase();
      const isOldDowntimeHero = text.includes("ASSOCIATE DOWNTIME TRACKING") && text.includes("SNAPSHOT RULE");
      const isOldInactivityPanel = child.id === "finishMorningInactivityPanel";
      if (isOldDowntimeHero || isOldInactivityPanel) child.remove();
    });
  }

  function statusClass(status) {
    return `status-${normalizeStatus(status).toLowerCase()}`;
  }

  function initials(name) {
    const parts = clean(name).split(/\s+/).filter(Boolean);
    return ((parts[0]?.[0] || "?") + (parts[1]?.[0] || "")).toUpperCase();
  }

  function ensurePanel() {
    const mount = document.getElementById("finishInactivityTabMount");
    if (!mount) return null;

    if (!document.getElementById("finishScanDowntimePanel")) {
      mount.innerHTML = `
        <section id="finishScanDowntimePanel" class="scan-advanced-shell">
          <section class="scan-advanced-panel">
            <div class="scan-advanced-panel-head">
              <div>
                <span class="scan-advanced-eyebrow">Scan Activity Tracker</span>
                <h3>Associate Downtime Log</h3>                
              </div>
              <div class="scan-advanced-actions">
                <div class="scan-rule-card"><span>Watch / Inactive</span><strong>5m / 10m</strong></div>
                <button id="finishRunInactivitySnapshotBtn" type="button">Run Scan Refresh</button>
                <button id="finishRefreshInactivityLogBtn" type="button">Refresh Log</button>
              </div>
            </div>

            <div class="scan-advanced-kpis">
              <article><span>Tracked</span><strong id="finishInactivityTrackedCount">0</strong><small>Operators log on/small></article>
              <article class="watch"><span>Watch</span><strong id="finishInactivityWatchCount">0</strong><small>5+ minutes no scan</small></article>
              <article class="danger"><span>Inactive</span><strong id="finishInactivityInactiveCount">0</strong><small>10+ minutes no scan</small></article>              
            </div>

            <div class="scan-filter-bar">
              <div class="scan-filter-group">
                <span>Access Point</span>
                <div class="scan-segmented">
                  <button class="scan-downtime-tab active" type="button" data-scan-access="Mounting">Mounting / Drill</button>
                  <button class="scan-downtime-tab" type="button" data-scan-access="Final Inspection">Final Inspection</button>
                </div>
              </div>
              <div class="scan-filter-group">
                <span>Status Filter</span>
                <div class="scan-status-filters">
                  <button type="button" data-scan-status="ALL" class="active">All</button>
                  <button type="button" data-scan-status="ACTIVE">Active</button>
                  <button type="button" data-scan-status="WATCH">Watch</button>
                  <button type="button" data-scan-status="INACTIVE">Inactive</button>
                </div>
              </div>
              <label class="scan-filter-field">
                <span>Sort By</span>
                <select id="finishScanSortSelect">
                  <option value="currentGap">Current Gap</option>
                  <option value="longestGap">Longest Gap</option>
                  <option value="multiScan">Most Multi-Scan</option>
                  <option value="totalScans">Most Scans</option>
                  <option value="name">Name A-Z</option>
                  <option value="status">Status</option>
                </select>
              </label>
              <label class="scan-filter-search">
                <span>Search</span>
                <input id="finishScanSearchInput" type="search" placeholder="Search associate..." autocomplete="off" />
              </label>
            </div>

            <div class="scan-advanced-list-card">
              <div class="scan-advanced-list-head">
                <div>
                  <span>Live Associate List</span>                  
                </div>
                <small id="finishScanListCount">0 visible</small>
              </div>
              <div class="scan-advanced-table-wrap">
                <div class="scan-advanced-table-head">
                  <span>Associate</span><span>Access / Last Scan</span><span>Total</span><span>This Hour</span><span>Current Gap</span><span>Longest</span><span>Avg</span><span>Multi</span><span>Status</span><span>Action</span>
                </div>
                <div id="finishInactivityList" class="scan-advanced-table-body">
                  <article class="finish-inactivity-empty">Loading scan tracker data...</article>
                </div>
              </div>
            </div>

            <section id="finishInactivityProfilePanel" class="scan-profile-panel">
              <div class="associate-profile-empty">Click an associate to open their scan-gap profile.</div>
            </section>
          </section>
        </section>`;

      wirePanel();
    }

    return mount;
  }

  function wirePanel() {
    document.querySelectorAll(".scan-downtime-tab[data-scan-access]").forEach(btn => {
      if (btn.dataset.wired) return;
      btn.dataset.wired = "true";
      btn.addEventListener("click", () => {
        state.activeAccess = btn.dataset.scanAccess || "Mounting";
        state.selectedKey = "";
        renderPanel();
      });
    });

    document.querySelectorAll("[data-scan-status]").forEach(btn => {
      if (btn.dataset.wired) return;
      btn.dataset.wired = "true";
      btn.addEventListener("click", () => {
        state.statusFilter = cleanKey(btn.dataset.scanStatus || "ALL");
        renderPanel();
      });
    });

    const sortSelect = document.getElementById("finishScanSortSelect");
    if (sortSelect && !sortSelect.dataset.wired) {
      sortSelect.dataset.wired = "true";
      sortSelect.value = state.sortBy || "currentGap";
      sortSelect.addEventListener("change", () => {
        state.sortBy = sortSelect.value || "currentGap";
        renderPanel();
      });
    }

    const searchInput = document.getElementById("finishScanSearchInput");
    if (searchInput && !searchInput.dataset.wired) {
      searchInput.dataset.wired = "true";
      searchInput.value = state.search || "";
      searchInput.addEventListener("input", () => {
        state.search = searchInput.value || "";
        renderPanel();
      });
    }

    const runBtn = document.getElementById("finishRunInactivitySnapshotBtn");
    if (runBtn && !runBtn.dataset.scanWired) {
      runBtn.dataset.scanWired = "true";
      runBtn.addEventListener("click", runFinishMorningInactivitySnapshot);
    }

    const refreshBtn = document.getElementById("finishRefreshInactivityLogBtn");
    if (refreshBtn && !refreshBtn.dataset.scanWired) {
      refreshBtn.dataset.scanWired = "true";
      refreshBtn.addEventListener("click", loadFinishMorningInactivityLog);
    }
  }

  function renderPanel() {
    ensurePanel();
    wirePanel();

    document.querySelectorAll(".scan-downtime-tab[data-scan-access]").forEach(btn => {
      btn.classList.toggle("active", cleanAccess(btn.dataset.scanAccess) === cleanAccess(state.activeAccess));
    });
    document.querySelectorAll("[data-scan-status]").forEach(btn => {
      btn.classList.toggle("active", cleanKey(btn.dataset.scanStatus) === cleanKey(state.statusFilter || "ALL"));
    });

    const access = cleanAccess(state.activeAccess || "Mounting");
    const rows = rowsForAccess(access);
    const summary = accessSummary(access);

    setTextSafe("finishInactivityTrackedCount", numberFmt(summary.tracked));
    setTextSafe("finishInactivityWatchCount", numberFmt(summary.watch));
    setTextSafe("finishInactivityInactiveCount", numberFmt(summary.inactive));
    setTextSafe("finishInactivityLastSync", formatTime(state.lastSync));
    setTextSafe("finishScanListCount", `${rows.length} visible`);

    const list = document.getElementById("finishInactivityList");
    if (!list) return;

    if (!rows.length) {
      list.innerHTML = `<article class="finish-inactivity-empty">No ${esc(access)} scan rows match the current filters.</article>`;
      const profile = document.getElementById("finishInactivityProfilePanel");
      if (profile) profile.innerHTML = `<div class="associate-profile-empty">No selected associate profile.</div>`;
      return;
    }

    const selected = getSelectedRow(rows);
    list.innerHTML = rows.map(row => renderRow(row, rowKey(row) === rowKey(selected))).join("");
    list.querySelectorAll("[data-scan-key]").forEach(el => {
      el.addEventListener("click", () => {
        state.selectedKey = el.dataset.scanKey || "";
        renderPanel();
      });
    });

    openProfile(selected);
  }

  function renderRow(row, selected) {
    const name = operatorName(row);
    const access = rowAccess(row);
    const status = normalizeStatus(row.CurrentStatus || row.currentStatus || row.Status || row.status);
    const totalScans = metric(row, "TotalScans", "totalScans");
    const scansThisHour = metric(row, "ScansThisHour", "scansThisHour", "ThisHour", "thisHour");
    const currentGap = metric(row, "CurrentGapMinutes", "currentGapMinutes");
    const longestGap = metric(row, "LongestGapMinutes", "longestGapMinutes");
    const avgGap = metric(row, "AverageGapMinutes", "averageGapMinutes");
    const multiSeconds = metric(row, "MultiScanSeconds", "multiScanSeconds");

    return `
      <button class="scan-operator-row ${statusClass(status)} ${selected ? "selected" : ""}" type="button" data-scan-key="${esc(rowKey(row))}">
        <span class="scan-row-name"><i></i><strong>${esc(name)}</strong></span>
        <span><b>${esc(access)}</b><small>${esc(formatTime(row.LastScanAt || row.lastScanAt))}</small></span>
        <span>${esc(numberFmt(totalScans))}</span>
        <span>${esc(numberFmt(scansThisHour))}</span>
        <span class="scan-gap-value">${esc(formatMinutes(currentGap))}</span>
        <span>${esc(formatMinutes(longestGap))}</span>
        <span>${esc(formatMinutes(avgGap))}</span>
        <span>${esc(numberFmt(multiSeconds))}</span>
        <span><em class="scan-status-pill">${esc(status)}</em></span>
        <span><em class="scan-open-pill">Open</em></span>
      </button>`;
  }

  function statusPercent(currentGap) {
    const n = Math.max(0, Number(currentGap) || 0);
    return Math.min(100, Math.round((n / 20) * 100));
  }

  function gapClass(minutes) {
    const n = Number(minutes) || 0;
    if (n >= 10) return "inactive";
    if (n >= 5) return "watch";
    return "active";
  }

  function openProfile(row) {
    const panel = document.getElementById("finishInactivityProfilePanel");
    if (!panel || !row) return;

    const name = operatorName(row);
    const access = rowAccess(row);
    const status = normalizeStatus(row.CurrentStatus || row.currentStatus || row.Status || row.status);
    const daily = dailyForRow(row);
    const totalScans = metric(row, "TotalScans", "totalScans");
    const scansThisHour = metric(row, "ScansThisHour", "scansThisHour", "ThisHour", "thisHour");
    const currentGap = metric(row, "CurrentGapMinutes", "currentGapMinutes");
    const longestGap = metric(row, "LongestGapMinutes", "longestGapMinutes");
    const avgGap = metric(row, "AverageGapMinutes", "averageGapMinutes");
    const multiSeconds = metric(row, "MultiScanSeconds", "multiScanSeconds");
    const downtime = metric(daily, "TotalDowntimeMinutes", "totalDowntimeMinutes", "InactiveMinutes", "inactiveMinutes");
    const barPct = statusPercent(currentGap);

    panel.innerHTML = `
      <article class="scan-profile-card ${statusClass(status)}">
        <header class="scan-profile-top">
          <div class="scan-profile-id">
            <div class="scan-avatar">${esc(initials(name))}</div>
            <div>
              <span>Associate Profile</span>
              <h3>${esc(name)}</h3>
              <p>Finish Department · ${esc(access)}</p>
            </div>
          </div>
          <em class="scan-status-pill">${esc(status)}</em>
        </header>

        <div class="scan-profile-metrics-strip">
          <article><span>Last Scan</span><strong>${esc(formatTime(row.LastScanAt || row.lastScanAt))}</strong></article>
          <article><span>Total Scans</span><strong>${esc(numberFmt(totalScans))}</strong></article>
          <article><span>This Hour</span><strong>${esc(numberFmt(scansThisHour))}</strong></article>
          <article><span>Current Gap</span><strong>${esc(formatMinutes(currentGap))}</strong></article>
          <article><span>Longest Gap</span><strong>${esc(formatMinutes(longestGap))}</strong></article>
          <article><span>Downtime Today</span><strong>${esc(formatMinutes(downtime))}</strong></article>
          <article><span>Multi-Scan</span><strong>${esc(numberFmt(multiSeconds))}</strong></article>
        </div>

        <div class="scan-gap-meter">
          <div class="scan-gap-meter-labels"><span>0m</span><span>5m Watch</span><span>10m Inactive</span><span>20m+</span></div>
          <div class="scan-gap-meter-track"><i style="width:${barPct}%"></i><b style="left:${Math.min(100, barPct)}%"></b></div>
        </div>

        <nav class="scan-profile-tabs">
          ${["overview", "gaps", "multi", "hourly"].map(tab => `
            <button type="button" class="${state.selectedDetailTab === tab ? "active" : ""}" data-profile-tab="${tab}">${tab === "overview" ? "Overview" : tab === "gaps" ? "Gap History" : tab === "multi" ? "Multi-Scan Events" : "Hourly Timeline"}</button>
          `).join("")}
        </nav>

        <div class="scan-profile-tab-body">
          ${renderDetailTab(row, state.selectedDetailTab || "overview", { avgGap, downtime, currentGap, longestGap, multiSeconds })}
        </div>
      </article>`;

    panel.querySelectorAll("[data-profile-tab]").forEach(btn => {
      btn.addEventListener("click", () => {
        state.selectedDetailTab = btn.dataset.profileTab || "overview";
        openProfile(row);
      });
    });
  }

  function renderDetailTab(row, tab, extras) {
    if (tab === "gaps") return renderGapHistory(row, 12);
    if (tab === "multi") return renderMultiHistory(row, 18);
    if (tab === "hourly") return renderHourlyTimeline(row);
    return renderOverview(row, extras);
  }

  function renderOverview(row, extras) {
    const gaps = operatorGapRows(row);
    const multi = operatorMultiRows(row);
    const recentGaps = gaps.slice(0, 4);
    const recentMulti = multi.slice(0, 3);
    return `
      <div class="scan-overview-grid">
        <section class="scan-mini-panel">
          <div class="scan-mini-head"><span>Gap Snapshot</span><strong>${esc(gaps.length)} 5m+ gaps</strong></div>
          <div class="scan-mini-bars">
            <div><span>Average Gap</span><b>${esc(formatMinutes(extras.avgGap))}</b></div>
            <div><span>Longest Gap</span><b>${esc(formatMinutes(extras.longestGap))}</b></div>
            <div><span>Downtime Today</span><b>${esc(formatMinutes(extras.downtime))}</b></div>
          </div>
        </section>
        <section class="scan-mini-panel">
          <div class="scan-mini-head"><span>Recent Gaps</span><strong>${esc(recentGaps.length)}</strong></div>
          <div class="scan-compact-list">
            ${recentGaps.length ? recentGaps.map(g => renderCompactGapRow(g)).join("") : `<p>No 5m+ gaps yet.</p>`}
          </div>
        </section>
        <section class="scan-mini-panel">
          <div class="scan-mini-head"><span>Multi-Scan</span><strong>${esc(multi.length)} events</strong></div>
          <div class="scan-mini-multi-row">
            ${recentMulti.length ? recentMulti.map(renderSmallMultiCard).join("") : `<p>No same-second events.</p>`}
          </div>
        </section>
      </div>`;
  }

  function renderGapHistory(row, limit = 30) {
    const gaps = operatorGapRows(row).slice(0, limit);
    if (!gaps.length) {
      return `<section class="scan-history-panel"><div class="scan-section-head"><span>Gap History</span><strong>No 5+ minute gaps yet</strong></div><article class="associate-session-empty">This associate has no watch/inactive scan gaps in the current tracker data.</article></section>`;
    }

    return `
      <section class="scan-history-panel">
        <div class="scan-section-head"><span>Gap History</span><strong>${esc(gaps.length)} 5m+ records</strong></div>
        <div class="scan-gap-table">
          <div class="scan-gap-table-head"><span>Start</span><span>End</span><span>Duration</span><span>Status</span><span>Movement</span></div>
          ${gaps.map(g => renderGapTableRow(g)).join("")}
        </div>
      </section>`;
  }

  function renderGapTableRow(g) {
    const status = normalizeStatus(g.Status || g.status);
    const start = g.GapStartAt || g.gapStartAt || g.PreviousScanTime || g.previousScanTime;
    const end = g.GapEndAt || g.gapEndAt || g.NextScanTime || g.nextScanTime || "";
    const minutes = metric(g, "GapMinutes", "gapMinutes");
    return `
      <article class="scan-gap-line ${statusClass(status)}">
        <span>${esc(formatTime(start))}</span>
        <span>${esc(end ? formatTime(end) : "Open")}</span>
        <span>${esc(formatMinutes(minutes))}</span>
        <span><em class="scan-status-pill">${esc(status)}</em></span>
        <span>${esc(formatTime(g.PreviousScanTime || g.previousScanTime))} → ${esc(g.NextScanTime || g.nextScanTime ? formatTime(g.NextScanTime || g.nextScanTime) : "Open")}</span>
      </article>`;
  }

  function renderCompactGapRow(g) {
    const status = normalizeStatus(g.Status || g.status);
    const start = g.GapStartAt || g.gapStartAt || g.PreviousScanTime || g.previousScanTime;
    const end = g.GapEndAt || g.gapEndAt || g.NextScanTime || g.nextScanTime || "";
    const minutes = metric(g, "GapMinutes", "gapMinutes");
    return `<div class="scan-compact-row ${statusClass(status)}"><span>${esc(formatTime(start))} → ${esc(end ? formatTime(end) : "Open")}</span><b>${esc(formatMinutes(minutes))}</b><em>${esc(status)}</em></div>`;
  }

  function renderMultiHistory(row, limit = 30) {
    const events = operatorMultiRows(row).slice(0, limit);
    if (!events.length) {
      return `<section class="scan-history-panel"><div class="scan-section-head"><span>Multi-Scan Events</span><strong>0 same-second events</strong></div><article class="associate-session-empty">No same-second multi-scan events for this associate.</article></section>`;
    }

    return `
      <section class="scan-history-panel">
        <div class="scan-section-head"><span>Multi-Scan Events</span><strong>${esc(events.length)} same-second events</strong></div>
        <div class="scan-multi-grid advanced">
          ${events.map(renderSmallMultiCard).join("")}
        </div>
      </section>`;
  }

  function renderSmallMultiCard(ev) {
    const count = ev.ScanCount || ev.scanCount || 0;
    return `
      <article class="scan-small-multi-card">
        <span>${esc(formatTime(ev.ScanTime || ev.scanTime))}</span>
        <strong>${esc(count)} jobs</strong>
        <small>${esc(cleanAccess(ev.AccessPoint || ev.accessPoint))}</small>
      </article>`;
  }

  function getScanTimelineShiftHours(row, rawRows = []) {
    const dateSource =
      row?.LastScanAt ||
      row?.lastScanAt ||
      rawRows.find(g => g.GapStartAt || g.PreviousScanTime || g.NextScanTime)?.GapStartAt ||
      rawRows.find(g => g.GapStartAt || g.PreviousScanTime || g.NextScanTime)?.PreviousScanTime ||
      rawRows.find(g => g.GapStartAt || g.PreviousScanTime || g.NextScanTime)?.NextScanTime ||
      new Date();

    const date = toDate(dateSource) || new Date();
    const day = date.getDay();
    const isWeekend = day === 0 || day === 6;

    // Weekday Finish = 7 AM start.
    // Weekend Finish = 6 AM start.
    return isWeekend
      ? [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
      : [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];
  }

  function getOperatorRawGapRows(row) {
    const keys = new Set(rowSourceKeys(row));

    // Use raw gapLog here, NOT operatorGapRows().
    // operatorGapRows() filters out normal ACTIVE gaps under 5 minutes.
    // The hourly timeline must still show those hours, especially 7 AM.
    return (state.gapLog || [])
      .filter(g => keys.has(baseRowKey(g)))
      .sort((a, b) => dateMs(a.GapStartAt || a.PreviousScanTime) - dateMs(b.GapStartAt || b.PreviousScanTime));
  }

  function renderHourlyTimeline(row) {
    const rawGaps = getOperatorRawGapRows(row);
    const multi = operatorMultiRows(row);
    const shiftHours = getScanTimelineShiftHours(row, rawGaps);
    const buckets = new Map();

    function bucketForHour(hour) {
      if (!buckets.has(hour)) {
        buckets.set(hour, {
          hour,
          gaps: [],
          watchGaps: [],
          multi: 0,
          scansEstimate: 0
        });
      }
      return buckets.get(hour);
    }

    // Build every shift hour first so 7 AM does not disappear when it only has normal gaps.
    shiftHours.forEach(hour => bucketForHour(hour));

    rawGaps.forEach(g => {
      const d = toDate(g.GapStartAt || g.PreviousScanTime || g.NextScanTime);
      if (!d) return;

      const hour = d.getHours();
      if (!shiftHours.includes(hour)) return;

      const gapMinutes = metric(g, "GapMinutes", "gapMinutes");
      const b = bucketForHour(hour);

      b.gaps.push(gapMinutes);
      b.scansEstimate += 1;

      if (isRealDisplayGap(g)) {
        b.watchGaps.push(gapMinutes);
      }
    });

    multi.forEach(m => {
      const d = toDate(m.ScanTime || m.scanTime);
      if (!d) return;

      const hour = d.getHours();
      if (!shiftHours.includes(hour)) return;

      const b = bucketForHour(hour);
      b.multi += 1;
    });

    const rows = shiftHours.map(hour => buckets.get(hour));

    return `
      <section class="scan-history-panel">
        <div class="scan-section-head">
          <span>Hourly Timeline</span>
          <strong>Full shift · gap / multi-scan pattern</strong>
        </div>
        <div class="scan-hourly-graph">
          ${rows.map(b => {
            const largest = b.gaps.length ? Math.max(...b.gaps) : 0;
            const avg = b.gaps.length ? Math.round((b.gaps.reduce((a, c) => a + c, 0) / b.gaps.length) * 10) / 10 : 0;
            const watchCount = b.watchGaps.length;
            const width = b.gaps.length
              ? Math.min(100, Math.max(8, (largest / 20) * 100))
              : 4;

            return `
              <div class="scan-hour-row ${gapClass(largest)}">
                <span>${esc(formatHourLabel(b.hour))}</span>
                <div class="scan-hour-bar"><i style="width:${width}%"></i></div>
                <b>${esc(formatMinutes(largest))}</b>
                <em>${esc(formatMinutes(avg))} avg · ${esc(numberFmt(b.scansEstimate))} scans</em>
                <strong>${esc(b.multi)} multi${watchCount ? ` · ${esc(watchCount)} watch` : ""}</strong>
              </div>`;
          }).join("")}
        </div>
      </section>`;
  }

  function formatHourLabel(hour24) {
    const suffix = hour24 >= 12 ? "PM" : "AM";
    let hour = hour24 % 12;
    if (!hour) hour = 12;
    return `${hour} ${suffix}`;
  }

  async function loadFinishMorningInactivityLog() {
    ensurePanel();
    const btn = document.getElementById("finishRefreshInactivityLogBtn");
    if (btn) { btn.disabled = true; btn.textContent = "Refreshing..."; }
    try {
      const payload = await fetchFinishScanActivityApi("getfinishscanactivity", { debug: true, cacheBust: Date.now() });
      applyPayload(payload);
    } catch (error) {
      console.error("Finish Scan Activity Tracker load failed:", error);
      const list = document.getElementById("finishInactivityList");
      if (list) list.innerHTML = `<article class="finish-inactivity-empty">Scan Tracker API failed: ${esc(error.message || String(error))}</article>`;
      if (typeof showFinishAssignmentToast === "function") showFinishAssignmentToast(error.message || String(error));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "Refresh Log"; }
    }
  }

  async function runFinishMorningInactivitySnapshot() {
    ensurePanel();
    if (state.loading) return;
    const btn = document.getElementById("finishRunInactivitySnapshotBtn");
    state.loading = true;
    if (btn) { btn.disabled = true; btn.textContent = "Running..."; }
    try {
      await fetchFinishScanActivityApi("runfinishscanactivityautorefresh", { cacheBust: Date.now() });
      const payload = await fetchFinishScanActivityApi("getfinishscanactivity", { debug: true, cacheBust: Date.now() });
      applyPayload(payload);
      if (typeof showFinishAssignmentToast === "function") showFinishAssignmentToast("Scan tracker refreshed.");
    } catch (error) {
      console.error("Finish Scan Activity refresh failed:", error);
      if (typeof showFinishAssignmentToast === "function") showFinishAssignmentToast(error.message || String(error));
      await loadFinishMorningInactivityLog();
    } finally {
      state.loading = false;
      if (btn) { btn.disabled = false; btn.textContent = "Run Scan Refresh"; }
    }
  }

  function refreshFinishInactivityAutomationStatus() {
    ensurePanel();
    return Promise.resolve({ success: true, enabled: true, source: "scan-tracker" });
  }

  function toggleFinishInactivityAutomationFromPage() {
    if (typeof showFinishAssignmentToast === "function") {
      showFinishAssignmentToast("Scan tracker uses its own backend refresh. Use Run Scan Refresh here.");
    }
    return Promise.resolve({ success: true });
  }

  function startFinishInactivityFrontendAutoRefresh() {
    if (window.finishInactivityAutoRefreshTimerFinal) clearInterval(window.finishInactivityAutoRefreshTimerFinal);
    if (window.finishInactivityAutoRefreshTimer) clearInterval(window.finishInactivityAutoRefreshTimer);
    loadFinishMorningInactivityLog();
    window.finishInactivityAutoRefreshTimerFinal = setInterval(loadFinishMorningInactivityLog, 120000);
  }

  function resetFinishInactivityButtons() {
    const runBtn = document.getElementById("finishRunInactivitySnapshotBtn");
    const refreshBtn = document.getElementById("finishRefreshInactivityLogBtn");
    if (runBtn) { runBtn.disabled = false; runBtn.textContent = "Run Scan Refresh"; }
    if (refreshBtn) { refreshBtn.disabled = false; refreshBtn.textContent = "Refresh Log"; }
  }

  window.fetchFinishScanActivityApi = fetchFinishScanActivityApi;
  window.finishScanEscapeHtml = esc;
  window.finishScanApplyPayload = applyPayload;
  window.loadFinishMorningInactivityLog = loadFinishMorningInactivityLog;
  window.runFinishMorningInactivitySnapshot = runFinishMorningInactivitySnapshot;
  window.refreshFinishInactivityAutomationStatus = refreshFinishInactivityAutomationStatus;
  window.toggleFinishInactivityAutomationFromPage = toggleFinishInactivityAutomationFromPage;
  window.startFinishInactivityFrontendAutoRefresh = startFinishInactivityFrontendAutoRefresh;
  window.resetFinishInactivityButtons = resetFinishInactivityButtons;
  window.cleanupFinishDowntimeSnapshotRuleCards = cleanupFinishDowntimeSnapshotRuleCards;

  const finishDowntimeCleanupObserver = new MutationObserver(() => cleanupFinishDowntimeSnapshotRuleCards());
  setTimeout(() => {
    const mount = document.getElementById("finishInactivityTabMount");
    if (mount) {
      finishDowntimeCleanupObserver.observe(mount, { childList: true, subtree: true });
      cleanupFinishDowntimeSnapshotRuleCards();
    }
  }, 500);

  try { eval("loadFinishMorningInactivityLog = window.loadFinishMorningInactivityLog"); } catch (e) {}
  try { eval("runFinishMorningInactivitySnapshot = window.runFinishMorningInactivitySnapshot"); } catch (e) {}
  try { eval("refreshFinishInactivityAutomationStatus = window.refreshFinishInactivityAutomationStatus"); } catch (e) {}
  try { eval("toggleFinishInactivityAutomationFromPage = window.toggleFinishInactivityAutomationFromPage"); } catch (e) {}
  try { eval("startFinishInactivityFrontendAutoRefresh = window.startFinishInactivityFrontendAutoRefresh"); } catch (e) {}

  window.addEventListener("error", resetFinishInactivityButtons);
  window.addEventListener("unhandledrejection", resetFinishInactivityButtons);

  setTimeout(() => {
    ensurePanel();
    wirePanel();
    cleanupFinishDowntimeSnapshotRuleCards();
    startFinishInactivityFrontendAutoRefresh();
  }, 250);
})();


/* =========================================================
   HOTFIX — HIDE ME SNAPSHOT RULE FROM FINISH DOWNTIME VIEW
========================================================= */
(function installFinishDowntimeSnapshotRuleHideStyle() {
  if (document.getElementById("finishDowntimeSnapshotRuleHideStyle")) return;
  const style = document.createElement("style");
  style.id = "finishDowntimeSnapshotRuleHideStyle";
  style.textContent = `
    #finishInactivityTabMount .scan-advanced-rule-card { display: none !important; }
    #finishInactivityTabMount #finishMorningInactivityPanel { display: none !important; }
  `;
  document.head.appendChild(style);
})();


/* =========================================================
   HOTFIX — REMOVE TOP ASSOCIATE DOWNTIME SNAPSHOT BANNER
========================================================= */
(function installFinishDowntimeTopBannerRemoveHotfix() {
  function removeTopBanner() {
    const mount = document.getElementById("finishInactivityTabMount");
    if (!mount) return;

    mount.querySelectorAll(".scan-advanced-hero").forEach(el => el.remove());

    mount.querySelectorAll("header, section, article, div").forEach(el => {
      const text = String(el.textContent || "").replace(/\s+/g, " ").trim().toUpperCase();
      if (!text) return;
      const hasOldSnapshot = text.includes("SNAPSHOT RULE") || text.includes("3 CHECKS");
      const insideMainLog = !!el.closest("#finishScanDowntimePanel .scan-advanced-panel");
      if (hasOldSnapshot && !insideMainLog) el.remove();
    });
  }

  if (!document.getElementById("finishDowntimeTopBannerRemoveStyle")) {
    const style = document.createElement("style");
    style.id = "finishDowntimeTopBannerRemoveStyle";
    style.textContent = `
      #finishInactivityTabMount .scan-advanced-hero { display: none !important; }
      #finishInactivityTabMount .scan-advanced-rule-card { display: none !important; }
    `;
    document.head.appendChild(style);
  }

  document.addEventListener("DOMContentLoaded", removeTopBanner);
  window.addEventListener("load", removeTopBanner);
  setTimeout(removeTopBanner, 100);
  setTimeout(removeTopBanner, 500);
  setTimeout(removeTopBanner, 1500);

  const observer = new MutationObserver(removeTopBanner);
  document.addEventListener("DOMContentLoaded", () => {
    const mount = document.getElementById("finishInactivityTabMount");
    if (mount) observer.observe(mount, { childList: true, subtree: true });
  });

  window.removeFinishDowntimeTopBanner = removeTopBanner;
})();


/*******************************************************
 * FINISH DOWNTIME STATUS EVENTS — STOP COUNTING / NOTES
 *
 * Adds supervisor/LMS note events to Associate Downtime Log:
 * - Moved
 * - Left for day
 * - No WIP
 * - Machine issue
 * - Training / coaching
 * - Break / lunch
 * - Support another area
 * - Other
 *
 * This does NOT edit raw scan data.
 * It adds an exception/status window and shows:
 * - Raw downtime
 * - Excluded time
 * - Adjusted downtime
 *******************************************************/
(function installFinishDowntimeStatusEvents() {
  const EVENT_ACTION_GET = "getfinishassociatestatusevents";
  const EVENT_ACTION_SAVE = "savefinishassociatestatusevent";
  const EVENT_ACTION_CLOSE = "closefinishassociatestatusevent";

  const state = window.finishScanDowntimeState = window.finishScanDowntimeState || {};
  state.statusEvents = Array.isArray(state.statusEvents) ? state.statusEvents : [];
  state.statusEventsLoadedAt = state.statusEventsLoadedAt || "";

  function esc(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function clean(value) {
    return String(value || "").trim();
  }

  function cleanKey(value) {
    return clean(value).toUpperCase().replace(/\s+/g, " ");
  }

  function cleanAccess(value) {
    const text = clean(value).toLowerCase();
    if (text.includes("mount") && text.includes("drill")) return "Mounting / Drill";
    if (text.includes("final")) return "Final Inspection";
    if (text.includes("drill")) return "Drill";
    if (text.includes("mount")) return "Mounting";
    return clean(value) || "Unknown";
  }

  function operatorName(row) {
    return clean(row?.OperatorName || row?.operatorName || row?.Operator || row?.operator || "Unknown Operator");
  }

  function rowAccess(row) {
    return cleanAccess(row?.AccessPoint || row?.accessPoint || row?.ScanStage || row?.scanStage || "");
  }

  function profileFromDom() {
    const panel = document.getElementById("finishInactivityProfilePanel");
    if (!panel) return null;

    const name = clean(panel.querySelector(".scan-profile-id h3")?.textContent || "");
    const profileText = clean(panel.querySelector(".scan-profile-id p")?.textContent || "");
    let access = profileText.split("·").pop() || "";
    access = cleanAccess(access);

    if (!name) return null;
    return { operatorName: name, accessPoint: access };
  }

  function todayIso() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  function timeValueNow() {
    const d = new Date();
    return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  }

  function toMinutes(timeText) {
    const text = clean(timeText);
    if (!text || text.toUpperCase() === "EOS") return null;

    // Accept HH:MM, H:MM AM, or Google date strings.
    const dateTry = new Date(text);
    if (!Number.isNaN(dateTry.getTime()) && /AM|PM|\/|-/.test(text)) {
      return dateTry.getHours() * 60 + dateTry.getMinutes();
    }

    const match = text.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/i);
    if (!match) return null;

    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const ampm = String(match[3] || "").toUpperCase();

    if (ampm === "PM" && hour < 12) hour += 12;
    if (ampm === "AM" && hour === 12) hour = 0;

    return hour * 60 + minute;
  }

  function minutesToLabel(total) {
    if (total == null || !Number.isFinite(total)) return "";
    const hour24 = Math.floor(total / 60);
    const minute = total % 60;
    const suffix = hour24 >= 12 ? "PM" : "AM";
    let h = hour24 % 12;
    if (!h) h = 12;
    return `${h}:${String(minute).padStart(2, "0")} ${suffix}`;
  }

  function formatMinutes(value) {
    const n = Math.max(0, Math.round(Number(value) || 0));
    if (n >= 60) {
      const h = Math.floor(n / 60);
      const m = n % 60;
      return `${h}h ${m}m`;
    }
    return `${n}m`;
  }

  function eventMatchesProfile(event, profile) {
    if (!event || !profile) return false;
    const evName = cleanKey(event.OperatorName || event.operatorName);
    const evAccess = cleanAccess(event.AccessPoint || event.accessPoint || event.Area || event.area);
    const profileName = cleanKey(profile.operatorName);
    const profileAccess = cleanAccess(profile.accessPoint);

    if (evName !== profileName) return false;

    // Mounting view intentionally includes Drill/Mounting.
    if (profileAccess === "Mounting" || profileAccess === "Mounting / Drill") {
      return evAccess === "Mounting" || evAccess === "Drill" || evAccess === "Mounting / Drill";
    }

    return evAccess === profileAccess;
  }

  function getProfileEvents(profile) {
    return (state.statusEvents || [])
      .filter(event => eventMatchesProfile(event, profile))
      .sort((a, b) => String(b.EnteredAt || b.enteredAt || "").localeCompare(String(a.EnteredAt || a.enteredAt || "")));
  }

  function isOpenEvent(event) {
    const status = clean(event.Status || event.status || "Open").toUpperCase();
    const end = clean(event.EndTime || event.endTime);
    return status !== "CLOSED" && (!end || end.toUpperCase() === "EOS");
  }

  function calculateExcludedMinutes(events) {
    const now = new Date();
    const nowMins = now.getHours() * 60 + now.getMinutes();

    return (events || []).reduce((sum, event) => {
      const start = toMinutes(event.StartTime || event.startTime);
      let end = toMinutes(event.EndTime || event.endTime);

      if (start == null) return sum;
      if (end == null) end = nowMins;
      if (end < start) return sum;

      return sum + Math.max(0, end - start);
    }, 0);
  }

  function readRawDowntimeFromProfile() {
    const cards = Array.from(document.querySelectorAll("#finishInactivityProfilePanel .scan-profile-metrics-strip article"));
    for (const card of cards) {
      const label = clean(card.querySelector("span")?.textContent || "").toLowerCase();
      if (!label.includes("downtime")) continue;

      const text = clean(card.querySelector("strong")?.textContent || "");
      let total = 0;

      const h = text.match(/(\d+)\s*h/i);
      const m = text.match(/(\d+)\s*m/i);

      if (h) total += Number(h[1]) * 60;
      if (m) total += Number(m[1]);
      if (!h && !m && /^\d+/.test(text)) total = Number(text.match(/^\d+/)[0]);

      return total;
    }

    return 0;
  }

  async function fetchStatusEvents() {
    if (typeof window.fetchFinishScanActivityApi !== "function") return [];

    try {
      const payload = await window.fetchFinishScanActivityApi(EVENT_ACTION_GET, {
        date: todayIso(),
        cacheBust: Date.now()
      });

      const rows =
        payload.statusEvents ||
        payload.events ||
        payload.data ||
        payload.payload?.statusEvents ||
        [];

      state.statusEvents = Array.isArray(rows) ? rows : [];
      state.statusEventsLoadedAt = new Date().toISOString();

      renderStatusEventEnhancement();
      return state.statusEvents;
    } catch (error) {
      console.warn("Could not load Finish status events:", error);
      return state.statusEvents || [];
    }
  }

  async function saveStatusEvent(profile, data) {
    if (typeof window.fetchFinishScanActivityApi !== "function") {
      throw new Error("Finish Scan Activity API helper is not available.");
    }

    const enteredBy = typeof getCurrentUsername === "function" ? getCurrentUsername() : "SYSTEM";

    await window.fetchFinishScanActivityApi(EVENT_ACTION_SAVE, {
      date: todayIso(),
      operatorName: profile.operatorName,
      accessPoint: profile.accessPoint,
      eventType: data.eventType,
      startTime: data.startTime,
      endTime: data.endTime,
      reason: data.reason || data.eventType,
      notes: data.notes,
      enteredBy,
      cacheBust: Date.now()
    });

    await fetchStatusEvents();
  }

  async function closeStatusEvent(eventId) {
    if (typeof window.fetchFinishScanActivityApi !== "function") {
      throw new Error("Finish Scan Activity API helper is not available.");
    }

    const enteredBy = typeof getCurrentUsername === "function" ? getCurrentUsername() : "SYSTEM";

    await window.fetchFinishScanActivityApi(EVENT_ACTION_CLOSE, {
      eventId,
      endTime: timeValueNow(),
      updatedBy: enteredBy,
      cacheBust: Date.now()
    });

    await fetchStatusEvents();
  }

  function ensureStatusModal() {
    let modal = document.getElementById("finishStatusEventModal");
    if (modal) return modal;

    modal = document.createElement("div");
    modal.id = "finishStatusEventModal";
    modal.className = "finish-status-modal";
    modal.innerHTML = `
      <div class="finish-status-modal-card">
        <header>
          <div>
            <span>Downtime Exception</span>
            <h3 id="finishStatusModalTitle">Stop Counting</h3>
            <p id="finishStatusModalSub">Add a status event. Raw scan data will not be changed.</p>
          </div>
          <button type="button" data-status-modal-close>×</button>
        </header>

        <div class="finish-status-grid">
          <label>
            Event Type
            <select id="finishStatusEventType">
              <option value="Moved">Moved</option>
              <option value="Left for day">Left for day</option>
              <option value="No WIP">No WIP</option>
              <option value="Machine issue">Machine issue</option>
              <option value="Training / coaching">Training / coaching</option>
              <option value="Break / lunch">Break / lunch</option>
              <option value="Support another area">Support another area</option>
              <option value="Other">Other</option>
            </select>
          </label>

          <label>
            Start Time
            <input id="finishStatusStartTime" type="time" />
          </label>

          <label>
            End Time
            <input id="finishStatusEndTime" type="time" />
          </label>

          <label class="wide">
            Reason / Notes
            <textarea id="finishStatusNotes" rows="4" placeholder="Example: moved to Final Inspection, left early, no WIP, machine issue..."></textarea>
          </label>
        </div>

        <footer>
          <button type="button" class="ghost" data-status-modal-close>Cancel</button>
          <button type="button" id="finishSaveStatusEventBtn">Save Event</button>
        </footer>
      </div>
    `;

    document.body.appendChild(modal);

    modal.querySelectorAll("[data-status-modal-close]").forEach(btn => {
      btn.addEventListener("click", () => modal.classList.remove("open"));
    });

    modal.addEventListener("click", event => {
      if (event.target === modal) modal.classList.remove("open");
    });

    return modal;
  }

  function openStatusModal(profile) {
    const modal = ensureStatusModal();

    modal.dataset.operatorName = profile.operatorName;
    modal.dataset.accessPoint = profile.accessPoint;

    const title = modal.querySelector("#finishStatusModalTitle");
    const sub = modal.querySelector("#finishStatusModalSub");
    const start = modal.querySelector("#finishStatusStartTime");
    const end = modal.querySelector("#finishStatusEndTime");
    const notes = modal.querySelector("#finishStatusNotes");
    const eventType = modal.querySelector("#finishStatusEventType");
    const saveBtn = modal.querySelector("#finishSaveStatusEventBtn");

    if (title) title.textContent = `Stop Counting · ${profile.operatorName}`;
    if (sub) sub.textContent = `${profile.accessPoint} · Add reason and time window.`;
    if (start) start.value = timeValueNow();
    if (end) end.value = "";
    if (notes) notes.value = "";
    if (eventType) eventType.value = "Moved";

    if (saveBtn) {
      saveBtn.onclick = async () => {
        saveBtn.disabled = true;
        saveBtn.textContent = "Saving...";

        try {
          await saveStatusEvent(profile, {
            eventType: eventType?.value || "Moved",
            startTime: start?.value || timeValueNow(),
            endTime: end?.value || "",
            reason: eventType?.value || "Moved",
            notes: notes?.value || ""
          });

          modal.classList.remove("open");
          if (typeof showFinishAssignmentToast === "function") {
            showFinishAssignmentToast("Status event saved. Downtime adjusted.");
          }
        } catch (error) {
          console.error(error);
          if (typeof showFinishAssignmentToast === "function") {
            showFinishAssignmentToast(error.message || String(error));
          }
        } finally {
          saveBtn.disabled = false;
          saveBtn.textContent = "Save Event";
        }
      };
    }

    modal.classList.add("open");
  }

  function renderStatusEventEnhancement() {
    const panel = document.getElementById("finishInactivityProfilePanel");
    const card = panel?.querySelector(".scan-profile-card");
    if (!panel || !card) return;

    const profile = profileFromDom();
    if (!profile) return;

    card.querySelectorAll(".finish-status-event-enhancement, .finish-status-event-actions, .finish-status-event-list").forEach(el => el.remove());

    const events = getProfileEvents(profile);
    const openEvents = events.filter(isOpenEvent);
    const excluded = calculateExcludedMinutes(events);
    const rawDowntime = readRawDowntimeFromProfile();
    const adjusted = Math.max(0, rawDowntime - excluded);

    const metrics = card.querySelector(".scan-profile-metrics-strip");
    if (metrics) {
      metrics.insertAdjacentHTML("beforeend", `
        <article class="finish-status-event-enhancement"><span>Excluded Time</span><strong>${esc(formatMinutes(excluded))}</strong></article>
        <article class="finish-status-event-enhancement"><span>Adjusted Downtime</span><strong>${esc(formatMinutes(adjusted))}</strong></article>
      `);
    }

    const top = card.querySelector(".scan-profile-top");
    if (top) {
      const activeBadge = openEvents.length
        ? `<em class="finish-status-active-badge">Excluded · ${esc(openEvents[0].EventType || openEvents[0].eventType || "Open")}</em>`
        : "";

      top.insertAdjacentHTML("beforeend", `
        <div class="finish-status-event-actions">
          ${activeBadge}
          <button type="button" id="finishAddStatusEventBtn">Stop Counting / Add Note</button>
          ${openEvents.length ? `<button type="button" id="finishCloseStatusEventBtn" data-event-id="${esc(openEvents[0].EventId || openEvents[0].eventId || "")}">Resume Counting</button>` : ""}
        </div>
      `);

      card.querySelector("#finishAddStatusEventBtn")?.addEventListener("click", () => openStatusModal(profile));
      card.querySelector("#finishCloseStatusEventBtn")?.addEventListener("click", async event => {
        const eventId = event.currentTarget.dataset.eventId || "";
        if (!eventId) return openStatusModal(profile);

        event.currentTarget.disabled = true;
        event.currentTarget.textContent = "Closing...";

        try {
          await closeStatusEvent(eventId);
          if (typeof showFinishAssignmentToast === "function") {
            showFinishAssignmentToast("Counting resumed.");
          }
        } catch (error) {
          console.error(error);
          if (typeof showFinishAssignmentToast === "function") {
            showFinishAssignmentToast(error.message || String(error));
          }
        }
      });
    }

    if (events.length) {
      const body = card.querySelector(".scan-profile-tab-body") || card;
      body.insertAdjacentHTML("beforebegin", `
        <section class="finish-status-event-list">
          <div class="scan-section-head">
            <span>Status Events</span>
            <strong>${esc(events.length)} note${events.length === 1 ? "" : "s"}</strong>
          </div>
          ${events.slice(0, 4).map(event => {
            const start = event.StartTime || event.startTime || "";
            const end = event.EndTime || event.endTime || (isOpenEvent(event) ? "Open" : "");
            const type = event.EventType || event.eventType || "Event";
            const notes = event.Notes || event.notes || "";
            const enteredBy = event.EnteredBy || event.enteredBy || "";
            return `
              <article class="finish-status-event-row ${isOpenEvent(event) ? "open" : "closed"}">
                <b>${esc(type)}</b>
                <span>${esc(start)}${end ? ` → ${esc(end)}` : ""}</span>
                <em>${esc(notes || "No notes")}</em>
                <small>${esc(enteredBy)}</small>
              </article>`;
          }).join("")}
        </section>
      `);
    }
  }

  function installObserver() {
    const target = document.getElementById("finishInactivityProfilePanel");
    if (!target || target.dataset.statusObserverReady) return;

    target.dataset.statusObserverReady = "true";
    const observer = new MutationObserver(() => {
      window.clearTimeout(installObserver._timer);
      installObserver._timer = window.setTimeout(renderStatusEventEnhancement, 80);
    });

    observer.observe(target, { childList: true, subtree: true });
  }

  function bootStatusEvents() {
    installObserver();
    fetchStatusEvents();
    renderStatusEventEnhancement();
  }

  document.addEventListener("DOMContentLoaded", () => {
    window.setTimeout(bootStatusEvents, 1200);
    window.setInterval(fetchStatusEvents, 120000);
  });

  document.addEventListener("click", event => {
    const tab = event.target.closest?.('.tab-btn[data-tab="inactivity"]');
    if (!tab) return;
    window.setTimeout(bootStatusEvents, 500);
  });

  window.refreshFinishDowntimeStatusEvents = fetchStatusEvents;
})();




function cleanFinishTqRoleOptionsOnce_() {
  const root = document.querySelector("[data-content='lms-control']") || document.body;

  root.querySelectorAll("select").forEach(select => {
    Array.from(select.options || []).forEach(option => {
      const text = String(option.textContent || "").trim().toUpperCase();
      const value = String(option.value || "").trim().toUpperCase();

      if (text === "TQ" || value === "TQ") {
        option.remove();
      }
    });
  });
}


document.addEventListener("DOMContentLoaded", () => {
  if (typeof forceFinishSetupTabVisibilityForAllowedRoles_ === "function") {
    forceFinishSetupTabVisibilityForAllowedRoles_();
  }
});
