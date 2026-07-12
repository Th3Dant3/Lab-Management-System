const AUTH_API =
  "https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/*
 * This value must exactly match AUTH_VERSION in the root authGuard.js.
 * Changing both values invalidates all existing sessions.
 */
const AUTH_VERSION = "2026-07-12-main-rebuild";

/**************************************************
 * LOADER MAP
 * login.html is inside /auth
 * loaders are inside /auth/loaders
 **************************************************/
const LOADER_MAP = {
  BLOPEZ: "loaders/loader_BLOPEZ.html",
  MLITTLE: "loaders/loader_MLITTLE.html",
  JBOOMERSHINE: "loaders/loader_JBOOMERSHINE.html",
  BKARR: "loaders/loader_BKARR.html",
  RTATE: "loaders/loader_RTATE.html",
  AIVANOVSKI: "loaders/loader_AIVANOVSKI.html",
  SANDERSON: "loaders/loader_SANDERSON.html",
  BHONICKER: "loaders/loader_BHONICKER.html",
  BDADE: "loaders/loader_BDADE.html",
  BBLAKE: "loaders/loader_BBLAKE.html",
  NPOSTON: "loaders/loader_NPOSTON.html",
  CSEARFOSS: "loaders/loader_CSEARFOSS.html",
  DWATTERS: "loaders/loader_DWATTERS.html",

  CDAY: "loaders/loader_TEAMLEAD.html",
  JJOHNSON: "loaders/loader_TEAMLEAD.html",
  CPATRICK: "loaders/loader_TEAMLEAD.html",
  CWOOD: "loaders/loader_TEAMLEAD.html",
  CFORBES: "loaders/loader_TEAMLEAD.html",
  JADAIR: "loaders/loader_TEAMLEAD.html",
  JMOLING: "loaders/loader_TEAMLEAD.html",
  KSMITH: "loaders/loader_TEAMLEAD.html"
};

document.addEventListener("DOMContentLoaded", () => {
  const userEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");
  const toggleBtn = document.getElementById("togglePassword");

  if (userEl) {
    userEl.addEventListener("input", () => {
      userEl.value = userEl.value.toUpperCase();
    });
  }

  if (toggleBtn && passwordEl) {
    toggleBtn.addEventListener("click", () => {
      const isPassword = passwordEl.type === "password";
      passwordEl.type = isPassword ? "text" : "password";
      toggleBtn.textContent = isPassword ? "Hide" : "Show";
    });
  }

  document.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      login();
    }
  });
});

/**************************************************
 * UI HELPERS
 **************************************************/
function setMessage(text, type = "warning") {
  const wrap = document.getElementById("messageWrap");
  const el = document.getElementById("message");

  if (!wrap || !el) return;

  wrap.classList.remove("warning", "success", "error");
  wrap.classList.add(type);
  el.textContent = text;
}

function setLoading(isLoading) {
  const btn = document.getElementById("loginBtn");
  const btnText = document.getElementById("loginBtnText");

  if (!btn || !btnText) return;

  btn.disabled = isLoading;
  btnText.textContent = isLoading ? "Signing In..." : "Sign In";
}

function setProgress(percent) {
  const progress = document.getElementById("progressBar");

  if (progress) {
    progress.style.width = `${percent}%`;
  }
}

/**************************************************
 * LOGIN
 **************************************************/
async function login() {
  const usernameEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");

  const username = usernameEl?.value.trim().toUpperCase();
  const password = passwordEl?.value.trim();

  if (!username || !password) {
    setMessage("Enter username and password.", "error");
    setProgress(0);
    return;
  }

  setLoading(true);
  setMessage("Authenticating secure session...", "warning");
  setProgress(25);

  const url =
    `${AUTH_API}?action=loginFull` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  try {
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();

    setProgress(70);
    await handleLoginResponse(data, password);
  } catch (error) {
    console.error("Login error:", error);
    setMessage("Connection error. Please try again.", "error");
    setProgress(0);
  } finally {
    setLoading(false);
  }
}

/**************************************************
 * HANDLE RESPONSE
 **************************************************/
async function handleLoginResponse(data, originalPassword) {
  if (!data || !data.status) {
    setMessage("Invalid server response.", "error");
    setProgress(0);
    return;
  }

  if (data.status === "ERROR") {
    setMessage(data.message || "Login failed.", "error");
    setProgress(0);
    return;
  }

  if (data.status === "SET_PASSWORD_REQUIRED") {
    setMessage("First login detected. Setting password...", "warning");
    setProgress(55);
    await setPassword(data.username, originalPassword);
    return;
  }

  if (data.status !== "SUCCESS") {
    setMessage(data.message || "Login failed.", "error");
    setProgress(0);
    return;
  }

  const username = String(data.username || "").trim().toUpperCase();
  const fullName =
    `${data.firstName || ""} ${data.lastName || ""}`.trim() ||
    username;

  const visibility = data.visibility || {};
  const departments = visibility.departments || {};
  const features = visibility.features || {};

  sessionStorage.setItem("lms_logged_in", "true");
  sessionStorage.setItem("lms_user", username);
  sessionStorage.setItem("lms_username", username);
  sessionStorage.setItem("lms_role", data.role || "");
  sessionStorage.setItem("lms_subrole", data.subRole || "");
  sessionStorage.setItem("lms_fullname", fullName);
  sessionStorage.setItem("lms_firstname", data.firstName || "");
  sessionStorage.setItem("lms_lastname", data.lastName || "");
  sessionStorage.setItem("lms_auth_version", AUTH_VERSION);

  sessionStorage.setItem(
    "lms_visibility",
    JSON.stringify(visibility)
  );
  sessionStorage.setItem(
    "lms_departments",
    JSON.stringify(departments)
  );
  sessionStorage.setItem(
    "lms_features",
    JSON.stringify(features)
  );

  setProgress(100);
  setMessage("Access granted. Redirecting...", "success");

  sessionStorage.removeItem("lms_redirect_after_login");

  const nextPage =
    LOADER_MAP[username] || "../index.html";

  setTimeout(() => {
    window.location.replace(nextPage);
  }, 400);
}

/**************************************************
 * SET PASSWORD
 **************************************************/
async function setPassword(username, password) {
  const url =
    `${AUTH_API}?action=setPassword` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  try {
    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const data = await response.json();

    if (data.status !== "PASSWORD_SET") {
      setMessage(
        data.message || "Password setup failed.",
        "error"
      );
      setProgress(0);
      return;
    }

    setMessage("Password set. Signing in...", "warning");
    setProgress(75);

    const loginUrl =
      `${AUTH_API}?action=loginFull` +
      `&username=${encodeURIComponent(username)}` +
      `&password=${encodeURIComponent(password)}`;

    const loginResponse = await fetch(loginUrl);

    if (!loginResponse.ok) {
      throw new Error(`HTTP ${loginResponse.status}`);
    }

    const loginData = await loginResponse.json();
    await handleLoginResponse(loginData, password);
  } catch (error) {
    console.error("Password setup error:", error);
    setMessage("Password setup failed.", "error");
    setProgress(0);
  }
}

/**************************************************
 * VISIBILITY HELPERS
 **************************************************/
function canAccessDept(deptKey) {
  try {
    const departments = JSON.parse(
      sessionStorage.getItem("lms_departments") || "{}"
    );

    return departments[deptKey] === true;
  } catch (error) {
    console.error("Department visibility parse error:", error);
    return false;
  }
}

function canAccessFeature(featureKey) {
  try {
    const features = JSON.parse(
      sessionStorage.getItem("lms_features") || "{}"
    );

    return features[featureKey] === true;
  } catch (error) {
    console.error("Feature visibility parse error:", error);
    return false;
  }
}

function applyVisibility() {
  document.querySelectorAll("[data-dept]").forEach(element => {
    const department = element.getAttribute("data-dept");
    element.style.display =
      canAccessDept(department) ? "" : "none";
  });

  document.querySelectorAll("[data-feature]").forEach(element => {
    const feature = element.getAttribute("data-feature");
    element.style.display =
      canAccessFeature(feature) ? "" : "none";
  });
}

function requireAuth() {
  const loggedIn =
    sessionStorage.getItem("lms_logged_in") === "true";

  const username =
    sessionStorage.getItem("lms_user") ||
    sessionStorage.getItem("lms_username") ||
    "";

  const visibility =
    sessionStorage.getItem("lms_visibility") || "";

  const savedVersion =
    sessionStorage.getItem("lms_auth_version") || "";

  if (
    !loggedIn ||
    !username ||
    !visibility ||
    savedVersion !== AUTH_VERSION
  ) {
    sessionStorage.clear();
    window.location.replace("../login.html");
    return false;
  }

  return true;
}