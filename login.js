/**************************************************
 * CONFIG
 **************************************************/
const AUTH_API =
"https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/**************************************************
 * INIT
 **************************************************/
document.addEventListener("DOMContentLoaded", () => {

  const userEl = document.getElementById("username");

  if (userEl) {
    userEl.addEventListener("input", () => {
      userEl.value = userEl.value.toUpperCase();
    });
  }

});

/**************************************************
 * UI HELPERS
 **************************************************/
function setMessage(text, type = "error") {

  const el = document.getElementById("message");

  el.textContent = text;
  el.className = "login-message " + (type || "");

}

function setLoading(isLoading) {

  const btn = document.getElementById("loginBtn");

  if (!btn) return;

  if (isLoading) {
    btn.disabled = true;
    btn.textContent = "Signing In...";
  } else {
    btn.disabled = false;
    btn.textContent = "Sign In";
  }

}

/**************************************************
 * LOGIN
 **************************************************/
async function login() {

  console.time("LOGIN_TOTAL");

  const usernameEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");

  const username = usernameEl.value.trim().toUpperCase();
  const password = passwordEl.value.trim();

  setMessage("");
  setLoading(true);

  if (!username || !password) {
    setMessage("Enter username and password", "error");
    setLoading(false);
    return;
  }

  const url =
    `${AUTH_API}?action=login` +
    `&username=${encodeURIComponent(username)}` +
    `&password=${encodeURIComponent(password)}`;

  console.time("LOGIN_API");

  try {

    const res = await fetch(url);
    const data = await res.json();

    console.timeEnd("LOGIN_API");

    await handleLoginResponse(data, password);

  } catch (err) {

    setMessage("Connection error. Try again.", "error");

  }

  setLoading(false);

}

/**************************************************
 * HANDLE RESPONSE
 **************************************************/
async function handleLoginResponse(data, originalPassword) {

  if (!data || !data.status) {
    setMessage("Invalid server response", "error");
    return;
  }

  if (data.status === "ERROR") {
    setMessage(data.message || "Login failed", "error");
    return;
  }

  if (data.status === "SET_PASSWORD_REQUIRED") {

    setMessage("First login — setting password…", "warning");

    await setPassword(data.username, originalPassword);
    return;
  }

  if (data.status === "SUCCESS") {

    setMessage("Access Granted", "success");

    sessionStorage.setItem("lms_logged_in", "true");
    sessionStorage.setItem("lms_user", data.username);
    sessionStorage.setItem("lms_role", data.role);

    await loadVisibility(data.username);

  }

}

/**************************************************
 * LOAD VISIBILITY
 **************************************************/
async function loadVisibility(username) {

  console.time("VISIBILITY_API");

  try {

    const res = await fetch(
      `${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`
    );

    const visData = await res.json();

    console.timeEnd("VISIBILITY_API");

    if (visData.status === "SUCCESS") {

      sessionStorage.setItem(
        "lms_visibility",
        JSON.stringify(visData.visibility)
      );

    } else {

      sessionStorage.setItem("lms_visibility", "{}");

    }

  } catch (err) {

    sessionStorage.setItem("lms_visibility", "{}");

  }

  console.timeEnd("LOGIN_TOTAL");

  // slight delay for UX smoothness
  setTimeout(() => {
    window.location.replace("index.html");
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

    const res = await fetch(url);
    const data = await res.json();

    if (data.status !== "PASSWORD_SET") {
      setMessage("Password setup failed", "error");
      return;
    }

    sessionStorage.setItem("lms_logged_in", "true");
    sessionStorage.setItem("lms_user", username);

    await loadVisibility(username);

  } catch (err) {

    setMessage("Password setup failed", "error");

  }

}

/**************************************************
 * ENTER KEY
 **************************************************/
document.addEventListener("keydown", e => {
  if (e.key === "Enter") login();
});