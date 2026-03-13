/**************************************************
 * CONFIG
 **************************************************/
const AUTH_API =
"https://script.google.com/macros/s/AKfycbzESjnpNzOyDP76Gm6atwBgh5txV5N2AI225kxz5Q8w7jXgVTIqZrDtIIpQigEE6250/exec";

/**************************************************
 * FORCE USERNAME UPPERCASE
 **************************************************/
document.addEventListener("DOMContentLoaded", () => {

  const userEl = document.getElementById("username");
  if (!userEl) return;

  userEl.addEventListener("input", () => {
    userEl.value = userEl.value.toUpperCase();
  });

});

/**************************************************
 * LOGIN
 **************************************************/
function login() {

  console.time("LOGIN_TOTAL");

  const usernameEl = document.getElementById("username");
  const passwordEl = document.getElementById("password");
  const messageEl = document.getElementById("message");

  const username = usernameEl.value.trim().toUpperCase();
  const password = passwordEl.value.trim();

  messageEl.textContent = "";
  messageEl.style.color = "#f87171";

  if (!username || !password) {
    messageEl.textContent = "Enter username and password";
    return;
  }

  const url =
  `${AUTH_API}?action=login` +
  `&username=${encodeURIComponent(username)}` +
  `&password=${encodeURIComponent(password)}`;

  console.time("LOGIN_API");

  fetch(url)
    .then(res => res.json())
    .then(data => {

      console.timeEnd("LOGIN_API");

      handleLoginResponse(data, password);

    })
    .catch(() => {
      messageEl.textContent = "Connection error";
    });

}

/**************************************************
 * HANDLE RESPONSE
 **************************************************/
function handleLoginResponse(data, originalPassword) {

  const messageEl = document.getElementById("message");

  if (data.status === "ERROR") {
    messageEl.textContent = data.message;
    return;
  }

  if (data.status === "SET_PASSWORD_REQUIRED") {

    messageEl.style.color = "#fbbf24";
    messageEl.textContent = "First login — setting password…";

    setPassword(data.username, originalPassword);
    return;
  }

  if (data.status === "SUCCESS") {

    messageEl.style.color = "#6ee7b7";
    messageEl.textContent = "Login successful";

    sessionStorage.setItem("lms_logged_in", "true");
    sessionStorage.setItem("lms_user", data.username);
    sessionStorage.setItem("lms_role", data.role);

    loadVisibility(data.username);
  }
}

/**************************************************
 * LOAD VISIBILITY
 **************************************************/
function loadVisibility(username) {

  console.time("VISIBILITY_API");

  fetch(`${AUTH_API}?action=visibility&username=${encodeURIComponent(username)}`)
    .then(res => res.json())
    .then(visData => {

      console.timeEnd("VISIBILITY_API");

      if (visData.status === "SUCCESS") {

        sessionStorage.setItem(
          "lms_visibility",
          JSON.stringify(visData.visibility)
        );

      } else {

        sessionStorage.setItem("lms_visibility", "{}");

      }

      console.timeEnd("LOGIN_TOTAL");

      window.location.replace("index.html");

    })
    .catch(() => {

      sessionStorage.setItem("lms_visibility", "{}");

      console.timeEnd("LOGIN_TOTAL");

      window.location.replace("index.html");

    });
}

/**************************************************
 * SET PASSWORD
 **************************************************/
function setPassword(username, password) {

  const messageEl = document.getElementById("message");

  const url =
  `${AUTH_API}?action=setPassword` +
  `&username=${encodeURIComponent(username)}` +
  `&password=${encodeURIComponent(password)}`;

  fetch(url)
    .then(res => res.json())
    .then(data => {

      if (data.status !== "PASSWORD_SET") {
        messageEl.textContent = "Password setup failed";
        return;
      }

      sessionStorage.setItem("lms_logged_in", "true");
      sessionStorage.setItem("lms_user", username);

      loadVisibility(username);

    })
    .catch(() => {
      messageEl.textContent = "Password setup failed";
    });

}

/**************************************************
 * ENTER KEY
 **************************************************/
document.addEventListener("keydown", e => {

  if (e.key === "Enter") login();

});