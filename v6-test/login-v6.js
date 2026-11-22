/* V6 TEST ONLY: Github login launcher. Identity is verified by the iframe Apps Script server, not this page. */
'use strict';
document.getElementById('loginBtn')?.addEventListener('click', () => {
  window.location.assign('./index.html');
});
