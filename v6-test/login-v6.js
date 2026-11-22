/* Zenni V6 staging — Google login popup; no client-entered email or password. */
'use strict';
const QA_SIGNIN_URL = 'https://script.google.com/a/macros/zennioptical.com/s/AKfycbzqOAt8pc3gHT3vcnk6G-P3HuxkTHiV4ZrB5u7Gmr9oLPoI-8S2aixMfBWEZa4fk00y/exec?page=authPopup';
const QA_ALLOWED_GOOGLE_ORIGINS = new Set(['https://script.google.com', 'https://script.googleusercontent.com']);
let qaPopup = null;
const qaBtn = document.getElementById('loginBtn');
const qaLabel = document.getElementById('loginBtnText');
const qaStatus = document.createElement('p');
qaStatus.setAttribute('role', 'status');
qaStatus.setAttribute('aria-live', 'polite');
qaStatus.style.cssText = 'color:#b8d7f0;font:13px/1.5 system-ui;margin:15px 0 0;';
qaBtn?.insertAdjacentElement('afterend', qaStatus);
function qaSetStatus(value) {qaStatus.textContent = value;}
qaBtn?.addEventListener('click', () => {
  if (qaPopup && !qaPopup.closed) {qaPopup.focus();return;}
  const narrow = matchMedia('(max-width: 700px)').matches;
  const width = Math.min(520, screen.availWidth || 520);
  const height = Math.min(690, screen.availHeight || 690);
  const left = Math.max(0, Math.round((screen.availWidth - width) / 2));
  const top = Math.max(0, Math.round((screen.availHeight - height) / 2));
  // Must open synchronously on user click to avoid popup blockers.
  qaPopup = window.open(QA_SIGNIN_URL, 'ZenniLMSGoogleSignIn',
    narrow ? undefined : `popup=yes,width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`);
  if (!qaPopup) {
    qaSetStatus('Your browser blocked the sign-in window. Allow pop-ups for this site and try again.');
    return;
  }
  qaSetStatus('Complete Google sign-in in the new window.');
  qaLabel.textContent = 'GOOGLE SIGN-IN OPEN';
  try { qaPopup.focus(); } catch (_) {}
});
window.addEventListener('message', (event) => {
  if (!qaPopup || event.source !== qaPopup || !QA_ALLOWED_GOOGLE_ORIGINS.has(event.origin)) return;
  const data = event.data;
  if (!data || data.product !== 'ZENNI_LMS_V6' || data.type !== 'AUTH_RESULT') return;
  if (data.status === 'AUTHORIZED') {
    qaSetStatus('Google verification completed. Opening LMS…');
    try {qaPopup.close();}catch(_){}
    // The destination independently re-verifies the Google identity and permissions.
    window.location.assign('./index.html');
  } else {
    qaLabel.textContent = 'SIGN IN WITH GOOGLE';
    qaSetStatus('Access denied. Contact LMS ADMIN to request registration or permission.');
    // No email or authorization decision is taken from this message for backend use.
  }
});
