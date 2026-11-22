/* Zenni LMS V6.5 — mobile-safe sign-in handoff.
 * Desktop: existing Google Apps Script popup and same-origin return.
 * Mobile: direct top-level navigation to Google's secure LMS page.
 * IMPORTANT: No authorization decision is made here. Apps Script must
 * independently check the signed-in Google account and PAGE_* permissions.
 */
'use strict';

const QA_V6_APP_BASE = 'https://script.google.com/a/macros/zennioptical.com/s/AKfycbzqOAt8pc3gHT3vcnk6G-P3HuxkTHiV4ZrB5u7Gmr9oLPoI-8S2aixMfBWEZa4fk00y/exec';
const QA_V6_POPUP_URL = QA_V6_APP_BASE + '?page=authPopup';
const QA_V6_MOBILE_URL = QA_V6_APP_BASE + '?page=lmsV6';
const QA_V6_RETURN_PATH = './index.html';
const QA_V6_ORIGIN = window.location.origin;
const QA_V6_BUTTON = document.getElementById('loginBtn');
const QA_V6_LABEL = document.getElementById('loginBtnText');
let qaV6Popup = null;

const qaV6Status = document.createElement('div');
qaV6Status.setAttribute('role', 'status');
qaV6Status.setAttribute('aria-live', 'polite');
qaV6Status.style.cssText = 'margin-top:14px;color:#b7d3e8;font:13px/1.5 system-ui;';
QA_V6_BUTTON?.insertAdjacentElement('afterend', qaV6Status);

function qaV6SetStatus(message) {
  qaV6Status.textContent = message;
}
function qaV6ShowLoginLabel() {
  if (QA_V6_LABEL) QA_V6_LABEL.textContent = 'SIGN IN WITH GOOGLE';
}
function qaV6IsMobile() {
  // Detect mobile devices, not merely a narrow desktop browser window.
  const ua = navigator.userAgent || '';
  return /Android|iPhone|iPod|Mobile/i.test(ua) ||
    (/iPad/i.test(ua)) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}
function qaV6Open() {
  if (qaV6IsMobile()) {
    // Avoid cross-origin popup/return chains in Safari, Android browsers,
    // and third-party in-app browsers. Google's server authorizes this page.
    qaV6SetStatus('Opening secure LMS with your Zenni Google account…');
    if (QA_V6_LABEL) QA_V6_LABEL.textContent = 'OPENING LMS…';
    window.location.assign(QA_V6_MOBILE_URL);
    return;
  }

  if (qaV6Popup && !qaV6Popup.closed) {
    qaV6Popup.focus();
    return;
  }
  const width = Math.min(560, window.screen.availWidth || 560);
  const height = Math.min(740, window.screen.availHeight || 740);
  const left = Math.max(0, Math.floor(((window.screen.availWidth || width) - width) / 2));
  const top = Math.max(0, Math.floor(((window.screen.availHeight || height) - height) / 2));
  qaV6Popup = window.open(
    QA_V6_POPUP_URL,
    'ZenniLMSGoogleSignIn',
    `popup=yes,width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`
  );
  if (!qaV6Popup) {
    qaV6SetStatus('Allow pop-ups for this page, then select Sign in with Google again.');
    qaV6ShowLoginLabel();
    return;
  }
  if (QA_V6_LABEL) QA_V6_LABEL.textContent = 'GOOGLE SIGN-IN OPEN';
  qaV6SetStatus('Finish Google verification in the new window.');
  try { qaV6Popup.focus(); } catch (_) {}
}
QA_V6_BUTTON?.addEventListener('click', qaV6Open);

// This message is a navigation hint, not an authentication token.
window.addEventListener('message', (event) => {
  if (event.origin !== QA_V6_ORIGIN || event.source !== qaV6Popup) return;
  const data = event.data;
  if (!data || data.product !== 'ZENNI_LMS_V6' || data.type !== 'RETURN_TO_LMS') return;
  qaV6SetStatus('Opening LMS. Google will verify access again.');
  try { qaV6Popup.close(); } catch (_) {}
  window.location.assign(QA_V6_RETURN_PATH);
});
window.addEventListener('pageshow', qaV6ShowLoginLabel);
