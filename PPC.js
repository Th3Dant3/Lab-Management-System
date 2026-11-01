"use strict";
/* PPC-APP.js: the PPC page. Google serves a tiny shell that loads this file from GitHub, so everything about how the page
   looks and behaves lives here. The server trusts nothing on this page: every action is re-checked against PPC_USERS and
   your role. Every value that comes from the sheet is inserted with textContent, never as HTML. */
const LMS_URL = "https://th3dant3.github.io/Lab-Management-System/index.html";
const SELF_SRC = (document.currentScript && document.currentScript.src) || "";
const ASSET_BASE = SELF_SRC.replace(/[^\/]*$/, "");
const ASSET_NAME = (SELF_SRC.split("/").pop().split("?")[0] || "PPC-APP.js").replace(/\.js$/i, "");   // PPC-APP: the layout file is PPC-APP.html next to this file
const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
function h(tag, props, kids) {
  const e = document.createElement(tag);
  if (props) for (const k in props) {
    const v = props[k];
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  (Array.isArray(kids) ? kids : kids == null ? [] : [kids]).forEach(c => { if (c != null && c !== false) e.append(c); });
  return e;
}
function setKids(el, kids) { el.replaceChildren(...kids.filter(k => k != null && k !== false)); }
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtDate(s) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? MON[+m[2] - 1] + ' ' + (+m[3]) + ', ' + m[1] : ''; }
function todayUTC() { const d = new Date(); return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()); }
function daysSince(s) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? Math.round((todayUTC() - Date.UTC(+m[1], +m[2] - 1, +m[3])) / 86400000) : null; }
function pl(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }
function yearOf(s) { const m = /^(\d{4})-/.exec(s || ''); return m ? m[1] : 'No date'; }

/* ---------- talking to the server (google.script.run; Google already signed you in when you opened this page) ---------- */
let D = null;
class ApiError extends Error { constructor(code, msg) { super(msg); this.code = code; } }
function setGate(msg, err) { const g = $('#gateStatus'); g.textContent = msg; g.classList.toggle('err', !!err); $('#retry').hidden = !err; }
function api(action, args) {
  return new Promise((resolve, reject) => {
    if (!(window.google && google.script && google.script.run)) return reject(new ApiError('NOT_HOSTED', 'Open this page from the PPC button in the LMS.'));
    google.script.run
      .withSuccessHandler(r => { if (r && r.ok) resolve(r.data); else { const er = (r && r.error) || {}; reject(new ApiError(er.code || 'SERVER', er.message || 'Unexpected reply from the server.')); } })
      .withFailureHandler(() => reject(new ApiError('NETWORK', 'Could not reach the PPC server. Reload the page. If it keeps happening, sign in to Google again.')))
      .ppcCall(action, args || {});
  });
}
function gateError(e) {
  const m = {
    FORBIDDEN: 'Your Google account is not set up for PPC. Ask the administrator to add you to PPC_USERS.',
    API_OFF: 'The PPC web page is switched off by the administrator.'
  };
  setGate(m[e.code] || e.message, true);
  $('#gate').hidden = false; $('#shellStaff').hidden = true; $('#shellReq').hidden = true;
}
async function loadApp() {
  setGate('Loading your PPC requests...');
  try { D = await api('bootstrap'); } catch (e) { gateError(e); return; }
  showApp();
}
async function refresh() { D = await api('bootstrap'); reqSel = null; renderCurrent(); }

/* ---------- shared pieces ---------- */
const STATUS_COLS = [
  { status: 'Not Started', title: 'Not started', cls: '' },
  { status: 'Rough Draft in Progress', title: 'Rough draft in progress', cls: 'draft' },
  { status: 'Ready for Signature', title: 'Ready for signature', cls: 'sign' },
  { status: 'In Routing', title: 'In routing', cls: 'route' },
  { status: 'Published', title: 'Published', cls: 'done' }
];
function ageBar(days, maxD) { const s = h('span'); s.style.width = (maxD ? Math.max(2, Math.round(days / maxD * 100)) : 0) + '%'; return h('div', { class: 'age', role: 'img', 'aria-label': days + ' days open' }, s); }
function btn(label, o) { o = o || {}; return h('button', { type: 'button', class: 'btn' + (o.quiet ? ' quiet' : '') + (o.primary ? ' primary' : ''), disabled: o.disabled, title: o.title, onclick: o.onclick, 'data-open': o.open, 'data-ppc': o.ppc }, label); }
function table(heads, rows, empty) {
  if (!rows.length) return h('p', { class: 'empty', text: empty });
  return h('div', { class: 'scroll' }, h('table', null, [h('thead', null, h('tr', null, heads.map(x => h('th', { text: x })))), h('tbody', null, rows)]));
}
function td(content, cls) { return h('td', { class: cls }, content); }
function openRows() { return D.register.filter(r => r.status !== 'Published'); }
function maxAge() { const a = openRows().concat(D.needsNumber).map(r => daysSince(r.submitted)).filter(x => x !== null); return a.length ? Math.max.apply(null, a) : 0; }
const NOT_BUILT = 'Not switched on yet. This button needs a backend step that is built after sign-in is confirmed.';

/* ---------- staff views ---------- */
const VIEWS = [['overview', 'Dashboard'], ['missing', 'Missing PPC, need update'], ['tests', 'Tests and reminders'], ['routing', 'Routing and emails'], ['updates', 'Update requests'], ['years', 'Year records']];
let currentView = 'overview', currentYear = null, mailSel = null, mailEv = 'routed';
function vOverview() {
  const open = openRows(), need = D.needsNumber, maxD = maxAge();
  const pubTotal = D.published.length + D.register.filter(r => r.status === 'Published').length;
  const items = [];
  if (need.length) items.push([h('span', null, [h('b', { text: pl(need.length, 'open request') + ' ' + (need.length === 1 ? 'has' : 'have') + ' no PPC number.' }), ' The oldest has waited ' + Math.max.apply(null, need.map(r => daysSince(r.submitted) || 0)) + ' days.']), 'Open Missing PPC, need update', 'missing', '']);
  const noTest = open.filter(r => r.stage === 'Testing' && !D.tests.some(t => t.ppc === r.number));
  if (noTest.length) items.push([h('span', null, [h('b', { text: pl(noTest.length, 'PPC') + ' ' + (noTest.length === 1 ? 'is' : 'are') + ' in the Testing stage with no test recorded: ' }), noTest.map(r => r.number).join(' and ') + '. Without a test end date, no reminder can be sent.']), 'Open Tests and reminders', 'tests', 'amber']);
  const routing = open.filter(r => r.status === 'In Routing'), late = routing.filter(r => (daysSince(r.routing) || 0) > 30);
  if (routing.length) items.push([h('span', null, [h('b', { text: pl(routing.length, 'PPC') + ' ' + (routing.length === 1 ? 'is' : 'are') + ' in routing.' }), late.length ? ' ' + late.length + ' ' + (late.length === 1 ? 'has' : 'have') + ' been there more than a month: ' + late.map(r => daysSince(r.routing) + ' days').join(' and ') + '.' : '']), 'Open Routing and emails', 'routing', 'blue']);
  const odd = D.register.filter(r => !STATUS_COLS.some(c => c.status === r.status));
  const att = h('section', null, [h('h2', { text: 'Needs attention' }), items.length ? h('ul', { class: 'attn' }, items.map(i => h('li', null, [h('div', null, [h('span', { class: 'dot ' + i[3] }), i[0]]), h('a', { class: 'btn', href: '#' + i[2], onclick: e => { e.preventDefault(); show(i[2]); } }, i[1])]))) : h('p', { class: 'empty', text: 'Nothing needs attention right now.' })]);
  const cols = STATUS_COLS.map(c => {
    const rows = D.register.filter(r => r.status === c.status), head = h('header', null, [h('h3', { text: c.title }), h('span', { class: 'c', text: String(c.status === 'Published' ? pubTotal : rows.length) })]);
    if (c.status === 'Published') return h('div', { class: 'col done' }, [head, h('div', { class: 'arch' }, h('p', { text: 'Completed requests live in the archive, not in this board.' }))]);
    return h('div', { class: 'col ' + c.cls }, [head, rows.length ? h('ul', null, rows.map(r => {
      const d = daysSince(r.submitted), kids = [h('p', { class: 'num', text: r.number }), h('p', { class: 'ttl', text: r.title || 'Title not recorded' }), h('p', { class: 'who', text: (r.requestor || 'Requester not recorded') + (r.wi ? ', ' + r.wi : '') })];
      if (r.status === 'In Routing' && r.routing) kids.push(h('p', { class: 'sub', text: 'In routing since ' + fmtDate(r.routing) + ' (' + daysSince(r.routing) + ' days)' })); else if (r.stage) kids.push(h('p', { class: 'sub', text: r.stage }));
      if (d !== null) { kids.push(ageBar(d, maxD)); kids.push(h('p', { class: 'days', text: d + ' days open' })); }
      return h('li', { class: 'tok ' + c.cls }, kids);
    })) : h('p', { class: 'none', text: c.status === 'Not Started' ? 'No PPC is waiting to start.' : 'No PPC is waiting for management signatures.' })]);
  });
  return h('div', { class: 'view', id: 'v-overview' }, [h('h1', { text: 'PPC command center' }),
    h('p', { class: 'lead' }, [h('strong', { text: pl(open.length + need.length, 'request') + (open.length + need.length === 1 ? ' is' : ' are') + ' open. ' }), open.length + ' ' + (open.length === 1 ? 'has' : 'have') + ' a PPC number and ' + need.length + ' ' + (need.length === 1 ? 'is' : 'are') + ' still waiting for one.' + (maxD ? ' The oldest open request has been waiting ' + maxD + ' days.' : '')]),
    odd.length ? h('div', { class: 'banner' }, 'Fix in the sheet: ' + pl(odd.length, 'register row') + ' ' + (odd.length === 1 ? 'has' : 'have') + ' a status that is not one of the five (' + odd.map(r => r.number + ' says "' + r.status + '"').join('; ') + '). ' + (odd.length === 1 ? 'It is' : 'They are') + ' not shown on the board.') : null,
    att, h('section', null, [h('h2', { text: 'Where the numbered PPCs are' }), h('p', { class: 'hint', text: 'Days open count from the submitted date in the sheet.' }), h('div', { class: 'flow' }, cols)])]);
}
function vMissing() {
  const need = D.needsNumber, maxD = Math.max.apply(null, [0].concat(need.map(r => daysSince(r.submitted) || 0)));
  const rows = need.map(r => { const d = daysSince(r.submitted); return h('tr', null, [td([h('b', { text: r.wi || 'No WI recorded' }), h('br'), h('span', { class: 'mut', text: r.title || 'Title not recorded' })]), td(r.requestor || ''), td(fmtDate(r.submitted)), td(d === null ? '' : [ageBar(d, maxD), h('span', { class: 'n', text: String(d) })], 'bar'), td(r.status + (r.stage ? ', ' + r.stage : '')), td([btn('Issue PPC number', { disabled: true, title: NOT_BUILT }), ' ', btn('Close request', { quiet: true, disabled: true, title: NOT_BUILT })], 'act')]); });
  return h('div', { class: 'view', id: 'v-missing', hidden: true }, [h('h1', { text: 'Missing PPC, need update' }),
    h('p', { class: 'lead' }, [h('strong', { text: need.length ? pl(need.length, 'open request') + ' ' + (need.length === 1 ? 'has' : 'have') + ' no PPC number.' : 'Every open request has a PPC number.' }), need.length ? ' New requests with no WI number will land here too.' : '']),
    h('section', null, [h('h2', { text: 'Decide each one' }), h('p', { class: 'hint', text: 'The buttons are visible so you can see the plan. They are switched on in a later step. No number is ever guessed.' }), table(['Request', 'Requested by', 'Submitted', 'Days open', 'Status in the sheet', 'Action'], rows, 'Nothing is waiting for a PPC number.')])]);
}
function vTests() {
  const open = openRows(), cfg = D.config, noTest = open.filter(r => r.stage === 'Testing' && !D.tests.some(t => t.ppc === r.number));
  const rows1 = noTest.map(r => h('tr', null, [td([h('b', { text: r.number }), h('br'), h('span', { class: 'mut', text: r.title })]), td(r.requestor || ''), td('No test recorded'), td(btn('Record a test', { open: 'test', ppc: r.number }), 'act')]));
  const rows2 = D.tests.map(t => { const d = daysSince(t.end); return h('tr', null, [td([h('b', { text: t.ppc }), h('br'), h('span', { class: 'mut', text: t.id })]), td([t.owner, h('br'), h('span', { class: 'mut', text: t.ownerEmail })]), td(fmtDate(t.end) + (d === null ? '' : d > 0 ? ', ended ' + pl(d, 'day') + ' ago' : d === 0 ? ', ends today' : ', ends in ' + pl(-d, 'day'))), td(t.result), td(btn('Send reminder now', { quiet: true, disabled: true, title: NOT_BUILT }), 'act')]); });
  const kv = [['Automatic check', 'Once a day, around 8 AM, after the daily trigger is installed.'], ['When it reminds', 'From ' + cfg.reminderLeadDays + ' days before the test end date, and every ' + cfg.reminderCooldownDays + ' days after it passes, until the PPC is published.'], ['Who gets it', 'The document specialist, or the requester if no specialist is recorded. Adding the test owner is a backend change.'], ['Minimum test length', cfg.minTestDays + ' calendar days, enforced by the server as well as the form.'], ['Email mode', cfg.emailMode === 'LIVE' ? 'Live. Reminders are emailed and logged.' : 'Dry run. Reminders are logged but not sent until PPC_CONFIG EMAIL_MODE is LIVE.']];
  return h('div', { class: 'view', id: 'v-tests', hidden: true }, [h('h1', { text: 'Tests and reminders' }),
    h('p', { class: 'lead' }, [h('strong', { text: D.tests.length ? pl(D.tests.length, 'test') + ' recorded.' : 'No test has been recorded yet, so no reminder is scheduled.' })]),
    h('section', null, [h('h2', { text: 'PPCs in the Testing stage with no test record' }), table(['PPC', 'Requested by', 'Test', 'Action'], rows1, 'Every PPC in the Testing stage has a test.')]),
    h('section', null, [h('h2', { text: 'Recorded tests' }), table(['PPC and test', 'Owner', 'Ends', 'Result', 'Action'], rows2, 'No tests recorded.')]),
    h('section', null, [h('h2', { text: 'How reminders work' }), h('dl', { class: 'kv' }, kv.flatMap(p => [h('dt', { text: p[0] }), h('dd', { text: p[1] })]))])]);
}
function mailBox() {
  const r = D.register.find(x => x.number === mailSel), box = h('div', { id: 'mailbox' });
  if (!r) return box;
  const note = r.publicNotes || '', imported = /^Imported from legacy source/.test(note);
  box.append(h('p', { class: 'mailhead' }, ['To: ', h('b', { text: r.requestorEmail || '(no email on file, nothing would be sent)' }), h('br'), 'Subject: ', h('b', { text: 'PPC ' + mailEv + ': ' + r.number })]),
    h('div', { class: 'mail', text: 'PPC ' + r.number + '\nDocument: ' + (r.title || '') + '\nStage: ' + (mailEv === 'published' ? 'Published' : 'Approval / Routing') + '\nPublic update: ' + (note || 'Please contact your document specialist.') }));
  if (imported) box.append(h('p', { class: 'warnbox', text: 'The public update is still the import note. Write a real update, or clear it, before email goes live. When it is blank, the email says "Please contact your document specialist."' }));
  box.append(h('p', { class: 'mut', text: 'Sent from the script owner\'s Google account. Email mode: ' + (D.config.emailMode === 'LIVE' ? 'Live.' : 'Dry run, so nothing is sent yet.') + ' Every attempt is logged in PPC_AUDIT.' }));
  return box;
}
function vRouting() {
  const routing = openRows().filter(r => r.status === 'In Routing'); if (!mailSel && routing.length) mailSel = routing[0].number;
  const rows = routing.map(r => { const d = daysSince(r.routing); return h('tr', null, [td([h('b', { text: r.number }), h('br'), h('span', { class: 'mut', text: r.title })]), td([r.requestor || '', h('br'), h('span', { class: 'mut', text: r.requestorEmail || 'no email on file' })]), td(r.routing ? [fmtDate(r.routing), h('br'), h('span', { class: 'mut', text: d + ' days' })] : 'No routing date'), td([btn('Preview email', { quiet: true, onclick: () => { mailSel = r.number; renderMail(); } }), ' ', btn('Mark published', { disabled: true, title: NOT_BUILT })], 'act')]); });
  const seg = (ev, label) => h('button', { type: 'button', class: 'seg' + (mailEv === ev ? ' on' : ''), 'aria-pressed': String(mailEv === ev), onclick: () => { mailEv = ev; renderMail(); } }, label);
  return h('div', { class: 'view', id: 'v-routing', hidden: true }, [h('h1', { text: 'Routing and emails' }),
    h('p', { class: 'lead' }, [h('strong', { text: pl(routing.length, 'PPC') + (routing.length === 1 ? ' is' : ' are') + ' in routing.' }), ' Marking a PPC routed or published emails the requester.']),
    h('section', null, [h('h2', { text: 'In routing now' }), table(['PPC', 'Requested by', 'In routing since', 'Action'], rows, 'No PPC is in routing.')]),
    h('section', null, [h('h2', { text: 'Email the requester receives' }), h('div', { class: 'tabs2', role: 'group', 'aria-label': 'Email event' }, [seg('routed', 'Routed'), seg('published', 'Published')]), mailBox()])]);
}
function renderMail() { const v = $('#v-routing'); if (!v) return; const old = v.children[3]; const fresh = vRouting(); v.replaceWith(fresh); fresh.hidden = currentView !== 'routing'; }
function vUpdates() {
  return h('div', { class: 'view', id: 'v-updates', hidden: true }, [h('h1', { text: 'Update requests' }), h('p', { class: 'lead' }, [h('strong', { text: 'Not built yet.' }), ' Requesters cannot ask for an update from the page until this exists.']),
    h('section', null, [h('h2', { text: 'The plan' }), h('ol', { class: 'steps4' }, ['A requester opens a request on their page and chooses Request an update.', 'It appears here, and the document specialist gets an email.', 'You write a short public update on the PPC. Internal notes are never shown to requesters.', 'You mark it answered, and the requester gets an email with the status and your note.'].map(t => h('li', { text: t })))])]);
}
function allRows() {
  return [].concat(D.register.map(r => Object.assign({ tab: r.status === 'Published' ? 'published' : 'register' }, r)), D.published.map(r => Object.assign({ tab: 'published' }, r)), D.blocked.map(r => Object.assign({ tab: 'blocked' }, r)), D.needsNumber.map(r => Object.assign({ tab: 'needsNumber' }, r)));
}
function yearStats(rows) {
  const s = { total: rows.length, done: 0, doneDated: 0, open: 0, need: 0, blocked: 0 };
  rows.forEach(r => { if (r.tab === 'published') { s.done++; if (r.published) s.doneDated++; } else if (r.tab === 'blocked') s.blocked++; else if (r.tab === 'needsNumber') s.need++; else s.open++; });
  return s;
}
function rowStatus(r) { return r.tab === 'published' ? 'Published' : r.tab === 'blocked' ? 'Marked blocked' : r.tab === 'needsNumber' ? 'Open, no PPC number' : (r.status || 'Open'); }
function vYears() {
  const rows = allRows(), by = {}; rows.forEach(r => { (by[yearOf(r.submitted)] = by[yearOf(r.submitted)] || []).push(r); });
  const years = Object.keys(by).sort(); if (!currentYear || !by[currentYear]) currentYear = years.filter(y => y !== 'No date').pop() || years[0];
  const idx = years.map(y => { const s = yearStats(by[y]); return h('tr', null, [td(h('b', { text: y })), td(String(s.total)), td(String(s.done)), td(String(s.open)), td(String(s.need)), td(String(s.blocked)), td(s.doneDated + ' of ' + s.done)]); });
  const s = currentYear ? yearStats(by[currentYear]) : null, facts = [], open = s ? s.open + s.need : 0;
  if (s) {
    facts.push(h('li', null, [h('b', { text: s.done + ' of ' + pl(s.total, 'request') + (s.done === 1 ? ' is' : ' are') + ' complete.' })]));
    if (s.done && s.doneDated < s.done) facts.push(h('li', null, [h('b', { text: (s.done - s.doneDated) + ' of those have no published date' }), ' in the notes, so their completion month is unknown.' + (s.doneDated ? ' The other ' + s.doneDated + ' do.' : '')]));
    if (s.open) facts.push(h('li', null, [h('b', { text: pl(s.open, 'PPC') + (s.open === 1 ? ' is' : ' are') + ' in progress.' })]));
    if (s.need) facts.push(h('li', null, [h('b', { text: pl(s.need, 'open request') + (s.need === 1 ? ' has' : ' have') + ' no PPC number.' })]));
    if (s.blocked) facts.push(h('li', null, [h('b', { text: pl(s.blocked, 'request') + (s.blocked === 1 ? ' is' : ' are') + ' marked blocked.' })]));
    if (!open) facts.push(h('li', null, [h('b', { text: 'Nothing is open.' }), s.blocked ? ' Close or resolve the blocked rows to finish this record.' : '']));
  }
  const yb = years.map(y => h('button', { type: 'button', class: 'seg' + (y === currentYear ? ' on' : ''), 'aria-pressed': String(y === currentYear), onclick: () => { currentYear = y; rebuild('years'); } }, y));
  let bar = null;
  if (s && s.total) { bar = h('div', { class: 'split tall', role: 'img', 'aria-label': 'Split of ' + currentYear }); [['done', s.done], ['draft', s.open], ['steel', s.need], ['block', s.blocked]].forEach(p => { if (p[1]) { const i = h('i', { class: p[0] }); i.style.width = (p[1] / s.total * 100) + '%'; bar.append(i); } }); }
  return h('div', { class: 'view', id: 'v-years', hidden: true }, [h('h1', { text: 'Year records' }), h('p', { class: 'lead', text: 'Each request belongs to the year it was submitted. A request submitted in one year and finished in the next stays in the year it was submitted.' }),
    h('section', null, [h('h2', { text: 'All years' }), table(['Year', 'Requests', 'Completed', 'In progress, numbered', 'Missing PPC number', 'Marked blocked', 'Completed with a published date'], idx, 'No requests yet.')]),
    s ? h('section', null, [h('div', { class: 'yr-head' }, [h('h2', { text: 'Open one year' }), h('div', { role: 'group', 'aria-label': 'Choose a year' }, yb)]), h('h3', { class: 'ytitle', text: currentYear + ' record' }), h('p', { class: 'state', text: open ? 'Open: ' + pl(open, 'request') + ' still in progress' : 'Nothing open' }), bar, h('ul', { class: 'facts' }, facts),
      h('div', { class: 'tools' }, [btn('Print year record', { onclick: () => printRecord(by[currentYear]) }), btn('Download as CSV', { quiet: true, onclick: () => downloadCsv(by[currentYear]) })])]) : null]);
}
function recordRows(list) { return list.slice().sort((a, b) => (a.submitted || '').localeCompare(b.submitted || '')).map(r => [r.number || 'Pending', r.wi || '', r.title || '', r.requestor || '', fmtDate(r.submitted), rowStatus(r), fmtDate(r.published)]); }
const RECORD_HEADS = ['PPC number', 'Work instruction', 'Title', 'Requested by', 'Submitted', 'Status', 'Published'];
function printRecord(list) {
  const pa = $('#printArea'); pa.replaceChildren(h('h1', { text: 'PPC record, ' + currentYear }), h('p', { text: pl(list.length, 'request') + '. Printed ' + fmtDate(new Date().toISOString().slice(0, 10)) + ' by ' + D.user.name + '.' }), table(RECORD_HEADS, recordRows(list).map(r => h('tr', null, r.map(c => h('td', { text: c })))), ''));
  document.body.classList.add('printing-record'); window.print();
}
function csvCell(v) { let s = String(v == null ? '' : v); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return '"' + s.replace(/"/g, '""') + '"'; }
function downloadCsv(list) {
  const lines = [RECORD_HEADS].concat(recordRows(list)).map(r => r.map(csvCell).join(','));
  const a = h('a', { href: URL.createObjectURL(new Blob([lines.join('\r\n')], { type: 'text/csv' })), download: 'ppc-record-' + currentYear + '.csv' }); document.body.append(a); a.click(); a.remove();
}
window.addEventListener('afterprint', () => document.body.classList.remove('printing-record', 'printing-wz'));
const BUILDERS = { overview: vOverview, missing: vMissing, tests: vTests, routing: vRouting, updates: vUpdates, years: vYears };
function rebuild(id) { const old = $('#v-' + id); const fresh = BUILDERS[id](); old.replaceWith(fresh); fresh.hidden = currentView !== id; }
function show(id) {
  if (!BUILDERS[id]) id = 'overview'; currentView = id;
  $$('#views .view').forEach(v => { v.hidden = v.id !== 'v-' + id; });
  $$('#nav a').forEach(a => { if (a.dataset.view === id) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
  try { history.replaceState(null, '', '#' + id); } catch (e) { /* the page lives in a sandboxed frame; the address bar is not ours */ } window.scrollTo(0, 0);
}
function banner(el) {
  el.hidden = D.writes; if (!D.writes) el.textContent = 'Read-only for now. You can look around and fill in the forms, but nothing is saved from this page yet.';
}
function renderStaff() {
  $('#nav').replaceChildren(...VIEWS.map(v => h('a', { href: '#' + v[0], 'data-view': v[0], onclick: e => { e.preventDefault(); show(v[0]); } }, [v[1], v[0] === 'missing' && D.needsNumber.length ? h('span', { class: 'b', text: String(D.needsNumber.length) }) : null])));
  setKids($('#topbar'), [D.can.test ? btn('New PPC test', { open: 'test' }) : null, D.can.request ? btn('New PPC request', { open: 'request', primary: true }) : null]);
  $('#who').textContent = D.user.name + ', ' + D.user.role.toLowerCase();
  banner($('#banner')); $('#views').replaceChildren(...Object.keys(BUILDERS).map(id => BUILDERS[id]()));
  wireOpeners(); show(currentView || (location.hash || '').slice(1));
}

/* ---------- requester view ---------- */
const STEPS = ['Request received', 'Drafting and testing', 'Ready for signature', 'In routing', 'Published'];
let reqSel = null;
function whereText(r) {
  if (r.tab === 'blocked') return 'On hold';
  if (r.state === 'Complete') return 'Published' + (r.published ? ' ' + fmtDate(r.published) : '');
  if (r.numberPending) return 'Waiting for a PPC number from document control';
  if (r.status === 'In Routing') return 'In routing' + (r.routing ? ' since ' + fmtDate(r.routing) : '');
  if (r.status === 'Ready for Signature') return 'Waiting for management signatures';
  if (r.status === 'Not Started') return 'Not started yet';
  return r.stage === 'Testing' ? 'Testing' : 'Drafting the work instruction';
}
function renderRequester() {
  $('#rwho').textContent = D.user.name; banner($('#rbanner'));
  setKids($('#rtop'), [D.can.request ? btn('New PPC request', { open: 'request', primary: true }) : null]);
  const mine = D.mine, cnt = s => mine.filter(r => r.state === s).length;
  const row = r => h('li', null, h('button', { type: 'button', class: 'req', 'aria-pressed': String(reqSel === r), onclick: () => { reqSel = r; renderRequester(); if (window.innerWidth < 980) $('#detail').scrollIntoView({ behavior: 'smooth', block: 'start' }); } }, [
    h('span', { class: 'rn' }, [h('b', { text: r.number || '' }), r.number ? null : h('span', { class: 'mut', text: 'PPC number pending' }), h('br'), h('span', { class: 'mut', text: r.wi })]), h('span', { class: 'rt', text: r.title || 'Title not recorded' }), h('span', { class: 'rd', text: fmtDate(r.submitted) }), h('span', { class: 'chip2 ' + (r.state === 'Complete' ? 'c' : 'p'), text: r.state })]));
  const group = (name) => { const l = mine.filter(r => r.state === name); return l.length ? [h('h3', { class: 'gh' }, [name + ' ', h('span', { class: 'mut', text: String(l.length) })]), h('ul', { class: 'reqs' }, l.map(row))] : []; };
  const det = h('aside', { class: 'detail', id: 'detail', 'aria-live': 'polite' });
  if (!reqSel) det.append(h('h2', { text: mine.length ? 'Choose a request' : 'No requests yet' }), h('p', { class: 'sub', text: mine.length ? 'Select one on the left to see its progress.' : 'Nothing is recorded under your email address.' }));
  else {
    const r = reqSel; det.append(h('h2', { text: r.number || 'PPC number pending' }), h('p', { class: 'sub', text: (r.title || 'Title not recorded') + (r.wi ? ' (' + r.wi + ')' : '') + ', submitted ' + fmtDate(r.submitted) }),
      h('ol', { class: 'tl' }, STEPS.map((s, i) => h('li', { class: r.state === 'Complete' ? (i < 4 ? 'done' : 'now') : (i < r.step ? 'done' : i === r.step ? 'now' : ''), text: s }))),
      h('p', null, [h('b', { text: 'Where it is: ' }), whereText(r)]), h('h3', { class: 'nh', text: 'Latest note from document control' }),
      r.publicNotes && !/^Imported from legacy source/.test(r.publicNotes) ? h('p', { class: 'note2', text: r.publicNotes }) : h('p', { class: 'note2 mut', text: 'No update has been written yet.' }),
      h('p', { class: 'mut', text: 'Asking document control for an update from this page is not built yet.' }));
  }
  setKids($('#rview'), [h('h1', { text: 'My requests' }), h('p', { class: 'lead' }, ['You have ', h('strong', { text: pl(mine.length, 'request') }), '. Pick one to see where it is.']),
    h('p', { class: 'keyline' }, ['Not started', 'In progress', 'Complete'].map(s => h('span', null, [h('b', { text: String(cnt(s)) }), s]))),
    D.can.request ? null : h('p', { class: 'mut', text: 'Sending a new request from this page is not turned on for your account.' }),
    h('div', { class: 'cols' }, [h('div', null, [].concat(group('In progress'), group('Complete'), mine.length ? [] : [h('p', { class: 'empty', text: 'No requests found.' })])), det])]);
  wireOpeners();
}
function renderCurrent() { if (!D) return; if (D.view === 'staff') renderStaff(); else renderRequester(); }
function showApp() {
  if (!WZ.request) initWizards();
  $('#gate').hidden = true; $('#shellStaff').hidden = D.view !== 'staff'; $('#shellReq').hidden = D.view !== 'requester';
  document.title = 'PPC command center, Zenni USA'; renderCurrent();
}

/* ---------- wizards (your 3-step format) ---------- */
const openers = new WeakSet();
function wireOpeners() { $$('[data-open]').forEach(b => { if (openers.has(b)) return; openers.add(b); b.addEventListener('click', () => (b.dataset.open === 'request' ? WZ.request : WZ.test).open(b, b.dataset.ppc || null)); }); }
const WZ = {};
function ppcsForWizard() { const src = D.view === 'staff' ? D.register : D.mine.filter(m => m.tab === 'register' && m.number); return src.filter(r => r.status !== 'Published' && r.number).map(r => ({ n: r.number, dept: r.number.split('-')[1], title: r.title, wi: r.wi, status: r.status })); }
function wizard(ov, o) {
  const wz = $('.wz', ov); let step = 1, opener = null;
  const label = $('[data-label]', wz);
  function set(n, force) {
    if (n < 1 || n > 3) return;
    if (!force && n > step) { const e = step === 1 ? o.v1() : []; if (e.length) { label.textContent = e[0]; return; } }
    step = n; wz.dataset.step = n;
    $$('[data-go]', wz).forEach(b => { const k = +b.dataset.go; b.classList.toggle('on', k === n); b.classList.toggle('done', k < n); });
    $('[data-back]', wz).style.visibility = n === 1 ? 'hidden' : 'visible'; label.textContent = 'Step ' + n + ' of 3';
    const nx = $('[data-next]', wz); nx.textContent = n === 1 ? o.l1 : n === 2 ? 'Next: Review & Submit →' : o.submit; nx.disabled = false;
    if (n === 3) o.review(); ov.scrollTop = 0;
  }
  $('[data-next]', wz).addEventListener('click', () => { if (step < 3) set(step + 1); else o.send(); });
  $('[data-back]', wz).addEventListener('click', () => set(step - 1, true));
  $$('[data-go]', wz).forEach(b => b.addEventListener('click', () => { const k = +b.dataset.go; if (k <= step) set(k, true); else if (k === step + 1) set(k); }));
  const close = () => { ov.hidden = true; if (opener && opener.focus) opener.focus(); };
  $$('[data-close]', wz).forEach(b => b.addEventListener('click', close));
  $('[data-print]', wz).addEventListener('click', () => { set(2, true); document.body.classList.add('printing-wz'); window.print(); });
  ov.addEventListener('click', e => { if (e.target === ov) close(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !ov.hidden) close(); });
  ['input', 'change'].forEach(t => ov.addEventListener(t, () => { if (!/^Step \d of 3$/.test(label.textContent)) label.textContent = 'Step ' + step + ' of 3'; }));
  $$('form', ov).forEach(f => f.addEventListener('submit', e => e.preventDefault()));
  return { open(btn, pre) { opener = btn || null; ov.hidden = false; if (o.reset) o.reset(pre); set(1, true); const f = ov.querySelector('select,input'); if (f) f.focus(); } };
}
function initWizards() {
  const rq = $('#rqOverlay'), ts = $('#tsOverlay'), esc = null;
  const dept = $('#rqDept'), act = $('#rqAction'), ex = $('#rqExist'), wi = $('#rqWi'), meth = $('#rqMethod'), man = $('#rqManual'), prev = $('#rqPrev');
  const pad = n => ('000' + n).slice(-3), val = s => $(s).value;
  function fillEx() { const d = dept.value, l = ppcsForWizard().filter(p => p.dept === d); ex.replaceChildren(h('option', { value: '', text: d ? (l.length ? 'Select an existing PPC' : 'No open PPC in this department') : 'Choose a department first' }), ...l.map(p => h('option', { value: p.n, text: p.n + ' (' + p.title + ')' }))); }
  function numText() {
    const d = dept.value, a = act.value; if (!d) return 'Choose a department';
    if (a === 'EXISTING') return ex.value || 'Select an existing PPC';
    if (meth.value === 'MANUAL') return man.value.trim().toUpperCase() || 'Enter a manual PPC number';
    const m = wi.value.match(/WI-([A-Za-z]{2,8})-(\d{1,6})/);
    if (!m) return 'Not issued yet. Document control adds the WI number first. The app never guesses one.';
    if (m[1].toUpperCase() !== d) return 'The WI says ' + m[1].toUpperCase() + ' but the department is ' + d + '. Fix one of them.';
    return D.year + '-' + d + '-' + pad(m[2]) + '-' + D.nextSequence + ' (preview only, the final number is issued on submit)';
  }
  function rqRefresh() {
    const a = act.value; $('#rqExistWrap').hidden = a !== 'EXISTING'; $('#rqMethodWrap').hidden = a === 'EXISTING'; $('#rqManualWrap').hidden = a === 'EXISTING' || meth.value !== 'MANUAL';
    if (a === 'EXISTING') { const p = ppcsForWizard().find(x => x.n === ex.value); wi.value = p ? p.wi : ''; wi.readOnly = true; wi.placeholder = 'Select an existing PPC to load its WI number(s)'; $('#rqWiHelp').textContent = 'WI numbers load automatically from the PPC selected above.'; }
    else { wi.readOnly = false; wi.placeholder = 'Example: WI-SUR-246 (leave blank if the document is new)'; $('#rqWiHelp').textContent = 'Existing document: enter its WI number. New document: leave blank.'; }
    const t = numText(); prev.textContent = t; $('#rqPpcPrint').value = t.split(' (')[0].split('. ')[0]; $('#rqWiPrint').value = wi.value;
  }
  dept.addEventListener('change', () => { fillEx(); rqRefresh(); }); [act, ex, meth].forEach(e => e.addEventListener('change', rqRefresh)); [wi, man].forEach(e => e.addEventListener('input', rqRefresh));
  WZ.request = wizard(rq, {
    l1: 'Next: Fill Request Form →', submit: 'Submit PPC Proposal',
    reset() { dept.value = ''; act.value = 'NEW'; meth.value = 'AUTO'; man.value = ''; wi.value = ''; $('#rqForm').reset(); $('#rqDone').hidden = true; $('#rqDate').value = new Date().toISOString().slice(0, 10); $('#rqInit').value = D.user.name; $('#rqEmail').value = D.user.email; $('.wz-nav [data-next]', rq).disabled = false; fillEx(); rqRefresh(); },
    v1() { const e = []; if (!dept.value) e.push('Select a requesting department.'); if (act.value === 'EXISTING' && !ex.value) e.push('Choose the existing PPC.'); if (act.value === 'NEW' && meth.value === 'MANUAL' && !/^\d{4}-[A-Z]{2,8}-\d{3,6}-\d{1,6}$/.test(man.value.trim().toUpperCase())) e.push('Enter a manual number like 2026-SUR-246-014.'); if (/^The WI says/.test(numText())) e.push('The WI department does not match.'); return e; },
    review() {
      const f = [['Department', dept.value ? dept.options[dept.selectedIndex].text : 'Not selected'], ['WI number(s)', wi.value || 'None yet'], ['PPC mode', act.value === 'NEW' ? 'New PPC' : 'Existing PPC'], ['PPC number', numText()], ['Initiator', val('#rqInit') || 'Missing'], ['Change title', val('#rqChTitle') || 'Missing'], ['Target date', val('#rqDue') || 'Not provided'], ['Description', val('#rqDesc') || 'Missing'], ['Justification', val('#rqJust') || 'Missing'], ['Email for updates', val('#rqEmail') || 'Missing']];
      $('#rqRev').replaceChildren(...f.flatMap(r => [h('dt', { text: r[0] }), h('dd', { text: r[1] })]));
      const m = []; if (!val('#rqInit').trim()) m.push('Initiator'); if (!val('#rqChTitle').trim()) m.push('Change title'); if (!val('#rqDesc').trim()) m.push('Description'); if (!val('#rqJust').trim()) m.push('Justification'); if (!val('#rqEmail').trim()) m.push('Email');
      const mm = $('#rqMiss'); mm.textContent = m.length ? 'Please complete: ' + m.join(', ') + '.' : 'All core fields are entered. Check the details, then submit.'; mm.className = m.length ? 'wz-miss' : 'wz-ok'; $('#rqDone').hidden = true;
    },
    async send() {
      const out = $('#rqDone'), nx = $('.wz-nav [data-next]', rq); out.hidden = false; out.textContent = 'Saving...'; nx.disabled = true;
      try {
        const r = await api('submitRequest', { ppcAction: act.value, department: dept.value, existingPpc: ex.value, wiNumbers: wi.value.trim(), assignment: act.value === 'NEW' ? meth.value : 'AUTO', manualPpc: man.value.trim(), requestor: val('#rqInit').trim(), email: val('#rqEmail').trim(), title: val('#rqChTitle').trim(), description: val('#rqDesc'), justification: val('#rqJust'), formPriority: ($('input[name=prio]:checked', rq) || {}).value || 'Low', requestDate: val('#rqDate'), dueDate: val('#rqDue'), manager: val('#rqMgr'), poc: val('#rqPoc'), changeType: val('#rqType'), effort: val('#rqEffort'), impactQuality: val('#rqImpQ'), impactCost: val('#rqImpC'), impactSchedule: val('#rqImpS'), impactResources: val('#rqImpR'), impactOther: val('#rqImpO') });
        out.textContent = 'Saved. PPC number: ' + r.id + '.'; await refresh();
      } catch (e) { out.textContent = e.code === 'WRITES_OFF' ? 'Saving from this page is switched off for now. Nothing was saved.' : e.message; nx.disabled = false; }
    }
  });
  /* test wizard */
  const tD = $('#tsDept'), tP = $('#tsPpc'), tDet = $('#tsDetails'), tS = $('#tsStart'), tE = $('#tsEnd'), tPoc = $('#tsPoc'), hint = $('#tsHint');
  const utc = s => Date.parse(s + 'T00:00:00Z'), span = (a, b) => Math.round((utc(b) - utc(a)) / 86400000), plus = (s, n) => { const d = new Date(utc(s)); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const cur = () => ppcsForWizard().find(p => p.n === tP.value);
  function tsFill() { const d = tD.value, l = ppcsForWizard().filter(p => p.dept === d); tP.replaceChildren(h('option', { value: '', text: d ? (l.length ? 'Select existing PPC' : 'No PPC in this department') : 'Choose a department first' }), ...l.map(p => h('option', { value: p.n, text: p.n + ' — ' + p.title }))); tsDetails(); }
  function tsDetails() { const p = cur(); tDet.textContent = p ? p.title + ' • ' + p.status + ' • WI: ' + p.wi : (tD.value ? 'Choose the PPC to test.' : 'Select a PPC to view its title, WI and current status.'); $('#tsPpcPrint').value = p ? p.n : ''; }
  function dates() {
    const s = tS.value, e = tE.value; if (s) tE.min = plus(s, D.config.minTestDays); const sp = (s && e) ? span(s, e) : null, bad = sp !== null && sp < D.config.minTestDays;
    hint.className = 'wz-hint' + (bad ? ' bad' : ''); hint.textContent = !s ? 'Minimum testing period: ' + D.config.minTestDays + ' calendar days. Select a start date to see the earliest end date.' : 'Test duration: ' + (sp === null ? 'choose an end date' : sp + ' day(s)') + '. Earliest permitted end date: ' + plus(s, D.config.minTestDays) + (bad ? '. Please extend the test.' : '.');
    $('#tsStart2').value = s; $('#tsEnd2').value = e; $('#tsPocPrint').value = tPoc.value;
  }
  tD.addEventListener('change', tsFill); tP.addEventListener('change', tsDetails); [tS, tE, tPoc].forEach(e => e.addEventListener('input', dates));
  $('#tsStart2').addEventListener('input', e => { tS.value = e.target.value; dates(); }); $('#tsEnd2').addEventListener('input', e => { tE.value = e.target.value; dates(); }); $('#tsPocPrint').addEventListener('input', e => { tPoc.value = e.target.value; });
  WZ.test = wizard(ts, {
    l1: 'Next: Fill Test Form →', submit: 'Submit PPC Test',
    reset(pre) { tD.value = ''; tP.replaceChildren(h('option', { value: '', text: 'Choose a department first' })); [tS, tE, tPoc].forEach(e => { e.value = ''; }); $('#tsForm').reset(); $('#tsDone').hidden = true; $('#tsNote').textContent = 'Reminders start ' + D.config.reminderLeadDays + ' days before the end date and repeat every ' + D.config.reminderCooldownDays + ' days. Email mode is ' + (D.config.emailMode === 'LIVE' ? 'live.' : 'dry run, so reminders are logged but not sent yet.'); if (pre) { const p = ppcsForWizard().find(x => x.n === pre); if (p) { tD.value = p.dept; tsFill(); tP.value = p.n; } } tsDetails(); dates(); },
    v1() { const e = []; if (!tD.value) e.push('Choose a department.'); if (!tP.value) e.push('Choose an existing PPC.'); if (!tS.value || !tE.value) e.push('Select test start and end dates.'); else if (span(tS.value, tE.value) < D.config.minTestDays) e.push('The test must span at least ' + D.config.minTestDays + ' calendar days.'); if (!tPoc.value.trim()) e.push('Enter the main test POC.'); return e; },
    review() {
      const p = cur(), f = [['PPC number', p ? p.n : 'Missing'], ['Department', tD.value || 'Missing'], ['Test reason', val('#tsReason') || 'Missing'], ['Main POC', tPoc.value || 'Missing'], ['Owner email', val('#tsEmail') || 'Missing'], ['Test dates', (tS.value || '—') + ' to ' + (tE.value || '—')], ['Define the test', val('#tsDefine') || 'Missing'], ['Expected results', val('#tsExpected') || 'Missing']];
      $('#tsRev').replaceChildren(...f.flatMap(r => [h('dt', { text: r[0] }), h('dd', { text: r[1] })]));
      const m = []; if (!val('#tsReason').trim()) m.push('Test reason'); if (!val('#tsDefine').trim()) m.push('Define the test'); if (!val('#tsExpected').trim()) m.push('Expected results'); if (!val('#tsEmail').trim()) m.push('Owner email');
      const mm = $('#tsMiss'); mm.textContent = m.length ? 'Complete before submitting: ' + m.join(', ') + '.' : 'Ready to submit. Review the values for accuracy.'; mm.className = m.length ? 'wz-miss' : 'wz-ok'; $('#tsDone').hidden = true;
    },
    async send() {
      const out = $('#tsDone'), nx = $('.wz-nav [data-next]', ts); out.hidden = false; out.textContent = 'Saving...'; nx.disabled = true;
      try {
        const r = await api('submitTest', { ppc: tP.value, owner: tPoc.value.trim(), ownerEmail: val('#tsEmail').trim(), reason: val('#tsReason'), start: tS.value, end: tE.value, defineTest: val('#tsDefine'), variants: val('#tsVar'), timeline: val('#tsTl'), resources: val('#tsRes'), expected: val('#tsExpected'), actual: val('#tsActual'), evidence: val('#tsEvid').trim(), result: val('#tsResult'), testNumber: ($('input[name=tn]:checked', ts) || {}).value || '1', signer1: val('#tsSig1'), signer2: val('#tsSig2'), signer3: val('#tsSig3') });
        out.textContent = 'Saved. Test ' + r.id + '.'; await refresh();
      } catch (e) { out.textContent = e.code === 'WRITES_OFF' ? 'Saving from this page is switched off for now. Nothing was saved.' : e.message; nx.disabled = false; }
    }
  });
}

/* ---------- start ---------- */
async function boot() {
  const bootNote = document.getElementById('boot'); if (bootNote) bootNote.remove();
  try {
    const res = await fetch(ASSET_BASE + ASSET_NAME + '.html', { credentials: 'omit' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    document.body.insertAdjacentHTML('beforeend', await res.text());     // our own file from our own repo
  } catch (e) { document.body.textContent = 'The page could not load its layout file (' + ASSET_NAME + '.html). Check that it is uploaded next to ' + ASSET_NAME + '.js on GitHub, with the same capital letters.'; return; }
  $$('a[data-lms]').forEach(a => { a.href = LMS_URL; a.target = '_top'; });
  $('#retry').addEventListener('click', loadApp);
  await loadApp();
}
boot();