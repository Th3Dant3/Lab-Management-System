"use strict";
/* PPC.js: one file, two jobs. (1) Opened from the LMS at github.io, PPC.html is your launcher: it checks the LMS session and opens
   the secure app. (2) Loaded by Google's tiny shell inside the secure app, it is the whole dashboard. Everything about how the
   page looks and behaves lives here in GitHub. The server trusts nothing on this page: every action is re-checked against PPC_USERS and
   your role. Every value that comes from the sheet is inserted with textContent, never as HTML. */
const APP_NAME = "PPC Document Control";          // the name shown everywhere on this page. Change this one line to rename it.
const LMS_URL = "https://th3dant3.github.io/Lab-Management-System/index.html";
const PPC_SECURE_URL = "https://script.google.com/a/macros/zennioptical.com/s/AKfycbybKUJ9OP708EbnplRJVrGPy9aenOEz07ZOWUEPJ8PrVCQprd6X2dFTffOxxpkXsqIiFg/exec";
const EMBED = true;                 // true: the dashboard opens INSIDE this page (the address bar stays on github.io). false: it opens in its own tab, as before.
const EMBED_TIMEOUT_MS = 15000;     // if nothing reports back by then, show how to open it in a new tab
const PARENT_ORIGIN = "https://th3dant3.github.io";   // the only website allowed to frame the secure app
const SELF_SRC = (document.currentScript && document.currentScript.src) || "";
const ASSET_BASE = SELF_SRC.replace(/[^\/]*$/, "");
const ASSET_NAME = (SELF_SRC.split("/").pop().split("?")[0] || "PPC.js").replace(/\.js$/i, "");   // PPC: the layout lives in PPC.html next to this file
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
  { status: 'In Progress', title: 'In progress', cls: 'draft' },
  { status: 'Ready for Signature', title: 'Ready for signature', cls: 'sign' },
  { status: 'Completed', title: 'Completed', cls: 'route' },
  { status: 'Published', title: 'Published', cls: 'done' }
];
const ADMIN_ONLY = ['Ready for Signature', 'Completed', 'Published'];   // only an administrator may set these (the server enforces it too)
/* The age clock. It runs while a PPC is In Progress, counted from the day it became In Progress.
   0 to 15 days: green. 16 to 29 days: yellow. 30 days or more: Overdue. */
const CLOCK_GREEN_MAX = 15, CLOCK_YELLOW_MAX = 29;
function isPending(r) { return /^PENDING-/.test((r && r.number) || ''); }
function clockStart(r) { return r.progressStart || r.submitted || ''; }
function clockDays(r) { return r.status === 'In Progress' ? daysSince(clockStart(r)) : null; }
function clockBand(d) { return d === null ? '' : d <= CLOCK_GREEN_MAX ? 'green' : d <= CLOCK_YELLOW_MAX ? 'yellow' : 'overdue'; }
const BAND_RANK = { green: 0, yellow: 1, overdue: 2, '': 3 };
/* Cards are ordered green, then yellow, then overdue, and within a colour the fewest days first.
   Rows with no running clock (not started, completed, published) come last, newest first. */
function byBand(a, b) {
  const da = clockDays(a), db = clockDays(b), r = BAND_RANK[clockBand(da)] - BAND_RANK[clockBand(db)];
  if (r) return r;
  if (da !== null && db !== null) return da - db;
  return (b.submitted || '').localeCompare(a.submitted || '');
}
const collapsed = {};                                  // which board columns are closed, remembered while the page is open
const bandShut = {};                                   // which colour groups inside In progress are closed
const BANDS = [['green', '0 to 15 days'], ['yellow', '16 to 29 days'], ['overdue', '30 days or more, Overdue']];
const deptOpen = {};                                   // which department groups are open on the Published page
function clockBar(d) { const s = h('span'); s.style.width = Math.min(100, Math.max(2, Math.round(d / 30 * 100))) + '%'; return h('div', { class: 'age ' + clockBand(d), role: 'img', 'aria-label': d + ' days in progress, ' + clockBand(d) }, s); }
function ageBar(days, maxD) { const s = h('span'); s.style.width = (maxD ? Math.max(2, Math.round(days / maxD * 100)) : 0) + '%'; return h('div', { class: 'age', role: 'img', 'aria-label': days + ' days open' }, s); }
function btn(label, o) { o = o || {}; return h('button', { type: 'button', class: 'btn' + (o.quiet ? ' quiet' : '') + (o.primary ? ' primary' : ''), disabled: o.disabled, title: o.title, onclick: o.onclick, 'data-open': o.open, 'data-ppc': o.ppc }, label); }
function table(heads, rows, empty) {
  if (!rows.length) return h('p', { class: 'empty', text: empty });
  return h('div', { class: 'scroll' }, h('table', null, [h('thead', null, h('tr', null, heads.map(x => h('th', { text: x })))), h('tbody', null, rows)]));
}
function td(content, cls) { return h('td', { class: cls }, content); }
function openRows() { return D.register.filter(r => r.status !== 'Completed' && r.status !== 'Published'); }   // work still going: not started, in progress, ready for signature
function pendingRows() { return openRows().filter(isPending).concat(D.needsNumber); }
function maxAge() { const a = openRows().map(clockDays).filter(x => x !== null); return a.length ? Math.max.apply(null, a) : 0; }
const EDIT_STAGES = ['Request Submitted', 'Proposal Review', 'Testing', 'WI Draft', 'Approval / Routing', 'Published'];
/* Edit one numbered PPC. Only status, stage, due date, blocker and the two notes can change here, and the server enforces that.
   "In Routing" and "Published" are offered only if the PPC is already there, because they have their own rules and buttons. */
function openEdit(r) {
  if (!D.can.edit) return;
  const invalid = !STATUS_COLS.some(c => c.status === r.status);
  const stSel = h('select', { id: 'edStatus' }, [invalid ? h('option', { value: '', text: 'Choose a status' }) : null].concat(STATUS_COLS.map(c => h('option', { value: c.status, text: c.title, disabled: !D.can.allStatuses && ADMIN_ONLY.indexOf(c.status) >= 0 && r.status !== c.status }))));
  stSel.value = invalid ? '' : r.status;
  const stages = EDIT_STAGES.slice(); if (stages.indexOf(r.stage || '') < 0) stages.unshift(r.stage || '');
  const sgSel = h('select', { id: 'edStage' }, stages.map(s => h('option', { value: s, text: s || 'Not set', disabled: s === 'Published' && r.stage !== 'Published' })));
  sgSel.value = r.stage || '';
  const due = h('input', { id: 'edDue', type: 'date', value: r.due || '' });
  const blocker = h('input', { id: 'edBlocker', type: 'text', maxlength: '600', value: r.blocker || '' });
  const pub = h('textarea', { id: 'edPublic', maxlength: '1500' }); pub.value = r.publicNotes || '';
  const internal = h('textarea', { id: 'edInternal', maxlength: '1500' }); internal.value = r.internalNotes || '';
  const msg = h('p', { class: 'editmsg', role: 'status' });
  const setMsg = (t, err) => { msg.textContent = t; msg.classList.toggle('err', !!err); };
  const syncStage = () => { Array.from(sgSel.options).forEach(o => { if (o.value === 'Published') o.disabled = stSel.value !== 'Published' && r.stage !== 'Published'; }); };
  stSel.addEventListener('change', () => {
    const s = stSel.value; let d = ({ 'Not Started': 'Request Submitted', 'Ready for Signature': 'Approval / Routing', 'Completed': 'Approval / Routing', 'Published': 'Published' })[s];
    if (s === 'In Progress') d = ['Proposal Review', 'Testing', 'WI Draft'].indexOf(sgSel.value) >= 0 ? sgSel.value : 'WI Draft';
    syncStage(); if (d) sgSel.value = d;
  });
  const onKey = e => { if (e.key === 'Escape' && !document.querySelector('.lightbox')) close(); };   // Escape closes the image viewer first, never this panel underneath it
  const close = () => { ov.remove(); document.removeEventListener('keydown', onKey); };
  let save;
  const commit = async () => {
    const ch = {};
    if (invalid && !stSel.value) { setMsg('Choose a status for this PPC.', true); return; }
    if (stSel.value && stSel.value !== r.status) ch.status = stSel.value;
    if (sgSel.value !== (r.stage || '')) ch.stage = sgSel.value;
    if (due.value !== (r.due || '')) ch.due = due.value;
    if (blocker.value !== (r.blocker || '')) ch.blocker = blocker.value;
    if (pub.value !== (r.publicNotes || '')) ch.publicNotes = pub.value;
    if (internal.value !== (r.internalNotes || '')) ch.internalNotes = internal.value;
    if (!Object.keys(ch).length) { setMsg('Nothing was changed.', true); return; }
    save.disabled = true; setMsg('Saving...');
    try {
      await api('updatePpc', { number: r.number, changes: ch, expect: { status: r.status, stage: r.stage } });
      FLASH = 'Saved ' + r.number + (ch.status ? '. Status is now ' + ch.status + '.' : '.'); close(); await refresh();
    } catch (e) { setMsg(e.code === 'WRITES_OFF' ? 'Saving from this page is switched off for now. Nothing was saved.' : e.message, true); save.disabled = false; }
  };
  save = btn('Save changes', { primary: true, onclick: commit });
  /* The five steps, in order, with one button that moves this PPC on by exactly one. A PPC starts at Not Started
     when it is submitted and ends at Published. Saving a step saves anything else typed in this panel at the same time. */
  const here = STATUS_COLS.findIndex(c => c.status === r.status);
  const stepTo = async (to) => {
    stSel.value = to; stSel.dispatchEvent(new Event('change'));
    setMsg('Moving to ' + to + '...');
    await commit();
  };
  const stepRow = h('ol', { class: 'steps' }, STATUS_COLS.map((c, i) => h('li', {
    class: (here < 0 ? '' : i < here ? 'done' : i === here ? 'now' : ''), 'aria-current': i === here ? 'step' : null
  }, [h('span', { class: 'sn', text: String(i + 1) }), h('span', { text: c.title })])));
  const next = here >= 0 && here < STATUS_COLS.length - 1 ? STATUS_COLS[here + 1] : null;
  const back = here > 0 ? STATUS_COLS[here - 1] : null;
  const blocked = s => !D.can.allStatuses && ADMIN_ONLY.indexOf(s.status) >= 0;
  const stepBtns = h('div', { class: 'stepbtns' }, [
    next ? btn('Move to ' + next.title, { primary: true, disabled: blocked(next), title: blocked(next) ? 'Only an administrator can set ' + next.status + '.' : '', onclick: () => stepTo(next.status) }) : h('span', { class: 'mut', text: 'This PPC is Published. It is the last step.' }),
    back ? btn('Back to ' + back.title, { quiet: true, disabled: blocked(back), title: blocked(back) ? 'Only an administrator can set ' + back.status + '.' : '', onclick: () => stepTo(back.status) }) : null,
    here < 0 ? h('span', { class: 'mut', text: 'This status is not one of the five, so choose one below first.' }) : null]);
  const stepper = h('section', { class: 'stepper' }, [h('h3', { class: 'nh', text: 'Where it is' }), stepRow, stepBtns,
    next && !blocked(next) ? h('p', { class: 'mut small', text: 'Or pick any status below to jump straight to it.' }) : null]);
  const field = (label, el, note) => h('label', null, [label, el, note ? h('small', { text: note }) : null]);
  const ov = h('div', { class: 'wz-overlay editov', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'Edit ' + r.number }, h('div', { class: 'editbox' }, [
    h('h2', { text: r.number }), h('p', { class: 'sub', text: (r.title || 'Title not recorded') + (r.wi ? ' (' + r.wi + ')' : '') }),
    stepper,
    field('Status', stSel, D.can.allStatuses ? 'In progress starts the age clock. Completed and Published stamp today\'s date if none is set.' : 'Ready for signature, Completed and Published can be set by an administrator only.'), field('Stage', sgSel),
    field('Due date', due), field('Blocker or next action', blocker),
    field('Public update (requesters see this, and it is emailed when the PPC is routed or published)', pub), field('Internal notes (staff only)', internal),
    msg, h('div', { class: 'actions' }, [save, btn('Cancel', { quiet: true, onclick: close })]),
    D.can.attach ? attachPanel(r.number, { canAdd: true, canDelete: true }) : h('p', { class: 'mut small', text: 'Images are not set up yet. An administrator must run setupPpcAttachmentsV7.' })]));
  ov.addEventListener('click', e => { if (e.target === ov) close(); });
  document.addEventListener('keydown', onKey); document.body.append(ov); stSel.focus();
}
const IMG_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
const IMG_MAX_BYTES = 4 * 1024 * 1024;
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => { const s = String(fr.result || ''); const i = s.indexOf(','); resolve(i >= 0 ? s.slice(i + 1) : ''); };
    fr.onerror = () => reject(new Error('That file could not be read.'));
    fr.readAsDataURL(file);
  });
}
/* Images chosen while filling in a form, before the PPC or test exists. Nothing is sent until the record is created,
   so a half-finished form never leaves files behind. After it is created the staged images are uploaded one by one. */
function stageImages(label) {
  const picked = [], list = h('div', { class: 'staged' }), msg = h('p', { class: 'attachmsg', role: 'status' });
  const setMsg = (t, err) => { msg.textContent = t || ''; msg.classList.toggle('err', !!err); };
  function draw() {
    if (!picked.length) { setKids(list, [h('p', { class: 'mut', text: 'No images chosen yet.' })]); return; }
    setKids(list, picked.map((f, i) => h('div', { class: 'stagerow' }, [
      h('span', { class: 'fn', text: f.name || 'screenshot.png' }),
      h('span', { class: 'mut', text: Math.max(1, Math.round(f.size / 1024)) + ' KB' }),
      btn('Remove', { quiet: true, onclick: () => { picked.splice(i, 1); draw(); setMsg(''); } })])));
  }
  function take(file) {
    if (!file) { setMsg('Choose an image first.', true); return false; }
    if (IMG_TYPES.indexOf(file.type) < 0) { setMsg('Only PNG, JPEG, GIF and WebP images can be attached.', true); return false; }
    if (file.size > IMG_MAX_BYTES) { setMsg('That image is ' + (Math.round(file.size / 1048576 * 10) / 10) + ' MB. The limit is 4 MB.', true); return false; }
    if (picked.length >= 20) { setMsg('You can attach 20 images. Remove one first.', true); return false; }
    picked.push(file); draw(); setMsg('Added ' + (file.name || 'screenshot.png') + '. It is uploaded when you submit.');
    return true;
  }
  const input = h('input', { type: 'file', accept: IMG_TYPES.join(',') });
  const add = btn('Add image', { onclick: () => { if (take(input.files && input.files[0])) input.value = ''; } });
  const paste = h('div', { class: 'pastebox', tabindex: '0', title: 'Click here, then press Ctrl+V to paste a screenshot' }, 'Or click here and press Ctrl+V to paste a screenshot');
  paste.addEventListener('paste', e => {
    const files = (e.clipboardData && e.clipboardData.files) || [];
    if (!files.length) { setMsg('There is no image on the clipboard. Take a screenshot first.', true); return; }
    e.preventDefault(); take(files[0]);
  });
  const wrap = h('section', { class: 'attach stage' }, [h('h3', { class: 'nh', text: label || 'Images (optional)' }), list,
    h('div', { class: 'attachadd' }, [input, add]), paste, msg,
    h('p', { class: 'mut small', text: 'PNG, JPEG, GIF or WebP, up to 4 MB each. They are uploaded only when you submit, and only people allowed to see this request can view them.' })]);
  draw();
  /* Upload everything that was staged. Returns a short message about anything that failed, so a lost image is never silent. */
  wrap.upload = async (number, test) => {
    const failed = [];
    for (let i = 0; i < picked.length; i++) {
      const f = picked[i];
      try {
        const data = await fileToBase64(f);
        await api('addAttachment', { number: number, test: test || '', name: f.name || 'screenshot.png', data: data });
      } catch (e) { failed.push((f.name || 'an image') + ' (' + e.message + ')'); }
    }
    const done = picked.length - failed.length;
    picked.length = 0; draw();
    return { done: done, failed: failed };
  };
  wrap.count = () => picked.length;
  wrap.clear = () => { picked.length = 0; draw(); setMsg(''); };
  return wrap;
}
/* Images for one PPC. The server re-checks who you are on every list, every view and every upload, so this panel
   only decides what to show. Pictures are fetched one at a time and never cached on disk. */
function attachPanel(ppc, opts) {
  opts = opts || {};
  const testId = opts.test || '';
  const grid = h('div', { class: 'thumbs' }), msg = h('p', { class: 'attachmsg', role: 'status' });
  const setMsg = (t, err) => { msg.textContent = t || ''; msg.classList.toggle('err', !!err); };
  const wrap = h('section', { class: 'attach' }, [h('h3', { class: 'nh', text: opts.title || 'Images' }), grid, msg]);
  let busy = false;
  async function open(it) {
    const box = h('div', { class: 'lightbox', role: 'dialog', 'aria-modal': 'true', 'aria-label': it.name });
    const close = () => { box.remove(); document.removeEventListener('keydown', onKey); };
    const onKey = e => { if (e.key === 'Escape') close(); };
    const body = h('div', { class: 'lbbody' }, h('p', { class: 'mut', text: 'Loading ' + it.name + '...' }));
    box.append(h('div', { class: 'lbbar' }, [h('b', { text: it.name }), h('span', { class: 'grow' }), btn('Close', { quiet: true, onclick: close })]), body);
    box.addEventListener('click', e => { if (e.target === box) close(); });
    document.addEventListener('keydown', onKey); document.body.append(box);
    try {
      const full = await api('getAttachment', { id: it.id });
      setKids(body, [h('img', { src: full.dataUrl, alt: it.caption || it.name }), it.caption ? h('p', { class: 'cap', text: it.caption }) : null]);
    } catch (e) { setKids(body, [h('p', { class: 'err', text: e.message })]); }
  }
  async function load() {
    setKids(grid, [h('p', { class: 'mut', text: 'Loading images...' })]);
    let items = [];
    try { items = (await api('listAttachments', { number: ppc, test: testId })).items || []; }
    catch (e) { setKids(grid, [h('p', { class: 'err', text: e.message })]); return; }
    if (!items.length) setKids(grid, [h('p', { class: 'mut', text: 'No images yet.' })]);
    else setKids(grid, items.map(it => {
      const fig = h('figure', { class: 'thumb' });
      const img = h('img', { alt: it.caption || it.name });
      const btns = [btn('View', { quiet: true, onclick: () => open(it) })];
      if (opts.canDelete) btns.push(btn('Delete', { quiet: true, onclick: async () => {
        if (!confirm('Delete ' + it.name + '? It is removed from this PPC and moved to the Drive bin.')) return;
        try { await api('deleteAttachment', { id: it.id }); setMsg('Deleted ' + it.name + '.'); load(); }
        catch (e) { setMsg(e.message, true); } } }));
      fig.append(img, h('figcaption', null, [h('span', { class: 'fn', text: it.caption || it.name }), h('span', { class: 'mut', text: Math.max(1, Math.round(it.bytes / 1024)) + ' KB, added by ' + it.by.split('@')[0] + (it.at ? ' on ' + fmtDate(it.at) : '') }), h('span', { class: 'tb' }, btns)]));
      let triedFull = false;
      const giveUp = () => { if (img.isConnected) img.replaceWith(h('div', { class: 'nothumb', text: 'Preview unavailable' })); };
      img.addEventListener('error', () => {            // Drive's small preview can be missing or unreadable; fall back to the real image once
        if (triedFull) { giveUp(); return; }
        triedFull = true;
        api('getAttachment', { id: it.id }).then(f => { img.src = f.dataUrl; }).catch(giveUp);
      });
      api('getAttachment', { id: it.id, size: 'thumb' }).then(t => { img.src = t.dataUrl; }).catch(giveUp);
      return fig;
    }));
  }
  async function send(file, caption) {
    if (busy) return;
    if (!file) { setMsg('Choose an image first.', true); return; }
    if (IMG_TYPES.indexOf(file.type) < 0) { setMsg('Only PNG, JPEG, GIF and WebP images can be attached.', true); return; }
    if (file.size > IMG_MAX_BYTES) { setMsg('That image is ' + (Math.round(file.size / 1048576 * 10) / 10) + ' MB. The limit is 4 MB.', true); return; }
    busy = true; setMsg('Uploading ' + (file.name || 'image') + '...');
    try {
      const data = await fileToBase64(file);
      await api('addAttachment', { number: ppc, test: testId, name: file.name || 'screenshot.png', data: data, caption: caption || '' });
      setMsg('Added ' + (file.name || 'image') + '.'); await load();
    } catch (e) { setMsg(e.code === 'WRITES_OFF' ? 'Saving from this page is switched off for now. Nothing was uploaded.' : e.message, true); }
    busy = false;
  }
  if (opts.canAdd) {
    const input = h('input', { type: 'file', id: 'imgFile-' + ppc + (testId ? '-' + testId : ''), accept: IMG_TYPES.join(',') });
    const cap = h('input', { type: 'text', maxlength: '300', placeholder: 'What does this image show? (optional)' });
    const add = btn('Add image', { onclick: async () => { await send(input.files && input.files[0], cap.value); input.value = ''; cap.value = ''; } });
    const paste = h('div', { class: 'pastebox', tabindex: '0', title: 'Click here, then press Ctrl+V to paste a screenshot' }, 'Or click here and press Ctrl+V to paste a screenshot');
    paste.addEventListener('paste', async e => {
      const items = (e.clipboardData && e.clipboardData.files) || [];
      if (!items.length) { setMsg('There is no image on the clipboard. Take a screenshot first.', true); return; }
      e.preventDefault(); await send(items[0], cap.value); cap.value = '';
    });
    wrap.append(h('div', { class: 'attachadd' }, [input, cap, add]), paste);
  }
  wrap.append(h('p', { class: 'mut small', text: 'PNG, JPEG, GIF or WebP, up to 4 MB each, 20 per PPC. Images are kept in the document control Drive folder and are only shown to people allowed to see this PPC.' }));
  load();
  return wrap;
}
/* A standalone Images window, used for a test and anywhere else that has no edit panel. */
function openImages(ppc, test, title) {
  const onKey = e => { if (e.key === 'Escape' && !document.querySelector('.lightbox')) close(); };
  const close = () => { ov.remove(); document.removeEventListener('keydown', onKey); };
  const ov = h('div', { class: 'wz-overlay editov', role: 'dialog', 'aria-modal': 'true', 'aria-label': title || 'Images' },
    h('div', { class: 'editbox' }, [h('h2', { text: title || 'Images' }), h('p', { class: 'sub', text: ppc }),
      attachPanel(ppc, { test: test, canAdd: true, canDelete: true, title: 'Images' }),
      h('div', { class: 'actions' }, btn('Close', { quiet: true, onclick: close }))]));
  ov.addEventListener('click', e => { if (e.target === ov) close(); });
  document.addEventListener('keydown', onKey); document.body.append(ov);
}
const NOT_BUILT = 'Not switched on yet. This button needs a backend step that is built after sign-in is confirmed.';

/* ---------- staff views ---------- */
const VIEWS = [['overview', 'Dashboard'], ['guide', 'Quick start guide'], ['published', 'Published documents'], ['missing', 'Missing PPC, need update'], ['tests', 'Tests and reminders'], ['routing', 'Routing and emails'], ['updates', 'Update requests'], ['years', 'Year records']];
let currentView = 'overview', currentYear = null, mailSel = null, mailEv = 'routed';
function token(r, cls) {
  const kids = [h('p', { class: 'num', text: isPending(r) ? 'PPC number pending' : r.number }), h('p', { class: 'ttl', text: r.title || 'Title not recorded' }), h('p', { class: 'who', text: (r.requestor || 'Requester not recorded') + (r.wi ? ', ' + r.wi : '') })];
  if (r.stage) kids.push(h('p', { class: 'sub', text: r.stage }));
  const d = clockDays(r);
  if (d !== null) {
    const band = clockBand(d);
    kids.push(clockBar(d));
    kids.push(h('p', { class: 'days ' + band, title: r.progressStart ? '' : 'Counted from the submitted date, because the day it became In Progress was not recorded.' }, [d + (d === 1 ? ' day' : ' days') + ' in progress', band === 'overdue' ? h('span', { class: 'flagover', text: 'Overdue' }) : null]));
  } else if (r.status === 'Completed' && r.routing) kids.push(h('p', { class: 'days', text: 'Completed ' + fmtDate(r.routing) }));
  else if (r.status === 'Published' && r.published) kids.push(h('p', { class: 'days', text: 'Published ' + fmtDate(r.published) }));
  else if (r.submitted) kids.push(h('p', { class: 'days', text: 'Submitted ' + fmtDate(r.submitted) }));
  const pics = (D.imageCounts || {})[r.number] || 0;
  if (pics) kids.push(h('p', { class: 'haspic', text: pics === 1 ? '1 image' : pics + ' images' }));
  if (D.can.edit) kids.push(h('button', { type: 'button', class: 'tokedit', onclick: () => openEdit(r) }, 'Edit'));
  return h('li', { class: 'tok ' + cls }, kids);
}
function vOverview() {
  const open = openRows(), pend = pendingRows(), numbered = open.filter(r => !isPending(r)), maxD = maxAge(), openTotal = open.length + D.needsNumber.length;
  const pubRows = D.register.filter(r => r.status === 'Published'), pubTotal = D.published.length + pubRows.length;
  const overdue = open.filter(r => clockBand(clockDays(r)) === 'overdue');
  const items = [];
  if (overdue.length) items.push([h('span', null, [h('b', { text: pl(overdue.length, 'PPC') + ' ' + (overdue.length === 1 ? 'is' : 'are') + ' overdue (30 days or more in progress): ' }), overdue.map(r => (isPending(r) ? (r.wi || r.title || 'pending request') : r.number) + ' (' + clockDays(r) + ' days)').join(', ') + '.']), null, null, 'red']);
  if (pend.length) items.push([h('span', null, [h('b', { text: pl(pend.length, 'open request') + ' ' + (pend.length === 1 ? 'has' : 'have') + ' no PPC number.' }), ' The longest has waited ' + Math.max.apply(null, pend.map(r => daysSince(clockStart(r)) || 0)) + ' days.']), 'Open Missing PPC, need update', 'missing', '']);
  const noTest = numbered.filter(r => r.stage === 'Testing' && !D.tests.some(t => t.ppc === r.number));
  if (noTest.length) items.push([h('span', null, [h('b', { text: pl(noTest.length, 'PPC') + ' ' + (noTest.length === 1 ? 'is' : 'are') + ' in the Testing stage with no test recorded: ' }), noTest.map(r => r.number).join(' and ') + '. Without a test end date, no reminder can be sent.']), 'Open Tests and reminders', 'tests', 'amber']);
  const sign = open.filter(r => r.status === 'Ready for Signature');
  if (sign.length) items.push([h('span', null, [h('b', { text: pl(sign.length, 'PPC') + ' ' + (sign.length === 1 ? 'is' : 'are') + ' ready for signature: ' }), sign.map(r => r.number).join(', ') + '.']), null, null, 'blue']);
  const odd = D.register.filter(r => !STATUS_COLS.some(c => c.status === r.status));
  const att = h('section', null, [h('h2', { text: 'Needs attention' }), items.length ? h('ul', { class: 'attn' }, items.map(i => h('li', null, [h('div', null, [h('span', { class: 'dot ' + i[3] }), i[0]]), i[1] ? h('a', { class: 'btn', href: '#' + i[2], onclick: e => { e.preventDefault(); show(i[2]); } }, i[1]) : null]))) : h('p', { class: 'empty', text: 'Nothing needs attention right now.' })]);
  const colFor = c => {
    const rows = (c.status === 'Published' ? pubRows : D.register.filter(r => r.status === c.status)).slice().sort(byBand);
    const count = c.status === 'Published' ? pubTotal : rows.length;
    const shut = !!collapsed[c.status];
    const head = h('button', { type: 'button', class: 'colhead', 'aria-expanded': String(!shut), title: (shut ? 'Open' : 'Close') + ' the ' + c.title + ' column',
      onclick: () => { collapsed[c.status] = !shut; rebuild('overview'); } },
      [h('span', { class: 'caret', 'aria-hidden': 'true', text: shut ? '+' : '\u2212' }), h('h3', { text: c.title }), h('span', { class: 'c', text: String(count) })]);
    if (shut) return h('div', { class: 'col ' + c.cls + ' collapsed' }, [head, h('p', { class: 'none', text: count === 1 ? '1 hidden' : count + ' hidden' })]);
    if (c.status === 'In Progress') {
      const kids = [head];
      BANDS.forEach(p => {
        const key = p[0], list = rows.filter(r => clockBand(clockDays(r)) === key), shutB = bandShut[key] === true;
        const bh = h('button', { type: 'button', class: 'bandhead ' + key, 'aria-expanded': String(!shutB), title: (shutB ? 'Open' : 'Close') + ' ' + p[1],
          onclick: () => { bandShut[key] = !shutB; rebuild('overview'); } },
          [h('span', { class: 'caret', 'aria-hidden': 'true', text: shutB ? '+' : '\u2212' }), h('span', { class: 'sw ' + key, 'aria-hidden': 'true' }), h('span', { class: 'bl', text: p[1] }), h('span', { class: 'n', text: String(list.length) })]);
        kids.push(h('div', { class: 'band ' + key }, [bh, shutB ? null : (list.length ? h('ul', null, list.map(r => token(r, c.cls))) : h('p', { class: 'none small', text: 'None in this range.' }))]));
      });
      const noClock = rows.filter(r => clockDays(r) === null);
      if (noClock.length) kids.push(h('div', { class: 'band' }, [h('p', { class: 'bandplain', text: pl(noClock.length, 'with no start date recorded') }), h('ul', null, noClock.map(r => token(r, c.cls)))]));
      return h('div', { class: 'col ' + c.cls }, kids);
    }
    if (c.status === 'Published') return h('div', { class: 'col done' }, [head, rows.length ? h('ul', null, rows.slice(0, 8).map(r => token(r, 'done'))) : null,
      h('div', { class: 'arch' }, [h('p', { text: D.published.length + ' older completed requests are not on this board.' }), h('a', { class: 'btn', href: '#published', onclick: e => { e.preventDefault(); show('published'); } }, 'See all published documents')])]);
    return h('div', { class: 'col ' + c.cls }, [head, rows.length ? h('ul', null, rows.map(r => token(r, c.cls))) : h('p', { class: 'none', text: ({ 'Not Started': 'No PPC is waiting to start.', 'Ready for Signature': 'No PPC is waiting for management signatures.', 'Completed': 'No PPC is marked Completed.' })[c.status] || 'Nothing here.' })]);
  };
  const cols = STATUS_COLS.map(colFor);
  const flow = h('div', { class: 'flow' }, cols);
  flow.style.gridTemplateColumns = STATUS_COLS.map(c => collapsed[c.status] ? 'minmax(118px,0.34fr)' : 'minmax(0,1fr)').join(' ');   // a closed column keeps enough room for its heading to read properly
  return h('div', { class: 'view', id: 'v-overview' }, [h('h1', { text: APP_NAME }),
    h('p', { class: 'lead' }, [h('strong', { text: pl(openTotal, 'request') + (openTotal === 1 ? ' is' : ' are') + ' open. ' }), numbered.length + ' ' + (numbered.length === 1 ? 'has' : 'have') + ' a PPC number and ' + pend.length + ' ' + (pend.length === 1 ? 'is' : 'are') + ' still waiting for one.' + (overdue.length ? ' ' + overdue.length + ' ' + (overdue.length === 1 ? 'is' : 'are') + ' overdue.' : '') + (maxD ? ' The longest in progress has been going ' + maxD + ' days.' : '')]),
    odd.length ? h('div', { class: 'banner' }, [pl(odd.length, 'register row') + ' ' + (odd.length === 1 ? 'has' : 'have') + ' a status that is not one of the five (' + odd.map(r => (isPending(r) ? 'pending ' + (r.wi || '') : r.number) + ' says "' + r.status + '"').join('; ') + '), so ' + (odd.length === 1 ? 'it is' : 'they are') + ' not shown on the board. ', ...(D.can.edit ? odd.map(r => btn('Fix ' + (isPending(r) ? (r.wi || 'pending request') : r.number), { onclick: () => openEdit(r) })) : ['Ask someone who can edit to fix it.'])]) : null,
    h('section', { class: 'boardfirst' }, [h('h2', { text: 'Where the PPCs are' }), h('p', { class: 'hint', text: 'The age clock runs while a PPC is In Progress, counted from the day it became In Progress. Click a column heading, or a colour heading, to close it.' }), flow]),
    att]);
}
function vMissing() {
  const need = pendingRows();
  const rows = need.map(r => { const d = r.status === 'In Progress' ? clockDays(r) : daysSince(r.submitted); const inReg = D.register.indexOf(r) >= 0; return h('tr', null, [td([h('b', { text: r.wi || 'No WI recorded' }), h('br'), h('span', { class: 'mut', text: r.title || 'Title not recorded' })]), td(r.requestor || ''), td(fmtDate(r.submitted)), td(d === null ? '' : [r.status === 'In Progress' ? clockBar(d) : null, h('span', { class: 'n', text: String(d) })], 'bar'), td(r.status + (r.stage ? ', ' + r.stage : '')), td([btn('Issue PPC number', { disabled: true, title: NOT_BUILT }), ' ', btn('Close request', { quiet: true, disabled: true, title: NOT_BUILT }), D.can.edit && inReg ? ' ' : null, D.can.edit && inReg ? btn('Edit', { quiet: true, onclick: () => openEdit(r) }) : null], 'act')]); });
  return h('div', { class: 'view', id: 'v-missing', hidden: true }, [h('h1', { text: 'Missing PPC, need update' }),
    h('p', { class: 'lead' }, [h('strong', { text: need.length ? pl(need.length, 'open request') + ' ' + (need.length === 1 ? 'has' : 'have') + ' no PPC number.' : 'Every open request has a PPC number.' }), need.length ? ' They are on the register with the status the source sheet gives them. New requests with no WI number will land here too.' : '']),
    h('section', null, [h('h2', { text: 'Decide each one' }), h('p', { class: 'hint', text: 'You can edit these now. Issue PPC number and Close request are switched on in a later step. No number is ever guessed.' }), table(['Request', 'Requested by', 'Submitted', 'Days in progress', 'Status, stage', 'Action'], rows, 'Nothing is waiting for a PPC number.')])]);
}
function vTests() {
  const open = openRows().filter(r => !isPending(r)), cfg = D.config, noTest = open.filter(r => r.stage === 'Testing' && !D.tests.some(t => t.ppc === r.number));
  const rows1 = noTest.map(r => h('tr', null, [td([h('b', { text: r.number }), h('br'), h('span', { class: 'mut', text: r.title })]), td(r.requestor || ''), td('No test recorded'), td(btn('Record a test', { open: 'test', ppc: r.number }), 'act')]));
  const rows2 = D.tests.map(t => { const d = daysSince(t.end); return h('tr', null, [td([h('b', { text: t.ppc }), h('br'), h('span', { class: 'mut', text: t.id })]), td([t.owner, h('br'), h('span', { class: 'mut', text: t.ownerEmail })]), td(fmtDate(t.end) + (d === null ? '' : d > 0 ? ', ended ' + pl(d, 'day') + ' ago' : d === 0 ? ', ends today' : ', ends in ' + pl(-d, 'day'))), td(t.result), td([D.can.attach ? btn('Images' + (((D.imageCounts || {})[t.id]) ? ' (' + D.imageCounts[t.id] + ')' : ''), { quiet: true, onclick: () => openImages(t.ppc, t.id, 'Images for test ' + t.id) }) : null, ' ', btn('Send reminder now', { quiet: true, disabled: true, title: NOT_BUILT })], 'act')]); });
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
  const done = D.register.filter(r => r.status === 'Completed'); if ((!mailSel || !done.some(r => r.number === mailSel)) && done.length) mailSel = done[0].number;
  const rows = done.map(r => { const d = daysSince(r.routing); return h('tr', null, [td([h('b', { text: r.number }), h('br'), h('span', { class: 'mut', text: r.title })]), td([r.requestor || '', h('br'), h('span', { class: 'mut', text: r.requestorEmail || 'no email on file' })]), td(r.routing ? [fmtDate(r.routing), h('br'), h('span', { class: 'mut', text: d + ' days ago' })] : 'No completed date'), td([btn('Preview email', { quiet: true, onclick: () => { mailSel = r.number; renderMail(); } }), ' ', D.can.edit ? btn('Edit', { quiet: true, onclick: () => openEdit(r) }) : null, ' ', btn('Send email', { disabled: true, title: NOT_BUILT })], 'act')]); });
  const seg = (ev, label) => h('button', { type: 'button', class: 'seg' + (mailEv === ev ? ' on' : ''), 'aria-pressed': String(mailEv === ev), onclick: () => { mailEv = ev; renderMail(); } }, label);
  return h('div', { class: 'view', id: 'v-routing', hidden: true }, [h('h1', { text: 'Routing and emails' }),
    h('p', { class: 'lead' }, [h('strong', { text: pl(done.length, 'PPC') + (done.length === 1 ? ' is' : ' are') + ' Completed.' }), ' Setting a status from the Edit panel does not send email. The email buttons are switched on in a later step.']),
    h('section', null, [h('h2', { text: 'Completed PPCs' }), table(['PPC', 'Requested by', 'Completed on', 'Action'], rows, 'No PPC is marked Completed.')]),
    h('section', null, [h('h2', { text: 'Email the requester will receive' }), h('div', { class: 'tabs2', role: 'group', 'aria-label': 'Email event' }, [seg('routed', 'Completed'), seg('published', 'Published')]), mailBox()])]);
}
function renderMail() { const v = $('#v-routing'); if (!v) return; const old = v.children[3]; const fresh = vRouting(); v.replaceWith(fresh); fresh.hidden = currentView !== 'routing'; }
let pubFilter = '';
function deptLabel(code) { const c = String(code || '').trim(); if (!c) return 'No department recorded'; const hit = (D.departments || []).filter(d => d.code === c)[0]; return hit ? hit.label + ' (' + c + ')' : c; }
function vPublished() {
  const all = D.published.concat(D.register.filter(r => r.status === 'Published'));
  const q = pubFilter.trim().toLowerCase();
  const match = r => !q || [r.number, r.title, r.wi, r.requestor, r.specialist].some(v => String(v || '').toLowerCase().indexOf(q) >= 0);
  const rows = all.filter(match);
  const groups = {}; rows.forEach(r => { const k = deptLabel(r.department); (groups[k] = groups[k] || []).push(r); });
  const names = Object.keys(groups).sort((a, b) => groups[b].length - groups[a].length || a.localeCompare(b));
  const dated = rows.filter(r => r.published).length;
  const box = h('input', { type: 'search', id: 'pubSearch', placeholder: 'Search PPC number, title, WI, requester or specialist', value: pubFilter });
  box.addEventListener('input', () => { pubFilter = box.value; rebuild('published'); const el = $('#pubSearch'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); } });
  const sections = names.map(name => {
    const list = groups[name].slice().sort((a, b) => (b.published || '').localeCompare(a.published || '') || (b.submitted || '').localeCompare(a.submitted || ''));
    const open = q ? true : !!deptOpen[name];
    const head = h('button', { type: 'button', class: 'dephead', 'aria-expanded': String(open), onclick: () => { deptOpen[name] = !open; rebuild('published'); } },
      [h('span', { class: 'caret', 'aria-hidden': 'true', text: open ? '\u2212' : '+' }), h('b', { text: name }), h('span', { class: 'n', text: String(list.length) })]);
    if (!open) return h('section', { class: 'depgroup' }, head);
    const trs = list.map(r => h('tr', null, [td(h('b', { text: r.number || 'No PPC number' })), td(r.wi || ''), td(r.title || h('span', { class: 'mut', text: 'Title not recorded' })), td(r.requestor || ''), td(r.specialist || ''), td(r.published ? fmtDate(r.published) : h('span', { class: 'mut', text: 'No date recorded' }))]));
    return h('section', { class: 'depgroup' }, [head, table(['PPC number', 'Work instruction', 'Title', 'Requested by', 'Specialist', 'Published'], trs, 'None.')]);
  });
  return h('div', { class: 'view', id: 'v-published', hidden: true }, [h('h1', { text: 'Published documents' }),
    h('p', { class: 'lead' }, [h('strong', { text: pl(all.length, 'published document') + '.' }), ' Grouped by department. ' + dated + ' of ' + rows.length + ' shown ' + (dated === 1 ? 'has' : 'have') + ' a published date recorded.']),
    h('div', { class: 'pubtools' }, [box, btn('Open every department', { quiet: true, onclick: () => { names.forEach(n => { deptOpen[n] = true; }); rebuild('published'); } }), btn('Close every department', { quiet: true, onclick: () => { names.forEach(n => { deptOpen[n] = false; }); rebuild('published'); } }), btn('Download as CSV', { quiet: true, onclick: () => downloadPublished(rows) })]),
    q && !rows.length ? h('p', { class: 'empty', text: 'Nothing matches "' + pubFilter + '".' }) : null,
    h('section', null, sections)]);
}
function downloadPublished(list) {
  const heads = ['Department', 'PPC number', 'Work instruction', 'Title', 'Requested by', 'Specialist', 'Published'];
  const lines = [heads].concat(list.slice().sort((a, b) => deptLabel(a.department).localeCompare(deptLabel(b.department)) || (b.published || '').localeCompare(a.published || ''))
    .map(r => [deptLabel(r.department), r.number || '', r.wi || '', r.title || '', r.requestor || '', r.specialist || '', fmtDate(r.published)])).map(r => r.map(csvCell).join(','));
  const a = h('a', { href: URL.createObjectURL(new Blob([lines.join('\r\n')], { type: 'text/csv' })), download: 'ppc-published-by-department.csv' });
  document.body.append(a); a.click(); a.remove();
}
/* The quick start guide. Each role is shown only the parts that apply to it, and the rules and numbers are read
   from the live settings rather than written out, so the guide cannot drift away from how the system behaves. */
function vGuide() {
  const cfg = D.config, role = D.user.role, isAdmin = !!D.can.allStatuses;
  const title = ({ ADMIN: 'administrator', DIRECTOR: 'director', MANAGER: 'manager', SPECIALIST: 'document specialist', SUPERVISOR: 'supervisor', REQUESTOR: 'requester' })[role] || role.toLowerCase();
  const sec = (heading, kids) => h('section', null, [h('h2', { text: heading })].concat(kids.filter(Boolean)));
  const dl = (pairs) => h('dl', { class: 'kv' }, pairs.filter(Boolean).flatMap(p => [h('dt', { text: p[0] }), h('dd', null, p[1])]));
  const steps = (items) => h('ol', { class: 'steps4' }, items.filter(Boolean).map(t => h('li', null, t)));
  const out = [h('h1', { text: 'Quick start guide' }),
    h('p', { class: 'lead' }, ['Prepared for ', h('strong', { text: D.user.name }), ', ' + title + '. It covers the parts of the system your account is able to use.',
      D.writes ? '' : ' The page is currently read-only, so changes cannot be saved.'])];

  /* Sending a request: everyone who is allowed to create one. */
  if (D.can.request) out.push(sec('Submitting a request', [
    steps(['Select New PPC request. Choose the requesting department and the work instruction number, if the document already exists.',
      'Complete the SOP-06-01 proposal form. The initiator, change title, description, justification and contact email are required.',
      'Attach any photographs at the foot of the form. They are uploaded when the request is submitted.',
      'Review the summary and submit. A PPC number is issued at that point and the request is recorded as Not started.']),
    h('p', { class: 'mut', text: 'A request without a work instruction number is held in Missing PPC, need update until document control assigns one. The system does not generate work instruction numbers, as these come from GLG.' })]));

  /* The five steps: anyone who can edit needs this. */
  if (D.can.edit) out.push(sec('The five stages of a PPC', [
    h('p', { class: 'hint', text: 'A PPC occupies one stage at any time. It is recorded as Not started on submission and concludes at Published.' }),
    h('ol', { class: 'steps' }, STATUS_COLS.map((c, i) => h('li', null, [h('span', { class: 'sn', text: String(i + 1) }), h('span', { text: c.title })]))),
    dl([['Advancing a PPC', 'Select Edit on the record, then use the button at the head of the panel. It advances the PPC by one stage and saves any other changes made in the panel at the same time.'],
      ['Reversing a stage', 'The Back button returns the PPC to the previous stage, for corrections.'],
      ['Your authority', isAdmin ? 'As an administrator you may set any of the five stages.' : 'You may set Not started and In progress. Ready for signature, Completed and Published are reserved for administrators.'],
      ['Recorded automatically', 'Reaching In progress starts the age clock. Reaching Completed or Published records that date if one is not already held.']])]));
  else out.push(sec('How work progresses', [
    h('p', { class: 'hint', text: 'Document control moves each PPC through five stages. Your account has view access, so the board shows progress but cannot change it.' }),
    h('ol', { class: 'steps' }, STATUS_COLS.map((c, i) => h('li', null, [h('span', { class: 'sn', text: String(i + 1) }), h('span', { text: c.title })])))]));

  /* Reading the board: everyone. */
  out.push(sec('Reading the board', [
    h('p', { class: 'hint', text: 'The age clock runs only while a PPC is In progress and is counted from the date it reached that stage.' }),
    h('ul', { class: 'clocklegend' }, [['green', '0 to 15 days'], ['yellow', '16 to 29 days'], ['overdue', '30 days or more, shown as Overdue']].map(p => h('li', null, [h('span', { class: 'sw ' + p[0] }), p[1]]))),
    h('p', { class: 'mut', text: 'The In progress column is divided into those three ranges, each with its own count, ordered from green through to overdue. Selecting a column heading, or a range heading, closes it and gives the remaining columns more width.' })]));

  /* Recording a test: only those who may. */
  if (D.can.test) out.push(sec('Recording a test', [
    steps(['Select New PPC test, or Record a test against a PPC shown in the Testing stage.',
      'Choose the PPC, then set the test dates and the person responsible.',
      'Complete the SOP-06-02 form and attach any photographs at the foot of it.',
      'Review and submit. The test is linked to the PPC and its end date is recorded.']),
    dl([['Minimum duration', cfg.minTestDays + ' calendar days. This is enforced by the form and again by the server.'],
      ['Reminders', 'Issued from ' + cfg.reminderLeadDays + ' days before the test end date and repeated every ' + cfg.reminderCooldownDays + ' days once it has passed.'],
      ['Current email setting', cfg.emailMode === 'LIVE' ? 'Live. Reminders and notifications are sent and recorded.' : 'Dry run. Reminders are recorded but not sent.']])]));

  /* Images: anyone who can attach. */
  if (D.can.attach) out.push(sec('Attaching photographs', [
    steps([D.can.request ? 'On a new request' + (D.can.test ? ' or test' : '') + ', use the Images panel at the foot of the form. Files are uploaded when the record is submitted.' : null,
      D.can.edit ? 'On an existing record, select Edit and scroll to Images.' : 'Images already held against a PPC are shown when document control opens the record. Your account cannot add to an existing PPC.',
      'Select a file, or use the paste area and press Ctrl+V to insert a screenshot from the clipboard.',
      'Select View to open an image at full size' + (D.can.edit ? ', or Delete to remove it.' : '. Images you added yourself can be deleted.')]),
    h('p', { class: 'mut', text: 'Accepted formats are PNG, JPEG, GIF and WebP, up to 4 MB per file and 20 per record. SVG files are not accepted. Images are visible only to those permitted to see the PPC.' })]));

  /* The pages: describe only what this role will actually use. */
  const pages = [['Dashboard', 'The stage board, with items requiring attention listed beneath it.'],
    ['Quick start guide', 'This page.'],
    ['Published documents', 'Completed documents grouped by department, with a search facility and a CSV export.'],
    ['Missing PPC, need update', 'Open requests received before PPC numbering began.' + (D.can.edit ? ' These can be edited; number assignment is not yet available.' : ' View only for your account.')],
    ['Tests and reminders', 'PPCs at the Testing stage and the tests recorded against them.' + (D.can.test ? ' Each test holds its own images.' : ' View only for your account.')],
    ['Routing and emails', 'PPCs marked Completed, with a preview of the notification a requester would receive.' + (D.can.edit ? ' Sending is not yet available.' : ' View only for your account.')],
    ['Update requests', 'Not yet available.'],
    ['Year records', 'Requests grouped by year of submission, with a printable record and a CSV export.']];
  out.push(sec('The pages in the menu', [dl(pages)]));

  /* Administrator-only settings. */
  if (isAdmin) out.push(sec('Administrator settings', [
    dl([['Saving from this page', D.writes ? 'Enabled. Set the script property PPC_API_WRITES to OFF to suspend it.' : 'Suspended. Set the script property PPC_API_WRITES to ON to enable saving.'],
      ['Email', 'Controlled by EMAIL_MODE in the PPC_CONFIG tab. It is currently ' + (cfg.emailMode === 'LIVE' ? 'LIVE.' : 'DRY_RUN, so nothing is sent.')],
      ['Roles', 'Held in the PPC_USERS tab. A requester sees only their own requests and never the internal notes.'],
      ['Suspending the page', 'Set PPC_API_ENABLED to OFF to close the page to all users.']])]));
  return h('div', { class: 'view guide', id: 'v-guide', hidden: true }, out);
}
/* A short guide for requesters, shown on their own page. It explains only what they can see and do. */
function requesterGuide() {
  const open = !!requesterGuide.open;
  const head = h('button', { type: 'button', class: 'dephead', 'aria-expanded': String(open), onclick: () => { requesterGuide.open = !open; renderRequester(); } },
    [h('span', { class: 'caret', 'aria-hidden': 'true', text: open ? '\u2212' : '+' }), h('b', { text: 'How this page works' })]);
  if (!open) return h('section', { class: 'depgroup reqguide' }, head);
  return h('section', { class: 'depgroup reqguide' }, [head, h('div', { class: 'reqguidebody' }, [
    h('p', { text: 'This page lists the document change requests recorded under your email address, and the stage each one has reached.' }),
    h('dl', { class: 'kv' }, [['Not started', 'Received by document control and awaiting work.'],
      ['In progress', 'Being drafted or tested.'], ['Complete', 'Published, or completed and awaiting publication.']]
      .flatMap(p => [h('dt', { text: p[0] }), h('dd', { text: p[1] })])),
    h('p', { text: 'Select a request to see its progress and the latest note from document control. Internal working notes are not shown here.' }),
    D.can.attach ? h('p', { text: 'Photographs may be added to your own requests from the Images panel, using either a file or Ctrl+V to paste a screenshot. PNG, JPEG, GIF and WebP are accepted, up to 4 MB each.' }) : null,
    h('p', { class: 'mut', text: 'Requesting an update from this page is not yet available. Contact document control directly in the meantime.' })])]);
}
function vUpdates() {
  return h('div', { class: 'view', id: 'v-updates', hidden: true }, [h('h1', { text: 'Update requests' }), h('p', { class: 'lead' }, [h('strong', { text: 'Not built yet.' }), ' Requesters cannot ask for an update from the page until this exists.']),
    h('section', null, [h('h2', { text: 'The plan' }), h('ol', { class: 'steps4' }, ['A requester opens a request on their page and chooses Request an update.', 'It appears here, and the document specialist gets an email.', 'You write a short public update on the PPC. Internal notes are never shown to requesters.', 'You mark it answered, and the requester gets an email with the status and your note.'].map(t => h('li', { text: t })))])]);
}
function allRows() {
  return [].concat(D.register.map(r => Object.assign({ tab: (r.status === 'Completed' || r.status === 'Published') ? 'published' : (isPending(r) ? 'needsNumber' : 'register') }, r)), D.published.map(r => Object.assign({ tab: 'published' }, r)), D.blocked.map(r => Object.assign({ tab: 'blocked' }, r)), D.needsNumber.map(r => Object.assign({ tab: 'needsNumber' }, r)));
}
function yearStats(rows) {
  const s = { total: rows.length, done: 0, doneDated: 0, open: 0, need: 0, blocked: 0 };
  rows.forEach(r => { if (r.tab === 'published') { s.done++; if (r.published) s.doneDated++; } else if (r.tab === 'blocked') s.blocked++; else if (r.tab === 'needsNumber') s.need++; else s.open++; });
  return s;
}
function rowStatus(r) { return r.tab === 'published' ? (r.status === 'Completed' ? 'Completed' : 'Published') : r.tab === 'blocked' ? 'Marked blocked' : r.tab === 'needsNumber' ? 'In progress, no PPC number' : (r.status || 'Open'); }
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
const BUILDERS = { overview: vOverview, guide: vGuide, published: vPublished, missing: vMissing, tests: vTests, routing: vRouting, updates: vUpdates, years: vYears };
function rebuild(id) { const old = $('#v-' + id); const fresh = BUILDERS[id](); old.replaceWith(fresh); fresh.hidden = currentView !== id; }
function show(id) {
  if (!BUILDERS[id]) id = 'overview'; currentView = id;
  $$('#views .view').forEach(v => { v.hidden = v.id !== 'v-' + id; });
  $$('#nav a').forEach(a => { if (a.dataset.view === id) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
  try { history.replaceState(null, '', '#' + id); } catch (e) { /* the page lives in a sandboxed frame; the address bar is not ours */ } window.scrollTo(0, 0);
}
let FLASH = '';
function banner(el) {
  el.classList.toggle('ok', !!FLASH);
  if (FLASH) { el.hidden = false; el.textContent = FLASH; FLASH = ''; return; }
  el.hidden = D.writes; if (!D.writes) el.textContent = 'Read-only for now. You can look around and fill in the forms, but nothing is saved from this page yet, including edits.';
}
function renderStaff() {
  $('#nav').replaceChildren(...VIEWS.map(v => h('a', { href: '#' + v[0], 'data-view': v[0], onclick: e => { e.preventDefault(); show(v[0]); } }, [v[1], v[0] === 'missing' && D.needsNumber.length ? h('span', { class: 'b', text: String(D.needsNumber.length) }) : null])));
  setKids($('#topbar'), [D.can.test ? btn('New PPC test', { open: 'test' }) : null, D.can.request ? btn('New PPC request', { open: 'request', primary: true }) : null]);
  $('#who').textContent = D.user.name + ', ' + D.user.role.toLowerCase();
  banner($('#banner')); $('#views').replaceChildren(...Object.keys(BUILDERS).map(id => BUILDERS[id]()));
  wireOpeners(); show(currentView || (location.hash || '').slice(1));
}

/* ---------- requester view ---------- */
const STEPS = ['Request received', 'In progress', 'Ready for signature', 'Completed', 'Published'];
let reqSel = null;
function whereText(r) {
  if (r.tab === 'blocked') return 'On hold';
  if (r.status === 'Published' || r.tab === 'published') return 'Published' + (r.published ? ' ' + fmtDate(r.published) : '');
  if (r.status === 'Completed') return 'Completed' + (r.routing ? ' ' + fmtDate(r.routing) : '');
  if (r.numberPending) return 'In progress, waiting for a PPC number from document control';
  if (r.status === 'Ready for Signature') return 'Waiting for management signatures';
  if (r.status === 'Not Started') return 'Not started yet';
  return r.stage === 'Testing' ? 'In progress, testing' : 'In progress, drafting the work instruction';
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
      h('ol', { class: 'tl' }, STEPS.map((s, i) => h('li', { class: i < r.step ? 'done' : i === r.step ? 'now' : '', text: s }))),
      h('p', null, [h('b', { text: 'Where it is: ' }), whereText(r)]), h('h3', { class: 'nh', text: 'Latest note from document control' }),
      r.publicNotes && !/^Imported from legacy source/.test(r.publicNotes) ? h('p', { class: 'note2', text: r.publicNotes }) : h('p', { class: 'note2 mut', text: 'No update has been written yet.' }),
      h('p', { class: 'mut', text: 'Asking document control for an update from this page is not built yet.' }));
    if (D.can.attach && r.number) det.append(attachPanel(r.number, { canAdd: true, canDelete: true }));
    else if (D.can.attach) det.append(h('p', { class: 'mut small', text: 'Images can be added once document control gives this request a PPC number.' }));
  }
  setKids($('#rview'), [h('h1', { text: 'My requests' }), h('p', { class: 'lead' }, ['You have ', h('strong', { text: pl(mine.length, 'request') }), '. Pick one to see where it is.']),
    h('p', { class: 'keyline' }, ['Not started', 'In progress', 'Complete'].map(s => h('span', null, [h('b', { text: String(cnt(s)) }), s]))),
    D.can.request ? null : h('p', { class: 'mut', text: 'Sending a new request from this page is not turned on for your account.' }),
    h('div', { class: 'cols' }, [h('div', null, [].concat(group('In progress'), group('Complete'), mine.length ? [] : [h('p', { class: 'empty', text: 'No requests found.' })], [requesterGuide()])), det])]);
  wireOpeners();
}
function renderCurrent() { if (!D) return; if (D.view === 'staff') renderStaff(); else renderRequester(); }
function showApp() {
  if (!WZ.request) initWizards();
  $('#gate').hidden = true; $('#shellStaff').hidden = D.view !== 'staff'; $('#shellReq').hidden = D.view !== 'requester';
  document.title = APP_NAME + ', Zenni USA'; renderCurrent();
}

/* ---------- wizards (your 3-step format) ---------- */
const openers = new WeakSet();
function wireOpeners() { $$('[data-open]').forEach(b => { if (openers.has(b)) return; openers.add(b); b.addEventListener('click', () => (b.dataset.open === 'request' ? WZ.request : WZ.test).open(b, b.dataset.ppc || null)); }); }
const WZ = {};
function ppcsForWizard() { const src = D.view === 'staff' ? D.register : D.mine.filter(m => m.tab === 'register' && m.number); return src.filter(r => r.status !== 'Published' && r.status !== 'Completed' && r.number && !isPending(r)).map(r => ({ n: r.number, dept: r.number.split('-')[1], title: r.title, wi: r.wi, status: r.status })); }
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
  const rqStage = stageImages('Images (optional)');
  $('#rqForm').append(rqStage);
  WZ.request = wizard(rq, {
    l1: 'Next: Fill Request Form →', submit: 'Submit PPC Proposal',
    reset() { dept.value = ''; act.value = 'NEW'; meth.value = 'AUTO'; man.value = ''; wi.value = ''; $('#rqForm').reset(); rqStage.clear(); $('#rqDone').hidden = true; $('#rqDate').value = new Date().toISOString().slice(0, 10); $('#rqInit').value = D.user.name; $('#rqEmail').value = D.user.email; $('.wz-nav [data-next]', rq).disabled = false; fillEx(); rqRefresh(); },
    v1() { const e = []; if (!dept.value) e.push('Select a requesting department.'); if (act.value === 'EXISTING' && !ex.value) e.push('Choose the existing PPC.'); if (act.value === 'NEW' && meth.value === 'MANUAL' && !/^\d{4}-[A-Z]{2,8}-\d{3,6}-\d{1,6}$/.test(man.value.trim().toUpperCase())) e.push('Enter a manual number like 2026-SUR-246-014.'); if (/^The WI says/.test(numText())) e.push('The WI department does not match.'); return e; },
    review() {
      const f = [['Department', dept.value ? dept.options[dept.selectedIndex].text : 'Not selected'], ['WI number(s)', wi.value || 'None yet'], ['PPC mode', act.value === 'NEW' ? 'New PPC' : 'Existing PPC'], ['PPC number', numText()], ['Initiator', val('#rqInit') || 'Missing'], ['Change title', val('#rqChTitle') || 'Missing'], ['Target date', val('#rqDue') || 'Not provided'], ['Description', val('#rqDesc') || 'Missing'], ['Justification', val('#rqJust') || 'Missing'], ['Email for updates', val('#rqEmail') || 'Missing'], ['Images', rqStage.count() ? pl(rqStage.count(), 'image') + ', uploaded when you submit' : 'None']];
      $('#rqRev').replaceChildren(...f.flatMap(r => [h('dt', { text: r[0] }), h('dd', { text: r[1] })]));
      const m = []; if (!val('#rqInit').trim()) m.push('Initiator'); if (!val('#rqChTitle').trim()) m.push('Change title'); if (!val('#rqDesc').trim()) m.push('Description'); if (!val('#rqJust').trim()) m.push('Justification'); if (!val('#rqEmail').trim()) m.push('Email');
      const mm = $('#rqMiss'); mm.textContent = m.length ? 'Please complete: ' + m.join(', ') + '.' : 'All core fields are entered. Check the details, then submit.'; mm.className = m.length ? 'wz-miss' : 'wz-ok'; $('#rqDone').hidden = true;
    },
    async send() {
      const out = $('#rqDone'), nx = $('.wz-nav [data-next]', rq); out.hidden = false; out.textContent = 'Saving...'; nx.disabled = true;
      try {
        const r = await api('submitRequest', { ppcAction: act.value, department: dept.value, existingPpc: ex.value, wiNumbers: wi.value.trim(), assignment: act.value === 'NEW' ? meth.value : 'AUTO', manualPpc: man.value.trim(), requestor: val('#rqInit').trim(), email: val('#rqEmail').trim(), title: val('#rqChTitle').trim(), description: val('#rqDesc'), justification: val('#rqJust'), formPriority: ($('input[name=prio]:checked', rq) || {}).value || 'Low', requestDate: val('#rqDate'), dueDate: val('#rqDue'), manager: val('#rqMgr'), poc: val('#rqPoc'), changeType: val('#rqType'), effort: val('#rqEffort'), impactQuality: val('#rqImpQ'), impactCost: val('#rqImpC'), impactSchedule: val('#rqImpS'), impactResources: val('#rqImpR'), impactOther: val('#rqImpO') });
        out.textContent = 'Saved. PPC number: ' + r.id + '.';
        if (rqStage.count()) {
          out.textContent = 'Saved as ' + r.id + '. Uploading ' + pl(rqStage.count(), 'image') + '...';
          const up = await rqStage.upload(r.id);
          out.textContent = 'Saved as ' + r.id + '. ' + (up.done ? pl(up.done, 'image') + ' attached.' : '') + (up.failed.length ? ' These did not upload: ' + up.failed.join('; ') + '. Add them again from Edit.' : '');
        }
        await refresh();
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
  const tsStage = stageImages('Test images (optional)');
  $('#tsForm').append(tsStage);
  WZ.test = wizard(ts, {
    l1: 'Next: Fill Test Form →', submit: 'Submit PPC Test',
    reset(pre) { tD.value = ''; tP.replaceChildren(h('option', { value: '', text: 'Choose a department first' })); [tS, tE, tPoc].forEach(e => { e.value = ''; }); $('#tsForm').reset(); tsStage.clear(); $('#tsDone').hidden = true; $('#tsNote').textContent = 'Reminders start ' + D.config.reminderLeadDays + ' days before the end date and repeat every ' + D.config.reminderCooldownDays + ' days. Email mode is ' + (D.config.emailMode === 'LIVE' ? 'live.' : 'dry run, so reminders are logged but not sent yet.'); if (pre) { const p = ppcsForWizard().find(x => x.n === pre); if (p) { tD.value = p.dept; tsFill(); tP.value = p.n; } } tsDetails(); dates(); },
    v1() { const e = []; if (!tD.value) e.push('Choose a department.'); if (!tP.value) e.push('Choose an existing PPC.'); if (!tS.value || !tE.value) e.push('Select test start and end dates.'); else if (span(tS.value, tE.value) < D.config.minTestDays) e.push('The test must span at least ' + D.config.minTestDays + ' calendar days.'); if (!tPoc.value.trim()) e.push('Enter the main test POC.'); return e; },
    review() {
      const p = cur(), f = [['PPC number', p ? p.n : 'Missing'], ['Department', tD.value || 'Missing'], ['Test reason', val('#tsReason') || 'Missing'], ['Main POC', tPoc.value || 'Missing'], ['Owner email', val('#tsEmail') || 'Missing'], ['Test dates', (tS.value || '—') + ' to ' + (tE.value || '—')], ['Define the test', val('#tsDefine') || 'Missing'], ['Expected results', val('#tsExpected') || 'Missing'], ['Images', tsStage.count() ? pl(tsStage.count(), 'image') + ', uploaded when you submit' : 'None']];
      $('#tsRev').replaceChildren(...f.flatMap(r => [h('dt', { text: r[0] }), h('dd', { text: r[1] })]));
      const m = []; if (!val('#tsReason').trim()) m.push('Test reason'); if (!val('#tsDefine').trim()) m.push('Define the test'); if (!val('#tsExpected').trim()) m.push('Expected results'); if (!val('#tsEmail').trim()) m.push('Owner email');
      const mm = $('#tsMiss'); mm.textContent = m.length ? 'Complete before submitting: ' + m.join(', ') + '.' : 'Ready to submit. Review the values for accuracy.'; mm.className = m.length ? 'wz-miss' : 'wz-ok'; $('#tsDone').hidden = true;
    },
    async send() {
      const out = $('#tsDone'), nx = $('.wz-nav [data-next]', ts); out.hidden = false; out.textContent = 'Saving...'; nx.disabled = true;
      try {
        const r = await api('submitTest', { ppc: tP.value, owner: tPoc.value.trim(), ownerEmail: val('#tsEmail').trim(), reason: val('#tsReason'), start: tS.value, end: tE.value, defineTest: val('#tsDefine'), variants: val('#tsVar'), timeline: val('#tsTl'), resources: val('#tsRes'), expected: val('#tsExpected'), actual: val('#tsActual'), evidence: val('#tsEvid').trim(), result: val('#tsResult'), testNumber: ($('input[name=tn]:checked', ts) || {}).value || '1', signer1: val('#tsSig1'), signer2: val('#tsSig2'), signer3: val('#tsSig3') });
        out.textContent = 'Saved. Test ' + r.id + '.';
        if (tsStage.count()) {
          out.textContent = 'Saved as ' + r.id + '. Uploading ' + pl(tsStage.count(), 'image') + '...';
          const up = await tsStage.upload(tP.value, r.id);
          out.textContent = 'Saved as ' + r.id + '. ' + (up.done ? pl(up.done, 'image') + ' attached to the test.' : '') + (up.failed.length ? ' These did not upload: ' + up.failed.join('; ') + '.' : '');
        }
        await refresh();
      } catch (e) { out.textContent = e.code === 'WRITES_OFF' ? 'Saving from this page is switched off for now. Nothing was saved.' : e.message; nx.disabled = false; }
    }
  });
}

/* ---------- launcher (what PPC.html does when you open it from the LMS) ---------- */
function flag(v) { return v === true || v === 1 || (typeof v === 'string' && ['true', 'yes', '1'].includes(v.toLowerCase().trim())); }
/* Browser storage can be unreadable: a privacy setting, or a frame the browser treats as third party. Reading it must
   never throw, or the entry page would stall with nothing on screen. */
let SS_BLOCKED = false;
function ssGet(k) { try { return sessionStorage.getItem(k); } catch (e) { SS_BLOCKED = true; return null; } }
function readSS(k) { try { return JSON.parse(ssGet(k) || 'null') || {}; } catch (e) { return {}; } }
/* The LMS stores its session in this tab before opening PPC. This reads a display name from it when one is there,
   checking the shapes an LMS commonly uses. It never invents one: with nothing to read, the page says so. */
function lmsWho() {
  const pick = o => {
    if (!o || typeof o !== 'object') return '';
    const direct = o.displayName || o.display_name || o.fullName || o.full_name || o.name || o.user_name || o.username || o.email || o.userEmail || o.user_email || '';
    if (typeof direct === 'string' && direct.trim()) return direct.trim();
    for (const k of ['user', 'profile', 'account', 'session', 'identity']) { const v = pick(o[k]); if (v) return v; }
    return '';
  };
  for (const k of ['lms_user', 'lms_profile', 'lms_account', 'lms_session', 'lms_identity', 'lms_visibility']) { const v = pick(readSS(k)); if (v) return v; }
  for (const k of ['lms_display_name', 'lms_name', 'lms_user_name', 'lms_full_name', 'lms_email', 'lms_user_email']) {
    const v = ssGet(k); if (v && v.trim() && v.trim() !== 'null') return v.trim().replace(/^"|"$/g, '');
  }
  return '';
}
function lmsProblem() {
  const flagv = ssGet('lms_logged_in');
  if (SS_BLOCKED) return 'This tab cannot read the LMS session, because the browser is blocking site storage here. Open PPC from the LMS dashboard, or allow storage for this site.';
  if (flagv !== 'true') return 'No active LMS login in this tab. Return to LMS login, then open PPC from the dashboard.';
  const vis = readSS('lms_visibility');
  const dept = Object.assign({}, vis.departments || {}, readSS('lms_departments'));
  const feat = Object.assign({}, vis.features || {}, readSS('lms_features'));
  if (!flag(dept.PPC_Document_Control) || !flag(feat.PPC_Document_Control_ViewHub)) return "PPC isn't enabled in this LMS session. Check the existing LMS permissions and sign back in.";
  return '';
}
function launcherMode() {
  const open = $('#open'); if (!open) return;
  document.title = APP_NAME + ', Zenni USA';
  $$('[data-appname]').forEach(el => { el.textContent = APP_NAME; });
  const problem = lmsProblem(), who = lmsWho();
  const line = (state, text) => h('div', { class: 'chk ' + state }, [h('span', { class: 'mk', 'aria-hidden': 'true', text: state === 'ok' ? '\u2713' : state === 'bad' ? '\u00d7' : '\u25cb' }), h('span', { text: text })]);
  const checks = [];
  if (problem) checks.push(line('bad', problem));
  else {
    checks.push(line('ok', who ? 'Signed in to the LMS as ' + who : 'Signed in to the LMS in this tab'));
    if (!who) checks.push(line('note', 'The LMS did not pass a name to this page, so it cannot be shown here.'));
    checks.push(line('note', 'Your Google Workspace account and PPC_USERS entry are checked by the PPC application when it opens.'));
  }
  setKids($('#checks'), checks);
  $('#statusbar').textContent = problem ? 'Access unavailable' : 'Ready';
  if (problem) return;
  open.disabled = false;
  open.addEventListener('click', () => { if (open.disabled) return; if (EMBED) showEmbed(); else window.location.assign(PPC_SECURE_URL); });
}

const GOOGLE_ORIGIN = /^https:\/\/([a-z0-9-]+\.)*(google\.com|googleusercontent\.com)$/;
function showEmbed() {
  $('#launcher').hidden = true;
  const hint = h('span', { class: 'embedhint', text: 'Loading the secure PPC application...' });
  const frame = h('iframe', { id: 'ppcFrame', title: APP_NAME, src: PPC_SECURE_URL, referrerpolicy: 'no-referrer' });
  const bar = h('div', { class: 'embedbar' }, [h('b', { text: APP_NAME }), hint, h('span', { class: 'grow' }), h('a', { href: PPC_SECURE_URL, target: '_blank', rel: 'noopener' }, 'Open in a new tab'), h('a', { href: 'index.html' }, 'Return to the LMS')]);
  document.body.append(h('div', { id: 'embedWrap', class: 'embedwrap' }, [bar, frame]));
  const timer = setTimeout(() => { hint.textContent = 'This is taking a while. If nothing appears, your browser may be blocking Google sign-in inside this page. Use "Open in a new tab".'; hint.classList.add('warn'); }, EMBED_TIMEOUT_MS);
  window.addEventListener('message', e => { if (GOOGLE_ORIGIN.test(e.origin) && e.data && e.data.ppc === 'ready') { clearTimeout(timer); hint.textContent = ''; } });
}
/* The secure app can be framed by Google (its own wrapper) and by your LMS site, and by nobody else. If any other site frames it, it refuses to draw anything. */
function framedByStranger() {
  try { const ao = window.location.ancestorOrigins; if (!ao) return false; for (let i = 0; i < ao.length; i++) { if (ao[i] !== PARENT_ORIGIN && !GOOGLE_ORIGIN.test(ao[i])) return true; } return false; } catch (e) { return false; }
}

/* ---------- start ---------- */
async function boot() {
  if (!(window.google && google.script && google.script.run)) { launcherMode(); return; }   // not inside Google's secure app: act as the launcher
  if (framedByStranger()) { document.body.textContent = 'This page can only be shown inside the Zenni LMS.'; return; }
  const bootNote = document.getElementById('boot'); if (bootNote) bootNote.remove();
  try {
    const res = await fetch(ASSET_BASE + ASSET_NAME + '.html', { credentials: 'omit' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const tpl = new DOMParser().parseFromString(await res.text(), 'text/html').querySelector('template#app-layout');
    if (!tpl) throw new Error('layout template missing');
    document.body.insertAdjacentHTML('beforeend', tpl.innerHTML);     // our own file from our own repo
  } catch (e) { document.body.textContent = 'The page could not load its layout (' + ASSET_NAME + '.html). Check that PPC.html, PPC.css and PPC.js on GitHub are all the new versions.'; return; }
  $$('a[data-lms]').forEach(a => { a.href = LMS_URL; a.target = '_top'; });
  $('#retry').addEventListener('click', loadApp);
  await loadApp();
  try { window.top.postMessage({ ppc: 'ready' }, '*'); } catch (e) { /* nothing to tell */ }   // carries no data; the launcher only listens to Google origins
}
boot();