const API_URL = 'https://script.google.com/a/macros/zennioptical.com/s/AKfycbwrNOktjZuXil1FhD6QPGo-gfsFOD0AAC1wdJOtxfhmQuGAltXYC4OT5y4EwgcATGfmpQ/exec';

const DASHBOARDS = [
  'LMS Command Center',
  'Investigation Hold Management',
  'Lens Intake & Upload',
  'True Curve & Power Analysis',
  'Surface Dashboard',
  'AR Dashboard',
  'AR Run Tracker',
  'Finish Dashboard',
  'Coating Breakage Analytics',
  'Power Analysis Tool',
  'Scanner Network Monitor',
  'Incoming Jobs Dashboard',
  'Picking Dashboard',
  'Productivity Hub',
  'Breakage / Quality Hub',
  'Facility Breakage',
  'Other'
];

const WORK_TYPES = [
  'Feature Added',
  'Bug Fixed',
  'Testing',
  'Automation',
  'Accuracy Validation',
  'Documentation',
  'UI Improvement',
  'Performance / Reliability',
  'Planning'
];

const STATUS_OPTIONS = [
  'Completed',
  'In Progress',
  'Needs Testing',
  'Open Issue',
  'Blocked'
];

const DOC_OPTIONS = [
  'Open',
  'In Progress',
  'Completed'
];

const DOC_STATUS_OPTIONS = [
  'Not Started',
  'In Progress',
  'Completed'
];

const NEXT_MILESTONES = [
  ['Automate snapshot cleanup', 'Next focus'],
  ['Document refresh scheduling', 'Documentation'],
  ['Create maintenance checklist', 'Reliability']
];

let appState = {
  actions: [],
  workLog: [],
  accuracyChecks: [],
  docs: [],
  monthlySummaries: [],
  progress: {
    percent: 0,
    completed: 0,
    total: 0
  }
};

document.addEventListener('DOMContentLoaded', () => {
  setDefaultDates();
  setupNavigation();
  setupDropdowns();
  setupButtons();
  refreshTracker();
});

function setupNavigation() {
  document.querySelectorAll('.nav-btn').forEach(button => {
    button.addEventListener('click', () => showView(button.dataset.view));
  });

  document.querySelectorAll('[data-view-jump]').forEach(button => {
    button.addEventListener('click', () => showView(button.dataset.viewJump));
  });
}

function showView(view) {
  if (!view) return;

  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === view);
  });

  document.querySelectorAll('.view').forEach(section => {
    section.classList.toggle('active', section.id === view);
  });
}

function setupDropdowns() {
  fillSelect('workDashboard', DASHBOARDS);
  fillSelect('checkDashboard', DASHBOARDS);
  fillSelect('docDashboard', DASHBOARDS);
  fillSelect('workType', WORK_TYPES);
  fillSelect('workStatus', STATUS_OPTIONS);
  fillSelect('docDataSources', DOC_OPTIONS);
  fillSelect('docFormulas', DOC_OPTIONS);
  fillSelect('docRefreshTiming', DOC_OPTIONS);
  fillSelect('docSnapshotLogic', DOC_OPTIONS);
  fillSelect('docApiDependencies', DOC_OPTIONS);
  fillSelect('docTroubleshooting', DOC_OPTIONS);
  fillSelect('docStatus', DOC_STATUS_OPTIONS);
}

function setupButtons() {
  bindClick('refreshBtn', refreshTracker);
  bindClick('saveWorkBtn', saveWorkLog);
  bindClick('saveAccuracyBtn', saveAccuracyCheck);
  bindClick('saveDocsBtn', saveDocumentation);
  bindClick('generateSummaryBtn', generateSummary);
  bindClick('copySummaryBtn', copySummary);

  const docDashboard = document.getElementById('docDashboard');
  if (docDashboard) {
    docDashboard.addEventListener('change', loadDocEditor);
  }
}

function bindClick(id, handler) {
  const element = document.getElementById(id);
  if (element) element.addEventListener('click', handler);
}

function setDefaultDates() {
  const today = new Date();
  setValue('workDate', today.toISOString().slice(0, 10));
  setValue('checkDate', today.toISOString().slice(0, 10));
  setValue('summaryMonth', today.toISOString().slice(0, 7));
}

async function refreshTracker() {
  showToast('Refreshing mission tracker...');

  try {
    const data = await apiGet({ action: 'getAppData' });
    appState = data;
    renderApp();
    showToast('Mission tracker refreshed.');
  } catch (error) {
    console.error(error);
    showToast('Refresh failed. Check API deployment and access.');
  }
}

async function saveWorkLog() {
  const payload = {
    action: 'saveWorkLog',
    workDate: getValue('workDate'),
    dashboard: getValue('workDashboard'),
    workType: getValue('workType'),
    status: getValue('workStatus'),
    description: getValue('workDescription'),
    evidence: getValue('workEvidence')
  };

  if (!payload.description) {
    showToast('Add a description first.');
    return;
  }

  try {
    showToast('Saving work log...');
    const data = await apiGet(payload);
    appState = data;
    setValue('workDescription', '');
    setValue('workEvidence', '');
    renderApp();
    showToast('Work log saved.');
  } catch (error) {
    console.error(error);
    showToast('Work log failed to save.');
  }
}

async function saveAccuracyCheck() {
  const payload = {
    action: 'saveAccuracyCheck',
    checkDate: getValue('checkDate'),
    dashboard: getValue('checkDashboard'),
    metricName: getValue('metricName'),
    sourceTotal: getValue('sourceTotal'),
    dashboardTotal: getValue('dashboardTotal'),
    notes: getValue('checkNotes')
  };

  if (!payload.metricName) {
    showToast('Add a metric name first.');
    return;
  }

  try {
    showToast('Saving accuracy check...');
    const data = await apiGet(payload);
    appState = data;
    setValue('metricName', '');
    setValue('sourceTotal', '');
    setValue('dashboardTotal', '');
    setValue('checkNotes', '');
    renderApp();
    showToast('Accuracy check saved.');
  } catch (error) {
    console.error(error);
    showToast('Accuracy check failed to save.');
  }
}

async function saveDocumentation() {
  const payload = {
    action: 'updateDocumentation',
    dashboard: getValue('docDashboard'),
    dataSources: getValue('docDataSources'),
    formulas: getValue('docFormulas'),
    refreshTiming: getValue('docRefreshTiming'),
    snapshotLogic: getValue('docSnapshotLogic'),
    apiDependencies: getValue('docApiDependencies'),
    troubleshooting: getValue('docTroubleshooting'),
    status: getValue('docStatus'),
    notes: getValue('docNotes')
  };

  try {
    showToast('Saving documentation status...');
    const data = await apiGet(payload);
    appState = data;
    renderApp();
    showToast('Documentation status saved.');
  } catch (error) {
    console.error(error);
    showToast('Documentation failed to save.');
  }
}

async function generateSummary() {
  const monthKey = getValue('summaryMonth');

  if (!monthKey) {
    showToast('Select a month first.');
    return;
  }

  try {
    showToast('Generating monthly summary...');
    const result = await apiGet({ action: 'generateMonthlySummary', monthKey });
    appState = result.data;
    setValue('generatedSummary', result.summary);
    renderApp();
    showToast('Monthly summary generated.');
  } catch (error) {
    console.error(error);
    showToast('Summary generation failed.');
  }
}

async function toggleActionStatus(actionId, currentStatus) {
  const nextStatus = currentStatus === 'Completed' ? 'Open' : 'Completed';

  try {
    showToast('Updating action...');
    const data = await apiGet({
      action: 'updateActionStatus',
      actionId,
      status: nextStatus
    });

    appState = data;
    renderApp();
    showToast('Action updated.');
  } catch (error) {
    console.error(error);
    showToast('Action update failed.');
  }
}

function renderApp() {
  renderProgress();
  renderHealthSummary();
  renderActions();
  renderWorkLogs();
  renderAccuracyChecks();
  renderDocumentationNotes();
  renderNextMilestones();
  renderMonthlySummaries();
  loadDocEditor();
}

function renderProgress() {
  const progress = appState.progress || {};
  const percent = Number(progress.percent || 0);
  const completed = Number(progress.completed || 0);
  const total = Number(progress.total || 0);
  const open = Math.max(total - completed, 0);

  setText('topProgress', `${percent}%`);
  setText('topCompleted', `${completed} / ${total}`);
  setText('topCompletedMirror', `${completed} / ${total}`);
  setText('openItemsCount', `${open}`);
  setText('progressPercent', `${percent}%`);
  setText('progressPercentMirror', `${percent}%`);
  setText('progressText', `${completed} of ${total} completed`);

  const xpFill = document.querySelector('.xp-fill');
  if (xpFill) xpFill.style.width = `${Math.max(0, Math.min(100, percent))}%`;

  const insight = percent >= 85
    ? 'Final stretch. Lock documentation and verify the last open actions.'
    : percent >= 50
      ? 'Mission is moving. Focus next on validation, documentation, and open blockers.'
      : 'Foundation is active. Complete the highest-impact actions before adding more scope.';

  setText('progressInsight', insight);

  const ring = document.getElementById('progressRing');
  if (ring) {
    const circumference = 314;
    ring.style.strokeDashoffset = circumference - (circumference * percent / 100);
  }
}

function renderHealthSummary() {
  const docs = appState.docs || [];
  const checks = appState.accuracyChecks || [];
  const work = appState.workLog || [];

  const docsTotal = docs.length;
  const docsDone = docs.filter(row => String(row.Status || '').toLowerCase() === 'completed').length;
  const docsPct = docsTotal ? Math.round((docsDone / docsTotal) * 100) : 0;

  const checksTotal = checks.length;
  const passCount = checks.filter(row => String(row.Result || '').toLowerCase() === 'pass').length;
  const qualityPct = checksTotal ? Math.round((passCount / checksTotal) * 100) : 0;

  const automationCount = work.filter(row => String(row.WorkType || '').toLowerCase().includes('automation')).length;
  const automationPct = work.length ? Math.min(100, Math.round((automationCount / work.length) * 100)) : 0;

  setText('docsHealthText', `${docsPct}%`);
  setText('docsHealthTextMirror', `${docsPct}%`);
  setText('qualityHealthText', checksTotal ? `${qualityPct}%` : '0%');
  setText('automationHealthText', work.length ? `${automationPct}%` : '0%');
  setText('dataHealthText', 'Good');

  const docsFill = document.querySelector('.docs-fill');
  if (docsFill) docsFill.style.width = `${docsPct}%`;

  const autoFill = document.querySelector('.auto-fill');
  if (autoFill) autoFill.style.width = `${automationPct}%`;
}

function renderActions() {
  const box = document.getElementById('actionsList');
  if (!box) return;

  box.innerHTML = '';
  const actions = appState.actions || [];

  if (!actions.length) {
    box.innerHTML = '<div class="record-item">No actions loaded yet. Replace Code.gs with the API-only version, run setupGoalTracker, then redeploy.</div>';
    return;
  }

  actions.forEach(action => {
    const status = action.Status || 'Open';
    const completed = status === 'Completed';
    const item = document.createElement('div');
    item.className = 'action-item';

    item.innerHTML = `
      <div class="action-id">#${escapeHtml(action.ActionID)}</div>
      <div class="action-text">${escapeHtml(action.Action)}</div>
      <div class="hide-md"><span class="badge ${completed ? 'completed' : 'open'}">${completed ? 'Completed' : 'Open'}</span></div>
      <div class="hide-md action-meta">BLOPEZ</div>
      <div><span class="badge ${completed ? 'completed' : 'open'}">${escapeHtml(status)}</span></div>
      <div>
        <button class="btn ghost" onclick="toggleActionStatus('${escapeAttr(action.ActionID)}', '${escapeAttr(status)}')">
          ${completed ? 'Reopen' : 'Complete'}
        </button>
      </div>
    `;

    box.appendChild(item);
  });
}

function renderWorkLogs() {
  const rows = appState.workLog || [];
  renderRecordList('workLogList', rows, 'work');
  renderRecordList('recentWorkLogs', rows.slice(0, 3), 'work');
}

function renderAccuracyChecks() {
  const rows = appState.accuracyChecks || [];
  renderRecordList('accuracyList', rows, 'accuracy');
  renderRecordList('recentAccuracyChecks', rows.slice(0, 3), 'accuracy');
}

function renderDocumentationNotes() {
  const box = document.getElementById('documentationNotes');
  if (!box) return;

  const rows = (appState.docs || []).filter(row => String(row.Status || '').toLowerCase() !== 'completed').slice(0, 3);
  box.innerHTML = '';

  if (!rows.length) {
    box.innerHTML = '<div class="record-item">No open documentation notes.</div>';
    return;
  }

  rows.forEach(row => {
    const item = document.createElement('div');
    item.className = 'record-item';
    item.innerHTML = `
      <div class="record-title">${escapeHtml(row.Dashboard || 'Documentation')}</div>
      <div class="record-meta">${escapeHtml(row.Status || 'Not Started')} | API: ${escapeHtml(row.APIDependencies || 'Open')}</div>
      <div class="record-body">${escapeHtml(row.Notes || 'Documentation still needs review.')}</div>
    `;
    box.appendChild(item);
  });
}

function renderNextMilestones() {
  const box = document.getElementById('nextMilestones');
  if (!box) return;

  box.innerHTML = '';
  NEXT_MILESTONES.forEach(row => {
    const item = document.createElement('div');
    item.className = 'record-item';
    item.innerHTML = `
      <div class="record-title">${escapeHtml(row[0])}</div>
      <div class="record-meta">${escapeHtml(row[1])}</div>
    `;
    box.appendChild(item);
  });
}

function renderMonthlySummaries() {
  const box = document.getElementById('monthlyList');
  if (!box) return;

  box.innerHTML = '';
  const rows = appState.monthlySummaries || [];

  if (!rows.length) {
    box.innerHTML = '<div class="record-item">No monthly summaries yet.</div>';
    return;
  }

  rows.forEach(row => {
    const item = document.createElement('div');
    item.className = 'record-item';
    item.innerHTML = `
      <div class="record-title">${escapeHtml(row.MonthKey || 'Monthly Summary')}</div>
      <div class="record-meta">Work Logs: ${escapeHtml(row.WorkLogCount || 0)} | Accuracy Checks: ${escapeHtml(row.AccuracyChecksCount || 0)}</div>
      <div class="record-body">${escapeHtml(row.GeneratedSummary || '')}</div>
    `;
    box.appendChild(item);
  });
}

function renderRecordList(id, rows, type) {
  const box = document.getElementById(id);
  if (!box) return;

  box.innerHTML = '';

  if (!rows.length) {
    box.innerHTML = '<div class="record-item">No records yet.</div>';
    return;
  }

  rows.forEach(row => {
    const item = document.createElement('div');
    item.className = 'record-item';

    if (type === 'work') {
      item.innerHTML = `
        <div class="record-title">${escapeHtml(row.Dashboard || 'Dashboard Work')}</div>
        <div class="record-meta">${escapeHtml(row.WorkDate || '')} | ${escapeHtml(row.WorkType || '')} | ${escapeHtml(row.Status || '')}</div>
        <div class="record-body">${escapeHtml(row.Description || '')}</div>
        ${row.Evidence ? `<div class="record-body muted">Evidence: ${escapeHtml(row.Evidence)}</div>` : ''}
      `;
    }

    if (type === 'accuracy') {
      const resultClass = String(row.Result || '').toLowerCase() === 'pass' ? 'completed' : 'review';
      item.innerHTML = `
        <div class="record-title">${escapeHtml(row.Dashboard || 'Accuracy Check')}</div>
        <div class="record-meta">${escapeHtml(row.CheckDate || '')} | ${escapeHtml(row.MetricName || '')} | <span class="badge ${resultClass}">${escapeHtml(row.Result || 'Review')}</span></div>
        <div class="record-body">Source: ${escapeHtml(row.SourceTotal)} | Dashboard: ${escapeHtml(row.DashboardTotal)} | Variance: ${escapeHtml(row.Variance)}</div>
        ${row.Notes ? `<div class="record-body muted">${escapeHtml(row.Notes)}</div>` : ''}
      `;
    }

    box.appendChild(item);
  });
}

function loadDocEditor() {
  const dashboard = getValue('docDashboard');
  const docs = appState.docs || [];
  const row = docs.find(item => item.Dashboard === dashboard);
  if (!row) return;

  setValue('docDataSources', row.DataSources || 'Open');
  setValue('docFormulas', row.Formulas || 'Open');
  setValue('docRefreshTiming', row.RefreshTiming || 'Open');
  setValue('docSnapshotLogic', row.SnapshotLogic || 'Open');
  setValue('docApiDependencies', row.APIDependencies || 'Open');
  setValue('docTroubleshooting', row.Troubleshooting || 'Open');
  setValue('docStatus', row.Status || 'Not Started');
  setValue('docNotes', row.Notes || '');
}

function copySummary() {
  const text = getValue('generatedSummary');
  if (!text) {
    showToast('No summary to copy.');
    return;
  }

  navigator.clipboard.writeText(text)
    .then(() => showToast('Summary copied.'))
    .catch(() => showToast('Copy failed. Select and copy manually.'));
}

function apiGet(params) {
  return new Promise((resolve, reject) => {
    const callbackName = 'goalTrackerCallback_' + Date.now() + '_' + Math.floor(Math.random() * 100000);
    const url = new URL(API_URL);

    Object.keys(params).forEach(key => {
      url.searchParams.set(key, params[key]);
    });

    url.searchParams.set('callback', callbackName);

    const script = document.createElement('script');
    script.src = url.toString();

    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error('API request timed out.'));
    }, 20000);

    window[callbackName] = function (data) {
      cleanup();

      if (!data || data.ok === false) {
        reject(new Error(data && data.error ? data.error : 'API returned an error.'));
        return;
      }

      resolve(data);
    };

    script.onerror = function () {
      cleanup();
      reject(new Error('API script request failed.'));
    };

    function cleanup() {
      clearTimeout(timeout);

      if (script.parentNode) {
        script.parentNode.removeChild(script);
      }

      try {
        delete window[callbackName];
      } catch (err) {
        window[callbackName] = undefined;
      }
    }

    document.body.appendChild(script);
  });
}

function fillSelect(id, values) {
  const select = document.getElementById(id);
  if (!select) return;

  select.innerHTML = '';

  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

function getValue(id) {
  const element = document.getElementById(id);
  return element ? element.value : '';
}

function setValue(id, value) {
  const element = document.getElementById(id);
  if (element) element.value = value || '';
}

function setText(id, value) {
  const element = document.getElementById(id);
  if (element) element.textContent = value || '';
}

function showToast(message) {
  const toast = document.getElementById('toast');
  if (!toast) return;

  toast.textContent = message;
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
  }, 2800);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/'/g, '&#039;');
}
