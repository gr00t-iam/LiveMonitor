/**
 * Tech Monitor shared backend.
 * Bind this script to the "Live Dashboard" spreadsheet and deploy as a web app.
 */
const SPREADSHEET_ID = '1Snz0D4WBgKX9P2PT8M6lyCzPw9F3LDk6ao1l30PX6xQ';
const TAB = Object.freeze({
  TECHNICIANS: 'Technicians',
  LOG: 'Activity Log',
  SETTINGS: 'Settings',
  STAFF: 'Command Staff'
});
const ADMIN_EMAILS = Object.freeze([
  'jmiller@rr-solutions.us',
  'dlowe@rr-solutions.us'
]);
const PROTECTED_ADMIN_NAMES = Object.freeze(['Jeremy Miller', 'David Lowe']);

const TECH_HEADERS = Object.freeze([
  'Technician ID', 'Name', 'Active', 'Card Visible', 'Status',
  'Shift Start UTC', 'Update Due UTC', 'Break Start UTC', 'Shift Ended',
  'Active Issue', 'Last Resolution', 'Last Resolution UTC',
  'Updated At UTC', 'Updated By', 'Version'
]);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tech Monitor · Command Suite')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialState() {
  return buildState_();
}

function getSharedState() {
  return buildState_();
}

function mutateTechnician(request) {
  request = request || {};
  return withWriteLock_(function () {
    const sheet = getSheet_(TAB.TECHNICIANS);
    const settings = readSettings_();
    const rowNumber = findTechnicianRow_(sheet, cleanText_(request.techId, 100));
    if (!rowNumber) throw new Error('Technician was not found. Refresh and try again.');

    const range = sheet.getRange(rowNumber, 1, 1, TECH_HEADERS.length);
    const values = range.getValues()[0];
    const tech = rowToTechnician_(values);
    const expectedVersion = Number(request.expectedVersion || 0);
    if (expectedVersion && expectedVersion !== tech.version) {
      return { ok: false, conflict: true, message: 'This technician changed on another screen.', state: buildState_() };
    }

    const now = new Date();
    const nowIso = now.toISOString();
    const actor = cleanText_(request.actor || 'Team member', 80);
    const action = String(request.action || '');
    if (action === 'setActive' || action === 'setVisible') requireAdmin_();
    let logAction = '';
    let logType = action;
    let logNote = cleanText_(request.note || '', 1000);

    switch (action) {
      case 'beginShift':
        tech.shiftStart = nowIso;
        tech.updateDue = addMinutes_(now, settings.updateMinutes).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = 'On Shift';
        logAction = 'Shift Started';
        logType = 'start';
        break;
      case 'confirmUpdate':
        if (!tech.shiftStart) tech.shiftStart = nowIso;
        tech.updateDue = addMinutes_(now, settings.updateMinutes).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = tech.activeIssue ? 'Ticket Open' : 'On Shift';
        logAction = 'Log Updated';
        logType = 'update';
        break;
      case 'startBreak':
        tech.breakStart = nowIso;
        tech.status = 'On Break';
        logAction = 'Break started';
        logType = 'break-start';
        break;
      case 'endBreak':
        tech.breakStart = '';
        tech.updateDue = addMinutes_(now, settings.updateMinutes).toISOString();
        tech.status = tech.activeIssue ? 'Ticket Open' : 'On Shift';
        logAction = 'Break ended';
        logType = 'break-end';
        break;
      case 'endShift':
        tech.shiftStart = '';
        tech.updateDue = '';
        tech.breakStart = '';
        tech.shiftEnded = true;
        tech.status = 'Off Shift';
        logAction = 'Shift Ended';
        logType = 'end';
        break;
      case 'setManualShift': {
        const manual = new Date(request.value);
        if (isNaN(manual.getTime())) throw new Error('The shift start time is invalid.');
        tech.shiftStart = manual.toISOString();
        tech.updateDue = addMinutes_(manual, settings.updateMinutes).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = 'On Shift';
        logAction = 'Manual shift start set';
        logType = 'update';
        logNote = manual.toISOString();
        break;
      }
      case 'setVisible':
        tech.cardVisible = Boolean(request.value);
        logAction = tech.cardVisible ? 'Technician card shown' : 'Technician card hidden';
        logType = 'display';
        break;
      case 'setActive':
        tech.active = Boolean(request.value);
        if (tech.active) tech.cardVisible = true;
        logAction = tech.active ? 'Technician activated' : 'Technician deactivated';
        logType = 'roster';
        break;
      case 'createIssue':
        if (!logNote) throw new Error('Enter an issue description.');
        tech.activeIssue = logNote;
        tech.status = 'Ticket Open';
        logAction = 'Ticket created';
        logType = 'issue';
        break;
      case 'resolveIssue':
        if (!tech.activeIssue) throw new Error('This ticket is already resolved.');
        logNote = tech.activeIssue + (logNote ? ' · ' + logNote : '');
        tech.lastResolution = cleanText_(request.note || 'No details', 1000);
        tech.lastResolutionUtc = nowIso;
        tech.activeIssue = '';
        tech.status = tech.breakStart ? 'On Break' : (tech.shiftStart ? 'On Shift' : (tech.shiftEnded ? 'Off Shift' : 'Not Started'));
        logAction = 'Ticket resolved';
        logType = 'resolve';
        break;
      case 'addNote':
        if (!logNote) throw new Error('Enter a note.');
        logAction = 'Note added';
        logType = 'note';
        break;
      default:
        throw new Error('Unsupported technician action.');
    }

    tech.updatedAt = nowIso;
    tech.updatedBy = actor;
    tech.version += 1;
    range.setValues([technicianToRow_(tech)]);
    appendLog_(tech, logAction, logType, logNote, actor, nowIso);
    SpreadsheetApp.flush();
    return { ok: true, technician: tech, state: buildState_() };
  });
}

function addTechnician(name, actor) {
  return withWriteLock_(function () {
    const viewer = requireAdmin_();
    name = cleanText_(name, 100);
    actor = viewer.name;
    if (!name) throw new Error('Enter a technician name.');
    const sheet = getSheet_(TAB.TECHNICIANS);
    const existing = readTechnicians_();
    const match = existing.find(function (t) { return t.name.toLowerCase() === name.toLowerCase(); });
    if (match && match.active) {
      throw new Error('That technician is already on the roster.');
    }
    if (match) {
      const rowNumber = findTechnicianRow_(sheet, match.id);
      match.active = true;
      match.cardVisible = true;
      match.updatedAt = new Date().toISOString();
      match.updatedBy = actor;
      match.version += 1;
      sheet.getRange(rowNumber, 1, 1, TECH_HEADERS.length).setValues([technicianToRow_(match)]);
      appendLog_(match, 'Restored to roster', 'roster', '', actor, match.updatedAt);
      return { ok: true, state: buildState_() };
    }
    const id = 'tech-' + Utilities.getUuid();
    const tech = {
      id: id, name: name, active: true, cardVisible: true, status: 'Not Started',
      shiftStart: '', updateDue: '', breakStart: '', shiftEnded: false,
      activeIssue: '', lastResolution: '', lastResolutionUtc: '',
      updatedAt: new Date().toISOString(), updatedBy: actor, version: 1
    };
    sheet.appendRow(technicianToRow_(tech));
    appendLog_(tech, 'Added to roster', 'roster', '', actor, tech.updatedAt);
    return { ok: true, state: buildState_() };
  });
}

function removeTechnician(techId) {
  return withWriteLock_(function () {
    const viewer = requireAdmin_();
    const sheet = getSheet_(TAB.TECHNICIANS);
    const rowNumber = findTechnicianRow_(sheet, cleanText_(techId, 100));
    if (!rowNumber) throw new Error('Technician was not found. Refresh and try again.');
    const range = sheet.getRange(rowNumber, 1, 1, TECH_HEADERS.length);
    const tech = rowToTechnician_(range.getValues()[0]);
    tech.active = false;
    tech.cardVisible = false;
    tech.updatedAt = new Date().toISOString();
    tech.updatedBy = viewer.name;
    tech.version += 1;
    range.setValues([technicianToRow_(tech)]);
    appendLog_(tech, 'Removed from roster', 'roster', '', viewer.name, tech.updatedAt);
    return { ok: true, state: buildState_() };
  });
}

function addCommandStaff(name, role) {
  return withWriteLock_(function () {
    requireAdmin_();
    name = cleanText_(name, 100);
    role = role === 'Admin' ? 'Admin' : 'Staff';
    if (!name) throw new Error('Enter a staff name.');
    const sheet = getSheet_(TAB.STAFF);
    const rowNumber = findStaffRow_(sheet, name);
    if (rowNumber && asBoolean_(sheet.getRange(rowNumber, 3).getValue())) {
      throw new Error('That person is already listed.');
    }
    if (rowNumber) {
      sheet.getRange(rowNumber, 2, 1, 2).setValues([[role, true]]);
      return { ok: true, state: buildState_() };
    }
    sheet.appendRow([name, role, true]);
    return { ok: true, state: buildState_() };
  });
}

function removeCommandStaff(name) {
  return withWriteLock_(function () {
    requireAdmin_();
    name = cleanText_(name, 100);
    if (PROTECTED_ADMIN_NAMES.some(function (adminName) { return adminName.toLowerCase() === name.toLowerCase(); })) {
      throw new Error('Jeremy Miller and David Lowe are protected administrators and cannot be removed here.');
    }
    const sheet = getSheet_(TAB.STAFF);
    const rowNumber = findStaffRow_(sheet, name);
    if (!rowNumber) throw new Error('Command staff member was not found. Refresh and try again.');
    sheet.getRange(rowNumber, 3).setValue(false);
    appendSystemLog_('Command staff removed', 'staff', name, currentViewer_().name);
    return { ok: true, state: buildState_() };
  });
}

function updateSettings(values, actor) {
  return withWriteLock_(function () {
    const viewer = requireAdmin_();
    values = values || {};
    actor = viewer.name;
    const rules = {
      update_minutes: [Number(values.updateMinutes), 5, 1440],
      shift_minutes: [Number(values.shiftMinutes), 30, 1440],
      warning_minutes: [Number(values.warningMinutes), 1, 180],
      critical_minutes: [Number(values.criticalMinutes), 1, 120],
      polling_seconds: [Number(values.pollingSeconds), 5, 300]
    };
    Object.keys(rules).forEach(function (key) {
      const rule = rules[key];
      if (!Number.isFinite(rule[0]) || rule[0] < rule[1] || rule[0] > rule[2]) {
        throw new Error('One or more timer settings are outside the allowed range.');
      }
    });
    if (rules.critical_minutes[0] > rules.warning_minutes[0]) {
      throw new Error('Critical minutes cannot exceed warning minutes.');
    }
    const sheet = getSheet_(TAB.SETTINGS);
    const data = sheet.getDataRange().getValues();
    Object.keys(rules).forEach(function (key) {
      const rowIndex = data.findIndex(function (row, i) { return i > 0 && String(row[0]) === key; });
      if (rowIndex < 0) throw new Error('Missing setting: ' + key);
      sheet.getRange(rowIndex + 1, 2).setValue(rules[key][0]);
    });
    appendSystemLog_('Timer settings updated', 'settings', JSON.stringify(values), actor);
    SpreadsheetApp.flush();
    return { ok: true, state: buildState_() };
  });
}

function runDailyReset(actor) {
  return withWriteLock_(function () {
    actor = requireAdmin_().name;
    const sheet = getSheet_(TAB.TECHNICIANS);
    const values = sheet.getDataRange().getValues();
    const nowIso = new Date().toISOString();
    for (let i = 1; i < values.length; i += 1) {
      if (!values[i][0]) continue;
      const tech = rowToTechnician_(values[i]);
      tech.status = 'Not Started';
      tech.shiftStart = '';
      tech.updateDue = '';
      tech.breakStart = '';
      tech.shiftEnded = false;
      tech.activeIssue = '';
      tech.updatedAt = nowIso;
      tech.updatedBy = actor;
      tech.version += 1;
      values[i] = technicianToRow_(tech);
    }
    if (values.length > 1) sheet.getRange(2, 1, values.length - 1, TECH_HEADERS.length).setValues(values.slice(1));
    appendSystemLog_('Daily reset completed', 'reset', '', actor);
    SpreadsheetApp.flush();
    return { ok: true, state: buildState_() };
  });
}

function buildState_() {
  const settings = readSettings_();
  const technicians = readTechnicians_();
  const activityLog = readRecentActivity_(250);
  const notesByTech = {};
  activityLog.forEach(function (entry) {
    if (entry.type !== 'note' && !(entry.type === 'update' && entry.note)) return;
    if (!notesByTech[entry.techId]) notesByTech[entry.techId] = [];
    if (notesByTech[entry.techId].length < 20) {
      notesByTech[entry.techId].push({ time: entry.timestampUtc, text: entry.note, actor: entry.actor });
    }
  });
  technicians.forEach(function (tech) { tech.noteEntries = notesByTech[tech.id] || []; });
  return {
    serverNowUtc: new Date().toISOString(),
    technicians: technicians,
    activityLog: activityLog,
    settings: settings,
    commandStaff: readStaff_(),
    viewer: currentViewer_()
  };
}

function readSettings_() {
  const rows = getSheet_(TAB.SETTINGS).getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(function (row) { if (row[0] !== '') map[String(row[0])] = row[1]; });
  return {
    updateMinutes: numberSetting_(map.update_minutes, 120),
    shiftMinutes: numberSetting_(map.shift_minutes, 480),
    warningMinutes: numberSetting_(map.warning_minutes, 15),
    criticalMinutes: numberSetting_(map.critical_minutes, 5),
    pollingSeconds: numberSetting_(map.polling_seconds, 10),
    timezone: String(map.timezone || 'America/Chicago'),
    dashboardTitle: String(map.dashboard_title || 'Tech Monitor · Command Suite')
  };
}

function readTechnicians_() {
  const values = getSheet_(TAB.TECHNICIANS).getDataRange().getValues();
  return values.slice(1).filter(function (row) { return row[0]; }).map(rowToTechnician_);
}

function readStaff_() {
  const values = getSheet_(TAB.STAFF).getDataRange().getValues();
  return values.slice(1).filter(function (row) { return row[0] && asBoolean_(row[2]); }).map(function (row) {
    const name = String(row[0]);
    const isProtectedAdmin = PROTECTED_ADMIN_NAMES.some(function (adminName) {
      return adminName.toLowerCase() === name.toLowerCase();
    });
    return { name: name, role: isProtectedAdmin ? 'Admin' : String(row[1] || 'Staff') };
  });
}

function readRecentActivity_(limit) {
  const sheet = getSheet_(TAB.LOG);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const count = Math.min(limit, lastRow - 1);
  const rows = sheet.getRange(lastRow - count + 1, 1, count, 8).getValues();
  return rows.reverse().map(function (row) {
    return {
      eventId: String(row[0] || ''), timestampUtc: isoValue_(row[1]),
      techId: String(row[2] || ''), tech: String(row[3] || 'System'),
      message: String(row[4] || ''), type: String(row[5] || ''),
      note: String(row[6] || ''), actor: String(row[7] || '')
    };
  });
}

function rowToTechnician_(row) {
  return {
    id: String(row[0] || ''), name: String(row[1] || ''),
    active: asBoolean_(row[2]), cardVisible: asBoolean_(row[3]),
    status: String(row[4] || 'Not Started'), shiftStart: isoValue_(row[5]),
    updateDue: isoValue_(row[6]), breakStart: isoValue_(row[7]),
    shiftEnded: asBoolean_(row[8]), activeIssue: String(row[9] || ''),
    lastResolution: String(row[10] || ''), lastResolutionUtc: isoValue_(row[11]),
    updatedAt: isoValue_(row[12]), updatedBy: String(row[13] || ''),
    version: Number(row[14] || 0)
  };
}

function technicianToRow_(tech) {
  return [
    tech.id, tech.name, Boolean(tech.active), Boolean(tech.cardVisible), tech.status,
    tech.shiftStart || '', tech.updateDue || '', tech.breakStart || '', Boolean(tech.shiftEnded),
    tech.activeIssue || '', tech.lastResolution || '', tech.lastResolutionUtc || '',
    tech.updatedAt || '', tech.updatedBy || '', Number(tech.version || 0)
  ];
}

function appendLog_(tech, action, type, note, actor, timestampUtc) {
  getSheet_(TAB.LOG).appendRow([
    Utilities.getUuid(), timestampUtc, tech.id, tech.name, action, type, note || '', actor
  ]);
}

function appendSystemLog_(action, type, note, actor) {
  getSheet_(TAB.LOG).appendRow([
    Utilities.getUuid(), new Date().toISOString(), '', 'System', action, type, note || '', actor
  ]);
}

function findTechnicianRow_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return 0;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i += 1) if (ids[i][0] === id) return i + 2;
  return 0;
}

function findStaffRow_(sheet, name) {
  if (!name || sheet.getLastRow() < 2) return 0;
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const target = String(name).trim().toLowerCase();
  for (let i = 0; i < names.length; i += 1) {
    if (String(names[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return 0;
}

function currentViewer_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const names = {
    'jmiller@rr-solutions.us': 'Jeremy Miller',
    'dlowe@rr-solutions.us': 'David Lowe'
  };
  return {
    email: email,
    name: names[email] || 'Team member',
    isAdmin: ADMIN_EMAILS.indexOf(email) !== -1
  };
}

function requireAdmin_() {
  const viewer = currentViewer_();
  if (!viewer.isAdmin) {
    throw new Error('Administrator access is required. Sign in as Jeremy Miller or David Lowe.');
  }
  return viewer;
}

function getSheet_(name) {
  const sheet = getDatabase_().getSheetByName(name);
  if (!sheet) throw new Error('Required sheet is missing: ' + name);
  return sheet;
}

let DATABASE_;
function getDatabase_() {
  if (!DATABASE_) DATABASE_ = SpreadsheetApp.openById(SPREADSHEET_ID);
  return DATABASE_;
}

function withWriteLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return callback(); } finally { lock.releaseLock(); }
}

function addMinutes_(date, minutes) { return new Date(date.getTime() + Number(minutes) * 60000); }
function cleanText_(value, maxLength) { return String(value == null ? '' : value).trim().slice(0, maxLength); }
function numberSetting_(value, fallback) { const n = Number(value); return Number.isFinite(n) ? n : fallback; }
function asBoolean_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function isoValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value.toISOString();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? String(value) : parsed.toISOString();
}
