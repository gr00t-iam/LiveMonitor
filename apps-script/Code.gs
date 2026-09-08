/**
 * Tech Monitor shared backend.
 * Bind this script to the "Live Dashboard" spreadsheet and deploy as a web app.
 */
const SPREADSHEET_ID = '1Snz0D4WBgKX9P2PT8M6lyCzPw9F3LDk6ao1l30PX6xQ';
const ARCHIVE_SPREADSHEET_ID = '16sdXMHsDKA979wli_IIwEiBE8xNeROtgsBAEwfCPJoE';
const TAB = Object.freeze({
  TECHNICIANS: 'Technicians',
  LOG: 'Activity Log',
  SETTINGS: 'Settings',
  STAFF: 'Command Staff'
});
const ADMIN_USERS_PROPERTY = 'ADMIN_USERS';
const PROTECTED_ADMIN_NAMES = Object.freeze(['Jeremy Miller', 'David Lowe']);
const PRESENCE_TIMEOUT_MS = 90 * 1000;
const SCHEDULE_API_URL = 'https://script.google.com/macros/s/AKfycbzQC6eWLLN_mRjO4jgUwCwDrMGe3Gy1fYzunf67gnCejfDfsd6nqgdt11wryl_ZNo-RCw/exec';
const SCHEDULE_CACHE_KEY = 'weekly-schedule-v1';
const SCHEDULE_SYNC_THROTTLE_KEY = 'weekly-schedule-sync-v1';
const SCHEDULE_CACHE_SECONDS = 120;

const TECH_HEADERS = Object.freeze([
  'Technician ID', 'Name', 'Active', 'Card Visible', 'Status',
  'Shift Start UTC', 'Update Due UTC', 'Break Start UTC', 'Shift Ended',
  'Active Issue', 'Last Resolution', 'Last Resolution UTC',
  'Updated At UTC', 'Updated By', 'Version', 'Last Update UTC',
  'Update Paused Milliseconds', 'Archived'
]);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tech Monitor · Command Suite')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialState(operator) {
  ensureTechnicianSchema_();
  ensureSettingsSchema_();
  recordPresence_(operator);
  synchronizeScheduledTechnicians_();
  return buildState_();
}

function getSharedState(operator) {
  recordPresence_(operator);
  synchronizeScheduledTechnicians_();
  return buildState_();
}

// ==========================================
// BACKGROUND TRIGGERS (NEW)
// ==========================================

/**
 * Run this function ONCE manually from the Apps Script editor 
 * to automatically create the scheduled background triggers.
 */
function setupTriggers() {
  // Clear any existing triggers to avoid duplicates
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(t => ScriptApp.deleteTrigger(t));

  // 1. Auto-Clock out every day between 5:00 AM and 6:00 AM
  ScriptApp.newTrigger('scheduledDailyReset')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();

  // 2. Archive old logs every Sunday around 2:00 AM
  ScriptApp.newTrigger('archiveOldLogs')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
}

function scheduledDailyReset() {
  return withWriteLock_(function () {
    const sheet = getSheet_(TAB.TECHNICIANS);
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    let recentUpdates = false;

    // Check if any active tech was updated in the last 4 hours
    for (let i = 1; i < values.length; i += 1) {
      if (!values[i][0]) continue;
      const tech = rowToTechnician_(values[i]);
      if (tech.updatedAt) {
        const updateTime = new Date(tech.updatedAt);
        if (now.getTime() - updateTime.getTime() < 4 * 60 * 60 * 1000) {
          recentUpdates = true;
          break; // Stop checking, someone is actively working
        }
      }
    }

    // Only run if no recent updates occurred
    if (!recentUpdates) {
      const nowIso = now.toISOString();
      for (let i = 1; i < values.length; i += 1) {
        if (!values[i][0]) continue;
        const tech = rowToTechnician_(values[i]);
        if (tech.archived || (!tech.active && tech.status === 'Site Complete')) continue;
        
        tech.status = 'Not Started';
        tech.shiftStart = '';
        tech.lastUpdate = '';
        tech.updatePausedMs = 0;
        tech.updateDue = '';
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.activeIssue = '';
        tech.updatedAt = nowIso;
        tech.updatedBy = 'System';
        tech.version += 1;
        values[i] = technicianToRow_(tech);
      }
      if (values.length > 1) sheet.getRange(2, 1, values.length - 1, TECH_HEADERS.length).setValues(values.slice(1));
      appendSystemLog_('Automated daily reset completed at 5:00 AM', 'reset', 'Triggered by inactivity', 'System');
      SpreadsheetApp.flush();
    }
  });
}

function archiveOldLogs() {
  return withWriteLock_(function () {
    const sheet = getSheet_(TAB.LOG);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Only headers

    const now = new Date();
    const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
    const rowsToArchive = [];
    const headers = data[0];
    
    // Logs append downward. Scan from top (oldest) downward.
    let firstKeepIndex = 1; 
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][1]); // Column B is timestampUtc
      if (now.getTime() - timestamp.getTime() > thirtyDaysMs) {
        rowsToArchive.push(data[i]);
      } else {
        firstKeepIndex = i; // Found the first row that is newer than 30 days
        break; 
      }
    }

    if (rowsToArchive.length > 0) {
      const archiveDb = SpreadsheetApp.openById(ARCHIVE_SPREADSHEET_ID);
      let archiveSheet = archiveDb.getSheetByName('Archive Log');
      
      if (!archiveSheet) {
        archiveSheet = archiveDb.insertSheet('Archive Log');
        archiveSheet.appendRow(headers);
      }
      
      // Append old rows to archive sheet
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
      
      // Delete rows from the live sheet (deleting from row 2 downwards)
      const numToDelete = firstKeepIndex - 1;
      if (numToDelete > 0) {
        sheet.deleteRows(2, numToDelete);
      }
      appendSystemLog_(`Archived ${rowsToArchive.length} entries older than 30 days`, 'archive', '', 'System');
    }
  });
}


// ==========================================
// CORE FUNCTIONS
// ==========================================

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
    normalizeUpdateClock_(tech, settings);
    const expectedVersion = Number(request.expectedVersion || 0);
    if (expectedVersion && expectedVersion !== tech.version) {
      return { ok: false, conflict: true, message: 'This technician changed on another screen.', state: buildState_() };
    }

    const now = new Date();
    const nowIso = now.toISOString();
    const actor = cleanText_(request.actor || 'Team member', 80);
    const action = String(request.action || '');
    if (action === 'setVisible' || (action === 'setActive' && !asBoolean_(request.value))) requireAdmin_();
    let logAction = '';
    let logType = action;
    let logNote = cleanText_(request.note || '', 1000);

    switch (action) {
      case 'beginShift':
        tech.shiftStart = nowIso;
        tech.lastUpdate = nowIso;
        tech.updatePausedMs = 0;
        tech.updateDue = addMinutes_(now, settings.updateMinutes).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = 'On Shift';
        logAction = 'Shift Started';
        logType = 'start';
        break;
      case 'confirmUpdate':
        if (tech.status === 'On Break') throw new Error('Return from lunch before recording an update.');
        if (!tech.shiftStart) tech.shiftStart = nowIso;
        tech.lastUpdate = nowIso;
        tech.updatePausedMs = 0;
        tech.updateDue = addMinutes_(now, settings.updateMinutes).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = tech.activeIssue ? 'Ticket Open' : 'On Shift';
        logAction = 'Log Updated';
        logType = 'update';
        break;
      case 'setLastUpdate': {
        const selectedUpdate = new Date(request.value);
        if (isNaN(selectedUpdate.getTime())) throw new Error('The last update time is invalid.');
        tech.lastUpdate = selectedUpdate.toISOString();
        // A new update begins a new two-hour cycle. Completed lunch time from
        // the previous cycle must not be added to the new deadline.
        tech.updatePausedMs = 0;
        const activePauseMs = tech.breakStart && tech.status === 'On Break'
          ? Math.max(0, now.getTime() - Math.max(new Date(tech.breakStart).getTime(), selectedUpdate.getTime()))
          : 0;
        tech.updateDue = new Date(
          addMinutes_(selectedUpdate, settings.updateMinutes).getTime() + tech.updatePausedMs + activePauseMs
        ).toISOString();
        logAction = 'Last update time changed';
        logType = 'last-update';
        logNote = selectedUpdate.toISOString();
        break;
      }
      case 'startBreak':
        tech.breakStart = nowIso;
        tech.status = 'On Break';
        logAction = 'Break started';
        logType = 'break-start';
        break;
      case 'endBreak':
        if (tech.breakStart) {
          const pauseStartMs = tech.lastUpdate
            ? Math.max(new Date(tech.breakStart).getTime(), new Date(tech.lastUpdate).getTime())
            : new Date(tech.breakStart).getTime();
          tech.updatePausedMs += Math.max(0, now.getTime() - pauseStartMs);
        }
        tech.breakStart = '';
        tech.updateDue = tech.lastUpdate
          ? new Date(addMinutes_(new Date(tech.lastUpdate), settings.updateMinutes).getTime() + tech.updatePausedMs).toISOString()
          : '';
        tech.status = tech.activeIssue ? 'Ticket Open' : 'On Shift';
        logAction = 'Break ended';
        logType = 'break-end';
        break;
      case 'endShift':
        tech.shiftStart = '';
        tech.lastUpdate = '';
        tech.updatePausedMs = 0;
        tech.updateDue = '';
        tech.breakStart = '';
        tech.shiftEnded = true;
        tech.status = 'Off Shift';
        logAction = 'Shift Ended';
        logType = 'end';
        break;
      case 'completeSite':
        tech.active = false;
        tech.cardVisible = false;
        tech.status = 'Site Complete';
        tech.shiftStart = '';
        tech.lastUpdate = '';
        tech.updatePausedMs = 0;
        tech.updateDue = '';
        tech.breakStart = '';
        tech.shiftEnded = true;
        tech.activeIssue = '';
        logAction = 'Site Completed';
        logType = 'site-complete';
        break;
      case 'setManualShift': {
        const manual = new Date(request.value);
        if (isNaN(manual.getTime())) throw new Error('The shift start time is invalid.');
        const alreadyStarted = Boolean(tech.shiftStart) && !tech.shiftEnded;
        tech.shiftStart = manual.toISOString();
        if (!alreadyStarted || !tech.lastUpdate) {
          tech.lastUpdate = manual.toISOString();
          tech.updatePausedMs = 0;
        }
        tech.updateDue = new Date(addMinutes_(new Date(tech.lastUpdate), settings.updateMinutes).getTime() + tech.updatePausedMs).toISOString();
        tech.breakStart = '';
        tech.shiftEnded = false;
        tech.status = 'On Shift';
        logAction = 'Manual shift start set';
        logType = 'update';
        logNote = manual.toISOString();
        break;
      }
      case 'setVisible':
        tech.cardVisible = asBoolean_(request.value);
        logAction = tech.cardVisible ? 'Technician card shown' : 'Technician card hidden';
        logType = 'display';
        break;
      case 'setActive':
        tech.active = asBoolean_(request.value);
        if (tech.active) {
          tech.cardVisible = true;
          if (tech.status === 'Site Complete') {
            tech.status = 'Not Started';
            tech.shiftEnded = false;
          }
        }
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

    if ((action === 'confirmUpdate' || action === 'setLastUpdate') && technicianIsOvertime_(tech, settings, now)) {
      logAction += ' · Overtime';
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
    const viewer = currentViewer_();
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
      match.archived = false;
      if (match.status === 'Site Complete') {
        match.status = 'Not Started';
        match.shiftEnded = false;
      }
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
      lastUpdate: '', updatePausedMs: 0, archived: false,
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
    tech.archived = true;
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
      break_minutes: [Number(values.breakMinutes), 15, 180],
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
      if (tech.archived || (!tech.active && tech.status === 'Site Complete')) continue;
      tech.status = 'Not Started';
      tech.shiftStart = '';
      tech.lastUpdate = '';
      tech.updatePausedMs = 0;
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
  const weeklySchedule = readWeeklySchedule_(settings.timezone);
  const assignmentsByTech = scheduleAssignmentsByTechnician_(weeklySchedule);
  const activityLog = readRecentActivity_(250);
  const notesByTech = {};
  activityLog.forEach(function (entry) {
    if (entry.type !== 'note' && !(entry.type === 'update' && entry.note)) return;
    if (!notesByTech[entry.techId]) notesByTech[entry.techId] = [];
    if (notesByTech[entry.techId].length < 20) {
      notesByTech[entry.techId].push({ time: entry.timestampUtc, text: entry.note, actor: entry.actor });
    }
  });
  technicians.forEach(function (tech) {
    normalizeUpdateClock_(tech, settings);
    tech.noteEntries = notesByTech[tech.id] || [];
    const assignment = assignmentsByTech[normalizePersonName_(tech.name)] || null;
    tech.scheduledSite = assignment ? assignment.siteNumber : '';
    tech.scheduledDate = assignment ? assignment.dateKey : '';
    tech.scheduleStatus = assignment ? assignment.status : '';
  });
  return {
    serverNowUtc: new Date().toISOString(),
    technicians: technicians,
    weeklySchedule: weeklySchedule,
    activityLog: activityLog,
    settings: settings,
    commandStaff: readStaff_(),
    viewer: currentViewer_()
  };
}


function readWeeklySchedule_(timezone) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SCHEDULE_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (error) {}
  }
  try {
    const response = UrlFetchApp.fetch(SCHEDULE_API_URL, {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true
    });
    const statusCode = response.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) throw new Error('Schedule service returned HTTP ' + statusCode);
    const payload = JSON.parse(response.getContentText());
    const rows = payload && Array.isArray(payload.schedule) ? payload.schedule : [];
    const bounds = currentWeekBounds_(timezone);
    const schedule = rows.map(normalizeScheduleRow_).filter(function (row) {
      return row.dateKey && row.dateKey >= bounds.monday && row.dateKey < bounds.nextMonday;
    }).sort(function (a, b) {
      return a.dateKey.localeCompare(b.dateKey) || a.technician.localeCompare(b.technician);
    });
    cache.put(SCHEDULE_CACHE_KEY, JSON.stringify(schedule), SCHEDULE_CACHE_SECONDS);
    return schedule;
  } catch (error) {
    console.error('Schedule sync failed: ' + error.message);
    cache.put(SCHEDULE_CACHE_KEY, '[]', 60);
    return [];
  }
}

function normalizeScheduleRow_(row) {
  row = row || {};
  return {
    dateKey: scheduleDateKey_(row.date),
    siteNumber: cleanText_(row.siteNumber, 40),
    technician: cleanText_(row.technician, 100),
    status: cleanText_(row.status || 'Scheduled', 60)
  };
}

function scheduleDateKey_(value) {
  if (!value) return '';
  const text = String(value).trim();
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + '-' + match[2] + '-' + match[3];
  match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return match[3] + '-' + String(match[1]).padStart(2, '0') + '-' + String(match[2]).padStart(2, '0');
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : Utilities.formatDate(parsed, 'America/Chicago', 'yyyy-MM-dd');
}

function currentWeekBounds_(timezone) {
  const today = Utilities.formatDate(new Date(), timezone || 'America/Chicago', 'yyyy-MM-dd');
  const parts = today.split('-').map(Number);
  const anchor = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12));
  const day = anchor.getUTCDay();
  const daysFromMonday = day === 0 ? 6 : day - 1;
  const mondayDate = new Date(anchor.getTime() - daysFromMonday * 86400000);
  const nextMondayDate = new Date(mondayDate.getTime() + 7 * 86400000);
  return {
    today: today,
    monday: Utilities.formatDate(mondayDate, 'UTC', 'yyyy-MM-dd'),
    nextMonday: Utilities.formatDate(nextMondayDate, 'UTC', 'yyyy-MM-dd')
  };
}

function normalizePersonName_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function scheduleAssignmentsByTechnician_(schedule) {
  const assignments = {};
  (schedule || []).forEach(function (row) {
    const key = normalizePersonName_(row.technician);
    if (!key) return;
    const current = assignments[key];
    if (!current || (isScheduleComplete_(current.status) && !isScheduleComplete_(row.status))) assignments[key] = row;
  });
  return assignments;
}

function isScheduleComplete_(status) {
  return /^(complete|completed|site complete|closed|done)$/i.test(String(status || '').trim());
}

function synchronizeScheduledTechnicians_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(SCHEDULE_SYNC_THROTTLE_KEY)) return;
  cache.put(SCHEDULE_SYNC_THROTTLE_KEY, '1', 45);
  const settings = readSettings_();
  const schedule = readWeeklySchedule_(settings.timezone);
  if (!schedule.length) return;
  const assignments = scheduleAssignmentsByTechnician_(schedule);
  const today = currentWeekBounds_(settings.timezone).today;

  return withWriteLock_(function () {
    const sheet = getSheet_(TAB.TECHNICIANS);
    const values = sheet.getDataRange().getValues();
    const properties = PropertiesService.getScriptProperties();
    let changed = false;
    for (let i = 1; i < values.length; i += 1) {
      if (!values[i][0]) continue;
      const tech = rowToTechnician_(values[i]);
      if (tech.archived) continue;
      const assignment = assignments[normalizePersonName_(tech.name)];
      if (!assignment || !assignment.dateKey || assignment.dateKey > today) continue;
      const assignmentKey = assignment.dateKey + '|' + assignment.siteNumber;
      const propertyKey = 'SCHEDULE_ACTIVATED_' + tech.id;
      const previousKey = properties.getProperty(propertyKey) || '';

      if (isScheduleComplete_(assignment.status)) {
        if (previousKey !== assignmentKey) properties.setProperty(propertyKey, assignmentKey);
        if (tech.status === 'Site Complete' && !tech.active && !tech.cardVisible) continue;
        tech.active = false;
        tech.cardVisible = false;
        tech.status = 'Site Complete';
        tech.shiftStart = '';
        tech.lastUpdate = '';
        tech.updatePausedMs = 0;
        tech.updateDue = '';
        tech.breakStart = '';
        tech.shiftEnded = true;
        tech.activeIssue = '';
        tech.updatedAt = new Date().toISOString();
        tech.updatedBy = 'Schedule Sync';
        tech.version += 1;
        values[i] = technicianToRow_(tech);
        appendLog_(tech, 'Site Completed from schedule', 'site-complete', 'Site ' + assignment.siteNumber, 'Schedule Sync', tech.updatedAt);
        changed = true;
        continue;
      }

      if (previousKey === assignmentKey) continue;
      properties.setProperty(propertyKey, assignmentKey);
      const shouldOpen = !tech.active || !tech.cardVisible || tech.status === 'Site Complete' || tech.status === 'Off Shift';
      if (!shouldOpen) continue;
      tech.active = true;
      tech.cardVisible = true;
      tech.archived = false;
      tech.status = 'Not Started';
      tech.shiftStart = '';
      tech.lastUpdate = '';
      tech.updatePausedMs = 0;
      tech.updateDue = '';
      tech.breakStart = '';
      tech.shiftEnded = false;
      tech.activeIssue = '';
      tech.updatedAt = new Date().toISOString();
      tech.updatedBy = 'Schedule Sync';
      tech.version += 1;
      values[i] = technicianToRow_(tech);
      appendLog_(tech, 'Card opened from weekly schedule', 'schedule', 'Site ' + assignment.siteNumber, 'Schedule Sync', tech.updatedAt);
      changed = true;
    }
    if (changed) {
      sheet.getRange(2, 1, values.length - 1, TECH_HEADERS.length).setValues(values.slice(1));
      SpreadsheetApp.flush();
    }
  });
}

function readSettings_() {
  const rows = getSheet_(TAB.SETTINGS).getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(function (row) { if (row[0] !== '') map[String(row[0])] = row[1]; });
  return {
    updateMinutes: numberSetting_(map.update_minutes, 120),
    shiftMinutes: numberSetting_(map.shift_minutes, 480),
    breakMinutes: numberSetting_(map.break_minutes, 60),
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
  const activeRows = values.slice(1).filter(function (row) { return row[0] && asBoolean_(row[2]); });
  const cache = CacheService.getScriptCache();
  const keys = activeRows.map(function (row) { return presenceKey_(row[0]); });
  const presence = keys.length ? cache.getAll(keys) : {};
  const now = Date.now();
  return activeRows.map(function (row) {
    const name = String(row[0]);
    const lastSeenUtc = presence[presenceKey_(name)] || '';
    const lastSeen = lastSeenUtc ? new Date(lastSeenUtc).getTime() : 0;
    const isProtectedAdmin = PROTECTED_ADMIN_NAMES.some(function (adminName) {
      return adminName.toLowerCase() === name.toLowerCase();
    });
    return {
      name: name,
      role: isProtectedAdmin ? 'Admin' : String(row[1] || 'Staff'),
      lastSeenUtc: lastSeenUtc,
      online: Boolean(lastSeen && now - lastSeen <= PRESENCE_TIMEOUT_MS)
    };
  });
}

function recordPresence_(operator) {
  operator = cleanText_(operator, 100);
  if (!operator) return;
  const sheet = getSheet_(TAB.STAFF);
  const rowNumber = findStaffRow_(sheet, operator);
  if (!rowNumber || !asBoolean_(sheet.getRange(rowNumber, 3).getValue())) return;
  CacheService.getScriptCache().put(presenceKey_(operator), new Date().toISOString(), 120);
}

function presenceKey_(name) {
  return 'presence:' + String(name || '').trim().toLowerCase();
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
    version: Number(row[14] || 0), lastUpdate: isoValue_(row[15]),
    updatePausedMs: Math.max(0, Number(row[16] || 0)), archived: asBoolean_(row[17])
  };
}

function technicianToRow_(tech) {
  return [
    tech.id, tech.name, Boolean(tech.active), Boolean(tech.cardVisible), tech.status,
    tech.shiftStart || '', tech.updateDue || '', tech.breakStart || '', Boolean(tech.shiftEnded),
    tech.activeIssue || '', tech.lastResolution || '', tech.lastResolutionUtc || '',
    tech.updatedAt || '', tech.updatedBy || '', Number(tech.version || 0), tech.lastUpdate || '',
    Math.max(0, Number(tech.updatePausedMs || 0)), Boolean(tech.archived)
  ];
}

function normalizeUpdateClock_(tech, settings) {
  if (!tech.lastUpdate && tech.updateDue) {
    const due = new Date(tech.updateDue);
    if (!isNaN(due.getTime())) tech.lastUpdate = addMinutes_(due, -settings.updateMinutes).toISOString();
  }
  if (!tech.lastUpdate && tech.shiftStart) tech.lastUpdate = tech.shiftStart;
  tech.updatePausedMs = Math.max(0, Number(tech.updatePausedMs || 0));
  const activePauseMs = tech.breakStart && tech.status === 'On Break'
    ? Math.max(0, Date.now() - Math.max(
        new Date(tech.breakStart).getTime(),
        tech.lastUpdate ? new Date(tech.lastUpdate).getTime() : 0
      ))
    : 0;
  tech.updateDue = tech.lastUpdate && !tech.shiftEnded
    ? new Date(addMinutes_(new Date(tech.lastUpdate), settings.updateMinutes).getTime() + tech.updatePausedMs + activePauseMs).toISOString()
    : '';
  return tech;
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
  const admins = adminDirectory_();
  return {
    email: email,
    name: admins[email] || 'Team member',
    isAdmin: Boolean(admins[email])
  };
}

function adminDirectory_() {
  const raw = PropertiesService.getScriptProperties().getProperty(ADMIN_USERS_PROPERTY) || '';
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    const admins = {};
    Object.keys(parsed || {}).forEach(function (email) {
      const normalizedEmail = String(email).trim().toLowerCase();
      const name = cleanText_(parsed[email], 100);
      if (normalizedEmail && name) admins[normalizedEmail] = name;
    });
    return admins;
  } catch (error) {
    return {};
  }
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

function ensureTechnicianSchema_() {
  const sheet = getSheet_(TAB.TECHNICIANS);
  for (let column = 1; column <= TECH_HEADERS.length; column += 1) {
    const current = String(sheet.getRange(1, column).getDisplayValue() || '').trim();
    if (current && current !== TECH_HEADERS[column - 1]) {
      throw new Error('The Technicians sheet has an unexpected heading in column ' + column + ': ' + current);
    }
    if (!current) sheet.getRange(1, column).setValue(TECH_HEADERS[column - 1]);
  }
}

function ensureSettingsSchema_() {
  const sheet = getSheet_(TAB.SETTINGS);
  const values = sheet.getDataRange().getDisplayValues();
  const hasBreakSetting = values.slice(1).some(function (row) { return String(row[0]).trim() === 'break_minutes'; });
  if (!hasBreakSetting) sheet.appendRow(['break_minutes', 60]);
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
function technicianIsOvertime_(tech, settings, at) {
  if (!tech.shiftStart || tech.shiftEnded) return false;
  const shiftStart = new Date(tech.shiftStart);
  if (isNaN(shiftStart.getTime())) return false;
  return at.getTime() >= addMinutes_(shiftStart, settings.shiftMinutes).getTime();
}

function cleanText_(value, maxLength) { return String(value == null ? '' : value).trim().slice(0, maxLength); }
function numberSetting_(value, fallback) { const n = Number(value); return Number.isFinite(n) ? n : fallback; }
function asBoolean_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function isoValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value.toISOString();
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? String(value) : parsed.toISOString();
}
