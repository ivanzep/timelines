// ============================================================
//  LA COSTA HOTEL — TIMELINE SYNC  |  Google Apps Script
//
//  VERSION HISTORY
//  ───────────────
//  V1.01  2026-05-18
//    • Added GANTT SETTINGS-DO NOT EDIT tab support
//    • doGet now returns saved chart settings alongside tasks
//    • doPost saves chart settings independently of task save —
//      settings write always runs even if task save errors
//    • readSettings() / writeSettings() helpers added
//    • No changes to importFromTaskList or saveBackToTaskList
//
//  V1.0   2026-05-18  (baseline — working release)
//    • Reads PROJECT TASK LIST; only rows with SCHEDULE or
//      MILESTONE checked appear on the Gantt
//    • Status text mapped to bar colours via STATUS_COLORS
//    • saveBackToTaskList updates dates + status; appends new tasks
//
//  ───────────────
//  SETUP (first time):
//  1. Open your Google Sheet
//     → https://docs.google.com/spreadsheets/d/1HShZAkZ7oV4_yDdRRbAeBG2p5Uh5EvH4Chs_64ZZ5WU
//  2. Extensions → Apps Script
//  3. Delete existing code, paste this entire file, Save
//  4. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  5. Copy the Web App URL → paste into timeline ⚙ Setup
//
//  UPDATING (already deployed):
//  Deploy → Manage deployments → pencil icon → New version → Deploy
//  (the URL stays the same)
// ============================================================

var SOURCE_SHEET   = 'PROJECT TASK LIST';          // your existing tab — never reformatted
var SETTINGS_SHEET = 'GANTT SETTINGS-DO NOT EDIT'; // created automatically on first Save

// ---- Status text → Gantt bar colour ----
var STATUS_COLORS = {
  'IN PROGRESS':        '#16a34a',   // green
  'UPCOMING':           '#f59e0b',   // amber
  'DOWNSTREAM':         '#8b5cf6',   // purple
  'PENDING':            '#d9c34a',   // yellow
  'COMPLETED':          '#94a3b8',   // grey
  'CANCELLED':          '#dc2626',   // red
  '75% COMPLETE':       '#22c55e',   // light green
  'WAITING ON OTHERS':  '#f97316',   // orange
  'WITING ON THE CITY': '#f97316',
  'UNDER REVIEW':       '#06b6d4',   // cyan
  'URGENT':             '#dc2626',
  'ON HOLD':            '#94a3b8',
  'STAND BY':           '#94a3b8',
  'NEEDS ATTENTION':    '#f97316',
  'TBD':                '#94a3b8',
  'NOTE':               '#60a5fa',   // light blue
  'RECEIVED':           '#22c55e',
  'APPROVED':           '#16a34a',
  'DECLINED':           '#dc2626'
};

// ============================================================
//  GET  →  Load tasks from PROJECT TASK LIST into the Gantt
// ============================================================
function doGet(e) {
  try {
    var result = importFromTaskList();
    result.settings = readSettings();   // null if tab doesn't exist yet
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
//  POST  →  Push date + status changes back to the sheet
// ============================================================
function doPost(e) {
  var taskMsg      = '';
  var taskErr      = '';
  var settingsErr  = '';

  try {
    var payload = JSON.parse(e.postData.contents);

    // ── 1. Save tasks (independent — errors here do NOT block settings) ──
    try {
      taskMsg = saveBackToTaskList(payload);
    } catch (tErr) {
      taskErr = tErr.toString();
    }

    // ── 2. Save chart settings (always runs, even if task save failed) ──
    if (payload.settings) {
      try {
        writeSettings(payload.settings);
      } catch (sErr) {
        settingsErr = sErr.toString();
      }
    }

    return buildResponse({
      success:     true,
      message:     taskMsg,
      taskError:   taskErr,
      settingsError: settingsErr
    });

  } catch (err) {
    // Only fires if JSON.parse itself fails
    return buildResponse({ success: false, error: err.toString() });
  }
}

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  IMPORT  — read PROJECT TASK LIST, return tasks array
// ============================================================
function importFromTaskList() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"' + SOURCE_SHEET + '" tab not found. Check the sheet name.');

  var raw = sheet.getDataRange().getValues();

  // --- Find the column-header row (the one with "DISCIPLINE", "TASK", etc.) ---
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') {
        hRow = i;
        break;
      }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) throw new Error('Could not find the DISCIPLINE header row in "' + SOURCE_SHEET + '".');

  // Map column name → zero-based index
  var cols = {};
  raw[hRow].forEach(function(h, idx) {
    cols[String(h).trim().toUpperCase()] = idx;
  });

  // Column indices (fall back to known positions if header missing)
  var CI = {
    discipline: ci(cols, 'DISCIPLINE',  1),
    id:         ci(cols, 'ID',          2),
    task:       ci(cols, 'TASK',        3),
    consultant: ci(cols, 'CONSULTANT',  4),
    person:     ci(cols, 'PERSON',      5),
    start:      ci(cols, 'START DATE',  7),
    end:        ci(cols, 'END DATE',    8),
    schedule:   ci(cols, 'SCHEDULE',   10),
    milestone:  ci(cols, 'MILESTONE',  11),
    status:     ci(cols, 'STATUS',     12),
    priority:   ci(cols, 'PRIORITY',   13),
    notes:      ci(cols, 'NOTES',      20)
  };

  var tasks      = [];
  var idCounter  = 1;
  var lastDisc   = '';

  for (var row = hRow + 1; row < raw.length; row++) {
    var r = raw[row];

    // Track current discipline (rows inherit the last non-empty value)
    var disc = String(r[CI.discipline] || '').trim();
    if (disc) lastDisc = disc;

    var taskName = String(r[CI.task] || '').trim();

    // Skip blanks, separators and section-header rows
    // (catches "-", "- -", "--", " - ", "---", etc.)
    if (!taskName || /^[\s\-]+$/.test(taskName)) continue;

    // Only include rows where BOTH Schedule and Milestone are checked (TRUE)
    var isScheduled = (r[CI.schedule] === true || String(r[CI.schedule]).toUpperCase() === 'TRUE');
    var isMilestone = (r[CI.milestone] === true || String(r[CI.milestone]).toUpperCase() === 'TRUE');
    if (!isScheduled && !isMilestone) continue;

    // Must have both a start and end date
    var startDate = parseSheetDate(r[CI.start]);
    var endDate   = parseSheetDate(r[CI.end]);
    if (!startDate || !endDate) continue;
    if (endDate.getTime() <= startDate.getTime()) continue;

    var status     = String(r[CI.status]     || '').trim();
    var consultant = String(r[CI.consultant] || '').trim();
    var person     = String(r[CI.person]     || '').trim();
    var notes      = String(r[CI.notes]      || '').trim();
    var priority   = String(r[CI.priority]   || '').trim();

    // Gantt row = the discipline code itself (direct match to spreadsheet)
    var groupKey = lastDisc || disc || 'UNCATEGORIZED';

    // Build the note shown on the bar: status + optional person
    var barNote = status;
    if (person && person.length < 20) barNote = status + (status ? ' · ' : '') + person;

    tasks.push({
      id:     idCounter++,
      name:   taskName,
      group:  groupKey,           // discipline code — exact match to spreadsheet DISCIPLINE column
      type:   'bar',
      start:  fmtDate(startDate),
      end:    fmtDate(endDate),
      color:  statusColor(status),
      note:   barNote,
      dashed: 'false',
      symbol: ''
    });
  }

  return {
    success: true,
    tasks:   tasks,
    meta: {
      name:     'LA COSTA HOTEL',
      subtitle: 'PROJECT SCHEDULE',
      updated:  fmtDate(new Date()),
      note:     'Live from Project Task List · ' + tasks.length + ' scheduled tasks'
    }
  };
}

// ============================================================
//  SAVE BACK  — update existing rows OR append new ones
//
//  • Existing tasks  → matched by DISCIPLINE|TASKNAME, updates
//                       START DATE, END DATE, STATUS only.
//  • New tasks       → not found in the sheet, appended as new
//                       rows at the bottom with DISCIPLINE, TASK,
//                       START DATE, END DATE, STATUS filled in.
//  All other columns are left completely unchanged.
// ============================================================
function saveBackToTaskList(payload) {
  if (!payload.tasks || !payload.tasks.length) return 'Nothing to save.';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"' + SOURCE_SHEET + '" tab not found.');

  var raw  = sheet.getDataRange().getValues();

  // Locate header row
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow = i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) throw new Error('Header row not found.');

  var cols = {};
  raw[hRow].forEach(function(h, idx) { cols[String(h).trim().toUpperCase()] = idx; });

  var CI = {
    discipline: ci(cols, 'DISCIPLINE',  1),
    task:       ci(cols, 'TASK',        3),
    start:      ci(cols, 'START DATE',  7),
    end:        ci(cols, 'END DATE',    8),
    schedule:   ci(cols, 'SCHEDULE',   10),
    milestone:  ci(cols, 'MILESTONE',  11),
    status:     ci(cols, 'STATUS',     12)
  };

  // Build lookup: "DISCIPLINE|TASKNAME" → 1-based sheet row
  // Also track the last sheet row that belongs to each discipline group
  var lookup      = {};
  var discLastRow = {};   // DISCIPLINE_KEY → last 1-based sheet row in that group
  var lastDisc    = '';
  for (var row = hRow + 1; row < raw.length; row++) {
    var disc = String(raw[row][CI.discipline] || '').trim();
    if (disc) lastDisc = disc;
    var task = String(raw[row][CI.task] || '').trim();
    if (task && task !== '-' && task !== '- -') {
      var key = normKey(lastDisc + '|' + task);
      if (!lookup[key]) lookup[key] = row + 1;            // Sheets rows are 1-based
      discLastRow[normKey(lastDisc)] = row + 1;           // keep updating → ends at last task row
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var updated  = 0;
  var appended = 0;

  // Collect new rows grouped by discipline so we can insert them together
  var newRowsByDisc    = {};   // disc → [ rowArray, ... ]
  var newRowsOrphaned  = [];   // discipline not found in sheet → bottom of sheet

  try {
    payload.tasks.forEach(function(t) {
      var discipline = (t.group || '').trim().toUpperCase();
      var taskName   = (t.name || '').trim();
      var key        = normKey(discipline + '|' + taskName);
      var sheetRow   = lookup[key];

      if (sheetRow) {
        // ---- UPDATE existing row ----
        if (t.start) sheet.getRange(sheetRow, CI.start + 1).setValue(t.start);
        if (t.end)   sheet.getRange(sheetRow, CI.end   + 1).setValue(t.end);
        if (t.note) {
          var statusCandidate = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[statusCandidate]) {
            sheet.getRange(sheetRow, CI.status + 1).setValue(statusCandidate);
          }
        }
        // Ensure SCHEDULE is checked; check MILESTONE only for milestone-type tasks
        sheet.getRange(sheetRow, CI.schedule  + 1).setValue(true);
        if (t.type === 'milestone') sheet.getRange(sheetRow, CI.milestone + 1).setValue(true);
        updated++;

      } else if (taskName) {
        // ---- NEW row — queue under its discipline ----
        var totalCols = raw[hRow].length;
        var newRow = new Array(totalCols).fill('');
        newRow[CI.discipline] = discipline;
        newRow[CI.task]       = taskName;
        newRow[CI.start]      = t.start || '';
        newRow[CI.end]        = t.end   || '';
        newRow[CI.schedule]   = true;                            // always checked
        if (t.type === 'milestone') newRow[CI.milestone] = true; // milestone flag
        if (t.note) {
          var sc = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[sc]) newRow[CI.status] = sc;
        }

        var discKey = normKey(discipline);
        if (discLastRow[discKey]) {
          if (!newRowsByDisc[discKey]) newRowsByDisc[discKey] = [];
          newRowsByDisc[discKey].push(newRow);
        } else {
          // Discipline not yet in the sheet — fall back to appending at bottom
          newRowsOrphaned.push(newRow);
        }
        appended++;
      }
    });

    // Insert new rows at the end of each discipline section.
    // Process disciplines sorted by their last-row position DESCENDING so that
    // inserting rows lower in the sheet doesn't shift the indices we still need.
    var discKeys = Object.keys(newRowsByDisc);
    discKeys.sort(function(a, b) { return (discLastRow[b] || 0) - (discLastRow[a] || 0); });

    discKeys.forEach(function(discKey) {
      var rows        = newRowsByDisc[discKey];
      var insertAfter = discLastRow[discKey];   // 1-based sheet row
      sheet.insertRowsAfter(insertAfter, rows.length);
      sheet.getRange(insertAfter + 1, 1, rows.length, rows[0].length).setValues(rows);
    });

    // Append any orphaned rows (unknown discipline) at the very bottom
    if (newRowsOrphaned.length) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRowsOrphaned.length, newRowsOrphaned[0].length)
           .setValues(newRowsOrphaned);
    }

  } finally {
    lock.releaseLock();
  }

  var msg = 'Updated ' + updated + ' task(s)';
  if (appended) msg += ', appended ' + appended + ' new task(s)';
  msg += ' in "' + SOURCE_SHEET + '".';
  return msg;
}

// ============================================================
//  HELPERS
// ============================================================

// Safe column-index lookup with fallback
function ci(cols, name, fallback) {
  return cols[name.toUpperCase()] !== undefined ? cols[name.toUpperCase()] : fallback;
}

// Parse a date value that may be a Date object, "2026/04/01", or "2026-04-01"
function parseSheetDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  var s = String(val).trim();
  if (!s || s === 'FALSE' || s === 'TRUE' || s === '-') return null;

  // Normalise slashes to dashes and zero-pad month/day
  s = s.replace(/\//g, '-');
  var parts = s.split('-');
  if (parts.length === 3) {
    var iso = parts[0].padStart(4, '0') + '-' +
              parts[1].padStart(2, '0') + '-' +
              parts[2].padStart(2, '0');
    var d = new Date(iso + 'T12:00:00');
    return isNaN(d.getTime()) ? null : d;
  }

  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d) {
  if (!(d instanceof Date)) return '';
  return d.getFullYear() + '-' +
         String(d.getMonth() + 1).padStart(2, '0') + '-' +
         String(d.getDate()).padStart(2, '0');
}

function statusColor(status) {
  if (!status) return '#64748b';
  return STATUS_COLORS[status.trim().toUpperCase()] || '#64748b';
}

function normKey(s) {
  return String(s).trim().toUpperCase().replace(/\s+/g, ' ');
}

// ============================================================
//  GANTT SETTINGS TAB — read / write
//  Only ever touches SETTINGS_SHEET. All other tabs untouched.
// ============================================================

var SETTINGS_KEYS = [
  'userLabelWidth', 'ganttBarFontSize', 'showTodayLine', 'darkMode',
  'flatMode', 'barTextColor', 'groupColors', 'collapsedGroups',
  'disciplineOrder', 'showName', 'showPhase', 'showNote'
];

var SETTINGS_DESCRIPTIONS = {
  userLabelWidth:   'Width of the task name column (px)',
  ganttBarFontSize: 'Font size for Gantt bar labels (pt)',
  showTodayLine:    'Show vertical Today line on chart (true/false)',
  darkMode:         'Dark theme enabled (true/false)',
  flatMode:         'Flat mode — suppress group header rows (true/false)',
  barTextColor:     'Bar label text colour hex — empty string = auto',
  groupColors:      'JSON object: per-discipline colour overrides',
  collapsedGroups:  'JSON object: which discipline groups are collapsed',
  disciplineOrder:  'JSON array: discipline group display order',
  showName:         'Show task name inside bar (true/false)',
  showPhase:        'Show phase/sub label inside bar (true/false)',
  showNote:         'Show note/status inside bar (true/false)'
};

// Return a plain object { key: value, ... } or null if the tab doesn't exist.
function readSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;

  var data   = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var result = {};
  for (var r = 0; r < data.length; r++) {
    var k = String(data[r][0]).trim();
    var v = String(data[r][1]).trim();
    if (k && k.toUpperCase() !== 'SETTING KEY') result[k] = v;
  }
  return Object.keys(result).length > 0 ? result : null;
}

// Write settings to SETTINGS_SHEET. Creates the tab if it doesn't exist.
// Never touches any other tab.
function writeSettings(settings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);

  sh.clearContents();

  var rows = [['SETTING KEY', 'VALUE']];
  SETTINGS_KEYS.forEach(function(key) {
    rows.push([key, settings[key] !== undefined ? String(settings[key]) : '']);
  });

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
}

// ============================================================
//  TEST: Run this manually from the Apps Script editor.
//  Select testCreateSettingsTab from the function dropdown
//  and click ▶ Run. Check the Execution Log for the result.
// ============================================================
function testWriteSettings() {
  try {
    var TAB_NAME = 'GANTT SETTINGS-DO NOT EDIT';
    var KEYS = [
      'userLabelWidth', 'ganttBarFontSize', 'showTodayLine', 'darkMode',
      'flatMode', 'barTextColor', 'groupColors', 'collapsedGroups',
      'disciplineOrder', 'showName', 'showPhase', 'showNote'
    ];
    var dummy = {
      userLabelWidth: '240', ganttBarFontSize: '12',
      showTodayLine: 'true', darkMode: 'false', flatMode: 'false',
      barTextColor: '', groupColors: '{}', collapsedGroups: '{}',
      disciplineOrder: '[]', showName: 'true', showPhase: 'false', showNote: 'true'
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('1. Got spreadsheet: ' + ss.getName());

    var sh = ss.getSheetByName(TAB_NAME);
    if (!sh) sh = ss.insertSheet(TAB_NAME);
    Logger.log('2. Got sheet: ' + sh.getName());

    sh.clearContents();
    Logger.log('3. Cleared contents');

    var rows = [['SETTING KEY', 'VALUE']];
    for (var i = 0; i < KEYS.length; i++) {
      rows.push([KEYS[i], dummy[KEYS[i]] !== undefined ? String(dummy[KEYS[i]]) : '']);
    }
    Logger.log('4. Built ' + rows.length + ' rows');

    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    Logger.log('5. Values written');

    SpreadsheetApp.flush();
    Logger.log('6. DONE — check the tab');

  } catch (e) {
    Logger.log('ERROR at step above: ' + e.toString());
  }
}

function testCreateSettingsTab() {
  // Completely self-contained — no global variables used
  var TAB_NAME = 'GANTT SETTINGS-DO NOT EDIT';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ss ? ss.getName() : 'NULL';
  Logger.log('Spreadsheet found: ' + name);

  if (!ss) {
    Logger.log('FAIL: getActiveSpreadsheet() returned null');
    return;
  }

  var existing = ss.getSheetByName(TAB_NAME);
  if (existing) {
    Logger.log('Tab already exists, deleting for clean test...');
    ss.deleteSheet(existing);
  }

  var sh = ss.insertSheet(TAB_NAME);
  Logger.log('insertSheet() called. Result: ' + (sh ? sh.getName() : 'NULL'));

  sh.getRange('A1').setValue('TEST OK ' + new Date().toString());
  SpreadsheetApp.flush();

  Logger.log('DONE — check your spreadsheet for the new tab.');
}

// ============================================================
//  OPTIONAL: Run this manually from the editor to verify
//  that your sheet is readable before deploying.
// ============================================================
function testImport() {
  var result = importFromTaskList();
  Logger.log('Tasks found: ' + result.tasks.length);
  result.tasks.forEach(function(t) {
    Logger.log(t.sub + ' | ' + t.name + ' | ' + t.start + ' → ' + t.end + ' | ' + t.color);
  });
}
