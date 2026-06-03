// ============================================================
//  LA COSTA HOTEL — TIMELINE SYNC  |  Google Apps Script
//
//  VERSION HISTORY
//  ───────────────
//  V1.02  2026-05-18
//    • Section (discipline) colors now saved as individual rows
//      in the GANTT SETTINGS tab: groupColor.DISCIPLINE → #hexcolor
//      instead of a single opaque JSON blob
//    • readSettings() reconstructs groupColors JSON from those rows
//      so the HTML frontend receives it in the same format as before
//    • writeSettings() no longer writes a groupColors JSON row —
//      replaced by individual groupColor.* rows (human-readable,
//      directly editable in the sheet)
//    • No changes to importFromTaskList or saveBackToTaskList
//
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

    var groupKey = lastDisc || disc || 'UNCATEGORIZED';

    var barNote = status;
    if (person && person.length < 20) barNote = status + (status ? ' · ' : '') + person;

    tasks.push({
      id:     idCounter++,
      name:   taskName,
      group:  groupKey,
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
// ============================================================
function saveBackToTaskList(payload) {
  if (!payload.tasks || !payload.tasks.length) return 'Nothing to save.';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"' + SOURCE_SHEET + '" tab not found.');

  var raw  = sheet.getDataRange().getValues();

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

  var lookup      = {};
  var discLastRow = {};
  var lastDisc    = '';
  for (var row = hRow + 1; row < raw.length; row++) {
    var disc = String(raw[row][CI.discipline] || '').trim();
    if (disc) lastDisc = disc;
    var task = String(raw[row][CI.task] || '').trim();
    if (task && task !== '-' && task !== '- -') {
      var key = normKey(lastDisc + '|' + task);
      if (!lookup[key]) lookup[key] = row + 1;
      discLastRow[normKey(lastDisc)] = row + 1;
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var updated  = 0;
  var appended = 0;

  var newRowsByDisc    = {};
  var newRowsOrphaned  = [];

  try {
    payload.tasks.forEach(function(t) {
      var discipline = (t.group || '').trim().toUpperCase();
      var taskName   = (t.name || '').trim();
      var key        = normKey(discipline + '|' + taskName);
      var sheetRow   = lookup[key];

      if (sheetRow) {
        if (t.start) sheet.getRange(sheetRow, CI.start + 1).setValue(t.start);
        if (t.end)   sheet.getRange(sheetRow, CI.end   + 1).setValue(t.end);
        if (t.note) {
          var statusCandidate = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[statusCandidate]) {
            sheet.getRange(sheetRow, CI.status + 1).setValue(statusCandidate);
          }
        }
        sheet.getRange(sheetRow, CI.schedule  + 1).setValue(true);
        if (t.type === 'milestone') sheet.getRange(sheetRow, CI.milestone + 1).setValue(true);
        updated++;

      } else if (taskName) {
        var totalCols = raw[hRow].length;
        var newRow = new Array(totalCols).fill('');
        newRow[CI.discipline] = discipline;
        newRow[CI.task]       = taskName;
        newRow[CI.start]      = t.start || '';
        newRow[CI.end]        = t.end   || '';
        newRow[CI.schedule]   = true;
        if (t.type === 'milestone') newRow[CI.milestone] = true;
        if (t.note) {
          var sc = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[sc]) newRow[CI.status] = sc;
        }

        var discKey = normKey(discipline);
        if (discLastRow[discKey]) {
          if (!newRowsByDisc[discKey]) newRowsByDisc[discKey] = [];
          newRowsByDisc[discKey].push(newRow);
        } else {
          newRowsOrphaned.push(newRow);
        }
        appended++;
      }
    });

    var discKeys = Object.keys(newRowsByDisc);
    discKeys.sort(function(a, b) { return (discLastRow[b] || 0) - (discLastRow[a] || 0); });

    discKeys.forEach(function(discKey) {
      var rows        = newRowsByDisc[discKey];
      var insertAfter = discLastRow[discKey];
      sheet.insertRowsAfter(insertAfter, rows.length);
      sheet.getRange(insertAfter + 1, 1, rows.length, rows[0].length).setValues(rows);
    });

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

function ci(cols, name, fallback) {
  return cols[name.toUpperCase()] !== undefined ? cols[name.toUpperCase()] : fallback;
}

function parseSheetDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  var s = String(val).trim();
  if (!s || s === 'FALSE' || s === 'TRUE' || s === '-') return null;

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
//
//  Section (discipline) colors are stored as individual rows:
//    groupColor.DISCIPLINE_NAME  →  #hexcolor
//  All other settings use the fixed SETTINGS_KEYS list.
// ============================================================

var SETTINGS_KEYS = [
  'userLabelWidth', 'ganttBarFontSize', 'showTodayLine', 'darkMode',
  'flatMode', 'barTextColor', 'collapsedGroups',
  'disciplineOrder', 'showName', 'showPhase', 'showNote'
];

var SETTINGS_DESCRIPTIONS = {
  userLabelWidth:   'Width of the task name column (px)',
  ganttBarFontSize: 'Font size for Gantt bar labels (pt)',
  showTodayLine:    'Show vertical Today line on chart (true/false)',
  darkMode:         'Dark theme enabled (true/false)',
  flatMode:         'Flat mode — suppress group header rows (true/false)',
  barTextColor:     'Bar label text colour hex — empty string = auto',
  collapsedGroups:  'JSON object: which discipline groups are collapsed',
  disciplineOrder:  'JSON array: discipline group display order',
  showName:         'Show task name inside bar (true/false)',
  showPhase:        'Show phase/sub label inside bar (true/false)',
  showNote:         'Show note/status inside bar (true/false)'
};

// Return a plain object { key: value, ... } or null if the tab doesn't exist.
// groupColor.* rows are collected and returned as groupColors JSON string
// so the HTML frontend receives the same format it always expected.
function readSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;

  var data          = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var result        = {};
  var groupColorsMap = {};

  for (var r = 0; r < data.length; r++) {
    var k = String(data[r][0]).trim();
    var v = String(data[r][1]).trim();
    if (!k || k.toUpperCase() === 'SETTING KEY') continue;

    if (k.indexOf('groupColor.') === 0) {
      // Individual discipline color row → collect into map
      groupColorsMap[k.substring('groupColor.'.length)] = v;
    } else {
      result[k] = v;
    }
  }

  // Reconstruct groupColors JSON from individual rows so the HTML
  // frontend receives it in the exact format applySettings() expects.
  if (Object.keys(groupColorsMap).length > 0) {
    result.groupColors = JSON.stringify(groupColorsMap);
  }

  return Object.keys(result).length > 0 ? result : null;
}

// Write settings to SETTINGS_SHEET. Creates the tab if it doesn't exist.
// Section colors (groupColors JSON) are expanded to individual readable rows.
// Never touches any other tab.
function writeSettings(settings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);

  sh.clearContents();

  var rows = [['SETTING KEY', 'VALUE']];

  // Fixed settings keys
  SETTINGS_KEYS.forEach(function(key) {
    rows.push([key, settings[key] !== undefined ? String(settings[key]) : '']);
  });

  // Section colors — one row per discipline, sorted alphabetically
  if (settings.groupColors) {
    try {
      var gc = JSON.parse(settings.groupColors);
      Object.keys(gc).sort().forEach(function(disc) {
        rows.push(['groupColor.' + disc, gc[disc]]);
      });
    } catch(e) {}
  }

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
}

// ============================================================
//  TEST: Run this manually from the Apps Script editor.
// ============================================================
function testWriteSettings() {
  try {
    var dummy = {
      userLabelWidth: '240', ganttBarFontSize: '12',
      showTodayLine: 'true', darkMode: 'false', flatMode: 'false',
      barTextColor: '', collapsedGroups: '{}',
      disciplineOrder: '[]', showName: 'true', showPhase: 'false', showNote: 'true',
      groupColors: '{"ARCHITECTURE":"#0284c7","STRUCTURAL":"#16a34a","MEP":"#f59e0b"}'
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('1. Got spreadsheet: ' + ss.getName());

    writeSettings(dummy);
    Logger.log('2. writeSettings() complete');

    var back = readSettings();
    Logger.log('3. readSettings() returned: ' + JSON.stringify(back));

    SpreadsheetApp.flush();
    Logger.log('4. DONE — check the GANTT SETTINGS tab');

  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
  }
}

function testCreateSettingsTab() {
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

function testImport() {
  var result = importFromTaskList();
  Logger.log('Tasks found: ' + result.tasks.length);
  result.tasks.forEach(function(t) {
    Logger.log(t.sub + ' | ' + t.name + ' | ' + t.start + ' → ' + t.end + ' | ' + t.color);
  });
}
