// ============================================================
//  LA COSTA HOTEL — TIMELINE SYNC  |  Google Apps Script
//  File: Code_V1.05.gs
//
//  PURPOSE
//  ───────
//  Two-way sync between the interactive Gantt HTML frontend
//  and a Google Sheet. Handles three categories of data:
//
//    1. PROJECT TASK LIST tab  — task schedule data
//       (discipline, task name, start/end dates, status)
//       Read on Load; dates + status written back on Save.
//
//    2. GANTT TASK PARAMS tab  — per-task Gantt display params
//       (color override, bar type, bar style, symbol)
//       Read on Load and merged into tasks; fully rewritten on Save.
//       These are Gantt-only values that don't belong in the main list.
//
//    3. GANTT SETTINGS tab     — chart-wide UI settings
//       (label width, font size, dark mode, section colors, etc.)
//       Read on Load; fully rewritten on Save.
//       Section colors stored as individual groupColor.DISC rows.
//
//  ENTRY POINTS
//  ────────────
//  doGet(e)   — called by HTML on Load. Returns tasks + task params
//               (merged) + chart settings as a single JSON response.
//  doPost(e)  — called by HTML on Save. Writes task schedule data,
//               task params, and chart settings independently so a
//               failure in one block never blocks the others.
//
//  VERSION HISTORY
//  ───────────────
//  V1.05  2026-05-18
//    • No backend changes — print mode is HTML/JS only
//
//  V1.04  2026-05-18
//    • Notes field returned from importFromTaskList and saved back
//      to PROJECT TASK LIST on Save
//    • Bar text renders LEFT of bar (text-anchor:end, no overlap)
//      when estimated label width exceeds bar width
//
//  V1.03  2026-05-18
//    • New GANTT TASK PARAMS tab: per-task color override, bar type,
//      bar style, and symbol stored separately from PROJECT TASK LIST
//    • readTaskParams() / writeTaskParams() helpers added
//    • doGet merges task params into tasks before returning
//    • doPost writes task params independently (errors don't block
//      the main task save or settings save)
//    • No changes to importFromTaskList or saveBackToTaskList
//
//  V1.02  2026-05-18
//    • Section colors saved as individual groupColor.DISC rows in
//      GANTT SETTINGS tab instead of a single opaque JSON blob
//    • readSettings() reconstructs groupColors JSON from those rows
//    • writeSettings() expanded to write individual groupColor.* rows
//
//  V1.01  2026-05-18
//    • Added GANTT SETTINGS-DO NOT EDIT tab support
//    • doGet returns saved chart settings alongside tasks
//    • doPost saves chart settings independently of task save
//    • readSettings() / writeSettings() helpers added
//
//  V1.0   2026-05-18  (baseline)
//    • Reads PROJECT TASK LIST; rows with SCHEDULE or MILESTONE
//      checkbox appear on the Gantt
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

// ---- Tab names ----
var SOURCE_SHEET      = 'PROJECT TASK LIST';          // master task data — never reformatted
var TASK_PARAMS_SHEET = 'GANTT TASK PARAMS';          // per-task Gantt display params
var SETTINGS_SHEET    = 'GANTT SETTINGS-DO NOT EDIT'; // chart-wide UI settings

// ---- Status text → Gantt bar colour ----
// Used as the fallback bar colour when no colour override is set on the task
// or its discipline group.
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
//  GET  →  Load tasks + task params + settings
// ============================================================
function doGet(e) {
  try {
    var result = importFromTaskList();

    // Merge per-task Gantt display params (colour override, type,
    // style, symbol) on top of the base task data from the main sheet.
    var taskParams = readTaskParams();
    if (Object.keys(taskParams).length > 0) {
      result.tasks.forEach(function(t) {
        var key = normKey((t.group || '') + '|' + (t.name || ''));
        var p   = taskParams[key];
        if (!p) return;
        if (p.color)  t.colorOverride = p.color;
        if (p.type)   t.type          = p.type;
        if (p.style === 'dashed')         { t.dashed = 'true';  t.dashedOutline = 'false'; }
        if (p.style === 'dashed-outline') { t.dashed = 'false'; t.dashedOutline = 'true';  }
        if (p.style === 'solid')          { t.dashed = 'false'; t.dashedOutline = 'false'; }
        if (p.symbol) t.symbol = p.symbol;
      });
    }

    result.settings = readSettings();
    return buildResponse(result);

  } catch (err) {
    return buildResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
//  POST  →  Save task schedule, task params, and settings
//           Each section runs independently — an error in one
//           block does not prevent the others from running.
// ============================================================
function doPost(e) {
  var taskMsg       = '';
  var taskErr       = '';
  var taskParamsErr = '';
  var settingsErr   = '';

  try {
    var payload = JSON.parse(e.postData.contents);

    // 1. Update dates + status in PROJECT TASK LIST
    try {
      taskMsg = saveBackToTaskList(payload);
    } catch (tErr) {
      taskErr = tErr.toString();
    }

    // 2. Write per-task Gantt params to GANTT TASK PARAMS
    if (payload.tasks) {
      try {
        writeTaskParams(payload.tasks);
      } catch (tpErr) {
        taskParamsErr = tpErr.toString();
      }
    }

    // 3. Write chart-wide settings to GANTT SETTINGS tab
    if (payload.settings) {
      try {
        writeSettings(payload.settings);
      } catch (sErr) {
        settingsErr = sErr.toString();
      }
    }

    return buildResponse({
      success:         true,
      message:         taskMsg,
      taskError:       taskErr,
      taskParamsError: taskParamsErr,
      settingsError:   settingsErr
    });

  } catch (err) {
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
//
//  Only rows where SCHEDULE=TRUE or MILESTONE=TRUE and with
//  valid START DATE + END DATE are included. Task name, dates,
//  and status are the only fields read — all other Gantt display
//  params come from the GANTT TASK PARAMS tab via doGet.
// ============================================================
function importFromTaskList() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"' + SOURCE_SHEET + '" tab not found. Check the sheet name.');

  var raw = sheet.getDataRange().getValues();

  // Find the header row (the one containing "DISCIPLINE")
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow = i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) throw new Error('Could not find the DISCIPLINE header row in "' + SOURCE_SHEET + '".');

  // Map column name → zero-based index
  var cols = {};
  raw[hRow].forEach(function(h, idx) {
    cols[String(h).trim().toUpperCase()] = idx;
  });

  // Column indices — auto-detected by header name, fallback to known positions
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

  var tasks     = [];
  var idCounter = 1;
  var lastDisc  = '';

  for (var row = hRow + 1; row < raw.length; row++) {
    var r = raw[row];

    var disc = String(r[CI.discipline] || '').trim();
    if (disc) lastDisc = disc;

    var taskName = String(r[CI.task] || '').trim();
    if (!taskName || /^[\s\-]+$/.test(taskName)) continue;

    var isScheduled = (r[CI.schedule] === true || String(r[CI.schedule]).toUpperCase() === 'TRUE');
    var isMilestone = (r[CI.milestone] === true || String(r[CI.milestone]).toUpperCase() === 'TRUE');
    if (!isScheduled && !isMilestone) continue;

    var startDate = parseSheetDate(r[CI.start]);
    var endDate   = parseSheetDate(r[CI.end]);
    if (!startDate || !endDate) continue;
    if (endDate.getTime() <= startDate.getTime()) continue;

    var status     = String(r[CI.status]     || '').trim();
    var person     = String(r[CI.person]     || '').trim();
    var groupKey   = lastDisc || disc || 'UNCATEGORIZED';
    var barNote    = status;
    if (person && person.length < 20) barNote = status + (status ? ' · ' : '') + person;

    // Base type from MILESTONE checkbox; may be overridden by GANTT TASK PARAMS
    var taskType = isMilestone && !isScheduled ? 'milestone' : 'bar';

    tasks.push({
      id:           idCounter++,
      name:         taskName,
      group:        groupKey,
      type:         taskType,
      start:        fmtDate(startDate),
      end:          fmtDate(endDate),
      color:        statusColor(status),   // fallback; overridden by groupColor or colorOverride
      note:         barNote,
      dashed:       'false',
      dashedOutline:'false',
      symbol:       '',
      notes:        String(r[CI.notes] || '').trim()
      // colorOverride added by doGet after readTaskParams() merge
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
//  Only START DATE, END DATE, and STATUS are written back to
//  PROJECT TASK LIST. All Gantt-specific display params
//  (color, style, symbol, type) are handled by writeTaskParams().
// ============================================================
function saveBackToTaskList(payload) {
  if (!payload.tasks || !payload.tasks.length) return 'Nothing to save.';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"' + SOURCE_SHEET + '" tab not found.');

  var raw = sheet.getDataRange().getValues();

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
    status:     ci(cols, 'STATUS',     12),
    notes:      ci(cols, 'NOTES',      20)
  };

  // Build lookup: "DISCIPLINE|TASKNAME" → 1-based sheet row
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
  var newRowsByDisc   = {};
  var newRowsOrphaned = [];

  try {
    payload.tasks.forEach(function(t) {
      var discipline = (t.group || '').trim().toUpperCase();
      var taskName   = (t.name  || '').trim();
      var key        = normKey(discipline + '|' + taskName);
      var sheetRow   = lookup[key];

      if (sheetRow) {
        // Update dates and status for existing row
        if (t.start) sheet.getRange(sheetRow, CI.start + 1).setValue(t.start);
        if (t.end)   sheet.getRange(sheetRow, CI.end   + 1).setValue(t.end);
        if (t.note) {
          var statusCandidate = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[statusCandidate]) {
            sheet.getRange(sheetRow, CI.status + 1).setValue(statusCandidate);
          }
        }
        if (typeof t.notes !== 'undefined') sheet.getRange(sheetRow, CI.notes + 1).setValue(t.notes);
        sheet.getRange(sheetRow, CI.schedule + 1).setValue(true);
        if (t.type === 'milestone') sheet.getRange(sheetRow, CI.milestone + 1).setValue(true);
        updated++;

      } else if (taskName) {
        // Queue new row under its discipline group
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
        if (t.notes) newRow[CI.notes] = t.notes;

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

    // Insert new rows after the last row of each discipline (descending order
    // to avoid shifting row indices for subsequent insertions)
    var discKeys = Object.keys(newRowsByDisc);
    discKeys.sort(function(a, b) { return (discLastRow[b] || 0) - (discLastRow[a] || 0); });
    discKeys.forEach(function(discKey) {
      var rows = newRowsByDisc[discKey];
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
//  GANTT TASK PARAMS TAB — read / write
//
//  Tab columns: KEY | COLOR | TYPE | STYLE | SYMBOL
//    KEY    — "DISCIPLINE|TASKNAME" (same key used in saveBackToTaskList)
//    COLOR  — hex color override (#rrggbb) or empty (inherit section color)
//    TYPE   — bar / milestone / flag  (overrides MILESTONE checkbox)
//    STYLE  — solid / dashed / dashed-outline
//    SYMBOL — emoji or short text displayed inside milestone / flag
//
//  This tab is created automatically on the first Save.
//  All rows are rewritten on every Save (no partial updates).
// ============================================================

// Return a map of normKey(DISC|TASK) → { color, type, style, symbol }
// Returns an empty object if the tab doesn't exist yet.
function readTaskParams() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_PARAMS_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};

  var data = sh.getDataRange().getValues();

  // Locate header row (expects "KEY" in column A)
  var hRow = -1;
  for (var i = 0; i < Math.min(data.length, 5); i++) {
    if (String(data[i][0]).trim().toUpperCase() === 'KEY') { hRow = i; break; }
  }
  if (hRow < 0) return {};

  // Map header names → column indices
  var cols = {};
  data[hRow].forEach(function(h, idx) {
    cols[String(h).trim().toUpperCase()] = idx;
  });
  var colKey    = cols['KEY']    !== undefined ? cols['KEY']    : 0;
  var colColor  = cols['COLOR']  !== undefined ? cols['COLOR']  : 1;
  var colType   = cols['TYPE']   !== undefined ? cols['TYPE']   : 2;
  var colStyle  = cols['STYLE']  !== undefined ? cols['STYLE']  : 3;
  var colSymbol = cols['SYMBOL'] !== undefined ? cols['SYMBOL'] : 4;

  var params = {};
  for (var r = hRow + 1; r < data.length; r++) {
    var key = normKey(String(data[r][colKey] || '').trim());
    if (!key) continue;
    params[key] = {
      color:  String(data[r][colColor]  || '').trim(),
      type:   String(data[r][colType]   || '').trim().toLowerCase(),
      style:  String(data[r][colStyle]  || '').trim().toLowerCase(),
      symbol: String(data[r][colSymbol] || '').trim()
    };
  }
  return params;
}

// Write all task display params to GANTT TASK PARAMS.
// Rewrites the entire tab — one row per task.
// Creates the tab if it doesn't exist.
function writeTaskParams(tasks) {
  if (!tasks || !tasks.length) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_PARAMS_SHEET);
  if (!sh) sh = ss.insertSheet(TASK_PARAMS_SHEET);

  sh.clearContents();

  var rows = [['KEY', 'COLOR', 'TYPE', 'STYLE', 'SYMBOL']];

  tasks.forEach(function(t) {
    var disc = String(t.group || '').trim().toUpperCase();
    var name = String(t.name  || '').trim();
    if (!name) return;

    var key   = normKey(disc + '|' + name);
    var color = String(t.colorOverride || '').trim();
    var type  = String(t.type  || 'bar').trim().toLowerCase();
    var style;
    if (String(t.dashed) === 'true' || t.dashed === true) {
      style = 'dashed';
    } else if (String(t.dashedOutline) === 'true' || t.dashedOutline === true) {
      style = 'dashed-outline';
    } else {
      style = 'solid';
    }
    var symbol = String(t.symbol || '').trim();

    rows.push([key, color, type, style, symbol]);
  });

  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.hideSheet();
}

// ============================================================
//  GANTT SETTINGS TAB — read / write
//
//  Fixed settings stored as KEY → VALUE rows.
//  Section colors stored as individual rows:
//    groupColor.DISCIPLINE_NAME → #hexcolor
//  readSettings() reconstructs groupColors JSON from those rows
//  so the HTML frontend receives it in the expected format.
// ============================================================

var SETTINGS_KEYS = [
  'userLabelWidth', 'ganttBarFontSize', 'showTodayLine', 'darkMode',
  'flatMode', 'barTextColor', 'collapsedGroups',
  'disciplineOrder', 'showName', 'showPhase', 'showNote',
  'groupTint', 'currentScale', 'zoomLevel', 'chartStart', 'chartEnd'
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

// Return { key: value, ... } or null if the tab doesn't exist.
// groupColor.* rows are collected and returned as a groupColors JSON
// string so the HTML frontend receives the format applySettings() expects.
function readSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;

  var data           = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var result         = {};
  var groupColorsMap = {};

  for (var r = 0; r < data.length; r++) {
    var k = String(data[r][0]).trim();
    var v = String(data[r][1]).trim();
    if (!k || k.toUpperCase() === 'SETTING KEY') continue;

    if (k.indexOf('groupColor.') === 0) {
      groupColorsMap[k.substring('groupColor.'.length)] = v;
    } else {
      result[k] = v;
    }
  }

  if (Object.keys(groupColorsMap).length > 0) {
    result.groupColors = JSON.stringify(groupColorsMap);
  }

  return Object.keys(result).length > 0 ? result : null;
}

// Write chart-wide settings and section colors to GANTT SETTINGS tab.
// Creates the tab if it doesn't exist. Fully rewrites on every Save.
function writeSettings(settings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);

  sh.clearContents();

  var rows = [['SETTING KEY', 'VALUE']];

  // Fixed chart-wide settings
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
  sh.hideSheet();
}

// ============================================================
//  HELPERS
// ============================================================

// Safe column-index lookup with positional fallback
function ci(cols, name, fallback) {
  return cols[name.toUpperCase()] !== undefined ? cols[name.toUpperCase()] : fallback;
}

// Parse a date value that may be a Date object or "YYYY-MM-DD" / "YYYY/MM/DD"
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

// Normalise a key string: trim, uppercase, collapse internal spaces
function normKey(s) {
  return String(s).trim().toUpperCase().replace(/\s+/g, ' ');
}

// ============================================================
//  TEST FUNCTIONS — run manually from the Apps Script editor
// ============================================================

// Verify the full read/write cycle for all three tabs.
function testWriteSettings() {
  try {
    var dummySettings = {
      userLabelWidth: '240', ganttBarFontSize: '12',
      showTodayLine: 'true', darkMode: 'false', flatMode: 'false',
      barTextColor: '', collapsedGroups: '{}',
      disciplineOrder: '[]', showName: 'true', showPhase: 'false', showNote: 'true',
      groupColors: '{"ARCHITECTURE":"#0284c7","STRUCTURAL":"#16a34a","MEP":"#f59e0b"}'
    };
    var dummyTasks = [
      { group: 'ARCHITECTURE', name: 'SCHEMATIC DESIGN', type: 'bar',
        dashed: false, dashedOutline: false, symbol: '',
        colorOverride: '#ff0000' },
      { group: 'STRUCTURAL', name: 'FOUNDATION REVIEW', type: 'milestone',
        dashed: false, dashedOutline: false, symbol: '⭐',
        colorOverride: '' }
    ];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet: ' + ss.getName());

    writeSettings(dummySettings);
    Logger.log('writeSettings() done');

    writeTaskParams(dummyTasks);
    Logger.log('writeTaskParams() done');

    var s = readSettings();
    Logger.log('readSettings(): ' + JSON.stringify(s));

    var p = readTaskParams();
    Logger.log('readTaskParams(): ' + JSON.stringify(p));

    SpreadsheetApp.flush();
    Logger.log('DONE — check GANTT SETTINGS and GANTT TASK PARAMS tabs');

  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
  }
}

// Verify sheet readability and task count before deploying.
function testImport() {
  var result = importFromTaskList();
  Logger.log('Tasks found: ' + result.tasks.length);
  result.tasks.forEach(function(t) {
    Logger.log(t.group + ' | ' + t.name + ' | ' + t.start + ' → ' + t.end + ' | ' + t.color);
  });
}

// Create / recreate the GANTT SETTINGS tab from scratch.
function testCreateSettingsTab() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SETTINGS_SHEET);
  if (existing) { ss.deleteSheet(existing); Logger.log('Deleted existing tab.'); }
  var sh = ss.insertSheet(SETTINGS_SHEET);
  sh.getRange('A1').setValue('TEST OK ' + new Date().toString());
  SpreadsheetApp.flush();
  Logger.log('DONE — check spreadsheet for new tab: ' + SETTINGS_SHEET);
}
