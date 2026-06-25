// ============================================================
//  TIMELINE SYNC  |  Google Apps Script
//  File: Code_V1.20.gs
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
//       (color override, bar type, bar style, symbol, taskId)
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
//  V1.20  2026-06-25 06:43
//    • GANTT TASK PARAMS tab now stores TASKID and KEY (DISC|TASKNAME) columns together.
//      readTaskParams() builds two indexes: byId (integer taskId) and byKey (normKey string).
//      doGet() merges using byId[t.taskId] || byKey[normKey] — taskId is primary (rename-safe)
//      and byKey is the fallback if IDs are regenerated or the IDs tab is lost.
//      writeTaskParams() builds all rows BEFORE clearing the sheet so a failed save never
//      wipes existing params; clearContents only runs when there is confirmed data to write.
//      COLOR column now stores colorOverride || color so ALL task colors are persisted on
//      every Save — the displayed colour is locked in regardless of its source.
//      Result: colour, type, style, and symbol survive renames, discipline changes, AND
//      accidental ID-tab loss. Backward-compatible with old KEY-only and TASKID-only tabs.
//    • SETTINGS_KEYS also gains showGroupBars (new toggle from HTML V1.19/V1.20).
//    • SETTINGS_KEYS gains showDateColumns (persists START/END date column visibility from HTML V1.21).
//
//  V1.19  2026-06-17
//    • SETTINGS_KEYS expanded with 2 new keys introduced in HTML V1.19:
//        matchHdrToGroupColor — group header bg adopts the group's bar colour when true
//        showRollupTicks      — show/hide vertical tick lines at rollup bar ends
//    • SETTINGS_DESCRIPTIONS updated for both new keys.
//    • No other logic changes.
//
//  V1.18  2026-06-06
//    • Persistent numeric TASKID assigned to each task on Load and
//      stored in a dedicated GANTT TASK IDS-DO NOT EDIT tab (KEY|TASKID)
//      that is NEVER touched by writeTaskParams — IDs survive every Save.
//    • doGet reads existing IDs from the IDs tab, assigns new sequential
//      IDs only to tasks without one, then writes back only if new IDs
//      were assigned.
//    • saveBackToTaskList() uses taskId for rename detection: when a
//      task's key is not found in the sheet but its taskId matches a
//      known IDs-tab entry, the existing sheet row is updated in-place
//      (discipline + task name rewritten) instead of appending a new row.
//
//  V1.17  2026-06-06  (HTML only — no backend changes)
//
//  V1.16  2026-06-03
//    • Task deletion syncs to spreadsheet — rows removed from the chart
//      are deleted from PROJECT TASK LIST on Save.
//    • MILESTONE column explicitly written on every Save (TRUE/FALSE) so
//      un-checking it in the sheet is respected on next Load.
//    • Fixed: stale GANTT TASK PARAMS type='milestone' override is
//      ignored when the sheet's MILESTONE column is FALSE (_sheetMilestone
//      guard added to doGet).
//
//  V1.14  2026-05-22
//    • SETTINGS_KEYS expanded with 4 new keys introduced in HTML V1.14:
//        ganttCollapsedGroups      — Gantt chart collapse state (separate from task list)
//        milestonesCollapsedGroups — Milestones tab collapse state
//        flagsCollapsedGroups      — Flags tab collapse state
//        ganttRollupMarkersVisible — Global show/hide of milestone/flag markers on rollup bars
//    • SETTINGS_DESCRIPTIONS updated with descriptions for all 4 new keys.
//    • No other logic changes — doGet / doPost / readSettings / writeSettings unchanged.
//
//  V1.13  2026-05-22 (updated)
//    • SETTINGS_DESCRIPTIONS completed for all 28 keys.
//    • writeSettings() now writes a 3rd DESCRIPTION column in
//      GANTT SETTINGS-DO NOT EDIT so the tab is self-documenting
//      (readSettings still reads cols A+B only — no breaking change).
//    • groupColor.* rows now include a per-discipline description.
//    • GanttSetup_AddOn.gs synced with same SETTINGS_DESCRIPTIONS
//      and 3-column writeSettings.
//
//  V1.13  2026-05-21
//    • (menu functions removed in V1.15)
//
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
//  5. Click the "📊 Gantt Timeline → Get Web App URL" menu in the sheet
//     (or copy the URL from the deployment dialog) → paste into timeline ⚙ Setup
//
//  UPDATING (already deployed):
//  Deploy → Manage deployments → pencil icon → New version → Deploy
//  (the URL stays the same)
// ============================================================

// ---- Tab names ----
var SOURCE_SHEET      = 'PROJECT TASK LIST';             // master task data — never reformatted
var TASK_PARAMS_SHEET = 'GANTT TASK PARAMS-DO NOT EDIT'; // per-task Gantt display params (color, type, style, symbol, deps)
var TASK_IDS_SHEET    = 'GANTT TASK IDS-DO NOT EDIT';    // persistent numeric task ID registry — never cleared by saves
var SETTINGS_SHEET    = 'GANTT SETTINGS-DO NOT EDIT';    // chart-wide UI settings

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

    // Assign persistent numeric IDs FIRST so params can be looked up by taskId.
    // IDs are stored in the dedicated IDs tab and never touched by writeTaskParams.
    var taskIds = readTaskIds(); // { ids: {normKey: taskId}, maxId: N }
    var newIdsAssigned = false;
    result.tasks.forEach(function(t) {
      var key = normKey((t.group || '') + '|' + (t.name || ''));
      if (taskIds.ids[key]) {
        t.taskId = taskIds.ids[key];
      } else {
        taskIds.maxId++;
        t.taskId = taskIds.maxId;
        taskIds.ids[key] = t.taskId;
        newIdsAssigned = true;
      }
    });
    if (newIdsAssigned) {
      try { writeTaskIds(taskIds); } catch(e) {}
    } else {
      try { hideTaskIdsSheet(); } catch(e) {}
    }

    // Merge per-task Gantt display params (colour override, type, style, symbol).
    // readTaskParams() returns { byId, byKey } — two indexes into the same param objects.
    // Lookup tries taskId first (rename-safe), then falls back to normKey(DISC|NAME)
    // so params are found even if the IDs tab was regenerated or the tab is in legacy format.
    var paramsResult = readTaskParams();
    result.tasks.forEach(function(t) {
      var taskKey = normKey((t.group || '') + '|' + (t.name || ''));
      var p = paramsResult.byId[t.taskId] || paramsResult.byKey[taskKey];
      if (p) {
        if (p.color)  t.colorOverride = p.color;
        // Only allow task-params to set type='milestone' when the sheet MILESTONE column
        // is still checked — prevents stale params from persisting a milestone after
        // the user unchecks it directly in the spreadsheet.
        if (p.type) {
          if (p.type !== 'milestone' || t._sheetMilestone) t.type = p.type;
        }
        if (p.style === 'dashed')         { t.dashed = 'true';  t.dashedOutline = 'false'; }
        if (p.style === 'dashed-outline') { t.dashed = 'false'; t.dashedOutline = 'true';  }
        if (p.style === 'solid')          { t.dashed = 'false'; t.dashedOutline = 'false'; }
        if (p.symbol) t.symbol = p.symbol;
        if (p.deps)   t.dependencies = p.deps;
      }
      delete t._sheetMilestone; // internal flag — not sent to frontend
    });

    result.settings = readSettings();
    result.statusColors = STATUS_COLORS; // send full status→color map so frontend stays in sync
    try { result.spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName(); } catch(e) {}
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
    // Skip only if end is strictly before start — same-day tasks are valid (render as 1-day bar).
    if (endDate.getTime() < startDate.getTime()) continue;

    var status     = String(r[CI.status]     || '').trim();
    var person     = String(r[CI.person]     || '').trim();
    var groupKey   = lastDisc || disc || 'UNCATEGORIZED';
    var barNote    = status;

    // Any row with MILESTONE checked loads as milestone type regardless of SCHEDULE.
    // GANTT TASK PARAMS can further override this per-task.
    var taskType = isMilestone ? 'milestone' : 'bar';

    tasks.push({
      id:            idCounter++,
      name:          taskName,
      group:         groupKey,
      type:          taskType,
      start:         fmtDate(startDate),
      end:           fmtDate(endDate),
      color:         statusColor(status),   // fallback; overridden by groupColor or colorOverride
      note:          barNote,
      dashed:        'false',
      dashedOutline: 'false',
      symbol:        '',
      notes:         String(r[CI.notes] || '').trim(),
      _sheetMilestone: isMilestone          // used by doGet to guard task-params type override
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

  // Build lookup: "DISCIPLINE|TASKNAME" → 1-based sheet row (scheduled/milestone rows only)
  var lookup      = {};
  var discLastRow = {};
  var lastDisc    = '';
  for (var row = hRow + 1; row < raw.length; row++) {
    var disc = String(raw[row][CI.discipline] || '').trim();
    if (disc) lastDisc = disc;
    var task = String(raw[row][CI.task] || '').trim();
    if (task && task !== '-' && task !== '- -') {
      var isRowScheduled  = (raw[row][CI.schedule]  === true || String(raw[row][CI.schedule]  || '').toUpperCase() === 'TRUE');
      var isRowMilestone  = (raw[row][CI.milestone] === true || String(raw[row][CI.milestone] || '').toUpperCase() === 'TRUE');
      var key = normKey(lastDisc + '|' + task);
      if ((isRowScheduled || isRowMilestone) && !lookup[key]) lookup[key] = row + 1;
      discLastRow[normKey(lastDisc)] = row + 1;
    }
  }

  // Keys present in the payload — used to detect tasks deleted from the chart
  var payloadKeys = {};
  payload.tasks.forEach(function(t) {
    var discipline = (t.group || '').trim().toUpperCase();
    var taskName   = (t.name  || '').trim();
    if (taskName) payloadKeys[normKey(discipline + '|' + taskName)] = true;
  });

  // Build taskId → old sheet key map for rename detection (reads the IDs tab)
  var savedTaskIds = readTaskIds();
  try { hideTaskIdsSheet(); } catch(e) {}
  var idToKey = {};
  Object.keys(savedTaskIds.ids).forEach(function(k) {
    idToKey[savedTaskIds.ids[k]] = k;
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  var updated   = 0;
  var appended  = 0;
  var deleted   = 0;
  var moveTasks = [];          // discipline-change moves: {oldRow, newRow[], targetDiscKey}
  var newRowsByDisc   = {};
  var newRowsOrphaned = [];

  try {
    // ── 1. Update existing rows ─────────────────────────────────
    payload.tasks.forEach(function(t) {
      var discipline = (t.group || '').trim().toUpperCase();
      var taskName   = (t.name  || '').trim();
      var key        = normKey(discipline + '|' + taskName);
      var sheetRow   = lookup[key];

      // Rename / discipline-change detection via persistent taskId.
      // • Name-only rename   → update in place (sheetRow set, payloadKeys[oldKey] protected)
      // • Discipline change  → queued in moveTasks for Phase 4 (explicit move)
      var isRename = false;
      var renamedOldKey = null;
      if (!sheetRow && t.taskId) {
        renamedOldKey = idToKey[parseInt(t.taskId, 10)];
        if (renamedOldKey && lookup[renamedOldKey] && !payloadKeys[renamedOldKey]) {
          isRename = true;
          var oldDiscNorm        = renamedOldKey.split('|')[0];
          var isDisciplineChange = (oldDiscNorm !== normKey(discipline));

          // Always protect old key so Phase 2 doesn't delete it
          payloadKeys[renamedOldKey] = true;

          if (!isDisciplineChange) {
            // Name-only rename: update the existing row in place
            sheetRow = lookup[renamedOldKey];
          } else {
            // Discipline change: build a move record; Phase 4 handles delete + insert
            var totalColsM = raw[hRow].length;
            var moveRow = new Array(totalColsM).fill('');
            moveRow[CI.discipline] = discipline;
            moveRow[CI.task]       = taskName;
            moveRow[CI.start]      = t.start || '';
            moveRow[CI.end]        = t.end   || '';
            moveRow[CI.schedule]   = true;
            moveRow[CI.milestone]  = (t.type === 'milestone');
            if (t.note) {
              var moveSc = t.note.split('·')[0].trim().toUpperCase();
              if (STATUS_COLORS[moveSc]) moveRow[CI.status] = moveSc;
            }
            if (t.notes) moveRow[CI.notes] = t.notes;
            moveTasks.push({ oldRow: lookup[renamedOldKey], newRow: moveRow, targetDiscKey: normKey(discipline) });
          }

          // Update IDs registry: new key inherits the same persistent ID
          if (savedTaskIds.ids[renamedOldKey]) {
            savedTaskIds.ids[key] = savedTaskIds.ids[renamedOldKey];
            delete savedTaskIds.ids[renamedOldKey];
            try { writeTaskIds(savedTaskIds); } catch(e) {}
          }
        }
      }

      if (sheetRow) {
        if (isRename) {
          // Name-only rename: overwrite discipline + task name in the existing row
          sheet.getRange(sheetRow, CI.discipline + 1).setValue(discipline);
          sheet.getRange(sheetRow, CI.task       + 1).setValue(taskName);
        }
        if (t.start) sheet.getRange(sheetRow, CI.start + 1).setValue(t.start);
        if (t.end)   sheet.getRange(sheetRow, CI.end   + 1).setValue(t.end);
        if (t.note) {
          var statusCandidate = t.note.split('·')[0].trim().toUpperCase();
          if (STATUS_COLORS[statusCandidate]) {
            sheet.getRange(sheetRow, CI.status + 1).setValue(statusCandidate);
          }
        }
        if (typeof t.notes !== 'undefined') sheet.getRange(sheetRow, CI.notes + 1).setValue(t.notes);
        sheet.getRange(sheetRow, CI.schedule  + 1).setValue(true);
        // Write MILESTONE column based on current type — clears it when task is no longer a milestone
        sheet.getRange(sheetRow, CI.milestone + 1).setValue(t.type === 'milestone');
        updated++;

      } else if (taskName && !moveTasks.some(function(m) { return m.newRow[CI.discipline] === discipline && m.newRow[CI.task] === taskName; })) {
        // Queue genuinely new row under its discipline group (not a discipline-change move)
        var totalCols = raw[hRow].length;
        var newRow = new Array(totalCols).fill('');
        newRow[CI.discipline] = discipline;
        newRow[CI.task]       = taskName;
        newRow[CI.start]      = t.start || '';
        newRow[CI.end]        = t.end   || '';
        newRow[CI.schedule]   = true;
        newRow[CI.milestone]  = (t.type === 'milestone');
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

    // ── 2. Delete rows that were removed from the chart ─────────
    // Rows in the lookup that are NOT in the payload were deleted
    // by the user. Remove them from the sheet top-to-bottom
    // reversed so row indices stay valid as we go.
    var rowsToDelete = [];
    Object.keys(lookup).forEach(function(key) {
      if (!payloadKeys[key]) rowsToDelete.push(lookup[key]);
    });
    rowsToDelete.sort(function(a, b) { return b - a; }); // descending
    rowsToDelete.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
      deleted++;
      // Adjust discLastRow for any discipline whose last row was at or after the deleted row
      Object.keys(discLastRow).forEach(function(dk) {
        if (discLastRow[dk] >= rowNum) discLastRow[dk]--;
      });
    });

    // ── 4. Move discipline-changed tasks ────────────────────────
    // Two-pass approach so row-number accounting stays exact:
    // Pass A: delete all old rows (descending order → no index drift)
    // Pass B: insert at target discipline's current last row (descending → safe)
    if (moveTasks.length) {
      // Pass A — delete old rows, adjusting discLastRow after each
      moveTasks.sort(function(a, b) { return b.oldRow - a.oldRow; });
      moveTasks.forEach(function(m) {
        sheet.deleteRow(m.oldRow);
        Object.keys(discLastRow).forEach(function(dk) {
          if (discLastRow[dk] >= m.oldRow) discLastRow[dk]--;
        });
      });

      // Pass B — insert at target position, adjusting discLastRow after each
      moveTasks.sort(function(a, b) {
        return (discLastRow[b.targetDiscKey] || 0) - (discLastRow[a.targetDiscKey] || 0);
      });
      moveTasks.forEach(function(m) {
        var insertAfter = discLastRow[m.targetDiscKey];
        if (insertAfter) {
          sheet.insertRowsAfter(insertAfter, 1);
          sheet.getRange(insertAfter + 1, 1, 1, m.newRow.length).setValues([m.newRow]);
          // Shift all disciplines whose last row is now at or beyond the inserted row
          Object.keys(discLastRow).forEach(function(dk) {
            if (discLastRow[dk] >= insertAfter + 1) discLastRow[dk]++;
          });
          discLastRow[m.targetDiscKey] = insertAfter + 1;
        } else {
          // Target discipline not yet in sheet — append at end
          var lastRow = sheet.getLastRow();
          sheet.getRange(lastRow + 1, 1, 1, m.newRow.length).setValues([m.newRow]);
          discLastRow[m.targetDiscKey] = lastRow + 1;
        }
        appended++;
      });
    }

    // ── 3. Append new rows ──────────────────────────────────────
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
  if (deleted)  msg += ', deleted ' + deleted  + ' task(s)';
  msg += ' in "' + SOURCE_SHEET + '".';
  return msg;
}

// ============================================================
//  GANTT TASK PARAMS TAB — read / write
//
//  Tab columns: TASKID | COLOR | TYPE | STYLE | SYMBOL | DEPS
//    TASKID — persistent integer task ID (from GANTT TASK IDS tab)
//    COLOR  — hex color override (#rrggbb) or empty (inherit section color)
//    TYPE   — bar / milestone / flag  (overrides MILESTONE checkbox)
//    STYLE  — solid / dashed / dashed-outline
//    SYMBOL — emoji or short text displayed inside milestone / flag
//    DEPS   — tilde-separated list of "DISCIPLINE|TASKNAME:S/E" predecessor keys
//
//  Keyed by TASKID so params survive task renames and discipline changes.
//  Tasks without a TASKID (newly created, not yet saved+loaded) are skipped.
//  This tab is created automatically on the first Save.
//  All rows are rewritten on every Save (no partial updates).
// ============================================================

// ============================================================
//  GANTT TASK IDS TAB — read / write
//
//  Separate from GANTT TASK PARAMS so that task IDs are NEVER
//  touched by writeTaskParams (which rewrites the params tab on
//  every Save). IDs are written here once by doGet and persist
//  indefinitely — only doGet and saveBackToTaskList (for renames)
//  ever modify this tab.
//
//  Tab columns: KEY | TASKID
// ============================================================
function readTaskIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_IDS_SHEET);
  var result = { ids: {}, maxId: 0 };
  if (!sh || sh.getLastRow() < 2) return result;
  var data = sh.getDataRange().getValues();
  // Find header row
  var hRow = -1;
  for (var i = 0; i < Math.min(data.length, 5); i++) {
    if (String(data[i][0]).trim().toUpperCase() === 'KEY') { hRow = i; break; }
  }
  if (hRow < 0) return result;
  var cols = {};
  data[hRow].forEach(function(h, idx) { cols[String(h).trim().toUpperCase()] = idx; });
  var colKey    = cols['KEY']    !== undefined ? cols['KEY']    : 0;
  var colTaskId = cols['TASKID'] !== undefined ? cols['TASKID'] : 1;
  for (var r = hRow + 1; r < data.length; r++) {
    var key = normKey(String(data[r][colKey] || '').trim());
    var id  = parseInt(String(data[r][colTaskId] || '').trim(), 10) || 0;
    if (key && id) {
      result.ids[key] = id;
      if (id > result.maxId) result.maxId = id;
    }
  }
  return result;
}

function writeTaskIds(taskIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_IDS_SHEET);
  if (!sh) sh = ss.insertSheet(TASK_IDS_SHEET);
  sh.clearContents();
  var rows = [['KEY', 'TASKID']];
  Object.keys(taskIds.ids).forEach(function(key) {
    if (taskIds.ids[key]) rows.push([key, taskIds.ids[key]]);
  });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.hideSheet();
}

function hideTaskIdsSheet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_IDS_SHEET);
  if (sh && !sh.isSheetHidden()) sh.hideSheet();
}

// ============================================================
//  GANTT TASK PARAMS TAB — read / write
//
//  Tab columns: KEY | COLOR | TYPE | STYLE | SYMBOL | DEPS
//  (No TASKID — IDs live in GANTT TASK IDS tab instead.)
//
//  This tab is created automatically on the first Save.
//  All rows are rewritten on every Save (no partial updates).
// ============================================================

// Return { byId: { taskId: params }, byKey: { normKey: params } }
//
// Three tab formats are handled transparently:
//   V1.20+ (new)   — headers: TASKID | KEY | COLOR | TYPE | STYLE | SYMBOL | DEPS
//   V1.20 (early)  — headers: TASKID | COLOR | TYPE | STYLE | SYMBOL | DEPS  (no KEY col)
//   Legacy (<V1.20) — headers: KEY | COLOR | TYPE | STYLE | SYMBOL | DEPS  (no TASKID col)
//
// doGet merges via: byId[t.taskId] || byKey[normKey(t.group+'|'+t.name)]
// taskId wins (rename-safe); key is the fallback if IDs were ever regenerated.
function readTaskParams() {
  var result = { byId: {}, byKey: {} };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_PARAMS_SHEET);
  if (!sh || sh.getLastRow() < 2) return result;

  var data = sh.getDataRange().getValues();

  // Locate header row — col A must be TASKID (new) or KEY (legacy).
  var hRow = -1;
  for (var i = 0; i < Math.min(data.length, 5); i++) {
    var col0 = String(data[i][0]).trim().toUpperCase();
    if (col0 === 'TASKID' || col0 === 'KEY') { hRow = i; break; }
  }
  if (hRow < 0) return result;

  var cols = {};
  data[hRow].forEach(function(h, idx) { cols[String(h).trim().toUpperCase()] = idx; });

  var colId     = cols['TASKID'] !== undefined ? cols['TASKID'] : -1;
  var colKey    = cols['KEY']    !== undefined ? cols['KEY']    : -1;
  // Column positions differ between old (no KEY) and new (KEY inserted at col 1) formats.
  // Header-name lookup handles both; the fallback indices cover the two known old schemas.
  var colColor  = cols['COLOR']  !== undefined ? cols['COLOR']  : (colId >= 0 ? 1 : 1);
  var colType   = cols['TYPE']   !== undefined ? cols['TYPE']   : (colId >= 0 ? 2 : 2);
  var colStyle  = cols['STYLE']  !== undefined ? cols['STYLE']  : (colId >= 0 ? 3 : 3);
  var colSymbol = cols['SYMBOL'] !== undefined ? cols['SYMBOL'] : (colId >= 0 ? 4 : 4);
  var colDeps   = cols['DEPS']   !== undefined ? cols['DEPS']   : (colId >= 0 ? 5 : 5);

  for (var r = hRow + 1; r < data.length; r++) {
    var params = {
      color:  String(data[r][colColor]  || '').trim(),
      type:   String(data[r][colType]   || '').trim().toLowerCase(),
      style:  String(data[r][colStyle]  || '').trim().toLowerCase(),
      symbol: String(data[r][colSymbol] || '').trim(),
      deps:   String(data[r][colDeps]   || '').trim()
    };

    if (colId >= 0) {
      var id = parseInt(String(data[r][colId] || '').trim(), 10);
      if (id) result.byId[id] = params;
    }
    if (colKey >= 0) {
      var k = normKey(String(data[r][colKey] || '').trim());
      if (k) result.byKey[k] = params;
    }
  }
  return result;
}

// Write all task display params to GANTT TASK PARAMS.
// Tab columns: TASKID | KEY | COLOR | TYPE | STYLE | SYMBOL | DEPS
// KEY = normKey(DISC|TASKNAME) — human-readable backup; also used as fallback on Load
// if task IDs are ever regenerated.
// Rows are built before touching the sheet — if no tasks have a persistent ID yet,
// the existing tab is left intact (no data-wipe on a failed save).
function writeTaskParams(tasks) {
  if (!tasks || !tasks.length) return;

  var rows = [['TASKID', 'KEY', 'COLOR', 'TYPE', 'STYLE', 'SYMBOL', 'DEPS']];

  tasks.forEach(function(t) {
    var taskId = parseInt(t.taskId, 10);
    if (!taskId) return; // new tasks without a persistent ID are skipped

    var key    = normKey((t.group || '') + '|' + (t.name || ''));
    // Save the effective displayed colour — override first, then status colour.
    // This locks in every visible bar colour so it persists across reloads.
    var color  = String(t.colorOverride || t.color || '').trim();
    var type   = String(t.type  || 'bar').trim().toLowerCase();
    var style;
    if (String(t.dashed) === 'true' || t.dashed === true) {
      style = 'dashed';
    } else if (String(t.dashedOutline) === 'true' || t.dashedOutline === true) {
      style = 'dashed-outline';
    } else {
      style = 'solid';
    }
    var symbol = String(t.symbol || '').trim();
    var deps   = String(t.dependencies || '').trim();

    rows.push([taskId, key, color, type, style, symbol, deps]);
  });

  if (rows.length < 2) return; // nothing to write — do NOT clear the existing tab

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_PARAMS_SHEET);
  if (!sh) sh = ss.insertSheet(TASK_PARAMS_SHEET);

  sh.clearContents();
  sh.getRange(1, 1, rows.length, 7).setValues(rows);
  sh.hideSheet();
}

// ============================================================
//  GANTT SETTINGS TAB — read / write
//
//  Fixed settings stored as KEY → VALUE rows.
//  Section colors stored as individual rows:
//    groupColor.DISCIPLINE_NAME → #hexcolor
//  readSettings() reconstructs groupColors JSON from those rows
//  so the HTML frontend receives the format applySettings() expects.
// ============================================================

var SETTINGS_KEYS = [
  'projectName', 'projectSubtitle', 'projectDate', 'projectNote',
  'userLabelWidth', 'ganttBarFontSize', 'showTodayLine', 'darkMode',
  'flatMode', 'barTextColor', 'collapsedGroups',
  'ganttCollapsedGroups', 'milestonesCollapsedGroups', 'flagsCollapsedGroups',
  'disciplineOrder', 'taskListOrder', 'showName', 'showPhase', 'showNote',
  'groupTint', 'taskRowTint', 'groupHeaderColor', 'groupHeaderTint', 'groupHeaderTextColor',
  'currentScale', 'zoomLevel', 'chartStart', 'chartEnd',
  'showDependencies', 'snapValue', 'toolbarCollapsed',
  'ganttRollupMarkersVisible',
  'matchHdrToGroupColor', 'showRollupTicks', 'showGroupBars',
  'showDateColumns', 'showDurationColumn',
  'flatTextWrap', 'flatBarHeight'
];

var SETTINGS_DESCRIPTIONS = {
  projectName:          'Project title shown in the chart header',
  projectSubtitle:      'Project subtitle shown in the chart header',
  projectDate:          'Updated date shown in the chart header (YYYY-MM-DD)',
  projectNote:          'Top note shown below the chart header',
  userLabelWidth:       'Width of the task-name label column (px)',
  ganttBarFontSize:     'Font size for Gantt bar labels (pt)',
  showTodayLine:        'Show vertical Today line on chart (true/false)',
  darkMode:             'Dark theme enabled (true/false)',
  flatMode:             'Flat mode — suppress group header rows (true/false)',
  barTextColor:         'Bar label text colour hex — empty string = auto-contrast',
  collapsedGroups:           'JSON object: which discipline groups are collapsed in the Task List tab',
  ganttCollapsedGroups:      'JSON object: which discipline groups are collapsed in the Gantt chart (independent of task list)',
  milestonesCollapsedGroups: 'JSON object: which discipline groups are collapsed in the Milestones tab',
  flagsCollapsedGroups:      'JSON object: which discipline groups are collapsed in the Flags tab',
  disciplineOrder:      'JSON array: discipline group display order on the Gantt chart',
  taskListOrder:        'JSON array: discipline group display order in the Task List section',
  showName:             'Show task name text inside Gantt bar (true/false)',
  showPhase:            'Show sub/phase label inside Gantt bar (true/false)',
  showNote:             'Show status/note text inside Gantt bar (true/false)',
  groupTint:            'Opacity of rollup-bar colour tint (0–100)',
  taskRowTint:          'Opacity of alternating task-row stripe in label column (0–100)',
  groupHeaderColor:     'Background fill colour for all group header rows (hex)',
  groupHeaderTint:      'Opacity of group header background fill (0–100)',
  groupHeaderTextColor: 'Text colour for group header discipline labels (hex)',
  currentScale:         'Timeline scale — weeks / months / quarters / years',
  zoomLevel:            'Horizontal zoom multiplier (integer, default 1)',
  chartStart:           'Chart date range start override (YYYY-MM-DD, empty = auto)',
  chartEnd:             'Chart date range end override (YYYY-MM-DD, empty = auto)',
  showDependencies:     'Show Finish-to-Start dependency arrows (true/false)',
  snapValue:            'Drag-snap interval in days (1 / 7 / 14 / 30)',
  toolbarCollapsed:     'Settings toolbar collapsed state (true/false)',
  ganttRollupMarkersVisible: 'Show milestone & flag markers on collapsed Gantt rollup bars (true/false)',
  matchHdrToGroupColor: 'Group header background uses the group bar colour when true (true/false)',
  showRollupTicks:      'Show vertical tick lines at rollup bar start/end (true/false)',
  showGroupBars:        'Show rollup bars on group header rows (true/false)',
  showDateColumns:      'Show START / END date columns beside the label column in the Gantt view (true/false)',
  showDurationColumn:   'Show DURATION column (days) beside END / label column in the Gantt view (true/false)',
  flatTextWrap:         'Wrap bar label text at word boundaries in flat mode (true/false)',
  flatBarHeight:        'Bar height in pixels in flat mode (12–72, default 24)'
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
    // Sheets may auto-parse date-looking strings back to Date objects; normalise to YYYY-MM-DD
    var raw = data[r][1];
    var v = (raw instanceof Date && !isNaN(raw.getTime())) ? fmtDate(raw) : String(raw).trim();
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

  var rows = [['SETTING KEY', 'VALUE', 'DESCRIPTION']];

  // Write the URL the HTML sent — this is exactly what the user configured.
  if (settings.appScriptURL) rows.push(['appScriptURL', settings.appScriptURL, 'Apps Script Web App URL (auto-populated on every Save)']);

  // Fixed chart-wide settings
  SETTINGS_KEYS.forEach(function(key) {
    rows.push([key,
               settings[key] !== undefined ? String(settings[key]) : '',
               SETTINGS_DESCRIPTIONS[key] || '']);
  });

  // Section colors — one row per discipline, sorted alphabetically
  if (settings.groupColors) {
    try {
      var gc = JSON.parse(settings.groupColors);
      Object.keys(gc).sort().forEach(function(disc) {
        rows.push(['groupColor.' + disc, gc[disc], 'Rollup-bar / separator accent colour for discipline: ' + disc]);
      });
    } catch(e) {}
  }

  sh.getRange(1, 1, rows.length, 3).setValues(rows);
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
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) { Logger.log('ERROR: "' + SOURCE_SHEET + '" tab not found'); return; }

  var raw = sheet.getDataRange().getValues();
  Logger.log('Total rows in data range: ' + raw.length);

  // Find header row
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow = i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) { Logger.log('ERROR: DISCIPLINE header not found in first 25 rows'); return; }
  Logger.log('Header row: ' + (hRow + 1) + ' (1-based)');

  var cols = {};
  raw[hRow].forEach(function(h, idx) { cols[String(h).trim().toUpperCase()] = idx; });
  Logger.log('Columns detected: ' + JSON.stringify(cols));

  var CI = {
    discipline: ci(cols,'DISCIPLINE',1), task: ci(cols,'TASK',3),
    start: ci(cols,'START DATE',7), end: ci(cols,'END DATE',8),
    schedule: ci(cols,'SCHEDULE',10), milestone: ci(cols,'MILESTONE',11)
  };
  Logger.log('Column indices → DISCIPLINE:' + CI.discipline + ' TASK:' + CI.task +
    ' START:' + CI.start + ' END:' + CI.end +
    ' SCHEDULE:' + CI.schedule + ' MILESTONE:' + CI.milestone);

  var loaded = 0, lastDisc = '';
  var groups = {};

  for (var row = hRow + 1; row < raw.length; row++) {
    var r = raw[row];
    var disc = String(r[CI.discipline] || '').trim();
    if (disc) lastDisc = disc;

    var taskName = String(r[CI.task] || '').trim();
    if (!taskName || /^[\s\-]+$/.test(taskName)) continue;

    var isSched = (r[CI.schedule] === true || String(r[CI.schedule]).toUpperCase() === 'TRUE');
    var isMil   = (r[CI.milestone] === true || String(r[CI.milestone]).toUpperCase() === 'TRUE');
    if (!isSched && !isMil) {
      Logger.log('  ROW ' + (row+1) + ' SKIPPED (no SCHEDULE/MILESTONE): "' + taskName + '" | SCHEDULE raw value: [' + r[CI.schedule] + ']');
      continue;
    }

    var sd = parseSheetDate(r[CI.start]);
    var ed = parseSheetDate(r[CI.end]);
    if (!sd || !ed) {
      Logger.log('  ROW ' + (row+1) + ' SKIPPED (bad dates): "' + taskName + '" | START:[' + r[CI.start] + '] END:[' + r[CI.end] + ']');
      continue;
    }
    if (ed.getTime() < sd.getTime()) {
      Logger.log('  ROW ' + (row+1) + ' SKIPPED (end before start): "' + taskName + '"');
      continue;
    }

    var grp = lastDisc || disc || 'UNCATEGORIZED';
    groups[grp] = (groups[grp] || 0) + 1;
    loaded++;
  }

  Logger.log('─── RESULT ──────────────────────────────');
  Logger.log('Tasks loaded: ' + loaded);
  Logger.log('Groups found:');
  Object.keys(groups).forEach(function(g) { Logger.log('  [' + g + '] — ' + groups[g] + ' task(s)'); });
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
