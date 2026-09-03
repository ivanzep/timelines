// ============================================================
//  TIMELINE SYNC  |  Google Apps Script
//  File: Code_V2.0.gs
//
//  PURPOSE
//  ───────
//  Two-way sync between the interactive Gantt HTML frontend
//  and a Google Sheet. Handles four categories of data:
//
//    1. PROJECT TASK LIST tab  — task schedule data
//       (discipline, task name, start/end dates, status)
//       Read on Load; dates + status written back on Save.
//       A project can have multiple "versions" of this tab (different
//       scenarios) — see VERSIONS below.
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
//    4. GANTT VERSIONS tab     — registry of task-list "versions"/scenarios
//       (version display name + its task sheet tab name). Read on Load so
//       the frontend can populate the Version dropdown.
//
//  VERSIONS
//  ────────
//  A "version" is just an alternate task-list tab (e.g. "PROJECT TASK LIST -
//  ACCELERATED"). Each version gets its OWN settings/task-params/task-IDs
//  tabs (suffixed with the version's task sheet name via _versionSuffixTab())
//  so edits to one version's chart settings or per-task colors never bleed
//  into another. The original default version (SOURCE_SHEET_DEFAULT =
//  'PROJECT TASK LIST') keeps the exact legacy unsuffixed tab names, so
//  existing single-version spreadsheets need no migration. doGet()/doPost()
//  resolve SOURCE_SHEET / SETTINGS_SHEET / TASK_PARAMS_SHEET / TASK_IDS_SHEET
//  for the requested version at the top of each call via _applyVersionTabs().
//
//  ENTRY POINTS
//  ────────────
//  doGet(e)   — called by HTML on Load. e.parameter.taskSheetName selects
//               which version to load (omit for the default version, or for
//               the very first load when the frontend doesn't know the last
//               active version yet — falls back to bootstrapping from the
//               default settings tab, same as pre-V1.28 behavior). Returns
//               tasks + task params + chart settings + the versions registry.
//  doPost(e)  — called by HTML on Save. Writes task schedule data, task
//               params, settings, and the versions registry independently so
//               a failure in one block never blocks the others. Creates the
//               target task-list tab (as a copy of payload.newVersionSourceSheet,
//               or the default version) if this is the first save into a
//               brand-new version.
//
//  VERSION HISTORY
//  ───────────────
//  V2.0  2026-09-03
//    • Baseline comparison. New GANTT BASELINE-DO NOT EDIT tab (one per
//      version, TASKID | KEY | TYPE | START | END) storing a frozen snapshot
//      of task dates, captured explicitly by the frontend (Set/Update/Clear
//      Baseline) — never touched by an ordinary task Save. readBaseline()/
//      writeBaseline() mirror readTaskParams()/writeTaskParams()'s taskId-
//      first, KEY-fallback matching convention, so a plain task rename
//      doesn't orphan its baseline entry. doGet() returns it as
//      result.baseline; doPost() writes it as step 6, independent try/catch
//      like every other save step. Two new SETTINGS_KEYS: showBaseline,
//      baselineCapturedAt (both automatically per-version, since
//      SETTINGS_SHEET is already version-suffixed). New BASELINE_SHEET
//      global + _applyVersionTabs() extension. testWriteBaseline() added
//      alongside the existing manual test functions.
//    • Cut as a new file pair (Code_V2.0.gs / TIMELINE-V2.0.html) rather than
//      extending V1.29 in place — V1.29 is now frozen for reference.
//
//  V1.29  2026-08-25
//    • PERSON column now round-trips. importFromTaskList() already READ the
//      PERSON column but dropped it before returning, so the value never
//      reached the frontend; it is now included on every emitted task.
//      saveBackToTaskList() gains a PERSON column mapping and writes the
//      value back (always written, so clearing an assignee sticks), wrapped
//      in its own try/catch like STATUS so a data-validation rejection on
//      one row can't abort the rest of the save. Multiple assignees are
//      stored as a plain comma-separated string — no parsing server-side.
//    • SETTINGS_KEYS gains kanbanBoards (JSON board definitions for the new
//      Kanban tab — frontend-only rendering state, stored per-version like
//      every other setting). Redeploy required for persistence.
//
//  V1.28  2026-08-18
//    • New multi-version task list support — see VERSIONS above. New tab-name
//      constant SOURCE_SHEET_DEFAULT and VERSIONS_SHEET; new helper
//      _versionSuffixTab() / _applyVersionTabs() / _ensureTaskSheetExists();
//      new readVersions() / writeVersions() reading/writing the GANTT
//      VERSIONS-DO NOT EDIT tab. doGet() accepts a taskSheetName query param
//      and returns result.versions + result.currentTaskSheetName. doPost()
//      ensures the target task-list tab exists (cloning payload.
//      newVersionSourceSheet when creating a brand-new version) before
//      saving, and writes payload.versions to the registry as an
//      independent, isolated write step (versionsError in the response).
//    • No changes to SETTINGS_KEYS/SETTINGS_DESCRIPTIONS, importFromTaskList,
//      or saveBackToTaskList's task-row logic — they already read SOURCE_SHEET
//      dynamically, so they work unchanged once it's resolved to the active
//      version. Redeploy required.
//    • doGet() now also returns result.spreadsheetUrl (SpreadsheetApp.
//      getUrl()) so the frontend's spreadsheet-name label can link directly
//      to the sheet. Redeploy required.
//    • _ensureTaskSheetExists() takes a new `blank` param — when true, the
//      cloned tab's task rows are cleared (via new _clearTaskRows()) right
//      after copying, so a "start blank" new version has no tasks of its
//      own instead of inheriting the source's. doPost() reads this from
//      payload.newVersionBlank.
//    • New _deleteVersionTabs(taskSheetName) — deletes a version's task-list
//      tab plus its settings/task-params/task-IDs tabs (no-op for the
//      default version). doPost() runs this as an independent step 0 when
//      payload.deleteVersionTaskSheetName is present (deleteVersionError in
//      the response), and skips _ensureTaskSheetExists for delete-only
//      requests that carry no payload.tasks/payload.settings. Redeploy
//      required.
//    • Fix: task color/type/style/symbol settings were silently dropped on
//      Save for any task without a persistent taskId yet — which is every
//      brand-new task (taskId is normally only ever assigned by doGet/Load),
//      most noticeably right after creating a new version and adding tasks
//      to it before the first reload. writeTaskParams() requires a taskId to
//      key each row on and was simply skipping tasks that didn't have one
//      yet. doPost() now assigns persistent IDs to any task missing one
//      (same GANTT TASK IDS lookup/increment logic doGet already uses)
//      immediately before writeTaskParams() runs, so nothing gets skipped.
//    • _ensureTaskSheetExists() now also clones the source version's GANTT
//      TASK IDS tab (when not blank) so a new version's tasks keep the same
//      persistent IDs their params were written with, instead of getting
//      renumbered on the new version's first Load. readTaskParams()'s KEY
//      fallback would still recover the params even without this, but this
//      avoids relying on that fallback for a case that happens on every
//      single new version. Redeploy required.
//    • SETTINGS_KEYS gains flatShowFlags (Flat mode "Flags" visibility
//      toggle — frontend-only rendering change, no other backend logic
//      involved). Redeploy required for persistence.
//    • SETTINGS_KEYS gains flatFlagOffset (Flat mode flag pole/icon offset
//      above its group's bar, px — frontend-only rendering change, no other
//      backend logic involved). Redeploy required for persistence.
//    • SETTINGS_KEYS gains flatShowNotes ("Notes" bar-label toggle, Bar
//      Labels section — draws the task's NOTES column text as an extra bar
//      label in both Flat and normal mode; frontend-only rendering change,
//      no other backend logic involved). Redeploy required for
//      persistence.
//
//  V1.27  2026-08-10
//    • saveBackToTaskList() — per-task write isolation: each task's row writes
//      are wrapped in an individual try/catch so one failed row no longer aborts
//      updates for all subsequent tasks in the Save payload.
//    • STATUS write has its own inner try/catch so a Google Sheets data-validation
//      rejection on the STATUS cell is logged (Logger.log) and skipped without
//      stopping the rest of that task's writes or any following tasks.
//    • STATUS always written unconditionally (was previously skipped when t.note
//      was falsy, preventing the status from being cleared to empty string).
//    • No new SETTINGS_KEYS. No redeploy needed for other logic; redeploy required
//      to get the more robust write behaviour.
//
//  V1.26  2026-07-07
//    • SETTINGS_KEYS expanded with 5 new keys:
//        taskSheetName        — user-configurable task sheet tab name
//        sortColumn           — task-table sort column (field name or empty)
//        sortDirection        — task-table sort direction (1 = asc, -1 = desc)
//        currentTab           — active view tab (gantt | tasks | milestones | flags)
//        taskListGroupSortMode — Task Properties group ordering mode (date | alpha | custom)
//    • SETTINGS_DESCRIPTIONS updated for all 5 new keys.
//    • doGet() reads settings before importing tasks so taskSheetName is applied
//      on every Load and Save.
//    • No other logic changes. Redeploy to activate persistence for these keys.
//
//  V1.25  2026-07-03
//    • SETTINGS_KEYS expanded with 3 keys that were collected by the HTML but
//      never written to the sheet: todayLineColor, metaDetailsCollapsed, statusColors.
//    • SETTINGS_DESCRIPTIONS updated for all 3 new keys.
//    • No other logic changes. Redeploy to activate settings persistence for these keys.
//
//  V1.24  2026-07-02
//    • No backend changes. Version bump to accompany TIMELINE-V1.24.html (Task
//      Properties table restyled as Excel-like grid; removed bubble-style inputs).
//
//  V1.23  2026-07-02
//    • SETTINGS_KEYS expanded with 3 new keys introduced in HTML V1.23:
//        flatLabelOverflow — extra px the outside flat-mode label clip extends beyond bar right edge
//        flatLabelsOutside — force all flat-mode labels outside bars (above/below)
//        useStatusColors   — flip bar/milestone/flag fills to STATUS_COLOR_MAP[task.note]
//    • SETTINGS_DESCRIPTIONS updated for all 3 new keys.
//    • No other logic changes. No redeploy required unless you want settings persistence.
//
//  V1.21  2026-07-01 22:33
//    • No backend logic changes. Version bump to accompany TIMELINE-V1.22.html (flat mode
//      labels moved above bars). Backend unchanged; redeploy not required.
//
//  V1.20  2026-06-25 21:50
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
// SOURCE_SHEET/SETTINGS_SHEET/TASK_PARAMS_SHEET/TASK_IDS_SHEET are all re-resolved
// at the top of doGet()/doPost() (via _applyVersionTabs) for whichever "version"
// (task-list tab) the current request targets. Every function below just reads
// these globals at call time, so nothing else needs to know versions exist.
var SOURCE_SHEET_DEFAULT = 'PROJECT TASK LIST'; // the original/default version — keeps unsuffixed legacy tab names
var VERSIONS_SHEET       = 'GANTT VERSIONS-DO NOT EDIT'; // registry of {version name, task sheet name} — not itself versioned

var SOURCE_SHEET      = SOURCE_SHEET_DEFAULT;            // master task data — never reformatted; overridable via taskSheetName setting
var TASK_PARAMS_SHEET = 'GANTT TASK PARAMS-DO NOT EDIT'; // per-task Gantt display params (color, type, style, symbol, deps)
var TASK_IDS_SHEET    = 'GANTT TASK IDS-DO NOT EDIT';    // persistent numeric task ID registry — never cleared by saves
var SETTINGS_SHEET    = 'GANTT SETTINGS-DO NOT EDIT';    // chart-wide UI settings
var BASELINE_SHEET    = 'GANTT BASELINE-DO NOT EDIT';    // frozen baseline date snapshot (V2.0)

// Resolve the version-specific tab name for a given base tab name. The default
// version (SOURCE_SHEET_DEFAULT) always keeps the exact unsuffixed base name, so
// existing single-version spreadsheets keep working with zero migration.
function _versionSuffixTab(baseName, taskSheetName) {
  if (!taskSheetName || normKey(taskSheetName) === normKey(SOURCE_SHEET_DEFAULT)) return baseName;
  return baseName + ' [' + taskSheetName + ']';
}

// Point SETTINGS_SHEET / TASK_PARAMS_SHEET / TASK_IDS_SHEET / BASELINE_SHEET at
// the tabs for the given task-list version. Call this immediately after
// resolving SOURCE_SHEET.
function _applyVersionTabs(taskSheetName) {
  SETTINGS_SHEET    = _versionSuffixTab('GANTT SETTINGS-DO NOT EDIT', taskSheetName);
  TASK_PARAMS_SHEET = _versionSuffixTab('GANTT TASK PARAMS-DO NOT EDIT', taskSheetName);
  TASK_IDS_SHEET    = _versionSuffixTab('GANTT TASK IDS-DO NOT EDIT', taskSheetName);
  BASELINE_SHEET     = _versionSuffixTab('GANTT BASELINE-DO NOT EDIT', taskSheetName);
}

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
    // Resolve which task-list "version" to load. If the frontend already knows
    // (normal reload / explicit version switch), it passes ?taskSheetName=...
    // and we go straight to that version's own settings tab. Otherwise (very
    // first load, frontend has no saved state yet) fall back to the pre-V1.28
    // bootstrap: read the DEFAULT settings tab to discover the last-active
    // taskSheetName, then re-read from that version's own settings tab.
    var requestedSheet = (e && e.parameter && e.parameter.taskSheetName) ? String(e.parameter.taskSheetName).trim() : '';
    var savedSettings;
    if (requestedSheet) {
      SOURCE_SHEET = requestedSheet;
      _applyVersionTabs(SOURCE_SHEET);
      savedSettings = readSettings() || {};
      savedSettings.taskSheetName = SOURCE_SHEET;
    } else {
      savedSettings = readSettings();
      if (savedSettings && savedSettings.taskSheetName) {
        SOURCE_SHEET = savedSettings.taskSheetName;
        _applyVersionTabs(SOURCE_SHEET);
        if (SOURCE_SHEET !== SOURCE_SHEET_DEFAULT) {
          var versionSettings = readSettings();
          if (versionSettings) { versionSettings.taskSheetName = SOURCE_SHEET; savedSettings = versionSettings; }
        }
      }
    }

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

    result.settings = savedSettings;
    // Build statusColors from STATUS_COLORS + any custom values found in the task list.
    // This ensures statuses typed directly in the sheet (not in STATUS_COLORS) reach the frontend.
    var allStatusColors = {};
    Object.keys(STATUS_COLORS).forEach(function(k) { allStatusColors[k] = STATUS_COLORS[k]; });
    result.tasks.forEach(function(t) {
      var s = String(t.note || '').split('·')[0].trim().toUpperCase();
      if (s && !allStatusColors[s]) allStatusColors[s] = '#64748b'; // slate default for unknown statuses
    });
    result.statusColors = allStatusColors;
    try {
      var activeSs = SpreadsheetApp.getActiveSpreadsheet();
      result.spreadsheetName = activeSs.getName();
      result.spreadsheetUrl  = activeSs.getUrl();
    } catch(e) {}
    try { result.versions = readVersions(); } catch(e) { result.versions = [{ name: 'Default', taskSheetName: SOURCE_SHEET_DEFAULT }]; }
    try { result.baseline = readBaseline(); } catch(e) { result.baseline = { byId: {}, byKey: {} }; }
    result.currentTaskSheetName = SOURCE_SHEET;
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

    // Apply taskSheetName from incoming settings before writing tasks —
    // saveBackToTaskList uses SOURCE_SHEET to find the tab. Resolve this
    // version's own settings/task-params/task-IDs tabs at the same time.
    SOURCE_SHEET = (payload.settings && payload.settings.taskSheetName) ? payload.settings.taskSheetName : SOURCE_SHEET_DEFAULT;
    _applyVersionTabs(SOURCE_SHEET);

    // 0. Delete a version's tabs (independent of whatever else this request saves —
    // used by the Version dropdown's delete action, which sends a minimal payload).
    var deleteVersionError = '';
    if (payload.deleteVersionTaskSheetName) {
      try {
        _deleteVersionTabs(payload.deleteVersionTaskSheetName);
      } catch (delErr) {
        deleteVersionError = delErr.toString();
        Logger.log('_deleteVersionTabs error: ' + delErr);
      }
    }

    // If this is the first save into a brand-new version, its task-list tab
    // doesn't exist yet — create it now (as a snapshot copy of the version the
    // frontend branched from, or the default version) before writing to it.
    // Skipped for delete-only requests (no task/settings payload to write).
    if (payload.tasks || payload.settings) {
      try { _ensureTaskSheetExists(SOURCE_SHEET, payload.newVersionSourceSheet, !!payload.newVersionBlank); } catch (ensureErr) {
        Logger.log('_ensureTaskSheetExists error: ' + ensureErr);
      }
    }

    // 1. Update dates + status in the configured task sheet
    try {
      taskMsg = saveBackToTaskList(payload);
    } catch (tErr) {
      taskErr = tErr.toString();
    }

    // 2. Write per-task Gantt params to GANTT TASK PARAMS
    if (payload.tasks) {
      try {
        // writeTaskParams() requires a persistent taskId to key each row on, but
        // taskId is normally only ever assigned during doGet (Load) — a brand-new
        // task saved for the first time (e.g. right after "+ Add Task", or any task
        // in a freshly created version before its first Load) still has taskId=null
        // at this point, so its color/type/style/symbol would otherwise be silently
        // skipped. Assign persistent IDs here too, same as doGet, before writing.
        var taskIdsForSave = readTaskIds();
        var taskIdsChanged = false;
        payload.tasks.forEach(function(t) {
          if (t.taskId) return;
          var pKey = normKey((t.group || '') + '|' + (t.name || ''));
          if (taskIdsForSave.ids[pKey]) {
            t.taskId = taskIdsForSave.ids[pKey];
          } else {
            taskIdsForSave.maxId++;
            t.taskId = taskIdsForSave.maxId;
            taskIdsForSave.ids[pKey] = t.taskId;
            taskIdsChanged = true;
          }
        });
        if (taskIdsChanged) {
          try { writeTaskIds(taskIdsForSave); } catch (idErr) { Logger.log('writeTaskIds error: ' + idErr); }
        }

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

    // 4. Write / delete group separator rows (handles empty groups + deletions)
    var groupSepError = '';
    try {
      writeGroupSeparators(payload);
    } catch (gsErr) {
      groupSepError = gsErr.toString();
      Logger.log('writeGroupSeparators error: ' + gsErr);
    }

    // 5. Sync the versions registry (independent of which version this save targets)
    var versionsError = '';
    if (payload.versions) {
      try {
        writeVersions(payload.versions);
      } catch (vErr) {
        versionsError = vErr.toString();
        Logger.log('writeVersions error: ' + vErr);
      }
    }

    // 6. Write the frozen baseline snapshot (V2.0) — only present when the
    // frontend is explicitly capturing/updating/clearing one (Set/Update/Clear
    // Baseline). An ordinary task Save omits payload.baseline entirely, so
    // writeBaseline()'s own "skip if undefined" guard means routine edits
    // never touch this tab.
    var baselineError = '';
    if (payload.baseline !== undefined) {
      try {
        writeBaseline(payload.baseline, !!payload.baselineClear);
      } catch (blErr) {
        baselineError = blErr.toString();
        Logger.log('writeBaseline error: ' + blErr);
      }
    }

    return buildResponse({
      success:            true,
      message:            taskMsg,
      taskError:          taskErr,
      taskParamsError:    taskParamsErr,
      settingsError:      settingsErr,
      groupSepError:      groupSepError,
      versionsError:      versionsError,
      deleteVersionError: deleteVersionError,
      baselineError:      baselineError
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
//  VERSIONS  — registry of task-list "versions" (alternate scenarios)
//
//  GANTT VERSIONS-DO NOT EDIT tab columns: VERSION NAME | TASK SHEET NAME.
//  The default version (SOURCE_SHEET_DEFAULT) is never stored as a row here —
//  it's always implicitly available so a spreadsheet with no versions tab yet
//  still returns exactly one ("Default") version.
// ============================================================
function readVersions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var versions = [{ name: 'Default', taskSheetName: SOURCE_SHEET_DEFAULT }];
  var sh = ss.getSheetByName(VERSIONS_SHEET);
  if (!sh || sh.getLastRow() < 2) return versions;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var seen = {};
  seen[normKey(SOURCE_SHEET_DEFAULT)] = true;
  data.forEach(function(row) {
    var name = String(row[0] || '').trim();
    var tsn  = String(row[1] || '').trim();
    if (!name || !tsn) return;
    var k = normKey(tsn);
    if (seen[k]) return; // skip duplicate / re-registered default
    versions.push({ name: name, taskSheetName: tsn });
    seen[k] = true;
  });
  return versions;
}

// Fully rewrites the versions registry from the array the frontend sends on
// every Save (its complete current version list). The default version is
// never written — it's implicit. Creates the tab on first use.
function writeVersions(versions) {
  if (!versions || !versions.length) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(VERSIONS_SHEET);
  if (!sh) sh = ss.insertSheet(VERSIONS_SHEET);
  sh.clearContents();

  var rows = [['VERSION NAME', 'TASK SHEET NAME']];
  var seen = {};
  versions.forEach(function(v) {
    var name = String((v && v.name) || '').trim();
    var tsn  = String((v && v.taskSheetName) || '').trim();
    if (!name || !tsn) return;
    var k = normKey(tsn);
    if (k === normKey(SOURCE_SHEET_DEFAULT) || seen[k]) return;
    rows.push([name, tsn]);
    seen[k] = true;
  });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.hideSheet();
}

// Creates the task-list tab for a brand-new version if it doesn't already
// exist, by copying an existing tab (the version the frontend branched from,
// falling back to the default version) so the new version starts as a
// snapshot of its source — header row, data validation, and current tasks
// all carry over, ready to edit/add on top of. When blank=true, the cloned
// tab's task rows are cleared (header/formatting/validation stay intact) so
// the version starts genuinely empty instead of as a copy of the source.
//
// Also clones the source version's TASK IDS tab (when not blank) so the new
// version's tasks keep the same persistent IDs their GANTT TASK PARAMS rows
// were written with, instead of getting renumbered on the new version's
// first Load. readTaskParams()'s KEY fallback would still recover color/
// type/style/symbol even without this (it doesn't require IDs to match),
// but preserving IDs avoids relying on that fallback at all.
function _ensureTaskSheetExists(taskSheetName, sourceSheetName, blank) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(taskSheetName)) return; // already exists — nothing to do

  var srcName = sourceSheetName || SOURCE_SHEET_DEFAULT;
  var src = ss.getSheetByName(srcName) || ss.getSheetByName(SOURCE_SHEET_DEFAULT);
  if (src) {
    var copy = src.copyTo(ss);
    copy.setName(taskSheetName);
    try { ss.setActiveSheet(copy); ss.moveActiveSheet(src.getIndex() + 1); } catch (moveErr) {}
    if (blank) {
      _clearTaskRows(copy);
    } else {
      try {
        var srcIdsSheet = ss.getSheetByName(_versionSuffixTab('GANTT TASK IDS-DO NOT EDIT', srcName));
        if (srcIdsSheet) {
          var idsCopy = srcIdsSheet.copyTo(ss);
          idsCopy.setName(_versionSuffixTab('GANTT TASK IDS-DO NOT EDIT', taskSheetName));
          try { ss.setActiveSheet(idsCopy); ss.moveActiveSheet(copy.getIndex() + 1); } catch (moveErr2) {}
        }
      } catch (idsCloneErr) {
        Logger.log('Task IDs clone error: ' + idsCloneErr);
      }
    }
  } else {
    ss.insertSheet(taskSheetName); // fallback: no source tab found — blank tab
  }
}

// Clears all task rows below the header (keeping header text, formatting, and
// data validation intact) — used to make a freshly cloned tab start empty.
function _clearTaskRows(sheet) {
  var raw = sheet.getDataRange().getValues();
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow = i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) return;
  var lastRow = sheet.getLastRow();
  if (lastRow > hRow + 1) {
    sheet.getRange(hRow + 2, 1, lastRow - hRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// Deletes a version's task-list tab and its settings/task-params/task-IDs
// tabs. The default version can never be deleted this way.
function _deleteVersionTabs(taskSheetName) {
  if (!taskSheetName || normKey(taskSheetName) === normKey(SOURCE_SHEET_DEFAULT)) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [
    taskSheetName,
    _versionSuffixTab('GANTT SETTINGS-DO NOT EDIT', taskSheetName),
    _versionSuffixTab('GANTT TASK PARAMS-DO NOT EDIT', taskSheetName),
    _versionSuffixTab('GANTT TASK IDS-DO NOT EDIT', taskSheetName)
  ];
  names.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    if (sh) ss.deleteSheet(sh);
  });
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
      // PERSON was read above but never emitted before V1.29, so assignments
      // never reached the frontend. Passed through verbatim — multiple
      // assignees are just a comma-separated string, split client-side.
      person:        person,
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
  var hasTasks        = payload.tasks        && payload.tasks.length        > 0;
  var hasDeletedGroups = payload.deletedGroups && payload.deletedGroups.length > 0;
  if (!hasTasks && !hasDeletedGroups) return 'Nothing to save.';

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
    person:     ci(cols, 'PERSON',      5),
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
  var discSepRow  = {}; // normKey(disc) → 1-based row of existing separator ("-") row
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
    } else if (task && /^[\s\-]+$/.test(task) && lastDisc) {
      // Separator / divider row ("-") — track so we know not to add a duplicate
      if (!discSepRow[normKey(lastDisc)]) discSepRow[normKey(lastDisc)] = row + 1;
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
              if (moveSc) moveRow[CI.status] = moveSc;
            }
            if (t.notes) moveRow[CI.notes] = t.notes;
            if (t.person) moveRow[CI.person] = String(t.person).trim();
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
        try {
          if (isRename) {
            // Name-only rename: overwrite discipline + task name in the existing row
            sheet.getRange(sheetRow, CI.discipline + 1).setValue(discipline);
            sheet.getRange(sheetRow, CI.task       + 1).setValue(taskName);
          }
          if (t.start) sheet.getRange(sheetRow, CI.start + 1).setValue(t.start);
          if (t.end)   sheet.getRange(sheetRow, CI.end   + 1).setValue(t.end);
          // Always write STATUS — allows clearing it when set to empty string.
          // Wrapped in its own try so a data-validation rejection on the STATUS cell
          // does not abort writes for this task or any subsequent tasks.
          try {
            var statusCandidate = String(t.note || '').split('·')[0].trim().toUpperCase();
            sheet.getRange(sheetRow, CI.status + 1).setValue(statusCandidate);
          } catch(statusErr) {
            // Log but continue — the status cell may have strict validation
            Logger.log('STATUS write failed for row ' + sheetRow + ': ' + statusErr);
          }
          if (typeof t.notes !== 'undefined') sheet.getRange(sheetRow, CI.notes + 1).setValue(t.notes);
          // Always write PERSON (like STATUS) so clearing an assignee sticks.
          // Own try/catch: the PERSON cell may carry data validation, and a
          // rejection there must not abort this row or any later task.
          if (typeof t.person !== 'undefined') {
            try {
              sheet.getRange(sheetRow, CI.person + 1).setValue(String(t.person || '').trim());
            } catch(personErr) {
              Logger.log('PERSON write failed for row ' + sheetRow + ': ' + personErr);
            }
          }
          sheet.getRange(sheetRow, CI.schedule  + 1).setValue(true);
          // Write MILESTONE column based on current type — clears it when task is no longer a milestone
          sheet.getRange(sheetRow, CI.milestone + 1).setValue(t.type === 'milestone');
          updated++;
        } catch(rowErr) {
          Logger.log('Row write failed for row ' + sheetRow + ' (' + discipline + '|' + taskName + '): ' + rowErr);
        }

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
          if (sc) newRow[CI.status] = sc;
        }
        if (t.notes) newRow[CI.notes] = t.notes;
        if (t.person) newRow[CI.person] = String(t.person).trim();

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
    // Rows in the lookup that are NOT in the payload were deleted by the user.
    // Also delete separator rows for groups removed from the chart.
    // Process all deletions together in descending order so indices stay valid.
    var rowsToDelete = [];
    Object.keys(lookup).forEach(function(key) {
      if (!payloadKeys[key]) rowsToDelete.push(lookup[key]);
    });
    if (payload.deletedGroups && payload.deletedGroups.length) {
      payload.deletedGroups.forEach(function(g) {
        var sr = discSepRow[normKey(g)];
        if (sr) rowsToDelete.push(sr);
      });
    }
    rowsToDelete.sort(function(a, b) { return b - a; }); // descending
    rowsToDelete.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
      deleted++;
      Object.keys(discLastRow).forEach(function(dk) {
        if (discLastRow[dk] >= rowNum) discLastRow[dk]--;
      });
      Object.keys(discSepRow).forEach(function(dk) {
        if (discSepRow[dk] >= rowNum) discSepRow[dk]--;
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
          // Target discipline not yet in sheet — write separator first, then task
          var lastRow = findLastContentRow(sheet, CI.discipline + 1, CI.task + 1);
          if (!discSepRow[m.targetDiscKey]) {
            var mSepRow = [];
            for (var _msi = 0; _msi < m.newRow.length; _msi++) mSepRow.push('');
            mSepRow[CI.discipline] = m.newRow[CI.discipline];
            mSepRow[CI.task]       = '-';
            sheet.getRange(lastRow + 1, 1, 1, mSepRow.length).setValues([mSepRow]);
            discSepRow[m.targetDiscKey] = lastRow + 1;
            lastRow++;
          }
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
      // Group orphaned rows by discipline so we can write one separator per discipline.
      var orphansByDisc = {};
      var orphanDiscOrder = [];
      newRowsOrphaned.forEach(function(rowArr) {
        var dk = normKey(String(rowArr[CI.discipline] || '').trim());
        if (!orphansByDisc[dk]) { orphansByDisc[dk] = []; orphanDiscOrder.push(dk); }
        orphansByDisc[dk].push(rowArr);
      });

      var lastRow = findLastContentRow(sheet, CI.discipline + 1, CI.task + 1);
      orphanDiscOrder.forEach(function(dk) {
        var rows = orphansByDisc[dk];
        // Write separator row first if none exists for this discipline
        if (!discSepRow[dk]) {
          var sepRow = [];
          for (var _si = 0; _si < rows[0].length; _si++) sepRow.push('');
          sepRow[CI.discipline] = String(rows[0][CI.discipline] || '').trim();
          sepRow[CI.task]       = '-';
          sheet.getRange(lastRow + 1, 1, 1, sepRow.length).setValues([sepRow]);
          discSepRow[dk] = lastRow + 1;
          lastRow++;
        }
        sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
        lastRow += rows.length;
      });
    }

    // ── 5. Separator rows for empty groups ─────────────────────
    // Any group in payload.groups that still has no separator row in the sheet
    // (e.g. a group created via Add Group with no tasks assigned yet).
    if (payload.groups && payload.groups.length) {
      var totalColsG = raw[hRow].length || 21;
      payload.groups.forEach(function(g) {
        var gkey = normKey(g);
        if (!discSepRow[gkey]) {
          var gLastRow = findLastContentRow(sheet, CI.discipline + 1, CI.task + 1);
          var gSepRow = [];
          for (var _gi = 0; _gi < totalColsG; _gi++) gSepRow.push('');
          gSepRow[CI.discipline] = g.toUpperCase();
          gSepRow[CI.task]       = '-';
          sheet.getRange(gLastRow + 1, 1, 1, gSepRow.length).setValues([gSepRow]);
          discSepRow[gkey] = gLastRow + 1;
        }
      });
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
//  GROUP SEPARATOR ROWS — write / delete
//
//  Groups/sections appear in the PROJECT TASK LIST tab as rows with
//  DISCIPLINE = <group name> and TASK = "-". These rows are skipped
//  on import but preserved visually in the spreadsheet. This function
//  adds separator rows for new groups and removes them for deleted ones.
// ============================================================
function writeGroupSeparators(payload) {
  var groups = (payload.groups || []).slice();
  if (!groups.length) return; // separator deletions handled in saveBackToTaskList

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) return;

  // ── Helper: scan raw data → sepRows and firstDataRow maps ──
  function scanSheet(rawData, hRowIdx, discIdx, taskIdx) {
    var sep  = {};  // normKey(disc) → 1-based row of the separator ("–" row)
    var first = {}; // normKey(disc) → 1-based row of the first non-separator data row
    for (var r = hRowIdx + 1; r < rawData.length; r++) {
      var d = String(rawData[r][discIdx] || '').trim();
      var t = String(rawData[r][taskIdx] || '').trim();
      if (!d) continue;
      var dk = normKey(d);
      if (/^[\s\-]+$/.test(t)) {
        if (!sep[dk])  sep[dk]  = r + 1;
      } else if (t) {
        if (!first[dk]) first[dk] = r + 1;
      }
    }
    return { sep: sep, first: first };
  }

  // ── Find header row ──
  var raw  = sheet.getDataRange().getValues();
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length, 10); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow = i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) return;

  var cols = {};
  raw[hRow].forEach(function(h, idx) { cols[String(h).trim().toUpperCase()] = idx; });
  var discIdx  = ci(cols, 'DISCIPLINE', 1);
  var taskIdx  = ci(cols, 'TASK',       3);
  var totalCols = raw[hRow].length || 21;

  // Separator row deletions for removed groups are handled inside saveBackToTaskList
  // (step 2), which already has the discSepRow map from its initial sheet scan and
  // processes all row deletions in one descending pass. We only handle insertion here.
  var maps = scanSheet(raw, hRow, discIdx, taskIdx);

  // ── Insert separator rows for new groups ──
  // A separator row has DISCIPLINE = <group name>, TASK = "-", all other cells blank.
  var newGroups = groups.filter(function(g) { return !maps.sep[normKey(g)]; });

  // Build a blank row template (loop, no Array.fill for GAS compatibility).
  function blankRow() {
    var r = [];
    for (var _i = 0; _i < totalCols; _i++) r.push('');
    return r;
  }

  // Compute the last content row from raw data (avoids counting checkbox-only empty rows).
  var contentLastRow = 0;
  for (var _cr = raw.length - 1; _cr >= 0; _cr--) {
    if (String(raw[_cr][discIdx] || '').trim() || String(raw[_cr][taskIdx] || '').trim()) {
      contentLastRow = _cr + 1; // 1-based
      break;
    }
  }
  if (!contentLastRow) contentLastRow = sheet.getLastRow();

  // Sort descending by insertion row so each insert doesn't shift the others.
  // Groups whose discipline already has data rows get inserted ABOVE that first data row.
  // Groups with no data rows yet get appended at the end (largest row number = last).
  newGroups.sort(function(a, b) {
    var ra = maps.first[normKey(a)] || (contentLastRow + 1);
    var rb = maps.first[normKey(b)] || (contentLastRow + 1);
    return rb - ra; // descending
  });

  newGroups.forEach(function(g) {
    var sepRow = blankRow();
    sepRow[discIdx] = g.toUpperCase();
    sepRow[taskIdx] = '-';

    var insertAt = maps.first[normKey(g)]; // 1-based row of first data row for this group
    if (insertAt) {
      // Insert a blank row above the first data row, then write the separator into it.
      sheet.insertRowBefore(insertAt);
      sheet.getRange(insertAt, 1, 1, totalCols).setValues([sepRow]);
    } else {
      // No existing data rows for this group — write after the last content row.
      sheet.getRange(contentLastRow + 1, 1, 1, totalCols).setValues([sepRow]);
      contentLastRow++;
    }
  });
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
//  GANTT BASELINE TAB — read / write (V2.0)
// ============================================================
// A frozen snapshot of task dates the user explicitly captures, compared
// against the live schedule to show drift. One tab per version (via
// _versionSuffixTab), never touched by an ordinary task Save — only written
// when the frontend sends payload.baseline (Set/Update/Clear Baseline).
//
// Tab columns: TASKID | KEY | TYPE | START | END
// Same TASKID-first, KEY-fallback matching convention as GANTT TASK PARAMS —
// KEY = normKey(DISC|TASKNAME) — so a plain rename doesn't orphan the ghost.
function readBaseline() {
  var result = { byId: {}, byKey: {} };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BASELINE_SHEET);
  if (!sh || sh.getLastRow() < 2) return result;

  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var entry = {
      type:  String(data[r][2] || 'bar').trim().toLowerCase(),
      start: String(data[r][3] || '').trim(),
      end:   String(data[r][4] || '').trim()
    };
    if (!entry.start) continue;

    var id = parseInt(String(data[r][0] || '').trim(), 10);
    if (id) result.byId[id] = entry;

    var k = normKey(String(data[r][1] || '').trim());
    if (k) result.byKey[k] = entry;
  }
  return result;
}

// Fully rewrites GANTT BASELINE from a frontend-supplied array of
// {taskId, group, name, type, start, end}. Skips entirely (leaves the
// existing tab untouched) when baselineTasks is undefined — an ordinary Save
// never sends this field, so it must never wipe a captured baseline. An
// explicit Clear Baseline sends an empty array WITH isClear=true, which is
// the only way to actually empty the tab.
function writeBaseline(baselineTasks, isClear) {
  if (!baselineTasks) return;
  if (!baselineTasks.length && !isClear) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BASELINE_SHEET);
  if (!sh) sh = ss.insertSheet(BASELINE_SHEET);
  sh.clearContents();

  if (!baselineTasks.length) { sh.hideSheet(); return; } // cleared — leave an empty, hidden tab

  var rows = [['TASKID', 'KEY', 'TYPE', 'START', 'END']];
  baselineTasks.forEach(function(t) {
    var taskId = parseInt(t.taskId, 10) || '';
    var key    = normKey((t.group || '') + '|' + (t.name || ''));
    var type   = String(t.type || 'bar').trim().toLowerCase();
    rows.push([taskId, key, type, t.start || '', t.end || '']);
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
  'flatTextWrap', 'flatBarHeight',
  'flatLabelOverflow', 'flatLabelsOutside', 'flatShowFlags', 'flatFlagOffset', 'flatShowNotes', 'useStatusColors',
  'todayLineColor', 'depsArrowColor', 'metaDetailsCollapsed', 'statusColors', 'kanbanBoards',
  'kanbanTextSize',
  'taskSheetName',
  'sortColumn', 'sortDirection', 'currentTab',
  'tabOrder', 'tabVisible',
  'showBaseline', 'baselineCapturedAt',
  'taskListGroupSortMode',
  'taskTableColWidths', 'taskTableColVisible',
  'taskListUseUniformRowColor', 'taskListUniformRowColor', 'taskListRowColorTint',
  'taskListGroupTextColor', 'taskListWrapText', 'taskPrintHeaderColor'
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
  flatBarHeight:        'Bar height in pixels in flat mode (12–72, default 24)',
  flatLabelOverflow:    'Extra px the outside flat-mode label clip extends beyond bar right edge (0–400, default 0)',
  flatLabelsOutside:    'Force all flat-mode labels outside bars (above/below) even when they would fit inside (true/false)',
  flatShowFlags:        'Show flag markers while Flat mode is on (true/false) — unrelated to the Flags tab filter',
  flatFlagOffset:       'Px the flag pole/triangle extends above its group bar top in flat mode; pole bottom is pinned to the bar bottom (0–40, default 8)',
  flatShowNotes:        'Show each task\'s NOTES column text as an extra bar label, in both Flat and normal mode (true/false)',
  useStatusColors:      'Use STATUS_COLOR_MAP[task.note] for bar/milestone/flag fill instead of assigned color (true/false)',
  todayLineColor:       'Hex color of the Today vertical line on the Gantt chart (default #ef4444)',
  depsArrowColor:       'Hex color of non-violated dependency arrows, in both the interactive chart and print (default #94a3b8); violated deps always render red',
  kanbanBoards:         'JSON array of Kanban board definitions (name, groupBy status/person, card limit, sort field + direction) — stored per-version like every other setting',
  kanbanTextSize:       'Kanban card/column text size — sm / md / lg (default md)',
  metaDetailsCollapsed: 'Whether the subtitle / date / note block in the project header is collapsed (true/false)',
  statusColors:         'JSON object mapping status names to hex colors — saved from the Status Colors editor popup',
  taskSheetName:           'Name of the Google Sheet tab that contains the task list (default: PROJECT TASK LIST)',
  sortColumn:              'Task-table sort column — field name or empty for unsorted',
  sortDirection:           'Task-table sort direction — 1 = ascending, -1 = descending',
  currentTab:              'Active view tab — gantt | tasks | milestones | flags | kanban',
  tabOrder:                'JSON array: display order of the top app tabs (gantt/tasks/milestones/flags/kanban)',
  tabVisible:              'JSON object: which top app tabs are shown — key -> false when hidden',
  showBaseline:            'Show the frozen baseline overlay under bars/milestones/flags (true/false)',
  baselineCapturedAt:      'ISO timestamp of the last baseline capture for this version (empty = no baseline captured)',
  taskListGroupSortMode:   'Task Properties group ordering mode — date | alpha | custom',
  taskTableColWidths:      'JSON object: Task Properties table column widths in px, keyed by column key',
  taskTableColVisible:     'JSON object: Task Properties table column visibility, keyed by column key (false = hidden)',
  taskListUseUniformRowColor: 'When true, every Task Properties group header row uses taskListUniformRowColor instead of that group\'s own assigned color (true/false)',
  taskListUniformRowColor:    'Hex color used for all Task Properties group header rows when taskListUseUniformRowColor is true',
  taskListRowColorTint:       'Opacity (0-100) of the Task Properties group header row color tint',
  taskListGroupTextColor:     'Hex text color for Task Properties group header row labels (also used in Task List print/export)',
  taskListWrapText:           'Wrap long TASK NAME / NOTES text in the Task Properties table instead of truncating with ellipsis (true/false)',
  taskPrintHeaderColor:       'Hex background color for the Task List print/export table\'s column header row'
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

// Returns the 1-based row number of the last row in the sheet where either
// the DISCIPLINE column (discCol, 1-based) or the TASK column (taskCol, 1-based)
// contains non-empty, non-whitespace content. Falls back to sheet.getLastRow()
// when no such row is found. Used to avoid appending after checkbox-only empty rows.
function findLastContentRow(sheet, discCol, taskCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;
  var discVals = sheet.getRange(1, discCol, lastRow, 1).getValues();
  var taskVals = sheet.getRange(1, taskCol, lastRow, 1).getValues();
  for (var r = lastRow - 1; r >= 0; r--) {
    if (String(discVals[r][0] || '').trim() || String(taskVals[r][0] || '').trim()) {
      return r + 1; // convert 0-based index to 1-based row
    }
  }
  return lastRow;
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

// Verify GANTT BASELINE read/write against a couple of dummy entries (V2.0).
function testWriteBaseline() {
  try {
    var dummyBaseline = [
      { taskId: 1, group: 'ARCHITECTURE', name: 'SCHEMATIC DESIGN', type: 'bar',
        start: '2026-01-01', end: '2026-01-15' },
      { taskId: 2, group: 'STRUCTURAL', name: 'FOUNDATION REVIEW', type: 'milestone',
        start: '2026-02-01', end: '2026-02-01' }
    ];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet: ' + ss.getName());

    writeBaseline(dummyBaseline, false);
    Logger.log('writeBaseline() done');

    var b = readBaseline();
    Logger.log('readBaseline(): ' + JSON.stringify(b));

    writeBaseline([], true);
    Logger.log('writeBaseline([], true) — clear — done');

    var bCleared = readBaseline();
    Logger.log('readBaseline() after clear: ' + JSON.stringify(bCleared));

    SpreadsheetApp.flush();
    Logger.log('DONE — check GANTT BASELINE tab');

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
