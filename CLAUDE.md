# LA COSTA HOTEL — Gantt Timeline Project

## What this is
An interactive HTML Gantt chart for the LA COSTA HOTEL project schedule, backed by a Google Sheet via a Google Apps Script Web App. The chart reads task data from the sheet on Load and writes changes back on Save. No build system — pure vanilla JS in a single HTML file + a single `.gs` script.

---

## File Overview

| File | Purpose |
|------|---------|
| `Code_V1.22.gs` | **Active backend** — deploy this. No logic changes from V1.21; version bump only. |
| `Code_V1.21.gs` | Previous backend. Keep for reference. |
| `Code_V1.20.gs` | Previous backend. Keep for reference. |
| `Code_V1.19.gs` | Previous backend. Keep for reference. |
| `Code_V1.18.gs` | Previous backend. Keep for reference. |
| `Code_V1.15.gs` | Previous backend. Keep for reference. |
| `Code_V1.14.gs` | Previous backend. Keep for reference. |
| `Code_V1.13.gs` | Previous backend. Keep for reference. |
| `Code_V1.12.gs` | Previous backend. Keep for reference. |
| `Code_V1.11.gs` | Previous backend. Keep for reference. |
| `Code_V1.10.gs` | Previous backend. Keep for reference. |
| `Code_V1.09.gs` | Previous backend. Keep for reference. |
| `Code_V1.08.gs` | Previous backend. Keep for reference. |
| `Code_V1.07.gs` | Previous backend. Keep for reference. |
| `Code_V1.06.gs` | Previous backend. Keep for reference. |
| `Code_V1.05.gs` | Previous backend. Keep for reference. |
| `Code_V1.04.gs` | Previous backend. Keep for reference. |
| `Code_V1.03.gs` | Previous backend. Keep for reference. |
| `Code_V1.02.gs` | Previous backend. Keep for reference. |
| `Code_V1.01.gs` | Previous backend. Keep for reference. |
| `Code_V1.0.gs` | Original baseline. Keep for reference. |
| `Code.gs` | Original base script. Keep for reference. |
| `TIMELINE-V1.23.html` | **Active HTML frontend** — open this in browser. Status Colors toggle + per-group flat row heights. |
| `TIMELINE-V1.22.html` | Previous HTML version. Keep for reference. |
| `TIMELINE-V1.21.html` | Previous HTML version. Keep for reference. |
| `TIMELINE-V1.20.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.15.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.14.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.13.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.12.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.11.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.10.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.09.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.08.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.07.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.06.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.05.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.04.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.03.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.02.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.01.html` | Previous HTML version. Keep for reference. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE.html` | Older HTML version. |
| `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE_V1.0.html` | Older HTML version. |
| `LA_COSTA_HOTEL_TIMELINE.html` | Static (non-interactive) version. |
| `Google_Apps_Script_Code.js` | Older standalone JS version — superseded by `.gs` files. |
| `SHEETS_SYNC_SETUP.md` | Full deployment guide for the Apps Script Web App. |
| `_copy_v101.bat` | Windows batch: copies V1.01 files to working locations. |

---

## Google Sheet

**URL:** https://docs.google.com/spreadsheets/d/1HShZAkZ7oV4_yDdRRbAeBG2p5Uh5EvH4Chs_64ZZ5WU

**Tabs:**
- `PROJECT TASK LIST` — source of truth for all tasks (never reformatted by script)
- `GANTT SETTINGS-DO NOT EDIT` — auto-created by script on first Save; stores chart UI settings

---

## Apps Script Architecture (`Code_V1.03.gs`)

### Entry points
- `doGet(e)` — called by HTML on Load. Reads tasks via `importFromTaskList()` and chart settings via `readSettings()`. Returns JSON.
- `doPost(e)` — called by HTML on Save. Runs `saveBackToTaskList(payload)` AND `writeSettings(payload.settings)` independently (settings save always runs even if task save fails).

### Key functions
- `importFromTaskList()` — scans PROJECT TASK LIST for rows where SCHEDULE=TRUE or MILESTONE=TRUE, with valid START DATE + END DATE. Returns tasks array + meta.
- `saveBackToTaskList(payload)` — matches tasks by `DISCIPLINE|TASKNAME` key and updates START DATE, END DATE, STATUS. Appends brand-new tasks after their discipline group. Uses `LockService` to prevent race conditions.
- `readSettings()` / `writeSettings(settings)` — read/write the GANTT SETTINGS tab.
- `readTaskParams()` / `writeTaskParams(tasks)` — read/write the GANTT TASK PARAMS tab (per-task colour override, type, style, symbol).

### Column mapping (PROJECT TASK LIST)
| Header | Default col index |
|--------|------------------|
| DISCIPLINE | 1 |
| ID | 2 |
| TASK | 3 |
| CONSULTANT | 4 |
| PERSON | 5 |
| START DATE | 7 |
| END DATE | 8 |
| SCHEDULE | 10 |
| MILESTONE | 11 |
| STATUS | 12 |
| PRIORITY | 13 |
| NOTES | 20 |

Script auto-detects columns by header name (falls back to index if header not found).

### Status → bar colour map
Defined in `STATUS_COLORS` at top of script. Key statuses:
- IN PROGRESS → green `#16a34a`
- UPCOMING → amber `#f59e0b`
- COMPLETED → grey `#94a3b8`
- CANCELLED / URGENT → red `#dc2626`
- NOTE → light blue `#60a5fa`

Default (unknown status) → `#64748b` (slate).

### Chart settings keys (saved to GANTT SETTINGS tab)
Fixed keys: `userLabelWidth`, `ganttBarFontSize`, `showTodayLine`, `darkMode`, `flatMode`, `barTextColor`, `collapsedGroups`, `disciplineOrder`, `showName`, `showPhase`, `showNote`

Section colors: one row per discipline → `groupColor.DISCIPLINE_NAME : #hexcolor` (sorted A–Z). Reconstructed into `groupColors` JSON on read.

---

## Deployment (Apps Script)

1. Open the Google Sheet → Extensions → Apps Script
2. Paste `Code_V1.01.gs` contents → Save
3. Deploy → New deployment → Web app
   - Execute as: **Me**
   - Who has access: **Anyone** ← must be exactly this, not "Anyone within [domain]"
4. Copy the Web App URL → paste into the HTML's ⚙ Setup panel
5. **To update after code changes:** Deploy → Manage deployments → pencil → New version → Deploy (URL stays the same)

> ⚠️ **Google Workspace gotcha:** On Workspace accounts the default access is often set to
> `Anyone within [domain].com` which generates a domain-locked URL:
> `https://script.google.com/a/macros/yourdomain.com/s/.../exec`
> This URL requires a signed-in Workspace account — the HTML tool calls it anonymously so
> it will always fail with a connection error. Fix: change access to **Anyone** (no domain),
> which generates `https://script.google.com/macros/s/.../exec` (no domain in path).

---

## Test Functions (run manually in Apps Script editor)
- `testImport()` — verifies sheet is readable; logs task count to Execution Log
- `testWriteSettings()` — writes dummy settings to GANTT SETTINGS tab; verifies write works
- `testCreateSettingsTab()` — creates/recreates the GANTT SETTINGS tab from scratch

---

## Version History
- **V1.23** (2026-07-02) — HTML only. **Status Colors toggle**: "Status Colors" checkbox in the toolbar switches all bar / milestone / flag fills between assigned colors (`colorOverride → task.color`) and status-derived colors (`STATUS_COLOR_MAP[task.status]`). `_effectiveBarColor(task)` helper centralises the colour logic; all rendering paths (interactive bars, milestones, flags, collapsed markers, flat disc rows, print individual task rows, print collapsed markers) and the task-table colour dot are wired to it. Toggle persisted to settings as `useStatusColors`. Backend unchanged. Also includes per-group variable flat-mode row heights: each discipline row height is computed from its own max outside-label line count (`_flatMaxLabelLines` changed from global scalar to per-group dict); `geom.rowYOffsets[]` replaces the old uniform stride; all rendering updated to use per-row heights.
- **V1.22** (2026-07-01) — HTML only. Flat mode bar labels repositioned **above** each bar instead of inside it. `getEffectiveRowHeight()` now adds a label area (`Math.round(ganttBarFontSize * 1.4) + 4` px) above the bar so rows are tall enough to accommodate the label without overlapping the bar row above. `renderTaskBar` flat-mode branch: `barY` shifted down by `_flatLabelH`; labels rendered at `yTop + ganttBarFontSize + 2` (above bar baseline); `lineGap` set to `ganttBarFontSize + 2`; clipPath covers only the label area (horizontally constrained to bar x-span). Label fill uses chart text colour (not bar-contrast) since labels sit on the row background. Same layout mirrored in `_buildPrintPageEl` disc rows. Backend unchanged (no redeploy needed).
- **V1.15** (2026-05-22) — HTML only. Shift+click multi-select on Gantt bars, milestones, and flags. `ganttMultiSelected` object tracks selected IDs. `onBarMouseDown` handles three cases: shift+click (toggle in multi-set, rebuild `drag.multiOriginals`), plain click on multi-selected item (start multi-drag), plain click on unselected item (clear selection, single drag). `onMouseMove` checks `drag.isMulti` and moves all tasks in `multiOriginals` by the same day delta, preserving bar durations. Multi-selected items render with amber stroke/glow (`.task-bar.multi-selected` CSS class). Status bar shows count + ESC hint when 2+ items selected. ESC key and background click clear `ganttMultiSelected`. Also includes all V1.14 features: Milestones/Flags tabs, independent collapse states per tab, global ◉ Markers toggle, print mode parity, 4 new settings keys.
- **V1.14** (2026-05-22) — HTML + backend. Milestones tab and Flags tab: filtered task-list views alongside the existing Tasks tab. Four independent collapse-state variables (`collapsedGroups`, `milestonesCollapsedGroups`, `flagsCollapsedGroups`, `ganttCollapsedGroups`). Global ◉ Markers toggle button shows/hides milestone & flag markers on collapsed rollup bars; button only visible when any group is collapsed. Print mode updated to respect Gantt collapse state, flag shapes, and markers toggle. Four new settings keys in `SETTINGS_KEYS`: `ganttCollapsedGroups`, `milestonesCollapsedGroups`, `flagsCollapsedGroups`, `ganttRollupMarkersVisible`. Fixed `applySettings()` scope bug (inline `getElementById` replaces closure call).
- **V1.13** (2026-05-21) — Backend only. Adds `onOpen()` which injects a permanent **📊 Gantt Timeline** custom menu into the Google Sheets UI. Menu items: **Get Web App URL** (shows deployed URL via `ScriptApp.getService().getUrl()`, falls back to `appScriptURL` row in GANTT SETTINGS tab, shows HTML dialog with Copy button; shows deployment instructions if not yet deployed) and **About / Setup Help** (5-step setup guide dialog). Requires one redeployment to activate; menu persists on every sheet open thereafter.
- **V1.12** (2026-05-21) — HTML only. Hover tooltip on Gantt bars: hovering any bar, milestone, or flag shows a floating white card with Group, Task Name, Start/End dates, and Note (if present). Tooltip follows cursor and stays within viewport. Dismissed on mouse-leave or when drag starts.
- **V1.11** (2026-05-20) — HTML only. Interactive drag-to-connect dependency grips: hover any bar to reveal small grey circles at its Start and End; click-drag from a circle to another bar's circle to create a link; rubber-band dashed blue line follows cursor during drag; target circles highlight blue when in range; drop creates `{id, to}` dep entry and re-renders arrows + task table. Works for bars, milestones, and flags.
- **V1.10** (2026-05-20) — HTML + backend. Print mode: two-row month ruler (years top, months bottom — no overlap); vertical grid lines rendered after row fills so they're always visible; grid line colors lightened; alternating row stripe vs grid line contrast fixed. Auto-load on every open/reload — fires from `initSheetsSync` before any other code can block it; `localStorage` URL takes priority over embedded URL so the last confirmed-working URL is reused on next open. **Task dependencies**: Finish-to-Start link arrows drawn on Gantt between predecessor end → successor start; elbow-routed SVG paths; arrowhead at successor; red dashed when constraint violated (successor starts before predecessor ends). DEPS column added to task table (🔗 button opens picker modal). Dependencies persisted to GANTT TASK PARAMS tab as tilde-separated `GROUP|NAME` keys (stable across reloads); resolved back to runtime IDs on Load.
- **V1.09** (2026-05-19) — HTML only. Removed outer border rect from Gantt print mode pages.
- **V1.08** (2026-05-19) — HTML + backend. App Script URL saved as first row of GANTT SETTINGS tab on every Save; `readSettings()` returns it; `applySettings()` adopts it as `sheetsURL` if it differs from the stored value (sheet is source of truth for URL). Title and subtitle word-wrap in the corner SVG. Date display strips time/timezone. Saved project meta (name, subtitle, date, note) restored from settings on Load. Group header backgrounds use SVG-compliant `fill` + `fill-opacity` attributes (fixes black export in SVG). V1.07 files kept for reference.
- **V1.06** (2026-05-19) — HTML only. Separate Task List print mode with its own overlay, controls (page size, orientation, font, row height, column toggles), auto vertical pagination, discipline color as left-border accent. Removed color background tint from discipline group rows in the task table.
- **V1.05** (2026-05-18) — HTML only. Print mode: full-screen preview overlay, page size/orientation controls, label/row/font sliders, date range, optional horizontal timeline split (N months/page), auto vertical pagination. Each page is SVG-rendered with group headers, bars, milestones, today line. `window.print()` + dynamic `@page` CSS.
- **V1.04** (2026-05-18) — HTML + backend. Bar text renders LEFT of bar (no overlap) when estimated label width exceeds bar width. NOTES column added to task table — editable inline, saved back to PROJECT TASK LIST column 20 on Save.
- **V1.03** (2026-05-18) — HTML + backend. New `GANTT TASK PARAMS` tab stores per-task colour override, bar type, bar style, and symbol. `doGet` merges params into tasks on Load; `doPost` writes them independently. Colour picker restored in task table; bar colour priority: override → section → status.
- **V1.02** (2026-05-18) — HTML + backend. Section colors saved as individual `groupColor.DISCIPLINE` rows in the settings tab (human-readable, directly editable in sheet). Section color tints group header background; task rows get matching left border. Bar color inherits section color. Removed per-task color picker. Fixed bar text color override.
- **V1.01** (2026-05-18) — Added GANTT SETTINGS tab. `doGet` now returns settings. `doPost` saves settings independently. `readSettings()` / `writeSettings()` added.
- **V1.0** (2026-05-18) — Baseline. Reads PROJECT TASK LIST, filters by SCHEDULE/MILESTONE checkbox. Status → bar colour. Two-way save.

---

## Common Tasks for Claude Code

**Edit the status colour map:** `STATUS_COLORS` object at top of `Code_V1.02.gs`

**Add a new chart setting:** Add the key to `SETTINGS_KEYS` array and `SETTINGS_DESCRIPTIONS` object in `Code_V1.02.gs`, then handle it in the HTML frontend.

**Change which rows appear on the chart:** Edit the filter logic in `importFromTaskList()` around lines 183–185 (currently: SCHEDULE=TRUE OR MILESTONE=TRUE, plus valid dates).

**Change how new tasks are appended:** See `saveBackToTaskList()` — new rows go after the last row of their discipline group; unknown disciplines go to the bottom.

**HTML frontend:** All chart rendering, drag-to-reschedule, Load/Save buttons, and ⚙ Setup modal are in `LA_COSTA_HOTEL_TIMELINE_INTERACTIVE-V1.02.html`. It's a single self-contained file — no external dependencies to install.
