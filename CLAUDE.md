# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An interactive HTML Gantt chart for construction/hospitality project schedules (LA COSTA HOTEL and other
projects), backed by a Google Sheet via a Google Apps Script Web App. The chart reads task data from the
sheet on Load and writes changes back on Save. No build system, no package manager, no test runner — pure
vanilla JS/HTML/CSS in a single self-contained HTML file, paired with a single `.gs` Google Apps Script
backend file. There is nothing to `npm install`, build, lint, or run test suites for; verification is done
by opening the HTML directly in a browser.

This repo is a git mirror of what was originally a Google Drive folder (note the stray `desktop.ini` files
under `PROJECTS/` and `ARCHIVED CHATGPT/` — these are Google Drive/Windows artifacts, not project config;
ignore them). There is no README.md and no CI/build config (`.github/workflows` does not exist) — a handful
of empty "trigger: retry GitHub Pages deploy" commits exist only to kick a stuck GitHub Pages build.

## Active files — always edit these

| File | Role |
|------|------|
| `TIMELINE-V1.26.html` | **Active HTML frontend.** All chart rendering, drag-to-reschedule, Load/Save, print mode, and the ⚙ Setup modal live here. Open this file directly in a browser to test. |
| `Code_V1.26.gs` | **Active backend.** Paste into the Google Sheet's Extensions → Apps Script editor and deploy as a Web App. |

Every other `TIMELINE-V*.html` / `Code_V*.gs` file in the repo root is a **superseded version kept for
reference only** — do not edit them, and do not assume the highest-numbered file you see is necessarily the
one in active use; check `## Version History` at the bottom of this doc (and the file's own in-file
`VERSION HISTORY` header comment) to confirm which pair is current before starting work. New feature/fix
work happens by editing `TIMELINE-V1.26.html` and `Code_V1.26.gs` in place — bump to a new `V1.XX` pair only
when explicitly asked to cut a new version; mid-flight branch versions (e.g. a stray `TIMELINE-V1.27.html`)
have historically been merged back into V1.26 rather than kept, so don't assume a new numbered file should
persist unless told to keep it.

### Deployed/mirrored snapshots — do not treat as sources of truth

- `index.html` (repo root) and `LA_COSTA/index.html`, `SANFORD_145/index.html` are **static GitHub Pages
  snapshots**, each frozen at whatever `TIMELINE-V*.html` state existed when it was copied in (all three
  were last touched in commit `271ff6e`, an old V1.15/V1.16-era snapshot — they are **not** kept in sync
  with `TIMELINE-V1.26.html` automatically). If asked to update a deployed/hosted timeline, that means
  manually copying the current active HTML over the relevant `index.html` and committing — this repo has no
  automation that does it for you.
- `PROJECTS/*.html` and the other files inside `LA_COSTA/` / `SANFORD_145/` are dated, per-client export
  snapshots (filenames encode project + version + date, e.g. `LA COSTA-TIMELINE-V1.21-20260625.html`).
  Treat them as historical exports, not code to modify.
- `ARCHIVED/` and `ARCHIVED CHATGPT/` hold genuinely obsolete early iterations (original `Code.gs`,
  `Google_Apps_Script_Code.js`, pre-V1.0 HTML). Reference only if tracing very old behavior; never edit.

### Setup Tool

`SETUP-TOOL-V0.03.html` (previous: `V0.0`–`V0.02`) is a standalone utility — unrelated to the Gantt chart
itself — that uses the browser File System Access API to bulk-patch a Web App URL and/or an embedded logo
(as base64) directly into one or more local `TIMELINE-V*.html` files on disk, so a deployed copy doesn't
need the ⚙ Setup panel filled in by hand. It has no server component and isn't wired into the Gantt HTML at
runtime.

`SHEETS_SYNC_SETUP.md` is an old two-way-sync setup guide written against a much earlier file layout
(`LA_COSTA_HOTEL_TIMELINE_INTERACTIVE.html` / `Google_Apps_Script_Code.js`, now in `ARCHIVED/`). The
deployment steps below (under "Deployment") are the current procedure; treat that file as historical.

## Google Sheet

**URL:** https://docs.google.com/spreadsheets/d/1HShZAkZ7oV4_yDdRRbAeBG2p5Uh5EvH4Chs_64ZZ5WU

**Tabs:**
- `PROJECT TASK LIST` — source of truth for all tasks (default name; user-overridable per-project via the
  `taskSheetName` setting — see below). Never reformatted by the script.
- `GANTT SETTINGS-DO NOT EDIT` — auto-created by the script on first Save; stores chart UI settings.
- `GANTT TASK PARAMS` — auto-created; stores per-task Gantt-only display overrides (color, bar type/style,
  symbol) keyed by task, independent of the main task list.

## Apps Script architecture (`Code_V1.26.gs`)

### Entry points
- `doGet(e)` — called by the HTML on Load. Reads chart settings via `readSettings()` **first** (so a
  custom `taskSheetName` is honored), then imports tasks via `importFromTaskList()` and merges in
  `readTaskParams()`. Returns tasks + task params + settings as one JSON response.
- `doPost(e)` — called by the HTML on Save. Runs `saveBackToTaskList(payload)`, `writeTaskParams(payload)`,
  and `writeSettings(payload.settings)` independently, each in its own try/catch, so a failure in one never
  blocks the others (e.g. a bad task write still lets settings persist).
- `onOpen()` — injects a **📊 Gantt Timeline** custom menu into the Sheets UI (Get Web App URL, About/Setup
  Help). Only takes effect after redeploying — menu registration doesn't apply retroactively.

### Key functions
- `importFromTaskList()` — scans the task sheet for rows where `SCHEDULE=TRUE` or `MILESTONE=TRUE`, with
  valid START DATE + END DATE. Returns `{ tasks: [...], meta: {...} }`.
- `saveBackToTaskList(payload)` — matches tasks by a `DISCIPLINE|TASKNAME` composite key; updates START
  DATE, END DATE, STATUS, MILESTONE (always written explicitly, TRUE or FALSE, so unchecking it in the
  sheet is respected next Load), and NOTES. Appends brand-new tasks after their discipline group (unknown
  disciplines go to the bottom). Deletes rows for tasks removed in the chart. Uses `LockService` to guard
  against concurrent Save race conditions.
- `readSettings()` / `writeSettings(settings)` — read/write the `GANTT SETTINGS` tab. The deployed Web App
  URL is itself stored as the first settings row on every Save; `applySettings()` in the HTML adopts it as
  `sheetsURL` if it differs from what's stored, since the sheet is treated as the source of truth for the
  URL, not the browser's local copy.
- `readTaskParams()` / `writeTaskParams(tasks)` — read/write the `GANTT TASK PARAMS` tab.

### Column mapping (task sheet)
Columns are auto-detected by header name, falling back to a fixed index only if the header isn't found:

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

### Status → bar color map (`STATUS_COLORS`, top of `Code_V1.26.gs`)
`IN PROGRESS` green `#16a34a` · `UPCOMING` amber `#f59e0b` · `DOWNSTREAM` purple `#8b5cf6` ·
`PENDING` yellow `#d9c34a` · `COMPLETED` / `ON HOLD` grey `#94a3b8` · `CANCELLED` / `URGENT` red `#dc2626` ·
`75% COMPLETE` light green `#22c55e` · `WAITING ON OTHERS` / `WITING ON THE CITY` [sic, do not "fix" the
typo without checking live sheet data for that exact string] orange `#f97316` · `UNDER REVIEW` cyan
`#06b6d4`. Unknown status falls back to slate `#64748b`. This map is duplicated conceptually on the HTML
side as `STATUS_COLOR_MAP` for the "Status Colors" display toggle — keep both in sync when editing.

### Chart settings (`SETTINGS_KEYS` array + `SETTINGS_DESCRIPTIONS` object, `Code_V1.26.gs`)
Persisted to the `GANTT SETTINGS` tab as flat key/value rows — covers project meta (name/subtitle/date/note),
layout (label width, font size, dark/flat mode, bar text color), independent per-tab collapse state
(`collapsedGroups`, `ganttCollapsedGroups`, `milestonesCollapsedGroups`, `flagsCollapsedGroups`), group
ordering, column visibility toggles, print/zoom state, dependencies, today-line color, status-color toggle,
the configurable `taskSheetName`, and persisted sort/tab state (`sortColumn`, `sortDirection`, `currentTab`).
**Section colors** are the one exception: stored as individual `groupColor.DISCIPLINE_NAME: #hexcolor` rows
(one per discipline, sorted A–Z) rather than as a key in `SETTINGS_KEYS`, so they're directly readable/
editable in the sheet; reconstructed into a `groupColors` JSON object on read.

## Deployment (Apps Script)

1. Open the Google Sheet → Extensions → Apps Script.
2. Paste the contents of `Code_V1.26.gs` → Save.
3. Deploy → New deployment → Web app:
   - Execute as: **Me**
   - Who has access: **Anyone** ← must be exactly this, not "Anyone within [domain]"
4. Copy the Web App URL → paste into the HTML's ⚙ Setup panel.
5. **To update after code changes:** Deploy → Manage deployments → pencil → New version → Deploy (URL stays
   the same — no need to re-paste into the HTML).

> ⚠️ **Google Workspace gotcha:** on Workspace accounts the default access is often
> `Anyone within [domain].com`, which generates a domain-locked URL
> (`https://script.google.com/a/macros/yourdomain.com/s/.../exec`) that requires a signed-in Workspace
> account. The HTML calls it anonymously, so it will always fail with a connection error. Fix: set access to
> **Anyone** (no domain), which generates `https://script.google.com/macros/s/.../exec`.

### Test functions (run manually in the Apps Script editor, not an automated suite)
- `testImport()` — verifies the sheet is readable; logs task count to the Execution Log.
- `testWriteSettings()` — writes dummy settings to the `GANTT SETTINGS` tab; verifies write works.
- `testCreateSettingsTab()` — creates/recreates the `GANTT SETTINGS` tab from scratch.

## Common tasks

- **Edit the status color map:** `STATUS_COLORS` in `Code_V1.26.gs`, and mirror in `STATUS_COLOR_MAP` inside
  `TIMELINE-V1.26.html` if the change should also affect the "Status Colors" display toggle.
- **Add a new chart setting:** add the key to `SETTINGS_KEYS` and describe it in `SETTINGS_DESCRIPTIONS` in
  `Code_V1.26.gs`, then read/write it from `collectSettings()` / `applySettings()` in the HTML frontend.
  Redeploy the Apps Script for the new key to actually persist.
- **Change which rows appear on the chart:** filter logic in `importFromTaskList()` in `Code_V1.26.gs`
  (currently: `SCHEDULE=TRUE` OR `MILESTONE=TRUE`, plus valid dates).
- **Change how new tasks are appended on Save:** `saveBackToTaskList()` in `Code_V1.26.gs` — new rows go
  after the last row of their discipline group; unknown disciplines go to the bottom.
- **Frontend changes (rendering, drag-to-reschedule, print mode, Load/Save, ⚙ Setup modal):** all in
  `TIMELINE-V1.26.html`, a single self-contained file with no external dependencies to install. Open it
  directly in a browser to iterate — there's no dev server or hot reload.

## Version history

See the `VERSION HISTORY` comment block at the top of `TIMELINE-V1.26.html` and `Code_V1.26.gs` for the
authoritative, detailed per-release changelog (both files carry their own). Highlights of the current
(V1.26) state relative to earlier majors documented in prior revisions of this file:

- **V1.26** — User-configurable task sheet name (`taskSheetName` setting; Setup panel field), applied to
  every Load/Save. Project title word-wrap fixed against the label column width so it no longer gets
  painted over when Date/Duration columns are shown. Active tab and task-table sort column/direction persist
  across loads. Group sort controls added to the Task Properties section (`taskListGroupSortMode`).
- **V1.25** — Print mode restyled to match the main settings-card UI and made collapsible; editable zoom %
  input with working per-page CSS zoom; Export PDF (lazy-loaded html2canvas + jsPDF from CDN) across all
  print modes; logo upload + size controls (topbar and print); today-line color picker; dark-mode axis-label
  fix; several settings that were collected but never persisted (`todayLineColor`, `metaDetailsCollapsed`,
  `statusColors`) wired into `SETTINGS_KEYS`.
- **V1.23** — "Status Colors" toolbar toggle switches all bar/milestone/flag/table fills between assigned
  colors and status-derived colors via a new `_effectiveBarColor(task)` helper used consistently across
  every render path. Per-group variable flat-mode row heights.
- **V1.15** — Shift+click multi-select and multi-drag across bars/milestones/flags; Milestones and Flags
  tabs with independent per-tab collapse state; global ◉ Markers toggle for collapsed rollup rows.
- **V1.10** — Print mode two-row month ruler; task dependency arrows (finish-to-start) persisted to
  `GANTT TASK PARAMS` by stable `GROUP|NAME` keys; auto-load on open using `localStorage`-cached URL.
- **V1.0–V1.08** — Baseline sheet import/export, `GANTT SETTINGS` tab, per-task color/type/style overrides
  in `GANTT TASK PARAMS`, section colors, NOTES column round-trip.

When making a change, prefer updating this summary only for genuinely repo-wide/structural shifts (new
active file pair, new folder convention, new deployment step) — routine feature work should just extend the
in-file `VERSION HISTORY` comments in the active `.html`/`.gs` pair, which are the detailed record.
