/**
 * PROJECT DASHBOARD — WEEKLY REPORTING VIEW
 *
 * HOW TO RUN:
 *   1. Open your Google Spreadsheet
 *   2. Extensions → Apps Script
 *   3. Create a new project, paste this entire file
 *   4. Click Run → createProjectDashboard
 *   5. Grant permissions when prompted
 *   6. A "DASHBOARD" tab is created (or replaced if it already exists)
 *
 * The dashboard uses live COUNTIF / QUERY formulas — it auto-updates
 * whenever data in "PROJECT TASK LIST" changes.
 */

// ── CONFIG ─────────────────────────────────────────────────────────────────
const DB_DB_SOURCE_SHEET  = 'PROJECT TASK LIST';
const DB_DB_DASH_NAME     = 'DASHBOARD';

// Column letters in DB_DB_SOURCE_SHEET (do not change unless headers move)
const DB_COL = {
  DISCIPLINE : 'A',
  ID         : 'B',
  TASK       : 'C',
  CONSULTANT : 'D',
  PERSON     : 'E',
  DATE_ADDED : 'F',
  START      : 'G',
  END        : 'H',
  DURATION   : 'I',
  SCHEDULE   : 'J',
  MILESTONE  : 'K',
  STATUS     : 'L',
  PRIORITY   : 'M',
  WEEKLY     : 'N',
  NOTES      : 'T',
};

// ── COLORS ─────────────────────────────────────────────────────────────────
const DB_C = {
  navy        : '#162032',
  navyLight   : '#1E3050',
  navyMid     : '#2A4A6E',
  blue        : '#5AAFF0',
  separator   : '#DCE5F0',

  // KPI cards
  orange      : '#E86A00',
  orangeDark  : '#BF4F00',
  cobalt      : '#1A72B8',
  cobaltDark  : '#145A8C',
  red         : '#C62828',
  redDark     : '#9B1C1C',
  purple      : '#7C3AED',
  purpleDark  : '#5B21B6',
  green       : '#059669',
  greenDark   : '#065F46',

  // Status badges
  statusIP    : { bg: '#FFF0D6', font: '#7A3800' },   // IN PROGRESS
  statusUP    : { bg: '#DBEAFE', font: '#1E3A5F' },   // UPCOMING
  statusDS    : { bg: '#E5E7EB', font: '#374151' },   // DOWNSTREAM
  statusPN    : { bg: '#FEF9C3', font: '#713F12' },   // PENDING
  status75    : { bg: '#D1FAE5', font: '#064E3B' },   // 75% COMPLETE

  // Priority badges
  high        : { bg: '#FEE2E2', font: '#991B1B' },
  medium      : { bg: '#FEF3C7', font: '#92400E' },
  low         : { bg: '#DBEAFE', font: '#1E3A5F' },
};

// ── MAIN FUNCTION ──────────────────────────────────────────────────────────
function createProjectDashboard() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(DB_SOURCE_SHEET);

  if (!src) {
    SpreadsheetApp.getUi().alert(`Sheet "${DB_SOURCE_SHEET}" not found.\nCheck the tab name and try again.`);
    return;
  }

  // Create or wipe dashboard tab
  let dash = ss.getSheetByName(DB_DASH_NAME);
  if (dash) {
    dash.clear();
    dash.clearFormats();
    dash.clearConditionalFormatRules();
  } else {
    dash = ss.insertSheet(DB_DASH_NAME, 0);
  }

  // ── COLUMN WIDTHS ──────────────────────────────
  const widths = [175, 320, 115, 135, 92, 92, 135, 105, 270];
  widths.forEach((w, i) => dash.setColumnWidth(i + 1, w));
  if (dash.getMaxColumns() > 9) {
    dash.hideColumns(10, dash.getMaxColumns() - 9);
  }

  // ── ROW 1 — TITLE ──────────────────────────────
  dash.setRowHeight(1, 52);
  dbCell(dash, 'A1:I1', {
    value     : 'PROJECT DASHBOARD  ·  WEEKLY REPORTING VIEW',
    bg        : DB_C.navy,
    fontColor : '#FFFFFF',
    fontSize  : 15,
    bold      : true,
    hAlign    : 'center',
    vAlign    : 'middle',
  });

  // ── ROW 2 — TIMESTAMP ──────────────────────────
  dash.setRowHeight(2, 20);
  dbCell(dash, 'A2:I2', {
    formula   : '="Last viewed: " & TEXT(NOW(),"MMM D, YYYY  ·  h:MM AM/PM")',
    bg        : DB_DB_C.navyLight,
    fontColor : '#7A9BC0',
    fontSize  : 8,
    hAlign    : 'center',
    vAlign    : 'middle',
  });

  // ── ROW 3 — SPACER ─────────────────────────────
  dash.setRowHeight(3, 8);
  dash.getRange('A3:I3').setBackground(DB_C.separator);

  // ── ROWS 4-5 — KPI CARDS ───────────────────────
  dash.setRowHeight(4, 18);
  dash.setRowHeight(5, 46);

  const kpis = [
    {
      col: 'A', span: 2,
      label   : 'IN PROGRESS',
      formula : `=COUNTIF('${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"IN PROGRESS")`,
      labelBg : DB_DB_C.orangeDark,
      numBg   : DB_C.orange,
    },
    {
      col: 'C', span: 2,
      label   : 'UPCOMING',
      formula : `=COUNTIF('${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"UPCOMING")`,
      labelBg : DB_DB_C.cobaltDark,
      numBg   : DB_C.cobalt,
    },
    {
      col: 'E', span: 2,
      label   : '1-HIGH ACTIVE',
      formula : `=COUNTIFS('${DB_SOURCE_SHEET}'!${DB_COL.PRIORITY}:${DB_COL.PRIORITY},"1-HIGH",`
              + `'${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"<>COMPLETED",`
              + `'${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"<>CANCELLED")`,
      labelBg : DB_DB_C.redDark,
      numBg   : DB_C.red,
    },
    {
      col: 'G', span: 2,
      label   : 'WEEKLY TRACKED',
      formula : `=COUNTIFS('${DB_SOURCE_SHEET}'!${DB_COL.WEEKLY}:${DB_COL.WEEKLY},TRUE,`
              + `'${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"<>COMPLETED",`
              + `'${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"<>CANCELLED")`,
      labelBg : DB_DB_C.purpleDark,
      numBg   : DB_C.purple,
    },
    {
      col: 'I', span: 1,
      label   : 'COMPLETED',
      formula : `=COUNTIF('${DB_SOURCE_SHEET}'!${DB_COL.STATUS}:${DB_COL.STATUS},"COMPLETED")`,
      labelBg : DB_DB_C.greenDark,
      numBg   : DB_C.green,
    },
  ];

  kpis.forEach(k => {
    const endCol   = dbColLetter(dbColIndex(k.col) + k.span - 1);
    const labelRng = `${k.col}4:${endCol}4`;
    const numRng   = `${k.col}5:${endCol}5`;

    dbCell(dash, labelRng, {
      value: k.label, bg: k.labelBg, fontColor: '#FFFFFF',
      fontSize: 8, bold: true, hAlign: 'center', vAlign: 'middle',
    });
    dbCell(dash, numRng, {
      formula: k.formula, bg: k.numBg, fontColor: '#FFFFFF',
      fontSize: 30, bold: true, hAlign: 'center', vAlign: 'middle',
    });
  });

  // ── ROW 6 — SPACER ─────────────────────────────
  dash.setRowHeight(6, 12);
  dash.getRange('A6:I6').setBackground(DB_C.separator);

  // ── ROW 7 — SECTION LABEL ──────────────────────
  dash.setRowHeight(7, 30);
  dbCell(dash, 'A7:I7', {
    value     : '   WEEKLY TASKS  —  ACTIVE PROJECT ITEMS',
    bg        : DB_C.navy,
    fontColor : DB_C.blue,
    fontSize  : 11,
    bold      : true,
    hAlign    : 'left',
    vAlign    : 'middle',
  });

  // ── ROW 8 — COLUMN HEADERS ─────────────────────
  dash.setRowHeight(8, 22);
  ['DISCIPLINE','TASK','CONSULTANT','PERSON','START','END DATE','STATUS','PRIORITY','NOTES']
    .forEach((h, i) => {
      dash.getRange(8, i + 1)
        .setValue(h)
        .setBackground(DB_DB_C.navyMid)
        .setFontColor('#FFFFFF')
        .setFontSize(8)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    });

  // ── ROW 9+ — WEEKLY TASKS (live QUERY) ─────────
  //   Source columns: A=DISCIPLINE C=TASK D=CONSULTANT E=PERSON
  //                   G=START H=END L=STATUS M=PRIORITY T=NOTES
  //   Filter:  N (WEEKLY) = true, L not COMPLETED/CANCELLED
  //   Sort:    M (PRIORITY) asc, H (END DATE) asc
  const q = [
    `=IFERROR(QUERY('${DB_SOURCE_SHEET}'!A:T,`,
    `"SELECT A, C, D, E, G, H, L, M, T`,
    ` WHERE N = true`,
    ` AND L <> 'COMPLETED'`,
    ` AND L <> 'CANCELLED'`,
    ` ORDER BY M ASC, H ASC`,
    ` LABEL A '',C '',D '',E '',G '',H '',L '',M '',T ''",0),`,
    `"— No active weekly tasks —")`,
  ].join('');

  dash.getRange('A9').setFormula(q);

  // Default styling for data area
  dash.getRange('A9:I400')
    .setFontSize(9)
    .setFontFamily('Arial')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  dash.getRange('A9:A400').setFontWeight('bold').setFontColor(DB_DB_C.navyMid);

  for (let r = 9; r <= 400; r++) dash.setRowHeight(r, 22);

  // ── FREEZE THROUGH COLUMN HEADERS ──────────────
  dash.setFrozenRows(8);

  // ── CONDITIONAL FORMATTING ─────────────────────
  const rules = [];
  const statusRng   = dash.getRange('G9:G400');
  const priorityRng = dash.getRange('H9:H400');

  [
    ['IN PROGRESS', DB_C.statusIP],
    ['UPCOMING',    DB_C.statusUP],
    ['DOWNSTREAM',  DB_C.statusDS],
    ['PENDING',     DB_C.statusPN],
    ['75% COMPLETE',DB_C.status75],
  ].forEach(([text, colors]) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(colors.bg)
        .setFontColor(colors.font)
        .setRanges([statusRng])
        .build()
    );
  });

  [
    ['1-HIGH',   DB_C.high],
    ['2-MEDIUM', DB_C.medium],
    ['3-LOW',    DB_C.low],
  ].forEach(([text, colors]) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(colors.bg)
        .setFontColor(colors.font)
        .setRanges([priorityRng])
        .build()
    );
  });

  dash.setConditionalFormatRules(rules);

  // ── DONE ───────────────────────────────────────
  ss.setActiveSheet(dash);
  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '✅  Dashboard created!\n\n' +
    'The WEEKLY TASKS table updates automatically whenever\n' +
    '"PROJECT TASK LIST" data changes — no refresh needed.'
  );
}

// ── HELPERS ────────────────────────────────────────────────────────────────

function dbCell(sheet, a1, opts) {
  const rng = sheet.getRange(a1);
  const [start, end] = a1.split(':');
  if (end && start !== end) rng.merge();
  if (opts.formula)              rng.setFormula(opts.formula);
  else if (opts.value !== undefined) rng.setValue(opts.value);
  if (opts.bg)        rng.setBackground(opts.bg);
  if (opts.fontColor) rng.setFontColor(opts.fontColor);
  if (opts.fontSize)  rng.setFontSize(opts.fontSize);
  if (opts.bold)      rng.setFontWeight('bold');
  if (opts.hAlign)    rng.setHorizontalAlignment(opts.hAlign);
  if (opts.vAlign)    rng.setVerticalAlignment(opts.vAlign);
}

function dbColIndex(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}

function dbColLetter(index) {
  return String.fromCharCode(64 + index);
}
