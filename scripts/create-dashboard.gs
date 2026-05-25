/**
 * PROJECT DASHBOARD — WEEKLY REPORTING VIEW
 *
 * HOW TO RUN:
 *   1. Open your Google Spreadsheet
 *   2. Extensions → Apps Script → add a new script file, paste this code
 *   3. Click Run → createProjectDashboard
 *   4. Grant permissions when prompted
 *   5. A "DASHBOARD" tab is created (or replaced if it already exists)
 *
 * All formulas (COUNTIF / QUERY) are live — the dashboard auto-updates
 * whenever data in "PROJECT TASK LIST" changes.
 */

function createProjectDashboard() {

  // ── CONFIG (all locals — no global scope pollution) ──────────────────────
  const SRC  = 'PROJECT TASK LIST';
  const DASH = 'DASHBOARD';

  // Source column letters
  const STATUS   = 'L';
  const PRIORITY = 'M';
  const WEEKLY   = 'N';

  // Colors
  const NAVY       = '#162032';
  const NAVY_LT    = '#1E3050';
  const NAVY_MID   = '#2A4A6E';
  const BLUE_ACC   = '#5AAFF0';
  const SEP        = '#DCE5F0';

  // ── SHEET SETUP ──────────────────────────────────────────────────────────
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SRC);

  if (!src) {
    SpreadsheetApp.getUi().alert(`Sheet "${SRC}" not found.\nCheck the tab name.`);
    return;
  }

  let dash = ss.getSheetByName(DASH);
  if (dash) {
    dash.clear();
    dash.clearFormats();
    dash.clearConditionalFormatRules();
  } else {
    dash = ss.insertSheet(DASH, 0);
  }

  // Column widths for 9 data columns (A–I)
  [175, 320, 115, 135, 92, 92, 135, 105, 270].forEach((w, i) => dash.setColumnWidth(i + 1, w));
  if (dash.getMaxColumns() > 9) dash.hideColumns(10, dash.getMaxColumns() - 9);

  // ── ROW 1 — TITLE ────────────────────────────────────────────────────────
  dash.setRowHeight(1, 52);
  mergeSet(dash, 'A1:I1', {
    value: 'PROJECT DASHBOARD  ·  WEEKLY REPORTING VIEW',
    bg: NAVY, font: '#FFFFFF', size: 15, bold: true, h: 'center', v: 'middle',
  });

  // ── ROW 2 — LIVE TIMESTAMP ───────────────────────────────────────────────
  dash.setRowHeight(2, 20);
  mergeSet(dash, 'A2:I2', {
    formula: '="Last viewed: "&TEXT(NOW(),"MMM D, YYYY  ·  h:MM AM/PM")',
    bg: NAVY_LT, font: '#7A9BC0', size: 8, h: 'center', v: 'middle',
  });

  // ── ROW 3 — SPACER ───────────────────────────────────────────────────────
  dash.setRowHeight(3, 8);
  dash.getRange('A3:I3').setBackground(SEP);

  // ── ROWS 4-5 — KPI CARDS ─────────────────────────────────────────────────
  dash.setRowHeight(4, 18);
  dash.setRowHeight(5, 46);

  const kpis = [
    {
      start: 1, span: 2,
      label: 'IN PROGRESS',
      formula: `=COUNTIF('${SRC}'!${STATUS}:${STATUS},"IN PROGRESS")`,
      lBg: '#BF4F00', nBg: '#E86A00',
    },
    {
      start: 3, span: 2,
      label: 'UPCOMING',
      formula: `=COUNTIF('${SRC}'!${STATUS}:${STATUS},"UPCOMING")`,
      lBg: '#145A8C', nBg: '#1A72B8',
    },
    {
      start: 5, span: 2,
      label: '1-HIGH ACTIVE',
      formula: `=COUNTIFS('${SRC}'!${PRIORITY}:${PRIORITY},"1-HIGH",`
             + `'${SRC}'!${STATUS}:${STATUS},"<>COMPLETED",`
             + `'${SRC}'!${STATUS}:${STATUS},"<>CANCELLED")`,
      lBg: '#9B1C1C', nBg: '#C62828',
    },
    {
      start: 7, span: 2,
      label: 'WEEKLY TRACKED',
      formula: `=COUNTIFS('${SRC}'!${WEEKLY}:${WEEKLY},TRUE,`
             + `'${SRC}'!${STATUS}:${STATUS},"<>COMPLETED",`
             + `'${SRC}'!${STATUS}:${STATUS},"<>CANCELLED")`,
      lBg: '#5B21B6', nBg: '#7C3AED',
    },
    {
      start: 9, span: 1,
      label: 'COMPLETED',
      formula: `=COUNTIF('${SRC}'!${STATUS}:${STATUS},"COMPLETED")`,
      lBg: '#065F46', nBg: '#059669',
    },
  ];

  kpis.forEach(k => {
    const c1 = colLtr(k.start);
    const c2 = colLtr(k.start + k.span - 1);
    mergeSet(dash, `${c1}4:${c2}4`, { value: k.label, bg: k.lBg, font: '#FFF', size: 8, bold: true, h: 'center', v: 'middle' });
    mergeSet(dash, `${c1}5:${c2}5`, { formula: k.formula, bg: k.nBg, font: '#FFF', size: 30, bold: true, h: 'center', v: 'middle' });
  });

  // ── ROW 6 — SPACER ───────────────────────────────────────────────────────
  dash.setRowHeight(6, 12);
  dash.getRange('A6:I6').setBackground(SEP);

  // ── ROW 7 — SECTION LABEL ────────────────────────────────────────────────
  dash.setRowHeight(7, 30);
  mergeSet(dash, 'A7:I7', {
    value: '   WEEKLY TASKS  —  ACTIVE PROJECT ITEMS',
    bg: NAVY, font: BLUE_ACC, size: 11, bold: true, h: 'left', v: 'middle',
  });

  // ── ROW 8 — COLUMN HEADERS ───────────────────────────────────────────────
  dash.setRowHeight(8, 22);
  ['DISCIPLINE','TASK','CONSULTANT','PERSON','START','END DATE','STATUS','PRIORITY','NOTES']
    .forEach((h, i) => dash.getRange(8, i + 1)
      .setValue(h)
      .setBackground(NAVY_MID).setFontColor('#FFF')
      .setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle'));

  // ── ROW 9+ — WEEKLY TASKS (live QUERY) ───────────────────────────────────
  // Source: A=DISCIPLINE C=TASK D=CONSULTANT E=PERSON G=START H=END
  //         L=STATUS M=PRIORITY N=WEEKLY T=NOTES
  // Filter: WEEKLY=true, STATUS not COMPLETED/CANCELLED
  // Sort:   PRIORITY asc, END DATE asc
  const query =
    `=IFERROR(QUERY('${SRC}'!A:T,` +
    `"SELECT A,C,D,E,G,H,L,M,T ` +
    `WHERE N=true ` +
    `AND L<>'COMPLETED' ` +
    `AND L<>'CANCELLED' ` +
    `ORDER BY M ASC,H ASC ` +
    `LABEL A'',C'',D'',E'',G'',H'',L'',M'',T''",0),` +
    `"— No active weekly tasks —")`;

  dash.getRange('A9').setFormula(query);
  dash.getRange('A9:I400')
    .setFontSize(9).setFontFamily('Arial')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  dash.getRange('A9:A400').setFontWeight('bold').setFontColor(NAVY_MID);
  for (let r = 9; r <= 400; r++) dash.setRowHeight(r, 22);

  // ── FREEZE HEADER ROWS ───────────────────────────────────────────────────
  dash.setFrozenRows(8);

  // ── CONDITIONAL FORMATTING ───────────────────────────────────────────────
  const statusRng   = dash.getRange('G9:G400');
  const priorityRng = dash.getRange('H9:H400');
  const rules = [];

  [
    ['IN PROGRESS',  '#FFF0D6', '#7A3800'],
    ['UPCOMING',     '#DBEAFE', '#1E3A5F'],
    ['DOWNSTREAM',   '#E5E7EB', '#374151'],
    ['PENDING',      '#FEF9C3', '#713F12'],
    ['75% COMPLETE', '#D1FAE5', '#064E3B'],
  ].forEach(([text, bg, font]) =>
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(font)
      .setRanges([statusRng]).build()));

  [
    ['1-HIGH',   '#FEE2E2', '#991B1B'],
    ['2-MEDIUM', '#FEF3C7', '#92400E'],
    ['3-LOW',    '#DBEAFE', '#1E3A5F'],
  ].forEach(([text, bg, font]) =>
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(font)
      .setRanges([priorityRng]).build()));

  dash.setConditionalFormatRules(rules);

  // ── DONE ─────────────────────────────────────────────────────────────────
  ss.setActiveSheet(dash);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅  Dashboard created!\n\n' +
    'The WEEKLY TASKS table updates automatically whenever\n' +
    '"PROJECT TASK LIST" data changes — no refresh needed.'
  );

  // ── LOCAL HELPERS (inside function — zero global footprint) ──────────────
  function mergeSet(sheet, a1, opts) {
    const rng = sheet.getRange(a1);
    const parts = a1.split(':');
    if (parts.length === 2 && parts[0] !== parts[1]) rng.merge();
    if (opts.formula)               rng.setFormula(opts.formula);
    else if (opts.value !== undefined) rng.setValue(opts.value);
    if (opts.bg)   rng.setBackground(opts.bg);
    if (opts.font) rng.setFontColor(opts.font);
    if (opts.size) rng.setFontSize(opts.size);
    if (opts.bold) rng.setFontWeight('bold');
    if (opts.h)    rng.setHorizontalAlignment(opts.h);
    if (opts.v)    rng.setVerticalAlignment(opts.v);
  }

  function colLtr(n) { return String.fromCharCode(64 + n); }
}
