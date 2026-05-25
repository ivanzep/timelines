/**
 * PROJECT DASHBOARD — WEEKLY REPORTING VIEW
 *
 * STEP 1 — Run debugDashboard() first to confirm the exact tab name and columns.
 * STEP 2 — Update SRC below if needed, then run createProjectDashboard().
 *
 * All formulas are live — the dashboard auto-updates whenever the source tab changes.
 */

/**
 * Run this first to verify the sheet name and column layout.
 * Output appears in the Apps Script execution log (View → Logs).
 */
function debugDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let report = '=== ALL TAB NAMES (copy the exact one you want) ===\n';
  sheets.forEach((s, i) => {
    report += `  [${i}] "${s.getName()}"  (${s.getLastRow()} rows, ${s.getLastColumn()} cols)\n`;
  });

  // Try to find the task list tab automatically
  const candidates = sheets.filter(s =>
    s.getName().toUpperCase().includes('TASK') ||
    s.getName().toUpperCase().includes('PROJECT')
  );

  if (candidates.length > 0) {
    const s = candidates[0];
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    report += `\n=== HEADERS IN "${s.getName()}" (row 1) ===\n`;
    headers.forEach((h, i) => {
      if (h) report += `  Col ${columnLetter_(i+1)} (${i+1}): "${h}"\n`;
    });

    // Sample 3 data rows to show WEEKLY column type
    const weeklyIdx = headers.findIndex(h => String(h).toUpperCase() === 'WEEKLY');
    const statusIdx = headers.findIndex(h => String(h).toUpperCase() === 'STATUS');
    if (weeklyIdx >= 0) {
      const sample = s.getRange(2, 1, 5, s.getLastColumn()).getValues();
      report += `\n=== SAMPLE ROWS (WEEKLY col ${columnLetter_(weeklyIdx+1)}, STATUS col ${columnLetter_(statusIdx+1)}) ===\n`;
      sample.forEach((row, i) => {
        const weekly = row[weeklyIdx];
        const status = row[statusIdx];
        const task   = row[2] || '';
        report += `  Row ${i+2}: WEEKLY="${weekly}" (${typeof weekly})  STATUS="${status}"  TASK="${String(task).substring(0,40)}"\n`;
      });
    }
  }

  Logger.log(report);
  SpreadsheetApp.getUi().alert(report);
}

function createProjectDashboard() {

  // ── CONFIG ───────────────────────────────────────────────────────────────
  const SRC  = 'PROJECT TASK LIST';
  const DASH = 'DASHBOARD';

  // Colors
  const NAVY     = '#162032';
  const NAVY_LT  = '#1E3050';
  const NAVY_MID = '#2A4A6E';
  const BLUE_ACC = '#5AAFF0';
  const SEP      = '#DCE5F0';

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

  // WEEKLY card uses summed COUNTIFS to handle both boolean TRUE and text "TRUE"
  const weeklyFormula =
    `=COUNTIFS('${SRC}'!N:N,TRUE,'${SRC}'!L:L,"<>COMPLETED",'${SRC}'!L:L,"<>CANCELLED")` +
    `+COUNTIFS('${SRC}'!N:N,"TRUE",'${SRC}'!L:L,"<>COMPLETED",'${SRC}'!L:L,"<>CANCELLED")`;

  const kpis = [
    {
      start: 1, span: 2,
      label: 'IN PROGRESS',
      formula: `=COUNTIF('${SRC}'!L:L,"IN PROGRESS")`,
      lBg: '#BF4F00', nBg: '#E86A00',
    },
    {
      start: 3, span: 2,
      label: 'UPCOMING',
      formula: `=COUNTIF('${SRC}'!L:L,"UPCOMING")`,
      lBg: '#145A8C', nBg: '#1A72B8',
    },
    {
      start: 5, span: 2,
      label: '1-HIGH ACTIVE',
      formula: `=COUNTIFS('${SRC}'!M:M,"1-HIGH",'${SRC}'!L:L,"<>COMPLETED",'${SRC}'!L:L,"<>CANCELLED")`,
      lBg: '#9B1C1C', nBg: '#C62828',
    },
    {
      start: 7, span: 2,
      label: 'WEEKLY TRACKED',
      formula: weeklyFormula,
      lBg: '#5B21B6', nBg: '#7C3AED',
    },
    {
      start: 9, span: 1,
      label: 'COMPLETED',
      formula: `=COUNTIF('${SRC}'!L:L,"COMPLETED")`,
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

  // ── ROW 9+ — WEEKLY TASKS (live FILTER) ──────────────────────────────────
  //
  // FILTER handles both boolean TRUE and text "TRUE" in the WEEKLY column (N).
  //   (N2:N=TRUE)+(N2:N="TRUE")  →  1 or 2 for any truthy value, 0 otherwise
  //
  // Source column map:
  //   A=DISCIPLINE  C=TASK  D=CONSULTANT  E=PERSON
  //   G=START  H=END  L=STATUS  M=PRIORITY  T=NOTES  N=WEEKLY
  //
  // Sorted by: PRIORITY asc (col 8), then END DATE asc (col 6)

  const s = `'${SRC}'!`;
  const filterFormula =
    `=IFERROR(` +
      `SORT(` +
        `FILTER(` +
          `{${s}A2:A,${s}C2:C,${s}D2:D,${s}E2:E,${s}G2:G,${s}H2:H,${s}L2:L,${s}M2:M,${s}T2:T},` +
          `(${s}N2:N=TRUE)+(${s}N2:N="TRUE"),` +   // WEEKLY = true (boolean or text)
          `${s}L2:L<>"COMPLETED",` +
          `${s}L2:L<>"CANCELLED",` +
          `${s}L2:L<>""` +                          // exclude blank/separator rows
        `)` +
      `,8,TRUE,6,TRUE),` +                          // sort: priority asc, end date asc
    `"— No active weekly tasks —")`;

  dash.getRange('A9').setFormula(filterFormula);
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

  // ── LOCAL HELPERS ─────────────────────────────────────────────────────────
  function mergeSet(sheet, a1, opts) {
    const rng = sheet.getRange(a1);
    const parts = a1.split(':');
    if (parts.length === 2 && parts[0] !== parts[1]) rng.merge();
    if (opts.formula)                  rng.setFormula(opts.formula);
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

// Shared helper (outside both functions so debugDashboard can use it)
function columnLetter_(n) { return String.fromCharCode(64 + n); }
