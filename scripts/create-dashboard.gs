/**
 * PROJECT DASHBOARD — WEEKLY REPORTING VIEW
 *
 * HOW TO RUN:
 *   1. Extensions → Apps Script → paste this into your create-dashboard file
 *   2. Run → createProjectDashboard
 *
 * The script auto-detects the source tab and every column position by
 * reading row 1 headers — no hardcoded column letters anywhere.
 */

// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — run this first if the dashboard shows all zeros
// It logs every tab name and every header it finds.
// ─────────────────────────────────────────────────────────────────────────────
function debugDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = '=== TAB NAMES ===\n';
  ss.getSheets().forEach((s, i) => {
    msg += `  [${i}] "${s.getName()}"\n`;
    const headers = s.getRange(1, 1, 1, Math.min(s.getLastColumn(), 30)).getValues()[0];
    headers.forEach((h, j) => {
      if (h) msg += `        col ${j+1} (${String.fromCharCode(65+j)}): "${h}"\n`;
    });
  });
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────────────────────
function createProjectDashboard() {

  // ── 1. FIND SOURCE SHEET (auto-detect by looking for WEEKLY + STATUS headers)
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheets().find(sheet => {
    const hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                      .map(h => String(h).toUpperCase().trim());
    return hdrs.includes('WEEKLY') && hdrs.includes('STATUS');
  });

  if (!src) {
    SpreadsheetApp.getUi().alert(
      'Could not find a tab with both WEEKLY and STATUS column headers.\n' +
      'Run debugDashboard() to see all tab names and their headers.'
    );
    return;
  }

  const SRC_NAME = src.getName();

  // ── 2. MAP HEADERS → COLUMN LETTERS ────────────────────────────────────────
  const rawHeaders = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];

  // Returns the spreadsheet column letter for the first matching header name.
  // Pass multiple aliases in priority order.
  function col() {
    for (const alias of arguments) {
      const idx = rawHeaders.findIndex(
        h => String(h).toUpperCase().trim() === alias.toUpperCase().trim()
      );
      if (idx >= 0) return String.fromCharCode(65 + idx);
    }
    return null;
  }

  const C = {
    discipline : col('DISCIPLINE'),
    task       : col('TASK'),
    consultant : col('CONSULTANT'),
    person     : col('PERSON'),
    start      : col('START DATE', 'START'),
    end        : col('END DATE', 'END'),
    status     : col('STATUS'),
    priority   : col('PRIORITY'),
    weekly     : col('WEEKLY'),
    notes      : col('NOTES'),
  };

  // Abort if critical columns are missing
  const missing = Object.entries(C)
    .filter(([k, v]) => !v && ['status','weekly','task','discipline'].includes(k))
    .map(([k]) => k.toUpperCase());

  if (missing.length) {
    SpreadsheetApp.getUi().alert(
      `Missing required columns: ${missing.join(', ')}\n\n` +
      `Headers found: ${rawHeaders.filter(h=>h).join(' | ')}\n\n` +
      'Run debugDashboard() to see the full column list.'
    );
    return;
  }

  // ── 3. BUILD FORMULA HELPERS ────────────────────────────────────────────────
  // Safely quote the sheet name for use inside formula strings
  const Q = `'${SRC_NAME.replace(/'/g, "''")}'`;

  // Full-column reference: e.g. range(C.status) → 'Sheet'!L:L
  function range(letter, fromRow) {
    if (!letter) return null;
    return fromRow
      ? `${Q}!${letter}${fromRow}:${letter}`
      : `${Q}!${letter}:${letter}`;
  }

  // ── 4. CREATE / RESET DASHBOARD SHEET ──────────────────────────────────────
  const DASH = 'DASHBOARD';
  let dash = ss.getSheetByName(DASH);
  if (dash) {
    dash.clear();
    dash.clearFormats();
    dash.clearConditionalFormatRules();
  } else {
    dash = ss.insertSheet(DASH, 0);
  }

  const COL_WIDTHS = [175, 320, 115, 135, 92, 92, 135, 105, 270];
  COL_WIDTHS.forEach((w, i) => dash.setColumnWidth(i + 1, w));
  if (dash.getMaxColumns() > 9) dash.hideColumns(10, dash.getMaxColumns() - 9);

  // ── 5. COLORS ───────────────────────────────────────────────────────────────
  const NAVY     = '#162032';
  const NAVY_LT  = '#1E3050';
  const NAVY_MID = '#2A4A6E';
  const BLUE_ACC = '#5AAFF0';
  const SEP      = '#DCE5F0';

  // ── 6. ROW 1 — TITLE ────────────────────────────────────────────────────────
  dash.setRowHeight(1, 52);
  put(dash, 'A1:I1', {
    value: 'PROJECT DASHBOARD  ·  WEEKLY REPORTING VIEW',
    bg: NAVY, font: '#FFF', size: 15, bold: true, h: 'center', v: 'middle',
  });

  // ── 7. ROW 2 — TIMESTAMP ────────────────────────────────────────────────────
  dash.setRowHeight(2, 20);
  put(dash, 'A2:I2', {
    formula: '="Last viewed: "&TEXT(NOW(),"MMM D, YYYY  ·  h:MM AM/PM")',
    bg: NAVY_LT, font: '#7A9BC0', size: 8, h: 'center', v: 'middle',
  });

  // ── 8. ROW 3 — SPACER ───────────────────────────────────────────────────────
  dash.setRowHeight(3, 8);
  dash.getRange('A3:I3').setBackground(SEP);

  // ── 9. ROWS 4-5 — KPI CARDS ─────────────────────────────────────────────────
  dash.setRowHeight(4, 18);
  dash.setRowHeight(5, 46);

  const sr = range(C.status);
  const pr = range(C.priority);
  const wr = range(C.weekly);

  // WEEKLY card sums both boolean TRUE and text "TRUE" to handle either storage type
  const weeklyKpi = wr
    ? `=COUNTIFS(${wr},TRUE,${sr},"<>COMPLETED",${sr},"<>CANCELLED")`
    + `+COUNTIFS(${wr},"TRUE",${sr},"<>COMPLETED",${sr},"<>CANCELLED")`
    : '=0';

  const kpis = [
    { start:1, span:2, label:'IN PROGRESS',   lBg:'#BF4F00', nBg:'#E86A00',
      formula: sr ? `=COUNTIF(${sr},"IN PROGRESS")` : '=0' },
    { start:3, span:2, label:'UPCOMING',       lBg:'#145A8C', nBg:'#1A72B8',
      formula: sr ? `=COUNTIF(${sr},"UPCOMING")` : '=0' },
    { start:5, span:2, label:'1-HIGH ACTIVE',  lBg:'#9B1C1C', nBg:'#C62828',
      formula: (sr && pr)
        ? `=COUNTIFS(${pr},"1-HIGH",${sr},"<>COMPLETED",${sr},"<>CANCELLED")`
        : '=0' },
    { start:7, span:2, label:'WEEKLY TRACKED', lBg:'#5B21B6', nBg:'#7C3AED',
      formula: weeklyKpi },
    { start:9, span:1, label:'COMPLETED',      lBg:'#065F46', nBg:'#059669',
      formula: sr ? `=COUNTIF(${sr},"COMPLETED")` : '=0' },
  ];

  kpis.forEach(k => {
    const c1 = ltr(k.start), c2 = ltr(k.start + k.span - 1);
    put(dash, `${c1}4:${c2}4`, { value:k.label,    bg:k.lBg, font:'#FFF', size:8,  bold:true, h:'center', v:'middle' });
    put(dash, `${c1}5:${c2}5`, { formula:k.formula, bg:k.nBg, font:'#FFF', size:30, bold:true, h:'center', v:'middle' });
  });

  // ── 10. ROW 6 — SPACER ──────────────────────────────────────────────────────
  dash.setRowHeight(6, 12);
  dash.getRange('A6:I6').setBackground(SEP);

  // ── 11. ROW 7 — SECTION LABEL ───────────────────────────────────────────────
  dash.setRowHeight(7, 30);
  put(dash, 'A7:I7', {
    value: '   WEEKLY TASKS  —  ACTIVE PROJECT ITEMS',
    bg: NAVY, font: BLUE_ACC, size: 11, bold: true, h: 'left', v: 'middle',
  });

  // ── 12. ROW 8 — COLUMN HEADERS ──────────────────────────────────────────────
  dash.setRowHeight(8, 22);
  ['DISCIPLINE','TASK','CONSULTANT','PERSON','START','END DATE','STATUS','PRIORITY','NOTES']
    .forEach((h, i) => dash.getRange(8, i+1)
      .setValue(h).setBackground(NAVY_MID).setFontColor('#FFF')
      .setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle'));

  // ── 13. ROW 9+ — WEEKLY TASKS TABLE ─────────────────────────────────────────
  // Build output column array using detected positions.
  // Falls back to empty string column if a column wasn't found.
  const outCols = [C.discipline, C.task, C.consultant, C.person,
                   C.start, C.end, C.status, C.priority, C.notes]
    .map(letter => letter ? `${Q}!${letter}2:${letter}` : '""');

  // SORT indices in the 9-column output: col 8 = PRIORITY, col 6 = END DATE
  const filterFormula =
    `=IFERROR(` +
      `SORT(` +
        `FILTER(` +
          `{${outCols.join(',')}},` +
          `(${range(C.weekly, 2)}=TRUE)+(${range(C.weekly, 2)}="TRUE"),` +
          `${range(C.status, 2)}<>"COMPLETED",` +
          `${range(C.status, 2)}<>"CANCELLED",` +
          `${range(C.status, 2)}<>""` +
        `)` +
      `,8,TRUE,6,TRUE),` +
    `"— No active weekly tasks —")`;

  dash.getRange('A9').setFormula(filterFormula);
  dash.getRange('A9:I400')
    .setFontSize(9).setFontFamily('Arial')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  dash.getRange('A9:A400').setFontWeight('bold').setFontColor(NAVY_MID);
  for (let r = 9; r <= 400; r++) dash.setRowHeight(r, 22);

  // ── 14. FREEZE ───────────────────────────────────────────────────────────────
  dash.setFrozenRows(8);

  // ── 15. CONDITIONAL FORMATTING ───────────────────────────────────────────────
  const statusRng   = dash.getRange('G9:G400');
  const priorityRng = dash.getRange('H9:H400');
  const rules = [];

  [
    ['IN PROGRESS',  '#FFF0D6', '#7A3800'],
    ['UPCOMING',     '#DBEAFE', '#1E3A5F'],
    ['DOWNSTREAM',   '#E5E7EB', '#374151'],
    ['PENDING',      '#FEF9C3', '#713F12'],
    ['75% COMPLETE', '#D1FAE5', '#064E3B'],
  ].forEach(([t, bg, font]) => rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(font)
      .setRanges([statusRng]).build()
  ));

  [
    ['1-HIGH',   '#FEE2E2', '#991B1B'],
    ['2-MEDIUM', '#FEF3C7', '#92400E'],
    ['3-LOW',    '#DBEAFE', '#1E3A5F'],
  ].forEach(([t, bg, font]) => rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(font)
      .setRanges([priorityRng]).build()
  ));

  dash.setConditionalFormatRules(rules);

  // ── 16. DONE ─────────────────────────────────────────────────────────────────
  ss.setActiveSheet(dash);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    `✅  Dashboard created!\n\nSource tab: "${SRC_NAME}"\n\n` +
    'Column positions detected automatically from row 1 headers.\n' +
    'The table updates whenever the source data changes.'
  );

  // ── LOCAL HELPERS ─────────────────────────────────────────────────────────────
  function put(sheet, a1, opts) {
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

  function ltr(n) { return String.fromCharCode(64 + n); }
}
