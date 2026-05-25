/**
 * PROJECT DASHBOARD — WEEKLY REPORTING VIEW
 *
 * HOW TO RUN:
 *   1. Extensions → Apps Script → paste this into your create-dashboard file
 *   2. Run → createProjectDashboard
 *
 * Scans up to row 20 to find headers — works even when headers are not in row 1.
 * All column positions are detected from the actual header row automatically.
 */

// ─────────────────────────────────────────────────────────────────────────────
// DIAGNOSTIC — run this if something looks wrong
// ─────────────────────────────────────────────────────────────────────────────
function debugDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = '=== ALL TABS ===\n';
  ss.getSheets().forEach((s, i) => {
    msg += `\n[${i}] "${s.getName()}"  (${s.getLastRow()} rows)\n`;
    const found = findHeaderRow_(s);
    if (found) {
      msg += `  → Header row: ${found.rowNum}\n`;
      found.headers.forEach((h, j) => {
        if (h) msg += `    col ${String.fromCharCode(65+j)} (${j+1}): "${h}"\n`;
      });
    } else {
      msg += '  → No header row found in first 20 rows\n';
    }
  });
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────────────────────
function createProjectDashboard() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. FIND SOURCE SHEET + HEADER ROW ──────────────────────────────────────
  let src = null, headerInfo = null;

  for (const sheet of ss.getSheets()) {
    if (sheet.getName() === 'DASHBOARD') continue;
    const found = findHeaderRow_(sheet);
    if (found) { src = sheet; headerInfo = found; break; }
  }

  if (!src || !headerInfo) {
    SpreadsheetApp.getUi().alert(
      'Could not find a tab with WEEKLY and STATUS headers in the first 20 rows.\n' +
      'Run debugDashboard() to inspect all tabs.'
    );
    return;
  }

  const SRC_NAME  = src.getName();
  const HDR_ROW   = headerInfo.rowNum;       // 1-based row number of headers
  const DATA_START = HDR_ROW + 1;            // first row of actual data

  // ── 2. MAP HEADERS → COLUMN LETTERS ────────────────────────────────────────
  const rawHeaders = headerInfo.headers;

  function col(/* ...aliases */) {
    for (const alias of arguments) {
      const idx = rawHeaders.findIndex(
        h => String(h).toUpperCase().trim() === alias.toUpperCase().trim()
      );
      if (idx >= 0) return String.fromCharCode(65 + idx); // 0-based → A,B,C…
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

  const missing = ['discipline','task','status','weekly']
    .filter(k => !C[k]).map(k => k.toUpperCase());

  if (missing.length) {
    SpreadsheetApp.getUi().alert(
      `Missing required columns: ${missing.join(', ')}\n\n` +
      `Headers found in row ${HDR_ROW}: ${rawHeaders.filter(h=>h).join(' | ')}\n\n` +
      'Run debugDashboard() for the full layout.'
    );
    return;
  }

  // ── 3. FORMULA HELPERS ──────────────────────────────────────────────────────
  const Q = `'${SRC_NAME.replace(/'/g, "''")}'`;  // safely quoted sheet name

  // e.g. ref('M', 15) → 'Sheet'!M15:M
  function ref(letter, fromRow) {
    return fromRow ? `${Q}!${letter}${fromRow}:${letter}` : `${Q}!${letter}:${letter}`;
  }

  // ── 4. CREATE / RESET DASHBOARD ─────────────────────────────────────────────
  const DASH = 'DASHBOARD';
  let dash = ss.getSheetByName(DASH);
  if (dash) {
    dash.clear(); dash.clearFormats(); dash.clearConditionalFormatRules();
  } else {
    dash = ss.insertSheet(DASH, 0);
  }

  [175, 320, 115, 135, 92, 92, 135, 105, 270].forEach((w, i) => dash.setColumnWidth(i+1, w));
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
    value:'PROJECT DASHBOARD  ·  WEEKLY REPORTING VIEW',
    bg:NAVY, font:'#FFF', size:15, bold:true, h:'center', v:'middle',
  });

  // ── 7. ROW 2 — TIMESTAMP ────────────────────────────────────────────────────
  dash.setRowHeight(2, 20);
  put(dash, 'A2:I2', {
    formula:'="Last viewed: "&TEXT(NOW(),"MMM D, YYYY  ·  h:MM AM/PM")',
    bg:NAVY_LT, font:'#7A9BC0', size:8, h:'center', v:'middle',
  });

  // ── 8. ROW 3 — SPACER ───────────────────────────────────────────────────────
  dash.setRowHeight(3, 8);
  dash.getRange('A3:I3').setBackground(SEP);

  // ── 9. ROWS 4-5 — KPI CARDS ─────────────────────────────────────────────────
  dash.setRowHeight(4, 18);
  dash.setRowHeight(5, 46);

  const SR = ref(C.status);
  const PR = ref(C.priority);
  const WR = ref(C.weekly);

  const weeklyKpi =
    `=COUNTIFS(${WR},TRUE,${SR},"<>COMPLETED",${SR},"<>CANCELLED")`
  + `+COUNTIFS(${WR},"TRUE",${SR},"<>COMPLETED",${SR},"<>CANCELLED")`;

  const kpis = [
    { s:1, n:2, label:'IN PROGRESS',   lBg:'#BF4F00', nBg:'#E86A00',
      formula:`=COUNTIF(${SR},"IN PROGRESS")` },
    { s:3, n:2, label:'UPCOMING',      lBg:'#145A8C', nBg:'#1A72B8',
      formula:`=COUNTIF(${SR},"UPCOMING")` },
    { s:5, n:2, label:'1-HIGH ACTIVE', lBg:'#9B1C1C', nBg:'#C62828',
      formula:`=COUNTIFS(${PR},"1-HIGH",${SR},"<>COMPLETED",${SR},"<>CANCELLED")` },
    { s:7, n:2, label:'WEEKLY TRACKED',lBg:'#5B21B6', nBg:'#7C3AED',
      formula:weeklyKpi },
    { s:9, n:1, label:'COMPLETED',     lBg:'#065F46', nBg:'#059669',
      formula:`=COUNTIF(${SR},"COMPLETED")` },
  ];

  kpis.forEach(k => {
    const c1 = ltr(k.s), c2 = ltr(k.s + k.n - 1);
    put(dash, `${c1}4:${c2}4`, { value:k.label,    bg:k.lBg, font:'#FFF', size:8,  bold:true, h:'center', v:'middle' });
    put(dash, `${c1}5:${c2}5`, { formula:k.formula, bg:k.nBg, font:'#FFF', size:30, bold:true, h:'center', v:'middle' });
  });

  // ── 10. ROW 6 — SPACER ──────────────────────────────────────────────────────
  dash.setRowHeight(6, 12);
  dash.getRange('A6:I6').setBackground(SEP);

  // ── 11. ROW 7 — SECTION LABEL ───────────────────────────────────────────────
  dash.setRowHeight(7, 30);
  put(dash, 'A7:I7', {
    value:'   WEEKLY TASKS  —  ACTIVE PROJECT ITEMS',
    bg:NAVY, font:BLUE_ACC, size:11, bold:true, h:'left', v:'middle',
  });

  // ── 12. ROW 8 — COLUMN HEADERS ──────────────────────────────────────────────
  dash.setRowHeight(8, 22);
  ['DISCIPLINE','TASK','CONSULTANT','PERSON','START','END DATE','STATUS','PRIORITY','NOTES']
    .forEach((h, i) => dash.getRange(8, i+1)
      .setValue(h).setBackground(NAVY_MID).setFontColor('#FFF')
      .setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle'));

  // ── 13. ROW 9+ — WEEKLY TASKS TABLE ─────────────────────────────────────────
  // Output column order: DISCIPLINE, TASK, CONSULTANT, PERSON, START, END, STATUS, PRIORITY, NOTES
  // Sort: col 8 (PRIORITY) asc, col 6 (END DATE) asc
  const outCols = [C.discipline, C.task, C.consultant, C.person,
                   C.start, C.end, C.status, C.priority, C.notes]
    .map(letter => letter ? `${Q}!${letter}${DATA_START}:${letter}` : `""`);

  const WD = ref(C.weekly, DATA_START);
  const SD = ref(C.status, DATA_START);

  const filterFormula =
    `=IFERROR(` +
      `SORT(` +
        `FILTER(` +
          `{${outCols.join(',')}},` +
          `(${WD}=TRUE)+(${WD}="TRUE"),` +   // handles boolean TRUE and text "TRUE"
          `${SD}<>"COMPLETED",` +
          `${SD}<>"CANCELLED",` +
          `${SD}<>""` +                       // skips blank separator rows
        `)` +
      `,8,TRUE,6,TRUE),` +
    `"— No active weekly tasks —")`;

  dash.getRange('A9').setFormula(filterFormula);
  dash.getRange('A9:I400')
    .setFontSize(9).setFontFamily('Arial')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  dash.getRange('A9:A400').setFontWeight('bold').setFontColor(NAVY_MID);
  dash.setRowHeights(9, 392, 22);  // single API call instead of 392

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
  ].forEach(([t, bg, f]) => rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(f)
      .setRanges([statusRng]).build()
  ));

  [
    ['1-HIGH',   '#FEE2E2', '#991B1B'],
    ['2-MEDIUM', '#FEF3C7', '#92400E'],
    ['3-LOW',    '#DBEAFE', '#1E3A5F'],
  ].forEach(([t, bg, f]) => rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(f)
      .setRanges([priorityRng]).build()
  ));

  dash.setConditionalFormatRules(rules);

  // ── 16. SUCCESS ──────────────────────────────────────────────────────────────
  ss.setActiveSheet(dash);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    `✅  Dashboard created!\n\n` +
    `Source tab : "${SRC_NAME}"\n` +
    `Header row : ${HDR_ROW}\n` +
    `Data starts: row ${DATA_START}\n\n` +
    `Columns detected:\n` +
    Object.entries(C).map(([k,v]) => `  ${k.padEnd(12)}: ${v||'not found'}`).join('\n')
  );

  // ── LOCAL HELPERS ─────────────────────────────────────────────────────────────
  function put(sheet, a1, opts) {
    const rng = sheet.getRange(a1);
    if (a1.includes(':') && a1.split(':')[0] !== a1.split(':')[1]) rng.merge();
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

// ─────────────────────────────────────────────────────────────────────────────
// SHARED HELPER — scans rows 1-20 for a row containing both WEEKLY and STATUS
// ─────────────────────────────────────────────────────────────────────────────
function findHeaderRow_(sheet) {
  const maxScan = Math.min(20, sheet.getLastRow());
  if (maxScan === 0) return null;
  const data = sheet.getRange(1, 1, maxScan, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const upper = data[i].map(h => String(h).toUpperCase().trim());
    if (upper.includes('WEEKLY') && upper.includes('STATUS')) {
      return { rowNum: i + 1, headers: data[i] };
    }
  }
  return null;
}
