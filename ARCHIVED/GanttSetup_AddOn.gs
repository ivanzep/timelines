// ============================================================
//  GANTT TIMELINE INSTALLER
//  Google Workspace Editor Add-on
//  File: GanttSetup_AddOn.gs
//
//  WHAT THIS DOES
//  ──────────────
//  Install this add-on once → a "Gantt Timeline" menu item appears
//  automatically under Extensions in EVERY Google Spreadsheet you open.
//
//  Click "🚀 Setup Gantt on This Sheet" to:
//    1. Create a new container-bound Apps Script project in the current sheet
//    2. Push the full Gantt sync code (Code_V1.13) to it
//    3. Set the web app manifest (executeAs: Me, access: Anyone)
//    4. Create a versioned snapshot
//    5. Deploy it as a Web App
//    6. Show a copyable URL dialog → paste that URL into the Gantt HTML ⚙ Setup
//
//  HOW TO INSTALL  (one time only — ~3 minutes)
//  ─────────────────────────────────────────────
//  1. Go to script.google.com → New project → name it "Gantt Timeline Installer"
//  2. Delete the default Code.gs content, paste THIS entire file
//  3. Click the gear icon (⚙ Project Settings) → check
//     "Show 'appsscript.json' manifest file in editor"
//  4. Click appsscript.json and REPLACE its contents with:
//
//       {
//         "timeZone": "America/New_York",
//         "dependencies": {},
//         "exceptionLogging": "STACKDRIVER",
//         "runtimeVersion": "V8",
//         "oauthScopes": [
//           "https://www.googleapis.com/auth/spreadsheets.currentonly",
//           "https://www.googleapis.com/auth/script.projects",
//           "https://www.googleapis.com/auth/script.deployments",
//           "https://www.googleapis.com/auth/script.external_request"
//         ]
//       }
//
//  5. Save (Ctrl+S)
//  6. You must also enable the Apps Script API for your account:
//     → script.google.com/home/usersettings  →  toggle "Google Apps Script API" ON
//  7. Back in the editor: Deploy → New deployment
//     → Type: Editor Add-on  → Description: "Gantt Timeline Installer"  → Deploy
//  8. On the "Test deployments" screen that opens, click "Install"
//     and follow the permission prompts
//
//  After step 8, reload any Google Spreadsheet.
//  You will see: Extensions → 📊 Gantt Timeline → 🚀 Setup Gantt on This Sheet
//  ============================================================

// ─────────────────────────────────────────────────────────────────────────────
//  EMBEDDED SYNC CODE  (Code_V1.13 — injected into every new spreadsheet)
//  Uses backtick template literal (V8 runtime). The sync code itself contains
//  no backtick characters so no escaping is needed.
// ─────────────────────────────────────────────────────────────────────────────
var GANTT_SYNC_CODE = `
// ============================================================
//  LA COSTA HOTEL — TIMELINE SYNC  |  Google Apps Script
//  Auto-installed by Gantt Timeline Installer add-on (V1.13)
// ============================================================

// ---- Tab names ----
var SOURCE_SHEET      = 'PROJECT TASK LIST';
var TASK_PARAMS_SHEET = 'GANTT TASK PARAMS-DO NOT EDIT';
var SETTINGS_SHEET    = 'GANTT SETTINGS-DO NOT EDIT';

// ============================================================
//  SPREADSHEET MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Gantt Timeline')
    .addItem('Get Web App URL',    'showWebAppUrl')
    .addSeparator()
    .addItem('About / Setup Help', 'showSetupHelp')
    .addToUi();
}

function showWebAppUrl() {
  var ui  = SpreadsheetApp.getUi();
  var svc = ScriptApp.getService();
  var url = svc ? svc.getUrl() : null;
  if (!url) {
    try {
      var ss  = SpreadsheetApp.getActiveSpreadsheet();
      var sht = ss.getSheetByName(SETTINGS_SHEET);
      if (sht) {
        var data = sht.getDataRange().getValues();
        for (var i = 0; i < data.length; i++) {
          if (String(data[i][0]).trim() === 'appScriptURL' && data[i][1]) {
            url = String(data[i][1]).trim(); break;
          }
        }
      }
    } catch (e) {}
  }
  if (url) {
    var html = HtmlService.createHtmlOutput(
      '<style>body{font-family:Arial,sans-serif;padding:16px;font-size:13px;color:#1e293b}' +
      'h3{margin:0 0 10px}p{margin:4px 0 10px;color:#475569;font-size:12px}' +
      '#url{width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:6px;font-size:11px;' +
      'background:#f8fafc;box-sizing:border-box;resize:none;height:52px}' +
      'button{margin-top:10px;padding:7px 18px;background:#0ea5e9;color:#fff;border:none;' +
      'border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;width:100%}' +
      'button:hover{background:#0284c7}' +
      '#msg{font-size:11px;color:#16a34a;margin-top:6px;text-align:center;min-height:16px}</style>' +
      '<h3>🔗 Web App URL</h3>' +
      '<p>Paste this into the Gantt HTML tool\'s ⚙ Setup panel.</p>' +
      '<textarea id="url" readonly>' + url + '</textarea>' +
      '<button onclick="copy()">Copy to Clipboard</button><div id="msg"></div>' +
      '<script>function copy(){var el=document.getElementById("url");el.select();' +
      'el.setSelectionRange(0,9999);document.execCommand("copy");' +
      'document.getElementById("msg").textContent="✓ Copied!";}<\/script>'
    ).setWidth(420).setHeight(200);
    ui.showModalDialog(html, 'Gantt Timeline — Web App URL');
  } else {
    ui.alert('Not Deployed', 'This script has not been deployed as a Web App yet.\\n\\n' +
      'Deploy via Extensions → Apps Script → Deploy → New deployment.', ui.ButtonSet.OK);
  }
}

function showSetupHelp() {
  SpreadsheetApp.getUi().alert('Gantt Timeline',
    'This script was auto-installed by the Gantt Timeline Installer add-on.\\n\\n' +
    'To connect the Gantt HTML tool:\\n' +
    '  1. Use "Get Web App URL" from this menu\\n' +
    '  2. Paste the URL into the HTML tool\'s ⚙ Setup panel\\n' +
    '  3. Click Connect', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---- Status → colour ----
var STATUS_COLORS = {
  'IN PROGRESS':'#16a34a','UPCOMING':'#f59e0b','DOWNSTREAM':'#8b5cf6',
  'PENDING':'#d9c34a','COMPLETED':'#94a3b8','CANCELLED':'#dc2626',
  '75% COMPLETE':'#22c55e','WAITING ON OTHERS':'#f97316','WITING ON THE CITY':'#f97316',
  'UNDER REVIEW':'#06b6d4','URGENT':'#dc2626','ON HOLD':'#94a3b8',
  'STAND BY':'#94a3b8','NEEDS ATTENTION':'#f97316','TBD':'#94a3b8',
  'NOTE':'#60a5fa','RECEIVED':'#22c55e','APPROVED':'#16a34a','DECLINED':'#dc2626'
};

// ============================================================
//  GET
// ============================================================
function doGet(e) {
  try {
    var result = importFromTaskList();
    var taskParams = readTaskParams();
    if (Object.keys(taskParams).length > 0) {
      result.tasks.forEach(function(t) {
        var key = normKey((t.group||'')+'|'+(t.name||''));
        var p = taskParams[key]; if (!p) return;
        if (p.color)  t.colorOverride = p.color;
        if (p.type)   t.type = p.type;
        if (p.style === 'dashed')         { t.dashed='true';  t.dashedOutline='false'; }
        if (p.style === 'dashed-outline') { t.dashed='false'; t.dashedOutline='true';  }
        if (p.style === 'solid')          { t.dashed='false'; t.dashedOutline='false'; }
        if (p.symbol) t.symbol = p.symbol;
        if (p.deps)   t.dependencies = p.deps;
      });
    }
    result.settings = readSettings();
    return buildResponse(result);
  } catch (err) { return buildResponse({ success:false, error:err.toString() }); }
}

// ============================================================
//  POST
// ============================================================
function doPost(e) {
  var taskMsg='', taskErr='', taskParamsErr='', settingsErr='';
  try {
    var payload = JSON.parse(e.postData.contents);
    try { taskMsg = saveBackToTaskList(payload); } catch(tErr) { taskErr = tErr.toString(); }
    if (payload.tasks) {
      try { writeTaskParams(payload.tasks); } catch(tpErr) { taskParamsErr = tpErr.toString(); }
    }
    if (payload.settings) {
      try { writeSettings(payload.settings); } catch(sErr) { settingsErr = sErr.toString(); }
    }
    return buildResponse({ success:true, message:taskMsg,
      taskError:taskErr, taskParamsError:taskParamsErr, settingsError:settingsErr });
  } catch(err) { return buildResponse({ success:false, error:err.toString() }); }
}

function buildResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  IMPORT
// ============================================================
function importFromTaskList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"'+SOURCE_SHEET+'" tab not found.');
  var raw = sheet.getDataRange().getValues();
  var hRow = -1;
  for (var i = 0; i < Math.min(raw.length,25); i++) {
    for (var c = 0; c < raw[i].length; c++) {
      if (String(raw[i][c]).trim().toUpperCase() === 'DISCIPLINE') { hRow=i; break; }
    }
    if (hRow >= 0) break;
  }
  if (hRow < 0) throw new Error('DISCIPLINE header not found in "'+SOURCE_SHEET+'".');
  var cols = {};
  raw[hRow].forEach(function(h,idx){ cols[String(h).trim().toUpperCase()]=idx; });
  var CI = {
    discipline:ci(cols,'DISCIPLINE',1), id:ci(cols,'ID',2), task:ci(cols,'TASK',3),
    consultant:ci(cols,'CONSULTANT',4), person:ci(cols,'PERSON',5),
    start:ci(cols,'START DATE',7), end:ci(cols,'END DATE',8),
    schedule:ci(cols,'SCHEDULE',10), milestone:ci(cols,'MILESTONE',11),
    status:ci(cols,'STATUS',12), priority:ci(cols,'PRIORITY',13), notes:ci(cols,'NOTES',20)
  };
  var tasks=[], idCounter=1, lastDisc='';
  for (var row=hRow+1; row<raw.length; row++) {
    var r=raw[row];
    var disc=String(r[CI.discipline]||'').trim(); if(disc) lastDisc=disc;
    var taskName=String(r[CI.task]||'').trim();
    if (!taskName||/^[\\s\\-]+$/.test(taskName)) continue;
    var isSched=(r[CI.schedule]===true||String(r[CI.schedule]).toUpperCase()==='TRUE');
    var isMil  =(r[CI.milestone]===true||String(r[CI.milestone]).toUpperCase()==='TRUE');
    if (!isSched&&!isMil) continue;
    var sd=parseSheetDate(r[CI.start]), ed=parseSheetDate(r[CI.end]);
    if (!sd||!ed) continue;
    if (ed.getTime()<sd.getTime()) continue;
    var status=String(r[CI.status]||'').trim();
    tasks.push({
      id:idCounter++, name:taskName, group:lastDisc||disc||'UNCATEGORIZED',
      type:isMil?'milestone':'bar', start:fmtDate(sd), end:fmtDate(ed),
      color:statusColor(status), note:status, dashed:'false', dashedOutline:'false',
      symbol:'', notes:String(r[CI.notes]||'').trim()
    });
  }
  return { success:true, tasks:tasks, meta:{
    name:'LA COSTA HOTEL', subtitle:'PROJECT SCHEDULE', updated:fmtDate(new Date()),
    note:'Live from Project Task List · '+tasks.length+' scheduled tasks'
  }};
}

// ============================================================
//  SAVE BACK
// ============================================================
function saveBackToTaskList(payload) {
  if (!payload.tasks||!payload.tasks.length) return 'Nothing to save.';
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('"'+SOURCE_SHEET+'" tab not found.');
  var raw=sheet.getDataRange().getValues(), hRow=-1;
  for (var i=0;i<Math.min(raw.length,25);i++) {
    for (var c=0;c<raw[i].length;c++) {
      if (String(raw[i][c]).trim().toUpperCase()==='DISCIPLINE'){hRow=i;break;}
    }
    if(hRow>=0)break;
  }
  if (hRow<0) throw new Error('Header row not found.');
  var cols={};
  raw[hRow].forEach(function(h,idx){cols[String(h).trim().toUpperCase()]=idx;});
  var CI={
    discipline:ci(cols,'DISCIPLINE',1), task:ci(cols,'TASK',3),
    start:ci(cols,'START DATE',7), end:ci(cols,'END DATE',8),
    schedule:ci(cols,'SCHEDULE',10), milestone:ci(cols,'MILESTONE',11),
    status:ci(cols,'STATUS',12), notes:ci(cols,'NOTES',20)
  };
  var lookup={}, discLastRow={}, lastDisc='';
  for (var row=hRow+1;row<raw.length;row++) {
    var disc=String(raw[row][CI.discipline]||'').trim(); if(disc)lastDisc=disc;
    var task=String(raw[row][CI.task]||'').trim();
    if (task&&task!=='-'&&task!=='- -') {
      var key=normKey(lastDisc+'|'+task);
      if(!lookup[key]) lookup[key]=row+1;
      discLastRow[normKey(lastDisc)]=row+1;
    }
  }
  var lock=LockService.getScriptLock(); lock.waitLock(15000);
  var updated=0,appended=0,newByDisc={},orphaned=[];
  try {
    payload.tasks.forEach(function(t) {
      var disc=(t.group||'').trim().toUpperCase();
      var name=(t.name||'').trim();
      var key=normKey(disc+'|'+name);
      var sr=lookup[key];
      if (sr) {
        if(t.start) sheet.getRange(sr,CI.start+1).setValue(t.start);
        if(t.end)   sheet.getRange(sr,CI.end+1).setValue(t.end);
        if(t.note){var sc=t.note.split('·')[0].trim().toUpperCase();
          if(STATUS_COLORS[sc]) sheet.getRange(sr,CI.status+1).setValue(sc);}
        if(typeof t.notes!=='undefined') sheet.getRange(sr,CI.notes+1).setValue(t.notes);
        sheet.getRange(sr,CI.schedule+1).setValue(true);
        if(t.type==='milestone') sheet.getRange(sr,CI.milestone+1).setValue(true);
        updated++;
      } else if(name) {
        var nr=new Array(raw[hRow].length).fill('');
        nr[CI.discipline]=disc; nr[CI.task]=name;
        nr[CI.start]=t.start||''; nr[CI.end]=t.end||''; nr[CI.schedule]=true;
        if(t.type==='milestone') nr[CI.milestone]=true;
        if(t.note){var sc2=t.note.split('·')[0].trim().toUpperCase();
          if(STATUS_COLORS[sc2]) nr[CI.status]=sc2;}
        if(t.notes) nr[CI.notes]=t.notes;
        var dk=normKey(disc);
        if(discLastRow[dk]){if(!newByDisc[dk])newByDisc[dk]=[];newByDisc[dk].push(nr);}
        else orphaned.push(nr);
        appended++;
      }
    });
    var dks=Object.keys(newByDisc);
    dks.sort(function(a,b){return(discLastRow[b]||0)-(discLastRow[a]||0);});
    dks.forEach(function(dk){
      var rows=newByDisc[dk], ins=discLastRow[dk];
      sheet.insertRowsAfter(ins,rows.length);
      sheet.getRange(ins+1,1,rows.length,rows[0].length).setValues(rows);
    });
    if(orphaned.length){
      var lr=sheet.getLastRow();
      sheet.getRange(lr+1,1,orphaned.length,orphaned[0].length).setValues(orphaned);
    }
  } finally { lock.releaseLock(); }
  var msg='Updated '+updated+' task(s)';
  if(appended) msg+=', appended '+appended+' new task(s)';
  return msg+' in "'+SOURCE_SHEET+'".';
}

// ============================================================
//  TASK PARAMS TAB
// ============================================================
function readTaskParams() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(TASK_PARAMS_SHEET);
  if(!sh||sh.getLastRow()<2) return {};
  var data=sh.getDataRange().getValues(), hRow=-1;
  for(var i=0;i<Math.min(data.length,5);i++){
    if(String(data[i][0]).trim().toUpperCase()==='KEY'){hRow=i;break;}
  }
  if(hRow<0) return {};
  var cols={};
  data[hRow].forEach(function(h,idx){cols[String(h).trim().toUpperCase()]=idx;});
  var cK=cols['KEY']!==undefined?cols['KEY']:0, cC=cols['COLOR']!==undefined?cols['COLOR']:1,
      cT=cols['TYPE']!==undefined?cols['TYPE']:2, cS=cols['STYLE']!==undefined?cols['STYLE']:3,
      cSy=cols['SYMBOL']!==undefined?cols['SYMBOL']:4, cD=cols['DEPS']!==undefined?cols['DEPS']:5;
  var params={};
  for(var r=hRow+1;r<data.length;r++){
    var key=normKey(String(data[r][cK]||'').trim()); if(!key) continue;
    params[key]={color:String(data[r][cC]||'').trim(), type:String(data[r][cT]||'').trim().toLowerCase(),
      style:String(data[r][cS]||'').trim().toLowerCase(), symbol:String(data[r][cSy]||'').trim(),
      deps:String(data[r][cD]||'').trim()};
  }
  return params;
}

function writeTaskParams(tasks) {
  if(!tasks||!tasks.length) return;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(TASK_PARAMS_SHEET);
  if(!sh) sh=ss.insertSheet(TASK_PARAMS_SHEET);
  sh.clearContents();
  var rows=[['KEY','COLOR','TYPE','STYLE','SYMBOL','DEPS']];
  tasks.forEach(function(t){
    var name=String(t.name||'').trim(); if(!name) return;
    var key=normKey(String(t.group||'').trim().toUpperCase()+'|'+name);
    var style=String(t.dashed)==='true'?'dashed':String(t.dashedOutline)==='true'?'dashed-outline':'solid';
    rows.push([key,String(t.colorOverride||'').trim(),String(t.type||'bar').trim().toLowerCase(),
      style,String(t.symbol||'').trim(),String(t.dependencies||'').trim()]);
  });
  sh.getRange(1,1,rows.length,6).setValues(rows);
  sh.hideSheet();
}

// ============================================================
//  SETTINGS TAB
// ============================================================
var SETTINGS_KEYS=[
  'projectName','projectSubtitle','projectDate','projectNote',
  'userLabelWidth','ganttBarFontSize','showTodayLine','darkMode',
  'flatMode','barTextColor','collapsedGroups','disciplineOrder','taskListOrder',
  'showName','showPhase','showNote','groupTint','taskRowTint',
  'groupHeaderColor','groupHeaderTint','groupHeaderTextColor',
  'currentScale','zoomLevel','chartStart','chartEnd',
  'showDependencies','snapValue','toolbarCollapsed'
];
var SETTINGS_DESCRIPTIONS={
  projectName:'Project title shown in the chart header',
  projectSubtitle:'Project subtitle shown in the chart header',
  projectDate:'Updated date shown in the chart header (YYYY-MM-DD)',
  projectNote:'Top note shown below the chart header',
  userLabelWidth:'Width of the task-name label column (px)',
  ganttBarFontSize:'Font size for Gantt bar labels (pt)',
  showTodayLine:'Show vertical Today line on chart (true/false)',
  darkMode:'Dark theme enabled (true/false)',
  flatMode:'Flat mode — suppress group header rows (true/false)',
  barTextColor:'Bar label text colour hex — empty string = auto-contrast',
  collapsedGroups:'JSON object: which discipline groups are collapsed',
  disciplineOrder:'JSON array: discipline group display order on the Gantt chart',
  taskListOrder:'JSON array: discipline group display order in the Task List section',
  showName:'Show task name text inside Gantt bar (true/false)',
  showPhase:'Show sub/phase label inside Gantt bar (true/false)',
  showNote:'Show status/note text inside Gantt bar (true/false)',
  groupTint:'Opacity of rollup-bar colour tint (0–100)',
  taskRowTint:'Opacity of alternating task-row stripe in label column (0–100)',
  groupHeaderColor:'Background fill colour for all group header rows (hex)',
  groupHeaderTint:'Opacity of group header background fill (0–100)',
  groupHeaderTextColor:'Text colour for group header discipline labels (hex)',
  currentScale:'Timeline scale — weeks / months / quarters / years',
  zoomLevel:'Horizontal zoom multiplier (integer, default 1)',
  chartStart:'Chart date range start override (YYYY-MM-DD, empty = auto)',
  chartEnd:'Chart date range end override (YYYY-MM-DD, empty = auto)',
  showDependencies:'Show Finish-to-Start dependency arrows (true/false)',
  snapValue:'Drag-snap interval in days (1 / 7 / 14 / 30)',
  toolbarCollapsed:'Settings toolbar collapsed state (true/false)'
};

function readSettings() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SETTINGS_SHEET);
  if(!sh||sh.getLastRow()<2) return null;
  var data=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  var result={}, gcMap={};
  for(var r=0;r<data.length;r++){
    var k=String(data[r][0]).trim();
    var raw=data[r][1];
    var v=(raw instanceof Date&&!isNaN(raw.getTime()))?fmtDate(raw):String(raw).trim();
    if(!k||k.toUpperCase()==='SETTING KEY') continue;
    if(k.indexOf('groupColor.')===0) gcMap[k.substring('groupColor.'.length)]=v;
    else result[k]=v;
  }
  if(Object.keys(gcMap).length>0) result.groupColors=JSON.stringify(gcMap);
  return Object.keys(result).length>0?result:null;
}

function writeSettings(settings) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(SETTINGS_SHEET);
  if(!sh) sh=ss.insertSheet(SETTINGS_SHEET);
  sh.clearContents();
  var rows=[['SETTING KEY','VALUE','DESCRIPTION']];
  if(settings.appScriptURL) rows.push(['appScriptURL',settings.appScriptURL,'Apps Script Web App URL (auto-populated on every Save)']);
  SETTINGS_KEYS.forEach(function(k){ rows.push([k,settings[k]!==undefined?String(settings[k]):'',SETTINGS_DESCRIPTIONS[k]||'']); });
  if(settings.groupColors){
    try{
      var gc=JSON.parse(settings.groupColors);
      Object.keys(gc).sort().forEach(function(d){ rows.push(['groupColor.'+d,gc[d],'Rollup-bar / separator accent colour for discipline: '+d]); });
    }catch(e){}
  }
  sh.getRange(1,1,rows.length,3).setValues(rows);
  sh.hideSheet();
}

// ============================================================
//  HELPERS
// ============================================================
function ci(cols,name,fallback){return cols[name.toUpperCase()]!==undefined?cols[name.toUpperCase()]:fallback;}
function parseSheetDate(val){
  if(!val) return null;
  if(val instanceof Date) return isNaN(val.getTime())?null:val;
  var s=String(val).trim();
  if(!s||s==='FALSE'||s==='TRUE'||s==='-') return null;
  s=s.replace(/\\//g,'-');
  var p=s.split('-');
  if(p.length===3){
    var iso=p[0].padStart(4,'0')+'-'+p[1].padStart(2,'0')+'-'+p[2].padStart(2,'0');
    var d=new Date(iso+'T12:00:00'); return isNaN(d.getTime())?null:d;
  }
  var d2=new Date(s); return isNaN(d2.getTime())?null:d2;
}
function fmtDate(d){
  if(!(d instanceof Date)) return '';
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function statusColor(s){if(!s) return '#64748b'; return STATUS_COLORS[s.trim().toUpperCase()]||'#64748b';}
function normKey(s){return String(s).trim().toUpperCase().replace(/\\s+/g,' ');}
`;

// ─────────────────────────────────────────────────────────────────────────────
//  WEB APP MANIFEST pushed to each target script
//  executeAs: USER_DEPLOYING = runs as the owner (can read/write the sheet)
//  access: ANYONE_ANONYMOUS = accessible without login (required for Gantt HTML)
// ─────────────────────────────────────────────────────────────────────────────
var GANTT_SYNC_MANIFEST = JSON.stringify({
  timeZone: 'America/New_York',
  dependencies: {},
  exceptionLogging: 'STACKDRIVER',
  runtimeVersion: 'V8',
  webapp: {
    executeAs: 'USER_DEPLOYING',
    access: 'ANYONE_ANONYMOUS'
  }
}, null, 2);

// ─────────────────────────────────────────────────────────────────────────────
//  ADD-ON HOOKS
// ─────────────────────────────────────────────────────────────────────────────

/** Called once when the add-on is installed. */
function onInstall(e) { onOpen(e); }

/**
 * Called every time a spreadsheet opens.
 * Creates the "Extensions → 📊 Gantt Timeline" menu.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('🚀 Setup Gantt on This Sheet', 'setupGanttTimeline')
    .addSeparator()
    .addItem('ℹ  About', 'showAbout')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────────────────
//  MAIN INSTALLER
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Full automated setup:
 *   create project → push code → create version → deploy → show URL
 */
function setupGanttTimeline() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var name = ss.getName();

  var confirm = ui.alert(
    '🚀 Setup Gantt Timeline',
    'This will automatically:\n\n' +
    '  1. Create a new Apps Script in "' + name + '"\n' +
    '  2. Push the Gantt sync code (V1.13)\n' +
    '  3. Deploy it as a Web App\n' +
    '  4. Show you the URL to paste into the Gantt HTML ⚙ Setup\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Show progress toast
  ss.toast('Creating Apps Script project…', '📊 Gantt Setup', -1);

  try {
    var token = ScriptApp.getOAuthToken();
    var api   = 'https://script.googleapis.com/v1';

    // ── Step 1: Create container-bound script project ──────────────────
    var createRes = _apiCall('POST', api + '/projects', {
      title:    'Gantt Timeline Sync',
      parentId: ssId
    }, token);

    if (!createRes.scriptId) {
      // Common failure: Apps Script API not enabled
      var detail = createRes.error ? ('\n\nError: ' + createRes.error.message) : '';
      throw new Error(
        'Could not create the Apps Script project.' + detail + '\n\n' +
        'ACTION REQUIRED:\n' +
        '  1. Go to script.google.com/home/usersettings\n' +
        '  2. Toggle "Google Apps Script API" ON\n' +
        '  3. Try again'
      );
    }
    var scriptId = createRes.scriptId;

    // ── Step 2: Push code + manifest ───────────────────────────────────
    ss.toast('Pushing Gantt sync code…', '📊 Gantt Setup', -1);
    var pushRes = _apiCall('PUT', api + '/projects/' + scriptId + '/content', {
      files: [
        { name: 'Code',       type: 'SERVER_JS', source: GANTT_SYNC_CODE     },
        { name: 'appsscript', type: 'JSON',       source: GANTT_SYNC_MANIFEST }
      ]
    }, token);
    if (pushRes.error) {
      throw new Error('Failed to push code: ' + JSON.stringify(pushRes.error));
    }

    // ── Step 3: Create a versioned snapshot ────────────────────────────
    ss.toast('Creating version snapshot…', '📊 Gantt Setup', -1);
    var versionRes = _apiCall('POST', api + '/projects/' + scriptId + '/versions', {
      description: 'Gantt Timeline — auto-installed v1'
    }, token);
    if (!versionRes.versionNumber) {
      throw new Error('Failed to create version: ' + JSON.stringify(versionRes));
    }

    // ── Step 4: Deploy as Web App ───────────────────────────────────────
    ss.toast('Deploying as Web App…', '📊 Gantt Setup', -1);
    var deployRes = _apiCall('POST', api + '/projects/' + scriptId + '/deployments', {
      versionNumber:    versionRes.versionNumber,
      manifestFileName: 'appsscript',
      description:      'Gantt Timeline Web App'
    }, token);

    // Extract URL from entryPoints array
    var webUrl = null;
    if (deployRes.entryPoints) {
      for (var i = 0; i < deployRes.entryPoints.length; i++) {
        var ep = deployRes.entryPoints[i];
        if (ep.webApp && ep.webApp.url) { webUrl = ep.webApp.url; break; }
      }
    }
    if (!webUrl) {
      throw new Error(
        'Script deployed but URL could not be read.\n\n' +
        'Script ID: ' + scriptId + '\n' +
        'Open Extensions → Apps Script to get the URL manually.'
      );
    }

    // ── Step 5: Done — show copyable URL ───────────────────────────────
    ss.toast('✅ Done!', '📊 Gantt Setup', 4);
    _showUrlDialog(webUrl, name);

  } catch (err) {
    ss.toast('Setup failed.', '📊 Gantt Setup', 4);
    ui.alert('⚠️ Setup Failed', err.message || err.toString(), ui.ButtonSet.OK);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/** Thin wrapper around UrlFetchApp for JSON REST calls. */
function _apiCall(method, url, payload, token) {
  var options = {
    method:             method,
    contentType:        'application/json',
    headers:            { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  return JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
}

/** HTML dialog showing the web app URL with a copy button. */
function _showUrlDialog(url, sheetName) {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:18px;color:#0f172a;font-size:13px;line-height:1.5}' +
    'h3{margin:0 0 6px;color:#16a34a;font-size:15px}' +
    'p{margin:4px 0 10px;color:#475569;font-size:12px}' +
    '.sheet{font-weight:700;color:#0f172a}' +
    '#url{width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:6px;' +
         'font-size:11px;background:#f8fafc;box-sizing:border-box;resize:none;height:56px}' +
    'button{margin-top:8px;padding:8px 0;background:#0ea5e9;color:#fff;border:none;' +
           'border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;width:100%}' +
    'button:hover{background:#0284c7}' +
    '#msg{font-size:11px;color:#16a34a;margin-top:5px;text-align:center;min-height:15px}' +
    '</style>' +
    '<h3>✅ Gantt Timeline Ready!</h3>' +
    '<p>Script created &amp; deployed for <span class="sheet">' + sheetName + '</span>.<br>' +
    'Copy this URL → paste into the Gantt HTML tool\'s ⚙ Setup panel.</p>' +
    '<textarea id="url" readonly>' + url + '</textarea>' +
    '<button onclick="copy()">📋  Copy URL to Clipboard</button>' +
    '<div id="msg"></div>' +
    '<script>function copy(){' +
    'var e=document.getElementById("url");e.select();e.setSelectionRange(0,9999);' +
    'document.execCommand("copy");' +
    'document.getElementById("msg").textContent="✓ URL copied!";' +
    '}<\/script>'
  ).setWidth(450).setHeight(235);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Gantt Timeline — Setup Complete');
}

/** About dialog. */
function showAbout() {
  SpreadsheetApp.getUi().alert(
    '📊 Gantt Timeline Installer  v1.0',
    'Automatically creates and deploys the Gantt sync script\n' +
    'in any Google Spreadsheet — no manual Apps Script work needed.\n\n' +
    'REQUIREMENTS\n' +
    '  • Google Apps Script API must be enabled:\n' +
    '    script.google.com/home/usersettings\n\n' +
    'After setup, the target sheet also gets its own\n' +
    '"📊 Gantt Timeline → Get Web App URL" menu.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
