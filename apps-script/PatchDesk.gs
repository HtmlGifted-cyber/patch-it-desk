/**
 * ==============================================================================
 * SCRIPT NAME: PatchDesk.gs
 * FILE: bound to your "IT Desk" Google Sheet
 * PROJECT: Patch IT Desk (public edition)
 *
 * PURPOSE:
 * Intake desk for technical issues that a team member or AI assistant could
 * NOT fix inside its own area and that need outside / hands-on work.
 * - Departments file a ticket by adding a row on the "Tickets" tab.
 * - Your watcher / dashboard reads open tickets through the web app (doGet)
 *   using a secret key, and your IT worker (human or AI) writes diagnosis and
 *   status back through the web app (doPost).
 *
 * SETUP (once): run setupPatchDesk, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). The key is in Script Properties
 * (PATCH_KEY) and is printed in the execution log by setupPatchDesk.
 * ==============================================================================
 */

const PATCH_DESK = {
  TICKETS_TAB: 'Tickets',
  RULES_TAB: 'How to File',
  LOG_TAB: 'Desk Log',
  TIMEZONE: 'America/New_York',
  HEADERS: [
    'Ticket ID', 'Filed', 'Filed By', 'Department', 'System / Script',
    'Problem', 'Exact Error', 'Already Tried', 'Urgency', 'Needs Owner Hands-On',
    'Status', 'IT Diagnosis', 'Fix / Next Step', 'Updated', 'Closed',
    'Occurrences', 'Last Seen'
  ],
  BACKUP_FOLDER: 'IT Backups',
  MONITOR_TAB: 'Monitor',
  MONITOR_THRESHOLD: 3,
  MONITOR_LOOKBACK_HOURS: 24,
  STATUSES: ['New', 'Diagnosing', 'Assigned to Worker', 'Waiting on Owner', 'Fixed', 'Closed', 'Test'],
  URGENCY: ['Low', 'Normal', 'High', 'Down'],
  FILERS: ['Gemini', 'ChatGPT', 'Claude', 'Owner', 'IT', 'Monitor'],
  OPEN_STATUSES: ['New', 'Diagnosing', 'Assigned to Worker', 'Waiting on Owner', 'Test'],
  BOARD_ID: '', // blank = board features off, everything else still works
  BOARD_FROM: 'IT Desk',
  SWEEP_SENDERS: [],
  HISTORY_TAB: 'Closed History',
  HISTORY_HEADERS: [
    'Ticket ID', 'First Closed', 'Last Closed', 'Filed By', 'Department', 'System / Script',
    'Problem', 'IT Diagnosis', 'Fix / Next Step', 'Times Closed', 'Times Reopened', 'Last Reopened', 'Repeats Filed'
  ],
  WORKER_NAME: 'AI worker'
};

function setupPatchDesk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(PATCH_DESK.TIMEZONE);

  let sheet = ss.getSheetByName(PATCH_DESK.TICKETS_TAB);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(PATCH_DESK.TICKETS_TAB);
  }
  const h = PATCH_DESK.HEADERS;
  sheet.getRange(1, 1, 1, h.length).setValues([h])
       .setFontWeight('bold').setBackground('#3E5F3A').setFontColor('#FAF6EC');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, h.length, 140);
  sheet.setColumnWidths(6, 3, 300);
  sheet.setColumnWidths(12, 2, 300);
  sheet.getRange(2, 6, 999, 8).setWrap(true);

  const list = (values) => SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build();
  sheet.getRange(2, col_('Filed By'), 999, 1).setDataValidation(list(PATCH_DESK.FILERS));
  sheet.getRange(2, col_('Urgency'), 999, 1).setDataValidation(list(PATCH_DESK.URGENCY));
  sheet.getRange(2, col_('Needs Owner Hands-On'), 999, 1).setDataValidation(list(['Yes', 'No', 'Unsure']));
  sheet.getRange(2, col_('Status'), 999, 1).setDataValidation(list(PATCH_DESK.STATUSES));

  let rules = ss.getSheetByName(PATCH_DESK.RULES_TAB) || ss.insertSheet(PATCH_DESK.RULES_TAB);
  rules.clear();
  const text = [
    ['PATCH IT DESK - HOW TO FILE A TECHNICAL ISSUE'],
    [''],
    ['1. Try to fix it inside your own area first.'],
    ['2. File here ONLY if you could not fix it yourself, it needs hands-on/outside work, or it keeps failing (3+ times).'],
    ['3. Add ONE row on the Tickets tab. Fill: Filed By, Department, System / Script, Problem, Exact Error (copy it word for word), Already Tried, Urgency, Needs Owner Hands-On. Leave Ticket ID blank and set Status = New.'],
    ['4. Urgency: Down = something customer-facing or a daily pipeline is stopped; High = failing daily; Normal = needs a fix soon; Low = cleanup.'],
    ['5. Do not edit IT Diagnosis / Fix / Closed columns. The IT worker updates those and the Status.'],
    ['6. One problem per row. If it happens again, add a note in Already Tried on the same row instead of a new row.'],
    ['7. Dropdown columns accept ONLY these exact values. Filed By: ' + PATCH_DESK.FILERS.join(', ') + '. Urgency: ' + PATCH_DESK.URGENCY.join(', ') + '. Needs Owner Hands-On: Yes, No, Unsure. Status when filing: New (or Test for a test entry).'],
    ['8. Ticket ID and Filed are filled in automatically within about 10 minutes. The watcher sends an alert for every new ticket.']
  ];
  rules.getRange(1, 1, text.length, 1).setValues(text);
  rules.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  rules.setColumnWidth(1, 900);
  rules.getRange(1, 1, text.length, 1).setWrap(true);

  if (!ss.getSheetByName(PATCH_DESK.LOG_TAB)) {
    ss.insertSheet(PATCH_DESK.LOG_TAB).appendRow(['Time', 'Action', 'Ticket ID', 'Detail']);
  }

  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('PATCH_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('PATCH_KEY', key);
  }
  if (!ss.getSheetByName(PATCH_DESK.MONITOR_TAB)) {
    ss.insertSheet(PATCH_DESK.MONITOR_TAB).appendRow(['Key', 'Project', 'Function', 'Failures (lifetime)', 'First Seen', 'Last Seen', 'Ticket ID']);
  }
  backupFolder_();
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'patchMonitor')) {
    ScriptApp.newTrigger('patchMonitor').timeBased().everyHours(1).create();
  }
  assignTicketIds_();
  Logger.log('Patch IT Desk ready. Spreadsheet ID: ' + ss.getId());
  Logger.log('PATCH_KEY (give to your watcher, keep private): ' + key);
}

/** GET ?key=...&view=open|all  -> JSON list of tickets */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'sweep') return json_(logSweep_(p));
  if (!keyOk_(p.key)) return json_({ ok: false, error: 'bad key' });
  if (p.view === 'replies') return json_(boardReplies_());
  assignTicketIds_();
  const rows = readTickets_();
  const view = p.view || 'open';
  const tickets = view === 'all' ? rows : rows.filter(t => PATCH_DESK.OPEN_STATUSES.indexOf(t['Status']) !== -1);
  const counts = {};
  rows.forEach(t => { counts[t['Status'] || 'New'] = (counts[t['Status'] || 'New'] || 0) + 1; });
  return json_({
    ok: true,
    desk: 'IT Desk',
    checked: Utilities.formatDate(new Date(), PATCH_DESK.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss"),
    open: rows.filter(t => PATCH_DESK.OPEN_STATUSES.indexOf(t['Status']) !== -1).length,
    counts: counts,
    tickets: tickets
  });
}

/**
 * POST JSON {key, action:'update', id, status, diagnosis, fix}
 *        or {key, action:'file', filedBy, department, system, problem, error, tried, urgency, handsOn}
 *        or {key, action:'backup', id, name, content}
 *        or {key, action:'notice', title, area, details, needsConfirm}
 *        or {key, action:'sweep', sweepId, status, rowsChecked, changes, anomalies, actions, details}
 */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (body.action === 'sweep') return json_(logSweep_(body));
  if (!keyOk_(body.key)) return json_({ ok: false, error: 'bad key' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy' });
  try {
    if (body.action === 'file') return json_(fileTicket_(body));
    if (body.action === 'update') return json_(updateTicket_(body));
    if (body.action === 'backup') return json_(saveBackup_(body));
    if (body.action === 'notice') return json_(postWorkNotice_(body));
    return json_({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

function fileTicket_(b) {
  const sheet = sheet_();
  const now = new Date();
  const row = PATCH_DESK.HEADERS.map(() => '');
  row[col_('Filed') - 1] = now;
  row[col_('Filed By') - 1] = b.filedBy || 'Monitor';
  row[col_('Department') - 1] = b.department || '';
  row[col_('System / Script') - 1] = b.system || '';
  row[col_('Problem') - 1] = b.problem || '';
  row[col_('Exact Error') - 1] = b.error || '';
  row[col_('Already Tried') - 1] = b.tried || '';
  row[col_('Urgency') - 1] = b.urgency || 'Normal';
  row[col_('Needs Owner Hands-On') - 1] = b.handsOn || 'Unsure';
  row[col_('Status') - 1] = 'New';
  row[col_('Updated') - 1] = now;
  sheet.appendRow(row);
  const ids = assignTicketIds_();
  const id = ids[ids.length - 1];
  log_('file', id, (b.filedBy || 'Monitor') + ': ' + (b.problem || '').slice(0, 200));
  return { ok: true, id: id };
}

function updateTicket_(b) {
  const sheet = sheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(b.id)) {
      const r = i + 1;
      const oldStatus = String(data[i][col_('Status') - 1] || '');
      if (b.status) sheet.getRange(r, col_('Status')).setValue(b.status);
      if (b.diagnosis) sheet.getRange(r, col_('IT Diagnosis')).setValue(b.diagnosis);
      if (b.fix) sheet.getRange(r, col_('Fix / Next Step')).setValue(b.fix);
      sheet.getRange(r, col_('Updated')).setValue(new Date());
      if (b.status === 'Closed' || b.status === 'Fixed') sheet.getRange(r, col_('Closed')).setValue(new Date());
      log_('update', b.id, (b.status || '') + ' ' + (b.diagnosis || '').slice(0, 150));
      const t = {};
      PATCH_DESK.HEADERS.forEach((h, c) => { t[h] = data[i][c]; });
      if (b.diagnosis) t['IT Diagnosis'] = b.diagnosis;
      if (b.fix) t['Fix / Next Step'] = b.fix;
      const done = s => s === 'Fixed' || s === 'Closed';
      if (done(b.status) && !done(oldStatus)) historyClosed_(t);
      if (b.status === 'New' && done(oldStatus)) historyReopened_(b.id);
      const notice = b.status && b.status === oldStatus ? 'none (already ' + oldStatus + ')' : boardNotice_(t, b.status);
      return { ok: true, id: b.id, board: notice };
    }
  }
  return { ok: false, error: 'ticket not found' };
}

/** Gives every filled row a ticket ID (PT-0001...) and a Filed date / Status if missing. */
function assignTicketIds_() {
  const sheet = sheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const range = sheet.getRange(2, 1, last - 1, PATCH_DESK.HEADERS.length);
  const data = range.getValues();
  let max = 0;
  data.forEach(r => { const m = String(r[0]).match(/^PT-(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  const ids = [];
  const fresh = [];
  let changed = false;
  data.forEach(r => {
    const filled = r.slice(2, 8).some(v => String(v).trim() !== '');
    if (!filled) return;
    if (!/^PT-\d+$/.test(String(r[0]))) { max++; r[0] = 'PT-' + ('000' + max).slice(-4); changed = true; fresh.push(r[0]); }
    if (!r[col_('Filed') - 1]) { r[col_('Filed') - 1] = new Date(); changed = true; }
    if (!r[col_('Status') - 1]) { r[col_('Status') - 1] = 'New'; changed = true; }
    ids.push(r[0]);
  });
  if (changed) range.setValues(data);
  fresh.forEach(id => { if (id !== undefined) try { checkRepeat_(id); } catch (e) { log_('repeat-error', id, String(e)); } });
  return ids;
}

function readTickets_() {
  const data = sheet_().getDataRange().getValues();
  const h = data[0];
  return data.slice(1)
    .filter(r => /^PT-\d+$/.test(String(r[0])))
    .map(r => {
      const o = {};
      h.forEach((name, i) => {
        o[name] = r[i] instanceof Date
          ? Utilities.formatDate(r[i], PATCH_DESK.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss") : r[i];
      });
      return o;
    });
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PATCH_DESK.TICKETS_TAB);
}

function col_(name) {
  return PATCH_DESK.HEADERS.indexOf(name) + 1;
}

function keyOk_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('PATCH_KEY');
  return !!real && key === real;
}

function log_(action, id, detail) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PATCH_DESK.LOG_TAB);
  if (s) s.appendRow([new Date(), action, id, detail]);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------------------------
// Backups: saves a copy of anything before changes are made.
// ---------------------------------------------------------------------------------------------
function backupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BACKUP_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ }
  }
  const parent = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  const it = parent.getFoldersByName(PATCH_DESK.BACKUP_FOLDER);
  const folder = it.hasNext() ? it.next() : parent.createFolder(PATCH_DESK.BACKUP_FOLDER);
  props.setProperty('BACKUP_FOLDER_ID', folder.getId());
  return folder;
}

function saveBackup_(b) {
  if (!b.id || !b.name || typeof b.content !== 'string') return { ok: false, error: 'need id, name, content' };
  const root = backupFolder_();
  const sub = root.getFoldersByName(b.id);
  const folder = sub.hasNext() ? sub.next() : root.createFolder(b.id);
  const stamp = Utilities.formatDate(new Date(), PATCH_DESK.TIMEZONE, 'yyyy-MM-dd HHmm');
  const file = folder.createFile(stamp + ' ' + String(b.name).replace(/[\/\\]/g, '_') + '.txt', b.content, MimeType.PLAIN_TEXT);
  log_('backup', b.id, b.name + ' -> ' + file.getUrl());
  return { ok: true, id: b.id, url: file.getUrl() };
}

// ---------------------------------------------------------------------------------------------
// Patch Monitor: hourly, reads failed Apps Script runs from failure emails.
// ---------------------------------------------------------------------------------------------
function patchMonitor() {
  const sinceMs = Date.now() - PATCH_DESK.MONITOR_LOOKBACK_HOURS * 3600 * 1000;
  const lastRun = PropertiesService.getScriptProperties().getProperty('MONITOR_LAST') || new Date(sinceMs).toISOString();
  try { ingestSweepEmails_(); } catch (err) { log_('sweep-error', '', String(err).slice(0, 200)); }
  const failed = failuresFromEmail_(sinceMs, lastRun);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mon = ss.getSheetByName(PATCH_DESK.MONITOR_TAB);
  if (!mon) return;
  const rows = mon.getDataRange().getValues();
  const index = {};
  for (let i = 1; i < rows.length; i++) index[rows[i][0]] = i + 1;
  const tickets = readTickets_();

  Object.keys(failed).forEach(k => {
    const f = failed[k];
    let r = index[k];
    if (!r) {
      mon.appendRow([k, f.project, f.fn, 0, f.first, f.last, '']);
      r = mon.getLastRow();
    }
    const cur = mon.getRange(r, 1, 1, 7).getValues()[0];
    const total = Number(cur[3] || 0) + f.fresh;
    mon.getRange(r, 4, 1, 3).setValues([[total, cur[4] || f.first, f.last]]);
    let tid = cur[6];
    const open = tid && tickets.some(t => t['Ticket ID'] === tid && ['Fixed', 'Closed'].indexOf(t['Status']) === -1);
    if (open) {
      setTicketCounts_(tid, total, f.last);
    } else if (f.n >= PATCH_DESK.MONITOR_THRESHOLD) {
      const res = fileTicket_({
        filedBy: 'Monitor', department: 'IT', system: f.project + ' -> ' + f.fn,
        problem: f.fn + ' failed ' + f.n + ' times in the last ' + PATCH_DESK.MONITOR_LOOKBACK_HOURS + ' h (' + f.status + ').',
        error: f.error ? f.error + ' (from Google failure email; see Apps Script executions for ' + f.project + ')' : 'See Apps Script executions for ' + f.project,
        tried: 'Auto-filed by Patch Monitor',
        urgency: f.n >= 12 ? 'High' : 'Normal', handsOn: 'Unsure'
      });
      if (res.ok) {
        mon.getRange(r, 7).setValue(res.id);
        setTicketCounts_(res.id, total, f.last);
      }
    }
  });
  PropertiesService.getScriptProperties().setProperty('MONITOR_LAST', new Date().toISOString());
}

function setTicketCounts_(id, count, lastSeen) {
  const sheet = sheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, col_('Occurrences'), 1, 2).setValues([[count, new Date(lastSeen)]]);
      return;
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Optional NOTICE BOARD (replaces private Back Office Ledger if BOARD_ID is configured)
// ---------------------------------------------------------------------------------------------
function getBoard_() {
  const id = PATCH_DESK.BOARD_ID || PropertiesService.getScriptProperties().getProperty('BOARD_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let notices = ss.getSheetByName('Notices');
  if (!notices) {
    notices = ss.insertSheet('Notices');
    notices.appendRow(['Date', 'From', 'To', 'Type', 'Ticket', 'Subject', 'Message', 'Confirm By']);
    notices.setFrozenRows(1);
  }
  let runLog = ss.getSheetByName('Run Log');
  if (!runLog) {
    runLog = ss.insertSheet('Run Log');
    runLog.appendRow(['Date', 'Worker', 'Check Type', 'Status', 'Counts', 'Actions', 'Notes']);
    runLog.setFrozenRows(1);
  }
  return ss;
}

function boardNotice_(t, status) {
  if (['Fixed', 'Closed', 'Assigned to Worker'].indexOf(status) === -1) return 'none';
  const board = getBoard_();
  if (!board) return 'board off (BOARD_ID blank)';
  try {
    const now = new Date();
    const id = t['Ticket ID'];
    const what = (t['System / Script'] || 'technical issue') + ': ' + String(t['Problem'] || '').slice(0, 300);
    const notices = board.getSheetByName('Notices');
    if (status === 'Assigned to Worker') {
      notices.appendRow([
        now, PATCH_DESK.BOARD_FROM, PATCH_DESK.WORKER_NAME, 'WORK ORDER', id,
        'IT work order ' + id + ' — ' + (t['System / Script'] || ''),
        'Ticket diagnosed. Needs worker action.\n\nISSUE: ' + what +
          '\n\nDIAGNOSIS: ' + (t['IT Diagnosis'] || '') + '\n\nSTEPS: ' + (t['Fix / Next Step'] || '') +
          '\n\nWhen done, reply "' + id + ' done".',
        'Owner'
      ]);
      return 'Worker work order added to board Notices';
    }
    notices.appendRow([
      now, PATCH_DESK.BOARD_FROM, t['Filed By'] || 'Filer', 'RESOLVED', id,
      id + ' ' + status.toUpperCase() + ' — ' + what,
      'Diagnosis: ' + String(t['IT Diagnosis'] || '').slice(0, 700) +
        '\nFix: ' + String(t['Fix / Next Step'] || '').slice(0, 700),
      t['Filed By'] || 'Filer'
    ]);
    return 'Resolution notice added to board Notices';
  } catch (err) {
    log_('board-error', t['Ticket ID'], String(err).slice(0, 200));
    return 'board notice failed: ' + err;
  }
}

function refreshStatusDropdown() {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(PATCH_DESK.STATUSES, true).build();
  sheet_().getRange(2, col_('Status'), 999, 1).setDataValidation(rule);
}

function postWorkNotice_(b) {
  const title = String(b.title || '').trim().slice(0, 200);
  if (!title) return { ok: false, error: 'title required' };
  const board = getBoard_();
  if (!board) return { ok: false, error: 'board off (BOARD_ID blank)' };
  try {
    const now = new Date();
    const confirm = String(b.needsConfirm || '').trim().slice(0, 100);
    board.getSheetByName('Notices').appendRow([
      now, PATCH_DESK.BOARD_FROM, confirm || 'All', 'WORK COMPLETED', '',
      title + (b.area ? ' (' + String(b.area).slice(0, 100) + ')' : ''),
      String(b.details || '').slice(0, 1500),
      confirm
    ]);
    log_('notice', '', title);
    return { ok: true, message: 'Completed-work notice added to board Notices' };
  } catch (err) {
    log_('board-error', '', String(err).slice(0, 200));
    return { ok: false, error: 'board notice failed: ' + err };
  }
}

function authorizeBoard() {
  const id = PATCH_DESK.BOARD_ID || PropertiesService.getScriptProperties().getProperty('BOARD_ID');
  if (!id) {
    Logger.log('BOARD_ID is blank in script properties. Set BOARD_ID first.');
    return;
  }
  const name = SpreadsheetApp.openById(id).getName();
  Logger.log('Board access OK: ' + name);
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone(PATCH_DESK.TIMEZONE);
  Logger.log('IT Desk time zone: ' + PATCH_DESK.TIMEZONE);
}

// ---------------------------------------------------------------------------------------------
// Closed History & Repeat Detection
// ---------------------------------------------------------------------------------------------
function historySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(PATCH_DESK.HISTORY_TAB);
  if (!h) {
    h = ss.insertSheet(PATCH_DESK.HISTORY_TAB);
    h.getRange(1, 1, 1, PATCH_DESK.HISTORY_HEADERS.length).setValues([PATCH_DESK.HISTORY_HEADERS]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function historyCol_(name) { return PATCH_DESK.HISTORY_HEADERS.indexOf(name) + 1; }

function historyRow_(h, id) {
  const ids = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getValues() : [];
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function historyClosed_(t) {
  const h = historySheet_();
  const now = new Date();
  const r = historyRow_(h, t['Ticket ID']);
  if (r) {
    const c = historyCol_;
    h.getRange(r, c('Last Closed')).setValue(now);
    h.getRange(r, c('IT Diagnosis')).setValue(t['IT Diagnosis'] || '');
    h.getRange(r, c('Fix / Next Step')).setValue(t['Fix / Next Step'] || '');
    h.getRange(r, c('Times Closed')).setValue((+h.getRange(r, c('Times Closed')).getValue() || 0) + 1);
  } else {
    h.appendRow([t['Ticket ID'], now, now, t['Filed By'] || '', t['Department'] || '', t['System / Script'] || '',
      String(t['Problem'] || '').slice(0, 1000), t['IT Diagnosis'] || '', t['Fix / Next Step'] || '', 1, 0, '', '']);
  }
}

function historyReopened_(id) {
  const h = historySheet_();
  const r = historyRow_(h, id);
  if (!r) return;
  h.getRange(r, historyCol_('Times Reopened')).setValue((+h.getRange(r, historyCol_('Times Reopened')).getValue() || 0) + 1);
  h.getRange(r, historyCol_('Last Reopened')).setValue(new Date());
}

function words_(text) {
  const stop = ['the', 'and', 'for', 'with', 'that', 'this', 'from', 'are', 'not', 'any', 'all', 'was', 'into', 'its', 'has', 'have'];
  const out = {};
  String(text || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/)
    .filter(w => w.length > 2 && stop.indexOf(w) === -1).forEach(w => { out[w] = true; });
  return out;
}

function checkRepeat_(id) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PATCH_DESK.HISTORY_TAB);
  if (!h || h.getLastRow() < 2) return '';
  const sheet = sheet_();
  const data = sheet.getDataRange().getValues();
  let r = 0, t = null;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]) === String(id)) { r = i + 1; t = data[i]; break; }
  if (!t) return '';
  const sys = String(t[col_('System / Script') - 1] || '').trim().toLowerCase();
  const who = [String(t[col_('Department') - 1] || '').toLowerCase(), String(t[col_('Filed By') - 1] || '').toLowerCase()];
  const mine = words_(t[col_('Problem') - 1]);
  const hist = h.getRange(2, 1, h.getLastRow() - 1, PATCH_DESK.HISTORY_HEADERS.length).getValues();
  const hc = name => historyCol_(name) - 1;
  let best = null, bestScore = 0;
  hist.forEach((row, k) => {
    if (String(row[0]) === String(id)) return;
    const theirs = words_(row[hc('Problem')]);
    const keys = Object.keys(mine);
    const overlap = keys.length ? keys.filter(w => theirs[w]).length / keys.length : 0;
    const sameSys = sys && sys === String(row[hc('System / Script')] || '').trim().toLowerCase();
    const sameWho = who.indexOf(String(row[hc('Department')] || '').toLowerCase()) !== -1 ||
                    who.indexOf(String(row[hc('Filed By')] || '').toLowerCase()) !== -1;
    const score = (sameSys ? 1 : 0) + (sameWho ? 0.3 : 0) + overlap;
    if ((sameSys || (sameWho && overlap >= 0.4) || overlap >= 0.6) && score > bestScore) { best = { row: row, at: k + 2 }; bestScore = score; }
  });
  if (!best) return '';
  const prev = best.row;
  const note = 'POSSIBLE REPEAT of ' + prev[0] + ' (last closed ' +
    Utilities.formatDate(new Date(prev[hc('Last Closed')]), PATCH_DESK.TIMEZONE, 'yyyy-MM-dd') + '). Earlier diagnosis: ' +
    String(prev[hc('IT Diagnosis')] || '').slice(0, 500) + ' | Earlier fix: ' + String(prev[hc('Fix / Next Step')] || '').slice(0, 500);
  const dCell = sheet.getRange(r, col_('IT Diagnosis'));
  dCell.setValue(dCell.getValue() ? note + '\n\n' + dCell.getValue() : note);
  const rep = h.getRange(best.at, historyCol_('Repeats Filed'));
  rep.setValue(rep.getValue() ? rep.getValue() + ', ' + id : id);
  log_('repeat', id, 'possible repeat of ' + prev[0]);
  return prev[0];
}

function backfillClosedHistory() {
  const data = sheet_().getDataRange().getValues();
  const h = historySheet_();
  data.slice(1).forEach(row => {
    const st = String(row[col_('Status') - 1]);
    if ((st === 'Fixed' || st === 'Closed') && !historyRow_(h, row[0])) {
      const t = {};
      PATCH_DESK.HEADERS.forEach((name, c) => { t[name] = row[c]; });
      historyClosed_(t);
    }
  });
  Logger.log('Closed History rows: ' + (h.getLastRow() - 1));
}

// ---------------------------------------------------------------------------------------------
// Sweep Write-Back & Email Intake
// ---------------------------------------------------------------------------------------------
function setupSweepKey() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('SWEEP_KEY');
  if (!k) { k = 'SWEEP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('SWEEP_KEY', k); }
  Logger.log('Sweep key: ' + k);
}

function logSweep_(p) {
  const real = PropertiesService.getScriptProperties().getProperty('SWEEP_KEY');
  if (!real || p.key !== real) return { ok: false, error: 'bad key' };
  return writeSweepRow_(p);
}

function writeSweepRow_(p) {
  const sweepId = String(p.sweepId || '').trim().slice(0, 60);
  if (!/^[A-Za-z0-9._:-]{4,60}$/.test(sweepId)) return { ok: false, error: 'sweepId required' };
  const n = v => (v === undefined || v === '' ? '' : (isNaN(+v) ? String(v).slice(0, 20) : +v));
  const status = String(p.status || 'COMPLETED').trim().toUpperCase().slice(0, 30);
  const board = getBoard_();
  if (!board) return { ok: false, error: 'board off (BOARD_ID blank)' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'busy' };
  try {
    const runLog = board.getSheetByName('Run Log');
    if (!runLog) return { ok: false, error: 'Run Log tab not found' };
    const tag = '[sweep ' + sweepId + ']';
    const last = runLog.getLastRow();
    if (last > 1) {
      const start = Math.max(2, last - 499);
      const notes = runLog.getRange(start, 7, last - start + 1, 1).getValues();
      if (notes.some(r => String(r[0]).indexOf(tag) !== -1)) return { ok: true, duplicate: true, message: 'sweep already logged' };
    }
    const issues = 'Rows checked: ' + n(p.rowsChecked) + ' | Changes detected: ' + n(p.changes) +
      ' | Anomalies found: ' + n(p.anomalies);
    runLog.appendRow([
      new Date(), PATCH_DESK.WORKER_NAME, 'MONITORING SWEEP', status, issues,
      'Actions taken: ' + String(p.actions || 'none').slice(0, 800),
      tag + ' ' + String(p.details || '').slice(0, 1200)
    ]);
    log_('sweep', '', sweepId + ' ' + status + (p.via ? ' via ' + p.via : ''));
    return { ok: true, message: 'sweep logged to board Run Log' };
  } finally {
    lock.releaseLock();
  }
}

function sweepInboxAddress() {
  Logger.log('Sweep email address: ' + Session.getEffectiveUser().getEmail());
}

function ingestSweepEmails_() {
  if (!PATCH_DESK.SWEEP_SENDERS || !PATCH_DESK.SWEEP_SENDERS.length) return;
  const board = getBoard_();
  if (!board) return;
  const from = PATCH_DESK.SWEEP_SENDERS.map(a => 'from:' + a).join(' OR ');
  const label = GmailApp.getUserLabelByName('IT Desk/Sweep logged') || GmailApp.createLabel('IT Desk/Sweep logged');
  const threads = GmailApp.search('subject:"[DESK SWEEP]" (' + from + ') -label:"IT Desk/Sweep logged" newer_than:7d', 0, 50);
  threads.forEach(th => {
    th.getMessages().forEach(msg => {
      const id = (msg.getSubject().match(/\[DESK SWEEP\]\s*([A-Za-z0-9._:-]{4,60})/) || [])[1];
      if (!id) return;
      const body = msg.getPlainBody();
      const get = name => ((body.match(new RegExp('^\\s*' + name + '\\s*:\\s*(.+)$', 'im')) || [])[1] || '').trim();
      const res = writeSweepRow_({
        sweepId: id, status: get('Status') || 'COMPLETED',
        rowsChecked: get('Rows checked'), changes: get('Changes detected'), anomalies: get('Anomalies found'),
        actions: get('Actions taken'), details: get('Details'), via: 'email'
      });
      if (!res.ok) log_('sweep-error', '', id + ': ' + res.error);
    });
    th.addLabel(label);
  });
}

function boardReplies_() {
  const board = getBoard_();
  if (!board) return { ok: true, replies: [] };
  const out = [];
  const scan = (tab, dateCol, fromCol, textCols) => {
    const sh = board.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    const last = sh.getLastRow(), start = Math.max(2, last - 199);
    sh.getRange(start, 1, last - start + 1, sh.getLastColumn()).getValues().forEach(r => {
      const from = String(r[fromCol] || '');
      if (from.indexOf(PATCH_DESK.BOARD_FROM) === 0) return;
      const text = textCols.map(c => String(r[c] || '')).join(' — ');
      const ids = text.match(/PT-\d{4}/g);
      if (!ids) return;
      const d = r[dateCol] instanceof Date ? Utilities.formatDate(r[dateCol], PATCH_DESK.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss") : String(r[dateCol]);
      [...new Set(ids)].forEach(id => out.push({ ticket: id, tab: tab, from: from, date: d, text: text.slice(0, 600),
        done: new RegExp(id + '\\W+(done|fixed|complete)', 'i').test(text) }));
    });
  };
  scan('Notices', 0, 1, [3, 4, 5, 6]);
  scan('Run Log', 0, 1, [3, 4, 5, 6]);
  out.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return { ok: true, replies: out.slice(0, 100) };
}

function failuresFromEmail_(sinceMs, lastRunIso) {
  const days = Math.max(1, Math.ceil(PATCH_DESK.MONITOR_LOOKBACK_HOURS / 24) + 1);
  const threads = GmailApp.search('from:apps-scripts-notifications@google.com newer_than:' + days + 'd', 0, 100);
  const failed = {}, seen = {};
  const lastRun = new Date(lastRunIso).getTime();
  const strip = h => String(h).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\s+/g, ' ').trim();
  threads.forEach(th => th.getMessages().forEach(msg => {
    const subj = msg.getSubject();
    const m = subj.match(/failures for Google Apps Script:?\s*(.+)$/i);
    const project = m ? m[1].trim() : '?';
    const rows = msg.getBody().match(/<tr[\s\S]*?<\/tr>/gi) || [];
    rows.forEach(tr => {
      const cells = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(strip);
      if (cells.length < 4) return;
      const [startTxt, fn, error, trigger] = cells;
      if (!fn || /^function$/i.test(fn)) return;
      let start = new Date(startTxt.replace(',', '')).getTime();
      if (isNaN(start)) start = msg.getDate().getTime();
      if (start < sinceMs) return;
      const once = project + '|' + fn + '|' + startTxt;
      if (seen[once]) return;
      seen[once] = true;
      const k = project + ' | ' + fn;
      const iso = new Date(start).toISOString();
      const f = failed[k] || (failed[k] = { project: project, fn: fn, n: 0, fresh: 0, first: iso, last: iso,
        status: 'failed, trigger: ' + (trigger || '?'), error: '' });
      f.n++;
      if (start > lastRun) f.fresh++;
      if (iso < f.first) f.first = iso;
      if (iso >= f.last) { f.last = iso; f.error = String(error || '').slice(0, 400); }
    });
  }));
  log_('monitor', '', 'email scan: ' + threads.length + ' failure emails, ' + Object.keys(failed).length + ' failing functions');
  return failed;
}
