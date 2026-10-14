/**
 * ==============================================================================
 * SCRIPT NAME: PatchDesk.gs
 * FILE: bound to your "IT Desk" Google Sheet
 * PROJECT: Patch IT Desk (bare-bones public edition)
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
  TIMEZONE: 'America/New_York',          // change to your timezone
  OWNER: 'the owner',                     // who gets hands-on requests, e.g. 'Sam'
  HEADERS: [
    'Ticket ID', 'Filed', 'Filed By', 'Department', 'System / Script',
    'Problem', 'Exact Error', 'Already Tried', 'Urgency', 'Needs Owner Hands-On',
    'Status', 'IT Diagnosis', 'Fix / Next Step', 'Updated', 'Closed'
  ],
  STATUSES: ['New', 'Diagnosing', 'Waiting on Owner', 'Fixed', 'Closed', 'Test'],
  URGENCY: ['Low', 'Normal', 'High', 'Down'],
  // Who may file tickets. Edit to match your team / AI assistants, then re-run setupPatchDesk.
  FILERS: ['Gemini', 'ChatGPT', 'Claude', 'Owner', 'IT', 'Monitor'],
  OPEN_STATUSES: ['New', 'Diagnosing', 'Waiting on Owner', 'Test']
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
  assignTicketIds_();
  Logger.log('Patch IT Desk ready. Spreadsheet ID: ' + ss.getId());
  Logger.log('PATCH_KEY (give to your watcher, keep private): ' + key);
}

/** GET ?key=...&view=open|all  -> JSON list of tickets */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!keyOk_(p.key)) return json_({ ok: false, error: 'bad key' });
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
 */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (!keyOk_(body.key)) return json_({ ok: false, error: 'bad key' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy' });
  try {
    if (body.action === 'file') return json_(fileTicket_(body));
    if (body.action === 'update') return json_(updateTicket_(body));
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
      if (b.status) sheet.getRange(r, col_('Status')).setValue(b.status);
      if (b.diagnosis) sheet.getRange(r, col_('IT Diagnosis')).setValue(b.diagnosis);
      if (b.fix) sheet.getRange(r, col_('Fix / Next Step')).setValue(b.fix);
      sheet.getRange(r, col_('Updated')).setValue(new Date());
      if (b.status === 'Closed' || b.status === 'Fixed') sheet.getRange(r, col_('Closed')).setValue(new Date());
      log_('update', b.id, (b.status || '') + ' ' + (b.diagnosis || '').slice(0, 150));
      return { ok: true, id: b.id };
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
  let changed = false;
  data.forEach(r => {
    const filled = r.slice(2, 8).some(v => String(v).trim() !== '');
    if (!filled) return;
    if (!/^PT-\d+$/.test(String(r[0]))) { max++; r[0] = 'PT-' + ('000' + max).slice(-4); changed = true; }
    if (!r[col_('Filed') - 1]) { r[col_('Filed') - 1] = new Date(); changed = true; }
    if (!r[col_('Status') - 1]) { r[col_('Status') - 1] = 'New'; changed = true; }
    ids.push(r[0]);
  });
  if (changed) range.setValues(data);
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
