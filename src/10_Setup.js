/**
 * 10_Setup — builds / repairs the spreadsheet non-destructively, seeds the demo scenario,
 * installs triggers. Safe to run any number of times (spec F09).
 */

const CONFIG_HEADERS = ['Key', 'Value', 'Description', 'Status'];

function setupSystem_(ss, opts) {
  opts = opts || {};
  try { ss.setSpreadsheetTimeZone(Session.getScriptTimeZone()); } catch (e) { /* not fatal */ }

  ensureStartSheet_(ss);
  ensureConfigSheet_(ss);
  const team = ensureTableSheet_(ss, SHEETS.TEAM, TEAM_FIELDS.map(function (f) { return f[1]; }));
  const tasks = ensureTableSheet_(ss, SHEETS.TASKS, TASK_FIELDS.map(function (f) { return f[1]; }));
  const cycles = ensureTableSheet_(ss, SHEETS.CYCLES, CYCLE_FIELDS.map(function (f) { return f[1]; }));
  ensureTableSheet_(ss, SHEETS.LOG, LOG_HEADERS);
  ensureTableSheet_(ss, SHEETS.EMAILS, EMAIL_HEADERS);
  ensureTableSheet_(ss, SHEETS.TESTS, TEST_HEADERS);

  if (opts.seed !== false && team.getLastRow() < 2 && tasks.getLastRow() < 2) {
    seedDemo_(ss, opts.now || new Date());
  }
  formatTeam_(team);
  formatTasks_(ss, tasks, team);
  formatCycles_(cycles);
  formatLogs_(ss);

  if (!ss.getSheetByName(SHEETS.DASHBOARD)) styleDashboard_(ss.insertSheet(SHEETS.DASHBOARD));
  const stray = ss.getSheetByName('Sheet1');
  if (stray && stray.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(stray);
  orderSheets_(ss);
}

function ensureTableSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ensureCols_(sh, headers.length);
  const lastCol = sh.getLastColumn();
  const existing = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  const nonEmpty = existing.filter(Boolean).length;
  if (!nonEmpty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const missing = headers.filter(function (h) { return existing.indexOf(h) < 0; });
    if (missing.length) {
      ensureCols_(sh, lastCol + missing.length);
      sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    }
  }
  const width = Math.max(sh.getLastColumn(), headers.length);
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground(COLORS.HEADER_BG);
  sh.setFrozenRows(1);
  return sh;
}

function ensureConfigSheet_(ss) {
  const sh = ensureTableSheet_(ss, SHEETS.CONFIG, CONFIG_HEADERS);
  ensureRows_(sh, CONFIG_SPEC.length + 5);
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat(FMT.TEXT);
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  const add = CONFIG_SPEC.filter(function (s) { return keys.indexOf(s[0]) < 0; }).map(function (s) { return [s[0], s[1], s[2], '']; });
  appendRows_(sh, add, 4);
  [230, 330, 560, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(2, 3, sh.getMaxRows() - 1, 2).setWrap(true).setFontColor(COLORS.MUTED);
  return sh;
}

/** Writes a ✅ / ⚠️ status next to each config key. */
function writeConfigStatus_(ctx) {
  const sh = ctx.ss.getSheetByName(SHEETS.CONFIG);
  if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1;
  const keys = sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const errors = ctx.cfg.errors;
  const out = keys.map(function (k) {
    if (!k) return [''];
    const e = errors.filter(function (x) { return x.indexOf(k + ' ') === 0 || x.indexOf(k + ' "') === 0; });
    return [e.length ? '⚠️ ' + e.join('; ') : '✅ OK'];
  });
  sh.getRange(2, 4, n, 1).setValues(out);
}

/**
 * insertCheckboxes() resets values to FALSE, so existing values are read first and written back.
 * Blank cells stay blank (an unchecked box with no value) so getLastRow() is not pushed down.
 */
function checkboxColumn_(sh, col, rows) {
  ensureRows_(sh, rows + 1);
  const rng = sh.getRange(2, col, rows, 1);
  const keep = rng.getValues().map(function (r) { return [r[0] === '' || r[0] === null ? '' : (r[0] === true || toBool_(r[0]))]; });
  rng.insertCheckboxes();
  rng.setValues(keep);
}

function formatTeam_(sh) {
  const idx = headerIndex_(sh);
  [170, 240, 240, 80, 300].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  if (idx['Active']) checkboxColumn_(sh, idx['Active'], 200);
  ['Email', 'Manager Email', 'Employee'].forEach(function (h) {
    if (idx[h]) sh.getRange(2, idx[h], 200, 1).setNumberFormat(FMT.TEXT);
  });
}

function formatTasks_(ss, sh, team) {
  const idx = headerIndex_(sh);
  [80, 230, 150, 130, 120, 380, 90, 130, 70, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  const rows = 300;
  ensureRows_(sh, rows + 1);
  if (idx['Task ID']) sh.getRange(2, idx['Task ID'], rows, 1).setNumberFormat(FMT.TEXT);
  if (idx['Due Time']) sh.getRange(2, idx['Due Time'], rows, 1).setNumberFormat(FMT.TEXT);
  if (idx['First Due Date']) sh.getRange(2, idx['First Due Date'], rows, 1).setNumberFormat(FMT.DATE);
  if (idx['Frequency']) {
    sh.getRange(2, idx['Frequency'], rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(FREQUENCY_OPTIONS, true).setAllowInvalid(true)
      .setHelpText('Daily, or "Every N days" (any N from 2 to 366).').build());
  }
  if (idx['Proof Type']) {
    sh.getRange(2, idx['Proof Type'], rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([PROOF.TICKET, PROOF.DAILY], true).setAllowInvalid(false).build());
  }
  if (idx['Assigned To']) {
    sh.getRange(2, idx['Assigned To'], rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(team.getRange(2, 1, 200, 1), true).setAllowInvalid(true).build());
  }
  if (idx['Active']) checkboxColumn_(sh, idx['Active'], rows);
  if (idx['Health']) sh.getRange(2, idx['Health'], rows, 1).setFontColor(COLORS.MUTED).setWrap(true);
  sh.getRange(2, 1, rows, sh.getLastColumn()).setVerticalAlignment('top');
  if (idx['Instructions']) sh.getRange(2, idx['Instructions'], rows, 1).setWrap(true);
}

function formatCycles_(sh) {
  const idx = headerIndex_(sh);
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  CYCLE_FIELDS.forEach(function (f) {
    const col = idx[f[1]];
    if (!col) return;
    let fm = FMT.TEXT;
    if (f[0] === 'slot') fm = FMT.DATE;
    else if (CYCLE_DATE_KEYS.indexOf(f[0]) >= 0) fm = FMT.LEDGER_DT;
    sh.getRange(2, col, rows, 1).setNumberFormat(fm);
  });
  sh.setTabColor('#9ca3af');
}

function formatLogs_(ss) {
  [SHEETS.LOG, SHEETS.EMAILS, SHEETS.TESTS].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    const rows = Math.max(sh.getMaxRows() - 1, 1);
    sh.getRange(2, 1, rows, 1).setNumberFormat(FMT.LEDGER_DT);
    // Text columns may carry user input (proof, notes): plain text so '=…' is never evaluated.
    if (n !== SHEETS.TESTS) sh.getRange(2, 2, rows, Math.max(sh.getLastColumn() - 1, 1)).setNumberFormat(FMT.TEXT);
    sh.setTabColor('#9ca3af');
  });
  const t = ss.getSheetByName(SHEETS.TESTS);
  if (t) { [60, 330, 80, 700, 110, 150].forEach(function (w, i) { t.setColumnWidth(i + 1, w); }); }
  const l = ss.getSheetByName(SHEETS.LOG);
  if (l) { [150, 170, 150, 230, 520].forEach(function (w, i) { l.setColumnWidth(i + 1, w); }); }
}

function headerIndex_(sh) {
  const lastCol = sh.getLastColumn();
  const out = {};
  if (!lastCol) return out;
  sh.getRange(1, 1, 1, lastCol).getValues()[0].forEach(function (h, i) { if (h !== '') out[String(h).trim()] = i + 1; });
  return out;
}

function ensureStartSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.START);
  if (!sh) sh = ss.insertSheet(SHEETS.START, 0);
  const lines = [
    ['Allo Recurring Tasks — start here'],
    [''],
    ['EMPLOYEES'],
    ['Open your own tab (👤 Your Name). Everything you need to act on is under ACTION NEEDED.'],
    ['When a task is done, type the Ticket ID (portal tasks) or paste the update you shared on WhatsApp (daily task) into the yellow cell and press Enter. That one entry is the whole update: the system records who and when, closes the task and schedules the next cycle.'],
    ['If the Result column shows ❌, read the reason, fix the entry and press Enter again.'],
    [''],
    ['MANAGER'],
    ['Manager Dashboard shows who is behind. If a task misses its cut-off you get ONE email the same day (cut-off + 30 min).'],
    ['Then: (1) check with the employee, (2) acknowledge by 11:00 the next working day (email link, or tick ✅ on the dashboard with a note), (3) get it done by the cut-off time on the next working day, or choose Waive if the person is on leave or it is a holiday.'],
    ['If you do not acknowledge, or the task is still open after that, leadership is emailed automatically (once).'],
    [''],
    ['ADMIN'],
    ['Tasks = the Task Master (add, change or deactivate recurring tasks). Team = people and their managers. Config = SLA times, email recipients and EMAIL_MODE (LOG_ONLY / LIVE).'],
    ['Menu ⚙️ Task System → Admin → Install triggers (once). The engine then runs every 15 minutes and edits are processed instantly.'],
    ['Demo: ⚙️ Task System → Demo clock moves a simulated time forward so you can watch reminders, alerts and escalations happen. Use "Back to real time" when finished.'],
    ['Proof of correctness: ⚙️ Task System → Automated tests writes PASS/FAIL with evidence to the Test Results sheet. Tests run on a throw-away copy, so your data is never touched.'],
    [''],
    ['System sheets (do not edit by hand): Cycles (ledger), Activity Log (audit trail), Email Log (every email), Test Results.'],
  ];
  ensureRows_(sh, lines.length + 1);
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.setColumnWidth(1, 980);
  sh.getRange(1, 1, lines.length, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  [3, 8, 13].forEach(function (r) { sh.getRange(r, 1).setFontWeight('bold').setBackground(COLORS.SECTION_BG); });
  sh.setHiddenGridlines(true);
  return sh;
}

function seedDemo_(ss, now) {
  const base = startOfDay_(now);
  const team = ss.getSheetByName(SHEETS.TEAM);
  const tasks = ss.getSheetByName(SHEETS.TASKS);
  const teamRows = [
    ['Employee A', 'employee.a@example.com', '', true, 'Owns the daily operations update'],
    ['Employee B', 'employee.b@example.com', '', true, ''],
    ['Employee C', 'employee.c@example.com', '', true, ''],
    ['Employee D', 'employee.d@example.com', '', true, ''],
  ];
  const portal = 'Complete the task on the portal and update the ticket, then enter the Ticket ID here.';
  const taskRows = [
    ['T01', 'Daily Operations Update', 'Employee A', 'Daily', PROOF.DAILY,
      'Post today\'s operations update in the Ops WhatsApp group, then paste the same text here.', '18:00', base, true, ''],
    ['T02', 'Portal Task A', 'Employee B', 'Every 7 days', PROOF.TICKET, portal, '18:00', base, true, ''],
    ['T03', 'Portal Task B', 'Employee C', 'Every 15 days', PROOF.TICKET, portal, '18:00', addDays_(base, 2), true, ''],
    ['T04', 'Portal Task C', 'Employee D', 'Every 30 days', PROOF.TICKET, portal, '18:00', addDays_(base, 5), true, ''],
  ];
  writeTableRows_(team, TEAM_FIELDS, teamRows);
  writeTableRows_(tasks, TASK_FIELDS, taskRows);
}

/** Writes rows given in *_FIELDS order into whatever column order the sheet actually has. */
function writeTableRows_(sh, fields, rows) {
  if (!rows.length) return;
  const idx = headerIndex_(sh);
  const width = sh.getLastColumn();
  const out = rows.map(function (r) {
    const line = [];
    for (let i = 0; i < width; i++) line.push('');
    fields.forEach(function (f, i) { if (idx[f[1]]) line[idx[f[1]] - 1] = r[i] === undefined || r[i] === null ? '' : r[i]; });
    return line;
  });
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + out.length);
  sh.getRange(start, 1, out.length, width).setValues(out);
}

function clearDataRows_(sh) {
  if (!sh) return;
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, Math.max(1, sh.getLastColumn())).clearContent();
}

function resetDemoData_(ss, now) {
  [SHEETS.TEAM, SHEETS.TASKS, SHEETS.CYCLES, SHEETS.LOG, SHEETS.EMAILS].forEach(function (n) { clearDataRows_(ss.getSheetByName(n)); });
  ss.getSheets().forEach(function (sh) { if (sh.getName().indexOf(EMPLOYEE_TAB_PREFIX) === 0) ss.deleteSheet(sh); });
  seedDemo_(ss, now);
}

function orderSheets_(ss) {
  const order = [SHEETS.START, SHEETS.DASHBOARD];
  ss.getSheets().forEach(function (sh) { if (sh.getName().indexOf(EMPLOYEE_TAB_PREFIX) === 0) order.push(sh.getName()); });
  order.push(SHEETS.TASKS, SHEETS.TEAM, SHEETS.CONFIG, SHEETS.CYCLES, SHEETS.LOG, SHEETS.EMAILS, SHEETS.TESTS);
  let pos = 1;
  order.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos++);
  });
  const start = ss.getSheetByName(SHEETS.START);
  if (start) ss.setActiveSheet(start);
}

function installTriggers_(ss) {
  const handlers = ['runEngine', 'handleEdit'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (handlers.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runEngine').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
}
