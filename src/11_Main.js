/**
 * 11_Main — entry points (triggers, menu, web app). Thin wrappers: lock → context → *Core_.
 * Public names (no trailing underscore) are what Apps Script triggers and the menu call.
 */

const MENU_NAME = '⚙️ Task System';

function withLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(waitMs || 25000)) return { skipped: true, reason: 'busy' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function safeUi_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

function alert_(msg) {
  const ui = safeUi_();
  if (ui) ui.alert(APP_NAME, msg, ui.ButtonSet.OK);
  else console.log(msg);
}

/* ───────────── core engine run ───────────── */

function runEngineCore_(ctx, opts) {
  opts = opts || {};
  const m = loadModel_(ctx);
  ensureCycles_(m);
  processNotifications_(ctx, m);
  saveModel_(ctx, m);
  writeConfigStatus_(ctx);
  PropertiesService.getScriptProperties().setProperty(lastRunKey_(ctx.ss), fmtShort_(m.now));
  if (opts.render !== false) renderAll_(ctx, m);
  const live = liveCycles_(m);
  return {
    cycles: live.length,
    open: live.filter(function (c) { return c.status === STATUS.OPEN; }).length,
    emailsSent: m.emailsSent || 0,
    warnings: uniq_(m.warnings),
  };
}

/* ───────────── triggers ───────────── */

/** Time-driven trigger (every 15 min) and menu "Run engine now". */
function runEngine() {
  return withLock_(function () {
    const ss = hostSpreadsheet_();
    return runEngineCore_(makeContext_(ss), { render: true });
  });
}

/** Installable onEdit trigger (created by installTriggers). Not named onEdit on purpose: no double firing. */
function handleEdit(e) {
  if (!e || !e.range) return;
  const ss = e.source || hostSpreadsheet_();
  const res = withLock_(function () {
    try {
      return handleEditCore_(makeContext_(ss), e);
    } catch (err) {
      console.error(err && err.stack ? err.stack : err);
      reportEditError_(ss, e, 'Something went wrong: ' + (err && err.message ? err.message : err) + '. Please try again or tell the admin.');
      return { handled: false, error: String(err) };
    }
  });
  if (res && res.skipped) reportEditError_(ss, e, '⏳ The system was busy. Please press Enter on the cell again in a moment.');
}

function reportEditError_(ss, e, msg) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName().indexOf(EMPLOYEE_TAB_PREFIX) === 0 && e.range.getRow() >= EMP_VIEW.FIRST_ROW) {
      sh.getRange(e.range.getRow(), EMP_VIEW.RESULT_COL).setValue(msg);
    }
    ss.toast(msg, APP_NAME, 8);
  } catch (x) { /* best effort */ }
}

/** Simple trigger: only builds the menu (simple triggers cannot do authorised work). */
function onOpen() {
  const ui = safeUi_();
  if (!ui) return;
  ui.createMenu(MENU_NAME)
    .addItem('▶ Run engine now', 'menuRunEngine')
    .addSubMenu(ui.createMenu('⏱ Demo clock')
      .addItem('Start of demo day (10:00 today, or next working day)', 'menuClockToday10')
      .addItem('+1 hour', 'menuClockPlus1h')
      .addItem('+3 hours', 'menuClockPlus3h')
      .addItem('+1 day', 'menuClockPlus1d')
      .addSeparator()
      .addItem('Jump to manager-alert time (cut-off + grace)', 'menuClockAlert')
      .addItem('Jump to ack deadline (next working day)', 'menuClockAckDeadline')
      .addItem('Jump to next cut-off + 1 min (resolve deadline)', 'menuClockResolveDeadline')
      .addSeparator()
      .addItem('Back to real time', 'menuClockReal'))
    .addSubMenu(testParts_().reduce(function (m, p) { return m.addItem(p.label, p.fn); }, ui.createMenu('🧪 Automated tests')))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠 Admin')
      .addItem('Setup / repair sheets', 'setupSystem')
      .addItem('Install triggers', 'installTriggers')
      .addItem('Reset demo data (sets demo clock)', 'menuResetDemo'))
    .addToUi();
}

/* ───────────── menu actions ───────────── */

function menuRunEngine() {
  const res = runEngine();
  if (res && res.skipped) return alert_('The engine is already running. Try again in a few seconds.');
  hostSpreadsheet_().toast('Engine ran · ' + res.open + ' open · ' + res.emailsSent + ' email batch(es)', APP_NAME, 5);
}

/** One-time install: builds sheets, registers the spreadsheet for the web app, renders views. */
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  withLock_(function () {
    setupSystem_(ss, { seed: true });
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    ackSecret_();
    runEngineCore_(makeContext_(ss), { render: true });
    orderSheets_(ss);
  });
  onOpen();
  alert_('Setup complete.\n\nNext: ⚙️ Task System → Admin → Install triggers.\n' +
    'Emails are in LOG_ONLY mode (see Email Log). Put real addresses in Team/Config and set EMAIL_MODE = LIVE to send them.');
}

function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  installTriggers_(ss);
  alert_('Triggers installed: engine every 15 minutes, edits processed instantly.');
}

function menuResetDemo() {
  const ui = safeUi_();
  if (ui) {
    const ok = ui.alert(APP_NAME, 'Reset Team, Tasks, Cycles and logs to the demo scenario? This deletes current data.', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }
  const ss = hostSpreadsheet_();
  const res = withLock_(function () {
    const start = demoStart_(readConfig_(ss));
    resetDemoData_(ss, start);
    setDemoClock_(ss, start);
    runEngineCore_(makeContext_(ss), { render: true });
    orderSheets_(ss);
    return { start: start };
  });
  if (res && res.skipped) return alert_('The engine is busy. Try again in a few seconds.');
  ss.toast('Demo data loaded. Demo clock set to ' + fmtLong_(res.start) + '. Use Demo clock → Back to real time when finished.', APP_NAME, 8);
}

/* Demo clock: moves Config.SIMULATED_NOW and runs the engine so time-based behaviour can be shown live. */

function setDemoClock_(ss, d) {
  const sh = requireSheet_(ss, SHEETS.CONFIG);
  const last = sh.getLastRow();
  const keys = sh.getRange(2, 1, Math.max(last - 1, 1), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const i = keys.indexOf('SIMULATED_NOW');
  if (i < 0) throw new Error('SIMULATED_NOW is missing from Config. Run Setup.');
  sh.getRange(i + 2, 2).setValue(d ? fmt_(d, 'yyyy-MM-dd HH:mm') : '');
}

function moveDemoClock_(fn) {
  const ss = hostSpreadsheet_();
  const res = withLock_(function () {
    const cfg = readConfig_(ss);
    const current = cfg.simulatedNow || new Date();
    const next = fn(current, cfg);
    setDemoClock_(ss, next);
    return { next: next, run: runEngineCore_(makeContext_(ss), { render: true }) };
  });
  if (res && res.skipped) return alert_('The engine is busy. Try again in a few seconds.');
  ss.toast((res.next ? 'Demo clock: ' + fmtLong_(res.next) : 'Back to real time') + ' · ' + res.run.emailsSent + ' email batch(es) this run', APP_NAME, 6);
}

/** 10:00 today, or on the next working day if today is a weekend/holiday (otherwise nothing would be due). */
function demoStart_(cfg) {
  return atTime_(onOrNextWorkingDay_(new Date(), cfg), { h: 10, m: 0 });
}

function menuClockToday10() { moveDemoClock_(function (d, cfg) { return demoStart_(cfg); }); }
function menuClockPlus1h() { moveDemoClock_(function (d) { return addMinutes_(d, 60); }); }
function menuClockPlus3h() { moveDemoClock_(function (d) { return addMinutes_(d, 180); }); }
function menuClockPlus1d() { moveDemoClock_(function (d) { return addDays_(d, 1); }); }
function menuClockAlert() {
  moveDemoClock_(function (d, cfg) { return addMinutes_(atTime_(d, cfg.defaultDueHM), cfg.graceMin + 1); });
}
function menuClockAckDeadline() {
  moveDemoClock_(function (d, cfg) { return addMinutes_(atTime_(nextWorkingDay_(d, cfg), cfg.ackHM), 1); });
}
/** Next cut-off + 1 min: from Mon 18:31 → Tue 18:01, from Tue 11:01 → Tue 18:01 (= the resolve deadline). */
function menuClockResolveDeadline() {
  moveDemoClock_(function (d, cfg) {
    const today = addMinutes_(atTime_(d, cfg.defaultDueHM), 1);
    if (isWorkingDay_(d, cfg) && d.getTime() < today.getTime()) return today;
    return addMinutes_(atTime_(nextWorkingDay_(d, cfg), cfg.defaultDueHM), 1);
  });
}
function menuClockReal() { moveDemoClock_(function () { return null; }); }

// One entry point per part of testParts_(): each must finish inside Google's 6-minute limit.
function menuTestsCore() { runTestsToSheet_('core'); }
function menuTestsBreak1() { runTestsToSheet_('break1'); }
function menuTestsBreak2() { runTestsToSheet_('break2'); }
function menuTestsBreak3() { runTestsToSheet_('break3'); }
function menuTestsBreak4() { runTestsToSheet_('break4'); }

/* ───────────── web app ───────────── */

function doGet(e) {
  const ss = hostSpreadsheet_();
  const page = doGetCore_(makeContext_(ss), e && e.parameter);
  return HtmlService.createHtmlOutput(page.html)
    .setTitle(page.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called from the confirmation page via google.script.run. */
function confirmAck(id, sig, note) {
  const res = withLock_(function () {
    return confirmAckCore_(makeContext_(hostSpreadsheet_()), id, sig, note);
  });
  if (res && res.skipped) return { ok: false, msg: 'The system is busy — please press the button again.' };
  return res;
}
