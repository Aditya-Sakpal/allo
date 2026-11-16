'use strict';
/**
 * npm test — runs (1) the shared in-sheet suite (T1–T8, B01–B22) against the strict GAS mock and
 * (2) Node-only integration checks of the public entry points (setup, triggers, menu, locks, LIVE mail,
 * demo clock, web app, test runner). Exit code 1 on any failure.
 */
process.env.TZ = 'Asia/Kolkata';
const assert = require('assert');
const { loadApp } = require('./load');

const results = [];
function check(name, fn) {
  const t0 = Date.now();
  try {
    fn();
    results.push({ name, ok: true, ms: Date.now() - t0 });
  } catch (e) {
    results.push({ name, ok: false, ms: Date.now() - t0, err: e && e.stack ? e.stack : String(e) });
  }
}

// ── 0. environment sanity ─────────────────────────────────────────────────────────────
check('Node TZ is Asia/Kolkata (matches appsscript.json)', () => {
  assert.strictEqual(new Date(2026, 10, 16, 10, 0).toISOString(), '2026-11-16T04:30:00.000Z');
});

// ── 1. shared suite ───────────────────────────────────────────────────────────────────
{
  const { app, gas } = loadApp();
  const sandbox = gas.newSpreadsheet('sandbox');
  const suite = app.runTestSuiteCore_(sandbox, 'all', 0);
  suite.forEach((r) => {
    results.push({ name: `[suite] ${r.id} ${r.name}`, ok: r.result === 'PASS', ms: r.ms, err: r.result === 'PASS' ? '' : r.evidence });
  });
  if (process.argv.includes('--evidence')) {
    suite.forEach((r) => console.log(`\n### ${r.id} ${r.name} — ${r.result}\n${r.evidence}`));
  }
}

// Sheets may keep a forcing apostrophe literally in plain-text cells instead of consuming it.
// The code must give the same results either way, so the whole suite runs again in that mode.
check('Whole suite also passes when Sheets keeps the apostrophe in plain-text cells', () => {
  const { app, gas } = loadApp({ keepApostropheInTextCells: true });
  const failed = app.runTestSuiteCore_(gas.newSpreadsheet('sandbox-apostrophe'), 'all', 0).filter((r) => r.result !== 'PASS');
  assert.strictEqual(failed.length, 0, failed.map((r) => r.id + ': ' + r.evidence).join('\n'));
});

// ── 2. node-only integration checks───────────────────────────────────────────────────
{
  const { app, gas } = loadApp({ activeUser: 'owner@allo.test' });
  const host = gas.newSpreadsheet('Allo Tasks');
  gas.setActive(host);
  const S = (n) => host.getSheetByName(n);
  const sheetsOf = () => host.getSheets().map((s) => s.getName());

  check('setupSystem() builds every sheet on a blank spreadsheet and seeds the demo', () => {
    app.setupSystem();
    const names = sheetsOf();
    ['Start Here', 'Manager Dashboard', 'Tasks', 'Team', 'Config', 'Cycles', 'Activity Log', 'Email Log', 'Test Results']
      .forEach((n) => assert.ok(names.includes(n), 'missing ' + n));
    assert.ok(!names.includes('Sheet1'), 'default Sheet1 removed');
    ['Employee A', 'Employee B', 'Employee C', 'Employee D'].forEach((e) => assert.ok(names.includes('👤 ' + e), 'tab for ' + e));
    assert.strictEqual(S('Team').getLastRow(), 5);
    assert.strictEqual(S('Tasks').getLastRow(), 5);
    assert.ok(S('Cycles').getLastRow() >= 2, 'cycles created');
    assert.strictEqual(gas.state.scriptProps.SPREADSHEET_ID, host.getId());
    assert.ok(gas.state.scriptProps.ACK_SECRET && gas.state.scriptProps.ACK_SECRET.length > 40, 'ack secret generated');
  });

  check('setupSystem() is idempotent (no duplicate config keys, data kept)', () => {
    const cyclesBefore = S('Cycles').getLastRow();
    S('Config').getRange(3, 2).setValue('ceo@allo.test');
    app.setupSystem();
    const keys = S('Config').getRange(2, 1, S('Config').getLastRow() - 1, 1).getValues().map((r) => r[0]);
    assert.strictEqual(new Set(keys).size, keys.length, 'duplicate config keys');
    assert.strictEqual(S('Config').getRange(3, 2).getValue(), 'ceo@allo.test', 'config edit kept');
    assert.strictEqual(S('Team').getLastRow(), 5);
    assert.strictEqual(S('Cycles').getLastRow(), cyclesBefore);
    assert.ok(S('Tasks').getRange(2, 9, 4, 1).getValues().every((r) => r[0] === true), 'Active checkboxes kept TRUE');
  });

  check('installTriggers() is idempotent: exactly one clock + one edit trigger', () => {
    app.installTriggers();
    app.installTriggers();
    const t = gas.state.triggers.map((x) => x.getHandlerFunction() + ':' + x.getEventType()).sort();
    assert.deepStrictEqual(t, ['handleEdit:ON_EDIT', 'runEngine:CLOCK']);
    assert.strictEqual(typeof app.runEngine, 'function');
    assert.strictEqual(typeof app.handleEdit, 'function');
  });

  check('onOpen() menu: every item points at an existing public function', () => {
    gas.state.menus = [];
    app.onOpen();
    const fns = [];
    const walk = (m) => m.items.forEach((i) => (i.sub ? walk(i.sub) : fns.push(i.fn)));
    walk(gas.state.menus[0]);
    assert.ok(fns.length >= 12, 'menu items');
    fns.forEach((f) => assert.strictEqual(typeof app[f], 'function', 'menu → missing function ' + f));
  });

  check('handleEdit (public trigger): employee completes from own tab, editor email recorded', () => {
    app.menuClockToday10();
    const tab = S('👤 Employee B');
    const ids = tab.getRange(1, 1, tab.getLastRow(), 1).getValues().map((r) => r[0]);
    const row = ids.findIndex((v) => String(v).startsWith('T02-')) + 1;
    assert.ok(row >= 7, 'T02 row on B tab');
    const range = tab.getRange(row, 5);
    range.setValue('PRT-90001');
    app.handleEdit({ range, source: host, value: 'PRT-90001', user: { getEmail: () => 'employee.b@example.com' } });
    const cyc = S('Cycles').getRange(1, 1, S('Cycles').getLastRow(), 23).getValues();
    const done = cyc.find((r) => String(r[0]).startsWith('T02-') && r[9] === 'DONE');
    assert.ok(done, 'T02 DONE in ledger');
    assert.strictEqual(done[11], 'Employee B <employee.b@example.com>');
  });

  check('Lock busy: engine skips safely; an edit is not lost silently (user told to retry)', () => {
    gas.state.lockBusy = true;
    try {
      const r = app.runEngine();
      assert.ok(r && r.skipped, 'engine skipped');
      const tab = S('👤 Employee A');
      const range = tab.getRange(7, 5);
      range.setValue('Ops update: everything shipped on time today, no issues.');
      app.handleEdit({ range, source: host });
      assert.ok(String(tab.getRange(7, 6).getValue()).includes('busy'), 'busy message shown in Result column');
    } finally {
      gas.state.lockBusy = false;
    }
  });

  check('handleEdit never throws to the user: unexpected error is reported in the Result cell', () => {
    const cycles = S('Cycles');
    const saved = cycles.getRange(1, 1, cycles.getLastRow(), 23).getValues();
    host.deleteSheet(cycles);
    const tab = S('👤 Employee A');
    const range = tab.getRange(7, 5);
    range.setValue('Ops update: everything shipped on time today, no issues.');
    assert.doesNotThrow(() => app.handleEdit({ range, source: host }));
    assert.ok(String(tab.getRange(7, 6).getValue()).includes('Setup'), 'message points to Setup');
    app.setupSystem_(host, { seed: false });
    S('Cycles').getRange(1, 1, saved.length, 23).setValues(saved);
  });

  check('Demo clock menu: moves SIMULATED_NOW, runs engine, alerts in LOG_ONLY, then back to real time', () => {
    app.menuClockToday10();
    const before = S('Email Log').getLastRow();
    app.menuClockAlert();
    const cfg = app.readConfig_(host);
    assert.strictEqual(cfg.simulatedNow.getHours(), 18);
    assert.strictEqual(cfg.simulatedNow.getMinutes(), 31);
    const log = S('Email Log').getRange(1, 1, S('Email Log').getLastRow(), 8).getValues();
    assert.ok(S('Email Log').getLastRow() > before, 'email logged');
    assert.ok(log.some((r) => r[1] === 'MANAGER_ALERT' && r[6] === 'LOG_ONLY'), 'alert recorded as LOG_ONLY');
    assert.strictEqual(gas.state.sentMail.length, 0, 'LOG_ONLY sends nothing');
    app.menuClockAckDeadline();
    assert.strictEqual(app.readConfig_(host).simulatedNow.getHours(), 11);
    app.menuClockResolveDeadline();
    assert.strictEqual(app.readConfig_(host).simulatedNow.getHours(), 18);
    app.menuClockReal();
    assert.strictEqual(app.readConfig_(host).simulatedNow, null);
  });

  check('LIVE mode: real MailApp calls with HTML + text; quota exhaustion is logged, not thrown', () => {
    app.menuResetDemo();
    const cfgSheet = S('Config');
    const keys = cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 1).getValues().map((r) => r[0]);
    cfgSheet.getRange(keys.indexOf('EMAIL_MODE') + 2, 2).setValue('LIVE');
    cfgSheet.getRange(keys.indexOf('MANAGER_EMAIL') + 2, 2).setValue('boss@allo.test');
    app.menuClockToday10();
    gas.state.sentMail = [];
    app.menuClockAlert();
    const alert = gas.state.sentMail.find((m) => /Action needed/.test(m.subject));
    assert.ok(alert, 'manager alert really sent');
    assert.strictEqual(alert.to, 'boss@allo.test');
    assert.ok(alert.htmlBody.includes('<table') && alert.body.includes('Pending after cut-off'), 'html + plain text');
    assert.ok(/Acknowledge<\/a>|✅ Ack/.test(alert.htmlBody), 'ack affordance');
    gas.state.mailQuota = 0;
    app.menuClockResolveDeadline();
    const log = S('Email Log').getRange(1, 1, S('Email Log').getLastRow(), 8).getValues();
    assert.ok(log.some((r) => r[6] === 'FAILED' && String(r[7]).includes('quota')), 'quota failure logged');
    gas.state.mailQuota = 100;
    cfgSheet.getRange(keys.indexOf('EMAIL_MODE') + 2, 2).setValue('LOG_ONLY');
    app.menuClockReal();
  });

  check('Web app: doGet renders HtmlOutput; confirmAck works end-to-end with the stored secret', () => {
    gas.state.webAppUrl = 'https://script.google.com/macros/s/LIVE/exec';
    app.menuResetDemo();
    app.menuClockToday10();
    app.menuClockAlert();
    const cyc = S('Cycles').getRange(2, 1, S('Cycles').getLastRow() - 1, 23).getValues();
    const alerted = cyc.find((r) => r[14] !== '' && r[9] === 'OPEN');
    assert.ok(alerted, 'an alerted cycle exists');
    const id = alerted[0];
    const sig = app.signId_(id);
    const page = app.doGet({ parameter: { a: 'ack', c: id, s: sig } });
    assert.ok(page.getContent().includes('Confirm acknowledgement'));
    assert.ok(page.getContent().includes('google.script.run'));
    const res = app.confirmAck(id, sig, 'Spoke to the employee');
    assert.ok(res.ok, res.msg);
    const after = S('Cycles').getRange(2, 1, S('Cycles').getLastRow() - 1, 23).getValues().find((r) => r[0] === id);
    assert.ok(after[15] !== '', 'ack stamped');
    assert.strictEqual(after[17], 'Spoke to the employee');
    const emailLog = S('Email Log').getRange(2, 1, S('Email Log').getLastRow() - 1, 8).getValues();
    assert.ok(emailLog.length > 0);
    app.menuClockReal();
    gas.state.webAppUrl = null;
  });

  check('Generated views never show "undefined", "NaN", "null" or "Invalid Date"', () => {
    app.menuResetDemo();
    app.menuClockToday10();
    app.menuClockAlert();
    app.menuClockResolveDeadline();
    host.getSheets().forEach((sh) => {
      const lr = sh.getLastRow();
      const lc = sh.getLastColumn();
      if (!lr || !lc) return;
      const text = sh.getRange(1, 1, lr, lc).getDisplayValues().map((r) => r.join('\t')).join('\n');
      ['undefined', 'NaN', 'Invalid Date', '[object Object]'].forEach((bad) => assert.ok(!text.includes(bad), `${sh.getName()} contains ${bad}`));
      if (sh.getName() !== 'Config') assert.ok(!/\bnull\b/.test(text), `${sh.getName()} contains null`);
    });
    app.menuClockReal();
  });

  check('Test menu: parts run separately (6-min limit), results accumulate to all 30, sandboxes trashed', () => {
    const before = Object.keys(gas.state.spreadsheets).length;
    const parts = app.testParts_();
    assert.ok(parts.length >= 4, 'suite split into parts');
    // every test belongs to exactly one part, and no part is bigger than 8 tests
    const ids = app.testCases_().map((tc) => tc.id);
    const owned = ids.map((id) => parts.filter((p) => app.testsInPart_(p.key).some((tc) => tc.id === id)).length);
    assert.ok(owned.every((n) => n === 1), 'each test in exactly one part');
    parts.forEach((p) => assert.ok(app.testsInPart_(p.key).length <= 8, p.key + ' too big'));
    // a leftover sandbox from a run Google killed is cleaned up by the next run
    gas.newSpreadsheet(app.testSandboxName_());
    app[parts[0].fn]();
    let sh = S('Test Results');
    let rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    assert.strictEqual(rows.length, 30, 'all tests always listed');
    assert.ok(rows.filter((r) => r[2] === 'NOT RUN').length === 30 - app.testsInPart_(parts[0].key).length, 'others shown as not run yet');
    assert.ok(String(rows[29][3]).includes(parts[parts.length - 1].label), 'not-run row says which menu item runs it');
    parts.slice(1).forEach((p) => app[p.fn]());
    sh = S('Test Results');
    rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    assert.strictEqual(rows.length, 30, 'no duplicate rows after re-runs');
    assert.strictEqual(rows.map((r) => r[0]).join(','), ids.join(','), 'kept in suite order');
    const failed = rows.filter((r) => r[2] !== 'PASS');
    assert.strictEqual(failed.length, 0, 'failures: ' + failed.map((r) => r[0]).join(','));
    assert.strictEqual(Object.keys(gas.state.spreadsheets).length, before + 1 + parts.length, 'one sandbox per part');
    assert.strictEqual(gas.state.trashed.length, parts.length + 1, 'every sandbox trashed, including the leftover');
    assert.ok(gas.state.alerts.some((a) => a.includes('30 / 30')), 'summary covers the whole suite');
  });

  check('Emails produced anywhere contain no "undefined"/"NaN"/"Invalid Date"', () => {
    const { app: a2, gas: g2 } = loadApp();
    const sb = g2.newSpreadsheet('sandbox2');
    const captured = [];
    const sbx = a2.makeSandbox_(sb);
    sbx.reset().at('2026-11-16 10:00').run();
    sbx.at('2026-11-16 16:00').run();
    sbx.at('2026-11-16 18:31').run();
    sbx.at('2026-11-17 11:01').run();
    sbx.at('2026-11-17 12:00');
    sbx.complete('T01-20261116', 'Ops update (late) with enough words in it to pass.');
    sbx.at('2026-11-17 12:15').run();
    captured.push(...sbx.mails);
    assert.ok(captured.length >= 4, 'reminder, alert, escalation, resolved');
    captured.forEach((m) => {
      [m.subject, m.htmlBody, m.body, m.to, m.cc].forEach((s) => {
        ['undefined', 'NaN', 'Invalid Date', '[object Object]'].forEach((bad) => assert.ok(!String(s).includes(bad), `${m.type} contains ${bad}`));
      });
    });
  });
}

// ── report ────────────────────────────────────────────────────────────────────────────
const failed = results.filter((r) => !r.ok);
results.forEach((r) => {
  console.log(`${r.ok ? '✔' : '✘'} ${r.name} (${r.ms} ms)`);
  if (!r.ok) console.log('    ' + String(r.err).split('\n').join('\n    '));
});
console.log(`\n${results.length - failed.length}/${results.length} passed`);
process.exit(failed.length ? 1 : 0);
