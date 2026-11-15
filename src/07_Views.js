/**
 * 07_Views — generated, disposable views: one tab per employee and the Manager Dashboard.
 * Spec: F02 (employee view), F06 (dashboard). Rendered with one batched write per property.
 */

function employeeTabName_(name) {
  return EMPLOYEE_TAB_PREFIX + sanitizeSheetName_(name);
}

/* ───────────── grid builder: values + formats + colours, flushed in a handful of calls ───────────── */

function newGrid_(width) {
  return { width: width, values: [], formats: [], bgs: [], weights: [], colors: [], checks: [], lists: [], wraps: [] };
}

/**
 * opts: {bold, bg, color, cellBg: {colIndex0: '#hex'}, cellColor: {...}}. Returns the 1-based row number.
 * Only explicit {formula: '=...'} cells become formulas. Every string is written as plain text ('@'),
 * so user-supplied text such as '=IMPORTDATA(...)' is never evaluated (formula-injection safe).
 */
function addRow_(g, cells, opts) {
  opts = opts || {};
  const vals = [];
  const fm = [];
  const bg = [];
  const wt = [];
  const fc = [];
  for (let i = 0; i < g.width; i++) {
    let v = cells ? cells[i] : '';
    if (v === undefined || v === null) v = '';
    let f = FMT.TEXT;
    if (v && typeof v === 'object' && !isDate_(v) && v.formula) { v = v.formula; f = FMT.PLAIN; }
    else if (isDate_(v)) f = FMT.VIEW_DT;
    else if (typeof v === 'boolean' || typeof v === 'number') f = FMT.PLAIN;
    vals.push(v);
    fm.push(f);
    bg.push((opts.cellBg && opts.cellBg[i]) || opts.bg || COLORS.WHITE);
    wt.push(opts.bold ? 'bold' : 'normal');
    fc.push((opts.cellColor && opts.cellColor[i]) || opts.color || COLORS.INK);
  }
  g.values.push(vals);
  g.formats.push(fm);
  g.bgs.push(bg);
  g.weights.push(wt);
  g.colors.push(fc);
  return g.values.length;
}

function flushGrid_(sh, g) {
  const n = g.values.length;
  const W = g.width;
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  ensureRows_(sh, n + 2);
  ensureCols_(sh, W);
  const rng = sh.getRange(1, 1, n, W);
  rng.setNumberFormats(g.formats);
  // insertCheckboxes resets cells to FALSE, so it must run before the values are written.
  g.checks.forEach(function (c) { if (c[2] > 0) sh.getRange(c[0], c[1], c[2], 1).insertCheckboxes(); });
  g.lists.forEach(function (l) {
    if (l[2] > 0) {
      sh.getRange(l[0], l[1], l[2], 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(l[3], true).setAllowInvalid(false).build());
    }
  });
  rng.setValues(g.values);
  rng.setBackgrounds(g.bgs);
  rng.setFontWeights(g.weights);
  rng.setFontColors(g.colors);
  rng.setVerticalAlignment('top');
  g.wraps.forEach(function (w) { if (w[2] > 0) sh.getRange(w[0], w[1], w[2], w[3]).setWrap(true); });
  sh.getRange(1, TITLE_COL).setFontSize(15);
}

// Literals (not COLORS.*) because top-level code must not depend on load order of other files.
const SECTION_STYLE = { bold: true, bg: '#e8edf5' };
const HEADER_STYLE = { bold: true, bg: '#f3f4f6' };
const MUTED_STYLE = { color: '#6b7280' };

/* ───────────── shared labels ───────────── */

function openStatus_(c, m) {
  const now = m.now;
  if (now.getTime() < c.opensAt.getTime()) return { label: 'Opens ' + fmtShort_(c.opensAt), bg: COLORS.QUIET_BG };
  if (now.getTime() > c.dueAt.getTime()) {
    const by = durationText_(now - c.dueAt);
    if (c.escalatedAt) return { label: '🔴 Overdue ' + by + ' — escalated to leadership', bg: COLORS.ESC_BG };
    if (c.alertAt) return { label: '🟠 Overdue ' + by + ' — manager alerted', bg: COLORS.BAD_BG };
    return { label: '⚠️ Overdue by ' + by, bg: COLORS.BAD_BG };
  }
  if (sameDay_(now, c.dueAt)) return { label: '⏰ Due today ' + fmt_(c.dueAt, 'HH:mm'), bg: COLORS.WARN_BG };
  const days = Math.round((startOfDay_(c.dueAt) - startOfDay_(now)) / 86400000);
  return { label: days === 1 ? 'Due tomorrow' : 'Due in ' + days + ' days', bg: COLORS.OK_BG };
}

function closedAt_(c) {
  return c.completedAt || c.updatedAt;
}

function waShareFormula_(text) {
  const url = 'https://wa.me/?text=' + encodeURIComponent(text);
  if (url.length > 1800) return 'Text too long for a share link';
  return { formula: '=HYPERLINK("' + url + '","📤 Share on WhatsApp")' };
}

/* ───────────── employee tabs ───────────── */
// Visible columns: B Task · C Status · D Proof required · E ✍️ input · F Result · G What to do · H Cut-off · I Frequency
// (A = hidden Cycle ID). Input and feedback sit next to the task so they fit on a phone screen.

function styleEmployeeTab_(sh) {
  [40, 200, 210, 220, 300, 300, 280, 140, 110].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(ID_COL);
  sh.setHiddenGridlines(true);
  sh.setTabColor('#f59e0b');
  try {
    const p = sh.protect().setDescription('Generated view — only the "✍️ Enter Ticket ID / update here" column is for input').setWarningOnly(true);
    p.setUnprotectedRanges([sh.getRange(EMP_VIEW.FIRST_ROW, EMP_VIEW.INPUT_COL, sh.getMaxRows() - EMP_VIEW.FIRST_ROW + 1, 1)]);
  } catch (e) {
    // Protection is a convenience; never block rendering on it.
  }
}

/** only: optional list of employee names to re-render (edits touch one tab, not all of them). */
function renderEmployeeTabs_(ctx, m, only) {
  const ss = ctx.ss;
  const wanted = {};
  m.team.filter(function (e) { return e.active; }).forEach(function (e) {
    const name = employeeTabName_(e.name);
    wanted[name] = true;
    if (only && only.indexOf(e.name) < 0 && ss.getSheetByName(name)) return;
    let sh = ss.getSheetByName(name);
    if (!sh) {
      const dash = ss.getSheetByName(SHEETS.DASHBOARD);
      const idx = dash ? ss.getSheets().indexOf(dash) + 1 : ss.getSheets().length;
      sh = ss.insertSheet(name, idx);
      styleEmployeeTab_(sh);
    }
    renderEmployeeTab_(ctx, m, e, sh);
  });
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (n.indexOf(EMPLOYEE_TAB_PREFIX) === 0 && !wanted[n]) ss.deleteSheet(sh);
  });
}

function renderEmployeeTab_(ctx, m, e, sh) {
  const now = m.now;
  const cfg = m.cfg;
  const mine = liveCycles_(m).filter(function (c) { return c.employee === e.name; });
  const open = mine.filter(function (c) { return c.status === STATUS.OPEN; }).sort(byDue_);
  const actionable = open.filter(function (c) { return now.getTime() >= c.opensAt.getTime(); });
  const upcoming = open.filter(function (c) { return now.getTime() < c.opensAt.getTime(); });
  const since = addDays_(startOfDay_(now), -cfg.recentDays);
  const done = mine.filter(function (c) {
    return c.status !== STATUS.OPEN && closedAt_(c) && closedAt_(c).getTime() >= since.getTime();
  }).sort(function (a, b) { return closedAt_(b) - closedAt_(a); });
  const overdue = actionable.filter(function (c) { return now.getTime() > c.dueAt.getTime(); }).length;

  const g = newGrid_(EMP_VIEW.WIDTH);
  addRow_(g, ['', e.name + ' — my recurring tasks'], { bold: true });
  addRow_(g, ['', 'When a task is done, type the Ticket ID (portal tasks) or paste the update you shared on WhatsApp (daily task) ' +
    'in the “✍️ Enter Ticket ID / update here” column and press Enter. That one entry records who and when, closes the task ' +
    'and schedules the next one. There is nothing else to update.'], MUTED_STYLE);
  addRow_(g, ['', 'As of ' + fmtLong_(now) + '  ·  ' + actionable.length + ' need action  ·  ' + overdue + ' overdue'],
    { color: overdue ? '#b42318' : COLORS.MUTED, bold: overdue > 0 });
  addRow_(g, []);
  addRow_(g, ['', 'ACTION NEEDED'], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Task', 'Status', 'Proof required', '✍️ Enter Ticket ID / update here', 'Result',
    'What to do', 'Cut-off', 'Frequency'], HEADER_STYLE);
  const firstAction = g.values.length + 1;
  if (!actionable.length) addRow_(g, ['', '🎉 Nothing needs your action right now.'], MUTED_STYLE);
  actionable.forEach(function (c) {
    const st = openStatus_(c, m);
    const t = m.taskById[c.taskId];
    addRow_(g, [c.id, c.taskName, st.label, proofHint_(c, cfg), '', '', t ? t.instructions : '', c.dueAt, c.frequency],
      { cellBg: { 2: st.bg, 4: COLORS.INPUT_BG } });
  });
  g.wraps.push([firstAction, 2, g.values.length - firstAction + 1, 6]);

  addRow_(g, []);
  addRow_(g, ['', 'COMING UP — not open for entry yet'], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Task', 'Opens for entry', 'Cut-off', '', '', '', '', 'Frequency'], HEADER_STYLE);
  if (!upcoming.length) addRow_(g, ['', 'Nothing scheduled yet.'], MUTED_STYLE);
  upcoming.forEach(function (c) {
    addRow_(g, [c.id, c.taskName, c.opensAt, c.dueAt, '', '', '', '', c.frequency], { color: COLORS.MUTED });
  });

  addRow_(g, []);
  addRow_(g, ['', 'COMPLETED — last ' + cfg.recentDays + ' days'], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Task', 'Result', 'Closed at', 'Share', 'Proof / update', '', 'Cut-off', ''], HEADER_STYLE);
  const firstDone = g.values.length + 1;
  if (!done.length) addRow_(g, ['', 'Nothing completed in this period yet.'], MUTED_STYLE);
  done.forEach(function (c) {
    const share = c.proofType === PROOF.DAILY && c.proof && (c.status === STATUS.DONE || c.status === STATUS.LATE)
      ? waShareFormula_(c.proof) : '';
    const bg = c.status === STATUS.DONE ? COLORS.OK_BG : (c.status === STATUS.LATE ? COLORS.WARN_BG : COLORS.QUIET_BG);
    addRow_(g, [c.id, c.taskName, (c.status === STATUS.DONE || c.status === STATUS.LATE ? '✅ ' : '') + statusLabel_(c.status),
      closedAt_(c), share, c.proof || c.closeNote, '', c.dueAt, ''], { cellBg: { 2: bg } });
  });
  g.wraps.push([firstDone, 6, g.values.length - firstDone + 1, 1]);
  flushGrid_(sh, g);
}

/* ───────────── manager dashboard ───────────── */
// Needs-attention columns: B Employee · C Task · D Overdue by · E Task status · F Manager note · G ✅ Ack · H Action ·
// I Escalation · J Cut-off · K Manager alerted · L Acknowledged at (A = hidden Cycle ID). Note sits before the tick
// so "type a note, then tick" reads left to right.

function styleDashboard_(sh) {
  [40, 130, 190, 110, 170, 240, 60, 170, 300, 140, 140, 140].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(ID_COL);
  sh.setHiddenGridlines(true);
  sh.setTabColor('#1f2933');
}

function lastRunKey_(ss) {
  return 'LAST_RUN_' + ss.getId();
}

function renderDashboard_(ctx, m) {
  const ss = ctx.ss;
  let sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.DASHBOARD, 0);
    styleDashboard_(sh);
  }
  const now = m.now;
  const cfg = m.cfg;
  const live = liveCycles_(m);
  const since = addDays_(startOfDay_(now), -cfg.recentDays);
  const lastRun = PropertiesService.getScriptProperties().getProperty(lastRunKey_(ss));
  const g = newGrid_(DASH.WIDTH);

  addRow_(g, ['', 'Manager Dashboard — recurring tasks'], { bold: true });
  addRow_(g, ['', 'For an overdue task: type a note (optional), tick ✅ Ack, and make sure it gets done. ' +
    'Choose Action → Waive only if the person is on leave or it is a holiday. Everything else updates automatically.'], MUTED_STYLE);
  addRow_(g, ['', 'As of ' + fmtLong_(now) + (lastRun ? '  ·  engine last ran ' + lastRun : '') + '  ·  email mode: ' + cfg.emailMode +
    (cfg.simulatedNow ? '  ·  ⏱ DEMO CLOCK ACTIVE (Config → SIMULATED_NOW)' : '')], MUTED_STYLE);
  const warnings = uniq_(m.warnings);
  if (warnings.length) addRow_(g, ['', '⚠️ ' + warnings.slice(0, 6).join('   |   ') + (warnings.length > 6 ? '   (+' + (warnings.length - 6) + ' more)' : '')], { bg: COLORS.WARN_BG });
  else addRow_(g, []);
  addRow_(g, []);

  // Team overview
  addRow_(g, ['', DASH.TITLE_OVERVIEW], SECTION_STYLE);
  addRow_(g, ['', 'Employee', 'Need action', 'Due today', 'Overdue', 'Awaiting your ack', 'Escalated', 'Done (7 days)', 'On-time % (30 days)'], HEADER_STYLE);
  const activeTeam = m.team.filter(function (e) { return e.active; });
  if (!activeTeam.length) addRow_(g, ['', 'No active employees on the Team sheet.'], MUTED_STYLE);
  activeTeam.forEach(function (e) {
    const mine = live.filter(function (c) { return c.employee === e.name; });
    const open = mine.filter(function (c) { return c.status === STATUS.OPEN; });
    const need = open.filter(function (c) { return now.getTime() >= c.opensAt.getTime(); }).length;
    const dueToday = open.filter(function (c) { return now.getTime() <= c.dueAt.getTime() && sameDay_(now, c.dueAt); }).length;
    const overdue = open.filter(function (c) { return now.getTime() > c.dueAt.getTime(); }).length;
    const awaiting = open.filter(function (c) { return c.alertAt && !c.ackAt; }).length;
    const escalated = open.filter(function (c) { return c.escalatedAt; }).length;
    const week = addDays_(now, -7).getTime();
    const done7 = mine.filter(function (c) { return (c.status === STATUS.DONE || c.status === STATUS.LATE) && c.completedAt && c.completedAt.getTime() >= week; }).length;
    const month = addDays_(now, -30).getTime();
    const due30 = mine.filter(function (c) {
      return c.dueAt.getTime() <= now.getTime() && c.dueAt.getTime() >= month && c.status !== STATUS.WAIVED;
    });
    const onTime = due30.filter(function (c) { return c.status === STATUS.DONE; }).length;
    const pct = due30.length ? Math.round(onTime * 100 / due30.length) + '%' : '—';
    addRow_(g, ['', e.name, need, dueToday, overdue, awaiting, escalated, done7, pct], {
      cellBg: {
        3: dueToday ? COLORS.WARN_BG : null,
        4: overdue ? COLORS.BAD_BG : COLORS.OK_BG,
        5: awaiting ? COLORS.BAD_BG : null,
        6: escalated ? COLORS.ESC_BG : null,
      },
    });
  });
  addRow_(g, []);

  // Needs attention
  addRow_(g, ['', DASH.TITLE_ATTENTION], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Employee', 'Task', 'Overdue by', 'Task status', 'Manager note', '✅ Ack', 'Action',
    'Escalation', 'Cut-off', 'Manager alerted', 'Acknowledged at'], HEADER_STYLE);
  const attn = live.filter(function (c) {
    if (c.status === STATUS.OPEN) return now.getTime() > c.dueAt.getTime();
    return !!c.alertAt && c.status !== STATUS.SKIPPED && closedAt_(c).getTime() >= since.getTime();
  }).sort(function (a, b) {
    const ao = a.status === STATUS.OPEN ? 0 : 1;
    const bo = b.status === STATUS.OPEN ? 0 : 1;
    return ao - bo || a.dueAt - b.dueAt;
  });
  const firstAttn = g.values.length + 1;
  if (!attn.length) addRow_(g, ['', '✅ Nothing overdue. No alerts in the last ' + cfg.recentDays + ' days.'], MUTED_STYLE);
  let openCount = 0;
  attn.forEach(function (c) {
    const isOpen = c.status === STATUS.OPEN;
    if (isOpen) openCount++;
    let status;
    if (isOpen) status = c.escalatedAt ? '🔴 Open — escalated' : '🟠 Open';
    else if (c.status === STATUS.WAIVED) status = '⚪ Waived';
    else status = '✅ ' + statusLabel_(c.status) + (c.alertAt && c.completedAt && c.completedAt.getTime() >= c.alertAt.getTime() ? ' (after alert)' : '');
    const overdueBy = isOpen ? durationText_(now - c.dueAt) : 'closed ' + fmtShort_(closedAt_(c));
    const alerted = c.alertAt || (isOpen ? 'due ' + fmtShort_(managerAlertAt_(c.dueAt, cfg)) : '—');
    let esc;
    if (c.escalatedAt) esc = '🔴 ' + fmtShort_(c.escalatedAt) + ' — ' + c.escalationReason;
    else if (!isOpen) esc = '—';
    else if (c.alertAt) {
      esc = (c.ackAt ? '' : 'Leadership at ' + fmtShort_(ackDeadline_(c.alertAt, cfg)) + ' if not acknowledged · ') +
        'at ' + fmtShort_(resolveDeadline_(c.dueAt, cfg)) + ' if not done';
    } else esc = 'Manager alert due ' + fmtShort_(managerAlertAt_(c.dueAt, cfg));
    const statusBg = isOpen ? (c.escalatedAt ? COLORS.ESC_BG : COLORS.BAD_BG) : COLORS.OK_BG;
    // Closed rows are a record, not a to-do: no checkbox, no yellow input affordance.
    const ackCell = isOpen ? !!c.ackAt : (c.ackAt ? '✅' : '—');
    addRow_(g, [c.id, c.employee, c.taskName, overdueBy, status, c.ackNote, ackCell, '', esc, c.dueAt, alerted, c.ackAt || ''], {
      cellBg: {
        4: statusBg,
        5: isOpen ? COLORS.INPUT_BG : COLORS.QUIET_BG,
        6: isOpen ? COLORS.INPUT_BG : COLORS.QUIET_BG,
        7: isOpen ? COLORS.INPUT_BG : null,
      },
    });
  });
  if (attn.length) {
    g.checks.push([firstAttn, DASH.ACK_COL, openCount]);
    g.lists.push([firstAttn, DASH.ACTION_COL, openCount, [ACTION_WAIVE]]);
    g.wraps.push([firstAttn, 3, attn.length, 1], [firstAttn, DASH.NOTE_COL, attn.length, 1], [firstAttn, 9, attn.length, 1]);
  }
  addRow_(g, []);

  // All open
  const allOpen = live.filter(function (c) { return c.status === STATUS.OPEN; }).sort(byDue_);
  addRow_(g, ['', DASH.TITLE_OPEN + ' (' + allOpen.length + ')'], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Employee', 'Task', 'Status', 'Cut-off', 'Frequency'], HEADER_STYLE);
  if (!allOpen.length) addRow_(g, ['', 'No open tasks.'], MUTED_STYLE);
  allOpen.forEach(function (c) {
    const st = openStatus_(c, m);
    addRow_(g, [c.id, c.employee, c.taskName, st.label, c.dueAt, c.frequency], { cellBg: { 3: st.bg } });
  });
  addRow_(g, []);

  // Recent completions
  addRow_(g, ['', DASH.TITLE_RECENT + ' — last ' + cfg.recentDays + ' days'], SECTION_STYLE);
  addRow_(g, ['Cycle ID', 'Employee', 'Task', 'Result', 'Completed at', 'Completed by', 'Proof / update', 'Cut-off'], HEADER_STYLE);
  const recent = live.filter(function (c) {
    return (c.status === STATUS.DONE || c.status === STATUS.LATE) && c.completedAt && c.completedAt.getTime() >= since.getTime();
  }).sort(function (a, b) { return b.completedAt - a.completedAt; });
  const firstRecent = g.values.length + 1;
  if (!recent.length) addRow_(g, ['', 'No completions in this period yet.'], MUTED_STYLE);
  recent.forEach(function (c) {
    addRow_(g, [c.id, c.employee, c.taskName, statusLabel_(c.status), c.completedAt, c.completedBy, c.proof, c.dueAt], {
      cellBg: { 3: c.status === STATUS.DONE ? COLORS.OK_BG : COLORS.WARN_BG },
    });
  });
  g.wraps.push([firstRecent, DASH.RECENT_PROOF_COL, g.values.length - firstRecent + 1, 1]);
  flushGrid_(sh, g);
}

/** only: optional employee names whose tabs changed; the dashboard is always refreshed. */
function renderAll_(ctx, m, only) {
  renderDashboard_(ctx, m);
  renderEmployeeTabs_(ctx, m, only);
}
