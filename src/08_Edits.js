/**
 * 08_Edits — routes a spreadsheet edit to the right action.
 *   👤 employee tab, column H  → completion (F02/F03/F04)
 *   Manager Dashboard          → acknowledge / note / waive (F06)
 *   Tasks / Team / Config      → full engine run so changes apply immediately (F01 R5)
 * Returns a plain result object so tests can assert on it.
 */

/** Best-effort identity of the editor. Consumer accounts often hide other users' emails. */
function actorFromEvent_(e, onBehalfOf) {
  let email = '';
  try {
    if (e && e.user) email = typeof e.user === 'string' ? e.user : (e.user.getEmail ? e.user.getEmail() : '');
  } catch (x) { email = ''; }
  if (!email) {
    try { email = Session.getActiveUser().getEmail(); } catch (x) { email = ''; }
  }
  email = String(email || '').trim().toLowerCase();
  if (!onBehalfOf) return email || '';
  if (!email) return onBehalfOf.name + ' (via own tab)';
  if (onBehalfOf.email && email === onBehalfOf.email) return onBehalfOf.name + ' <' + email + '>';
  return email + ' (on behalf of ' + onBehalfOf.name + ')';
}

function toast_(ctx, msg, title) {
  try { ctx.ss.toast(truncate_(msg, 250), title || APP_NAME, 6); } catch (e) { /* toast is cosmetic */ }
}

function handleEditCore_(ctx, e) {
  if (!e || !e.range) return { handled: false, reason: 'no range' };
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (name.indexOf(EMPLOYEE_TAB_PREFIX) === 0) return handleEmployeeEdit_(ctx, e, sh);
  if (name === SHEETS.DASHBOARD) return handleDashboardEdit_(ctx, e, sh);
  if (name === SHEETS.TASKS || name === SHEETS.TEAM || name === SHEETS.CONFIG) {
    if (name === SHEETS.TASKS && e.range.getColumn() === TASK_FIELDS.length && e.range.getNumColumns() === 1) {
      return { handled: false, reason: 'health column is system-written' };
    }
    const summary = runEngineCore_(ctx, { render: true });
    return { handled: true, kind: 'MASTER_DATA', summary: summary };
  }
  return { handled: false, reason: 'sheet not managed' };
}

function handleEmployeeEdit_(ctx, e, sh) {
  const r0 = e.range.getRow();
  const c0 = e.range.getColumn();
  const nr = e.range.getNumRows();
  const nc = e.range.getNumColumns();
  const IC = EMP_VIEW.INPUT_COL;
  if (c0 > IC || c0 + nc - 1 < IC) return { handled: false, reason: 'not the input column' };
  const first = Math.max(r0, EMP_VIEW.FIRST_ROW);
  const last = r0 + nr - 1;
  if (last < first) return { handled: false, reason: 'title/header rows' };

  const block = sh.getRange(first, 1, last - first + 1, EMP_VIEW.RESULT_COL).getValues()
    .map(function (row) { return row.map(fromCell_); });
  // The 15-minute refresh may have redrawn the tab while the user was typing: the cell then no longer
  // holds what they typed and the row may now be a different task. Never guess — ask them to re-enter.
  if (nr === 1 && nc === 1 && e.value !== undefined && e.value !== null &&
      String(block[0][IC - 1]) !== String(e.value)) {
    const msg = 'The sheet refreshed while you were typing, so nothing was saved. Please enter "' +
      truncate_(e.value, 40) + '" again on the right row.';
    toast_(ctx, msg);
    return { handled: true, kind: 'COMPLETION', outcomes: [{ row: r0, ok: false, code: 'REFRESHED', msg: msg }] };
  }

  const m = loadModel_(ctx);
  const tabName = sh.getName();
  const emp = m.team.find(function (t) { return employeeTabName_(t.name) === tabName; });
  if (!emp) return { handled: false, reason: 'tab does not match an employee' };
  const actor = actorFromEvent_(e, emp);

  const outcomes = [];
  block.forEach(function (row, i) {
    const input = row[IC - 1];
    if (input === '' || input === null || input === undefined) return;
    const id = String(row[0] || '').trim();
    const rowNum = first + i;
    let res;
    if (!id) res = { ok: false, code: 'NO_ROW', msg: 'Enter the ticket/update on a task row (one with a Cycle ID in column A).' };
    else if (!m.byId[id]) res = { ok: false, code: 'UNKNOWN', msg: 'This row is out of date — the sheet has been refreshed, please enter it again on the right row.' };
    else if (m.byId[id].employee !== emp.name) res = { ok: false, code: 'NOT_YOURS', msg: 'This task belongs to ' + m.byId[id].employee + '.' };
    else res = completeCycle_(m, m.byId[id], input, actor);
    outcomes.push(Object.assign({ row: rowNum, id: id, input: input }, res));
  });
  if (!outcomes.length) return { handled: false, reason: 'cell cleared' };

  const anyOk = outcomes.some(function (o) { return o.ok; });
  const stale = outcomes.some(function (o) { return o.code === 'NOT_OPEN' || o.code === 'UNKNOWN'; });
  const failures = outcomes.filter(function (o) { return !o.ok; });
  if (anyOk || stale) {
    saveModel_(ctx, m);
    renderAll_(ctx, m, [emp.name]);
    annotateEmployeeRows_(ctx.ss.getSheetByName(tabName), failures.filter(function (o) { return o.code !== 'NOT_OPEN' && o.code !== 'UNKNOWN'; }));
  } else {
    flushLogs_(ctx, m);
    failures.forEach(function (o) { sh.getRange(o.row, EMP_VIEW.RESULT_COL).setValue('❌ ' + o.msg); });
  }
  const okMsgs = outcomes.filter(function (o) { return o.ok; }).map(function (o) { return o.cycle.taskName + ': ' + o.msg; });
  const badMsgs = failures.map(function (o) { return '❌ ' + o.msg; });
  toast_(ctx, okMsgs.concat(badMsgs).join('  ·  '));
  return { handled: true, kind: 'COMPLETION', outcomes: outcomes };
}

/** After a re-render, put a rejected entry back on its (moved) row with the reason, so nothing typed is lost. */
function annotateEmployeeRows_(sh, failures) {
  if (!sh || !failures.length) return;
  const last = sh.getLastRow();
  if (last < EMP_VIEW.FIRST_ROW) return;
  const ids = sh.getRange(EMP_VIEW.FIRST_ROW, 1, last - EMP_VIEW.FIRST_ROW + 1, 1).getValues().map(function (r) { return String(r[0]); });
  failures.forEach(function (o) {
    const i = o.id ? ids.indexOf(o.id) : -1;
    if (i < 0) return;
    const r = EMP_VIEW.FIRST_ROW + i;
    sh.getRange(r, EMP_VIEW.INPUT_COL, 1, 2).setValues([[textCell_(String(o.input)), '❌ ' + o.msg]]);
  });
}

function sectionAt_(colA, row) {
  const titles = [DASH.TITLE_OVERVIEW, DASH.TITLE_ATTENTION, DASH.TITLE_OPEN, DASH.TITLE_RECENT];
  for (let r = row; r >= 1; r--) {
    const v = String(colA[r - 1] || '');
    for (let i = 0; i < titles.length; i++) if (v.indexOf(titles[i]) === 0) return titles[i];
  }
  return '';
}

function handleDashboardEdit_(ctx, e, sh) {
  const r0 = e.range.getRow();
  const c0 = e.range.getColumn();
  const nr = e.range.getNumRows();
  const nc = e.range.getNumColumns();
  const touches = function (col) { return c0 <= col && col <= c0 + nc - 1; };
  if (!touches(DASH.ACK_COL) && !touches(DASH.NOTE_COL) && !touches(DASH.ACTION_COL)) return { handled: false, reason: 'not an editable column' };

  const titles = sh.getRange(1, TITLE_COL, r0 + nr - 1, 1).getValues().map(function (r) { return r[0]; });
  const block = sh.getRange(r0, 1, nr, DASH.WIDTH).getValues().map(function (row) { return row.map(fromCell_); });
  const m = loadModel_(ctx);
  const actor = actorFromEvent_(e, null) || 'Manager (dashboard)';
  const outcomes = [];
  const touched = [];
  let changed = false;

  block.forEach(function (row, i) {
    if (sectionAt_(titles, r0 + i) !== DASH.TITLE_ATTENTION) return;
    const c = m.byId[String(row[0] || '').trim()];
    if (!c) return;
    touched.push(c.employee);
    const note = row[DASH.NOTE_COL - 1];
    if (touches(DASH.NOTE_COL) && !touches(DASH.ACK_COL)) {
      if (saveNote_(m, c, actor, note)) changed = true;
    }
    if (touches(DASH.ACK_COL)) {
      if (row[DASH.ACK_COL - 1] === true) {
        const res = ackCycle_(m, c, actor, note);
        outcomes.push(Object.assign({ id: c.id, action: 'ACK' }, res));
        if (res.ok) changed = true;
      } else if (c.ackAt) {
        outcomes.push({ id: c.id, action: 'UNACK', ok: false, msg: 'An acknowledgement is a record and cannot be withdrawn.' });
        changed = true; // re-render restores the tick
      } else {
        changed = true; // stray edit on a read-only cell: re-render restores it
      }
    }
    if (touches(DASH.ACTION_COL) && String(row[DASH.ACTION_COL - 1]).trim() === ACTION_WAIVE) {
      const res = waiveCycle_(m, c, actor, note);
      outcomes.push(Object.assign({ id: c.id, action: 'WAIVE' }, res));
      changed = true;
    }
  });
  if (!changed && !outcomes.length) return { handled: false, reason: 'no actionable rows' };
  saveModel_(ctx, m);
  renderAll_(ctx, m, uniq_(touched));
  toast_(ctx, outcomes.map(function (o) { return (o.ok ? '✅ ' : '❌ ') + o.id + ': ' + o.msg; }).join('  ·  ') || 'Saved.');
  return { handled: true, kind: 'DASHBOARD', outcomes: outcomes };
}
