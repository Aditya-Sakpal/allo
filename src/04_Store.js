/**
 * 04_Store — table I/O (by header name, batched) and the in-memory model.
 * The Cycles sheet is the ledger and the single source of truth; everything else is derived.
 */

function requireSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run ⚙️ Task System → Admin → Setup / repair sheets.');
  return sh;
}

/** Grow the sheet so rows 1..n exist (getRange beyond max rows throws in Apps Script). */
function ensureRows_(sh, n) {
  const max = sh.getMaxRows();
  if (max < n) sh.insertRowsAfter(max, n - max);
}

function ensureCols_(sh, n) {
  const max = sh.getMaxColumns();
  if (max < n) sh.insertColumnsAfter(max, n - max);
}

/**
 * Reads a header-row table. Rows that are entirely empty (or only unticked checkboxes) are skipped.
 * Each row object: {_row, <key>: raw value, _d: {<key>: display value}}.
 */
function readTable_(sheet, fields) {
  const res = { rows: [], missing: [], idx: {}, lastRow: sheet.getLastRow() };
  const lastRow = res.lastRow;
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    res.missing = fields.map(function (f) { return f[1]; });
    return res;
  }
  const rng = sheet.getRange(1, 1, lastRow, lastCol);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();
  const header = vals[0].map(function (h) { return String(h).trim(); });
  fields.forEach(function (f) {
    const i = header.indexOf(f[1]);
    if (i < 0) res.missing.push(f[1]);
    else res.idx[f[0]] = i;
  });
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const empty = row.every(function (v) { return v === '' || v === null || v === false; });
    if (empty) continue;
    const o = { _row: r + 1, _d: {} };
    fields.forEach(function (f) {
      const i = res.idx[f[0]];
      o[f[0]] = i === undefined ? '' : row[i];
      o._d[f[0]] = i === undefined ? '' : disp[r][i];
    });
    res.rows.push(o);
  }
  return res;
}

function appendRows_(sh, rows, width) {
  if (!rows.length) return;
  const start = sh.getLastRow() + 1;
  ensureRows_(sh, start + rows.length - 1);
  const data = rows.map(function (r) {
    const out = [];
    for (let i = 0; i < width; i++) out.push(r[i] === undefined || r[i] === null ? '' : r[i]);
    return out;
  });
  sh.getRange(start, 1, data.length, width).setValues(data);
}

function toDateOrNull_(v) {
  if (isDate_(v)) return v;
  if (v === '' || v === null || v === undefined) return null;
  return parseDateTime_(v);
}

function normalizeCycle_(r) {
  const s = function (v) { return v === null || v === undefined ? '' : String(v).trim(); };
  const c = { _row: r._row };
  CYCLE_FIELDS.forEach(function (f) {
    const k = f[0];
    c[k] = CYCLE_DATE_KEYS.indexOf(k) >= 0 ? toDateOrNull_(r[k]) : s(r[k]);
  });
  c.status = c.status.toUpperCase();
  c.proof = r.proof === null || r.proof === undefined ? '' : String(r.proof).trim();
  const okStatus = Object.keys(STATUS).some(function (k) { return STATUS[k] === c.status; });
  if (!c.id || !c.taskId || !c.slot || !c.dueAt || !c.opensAt || !okStatus) {
    // Keep the row exactly as found so a manual mistake is never destroyed by a rewrite.
    c._bad = true;
    c._orig = {};
    CYCLE_FIELDS.forEach(function (f) { c._orig[f[0]] = r[f[0]] === undefined ? '' : r[f[0]]; });
  }
  return c;
}

function cycleToRow_(c) {
  if (c._bad) return CYCLE_FIELDS.map(function (f) { return c._orig[f[0]]; });
  return CYCLE_FIELDS.map(function (f) {
    const v = c[f[0]];
    return v === null || v === undefined ? '' : v;
  });
}

function log_(m, event, cycleId, actor, details) {
  m.logs.push([m.now, event, cycleId || '', actor || 'system', details || '']);
}

/** Loads Team, Tasks and Cycles into a model. Validation of tasks happens in 05_Cycles. */
function loadModel_(ctx) {
  const ss = ctx.ss;
  const cfg = ctx.cfg;
  const m = {
    now: ctx.now(),
    cfg: cfg,
    logs: [],
    warnings: cfg.errors.map(function (e) { return 'Config: ' + e; }),
    healthComputed: false,
  };

  const teamT = readTable_(requireSheet_(ss, SHEETS.TEAM), TEAM_FIELDS);
  if (teamT.missing.indexOf('Employee') >= 0) throw new Error('Team sheet has no "Employee" column. Run Setup.');
  m.team = [];
  m.teamByName = {};
  teamT.rows.forEach(function (r) {
    const name = normalizeSpace_(r._d.name);
    if (!name) return;
    if (m.teamByName[name]) { m.warnings.push('Team: "' + name + '" is listed twice — row ' + r._row + ' ignored'); return; }
    const e = {
      _row: r._row,
      name: name,
      email: String(r._d.email || '').trim().toLowerCase(),
      managerEmail: String(r._d.managerEmail || '').trim().toLowerCase(),
      active: teamT.idx.active === undefined ? true : toBool_(r.active),
    };
    if (e.email && !isEmail_(e.email)) { m.warnings.push('Team: email for ' + name + ' is invalid'); e.email = ''; }
    if (e.managerEmail && !isEmail_(e.managerEmail)) { m.warnings.push('Team: manager email for ' + name + ' is invalid — using default manager'); e.managerEmail = ''; }
    m.team.push(e);
    m.teamByName[name] = e;
  });

  const taskT = readTable_(requireSheet_(ss, SHEETS.TASKS), TASK_FIELDS);
  const reqTask = ['Task ID', 'Task Name', 'Assigned To', 'Frequency', 'Proof Type', 'First Due Date'];
  const missingTask = reqTask.filter(function (h) { return taskT.missing.indexOf(h) >= 0; });
  if (missingTask.length) throw new Error('Tasks sheet is missing column(s): ' + missingTask.join(', ') + '. Run Setup.');
  m.taskIdx = taskT.idx;
  m.taskLastRow = taskT.lastRow;
  m.tasks = taskT.rows.map(function (r) {
    return {
      _row: r._row,
      id: normalizeSpace_(r._d.id),
      name: normalizeSpace_(r._d.name),
      assignedTo: normalizeSpace_(r._d.assignedTo),
      frequency: normalizeSpace_(r._d.frequency),
      proofType: normalizeSpace_(r._d.proofType),
      instructions: String(r._d.instructions || '').trim(),
      dueTimeRaw: String(r._d.dueTime || '').trim(),
      firstDueRaw: r.firstDue,
      firstDueDisp: String(r._d.firstDue || '').trim(),
      active: taskT.idx.active === undefined ? true : toBool_(r.active),
    };
  });

  const cycT = readTable_(requireSheet_(ss, SHEETS.CYCLES), CYCLE_FIELDS);
  if (cycT.missing.length && cycT.lastRow > 1) {
    throw new Error('Cycles sheet is missing column(s): ' + cycT.missing.join(', ') + '. Run Setup.');
  }
  m.cycles = [];
  m.byId = {};
  cycT.rows.forEach(function (r) {
    const c = normalizeCycle_(r);
    if (c._bad) {
      m.warnings.push('Cycles row ' + r._row + ' is incomplete or has an unknown status — it is kept but ignored');
    } else if (m.byId[c.id]) {
      c._bad = true;
      c._orig = {};
      CYCLE_FIELDS.forEach(function (f) { c._orig[f[0]] = r[f[0]] === undefined ? '' : r[f[0]]; });
      m.warnings.push('Cycles row ' + r._row + ' duplicates Cycle ID ' + c.id + ' — it is kept but ignored');
    } else {
      m.byId[c.id] = c;
    }
    m.cycles.push(c);
  });

  validateTasks_(m);
  return m;
}

/** Active (non-corrupt) cycles only. */
function liveCycles_(m) {
  return m.cycles.filter(function (c) { return !c._bad; });
}

function saveCycles_(ctx, m) {
  const sh = requireSheet_(ctx.ss, SHEETS.CYCLES);
  const rows = [CYCLE_FIELDS.map(function (f) { return f[1]; })].concat(m.cycles.map(cycleToRow_));
  ensureRows_(sh, rows.length);
  ensureCols_(sh, CYCLE_FIELDS.length);
  sh.getRange(1, 1, rows.length, CYCLE_FIELDS.length).setValues(rows);
  const last = sh.getLastRow();
  if (last > rows.length) {
    sh.getRange(rows.length + 1, 1, last - rows.length, Math.max(sh.getLastColumn(), CYCLE_FIELDS.length)).clearContent();
  }
}

function flushLogs_(ctx, m) {
  if (!m.logs.length) return;
  appendRows_(requireSheet_(ctx.ss, SHEETS.LOG), m.logs, LOG_HEADERS.length);
  m.logs = [];
}

function appendEmailLog_(ctx, row) {
  appendRows_(requireSheet_(ctx.ss, SHEETS.EMAILS), [row], EMAIL_HEADERS.length);
}

function writeTaskHealth_(ctx, m) {
  if (!m.healthComputed || m.taskIdx.health === undefined || m.taskLastRow < 2) return;
  const sh = requireSheet_(ctx.ss, SHEETS.TASKS);
  const vals = [];
  for (let r = 2; r <= m.taskLastRow; r++) vals.push(['']);
  m.tasks.forEach(function (t) { vals[t._row - 2] = [t.health || '']; });
  sh.getRange(2, m.taskIdx.health + 1, vals.length, 1).setValues(vals);
}

function saveModel_(ctx, m) {
  saveCycles_(ctx, m);
  writeTaskHealth_(ctx, m);
  flushLogs_(ctx, m);
}

function managerFor_(m, employeeName) {
  const e = m.teamByName[employeeName];
  return (e && e.managerEmail) || m.cfg.managerEmail || '';
}
