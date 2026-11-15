/**
 * 05_Cycles — cycle lifecycle: task validation, cycle creation, completion, acknowledgement, waive.
 * Spec: F01 (engine), F03 (ticket), F04 (daily), F05 (tracking), F06 (ack / waive).
 * Every function here mutates the in-memory model only; callers persist with saveModel_.
 */

const MAX_PROOF_CHARS = 4000;
const MAX_SKIPPED_ROWS = 60;

function statusLabel_(s) {
  return {
    OPEN: 'Open', DONE: 'Done on time', LATE_DONE: 'Done late', SKIPPED: 'Skipped (missed)', WAIVED: 'Waived',
  }[s] || s;
}

function validateTasks_(m) {
  const seen = {};
  m.taskById = {};
  m.tasks.forEach(function (t) {
    const errs = [];
    if (!t.id) errs.push('Task ID is blank');
    else if (!/^[A-Za-z0-9_.-]{1,30}$/.test(t.id)) errs.push('Task ID may only use letters, digits, - _ . (max 30)');
    else if (seen[t.id]) errs.push('Duplicate Task ID (first used on row ' + seen[t.id] + ')');
    if (t.id && !seen[t.id]) seen[t.id] = t._row;
    if (!t.name) errs.push('Task Name is blank');
    const emp = m.teamByName[t.assignedTo];
    if (!t.assignedTo) errs.push('Assigned To is blank');
    else if (!emp) errs.push('"' + t.assignedTo + '" is not on the Team sheet');
    else if (!emp.active) errs.push(t.assignedTo + ' is inactive — reassign this task');
    t.freq = parseFrequency_(t.frequency);
    if (!t.freq) errs.push('Frequency "' + t.frequency + '" not understood (use Daily or Every N days)');
    if (t.proofType !== PROOF.TICKET && t.proofType !== PROOF.DAILY) errs.push('Proof Type must be "Ticket ID" or "Daily Update"');
    t.dueHM = t.dueTimeRaw ? parseHM_(t.dueTimeRaw) : m.cfg.defaultDueHM;
    if (!t.dueHM) errs.push('Due Time "' + t.dueTimeRaw + '" must be HH:mm');
    const fd = isDate_(t.firstDueRaw) ? t.firstDueRaw : parseDateTime_(t.firstDueDisp);
    t.firstDue = fd ? startOfDay_(fd) : null;
    if (!t.firstDue) errs.push('First Due Date missing or invalid (use YYYY-MM-DD)');
    t.errors = errs;
    t.valid = errs.length === 0;
    // The engine runs every 15 minutes, so the alert can only be guaranteed the same day
    // if cut-off + grace is at or before 23:44.
    t.note = '';
    if (t.dueHM && t.dueHM.h * 60 + t.dueHM.m + m.cfg.graceMin > 23 * 60 + 44) {
      t.note = ' · ⚠️ cut-off is so late that the manager alert may arrive after midnight';
    }
    if (t.id && !m.taskById[t.id]) m.taskById[t.id] = t;
  });
}

function cyclesOfTask_(m, taskId) {
  return liveCycles_(m).filter(function (c) { return c.taskId === taskId; });
}

function openCyclesOfTask_(m, taskId) {
  return cyclesOfTask_(m, taskId).filter(function (c) { return c.status === STATUS.OPEN; });
}

function newCycle_(m, t, slot, status) {
  const s = startOfDay_(slot);
  return {
    id: cycleId_(t.id, s),
    taskId: t.id,
    taskName: t.name,
    employee: t.assignedTo,
    frequency: t.freq.label,
    proofType: t.proofType,
    slot: s,
    opensAt: windowOpensAt_(s, t.freq, t.dueHM, m.cfg),
    dueAt: slotDueAt_(s, t.freq, t.dueHM, m.cfg),
    status: status,
    completedAt: null, completedBy: '', proof: '',
    reminderAt: null, alertAt: null, ackAt: null, ackBy: '', ackNote: '',
    escalatedAt: null, escalationReason: '', resolvedNoticeAt: null,
    closeNote: '',
    updatedAt: m.now,
  };
}

function addCycle_(m, c, event, details) {
  m.cycles.push(c);
  m.byId[c.id] = c;
  log_(m, event, c.id, 'system', details);
}

function closeCycle_(m, c, status, note, actor) {
  c.status = status;
  c.closeNote = note;
  c.updatedAt = m.now;
  log_(m, status === STATUS.WAIVED ? 'WAIVED' : 'CLOSED', c.id, actor, note);
}

/** Keeps the denormalised copy on OPEN cycles in line with the Task Master. */
function syncCycleFromTask_(m, c, t) {
  if (c.employee !== t.assignedTo) {
    log_(m, 'REASSIGNED', c.id, 'system', c.employee + ' → ' + t.assignedTo);
    c.employee = t.assignedTo;
    c.updatedAt = m.now;
  }
  c.taskName = t.name;
  c.proofType = t.proofType;
}

/** Idempotent: safe to call any number of times. Creates only what is missing. */
function ensureCycles_(m) {
  m.tasks.forEach(function (t) {
    if (!t.id || m.taskById[t.id] !== t) {
      t.health = '⚠️ ' + (t.errors.length ? t.errors.join('; ') : 'Duplicate row');
      return;
    }
    const open = openCyclesOf_(m, t);
    if (!t.active) {
      open.forEach(function (c) { closeCycle_(m, c, STATUS.WAIVED, 'Task deactivated in Tasks sheet', 'system'); });
      t.health = '⏸ Inactive';
      return;
    }
    if (!t.valid) {
      t.health = '⚠️ ' + t.errors.join('; ');
      return;
    }
    open.forEach(function (c) { syncCycleFromTask_(m, c, t); });
    if (t.freq.kind === 'DAILY') ensureDaily_(m, t);
    else ensureInterval_(m, t);
    const nowOpen = openCyclesOf_(m, t).sort(byDue_);
    t.health = nowOpen.length
      ? '✅ OK · open cycle due ' + fmtShort_(nowOpen[0].dueAt) + (nowOpen.length > 1 ? ' (+' + (nowOpen.length - 1) + ' older open)' : '')
      : '✅ OK · next cycle is created automatically';
    t.health += t.note;
  });
  // Open cycles whose task row was deleted would otherwise alert forever.
  liveCycles_(m).forEach(function (c) {
    if (c.status === STATUS.OPEN && !m.taskById[c.taskId]) {
      closeCycle_(m, c, STATUS.WAIVED, 'Task ' + c.taskId + ' no longer exists in the Tasks sheet', 'system');
    }
  });
  m.healthComputed = true;
}

function openCyclesOf_(m, t) {
  return openCyclesOfTask_(m, t.id);
}

function ensureDaily_(m, t) {
  const today = startOfDay_(m.now);
  if (!isWorkingDay_(today, m.cfg)) return;
  if (t.firstDue.getTime() > today.getTime()) return;
  const id = cycleId_(t.id, today);
  if (m.byId[id]) return;
  const history = cyclesOfTask_(m, t.id);
  // A brand-new daily task added after today's cut-off starts tomorrow instead of being born overdue.
  if (!history.length && slotDueAt_(today, t.freq, t.dueHM, m.cfg).getTime() <= m.now.getTime()) return;
  addCycle_(m, newCycle_(m, t, today, STATUS.OPEN), 'CYCLE_CREATED', 'Daily cycle for ' + fmt_(today, 'EEE d MMM'));
}

function ensureInterval_(m, t) {
  if (openCyclesOf_(m, t).length) return;
  const history = cyclesOfTask_(m, t.id);
  let slot;
  if (!history.length) {
    slot = t.firstDue;
  } else {
    const last = history.reduce(function (a, c) { return c.slot.getTime() > a.getTime() ? c.slot : a; }, history[0].slot);
    slot = nextSlot_(last, t.freq, m.cfg);
  }
  let rolled = 0;
  while (slotDueAt_(slot, t.freq, t.dueHM, m.cfg).getTime() <= m.now.getTime() && rolled < 5000) {
    slot = nextSlot_(slot, t.freq, m.cfg);
    rolled++;
  }
  const id = cycleId_(t.id, slot);
  if (m.byId[id]) return;
  addCycle_(m, newCycle_(m, t, slot, STATUS.OPEN), 'CYCLE_CREATED',
    rolled ? 'Scheduled (rolled forward ' + rolled + ' period(s) that had already passed)' : 'Scheduled');
}

/** After an interval cycle closes: record slots that already passed as SKIPPED, open the next future slot. */
function scheduleNextInterval_(m, t, closed) {
  if (openCyclesOf_(m, t).length) return null;
  let slot = nextSlot_(closed.slot, t.freq, m.cfg);
  let skipped = 0;
  let guard = 0;
  while (slotDueAt_(slot, t.freq, t.dueHM, m.cfg).getTime() <= m.now.getTime() && guard++ < 5000) {
    const id = cycleId_(t.id, slot);
    if (!m.byId[id] && skipped < MAX_SKIPPED_ROWS) {
      const s = newCycle_(m, t, slot, STATUS.SKIPPED);
      s.closeNote = 'Cut-off passed while ' + closed.id + ' was still open';
      addCycle_(m, s, 'CYCLE_SKIPPED', s.closeNote);
      skipped++;
    }
    slot = nextSlot_(slot, t.freq, m.cfg);
  }
  const id = cycleId_(t.id, slot);
  if (m.byId[id]) return m.byId[id];
  const n = newCycle_(m, t, slot, STATUS.OPEN);
  addCycle_(m, n, 'CYCLE_CREATED', 'Next cycle after ' + closed.id);
  return n;
}

function validateProof_(m, c, raw) {
  const isDaily = c.proofType === PROOF.DAILY;
  const value = isDaily ? String(raw == null ? '' : raw).trim() : normalizeSpace_(raw);
  if (!value) return { ok: false, code: 'EMPTY', msg: 'Nothing was entered.' };
  if (value.length > MAX_PROOF_CHARS) {
    return { ok: false, code: 'TOO_LONG', msg: 'That is ' + value.length + ' characters — please keep it under ' + MAX_PROOF_CHARS + '.' };
  }
  const others = liveCycles_(m).filter(function (x) { return x !== c && x.proof; });
  if (!isDaily) {
    if (!m.cfg.ticketRe.test(value)) {
      return {
        ok: false, code: 'BAD_TICKET',
        msg: '"' + truncate_(value, 40) + '" is not a valid Ticket ID. Use the portal ticket number, e.g. PRT-10234 (3–40 letters/digits, at least one digit, no spaces).',
      };
    }
    const dup = others.find(function (x) { return x.proofType === PROOF.TICKET && x.proof.toUpperCase() === value.toUpperCase(); });
    if (dup) {
      return {
        ok: false, code: 'DUP_TICKET',
        msg: 'Ticket ' + value + ' was already used for ' + dup.taskName + ' (' + fmt_(dup.slot, 'd MMM') + '). Enter the ticket for this cycle.',
      };
    }
  } else {
    if (value.length < m.cfg.dailyMinChars) {
      return {
        ok: false, code: 'TOO_SHORT',
        msg: 'Paste the actual update you shared (at least ' + m.cfg.dailyMinChars + ' characters). "' + truncate_(value, 30) + '" is too short.',
      };
    }
    const norm = normalizeSpace_(value).toLowerCase();
    const dup = others.find(function (x) { return x.taskId === c.taskId && normalizeSpace_(x.proof).toLowerCase() === norm; });
    if (dup) {
      return {
        ok: false, code: 'DUP_UPDATE',
        msg: 'This is identical to the update for ' + fmt_(dup.slot, 'EEE d MMM') + '. Please paste today\'s update.',
      };
    }
  }
  return { ok: true, value: value };
}

/** Closes an OPEN cycle with proof. Returns {ok, code?, msg, next?}. */
function completeCycle_(m, c, raw, actor) {
  if (!c || c._bad) return { ok: false, code: 'UNKNOWN', msg: 'This row is out of date — the sheet will refresh in a moment.' };
  if (c.status !== STATUS.OPEN) {
    return {
      ok: false, code: 'NOT_OPEN',
      msg: 'Already ' + statusLabel_(c.status).toLowerCase() + (c.completedAt ? ' (' + fmtShort_(c.completedAt) + ')' : '') + ' — nothing to do.',
    };
  }
  if (m.now.getTime() < c.opensAt.getTime()) {
    return { ok: false, code: 'NOT_YET', msg: 'This cycle opens on ' + fmtShort_(c.opensAt) + ' — it cannot be closed early.' };
  }
  const v = validateProof_(m, c, raw);
  if (!v.ok) {
    log_(m, 'PROOF_REJECTED', c.id, actor, v.code + ': ' + truncate_(String(raw == null ? '' : raw), 80));
    return v;
  }
  const late = m.now.getTime() > c.dueAt.getTime();
  c.status = late ? STATUS.LATE : STATUS.DONE;
  c.completedAt = m.now;
  c.completedBy = actor;
  c.proof = v.value;
  c.updatedAt = m.now;
  log_(m, 'COMPLETED', c.id, actor, (late ? 'Late by ' + durationText_(m.now - c.dueAt) + ' · ' : 'On time · ') + truncate_(v.value, 80));

  let next = null;
  const t = m.taskById[c.taskId];
  if (t && t.valid && t.active && t.freq.kind === 'INTERVAL') next = scheduleNextInterval_(m, t, c);
  return {
    ok: true,
    cycle: c,
    next: next,
    msg: (late ? '✅ Recorded (late)' : '✅ Recorded on time') + (next ? ' · next cycle due ' + fmtShort_(next.dueAt) : ''),
  };
}

function canAcknowledge_(m, c) {
  return !!c.alertAt || (c.status === STATUS.OPEN && m.now.getTime() > c.dueAt.getTime());
}

/** Idempotent. A later note updates the note but never moves the original acknowledgement time. */
function ackCycle_(m, c, actor, note) {
  if (!c || c._bad) return { ok: false, msg: 'Unknown task cycle.' };
  note = normalizeSpace_(note).slice(0, 500);
  if (c.ackAt) {
    if (note && note !== c.ackNote) {
      c.ackNote = note;
      c.updatedAt = m.now;
      log_(m, 'ACK_NOTE', c.id, actor, note);
      return { ok: true, already: true, msg: 'Note saved.' };
    }
    return { ok: true, already: true, msg: 'Already acknowledged ' + fmtShort_(c.ackAt) + ' by ' + c.ackBy + '. Nothing changed.' };
  }
  if (!canAcknowledge_(m, c)) return { ok: false, msg: 'Nothing to acknowledge yet — this task is not overdue.' };
  c.ackAt = m.now;
  c.ackBy = actor;
  if (note) c.ackNote = note;
  c.updatedAt = m.now;
  log_(m, 'ACKNOWLEDGED', c.id, actor, 'Task status at acknowledgement: ' + statusLabel_(c.status) + (note ? ' · ' + note : ''));
  return {
    ok: true,
    msg: c.status === STATUS.OPEN
      ? 'Acknowledged. Please make sure it is done by ' + fmtShort_(resolveDeadline_(c.dueAt, m.cfg)) + '.'
      : 'Acknowledged. The task itself is already ' + statusLabel_(c.status).toLowerCase() + '.',
  };
}

function saveNote_(m, c, actor, note) {
  note = normalizeSpace_(note).slice(0, 500);
  if (note === c.ackNote) return false;
  c.ackNote = note;
  c.updatedAt = m.now;
  log_(m, 'ACK_NOTE', c.id, actor, note || '(cleared)');
  return true;
}

/** Manager closes an OPEN cycle without proof (leave / holiday). Counts as acknowledgement. */
function waiveCycle_(m, c, actor, note) {
  if (!c || c._bad || c.status !== STATUS.OPEN) return { ok: false, msg: 'Only open tasks can be waived.' };
  note = normalizeSpace_(note);
  closeCycle_(m, c, STATUS.WAIVED, 'Waived by ' + actor + (note ? ': ' + note : ''), actor);
  if (!c.ackAt) { c.ackAt = m.now; c.ackBy = actor; }
  const t = m.taskById[c.taskId];
  let next = null;
  if (t && t.valid && t.active && t.freq.kind === 'INTERVAL') next = scheduleNextInterval_(m, t, c);
  return { ok: true, next: next, msg: 'Waived.' + (next ? ' Next cycle due ' + fmtShort_(next.dueAt) + '.' : '') };
}
