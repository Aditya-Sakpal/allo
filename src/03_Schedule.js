/**
 * 03_Schedule — slot maths, working days and SLA deadlines. Pure functions of (date, cfg).
 * Spec: specs/features/F01-recurring-cycle-engine.md, F07-notifications-sla-escalation.md
 */

/** "Daily" | "Every N days" | "N" | "Weekly" | "Monthly" → {kind, days, label} or null. */
function parseFrequency_(s) {
  const t = normalizeSpace_(s).toLowerCase();
  if (!t) return null;
  const daily = { kind: 'DAILY', days: 1, label: 'Daily' };
  if (t === 'daily' || t === 'every day') return daily;
  if (t === 'weekly') return { kind: 'INTERVAL', days: 7, label: 'Every 7 days' };
  if (t === 'fortnightly') return { kind: 'INTERVAL', days: 14, label: 'Every 14 days' };
  if (t === 'monthly') return { kind: 'INTERVAL', days: 30, label: 'Every 30 days' };
  const m = t.match(/^(?:every\s+)?(\d{1,3})(?:\s*days?)?$/);
  if (!m) return null;
  const n = Number(m[1]);
  if (n < 1 || n > 366) return null;
  if (n === 1) return daily;
  return { kind: 'INTERVAL', days: n, label: 'Every ' + n + ' days' };
}

function isWorkingDay_(d, cfg) {
  return cfg.workingDays.has(d.getDay()) && !(cfg.holidays && cfg.holidays.has(ymd_(d)));
}

/** The same day if it is a working day, otherwise the next working day (date only). */
function onOrNextWorkingDay_(d, cfg) {
  let x = startOfDay_(d);
  // 60 days covers any realistic run of weekends + holidays; the cap prevents an infinite loop
  // if someone lists every day of a month as a holiday.
  for (let i = 0; i < 60 && !isWorkingDay_(x, cfg); i++) x = addDays_(x, 1);
  return x;
}

/** The first working day strictly after d's date. */
function nextWorkingDay_(d, cfg) {
  return onOrNextWorkingDay_(addDays_(startOfDay_(d), 1), cfg);
}

function addWorkingDays_(d, n, cfg) {
  let x = startOfDay_(d);
  for (let i = 0; i < n; i++) x = nextWorkingDay_(x, cfg);
  return x;
}

function cycleId_(taskId, slot) {
  return taskId + '-' + ymd_(slot);
}

/** Interval slots falling on a non-working day are due on the next working day. The anchor does not move. */
function slotDueAt_(slot, freq, dueHM, cfg) {
  const day = freq.kind === 'DAILY' ? startOfDay_(slot) : onOrNextWorkingDay_(slot, cfg);
  return atTime_(day, dueHM);
}

function nextSlot_(slot, freq, cfg) {
  return freq.kind === 'DAILY' ? nextWorkingDay_(slot, cfg) : addDays_(startOfDay_(slot), freq.days);
}

/** Daily: from 00:00 of the day. Interval: from the previous slot's cut-off. */
function windowOpensAt_(slot, freq, dueHM, cfg) {
  if (freq.kind === 'DAILY') return startOfDay_(slot);
  return slotDueAt_(addDays_(startOfDay_(slot), -freq.days), freq, dueHM, cfg);
}

/** Manager must acknowledge by ACK_DEADLINE_TIME on the working day after the alert. */
function ackDeadline_(alertAt, cfg) {
  return atTime_(nextWorkingDay_(alertAt, cfg), cfg.ackHM);
}

/** Task must be done by its own cut-off time, RESOLVE_GRACE_WORKING_DAYS working days after the due date. */
function resolveDeadline_(dueAt, cfg) {
  return atTime_(addWorkingDays_(dueAt, cfg.resolveGraceDays, cfg), { h: dueAt.getHours(), m: dueAt.getMinutes() });
}

function managerAlertAt_(dueAt, cfg) {
  return addMinutes_(dueAt, cfg.graceMin);
}

function reminderAt_(dueAt, cfg) {
  return addMinutes_(dueAt, -cfg.reminderLeadMin);
}
