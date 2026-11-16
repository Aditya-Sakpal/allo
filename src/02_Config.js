/**
 * 02_Config — typed configuration with validation, the mailer, and the execution context.
 * Bad values never crash the engine: they fall back to defaults and are reported as warnings.
 */

/** [key, default, description] */
const CONFIG_SPEC = [
  ['MANAGER_EMAIL', 'manager@example.com', 'Default manager who receives the same-day pending alert (a Team row can override it).'],
  ['LEADERSHIP_EMAILS', 'leadership@example.com', 'Comma-separated leadership recipients for escalations.'],
  ['EMAIL_MODE', 'LOG_ONLY', 'LIVE = send real emails. LOG_ONLY = only write them to the Email Log (safe for demos).'],
  ['WORKING_DAYS', 'Mon,Tue,Wed,Thu,Fri,Sat', 'Days on which daily tasks run and SLA deadlines count.'],
  ['HOLIDAYS', '', 'Comma-separated YYYY-MM-DD dates treated as non-working (no daily task, SLA clock paused), e.g. 2026-12-25, 2027-01-26.'],
  ['DEFAULT_DUE_TIME', '18:00', 'Cut-off time (HH:mm) for tasks without their own Due Time.'],
  ['REMINDER_LEAD_MINUTES', '120', 'One reminder to the employee this many minutes before the cut-off. 0 = off.'],
  ['MANAGER_ALERT_GRACE_MINUTES', '30', 'Manager is alerted this many minutes after the cut-off (same day).'],
  ['ACK_DEADLINE_TIME', '11:00', 'Manager must acknowledge by this time on the NEXT working day, otherwise leadership is informed.'],
  ['RESOLVE_GRACE_WORKING_DAYS', '1', 'Working days after the cut-off by which the task must be done, otherwise leadership is informed.'],
  ['TICKET_ID_PATTERN', '^(?=.*[0-9])[A-Za-z0-9][A-Za-z0-9#/_.-]{2,39}$', 'Regular expression a Ticket ID must match.'],
  ['DAILY_UPDATE_MIN_CHARS', '20', 'Minimum length of the daily update text.'],
  ['RECENT_DAYS', '7', 'Days of history shown on the views.'],
  ['SIMULATED_NOW', '', 'DEMO ONLY: pretend the time is this (YYYY-MM-DD HH:mm). Leave BLANK in production.'],
  ['WEBAPP_URL', '', 'Web-app URL used for one-click Acknowledge links (auto-detected once deployed).'],
];

function configSpecByKey_() {
  const o = {};
  CONFIG_SPEC.forEach(function (s) { o[s[0]] = s; });
  return o;
}

function readConfigRaw_(ss) {
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  const raw = {};
  if (!sh || sh.getLastRow() < 2) return raw;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().map(function (r) { return r.map(fromCell_); });
  vals.forEach(function (r) {
    const k = String(r[0]).trim();
    if (k) raw[k] = String(r[1]);
  });
  return raw;
}

function readConfig_(ss) {
  return parseConfig_(readConfigRaw_(ss));
}

function parseWorkingDays_(s) {
  const tokens = String(s == null ? '' : s).split(/[,\s]+/).filter(Boolean);
  if (!tokens.length) return null;
  const set = new Set();
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i].slice(0, 3).toLowerCase();
    const idx = DOW_NAMES.findIndex(function (d) { return d.toLowerCase() === t; });
    if (idx < 0) return null;
    set.add(idx);
  }
  return set;
}

/** raw: {KEY: string}. Returns a typed cfg with cfg.errors (human-readable warnings). */
function parseConfig_(raw) {
  const spec = configSpecByKey_();
  const errors = [];
  const val = function (key) {
    const v = raw[key];
    return v === undefined || v === null ? String(spec[key][1]) : String(v).trim();
  };
  const def = function (key) { return String(spec[key][1]); };
  const intOr = function (key, min, max) {
    const s = val(key);
    const n = Number(s);
    if (s === '' || !isFinite(n) || Math.floor(n) !== n || n < min || n > max) {
      errors.push(key + ' "' + s + '" must be a whole number between ' + min + ' and ' + max + ' — using ' + def(key));
      return Number(def(key));
    }
    return n;
  };
  const timeOr = function (key) {
    const hm = parseHM_(val(key));
    if (!hm) {
      errors.push(key + ' "' + val(key) + '" must be HH:mm — using ' + def(key));
      return parseHM_(def(key));
    }
    return hm;
  };

  const cfg = { errors: errors, raw: raw };

  cfg.managerEmail = val('MANAGER_EMAIL').toLowerCase();
  if (!isEmail_(cfg.managerEmail)) {
    errors.push('MANAGER_EMAIL "' + cfg.managerEmail + '" is not a valid email — manager alerts cannot be sent');
    cfg.managerEmail = '';
  }

  const lead = splitEmails_(val('LEADERSHIP_EMAILS')).map(function (e) { return e.toLowerCase(); });
  const badLead = lead.filter(function (e) { return !isEmail_(e); });
  if (badLead.length) errors.push('LEADERSHIP_EMAILS has invalid address(es): ' + badLead.join(', '));
  cfg.leadershipEmails = uniq_(lead.filter(isEmail_));
  if (!cfg.leadershipEmails.length) errors.push('LEADERSHIP_EMAILS is empty — escalations cannot be sent');

  const mode = val('EMAIL_MODE').toUpperCase();
  if (mode === 'LIVE' || mode === 'LOG_ONLY') {
    cfg.emailMode = mode;
  } else {
    errors.push('EMAIL_MODE "' + mode + '" must be LIVE or LOG_ONLY — using LOG_ONLY');
    cfg.emailMode = 'LOG_ONLY';
  }

  cfg.workingDays = parseWorkingDays_(val('WORKING_DAYS'));
  if (!cfg.workingDays) {
    errors.push('WORKING_DAYS "' + val('WORKING_DAYS') + '" not understood (use e.g. Mon,Tue,Wed,Thu,Fri) — using ' + def('WORKING_DAYS'));
    cfg.workingDays = parseWorkingDays_(def('WORKING_DAYS'));
  }

  cfg.holidays = new Set();
  const badHolidays = [];
  splitEmails_(val('HOLIDAYS')).forEach(function (s) {
    const d = /^\d{4}-\d{1,2}-\d{1,2}$/.test(s) ? parseDateTime_(s) : null;
    if (d) cfg.holidays.add(ymd_(d));
    else badHolidays.push(s);
  });
  if (badHolidays.length) errors.push('HOLIDAYS has invalid date(s) (use YYYY-MM-DD): ' + badHolidays.join(', '));

  cfg.defaultDueHM = timeOr('DEFAULT_DUE_TIME');
  cfg.reminderLeadMin = intOr('REMINDER_LEAD_MINUTES', 0, 1440);
  cfg.graceMin = intOr('MANAGER_ALERT_GRACE_MINUTES', 0, 720);
  cfg.ackHM = timeOr('ACK_DEADLINE_TIME');
  cfg.resolveGraceDays = intOr('RESOLVE_GRACE_WORKING_DAYS', 0, 30);
  cfg.dailyMinChars = intOr('DAILY_UPDATE_MIN_CHARS', 1, 1000);
  cfg.recentDays = intOr('RECENT_DAYS', 1, 90);

  try {
    cfg.ticketRe = new RegExp(val('TICKET_ID_PATTERN'));
  } catch (e) {
    errors.push('TICKET_ID_PATTERN is not a valid regular expression — using the default');
    cfg.ticketRe = new RegExp(def('TICKET_ID_PATTERN'));
  }

  const simRaw = val('SIMULATED_NOW');
  cfg.simulatedNow = null;
  if (simRaw) {
    cfg.simulatedNow = parseDateTime_(simRaw);
    if (!cfg.simulatedNow) errors.push('SIMULATED_NOW "' + simRaw + '" must be YYYY-MM-DD HH:mm — using real time');
  }

  cfg.webAppUrl = val('WEBAPP_URL');
  if (cfg.webAppUrl && !/^https:\/\//.test(cfg.webAppUrl)) {
    errors.push('WEBAPP_URL must start with https:// — acknowledgement links disabled');
    cfg.webAppUrl = '';
  }
  return cfg;
}

/** LIVE → MailApp; LOG_ONLY → nothing leaves the sheet. Tests inject their own capture mailer. */
function makeMailer_(cfg) {
  if (cfg.emailMode === 'LIVE') {
    return {
      mode: 'LIVE',
      send: function (msg) {
        const needed = splitEmails_(msg.to).length + splitEmails_(msg.cc).length;
        const left = MailApp.getRemainingDailyQuota();
        if (left < needed) throw new Error('Daily email quota exhausted (' + left + ' left, ' + needed + ' needed)');
        const opts = { to: msg.to, subject: msg.subject, htmlBody: msg.htmlBody, body: msg.body, name: APP_NAME };
        if (msg.cc) opts.cc = msg.cc;
        MailApp.sendEmail(opts);
        return { delivery: 'SENT' };
      },
    };
  }
  return { mode: 'LOG_ONLY', send: function () { return { delivery: 'LOG_ONLY' }; } };
}

/**
 * The execution context every core function receives.
 * opts: { cfg?, now?: () => Date, mailer?, webAppUrl? }
 */
function makeContext_(ss, opts) {
  opts = opts || {};
  const cfg = opts.cfg || readConfig_(ss);
  const now = opts.now || function () {
    return cfg.simulatedNow ? new Date(cfg.simulatedNow.getTime()) : new Date();
  };
  return {
    ss: ss,
    cfg: cfg,
    now: now,
    mailer: opts.mailer || makeMailer_(cfg),
    webAppUrl: opts.webAppUrl || '',
  };
}
