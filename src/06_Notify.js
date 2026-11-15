/**
 * 06_Notify — reminders, same-day manager alerts, leadership escalation, resolved notices.
 * Spec: specs/features/F07-notifications-sla-escalation.md, F08-idempotency-no-duplicates.md
 *
 * One email per recipient per event type per run; each cycle is stamped once per event type.
 * Stamps are written and saved straight after a successful send, so a failure later in the
 * run can never cause a resend, and a failed send is retried on the next run.
 */

function escalationReasons_(c, m) {
  const reasons = [];
  const t = m.now.getTime();
  const ackDl = ackDeadline_(c.alertAt, m.cfg);
  if (!c.ackAt && t >= ackDl.getTime()) reasons.push('Manager did not acknowledge by ' + fmtShort_(ackDl));
  const resDl = resolveDeadline_(c.dueAt, m.cfg);
  if (t >= resDl.getTime()) reasons.push('Task still not completed by ' + fmtShort_(resDl));
  return reasons;
}

function processNotifications_(ctx, m) {
  const cfg = m.cfg;
  const now = m.now;
  const t = now.getTime();
  const reminders = {};
  const alerts = {};
  const escalations = [];
  const resolved = [];
  const push = function (map, key, item) { (map[key] = map[key] || []).push(item); };
  const noEmail = {};
  m.emailsSent = m.emailsSent || 0;

  liveCycles_(m).forEach(function (c) {
    if (c.status === STATUS.OPEN) {
      if (cfg.reminderLeadMin > 0 && !c.reminderAt && t < c.dueAt.getTime() &&
          t >= reminderAt_(c.dueAt, cfg).getTime() && t >= c.opensAt.getTime()) {
        const emp = m.teamByName[c.employee];
        if (emp && emp.email) push(reminders, emp.email, c);
        else noEmail[c.employee] = true;
      }
      if (!c.alertAt && t >= managerAlertAt_(c.dueAt, cfg).getTime()) {
        if (c.ackAt) {
          c.alertAt = now;
          c.updatedAt = now;
          log_(m, 'ALERT_SKIPPED', c.id, 'system', 'Manager had already acknowledged — no alert email needed');
        } else {
          const mgr = managerFor_(m, c.employee);
          if (mgr) push(alerts, mgr, c);
          else m.warnings.push('No manager email for ' + c.employee + ' — alert for ' + c.id + ' cannot be sent');
        }
      }
      if (c.alertAt && !c.escalatedAt) {
        const reasons = escalationReasons_(c, m);
        if (reasons.length) escalations.push({ c: c, reasons: reasons });
      }
    } else if (c.escalatedAt && !c.resolvedNoticeAt && c.status !== STATUS.SKIPPED) {
      resolved.push(c);
    }
  });
  Object.keys(noEmail).forEach(function (n) { m.warnings.push('Team: ' + n + ' has no email — reminders are skipped'); });

  Object.keys(reminders).forEach(function (to) {
    const items = reminders[to].sort(byDue_);
    deliver_(ctx, m, 'REMINDER', buildReminderEmail_(ctx, m, to, items), items,
      function (c) { c.reminderAt = now; }, 'REMINDER_SENT');
  });

  Object.keys(alerts).forEach(function (to) {
    const items = alerts[to].sort(byDue_);
    deliver_(ctx, m, 'MANAGER_ALERT', buildManagerAlertEmail_(ctx, m, to, items), items,
      function (c) { c.alertAt = now; }, 'MANAGER_ALERT_SENT');
  });

  if (escalations.length) {
    if (!cfg.leadershipEmails.length) {
      m.warnings.push(escalations.length + ' item(s) are due for leadership escalation but LEADERSHIP_EMAILS is empty');
    } else {
      const items = escalations.map(function (e) { return e.c; }).sort(byDue_);
      const reasonById = {};
      escalations.forEach(function (e) { reasonById[e.c.id] = e.reasons.join('; '); });
      deliver_(ctx, m, 'LEADERSHIP_ESCALATION', buildEscalationEmail_(ctx, m, items, reasonById), items,
        function (c) { c.escalatedAt = now; c.escalationReason = reasonById[c.id]; }, 'ESCALATED');
    }
  }

  if (resolved.length && cfg.leadershipEmails.length) {
    const items = resolved.sort(byDue_);
    deliver_(ctx, m, 'RESOLVED_NOTICE', buildResolvedEmail_(ctx, m, items), items,
      function (c) { c.resolvedNoticeAt = now; }, 'RESOLVED_NOTICE_SENT');
  }
}

function deliver_(ctx, m, type, msg, cycles, mark, event) {
  const ids = cycles.map(function (c) { return c.id; });
  msg.type = type;
  msg.cycleIds = ids;
  let res;
  try {
    res = ctx.mailer.send(msg) || {};
  } catch (err) {
    const why = String((err && err.message) || err);
    appendEmailLog_(ctx, [m.now, type, msg.to, msg.cc || '', msg.subject, ids.join(', '), 'FAILED', why]);
    cycles.forEach(function (c) { log_(m, 'EMAIL_FAILED', c.id, 'system', type + ': ' + why); });
    m.warnings.push(type + ' email to ' + msg.to + ' failed (' + why + ') — it will be retried on the next run');
    return false;
  }
  cycles.forEach(function (c) {
    mark(c);
    c.updatedAt = m.now;
    log_(m, event, c.id, 'system', 'To ' + msg.to + (msg.cc ? ' · cc ' + msg.cc : ''));
  });
  saveCycles_(ctx, m);
  appendEmailLog_(ctx, [m.now, type, msg.to, msg.cc || '', msg.subject, ids.join(', '), res.delivery || 'SENT', '']);
  m.emailsSent++;
  return true;
}

/* ───────────────────────── email builders ───────────────────────── */

function sheetLink_(ctx, name) {
  const url = ctx.ss.getUrl();
  const sh = name ? ctx.ss.getSheetByName(name) : null;
  return sh ? url + '#gid=' + sh.getSheetId() : url;
}

function managersCc_(m, items) {
  return uniq_(items.map(function (c) { return managerFor_(m, c.employee); }).filter(function (e) {
    return e && m.cfg.leadershipEmails.indexOf(e) < 0;
  })).join(', ');
}

function employeesOf_(items) {
  return uniq_(items.map(function (c) { return c.employee; })).join(', ');
}

/** o: {accent, heading, intro, headers[], rows[][] (already-escaped HTML), outro, cta:{label,url}} */
function emailHtml_(o) {
  const th = o.headers.map(function (h) {
    return '<th style="text-align:left;padding:8px 10px;font-size:12px;font-weight:600;color:#4b5563;' +
      'background:#f3f4f6;border-bottom:1px solid #e5e7eb">' + escapeHtml_(h) + '</th>';
  }).join('');
  const trs = o.rows.map(function (r) {
    return '<tr>' + r.map(function (cell) {
      return '<td style="padding:8px 10px;font-size:13px;color:#1f2933;border-bottom:1px solid #eef0f3;vertical-align:top">' + cell + '</td>';
    }).join('') + '</tr>';
  }).join('');
  // Table-cell ("bulletproof") button: Outlook ignores padding on <a>, so the padded cell carries the shape.
  const cta = o.cta && o.cta.url
    ? '<table role="presentation" cellspacing="0" cellpadding="0" style="margin:20px 0 0"><tr><td bgcolor="#1f2933" ' +
      'style="background:#1f2933;border-radius:6px;padding:12px 20px"><a href="' + escapeHtml_(o.cta.url) +
      '" style="color:#ffffff;text-decoration:none;font-size:15px;font-weight:bold;display:inline-block">' +
      escapeHtml_(o.cta.label) + '</a></td></tr></table>'
    : '';
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1f2933;max-width:680px;margin:0 auto;padding:8px">' +
    '<div style="padding-top:8px">' +
    '<h2 style="font-size:21px;line-height:1.3;margin:0 0 12px;color:' + o.accent + '">' + escapeHtml_(o.heading) + '</h2>' +
    '<p style="font-size:15px;line-height:1.5;margin:0 0 14px">' + o.intro + '</p>' +
    '<table role="presentation" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb">' +
    '<thead><tr>' + th + '</tr></thead><tbody>' + trs + '</tbody></table>' +
    cta + (o.outro || '') +
    '<p style="font-size:12px;color:#6b7280;margin-top:24px">' + escapeHtml_(APP_NAME) +
    ' · automated message · each event is emailed once, never repeated.</p>' +
    '</div></div>';
}

function textTable_(headers, rows) {
  return rows.map(function (r) {
    return '- ' + r.map(function (v, i) { return headers[i] + ': ' + v; }).join(' | ');
  }).join('\n');
}

function proofHint_(c, cfg) {
  return c.proofType === PROOF.DAILY
    ? 'Paste the update you posted on WhatsApp (min ' + cfg.dailyMinChars + ' characters)'
    : 'Portal Ticket ID (e.g. PRT-10234)';
}

function buildReminderEmail_(ctx, m, to, items) {
  const emp = items[0].employee;
  const headers = ['Task', 'Cut-off', 'What to enter'];
  const plain = items.map(function (c) { return [c.taskName, fmtShort_(c.dueAt), proofHint_(c, m.cfg)]; });
  const link = sheetLink_(ctx, employeeTabName_(emp));
  const first = items[0].dueAt;
  return {
    to: to,
    cc: '',
    subject: 'Reminder: ' + (items.length === 1 ? items[0].taskName : items.length + ' tasks') + ' due at ' + fmt_(first, 'HH:mm') + ' today',
    htmlBody: emailHtml_({
      accent: '#b45309',
      heading: 'Hi ' + emp + ', ' + (items.length === 1 ? 'one task is' : items.length + ' tasks are') + ' due at ' + fmt_(first, 'HH:mm') + ' today',
      intro: 'Finish the task, then enter the Ticket ID or your update in the “✍️ Enter Ticket ID / update here” column on your tab and press Enter. That single entry closes it. There is nothing else to update.',
      headers: headers,
      rows: plain.map(function (r) { return r.map(escapeHtml_); }),
      outro: '',
      cta: { label: 'Open my tasks', url: link },
    }),
    body: 'Hi ' + emp + ',\n\nDue today:\n' + textTable_(headers, plain) + '\n\nOpen your tab: ' + link,
  };
}

function buildManagerAlertEmail_(ctx, m, to, items) {
  const ackDl = ackDeadline_(m.now, m.cfg);
  const n = items.length;
  const single = n === 1;
  const dash = sheetLink_(ctx, SHEETS.DASHBOARD);
  const firstUrl = ackUrl_(ctx, m, items[0].id);
  // One task: the acknowledgement is the big button. Several: a per-row link in the FIRST column (thumb-reachable).
  const headers = single ? ['Employee', 'Task', 'Overdue by', 'Must be done by'] : ['Acknowledge', 'Employee', 'Task', 'Overdue by', 'Must be done by'];
  const rowsPlain = [];
  const rowsHtml = items.map(function (c) {
    const url = ackUrl_(ctx, m, c.id);
    const cells = [c.employee, c.taskName, durationText_(m.now - c.dueAt) + ' (cut-off ' + fmt_(c.dueAt, 'HH:mm') + ')',
      fmtShort_(resolveDeadline_(c.dueAt, m.cfg))];
    rowsPlain.push(cells.concat([url || 'tick ✅ Ack on the Manager Dashboard']));
    const htmlCells = cells.map(escapeHtml_);
    if (single) return htmlCells;
    return [url ? '<a href="' + escapeHtml_(url) + '" style="color:#1d4ed8;font-weight:bold">Acknowledge</a>' : 'On dashboard'].concat(htmlCells);
  });
  const ackHow = firstUrl
    ? (single ? 'press <b>Acknowledge now</b> above' : 'use the Acknowledge link on each row') + ', or tick ✅ Ack on the Manager Dashboard'
    : 'tick ✅ Ack on the Manager Dashboard';
  const outro =
    '<p style="font-size:15px;margin:20px 0 6px"><b>What we need from you</b></p>' +
    '<ol style="font-size:15px;line-height:1.6;margin:0;padding-left:20px">' +
    '<li>Check with the employee why it is still pending.</li>' +
    '<li><b>Acknowledge by ' + escapeHtml_(fmtShort_(ackDl)) + '</b>: ' + ackHow + ' (a short note helps).</li>' +
    '<li>Make sure it is completed by the “Must be done by” time. If the person is on leave, choose <i>Waive</i> on the dashboard.</li>' +
    '</ol><p style="font-size:15px;line-height:1.5">If either deadline is missed, leadership is informed automatically. ' +
    '<a href="' + escapeHtml_(dash) + '" style="color:#1d4ed8">Open the Manager Dashboard</a></p>';
  return {
    to: to,
    cc: '',
    subject: '[Action needed] ' + n + ' recurring task' + (n > 1 ? 's' : '') + ' pending today — ' + employeesOf_(items),
    htmlBody: emailHtml_({
      accent: '#dc2626',
      heading: single ? items[0].employee + ' has not completed “' + items[0].taskName + '”' : n + ' recurring tasks were not completed by their cut-off',
      intro: 'Still open at ' + escapeHtml_(fmtShort_(m.now)) + ', after the cut-off and a ' + m.cfg.graceMin + '-minute grace period.',
      headers: headers,
      rows: rowsHtml,
      outro: outro,
      cta: single && firstUrl ? { label: 'Acknowledge now', url: firstUrl } : null,
    }),
    body: 'Pending after cut-off:\n' + textTable_(['Employee', 'Task', 'Overdue by', 'Must be done by', 'Acknowledge'], rowsPlain) +
      '\n\nPlease (1) check with the employee, (2) acknowledge by ' + fmtShort_(ackDl) +
      ', (3) make sure each task is completed by its "Must be done by" time (or waive it on the dashboard).\nDashboard: ' + dash,
  };
}

function buildEscalationEmail_(ctx, m, items, reasonById) {
  const headers = ['Why escalated', 'Employee', 'Task', 'Overdue by', 'Manager acknowledged?'];
  const plain = items.map(function (c) {
    return [reasonById[c.id], c.employee, c.taskName, durationText_(m.now - c.dueAt),
      c.ackAt ? 'Yes, ' + fmtShort_(c.ackAt) + (c.ackNote ? ' — "' + c.ackNote + '"' : '') : 'No (alerted ' + fmtShort_(c.alertAt) + ')'];
  });
  const dash = sheetLink_(ctx, SHEETS.DASHBOARD);
  const n = items.length;
  return {
    to: m.cfg.leadershipEmails.join(', '),
    cc: managersCc_(m, items),
    subject: '[Escalation] ' + n + ' recurring task' + (n > 1 ? 's' : '') + ' unresolved — ' + employeesOf_(items),
    htmlBody: emailHtml_({
      accent: '#7f1d1d',
      heading: 'Escalation: ' + (n === 1 ? 'a recurring task is' : n + ' recurring tasks are') + ' still unresolved',
      intro: 'These items missed the agreed deadlines after the manager was alerted. No action is needed from you unless you want to follow up: reply to the manager, who is copied.',
      headers: headers,
      rows: plain.map(function (r) { return r.map(escapeHtml_); }),
      outro: '<p style="font-size:15px;line-height:1.5;margin-top:16px">You will get one more message when each item is resolved. No repeat reminders will be sent.</p>',
      cta: { label: 'Open Manager Dashboard', url: dash },
    }),
    body: 'Escalated:\n' + textTable_(headers, plain) + '\n\nDashboard: ' + dash,
  };
}

function buildResolvedEmail_(ctx, m, items) {
  const headers = ['Employee', 'Task', 'Cut-off', 'Outcome', 'Closed at'];
  const plain = items.map(function (c) {
    const outcome = c.status === STATUS.WAIVED ? c.closeNote : statusLabel_(c.status) + (c.proof ? ' · ' + truncate_(c.proof, 60) : '');
    return [c.employee, c.taskName, fmtShort_(c.dueAt), outcome, fmtShort_(c.completedAt || c.updatedAt)];
  });
  const n = items.length;
  return {
    to: m.cfg.leadershipEmails.join(', '),
    cc: managersCc_(m, items),
    subject: '[Resolved] ' + n + ' escalated task' + (n > 1 ? 's' : '') + ' closed — ' + employeesOf_(items),
    htmlBody: emailHtml_({
      accent: '#15803d',
      heading: n === 1 ? 'Resolved: ' + items[0].taskName + ' (' + items[0].employee + ')' : n + ' escalated tasks are now closed',
      intro: 'Closing the loop on the earlier escalation.',
      headers: headers,
      rows: plain.map(function (r) { return r.map(escapeHtml_); }),
      outro: '',
      cta: { label: 'Open Manager Dashboard', url: sheetLink_(ctx, SHEETS.DASHBOARD) },
    }),
    body: 'Resolved:\n' + textTable_(headers, plain),
  };
}
