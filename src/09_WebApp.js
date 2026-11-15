/**
 * 09_WebApp — one-click acknowledgement from the manager-alert email.
 * GET only renders a confirmation page (email link scanners can't acknowledge on the manager's behalf);
 * the Confirm button calls confirmAck() via google.script.run. Links are HMAC-signed per cycle.
 * Spec: specs/features/F06-manager-dashboard-and-ack.md
 */

function ackSecret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('ACK_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('ACK_SECRET', s);
  }
  return s;
}

function signId_(id) {
  const raw = Utilities.computeHmacSha256Signature(String(id), ackSecret_());
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
}

function verifySig_(id, sig) {
  if (!id || !sig) return false;
  const expected = signId_(id);
  sig = String(sig);
  if (expected.length !== sig.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ sig.charCodeAt(i);
  return diff === 0;
}

function webAppBaseUrl_(ctx, m) {
  if (ctx.webAppUrl) return ctx.webAppUrl;
  if (m.cfg.webAppUrl) return m.cfg.webAppUrl;
  if (m._autoUrl === undefined) {
    try { m._autoUrl = ScriptApp.getService().getUrl() || ''; } catch (e) { m._autoUrl = ''; }
  }
  return m._autoUrl;
}

function ackUrl_(ctx, m, id) {
  const base = webAppBaseUrl_(ctx, m);
  if (!base) return '';
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'a=ack&c=' + encodeURIComponent(id) + '&s=' + signId_(id);
}

function hostSpreadsheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet not registered. Open the sheet and run Setup once.');
  return SpreadsheetApp.openById(id);
}

/* ───────────── page rendering ───────────── */

/** JSON string literal that is also safe inside an inline <script>. */
function jsString_(s) {
  return JSON.stringify(String(s)).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');
}

function ackPageHtml_(title, bodyHtml) {
  return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + escapeHtml_(title) + '</title><style>' +
    ':root{color-scheme:light dark;--ink:#1f2933;--muted:#5f6b7a;--line:#e5e7eb;--bg:#f4f5f7;--card:#ffffff;--accent:#1f2933;--on-accent:#ffffff;--ok:#15803d;--bad:#b42318;--warn-bg:#fff4e5}' +
    '@media (prefers-color-scheme: dark){:root{--ink:#e7eaee;--muted:#a3adb9;--line:#353c45;--bg:#0f1216;--card:#191e24;--accent:#e7eaee;--on-accent:#111418;--ok:#4ade80;--bad:#f87171;--warn-bg:#2a2216}}' +
    '*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 -apple-system,"Segoe UI",Roboto,Arial,sans-serif}' +
    '.wrap{max-width:560px;margin:0 auto;padding:24px 16px 40px}.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:22px}' +
    '.label{font-size:13px;font-weight:600;color:var(--muted);margin:0 0 4px}h1{font-size:22px;line-height:1.3;margin:0 0 16px}' +
    'dl{display:grid;grid-template-columns:auto 1fr;gap:6px 16px;margin:0 0 18px}dt{color:var(--muted)}dd{margin:0;font-weight:600}' +
    '.deadline{background:var(--warn-bg);border-radius:8px;padding:10px 12px;margin:0 0 18px;font-size:15px}' +
    'label{display:block;font-size:14px;color:var(--muted);margin-bottom:6px}textarea{width:100%;min-height:96px;padding:10px;border:1px solid var(--line);' +
    'border-radius:8px;background:transparent;color:inherit;font:inherit}textarea:focus{outline:2px solid var(--accent);outline-offset:1px}' +
    'button{margin-top:14px;width:100%;min-height:48px;padding:12px 16px;border:0;border-radius:8px;background:var(--accent);color:var(--on-accent);' +
    'font:inherit;font-weight:600;line-height:1.2;cursor:pointer}button:focus-visible{outline:3px solid var(--muted);outline-offset:2px}' +
    'button:disabled{opacity:.6;cursor:progress}.msg{margin-top:14px;font-weight:600;min-height:1.5em}.ok{color:var(--ok)}.bad{color:var(--bad)}' +
    '.note{color:var(--muted);font-size:14px;margin:16px 0 0}' +
    '</style></head><body><main class="wrap"><div class="card">' + bodyHtml +
    '</div><p class="note" style="text-align:center">' + escapeHtml_(APP_NAME) + '</p></main></body></html>';
}

function simplePage_(title, heading, text) {
  return { title: title, html: ackPageHtml_(title, '<h1>' + escapeHtml_(heading) + '</h1><p>' + escapeHtml_(text) + '</p>') };
}

/** Pure: builds the page for a GET. Never changes state (email link scanners must not acknowledge). */
function doGetCore_(ctx, params) {
  params = params || {};
  if (params.a !== 'ack') return simplePage_(APP_NAME, 'Nothing to do here', 'Open the spreadsheet to see your tasks.');
  const id = String(params.c || '');
  const sig = String(params.s || '');
  if (!verifySig_(id, sig)) {
    return simplePage_('Link not valid', 'This link is not valid',
      'It may have been changed or copied incompletely. Please acknowledge from the Manager Dashboard instead.');
  }
  const m = loadModel_(ctx);
  const c = m.byId[id];
  if (!c) return simplePage_('Task not found', 'This task could not be found', 'It may have been removed. Please check the Manager Dashboard.');

  const isOpen = c.status === STATUS.OPEN;
  const details = '<dl><dt>Employee</dt><dd>' + escapeHtml_(c.employee) + '</dd>' +
    '<dt>Cut-off</dt><dd>' + escapeHtml_(fmtLong_(c.dueAt)) + '</dd>' +
    '<dt>Task status</dt><dd>' + escapeHtml_(isOpen ? 'Still open — overdue by ' + durationText_(m.now - c.dueAt) : statusLabel_(c.status)) + '</dd>' +
    (c.ackAt ? '<dt>Acknowledged</dt><dd>' + escapeHtml_(fmtShort_(c.ackAt) + ' by ' + c.ackBy) + '</dd>' : '') + '</dl>';
  const heading = '<p class="label">' + (c.ackAt ? 'Acknowledged' : 'Acknowledge pending task') + '</p>' +
    '<h1>' + escapeHtml_(c.employee + ': ' + c.taskName) + '</h1>';

  if (!isOpen && !c.ackAt) {
    return {
      title: 'Already completed',
      html: ackPageHtml_('Already completed', '<p class="label">Nothing to do</p><h1>' + escapeHtml_(c.employee + ': ' + c.taskName) +
        '</h1>' + details + '<p>This task is already closed, so no acknowledgement is needed.</p>'),
    };
  }
  const deadline = isOpen
    ? '<p class="deadline">Make sure it is done by <b>' + escapeHtml_(fmtShort_(resolveDeadline_(c.dueAt, m.cfg))) +
      '</b>. If it is not, leadership is informed automatically.</p>'
    : '';
  const form = '<label for="note">' + (c.ackAt ? 'Update your note' : 'Note for the record (optional): the reason, or when it will be done') + '</label>' +
    '<textarea id="note" maxlength="500">' + escapeHtml_(c.ackNote) + '</textarea>' +
    '<button id="go" type="button">' + (c.ackAt ? 'Save note' : 'Confirm acknowledgement') + '</button>' +
    '<div id="msg" class="msg" role="status" aria-live="polite"></div>' +
    '<script>(function(){var b=document.getElementById("go"),m=document.getElementById("msg");b.onclick=function(){b.disabled=true;m.className="msg";m.textContent="Saving…";' +
    'google.script.run.withSuccessHandler(function(r){m.className="msg "+(r.ok?"ok":"bad");m.textContent=r.msg;b.disabled=false;})' +
    '.withFailureHandler(function(){m.className="msg bad";m.textContent="Could not save. Please try again, or tick ✅ Ack on the Manager Dashboard.";b.disabled=false;})' +
    '.confirmAck(' + jsString_(id) + ',' + jsString_(sig) + ',document.getElementById("note").value);};})();</script>';
  return { title: 'Acknowledge: ' + c.taskName, html: ackPageHtml_('Acknowledge', heading + details + deadline + form) };
}

/** State change for the Confirm button. Same code path as the dashboard tick. */
function confirmAckCore_(ctx, id, sig, note) {
  if (!verifySig_(id, sig)) return { ok: false, msg: 'This link is not valid.' };
  const m = loadModel_(ctx);
  const c = m.byId[String(id)];
  if (!c) return { ok: false, msg: 'This task could not be found.' };
  let actor = '';
  try { actor = Session.getActiveUser().getEmail(); } catch (e) { actor = ''; }
  actor = actor ? actor.toLowerCase() + ' (email link)' : 'Manager (email link)';
  const res = ackCycle_(m, c, actor, note);
  if (res.ok) {
    saveModel_(ctx, m);
    renderAll_(ctx, m, []);
  }
  return { ok: res.ok, msg: res.msg };
}
