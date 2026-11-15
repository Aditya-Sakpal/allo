/**
 * 01_Util — small pure helpers (dates, strings, validation). No Apps Script services here.
 * All Date maths uses the script's local time zone (Asia/Kolkata, pinned in appsscript.json).
 */

const DOW_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Realm-safe Date check (values can come from another JS realm in tests). */
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

/** Tokens: yyyy MMM MM dd d EEE HH mm. Patterns must not contain other letters that collide. */
function fmt_(d, pattern) {
  if (!isDate_(d)) return '';
  const map = {
    yyyy: String(d.getFullYear()),
    MMM: MONTH_NAMES[d.getMonth()],
    MM: pad2_(d.getMonth() + 1),
    dd: pad2_(d.getDate()),
    d: String(d.getDate()),
    EEE: DOW_NAMES[d.getDay()],
    HH: pad2_(d.getHours()),
    mm: pad2_(d.getMinutes()),
  };
  return pattern.replace(/yyyy|MMM|MM|dd|EEE|HH|mm|d/g, function (t) { return map[t]; });
}

/** "Mon 16 Nov, 18:00" */
function fmtShort_(d) { return fmt_(d, 'EEE d MMM, HH:mm'); }
/** "Mon 16 Nov 2026, 18:00" */
function fmtLong_(d) { return fmt_(d, 'EEE d MMM yyyy, HH:mm'); }
function ymd_(d) { return fmt_(d, 'yyyyMMdd'); }

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** Calendar-day arithmetic that keeps wall-clock time (DST-safe). */
function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n, d.getHours(), d.getMinutes(), d.getSeconds());
}

function addMinutes_(d, n) {
  return new Date(d.getTime() + n * 60000);
}

function sameDay_(a, b) {
  return isDate_(a) && isDate_(b) && a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function atTime_(d, hm) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), hm.h, hm.m, 0, 0);
}

/** "18:00" → {h:18, m:0}; anything else → null. */
function parseHM_(s) {
  const m = String(s == null ? '' : s).trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const mi = Number(m[2]);
  if (h > 23 || mi > 59) return null;
  return { h: h, m: mi };
}

/** Accepts a Date, "YYYY-MM-DD", "YYYY-MM-DD HH:mm[:ss]" or "YYYY-MM-DDTHH:mm". Rejects impossible dates. */
function parseDateTime_(v) {
  if (isDate_(v)) return new Date(v.getTime());
  const s = String(v == null ? '' : v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const da = Number(m[3]);
  const h = m[4] ? Number(m[4]) : 0;
  const mi = m[5] ? Number(m[5]) : 0;
  const se = m[6] ? Number(m[6]) : 0;
  if (h > 23 || mi > 59 || se > 59) return null;
  const d = new Date(y, mo, da, h, mi, se);
  if (d.getFullYear() !== y || d.getMonth() !== mo || d.getDate() !== da) return null;
  return d;
}

function toBool_(v) {
  if (v === true) return true;
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

/** 0 → "0m", 135 min → "2h 15m", 2 days → "2d 0h". */
function durationText_(ms) {
  const mins = Math.floor(Math.abs(ms) / 60000);
  const d = Math.floor(mins / 1440);
  const h = Math.floor((mins % 1440) / 60);
  const m = mins % 60;
  if (d) return d + 'd ' + h + 'h';
  if (h) return h + 'h ' + m + 'm';
  return m + 'm';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function truncate_(s, n) {
  s = String(s == null ? '' : s);
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

function normalizeSpace_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}

function isEmail_(s) {
  return /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(String(s == null ? '' : s).trim());
}

function splitEmails_(s) {
  return String(s == null ? '' : s).split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(Boolean);
}

function uniq_(arr) {
  const seen = {};
  return arr.filter(function (x) { if (seen[x]) return false; seen[x] = true; return true; });
}

function sanitizeSheetName_(s) {
  return String(s).replace(/[\[\]*?\/\\:]/g, '-').slice(0, 90);
}

function byDue_(a, b) {
  return a.dueAt.getTime() - b.dueAt.getTime();
}
