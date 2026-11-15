'use strict';
/**
 * A deliberately STRICT in-memory mock of the Google Apps Script services this project uses.
 * Faithful behaviours that matter for correctness:
 *  - getRange() beyond the sheet's max rows/cols throws (as in Apps Script)
 *  - setValues()/setBackgrounds()/... throw on dimension mismatch
 *  - insertCheckboxes() sets every cell in the range to FALSE
 *  - a "reject input" list validation (setAllowInvalid(false)) makes setValue(s) of any other value throw,
 *    with Sheets' own message; Sheet.clear() keeps validations (only clearDataValidations() removes them)
 *  - getLastRow()/getLastColumn() count any non-empty value (FALSE included)
 *  - unknown getX / isX / hasX methods throw, so code can't silently rely on unmocked reads
 * Formatting/setter methods without data semantics are no-ops that return `this` for chaining.
 */
const crypto = require('crypto');

const isDateLike = (v) => Object.prototype.toString.call(v) === '[object Date]';

function strict(target, label) {
  return new Proxy(target, {
    get(t, p, recv) {
      if (p in t) return t[p];
      if (typeof p === 'symbol' || p === 'then' || p === 'toJSON' || p === 'inspect' || p === 'constructor') return undefined;
      if (/^(get|is|has|can)[A-Z]/.test(p)) {
        return () => { throw new Error(`[gas-mock] ${label}.${String(p)}() is not implemented in the mock`); };
      }
      return function () { return recv; };
    },
  });
}

function pad(n) { return (n < 10 ? '0' : '') + n; }

function displayOf(v) {
  if (v === undefined || v === null) return '';
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  if (isDateLike(v)) {
    const d = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    return v.getHours() || v.getMinutes() ? `${d} ${pad(v.getHours())}:${pad(v.getMinutes())}` : d;
  }
  return String(v);
}

function checkDims(data, nr, nc, what) {
  if (!Array.isArray(data) || data.length !== nr) {
    throw new Error(`[gas-mock] ${what}: the data has ${Array.isArray(data) ? data.length : '?'} rows but the range has ${nr}.`);
  }
  data.forEach((row, i) => {
    if (!Array.isArray(row) || row.length !== nc) {
      throw new Error(`[gas-mock] ${what}: row ${i} has ${Array.isArray(row) ? row.length : '?'} columns but the range has ${nc}.`);
    }
  });
}

class MockRange {
  constructor(sheet, row, col, nr, nc) {
    this._sheet = sheet; this._r = row; this._c = col; this._nr = nr; this._nc = nc;
  }
  getRow() { return this._r; }
  getColumn() { return this._c; }
  getNumRows() { return this._nr; }
  getNumColumns() { return this._nc; }
  getLastRow() { return this._r + this._nr - 1; }
  getLastColumn() { return this._c + this._nc - 1; }
  getSheet() { return this._sheet._proxy; }
  getA1Notation() {
    const col = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    const a = col(this._c) + this._r;
    return this._nr === 1 && this._nc === 1 ? a : `${a}:${col(this._c + this._nc - 1)}${this._r + this._nr - 1}`;
  }
  _each(fn) {
    for (let i = 0; i < this._nr; i++) for (let j = 0; j < this._nc; j++) fn(this._r + i, this._c + j, i, j);
  }
  getValues() {
    const out = [];
    for (let i = 0; i < this._nr; i++) {
      const row = [];
      for (let j = 0; j < this._nc; j++) {
        const v = this._sheet._get(this._r + i, this._c + j);
        row.push(isDateLike(v) ? new v.constructor(v.getTime()) : v);
      }
      out.push(row);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map((r) => r.map(displayOf)); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return displayOf(this.getValue()); }
  /** Like Sheets: a string starting with "=" is a formula unless the cell is formatted as plain text ('@'). */
  getFormula() {
    const v = this._sheet._get(this._r, this._c);
    const fmt = this._sheet._fmt.get(`${this._r}:${this._c}`);
    return typeof v === 'string' && v.charAt(0) === '=' && fmt !== '@' ? v : '';
  }
  getNumberFormat() { return this._sheet._fmt.get(`${this._r}:${this._c}`) || 'General'; }
  setValues(data) {
    checkDims(data, this._nr, this._nc, 'setValues');
    data.forEach((row, i) => row.forEach((v, j) => {
      if (v === undefined) throw new Error(`[gas-mock] setValues: undefined at [${i}][${j}]`);
      if (v !== null && typeof v === 'object' && !isDateLike(v)) throw new Error(`[gas-mock] setValues: object at [${i}][${j}]`);
      if (typeof v === 'number' && !isFinite(v)) throw new Error(`[gas-mock] setValues: non-finite number at [${i}][${j}]`);
      if (typeof v === 'string' && v.length > 50000) throw new Error('[gas-mock] setValues: cell exceeds 50000 characters');
    }));
    data.forEach((row, i) => row.forEach((v, j) => this._sheet._checkRule(this._r + i, this._c + j, v)));
    data.forEach((row, i) => row.forEach((v, j) => this._sheet._set(this._r + i, this._c + j, v === null ? '' : v)));
    return this;
  }
  setValue(v) {
    if (v === undefined) throw new Error('[gas-mock] setValue: undefined');
    this._each((r, c) => this._sheet._checkRule(r, c, v));
    this._each((r, c) => this._sheet._set(r, c, v === null ? '' : v));
    return this;
  }
  clearContent() { this._each((r, c) => this._sheet._set(r, c, '')); return this; }
  clear() { return this.clearContent(); }
  insertCheckboxes() {
    this._each((r, c) => { this._sheet._checks.add(`${r}:${c}`); this._sheet._set(r, c, false); });
    return this;
  }
  clearDataValidations() {
    this._each((r, c) => { this._sheet._checks.delete(`${r}:${c}`); this._sheet._dv.delete(`${r}:${c}`); });
    return this;
  }
  setNumberFormats(f) {
    checkDims(f, this._nr, this._nc, 'setNumberFormats');
    this._each((r, c, i, j) => this._sheet._fmt.set(`${r}:${c}`, f[i][j]));
    return this;
  }
  setNumberFormat(f) {
    if (typeof f !== 'string') throw new Error('[gas-mock] setNumberFormat expects a string');
    this._each((r, c) => this._sheet._fmt.set(`${r}:${c}`, f));
    return this;
  }
  setBackgrounds(f) { checkDims(f, this._nr, this._nc, 'setBackgrounds'); return this; }
  setFontWeights(f) { checkDims(f, this._nr, this._nc, 'setFontWeights'); return this; }
  setFontColors(f) { checkDims(f, this._nr, this._nc, 'setFontColors'); return this; }
  /** null removes the rule (as in Apps Script); anything else must come from builder.build(). */
  setDataValidation(rule) {
    if (rule !== null && (!rule || !rule.__dv)) throw new Error('[gas-mock] setDataValidation needs a built rule or null');
    this._each((r, c) => (rule === null ? this._sheet._dv.delete(`${r}:${c}`) : this._sheet._dv.set(`${r}:${c}`, rule)));
    return this;
  }
}

class MockSheet {
  constructor(ss, name, id) {
    this._ss = ss; this._name = name; this._id = id;
    this._cells = []; this._maxRows = 1000; this._maxCols = 26;
    this._checks = new Set();
    this._fmt = new Map();
    this._dv = new Map();
    this._proxy = strict(this, 'Sheet');
  }
  getName() { return this._name; }
  setName(n) { this._name = n; return this._proxy; }
  getSheetId() { return this._id; }
  getParent() { return this._ss._proxy; }
  getMaxRows() { return this._maxRows; }
  getMaxColumns() { return this._maxCols; }
  insertRowsAfter(after, n) {
    if (after < 1 || after > this._maxRows || n < 1) throw new Error('[gas-mock] insertRowsAfter: bad arguments');
    this._cells.splice(after, 0, ...new Array(n).fill(undefined));
    this._maxRows += n;
    return this._proxy;
  }
  insertColumnsAfter(after, n) { this._maxCols += n; return this._proxy; }
  getLastRow() {
    for (let r = this._cells.length - 1; r >= 0; r--) {
      const row = this._cells[r];
      if (row && row.some((v) => v !== undefined && v !== '')) return r + 1;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this._cells.forEach((row) => {
      if (!row) return;
      for (let c = row.length - 1; c >= 0; c--) if (row[c] !== undefined && row[c] !== '') { max = Math.max(max, c + 1); break; }
    });
    return max;
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') return this._a1(a);
    const nr = c === undefined ? 1 : c;
    const nc = d === undefined ? 1 : d;
    if (![a, b, nr, nc].every((x) => Number.isInteger(x))) throw new Error(`[gas-mock] getRange: non-integer arguments (${a},${b},${c},${d})`);
    if (a < 1 || b < 1) throw new Error('[gas-mock] getRange: coordinates start at 1');
    if (nr < 1) throw new Error('[gas-mock] The number of rows in the range must be at least 1.');
    if (nc < 1) throw new Error('[gas-mock] The number of columns in the range must be at least 1.');
    if (a + nr - 1 > this._maxRows || b + nc - 1 > this._maxCols) {
      throw new Error(`[gas-mock] Range ${a},${b} ${nr}x${nc} is outside the sheet "${this._name}" (${this._maxRows}x${this._maxCols}).`);
    }
    return strict(new MockRange(this, a, b, nr, nc), 'Range');
  }
  _a1(s) {
    const m = String(s).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) throw new Error('[gas-mock] unsupported A1 ' + s);
    const col = (t) => t.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const r1 = Number(m[2]); const c1 = col(m[1]);
    const r2 = m[4] ? Number(m[4]) : r1; const c2 = m[3] ? col(m[3]) : c1;
    return this.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  clear() { this._cells = []; this._fmt = new Map(); return this._proxy; }
  clearContents() { this._cells = []; return this._proxy; }
  appendRow(row) {
    const r = this.getLastRow() + 1;
    if (r > this._maxRows) this._maxRows = r;
    row.forEach((v, i) => this._set(r, i + 1, v));
    return this._proxy;
  }
  protect() {
    const p = strict({ __protection: true }, 'Protection');
    return p;
  }
  /** Sheets rejects a scripted write that breaks a "reject input" list rule (blank is always allowed). */
  _checkRule(r, c, v) {
    const rule = this._dv.get(`${r}:${c}`);
    if (!rule || rule.allowInvalid || !rule.list || v === '' || v === null) return;
    if (rule.list.indexOf(String(v)) >= 0) return;
    const a1 = new MockRange(this, r, c, 1, 1).getA1Notation();
    throw new Error(`The data you entered in cell ${a1} violates the data validation rules set on this cell. ` +
      `Please enter one of the following values: ${rule.list.join(', ')}.`);
  }
  _get(r, c) {
    const row = this._cells[r - 1];
    const v = row ? row[c - 1] : undefined;
    return v === undefined ? '' : v;
  }
  _set(r, c, v) {
    if (!this._cells[r - 1]) this._cells[r - 1] = [];
    this._cells[r - 1][c - 1] = v;
  }
}

class MockSpreadsheet {
  constructor(state, name, id) {
    this._state = state; this._name = name; this._id = id;
    this._sheets = []; this._nextSheetId = 100; this._tz = 'America/New_York';
    this._toasts = []; this._active = null;
    this._proxy = strict(this, 'Spreadsheet');
    this.insertSheet('Sheet1');
  }
  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return `https://docs.google.com/spreadsheets/d/${this._id}/edit`; }
  getSheetByName(n) { const s = this._sheets.find((x) => x._name === n); return s ? s._proxy : null; }
  getSheets() { return this._sheets.map((s) => s._proxy); }
  insertSheet(name, index) {
    if (typeof name !== 'string' || !name) throw new Error('[gas-mock] insertSheet needs a name');
    if (this._sheets.some((s) => s._name === name)) throw new Error(`A sheet with the name "${name}" already exists. Please enter another name.`);
    const s = new MockSheet(this, name, this._nextSheetId++);
    if (index === undefined || index > this._sheets.length) this._sheets.push(s);
    else this._sheets.splice(index, 0, s);
    return s._proxy;
  }
  deleteSheet(sheetProxy) {
    const i = this._sheets.findIndex((s) => s._proxy === sheetProxy);
    if (i < 0) throw new Error('[gas-mock] deleteSheet: sheet not in this spreadsheet');
    if (this._sheets.length === 1) throw new Error("You can't remove all the sheets in a document.");
    this._sheets.splice(i, 1);
  }
  setActiveSheet(s) { this._active = s; return s; }
  getActiveSheet() { return this._active || this._sheets[0]._proxy; }
  setSpreadsheetTimeZone(tz) { this._tz = tz; }
  getSpreadsheetTimeZone() { return this._tz; }
  toast(msg) { this._toasts.push(String(msg)); }
}

function createGas(options = {}) {
  const state = {
    spreadsheets: {},
    active: null,
    triggers: [],
    scriptProps: {},
    docProps: {},
    sentMail: [],
    mailThrows: null,
    mailQuota: 100,
    lockBusy: false,
    activeUser: options.activeUser || '',
    webAppUrl: null,
    trashed: [],
    menus: [],
    alerts: [],
    uiResponse: 'YES',
    seq: 0,
  };

  const newSpreadsheet = (name) => {
    const id = 'ss' + (++state.seq).toString().padStart(4, '0');
    const ss = new MockSpreadsheet(state, name, id);
    state.spreadsheets[id] = ss;
    return ss._proxy;
  };

  const props = (store) => strict({
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null),
    setProperty(k, v) { store[k] = String(v); return this; },
    deleteProperty(k) { delete store[k]; return this; },
    getProperties: () => Object.assign({}, store),
  }, 'Properties');

  const ui = strict({
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    Button: { OK: 'OK', YES: 'YES', NO: 'NO', CANCEL: 'CANCEL' },
    alert: (...args) => { state.alerts.push(args.map(String).join(' | ')); return state.uiResponse; },
    createMenu: (name) => {
      const menu = { name, items: [] };
      const api = strict({
        addItem: (label, fn) => { menu.items.push({ label, fn }); return api; },
        addSeparator: () => api,
        addSubMenu: (sub) => { menu.items.push({ sub: sub.__menu }); return api; },
        addToUi: () => { state.menus.push(menu); },
        __menu: menu,
      }, 'Menu');
      return api;
    },
  }, 'Ui');

  // Records only what has data semantics: list rules and whether invalid input is allowed (Sheets default: true).
  const dvBuilder = () => {
    const spec = { list: null, allowInvalid: true };
    const b = strict({
      requireValueInList: (list) => { spec.list = list.map(String); return b; },
      setAllowInvalid: (yes) => { spec.allowInvalid = !!yes; return b; },
      build: () => ({ __dv: true, list: spec.list, allowInvalid: spec.allowInvalid }),
    }, 'DataValidationBuilder');
    return b;
  };

  const SpreadsheetApp = strict({
    getActive: () => state.active,
    getActiveSpreadsheet: () => state.active,
    create: (name) => newSpreadsheet(name),
    openById: (id) => { const s = state.spreadsheets[id]; if (!s) throw new Error('No spreadsheet ' + id); return s._proxy; },
    getUi: () => { if (options.noUi) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); return ui; },
    newDataValidation: dvBuilder,
    flush: () => {},
  }, 'SpreadsheetApp');

  const LockService = strict({
    getScriptLock: () => strict({
      tryLock: () => !state.lockBusy,
      waitLock: () => { if (state.lockBusy) throw new Error('Lock timeout'); },
      releaseLock: () => {},
      hasLock: () => !state.lockBusy,
    }, 'Lock'),
  }, 'LockService');

  const PropertiesService = strict({
    getScriptProperties: () => props(state.scriptProps),
    getDocumentProperties: () => props(state.docProps),
  }, 'PropertiesService');

  const trigger = (fn, type, source) => {
    const t = strict({
      getHandlerFunction: () => fn,
      getEventType: () => type,
      getUniqueId: () => fn + '#' + state.triggers.length,
      getTriggerSourceId: () => source || '',
    }, 'Trigger');
    return t;
  };
  const ScriptApp = strict({
    getProjectTriggers: () => state.triggers.slice(),
    deleteTrigger: (t) => { state.triggers = state.triggers.filter((x) => x !== t); },
    newTrigger: (fn) => strict({
      timeBased: () => {
        const tb = strict({
          everyMinutes: (n) => { if (![1, 5, 10, 15, 30].includes(n)) throw new Error('everyMinutes must be 1,5,10,15,30'); return tb; },
          everyHours: () => tb,
          create: () => { const t = trigger(fn, 'CLOCK'); state.triggers.push(t); return t; },
        }, 'ClockTriggerBuilder');
        return tb;
      },
      forSpreadsheet: (ss) => strict({
        onEdit: () => strict({ create: () => { const t = trigger(fn, 'ON_EDIT', ss.getId()); state.triggers.push(t); return t; } }, 'TB'),
        onOpen: () => strict({ create: () => { const t = trigger(fn, 'ON_OPEN', ss.getId()); state.triggers.push(t); return t; } }, 'TB'),
      }, 'SpreadsheetTriggerBuilder'),
    }, 'TriggerBuilder'),
    getService: () => strict({ getUrl: () => state.webAppUrl }, 'Service'),
    EventType: { CLOCK: 'CLOCK', ON_EDIT: 'ON_EDIT', ON_OPEN: 'ON_OPEN' },
  }, 'ScriptApp');

  const MailApp = strict({
    sendEmail: (arg) => {
      if (state.mailThrows) throw new Error(state.mailThrows);
      if (!arg || typeof arg !== 'object' || !arg.to || !arg.subject) throw new Error('[gas-mock] sendEmail expects {to, subject, ...}');
      state.sentMail.push(arg);
      state.mailQuota -= 1;
    },
    getRemainingDailyQuota: () => state.mailQuota,
  }, 'MailApp');

  const Session = strict({
    getScriptTimeZone: () => 'Asia/Kolkata',
    getActiveUser: () => strict({ getEmail: () => state.activeUser }, 'User'),
    getEffectiveUser: () => strict({ getEmail: () => 'owner@allo.test' }, 'User'),
  }, 'Session');

  const Utilities = strict({
    computeHmacSha256Signature: (value, key) => Array.from(crypto.createHmac('sha256', String(key)).update(String(value), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b)),
    base64EncodeWebSafe: (bytes) => Buffer.from(bytes.map((b) => b & 255)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
    getUuid: () => crypto.randomUUID(),
    sleep: () => {},
  }, 'Utilities');

  const HtmlService = strict({
    createHtmlOutput: (html) => {
      const out = strict({
        _html: html, _title: '',
        setTitle(t) { this._title = t; return out; },
        addMetaTag() { return out; },
        setXFrameOptionsMode() { return out; },
        getContent() { return html; },
        getTitle() { return this._title; },
      }, 'HtmlOutput');
      return out;
    },
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
  }, 'HtmlService');

  const DriveApp = strict({
    getFileById: (id) => strict({ setTrashed: (b) => { if (b) state.trashed.push(id); } }, 'File'),
    getFilesByName: (name) => {
      const ids = Object.keys(state.spreadsheets).filter((id) => state.spreadsheets[id]._name === name && !state.trashed.includes(id));
      return strict({
        hasNext: () => ids.length > 0,
        next: () => { const id = ids.shift(); if (!id) throw new Error('No more files'); return DriveApp.getFileById(id); },
      }, 'FileIterator');
    },
  }, 'DriveApp');

  return {
    state,
    newSpreadsheet,
    setActive: (ss) => { state.active = ss; },
    globals: { SpreadsheetApp, LockService, PropertiesService, ScriptApp, MailApp, Session, Utilities, HtmlService, DriveApp },
  };
}

module.exports = { createGas, displayOf };
