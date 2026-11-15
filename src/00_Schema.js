/**
 * 00_Schema — sheet names, column layouts and enums shared by every module.
 * Top level holds literals only: Apps Script loads files in arbitrary order.
 */

const APP_NAME = 'Allo Recurring Tasks';

const SHEETS = {
  START: 'Start Here',
  DASHBOARD: 'Manager Dashboard',
  TASKS: 'Tasks',
  TEAM: 'Team',
  CONFIG: 'Config',
  CYCLES: 'Cycles',
  LOG: 'Activity Log',
  EMAILS: 'Email Log',
  TESTS: 'Test Results',
};

const EMPLOYEE_TAB_PREFIX = '👤 ';

const STATUS = {
  OPEN: 'OPEN',
  DONE: 'DONE',
  LATE: 'LATE_DONE',
  SKIPPED: 'SKIPPED',
  WAIVED: 'WAIVED',
};

const PROOF = { TICKET: 'Ticket ID', DAILY: 'Daily Update' };

const FREQUENCY_OPTIONS = ['Daily', 'Every 7 days', 'Every 15 days', 'Every 30 days'];

const TEAM_FIELDS = [
  ['name', 'Employee'],
  ['email', 'Email'],
  ['managerEmail', 'Manager Email'],
  ['active', 'Active'],
  ['notes', 'Notes'],
];

const TASK_FIELDS = [
  ['id', 'Task ID'],
  ['name', 'Task Name'],
  ['assignedTo', 'Assigned To'],
  ['frequency', 'Frequency'],
  ['proofType', 'Proof Type'],
  ['instructions', 'Instructions'],
  ['dueTime', 'Due Time'],
  ['firstDue', 'First Due Date'],
  ['active', 'Active'],
  ['health', 'Health'],
];

const CYCLE_FIELDS = [
  ['id', 'Cycle ID'],
  ['taskId', 'Task ID'],
  ['taskName', 'Task Name'],
  ['employee', 'Employee'],
  ['frequency', 'Frequency'],
  ['proofType', 'Proof Type'],
  ['slot', 'Slot Date'],
  ['opensAt', 'Window Opens'],
  ['dueAt', 'Due At'],
  ['status', 'Status'],
  ['completedAt', 'Completed At'],
  ['completedBy', 'Completed By'],
  ['proof', 'Proof'],
  ['reminderAt', 'Reminder Sent At'],
  ['alertAt', 'Manager Alert Sent At'],
  ['ackAt', 'Ack At'],
  ['ackBy', 'Ack By'],
  ['ackNote', 'Ack Note'],
  ['escalatedAt', 'Escalated At'],
  ['escalationReason', 'Escalation Reason'],
  ['resolvedNoticeAt', 'Resolved Notice At'],
  ['closeNote', 'Close Note'],
  ['updatedAt', 'Updated At'],
];

const CYCLE_DATE_KEYS = ['slot', 'opensAt', 'dueAt', 'completedAt', 'reminderAt', 'alertAt', 'ackAt',
  'escalatedAt', 'resolvedNoticeAt', 'updatedAt'];

const LOG_HEADERS = ['Timestamp', 'Event', 'Cycle ID', 'Actor', 'Details'];
const EMAIL_HEADERS = ['Timestamp', 'Type', 'To', 'Cc', 'Subject', 'Cycle IDs', 'Delivery', 'Error'];
const TEST_HEADERS = ['ID', 'Test', 'Result', 'Evidence', 'Duration (ms)', 'Ran At'];

/**
 * View geometry (1-based). Column A of every generated view holds the machine Cycle ID and is
 * hidden; all visible text starts in column B. The input + result columns sit early (D/E) so they
 * are reachable on a phone without sideways scrolling.
 */
const ID_COL = 1;
const TITLE_COL = 2;
const EMP_VIEW = { HEADER_ROW: 6, FIRST_ROW: 7, INPUT_COL: 5, RESULT_COL: 6, WIDTH: 9 };

/** Manager dashboard geometry. Section titles are looked up in column B to route edits. */
const DASH = {
  WIDTH: 12,
  TITLE_OVERVIEW: 'TEAM OVERVIEW',
  TITLE_ATTENTION: 'NEEDS ATTENTION — overdue, alerted & escalated',
  TITLE_OPEN: 'ALL OPEN TASKS',
  TITLE_RECENT: 'RECENT COMPLETIONS',
  NOTE_COL: 6,
  ACK_COL: 7,
  ACTION_COL: 8,
  RECENT_PROOF_COL: 7,
};

const ACTION_WAIVE = 'Waive (leave / holiday)';

/** Number formats used by renderers and setup (Google Sheets format syntax). */
const FMT = {
  TEXT: '@',
  PLAIN: '0',
  VIEW_DT: 'ddd d mmm, hh:mm',
  LEDGER_DT: 'yyyy-mm-dd hh:mm',
  DATE: 'yyyy-mm-dd',
};

/** Palette for sheet views (kept in one place so views stay consistent). */
const COLORS = {
  INK: '#1f2933',
  MUTED: '#6b7280',
  WHITE: '#ffffff',
  TITLE_BG: '#ffffff',
  SECTION_BG: '#e8edf5',
  HEADER_BG: '#f3f4f6',
  INPUT_BG: '#fff4c2',
  BAD_BG: '#fde4e1',
  WARN_BG: '#fff1d6',
  OK_BG: '#e6f4ea',
  QUIET_BG: '#f7f7f8',
  ESC_BG: '#f9d0cc',
};
