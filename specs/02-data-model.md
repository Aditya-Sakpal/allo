# 02 — Data Model

All tables have the header in row 1. Code finds columns **by header name**, so reordering columns in the input sheets is safe.

## 1. `Config` (Key | Value | Description | Status)

| Key | Default | Meaning |
|---|---|---|
| MANAGER_EMAIL | manager@example.com | Default manager for same-day alerts (a Team row can override it) |
| LEADERSHIP_EMAILS | leadership@example.com | Comma-separated escalation recipients |
| EMAIL_MODE | LOG_ONLY | `LIVE` sends through MailApp; `LOG_ONLY` writes to the Email Log only |
| WORKING_DAYS | Mon,Tue,Wed,Thu,Fri,Sat | Days daily tasks run and deadlines count |
| HOLIDAYS | *(blank)* | `YYYY-MM-DD, …` dates treated as non-working (no daily cycle; SLA clock skips them) |
| DEFAULT_DUE_TIME | 18:00 | Cut-off when a task has no Due Time |
| REMINDER_LEAD_MINUTES | 120 | Employee reminder before the cut-off (0 = off) |
| MANAGER_ALERT_GRACE_MINUTES | 30 | Manager alert after the cut-off |
| ACK_DEADLINE_TIME | 11:00 | Manager must acknowledge by this time on the next working day |
| RESOLVE_GRACE_WORKING_DAYS | 1 | Working days after the cut-off by which the task must be done |
| TICKET_ID_PATTERN | `^(?=.*[0-9])[A-Za-z0-9][A-Za-z0-9#/_.-]{2,39}$` | Ticket ID validation |
| DAILY_UPDATE_MIN_CHARS | 20 | Minimum length of a daily update |
| RECENT_DAYS | 7 | History window on the views |
| SIMULATED_NOW | *(blank)* | Demo clock `YYYY-MM-DD HH:mm`; blank = real time |
| WEBAPP_URL | *(blank)* | Acknowledgement link base (auto-detected when deployed) |

Invalid values fall back to the default and are reported in the Status column and on the dashboard. They never crash the engine.

## 2. `Team`

| Employee | Email | Manager Email | Active | Notes |
|---|---|---|---|---|

`Employee` is the key used by `Tasks.Assigned To`. A blank Manager Email means `Config.MANAGER_EMAIL` is used.

## 3. `Tasks` (Task Master)

| Task ID | Task Name | Assigned To | Frequency | Proof Type | Instructions | Due Time | First Due Date | Active | Health |
|---|---|---|---|---|---|---|---|---|---|

- **Frequency:** `Daily` (working days) or `Every N days` (N = 2…366; also `7`, `Weekly`, `Monthly` = 30).
- **Proof Type:** `Ticket ID` or `Daily Update`.
- **First Due Date:** the anchor. Slot *k* = First Due + k·N days. If a slot falls on a non-working day, it is **due** on the next working day, but the anchor does not move.
- **Health** is written by the engine: `✅ OK · open cycle due …`, `⚠️ <problem>`, or `⏸ Inactive`.

## 4. `Cycles` (system ledger: one row per task occurrence)

| Column | Notes |
|---|---|
| Cycle ID | `TaskID-yyyymmdd` of the slot. Unique and deterministic, which is the idempotency key. |
| Task ID, Task Name, Employee, Frequency, Proof Type | Copied from the task; kept in sync while the cycle is OPEN |
| Slot Date | Nominal schedule date |
| Window Opens | Earliest moment the cycle can be completed (daily: 00:00 of the day; interval: the previous slot's cut-off) |
| Due At | Cut-off date-time |
| Status | `OPEN`, `DONE`, `LATE_DONE`, `SKIPPED`, `WAIVED` |
| Completed At / Completed By / Proof | Captured once, at completion |
| Reminder Sent At | Employee reminder stamp |
| Manager Alert Sent At | Same-day manager alert stamp |
| Ack At / Ack By / Ack Note | Manager acknowledgement |
| Escalated At / Escalation Reason | Leadership escalation stamp |
| Resolved Notice At | "Resolved" notice to leadership after an escalated item closes |
| Close Note | Why it was skipped or waived |
| Updated At | Last change |

### Cycle state machine

```
             create (engine / previous completion)
                         │
                         ▼
   ┌──────────────────  OPEN  ──────────────────┐
   │ proof ≤ due        │ proof > due           │ manager "Waive" / task deactivated
   ▼                    ▼                       ▼
  DONE              LATE_DONE                 WAIVED
                                    (interval only) slots that passed while the previous
                                    cycle was overdue are written directly as SKIPPED
```

Notification stamps are orthogonal to status. Status values are terminal except OPEN.

## 5. Logs

- **Activity Log:** `Timestamp | Event | Cycle ID | Actor | Details`. Events include CYCLE_CREATED, CYCLE_SKIPPED, COMPLETED, PROOF_REJECTED, REASSIGNED, ACKNOWLEDGED, WAIVED, REMINDER_SENT, MANAGER_ALERT_SENT, ALERT_SKIPPED, ESCALATED, RESOLVED_NOTICE_SENT, EMAIL_FAILED, CONFIG_WARNING.
- **Email Log:** `Timestamp | Type | To | Cc | Subject | Cycle IDs | Delivery | Error`. Delivery is `SENT`, `LOG_ONLY`, `CAPTURED` (tests) or `FAILED`.

## 6. Generated views

- **`👤 <Employee>`:** ACTION NEEDED (input column E, result column F; Cycle ID in hidden column A), COMING UP, COMPLETED (last N days, with the WhatsApp share link for daily updates).
- **`Manager Dashboard`:** TEAM OVERVIEW, NEEDS ATTENTION (✅ ack checkbox, note, Waive action), ALL OPEN TASKS, RECENT COMPLETIONS.

Views are disposable. They are rebuilt from the ledger on every engine run and after every handled edit.
