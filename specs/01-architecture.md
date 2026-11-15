# 01 — Architecture

## 1. Component view

```
┌────────────────────────────── Google Spreadsheet (single file) ─────────────────────────────┐
│                                                                                              │
│  INPUT (humans)                 SYSTEM-OWNED (script writes)             OUTPUT (humans read) │
│  ─────────────                  ─────────────────────────────            ─────────────────── │
│  Config   (admin)               Cycles      ← ledger, 1 row per cycle    👤 <Employee> tabs   │
│  Team     (admin)               Activity Log ← audit trail               Manager Dashboard    │
│  Tasks    (admin: Task Master)  Email Log   ← every email + delivery     Start Here           │
│                                 Test Results ← in-sheet test suite                            │
└──────────────────────────────────────────────────────────────────────────────────────────────┘
            ▲ edits (installable onEdit)             ▲ every 15 min (time trigger)
            │                                        │
┌───────────┴────────────────────────────────────────┴─────────────────────────────────────────┐
│ Apps Script (bound)                                                                            │
│  11_Main      entry points: runEngine · handleEdit · onOpen menu · demo clock · tests          │
│  08_Edits     routes an edit → completion / acknowledgement / waive / master-data sync         │
│  05_Cycles    lifecycle: ensure cycles, complete, acknowledge, waive, validate proof          │
│  06_Notify    reminder / manager alert / leadership escalation / resolved notice + SLA rules   │
│  07_Views     renders employee tabs + manager dashboard (batch writes)                         │
│  09_WebApp    doGet → confirm page → confirmAck (HMAC-signed one-click acknowledgement)        │
│  04_Store     table read/write by header name, model load/save                                │
│  03_Schedule  slot maths, working days, deadlines                                             │
│  02_Config    typed config + validation, context (ss, clock, mailer)                          │
│  10_Setup     builds/repairs sheets non-destructively, demo seed, triggers                     │
│  12_Tests     self-contained test suite (runs in Sheets on a sandbox copy, and in Node)        │
└────────────────────────────────────────────────────────────────────────────────────────────────┘
            │ MailApp (LIVE) or Email Log only (LOG_ONLY)
            ▼
     Employee · Manager · Leadership inboxes
```

## 2. The context object (dependency injection)

Every core function receives `ctx`:

```js
ctx = {
  ss,          // Spreadsheet to operate on (the real one, or a sandbox in tests)
  cfg,         // parsed + validated Config
  now(),       // clock: real time, Config.SIMULATED_NOW (demo), or a test clock
  mailer,      // LIVE → MailApp, LOG_ONLY → no-op, CAPTURE → array (tests)
  webAppUrl,   // optional override for acknowledgement links
}
```

Entry points (`runEngine`, `handleEdit`, `confirmAck`) are thin wrappers. Each one takes the script lock, builds the `ctx` and calls the `…Core_` function. Tests call the `…Core_` functions directly with a sandbox `ctx`.

## 3. Runtime flows

### 3.1 Engine run (time trigger, every 15 min; also after master-data edits and demo-clock moves)

```
lock → load model (Team, Tasks, Cycles) → validate tasks (writes Health)
     → ensureCycles: create today's daily cycles / the next interval cycle (idempotent IDs)
     → processNotifications: reminders, manager alerts, escalations, resolved notices
          (for each batch: send → stamp cycles → save immediately → Email Log)
     → save (Cycles, Activity Log, Task health) → render views → record "last run" → unlock
```

### 3.2 Employee completes a task (installable onEdit)

```
edit in 👤 tab, column E (yellow input) → lock → read Cycle ID from hidden column A of that row
  → checks: belongs to this employee? cycle OPEN? window open? proof valid & not reused?
     ├─ no  → write ❌ reason in column F (Result); nothing else changes
     └─ yes → status DONE/LATE_DONE, completed at/by, proof → (interval) schedule next cycle
              → save → re-render that tab and the dashboard → toast ✅
```

### 3.3 Manager acknowledges

- **From the dashboard:** tick ✅ in *Needs attention* (the note column is saved as typed). The edit handler stamps `Ack At / Ack By / Ack Note`.
- **From the email:** each row has an *Acknowledge* link → `doGet` shows a confirmation page (no state change, so link scanners can't acknowledge) → the **Confirm** button calls `confirmAck(id, sig, note)` → HMAC check → same acknowledgement code path.

### 3.4 Escalation (part of the engine run)

See `features/F07-notifications-sla-escalation.md` for the timing rules.

## 4. Concurrency & idempotency

| Hazard | Guard |
|---|---|
| Engine and an employee edit at the same moment | `LockService.getScriptLock()` around every write path. An edit that can't get the lock in 25 s writes "⏳ busy, please re-enter" instead of being lost. |
| Trigger fires twice / engine run repeatedly | Cycle IDs are deterministic (`TaskID-yyyymmdd`); creation checks for the ID first. Interval tasks never get a second OPEN cycle. |
| Same email twice | Each cycle carries `Reminder Sent At`, `Manager Alert Sent At`, `Escalated At`, `Resolved Notice At`; a cycle is only added to a batch while its stamp is empty. The stamp is written and saved straight after a successful send. |
| Email send fails | The stamp stays empty and the Email Log records `FAILED`. The next run retries, so each email is sent once it succeeds. |
| Same completion twice (double trigger, double paste) | Completion requires `status = OPEN`; a second attempt reports "Already done" and changes nothing. |
| Same Ticket ID reused for a different cycle | Rejected: Ticket IDs are unique across the ledger. |

## 5. Deployment

1. Create a Google Sheet → Extensions → Apps Script → paste `dist/Code.gs`, and replace `appsscript.json` with `dist/appsscript.json`.
2. Run `setupSystem` (authorise) → menu **⚙️ Task System** appears → **Install triggers**.
3. Optional: Deploy → Web app (execute as me, anyone with a Google account) so the email Acknowledge links work. The URL is auto-detected.

Alternatively push with `clasp` from `src/` (`.clasp.json` points `rootDir` at `src`).

## 6. Non-goals (prototype scope)

- Reading the external portal to verify that a Ticket ID really exists (needs portal API access; the extension point is `validateProof_`).
- Reading WhatsApp group messages automatically (see F04 §5, roadmap).
- Per-user sheet permissions beyond warning-level protection (needs Google Workspace accounts for every employee).
