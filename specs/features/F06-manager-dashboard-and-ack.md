# F06 — Manager Dashboard & Acknowledgement

**Assignment link:** §4 "Manager visibility", "Manager acknowledgement"; Test 5
**Code:** `src/07_Views.js` (`renderDashboard_`), `src/08_Edits.js`, `src/09_WebApp.js`

## Dashboard sections
1. **Team overview:** one row per employee with Open, Due today, Overdue, Awaiting your ack, Escalated, Completed (7 d), On-time % (30 d). This answers "who is behind?" at a glance.
2. **Needs attention:** every OPEN overdue cycle, plus alerted cycles closed in the last 7 days (so the manager can see that "it got done after my alert"). Columns: Employee, Task, Due, Task status, Overdue by / Done at, Manager alerted, **✅ Acknowledged** (checkbox), **Manager note**, Ack at, Escalation (deadlines or what fired), **Action** (`Waive (leave / holiday)`).
3. **All open tasks:** the full list, including items not yet due.
4. **Recent completions:** who, when and the proof (the daily update text is visible here).
The header shows "Engine last ran …" and any config warnings.

## What the manager is expected to do after an alert email (Test 5)
1. **Contact the employee** and find out why (forgot, blocked, on leave).
2. **Acknowledge** before **11:00 next working day**: click *Acknowledge* in the email (confirmation page with an optional note) **or** tick ✅ on the dashboard, adding a note such as "Blocked on portal access, done by 3 pm".
3. **Get it resolved** by the **cut-off time on the next working day**, or **Waive** the cycle if the employee is on leave or it's a holiday. Waiving closes the cycle and records the reason.

If (2) or (3) doesn't happen, leadership is informed automatically (F07).

## How the system knows
| Question | Field |
|---|---|
| Did the manager review it? | `Ack At`, `Ack By`, `Ack Note` (Cycles) + ACKNOWLEDGED in the Activity Log |
| Was the task done afterwards? | `Status`, `Completed At` vs `Manager Alert Sent At`, shown as "✅ Done <time> (after alert)" |
| Was it escalated anyway? | `Escalated At`, `Escalation Reason` |

## Rules
- Acknowledgement is a fact: unticking is ignored and the tick is restored on the next render.
- A note typed before ticking is kept (saved as typed).
- Acknowledging an overdue item *before* its alert goes out suppresses that alert (the manager already knows); the resolution deadline still applies.
- Waive is only available on OPEN rows. It also counts as an acknowledgement, and it schedules the next interval cycle.
- Email links are HMAC-signed per cycle. GET only shows a confirmation page, so link scanners can't acknowledge; a forged or edited link is rejected.

## Acceptance
- **T5:** acknowledgement via the dashboard and via the web confirmation; forged signature rejected; GET has no side effects; later completion is reflected; no leadership escalation after ack + completion.
- **B11:** waive. **B13:** untick ignored. **B20:** link security.
