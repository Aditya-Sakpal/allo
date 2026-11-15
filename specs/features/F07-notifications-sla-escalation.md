# F07 — Notifications, SLA & Leadership Escalation

**Assignment link:** Tests 4, 6, 7
**Code:** `src/06_Notify.js`, `src/03_Schedule.js` (`ackDeadline_`, `resolveDeadline_`)

## 1. The timeline (defaults; all configurable)

| When | Event | To | Condition |
|---|---|---|---|
| **Cut-off − 2 h** (16:00) | Reminder | Employee | Cycle open, window open, not reminded yet |
| **Cut-off** (18:00) | Task becomes *overdue* on views | — | Not done |
| **Cut-off + 30 min** (18:30, **same day**) | **Manager alert** | Employee's manager | Still not done, not alerted yet |
| **11:00 next working day** | **Leadership escalation**: "manager has not acknowledged" | Leadership (cc manager) | Alerted, **no acknowledgement** |
| **Cut-off time next working day** (18:00) | **Leadership escalation**: "still not completed" | Leadership (cc manager) | Alerted, acknowledged or not, **still open** |
| When an escalated item closes | Resolved notice | Leadership (cc manager) | Once |

Weekends and holidays (`WORKING_DAYS`, `HOLIDAYS`) are skipped by both deadlines. A Friday or Saturday alert gets its deadlines on the next working day.

**Granularity:** the engine runs every 15 minutes, so each event fires at the first run after its time (the 18:30 alert by 18:45 at the latest). A task whose cut-off + grace falls after 23:44 can't be guaranteed a same-day alert, so its Health shows a warning.

## 2. Reasoning
- **Reminder at −2 h:** prevention is cheaper than escalation. Two hours is enough to finish a portal task, and it fires only once per cycle.
- **Manager at +30 min, same day:** the assignment requires a same-day alert. The 30-minute grace absorbs someone finishing at 18:05 without a false alarm, and the manager still hears about it before the day ends.
- **Ack by 11:00 next working day:** gives the manager the evening plus the first working hours to see the email. If they haven't even acknowledged by late morning, the process (not just the task) has failed, and leadership should know.
- **Resolve by the next cut-off:** a full extra working day to fix the underlying problem. Past that, the task is late by more than a day, a real operational risk worth leadership attention, even if the manager is engaged.
- **Two independent triggers:** an engaged manager isn't penalised for a stubborn task until the resolve deadline, and an unresponsive manager is surfaced quickly.

## 3. Avoiding unnecessary or repeated emails
1. **One email per event per cycle**, enforced by the stamps (`Reminder Sent At`, `Manager Alert Sent At`, `Escalated At`, `Resolved Notice At`).
2. **Batching:** one email per recipient per event type per engine run lists all affected tasks (a manager with three late tasks gets one email).
3. **No nagging loops:** after escalation there are no repeat emails; the dashboard carries the state. The only follow-up is the single "resolved" notice that closes the loop for leadership.
4. **Suppression:** a cycle acknowledged before its alert is due gets no alert. Completed or waived cycles get no further emails, except "resolved" if they had been escalated.
5. **Failure-safe:** a failed send leaves the stamp empty and logs `FAILED`; the next run retries. Stamps are saved immediately after each batch, so a later crash can't cause a resend.
6. **Quota-aware:** in LIVE mode the remaining MailApp quota is checked before sending.

## 4. Email content
- **Manager alert:** a table of Employee, Task, Cut-off, Overdue by and Required proof, with an **Acknowledge** link per row. It includes "What we need from you" with the exact ack and resolve deadlines, and a link to the dashboard.
- **Leadership escalation:** the same table plus Manager alerted at, Acknowledged? (with the note) and Reason.
- **Reminder:** the employee's tasks due soon, with a link to their own tab.

## Acceptance
- **T4:** reminder at 16:00 (not 15:59); manager alert at 18:31, none at 18:29; same date as the cut-off; one batched email.
- **T6:** (a) unacknowledged → leadership at 11:01, not at 10:59; (b) acknowledged but still open → leadership at 18:01 next day; resolved notice exactly once.
- **T7:** deadline maths across weekends for Mon–Sat and Mon–Fri calendars; no escalation on a non-working day.
- **T8:** see F08.
