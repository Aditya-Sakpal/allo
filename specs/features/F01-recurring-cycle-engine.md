# F01 — Recurring Cycle Engine

**Assignment link:** §4 "Recurring cycle logic", Test 1
**Code:** `src/03_Schedule.js`, `src/05_Cycles.js` (`ensureCycles_`, `scheduleNextInterval_`)

## Goal
Every task becomes due again automatically. Nobody ever calculates a next due date by hand.

## Rules

### R1 — Slots are anchored, not chained
- Interval task (Every N days): slot *k* = `First Due Date + k·N` calendar days.
- Due At = slot date at the task's Due Time (or DEFAULT_DUE_TIME). If the slot falls on a non-working day, Due At moves to the next working day, but the next slot is still computed from the *nominal* date, so the schedule never drifts.
- Daily task: one slot per working day.

### R2 — Interval tasks: exactly one OPEN cycle
- When a cycle closes (DONE, LATE_DONE or WAIVED), the next cycle is created **in the same transaction**. The employee immediately sees it under *Coming up*.
- Window Opens of the next cycle = cut-off of the cycle it replaces. Work for the next period can't be submitted before the current period ends.
- **Late completion:** slots whose cut-off already passed at completion time are written as `SKIPPED` (with a note). The new OPEN cycle is the first slot whose cut-off is still in the future. One late piece of work closes one cycle; it does not create a backlog of duplicate work. The miss is still recorded honestly.

### R3 — Daily tasks: one cycle per working day
- The engine creates `TaskID-yyyymmdd` for **today** if today is a working day and the date is on or after First Due Date.
- A new daily task added after today's cut-off starts on the next working day, so it isn't instantly overdue.
- Earlier unfinished days stay OPEN (overdue) until completed or waived; they keep following the escalation path.

### R4 — First cycle / reactivation / repair
If an active, valid interval task has no OPEN cycle (new task, reactivated task, or manual damage), the engine creates one at the first slot after the last known slot (or the first slot from First Due Date). It rolls forward past slots already gone, **silently**. A pause is not reported as a series of misses.

### R5 — Master-data changes
| Change | Effect |
|---|---|
| Task deactivated | OPEN cycles become `WAIVED` ("Task deactivated"); no more alerts |
| Task reactivated | R4: fresh cycle in the future, no backfilled alerts |
| Assigned To changed | OPEN cycles move to the new employee (logged REASSIGNED) |
| Frequency / Due Time changed | Applies from the next cycle; the current due date is kept |
| Invalid row (unknown employee, bad frequency, duplicate ID, …) | Row skipped and Health shows the reason; existing cycles are untouched |

## Acceptance
- **T1:** completing a 7-day cycle creates the next one at +7 days with the correct window; 15- and 30-day tasks step +15/+30; repeated engine runs keep exactly one OPEN cycle.
- **B04:** late completion records skipped slots and opens the next future slot.
- **B05/B06/B09/B14:** deactivation, reassignment, invalid rows, weekend shift.
