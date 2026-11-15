# F05 — Completion Tracking & Audit Trail

**Assignment link:** §4 "Completion tracking"
**Code:** `src/05_Cycles.js`, `src/04_Store.js`

## Captured on every completion
| Field | Source |
|---|---|
| Completed At | Engine clock at the moment of the edit |
| Completed By | `Employee (editor-email)`. If the editor is someone else: `editor (on behalf of Employee)`. If Google doesn't expose the email (consumer accounts): `Employee (via own tab)` |
| Proof | Ticket ID or the update text |
| Status | `DONE` if at or before Due At, else `LATE_DONE` |

## Audit
- Every state change appends to **Activity Log** (who, what, which cycle, details). Nothing is overwritten.
- Every email (sent, log-only or failed) appends to **Email Log** with its recipients and cycle IDs.
- Rejected proofs are logged as `PROOF_REJECTED`, so attempts to game the system are visible.

## Cycle closing → next cycle
The completion closes the current cycle and (interval tasks) schedules the next one in the same locked transaction. Daily tasks get their next cycle from the engine on the next working day.

## Acceptance
- **T1/T2/T3:** fields populated; Activity Log has exactly one COMPLETED entry per cycle.
