# F03 — Ticket-Based Completion

**Assignment link:** §4 "Ticket-based task completion"; Tests 1–2
**Code:** `src/05_Cycles.js` (`validateProof_`, `completeCycle_`)

## Goal
For portal tasks, the Ticket ID *is* the completion. Entering a valid, unused Ticket ID closes the cycle.

## Validation (in order)
1. Not blank.
2. Matches `Config.TICKET_ID_PATTERN` (default: 3–40 chars, letters, digits and `# / _ . -`, at least one digit, no spaces). Rejects `done`, `ok`, `NA`, `12`, HTML or script fragments.
3. **Unique across the whole ledger** (case-insensitive). Re-using last week's ticket is rejected with the date it was used.
4. The cycle window must be open (it can't close the next period early).

## On success
`Status = DONE` (cut-off not passed) or `LATE_DONE`, with `Completed At = now`, `Completed By = editor` and `Proof = ticket`. For interval tasks, the next cycle is scheduled (F01 R2).

## Extension point
`validateProof_` is the single place to add a portal lookup (for example `UrlFetchApp` to the ticketing API) that confirms the ticket exists, is assigned to the employee and is in a closed state.

## Acceptance
- **T1/T2:** valid ticket → DONE and the next cycle exists.
- **B01:** invalid formats are rejected with a clear message and the cycle stays OPEN.
- **B02:** a re-used ticket is rejected.
- **B03:** submitting before the window opens is rejected.
