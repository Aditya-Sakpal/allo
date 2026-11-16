# F02 — Employee Pending View & Single-Point Update

**Assignment link:** §4 "Employee pending report", "Single-point updating"; Test 2
**Code:** `src/07_Views.js` (`renderEmployeeTab_`), `src/08_Edits.js` (`handleEmployeeEdit_`)

## Goal
Each employee has **one place** that shows what needs action and accepts the completion proof. They never update a second sheet.

## Layout — tab `👤 <Employee>`

```
Row 1  <Employee> — my recurring tasks
Row 2  How it works: type the Ticket ID or paste your update in the yellow cell, press Enter. Done.
Row 3  As of <time> · N need action · M overdue
Row 5  ACTION NEEDED
Row 6  [A hidden: Cycle ID] | Task | Status | Proof required | ✍️ Enter Ticket ID / update here | Result | What to do | Cut-off | Frequency
Row 7+ one row per OPEN cycle whose window is open (sorted by cut-off). Input (col E) and feedback (col F) sit next to the task so they fit on a phone screen.
       COMING UP  — open cycles whose window hasn't opened yet (read-only)
       COMPLETED — last RECENT_DAYS days: result, time, 📤 WhatsApp share link (daily), proof
```

Status wording: `Due in N days`, `Due today 18:00`, `OVERDUE by 2h 10m`, `OVERDUE — manager alerted`, `ESCALATED to leadership`.

## Behaviour
1. The employee types or pastes into column E on an ACTION NEEDED row and presses Enter.
2. The installable `onEdit` reads the Cycle ID from column A of the **same row** (what the employee sees is what gets updated).
3. Validation (F03/F04). On failure, column F shows `❌ <reason>` and the typed text stays, so they can fix it.
4. On success the cycle closes and the tab re-renders: the row moves to COMPLETED, and an interval task's next cycle appears under COMING UP. The dashboard refreshes too.
5. Pasting several rows at once is supported; every row in the pasted block is processed.

## Guard rails
- Edits outside column E, or on title/header rows, are ignored.
- The Cycle ID must belong to this tab's employee; otherwise ❌ "This row belongs to …".
- A stale row (the cycle is already closed) gives a friendly "Already done …" and a re-render.
- If the lock is busy for more than 25 s: `⏳ System busy — please press Enter on the cell again`.
- **Refresh race:** if the engine redrew the tab between the keystroke and the handler (the cell no longer holds what was typed), nothing is applied, because the row may now be a different task. The user is asked to re-enter (B21).
- Only the edited employee's tab and the dashboard are re-rendered after an edit, which keeps response time low.
- All user text is written as plain text (`@`), and any text starting with `=`, `+`, `-` or `@` is written with a leading apostrophe (`textCell_`), because a script write turns `=…` into a formula even in a plain-text cell (found by B22 on a real sheet). Readers strip that apostrophe again (`fromCell_`), so the result is identical whether Sheets consumes it or keeps it. Only the system's own share link is a formula (B22; `npm test` runs the suite under both apostrophe behaviours).
- Input cells are plain-text formatted, so `11-12-2026` stays text and is not turned into a date.
- The tab has warning-level protection on everything except the input column.
- Tabs are created for new active employees and deleted for removed or inactive ones. They are views only, so no data is lost.

## Acceptance
- **T2:** an edit on the employee tab completes the cycle and updates the ledger, the Activity Log and the dashboard; replaying the same edit doesn't duplicate anything.
- **B07:** a multi-row paste completes each row; an edit in another column is ignored.
- **B08:** a stale or unknown row is handled gracefully. **B16:** a deleted tab is regenerated.
