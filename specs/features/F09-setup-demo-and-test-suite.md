# F09 — Setup, Demo Clock & In-Sheet Test Suite

**Assignment link:** §7 "Working Demonstration / Test Cases"
**Code:** `src/10_Setup.js`, `src/11_Main.js`, `src/12_Tests.js`

## Setup (`setupSystem`)
- Creates any missing sheet and adds any missing header or config key. **Non-destructive**, so it is safe to re-run as a "repair".
- Applies formats: plain text for IDs and inputs, date formats, checkboxes, dropdowns (Frequency, Proof Type, Assigned To ← Team).
- Seeds the assignment's dummy scenario if Team and Tasks are empty (Employee A–D; Daily / 7 / 15 / 30-day tasks).
- Aligns the spreadsheet timezone with the script timezone and stores the spreadsheet ID for the web app.

## Menu `⚙️ Task System`
- Run engine now
- Demo clock ▸ start of demo day (10:00, next working day if needed) · +1 hour · +3 hours · +1 day · jump to manager-alert time · jump to ack deadline · jump to next cut-off + 1 min (resolve deadline) · back to real time. Each one moves `SIMULATED_NOW` and runs the engine, so Tests 4–6 can be demonstrated in minutes.
- Run automated tests ▸ core (T1–T8) · robustness (B01–B22) · all
- Admin ▸ Setup / repair · Install triggers · Reset demo data (also sets the demo clock)

## Test suite
- Runs on a **temporary sandbox spreadsheet** (created, then trashed), so real data is never touched. Results go to `Test Results` in the real sheet: ID, Test, PASS/FAIL, evidence lines, duration.
- Uses a test clock and a capture mailer, so it runs in about a minute rather than several days.
- The same code runs locally: `npm test` loads `src/` into a strict Google Apps Script mock.
- A time guard stops before the 6-minute Apps Script limit and marks the remaining tests "NOT RUN — run this group separately".
