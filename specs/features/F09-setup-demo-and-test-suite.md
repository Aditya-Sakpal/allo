# F09 — Setup, Demo Clock & In-Sheet Test Suite

**Assignment link:** §7 "Working Demonstration / Test Cases"
**Code:** `src/10_Setup.js`, `src/11_Main.js`, `src/12_Tests.js`

## Setup (`setupSystem`)
- Creates any missing sheet and adds any missing header or config key. **Non-destructive**, so it is safe to re-run as a "repair".
- Applies formats: plain text for IDs and inputs, date formats, checkboxes, dropdowns (Frequency, Proof Type, Assigned To ← Team). Every dropdown **warns** (red corner) rather than rejects: a rejected value never reaches the engine, so the Health column could not explain it, and Sheets also refuses scripted writes into a rejecting cell (found by B09 on a real sheet).
- Seeds the assignment's dummy scenario if Team and Tasks are empty (Employee A–D; Daily / 7 / 15 / 30-day tasks).
- Aligns the spreadsheet timezone with the script timezone and stores the spreadsheet ID for the web app.

## Menu `⚙️ Task System`
- Run engine now
- Demo clock ▸ start of demo day (10:00, next working day if needed) · +1 hour · +3 hours · +1 day · jump to manager-alert time · jump to ack deadline · jump to next cut-off + 1 min (resolve deadline) · back to real time. Each one moves `SIMULATED_NOW` and runs the engine, so Tests 4–6 can be demonstrated in minutes.
- Automated tests ▸ core (T1–T8) · robustness 1/4 (B01–B06) · 2/4 (B07–B12) · 3/4 (B13–B17) · 4/4 (B18–B22)
- Admin ▸ Setup / repair · Install triggers · Reset demo data (also sets the demo clock)

## Test suite
- Runs on a **temporary sandbox spreadsheet** (created, then trashed), so real data is never touched. Results go to `Test Results` in the real sheet: ID, Test, PASS/FAIL, evidence lines, duration.
- Uses a test clock and a capture mailer, so it runs in minutes rather than several days.
- The same code runs locally: `npm test` loads `src/` into a strict Google Apps Script mock.
- **Parts, not one run.** Google stops any single run after 6 minutes, and in a real spreadsheet each test takes about 30 s (sheet creation, protection, rendering). So the suite is split into parts of at most 8 tests (`testParts_()`), one menu item each.
- **Results accumulate.** Each run merges its results into `Test Results`, which always lists all 30 tests in suite order. Tests not run yet say which menu item runs them. The closing pop-up shows this part's score and the whole suite's.
- A 4-minute time guard inside each part marks anything left as "NOT RUN" instead of letting Google kill the run.
- A run that Google kills never reaches its clean-up, so every run first trashes leftover sandbox files by name.
