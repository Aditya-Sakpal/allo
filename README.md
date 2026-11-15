# Allo Innoware: Recurring Task Tracking, Pending Reports & Escalation

A Google Sheets + Apps Script system. Recurring tasks become due automatically. Each employee sees and closes their work **in one place**. The manager is emailed **the same day** when something is missed, and unresolved items are escalated to leadership automatically, **with no duplicate work or duplicate emails**.

- Design and reasoning: [`specs/`](specs/README.md). Start with the [plan](specs/00-plan.md), the [daily-task method](specs/features/F04-daily-task-completion.md) and the [SLA / escalation design](specs/features/F07-notifications-sla-escalation.md).
- Paste-ready code: [`dist/Code.gs`](dist/Code.gs) + [`dist/appsscript.json`](dist/appsscript.json). Source is in [`src/`](src/).
- Project site (Cloudflare): **https://allo-recurring-tasks.salesdesk-app.workers.dev**. It covers the overview, the SLA timeline and install steps, and serves the downloadable `Code.gs`. Deploy with `npm run deploy`, which rebuilds `dist/` and runs `wrangler deploy` on `site/`.
- Tests: 30 scenario tests (the assignment's Tests 1–8 plus 22 "try to break it" cases) that run **inside the Google Sheet** from the menu and locally with `npm test`.

---

## 1. Install (about 5 minutes)

1. Create a new Google Sheet → **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`dist/Code.gs`](dist/Code.gs).
3. Project Settings → tick **Show "appsscript.json"** → replace it with [`dist/appsscript.json`](dist/appsscript.json). This pins the timezone to Asia/Kolkata. Save.
4. Select the function **`setupSystem`** → **Run** → authorise. This builds every sheet and loads the demo scenario. Reload the spreadsheet: a **⚙️ Task System** menu appears.
5. **⚙️ Task System → Admin → Install triggers.** The engine then runs every 15 minutes and edits are processed instantly.
6. *(Optional, for one-click acknowledgement from email)* **Deploy → New deployment → Web app**, execute as **Me**, access **Anyone with a Google account** → Deploy. The URL is detected automatically.
7. To receive real emails: put real addresses in **Team** (employees) and **Config** (`MANAGER_EMAIL`, `LEADERSHIP_EMAILS`), then set `EMAIL_MODE = LIVE`. Until then every email is written to **Email Log** only, which is safe for demos. Gmail `+` aliases work well for testing (e.g. `you+manager@gmail.com`, `you+ceo@gmail.com`).

> Using clasp instead: `clasp create --type sheets --rootDir src`, then `clasp push`.

## 2. The sheets

| Sheet | Who | Purpose |
|---|---|---|
| **Start Here** | everyone | How to use it, per role |
| **Manager Dashboard** | manager | Team overview · Needs attention (✅ acknowledge, note, Waive) · All open · Recent completions |
| **👤 Employee A … D** | each employee | Their **only** working place: ACTION NEEDED (type the proof in the yellow cell) · COMING UP · COMPLETED |
| **Tasks** | admin | Task Master: owner, frequency (Daily / Every N days), proof type, cut-off, first due date, active, Health |
| **Team** | admin | Employees, emails, manager override |
| **Config** | admin | SLA times, recipients, working days, holidays, email mode, demo clock |
| **Cycles** | system | Ledger: one row per task occurrence, with who, when, proof, alert, acknowledgement and escalation stamps |
| **Activity Log / Email Log** | system | Full audit trail, including every email and whether it was sent |
| **Test Results** | system | PASS/FAIL + evidence from the automated test suite |

## 3. How each requirement is met

| Requirement | How |
|---|---|
| Recurring cycle logic | Schedule anchored on *First Due Date* + k × N days. Completing a cycle creates the next one immediately; daily tasks get one cycle per working day. Nobody calculates dates. |
| Employee pending report | Personal tab with what is due, the status (due today / overdue / manager alerted / escalated), what to do and the proof required. |
| Single-point update | The employee types the Ticket ID or update **into the yellow cell on that same row** and presses Enter. That edit closes the cycle, writes the ledger and audit log, refreshes the dashboard and schedules the next cycle. Nothing else is updated anywhere. |
| Ticket completion | Ticket ID is validated (format, at least one digit, never re-used) and becomes the completion record. |
| **Daily task (no ticket)** | The employee **pastes the exact update they posted on WhatsApp**. It is stored with a timestamp as verifiable evidence (min length, no copy of yesterday's text), and a **📤 Share on WhatsApp** link is generated for the write-first flow. See [F04](specs/features/F04-daily-task-completion.md). |
| Completion tracking | Who (editor identity), when, proof, on time vs late → ledger + Activity Log. |
| Manager visibility | Manager Dashboard: one row per person with Need action, Due today, Overdue, Awaiting ack, Escalated, On-time %. |
| Manager acknowledgement | ✅ checkbox (with note) on the dashboard, or a signed link in the email → confirmation page. Acknowledgement time, person and note are stored, and whether the task was later done is shown. |
| Escalation + SLA | Reminder at cut-off − 2 h → **manager email at cut-off + 30 min (same day)** → **leadership at 11:00 next working day if not acknowledged**, or **at cut-off time next working day if still not done** → one "resolved" notice. Weekends and holidays are skipped. See [F07](specs/features/F07-notifications-sla-escalation.md). |
| No duplicates | Deterministic cycle IDs, an "already done" guard, one stamp per email type per cycle (saved right after sending), and a script lock on every write path. Failed sends are retried once the outage ends, then never again. |

## 4. Live demo walkthrough (about 10 minutes, no waiting)

The **Demo clock** moves a simulated time (`Config → SIMULATED_NOW`) and runs the engine, so you can watch a whole day pass.

1. **⚙️ Task System → Admin → Reset demo data.** The clock is set to 10:00 on a working day. Each employee tab shows its task.
2. **Test 2 / Test 1:** open **👤 Employee B**, type `PRT-10234` in the yellow cell and press Enter. The row moves to COMPLETED, next week's cycle appears under COMING UP, and the dashboard shows the completion.
3. **Test 3:** open **👤 Employee A**, type `done` → ❌ with the reason. Paste a real update (e.g. `Ops update: 42 orders dispatched, 3 delayed by courier, stock OK`) → ✅. Use the **📤 Share on WhatsApp** link in COMPLETED.
4. **Test 4:** before step 3, or for another task, **Demo clock → Jump to manager-alert time**. The manager alert appears in **Email Log** (or in the inbox in LIVE mode) the same evening.
5. **Test 5:** on **Manager Dashboard → Needs attention**, type a note, tick **✅ Ack**. Acknowledged at / by / note are recorded. Or use the Acknowledge link in the email.
6. **Test 6:** leave one item unacknowledged → **Demo clock → Jump to ack deadline** (11:01 next working day) → leadership escalation in the Email Log. For an acknowledged-but-open item, **Jump to next cut-off + 1 min** → escalation "still not completed".
7. **Test 8:** **Run engine now** as many times as you like. The Email Log and Cycles don't grow.
8. **Demo clock → Back to real time** when finished.

## 5. Automated proof

**⚙️ Task System → 🧪 Automated tests**, then run each of the five items in turn: core tests (T1–T8), then robustness tests 1/4 to 4/4 (B01–B22). Google stops any single script run after 6 minutes, so the suite is split into parts of at most 8 tests (2–4 minutes each). The suite runs on a temporary copy, so your data is never touched. Each part merges its PASS/FAIL and evidence lines into **Test Results**, which always lists all 30 tests; ones not run yet say which menu item runs them.

Locally:

```bash
npm test             # 30 suite tests + 14 integration checks against a strict Apps Script mock
npm run test:evidence  # same, printing the evidence for every test
npm run build        # regenerate dist/
```

## 6. Known limits (prototype scope)

- A Ticket ID is checked for format and uniqueness, not looked up on the portal (extension point: `validateProof_`).
- WhatsApp content can't be read automatically without a WhatsApp Business group integration; the pasted text is the evidence ([F04 §2](specs/features/F04-daily-task-completion.md)).
- On consumer Gmail accounts Google may hide editors' emails from scripts. "Completed by" then records the tab owner ("Employee A (via own tab)").
- Timing granularity is the 15-minute trigger: the "18:30" alert is sent by 18:45 at the latest.
