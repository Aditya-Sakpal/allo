# F04 — Daily Task Completion (no Ticket ID): the proposal

**Assignment link:** §4 "Daily-task completion method"; Test 3
**Code:** `src/05_Cycles.js` (`validateProof_` Daily branch), `src/07_Views.js` (share link)

## 1. The problem
The daily task is "share an operations update in the WhatsApp group". It produces no ticket, and the WhatsApp group can't be read by Google Sheets. A plain "Done ✓" checkbox records a claim, not evidence.

## 2. Options considered

| Option | Employee effort | Evidence quality | Verdict |
|---|---|---|---|
| A. Tick a "Done" checkbox | 1 click | None: self-declared | ❌ too easy to game |
| B. Upload a screenshot (Form/Drive) | 4–6 taps, a second app | Good, but not machine-checkable | ❌ effort, and it's a second place to update |
| C. WhatsApp Business API / bot reads the group | 0 | Excellent | 🟡 roadmap: needs Meta business verification and a bot in the group |
| **D. Paste the update text into the sheet** (the same text that was posted) | **1 paste** | **Good: the exact content is stored with a timestamp and can be matched against the group** | ✅ **chosen** |

## 3. The method (what the employee does)

**Either direction works. Both take one paste.**

1. **Post first:** post the update in the WhatsApp group → long-press → Copy → paste into the yellow cell of the *Daily Operations Update* row → Enter.
2. **Write first:** type the update in the yellow cell → Enter → in *COMPLETED*, tap **📤 Share on WhatsApp** → pick the group → Send. The link opens WhatsApp with the text pre-filled.

## 4. What the system does
- Validates the text: at least `DAILY_UPDATE_MIN_CHARS` (20) characters, so "done" or "ok" are rejected, and **not identical to a previous day's update**, so copy-paste of yesterday is rejected.
- Records `Completed At` (server time), `Completed By` (editor) and the full text in `Proof`. On time if before 18:00, otherwise `LATE_DONE`.
- Shows the text on the Manager Dashboard (*Recent completions*). The manager can compare it with the group in seconds.
- Builds a `https://wa.me/?text=…` share link from the stored text.
- No text by the cut-off → the standard escalation path (F07).

## 5. Why this is the most practical choice
- **Minimum effort:** the employee writes the update anyway. Pasting it once is the whole "reporting" step.
- **Auditable:** the record contains *what* was shared and *when*, not just a tick.
- **Fits the single-point rule:** it uses the same yellow cell as the portal tasks, so there's one habit to learn.
- **Upgrade path:** when a WhatsApp group bot (option C) is available, it calls the same `completeCycle_()` with the message text. The sheet, SLA and escalation logic don't change.

## Acceptance
- **T3:** a daily cycle is created per working day (none on non-working days); a too-short update is rejected; a valid one completes with the text stored and a share link rendered; an identical repeat the next day is rejected.
