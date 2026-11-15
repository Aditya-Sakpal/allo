# CLAUDE.md — how development works in this repo

## Mission
Build and maintain a **Google Sheets + Apps Script** system that tracks recurring tasks, captures completion in **one place**, alerts the manager **the same day**, and escalates unresolved items to leadership, **without duplicate work or duplicate emails**. The source of truth for behaviour is [`specs/`](specs/README.md). Read `specs/00-plan.md` and the relevant `specs/features/F0x-*.md` before touching code.

## Mandatory skills (every session)
This project uses the **superpowers** process skills and **impeccable** for anything user-facing. Invoke them with the Skill tool; don't paraphrase them from memory.

| When | Skill |
|---|---|
| Start of any new feature / behaviour change | `superpowers:brainstorming` → then update the spec in `specs/` |
| Turning a spec into work | `superpowers:writing-plans` (plan goes in `specs/00-plan.md` or a new `specs/features/F1x-*.md`) |
| Writing any code | `superpowers:test-driven-development`: write or extend the test in `src/12_Tests.js` first and watch it fail |
| Any failing test / bug / surprising behaviour | `superpowers:systematic-debugging`: root cause before fix |
| Independent parallel work | `superpowers:dispatching-parallel-agents` / `superpowers:subagent-driven-development` |
| Before saying "done" / "fixed" / "passing" | `superpowers:verification-before-completion`: run `npm test` and quote the result |
| After a meaningful change | `superpowers:requesting-code-review` |
| **Any front end**: employee tabs, Manager Dashboard layout, email HTML, the acknowledgement web page | `impeccable` (or `impeccable:impeccable`): hierarchy, copy, colour, accessibility, mobile |

## Commands
```bash
npm test          # loads src/*.js into the strict GAS mock and runs the full suite (T1–T8, B01–B22, node checks)
npm run build     # writes dist/Code.gs + dist/appsscript.json (single paste-ready bundle)
```
`npm test` must be green before any commit. Never weaken an assertion to make it pass. Fix the code, or change the spec first and say why.

## Architecture rules (non-negotiable)
1. **Dependency injection:** core functions take `ctx` (`ss`, `cfg`, `now()`, `mailer`). Only entry points in `11_Main.js` / `09_WebApp.js` may touch `SpreadsheetApp.getActive…`, `LockService`, or real time. This is what makes the suite runnable in Node and on a sandbox.
2. **Every write path is locked** (`withLock_`) and **idempotent**: deterministic Cycle IDs, the `status === OPEN` precondition, a stamp column for every email type. Running anything twice must be harmless. Add a test that proves it.
3. **The ledger (`Cycles`) is the truth; views are disposable.** Employee tabs and the dashboard are fully re-rendered from the model. Never store state only in a view.
4. **Batch I/O:** read a table with one `getValues()`, write with one `setValues()`. No per-cell loops against the Sheets API.
5. **Columns are found by header name** (`readTable_`), never by hard-coded index, for input sheets.
6. **Top level of every `src/*.js` holds only literals and function declarations.** Apps Script loads files in arbitrary order.
7. **Times:** use `ctx.now()` and the helpers in `03_Schedule.js`. Script and sheet timezone = `Asia/Kolkata`. Never `new Date()` in core logic.
8. **User text into cells:** plain-text format (`@`) for IDs and input, so Sheets doesn't auto-convert `11-12-2026` to a date.
9. **Fail soft for data, loud for code:** bad config or task rows → warnings in Health / dashboard and the engine continues. Programming errors → throw, so the test suite catches them.
10. **Emails:** batch per recipient per event per run; stamp and save immediately after a successful send; never send in a loop without a stamp check.

## Coding conventions
- Plain Apps Script V8 JavaScript in `src/NN_Module.js` (clasp pushes `.js` as `.gs`). No modules, no `import`.
- Private helpers end with `_` (Apps Script hides them from the Run menu). Public entry points (triggers, menu, web app) have no underscore.
- Small, pure functions where possible (schedule maths, validation, labels). Keep functions under ~60 lines.
- Names: `camelCase` functions, `UPPER_SNAKE` constants, sheet and column names only via `SHEETS` / `*_FIELDS` in `00_Schema.js`.
- User-facing messages are plain English, say what to do next, and never include stack traces.
- No secrets in code. The HMAC secret lives in Script Properties (`ACK_SECRET`) and is generated on first use.
- Comments explain **why**, not what.

## Testing conventions
- Tests live in `src/12_Tests.js` (`T*` = assignment tests, `B*` = break-it tests). Each one resets the sandbox fixture, sets the clock explicitly (`sb.at('2026-11-16 18:31')`), and asserts with `t.eq / t.ok`, which also produce the evidence text shown in the `Test Results` sheet.
- Every bug fix gets a regression test first.
- The mock (`test/gas-mock.js`) is deliberately strict: unknown getters throw, ranges out of bounds throw, and `setValues` size mismatches throw. If you need a new Apps Script API, add it to the mock **faithfully** (check the Apps Script reference) rather than loosening the mock.

## Definition of done for any change
1. Spec updated (`specs/`).
2. Test written first and seen failing, then passing.
3. `npm test` green, `npm run build` regenerated `dist/`.
4. UI changes reviewed with `impeccable`.
5. `README.md` demo steps still accurate.
