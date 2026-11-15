# Specs index

| Doc | What it covers |
|---|---|
| [00-plan.md](00-plan.md) | Phases, key decisions, definition of done, risks |
| [01-architecture.md](01-architecture.md) | Components, runtime flows, concurrency, deployment |
| [02-data-model.md](02-data-model.md) | Every sheet and column, cycle state machine |
| [03-test-plan.md](03-test-plan.md) | Assignment tests T1–T8, break tests B01–B22 |
| [features/F01](features/F01-recurring-cycle-engine.md) | Recurring cycle engine |
| [features/F02](features/F02-employee-pending-view.md) | Employee pending view & single-point update |
| [features/F03](features/F03-ticket-completion.md) | Ticket-based completion |
| [features/F04](features/F04-daily-task-completion.md) | **Daily task completion method (proposal)** |
| [features/F05](features/F05-completion-tracking-audit.md) | Completion tracking & audit |
| [features/F06](features/F06-manager-dashboard-and-ack.md) | Manager dashboard & acknowledgement |
| [features/F07](features/F07-notifications-sla-escalation.md) | **Notifications, SLA & escalation (proposal)** |
| [features/F08](features/F08-idempotency-no-duplicates.md) | No duplicate cycles/completions/emails |
| [features/F09](features/F09-setup-demo-and-test-suite.md) | Setup, demo clock, in-sheet test suite |

Rule: **specs change first, code second.** A behaviour change that isn't reflected here is a bug.
