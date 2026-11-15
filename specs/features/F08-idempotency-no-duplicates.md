# F08 — Idempotency: No Duplicate Cycles, Completions or Emails

**Assignment link:** Test 8
**Code:** throughout; see `specs/01-architecture.md` §4

## Guarantees
| Thing | Guarantee | Mechanism |
|---|---|---|
| Cycles | At most one per (task, slot); interval tasks have at most one OPEN | Deterministic Cycle ID; existence check before create; OPEN check |
| Completions | At most one per cycle | `status === OPEN` precondition; second attempt → "Already done" |
| Emails | At most one per (cycle, event type) once delivered | Stamp columns; batch send → stamp → immediate save |
| Concurrent runs | Serialised | Script lock on every entry point |
| Failed email | Retried until delivered, then never again | Stamp only on success; FAILED row in Email Log |

## Test scenario (T8)
1. Engine ×10 at the same instant → cycle count unchanged.
2. Move past the cut-off, engine ×5 → exactly one manager alert email.
3. Mail outage at the escalation time → FAILED logged, cycle not stamped → outage ends → exactly one escalation → engine ×3 more → still one.
4. The same completion submitted twice → one COMPLETED log entry, one next cycle.
