# WP-03 - S02: walk-through part 2, then the models on both engines

- Date: 10-Sep-2026. Depends on: Codex review of S01. Increment 1, stage S02 in docs/STAGES.md.

## Scope
- Part A (documents, first half of the session): docs/walkthroughs/W2_spine_part2.md covering DISC_03 sections 3 to 8 in the W1 format (one row is, columns with units, links, sample row, invariants), under 300 lines, including: rate_record types and comparable flag; package, rfq, quote, commitment; cost ledger authority rules; income rows kept apart; variation axes and variation_cost_budget; manufacture; progress_record, measured_progress, inspection_request, evidence_check, defect, evidence.
- Part B (code): Django models for W1 and W2 in core/ (organisation, cross-cutting) and projects/ (scheme register and everything scheme-scoped); migrations; admin registration; django-simple-history on every model; constraints in the models (unique keys, check constraints for shares and quantities, PROTECT on foreign keys); model clean() for I-3 (same-company foreign keys) and I-2 (no clamping).
- Tests: migrations run on SQLite and PostgreSQL in the suite; every table appears in admin; a change appears in history; I-1, I-2, I-3, I-5, I-6, I-7 each have a failing-then-passing test on synthetic rows; tests/test_no_raw_sql.py stays green; the dump-load round trip from S00 extended to the new tables.

## Not in scope
- Screens, imports, exports, the BOQ import (S04), login and permissions on views (S03). Real data.

## Acceptance
- Fresh SQLite and PostgreSQL both migrate; every table in admin; the invariant tests pass; hygiene green; handover records the engines actually run.

## Estimate (DISC_05 basis, uncalibrated)
- Agent active 3 hours; human review 2 hours (Alaa and Mo). One session of about 45 tool turns; a retry is likely because the model set is large, so plan for two sessions.

## Governance
- W-13: synthetic data only. Codex review is the second review. Mo's line stays blank.
