# CW change log

- One line per change to a planning document. Newest first. Decisions are logged in DECISIONS.md; this file records which files changed and why.

## 09-Sep-2026 - Tony Project discovery, second package: product and delivery map, estimates, re-cut (Alaa via Codex)
- Added docs/discovery/DISC_05_Estimates_and_Usage_Ledger_v0.md: agent active hours separate from human review hours per increment, two calendar scenarios, token estimate with explicit low / base / high assumptions, API-equivalent cost from rates verified on the official platform pricing page, Max subscription treated as unbilled consumption with unverified limits left editable, usage ledger for calibration. The 231 to 320 hour figure is withdrawn with the reason.
- Added docs/maps/tony_project_map.html: interactive product mind map, data-flow map and build roadmap with estimates, ledger and governance views; standalone copy committed here.
- docs/STAGES.md re-cut against DISC_04: 31 stages S00 to S30 in nine increments; later stages S31 to S33 parked; old-to-new map at the foot.
- DISC_01 corrected after management review: summary actuals in F1 and F3 distinguished from the missing transaction-level ledger, PO, GRN, programme and inspection data; external link names distinguished from proof of contents; the F1 / F3 cloning wording softened to common-template origin.
- DISC_04 corrected: SQLite to PostgreSQL migration explained and validated by test matrix, model-level constraints, dump-load comparison; permissions and integrity designed from Increment 1; G1 replaced by the exact register text and the substantive dependency that remains.
- DECISIONS.md: dated lines under D-06, D-07, W-10 and D-02 recording what each now blocks after Alaa's authorisation of incremental building. No LOCKED line rewritten; Mo's signature not written.
- Added docs/packages/WP-01_S00_bootstrap.md: first bounded implementation package for Codex review.
- Codex review corrections applied before handover: (1) DISC_03 section 8 separates measured progress, internal verification, the inspection cycle and the QS application decision; the universal "only accepted feeds progress" and "submitted does not count" rules are removed and replaced by configurable evidence_checks with a recorded QS override; acceptance is never deemed; measured_progress supersedes rather than sums so cumulative installed cannot double count; (2) DISC_03 section 5 makes the imported ledger row the authority and a matching invoice the same cost, with tests for partial accrual reversal, partial payment and receipt, credits and allocations; (3) DISC_03 section 6 splits VO selling lines from an internal variation_cost_budget seeded without overhead recovery or profit, models instruction, commercial, valuation and site completion independently, and keeps work before price agreement trackable; (4) DISC_02, DISC_04 gates, STAGES.md S10, S12, S15, S17, S20, S21 and the map updated to match; (5) W-13 records the bounded execution exception without asserting employer clearance, deciding Mo's split or signing for him; (6) DISC_05 cites the official Max usage guidance as read by Codex and states that API-equivalent figures are not a subscription bill.

## 09-Sep-2026 - Tony Project discovery, first work package (Alaa via Codex)
- Added docs/discovery/DISC_01_Source_Mapping_v0.md: direct inspection of F1 to F6; defects located to sheet and cell; source-to-workflow mapping with confidence; supersedes CW_06 sections 2 to 4.
- Added docs/discovery/DISC_02_Requirements_Coverage_v0.md: the seven Stage 1 areas against D-14 and STAGES.md; every gap assigned to an increment; deltas from D-14 listed.
- Added docs/discovery/DISC_03_Record_Model_v0.md: shared record model, statuses, financial and progress rules.
- Added docs/discovery/DISC_04_Build_Increments_v0.md: nine increments with testable gates; thin demonstration on one anonymised scheme kept separate from full Stage 1 depth; technology kept; blockers Q1 to Q4 and G1.
- DECISIONS.md: D-15 added (Stage 1 scope, LOCKED on instruction); D-03 and D-02 carry dated lines for the increment order and the moved D-02 block. No LOCKED line rewritten (W-11).
- STAGES.md: note that the file is re-cut against DISC_04 on acceptance; S00 unchanged.
- FINDINGS.md: F-003. RISKS.md: R-15, R-16. CLAUDE.md: Stage 1 scope bullet. data/legacy/README.md: dated line.
- Not done, by instruction: no repository created, no collaborator invited, no deployment, no new canvas, no contact with Tony or Mo.

## 08-Sep-2026 - Scope widened to the whole contractor commercial process (D-14)
- STAGES.md rewritten to 35 stages plus one optional; Phases 3 (Costs and CVR) and 4 (Cash flow and forecasts) inserted; later stages renumbered.
- DECISIONS.md: D-14 added; D-03, D-10, D-11 dated lines. RISKS.md: R-13, R-14; R-05, R-09 renumbered. CLAUDE.md scope bullet. data/legacy/README.md: F1 to F3 as reference layouts.

## 08-Sep-2026 - S00 pre-work (Reply 2)
- DECISIONS.md, STAGES.md, S00 stage file replaced with patched versions: build order reversed (D-03), D-05 two-tier, D-09 closed, W-03 hygiene scope, W-12.
- FINDINGS.md: F-001, F-002. RISKS.md: R-06 line. data/legacy/README.md: F1 to F3 rule.

## 08-Sep-2026 - S00 pre-work (Reply 1)
- DECISIONS.md patched (D-01 reopen line, D-13, W-11); CLAUDE.md two sections; S00 set to TODO; STAGES.md S00 row aligned; CCW set moved to docs/plan/history/; .gitignore fixed for data/legacy/README.md.
