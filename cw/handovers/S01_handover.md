# S01 handover - spine walk-through, part 1

- Delivered: docs/walkthroughs/W1_spine_part1.md (organisation, scheme register, cross-cutting; sample rows; invariants I-1 to I-7; defaults; three open questions). docs/stages/S01_walkthrough_part1.md. No code, no model, no data.
- Settled per Codex's 10-Sep-2026 direction: cost allocation (value by cost code, shares = 100 percent) and WBS allocation (quantity by activity and location, sum <= item quantity) are separate tables (I-1); aggregates never exceed the authorised quantity and a breach is refused with the excess named, never clamped (I-2); membership and foreign-key ownership agree at the model layer (I-3); item_ref is the stable identity across revisions with one version row per revision (I-4); money is fixed-point decimal with an explicit currency resolved through the scheme (I-5); history and files cannot cross schemes, proven by a two-company test at S02 (I-6); revisions are single-headed (I-7).
- Decisions touched: none reopened. Defaults GEMMA and GEM-### recorded as editable settings under Q3.
- Findings: none raised. Checks run: hygiene and docs tests on the new document (see the close commit message for counts).
- Other-face note: consultant-party membership on a contractor scheme is open for Increment 9.
- Review record: Alaa authorised under Codex management (10-Sep-2026); Codex review pending; Mo not signed.
- Next stage: S02 (walk-through part 2, then models on both engines). Brief: docs/packages/WP-03_S02_models.md. Opens on Codex's review of this handover.
