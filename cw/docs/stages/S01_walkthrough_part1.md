# S01 - Spine walk-through, part 1

- Status: AUDIT (delivered 10-Sep-2026; awaiting Codex review)
- Reviewed by (W-10): Alaa - building authorised under Codex management, instruction of 10-Sep-2026, no separate signature; Codex - review pending; Mo - not signed
- Branch: claude/tony-project-files-p49w3l (holding repository). Tag on close: S01-done.

## 1. Objective
- Write the plain-language walk-through of DISC_03 sections 1, 2 and 9 with one sample row per table and explicit invariants, so that S02 can write the models without rederiving anything.

## 2. Inputs in the repository
- docs/discovery/DISC_03_Record_Model_v0.md sections 1, 2, 9; docs/DECISIONS.md (D-15, W-13); handovers/S00_handover.md.

## 3. Decisions this stage depends on
- D-15 (LOCKED), W-13 (LOCKED). Defaults under Q3 (GEMMA, GEM-###) used as editable settings.

## 4. Changes
- docs/walkthroughs/W1_spine_part1.md: 21 tables, seven invariants I-1 to I-7, defaults kept, three questions with defaults. No code.

## 5. Tests to write
- None new. tests/test_hygiene.py and tests/test_docs.py cover the new document.

## 6. Other-face note (W-07)
- The consultant face needs a way for a consultant party's user to hold a scheme_membership on a contractor's scheme; kept open in W1 section 5 for Increment 9.

## 7. Risks touched
- R-08 (nothing hard-coded: formats and heads are settings), R-15 (WBS source).

## 8. Out of scope
- DISC_03 sections 3 to 8; any model or screen; real Gamma data.

## 9. Definition of done
- Walk-through on the branch; hygiene green; Codex review recorded; STAGES.md S01 = DONE after that review.
