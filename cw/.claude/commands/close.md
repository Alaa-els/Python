---
description: Close a stage - tests green, merge, tag, handover, status
---
Close stage $ARGUMENTS.

1. Confirm handovers/$ARGUMENTS_audit.md exists and no BLOCKER or DEFECT in docs/FINDINGS.md raised at this stage is open. If any is open, stop.
2. Run `pytest` and `python manage.py check`. Both must be green.
3. Write handovers/$ARGUMENTS_handover.md, under two pages, bullets only:
   - Delivered (files, screens, exports, tests) and what was not delivered from the stage file.
   - Decisions touched, with any dated lines added to docs/DECISIONS.md.
   - Findings raised at this stage and their status.
   - Other-face note (W-07): what the other face still needs.
   - CMD run block: exact commands to run and test what was delivered on a fresh clone.
   - Codex review brief pointer: handovers/$ARGUMENTS_codex_brief.md (write it: what was built, what to check, where the risk is; never what to conclude).
   - Next stage: its id, whether its stage file exists, what must be true before it opens.
4. Merge stage/$ARGUMENTS into main with a merge commit, tag $ARGUMENTS-done, delete the stage branch.
5. Set the stage file status to DONE and update the row in docs/STAGES.md. Commit "$ARGUMENTS: close".
6. Print the handover path and stop. The next stage starts in a new session.
