---
description: Open a build stage - read the stage file, restate, plan, then gate before any write
---
Open stage $ARGUMENTS of the Commercial Workbench.

Do these in order and stop at step 6 until the plan is approved:

1. Read CLAUDE.md, docs/FOUNDATION.md (section 2), docs/DECISIONS.md and docs/stages/$ARGUMENTS_*.md (the stage file). Read the latest file in handovers/ named *_handover.md.
2. If the stage file status is not READY, or a decision it depends on is OPEN, stop and say which one. Do not proceed.
3. Restate in bullets: the objective in one sentence; the decisions this stage depends on with their statuses; the tests in section 5 as the acceptance criteria; what is out of scope; what the other face will need.
4. Run the pre-write gate: `git status` must be clean on main; `pytest` must be green; `python manage.py check` must be green (skip the last two only at S00 before they exist). Report each result. A red gate halts the stage.
5. Create and switch to branch stage/$ARGUMENTS. Set the stage file status to OPEN.
6. Enter plan mode and present the implementation plan: files to create or change, tests to write first, commits in order. Wait for approval. Write nothing before approval.

While building: small commits prefixed "$ARGUMENTS:"; write each test in section 5 before or with its code; present each deliverable when its own test passes; never add a dependency the stage file does not name; never touch docs/DECISIONS.md except to add a dated line when told to; never fix an audit finding without recording it in docs/FINDINGS.md.
