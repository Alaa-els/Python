---
description: Audit a stage with the read-only auditor subagent and record findings
---
Audit stage $ARGUMENTS.

1. Delegate to the auditor subagent (.claude/agents/auditor.md) with this brief: stage file docs/stages/$ARGUMENTS_*.md, CLAUDE.md rules, docs/FOUNDATION.md section 3. It must run `pytest` on a scratch database, check every test named in section 5 exists and passes, check every acceptance criterion, check hygiene, check that no settings value or layout is hard-coded, check the other-face note exists, and confirm the working tree is byte-unchanged after its run (`git status`).
2. The auditor writes handovers/$ARGUMENTS_audit.md with findings classified BLOCKER, DEFECT or OBSERVATION, each with file and line where relevant. It edits no other file.
3. Back in the main session: append every finding to docs/FINDINGS.md with source AUDIT and class. Fix BLOCKERs and DEFECTs in the build on the stage branch, one commit each, and record the closure against the finding (stage and commit). Leave OBSERVATIONs open unless trivial.
4. Re-run `pytest`. If a DEFECT needed a test change, record it as an audit-side finding (FOUNDATION section 3).
5. Report: findings raised, closed, still open. Do not close the stage here; /close does that.
