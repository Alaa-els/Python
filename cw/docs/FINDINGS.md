# CW Findings Register

- Classification: BLOCKER (the stage cannot close until fixed), DEFECT (scheduled to a named stage, logged), OBSERVATION (recorded, no gate).
- Source: AUDIT (auditor subagent), CODEX (external review), PILOT (Gamma feedback), OWNER (Alaa or Mo).
- No finding is silently resolved. Closure names the stage and the commit.

| ID | Raised at | Source | Class | Finding | Closure |
|---|---|---|---|---|---|
| F-001 (A-001) | S00 pre-work, 08-Sep-2026 | OWNER | OBSERVATION | docs/STAGES.md S00 row read READY while docs/stages/S00_bootstrap.md read TODO; W-10 governs, Mo had not signed | CLOSED 08-Sep-2026, S00 pre-work commit: STAGES.md row set to TODO |
| F-002 | S00 pre-work, 08-Sep-2026 | AUDIT | OBSERVATION | CW_06 X10: F5 macros of unknown purpose. Inspected: Module1 builds the client copy (values-paste, delete red rows, columns and tabs, estimator checklist prompt); Module2 and Module3 toggle columns and trim rows on a sheet named Pricing Matrix that no longer exists. Nothing the workbench reads is touched | CLOSED 08-Sep-2026: no action; F5 pricing engine stays out of scope |
