# CW Findings Register

- Classification: BLOCKER (the stage cannot close until fixed), DEFECT (scheduled to a named stage, logged), OBSERVATION (recorded, no gate).
- Source: AUDIT (auditor subagent), CODEX (external review), PILOT (Gamma feedback), OWNER (Alaa or Mo).
- No finding is silently resolved. Closure names the stage and the commit.

| ID | Raised at | Source | Class | Finding | Closure |
|---|---|---|---|---|---|
| F-001 (A-001) | S00 pre-work, 08-Sep-2026 | OWNER | OBSERVATION | docs/STAGES.md S00 row read READY while docs/stages/S00_bootstrap.md read TODO; W-10 governs, Mo had not signed | CLOSED 08-Sep-2026, S00 pre-work commit: STAGES.md row set to TODO |
| F-002 | S00 pre-work, 08-Sep-2026 | AUDIT | OBSERVATION | CW_06 X10: F5 macros of unknown purpose. Inspected: Module1 builds the client copy (values-paste, delete red rows, columns and tabs, estimator checklist prompt); Module2 and Module3 toggle columns and trim rows on a sheet named Pricing Matrix that no longer exists. Nothing the workbench reads is touched | CLOSED 08-Sep-2026: no action; F5 pricing engine stays out of scope |
| F-003 | Discovery, 09-Sep-2026 | OWNER | OBSERVATION | Direct inspection of F1 to F6 (docs/discovery/DISC_01) confirms and locates CW_06 X1 (F1 0.4 Manufacturing v F3 3 Manuf PP: 60 cells differ, both fed by one external report), X3 (F2 keys schemes by free text, 504 strings, four exact matches) and X5 (F3 2A and 2b carry 813 and 414 #REF! and are dead), and adds two: F5 Manuf. (Gamyba) carries 2,924 #REF!, and no file holds programme, purchase order, inspection or actual-cost data | CLOSED 09-Sep-2026: recorded; DISC_01 supersedes CW_06 sections 2 to 4; no build action |
