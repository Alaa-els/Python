# CW Findings Register

- Classification: BLOCKER (the stage cannot close until fixed), DEFECT (scheduled to a named stage, logged), OBSERVATION (recorded, no gate).
- Source: AUDIT (auditor subagent), CODEX (external review), PILOT (Gamma feedback), OWNER (Alaa or Mo).
- No finding is silently resolved. Closure names the stage and the commit.

| ID | Raised at | Source | Class | Finding | Closure |
|---|---|---|---|---|---|
| F-001 (A-001) | S00 pre-work, 08-Sep-2026 | OWNER | OBSERVATION | docs/STAGES.md S00 row read READY while docs/stages/S00_bootstrap.md read TODO; W-10 governs, Mo had not signed | CLOSED 08-Sep-2026, S00 pre-work commit: STAGES.md row set to TODO |
| F-002 | S00 pre-work, 08-Sep-2026 | AUDIT | OBSERVATION | CW_06 X10: F5 macros of unknown purpose. Inspected: Module1 builds the client copy (values-paste, delete red rows, columns and tabs, estimator checklist prompt); Module2 and Module3 toggle columns and trim rows on a sheet named Pricing Matrix that no longer exists. Nothing the workbench reads is touched | CLOSED 08-Sep-2026: no action; F5 pricing engine stays out of scope |
| F-003 | Discovery, 09-Sep-2026 | OWNER | OBSERVATION | Direct inspection of F1 to F6 (docs/discovery/DISC_01) confirms and locates CW_06 X1 (F1 0.4 Manufacturing v F3 3 Manuf PP: 60 cells differ, both fed by one external report), X3 (F2 keys schemes by free text, 504 strings, four exact matches) and X5 (F3 2A and 2b carry 813 and 414 #REF! and are dead), and adds two: F5 Manuf. (Gamyba) carries 2,924 #REF!, and no file holds programme, purchase order, inspection or actual-cost data | CLOSED 09-Sep-2026: recorded; DISC_01 supersedes CW_06 sections 2 to 4; no build action |
| F-004 (A-01) | S00 audit, 10-Sep-2026 | AUDIT | BLOCKER | handovers/S00_handover.md missing at audit time | CLOSED 10-Sep-2026 at close: handover written |
| F-005 (A-02) | S00 audit, 10-Sep-2026 | AUDIT | DEFECT | CLAUDE.md required Python 3.12; the suite ran on 3.11.15 | CLOSED 10-Sep-2026: D-04 dated line accepts 3.11 or later for development with 3.12 the target; CLAUDE.md rule amended; S24 fixes the hosted runtime |
| F-006 (A-03) | S00 audit, 10-Sep-2026 | AUDIT | OBSERVATION | No test opened a database, so the PostgreSQL pass proved settings resolution only | CLOSED 10-Sep-2026: tests/test_database.py adds a create-and-read round trip run on both engines; the handover states exactly what each engine run proved |
| F-007 (A-04) | S00 audit, 10-Sep-2026 | AUDIT | OBSERVATION | Stage file section 9 says "first commit on main" while CLAUDE.md forbids direct commits to main; work is on the holding branch claude/tony-project-files-p49w3l; tag S00-done | OPEN, recorded under W-06: delivered state governs; commits are on the holding branch, the tag is created locally and pushed only when the dedicated cw repository exists |
| F-008 (A-05) | S00 audit, 10-Sep-2026 | AUDIT | OBSERVATION | data/legacy/README.md is tracked by an explicit .gitignore negation against the "never touch data/legacy in git" wording | OPEN: deliberate (Reply 2, 08-Sep-2026); the README holds instructions only; CLAUDE.md wording to be tightened at S02 |
| F-009 (A-06) | S00 audit, 10-Sep-2026 | AUDIT | OBSERVATION | DEBUG defaults on and SECRET_KEY falls back to a development string when DEBUG or tests are on | OPEN, scheduled S24: settings already raise when DEBUG is off and no key is set |

