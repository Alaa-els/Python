# CLAUDE.md - Commercial Workbench (CW)

You are building the Commercial Workbench with Alaa Elsayed and Mo Ashour. Read this file, then docs/FOUNDATION.md, then the stage file named in the session. Nothing else is assumed.

## What this is
- One Django codebase, two faces on the same records: the Contractor edition (product for small contractors, they pay) and the Consultant edition (Alaa's and Mo's own work). A project setting `my_role` decides the face.
- Scope (D-14): a small contractor's whole commercial process on one database - handover budgets, bill and rates, variations raised on the phone, applications and certification, costs and CVR, cash flow and forecasts, records, correspondence and claims, reports. Bookkeeping, the estimating engine and drawing reading stay out.
- Definition: docs/plan/CW_02_Blueprint_v0.md. Order and gates: docs/STAGES.md. Locked decisions: docs/DECISIONS.md (D-13 records that the CW set supersedes the earlier CCW set; a CCW id in an old document resolves through it). Never rederive these; if one does not fit the stage, stop and say so.

## How work happens (the stage protocol, docs/FOUNDATION.md section 2)
- One stage per session, on branch `stage/Snn`, from `main` with a clean tree. Never two stages in one session. Stop at the stage boundary.
- Start of session: `/stage Snn`. It restates the objective, the locked decisions touched, the acceptance tests, then plans. No file is written before the plan is approved.
- Pre-write gate: `git status` clean, `pytest` green on main, `python manage.py check` green. Any failure halts the stage.
- Build in small commits. Write the tests named in the stage file before or with the code. Present each deliverable when it passes its own check, not at the end.
- End of session: `/audit Snn` (the auditor subagent, read-only) then `/close Snn` (handover, tag, status). Findings go to docs/FINDINGS.md; nothing is fixed silently.
- Steer by editing the stage file, never by mid-build correction. Where a stage premise diverges from the delivered state, the delivered state governs and the divergence is recorded.

## Verifiable rules
- Python 3.12, Django 5, PostgreSQL in use, SQLite for development and tests (`DATABASE_URL` env, default SQLite). HTMX templates, no JavaScript framework. openpyxl and python-docx for exports.
- Every model: `company`, `created_by`, `created_at`, `updated_by`, `updated_at`, history via django-simple-history. Imported rows: `source_file`, `source_tab`, `source_row`, `imported_at`. No hard deletes: `status` closed or superseded.
- Every view checks role plus project membership. A company never sees another company's rows. Tests prove both.
- Contract terms, thresholds and alarm settings are rows in settings tables, never constants in code. Export layouts are per-company templates, never hard-coded.
- Money fields carry a `status` (draft, submitted, claimed, assessed, certified, agreed, instructed, paid, forecast). Contractor figures are claims until agreed; certified and instructed figures are labelled as such.
- Every screen: one question, Next and Back, where you are, what is next, source shown beside each figure. Every export: readable first sheet (purpose, result, basis, action), no named ranges, neutral voice, metadata author `Alaa Elsayed`.
- Delivered code and documents: British English, hyphens (never en or em dashes), pure ASCII, double-quoted Python strings, no AI or library traces in any file. `pytest tests/test_hygiene.py` enforces this.
- A feature built on one face states in the handover what the other face still needs.
- Test data: `data/legacy/` (gitignored) holds Tony's files; fixtures under `tests/fixtures/` are anonymised extracts only. Never employer client data; if any appears, say so and stop.
- Absorb Alaa's existing tools (listed in docs/plan/CW_01_Scope_and_Filter_v0.md section 4); never rewrite them.

## Commands
- Run: `python manage.py runserver` - Tests: `pytest` - Checks: `python manage.py check && pytest tests/test_hygiene.py`
- Stage commands: `/stage Snn`, `/audit Snn`, `/close Snn` (see .claude/commands/).

## Do not
- Start a stage without `/stage`. Skip the pre-write gate. Fix an audit finding without recording it. Commit to `main` directly. Touch `data/legacy/` in git. Add a dependency not named in the stage file.
- Rewrite a LOCKED decision in place - reopen it with a dated line under it (W-11).
- Build a feature only Gamma would ever use. A feature every small contractor needs is the product; a feature specific to Gamma's own way of working is quoted separately (D-01, CW_04 section 4).
- Use Gamma's real data before written permission exists (D-05). Until then: the blank F4 template and anonymised fixtures only.
