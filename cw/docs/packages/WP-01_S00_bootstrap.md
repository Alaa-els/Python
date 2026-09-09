# WP-01 - S00 repository bootstrap: first bounded implementation package

- Date: 09-Sep-2026. Author: Alaa Elsayed. For: Codex review before the session opens. Stage file: docs/stages/S00_bootstrap.md (unchanged in scope; two additions below).
- Why this first: it has no external dependency, no data, no decision to close, and it makes every later gate runnable. It is also the first calibration point for DISC_05.

## Scope (from the stage file, plus two additions)
- Django 5 project `cw/` with apps `core` and `projects`; settings from env with SQLite default and `DATABASE_URL` override; `MEDIA_ROOT` outside the repo; timezone Europe/London; British English locale.
- Tooling: requirements.txt (django, django-simple-history, django-htmx, openpyxl, python-docx, pytest, pytest-django, dj-database-url, python-dotenv), pytest.ini, .env.example.
- Tests named in the stage file section 5: hygiene (four), settings (two), docs (two).
- Addition 1 (DISC_04 section 5): a second pytest configuration that runs the same suite against PostgreSQL when `TEST_DATABASE_URL` is set, skipped otherwise, so S02 can enforce the two-engine rule without changing S00's gate. Adds the driver `psycopg` to requirements.txt; named here so W-09 is satisfied.
- Addition 2: a `tests/test_no_raw_sql.py` that greps the code for raw SQL and SQLite-specific functions and passes on an empty project; S02 keeps it passing.
- Out of scope: any model, screen, export, server, data.

## Deliverables
- Files: manage.py, cw/settings.py, cw/urls.py, core/, projects/, requirements.txt, pytest.ini, .env.example, tests/test_hygiene.py, tests/test_settings.py, tests/test_docs.py, tests/test_no_raw_sql.py, handovers/S00_audit.md, handovers/S00_codex_brief.md, handovers/S00_handover.md.
- Commands in the handover: fresh clone, venv, install, `python manage.py check`, `pytest`.

## Acceptance proof
- `python manage.py check` green; `pytest` green with every named test present; `tests/test_docs.py::test_named_plan_files_exist` passes only once the seven CW_ plan files are placed (Q: still not supplied; the test is written and its failure is the recorded gap, not a skip).
- Tag S00-done; docs/STAGES.md S00 = DONE.

## Estimate (DISC_05 basis)
- Agent active: 2 hours (1.5 to 3). Human review: 1.5 hours (1 to 2.5). Tokens, base: one session, about 40 tool turns, about 100,000 cached context, 60,000 fresh input, 25,000 output.

## Governance
- W-10: Alaa signed 08-Sep-2026; Codex review is the second review (W-10 line of 09-Sep-2026); Mo's line stays blank until Mo signs.
- Blocked by nothing external. Opens on Codex's go.
