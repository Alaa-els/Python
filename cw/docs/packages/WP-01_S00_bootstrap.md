# WP-01 - S00 repository bootstrap: first bounded implementation package

- Date: 09-Sep-2026. Author: Alaa Elsayed. For: Codex review before the session opens. Stage file: docs/stages/S00_bootstrap.md (unchanged in scope; two additions below).
- Why this first: it has no external dependency, no data, no decision to close, and it makes every later gate runnable. It is also the first calibration point for DISC_05.

## Scope (from the stage file, plus two additions)
- Django 5 project `cw/` with apps `core` and `projects`; settings from env with SQLite default and `DATABASE_URL` override; `MEDIA_ROOT` outside the repo; timezone Europe/London; British English locale.
- Tooling: requirements.txt (django, django-simple-history, django-htmx, openpyxl, python-docx, pytest, pytest-django, dj-database-url, python-dotenv), pytest.ini, .env.example.
- Tests named in the stage file section 5: hygiene (four), settings (two), docs (two).
- Addition 1 (DISC_04 section 5): the same suite runs against PostgreSQL when `TEST_DATABASE_URL` is set, and against SQLite otherwise. The handover records which engines were actually run; a run that did not happen is reported as not run, never as passed. Adds the driver `psycopg` to requirements.txt; named here so W-09 is satisfied.
- Addition 2: a `tests/test_no_raw_sql.py` that greps the code for raw SQL and engine-specific functions and passes on an empty project; S02 keeps it passing. It is a hygiene check, not proof of portability.
- Out of scope: any model, screen, export, server, data.

## Deliverables
- Files: manage.py, cw/settings.py, cw/urls.py, core/, projects/, requirements.txt, pytest.ini, .env.example, tests/test_hygiene.py, tests/test_settings.py, tests/test_docs.py, tests/test_no_raw_sql.py, handovers/S00_audit.md, handovers/S00_codex_brief.md, handovers/S00_handover.md.
- Commands in the handover: fresh clone, venv, install, `python manage.py check`, `pytest`.

## Acceptance proof
- `python manage.py check` green; `pytest` green with every named test present. The former CW_ file gate is withdrawn (Codex review 10-Sep-2026): tests/test_docs.py checks that the authoritative documents exist and that CLAUDE.md and SETUP.md references resolve; the supersession is recorded under D-13; no placeholder file is created and no failing gate is skipped.
- Tag S00-done; docs/STAGES.md S00 = DONE.

## Estimate (DISC_05 basis)
- Agent active: 2 hours (1.5 to 3). Human review: 1.5 hours (1 to 2.5). Tokens, base: one session, about 40 tool turns, about 100,000 cached context, 60,000 fresh input, 25,000 output.

## Governance
- W-10: Alaa signed 08-Sep-2026; Codex review is the second review (W-10 line of 09-Sep-2026); Mo's line stays blank until Mo signs.
- External dependencies: none. Mo's signature line stays blank; Codex's review of 10-Sep-2026 is the second review under the W-10 line of 09-Sep-2026. Opened 10-Sep-2026 on Codex's go.
