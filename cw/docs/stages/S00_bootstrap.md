# S00 - Repository bootstrap

- Status: READY (W-10 as amended 09-Sep-2026: Codex review is the second review; Mo's line stays blank until he signs)
- Reviewed by (W-10): Alaa 08-Sep-2026, Codex 10-Sep-2026, Mo <not signed>
- Branch: stage/S00. Tag on close: S00-done.

## 1. Objective
- Create the empty, tested, documented Django repository so that every later stage starts from a green gate.

## 2. Inputs in the repository
- CLAUDE.md, docs/FOUNDATION.md, docs/DECISIONS.md, docs/STAGES.md, this file.
- docs/plan/CW_02_Blueprint_v0.md section 7 (stack) only.

## 3. Decisions this stage depends on
- D-04 (DEFAULT) stack. D-13 (LOCKED) document set. W-01 to W-12 (LOCKED).
- D-06 and D-07 are OPEN and block S01, not S00: S00 creates no product feature and holds no data.

## 4. Changes
- Project: `cw/` Django project with apps `core` (companies, users, roles, membership, settings, history helpers) and `projects` (empty for now). No models yet beyond Django's user.
- Settings: `DATABASE_URL` from env, default `sqlite:///dev.sqlite3`; `MEDIA_ROOT` outside the repo; timezone Europe/London; British English locale; secret key from env.
- Tooling: `requirements.txt` (django, django-simple-history, django-htmx, openpyxl, python-docx, pytest, pytest-django, dj-database-url, python-dotenv), `pytest.ini`, `.env.example`, `.gitignore` (data/legacy, *.sqlite3, .env, media, __pycache__).
- Tests: `tests/test_hygiene.py` scanning repo-authored .py, .md, .html and .txt files for non-ASCII characters other than currency symbols, en or em dashes, single-quoted Python string literals outside comments, and the strings "Claude", "Anthropic", "python-docx", "openpyxl" in product-facing files (templates/, docs/guides/; W-03 line of 10-Sep-2026). `docs/plan/history/` is superseded reference content and is excluded from the scan (W-03). `tests/test_settings.py` proving the database URL default and the media root location.
- Docs check (amended 10-Sep-2026, Codex): `tests/test_docs.py` proving the authoritative set exists (docs/discovery/DISC_01 to DISC_05, docs/STAGES.md, docs/DECISIONS.md, docs/FOUNDATION.md, docs/FINDINGS.md, docs/RISKS.md, docs/CHANGELOG.md, .claude/commands/stage.md, audit.md, close.md, .claude/agents/auditor.md), that every repository path cited in CLAUDE.md and SETUP.md resolves, and that no file directly in docs/plan/ is named CCW_* (D-13). The former CW_ file-existence gate is withdrawn: those files were never placed and are superseded (D-13 line of 10-Sep-2026).
- Database portability (amended 10-Sep-2026): `tests/test_no_raw_sql.py` greps for raw SQL and engine-specific functions; that is a hygiene check, not proof of portability. Two-engine proof is a separate, honestly recorded run: pytest against PostgreSQL when `TEST_DATABASE_URL` is set, reported in the handover as run or not run.
- Dependency versions (amended 10-Sep-2026): Django 5.2 LTS series pinned to the patch version installed at S00 (recorded in requirements.txt and the handover); support status verified from the official Django site if reachable, otherwise recorded as unverified from documentation and taken from the package index only.
- Stage tooling: verify `.claude/commands/` and `.claude/agents/auditor.md` load (`/help`).
- Data folders: `data/legacy/README.md`, `tests/fixtures/README.md`.
- Dependencies added (W-09): the requirements list above plus psycopg (driver for the PostgreSQL test run), nothing else.

## 5. Tests to write
- tests/test_hygiene.py::test_no_non_ascii_except_currency_in_authored_files
- tests/test_hygiene.py::test_history_folder_excluded_from_scan
- tests/test_hygiene.py::test_no_en_or_em_dashes
- tests/test_hygiene.py::test_python_strings_double_quoted
- tests/test_hygiene.py::test_no_tool_traces_in_docs_and_templates
- tests/test_settings.py::test_database_defaults_to_sqlite
- tests/test_settings.py::test_media_root_outside_repo
- tests/test_docs.py::test_authoritative_documents_exist
- tests/test_docs.py::test_claude_md_and_setup_references_resolve
- tests/test_docs.py::test_no_superseded_ccw_files_in_plan
- tests/test_no_raw_sql.py::test_no_raw_sql_or_engine_specific_functions

## 6. Other-face note (W-07)
- None; no face exists yet.

## 7. Risks touched
- R-10 (branch discipline set up here), R-11 (session size rule adopted).

## 8. Out of scope
- Any model, screen or export. Server setup (S05). Anything touching Gamma data (D-05).

## 9. Definition of done
- `python manage.py check` green; `pytest` green; first commit on main; handovers/S00_handover.md written; tag S00-done; docs/STAGES.md S00 = DONE.
