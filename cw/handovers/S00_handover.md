# S00 handover - repository bootstrap

- Closed 10-Sep-2026 under W-13 (synthetic-data contractor development, Codex review as second review; Mo's signature line blank). Branch: claude/tony-project-files-p49w3l in the holding repository; cw/ is the project root. Close commit: see git log ("S00: close").

## Delivered
- manage.py; cw/settings.py (environment-based: DEBUG, SECRET_KEY, ALLOWED_HOSTS, DATABASE_URL, TEST_DATABASE_URL under pytest, MEDIA_ROOT outside the repository; Europe/London; en-gb), cw/urls.py, cw/wsgi.py, cw/asgi.py; core/ and projects/ app shells with no models.
- requirements.txt pinned: django 5.2.17, django-simple-history 3.13.0, django-htmx 1.29.0, openpyxl 3.1.5, python-docx 1.2.0, pytest 9.1.1, pytest-django 4.14.0, dj-database-url 3.1.2, python-dotenv 1.2.3, psycopg[binary] 3.3.5. pytest.ini, .env.example, .gitignore additions (cw_media/, .venv/, *.egg-info/).
- tests/: hygiene (5), settings (3), docs (3), no-raw-sql (1), database round trip (2): 14 tests.
- Not delivered from the stage file: nothing; the CW_ file gate was withdrawn before the build (D-13 line of 10-Sep-2026).

## Check results, exactly
- `python manage.py check`: System check identified no issues (0 silenced).
- SQLite: `DEBUG=1 python -m pytest -p no:cacheprovider -q` -> 14 passed. `python manage.py migrate` applied Django's built-in migrations to dev.sqlite3 (gitignored).
- PostgreSQL 16.13 (local cluster started for this session as a non-root user, socket connection, database cwtest): `TEST_DATABASE_URL="postgres://pgtest@/cwtest?host=/tmp/pgcw&port=5433" DEBUG=1 python -m pytest -p no:cacheprovider -q` -> 14 passed, including the create-and-read round trip in tests/test_database.py; `DATABASE_URL=... python manage.py migrate` applied the built-in migrations (10 tables in schema public); `manage.py check` clean.
- What the two-engine run proves: settings resolve to each engine, Django's built-in migrations apply on both, and one row round-trips on both. It does not prove portability of product models, which do not exist yet; S02 extends the check to every model.
- Python in use: 3.11.15 (D-04 line of 10-Sep-2026; 3.12 remains the target).
- Django support status: closed 10-Sep-2026. Codex verified on djangoproject.com/download/ that 5.2 is the LTS series, latest 5.2.17, extended support to April 2028. This session could not reach that site (network policy); the verification is Codex's document check, not a rerun here.

## Decisions touched
- D-13 dated line (CW_ set never placed; discovery set authoritative). W-03 dated line (tool-trace check scoped to product-facing files). D-04 dated line (Python 3.11 accepted for development). W-10 line of 09-Sep-2026 applied; W-13 applies.

## Findings raised at this stage
- F-004 BLOCKER closed (this handover). F-005 DEFECT closed (D-04 line, CLAUDE.md). F-006 OBSERVATION closed (tests/test_database.py). F-007, F-008, F-009 OBSERVATION open, recorded with owners.

## Other-face note (W-07)
- None; no face exists yet.

## Run block (fresh clone)

POSIX shell (Linux, macOS, WSL, Git Bash):

```
git clone <repository> && cd cw
python -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python manage.py check
python -m pytest -q
# optional PostgreSQL pass (needs a reachable server and an empty database)
TEST_DATABASE_URL=postgres://user:password@localhost:5432/cwtest python -m pytest -q
```

PowerShell (Windows):

```
git clone <repository>; cd cw
python -m venv .venv; .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
python manage.py check
python -m pytest -q
# optional PostgreSQL pass
$env:TEST_DATABASE_URL = "postgres://user:password@localhost:5432/cwtest"; python -m pytest -q
```

- Keep DEBUG=1 and the development key only on a developer machine. Any hosted build (S24) sets DEBUG=0 and a real SECRET_KEY; the settings refuse to start otherwise.
- The default MEDIA_ROOT is a cw_media folder beside the cw project folder; while cw sits inside the holding repository that folder is still inside that repository's root and is gitignored there. Set MEDIA_ROOT explicitly on any shared machine.

## Codex review brief
- handovers/S00_codex_brief.md.

## Next stage
- S01 spine walk-through part 1 (docs only). Stage file: not yet written; brief at docs/packages/WP-02_S01_walkthrough.md. Must be true before it opens: S00 DONE (this file), Codex go on WP-02, hygiene green.

## Usage, observed only
- Read-only audit subagent: 54,967 tokens, 22 tool uses, 183 seconds. Main session figures are not shown on this surface and are not estimated here.
