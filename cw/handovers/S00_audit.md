# S00 audit - repository bootstrap

Audited 10-Sep-2026 against docs/stages/S00_bootstrap.md sections 4, 5, 8 and 9 and the
"Verifiable rules" and "Do not" sections of CLAUDE.md.

## Runs

- SQLite: `DEBUG=1 python -m pytest -q` - 12 passed, 0 failed.
- PostgreSQL: `TEST_DATABASE_URL=postgres://pgtest@/cwtest?host=/tmp/pgcw&port=5433` - 12 passed, 0 failed.
- `python manage.py check` - System check identified no issues (0 silenced).
- Working tree: `git status --short` identical before and after the runs. No new tracked or
  untracked entries; `__pycache__` folders were refreshed but are gitignored, and no
  `.pytest_cache` appeared (runs used `-p no:cacheprovider`).

## Findings

| id | class | file and line | finding | evidence |
| --- | --- | --- | --- | --- |
| A-01 | BLOCKER | docs/stages/S00_bootstrap.md:53 | Definition of done requires handovers/S00_handover.md; it does not exist. | `ls handovers/` returns README.md only. |
| A-02 | DEFECT | CLAUDE.md:20 | Rule states Python 3.12; the suite and checks ran on Python 3.11.15. | `python --version` = Python 3.11.15; compiled files are cpython-311. |
| A-03 | OBSERVATION | tests/ (all files) | The PostgreSQL run proves nothing about portability: no test opens a database. | No `django_db` marker or database fixture anywhere in tests/; both engines report the same 12 passes in 0.08s. Settings do resolve to `django.db.backends.postgresql`, name cwtest, under pytest. |
| A-04 | OBSERVATION | docs/stages/S00_bootstrap.md:53 | Remaining definition-of-done items are unmet at audit time and are close-time actions: no commit carrying the S00 files, no tag S00-done, docs/STAGES.md S00 still OPEN. Section 9 also asks for a "first commit on main", which sits against the CLAUDE.md rule not to commit to main directly. | Branch is claude/tony-project-files-p49w3l; `git tag` empty; all S00 files untracked; docs/STAGES.md:12 reads OPEN 10-Sep-2026. |
| A-05 | OBSERVATION | .gitignore:2 | data/legacy/README.md is tracked by an explicit negation. Section 4 asks for that file, and it holds instructions only, no client data, but the CLAUDE.md line is "never touch data/legacy in git". | `git ls-files cw/data` returns cw/data/legacy/README.md. |
| A-06 | OBSERVATION | cw/settings.py:18 | DEBUG defaults to on and the secret key falls back to a shared development string when DEBUG or tests are on. Safe for S00, but server work (S05) must close this. | `DEBUG = os.environ.get("DEBUG", "1") == "1"`. |

No further BLOCKER items. Only one DEFECT (A-02); no other hygiene or test-coverage defect was found.

## Checks that passed

- All eleven tests named in section 5 exist and run: five in tests/test_hygiene.py, two of the
  named settings tests plus an extra PostgreSQL URL test in tests/test_settings.py, three in
  tests/test_docs.py, one in tests/test_no_raw_sql.py.
- Hygiene beyond the tests: no single-quoted strings in manage.py, cw/, core/, projects/ or
  tests/ (only an apostrophe inside a docstring and a literal quote character in the checker);
  no en or em dashes and no non-ASCII bytes in any of those files or in requirements.txt,
  pytest.ini and .env.example.
- requirements.txt pins every package with `==`, lists exactly the dependencies named in
  section 4 plus psycopg, and every pin matches the installed version.
- .env.example carries no secret: SECRET_KEY is empty and the database lines are commented
  examples.
- Settings read DATABASE_URL, SECRET_KEY and MEDIA_ROOT from the environment
  (cw/settings.py lines 20, 82 and 107); MEDIA_ROOT defaults outside the repository.
- Section 6 (other-face note, W-07) and section 8 (out of scope) are both present in the stage
  file, and scope was held: no model, screen or export was added
  (no `models.Model` in core/ or projects/; both apps hold only an AppConfig).

## Verdict

Not ready to close. The stage work itself is sound - both runs are green, the named tests all
exist, hygiene is clean and scope was held - but the definition of done is not met while
handovers/S00_handover.md is missing (A-01), and the Python version in use does not match the
stated rule (A-02); the two-engine claim should be recorded honestly as a settings resolution
check rather than a portability proof (A-03), and the close-time items in A-04 remain open.
