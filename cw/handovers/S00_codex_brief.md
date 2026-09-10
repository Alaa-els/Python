# S00 - Codex review brief

- Stage: S00 repository bootstrap. Branch: claude/tony-project-files-p49w3l (holding repository; the cw/ folder is the project root). Commit: see handovers/S00_handover.md.
- Scope built: Django 5.2.17 project `cw/` with apps `core` and `projects` (no models), environment-based settings, pytest with pytest-django, four test modules, requirements.txt pinned, .env.example, .gitignore additions. No screen, export, model or product feature.

## What to check
- cw/settings.py: SECRET_KEY, DEBUG, ALLOWED_HOSTS, DATABASE_URL, TEST_DATABASE_URL, MEDIA_ROOT all come from the environment; the development fallback key is used only when DEBUG is on or under pytest, and the settings raise otherwise. Is `RUNNING_TESTS = "pytest" in sys.modules` an acceptable way to let TEST_DATABASE_URL override DATABASE_URL under test only, or should it be a pytest option?
- tests/test_hygiene.py: the scope in tests/hygiene_scope.py excludes docs/plan/history/ and data/legacy/; the tool-trace check is limited to product-facing folders (templates/, docs/guides/) per the W-03 line of 10-Sep-2026. Is that scoping right, or too narrow?
- tests/test_docs.py: the authoritative list replaces the withdrawn CW_ file gate (D-13 line of 10-Sep-2026); the reference resolver skips "Snn" and "<" placeholders. Check the list matches what you consider authoritative.
- tests/test_no_raw_sql.py: pattern list; it is a hygiene check only and the handover says so.
- Two-engine run: the handover records the PostgreSQL run (local 16.13 cluster, socket connection). Confirm the claim matches the recorded command and counts.

## Where the risk is
- The hygiene scope is a judgement; a too-wide scope would fail on the repository's own method documents, a too-narrow one would let a tool name into an export later. The S02 stage file should extend product_facing_files() to exports when they exist.
- The development SECRET_KEY fallback exists so `manage.py check` and pytest run with no .env; it must never reach a hosted build (S24 sets DEBUG=0 and a real key; the settings raise if not).
- Django support status was not verified from djangoproject.com (blocked by the network policy); 5.2 as the LTS series is a working assumption recorded in the handover.

## Not for you to conclude
- Whether S01 opens: that is Alaa's and yours after this brief and the audit.
