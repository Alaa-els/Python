"""A hygiene check, not proof of portability: application code carries no raw SQL or engine-specific calls."""
import re

from tests.hygiene_scope import REPO_ROOT

APP_FOLDERS = ("cw", "core", "projects")
PATTERNS = [
    re.compile(r"\.raw\("),
    re.compile(r"RawSQL"),
    re.compile(r"connection\.cursor"),
    re.compile(r"cursor\.execute"),
    re.compile(r"\bsqlite_[a-z]+\("),
    re.compile(r"\bjson_extract\("),
    re.compile(r"\bstrftime\("),
    re.compile(r"\bILIKE\b"),
]


def test_no_raw_sql_or_engine_specific_functions():
    offenders = []
    for folder in APP_FOLDERS:
        for path in sorted((REPO_ROOT / folder).rglob("*.py")):
            text = path.read_text(encoding="utf-8")
            for pattern in PATTERNS:
                if pattern.search(text):
                    offenders.append(f"{path.relative_to(REPO_ROOT).as_posix()}: {pattern.pattern}")
    assert not offenders, "raw SQL or engine-specific calls: " + "; ".join(offenders)
