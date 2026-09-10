"""Which files the hygiene tests look at.

Authored files: every .py, .md, .html and .txt file in the repository except
generated or third-party folders and docs/plan/history/, which is superseded
reference content (W-03).
"""
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
AUTHORED_SUFFIXES = {".py", ".md", ".html", ".txt"}
EXCLUDED_PARTS = {".git", "__pycache__", ".pytest_cache", "node_modules", "media", ".venv", "venv"}
EXCLUDED_PREFIXES = ("docs/plan/history", "data/legacy")
PRODUCT_FACING_PREFIXES = ("templates", "docs/guides")


def authored_files():
    for path in sorted(REPO_ROOT.rglob("*")):
        if not path.is_file() or path.suffix not in AUTHORED_SUFFIXES:
            continue
        rel = path.relative_to(REPO_ROOT).as_posix()
        if any(part in EXCLUDED_PARTS for part in path.parts):
            continue
        if rel.startswith(EXCLUDED_PREFIXES):
            continue
        yield path


def product_facing_files():
    for path in authored_files():
        rel = path.relative_to(REPO_ROOT).as_posix()
        if rel.startswith(PRODUCT_FACING_PREFIXES):
            yield path
