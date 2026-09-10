"""Hygiene is a test (W-03): ASCII with currency symbols, hyphens only, double-quoted Python strings,
no tool traces in product-facing files."""
import io
import tokenize

from tests.hygiene_scope import REPO_ROOT, authored_files, product_facing_files

ALLOWED_NON_ASCII = {"\u00a3", "\u20ac"}
DASHES = {"\u2013", "\u2014"}
TOOL_STRINGS = ("Claude", "Anthropic", "python-docx", "openpyxl")


def _rel(path):
    return path.relative_to(REPO_ROOT).as_posix()


def test_no_non_ascii_except_currency_in_authored_files():
    offenders = []
    for path in authored_files():
        text = path.read_text(encoding="utf-8")
        for lineno, line in enumerate(text.splitlines(), 1):
            bad = [c for c in line if ord(c) > 127 and c not in ALLOWED_NON_ASCII]
            if bad:
                offenders.append(f"{_rel(path)}:{lineno} {bad[:3]}")
    assert not offenders, "non-ASCII characters outside the currency allowance: " + "; ".join(offenders[:20])


def test_history_folder_excluded_from_scan():
    scanned = {_rel(p) for p in authored_files()}
    assert not any(rel.startswith("docs/plan/history") for rel in scanned)
    assert (REPO_ROOT / "docs" / "plan" / "history").is_dir()


def test_no_en_or_em_dashes():
    offenders = []
    for path in authored_files():
        text = path.read_text(encoding="utf-8")
        for lineno, line in enumerate(text.splitlines(), 1):
            if any(d in line for d in DASHES):
                offenders.append(f"{_rel(path)}:{lineno}")
    assert not offenders, "en or em dashes found: " + "; ".join(offenders[:20])


def test_python_strings_double_quoted():
    offenders = []
    for path in authored_files():
        if path.suffix != ".py":
            continue
        source = path.read_text(encoding="utf-8")
        for tok in tokenize.generate_tokens(io.StringIO(source).readline):
            if tok.type != tokenize.STRING:
                continue
            body = tok.string.lstrip("rRbBfFuU")
            if body.startswith("'"):
                offenders.append(f"{_rel(path)}:{tok.start[0]}")
    assert not offenders, "single-quoted Python strings: " + "; ".join(offenders[:20])


def test_no_tool_traces_in_docs_and_templates():
    offenders = []
    for path in product_facing_files():
        text = path.read_text(encoding="utf-8")
        for needle in TOOL_STRINGS:
            if needle in text:
                offenders.append(f"{_rel(path)}: {needle}")
    assert not offenders, "tool or library names in product-facing files: " + "; ".join(offenders)
