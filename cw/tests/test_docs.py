"""The authoritative documents exist and the standing brief's references resolve (D-13 line of 10-Sep-2026)."""
import re

from tests.hygiene_scope import REPO_ROOT

AUTHORITATIVE = [
    "CLAUDE.md",
    "SETUP.md",
    "docs/FOUNDATION.md",
    "docs/DECISIONS.md",
    "docs/STAGES.md",
    "docs/FINDINGS.md",
    "docs/RISKS.md",
    "docs/CHANGELOG.md",
    "docs/discovery/DISC_01_Source_Mapping_v0.md",
    "docs/discovery/DISC_02_Requirements_Coverage_v0.md",
    "docs/discovery/DISC_03_Record_Model_v0.md",
    "docs/discovery/DISC_04_Build_Increments_v0.md",
    "docs/discovery/DISC_05_Estimates_and_Usage_Ledger_v0.md",
    "docs/stages/_TEMPLATE.md",
    ".claude/commands/stage.md",
    ".claude/commands/audit.md",
    ".claude/commands/close.md",
    ".claude/agents/auditor.md",
]

PATH_PATTERN = re.compile(r"(?<![\w/])((?:docs|tests|data|handovers|\.claude)/[\w./-]+)")


def test_authoritative_documents_exist():
    missing = [rel for rel in AUTHORITATIVE if not (REPO_ROOT / rel).is_file()]
    assert not missing, "missing authoritative documents: " + ", ".join(missing)


def test_claude_md_and_setup_references_resolve():
    unresolved = []
    for name in ("CLAUDE.md", "SETUP.md"):
        text = (REPO_ROOT / name).read_text(encoding="utf-8")
        for match in PATH_PATTERN.finditer(text):
            rel = match.group(1).rstrip(".,;:)`")
            if rel.endswith("/"):
                rel = rel[:-1]
            if "Snn" in rel or "<" in rel:
                continue
            if not (REPO_ROOT / rel).exists():
                unresolved.append(f"{name}: {rel}")
    assert not unresolved, "references that do not resolve: " + "; ".join(unresolved)


def test_no_superseded_ccw_files_in_plan():
    plan = REPO_ROOT / "docs" / "plan"
    stray = sorted(p.name for p in plan.glob("CCW_*"))
    assert not stray, "CCW files directly in docs/plan/: " + ", ".join(stray)
