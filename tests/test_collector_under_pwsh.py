"""Run the embedded PowerShell collector under PowerShell 7 with mocked cmdlets.

Skipped when pwsh is not on PATH (or POWERSHELL_EXE is unset). This checks the
collector's control flow, JSON streaming and per-check status reporting. It
cannot prove Windows PowerShell 5.1 behaviour; the Windows smoke test in the
README remains mandatory.
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_pc_health_audit import m  # noqa: E402

PWSH = os.environ.get("POWERSHELL_EXE") or shutil.which("pwsh")
PRELUDE = Path(__file__).resolve().parent / "mock_cmdlets.ps1"

pytestmark = pytest.mark.skipif(not PWSH, reason="pwsh not available")


def launcher(script_path, script_args):
    quoted = " ".join(f"'{a}'" if not a.startswith("-") else a for a in script_args)
    return [PWSH, "-NoProfile", "-NonInteractive", "-Command", f". '{PRELUDE}'; & '{script_path}' {quoted}"]


def test_collector_streams_every_section_and_reports_statuses():
    out = tempfile.mkdtemp(prefix="pchealth_pwsh_")
    options = m.AuditOptions(mode="Standard", sample_count=5, output_dir=out, workload_label="Large Excel workbook")
    messages = []
    result = m.run_audit(options, lambda message, fraction=None: messages.append(message), launcher=launcher)
    report = result["Report"]
    assert result["RunMeta"]["ExitCode"] == 0 and not result["RunMeta"]["Cancelled"]
    assert all(entry["Deleted"] for entry in result["RunMeta"]["TempCleanup"])
    statuses = {row["Key"]: row["Status"] for row in report["Coverage"]}
    assert statuses["system"] == "collected" and statuses["secure_boot"] == "requires_elevation"
    assert statuses["samples_done"] == "collected" and report["Performance"]["SamplesTaken"] == 5
    assert report["MemorySummary"]["MaximumSupported"] == "64.00 GB"
    assert report["GPUs"][0]["DedicatedMemory"] == "8.00 GB"
    assert report["Performance"]["Throttling"]["Suspected"] is False
    assert report["EventDiagnostics"]["ProcessorPower37"] == 1 and report["EventDiagnostics"]["Storage_Warnings"] == 1
    processes = {row["Process"]: row for row in report["Performance"]["TopCPUProcesses"]}
    assert processes["msedge"]["Instances"] == 2 and processes["msedge"]["AverageCPU"] == 30.0
    latency = next(d for d in report["Performance"]["PerDisk"] if d["Disk"] == "0 C:")
    assert latency["Latency_ms_P95"] == pytest.approx(40.0)
    assert report["Completeness"]["CoreMissing"] == []
    assert any("Performance sample 5/5" in message for message in messages)
    for path in result["Paths"].values():
        assert os.path.exists(path)
