"""Tests for PC_Health_Audit.pyw (analysis, scoring and comparison logic).

Run with:  python -m pytest tests -q
The audit collector itself needs Windows; these tests feed synthetic records.
"""
from __future__ import annotations

import importlib.machinery
import importlib.util
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def load_module():
    path = ROOT / "PC_Health_Audit.pyw"
    loader = importlib.machinery.SourceFileLoader("pchealth", str(path))
    spec = importlib.util.spec_from_loader("pchealth", loader)
    module = importlib.util.module_from_spec(spec)
    sys.modules["pchealth"] = module
    loader.exec_module(module)
    return module


m = load_module()


# ---------------------------------------------------------------------------
# Synthetic collector records
# ---------------------------------------------------------------------------

def ok(section, data):
    return {"section": section, "status": "ok", "data": data, "error": None}


def base_records(*, samples=None, deep=False):
    now = datetime.now(timezone.utc).astimezone()
    records = [
        ok("hello", {"isAdmin": False}),
        ok("context", {"isAdmin": False, "psVersion": "5.1", "is64BitProcess": True, "computerName": "[redacted]", "userName": "[redacted]"}),
        ok("system", {
            "manufacturer": "LENOVO", "model": "21CB", "pcSystemType": 2, "totalPhysicalMemory": 16 * 1024 ** 3, "logicalProcessors": 8,
            "windowsEdition": "Windows 11 Pro", "buildNumber": "26100", "displayVersion": "24H2", "ubr": 4351, "architecture": "64-bit",
            "lastBoot": (now - timedelta(days=2)).isoformat(), "deviceFingerprint": "abc123def456", "pendingReboot": False, "timeZone": "Arab Standard Time",
        }),
        {"section": "secure_boot", "status": "requires_elevation", "data": None, "error": "Access was denied"},
        ok("cpu", {"name": "i7-1260P", "sockets": 1, "physicalCores": 12, "logicalProcessors": 16, "maxClockMHz": 2100}),
        ok("memory", {"modules": [{"slot": "A", "capacityBytes": 8 * 1024 ** 3}], "slots": 2, "maxCapacityKB": 67108864, "maxCapacityExKB": 67108864, "pageFiles": []}),
        ok("gpu", {"controllers": [{"name": "NVIDIA RTX A2000", "adapterRAM": 4293918720, "pnpDeviceId": "PCI\\VEN_10DE&DEV_25B6&SUBSYS_1&REV_A1\\4&1"}],
                   "registry": [{"key": "0001", "driverDesc": "NVIDIA RTX A2000", "matchingDeviceId": "pci\\ven_10de&dev_25b6", "qwMemorySize": 8 * 1024 ** 3}]}),
        ok("physical_disks", {"items": [{"name": "NVMe", "mediaType": "SSD", "busType": "NVMe", "sizeBytes": 10 ** 12, "health": "Healthy", "operationalStatus": "OK"}], "source": "Get-PhysicalDisk", "reliabilityStatus": "requires_elevation", "reliabilityError": "Access denied"}),
        ok("logical_disks", {"items": [{"drive": "C:", "fileSystem": "NTFS", "sizeBytes": 1000, "freeBytes": 300}]}),
        ok("battery", {"present": True, "batteryStatus": 2, "chargePercent": 90, "designCapacity": 57000, "fullCapacity": 52000, "capacitySource": "Windows WMI (root/wmi)", "powercfgStatus": "collected", "powercfgReportDeleted": True, "wmiErrors": []}),
        {"section": "thermal", "status": "requires_elevation", "data": None, "error": "Access denied"},
        ok("network", {"items": []}), ok("wifi", {"present": False, "parsed": False}), ok("problem_devices", {"items": []}),
        {"section": "tpm", "status": "requires_elevation", "data": None, "error": "Access denied"},
        ok("antivirus", {"items": [{"product": "Windows Defender", "productState": 397568}]}),
        ok("firewall", {"items": [{"name": "Domain", "enabled": "True"}]}),
        {"section": "bitlocker", "status": "requires_elevation", "data": None, "error": "Access is denied"},
        ok("defender", {"signatureAgeDays": 1}), ok("startup", {"items": []}), ok("services", {"items": []}), ok("hotfixes", {"items": []}),
        ok("events", {"days": 14, "logs": [{"log": "System", "status": "collected", "totalEvents": 0, "cap": 2000, "capped": False, "groups": []}], "hardware": {"status": "collected", "totalEvents": 0, "cap": 1000, "capped": False, "groups": []}}),
        ok("reliability", {"total": 0, "cap": 40, "items": []}),
        {"section": "installed_apps", "status": "not_run", "data": None, "error": "Not requested."},
    ]
    samples = samples if samples is not None else default_samples(10)
    for index, sample in enumerate(samples):
        records.append({"section": "sample", "status": "ok", "data": sample, "index": index, "total": len(samples)})
    records.append(ok("samples_done", {"requested": len(samples), "elapsedSeconds": len(samples) * 2.5, "logicalProcessors": 8}))
    records.append(ok("top_ram", {"items": [{"process": "EXCEL", "instances": 1, "ramMB": 2100, "privateMB": 2000}]}))
    if deep:
        records.append(ok("dism_checkhealth", {"exitCode": 0, "output": "The component store is repairable.\nThe operation completed successfully."}))
        records.append(ok("sfc_verifyonly", {"exitCode": 0, "output": "Windows Resource Protection did not find any integrity violations."}))
    else:
        records.append({"section": "dism_checkhealth", "status": "not_run", "data": None, "error": "Deep mode not selected."})
        records.append({"section": "sfc_verifyonly", "status": "not_run", "data": None, "error": "Deep mode not selected."})
    records.append(ok("done", {"isAdmin": False}))
    return records


def make_sample(index, cpu=20.0, performance=100.0, available_mb=8000.0, pages_in=10.0, on_mains=True, cores=None, processes=None, disks=None, raw=None):
    now = datetime.now(timezone.utc).astimezone()
    return {
        "index": index, "time": (now + timedelta(seconds=index)).isoformat(), "elapsedMs": index * 1000, "durationMs": 900,
        "cpu": {"total": cpu, "cores": cores if cores is not None else [{"name": "0", "percent": cpu}, {"name": "1", "percent": cpu}]},
        "frequency": {"percentPerformance": performance, "percentMaxFrequency": 100, "frequencyMHz": 2100},
        "memory": {"availableMB": available_mb, "percentCommitted": 50, "pagesInputPerSec": pages_in, "pagesPerSec": pages_in},
        "disks": disks if disks is not None else [{"name": "_Total", "percentDiskTime": 5, "currentQueue": 0, "avgQueue": 0, "readBytesPerSec": 0, "writeBytesPerSec": 0, "transfersPerSec": 0}],
        "disksRaw": raw if raw is not None else [],
        "processes": processes if processes is not None else [{"name": "EXCEL", "pid": 1, "cpu": 80, "privateMB": 1000, "ioReadBytesPerSec": 0, "ioWriteBytesPerSec": 0}],
        "power": {"batteryPresent": True, "batteryStatus": 2 if on_mains else 1, "chargePercent": 80},
    }


def default_samples(count):
    return [make_sample(i) for i in range(count)]


def analyse(records, **kwargs):
    options = m.AuditOptions(mode=kwargs.pop("mode", "Standard"), output_dir=str(ROOT), **kwargs)
    return m.analyse_records(records, options, {"ExitCode": 0, "TempCleanup": []})


def finding_titles(report, severity=None):
    return [(f["Severity"], f["Category"], f["Finding"]) for f in report["Findings"] if severity is None or f["Severity"] == severity]


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def test_parse_timestamp_variants():
    aware, assumed = m.parse_timestamp("2026-09-03T10:25:55.1234567+03:00")
    assert aware.utcoffset() == timedelta(hours=3) and not assumed
    aware, assumed = m.parse_timestamp("/Date(1756880755000)/")
    assert aware.tzinfo is not None and not assumed
    naive, assumed = m.parse_timestamp("2026-09-03 10:25:55")
    assert naive.tzinfo is not None and assumed
    assert m.parse_timestamp("garbage") == (None, False)
    assert m.parse_timestamp(None) == (None, False)


def test_process_base_name_and_numbers():
    assert m.process_base_name("msedge#3") == "msedge"
    assert m.process_base_name("EXCEL.EXE") == "EXCEL"
    assert m.to_number(True) is None and m.to_number("nan") is None and m.to_number("12.5") == 12.5
    assert m.percentile([1, 2, 3, 4, 5], 95) == 5 and m.percentile([], 50) is None


# ---------------------------------------------------------------------------
# Audit analysis
# ---------------------------------------------------------------------------

def test_no_records_is_inconclusive_not_healthy():
    report = analyse([])
    assert report["Completeness"]["Status"] == "Incomplete"
    assert report["Rating"].startswith("Inconclusive")
    assert "Live performance samples" in report["Completeness"]["CoreMissing"]
    assert not any(f["Severity"] == "Good" for f in report["Findings"])


def test_standard_user_healthy_laptop_is_good_but_discloses_elevation_gaps():
    report = analyse(base_records())
    assert report["Completeness"]["Status"] == "Complete except elevation-only checks"
    assert report["Completeness"]["CoreMissing"] == []
    assert "Secure Boot" in report["Completeness"]["ElevationMissing"]
    assert report["Rating"] == "Healthy snapshot"
    assert "need elevation" in report["Headline"]
    good = [f for f in report["Findings"] if f["Severity"] == "Good"]
    assert good and "elevation only" in good[0]["Evidence"]


def test_max_ram_uses_kilobytes_and_gpu_uses_registry_qword():
    report = analyse(base_records())
    assert report["MemorySummary"]["MaximumSupported"] == "64.00 GB"
    gpu = report["GPUs"][0]
    assert gpu["DedicatedMemory"] == "8.00 GB" and "qwMemorySize" in gpu["MemorySource"]


def test_gpu_adapter_ram_saturation_is_labelled_when_registry_missing():
    records = base_records()
    for record in records:
        if record["section"] == "gpu":
            record["data"]["registry"] = []
    report = analyse(records)
    assert "saturated" in report["GPUs"][0]["MemorySource"]


def test_throttling_not_flagged_when_high_frequency_under_load():
    # ChatGPT's case: high frequency during every high-load sample, low frequency only when idle.
    samples = []
    for index in range(20):
        if index % 2 == 0:
            samples.append(make_sample(index, cpu=90, performance=130))
        else:
            samples.append(make_sample(index, cpu=5, performance=40))
    report = analyse(base_records(samples=samples))
    throttle = report["Performance"]["Throttling"]
    assert throttle["OverlapSamples"] == 0 and throttle["Suspected"] is False
    assert not any(f["Category"] == "Thermal" for f in report["Findings"])


def test_throttling_flagged_with_sustained_overlap_on_mains():
    samples = [make_sample(i, cpu=95, performance=45) for i in range(12)]
    report = analyse(base_records(samples=samples))
    throttle = report["Performance"]["Throttling"]
    assert throttle["Suspected"] is True and throttle["OverlapSamples"] == 12
    assert ("Warning", "Thermal") in [(f["Severity"], f["Category"]) for f in report["Findings"]]


def test_throttling_on_battery_is_information_only():
    samples = [make_sample(i, cpu=95, performance=45, on_mains=False) for i in range(12)]
    report = analyse(base_records(samples=samples))
    thermal = [f for f in report["Findings"] if f["Category"] == "Thermal"]
    assert thermal and thermal[0]["Severity"] == "Information"
    assert report["Performance"]["PowerSource"] == "Battery"


def test_throttling_ignores_zero_performance_counter():
    samples = [make_sample(i, cpu=95, performance=0) for i in range(12)]
    report = analyse(base_records(samples=samples))
    assert report["Performance"]["Throttling"]["PairedSamples"] == 0
    assert report["Performance"]["CPU_Performance_P05Percent"] is None
    assert any("could not be evaluated" in f["Finding"] for f in report["Findings"])


def test_paging_needs_low_ram_to_be_a_warning():
    high_pages_low_ram = [make_sample(i, available_mb=800, pages_in=500) for i in range(10)]
    report = analyse(base_records(samples=high_pages_low_ram))
    paging = [f for f in report["Findings"] if "read pages from disk" in f["Finding"]]
    assert paging and paging[0]["Severity"] == "Warning"
    high_pages_free_ram = [make_sample(i, available_mb=9000, pages_in=500) for i in range(10)]
    report = analyse(base_records(samples=high_pages_free_ram))
    info = [f for f in report["Findings"] if "without RAM pressure" in f["Finding"]]
    assert info and info[0]["Severity"] == "Information"


def test_processes_are_summed_per_application_within_a_sample():
    processes = [
        {"name": "msedge", "pid": 1, "cpu": 80, "privateMB": 300},
        {"name": "msedge#3", "pid": 2, "cpu": 160, "privateMB": 200},
        {"name": "EXCEL", "pid": 3, "cpu": 400, "privateMB": 2000},
    ]
    samples = [make_sample(i, processes=processes) for i in range(4)]
    report = analyse(base_records(samples=samples))
    rows = {row["Process"]: row for row in report["Performance"]["TopCPUProcesses"]}
    # logical processors = 8 from the system record: (80 + 160) / 8 = 30, 400 / 8 = 50
    assert rows["msedge"]["AverageCPU"] == 30.0 and rows["msedge"]["Instances"] == 2 and rows["msedge"]["PeakRAM_MB"] == 500
    assert rows["EXCEL"]["AverageCPU"] == 50.0


def test_disk_latency_from_raw_counters():
    def raw(t):
        # 800 transfers per interval at 40 ms each with a 10 MHz clock.
        return [{"name": "0 C:", "secPerTransfer": 320_000_000 * t, "secPerTransferBase": 800 * t, "secPerRead": 0, "secPerReadBase": 0, "secPerWrite": 0, "secPerWriteBase": 0, "frequencyPerfTime": 10_000_000, "timestampPerfTime": 10_000_000 * t}]
    disks = [{"name": "_Total", "percentDiskTime": 90, "currentQueue": 1, "avgQueue": 1, "readBytesPerSec": 0, "writeBytesPerSec": 0, "transfersPerSec": 800},
             {"name": "0 C:", "percentDiskTime": 90, "currentQueue": 1, "avgQueue": 1, "readBytesPerSec": 0, "writeBytesPerSec": 0, "transfersPerSec": 800}]
    samples = [make_sample(i, disks=disks, raw=raw(i + 1)) for i in range(6)]
    report = analyse(base_records(samples=samples))
    disk = next(d for d in report["Performance"]["PerDisk"] if d["Disk"] == "0 C:")
    assert disk["Latency_ms_P95"] == pytest.approx(40.0) and disk["LatencyIntervals"] == 5
    assert any("responded slowly" in f["Finding"] for f in report["Findings"])


def test_single_thread_bound_detection():
    cores = [{"name": "0", "percent": 98}] + [{"name": str(i), "percent": 5} for i in range(1, 8)]
    samples = [make_sample(i, cpu=18, cores=cores) for i in range(10)]
    report = analyse(base_records(samples=samples))
    assert report["Performance"]["SingleThreadBound"] is True
    assert any("One logical processor" in f["Finding"] for f in report["Findings"])


def test_critical_finding_bounds_rating_and_headline():
    records = base_records()
    for record in records:
        if record["section"] == "physical_disks":
            record["data"]["items"][0]["health"] = "Warning"
    report = analyse(records)
    assert report["Rating"] == "Attention required" and "critical" in report["Headline"]


def test_unknown_disk_health_is_not_unhealthy():
    records = base_records()
    for record in records:
        if record["section"] == "physical_disks":
            record["data"]["items"][0]["health"] = None
    report = analyse(records)
    assert report["PhysicalDisks"][0]["Health"] == "Unknown" and report["PhysicalDisks"][0]["HealthKnown"] is False
    assert not any(f["Severity"] == "Critical" for f in report["Findings"])


def test_event_diagnostics_use_warnings_and_all_groups():
    records = base_records()
    groups = [{"provider": "Microsoft-Windows-Kernel-Processor-Power", "eventId": 37, "level": 3, "levelName": "Warning", "count": 9},
              {"provider": "disk", "eventId": 153, "level": 3, "levelName": "Warning", "count": 4},
              {"provider": "Microsoft-Windows-WHEA-Logger", "eventId": 19, "level": 3, "levelName": "Warning", "count": 2},
              {"provider": "Microsoft-Windows-Kernel-Power", "eventId": 41, "level": 1, "levelName": "Critical", "count": 1}]
    groups += [{"provider": f"Other{i}", "eventId": i, "level": 2, "levelName": "Error", "count": 100 + i} for i in range(40)]
    for record in records:
        if record["section"] == "events":
            record["data"]["hardware"]["groups"] = groups
    report = analyse(records)
    diag = report["EventDiagnostics"]
    assert diag["ProcessorPower37"] == 9 and diag["Storage_Warnings"] == 4 and diag["WHEA_Warnings"] == 2 and diag["KernelPower41"] == 1
    categories = [(f["Severity"], f["Category"]) for f in report["Findings"]]
    assert ("Warning", "Storage events") in categories and ("Warning", "Hardware events") in categories and ("Warning", "Stability") in categories
    assert any("firmware-limited" in f["Finding"] for f in report["Findings"])


def test_deep_checks_classified_and_shown():
    report = analyse(base_records(deep=True), mode="Deep")
    assert report["DeepChecks"]["DISMCheckHealth"]["Verdict"].startswith("Corruption detected")
    assert report["DeepChecks"]["SFCVerifyOnly"]["Verdict"] == "No integrity violations"
    assert ("Critical", "Windows integrity") in [(f["Severity"], f["Category"]) for f in report["Findings"]]
    html = m.html_report(report)
    assert "component store is repairable" in html and "DISM /CheckHealth" in html


def test_sfc_inconclusive_output_is_information():
    records = base_records(deep=True)
    for record in records:
        if record["section"] == "sfc_verifyonly":
            record["data"]["output"] = "Ueberpruefung 100 % abgeschlossen."
    report = analyse(records, mode="Deep")
    assert report["DeepChecks"]["SFCVerifyOnly"]["Verdict"].startswith("Inconclusive")
    assert any(f["Category"] == "Windows integrity" and f["Severity"] == "Information" for f in report["Findings"])


def test_failed_sample_collection_marks_core_evidence_missing():
    records = [r for r in base_records(samples=[]) if r["section"] != "samples_done"]
    records.append({"section": "samples_done", "status": "ok", "data": {"requested": 10, "elapsedSeconds": 1}})
    report = analyse(records)
    assert report["Completeness"]["Status"] == "Incomplete"
    assert "Live performance samples" in report["Completeness"]["CoreMissing"]
    assert report["Rating"].startswith("Inconclusive")


def test_temp_cleanup_failure_is_reported():
    options = m.AuditOptions(output_dir=str(ROOT))
    report = m.analyse_records(base_records(), options, {"TempCleanup": [{"Path": "C:\\Temp\\x.xml", "Deleted": False, "Error": "in use"}]})
    assert any(row["Key"] == "temp_cleanup" and row["Status"] == "failed" for row in report["Coverage"])
    assert "Temporary file clean-up" in report["Completeness"]["CoreMissing"]


def test_report_files_round_trip(tmp_path):
    report = analyse(base_records())
    paths = m.write_reports(report, str(tmp_path / "out"), base_records())
    data = json.loads(Path(paths["json"]).read_text(encoding="utf-8"))
    assert data["SchemaVersion"] == "1.2" and data["Tool"] == "PC Health Audit"
    html = Path(paths["html"]).read_text(encoding="utf-8")
    assert "class='badge good'" in html and "Coverage" in html and "Throttling evaluation" in html
    summary = Path(paths["summary"]).read_text(encoding="utf-8")
    assert "Evidence:" in summary and "COVERAGE GAPS" in summary
    assert Path(paths["records"]).read_text(encoding="utf-8").count("\n") == len(base_records())


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def legacy_report(schema, when, device="A", score=90, memory_p95=70.0, processes=None, disks=None):
    return {
        "SchemaVersion": schema, "Tool": "PC Health Audit", "HealthScore": score, "Rating": "Healthy snapshot",
        "System": {"Manufacturer": "LENOVO", "Model": f"Model-{device}", "AuditTime": when},
        "CPU": {"Name": "i7"}, "MemorySummary": {"InstalledBytes": 16 * 1024 ** 3},
        "Performance": {"CPU_AveragePercent": 20.0, "CPU_P95Percent": 50.0, "Memory_AveragePercent": 60.0, "Memory_P95Percent": memory_p95, "DiskActive_P95": 30.0,
                        "TopCPUProcesses": processes or [], "TopRAMProcesses": processes or []},
        "PhysicalDisks": disks if disks is not None else [{"Name": "SSD", "Health": "Healthy"}],
        "LogicalDisks": [{"Drive": "C:", "FreePercent": 40.0}], "Battery": {"HealthPercent": 80.0}, "Findings": [],
    }


def write_reports(tmp_path, reports):
    paths = []
    for index, report in enumerate(reports):
        folder = tmp_path / f"PC_Health_Audit_{index}"
        folder.mkdir()
        path = folder / "PC_Health_Data.json"
        path.write_text(json.dumps(report), encoding="utf-8")
        paths.append(path)
    return paths


def test_compare_orders_by_timezone_aware_time(tmp_path):
    early = legacy_report("1.1", "2026-09-01T23:00:00+03:00", score=95)   # 20:00 UTC
    late = legacy_report("1.1", "2026-09-01T22:00:00+00:00", score=60)    # 22:00 UTC, later despite the smaller clock time
    write_reports(tmp_path, [late, early])
    result = m.compare_files([str(tmp_path)], str(tmp_path / "cmp.md"))
    text = (tmp_path / "cmp.md").read_text(encoding="utf-8")
    assert result["Reports"] == 2 and "Latest result:** 60/100" in text


def test_compare_separates_devices(tmp_path):
    a = legacy_report("1.1", "2026-09-01T10:00:00+03:00", device="A")
    b = legacy_report("1.1", "2026-09-02T10:00:00+03:00", device="B")
    write_reports(tmp_path, [a, b])
    result = m.compare_files([str(tmp_path)], str(tmp_path / "cmp.md"))
    text = (tmp_path / "cmp.md").read_text(encoding="utf-8")
    assert result["Devices"] == 2 and "different devices" in text and "Model-A" in text and "Model-B" in text


def test_compare_sums_legacy_process_entries():
    processes = [{"Process": "msedge", "AverageCPU": 10, "RAM_MB": 1000}, {"Process": "msedge#3", "AverageCPU": 20, "RAM_MB": 2000}]
    report = m.LoadedReport(Path("x.json"), legacy_report("1.0", "2026-09-01T10:00:00+03:00", processes=processes))
    cpu, ram, malformed = m.report_processes(report)
    assert cpu["msedge"] == 30 and ram["msedge"] == 3000 and malformed == 0
    rows = m.recurring_processes([report])
    assert rows[0][0] == "msedge" and rows[0][2] == "30.0%" and rows[0][3] == "3000 MB"


def test_compare_memory_metric_mismatch_is_not_compared(tmp_path):
    old = legacy_report("1.0", "2026-08-01T10:00:00+03:00", memory_p95=50.0)
    new = legacy_report("1.1", "2026-09-01T10:00:00+03:00", memory_p95=90.0)
    write_reports(tmp_path, [old, new])
    m.compare_files([str(tmp_path)], str(tmp_path / "cmp.md"))
    text = (tmp_path / "cmp.md").read_text(encoding="utf-8")
    assert "Memory not compared" in text and "(commit)" in text and "memory pressure increased" not in text


def test_compare_unknown_and_malformed_disks(tmp_path):
    report = legacy_report("1.1", "2026-09-01T10:00:00+03:00", disks=[{"Name": "A", "Health": None}, "garbage", {"Name": "B", "Health": "Warning"}])
    write_reports(tmp_path, [report])
    m.compare_files([str(tmp_path)], str(tmp_path / "cmp.md"))
    text = (tmp_path / "cmp.md").read_text(encoding="utf-8")
    assert "1 unhealthy / 1 unknown" in text and "malformed disk entries ignored" in text


def test_compare_rejects_bad_files(tmp_path):
    (tmp_path / "PC_Health_Data.json").write_text('{"Tool": "Other"}', encoding="utf-8")
    (tmp_path / "broken.json").write_text("{", encoding="utf-8")
    with pytest.raises(RuntimeError):
        m.compare_files([str(tmp_path), str(tmp_path / "broken.json")], str(tmp_path / "cmp.md"))


def test_compare_accepts_new_schema_output(tmp_path):
    report = analyse(base_records())
    paths = m.write_reports(report, str(tmp_path / "PC_Health_Audit_x"))
    result = m.compare_files([paths["json"]], str(tmp_path / "cmp.md"))
    text = (tmp_path / "cmp.md").read_text(encoding="utf-8")
    assert result["Reports"] == 1 and "Complete except elevation-only checks" in text
