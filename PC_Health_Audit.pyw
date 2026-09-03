#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""PC Health Audit - single-file Windows laptop diagnostic with a window.

Read-only, offline and privacy-conscious. It makes no configuration changes and
uses no network connection. On a company-managed device, obtain IT approval
before running it.

Double-click the .pyw file to open the window (pythonw.exe, no console). The
window runs the audit, writes the reports and can compare earlier audits.

What it writes
  <output folder>\PC_Health_Audit_<timestamp>\PC_Health_Report.html
  <output folder>\PC_Health_Audit_<timestamp>\PC_Health_Data.json   (schema 1.2)
  <output folder>\PC_Health_Audit_<timestamp>\PC_Health_Summary.txt
  optional ZIP next to the folder, optional Battery_Report.html (sensitive)
  A temporary collector script and battery XML are created under %TEMP% and
  deleted afterwards; any deletion failure is reported in the coverage table.

How it collects
  Python starts one hidden Windows PowerShell 5.1 process that streams JSON
  lines (one record per check and one per performance sample). Every check is
  recorded as collected, unavailable, requires elevation, failed or not
  applicable, so a missing check can never masquerade as a clean result. All
  analysis, scoring and report assembly happen in Python.

Command line (optional; the window is the default)
  PC_Health_Audit.pyw audit [--mode Quick|Standard|Deep] [--samples N]
                            [--output DIR] [--workload LABEL] [--sensitive]
                            [--apps] [--zip] [--open]
  PC_Health_Audit.pyw compare INPUT [INPUT ...] [-o PC_Health_Comparison.md]
  PC_Health_Audit.pyw replay RECORDS.jsonl [--output DIR]   (offline analysis)

Requires Python 3.9 or newer with tkinter (included in the python.org
installer). No third-party packages.
"""

from __future__ import annotations

import argparse
import ctypes
import glob
import hashlib
import html
import json
import math
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

TOOL_NAME = "PC Health Audit"
TOOL_VERSION = "1.2.0"
SCHEMA_VERSION = "1.2"
SUPPORTED_SCHEMAS = ("1.0", "1.1", "1.2")
IS_WINDOWS = sys.platform.startswith("win")

MODES = ("Quick", "Standard", "Deep")
DEFAULT_SAMPLES = {"Quick": 10, "Standard": 30, "Deep": 60}
EVENT_DAYS = {"Quick": 3, "Standard": 14, "Deep": 30}
SAMPLE_MIN, SAMPLE_MAX = 5, 180

WORKLOAD_LABELS = (
    "Idle after restart",
    "Normal work",
    "Large Excel workbook",
    "Civil 3D / AutoCAD",
    "Active slowdown (run while it feels slow)",
    "Other (describe in notes)",
)

# Triage thresholds. These are heuristics for evidence gathering, not standards.
THRESHOLDS = {
    "disk_free_critical": 10.0,
    "disk_free_warning": 20.0,
    "battery_critical": 60.0,
    "battery_warning": 75.0,
    "cpu_avg_warning": 80.0,
    "core_p95_single_thread": 90.0,
    "ram_p95_critical": 90.0,
    "ram_p95_warning": 80.0,
    "paging_avg_pages_in": 200.0,
    "throttle_load": 70.0,
    "throttle_performance": 70.0,
    "throttle_min_samples": 3,
    "throttle_min_fraction": 0.10,
    "disk_active_p95": 95.0,
    "disk_queue_p95": 2.0,
    "disk_latency_ms_p95": 30.0,
    "disk_latency_min_active": 30.0,
    "uptime_days": 14.0,
    "storage_temp": 70.0,
    "ssd_wear": 80.0,
    "thermal_critical": 95.0,
    "thermal_warning": 85.0,
    "reliability_records": 10,
}

SEVERITY_ORDER = {"Critical": 0, "Warning": 1, "Information": 2, "Good": 3}
COVERAGE_STATUSES = ("collected", "partial", "unavailable", "requires_elevation", "failed", "not_applicable", "not_run")

METRIC_DEFINITIONS = {
    "Memory_AveragePercent": "Physical RAM in use, 100 minus available (schema 1.1+). In schema 1.0 this field was commit charge and is not comparable.",
    "Commit_AveragePercent": "Commit charge as a percentage of the commit limit (includes the pagefile).",
    "PagesInputPerSec": "Pages read from disk per second (Memory\\Pages Input/sec). Sustained values with low available RAM indicate paging pressure; not a count of hard faults.",
    "CPU_AveragePercent": "Processor Time, all logical processors combined.",
    "CPU_Performance_AveragePercent": "Processor Performance, actual frequency as a percentage of nominal (can exceed 100 with turbo).",
    "DiskActive_P95": "Physical disk _Total % Disk Time, 95th percentile.",
    "DiskLatency_ms": "Average seconds per transfer computed from raw counters between consecutive samples, in milliseconds.",
    "AverageCPU (process)": "Average over every sample, all instances of the application combined, divided by logical processor count.",
}


# ---------------------------------------------------------------------------
# Small utilities
# ---------------------------------------------------------------------------

def to_number(value: Any) -> Optional[float]:
    """Return a finite float or None. Booleans are not numbers here."""
    if value is None or isinstance(value, bool):
        return None
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def nested(data: Any, *keys: str, default: Any = None) -> Any:
    value = data
    for key in keys:
        if not isinstance(value, dict) or key not in value:
            return default
        value = value[key]
    return value


def as_list(value: Any) -> List[Any]:
    """PowerShell serialises a one-item collection as a scalar; normalise."""
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def as_dict_list(value: Any) -> List[Dict[str, Any]]:
    return [item for item in as_list(value) if isinstance(item, dict)]


def fmt_bytes(value: Any) -> str:
    numeric = to_number(value)
    if numeric is None:
        return "N/A"
    for unit, size in (("TB", 1024 ** 4), ("GB", 1024 ** 3), ("MB", 1024 ** 2), ("KB", 1024)):
        if numeric >= size:
            return f"{numeric / size:.2f} {unit}" if unit != "KB" else f"{numeric / size:.0f} KB"
    return f"{numeric:.0f} bytes"


def fmt_value(value: Any, suffix: str = "", decimals: int = 1) -> str:
    numeric = to_number(value)
    if numeric is None:
        return "N/A"
    return f"{numeric:.{decimals}f}{suffix}"


def mean(values: Sequence[float]) -> Optional[float]:
    values = [v for v in values if v is not None]
    return sum(values) / len(values) if values else None


def percentile(values: Sequence[float], pct: float) -> Optional[float]:
    """Nearest-rank percentile; returns None for an empty sequence."""
    values = sorted(v for v in values if v is not None)
    if not values:
        return None
    index = int(math.ceil((pct / 100.0) * len(values))) - 1
    index = min(max(index, 0), len(values) - 1)
    return float(values[index])


def rounded(value: Optional[float], decimals: int = 1) -> Optional[float]:
    if value is None:
        return None
    return round(value, decimals) if decimals else float(round(value))


def clean_text(value: Any, limit: int = 0) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\x00", "").strip()
    if limit and len(text) > limit:
        text = text[: limit - 3] + "..."
    return text


def single_line(value: Any, limit: int = 0) -> str:
    """clean_text with all whitespace runs collapsed to one space (for table cells)."""
    return clean_text(re.sub(r"\s+", " ", "" if value is None else str(value)), limit)


def now_local() -> datetime:
    return datetime.now(timezone.utc).astimezone()


_ISO_FRACTION = re.compile(r"(\.\d{1,6})\d*")


def parse_timestamp(value: Any) -> Tuple[Optional[datetime], bool]:
    """Parse ISO 8601 or PowerShell /Date(ms)/ values.

    Returns (aware datetime or None, assumed_local). Naive values are assumed to
    be local time of the machine running the analysis and flagged.
    """
    if value is None:
        return None, False
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.astimezone(), True
        return value, False
    text = str(value).strip()
    if not text:
        return None, False
    match = re.match(r"^/Date\((-?\d+)([+-]\d{4})?\)/$", text)
    if match:
        try:
            return datetime.fromtimestamp(int(match.group(1)) / 1000.0, tz=timezone.utc), False
        except (ValueError, OverflowError, OSError):
            return None, False
    candidate = text
    if candidate.endswith("Z"):
        candidate = candidate[:-1] + "+00:00"
    candidate = _ISO_FRACTION.sub(r"\1", candidate)
    candidate = re.sub(r"([+-]\d{2})(\d{2})$", r"\1:\2", candidate)
    try:
        parsed = datetime.fromisoformat(candidate)
    except ValueError:
        for pattern in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %I:%M:%S %p", "%Y-%m-%dT%H:%M:%S"):
            try:
                parsed = datetime.strptime(text, pattern)
                break
            except ValueError:
                continue
        else:
            return None, False
    if parsed.tzinfo is None:
        return parsed.astimezone(), True
    return parsed, False


def iso(value: Optional[datetime]) -> Optional[str]:
    return value.isoformat(timespec="seconds") if value else None


def escape_markdown(value: Any) -> str:
    return str(value if value is not None else "N/A").replace("|", "\\|").replace("\n", " ")


def markdown_table(headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> str:
    output = [
        "| " + " | ".join(escape_markdown(value) for value in headers) + " |",
        "|" + "|".join("---" for _ in headers) + "|",
    ]
    output.extend("| " + " | ".join(escape_markdown(value) for value in row) + " |" for row in rows)
    return "\n".join(output)


def fingerprint(*parts: Any) -> str:
    digest = hashlib.sha256("|".join(clean_text(p) for p in parts).encode("utf-8")).hexdigest()
    return digest[:12]


def process_base_name(name: Any) -> str:
    """Counter instances are named msedge#3; Get-Process names them msedge."""
    text = clean_text(name) or "Unknown"
    text = re.sub(r"#\d+$", "", text)
    if text.lower().endswith(".exe"):
        text = text[:-4]
    return text or "Unknown"


def default_output_dir() -> str:
    if IS_WINDOWS:
        try:
            buffer = ctypes.create_unicode_buffer(1024)
            # CSIDL_DESKTOPDIRECTORY = 0x0010
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buffer) == 0 and buffer.value:
                return buffer.value
        except Exception:
            pass
    desktop = Path.home() / "Desktop"
    return str(desktop if desktop.is_dir() else Path.home())


def crash_log_path() -> str:
    return os.path.join(tempfile.gettempdir(), "PC_Health_Audit_crash.log")


def show_fatal_error(message: str) -> None:
    """Last-resort error display that works without a console or tkinter."""
    try:
        with open(crash_log_path(), "a", encoding="utf-8") as handle:
            handle.write(f"\n---- {now_local().isoformat()} ----\n{message}\n")
    except OSError:
        pass
    if IS_WINDOWS:
        try:
            ctypes.windll.user32.MessageBoxW(None, message[:3000], f"{TOOL_NAME} - error", 0x10)
            return
        except Exception:
            pass
    try:
        sys.stderr.write(message + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Embedded Windows PowerShell 5.1 collector. It only collects: one JSON line
# per check ("section") and one per performance sample. It never returns
# collections from functions (the construct that failed in v1.1), makes no
# configuration changes and touches the network for nothing.
# ---------------------------------------------------------------------------

POWERSHELL_COLLECTOR = r"""
#requires -version 5.1
[CmdletBinding()]
param(
    [string]$Mode = 'Standard',
    [int]$SampleCount = 30,
    [int]$EventDays = 14,
    [string]$BatteryReportPath = '',
    [switch]$IncludeSensitiveData,
    [switch]$IncludeInstalledApps,
    [switch]$SkipSamples
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

$script:Stdout = [Console]::OpenStandardOutput()
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:IsAdmin = $false
try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $script:IsAdmin = [bool]$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { }

function Emit {
    param($Record)
    $json = ConvertTo-Json -InputObject $Record -Compress -Depth 12
    $bytes = $script:Utf8.GetBytes([string]$json + "`n")
    $script:Stdout.Write($bytes, 0, $bytes.Length)
    $script:Stdout.Flush()
}

function Send-Progress {
    param([string]$Message)
    Emit @{ section = 'progress'; status = 'ok'; message = $Message; time = (Get-Date).ToString('o') }
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $property = $null
    try { $property = $Object.PSObject.Properties[$Name] } catch { return $Default }
    if ($null -eq $property) { return $Default }
    $value = $property.Value
    if ($null -eq $value) { return $Default }
    return $value
}

function Str {
    param($Value)
    if ($null -eq $Value) { return $null }
    return [string]$Value
}

function Num {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { if ($Value) { return 1.0 } else { return 0.0 } }
    try { return [double]$Value } catch { return $null }
}

function NumOr {
    param($Value, [double]$Default = 0)
    $n = Num $Value
    if ($null -eq $n) { return $Default }
    return $n
}

function Iso {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { $Value = [DateTime]::SpecifyKind($Value, [DateTimeKind]::Local) }
        return $Value.ToString('o')
    }
    $text = [string]$Value
    if ($text -match '^\d{14}\.\d{6}[+-]\d{3}$') {
        try { return ([Management.ManagementDateTimeConverter]::ToDateTime($text)).ToString('o') } catch { }
    }
    try { return ([datetime]$text).ToString('o') } catch { return $text }
}

function Redact {
    param($Value)
    if ($IncludeSensitiveData) { return (Str $Value) }
    if ($null -eq $Value) { return $null }
    return '[redacted]'
}

function Get-FailureStatus {
    param([string]$Message, [bool]$NeedsElevation)
    if ($Message -match 'not supported on this platform') { return 'not_applicable' }
    if ($NeedsElevation -and -not $script:IsAdmin) { return 'requires_elevation' }
    if ($Message -match 'Access is denied|Access denied|AccessDenied|UnauthorizedAccess|Insufficient privilege|elevat|administrator|0x80070005') { return 'requires_elevation' }
    if ($Message -match 'Invalid class|Invalid namespace|not supported|Not supported|is not recognized|could not be loaded|could not be found|does not exist|No such interface|not installed|Not found|not found') { return 'unavailable' }
    return 'failed'
}

function Invoke-Section {
    param([string]$Name, [scriptblock]$Body, [bool]$NeedsElevation = $false, [string]$Note = '')
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $record = @{ section = $Name; status = 'ok'; data = $null; error = $null; note = $Note; ms = 0 }
    try {
        $record.data = & $Body
    }
    catch {
        $message = ''
        try { $message = [string]$_.Exception.Message } catch { $message = [string]$_ }
        if (-not $message) { $message = 'Unknown error' }
        $record.status = Get-FailureStatus $message $NeedsElevation
        $record.error = $message
    }
    $record.ms = [int]$watch.ElapsedMilliseconds
    Emit $record
}

function Get-NetshValue {
    param([string[]]$Lines, [string]$Pattern)
    $line = $Lines | Where-Object { $_ -match $Pattern } | Select-Object -First 1
    if ($line) { return (($line -split ':', 2)[1]).Trim() }
    return $null
}

try {

Emit @{ section = 'hello'; status = 'ok'; data = @{ tool = 'PC Health Audit collector'; version = '1.2.0'; isAdmin = $script:IsAdmin } }

Send-Progress 'Reading system information'
Invoke-Section 'context' {
    $psv = $PSVersionTable.PSVersion
    @{
        isAdmin = $script:IsAdmin
        psVersion = ('{0}.{1}' -f $psv.Major, $psv.Minor)
        is64BitProcess = [Environment]::Is64BitProcess
        is64BitOS = [Environment]::Is64BitOperatingSystem
        computerName = (Redact $env:COMPUTERNAME)
        userName = (Redact ([Environment]::UserName))
        startTime = (Get-Date).ToString('o')
        mode = $Mode
        sampleCount = $SampleCount
        eventDays = $EventDays
        includeSensitiveData = [bool]$IncludeSensitiveData
        includeInstalledApps = [bool]$IncludeInstalledApps
    }
}

Invoke-Section 'system' {
    $cs = @(Get-CimInstance Win32_ComputerSystem -ErrorAction Stop) | Select-Object -First 1
    $os = @(Get-CimInstance Win32_OperatingSystem -ErrorAction Stop) | Select-Object -First 1
    $bios = $null
    try { $bios = @(Get-CimInstance Win32_BIOS -ErrorAction Stop) | Select-Object -First 1 } catch { }
    $product = $null
    try { $product = @(Get-CimInstance Win32_ComputerSystemProduct -ErrorAction Stop) | Select-Object -First 1 } catch { }
    $reg = $null
    try { $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop } catch { }
    $uuid = Str (Get-Prop $product 'UUID')
    $fingerprint = $null
    if ($uuid) {
        try {
            $sha = [Security.Cryptography.SHA256]::Create()
            $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes('pc-health|' + $uuid))
            $fingerprint = (([BitConverter]::ToString($hash)) -replace '-', '').Substring(0, 12).ToLowerInvariant()
        } catch { }
    }
    $pendingReboot = $false
    foreach ($path in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')) {
        if (Test-Path -LiteralPath $path) { $pendingReboot = $true }
    }
    try {
        $sm = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
        if ($null -ne (Get-Prop $sm 'PendingFileRenameOperations')) { $pendingReboot = $true }
    } catch { }
    $tz = [TimeZoneInfo]::Local
    @{
        manufacturer = Str (Get-Prop $cs 'Manufacturer')
        model = Str (Get-Prop $cs 'Model')
        systemFamily = Str (Get-Prop $cs 'SystemFamily')
        systemType = Str (Get-Prop $cs 'SystemType')
        pcSystemType = Num (Get-Prop $cs 'PCSystemType')
        totalPhysicalMemory = Num (Get-Prop $cs 'TotalPhysicalMemory')
        logicalProcessors = Num (Get-Prop $cs 'NumberOfLogicalProcessors')
        windowsEdition = Str (Get-Prop $os 'Caption')
        windowsVersion = Str (Get-Prop $os 'Version')
        buildNumber = Str (Get-Prop $os 'BuildNumber')
        displayVersion = Str (Get-Prop $reg 'DisplayVersion')
        ubr = Num (Get-Prop $reg 'UBR')
        architecture = Str (Get-Prop $os 'OSArchitecture')
        lastBoot = Iso (Get-Prop $os 'LastBootUpTime')
        installDate = Iso (Get-Prop $os 'InstallDate')
        totalVisibleMemoryKB = Num (Get-Prop $os 'TotalVisibleMemorySize')
        freePhysicalMemoryKB = Num (Get-Prop $os 'FreePhysicalMemory')
        totalVirtualMemoryKB = Num (Get-Prop $os 'TotalVirtualMemorySize')
        freeVirtualMemoryKB = Num (Get-Prop $os 'FreeVirtualMemory')
        biosVersion = Str (Get-Prop $bios 'SMBIOSBIOSVersion')
        biosDate = Iso (Get-Prop $bios 'ReleaseDate')
        biosSerial = Redact (Get-Prop $bios 'SerialNumber')
        deviceFingerprint = $fingerprint
        pendingReboot = $pendingReboot
        timeZone = Str $tz.Id
        utcOffsetMinutes = [int]$tz.GetUtcOffset((Get-Date)).TotalMinutes
    }
}

Invoke-Section 'secure_boot' {
    $enabled = Confirm-SecureBootUEFI -ErrorAction Stop
    @{ enabled = [bool]$enabled }
} -NeedsElevation $true

Invoke-Section 'cpu' {
    $items = @(Get-CimInstance Win32_Processor -ErrorAction Stop)
    $p = $items | Select-Object -First 1
    @{
        name = Str (Get-Prop $p 'Name')
        sockets = $items.Count
        physicalCores = Num (Get-Prop $p 'NumberOfCores')
        logicalProcessors = Num (Get-Prop $p 'NumberOfLogicalProcessors')
        maxClockMHz = Num (Get-Prop $p 'MaxClockSpeed')
        currentClockMHz = Num (Get-Prop $p 'CurrentClockSpeed')
        l2CacheKB = Num (Get-Prop $p 'L2CacheSize')
        l3CacheKB = Num (Get-Prop $p 'L3CacheSize')
        socket = Str (Get-Prop $p 'SocketDesignation')
        virtualisationFirmware = Get-Prop $p 'VirtualizationFirmwareEnabled'
        loadPercentage = Num (Get-Prop $p 'LoadPercentage')
    }
}

Invoke-Section 'memory' {
    $modules = @()
    $moduleError = $null
    try { $modules = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop) } catch { $moduleError = $_.Exception.Message }
    $array = $null
    $arrayError = $null
    try { $array = @(Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction Stop) | Select-Object -First 1 } catch { $arrayError = $_.Exception.Message }
    $pageFiles = @()
    try { $pageFiles = @(Get-CimInstance Win32_PageFileUsage -ErrorAction Stop) } catch { }
    @{
        modules = @($modules | ForEach-Object {
            @{
                slot = Str (Get-Prop $_ 'DeviceLocator')
                bank = Str (Get-Prop $_ 'BankLabel')
                capacityBytes = Num (Get-Prop $_ 'Capacity')
                speedMHz = Num (Get-Prop $_ 'Speed')
                configuredMHz = Num (Get-Prop $_ 'ConfiguredClockSpeed')
                manufacturer = Str (Get-Prop $_ 'Manufacturer')
                partNumber = Str (Get-Prop $_ 'PartNumber')
                memoryType = Num (Get-Prop $_ 'SMBIOSMemoryType')
                formFactor = Num (Get-Prop $_ 'FormFactor')
            }
        })
        moduleError = $moduleError
        arrayError = $arrayError
        slots = Num (Get-Prop $array 'MemoryDevices')
        maxCapacityKB = Num (Get-Prop $array 'MaxCapacity')
        maxCapacityExKB = Num (Get-Prop $array 'MaxCapacityEx')
        pageFiles = @($pageFiles | ForEach-Object {
            @{
                name = Str (Get-Prop $_ 'Name')
                allocatedMB = Num (Get-Prop $_ 'AllocatedBaseSize')
                currentUsageMB = Num (Get-Prop $_ 'CurrentUsage')
                peakUsageMB = Num (Get-Prop $_ 'PeakUsage')
            }
        })
    }
}

Invoke-Section 'gpu' {
    $controllers = @(Get-CimInstance Win32_VideoController -ErrorAction Stop)
    $registry = @()
    $registryError = $null
    try {
        $classKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
        $registry = @(Get-ChildItem -Path $classKey -ErrorAction Stop | Where-Object { $_.PSChildName -match '^\d{4}$' } | ForEach-Object {
            $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            $qw = Get-Prop $props 'HardwareInformation.qwMemorySize'
            $qwValue = $null
            if ($qw -is [byte[]]) { if ($qw.Length -ge 8) { $qwValue = [double][BitConverter]::ToUInt64($qw, 0) } }
            elseif ($null -ne $qw) { $qwValue = Num $qw }
            @{
                key = [string]$_.PSChildName
                driverDesc = Str (Get-Prop $props 'DriverDesc')
                matchingDeviceId = Str (Get-Prop $props 'MatchingDeviceId')
                qwMemorySize = $qwValue
            }
        })
    } catch { $registryError = $_.Exception.Message }
    @{
        controllers = @($controllers | ForEach-Object {
            @{
                name = Str (Get-Prop $_ 'Name')
                adapterRAM = Num (Get-Prop $_ 'AdapterRAM')
                driverVersion = Str (Get-Prop $_ 'DriverVersion')
                driverDate = Iso (Get-Prop $_ 'DriverDate')
                horizontal = Num (Get-Prop $_ 'CurrentHorizontalResolution')
                vertical = Num (Get-Prop $_ 'CurrentVerticalResolution')
                refreshHz = Num (Get-Prop $_ 'CurrentRefreshRate')
                pnpDeviceId = Str (Get-Prop $_ 'PNPDeviceID')
                status = Str (Get-Prop $_ 'Status')
                videoProcessor = Str (Get-Prop $_ 'VideoProcessor')
            }
        })
        registry = $registry
        registryError = $registryError
    }
}

Send-Progress 'Reading storage and battery'
Invoke-Section 'physical_disks' {
    $disks = @()
    $source = 'Get-PhysicalDisk'
    $reliabilityStatus = 'collected'
    $reliabilityError = $null
    try {
        $physical = @(Get-PhysicalDisk -ErrorAction Stop)
        foreach ($disk in $physical) {
            $rel = $null
            try { $rel = $disk | Get-StorageReliabilityCounter -ErrorAction Stop }
            catch {
                $rel = $null
                if ($reliabilityStatus -eq 'collected') {
                    $reliabilityError = [string]$_.Exception.Message
                    $reliabilityStatus = Get-FailureStatus $reliabilityError $true
                }
            }
            $disks += @{
                name = Str (Get-Prop $disk 'FriendlyName')
                mediaType = Str (Get-Prop $disk 'MediaType')
                busType = Str (Get-Prop $disk 'BusType')
                sizeBytes = Num (Get-Prop $disk 'Size')
                health = Str (Get-Prop $disk 'HealthStatus')
                operationalStatus = ((@(Get-Prop $disk 'OperationalStatus') | ForEach-Object { [string]$_ }) -join ', ')
                firmware = Str (Get-Prop $disk 'FirmwareVersion')
                serial = Redact (Get-Prop $disk 'SerialNumber')
                spindleSpeed = Num (Get-Prop $disk 'SpindleSpeed')
                deviceId = Str (Get-Prop $disk 'DeviceId')
                temperatureC = Num (Get-Prop $rel 'Temperature')
                temperatureMaxC = Num (Get-Prop $rel 'TemperatureMax')
                wearPercent = Num (Get-Prop $rel 'Wear')
                readErrorsTotal = Num (Get-Prop $rel 'ReadErrorsTotal')
                readErrorsUncorrected = Num (Get-Prop $rel 'ReadErrorsUncorrected')
                writeErrorsTotal = Num (Get-Prop $rel 'WriteErrorsTotal')
                writeErrorsUncorrected = Num (Get-Prop $rel 'WriteErrorsUncorrected')
                powerOnHours = Num (Get-Prop $rel 'PowerOnHours')
                startStopCycles = Num (Get-Prop $rel 'StartStopCycleCount')
            }
        }
    }
    catch {
        $source = 'Win32_DiskDrive'
        $reliabilityStatus = 'unavailable'
        $reliabilityError = [string]$_.Exception.Message
        foreach ($disk in @(Get-CimInstance Win32_DiskDrive -ErrorAction Stop)) {
            $disks += @{
                name = Str (Get-Prop $disk 'Model')
                mediaType = Str (Get-Prop $disk 'MediaType')
                busType = Str (Get-Prop $disk 'InterfaceType')
                sizeBytes = Num (Get-Prop $disk 'Size')
                health = $null
                smartStatus = Str (Get-Prop $disk 'Status')
                operationalStatus = $null
                firmware = Str (Get-Prop $disk 'FirmwareRevision')
                serial = Redact (Get-Prop $disk 'SerialNumber')
                spindleSpeed = $null
                deviceId = Str (Get-Prop $disk 'Index')
                temperatureC = $null
                temperatureMaxC = $null
                wearPercent = $null
                readErrorsTotal = $null
                readErrorsUncorrected = $null
                writeErrorsTotal = $null
                writeErrorsUncorrected = $null
                powerOnHours = $null
                startStopCycles = $null
            }
        }
    }
    @{ items = $disks; source = $source; reliabilityStatus = $reliabilityStatus; reliabilityError = $reliabilityError }
}

Invoke-Section 'logical_disks' {
    $items = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
    @{
        items = @($items | ForEach-Object {
            @{
                drive = Str (Get-Prop $_ 'DeviceID')
                label = Redact (Get-Prop $_ 'VolumeName')
                fileSystem = Str (Get-Prop $_ 'FileSystem')
                sizeBytes = Num (Get-Prop $_ 'Size')
                freeBytes = Num (Get-Prop $_ 'FreeSpace')
            }
        })
    }
}

Invoke-Section 'battery' {
    $battery = $null
    $batteryError = $null
    try { $battery = @(Get-CimInstance Win32_Battery -ErrorAction Stop) | Select-Object -First 1 } catch { $batteryError = [string]$_.Exception.Message }
    $result = @{
        present = ($null -ne $battery)
        queryError = $batteryError
        batteryStatus = Num (Get-Prop $battery 'BatteryStatus')
        statusText = Str (Get-Prop $battery 'Status')
        chargePercent = Num (Get-Prop $battery 'EstimatedChargeRemaining')
        estimatedMinutes = Num (Get-Prop $battery 'EstimatedRunTime')
        chemistry = Num (Get-Prop $battery 'Chemistry')
        designCapacity = $null
        fullCapacity = $null
        cycleCount = $null
        capacitySource = $null
        wmiErrors = @()
        powercfgStatus = 'not_run'
        powercfgError = $null
        powercfgDesign = $null
        powercfgFull = $null
        powercfgReportDeleted = $null
        activePowerPlan = $null
    }
    $design = 0.0
    $full = 0.0
    try {
        $static = @(Get-CimInstance -Namespace root/wmi -ClassName BatteryStaticData -ErrorAction Stop) | Select-Object -First 1
        $design = NumOr (Get-Prop $static 'DesignedCapacity') 0
    } catch { $result.wmiErrors += ('BatteryStaticData: ' + [string]$_.Exception.Message) }
    try {
        $fullClass = @(Get-CimInstance -Namespace root/wmi -ClassName BatteryFullChargedCapacity -ErrorAction Stop) | Select-Object -First 1
        $full = NumOr (Get-Prop $fullClass 'FullChargedCapacity') 0
    } catch { $result.wmiErrors += ('BatteryFullChargedCapacity: ' + [string]$_.Exception.Message) }
    try {
        $cycle = @(Get-CimInstance -Namespace root/wmi -ClassName BatteryCycleCount -ErrorAction Stop) | Select-Object -First 1
        $cycles = NumOr (Get-Prop $cycle 'CycleCount') 0
        if ($cycles -gt 0) { $result.cycleCount = $cycles }
    } catch { $result.wmiErrors += ('BatteryCycleCount: ' + [string]$_.Exception.Message) }
    if ($design -gt 0 -and $full -gt 0) { $result.capacitySource = 'Windows WMI (root/wmi)' }
    if ($result.present -and $BatteryReportPath) {
        try {
            $null = & "$env:SystemRoot\System32\powercfg.exe" /batteryreport /xml /output "$BatteryReportPath" 2>&1
            if (Test-Path -LiteralPath $BatteryReportPath) {
                [xml]$doc = Get-Content -LiteralPath $BatteryReportPath -Raw -ErrorAction Stop
                $reportBattery = @($doc.BatteryReport.Batteries.Battery) | Select-Object -First 1
                if ($null -ne $reportBattery) {
                    $reportDesign = NumOr (Get-Prop $reportBattery 'DesignCapacity') 0
                    $reportFull = NumOr (Get-Prop $reportBattery 'FullChargeCapacity') 0
                    $reportCycles = NumOr (Get-Prop $reportBattery 'CycleCount') 0
                    $result.powercfgStatus = 'collected'
                    $result.powercfgDesign = $reportDesign
                    $result.powercfgFull = $reportFull
                    if (($design -le 0 -or $full -le 0) -and $reportDesign -gt 0 -and $reportFull -gt 0) {
                        $design = $reportDesign
                        $full = $reportFull
                        $result.capacitySource = 'powercfg /batteryreport'
                    }
                    if ($null -eq $result.cycleCount -and $reportCycles -gt 0) { $result.cycleCount = $reportCycles }
                }
                else {
                    $result.powercfgStatus = 'partial'
                    $result.powercfgError = 'Battery element not found in the report XML'
                }
            }
            else {
                $result.powercfgStatus = 'failed'
                $result.powercfgError = 'powercfg did not create the XML report'
            }
        }
        catch {
            $result.powercfgStatus = 'failed'
            $result.powercfgError = [string]$_.Exception.Message
        }
        finally {
            if (Test-Path -LiteralPath $BatteryReportPath) {
                try { Remove-Item -LiteralPath $BatteryReportPath -Force -ErrorAction Stop; $result.powercfgReportDeleted = $true }
                catch { $result.powercfgReportDeleted = $false }
            }
            else { $result.powercfgReportDeleted = $true }
        }
    }
    if ($design -gt 0 -and $full -gt 0) {
        $result.designCapacity = $design
        $result.fullCapacity = $full
    }
    try { $result.activePowerPlan = ((@(& "$env:SystemRoot\System32\powercfg.exe" /getactivescheme 2>&1 | ForEach-Object { [string]$_ })) -join ' ').Trim() } catch { }
    $result
}

Invoke-Section 'thermal' {
    $zones = @(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop)
    @{
        items = @($zones | ForEach-Object {
            $raw = Num (Get-Prop $_ 'CurrentTemperature')
            $c = $null
            if ($null -ne $raw -and $raw -gt 0) { $c = [math]::Round(($raw / 10.0) - 273.15, 1) }
            @{ zone = Str (Get-Prop $_ 'InstanceName'); temperatureC = $c; active = Get-Prop $_ 'Active' }
        })
    }
} -NeedsElevation $true

Send-Progress 'Reading network, devices and security posture'
Invoke-Section 'network' {
    $adapters = @(Get-NetAdapter -Physical -ErrorAction Stop)
    @{
        items = @($adapters | ForEach-Object {
            @{
                name = Redact (Get-Prop $_ 'Name')
                description = Str (Get-Prop $_ 'InterfaceDescription')
                status = Str (Get-Prop $_ 'Status')
                linkSpeed = Str (Get-Prop $_ 'LinkSpeed')
                mediaType = Str (Get-Prop $_ 'MediaType')
                driver = Str (Get-Prop $_ 'DriverDescription')
                driverDate = Str (Get-Prop $_ 'DriverDate')
                driverVersion = Str (Get-Prop $_ 'DriverVersion')
                macAddress = Redact (Get-Prop $_ 'MacAddress')
            }
        })
    }
}

Invoke-Section 'wifi' {
    $lines = @(& "$env:SystemRoot\System32\netsh.exe" wlan show interfaces 2>&1 | ForEach-Object { [string]$_ })
    $state = Get-NetshValue $lines '^\s*State\s*:'
    $signal = Get-NetshValue $lines '^\s*Signal\s*:'
    $present = $null
    if ($signal -or $state) { $present = $true }
    elseif (($lines -join ' ') -match 'no wireless interface|not running|is not started') { $present = $false }
    @{
        present = $present
        parsed = [bool]($signal -or $state)
        lineCount = $lines.Count
        state = $state
        signal = $signal
        receiveMbps = Get-NetshValue $lines '^\s*Receive rate.*:'
        transmitMbps = Get-NetshValue $lines '^\s*Transmit rate.*:'
        radioType = Get-NetshValue $lines '^\s*Radio type\s*:'
        band = Get-NetshValue $lines '^\s*Band\s*:'
        channel = Get-NetshValue $lines '^\s*Channel\s*:'
        authentication = Get-NetshValue $lines '^\s*Authentication\s*:'
        ssid = Redact (Get-NetshValue $lines '^\s*SSID\s*:')
    }
}

Invoke-Section 'problem_devices' {
    $items = @(Get-CimInstance Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' -ErrorAction Stop)
    @{
        items = @($items | ForEach-Object {
            @{ device = Str (Get-Prop $_ 'Name'); errorCode = Num (Get-Prop $_ 'ConfigManagerErrorCode'); status = Str (Get-Prop $_ 'Status'); pnpClass = Str (Get-Prop $_ 'PNPClass') }
        })
    }
}

Invoke-Section 'tpm' {
    $t = Get-Tpm -ErrorAction Stop
    @{ present = [bool](Get-Prop $t 'TpmPresent' $false); ready = [bool](Get-Prop $t 'TpmReady' $false); enabled = [bool](Get-Prop $t 'TpmEnabled' $false); activated = [bool](Get-Prop $t 'TpmActivated' $false) }
} -NeedsElevation $true

Invoke-Section 'antivirus' {
    $items = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct -ErrorAction Stop)
    @{ items = @($items | ForEach-Object { @{ product = Str (Get-Prop $_ 'displayName'); productState = Num (Get-Prop $_ 'productState'); timestamp = Str (Get-Prop $_ 'timestamp') } }) }
}

Invoke-Section 'firewall' {
    $items = @(Get-NetFirewallProfile -ErrorAction Stop)
    @{ items = @($items | ForEach-Object { @{ name = Str (Get-Prop $_ 'Name'); enabled = Str (Get-Prop $_ 'Enabled'); defaultInbound = Str (Get-Prop $_ 'DefaultInboundAction'); defaultOutbound = Str (Get-Prop $_ 'DefaultOutboundAction') } }) }
}

Invoke-Section 'bitlocker' {
    $items = @(Get-BitLockerVolume -ErrorAction Stop)
    @{ items = @($items | ForEach-Object { @{ mountPoint = Str (Get-Prop $_ 'MountPoint'); volumeStatus = Str (Get-Prop $_ 'VolumeStatus'); protectionStatus = Str (Get-Prop $_ 'ProtectionStatus'); encryptionMethod = Str (Get-Prop $_ 'EncryptionMethod'); encryptionPercentage = Num (Get-Prop $_ 'EncryptionPercentage') } }) }
} -NeedsElevation $true

Invoke-Section 'defender' {
    $s = Get-MpComputerStatus -ErrorAction Stop
    @{
        amServiceEnabled = Get-Prop $s 'AMServiceEnabled'
        antivirusEnabled = Get-Prop $s 'AntivirusEnabled'
        realTimeProtection = Get-Prop $s 'RealTimeProtectionEnabled'
        signatureAgeDays = Num (Get-Prop $s 'AntivirusSignatureAge')
        quickScanAgeDays = Num (Get-Prop $s 'QuickScanAge')
        fullScanAgeDays = Num (Get-Prop $s 'FullScanAge')
        tamperProtected = Get-Prop $s 'IsTamperProtected'
        amRunningMode = Str (Get-Prop $s 'AMRunningMode')
    }
}

Send-Progress 'Reading startup items, services and updates'
Invoke-Section 'startup' {
    $items = @(Get-CimInstance Win32_StartupCommand -ErrorAction Stop)
    @{ items = @($items | ForEach-Object { @{ name = Str (Get-Prop $_ 'Name'); command = Redact (Get-Prop $_ 'Command'); location = Redact (Get-Prop $_ 'Location'); user = Redact (Get-Prop $_ 'User') } }) }
}

Invoke-Section 'services' {
    $items = @(Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -ErrorAction Stop)
    @{ items = @($items | ForEach-Object { @{ service = Str (Get-Prop $_ 'DisplayName'); name = Str (Get-Prop $_ 'Name'); state = Str (Get-Prop $_ 'State'); startMode = Str (Get-Prop $_ 'StartMode'); delayed = Get-Prop $_ 'DelayedAutoStart' } }) }
}

Invoke-Section 'hotfixes' {
    $items = @(Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 15)
    @{ items = @($items | ForEach-Object { @{ hotFixId = Str (Get-Prop $_ 'HotFixID'); description = Str (Get-Prop $_ 'Description'); installedOn = Iso (Get-Prop $_ 'InstalledOn') } }) }
}

Send-Progress ('Reading event logs for the last {0} days' -f $EventDays)
Invoke-Section 'events' {
    $start = (Get-Date).AddDays(-$EventDays)
    $logs = @()
    foreach ($log in @('System', 'Application')) {
        $entry = @{ log = $log; status = 'collected'; error = $null; totalEvents = 0; cap = 2000; capped = $false; groups = @() }
        $events = @()
        try {
            $events = @(Get-WinEvent -FilterHashtable @{ LogName = $log; Level = @(1, 2); StartTime = $start } -MaxEvents 2000 -ErrorAction Stop)
        }
        catch {
            $events = @()
            $message = [string]$_.Exception.Message
            if ($message -notmatch 'No events were found') {
                $entry.status = Get-FailureStatus $message $false
                $entry.error = $message
            }
        }
        $entry.totalEvents = $events.Count
        $entry.capped = ($events.Count -ge 2000)
        $entry.groups = @($events | Group-Object ProviderName, Id | Sort-Object Count -Descending | ForEach-Object {
            $sample = $_.Group | Select-Object -First 1
            $message = $null
            if ($IncludeSensitiveData) { $message = Str (Get-Prop $sample 'Message') }
            if ($message -and $message.Length -gt 240) { $message = $message.Substring(0, 240) }
            @{
                provider = Str (Get-Prop $sample 'ProviderName')
                eventId = Num (Get-Prop $sample 'Id')
                level = Num (Get-Prop $sample 'Level')
                levelName = Str (Get-Prop $sample 'LevelDisplayName')
                count = $_.Count
                mostRecent = Iso (Get-Prop $sample 'TimeCreated')
                message = $message
            }
        })
        $logs += $entry
    }
    $providers = @('Microsoft-Windows-WHEA-Logger', 'Microsoft-Windows-Kernel-Power', 'Microsoft-Windows-Kernel-Processor-Power', 'Microsoft-Windows-Kernel-Boot', 'Microsoft-Windows-Kernel-General', 'Microsoft-Windows-Kernel-PnP', 'Microsoft-Windows-Power-Troubleshooter', 'disk', 'Disk', 'Microsoft-Windows-Disk', 'stornvme', 'storahci', 'storport', 'iaStorA', 'iaStorAC', 'iaStorAVC', 'iaStorV', 'iaStorE', 'Ntfs', 'Microsoft-Windows-Ntfs', 'volmgr', 'partmgr', 'Microsoft-Windows-StorDiag', 'nvlddmkm', 'amdkmdag', 'igfx', 'igfxn', 'Display')
    $hardware = @{ status = 'collected'; error = $null; totalEvents = 0; cap = 1000; capped = $false; providers = $providers; groups = @() }
    $warn = @()
    try {
        $warn = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = $providers; Level = @(1, 2, 3); StartTime = $start } -MaxEvents 1000 -ErrorAction Stop)
    }
    catch {
        $warn = @()
        $message = [string]$_.Exception.Message
        if ($message -notmatch 'No events were found') {
            $hardware.status = Get-FailureStatus $message $false
            $hardware.error = $message
        }
    }
    $hardware.totalEvents = $warn.Count
    $hardware.capped = ($warn.Count -ge 1000)
    $hardware.groups = @($warn | Group-Object ProviderName, Id, Level | Sort-Object Count -Descending | ForEach-Object {
        $sample = $_.Group | Select-Object -First 1
        $message = Str (Get-Prop $sample 'Message')
        if ($message -and $message.Length -gt 200) { $message = $message.Substring(0, 200) }
        @{
            provider = Str (Get-Prop $sample 'ProviderName')
            eventId = Num (Get-Prop $sample 'Id')
            level = Num (Get-Prop $sample 'Level')
            levelName = Str (Get-Prop $sample 'LevelDisplayName')
            count = $_.Count
            mostRecent = Iso (Get-Prop $sample 'TimeCreated')
            message = $message
        }
    })
    @{ days = $EventDays; logs = $logs; hardware = $hardware }
}

if ($Mode -ne 'Quick') {
    Invoke-Section 'reliability' {
        $start = (Get-Date).AddDays(-$EventDays)
        $records = @(Get-CimInstance Win32_ReliabilityRecords -ErrorAction Stop |
            Where-Object { $_.TimeGenerated -ge $start -and $_.SourceName -match 'Application (Hang|Error)|Windows Error Reporting|Hardware error' } |
            Sort-Object TimeGenerated -Descending)
        @{
            total = $records.Count
            cap = 40
            items = @($records | Select-Object -First 40 | ForEach-Object {
                @{ time = Iso (Get-Prop $_ 'TimeGenerated'); source = Str (Get-Prop $_ 'SourceName'); product = Str (Get-Prop $_ 'ProductName'); eventId = Num (Get-Prop $_ 'EventIdentifier') }
            })
        }
    }
}
else {
    Emit @{ section = 'reliability'; status = 'not_run'; data = $null; error = 'Not collected in Quick mode.' }
}

if ($IncludeInstalledApps) {
    Invoke-Section 'installed_apps' {
        $roots = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        if ($IncludeSensitiveData) { $roots += 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' }
        $apps = @($roots | ForEach-Object { Get-ItemProperty -Path $_ -ErrorAction SilentlyContinue } |
            Where-Object { $_.DisplayName } |
            Sort-Object DisplayName, DisplayVersion -Unique)
        @{ items = @($apps | ForEach-Object { @{ application = Str (Get-Prop $_ 'DisplayName'); version = Str (Get-Prop $_ 'DisplayVersion'); publisher = Str (Get-Prop $_ 'Publisher'); installDate = Str (Get-Prop $_ 'InstallDate') } }) }
    }
}
else {
    Emit @{ section = 'installed_apps'; status = 'not_run'; data = $null; error = 'Not requested.' }
}

if (-not $SkipSamples) {
    Send-Progress ('Collecting {0} live performance samples (each sample takes several counter reads)' -f $SampleCount)
    $logical = 1
    try { $logical = [int](NumOr ((@(Get-CimInstance Win32_ComputerSystem -ErrorAction Stop) | Select-Object -First 1).NumberOfLogicalProcessors) 1) } catch { }
    foreach ($className in @('Win32_PerfFormattedData_PerfOS_Processor', 'Win32_PerfFormattedData_Counters_ProcessorInformation', 'Win32_PerfFormattedData_PerfOS_Memory', 'Win32_PerfFormattedData_PerfDisk_PhysicalDisk', 'Win32_PerfRawData_PerfDisk_PhysicalDisk', 'Win32_PerfFormattedData_PerfProc_Process')) {
        try { $null = Get-CimInstance $className -ErrorAction Stop } catch { }
    }
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $SampleCount; $i++) {
        $iterationStart = $watch.Elapsed
        $sample = @{ index = $i; time = (Get-Date).ToString('o'); elapsedMs = [int]$watch.ElapsedMilliseconds; cpu = $null; cpuError = $null; frequency = $null; frequencyError = $null; memory = $null; memoryError = $null; disks = $null; disksError = $null; disksRaw = $null; disksRawError = $null; processes = $null; processesError = $null; power = $null; powerError = $null; durationMs = 0 }
        try {
            $cores = @(Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -ErrorAction Stop)
            $total = $null
            $perCore = @()
            foreach ($core in $cores) {
                $name = [string]$core.Name
                $value = Num $core.PercentProcessorTime
                if ($name -eq '_Total') { $total = $value } else { $perCore += @{ name = $name; percent = $value } }
            }
            $sample.cpu = @{ total = $total; cores = $perCore }
        } catch { $sample.cpuError = [string]$_.Exception.Message }
        try {
            $info = @(Get-CimInstance Win32_PerfFormattedData_Counters_ProcessorInformation -Filter "Name='_Total'" -ErrorAction Stop) | Select-Object -First 1
            $sample.frequency = @{
                percentPerformance = Num (Get-Prop $info 'PercentProcessorPerformance')
                percentMaxFrequency = Num (Get-Prop $info 'PercentofMaximumFrequency')
                frequencyMHz = Num (Get-Prop $info 'ProcessorFrequency')
                percentProcessorTime = Num (Get-Prop $info 'PercentProcessorTime')
            }
        } catch { $sample.frequencyError = [string]$_.Exception.Message }
        try {
            $mem = @(Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory -ErrorAction Stop) | Select-Object -First 1
            $sample.memory = @{
                availableMB = Num (Get-Prop $mem 'AvailableMBytes')
                percentCommitted = Num (Get-Prop $mem 'PercentCommittedBytesInUse')
                committedBytes = Num (Get-Prop $mem 'CommittedBytes')
                commitLimit = Num (Get-Prop $mem 'CommitLimit')
                pagesInputPerSec = Num (Get-Prop $mem 'PagesInputPersec')
                pagesPerSec = Num (Get-Prop $mem 'PagesPersec')
                pageFaultsPerSec = Num (Get-Prop $mem 'PageFaultsPersec')
                cacheBytes = Num (Get-Prop $mem 'CacheBytes')
            }
        } catch { $sample.memoryError = [string]$_.Exception.Message }
        try {
            $disks = @(Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -ErrorAction Stop)
            $sample.disks = @($disks | ForEach-Object {
                @{
                    name = [string]$_.Name
                    percentDiskTime = Num $_.PercentDiskTime
                    percentIdleTime = Num $_.PercentIdleTime
                    currentQueue = Num $_.CurrentDiskQueueLength
                    avgQueue = Num $_.AvgDiskQueueLength
                    readBytesPerSec = Num $_.DiskReadBytesPersec
                    writeBytesPerSec = Num $_.DiskWriteBytesPersec
                    transfersPerSec = Num $_.DiskTransfersPersec
                    splitIOPerSec = Num $_.SplitIOPerSec
                }
            })
        } catch { $sample.disksError = [string]$_.Exception.Message }
        try {
            $raw = @(Get-CimInstance Win32_PerfRawData_PerfDisk_PhysicalDisk -ErrorAction Stop)
            $sample.disksRaw = @($raw | ForEach-Object {
                @{
                    name = [string]$_.Name
                    secPerTransfer = Num $_.AvgDisksecPerTransfer
                    secPerTransferBase = Num $_.AvgDisksecPerTransfer_Base
                    secPerRead = Num $_.AvgDisksecPerRead
                    secPerReadBase = Num $_.AvgDisksecPerRead_Base
                    secPerWrite = Num $_.AvgDisksecPerWrite
                    secPerWriteBase = Num $_.AvgDisksecPerWrite_Base
                    frequencyPerfTime = Num $_.Frequency_PerfTime
                    timestampPerfTime = Num $_.Timestamp_PerfTime
                }
            })
        } catch { $sample.disksRawError = [string]$_.Exception.Message }
        try {
            $procs = @(Get-CimInstance Win32_PerfFormattedData_PerfProc_Process -ErrorAction Stop)
            $sample.processes = @($procs | Where-Object { $_.Name -ne '_Total' -and $_.Name -ne 'Idle' -and $_.IDProcess -gt 0 } | ForEach-Object {
                @{ name = [string]$_.Name; pid = [int]$_.IDProcess; cpu = Num $_.PercentProcessorTime; privateMB = [math]::Round((NumOr $_.WorkingSetPrivate 0) / 1MB, 1); ioReadBytesPerSec = Num $_.IOReadBytesPersec; ioWriteBytesPerSec = Num $_.IOWriteBytesPersec }
            })
        } catch { $sample.processesError = [string]$_.Exception.Message }
        try {
            $bat = @(Get-CimInstance Win32_Battery -ErrorAction Stop) | Select-Object -First 1
            if ($null -eq $bat) { $sample.power = @{ batteryPresent = $false; batteryStatus = $null; chargePercent = $null } }
            else { $sample.power = @{ batteryPresent = $true; batteryStatus = Num (Get-Prop $bat 'BatteryStatus'); chargePercent = Num (Get-Prop $bat 'EstimatedChargeRemaining') } }
        } catch { $sample.powerError = [string]$_.Exception.Message }
        $sample.durationMs = [int]($watch.Elapsed - $iterationStart).TotalMilliseconds
        Emit @{ section = 'sample'; status = 'ok'; data = $sample; index = $i; total = $SampleCount }
        $iterationSeconds = ($watch.Elapsed - $iterationStart).TotalSeconds
        if ($iterationSeconds -lt 1 -and $i -lt ($SampleCount - 1)) { Start-Sleep -Milliseconds ([int]((1 - $iterationSeconds) * 1000)) }
    }
    Emit @{ section = 'samples_done'; status = 'ok'; data = @{ requested = $SampleCount; elapsedSeconds = [math]::Round($watch.Elapsed.TotalSeconds, 1); logicalProcessors = $logical } }

    Invoke-Section 'top_ram' {
        $groups = @(Get-Process -ErrorAction Stop | Group-Object ProcessName)
        @{
            items = @($groups | ForEach-Object {
                $ws = ($_.Group | Measure-Object -Property WorkingSet64 -Sum).Sum
                $priv = ($_.Group | Measure-Object -Property PrivateMemorySize64 -Sum).Sum
                @{ process = [string]$_.Name; instances = [int]$_.Count; ramMB = [math]::Round((NumOr $ws 0) / 1MB, 0); privateMB = [math]::Round((NumOr $priv 0) / 1MB, 0) }
            })
        }
    }
}
else {
    Emit @{ section = 'samples_done'; status = 'not_run'; data = $null; error = 'Sampling skipped.' }
}

if ($Mode -eq 'Deep') {
    if ($script:IsAdmin) {
        Send-Progress 'Running DISM /CheckHealth (verification only)'
        Invoke-Section 'dism_checkhealth' {
            $output = @(& "$env:SystemRoot\System32\Dism.exe" /Online /Cleanup-Image /CheckHealth 2>&1 | ForEach-Object { [string]$_ })
            $code = $LASTEXITCODE
            @{ exitCode = $code; output = (($output -join "`n") -replace "`0", '') }
        }
        Send-Progress 'Running SFC /verifyonly (verification only; this can take 5 to 15 minutes)'
        Invoke-Section 'sfc_verifyonly' {
            $previous = $null
            try { $previous = [Console]::OutputEncoding } catch { }
            $output = @()
            $code = $null
            try {
                try { [Console]::OutputEncoding = [Text.Encoding]::Unicode } catch { }
                $output = @(& "$env:SystemRoot\System32\sfc.exe" /verifyonly 2>&1 | ForEach-Object { [string]$_ })
                $code = $LASTEXITCODE
            }
            finally {
                if ($null -ne $previous) { try { [Console]::OutputEncoding = $previous } catch { } }
            }
            @{ exitCode = $code; output = (($output -join "`n") -replace "`0", '') }
        }
    }
    else {
        Emit @{ section = 'dism_checkhealth'; status = 'requires_elevation'; data = $null; error = 'Deep verification requires Run as administrator.' }
        Emit @{ section = 'sfc_verifyonly'; status = 'requires_elevation'; data = $null; error = 'Deep verification requires Run as administrator.' }
    }
}
else {
    Emit @{ section = 'dism_checkhealth'; status = 'not_run'; data = $null; error = 'Deep mode not selected.' }
    Emit @{ section = 'sfc_verifyonly'; status = 'not_run'; data = $null; error = 'Deep mode not selected.' }
}

Emit @{ section = 'done'; status = 'ok'; data = @{ endTime = (Get-Date).ToString('o'); isAdmin = $script:IsAdmin } }

}
catch {
    $message = ''
    try { $message = [string]$_.Exception.Message } catch { $message = [string]$_ }
    $position = ''
    try { $position = [string]$_.InvocationInfo.PositionMessage } catch { }
    Emit @{ section = 'fatal'; status = 'failed'; data = $null; error = $message; position = $position }
}
"""


# ---------------------------------------------------------------------------
# Collector runner
# ---------------------------------------------------------------------------

class AuditOptions:
    """Settings for one audit run."""

    def __init__(
        self,
        mode: str = "Standard",
        sample_count: Optional[int] = None,
        output_dir: Optional[str] = None,
        include_sensitive: bool = False,
        include_apps: bool = False,
        create_zip: bool = False,
        open_report: bool = False,
        workload_label: str = "Normal work",
        workload_notes: str = "",
        docked: Optional[bool] = None,
        skip_samples: bool = False,
    ) -> None:
        self.mode = mode if mode in MODES else "Standard"
        count = sample_count if sample_count is not None else DEFAULT_SAMPLES[self.mode]
        self.sample_count = int(min(max(int(count), SAMPLE_MIN), SAMPLE_MAX))
        self.event_days = EVENT_DAYS[self.mode]
        self.output_dir = output_dir or default_output_dir()
        self.include_sensitive = bool(include_sensitive)
        self.include_apps = bool(include_apps)
        self.create_zip = bool(create_zip)
        self.open_report = bool(open_report)
        self.workload_label = clean_text(workload_label, 120) or "Not stated"
        self.workload_notes = clean_text(workload_notes, 1000)
        self.docked = docked
        self.skip_samples = bool(skip_samples)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "Mode": self.mode,
            "SampleCount": self.sample_count,
            "EventDays": self.event_days,
            "IncludeSensitiveData": self.include_sensitive,
            "IncludeInstalledApps": self.include_apps,
            "WorkloadLabel": self.workload_label,
            "WorkloadNotes": self.workload_notes,
            "DockConnected": self.docked,
        }


def powershell_executable() -> str:
    """Prefer the 64-bit Windows PowerShell 5.1 host; fall back to PATH."""
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    candidates = [
        os.path.join(system_root, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe"),
        os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    found = shutil.which("powershell.exe") or shutil.which("powershell")
    if found:
        return found
    raise RuntimeError(
        "Windows PowerShell 5.1 (powershell.exe) was not found. The audit needs it to read Windows counters."
    )


def default_launcher(script_path: str, script_args: List[str]) -> List[str]:
    return [
        powershell_executable(),
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_path,
    ] + script_args


def hidden_popen_kwargs() -> Dict[str, Any]:
    """Keep the PowerShell child from opening a console window under pythonw."""
    kwargs: Dict[str, Any] = {}
    if IS_WINDOWS:
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            kwargs["startupinfo"] = startupinfo
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        except AttributeError:
            pass
    return kwargs


def run_collector(
    options: AuditOptions,
    on_record: Callable[[Dict[str, Any]], None],
    cancel_event: Optional[threading.Event] = None,
    launcher: Callable[[str, List[str]], List[str]] = default_launcher,
) -> Dict[str, Any]:
    """Run the embedded collector once and stream its records to on_record.

    Returns run metadata: exit code, stderr text, temporary-file clean-up
    results and whether the run was cancelled. Never raises for collector
    failures; a launch failure is raised because nothing was collected.
    """
    meta: Dict[str, Any] = {
        "ExitCode": None,
        "Stderr": "",
        "TempCleanup": [],
        "Cancelled": False,
        "Started": iso(now_local()),
        "Finished": None,
        "Command": None,
    }
    temp_dir = tempfile.mkdtemp(prefix="PC_Health_Audit_")
    script_path = os.path.join(temp_dir, "PC_Health_Collector.ps1")
    battery_path = os.path.join(temp_dir, "PC_Health_Battery.xml")
    process: Optional[subprocess.Popen] = None
    try:
        with open(script_path, "w", encoding="utf-8-sig", newline="\r\n") as handle:
            handle.write(POWERSHELL_COLLECTOR)
        script_args = [
            "-Mode", options.mode,
            "-SampleCount", str(options.sample_count),
            "-EventDays", str(options.event_days),
            "-BatteryReportPath", battery_path,
        ]
        if options.include_sensitive:
            script_args.append("-IncludeSensitiveData")
        if options.include_apps:
            script_args.append("-IncludeInstalledApps")
        if options.skip_samples:
            script_args.append("-SkipSamples")
        command = launcher(script_path, script_args)
        meta["Command"] = " ".join(command[:6]) + " ..."
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            **hidden_popen_kwargs(),
        )
        stderr_chunks: List[bytes] = []

        def drain_stderr() -> None:
            try:
                assert process is not None and process.stderr is not None
                stderr_chunks.append(process.stderr.read())
            except Exception:
                pass

        def watch_cancel() -> None:
            assert process is not None and cancel_event is not None
            while process.poll() is None:
                if cancel_event.wait(0.25):
                    meta["Cancelled"] = True
                    try:
                        process.kill()
                    except Exception:
                        pass
                    return

        stderr_thread = threading.Thread(target=drain_stderr, daemon=True)
        stderr_thread.start()
        if cancel_event is not None:
            threading.Thread(target=watch_cancel, daemon=True).start()

        assert process.stdout is not None
        for raw_line in iter(process.stdout.readline, b""):
            line = raw_line.decode("utf-8", errors="replace").strip()
            if not line:
                continue
            try:
                record = json.loads(line)
            except ValueError:
                record = {"section": "noise", "status": "ok", "text": line[:500]}
            if not isinstance(record, dict) or "section" not in record:
                record = {"section": "noise", "status": "ok", "text": line[:500]}
            on_record(record)
        process.wait()
        stderr_thread.join(timeout=5)
        meta["ExitCode"] = process.returncode
        meta["Stderr"] = b"".join(stderr_chunks).decode("utf-8", errors="replace").replace("\x00", "").strip()[:4000]
    finally:
        meta["Finished"] = iso(now_local())
        if process is not None and process.poll() is None:
            try:
                process.kill()
            except Exception:
                pass
        for path in (battery_path, script_path):
            if os.path.exists(path):
                try:
                    os.remove(path)
                    meta["TempCleanup"].append({"Path": path, "Deleted": True, "Error": None})
                except OSError as exc:
                    meta["TempCleanup"].append({"Path": path, "Deleted": False, "Error": str(exc)})
        try:
            os.rmdir(temp_dir)
        except OSError as exc:
            if os.path.isdir(temp_dir):
                meta["TempCleanup"].append({"Path": temp_dir, "Deleted": False, "Error": str(exc)})
    return meta


# ---------------------------------------------------------------------------
# Analysis: turn collector records into a report (schema 1.2)
# ---------------------------------------------------------------------------

SECTION_LABELS = {
    "context": "Collector context",
    "system": "System and Windows",
    "secure_boot": "Secure Boot",
    "cpu": "Processor",
    "memory": "Memory modules",
    "gpu": "Graphics",
    "physical_disks": "Physical disks",
    "logical_disks": "Volumes",
    "battery": "Battery",
    "thermal": "ACPI thermal zones",
    "network": "Network adapters",
    "wifi": "Wi-Fi link",
    "problem_devices": "Device Manager errors",
    "tpm": "TPM",
    "antivirus": "Antivirus products",
    "firewall": "Firewall profiles",
    "bitlocker": "BitLocker",
    "defender": "Microsoft Defender status",
    "startup": "Startup entries",
    "services": "Automatic services not running",
    "hotfixes": "Recent hotfixes",
    "events": "Event logs",
    "reliability": "Reliability records",
    "installed_apps": "Installed applications",
    "samples_done": "Live performance samples",
    "top_ram": "RAM by application",
    "dism_checkhealth": "DISM /CheckHealth",
    "sfc_verifyonly": "SFC /verifyonly",
}
DIAGNOSTIC_SECTIONS = (
    "system", "cpu", "memory", "physical_disks", "logical_disks", "battery", "events", "samples_done", "top_ram",
    "problem_devices", "firewall", "antivirus", "gpu", "network", "startup", "services", "hotfixes", "reliability",
)
ELEVATION_SECTIONS = ("secure_boot", "tpm", "bitlocker", "thermal")

BATTERY_STATUS_TEXT = {
    1: "Discharging", 2: "On mains (not charging)", 3: "Fully charged", 4: "Low", 5: "Critical", 6: "Charging",
    7: "Charging and high", 8: "Charging and low", 9: "Charging and critical", 10: "Undefined", 11: "Partially charged",
}
MAINS_STATUSES = {2, 3, 6, 7, 8, 9}
BATTERY_STATUSES = {1, 4, 5}


class RecordSet:
    """Indexes collector records by section for analysis."""

    def __init__(self, records: Iterable[Dict[str, Any]]) -> None:
        self.sections: Dict[str, Dict[str, Any]] = {}
        self.samples: List[Dict[str, Any]] = []
        self.progress: List[str] = []
        self.noise: List[str] = []
        self.fatal: Optional[Dict[str, Any]] = None
        for record in records:
            if not isinstance(record, dict):
                continue
            name = str(record.get("section", ""))
            if name == "sample":
                data = record.get("data")
                if isinstance(data, dict):
                    self.samples.append(data)
            elif name == "progress":
                self.progress.append(clean_text(record.get("message")))
            elif name == "noise":
                self.noise.append(clean_text(record.get("text"), 500))
            elif name == "fatal":
                self.fatal = record
            elif name:
                self.sections[name] = record

    def status(self, name: str) -> str:
        """Absent sections count as failed: the collector emits explicit not_run records."""
        record = self.sections.get(name)
        if record is None:
            return "failed"
        status = str(record.get("status", "failed"))
        return "collected" if status == "ok" else status

    def data(self, name: str) -> Optional[Dict[str, Any]]:
        record = self.sections.get(name)
        if record is None or record.get("status") != "ok":
            return None
        data = record.get("data")
        return data if isinstance(data, dict) else None

    def error(self, name: str) -> str:
        record = self.sections.get(name)
        if record is None:
            return "Not collected: the collector ended before this check ran."
        return single_line(record.get("error") or record.get("note"), 400)

    def items(self, name: str) -> List[Dict[str, Any]]:
        data = self.data(name)
        return as_dict_list(data.get("items")) if data else []


def status_label(status: str) -> str:
    return {
        "collected": "Collected",
        "partial": "Partially collected",
        "unavailable": "Unavailable on this device",
        "requires_elevation": "Requires elevation",
        "failed": "Failed",
        "not_applicable": "Not applicable",
        "not_run": "Not run",
    }.get(status, status)


def secure_boot_text(records: RecordSet) -> str:
    status = records.status("secure_boot")
    if status == "collected":
        data = records.data("secure_boot") or {}
        return "Enabled" if data.get("enabled") else "Disabled"
    if status == "requires_elevation":
        return "Not checked (requires elevation)"
    if status == "not_applicable":
        return "Not applicable (legacy BIOS boot)"
    return f"Unknown ({status_label(status)})"


def decode_antivirus_state(state: Any) -> str:
    value = to_number(state)
    if value is None:
        return "Unknown"
    code = int(value)
    enabled = ((code >> 12) & 0xF) == 1
    up_to_date = (code & 0xF0) == 0
    return f"{'Enabled' if enabled else 'Disabled or snoozed'}; {'up to date' if up_to_date else 'signatures out of date'} (decoded from productState 0x{code:06X}, indicative)"


def match_gpu_memory(controller: Dict[str, Any], registry: List[Dict[str, Any]]) -> Tuple[Optional[float], str]:
    name = clean_text(controller.get("name")).lower()
    pnp = clean_text(controller.get("pnpDeviceId")).upper()
    by_name = [entry for entry in registry if clean_text(entry.get("driverDesc")).lower() == name and to_number(entry.get("qwMemorySize"))]
    if by_name:
        return to_number(by_name[0].get("qwMemorySize")), "Registry HardwareInformation.qwMemorySize (64-bit)"
    for entry in registry:
        match_id = clean_text(entry.get("matchingDeviceId")).upper()
        if match_id and pnp.startswith(match_id) and to_number(entry.get("qwMemorySize")):
            return to_number(entry.get("qwMemorySize")), "Registry HardwareInformation.qwMemorySize (64-bit)"
    adapter_ram = to_number(controller.get("adapterRAM"))
    if adapter_ram:
        note = "Win32_VideoController.AdapterRAM (32-bit field; unverified for adapters with 4 GB or more)"
        if adapter_ram >= 4 * 1024 ** 3 - 16 * 1024 ** 2:
            note = "Win32_VideoController.AdapterRAM saturated at the 32-bit limit; true size is larger"
        return adapter_ram, note
    return None, "Not reported"


def classify_dism(status: str, data: Optional[Dict[str, Any]], error: str) -> Dict[str, Any]:
    result = {"Status": status_label(status), "Verdict": "Not run", "ExitCode": None, "Output": error}
    if status != "collected" or not data:
        return result
    output = clean_text(data.get("output"), 4000)
    code = to_number(data.get("exitCode"))
    result["ExitCode"] = int(code) if code is not None else None
    result["Output"] = output
    lower = output.lower()
    if "no component store corruption detected" in lower:
        result["Verdict"] = "No corruption detected"
    elif "component store is repairable" in lower:
        result["Verdict"] = "Corruption detected (repairable)"
    elif "not repairable" in lower:
        result["Verdict"] = "Corruption detected (not repairable)"
    elif code not in (None, 0):
        result["Verdict"] = f"Command failed (exit code {int(code)})"
    else:
        result["Verdict"] = "Inconclusive (output not recognised; non-English Windows or unexpected text)"
    return result


def classify_sfc(status: str, data: Optional[Dict[str, Any]], error: str) -> Dict[str, Any]:
    result = {"Status": status_label(status), "Verdict": "Not run", "ExitCode": None, "Output": error}
    if status != "collected" or not data:
        return result
    output = clean_text(data.get("output"), 4000)
    code = to_number(data.get("exitCode"))
    result["ExitCode"] = int(code) if code is not None else None
    result["Output"] = output
    lower = output.lower()
    if "did not find any integrity violations" in lower:
        result["Verdict"] = "No integrity violations"
    elif "found integrity violations" in lower or "found corrupt files" in lower:
        result["Verdict"] = "Integrity violations found"
    elif "could not perform the requested operation" in lower or "could not start the repair service" in lower:
        result["Verdict"] = "Command could not run"
    elif "another servicing or repair operation" in lower or "pending system repair" in lower:
        result["Verdict"] = "Blocked by a pending repair or restart"
    elif code not in (None, 0):
        result["Verdict"] = f"Command failed (exit code {int(code)})"
    else:
        result["Verdict"] = "Inconclusive (output not recognised; non-English Windows or unexpected text)"
    return result


def build_coverage(records: RecordSet, run_meta: Dict[str, Any], options: AuditOptions) -> List[Dict[str, Any]]:
    coverage: List[Dict[str, Any]] = []
    for name, label in SECTION_LABELS.items():
        if name == "context":
            continue
        status = records.status(name)
        detail = ""
        if status == "collected":
            data = records.data(name) or {}
            if name == "physical_disks":
                rel = str(data.get("reliabilityStatus", "collected"))
                detail = f"{len(as_dict_list(data.get('items')))} disk(s) via {data.get('source')}"
                if rel != "collected":
                    status = "partial"
                    detail += f"; SMART temperature/wear counters {status_label(rel).lower()}: {clean_text(data.get('reliabilityError'), 160)}"
            elif name == "battery":
                if data.get("present"):
                    detail = f"capacity source: {data.get('capacitySource') or 'not exposed'}; powercfg report {data.get('powercfgStatus')}"
                    if data.get("powercfgReportDeleted") is False:
                        detail += "; WARNING temporary battery XML could not be deleted"
                        status = "partial"
                    if not data.get("capacitySource"):
                        status = "partial"
                else:
                    detail = "no battery reported (desktop or virtual machine)"
                    status = "not_applicable"
            elif name == "events":
                parts = []
                for log in as_dict_list(data.get("logs")):
                    log_status = str(log.get("status", "collected"))
                    text = f"{log.get('log')}: {int(to_number(log.get('totalEvents')) or 0)} error/critical events"
                    if log.get("capped"):
                        text += f" (capped at {int(to_number(log.get('cap')) or 0)})"
                    if log_status != "collected":
                        text += f" {status_label(log_status).lower()}: {clean_text(log.get('error'), 120)}"
                        status = "partial"
                    parts.append(text)
                hardware = data.get("hardware") if isinstance(data.get("hardware"), dict) else {}
                if hardware:
                    text = f"hardware providers: {int(to_number(hardware.get('totalEvents')) or 0)} events incl. warnings"
                    if hardware.get("capped"):
                        text += f" (capped at {int(to_number(hardware.get('cap')) or 0)})"
                    if str(hardware.get("status", "collected")) != "collected":
                        text += f" {status_label(str(hardware.get('status'))).lower()}"
                        status = "partial"
                    parts.append(text)
                detail = "; ".join(parts)
            elif name == "samples_done":
                taken = len(records.samples)
                requested = int(to_number(data.get("requested")) or options.sample_count)
                counts = sample_validity(records.samples)
                detail = (
                    f"{taken}/{requested} samples in {fmt_value(data.get('elapsedSeconds'), ' s', 0)}; valid reads: "
                    f"CPU {counts['cpu']}, per-core {counts['cores']}, frequency {counts['frequency']}, memory {counts['memory']}, "
                    f"disk {counts['disks']}, disk latency {counts['disksRaw']}, processes {counts['processes']}, power {counts['power']}"
                )
                if taken < requested or any(counts[key] < taken for key in ("cpu", "memory", "disks", "processes")):
                    status = "partial"
                if taken == 0:
                    status = "failed"
                    detail = "no samples were returned"
            elif name == "wifi":
                if data.get("present") is False:
                    status = "not_applicable"
                    detail = "no wireless interface"
                elif not data.get("parsed"):
                    status = "partial"
                    detail = "netsh output could not be parsed (non-English Windows?)"
            elif name == "gpu":
                if data.get("registryError"):
                    detail = f"registry memory size unavailable: {clean_text(data.get('registryError'), 120)}"
            elif name == "memory":
                if data.get("moduleError") or data.get("arrayError"):
                    status = "partial"
                    detail = clean_text(data.get("moduleError") or data.get("arrayError"), 160)
            elif "items" in data:
                detail = f"{len(as_dict_list(data.get('items')))} item(s)"
        else:
            detail = records.error(name)
        coverage.append({"Check": label, "Key": name, "Status": status, "StatusLabel": status_label(status), "Detail": single_line(detail, 600)})
    for entry in run_meta.get("TempCleanup", []):
        if not entry.get("Deleted"):
            coverage.append({
                "Check": "Temporary file clean-up",
                "Key": "temp_cleanup",
                "Status": "failed",
                "StatusLabel": "Failed",
                "Detail": f"{entry.get('Path')} could not be deleted: {entry.get('Error')}",
            })
    if run_meta.get("Cancelled"):
        coverage.append({"Check": "Audit run", "Key": "run", "Status": "partial", "StatusLabel": "Cancelled", "Detail": "The audit was cancelled before it finished."})
    elif records.fatal is not None:
        coverage.append({"Check": "Collector", "Key": "collector", "Status": "failed", "StatusLabel": "Failed", "Detail": clean_text(records.fatal.get("error"), 300) + " " + clean_text(records.fatal.get("position"), 200)})
    elif "done" not in records.sections:
        detail = "The collector ended before reporting completion."
        if run_meta.get("Stderr"):
            detail += " Stderr: " + clean_text(run_meta.get("Stderr"), 300)
        if run_meta.get("ExitCode") not in (None, 0):
            detail += f" Exit code {run_meta.get('ExitCode')}."
        coverage.append({"Check": "Collector", "Key": "collector", "Status": "failed", "StatusLabel": "Failed", "Detail": detail})
    return coverage


def sample_validity(samples: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = {"cpu": 0, "cores": 0, "frequency": 0, "memory": 0, "disks": 0, "disksRaw": 0, "processes": 0, "power": 0}
    for sample in samples:
        cpu = sample.get("cpu") if isinstance(sample.get("cpu"), dict) else None
        if cpu and to_number(cpu.get("total")) is not None:
            counts["cpu"] += 1
        if cpu and as_dict_list(cpu.get("cores")):
            counts["cores"] += 1
        frequency = sample.get("frequency") if isinstance(sample.get("frequency"), dict) else None
        if frequency and (to_number(frequency.get("percentPerformance")) or 0) > 0:
            counts["frequency"] += 1
        memory = sample.get("memory") if isinstance(sample.get("memory"), dict) else None
        if memory and to_number(memory.get("availableMB")) is not None:
            counts["memory"] += 1
        if as_dict_list(sample.get("disks")):
            counts["disks"] += 1
        if as_dict_list(sample.get("disksRaw")):
            counts["disksRaw"] += 1
        if isinstance(sample.get("processes"), list):
            counts["processes"] += 1
        power = sample.get("power") if isinstance(sample.get("power"), dict) else None
        if power:
            counts["power"] += 1
    return counts


def power_source_from_status(status: Any, battery_present: Any) -> Optional[bool]:
    """True on mains, False on battery, None unknown."""
    if battery_present is False:
        return True
    value = to_number(status)
    if value is None:
        return None
    code = int(value)
    if code in MAINS_STATUSES:
        return True
    if code in BATTERY_STATUSES:
        return False
    return None


def analyse_performance(records: RecordSet, options: AuditOptions, logical_processors: int, total_memory_mb: float) -> Dict[str, Any]:
    """Aggregate timestamped samples into performance metrics."""
    samples = records.samples
    done = records.data("samples_done") or {}
    counts = sample_validity(samples)
    logical = max(1, int(logical_processors or to_number(done.get("logicalProcessors")) or 1))

    cpu_values: List[float] = []
    core_series: Dict[str, List[float]] = defaultdict(list)
    performance_values: List[float] = []
    frequency_values: List[float] = []
    memory_percent: List[float] = []
    available_mb: List[float] = []
    commit_percent: List[float] = []
    pages_in: List[float] = []
    pages_per_sec: List[float] = []
    disk_series: Dict[str, Dict[str, List[float]]] = defaultdict(lambda: defaultdict(list))
    latency_series: Dict[str, Dict[str, List[float]]] = defaultdict(lambda: defaultdict(list))
    previous_raw: Dict[str, Dict[str, Any]] = {}
    process_cpu_totals: Dict[str, float] = defaultdict(float)
    process_io_totals: Dict[str, float] = defaultdict(float)
    process_samples: Dict[str, int] = defaultdict(int)
    process_peak_ram: Dict[str, float] = defaultdict(float)
    process_peak_instances: Dict[str, int] = defaultdict(int)
    process_sample_total = 0
    paired: List[Tuple[float, float, Optional[bool]]] = []
    mains_samples = 0
    battery_samples = 0
    unknown_power = 0
    timeline: List[Dict[str, Any]] = []

    for sample in samples:
        cpu = sample.get("cpu") if isinstance(sample.get("cpu"), dict) else {}
        cpu_total = to_number(cpu.get("total"))
        if cpu_total is not None:
            cpu_values.append(cpu_total)
        core_max = None
        for core in as_dict_list(cpu.get("cores")):
            value = to_number(core.get("percent"))
            if value is not None:
                core_series[clean_text(core.get("name"), 16) or "?"].append(value)
                core_max = value if core_max is None else max(core_max, value)
        frequency = sample.get("frequency") if isinstance(sample.get("frequency"), dict) else {}
        performance = to_number(frequency.get("percentPerformance"))
        if performance is not None and performance > 0:
            performance_values.append(performance)
        else:
            performance = None
        mhz = to_number(frequency.get("frequencyMHz"))
        if mhz:
            frequency_values.append(mhz)
        memory = sample.get("memory") if isinstance(sample.get("memory"), dict) else {}
        available = to_number(memory.get("availableMB"))
        if available is not None:
            available_mb.append(available)
            if total_memory_mb > 0:
                memory_percent.append(max(0.0, 100.0 - (available / total_memory_mb) * 100.0))
        commit = to_number(memory.get("percentCommitted"))
        if commit is not None:
            commit_percent.append(commit)
        pin = to_number(memory.get("pagesInputPerSec"))
        if pin is not None:
            pages_in.append(pin)
        pps = to_number(memory.get("pagesPerSec"))
        if pps is not None:
            pages_per_sec.append(pps)
        disk_total_active = None
        disk_total_queue = None
        for disk in as_dict_list(sample.get("disks")):
            name = clean_text(disk.get("name"), 40) or "?"
            series = disk_series[name]
            for key, source in (("active", "percentDiskTime"), ("queue", "currentQueue"), ("avgQueue", "avgQueue"), ("read", "readBytesPerSec"), ("write", "writeBytesPerSec"), ("transfers", "transfersPerSec")):
                value = to_number(disk.get(source))
                if value is not None:
                    series[key].append(value)
            if name == "_Total":
                disk_total_active = to_number(disk.get("percentDiskTime"))
                disk_total_queue = to_number(disk.get("currentQueue"))
        for raw in as_dict_list(sample.get("disksRaw")):
            name = clean_text(raw.get("name"), 40) or "?"
            prev = previous_raw.get(name)
            if prev:
                freq = to_number(raw.get("frequencyPerfTime")) or 0
                for key, counter, base in (("transfer", "secPerTransfer", "secPerTransferBase"), ("read", "secPerRead", "secPerReadBase"), ("write", "secPerWrite", "secPerWriteBase")):
                    c1, c0 = to_number(raw.get(counter)), to_number(prev.get(counter))
                    b1, b0 = to_number(raw.get(base)), to_number(prev.get(base))
                    if None in (c1, c0, b1, b0) or freq <= 0:
                        continue
                    delta_base = b1 - b0
                    delta_counter = c1 - c0
                    if delta_base > 0 and delta_counter >= 0:
                        latency_series[name][key].append((delta_counter / freq) / delta_base * 1000.0)
            previous_raw[name] = raw
        processes = sample.get("processes")
        if isinstance(processes, list):
            process_sample_total += 1
            sample_cpu: Dict[str, float] = defaultdict(float)
            sample_ram: Dict[str, float] = defaultdict(float)
            sample_io: Dict[str, float] = defaultdict(float)
            sample_instances: Dict[str, int] = defaultdict(int)
            for process in as_dict_list(processes):
                name = process_base_name(process.get("name"))
                sample_cpu[name] += (to_number(process.get("cpu")) or 0.0) / logical
                sample_ram[name] += to_number(process.get("privateMB")) or 0.0
                sample_io[name] += ((to_number(process.get("ioReadBytesPerSec")) or 0.0) + (to_number(process.get("ioWriteBytesPerSec")) or 0.0)) / (1024 ** 2)
                sample_instances[name] += 1
            for name in sample_cpu:
                process_cpu_totals[name] += sample_cpu[name]
                process_io_totals[name] += sample_io[name]
                process_samples[name] += 1
                process_peak_ram[name] = max(process_peak_ram[name], sample_ram[name])
                process_peak_instances[name] = max(process_peak_instances[name], sample_instances[name])
        power = sample.get("power") if isinstance(sample.get("power"), dict) else {}
        on_mains = power_source_from_status(power.get("batteryStatus"), power.get("batteryPresent")) if power else None
        if on_mains is True:
            mains_samples += 1
        elif on_mains is False:
            battery_samples += 1
        else:
            unknown_power += 1
        if cpu_total is not None and performance is not None:
            paired.append((cpu_total, performance, on_mains))
        timeline.append({
            "Time": clean_text(sample.get("time"), 40) or None,
            "ElapsedMs": to_number(sample.get("elapsedMs")),
            "CPU": cpu_total,
            "CoreMax": core_max,
            "Performance": performance,
            "AvailableMB": available,
            "PagesInputPerSec": pin,
            "DiskActive": disk_total_active,
            "DiskQueue": disk_total_queue,
            "OnMains": on_mains,
        })

    threshold_load = THRESHOLDS["throttle_load"]
    threshold_perf = THRESHOLDS["throttle_performance"]
    loaded = [entry for entry in paired if entry[0] >= threshold_load]
    overlap = [entry for entry in loaded if entry[1] <= threshold_perf]
    overlap_on_battery = sum(1 for entry in overlap if entry[2] is False)
    minimum_overlap = max(THRESHOLDS["throttle_min_samples"], int(math.ceil(THRESHOLDS["throttle_min_fraction"] * len(paired)))) if paired else 0
    throttle = {
        "PairedSamples": len(paired),
        "LoadedSamples": len(loaded),
        "OverlapSamples": len(overlap),
        "OverlapOnBattery": overlap_on_battery,
        "MinimumOverlap": minimum_overlap,
        "Suspected": bool(paired) and len(overlap) >= minimum_overlap,
        "Evaluable": len(paired) >= max(3, len(samples) // 2) if samples else False,
        "Rule": f"load >= {threshold_load:.0f}% and performance <= {threshold_perf:.0f}% in the same timestamped sample, in at least {minimum_overlap} samples",
    }

    per_core = []
    for name, values in sorted(core_series.items(), key=lambda item: (len(item[0]), item[0])):
        per_core.append({"Core": name, "Average": rounded(mean(values)), "P95": rounded(percentile(values, 95)), "Max": rounded(max(values)), "Samples": len(values)})
    core_p95_max = max((entry["P95"] for entry in per_core if entry["P95"] is not None), default=None)
    hot_core = next((entry["Core"] for entry in per_core if entry["P95"] == core_p95_max), None) if core_p95_max is not None else None
    cpu_average = mean(cpu_values)
    single_thread_bound = bool(core_p95_max is not None and cpu_average is not None and core_p95_max >= THRESHOLDS["core_p95_single_thread"] and cpu_average < 50.0)

    per_disk = []
    for name, series in disk_series.items():
        latency = latency_series.get(name, {})
        per_disk.append({
            "Disk": name,
            "Active_Average": rounded(mean(series.get("active", []))),
            "Active_P95": rounded(percentile(series.get("active", []), 95)),
            "Queue_P95": rounded(percentile(series.get("queue", []), 95)),
            "AvgQueue_P95": rounded(percentile(series.get("avgQueue", []), 95)),
            "ReadMBps_Average": rounded((mean(series.get("read", [])) or 0.0) / (1024 ** 2), 2) if series.get("read") else None,
            "WriteMBps_Average": rounded((mean(series.get("write", [])) or 0.0) / (1024 ** 2), 2) if series.get("write") else None,
            "TransfersPerSec_Average": rounded(mean(series.get("transfers", [])), 0),
            "Latency_ms_Average": rounded(mean(latency.get("transfer", [])), 2),
            "Latency_ms_P95": rounded(percentile(latency.get("transfer", []), 95), 2),
            "ReadLatency_ms_P95": rounded(percentile(latency.get("read", []), 95), 2),
            "WriteLatency_ms_P95": rounded(percentile(latency.get("write", []), 95), 2),
            "LatencyIntervals": len(latency.get("transfer", [])),
            "Samples": len(series.get("active", [])),
        })
    per_disk.sort(key=lambda entry: (entry["Disk"] != "_Total", entry["Disk"]))
    total_disk = next((entry for entry in per_disk if entry["Disk"] == "_Total"), None)

    top_cpu = []
    for name in process_cpu_totals:
        top_cpu.append({
            "Process": name,
            "Instances": process_peak_instances[name],
            "AverageCPU": rounded(process_cpu_totals[name] / max(1, process_sample_total)),
            "SamplesPresent": f"{process_samples[name]}/{process_sample_total}",
            "PeakRAM_MB": rounded(process_peak_ram[name], 0),
            "IO_MBps_Average": rounded(process_io_totals[name] / max(1, process_sample_total), 2),
        })
    top_cpu.sort(key=lambda row: (row["AverageCPU"] or 0, row["PeakRAM_MB"] or 0), reverse=True)
    top_io = sorted(top_cpu, key=lambda row: row["IO_MBps_Average"] or 0, reverse=True)[:10]
    top_io = [row for row in top_io if (row["IO_MBps_Average"] or 0) >= 0.5]

    top_ram = []
    for item in records.items("top_ram"):
        top_ram.append({
            "Process": process_base_name(item.get("process")),
            "Instances": int(to_number(item.get("instances")) or 0),
            "RAM_MB": rounded(to_number(item.get("ramMB")), 0),
            "Private_MB": rounded(to_number(item.get("privateMB")), 0),
        })
    top_ram.sort(key=lambda row: row["RAM_MB"] or 0, reverse=True)

    if not samples:
        power_source = "Unknown (no samples)"
    elif mains_samples and not battery_samples:
        power_source = "Mains"
    elif battery_samples and not mains_samples:
        power_source = "Battery"
    elif mains_samples and battery_samples:
        power_source = "Mixed (mains and battery during the sample)"
    else:
        power_source = "Unknown"

    return {
        "SamplesRequested": int(to_number(done.get("requested")) or options.sample_count),
        "SamplesTaken": len(samples),
        "ElapsedSeconds": rounded(to_number(done.get("elapsedSeconds")), 0),
        "AverageSampleSeconds": rounded(mean([to_number(s.get("durationMs")) or 0 for s in samples]) / 1000.0, 2) if samples else None,
        "ValidSamples": counts,
        "LogicalProcessorsUsed": logical,
        "CPU_AveragePercent": rounded(cpu_average),
        "CPU_P95Percent": rounded(percentile(cpu_values, 95)),
        "CPU_MaxPercent": rounded(max(cpu_values)) if cpu_values else None,
        "Core_P95Max": rounded(core_p95_max),
        "Core_Hottest": hot_core,
        "SingleThreadBound": single_thread_bound,
        "PerCore": per_core,
        "CPU_Performance_AveragePercent": rounded(mean(performance_values), 0),
        "CPU_Performance_P05Percent": rounded(percentile(performance_values, 5), 0),
        "CPU_Performance_MinPercent": rounded(min(performance_values), 0) if performance_values else None,
        "CPU_Performance_ValidSamples": len(performance_values),
        "CPU_FrequencyMHz_Average": rounded(mean(frequency_values), 0),
        "Throttling": throttle,
        "Memory_AveragePercent": rounded(mean(memory_percent)),
        "Memory_P95Percent": rounded(percentile(memory_percent, 95)),
        "AvailableRAM_MinMB": rounded(min(available_mb), 0) if available_mb else None,
        "AvailableRAM_AverageMB": rounded(mean(available_mb), 0),
        "Commit_AveragePercent": rounded(mean(commit_percent)),
        "Commit_P95Percent": rounded(percentile(commit_percent, 95)),
        "PagesInputPerSec_Average": rounded(mean(pages_in), 0),
        "PagesInputPerSec_P95": rounded(percentile(pages_in, 95), 0),
        "PagesPerSec_Average": rounded(mean(pages_per_sec), 0),
        "DiskActive_Average": total_disk["Active_Average"] if total_disk else None,
        "DiskActive_P95": total_disk["Active_P95"] if total_disk else None,
        "DiskQueue_P95": total_disk["Queue_P95"] if total_disk else None,
        "DiskLatency_ms_P95": total_disk["Latency_ms_P95"] if total_disk else None,
        "DiskReadMBps_Average": total_disk["ReadMBps_Average"] if total_disk else None,
        "DiskWriteMBps_Average": total_disk["WriteMBps_Average"] if total_disk else None,
        "PerDisk": per_disk,
        "PowerSource": power_source,
        "OnMainsSamples": mains_samples,
        "OnBatterySamples": battery_samples,
        "UnknownPowerSamples": unknown_power,
        "TopCPUProcesses": top_cpu[:20],
        "TopIOProcesses": top_io,
        "TopRAMProcesses": top_ram[:20],
        "Samples": timeline,
    }


class FindingList:
    def __init__(self) -> None:
        self.items: List[Dict[str, Any]] = []

    def add(self, severity: str, category: str, finding: str, evidence: str, recommendation: str, penalty: int = 0) -> None:
        self.items.append({
            "Severity": severity,
            "Category": category,
            "Finding": finding,
            "Evidence": evidence,
            "Recommendation": recommendation,
            "Penalty": int(penalty),
        })

    def count(self, *severities: str) -> int:
        return sum(1 for item in self.items if item["Severity"] in severities)

    def sorted(self) -> List[Dict[str, Any]]:
        return sorted(self.items, key=lambda item: (SEVERITY_ORDER.get(item["Severity"], 9), item["Category"]))


def event_diagnostics(events: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Diagnose hardware providers over every collected group, before any display cap."""
    result = {
        "WHEA_Errors": 0, "WHEA_Warnings": 0, "Storage_Errors": 0, "Storage_Warnings": 0, "KernelPower41": 0,
        "ProcessorPower37": 0, "DisplayResets": 0, "PnP219": 0, "Details": [],
    }
    if not events:
        return result
    hardware = events.get("hardware") if isinstance(events.get("hardware"), dict) else {}
    storage_pattern = re.compile(r"disk|stornvme|storahci|storport|iaStor|Ntfs|volmgr|partmgr|StorDiag", re.IGNORECASE)
    display_pattern = re.compile(r"Display|nvlddmkm|amdkmdag|igfx", re.IGNORECASE)
    seen = set()
    for group in as_dict_list(hardware.get("groups")):
        provider = clean_text(group.get("provider"))
        event_id = int(to_number(group.get("eventId")) or 0)
        level = int(to_number(group.get("level")) or 0)
        count = int(to_number(group.get("count")) or 0)
        key = (provider, event_id, level)
        if key in seen or count <= 0:
            continue
        seen.add(key)
        is_error = level in (1, 2)
        if "WHEA" in provider:
            result["WHEA_Errors" if is_error else "WHEA_Warnings"] += count
        elif storage_pattern.search(provider):
            result["Storage_Errors" if is_error else "Storage_Warnings"] += count
        if "Kernel-Power" in provider and event_id == 41:
            result["KernelPower41"] += count
        if "Kernel-Processor-Power" in provider and event_id == 37:
            result["ProcessorPower37"] += count
        if display_pattern.search(provider) and (event_id in (4101, 13, 14) or is_error):
            result["DisplayResets"] += count
        if "Kernel-PnP" in provider and event_id == 219:
            result["PnP219"] += count
        result["Details"].append({"Provider": provider, "EventId": event_id, "Level": clean_text(group.get("levelName")) or level, "Count": count, "MostRecent": clean_text(group.get("mostRecent"), 40), "Message": clean_text(group.get("message"), 200)})
    result["Details"].sort(key=lambda row: row["Count"], reverse=True)
    return result


def classify_coverage(coverage: List[Dict[str, Any]], samples_status: str) -> Tuple[List[str], List[str]]:
    """Split coverage gaps into core (diagnostic evidence missing) and elevation-only."""
    core_missing: List[str] = []
    elevation_missing: List[str] = []
    for row in coverage:
        key, status = row["Key"], row["Status"]
        if status == "requires_elevation" or (status == "partial" and key == "physical_disks" and "elevation" in row["Detail"].lower()):
            elevation_missing.append(row["Check"])
        elif status in ("failed", "unavailable", "partial") and (key in DIAGNOSTIC_SECTIONS or key in ELEVATION_SECTIONS or key in ("collector", "run", "temp_cleanup")):
            core_missing.append(row["Check"])
    if samples_status not in ("collected", "partial") and "Live performance samples" not in core_missing:
        core_missing.append("Live performance samples")
    return core_missing, elevation_missing


def evaluate_findings(report: Dict[str, Any], records: RecordSet, coverage: List[Dict[str, Any]]) -> FindingList:
    findings = FindingList()
    t = THRESHOLDS
    performance = report["Performance"]
    battery = report["Battery"]
    system = report["System"]
    diagnostics = report["EventDiagnostics"]
    days = report["EventDays"]

    for disk in report["LogicalDisks"]:
        free = to_number(disk.get("FreePercent"))
        if free is None:
            continue
        if free < t["disk_free_critical"]:
            findings.add("Critical", "Storage", f"Drive {disk['Drive']} is nearly full.", f"{disk['Free']} free ({free:.1f}%).", "Ask IT to clean approved temporary or cached data, or increase capacity.", 18)
        elif free < t["disk_free_warning"]:
            findings.add("Warning", "Storage", f"Drive {disk['Drive']} has limited free space.", f"{disk['Free']} free ({free:.1f}%).", "Keep at least 20% free where practical; ask IT to review safe clean-up options.", 8)

    for disk in report["PhysicalDisks"]:
        health = clean_text(disk.get("Health"))
        if disk.get("HealthKnown") and health.lower() not in ("healthy", "ok"):
            findings.add("Critical", "Storage", "A physical disk does not report healthy status.", f"{disk['Name']}: {health}, {disk.get('OperationalStatus') or 'status unknown'}.", "Back up company data through approved channels and contact IT promptly.", 25)
        temp = to_number(disk.get("TemperatureC"))
        if temp is not None and temp >= t["storage_temp"]:
            findings.add("Warning", "Thermal", "Storage temperature is high.", f"{disk['Name']}: {temp:.0f} C.", "Check ventilation and ask IT to inspect cooling.", 8)
        wear = to_number(disk.get("WearPercent"))
        if wear is not None and wear >= t["ssd_wear"]:
            findings.add("Warning", "Storage", "SSD wear indicator is high.", f"{disk['Name']}: {wear:.0f}% wear.", "Ask IT to verify the vendor SMART data and the replacement plan.", 12)
        uncorrected = (to_number(disk.get("ReadErrorsUncorrected")) or 0) + (to_number(disk.get("WriteErrorsUncorrected")) or 0)
        if uncorrected > 0:
            findings.add("Warning", "Storage", "The disk reports uncorrected read or write errors.", f"{disk['Name']}: {uncorrected:.0f} uncorrected error(s).", "Back up data through approved channels and ask IT to review SMART data.", 12)

    health_percent = to_number(battery.get("HealthPercent"))
    if health_percent is not None:
        if health_percent < t["battery_critical"]:
            findings.add("Critical", "Battery", "Battery capacity is severely degraded.", f"Full-charge capacity is {health_percent:.0f}% of design ({battery.get('CapacitySource')}).", "Request an IT battery assessment or replacement; report swelling immediately.", 18)
        elif health_percent < t["battery_warning"]:
            findings.add("Warning", "Battery", "Battery wear is material.", f"Full-charge capacity is {health_percent:.0f}% of design ({battery.get('CapacitySource')}).", "Ask IT to assess battery replacement, especially if runtime is poor.", 8)

    taken = performance["SamplesTaken"]
    valid = performance["ValidSamples"]
    if taken == 0:
        findings.add("Information", "Coverage", "No live performance samples were collected.", "The performance section is empty, so CPU, RAM, disk and process evidence is missing.", "Rerun the audit; if it fails again, run it from a console to read the collector error.", 0)
    else:
        cpu_avg = to_number(performance.get("CPU_AveragePercent"))
        if cpu_avg is not None and cpu_avg >= t["cpu_avg_warning"]:
            findings.add("Warning", "Performance", "CPU load remained high during the sample.", f"Average {cpu_avg:.0f}%, 95th percentile {fmt_value(performance.get('CPU_P95Percent'), '%', 0)} over {valid['cpu']} samples; workload: {report['Workload']['Label']}.", "Repeat while the laptop feels slow and give IT the top-process evidence.", 10)
        if performance.get("SingleThreadBound"):
            findings.add("Information", "Performance", "One logical processor ran near saturation while overall CPU load was moderate.", f"Core {performance.get('Core_Hottest')} P95 {fmt_value(performance.get('Core_P95Max'), '%', 0)} vs overall average {fmt_value(cpu_avg, '%', 0)}; workload: {report['Workload']['Label']}.", "Typical of single-threaded Excel calculation or Civil 3D operations; higher single-core speed or workload changes help more than extra cores.", 0)
        ram_p95 = to_number(performance.get("Memory_P95Percent"))
        ram_avg = to_number(performance.get("Memory_AveragePercent"))
        min_available = to_number(performance.get("AvailableRAM_MinMB"))
        if ram_p95 is not None:
            if ram_p95 >= t["ram_p95_critical"]:
                findings.add("Critical", "Performance", "Physical memory was nearly exhausted during the sample.", f"Physical RAM in use: average {fmt_value(ram_avg, '%')}, 95th percentile {ram_p95:.1f}%; minimum available {fmt_value(min_available, ' MB', 0)}; workload: {report['Workload']['Label']}.", "Close only approved unused applications; ask IT to review RAM capacity against the top applications by RAM.", 16)
            elif ram_p95 >= t["ram_p95_warning"]:
                findings.add("Warning", "Performance", "Physical memory use was high during the sample.", f"Physical RAM in use: average {fmt_value(ram_avg, '%')}, 95th percentile {ram_p95:.1f}%; minimum available {fmt_value(min_available, ' MB', 0)}.", "Repeat with the normal workload and ask IT to review the top applications by RAM.", 8)
        pages_in = to_number(performance.get("PagesInputPerSec_Average"))
        total_mb = to_number(report["MemorySummary"].get("InstalledBytes"))
        total_mb = total_mb / (1024 ** 2) if total_mb else None
        low_ram = (ram_p95 is not None and ram_p95 >= t["ram_p95_warning"]) or (min_available is not None and total_mb and min_available < 0.10 * total_mb)
        if pages_in is not None and pages_in >= t["paging_avg_pages_in"]:
            if low_ram:
                findings.add("Warning", "Performance", "Windows read pages from disk continuously while physical RAM was nearly full.", f"Pages input average {pages_in:.0f}/s, 95th percentile {fmt_value(performance.get('PagesInputPerSec_P95'), '/s', 0)}; RAM 95th percentile {fmt_value(ram_p95, '%')}; minimum available {fmt_value(min_available, ' MB', 0)}.", "Consistent with the workload exceeding installed memory, but not proof: repeat during normal work on several days and give IT every snapshot.", 10)
            else:
                findings.add("Information", "Performance", "Pages were read from disk at a high rate without RAM pressure.", f"Pages input average {pages_in:.0f}/s while RAM 95th percentile was {fmt_value(ram_p95, '%')}.", "Usually file loading or mapped-file reads (application start, sync, scans) rather than memory shortage.", 0)
        throttle = performance["Throttling"]
        if throttle.get("Suspected"):
            evidence = (
                f"{throttle['OverlapSamples']} of {throttle['PairedSamples']} timestamped samples had CPU load >= {t['throttle_load']:.0f}% with processor performance <= {t['throttle_performance']:.0f}% of nominal "
                f"(minimum {throttle['MinimumOverlap']}); performance 5th percentile {fmt_value(performance.get('CPU_Performance_P05Percent'), '%', 0)}; power source: {performance.get('PowerSource')}."
            )
            if diagnostics["ProcessorPower37"]:
                evidence += f" Windows also logged {diagnostics['ProcessorPower37']} Kernel-Processor-Power event 37 (firmware limited processor speed) in the last {days} days."
            if throttle["OverlapOnBattery"] >= max(1, throttle["OverlapSamples"] // 2):
                findings.add("Information", "Thermal", "The CPU ran well below nominal frequency under load, but mostly on battery power.", evidence, "Power saving on battery is expected. Repeat the audit on mains with a balanced or high-performance plan before concluding thermal throttling.", 0)
            else:
                findings.add("Warning", "Thermal", "The CPU ran well below nominal frequency while under load, consistent with thermal or power throttling.", evidence, "Confirm the laptop was on mains with a balanced or high-performance plan; if so, ask IT to inspect cooling, fan behaviour and charger wattage.", 10)
        elif diagnostics["ProcessorPower37"]:
            findings.add("Information", "Thermal", "Windows logged firmware-limited processor speed events, but throttling was not observed during this sample.", f"{diagnostics['ProcessorPower37']} Kernel-Processor-Power event 37 occurrence(s) in the last {days} days; overlap samples {throttle['OverlapSamples']} of {throttle['PairedSamples']}.", "Rerun the audit during heavy work on mains; if throttling recurs, ask IT to review cooling and the charger rating.", 0)
        elif not throttle.get("Evaluable") and taken:
            findings.add("Information", "Coverage", "Throttling could not be evaluated.", f"Only {throttle['PairedSamples']} of {taken} samples had both CPU load and a non-zero processor performance counter.", "The processor performance counter is not exposed on every device; the CPU frequency columns are unverified.", 0)
        disk_p95 = to_number(performance.get("DiskActive_P95"))
        queue_p95 = to_number(performance.get("DiskQueue_P95"))
        if disk_p95 is not None and disk_p95 >= t["disk_active_p95"] and queue_p95 is not None and queue_p95 >= t["disk_queue_p95"]:
            findings.add("Warning", "Performance", "The disk subsystem was saturated during the sample.", f"Disk activity 95th percentile {disk_p95:.0f}%; queue 95th percentile {queue_p95:.1f}; read {fmt_value(performance.get('DiskReadMBps_Average'), ' MB/s')}, write {fmt_value(performance.get('DiskWriteMBps_Average'), ' MB/s')}.", "Use the top-process I/O list and event evidence to identify sync, update, security scan or storage issues with IT.", 10)
        for disk in performance["PerDisk"]:
            if disk["Disk"] == "_Total":
                continue
            latency = to_number(disk.get("Latency_ms_P95"))
            active = to_number(disk.get("Active_Average"))
            if latency is not None and latency >= t["disk_latency_ms_p95"] and active is not None and active >= t["disk_latency_min_active"] and (disk.get("LatencyIntervals") or 0) >= 3:
                findings.add("Warning", "Performance", f"Disk {disk['Disk']} responded slowly under load.", f"Latency 95th percentile {latency:.0f} ms per transfer over {disk['LatencyIntervals']} intervals; activity average {active:.0f}%.", "Compare with the physical disk media type: tens of milliseconds is normal for a hard disk but not for an SSD. Ask IT to check drivers, SMART data and encryption software.", 8)

    uptime = to_number(system.get("UptimeDays"))
    if uptime is not None and uptime >= t["uptime_days"]:
        findings.add("Information", "Maintenance", "The laptop has not completed a full restart recently.", f"Uptime is {uptime:.1f} days.", "Save work and restart at an approved time; use Restart, not Shut down, to clear uptime.", 0)
    if system.get("PendingReboot"):
        findings.add("Information", "Maintenance", "Windows reports a pending restart.", "A servicing or file-replacement reboot marker is present.", "Save work and restart at an approved time.", 0)
    if report["ProblemDevices"]:
        findings.add("Warning", "Drivers/devices", "Device Manager reports one or more device errors.", f"{len(report['ProblemDevices'])} device(s) have a non-zero error code.", "Ask IT to review the listed device error codes and approved drivers.", 10)
    for zone in report["ThermalZones"]:
        temp = to_number(zone.get("TemperatureC"))
        if temp is None:
            continue
        if temp >= t["thermal_critical"]:
            findings.add("Critical", "Thermal", "A reported thermal zone is critically hot.", f"{zone.get('Zone')}: {temp:.0f} C.", "Stop heavy work, keep vents clear and contact IT if the temperature persists.", 18)
        elif temp >= t["thermal_warning"]:
            findings.add("Warning", "Thermal", "A reported thermal zone is hot.", f"{zone.get('Zone')}: {temp:.0f} C.", "Check ventilation and ask IT to inspect fan and cooling behaviour if persistent.", 8)

    if diagnostics["WHEA_Errors"]:
        findings.add("Critical", "Hardware events", "Windows Hardware Error Architecture errors were recorded.", f"{diagnostics['WHEA_Errors']} error/critical WHEA event(s) in the last {days} days.", "Contact IT; WHEA events can indicate CPU, memory, PCIe, power or firmware faults.", 20)
    elif diagnostics["WHEA_Warnings"]:
        findings.add("Warning", "Hardware events", "Corrected hardware (WHEA) warnings were recorded.", f"{diagnostics['WHEA_Warnings']} WHEA warning event(s) in the last {days} days.", "Corrected errors are tolerated by Windows but recurring ones deserve an IT hardware check.", 8)
    if diagnostics["Storage_Errors"]:
        findings.add("Warning", "Storage events", "Storage-related error events were recorded.", f"{diagnostics['Storage_Errors']} error/critical storage event(s) in the last {days} days.", "Ask IT to correlate the event IDs with SMART health, drivers and backups.", 12)
    elif diagnostics["Storage_Warnings"]:
        findings.add("Warning", "Storage events", "Storage-related warnings were recorded (for example retried I/O).", f"{diagnostics['Storage_Warnings']} storage warning event(s) in the last {days} days.", "Retried or reset I/O warnings often precede disk faults; ask IT to review SMART data and the disk driver.", 6)
    if diagnostics["KernelPower41"]:
        findings.add("Warning", "Stability", "Unexpected shutdown or restart events were recorded.", f"{diagnostics['KernelPower41']} Kernel-Power event 41 occurrence(s) in the last {days} days.", "Report unexpected shutdowns to IT, including dock, charger use and workload at the time.", 8)
    if diagnostics["DisplayResets"]:
        findings.add("Warning", "Drivers/devices", "Graphics driver errors or resets were recorded.", f"{diagnostics['DisplayResets']} display-driver event(s) in the last {days} days.", "Relevant to Civil 3D and multi-monitor docks; ask IT to review the approved graphics driver version.", 6)
    if diagnostics["PnP219"]:
        findings.add("Information", "Drivers/devices", "A driver failed to load for a device at least once.", f"{diagnostics['PnP219']} Kernel-PnP event 219 occurrence(s) in the last {days} days.", "Usually a dock or USB device; mention it to IT if the device misbehaves.", 0)
    if report["ReliabilityTotal"] >= t["reliability_records"]:
        findings.add("Warning", "Stability", "Frequent application or Windows reliability failures were recorded.", f"{report['ReliabilityTotal']} relevant records in the last {days} days (report shows up to 40).", "Ask IT to review the repeated product or source names and timestamps.", 7)

    disabled = [profile["Name"] for profile in report["Security"]["FirewallProfiles"] if clean_text(profile.get("Enabled")).lower() == "false"]
    if disabled:
        findings.add("Warning", "Security", "One or more Windows Firewall profiles report disabled.", ", ".join(disabled), "Do not change company policy; ask IT to confirm whether another managed control intentionally replaces it.", 10)
    for product in report["Security"]["Antivirus"]:
        state = clean_text(product.get("State"))
        if state.startswith("Disabled") or "out of date" in state:
            findings.add("Information", "Security", f"Antivirus product state needs confirmation: {product.get('Product')}.", state, "Managed devices may rely on another control; ask IT rather than changing anything.", 0)
    defender = report["Security"]["Defender"]
    signature_age = to_number(defender.get("SignatureAgeDays")) if isinstance(defender, dict) else None
    if signature_age is not None and signature_age > 7:
        findings.add("Information", "Security", "Microsoft Defender signatures are more than a week old.", f"Signature age {signature_age:.0f} days.", "Ask IT whether another product is the primary antivirus; do not change settings.", 0)

    deep = report["DeepChecks"]
    for label, entry, corrupt_severity, corrupt_penalty in (("DISM /CheckHealth", deep["DISMCheckHealth"], "Critical", 15), ("SFC /verifyonly", deep["SFCVerifyOnly"], "Warning", 10)):
        verdict = entry["Verdict"]
        if verdict.startswith(("Corruption detected", "Integrity violations")):
            findings.add(corrupt_severity, "Windows integrity", f"{label} reported a problem.", f"{verdict}; exit code {entry.get('ExitCode')}.", "Do not repair on a managed device yourself; give IT the output so they can run the approved repair.", corrupt_penalty)
        elif verdict.startswith(("Command", "Blocked", "Inconclusive")):
            findings.add("Information", "Windows integrity", f"{label} did not produce a usable result.", f"{verdict}; exit code {entry.get('ExitCode')}.", "Ask IT to run the check; the tool cannot conclude either way.", 0)

    elevation = report["Completeness"]["ElevationMissing"]
    if elevation:
        findings.add("Information", "Coverage", "Some checks need an elevated PowerShell and were not performed.", "; ".join(elevation), "Use 'Restart as administrator' in the window if company policy permits, or ask IT to confirm these controls only if they matter to the ticket.", 0)
    labels = {row["Check"]: row["StatusLabel"].lower() for row in coverage}
    missing = [f"{check} ({labels.get(check, 'missing')})" for check in report["Completeness"]["CoreMissing"]]
    other = [f"{row['Check']} ({row['StatusLabel'].lower()})" for row in coverage if row["Status"] in ("failed", "unavailable") and row["Check"] not in report["Completeness"]["CoreMissing"] and row["Check"] not in elevation]
    if missing:
        findings.add("Information", "Coverage", "Some diagnostic checks were unavailable, incomplete or failed on this device.", "; ".join(missing)[:600], "Treat the related sections as absent evidence rather than as a clean result; see the coverage table.", 0)
    if other:
        findings.add("Information", "Coverage", "Some optional checks were unavailable on this device.", "; ".join(other)[:400], "No action; listed for completeness.", 0)

    return findings


def score_and_rating(findings: FindingList, core_missing: Sequence[str], elevation_missing: Sequence[str]) -> Tuple[int, str, str]:
    """Score from penalties; rating bounded by the worst finding and by missing evidence.

    core_missing: diagnostic checks that failed or were unavailable (not elevation).
    elevation_missing: checks skipped only because PowerShell was not elevated.
    """
    penalty = sum(item["Penalty"] for item in findings.items)
    score = max(0, 100 - int(penalty))
    critical = findings.count("Critical")
    warning = findings.count("Warning")
    if score >= 90:
        rating = "Healthy snapshot"
    elif score >= 75:
        rating = "Review advised"
    elif score >= 50:
        rating = "Attention required"
    else:
        rating = "Escalate to IT"
    if critical:
        rating = "Escalate to IT" if score < 50 else "Attention required"
    elif warning and rating == "Healthy snapshot":
        rating = "Review advised"
    if core_missing and not critical and not warning:
        rating = "Inconclusive (evidence incomplete)"
    headline = f"{score}/100 - {rating}"
    if critical:
        headline += f" ({critical} critical finding{'s' if critical != 1 else ''})"
    if core_missing:
        headline += f" - {len(core_missing)} check(s) missing"
    if elevation_missing:
        headline += f" - {len(elevation_missing)} check(s) need elevation"
    return score, rating, headline


def analyse_records(records_iterable: Iterable[Dict[str, Any]], options: AuditOptions, run_meta: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Build the schema 1.2 report from collector records. Pure; no I/O."""
    run_meta = run_meta or {}
    records = RecordSet(records_iterable)
    context = records.data("context") or {}
    system_data = records.data("system") or {}
    cpu_data = records.data("cpu") or {}
    memory_data = records.data("memory") or {}
    gpu_data = records.data("gpu") or {}
    battery_data = records.data("battery") or {}
    events_data = records.data("events")
    audit_time = now_local()
    elevated = bool(context.get("isAdmin") if context else (records.data("hello") or {}).get("isAdmin", False))

    last_boot, _ = parse_timestamp(system_data.get("lastBoot"))
    uptime_days = rounded((audit_time - last_boot).total_seconds() / 86400.0) if last_boot else None
    total_ram_bytes = to_number(system_data.get("totalPhysicalMemory")) or ((to_number(system_data.get("totalVisibleMemoryKB")) or 0) * 1024)
    logical_processors = int(to_number(system_data.get("logicalProcessors")) or to_number(cpu_data.get("logicalProcessors")) or 1)
    pc_type = {1: "Desktop", 2: "Mobile (laptop)", 3: "Workstation", 4: "Enterprise server", 5: "SOHO server", 6: "Appliance", 7: "Performance server", 8: "Maximum"}.get(int(to_number(system_data.get("pcSystemType")) or 0), "Unknown")
    build = clean_text(system_data.get("buildNumber"))
    ubr = to_number(system_data.get("ubr"))
    if build and ubr is not None:
        build = f"{build}.{int(ubr)}"

    system = {
        "Manufacturer": clean_text(system_data.get("manufacturer")) or None,
        "Model": clean_text(system_data.get("model")) or None,
        "SystemFamily": clean_text(system_data.get("systemFamily")) or None,
        "SystemType": clean_text(system_data.get("systemType")) or None,
        "ChassisType": pc_type,
        "ComputerName": clean_text(context.get("computerName")) or "[redacted]",
        "CurrentUser": clean_text(context.get("userName")) or "[redacted]",
        "WindowsEdition": clean_text(system_data.get("windowsEdition")) or None,
        "WindowsVersion": clean_text(system_data.get("displayVersion")) or clean_text(system_data.get("windowsVersion")) or None,
        "BuildNumber": build or None,
        "Architecture": clean_text(system_data.get("architecture")) or None,
        "LastBoot": iso(last_boot),
        "UptimeDays": uptime_days,
        "InstallDate": clean_text(system_data.get("installDate")) or None,
        "BIOSVersion": clean_text(system_data.get("biosVersion")) or None,
        "BIOSDate": clean_text(system_data.get("biosDate")) or None,
        "BIOSSerial": clean_text(system_data.get("biosSerial")) or None,
        "SecureBoot": secure_boot_text(records),
        "Elevated": elevated,
        "PendingReboot": bool(system_data.get("pendingReboot")),
        "AuditMode": options.mode,
        "AuditTime": audit_time.isoformat(timespec="seconds"),
        "AuditTimeUTC": audit_time.astimezone(timezone.utc).isoformat(timespec="seconds"),
        "TimeZone": clean_text(system_data.get("timeZone")) or None,
        "SensitiveData": options.include_sensitive,
        "DeviceFingerprint": clean_text(system_data.get("deviceFingerprint")) or None,
        "PowerShellVersion": clean_text(context.get("psVersion")) or None,
        "PowerShell64Bit": context.get("is64BitProcess"),
        "ToolVersion": TOOL_VERSION,
    }
    if not system["DeviceFingerprint"] and (system["Manufacturer"] or system["Model"]):
        system["DeviceFingerprint"] = "f-" + fingerprint(system["Manufacturer"], system["Model"], cpu_data.get("name"), total_ram_bytes)

    cpu = {
        "Name": clean_text(cpu_data.get("name")) or None,
        "Sockets": int(to_number(cpu_data.get("sockets")) or 0) or None,
        "PhysicalCores": to_number(cpu_data.get("physicalCores")),
        "LogicalProcessors": to_number(cpu_data.get("logicalProcessors")),
        "MaximumClockMHz": to_number(cpu_data.get("maxClockMHz")),
        "CurrentClockMHz": to_number(cpu_data.get("currentClockMHz")),
        "L2CacheKB": to_number(cpu_data.get("l2CacheKB")),
        "L3CacheKB": to_number(cpu_data.get("l3CacheKB")),
        "Socket": clean_text(cpu_data.get("socket")) or None,
        "VirtualisationFW": cpu_data.get("virtualisationFirmware"),
    }

    modules = []
    for module in as_dict_list(memory_data.get("modules")):
        modules.append({
            "Slot": clean_text(module.get("slot")) or None,
            "Bank": clean_text(module.get("bank")) or None,
            "Capacity": fmt_bytes(module.get("capacityBytes")),
            "CapacityBytes": to_number(module.get("capacityBytes")),
            "SpeedMHz": to_number(module.get("speedMHz")),
            "ConfiguredMHz": to_number(module.get("configuredMHz")),
            "Manufacturer": clean_text(module.get("manufacturer")) or None,
            "PartNumber": clean_text(module.get("partNumber")) or None,
        })
    max_kb = to_number(memory_data.get("maxCapacityExKB")) or to_number(memory_data.get("maxCapacityKB"))
    max_bytes = max_kb * 1024 if max_kb else None
    max_note = "Firmware-reported (Win32_PhysicalMemoryArray, kilobytes); provisional, confirm with the vendor specification."
    if max_bytes and total_ram_bytes and max_bytes < total_ram_bytes:
        max_note = "Firmware value is lower than the installed RAM, so it is unreliable."
    page_files = as_dict_list(memory_data.get("pageFiles"))
    memory_summary = {
        "Installed": fmt_bytes(total_ram_bytes) if total_ram_bytes else "N/A",
        "InstalledBytes": total_ram_bytes or None,
        "PopulatedModules": len(modules),
        "TotalSlots": to_number(memory_data.get("slots")),
        "MaximumSupported": fmt_bytes(max_bytes) if max_bytes else "Unknown",
        "MaximumSupportedBytes": max_bytes,
        "MaximumSupportedNote": max_note if max_bytes else "Not reported by firmware.",
        "PageFiles": "; ".join(f"{clean_text(p.get('name'))}: {fmt_value((to_number(p.get('allocatedMB')) or 0) / 1024.0, ' GB')} (peak use {fmt_value(p.get('peakUsageMB'), ' MB', 0)})" for p in page_files) or "None reported",
        "PageFilePeakMB": max((to_number(p.get("peakUsageMB")) or 0 for p in page_files), default=None) if page_files else None,
        "CommitLimitGB": rounded((to_number(system_data.get("totalVirtualMemoryKB")) or 0) / (1024 ** 2), 1) if system_data.get("totalVirtualMemoryKB") else None,
    }

    gpus = []
    registry = as_dict_list(gpu_data.get("registry"))
    for controller in as_dict_list(gpu_data.get("controllers")):
        memory_bytes, source = match_gpu_memory(controller, registry)
        gpus.append({
            "Name": clean_text(controller.get("name")) or None,
            "DedicatedMemory": fmt_bytes(memory_bytes) if memory_bytes else "N/A",
            "MemoryBytes": memory_bytes,
            "MemorySource": source,
            "DriverVersion": clean_text(controller.get("driverVersion")) or None,
            "DriverDate": clean_text(controller.get("driverDate")) or None,
            "CurrentMode": f"{fmt_value(controller.get('horizontal'), '', 0)} x {fmt_value(controller.get('vertical'), '', 0)} @ {fmt_value(controller.get('refreshHz'), ' Hz', 0)}",
            "Status": clean_text(controller.get("status")) or None,
        })

    physical = []
    disks_data = records.data("physical_disks") or {}
    for disk in as_dict_list(disks_data.get("items")):
        health = clean_text(disk.get("health"))
        physical.append({
            "Name": clean_text(disk.get("name")) or "Unknown disk",
            "MediaType": clean_text(disk.get("mediaType")) or None,
            "BusType": clean_text(disk.get("busType")) or None,
            "Capacity": fmt_bytes(disk.get("sizeBytes")),
            "CapacityBytes": to_number(disk.get("sizeBytes")),
            "Health": health or "Unknown",
            "HealthKnown": bool(health),
            "SMARTStatus": clean_text(disk.get("smartStatus")) or None,
            "OperationalStatus": clean_text(disk.get("operationalStatus")) or None,
            "Firmware": clean_text(disk.get("firmware")) or None,
            "Serial": clean_text(disk.get("serial")) or None,
            "TemperatureC": to_number(disk.get("temperatureC")),
            "WearPercent": to_number(disk.get("wearPercent")),
            "ReadErrorsTotal": to_number(disk.get("readErrorsTotal")),
            "ReadErrorsUncorrected": to_number(disk.get("readErrorsUncorrected")),
            "WriteErrorsTotal": to_number(disk.get("writeErrorsTotal")),
            "WriteErrorsUncorrected": to_number(disk.get("writeErrorsUncorrected")),
            "PowerOnHours": to_number(disk.get("powerOnHours")),
            "Source": clean_text(disks_data.get("source")) or None,
        })

    logical = []
    for disk in records.items("logical_disks"):
        size = to_number(disk.get("sizeBytes")) or 0
        free = to_number(disk.get("freeBytes")) or 0
        logical.append({
            "Drive": clean_text(disk.get("drive")) or "?",
            "Label": clean_text(disk.get("label")) or None,
            "FileSystem": clean_text(disk.get("fileSystem")) or None,
            "Capacity": fmt_bytes(size),
            "Free": fmt_bytes(free),
            "FreePercent": rounded(free / size * 100.0) if size > 0 else None,
        })

    design = to_number(battery_data.get("designCapacity"))
    full = to_number(battery_data.get("fullCapacity"))
    battery = {
        "Present": bool(battery_data.get("present")),
        "Status": BATTERY_STATUS_TEXT.get(int(to_number(battery_data.get("batteryStatus")) or 0), clean_text(battery_data.get("statusText")) or None) if battery_data.get("present") else None,
        "ChargePercent": to_number(battery_data.get("chargePercent")),
        "EstimatedMinutes": to_number(battery_data.get("estimatedMinutes")),
        "DesignCapacity": design,
        "FullCapacity": full,
        "HealthPercent": rounded(full / design * 100.0) if design and full else None,
        "CycleCount": to_number(battery_data.get("cycleCount")),
        "CapacitySource": clean_text(battery_data.get("capacitySource")) or ("Not exposed" if battery_data.get("present") else None),
        "PowercfgStatus": clean_text(battery_data.get("powercfgStatus")) or None,
        "PowercfgCrossCheck": None,
        "ActivePowerPlan": clean_text(battery_data.get("activePowerPlan")) or None,
        "PowerSourceAtStart": {True: "Mains", False: "Battery", None: "Unknown"}[power_source_from_status(battery_data.get("batteryStatus"), battery_data.get("present"))] if battery_data else "Unknown",
        "WmiErrors": [clean_text(e, 200) for e in as_list(battery_data.get("wmiErrors"))],
    }
    powercfg_full = to_number(battery_data.get("powercfgFull"))
    powercfg_design = to_number(battery_data.get("powercfgDesign"))
    if design and full and powercfg_full and powercfg_design and battery["CapacitySource"] != "powercfg /batteryreport":
        battery["PowercfgCrossCheck"] = f"powercfg reports {powercfg_full / powercfg_design * 100.0:.0f}% (full {powercfg_full:.0f} / design {powercfg_design:.0f})"

    thermal = [{"Zone": clean_text(z.get("zone")) or "ACPI thermal zone", "TemperatureC": to_number(z.get("temperatureC"))} for z in records.items("thermal")]
    network = [{
        "Name": clean_text(a.get("name")) or None,
        "Description": clean_text(a.get("description")) or None,
        "Status": clean_text(a.get("status")) or None,
        "LinkSpeed": clean_text(a.get("linkSpeed")) or None,
        "Driver": clean_text(a.get("driver")) or None,
        "DriverDate": clean_text(a.get("driverDate")) or None,
        "DriverVersion": clean_text(a.get("driverVersion")) or None,
    } for a in records.items("network")]
    wifi_data = records.data("wifi")
    wifi = None
    if wifi_data and wifi_data.get("parsed"):
        wifi = {key: clean_text(wifi_data.get(source)) or None for key, source in (("State", "state"), ("Signal", "signal"), ("ReceiveMbps", "receiveMbps"), ("TransmitMbps", "transmitMbps"), ("RadioType", "radioType"), ("Band", "band"), ("Channel", "channel"), ("SSID", "ssid"))}
    problem_devices = [{"Device": clean_text(d.get("device")) or None, "ErrorCode": to_number(d.get("errorCode")), "Status": clean_text(d.get("status")) or None, "Class": clean_text(d.get("pnpClass")) or None} for d in records.items("problem_devices")]

    tpm_status = records.status("tpm")
    tpm_data = records.data("tpm") or {}
    if tpm_status == "collected":
        tpm_text = f"Present={tpm_data.get('present')}; Ready={tpm_data.get('ready')}; Enabled={tpm_data.get('enabled')}"
    elif tpm_status == "requires_elevation":
        tpm_text = "Not checked (requires elevation)"
    else:
        tpm_text = f"Unknown ({status_label(tpm_status)})"
    bitlocker_status = records.status("bitlocker")
    defender_data = records.data("defender")
    security = {
        "SecureBoot": system["SecureBoot"],
        "TPM": tpm_text,
        "Antivirus": [{"Product": clean_text(a.get("product")) or None, "State": decode_antivirus_state(a.get("productState")), "ProductState": to_number(a.get("productState"))} for a in records.items("antivirus")],
        "FirewallProfiles": [{"Name": clean_text(f.get("name")) or None, "Enabled": clean_text(f.get("enabled")) or None, "DefaultInboundAction": clean_text(f.get("defaultInbound")) or None, "DefaultOutboundAction": clean_text(f.get("defaultOutbound")) or None} for f in records.items("firewall")],
        "BitLocker": [{"MountPoint": clean_text(b.get("mountPoint")) or None, "VolumeStatus": clean_text(b.get("volumeStatus")) or None, "ProtectionStatus": clean_text(b.get("protectionStatus")) or None, "EncryptionMethod": clean_text(b.get("encryptionMethod")) or None, "EncryptionPercentage": to_number(b.get("encryptionPercentage"))} for b in records.items("bitlocker")],
        "BitLockerStatus": "Collected" if bitlocker_status == "collected" else ("Not checked (requires elevation)" if bitlocker_status == "requires_elevation" else status_label(bitlocker_status)),
        "Defender": {
            "AMServiceEnabled": defender_data.get("amServiceEnabled"),
            "AntivirusEnabled": defender_data.get("antivirusEnabled"),
            "RealTimeProtection": defender_data.get("realTimeProtection"),
            "SignatureAgeDays": to_number(defender_data.get("signatureAgeDays")),
            "QuickScanAgeDays": to_number(defender_data.get("quickScanAgeDays")),
            "RunningMode": clean_text(defender_data.get("amRunningMode")) or None,
        } if defender_data else {"Status": status_label(records.status("defender"))},
    }

    startup = [{"Name": clean_text(s.get("name")) or None, "Command": clean_text(s.get("command"), 300) or None, "Location": clean_text(s.get("location")) or None, "User": clean_text(s.get("user")) or None} for s in records.items("startup")]
    services = [{"Service": clean_text(s.get("service")) or None, "Name": clean_text(s.get("name")) or None, "State": clean_text(s.get("state")) or None, "StartMode": clean_text(s.get("startMode")) or None, "DelayedStart": s.get("delayed")} for s in records.items("services")]
    hotfixes = [{"HotFixID": clean_text(h.get("hotFixId")) or None, "Description": clean_text(h.get("description")) or None, "InstalledOn": clean_text(h.get("installedOn"), 30) or None} for h in records.items("hotfixes")]

    event_summary = []
    if events_data:
        for log in as_dict_list(events_data.get("logs")):
            for group in as_dict_list(log.get("groups")):
                event_summary.append({
                    "Log": clean_text(log.get("log")),
                    "Provider": clean_text(group.get("provider")) or None,
                    "EventId": int(to_number(group.get("eventId")) or 0),
                    "Level": clean_text(group.get("levelName")) or None,
                    "Count": int(to_number(group.get("count")) or 0),
                    "MostRecent": clean_text(group.get("mostRecent"), 40) or None,
                    "Message": clean_text(group.get("message"), 240) or None,
                })
    event_summary.sort(key=lambda row: row["Count"], reverse=True)
    diagnostics = event_diagnostics(events_data)
    event_days = int(to_number(events_data.get("days")) or options.event_days) if events_data else options.event_days

    reliability_data = records.data("reliability") or {}
    reliability = [{"Time": clean_text(r.get("time"), 40) or None, "Source": clean_text(r.get("source")) or None, "Product": clean_text(r.get("product")) or None, "EventId": to_number(r.get("eventId"))} for r in as_dict_list(reliability_data.get("items"))]
    reliability_total = int(to_number(reliability_data.get("total")) or len(reliability))
    installed = [{"Application": clean_text(a.get("application")) or None, "Version": clean_text(a.get("version")) or None, "Publisher": clean_text(a.get("publisher")) or None} for a in records.items("installed_apps")]

    deep = {
        "DISMCheckHealth": classify_dism(records.status("dism_checkhealth"), records.data("dism_checkhealth"), records.error("dism_checkhealth")),
        "SFCVerifyOnly": classify_sfc(records.status("sfc_verifyonly"), records.data("sfc_verifyonly"), records.error("sfc_verifyonly")),
    }

    performance = analyse_performance(records, options, logical_processors, total_ram_bytes / (1024 ** 2) if total_ram_bytes else 0.0)
    workload = {
        "Label": options.workload_label,
        "Notes": options.workload_notes or None,
        "DockConnected": options.docked,
        "PowerSource": performance["PowerSource"],
        "ActivePowerPlan": battery["ActivePowerPlan"],
    }
    coverage = build_coverage(records, run_meta, options)

    report: Dict[str, Any] = {
        "SchemaVersion": SCHEMA_VERSION,
        "Tool": TOOL_NAME,
        "GeneratedBy": f"{TOOL_NAME} {TOOL_VERSION} (single-file .pyw)",
        "HealthScore": None,
        "Rating": None,
        "Headline": None,
        "Completeness": None,
        "Findings": [],
        "Coverage": coverage,
        "Workload": workload,
        "System": system,
        "CPU": cpu,
        "MemorySummary": memory_summary,
        "MemoryModules": modules,
        "GPUs": gpus,
        "PhysicalDisks": physical,
        "LogicalDisks": logical,
        "Battery": battery,
        "ThermalZones": thermal,
        "Performance": performance,
        "NetworkAdapters": network,
        "WiFi": wifi,
        "ProblemDevices": problem_devices,
        "Security": security,
        "StartupItems": startup,
        "AutoServicesStopped": services,
        "RecentHotfixes": hotfixes,
        "EventDays": event_days,
        "EventSummary": event_summary,
        "HardwareEvents": diagnostics["Details"],
        "EventDiagnostics": {key: value for key, value in diagnostics.items() if key != "Details"},
        "ReliabilityTotal": reliability_total,
        "ReliabilityRecords": reliability,
        "InstalledApps": installed,
        "DeepChecks": deep,
        "CollectionNotes": [],
        "MetricDefinitions": METRIC_DEFINITIONS,
        "Options": options.as_dict(),
        "RunMeta": {key: run_meta.get(key) for key in ("ExitCode", "Cancelled", "Started", "Finished", "TempCleanup") if key in run_meta},
        "Limitations": [
            "A short snapshot cannot prove that a fault never occurs; repeat with the same workload label on different days.",
            "Standard Windows interfaces often do not expose accurate CPU/GPU temperature or fan speed.",
            "Physical weight, dust, fan obstruction, thermal paste, battery swelling and charger wattage require physical inspection.",
            "The score is a triage aid, not a warranty, benchmark or replacement for company IT diagnosis.",
            "Memory_*Percent is physical RAM in use (100 minus available); Commit_*Percent includes the pagefile.",
            "The sample count is fixed and the elapsed time is recorded, because each sample takes several seconds of counter reads.",
            "Top-process CPU combines all instances of an application and is averaged over every sample, not only samples in which the process ran; the RAM list is a single reading at the end of sampling.",
            "Disk latency is derived from raw counters between consecutive samples and needs at least two samples per disk.",
            "Threshold values are triage heuristics, not standards.",
        ],
    }

    notes = report["CollectionNotes"]
    for row in coverage:
        if row["Status"] not in ("collected", "not_run", "not_applicable"):
            notes.append(single_line(f"{row['Check']}: {row['StatusLabel']}. {row['Detail']}"))
    if run_meta.get("Stderr"):
        notes.append("Collector stderr: " + clean_text(run_meta.get("Stderr"), 500))
    for line in records.noise[:10]:
        notes.append("Unparsed collector output: " + line)
    if records.fatal:
        notes.append("Collector fatal error: " + clean_text(records.fatal.get("error"), 400))

    problem_statuses = {row["Key"]: row["Status"] for row in coverage}
    core_missing, elevation_missing = classify_coverage(coverage, str(problem_statuses.get("samples_done")))
    missing_checks = core_missing + elevation_missing
    report["Completeness"] = {"Status": "pending", "MissingChecks": missing_checks, "CoreMissing": core_missing, "ElevationMissing": elevation_missing}
    findings = evaluate_findings(report, records, coverage)
    if findings.count("Critical", "Warning") == 0:
        if not missing_checks:
            findings.add("Good", "Overall", "No material issue was detected in this snapshot.", "All checks were collected and remained within the tool thresholds.", "If slowness is intermittent, rerun Standard mode while the problem is occurring, with the workload label set.", 0)
        elif not core_missing:
            findings.add("Good", "Overall", "No material issue was detected in the collected checks.", "Every diagnostic check was collected and remained within the tool thresholds. Not performed (elevation only): " + "; ".join(elevation_missing)[:300] + ".", "Rerun as administrator only if Secure Boot, TPM, BitLocker, thermal zones or SSD wear matter to the ticket.", 0)
        else:
            findings.add("Information", "Overall", "No material issue was detected in the checks that were collected, but evidence is missing.", f"{len(core_missing)} diagnostic check(s) were not collected: " + "; ".join(core_missing)[:400], "This is not a clean bill of health; see the coverage table for the missing evidence and rerun the audit.", 0)
    score, rating, headline = score_and_rating(findings, core_missing, elevation_missing)
    report["Findings"] = findings.sorted()
    report["HealthScore"] = score
    report["Rating"] = rating
    report["Headline"] = headline
    report["Completeness"] = {
        "Status": "Complete" if not missing_checks else ("Complete except elevation-only checks" if not core_missing else "Incomplete"),
        "MissingChecks": missing_checks,
        "CoreMissing": core_missing,
        "ElevationMissing": elevation_missing,
        "Collected": sum(1 for row in coverage if row["Status"] == "collected"),
        "Partial": sum(1 for row in coverage if row["Status"] == "partial"),
        "Unavailable": sum(1 for row in coverage if row["Status"] == "unavailable"),
        "RequiresElevation": sum(1 for row in coverage if row["Status"] == "requires_elevation"),
        "Failed": sum(1 for row in coverage if row["Status"] == "failed"),
        "NotApplicable": sum(1 for row in coverage if row["Status"] == "not_applicable"),
        "NotRun": sum(1 for row in coverage if row["Status"] == "not_run"),
    }
    return report


# ---------------------------------------------------------------------------
# Report writers
# ---------------------------------------------------------------------------

def h(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def html_table(title: str, rows: Sequence[Dict[str, Any]], columns: Sequence[Tuple[str, str]], note: str = "", limit: int = 0) -> str:
    """columns: (key, header). Values are formatted plainly; numbers keep one decimal."""
    if not rows:
        return f"<section><h2>{h(title)}</h2><p class='muted'>No data available.</p></section>"
    shown = list(rows[:limit]) if limit else list(rows)
    header = "".join(f"<th>{h(header_text)}</th>" for _, header_text in columns)
    body = []
    for row in shown:
        cells = []
        for key, _ in columns:
            value = row.get(key)
            if isinstance(value, float):
                value = f"{value:.1f}" if abs(value) < 1000 else f"{value:,.0f}"
            elif isinstance(value, bool):
                value = "Yes" if value else "No"
            elif value is None:
                value = ""
            cells.append(f"<td>{h(value)}</td>")
        body.append("<tr>" + "".join(cells) + "</tr>")
    footer = ""
    if limit and len(rows) > limit:
        footer = f"<p class='muted'>Showing {limit} of {len(rows)} rows; the JSON file holds every row.</p>"
    if note:
        footer += f"<p class='muted'>{h(note)}</p>"
    return f"<section><h2>{h(title)}</h2><div class='table-wrap'><table><thead><tr>{header}</tr></thead><tbody>{''.join(body)}</tbody></table></div>{footer}</section>"


def html_key_values(title: str, data: Dict[str, Any], keys: Sequence[str], note: str = "") -> str:
    rows = [{"Field": key, "Value": data.get(key)} for key in keys if key in data]
    return html_table(title, rows, (("Field", "Field"), ("Value", "Value")), note)


HTML_STYLE = (
    ":root{--ink:#172033;--muted:#657084;--line:#dfe4ec;--panel:#fff;--bg:#f4f7fb;--blue:#155eef;--critical:#b42318;--warning:#b54708;--good:#067647;--info:#175cd3}"
    "*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 'Segoe UI',Arial,sans-serif}.shell{max-width:1280px;margin:auto;padding:28px}"
    ".hero{background:linear-gradient(135deg,#132238,#244f85);color:#fff;padding:28px;border-radius:16px}.hero h1{margin:0 0 6px;font-size:30px}.hero p{margin:4px 0;color:#dbeafe}"
    ".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:16px 0}.card,section{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px}"
    ".metric{font-size:26px;font-weight:700}.label,.muted{color:var(--muted)}section{margin:14px 0}h2{margin:0 0 12px;font-size:18px}.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:620px}"
    "th,td{text-align:left;vertical-align:top;border-bottom:1px solid var(--line);padding:8px 10px}th{background:#f8fafc;color:#475467;font-size:12px;text-transform:uppercase;letter-spacing:.03em}"
    ".badge{display:inline-block;border-radius:999px;padding:3px 8px;font-size:12px;font-weight:700}.critical{color:var(--critical);background:#fee4e2}.warning{color:var(--warning);background:#fef0c7}"
    ".information{color:var(--info);background:#dbeafe}.good{color:var(--good);background:#d1fadf}.collected{color:var(--good)}.partial,.requires_elevation{color:var(--warning)}.failed,.unavailable{color:var(--critical)}"
    ".incomplete{background:#fef0c7;border:1px solid #f5c66b;border-radius:10px;padding:12px 16px;margin:14px 0}pre{white-space:pre-wrap;font:12px/1.4 Consolas,monospace;background:#f8fafc;padding:10px;border-radius:8px;max-height:320px;overflow:auto}"
    ".foot{padding:14px 4px;color:var(--muted);font-size:12px}@media print{body{background:#fff}.shell{max-width:none;padding:0}section{break-inside:avoid}}"
)


def html_report(report: Dict[str, Any]) -> str:
    system = report["System"]
    performance = report["Performance"]
    battery = report["Battery"]
    completeness = report["Completeness"]
    sections: List[str] = []

    finding_rows = []
    for finding in report["Findings"]:
        css = finding["Severity"].lower()
        finding_rows.append(
            f"<tr><td><span class='badge {css}'>{h(finding['Severity'])}</span></td><td>{h(finding['Category'])}</td><td>{h(finding['Finding'])}</td><td>{h(finding['Evidence'])}</td><td>{h(finding['Recommendation'])}</td></tr>"
        )
    sections.append(
        "<section><h2>Findings and recommended next steps</h2><div class='table-wrap'><table><thead><tr><th>Severity</th><th>Category</th><th>Finding</th><th>Evidence</th><th>Recommendation</th></tr></thead><tbody>"
        + "".join(finding_rows) + "</tbody></table></div></section>"
    )
    coverage_rows = []
    for row in report["Coverage"]:
        coverage_rows.append(f"<tr><td>{h(row['Check'])}</td><td class='{h(row['Status'])}'>{h(row['StatusLabel'])}</td><td>{h(row['Detail'])}</td></tr>")
    sections.append(
        "<section><h2>Coverage: what was collected and what was not</h2><div class='table-wrap'><table><thead><tr><th>Check</th><th>Status</th><th>Detail</th></tr></thead><tbody>"
        + "".join(coverage_rows) + "</tbody></table></div></section>"
    )
    workload = dict(report["Workload"])
    workload["DockConnected"] = {True: "Yes", False: "No", None: "Not stated"}.get(workload.get("DockConnected"), "Not stated")
    sections.append(html_key_values("Workload during the sample (user-stated and measured)", workload, ["Label", "Notes", "DockConnected", "PowerSource", "ActivePowerPlan"]))
    sections.append(html_key_values("System", system, ["Manufacturer", "Model", "SystemFamily", "ChassisType", "WindowsEdition", "WindowsVersion", "BuildNumber", "Architecture", "BIOSVersion", "BIOSDate", "SecureBoot", "LastBoot", "UptimeDays", "PendingReboot", "Elevated", "PowerShellVersion", "DeviceFingerprint", "AuditTime"]))
    sections.append(html_key_values("Processor", report["CPU"], ["Name", "PhysicalCores", "LogicalProcessors", "MaximumClockMHz", "CurrentClockMHz", "L3CacheKB", "VirtualisationFW"]))
    sections.append(html_key_values("Memory overview", report["MemorySummary"], ["Installed", "PopulatedModules", "TotalSlots", "MaximumSupported", "MaximumSupportedNote", "PageFiles", "CommitLimitGB"]))
    sections.append(html_table("Memory modules", report["MemoryModules"], (("Slot", "Slot"), ("Capacity", "Capacity"), ("SpeedMHz", "Speed MHz"), ("ConfiguredMHz", "Configured MHz"), ("Manufacturer", "Manufacturer"), ("PartNumber", "Part number"))))
    sections.append(html_table("Physical storage", report["PhysicalDisks"], (("Name", "Name"), ("MediaType", "Media"), ("BusType", "Bus"), ("Capacity", "Capacity"), ("Health", "Health"), ("OperationalStatus", "Operational"), ("TemperatureC", "Temp C"), ("WearPercent", "Wear %"), ("ReadErrorsUncorrected", "Read err (uncorr.)"), ("WriteErrorsUncorrected", "Write err (uncorr.)"), ("PowerOnHours", "Power-on h")), "Health 'Unknown' means the value was not reported, not that the disk is unhealthy."))
    sections.append(html_table("Volumes", report["LogicalDisks"], (("Drive", "Drive"), ("FileSystem", "File system"), ("Capacity", "Capacity"), ("Free", "Free"), ("FreePercent", "Free %"))))
    sections.append(html_table("Graphics", report["GPUs"], (("Name", "Name"), ("DedicatedMemory", "Memory"), ("MemorySource", "Memory source"), ("DriverVersion", "Driver"), ("DriverDate", "Driver date"), ("CurrentMode", "Mode"))))
    sections.append(html_key_values("Battery and power", battery, ["Present", "Status", "ChargePercent", "DesignCapacity", "FullCapacity", "HealthPercent", "CycleCount", "CapacitySource", "PowercfgCrossCheck", "PowerSourceAtStart", "ActivePowerPlan"]))
    sections.append(html_table("ACPI thermal readings", report["ThermalZones"], (("Zone", "Zone"), ("TemperatureC", "Temperature C"))))

    perf_keys = ["SamplesTaken", "SamplesRequested", "ElapsedSeconds", "AverageSampleSeconds", "CPU_AveragePercent", "CPU_P95Percent", "CPU_MaxPercent", "Core_P95Max", "Core_Hottest", "SingleThreadBound", "CPU_Performance_AveragePercent", "CPU_Performance_P05Percent", "CPU_Performance_ValidSamples", "CPU_FrequencyMHz_Average", "Memory_AveragePercent", "Memory_P95Percent", "AvailableRAM_MinMB", "AvailableRAM_AverageMB", "Commit_AveragePercent", "Commit_P95Percent", "PagesInputPerSec_Average", "PagesInputPerSec_P95", "DiskActive_Average", "DiskActive_P95", "DiskQueue_P95", "DiskLatency_ms_P95", "DiskReadMBps_Average", "DiskWriteMBps_Average", "PowerSource", "OnMainsSamples", "OnBatterySamples"]
    sections.append(html_key_values("Live performance sample (Memory = physical RAM in use; CPU_Performance = frequency vs nominal)", performance, perf_keys))
    throttle = performance["Throttling"]
    sections.append(html_key_values("Throttling evaluation (load and frequency compared within the same timestamped sample)", throttle, ["Rule", "PairedSamples", "LoadedSamples", "OverlapSamples", "OverlapOnBattery", "MinimumOverlap", "Suspected", "Evaluable"]))
    sections.append(html_table("Per logical processor", performance["PerCore"], (("Core", "Core"), ("Average", "Average %"), ("P95", "P95 %"), ("Max", "Max %"), ("Samples", "Samples"))))
    sections.append(html_table("Per physical disk (latency from raw counters)", performance["PerDisk"], (("Disk", "Disk"), ("Active_Average", "Active avg %"), ("Active_P95", "Active P95 %"), ("Queue_P95", "Queue P95"), ("ReadMBps_Average", "Read MB/s"), ("WriteMBps_Average", "Write MB/s"), ("TransfersPerSec_Average", "Transfers/s"), ("Latency_ms_Average", "Latency avg ms"), ("Latency_ms_P95", "Latency P95 ms"), ("ReadLatency_ms_P95", "Read P95 ms"), ("WriteLatency_ms_P95", "Write P95 ms"), ("LatencyIntervals", "Intervals"))))
    sections.append(html_table("Top applications by measured CPU (all instances combined)", performance["TopCPUProcesses"], (("Process", "Process"), ("Instances", "Instances"), ("AverageCPU", "Average CPU %"), ("SamplesPresent", "Samples present"), ("PeakRAM_MB", "Peak private MB"), ("IO_MBps_Average", "I/O MB/s"))))
    sections.append(html_table("Top applications by disk/network I/O", performance["TopIOProcesses"], (("Process", "Process"), ("IO_MBps_Average", "I/O MB/s"), ("AverageCPU", "Average CPU %"))))
    sections.append(html_table("Top applications by current RAM (all instances combined)", performance["TopRAMProcesses"], (("Process", "Process"), ("Instances", "Instances"), ("RAM_MB", "Working set MB"), ("Private_MB", "Private MB"))))
    sections.append(html_table("Sample timeline", performance["Samples"], (("Time", "Time"), ("CPU", "CPU %"), ("CoreMax", "Busiest core %"), ("Performance", "Perf %"), ("AvailableMB", "Available MB"), ("PagesInputPerSec", "Pages in/s"), ("DiskActive", "Disk %"), ("DiskQueue", "Queue"), ("OnMains", "Mains")), limit=200))
    sections.append(html_table("Hardware and power events (all levels, providers of interest)", report["HardwareEvents"], (("Provider", "Provider"), ("EventId", "Event"), ("Level", "Level"), ("Count", "Count"), ("MostRecent", "Most recent"), ("Message", "Message (truncated)")), f"Last {report['EventDays']} days. Diagnosis uses every collected group; the display is capped.", limit=40))
    sections.append(html_table(f"Critical and error event groups - last {report['EventDays']} days", report["EventSummary"], (("Log", "Log"), ("Provider", "Provider"), ("EventId", "Event"), ("Count", "Count"), ("MostRecent", "Most recent")), limit=40))
    sections.append(html_table("Reliability records", report["ReliabilityRecords"], (("Time", "Time"), ("Source", "Source"), ("Product", "Product"), ("EventId", "Event")), f"{report['ReliabilityTotal']} matching record(s) in total.", limit=40))
    sections.append(html_table("Problem devices", report["ProblemDevices"], (("Device", "Device"), ("Class", "Class"), ("ErrorCode", "Error code"), ("Status", "Status"))))
    sections.append(html_table("Physical network adapters", report["NetworkAdapters"], (("Name", "Name"), ("Description", "Description"), ("Status", "Status"), ("LinkSpeed", "Link speed"), ("Driver", "Driver"), ("DriverDate", "Driver date"), ("DriverVersion", "Version"))))
    if report["WiFi"]:
        sections.append(html_key_values("Wi-Fi link (no internet test)", report["WiFi"], ["State", "Signal", "ReceiveMbps", "TransmitMbps", "RadioType", "Band", "Channel", "SSID"]))
    security = report["Security"]
    sections.append(html_key_values("Security posture", {"SecureBoot": security["SecureBoot"], "TPM": security["TPM"], "BitLockerStatus": security["BitLockerStatus"]}, ["SecureBoot", "TPM", "BitLockerStatus"]))
    sections.append(html_table("Antivirus products", security["Antivirus"], (("Product", "Product"), ("State", "State"))))
    if isinstance(security.get("Defender"), dict):
        sections.append(html_key_values("Microsoft Defender status", security["Defender"], list(security["Defender"].keys())))
    sections.append(html_table("Firewall profiles", security["FirewallProfiles"], (("Name", "Name"), ("Enabled", "Enabled"), ("DefaultInboundAction", "Inbound"), ("DefaultOutboundAction", "Outbound"))))
    sections.append(html_table("BitLocker status", security["BitLocker"], (("MountPoint", "Volume"), ("VolumeStatus", "Status"), ("ProtectionStatus", "Protection"), ("EncryptionMethod", "Method"), ("EncryptionPercentage", "%"))))
    sections.append(html_table("Startup entries", report["StartupItems"], (("Name", "Name"), ("Command", "Command"), ("Location", "Location"), ("User", "User"))))
    sections.append(html_table("Automatic services not running", report["AutoServicesStopped"], (("Service", "Service"), ("Name", "Name"), ("State", "State"), ("StartMode", "Start mode"), ("DelayedStart", "Delayed"))))
    sections.append(html_table("Recent Windows hotfixes", report["RecentHotfixes"], (("HotFixID", "Hotfix"), ("Description", "Description"), ("InstalledOn", "Installed"))))
    deep = report["DeepChecks"]
    deep_rows = [{"Check": "DISM /Online /Cleanup-Image /CheckHealth", **deep["DISMCheckHealth"]}, {"Check": "SFC /verifyonly", **deep["SFCVerifyOnly"]}]
    deep_html = html_table("Windows integrity checks (verification only)", deep_rows, (("Check", "Check"), ("Status", "Status"), ("Verdict", "Verdict"), ("ExitCode", "Exit code")))
    for row in deep_rows:
        if row.get("Output") and row["Status"] == "Collected":
            deep_html += f"<section><h2>{h(row['Check'])} output</h2><pre>{h(clean_text(row['Output'], 4000))}</pre></section>"
    sections.append(deep_html)
    if report["InstalledApps"]:
        sections.append(html_table("Installed applications", report["InstalledApps"], (("Application", "Application"), ("Version", "Version"), ("Publisher", "Publisher"))))
    sections.append(html_table("Collection notes", [{"Note": note} for note in report["CollectionNotes"]], (("Note", "Note"),)))
    sections.append(html_table("Metric definitions", [{"Metric": key, "Definition": value} for key, value in report["MetricDefinitions"].items()], (("Metric", "Metric"), ("Definition", "Definition"))))
    sections.append(html_table("Interpretation limits", [{"Limit": item} for item in report["Limitations"]], (("Limit", "Limit"),)))

    incomplete = ""
    if completeness.get("CoreMissing"):
        incomplete = f"<div class='incomplete'><strong>Evidence incomplete.</strong> Not collected: {h('; '.join(completeness['CoreMissing'])[:600])}. Missing checks are absent evidence, not a clean result.</div>"
    elif completeness.get("ElevationMissing"):
        incomplete = f"<div class='incomplete'><strong>Elevation-only checks not performed:</strong> {h('; '.join(completeness['ElevationMissing'])[:600])}. Rerun as administrator if they matter to the ticket.</div>"
    battery_metric = f"{fmt_value(battery.get('HealthPercent'), '%', 0)}" if battery.get("HealthPercent") is not None else "N/A"
    evidence_short = {"Complete": "Complete", "Complete except elevation-only checks": "Elevation gaps", "Incomplete": "Incomplete"}.get(completeness["Status"], completeness["Status"])
    return (
        "<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
        f"<title>PC Health Audit</title><style>{HTML_STYLE}</style></head><body><main class='shell'>"
        f"<div class='hero'><h1>PC Health Audit</h1><p>{h((system.get('Manufacturer') or '') + ' ' + (system.get('Model') or ''))}</p>"
        f"<p>Generated {h(system.get('AuditTime'))} | {h(system.get('AuditMode'))} mode | {h(performance.get('SamplesTaken'))} live samples over {h(performance.get('ElapsedSeconds'))} s | workload: {h(report['Workload']['Label'])}</p></div>"
        f"{incomplete}<div class='cards'>"
        f"<div class='card'><div class='label'>Triage result</div><div class='metric'>{h(report['HealthScore'])}/100</div><div>{h(report['Rating'])}</div></div>"
        f"<div class='card'><div class='label'>Evidence</div><div class='metric'>{h(evidence_short)}</div><div>{h(completeness['Collected'])} collected, {h(completeness['Partial'] + completeness['Failed'] + completeness['Unavailable'] + completeness['RequiresElevation'])} gaps</div></div>"
        f"<div class='card'><div class='label'>CPU average</div><div class='metric'>{h(fmt_value(performance.get('CPU_AveragePercent'), '%', 0))}</div><div>P95 {h(fmt_value(performance.get('CPU_P95Percent'), '%', 0))}; busiest core P95 {h(fmt_value(performance.get('Core_P95Max'), '%', 0))}</div></div>"
        f"<div class='card'><div class='label'>Physical RAM average</div><div class='metric'>{h(fmt_value(performance.get('Memory_AveragePercent'), '%', 0))}</div><div>{h(report['MemorySummary']['Installed'])} installed; min available {h(fmt_value(performance.get('AvailableRAM_MinMB'), ' MB', 0))}</div></div>"
        f"<div class='card'><div class='label'>Battery health</div><div class='metric'>{h(battery_metric)}</div><div>Full vs design capacity</div></div>"
        f"</div>{''.join(sections)}"
        "<div class='foot'>This is a read-only triage snapshot. Review the report before sharing. Do not disable company security or management software based on this report.</div>"
        "</main></body></html>"
    )


def text_summary(report: Dict[str, Any]) -> str:
    system = report["System"]
    cpu = report["CPU"]
    memory = report["MemorySummary"]
    battery = report["Battery"]
    performance = report["Performance"]
    completeness = report["Completeness"]
    lines = [
        "PC HEALTH AUDIT - EXECUTIVE SUMMARY",
        "=" * 70,
        f"Generated: {system.get('AuditTime')} ({system.get('AuditMode')} mode, tool {TOOL_VERSION}, schema {SCHEMA_VERSION})",
        f"Device: {system.get('Manufacturer')} {system.get('Model')} | fingerprint {system.get('DeviceFingerprint')}",
        f"OS: {system.get('WindowsEdition')}, version {system.get('WindowsVersion')}, build {system.get('BuildNumber')}",
        f"CPU: {cpu.get('Name')} | {fmt_value(cpu.get('PhysicalCores'), '', 0)} cores / {fmt_value(cpu.get('LogicalProcessors'), '', 0)} logical processors",
        f"RAM: {memory.get('Installed')} | modules {memory.get('PopulatedModules')} / slots {fmt_value(memory.get('TotalSlots'), '', 0)} | firmware maximum {memory.get('MaximumSupported')} (provisional)",
        f"Battery health: {fmt_value(battery.get('HealthPercent'), '%') if battery.get('HealthPercent') is not None else 'Not exposed'} ({battery.get('CapacitySource')})",
        f"Elevated: {system.get('Elevated')} (Secure Boot, TPM, BitLocker, thermal zones and SSD wear need elevation)",
        f"Workload: {report['Workload']['Label']} | power: {performance.get('PowerSource')} | dock: {report['Workload'].get('DockConnected')}",
        f"Performance sample: {performance.get('SamplesTaken')}/{performance.get('SamplesRequested')} samples over {fmt_value(performance.get('ElapsedSeconds'), ' s', 0)} | CPU {fmt_value(performance.get('CPU_AveragePercent'), '%')} avg, {fmt_value(performance.get('CPU_P95Percent'), '%')} P95, busiest core P95 {fmt_value(performance.get('Core_P95Max'), '%')} | physical RAM {fmt_value(performance.get('Memory_AveragePercent'), '%')} avg, {fmt_value(performance.get('Memory_P95Percent'), '%')} P95 | pages in {fmt_value(performance.get('PagesInputPerSec_Average'), '/s', 0)} avg | disk {fmt_value(performance.get('DiskActive_Average'), '%')} avg, latency P95 {fmt_value(performance.get('DiskLatency_ms_P95'), ' ms')}",
        f"CPU frequency vs nominal (5th percentile / average): {fmt_value(performance.get('CPU_Performance_P05Percent'), '%', 0)} / {fmt_value(performance.get('CPU_Performance_AveragePercent'), '%', 0)} over {performance.get('CPU_Performance_ValidSamples')} valid samples; throttling suspected: {performance['Throttling'].get('Suspected')}",
        "Top applications by measured CPU: " + "; ".join(f"{row['Process']} {fmt_value(row['AverageCPU'], '%')}" for row in performance["TopCPUProcesses"][:5]),
        "Top applications by RAM: " + "; ".join(f"{row['Process']} {fmt_value(row['RAM_MB'], ' MB', 0)}" for row in performance["TopRAMProcesses"][:5]),
        f"Triage result: {report['Headline']}",
        f"Evidence: {completeness['Status']} ({completeness['Collected']} collected, {completeness['Partial']} partial, {completeness['Unavailable']} unavailable, {completeness['RequiresElevation']} need elevation, {completeness['Failed']} failed)",
        "",
        "FINDINGS",
        "-" * 70,
    ]
    for finding in report["Findings"]:
        lines.append(f"[{finding['Severity']}] {finding['Category']}: {finding['Finding']}")
        lines.append(f"Evidence: {finding['Evidence']}")
        lines.append(f"Next step: {finding['Recommendation']}")
        lines.append("")
    lines.extend(["COVERAGE GAPS", "-" * 70])
    gaps = [row for row in report["Coverage"] if row["Status"] not in ("collected", "not_run", "not_applicable")]
    if gaps:
        lines.extend(f"{row['Check']}: {row['StatusLabel']} - {row['Detail']}" for row in gaps)
    else:
        lines.append("None. Every check was collected.")
    lines.extend([
        "",
        "PRIVACY",
        "-" * 70,
        "Computer name, user name, serial numbers, volume labels, Wi-Fi identity, MAC addresses and startup commands are redacted by default.",
        "Review every report before sharing outside your company.",
        "",
    ])
    return "\n".join(lines)


def write_reports(report: Dict[str, Any], folder: str, records: Optional[List[Dict[str, Any]]] = None) -> Dict[str, str]:
    os.makedirs(folder, exist_ok=True)
    paths = {
        "html": os.path.join(folder, "PC_Health_Report.html"),
        "json": os.path.join(folder, "PC_Health_Data.json"),
        "summary": os.path.join(folder, "PC_Health_Summary.txt"),
    }
    with open(paths["json"], "w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False, default=str)
    with open(paths["summary"], "w", encoding="utf-8") as handle:
        handle.write(text_summary(report))
    with open(paths["html"], "w", encoding="utf-8") as handle:
        handle.write(html_report(report))
    if records is not None:
        paths["records"] = os.path.join(folder, "PC_Health_Records.jsonl")
        with open(paths["records"], "w", encoding="utf-8") as handle:
            for record in records:
                handle.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")
    return paths


def create_zip(folder: str, zip_path: str) -> str:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for root, _, files in os.walk(folder):
            for name in files:
                full = os.path.join(root, name)
                archive.write(full, os.path.relpath(full, os.path.dirname(folder)))
    return zip_path


def write_battery_html_report(folder: str) -> Optional[str]:
    """Only with sensitive data enabled: the HTML battery report names the machine and serial."""
    if not IS_WINDOWS:
        return None
    target = os.path.join(folder, "Battery_Report.html")
    executable = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "powercfg.exe")
    try:
        subprocess.run([executable, "/batteryreport", "/output", target], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60, **hidden_popen_kwargs())
    except Exception:
        return None
    return target if os.path.exists(target) else None


def run_audit(
    options: AuditOptions,
    progress: Optional[Callable[[str, Optional[float]], None]] = None,
    cancel_event: Optional[threading.Event] = None,
    launcher: Callable[[str, List[str]], List[str]] = default_launcher,
) -> Dict[str, Any]:
    """Collect, analyse and write. Returns the report and file paths."""
    notify = progress or (lambda message, fraction=None: None)
    timestamp = now_local().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(options.output_dir, f"PC_Health_Audit_{timestamp}")
    os.makedirs(folder, exist_ok=True)
    notify(f"Output folder: {folder}", 0.0)
    records: List[Dict[str, Any]] = []
    static_steps = len(SECTION_LABELS)
    seen_sections = 0

    def on_record(record: Dict[str, Any]) -> None:
        nonlocal seen_sections
        records.append(record)
        section = str(record.get("section", ""))
        status = str(record.get("status", ""))
        if section == "progress":
            notify(clean_text(record.get("message")), None)
        elif section == "sample":
            index = int(to_number(record.get("index")) or 0) + 1
            total = int(to_number(record.get("total")) or options.sample_count) or 1
            notify(f"Performance sample {index}/{total}", 0.35 + 0.6 * (index / total))
        elif section == "fatal":
            notify("Collector error: " + clean_text(record.get("error"), 300), None)
        elif section in SECTION_LABELS:
            seen_sections += 1
            label = SECTION_LABELS[section]
            if status in ("ok", "not_run"):
                notify(f"{label}: collected" if status == "ok" else f"{label}: not run", min(0.35, 0.35 * seen_sections / static_steps))
            else:
                notify(f"{label}: {status_label(status)} - {clean_text(record.get('error'), 160)}", None)

    run_meta = run_collector(options, on_record, cancel_event, launcher)
    if run_meta.get("Cancelled"):
        notify("Cancelled. Writing the partial report.", None)
    notify("Analysing evidence", 0.96)
    report = analyse_records(records, options, run_meta)
    paths = write_reports(report, folder, records)
    if options.include_sensitive and report["Battery"].get("Present"):
        battery_html = write_battery_html_report(folder)
        if battery_html:
            paths["battery"] = battery_html
    if options.create_zip:
        zip_path = os.path.join(options.output_dir, f"PC_Health_Audit_{timestamp}.zip")
        try:
            paths["zip"] = create_zip(folder, zip_path)
        except OSError as exc:
            notify(f"ZIP not created: {exc}", None)
    notify(f"Done: {report['Headline']}", 1.0)
    return {"Report": report, "Folder": folder, "Paths": paths, "RunMeta": run_meta}


# ---------------------------------------------------------------------------
# Comparison of earlier audits
# ---------------------------------------------------------------------------

class LoadedReport:
    def __init__(self, path: Path, data: Dict[str, Any]) -> None:
        self.path = path
        self.data = data
        self.warnings: List[str] = []
        self.schema = clean_text(data.get("SchemaVersion")) or "1.0"
        self.time, self.assumed_local = parse_timestamp(nested(data, "System", "AuditTime"))
        if self.time is None:
            self.warnings.append("audit time missing or unreadable; ordered last")
        elif self.assumed_local:
            self.warnings.append("audit time has no UTC offset; assumed local time of this computer")
        system = data.get("System") if isinstance(data.get("System"), dict) else {}
        self.manufacturer = clean_text(system.get("Manufacturer"))
        self.model = clean_text(system.get("Model"))
        explicit = clean_text(system.get("DeviceFingerprint"))
        cpu_name = clean_text(nested(data, "CPU", "Name"))
        ram = to_number(nested(data, "MemorySummary", "InstalledBytes"))
        if explicit:
            self.device_key = explicit
        elif self.manufacturer or self.model or cpu_name:
            self.device_key = "f-" + fingerprint(self.manufacturer, self.model, cpu_name, ram)
        else:
            self.device_key = "unknown"
        self.device_label = f"{self.manufacturer} {self.model}".strip() or "Unknown device"
        self.memory_metric = "commit" if self.schema == "1.0" else "physical"

    @property
    def sort_key(self) -> Tuple[int, float]:
        return (0, self.time.timestamp()) if self.time else (1, 0.0)

    def time_text(self) -> str:
        if not self.time:
            return "Unknown"
        return self.time.isoformat(timespec="minutes") + (" (assumed local)" if self.assumed_local else "")


def validate_report(data: Any) -> List[str]:
    """Structural checks. Returns a list of problems; empty means acceptable."""
    problems = []
    if not isinstance(data, dict):
        return ["not a JSON object"]
    if data.get("Tool") != TOOL_NAME:
        problems.append("not a PC Health Audit JSON file (Tool field)")
        return problems
    schema = clean_text(data.get("SchemaVersion")) or "1.0"
    if schema not in SUPPORTED_SCHEMAS:
        problems.append(f"unsupported schema version {schema}; supported: {', '.join(SUPPORTED_SCHEMAS)}")
    for key in ("System", "Performance"):
        if key in data and not isinstance(data[key], dict):
            problems.append(f"{key} is not an object")
    for key in ("Findings", "PhysicalDisks", "LogicalDisks"):
        if key in data and data[key] is not None and not isinstance(data[key], list):
            problems.append(f"{key} is not a list")
    return problems


def discover_inputs(inputs: Iterable[str]) -> List[Path]:
    files: set = set()
    for raw in inputs:
        path = Path(str(raw)).expanduser()
        if path.is_dir():
            files.update(path.rglob("PC_Health_Data.json"))
        elif path.is_file():
            files.add(path)
        else:
            files.update(Path(match) for match in glob.glob(str(raw), recursive=True) if Path(match).is_file())
    return sorted(files)


def load_reports(paths: Iterable[Path]) -> Tuple[List[LoadedReport], List[str]]:
    reports: List[LoadedReport] = []
    errors: List[str] = []
    for path in paths:
        try:
            with path.open("r", encoding="utf-8-sig") as handle:
                data = json.load(handle)
        except (OSError, ValueError) as exc:
            errors.append(f"{path}: {exc}")
            continue
        problems = validate_report(data)
        if problems:
            errors.append(f"{path}: " + "; ".join(problems))
            continue
        reports.append(LoadedReport(path, data))
    reports.sort(key=lambda report: report.sort_key)
    return reports, errors


def disk_health_counts(report: LoadedReport) -> Tuple[int, int, int]:
    """(unhealthy, unknown, malformed)."""
    unhealthy = unknown = malformed = 0
    for disk in as_list(report.data.get("PhysicalDisks")):
        if not isinstance(disk, dict):
            malformed += 1
            continue
        health = clean_text(disk.get("Health")).lower()
        if not health or health == "unknown" or disk.get("HealthKnown") is False:
            unknown += 1
        elif health not in ("healthy", "ok"):
            unhealthy += 1
    return unhealthy, unknown, malformed


def lowest_free_percent(report: LoadedReport) -> Tuple[Optional[float], int]:
    values = []
    malformed = 0
    for disk in as_list(report.data.get("LogicalDisks")):
        if not isinstance(disk, dict):
            malformed += 1
            continue
        value = to_number(disk.get("FreePercent"))
        if value is not None:
            values.append(value)
    return (min(values) if values else None), malformed


def report_processes(report: LoadedReport) -> Tuple[Dict[str, float], Dict[str, float], int]:
    """Per-report CPU and RAM by application, summing every entry with the same base name."""
    cpu: Dict[str, float] = defaultdict(float)
    ram: Dict[str, float] = defaultdict(float)
    malformed = 0
    performance = report.data.get("Performance") if isinstance(report.data.get("Performance"), dict) else {}
    for process in as_list(performance.get("TopCPUProcesses")):
        if not isinstance(process, dict):
            malformed += 1
            continue
        value = to_number(process.get("AverageCPU"))
        if value is not None:
            cpu[process_base_name(process.get("Process"))] += value
    for process in as_list(performance.get("TopRAMProcesses")):
        if not isinstance(process, dict):
            malformed += 1
            continue
        value = to_number(process.get("RAM_MB"))
        if value is not None:
            ram[process_base_name(process.get("Process"))] += value
    return cpu, ram, malformed


def recurring_processes(reports: List[LoadedReport]) -> List[List[str]]:
    appearances: Counter = Counter()
    cpu_values: Dict[str, List[float]] = defaultdict(list)
    ram_values: Dict[str, List[float]] = defaultdict(list)
    for report in reports:
        cpu, ram, malformed = report_processes(report)
        if malformed:
            report.warnings.append(f"{malformed} malformed process entries ignored")
        for name, value in cpu.items():
            cpu_values[name].append(value)
        for name, value in ram.items():
            ram_values[name].append(value)
        appearances.update(set(cpu) | set(ram))
    rows: List[Tuple[float, float, List[str]]] = []
    for name, count in appearances.most_common():
        if count < 2 and len(reports) > 1:
            continue
        average_cpu = mean(cpu_values[name]) if cpu_values[name] else None
        peak_ram = max(ram_values[name]) if ram_values[name] else None
        rows.append((average_cpu or 0.0, peak_ram or 0.0, [name, f"{count}/{len(reports)}", fmt_value(average_cpu, "%"), fmt_value(peak_ram, " MB", 0)]))
    rows.sort(key=lambda entry: (entry[0], entry[1]), reverse=True)
    return [row for _, _, row in rows[:20]]


def trend_statement(label: str, first: Any, last: Any, unit: str, adverse_when_higher: bool = True) -> Optional[str]:
    start, end = to_number(first), to_number(last)
    if start is None or end is None or abs(end - start) < 0.5:
        return None
    direction = "increased" if end > start else "decreased"
    adverse = (end > start) == adverse_when_higher
    return f"- {label} {direction} from {start:.1f}{unit} to {end:.1f}{unit} ({'worse' if adverse else 'better'} on this measure)."


def metric_rows(reports: List[LoadedReport]) -> List[List[str]]:
    rows = []
    for report in reports:
        performance = report.data.get("Performance") if isinstance(report.data.get("Performance"), dict) else {}
        unhealthy, unknown, malformed = disk_health_counts(report)
        lowest, malformed_logical = lowest_free_percent(report)
        if malformed or malformed_logical:
            report.warnings.append(f"{malformed + malformed_logical} malformed disk entries ignored")
        memory_suffix = " (commit)" if report.memory_metric == "commit" else ""
        completeness = nested(report.data, "Completeness", "Status", default="n/a (schema " + report.schema + ")")
        rows.append([
            report.time_text(),
            str(report.data.get("HealthScore", "N/A")),
            str(completeness),
            fmt_value(performance.get("CPU_AveragePercent"), "%"),
            fmt_value(performance.get("CPU_P95Percent"), "%"),
            fmt_value(performance.get("Memory_AveragePercent"), "%") + memory_suffix,
            fmt_value(performance.get("Memory_P95Percent"), "%") + memory_suffix,
            fmt_value(performance.get("DiskActive_P95"), "%"),
            fmt_value(lowest, "%"),
            fmt_value(nested(report.data, "Battery", "HealthPercent"), "%"),
            f"{unhealthy} unhealthy / {unknown} unknown",
        ])
    return rows


def build_device_section(reports: List[LoadedReport]) -> List[str]:
    first, latest = reports[0], reports[-1]
    lines = [
        f"**Audits compared:** {len(reports)}  ",
        f"**Period:** {first.time_text()} to {latest.time_text()}  ",
        f"**Latest result:** {latest.data.get('HealthScore', 'N/A')}/100 - {latest.data.get('Rating', 'N/A')}"
        + (f" ({nested(latest.data, 'Completeness', 'Status')} evidence)" if nested(latest.data, "Completeness", "Status") else ""),
        "",
        "### Trend table",
        "",
        markdown_table(["Audit time", "Score", "Evidence", "CPU avg", "CPU P95", "RAM avg", "RAM P95", "Disk P95", "Lowest free", "Battery", "Disks"], metric_rows(reports)),
        "",
        "### Material changes (first vs latest)",
        "",
    ]
    changes = [
        trend_statement("Triage score", first.data.get("HealthScore"), latest.data.get("HealthScore"), " points", adverse_when_higher=False),
        trend_statement("Average CPU", nested(first.data, "Performance", "CPU_AveragePercent"), nested(latest.data, "Performance", "CPU_AveragePercent"), "%"),
        trend_statement("Battery health", nested(first.data, "Battery", "HealthPercent"), nested(latest.data, "Battery", "HealthPercent"), "%", adverse_when_higher=False),
    ]
    if first.memory_metric == latest.memory_metric:
        changes.append(trend_statement("P95 memory pressure" + (" (commit charge)" if first.memory_metric == "commit" else " (physical RAM)"), nested(first.data, "Performance", "Memory_P95Percent"), nested(latest.data, "Performance", "Memory_P95Percent"), "%"))
    else:
        lines.append("- Memory not compared: schema 1.0 reports commit charge, later schemas report physical RAM in use. The values are not the same measurement.")
    lines.extend(change for change in changes if change)
    if not any(changes):
        lines.append("- No material change was measurable across the selected snapshots.")
    lines.extend(["", "### Latest findings", ""])
    for finding in as_list(latest.data.get("Findings")):
        if not isinstance(finding, dict):
            latest.warnings.append("a malformed finding entry was ignored")
            continue
        lines.append(
            f"- **{escape_markdown(finding.get('Severity', 'Unknown'))} - {escape_markdown(finding.get('Category', 'General'))}:** "
            f"{escape_markdown(finding.get('Finding', ''))} Evidence: {escape_markdown(finding.get('Evidence', ''))} "
            f"Next step: {escape_markdown(finding.get('Recommendation', ''))}"
        )
    process_rows = recurring_processes(reports)
    lines.extend(["", "### Recurrent heavy processes", ""])
    if process_rows:
        lines.append(markdown_table(["Process", "Audits present", "Avg measured CPU", "Peak observed RAM"], process_rows))
        lines.extend(["", "Entries with the same application name are summed within each audit. Older audits kept only their top entries, so their totals are lower bounds. Presence is evidence for IT correlation, not permission to end, remove or disable anything."])
    else:
        lines.append("No process recurred in the available top-process lists.")
    warnings = [f"- {report.path.name if report.path else 'report'} ({report.time_text()}): {warning}" for report in reports for warning in report.warnings]
    if warnings:
        lines.extend(["", "### Data quality notes", ""])
        lines.extend(sorted(set(warnings)))
    return lines


def build_comparison(reports: List[LoadedReport], errors: List[str]) -> str:
    groups: Dict[str, List[LoadedReport]] = defaultdict(list)
    for report in reports:
        groups[report.device_key].append(report)
    lines = ["# PC Health Comparison", ""]
    if len(groups) > 1:
        lines.append(f"**Warning:** the selected audits come from {len(groups)} different devices. They are compared separately below; do not read trends across devices.")
        lines.append("")
    for key, device_reports in groups.items():
        device_reports.sort(key=lambda report: report.sort_key)
        label = device_reports[-1].device_label
        lines.append(f"## Device: {escape_markdown(label)} (id {key})")
        lines.append("")
        lines.extend(build_device_section(device_reports))
        lines.append("")
    if errors:
        lines.extend(["## Files not analysed", ""])
        lines.extend(f"- {escape_markdown(error)}" for error in errors)
        lines.append("")
    lines.extend([
        "## Interpretation limits",
        "",
        "Compare like with like: idle after restart, normal work, and an active slowdown, using the workload label recorded in each audit. Incomplete evidence is shown in the Evidence column. This report does not authorise changes to a managed device.",
        "",
    ])
    return "\n".join(lines)


def compare_files(inputs: Sequence[str], output_path: str) -> Dict[str, Any]:
    paths = discover_inputs(inputs)
    reports, errors = load_reports(paths)
    if not reports:
        detail = "\n".join(errors) if errors else "No PC_Health_Data.json files found."
        raise RuntimeError(detail)
    text = build_comparison(reports, errors)
    output = Path(output_path).expanduser()
    output.write_text(text, encoding="utf-8")
    return {"Output": str(output.resolve()), "Reports": len(reports), "Errors": errors, "Devices": len({r.device_key for r in reports})}


# ---------------------------------------------------------------------------
# Window (tkinter). Imported lazily so the command line works without it.
# ---------------------------------------------------------------------------

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except Exception:  # pragma: no cover - depends on the Python build
    tk = None  # type: ignore


def open_path(path: str) -> None:
    if not path or not os.path.exists(path):
        raise FileNotFoundError(path)
    if IS_WINDOWS:
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def is_elevated() -> Optional[bool]:
    if not IS_WINDOWS:
        return None
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return None


def relaunch_as_administrator() -> Tuple[bool, str]:
    if not IS_WINDOWS:
        return False, "Elevation is a Windows feature."
    script = os.path.abspath(__file__)
    executable = sys.executable
    try:
        result = ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, f'"{script}"', None, 1)
    except Exception as exc:
        return False, str(exc)
    if int(result) <= 32:
        return False, f"Windows refused the elevation request (code {int(result)}). The UAC prompt may have been declined or blocked by policy."
    return True, "Elevated window requested."


def load_records_file(path: str) -> List[Dict[str, Any]]:
    records = []
    with open(path, "r", encoding="utf-8-sig") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
            except ValueError:
                record = {"section": "noise", "status": "ok", "text": line[:500]}
            if isinstance(record, dict):
                records.append(record)
    return records


def replay_records(path: str, options: AuditOptions) -> Dict[str, Any]:
    records = load_records_file(path)
    if not records:
        raise RuntimeError("The records file is empty or unreadable.")
    timestamp = now_local().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(options.output_dir, f"PC_Health_Audit_replay_{timestamp}")
    report = analyse_records(records, options, {"Started": iso(now_local()), "Finished": iso(now_local()), "ExitCode": 0})
    report["CollectionNotes"].insert(0, f"Re-analysed from records file {os.path.basename(path)}; the audit time above is the analysis time.")
    paths = write_reports(report, folder)
    return {"Report": report, "Folder": folder, "Paths": paths, "RunMeta": {}}


class App:
    """Main window. All widget access happens on the tkinter thread."""

    POLL_MS = 120

    def __init__(self, root: "tk.Tk") -> None:
        self.root = root
        self.queue: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.cancel_event = threading.Event()
        self.last_result: Optional[Dict[str, Any]] = None
        self.compare_result: Optional[Dict[str, Any]] = None
        self.sample_edited = False
        root.title(f"{TOOL_NAME} {TOOL_VERSION}")
        root.minsize(860, 640)
        root.geometry("960x720")
        root.report_callback_exception = self.report_exception  # type: ignore[assignment]
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        style = ttk.Style()
        try:
            if "vista" in style.theme_names():
                style.theme_use("vista")
        except tk.TclError:
            pass
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.audit_tab = ttk.Frame(self.notebook, padding=10)
        self.compare_tab = ttk.Frame(self.notebook, padding=10)
        self.about_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.audit_tab, text="Run audit")
        self.notebook.add(self.compare_tab, text="Compare audits")
        self.notebook.add(self.about_tab, text="About and privacy")
        self.build_audit_tab()
        self.build_compare_tab()
        self.build_about_tab()
        self.root.after(self.POLL_MS, self.poll_queue)

    # ----- error handling -------------------------------------------------
    def report_exception(self, exc_type: Any, exc_value: Any, exc_tb: Any) -> None:
        text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        try:
            with open(crash_log_path(), "a", encoding="utf-8") as handle:
                handle.write(f"\n---- {now_local().isoformat()} ----\n{text}\n")
        except OSError:
            pass
        messagebox.showerror(f"{TOOL_NAME} - unexpected error", f"{exc_value}\n\nDetails were written to:\n{crash_log_path()}")

    # ----- audit tab -------------------------------------------------------
    def build_audit_tab(self) -> None:
        frame = self.audit_tab
        frame.columnconfigure(1, weight=1)
        banner = ttk.Label(frame, text="Read-only, offline and privacy-conscious. On a company-managed device, obtain IT approval before running.", wraplength=880)
        banner.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="Mode").grid(row=1, column=0, sticky="w")
        mode_frame = ttk.Frame(frame)
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w")
        self.mode_var = tk.StringVar(value="Standard")
        for mode in MODES:
            ttk.Radiobutton(mode_frame, text=mode, value=mode, variable=self.mode_var, command=self.on_mode_change).pack(side="left", padx=(0, 12))
        self.mode_hint = ttk.Label(frame, text="", foreground="#555555", wraplength=700)
        self.mode_hint.grid(row=2, column=1, columnspan=2, sticky="w")

        ttk.Label(frame, text="Samples (one per second nominal)").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.sample_var = tk.StringVar(value=str(DEFAULT_SAMPLES["Standard"]))
        sample_box = ttk.Spinbox(frame, from_=SAMPLE_MIN, to=SAMPLE_MAX, textvariable=self.sample_var, width=8, command=lambda: setattr(self, "sample_edited", True))
        sample_box.grid(row=3, column=1, sticky="w", pady=(6, 0))
        sample_box.bind("<KeyRelease>", lambda _e: setattr(self, "sample_edited", True))

        ttk.Label(frame, text="Output folder").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.output_var = tk.StringVar(value=default_output_dir())
        ttk.Entry(frame, textvariable=self.output_var).grid(row=4, column=1, sticky="ew", pady=(6, 0))
        ttk.Button(frame, text="Browse...", command=self.browse_output).grid(row=4, column=2, sticky="w", padx=(6, 0), pady=(6, 0))

        ttk.Label(frame, text="Workload during the sample").grid(row=5, column=0, sticky="w", pady=(6, 0))
        self.workload_var = tk.StringVar(value=WORKLOAD_LABELS[1])
        ttk.Combobox(frame, textvariable=self.workload_var, values=WORKLOAD_LABELS, state="readonly").grid(row=5, column=1, sticky="ew", pady=(6, 0))
        ttk.Label(frame, text="Notes (optional)").grid(row=6, column=0, sticky="w", pady=(6, 0))
        self.notes_var = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.notes_var).grid(row=6, column=1, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(frame, text="Dock connected").grid(row=7, column=0, sticky="w", pady=(6, 0))
        self.dock_var = tk.StringVar(value="Not stated")
        ttk.Combobox(frame, textvariable=self.dock_var, values=("Not stated", "Yes", "No"), state="readonly", width=12).grid(row=7, column=1, sticky="w", pady=(6, 0))

        options = ttk.Frame(frame)
        options.grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.sensitive_var = tk.BooleanVar(value=False)
        self.apps_var = tk.BooleanVar(value=False)
        self.zip_var = tk.BooleanVar(value=False)
        self.open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options, text="Include sensitive data (names, serials, SSID, startup commands)", variable=self.sensitive_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options, text="Include installed applications", variable=self.apps_var).grid(row=0, column=1, sticky="w", padx=(16, 0))
        ttk.Checkbutton(options, text="Create ZIP", variable=self.zip_var).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(options, text="Open the HTML report when done", variable=self.open_var).grid(row=1, column=1, sticky="w", padx=(16, 0))

        elevation = is_elevated()
        elevation_text = {True: "Running elevated (administrator): all checks available.", False: "Not elevated: Secure Boot, TPM, BitLocker, thermal zones, SSD wear and Deep integrity checks will be reported as 'requires elevation'.", None: "Elevation state unknown on this platform."}[elevation]
        elevation_row = ttk.Frame(frame)
        elevation_row.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(elevation_row, text=elevation_text, wraplength=700).pack(side="left")
        if elevation is False:
            ttk.Button(elevation_row, text="Restart as administrator", command=self.restart_elevated).pack(side="right")

        buttons = ttk.Frame(frame)
        buttons.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.run_button = ttk.Button(buttons, text="Run audit", command=self.start_audit)
        self.run_button.pack(side="left")
        self.cancel_button = ttk.Button(buttons, text="Cancel", command=self.cancel_audit, state="disabled")
        self.cancel_button.pack(side="left", padx=(6, 0))
        self.open_folder_button = ttk.Button(buttons, text="Open report folder", command=self.open_folder, state="disabled")
        self.open_folder_button.pack(side="left", padx=(18, 0))
        self.open_html_button = ttk.Button(buttons, text="Open HTML report", command=self.open_html, state="disabled")
        self.open_html_button.pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="Re-analyse a records file...", command=self.replay_file).pack(side="right")

        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1000)
        self.progress.grid(row=11, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(frame, textvariable=self.status_var, wraplength=880).grid(row=12, column=0, columnspan=3, sticky="w")
        self.log = scrolledtext.ScrolledText(frame, height=14, wrap="word", state="disabled", font=("Consolas", 9) if IS_WINDOWS else None)
        self.log.grid(row=13, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        frame.rowconfigure(13, weight=1)
        if not IS_WINDOWS:
            self.run_button.configure(state="disabled")
            self.status_var.set("The audit runs on Windows only. Comparison and re-analysis work on this platform.")
        self.on_mode_change()

    def on_mode_change(self) -> None:
        mode = self.mode_var.get()
        hints = {
            "Quick": "About 1 minute: 10 samples, 3 days of events, no reliability records.",
            "Standard": "Recommended: 30 samples (2 to 5 minutes), 14 days of events, reliability records.",
            "Deep": "60 samples, 30 days of events, plus DISM /CheckHealth and SFC /verifyonly (verification only; needs administrator; SFC can take 5 to 15 minutes).",
        }
        self.mode_hint.configure(text=hints.get(mode, ""))
        if not self.sample_edited:
            self.sample_var.set(str(DEFAULT_SAMPLES.get(mode, 30)))

    def browse_output(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_var.get() or default_output_dir(), title="Choose the output folder")
        if chosen:
            self.output_var.set(chosen)

    def restart_elevated(self) -> None:
        if not messagebox.askyesno("Restart as administrator", "Windows will show a User Account Control prompt. The elevated window collects the same read-only evidence plus the elevation-only checks. Continue?"):
            return
        ok, message = relaunch_as_administrator()
        if ok:
            self.root.destroy()
        else:
            messagebox.showwarning("Elevation not started", message)

    def append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"[{now_local().strftime('%H:%M:%S')}] {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def collect_options(self) -> AuditOptions:
        try:
            samples = int(self.sample_var.get())
        except ValueError:
            samples = DEFAULT_SAMPLES.get(self.mode_var.get(), 30)
        dock = {"Yes": True, "No": False}.get(self.dock_var.get())
        return AuditOptions(
            mode=self.mode_var.get(),
            sample_count=samples,
            output_dir=self.output_var.get().strip() or default_output_dir(),
            include_sensitive=self.sensitive_var.get(),
            include_apps=self.apps_var.get(),
            create_zip=self.zip_var.get(),
            open_report=self.open_var.get(),
            workload_label=self.workload_var.get(),
            workload_notes=self.notes_var.get(),
            docked=dock,
        )

    def start_audit(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        options = self.collect_options()
        if not os.path.isdir(options.output_dir):
            try:
                os.makedirs(options.output_dir, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Output folder", f"Cannot use the output folder:\n{exc}")
                return
        if options.mode == "Deep" and is_elevated() is False:
            if not messagebox.askyesno("Deep mode without elevation", "Deep mode's DISM and SFC checks need administrator rights and will be reported as 'requires elevation'. Run anyway?"):
                return
        self.sample_var.set(str(options.sample_count))
        self.cancel_event.clear()
        self.last_result = None
        self.set_running(True)
        self.progress["value"] = 0
        self.append_log(f"Starting {options.mode} audit with {options.sample_count} samples; workload: {options.workload_label}")
        self.worker = threading.Thread(target=self.audit_worker, args=(options,), daemon=True)
        self.worker.start()

    def audit_worker(self, options: AuditOptions) -> None:
        def progress(message: str, fraction: Optional[float] = None) -> None:
            self.queue.put(("progress", (message, fraction)))
        try:
            result = run_audit(options, progress, self.cancel_event)
            self.queue.put(("done", result))
        except Exception as exc:
            self.queue.put(("error", f"{exc}\n\n{traceback.format_exc()}"))

    def cancel_audit(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
            self.append_log("Cancel requested; the collector is being stopped and a partial report will be written.")
            self.cancel_button.configure(state="disabled")

    def set_running(self, running: bool) -> None:
        self.run_button.configure(state="disabled" if running or not IS_WINDOWS else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")

    def poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    message, fraction = payload
                    if message:
                        self.append_log(message)
                        self.status_var.set(message)
                    if fraction is not None:
                        self.progress["value"] = max(0, min(1000, int(fraction * 1000)))
                elif kind == "done":
                    self.finish_audit(payload)
                elif kind == "error":
                    self.set_running(False)
                    self.append_log("Audit failed: " + str(payload).splitlines()[0])
                    self.status_var.set("Audit failed. See the message and the crash log.")
                    messagebox.showerror("Audit failed", str(payload)[:2500])
                elif kind == "compare_done":
                    self.finish_compare(payload)
                elif kind == "compare_error":
                    self.compare_button.configure(state="normal")
                    self.compare_log_add("Comparison failed: " + str(payload))
                    messagebox.showerror("Comparison failed", str(payload)[:2500])
        except queue.Empty:
            pass
        self.root.after(self.POLL_MS, self.poll_queue)

    def finish_audit(self, result: Dict[str, Any]) -> None:
        self.set_running(False)
        self.last_result = result
        report = result["Report"]
        self.progress["value"] = 1000
        self.open_folder_button.configure(state="normal")
        self.open_html_button.configure(state="normal")
        self.append_log(f"Result: {report['Headline']}")
        self.append_log(f"Evidence: {report['Completeness']['Status']}; missing: {', '.join(report['Completeness']['MissingChecks']) or 'none'}")
        for finding in report["Findings"]:
            if finding["Severity"] in ("Critical", "Warning"):
                self.append_log(f"[{finding['Severity']}] {finding['Category']}: {finding['Finding']}")
        self.append_log("Reports: " + "; ".join(result["Paths"].values()))
        self.status_var.set(f"Done: {report['Headline']}. Reports in {result['Folder']}")
        if result["RunMeta"].get("Cancelled"):
            self.append_log("The run was cancelled; the report covers the evidence collected before cancellation.")
        if self.open_var.get():
            try:
                open_path(result["Paths"]["html"])
            except Exception as exc:
                self.append_log(f"Could not open the HTML report automatically: {exc}")

    def open_folder(self) -> None:
        if self.last_result:
            try:
                open_path(self.last_result["Folder"])
            except Exception as exc:
                messagebox.showwarning("Open folder", str(exc))

    def open_html(self) -> None:
        if self.last_result:
            try:
                open_path(self.last_result["Paths"]["html"])
            except Exception as exc:
                messagebox.showwarning("Open report", str(exc))

    def replay_file(self) -> None:
        path = filedialog.askopenfilename(title="Choose a PC_Health_Records.jsonl file", filetypes=[("Records", "*.jsonl"), ("All files", "*.*")])
        if not path:
            return
        options = self.collect_options()
        try:
            result = replay_records(path, options)
        except Exception as exc:
            messagebox.showerror("Re-analysis failed", str(exc))
            return
        self.finish_audit(result)

    # ----- compare tab -----------------------------------------------------
    def build_compare_tab(self) -> None:
        frame = self.compare_tab
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="Select audit folders (searched for PC_Health_Data.json) or JSON files from schema 1.0, 1.1 or 1.2. Audits from different devices are reported separately.", wraplength=880).grid(row=0, column=0, columnspan=2, sticky="w")
        self.inputs = tk.Listbox(frame, height=8, selectmode="extended")
        self.inputs.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        side = ttk.Frame(frame)
        side.grid(row=1, column=1, sticky="n", padx=(6, 0), pady=(6, 0))
        ttk.Button(side, text="Add folder...", command=self.add_folder).pack(fill="x")
        ttk.Button(side, text="Add files...", command=self.add_files).pack(fill="x", pady=(4, 0))
        ttk.Button(side, text="Remove selected", command=self.remove_selected).pack(fill="x", pady=(4, 0))
        ttk.Button(side, text="Clear", command=lambda: self.inputs.delete(0, "end")).pack(fill="x", pady=(4, 0))
        ttk.Label(frame, text="Output Markdown file").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.compare_output_var = tk.StringVar(value=os.path.join(default_output_dir(), "PC_Health_Comparison.md"))
        row = ttk.Frame(frame)
        row.grid(row=3, column=0, columnspan=2, sticky="ew")
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.compare_output_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="Browse...", command=self.browse_compare_output).grid(row=0, column=1, padx=(6, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.compare_button = ttk.Button(buttons, text="Run comparison", command=self.start_compare)
        self.compare_button.pack(side="left")
        self.open_compare_button = ttk.Button(buttons, text="Open result", command=self.open_compare_result, state="disabled")
        self.open_compare_button.pack(side="left", padx=(6, 0))
        self.compare_log = scrolledtext.ScrolledText(frame, height=14, wrap="word", state="disabled", font=("Consolas", 9) if IS_WINDOWS else None)
        self.compare_log.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        frame.rowconfigure(5, weight=1)

    def compare_log_add(self, text: str) -> None:
        self.compare_log.configure(state="normal")
        self.compare_log.insert("end", text + "\n")
        self.compare_log.see("end")
        self.compare_log.configure(state="disabled")

    def add_folder(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_var.get() or default_output_dir(), title="Choose a folder containing audits")
        if chosen:
            self.inputs.insert("end", chosen)

    def add_files(self) -> None:
        chosen = filedialog.askopenfilenames(title="Choose PC_Health_Data.json files", filetypes=[("JSON", "*.json"), ("All files", "*.*")])
        for path in chosen:
            self.inputs.insert("end", path)

    def remove_selected(self) -> None:
        for index in reversed(self.inputs.curselection()):
            self.inputs.delete(index)

    def browse_compare_output(self) -> None:
        chosen = filedialog.asksaveasfilename(title="Save the comparison as", defaultextension=".md", initialfile="PC_Health_Comparison.md", filetypes=[("Markdown", "*.md"), ("All files", "*.*")])
        if chosen:
            self.compare_output_var.set(chosen)

    def start_compare(self) -> None:
        inputs = list(self.inputs.get(0, "end"))
        if not inputs:
            messagebox.showinfo("Compare audits", "Add at least one folder or JSON file first.")
            return
        output = self.compare_output_var.get().strip()
        if not output:
            messagebox.showinfo("Compare audits", "Choose an output file.")
            return
        self.compare_button.configure(state="disabled")
        self.compare_log_add(f"Comparing {len(inputs)} input(s)...")

        def worker() -> None:
            try:
                self.queue.put(("compare_done", compare_files(inputs, output)))
            except Exception as exc:
                self.queue.put(("compare_error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def finish_compare(self, result: Dict[str, Any]) -> None:
        self.compare_button.configure(state="normal")
        self.compare_result = result
        self.open_compare_button.configure(state="normal")
        self.compare_log_add(f"Compared {result['Reports']} audit(s) from {result['Devices']} device(s). Report: {result['Output']}")
        for error in result.get("Errors", []):
            self.compare_log_add("Skipped: " + error)

    def open_compare_result(self) -> None:
        if self.compare_result:
            try:
                open_path(self.compare_result["Output"])
            except Exception as exc:
                messagebox.showwarning("Open result", str(exc))

    # ----- about tab -------------------------------------------------------
    def build_about_tab(self) -> None:
        text = scrolledtext.ScrolledText(self.about_tab, wrap="word", state="normal")
        text.pack(fill="both", expand=True)
        text.insert("end", ABOUT_TEXT)
        text.configure(state="disabled")

    def on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Audit running", "An audit is still running. Stop it and close?"):
                return
            self.cancel_event.set()
        self.root.destroy()


ABOUT_TEXT = f"""{TOOL_NAME} {TOOL_VERSION}

WHAT IT DOES
Collects hardware specifications, sampled performance (CPU total and per core, processor frequency vs nominal, physical RAM, paging, per-disk activity and latency, top applications), battery wear, storage health, security posture, device faults and recent event-log evidence. Everything is analysed in this Python file; a hidden Windows PowerShell 5.1 process only reads counters and streams them back as JSON lines.

READ-ONLY
No configuration change is made. Deep mode runs DISM /CheckHealth and SFC /verifyonly, which verify only. The tool writes the report folder you choose, and briefly writes a collector script and a battery XML under %TEMP%, both deleted afterwards (any deletion failure is listed in the coverage table). No network connection is used.

PRIVACY
Computer name, user name, serial numbers, volume labels, Wi-Fi SSID, MAC addresses, startup commands and event message text are redacted unless 'Include sensitive data' is ticked. Review every report before sharing.

EVIDENCE, NOT VERDICTS
Every check is recorded as collected, partially collected, unavailable, requires elevation, failed, not applicable or not run. A missing check makes the evidence 'Incomplete' and can never produce a clean result. The headline respects the most serious finding. Thresholds are triage heuristics.

ELEVATION
Secure Boot, TPM, BitLocker, ACPI thermal zones, SSD wear/temperature and the Deep integrity checks need an administrator PowerShell. Use 'Restart as administrator' if company policy permits.

COMPANY-MANAGED DEVICES
Obtain IT approval before running. Managed security, sync and update agents may be mandatory; nothing in the report authorises ending, removing or disabling them.

COMPARISON
The Compare tab reads audits from schema 1.0, 1.1 and 1.2, orders them by timezone-aware audit time, groups them by device, sums process entries correctly and refuses to compare memory figures whose definitions differ.

FILES
PC_Health_Report.html (readable report), PC_Health_Data.json (schema {SCHEMA_VERSION}), PC_Health_Summary.txt (executive summary), PC_Health_Records.jsonl (raw collector records for re-analysis).

Crash details, if any, are written to {crash_log_path()}.
"""


def launch_gui() -> int:
    if tk is None:
        show_fatal_error(
            "tkinter is not available in this Python installation, so the window cannot open.\n\n"
            "Reinstall Python from python.org with the 'tcl/tk and IDLE' option ticked, or use the command line:\n"
            "  py -3 PC_Health_Audit.pyw audit\n  py -3 PC_Health_Audit.pyw compare <folder>"
        )
        return 2
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="PC_Health_Audit.pyw", description="PC Health Audit. Without arguments the window opens.")
    sub = parser.add_subparsers(dest="command")
    audit = sub.add_parser("audit", help="run an audit without the window")
    audit.add_argument("--mode", choices=MODES, default="Standard")
    audit.add_argument("--samples", type=int, default=None, help=f"sample count ({SAMPLE_MIN}-{SAMPLE_MAX})")
    audit.add_argument("--output", default=None, help="output folder (default: Desktop)")
    audit.add_argument("--workload", default="Not stated", help="workload label, for example 'Large Excel workbook'")
    audit.add_argument("--notes", default="")
    audit.add_argument("--docked", choices=("yes", "no"), default=None)
    audit.add_argument("--sensitive", action="store_true")
    audit.add_argument("--apps", action="store_true")
    audit.add_argument("--zip", action="store_true")
    audit.add_argument("--open", action="store_true")
    compare = sub.add_parser("compare", help="compare audit JSON files or folders")
    compare.add_argument("inputs", nargs="+")
    compare.add_argument("-o", "--output", default="PC_Health_Comparison.md")
    replay = sub.add_parser("replay", help="re-analyse a PC_Health_Records.jsonl file")
    replay.add_argument("records")
    replay.add_argument("--output", default=None)
    replay.add_argument("--mode", choices=MODES, default="Standard")
    replay.add_argument("--workload", default="Not stated")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        return launch_gui()
    parsed = build_parser().parse_args(args)
    if parsed.command == "audit":
        options = AuditOptions(
            mode=parsed.mode, sample_count=parsed.samples, output_dir=parsed.output, include_sensitive=parsed.sensitive,
            include_apps=parsed.apps, create_zip=parsed.zip, open_report=parsed.open, workload_label=parsed.workload,
            workload_notes=parsed.notes, docked={"yes": True, "no": False}.get(parsed.docked),
        )
        result = run_audit(options, lambda message, fraction=None: print(message, flush=True))
        print("HTML report: " + result["Paths"]["html"])
        print("Structured data: " + result["Paths"]["json"])
        if parsed.open:
            open_path(result["Paths"]["html"])
        return 0
    if parsed.command == "compare":
        result = compare_files(parsed.inputs, parsed.output)
        print(f"Compared {result['Reports']} audit(s) from {result['Devices']} device(s). Report created: {result['Output']}")
        for error in result["Errors"]:
            print("Skipped: " + error)
        return 0
    if parsed.command == "replay":
        options = AuditOptions(mode=parsed.mode, output_dir=parsed.output, workload_label=parsed.workload)
        result = replay_records(parsed.records, options)
        print(f"{result['Report']['Headline']}\nHTML report: {result['Paths']['html']}")
        return 0
    return launch_gui()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        show_fatal_error(traceback.format_exc())
        sys.exit(1)
