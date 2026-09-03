# Review Brief - PC Health Audit v1.2 (single .pyw)

Author: Alaa Elsayed
Date: 03-Sep-2026
File: `PC_Health_Audit.pyw` (schema 1.2). Replaces `PC_Health_Audit.ps1` and `compare_pc_health.py`.

## Purpose

One double-clickable file with a window. The v1.1 `.py` closed at once because it was a command-line tool with no arguments; the v1.1 `.ps1` stopped in Windows PowerShell 5.1 with "Argument types do not match" at `return @($results)`. Both are replaced by a tkinter window that runs the audit, shows progress, writes the reports and compares earlier audits.

## Root cause of the v1.1 PowerShell failure

`Get-EventSummary` built a `System.Collections.Generic.List[object]` and returned it through the array sub-expression operator. Windows PowerShell 5.1 raised `ArgumentException: Argument types do not match` at that statement. This is consistent with the known 5.1 binder failures involving generic collections (the same family as the reported ConvertTo-Json failure). The v1.2 collector removes the pattern entirely: no function returns a collection; each check emits one JSON line; only plain PowerShell arrays are used.

## Changes against the v1.1 review findings

| # | Review finding | v1.2 change | Where |
|---|---|---|---|
| 1 | Potential 5.1 runtime failure (`@($list)` on a generic List) | Collector streams JSON per check; no generic lists; no collection returns; Python does all assembly | `POWERSHELL_COLLECTOR`, `Invoke-Section` |
| 2 | Missing evidence could read as 100/100 Healthy | Every check has a status (collected, partial, unavailable, requires elevation, failed, not applicable, not run). Coverage table in HTML, JSON and text. Rating is bounded by the worst finding; missing diagnostic evidence gives "Inconclusive (evidence incomplete)"; elevation-only gaps are named in the headline and never yield a plain "Healthy" without disclosure | `build_coverage`, `classify_coverage`, `score_and_rating` |
| 3 | Throttling rule combined load and frequency from different moments | Load and processor performance are compared within the same timestamped sample. Finding requires at least max(3, 10 percent of paired samples) overlapping samples; zero or absent performance counters are excluded; overlap mostly on battery downgrades to Information; Kernel-Processor-Power event 37 is used as corroboration | `analyse_performance`, `evaluate_findings` |
| 4 | Analyser compared 1.0 commit memory with 1.1 physical memory | Schema validation; `memory_metric` per report; the memory trend is skipped with a warning when metrics differ; commit values labelled in the table | `LoadedReport`, `build_device_section` |
| 5 | `MaxCapacityEx` treated as bytes | Kilobytes multiplied by 1024; value labelled provisional; flagged unreliable if lower than installed RAM | `analyse_records` |
| 6 | Deep-check results absent from the report | Exit codes captured; verdicts classified (no corruption, repairable, not repairable, command failed, blocked, inconclusive); output shown in the HTML; findings raised. SFC UTF-16 output decoded and nulls stripped | `classify_dism`, `classify_sfc` |
| 7 | Hardware events missed by the 25-group display cap and by excluding warnings | Hardware providers queried separately at levels 1-3 (cap 1000, disclosed); diagnosis runs over every group before any display cap; counts for WHEA, storage, Kernel-Power 41, Kernel-Processor-Power 37, display driver resets, Kernel-PnP 219 | `event_diagnostics` |
| 8 | `PagesInputPersec` mislabelled as hard faults | Renamed `PagesInputPerSec`; defined in `MetricDefinitions`; warning only with low available RAM, otherwise Information | `METRIC_DEFINITIONS`, `evaluate_findings` |
| 9 | Analyser: legacy aggregation averaged instead of summed; mixed laptops; text-sorted timestamps; null disk health counted unhealthy; malformed entries crashed | Entries with the same base name are summed within an audit; audits grouped by device fingerprint with a warning; timezone-aware ordering with `/Date(ms)/` and naive handling; unknown disk health counted separately; malformed entries skipped and reported | comparison functions |
| 10 | Per-core CPU, per-disk activity and latency, timestamps and workload labels, GPU memory, coverage table, `SampleCount`, temp-file failures | All added: per-core table; per-disk table with latency from raw `AvgDisksec*` counters between consecutive samples; sample timeline; workload label, notes, dock and measured power source; GPU memory from the registry `qwMemorySize` with the 32-bit `AdapterRAM` labelled when used; coverage table; `SampleCount`; temp-file deletion results in coverage | throughout |

## Not changed or still open

- Windows PowerShell 5.1 has not been executed in the development environment. The collector parses under PowerShell 7 and runs end to end with mocked cmdlets; the Windows smoke test in the README is mandatory.
- `Win32_PerfFormattedData_Counters_ProcessorInformation` availability on the target build remains to be confirmed on the laptop; the report states the valid sample count and says "could not be evaluated" when it is absent.
- Thresholds remain triage heuristics.
- Disk latency needs at least two samples per disk; the report shows the number of intervals used.

## Verification on Windows (mandatory)

1. Double-click `PC_Health_Audit.pyw`; the window must open. If not, run `py -3 PC_Health_Audit.pyw` from a console and read the error.
2. Quick mode, 5 samples, mains power, standard user. Confirm the coverage table statuses, `Memory_AveragePercent` against Task Manager, non-null `CPU_Performance_*`, battery capacity source, msedge or Teams listed once with Instances > 1, and no `PC_Health_Audit_*` folder left in `%TEMP%`.
3. Run as administrator once, Deep mode, only if IT permits: confirm DISM and SFC verdicts and exit codes appear.
4. Compare tab against the Desktop: confirm one device section and a data-quality note for any schema 1.0 audit.

## Review focus

- Windows PowerShell 5.1 behaviour of `ConvertTo-Json -Compress -Depth 12` on nested hashtables and of `[Console]::OpenStandardOutput()` with redirected output.
- `Get-WinEvent -FilterHashtable` with `Level = @(1, 2, 3)` and a `ProviderName` array on the target build.
- Correctness of the raw-counter latency formula `((N1 - N0) / F) / (D1 - D0)`.
