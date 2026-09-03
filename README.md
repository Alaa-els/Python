# PC Health Audit

Single-file, read-only Windows laptop diagnostic with a window. Double-click
`PC_Health_Audit.pyw` to open it. It collects hardware, performance, battery,
storage, security and event-log evidence, scores it, and writes an HTML report,
a JSON data file and a text summary. The same window compares earlier audits.

Company-managed device? Obtain IT approval before running it.

## Requirements

- Windows 10 or 11 with Windows PowerShell 5.1 (built in).
- Python 3.9 or newer from python.org with the "tcl/tk and IDLE" option ticked
  (the default). No third-party packages.
- Standard user is enough. Secure Boot, TPM, BitLocker, ACPI thermal zones, SSD
  wear counters and the Deep integrity checks need "Restart as administrator".

## Use

1. Double-click `PC_Health_Audit.pyw`. If nothing opens, run
   `py -3 PC_Health_Audit.pyw` from a Command Prompt to see the error, and check
   `%TEMP%\PC_Health_Audit_crash.log`.
2. Choose a mode (Quick 10 samples, Standard 30, Deep 60 plus DISM and SFC
   verification), set the workload label (for example "Large Excel workbook" or
   "Civil 3D") and whether a dock is connected, then click Run audit.
3. Read the HTML report. The first table is the findings; the second is the
   coverage table showing what was collected, what needs elevation and what
   failed. Missing evidence is reported as missing, never as a clean result.
4. Compare audits: on the Compare tab add the Desktop (or the audit folders)
   and click Run comparison. Audits from different devices are reported
   separately; schema 1.0 memory figures (commit charge) are not compared with
   later physical-RAM figures.

Command line (optional):

```
py -3 PC_Health_Audit.pyw audit --mode Standard --samples 30 --workload "Large Excel workbook"
py -3 PC_Health_Audit.pyw compare "%USERPROFILE%\Desktop" -o PC_Health_Comparison.md
py -3 PC_Health_Audit.pyw replay PC_Health_Records.jsonl
```

## What it writes

`<output>\PC_Health_Audit_<timestamp>\` containing `PC_Health_Report.html`,
`PC_Health_Data.json` (schema 1.2), `PC_Health_Summary.txt` and
`PC_Health_Records.jsonl` (raw collector records for re-analysis). Optional ZIP
next to the folder. A collector script and a battery XML are written under
`%TEMP%` for the duration of the run and deleted; any deletion failure appears
in the coverage table. Nothing else is written and no setting is changed.

## Privacy

Computer name, user name, serial numbers, volume labels, Wi-Fi SSID, MAC
addresses, startup commands and event message text are redacted unless
"Include sensitive data" is ticked. Review every report before sharing.

## How it works

Python starts one hidden Windows PowerShell 5.1 process that only reads
counters and streams one JSON line per check and per performance sample.
Every check records a status: collected, partially collected, unavailable,
requires elevation, failed, not applicable or not run. Python does all the
analysis: timestamped throttling evaluation (load and frequency in the same
sample), per-core CPU, per-disk latency from raw counters, physical RAM versus
commit, paging correlated with available RAM, hardware event diagnosis over
every collected group (including warnings such as Kernel-Processor-Power 37),
deep-check classification with exit codes, and a score whose rating respects
the most serious finding and the completeness of the evidence.

## Verification on Windows (mandatory before relying on the output)

The Python logic is unit-tested and the collector has been run under
PowerShell 7 with mocked cmdlets, but Windows PowerShell 5.1 has not been
executed in the development environment.

1. Run a Quick audit with 5 samples on mains power.
2. In the coverage table confirm: "Live performance samples" collected with
   5/5 and valid reads for CPU, per-core, frequency, memory, disk, disk
   latency, processes and power; "Event logs" collected; "Battery" shows a
   capacity source; every "requires elevation" row is one of the elevation
   checks listed above.
3. Confirm `Memory_AveragePercent` is within a few points of Task Manager's
   memory percentage and `CPU_Performance_AveragePercent` is non-null.
4. Confirm no `PC_Health_Audit_*` folder is left under `%TEMP%`.
5. Repeat as administrator only if the elevation-only checks matter.

## Troubleshooting

- Window does not open: run `py -3 PC_Health_Audit.pyw` from a Command Prompt
  and read the error; also check `%TEMP%\PC_Health_Audit_crash.log`.
- Coverage row "Collector: Failed" mentioning execution policy or AppLocker:
  company policy blocks PowerShell scripts for your account. Ask IT to run the
  audit or to allow it; the tool cannot and does not bypass a machine policy.
- Every performance metric N/A with "Live performance samples" failed: the
  WMI performance classes are disabled or the Performance Counters are
  corrupted on the device (IT can rebuild them with `lodctr /R`).
- "CPU_Performance_* could not be evaluated": the processor performance
  counter is not exposed on this build; the frequency columns stay unverified.

## Tests

```
py -3 -m pytest tests -q
```

`tests/test_collector_under_pwsh.py` runs the embedded collector with mocked
cmdlets when `pwsh` (PowerShell 7) is available or `POWERSHELL_EXE` is set.
