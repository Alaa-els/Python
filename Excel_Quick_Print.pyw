"""
Excel Quick Print
=================

Quick PDF print utility for multi-tab Excel workbooks where per-tab
page setup matters.

Features:
- Native Excel PDF export (no LibreOffice, no font substitution)
- Sequential per-workbook wizard (one workbook at a time)
- Bulk operations: multi-select rows then set range / orientation / fit /
  page breaks / repeat rows / copy-from / toggle / reset-to-auto
- Single-row edit dialog and "Click in Excel" range capture with overlap
  detection
- Per-tab Preview in Excel before export
- Page numbering footer with custom left-side label per workbook
- Saves config next to each workbook for one-click reuse next time
- Never overwrites the source: produces _rev.NN copies of both workbook
  and PDF when a file with the target name already exists
- Companion mode (v10): the Document Bundle Builder can launch this tool
  with a list of workbooks, an output folder and a result file:

      Excel_Quick_Print.pyw --companion-result <result.json>
                            --output-dir <folder> [--no-workbook-copy]
                            [--title <text>] -- <workbook> [<workbook> ...]

  The wizard runs exactly as it does standalone (sheet selection, print
  areas, preview, saved .printcfg.json settings). Only the PDF output
  folder is redirected and the _rev.NN workbook copy is suppressed. On
  exit a JSON result is written for the caller. Without those arguments
  the tool behaves exactly as before.

Author:  Alaa Elsayed
Created: May 2025
Revised: v10 - companion mode, password no-prompt on open

Save as excel_quick_print.pyw and double-click. First run auto-installs
the needed packages into the user profile (no admin rights required).
"""

import os
import sys
import json
import re
import atexit
import subprocess
import importlib
import traceback
from datetime import datetime


# ---------------------------------------------------------------------------
# Companion mode (Document Bundle Builder integration)
#
# Parsed before anything else so that every exit path - including a
# failed bootstrap - can still hand a result back to the caller. Without
# the --companion-result argument COMPANION is None and nothing below
# changes the standalone behaviour.
# ---------------------------------------------------------------------------

COMPANION_PROTOCOL = 1


def parse_companion_args(argv):
    """Return a companion dict or None.

    Recognised: --companion-result PATH  --output-dir DIR
                [--no-workbook-copy] [--title TEXT] -- FILE [FILE ...]
    Files may also follow directly without the '--' marker."""
    if "--companion-result" not in argv:
        return None
    comp = {"result_path": None, "output_dir": None, "workbook_copy": True,
            "title": None, "files": [], "written": False}
    i = 0
    files_mode = False
    while i < len(argv):
        a = argv[i]
        if files_mode:
            comp["files"].append(a)
        elif a == "--":
            files_mode = True
        elif a == "--companion-result" and i + 1 < len(argv):
            comp["result_path"] = argv[i + 1]
            i += 1
        elif a == "--output-dir" and i + 1 < len(argv):
            comp["output_dir"] = argv[i + 1]
            i += 1
        elif a == "--title" and i + 1 < len(argv):
            comp["title"] = argv[i + 1]
            i += 1
        elif a == "--no-workbook-copy":
            comp["workbook_copy"] = False
        elif not a.startswith("--"):
            comp["files"].append(a)
        i += 1
    if not comp["result_path"]:
        return None
    if not comp["output_dir"]:
        comp["output_dir"] = os.path.dirname(comp["result_path"])
    comp["files"] = [os.path.abspath(f) for f in comp["files"]]
    return comp


def write_companion_result(comp, status, results=None, error=None):
    """Write the result JSON atomically. status: done | cancelled | error."""
    if not comp or not comp.get("result_path"):
        return False
    payload = {
        "protocol": COMPANION_PROTOCOL,
        "status": status,
        "error": error,
        "written_at": datetime.now().isoformat(),
        "output_dir": comp.get("output_dir"),
        "results": [],
    }
    for r in results or []:
        payload["results"].append({
            "path": r.get("path"),
            "ok": bool(r.get("ok")),
            "pdf_path": r.get("pdf_path"),
            "error": r.get("error"),
            "warning": r.get("warning"),
            "skipped_tabs": list(r.get("skipped_tabs") or []),
            "recovery_pdf_path": r.get("recovery_pdf_path"),
        })
    target = comp["result_path"]
    tmp = f"{target}.tmp.{os.getpid()}"
    try:
        os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        os.replace(tmp, target)
        comp["written"] = True
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False


COMPANION = parse_companion_args(sys.argv[1:])


def _companion_finalise():
    """atexit hook: if the tool exits before an export result was written
    (window closed, bootstrap failed, startup error), tell the caller so
    it does not wait for a file that never comes."""
    if COMPANION and not COMPANION.get("written"):
        write_companion_result(
            COMPANION, "cancelled",
            error="Excel Quick Print exited before exporting")


if COMPANION:
    atexit.register(_companion_finalise)


# ---------------------------------------------------------------------------
# Last-resort error display
# ---------------------------------------------------------------------------

def _fatal_error(title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(title, message)
        r.destroy()
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
        except Exception:
            print(f"[FATAL - {title}]\n{message}", file=sys.stderr)
    sys.exit(1)


try:
    # -----------------------------------------------------------------------
    # Startup bootstrap
    #
    # Three concerns, in order:
    #
    # 1. Tkinter must be importable. It ships with the python.org Windows
    # installer when Tcl/Tk is selected (the default), but can be
    # missing on Microsoft Store builds or stripped installs. We can't
    # show a Tk dialog without Tkinter, so fall back to ctypes
    # MessageBoxW and tell the user to reinstall Python.
    #
    # 2. pywin32 must be importable. Specifically, both 'win32com.client'
    # (for Excel COM) AND 'pythoncom' (for CoInitialize) must work,
    # because partial pywin32 installs can provide one without the
    # other. If missing, show a Tk yes/no prompt, then run:
    # python -m pip install --user --upgrade pywin32
    # and retry imports (with importlib.invalidate_caches so the
    # newly-installed package is visible without restarting).
    #
    # 3. If retry still fails, surface the exact manual command:
    # python -m pip install --user --upgrade pywin32
    # and exit cleanly.
    #
    # We do NOT attempt to install Python itself - by definition Python is
    # already present, since this code is running.
    # -----------------------------------------------------------------------

    def _native_error_box(title, message):
        """Display an error via the Windows MessageBox API.

        Used only when Tkinter itself is not available. ctypes ships with
        the Python standard library and works on every Windows install.
        """
        try:
            import ctypes
            # 0x10 = MB_ICONERROR
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
        except Exception:
            print(f"[{title}]\n{message}", file=sys.stderr)

    # -- Concern 1: Tkinter -------------------------------------------------
    try:
        import tkinter as _tk_check
        from tkinter import messagebox as _mb_check
    except Exception:
        _native_error_box(
            "Tkinter not available",
            "This tool needs Tkinter, which is part of the standard "
            "Windows Python installer.\n\n"
            "Tkinter appears to be missing. The most common causes are:\n"
            "  - Python was installed from the Microsoft Store (which "
            "excludes Tcl/Tk)\n"
            "  - Tcl/Tk was unticked during the python.org installer\n"
            "  - An embeddable or minimal Python distribution is in use\n\n"
            "Please install Python from https://www.python.org/ and at "
            "the 'Optional Features' step, keep 'tcl/tk and IDLE' "
            "selected (it is on by default). Then run this tool again."
        )
        sys.exit(1)

    # -- Concern 2: pywin32 (both submodules) -------------------------------
    #
    # Map of pip package name -> list of import names that MUST all work
    # to consider the package healthy. pywin32 needs both win32com.client
    # AND pythoncom for this tool to function.
    _REQUIRED = {
        "pywin32": ["win32com.client", "pythoncom"],
    }

    def _failed_imports(import_names):
        """Return the subset of import_names that cannot be imported."""
        out = []
        for name in import_names:
            try:
                importlib.import_module(name)
            except Exception:
                out.append(name)
        return out

    def _check_all_required():
        """Return dict of {package: [failing_imports]} for anything missing."""
        out = {}
        for pkg, imports in _REQUIRED.items():
            failed = _failed_imports(imports)
            if failed:
                out[pkg] = failed
        return out

    _missing_packages = _check_all_required()

    if _missing_packages:
        # Detect virtualenv up front. `pip install --user` inside a venv
        # is rejected by pip with a clear error, so we drop `--user`
        # when in a venv. Determining this NOW (before the consent
        # dialog) lets the dialog show the exact command we will run,
        # rather than always saying --user. An earlier approach was
        # inconsistent: it dropped --user during install but the prompt
        # still claimed --user was being used.
        _in_venv = (hasattr(sys, "real_prefix")
                    or (hasattr(sys, "base_prefix")
                        and sys.base_prefix != sys.prefix))

        # Build the actual command string we will use, both for display
        # in the consent dialog and for any later "run this manually"
        # error dialogs. Single source of truth for what the install
        # actually does in this environment.
        if _in_venv:
            _display_cmd_template = (
                "python -m pip install --upgrade <package>")
            _manual_cmd = "python -m pip install --upgrade pywin32"
        else:
            _display_cmd_template = (
                "python -m pip install --user --upgrade <package>")
            _manual_cmd = (
                "python -m pip install --user --upgrade pywin32")

        # Build a single hidden Tk root for the install prompt dialogs.
        # We re-use this root across multiple message boxes so they all
        # appear with the right parent and modal behaviour.
        _boot_root = _tk_check.Tk()
        _boot_root.withdraw()

        _pkg_list = "\n".join(f"  - {p}" for p in _missing_packages.keys())
        _consent = _mb_check.askyesno(
            "Install required packages",
            f"This tool needs the following Python package(s):\n\n"
            f"{_pkg_list}\n\n"
            f"Install them now?\n\n"
            f"The installer will run:\n"
            f"  {_display_cmd_template}\n\n"
            f"An internet connection is required. This may take a few "
            f"minutes the first time.\n\n"
            f"No administrator rights are needed.",
            parent=_boot_root,
        )
        if not _consent:
            _mb_check.showinfo(
                "Cancelled",
                "Cannot continue without the required package(s). "
                "Closing.",
                parent=_boot_root,
            )
            _boot_root.destroy()
            # Exit code 1: "user declined required setup". Previously
            # this was sys.exit(0) which signalled success to any
            # launcher checking the return code - misleading when the
            # tool didn't actually start.
            sys.exit(1)

        # Show a visible "Installing..." window while pip runs.
        # subprocess.run() below blocks the main thread, so this Toplevel
        # won't be interactive (no minimize/move) during the install,
        # but at least the user sees that something is happening rather
        # than a frozen-app feel.
        _progress = _tk_check.Toplevel(_boot_root)
        _progress.title("Installing")
        _progress.resizable(False, False)
        # Centre roughly on screen
        _progress.geometry("+%d+%d" % (
            _progress.winfo_screenwidth() // 2 - 200,
            _progress.winfo_screenheight() // 2 - 60))
        _progress_label = _tk_check.Label(
            _progress,
            text="Installing required package(s)...\n"
                 "This may take a few minutes. Please wait.",
            padx=20, pady=20, justify="center")
        _progress_label.pack()
        _progress.update_idletasks()
        _progress.update()  # force a render before we block on subprocess

        _install_errors = []
        try:
            for _pkg in _missing_packages.keys():
                _progress_label.configure(
                    text=f"Installing {_pkg}...\n"
                         f"This may take a few minutes. Please wait.")
                _progress.update()

                _cmd = [sys.executable, "-m", "pip", "install"]
                if not _in_venv:
                    _cmd.append("--user")
                _cmd += ["--upgrade", _pkg]

                try:
                    # Capture output so we can surface it in any error
                    # dialog. The user sees no live progress; this is
                    # the trade-off for keeping the bootstrap simple
                    # (no threading).
                    _result = subprocess.run(
                        _cmd,
                        capture_output=True,
                        text=True,
                        timeout=300,
                    )
                    if _result.returncode != 0:
                        _install_errors.append(
                            (_pkg,
                             (_result.stdout or "")
                             + "\n"
                             + (_result.stderr or ""))
                        )
                except subprocess.TimeoutExpired:
                    _install_errors.append(
                        (_pkg, "pip install timed out after 5 minutes")
                    )
                except FileNotFoundError as _e:
                    # Extremely unlikely: sys.executable itself could not
                    # be launched. (Missing pip does NOT raise
                    # FileNotFoundError - it raises CalledProcessError
                    # with 'No module named pip' in stderr, which we
                    # catch via the non-zero returncode path above.)
                    _install_errors.append(
                        (_pkg,
                         f"could not launch Python to run pip ({_e})")
                    )
                except Exception as _e:
                    _install_errors.append((_pkg, str(_e)))
        finally:
            try:
                _progress.destroy()
            except Exception:
                pass

        if _install_errors:
            _details = "\n\n".join(
                f"{_pkg}:\n{(_err or '').strip()[:1500]}"
                for _pkg, _err in _install_errors
            )
            _mb_check.showerror(
                "Install failed",
                f"One or more required packages could not be "
                f"installed.\n\n{_details}\n\n"
                f"You can install manually. Open a Command Prompt and "
                f"run:\n\n"
                f"  {_manual_cmd}\n\n"
                f"If that also fails, common causes are:\n"
                f"  - Corporate proxy blocking pypi.org\n"
                f"  - Antivirus blocking pip\n"
                f"  - Python from the Microsoft Store (which "
                f"sandboxes --user installs)\n"
                f"  - pip itself not installed (run: python -m "
                f"ensurepip --upgrade)",
                parent=_boot_root,
            )
            _boot_root.destroy()
            sys.exit(1)

        # Make the newly-installed packages visible to the running
        # interpreter. THIS IS THE MOST IMPORTANT POST-INSTALL STEP.
        #
        # invalidate_caches() alone is NOT enough: it clears the import
        # system's negative cache, but does not add anything to sys.path.
        # If the user site-packages directory did not exist when Python
        # started, site.py never added it to sys.path, so even after a
        # successful `pip install --user`, the import still fails.
        #
        # site.addsitedir() does the right thing: it appends the
        # directory to sys.path AND processes any .pth files in it
        # (pywin32 ships a .pth file that sets up additional submodule
        # paths, so this matters for pywin32 specifically).
        #
        # Without this fix, a clean first install would succeed at the
        # pip step but still fail the import retry below, forcing the
        # user to restart the tool to pick up the new packages.
        # Flagged as the highest-impact bootstrap
        # bug in an earlier approach.
        importlib.invalidate_caches()
        if not _in_venv:
            try:
                import site
                _user_site = site.getusersitepackages()
                if _user_site and os.path.isdir(_user_site):
                    site.addsitedir(_user_site)
            except Exception:
                # If addsitedir fails for any reason, fall through to
                # the retry below. Worst case: the retry surfaces a
                # clear "install incomplete" error.
                pass

        _still_missing = _check_all_required()
        if _still_missing:
            _details = "\n".join(
                f"  - {_pkg}: still cannot import "
                f"{', '.join(_imps)}"
                for _pkg, _imps in _still_missing.items()
            )
            _mb_check.showerror(
                "Install incomplete",
                f"The package(s) were installed but cannot be "
                f"imported:\n\n{_details}\n\n"
                f"This sometimes happens when pip installs into a "
                f"site-packages folder that the running Python does "
                f"not search. It can also happen if pywin32 needs its "
                f"post-install step to register DLLs.\n\n"
                f"Please close this tool, open a Command Prompt, and "
                f"run:\n\n"
                f"  {_manual_cmd}\n\n"
                f"Verify it completes successfully, then restart the "
                f"tool. If the problem persists, run:\n\n"
                f"  python -m pywin32_postinstall -install\n\n"
                f"to register the pywin32 DLLs (this may need an "
                f"elevated Command Prompt).",
                parent=_boot_root,
            )
            _boot_root.destroy()
            sys.exit(1)

        # All good - tear down the temporary root and continue.
        _boot_root.destroy()

    # End of bootstrap. From here, all imports are safe.
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import win32com.client
    import pythoncom

    # -----------------------------------------------------------------------
    # Excel constants
    # -----------------------------------------------------------------------

    xlA3 = 8
    xlA4 = 9
    xlLetter = 1
    xlPortrait = 1
    xlLandscape = 2
    xlTypePDF = 0
    xlQualityStandard = 0
    xlSheetVisible = -1
    xlSheetHidden = 0
    xlSheetVeryHidden = 2

    # Page-break types. The VPageBreaks collection contains BOTH manual
    # and automatic breaks; only manual ones can be deleted. Automatic
    # breaks are regenerated by Excel on every repagination.
    xlPageBreakManual = -4135
    xlPageBreakAutomatic = -4105
    xlPageBreakNone = -4142

    # 72 points per inch. Fixed by definition, so margins are computed in
    # Python rather than via Application.InchesToPoints, which costs a COM
    # round trip per call (six per tab).
    POINTS_PER_INCH = 72.0

    # Paper size choices exposed to the UI. Display label -> xl constant.
    PAPER_SIZES = {
        "A4": xlA4,
        "A3": xlA3,
        "Letter": xlLetter,
    }
    PAPER_SIZE_LABELS = {v: k for k, v in PAPER_SIZES.items()}

    # Use 9999 for "unlimited pages tall". Excel rejects 0 with the
    # "Unable to set the FitToPagesTall property of the PageSetup class"
    # error, which was the root cause of the bug seen on ~80% of per-item
    # tabs in the Progress Statement workbook.
    FIT_UNLIMITED = 9999

    # -----------------------------------------------------------------------
    # Exceptions
    # -----------------------------------------------------------------------

    class PrintCommunicationError(RuntimeError):
        """Application.PrintCommunication could not be restored to True.

        Distinct from a normal per-tab failure because PrintCommunication
        is an APPLICATION-wide setting, not a per-worksheet one. If it is
        stuck at False, every subsequent PageSetup write in the session -
        on later tabs, later workbooks, and in the user's own Excel
        windows - is silently cached rather than committed. Continuing to
        configure tabs in that state would produce a PDF whose layout
        does not match the configuration, with no error to show for it.

        The caller must therefore treat this as fatal to the whole run
        and stop immediately, not swallow it per tab.
        """
        pass

    # -----------------------------------------------------------------------
    # Config storage
    # -----------------------------------------------------------------------

    APP_DIR = os.path.join(os.path.expanduser("~"), ".excel_quick_print")

    # One-time migration from the old name (.ipc_pdf_printer). If the
    # new folder doesn't exist but the old one does, copy the contents
    # across so existing users keep their last_files list and footer
    # preferences. The old folder is left in place untouched - users
    # can clean it up themselves once they've confirmed everything
    # works. This is a no-op for fresh installs.
    _LEGACY_APP_DIR = os.path.join(os.path.expanduser("~"),
                                   ".ipc_pdf_printer")
    if not os.path.exists(APP_DIR) and os.path.isdir(_LEGACY_APP_DIR):
        try:
            os.makedirs(APP_DIR, exist_ok=True)
            for _name in os.listdir(_LEGACY_APP_DIR):
                _src = os.path.join(_LEGACY_APP_DIR, _name)
                _dst = os.path.join(APP_DIR, _name)
                if os.path.isfile(_src) and not os.path.exists(_dst):
                    try:
                        with open(_src, "rb") as _fr:
                            _data = _fr.read()
                        with open(_dst, "wb") as _fw:
                            _fw.write(_data)
                    except Exception:
                        # Best-effort - if a single file fails to copy
                        # we still want the rest to migrate.
                        pass
        except Exception:
            # Migration is best-effort; on any unexpected failure we
            # silently fall back to a fresh empty config.
            pass

    os.makedirs(APP_DIR, exist_ok=True)
    APP_CONFIG = os.path.join(APP_DIR, "settings.json")

    def load_app_config():
        if not os.path.exists(APP_CONFIG):
            return {}
        try:
            with open(APP_CONFIG, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _atomic_write_json(target_path, data):
        """Write `data` as JSON to `target_path` atomically.

        Writes to `<target>.tmp` first, fsyncs (best-effort), then uses
        `os.replace` to swap the tmp file into the final name. This
        gives all-or-nothing semantics: if anything fails partway, the
        target either contains the previous good content or doesn't
        exist yet, but is never left truncated.

        Hardening: write-then-rename for crash-safety. Real-world scenarios that
        could truncate a non-atomic write:
        - Process killed mid-write (Task Manager, BSOD, power loss)
        - Antivirus interrupts the write
        - Laptop enters sleep at the wrong moment
        - SharePoint sync agent grabs the file mid-write
        - Out-of-disk during write

        Returns (ok, error_msg). Best-effort cleans up the tmp file on
        failure so a stale tmp doesn't accumulate next to the target.
        """
        # Same folder as target so os.replace stays on the same volume
        # (cross-volume falls back to copy+delete which defeats the
        # atomicity). Include PID in the tmp name so two simultaneous
        # instances of the tool writing to the same sidecar can't
        # fight over the same tmp slot - each process gets its own.
        tmp_path = f"{target_path}.tmp.{os.getpid()}"
        try:
            # Pre-emptively clear any stale tmp from a previous failed
            # write (HHMMSS-style is overkill here; a single tmp slot
            # is fine because only one write at a time per file).
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
                # Best-effort fsync. On Windows, file.flush() then
                # os.fsync() pushes the kernel buffer to disk. If
                # fsync isn't available or fails, the write still
                # completed at the Python level - we just lose a
                # belt-and-braces guarantee against an OS crash in
                # the next few ms.
                try:
                    f.flush()
                    os.fsync(f.fileno())
                except Exception:
                    pass

            # Atomic rename. On Windows, os.replace overwrites the
            # destination (unlike os.rename, which raises). Same
            # volume means this is a metadata-only operation: either
            # it succeeds entirely or the destination is untouched.
            os.replace(tmp_path, target_path)
            return True, None
        except Exception as e:
            # Clean up tmp so we don't leave debris next to the target
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            return False, str(e)

    def save_app_config(data):
        # App-level config is non-critical; if save fails we silently
        # accept it. But still use atomic write so a crash mid-write
        # doesn't leave the file truncated and force the user through
        # a fresh-defaults experience next launch.
        _atomic_write_json(APP_CONFIG, data)

    def workbook_config_path(xlsx_path):
        base = os.path.splitext(xlsx_path)[0]
        return f"{base}.printcfg.json"

    def load_workbook_config(xlsx_path):
        """Return (data, error_msg).

        - (None, None): file doesn't exist (first-time analysis)
        - (None, "..."): file exists but couldn't be read or has the
                         wrong top-level shape
        - (dict, None): loaded successfully and schema-validated

        Schema validation: top level must be a dict containing a "tabs"
        list. Each tab must be a dict with a "name" string. Tab entries
        that don't meet that minimum are dropped (with the loss noted in
        the error message). Other fields are coerced/defaulted by
        downstream code via setdefault and safe_int.

        An earlier approach returned None on bad JSON but didn't catch
        valid-JSON-wrong-shape (a list at top level, missing "tabs",
        tabs without "name"). That could crash analysis or produce
        degenerate configs.
        """
        p = workbook_config_path(xlsx_path)
        if not os.path.exists(p):
            return None, None
        try:
            with open(p, encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            return None, f"file unreadable: {e}"

        if not isinstance(raw, dict):
            return None, (f"top-level is {type(raw).__name__}, "
                          f"expected an object with a 'tabs' list")
        tabs_raw = raw.get("tabs")
        if not isinstance(tabs_raw, list):
            return None, "missing or non-list 'tabs' key"

        # Filter to entries that have a name. Track how many we dropped
        # so the caller can warn if data was lost.
        clean_tabs = []
        dropped = 0
        for t in tabs_raw:
            if isinstance(t, dict) and isinstance(t.get("name"), str) \
                    and t["name"]:
                clean_tabs.append(t)
            else:
                dropped += 1

        result = dict(raw)
        result["tabs"] = clean_tabs
        err = None
        if dropped:
            err = (f"{dropped} tab entry(ies) in saved config were "
                   f"missing a usable 'name' and were dropped")
        return result, err

    def save_workbook_config(xlsx_path, data):
        """Return (ok, error_msg).

        Uses atomic write (tmp file + os.replace) so a crash or
        interrupt mid-write cannot leave the sidecar truncated.
        Note: the previous non-atomic write as the
        next hardening step. A truncated sidecar would then fail
        load_workbook_config's JSON parser, surfacing as a
        'corrupt config' dialog and losing the user's tab settings
        for that workbook.
        """
        p = workbook_config_path(xlsx_path)
        return _atomic_write_json(p, data)

    # -----------------------------------------------------------------------
    # Excel helpers
    # -----------------------------------------------------------------------

    class ExcelSession:
        """Wraps an Excel.Application instance. Headless or visible."""

        def __init__(self, visible=False):
            # Balance CoInitialize with CoUninitialize on init failure
            # so we don't leak COM apartments. If DispatchEx succeeded
            # but a later app property setter fails, Excel is now
            # running too - we must Quit() it before uninit, otherwise
            # the process is orphaned.
            pythoncom.CoInitialize()
            self.app = None
            try:
                self.app = win32com.client.DispatchEx("Excel.Application")
                self.app.Visible = visible
                self.app.DisplayAlerts = False
                self.app.ScreenUpdating = visible
            except Exception:
                if self.app is not None:
                    try:
                        self.app.Quit()
                    except Exception:
                        pass
                    self.app = None
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                raise
            # AutomationSecurity = 3 (msoAutomationSecurityForceDisable)
            # prevents any Workbook_Open or Auto_Open macros from running
            # when we open .xlsm files via COM. Without this, opening an
            # .xlsm could execute arbitrary VBA before we touch anything.
            # We track whether it set successfully so the App layer can
            # refuse to proceed with .xlsm files if security failed.
            self.macros_disabled = False
            try:
                self.app.AutomationSecurity = 3
                # Verify it took (some Excel versions silently accept and
                # ignore the assignment)
                try:
                    self.macros_disabled = (
                        int(self.app.AutomationSecurity) == 3)
                except Exception:
                    self.macros_disabled = True  # best-effort assume OK
            except Exception:
                self.macros_disabled = False
            self._owned_workbooks = []

        def open(self, path):
            # UpdateLinks=0: don't refresh linked data on open. Without
            # this, external links could be refreshed during our COM open
            # and then persisted when we Save() - silently modifying data.
            # IgnoreReadOnlyRecommended=True: bypass the modal prompt for
            # workbooks flagged read-only recommended.
            # CorruptLoad=0: normal load (the default).
            # Password: a value no real workbook uses. Excel ignores it
            # for unprotected files; for password-protected files it
            # raises an error instead of showing a modal prompt inside an
            # invisible Excel instance (which would hang the tool).
            wb = self.app.Workbooks.Open(
                Filename=os.path.abspath(path),
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                Password="\x01excel-quick-print-no-prompt\x02",
            )
            self._owned_workbooks.append(wb)
            return wb

        def quit(self):
            try:
                for wb in self._owned_workbooks:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
            finally:
                try:
                    self.app.Quit()
                except Exception:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def col_letter(n):
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    def col_index(letter):
        letter = letter.upper().strip()
        if not letter or not letter.isalpha():
            return 0
        n = 0
        for c in letter:
            n = n * 26 + (ord(c) - 64)
        return n

    _A1_PIECE_RE = re.compile(
        r"^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$")

    def parse_a1_piece(piece):
        """Parse a single A1 address like '$A$1:$J$45' or 'A1' into
        (row1, col1, row2, col2). Returns None if the piece is not a
        recognisable range. The returned coords are normalised so
        row1 <= row2 and col1 <= col2.
        """
        piece = piece.strip()
        if not piece:
            return None
        # Strip any SheetName! prefix if present
        if "!" in piece:
            piece = piece.rsplit("!", 1)[1]
        m = _A1_PIECE_RE.match(piece)
        if not m:
            return None
        c1_let, r1, c2_let, r2 = m.groups()
        c1 = col_index(c1_let)
        r1 = int(r1)
        if c2_let is None:
            # Single cell like "A1"
            c2 = c1
            r2 = r1
        else:
            c2 = col_index(c2_let)
            r2 = int(r2)
        # Normalise direction
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        if c1 == 0 or c2 == 0:
            return None
        return (r1, c1, r2, c2)

    def rects_overlap(a, b):
        """True if two (r1,c1,r2,c2) rectangles share any cell."""
        ar1, ac1, ar2, ac2 = a
        br1, bc1, br2, bc2 = b
        if ar2 < br1 or br2 < ar1:
            return False
        if ac2 < bc1 or bc2 < ac1:
            return False
        return True

    def rect_contains(outer, inner):
        """True if `inner` rectangle is entirely contained within
        `outer` (inclusive)."""
        or1, oc1, or2, oc2 = outer
        ir1, ic1, ir2, ic2 = inner
        return (or1 <= ir1 and ir2 <= or2
                and oc1 <= ic1 and ic2 <= oc2)

    def find_overlapping_pieces(pieces):
        """Given a list of A1 pieces (strings), return a list of
        (i, j, kind) tuples for overlapping pairs.
        kind is 'contained' if pieces[j] is wholly inside pieces[i],
        'contains' if pieces[i] is wholly inside pieces[j],
        or 'partial' otherwise.
        Pieces that fail to parse are ignored (caller is expected to
        validate parseability separately).
        """
        parsed = [(p, parse_a1_piece(p)) for p in pieces]
        out = []
        for i in range(len(parsed)):
            _, ra = parsed[i]
            if ra is None:
                continue
            for j in range(i + 1, len(parsed)):
                _, rb = parsed[j]
                if rb is None:
                    continue
                if not rects_overlap(ra, rb):
                    continue
                if rect_contains(ra, rb):
                    out.append((i, j, "contained"))
                elif rect_contains(rb, ra):
                    out.append((i, j, "contains"))
                else:
                    out.append((i, j, "partial"))
        return out

    def find_revisioned_path(desired_path, max_revisions=99):
        """Given a desired path, return either the same path (if it
        doesn't exist) or a path with a _rev.NN suffix inserted before
        the extension (if the desired path already exists).

        Examples:
            /folder/Report.pdf (doesn't exist)
                -> /folder/Report.pdf
            /folder/Report.pdf (exists, no revisions yet)
                -> /folder/Report_rev.01.pdf
            /folder/Report.pdf (Report.pdf and Report_rev.01.pdf exist)
                -> /folder/Report_rev.02.pdf

        Suffix is zero-padded to 2 digits (rev.01..rev.99) for clean
        sorting in Windows Explorer. Beyond rev.99 we just keep
        incrementing without padding (rev.100, rev.101, ...).

        Returns None if max_revisions is exhausted (should not happen
        in any sane workflow - 99 revisions of a single IPC04 export
        in one folder would itself be a problem worth surfacing).

        Field-requested by user: when a target file already exists in
        the same directory (e.g. a previous run of this month's IPC04
        export sitting alongside the workbook), produce a new
        revisioned name rather than overwriting the existing file.
        """
        if not os.path.exists(desired_path):
            return desired_path
        base, ext = os.path.splitext(desired_path)
        for n in range(1, max_revisions + 1):
            candidate = f"{base}_rev.{n:02d}{ext}"
            if not os.path.exists(candidate):
                return candidate
        # Beyond the padded range - fall through to unpadded
        n = max_revisions + 1
        while True:
            candidate = f"{base}_rev.{n}{ext}"
            if not os.path.exists(candidate):
                return candidate
            n += 1
            if n > 9999:
                # Safety bail
                return None

    def detect_print_setup(ws):
        """Best-effort detection of the print range for a worksheet.

        Returns (print_area, last_row, last_col) where:
          - print_area is a string like "$A$1:$Z$39" or
            "$A$1:$K$45,$M$1:$X$39" (multi-area), OR None if no
            PrintArea has been set on the worksheet
          - last_row, last_col are the bounding-box extents used for
            orientation and paper-size defaults

        Strategy:
          1. If the worksheet has an existing PageSetup.PrintArea, use
             that. This is far more reliable than UsedRange - the
             workbook author has explicitly defined what to print, and
             UsedRange would include any cell ever touched (formatting,
             cleared values, comments) which often extends far beyond
             real data. Root cause of the "auto-detect says row 64 when data ends at row 39" for IPC
             sub-tabs that have formatted blank rows below the data.
          2. Fall back to UsedRange only when no PrintArea is set.
        """
        # Strategy 1: existing PrintArea
        try:
            pa = ws.PageSetup.PrintArea
            if pa:
                # Compute bounding box across all areas (multi-area
                # PrintArea is comma-separated). Strip any
                # SheetName! prefix on each piece.
                min_row = min_col = None
                max_row = max_col = 0
                for piece in pa.split(","):
                    piece = piece.strip()
                    if not piece:
                        continue
                    if "!" in piece:
                        piece = piece.rsplit("!", 1)[1]
                    try:
                        rng = ws.Range(piece)
                    except Exception:
                        continue
                    r1 = rng.Row
                    c1 = rng.Column
                    r2 = r1 + rng.Rows.Count - 1
                    c2 = c1 + rng.Columns.Count - 1
                    if min_row is None or r1 < min_row:
                        min_row = r1
                    if min_col is None or c1 < min_col:
                        min_col = c1
                    if r2 > max_row:
                        max_row = r2
                    if c2 > max_col:
                        max_col = c2
                if min_row is not None:
                    return pa, max_row, max_col
        except Exception:
            pass

        # Strategy 2: UsedRange fallback
        try:
            ur = ws.UsedRange
            last_row = ur.Row + ur.Rows.Count - 1
            last_col = ur.Column + ur.Columns.Count - 1
            return None, last_row, last_col
        except Exception:
            return None, 1, 1

    # Backwards compatibility shim: a few call sites use detect_used_range
    # as a quick "what's the size of this sheet" check (e.g. when
    # refreshing bounds on workbook re-analysis). They don't need the
    # PrintArea string. Keep this thin wrapper so we don't have to touch
    # every caller.
    def detect_used_range(ws):
        _pa, last_row, last_col = detect_print_setup(ws)
        return last_row, last_col

    def auto_orientation(last_col):
        return xlLandscape if last_col > 12 else xlPortrait

    def auto_cfg(name, last_row, last_col, print_area=None):
        """Default auto-detected config for a fresh tab.

        If print_area is provided (e.g. from an existing PageSetup
        PrintArea on the worksheet), use that verbatim. Otherwise
        construct a rectangle from $A$1 to the detected bottom-right.
        """
        if not print_area:
            print_area = f"$A$1:${col_letter(last_col)}${last_row}"
        return {
            "name": name,
            "include": True,
            "print_area": print_area,
            "orientation": auto_orientation(last_col),
            # Default to A4 unless the tab is genuinely wide. 16+ cols on
            # landscape A4 starts to look cramped at sensible font sizes,
            # so jump to A3 there.
            "paper_size": xlA4 if last_col <= 16 else xlA3,
            "fit_width": 1,
            # fit_height = 1 means "fit to one page tall".
            # fit_height = FIT_UNLIMITED means "fit width, unlimited tall".
            "fit_height": 1 if last_row <= 50 else FIT_UNLIMITED,
            "title_rows": "",
            "splits": [],
            "_detected_rows": last_row,
            "_detected_cols": last_col,
        }

    def sanitise_fit_height(value):
        """Coerce any bad value to FIT_UNLIMITED.

        Excel rejects FitToPagesTall = 0 with a COM error. Legacy configs
        stored 0 to mean 'unlimited tall', so 0 maps to FIT_UNLIMITED.

        None, empty string, or unparseable values also map to FIT_UNLIMITED
        (the safer default - printing unlimited tall never truncates rows;
        forcing fit-to-1-page could squash a 200-row tab into illegible
        microprint).

        Positive ints pass through unchanged.

        Note: an earlier approach of this function returned 1 for
        None / "" / bad values. That was a real bug: a saved config
        with null/empty fit_height would silently become 'fit to 1
        page tall' and produce unreadable output.
        """
        if value is None:
            return FIT_UNLIMITED
        if isinstance(value, str) and not value.strip():
            return FIT_UNLIMITED
        try:
            v = int(value)
        except (TypeError, ValueError):
            return FIT_UNLIMITED
        if v <= 0:
            return FIT_UNLIMITED
        return v

    def safe_int(value, default):
        """Coerce a value to int, returning default on any failure.

        Used when loading saved configs - a corrupt or hand-edited JSON
        with an unexpected type for fit_width etc. must not crash
        workbook analysis.
        """
        if value is None:
            return default
        if isinstance(value, str) and not value.strip():
            return default
        try:
            v = int(value)
        except (TypeError, ValueError):
            return default
        return v if v > 0 else default

    # -----------------------------------------------------------------------
    # Branding
    # -----------------------------------------------------------------------

    # Small attribution label shown at the bottom-left of every dialog
    # and the main window. Stays visible without being intrusive.
    BRAND_TEXT = "Sr. QS - Alaa Elsayed"
    BRAND_FONT = ("Calibri", 8, "italic")
    BRAND_FG = "#6b6b6b"  # muted grey so it doesn't compete with content

    def add_brand_label(window):
        """Pack a small attribution label at the bottom-left of the
        given Tk Toplevel / Tk root. Idempotent: safe to call twice on
        the same window without producing two labels.

        Uses .pack(side='bottom', anchor='w') so the label sits below
        whatever was packed before it - call this AFTER you've packed
        the dialog's main content for it to land at the bottom.
        """
        if getattr(window, "_brand_label_added", False):
            return
        try:
            lbl = tk.Label(
                window,
                text=BRAND_TEXT,
                font=BRAND_FONT,
                fg=BRAND_FG,
                anchor="w",
            )
            lbl.pack(side="bottom", anchor="w",
                     padx=8, pady=(2, 4),
                     fill="x")
            window._brand_label_added = True
        except Exception:
            # If a window doesn't support the label (e.g. it was
            # destroyed mid-call) just swallow - branding is cosmetic
            # and must not break a working dialog.
            pass

    # -----------------------------------------------------------------------
    # Main App
    # -----------------------------------------------------------------------

    class App:

        def __init__(self):
            self.root = tk.Tk()
            self.root.title("Excel Quick Print")
            self.root.geometry("1080x740")
            self.root.minsize(900, 600)

            # ttk style
            style = ttk.Style()
            try:
                style.theme_use("vista")
            except Exception:
                pass

            self.files = []                # selected xlsx paths
            self.workbook_data = {}        # path -> [tab_config dicts]
            self.workbook_footers = {}     # path -> footer string (per-workbook
                                            # left-side PDF footer; populated
                                            # at start of Step 3 via prompt)
            self.app_config = load_app_config()

            # One-time migration: if the loaded settings still carry
            # the old project-specific footer default that the tool
            # used in its earlier incarnation, blank it. The tool's
            # generalised form defaults to empty footer. Gated on a
            # version marker in the settings file so a user who later
            # types this string deliberately is not affected.
            if not self.app_config.get("_footer_default_migrated"):
                if self.app_config.get("left_footer") == "D-18 | IPC 04":
                    self.app_config["left_footer"] = ""
                self.app_config["_footer_default_migrated"] = True
                save_app_config(self.app_config)

            # Companion mode (launched by the Document Bundle Builder).
            # Pre-load the requested workbooks; do not merge last-session
            # files, which belong to a different job.
            self.companion = COMPANION
            self._companion_results = None
            if self.companion:
                self.files = [f for f in self.companion["files"]
                              if os.path.isfile(f)]
                try:
                    os.makedirs(self.companion["output_dir"], exist_ok=True)
                except Exception:
                    pass

            self.excel = None              # ExcelSession during configure stage
            self.current_wb_index = 0      # for sequential wizard
            self.current_tv = None         # current treeview reference
            self._exporting = False        # set True during Step-3 export

            # Register window-close handler from the start so it covers
            # every wizard step, not just Step 3. Registering only in
            # _build_step3 (as an earlier approach did) would mean
            # closing during Step 1 or Step 2 (e.g. mid-analysis) could
            # orphan the Excel process.
            self.root.protocol("WM_DELETE_WINDOW", self._on_close_window)

            self._build_step1()

        # ------------------------------------------------------------------
        # STEP 1 - Choose files
        # ------------------------------------------------------------------

        def _build_step1(self):
            self._clear()
            self.root.title("Excel Quick Print - Step 1 of 3")

            frm = ttk.Frame(self.root, padding=20)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="Excel Quick Print",
                      font=("Calibri", 18, "bold")).pack(anchor="w",
                                                         pady=(0, 4))
            ttk.Label(frm, text="Excel-to-PDF using Excel's native print "
                                "engine (best quality, real fonts, no "
                                "substitution).",
                      foreground="#555").pack(anchor="w", pady=(0, 16))

            if self.companion:
                ttk.Label(
                    frm,
                    text=("Companion mode: exporting for the Document Bundle "
                          "Builder" + (f" ({self.companion['title']})"
                                       if self.companion.get("title") else "")
                          + f".\nPDFs go to: {self.companion['output_dir']}\n"
                          "Workbook _rev copies are not written; print "
                          "settings are still saved next to each workbook."),
                    foreground="#7a4a00", wraplength=900, justify="left",
                ).pack(anchor="w", pady=(0, 12))

            ttk.Label(frm, text="Step 1.  Choose Excel files to print:",
                      font=("Calibri", 11, "bold")).pack(anchor="w")

            list_frame = ttk.Frame(frm)
            list_frame.pack(fill="both", expand=True, pady=8)

            self.files_box = tk.Listbox(list_frame, height=10,
                                        selectmode="extended",
                                        font=("Consolas", 9))
            self.files_box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.files_box.yview)
            sb.pack(side="right", fill="y")
            self.files_box.configure(yscrollcommand=sb.set)

            # Populate the file list. Dedup against current self.files
            # (which may be populated already if we're returning from
            # Step 2). Without this guard, _back_to_step1 would append
            # last_files to the existing list, producing duplicates.
            #
            current_paths = set(self.files)
            for box_path in list(self.files_box.get(0, tk.END)):
                # Treeview/Listbox may already display the existing
                # selection; if it does, leave it alone. If it doesn't
                # (fresh build), it's empty and we won't iterate.
                current_paths.add(box_path)

            # If files_box is empty but self.files is not (we just
            # re-entered Step 1 with state), repopulate the listbox to
            # match self.files first.
            if not self.files_box.get(0, tk.END) and self.files:
                for p in self.files:
                    self.files_box.insert(tk.END, p)

            # Now add any last-session files that aren't already shown
            # (standalone only - a companion run has its own file list)
            last_files = ([] if self.companion
                          else self.app_config.get("last_files", []))
            for f in last_files:
                if not os.path.exists(f):
                    continue
                if f in current_paths:
                    continue
                self.files.append(f)
                self.files_box.insert(tk.END, f)
                current_paths.add(f)

            btns = ttk.Frame(frm)
            btns.pack(fill="x", pady=(8, 0))
            ttk.Button(btns, text="Add files...",
                       command=self._add_files).pack(side="left", padx=(0, 6))
            ttk.Button(btns, text="Remove selected",
                       command=self._remove_files).pack(side="left", padx=6)
            ttk.Button(btns, text="Clear all",
                       command=self._clear_files).pack(side="left", padx=6)

            footer = ttk.Frame(frm)
            footer.pack(fill="x", pady=(20, 0))

            self.left_footer_var = tk.StringVar(
                value=self.app_config.get("left_footer", ""))
            ttk.Label(footer, text="Footer text (left):").grid(
                row=0, column=0, sticky="w")
            ttk.Entry(footer, textvariable=self.left_footer_var,
                      width=40).grid(row=0, column=1, padx=8, sticky="w")
            ttk.Label(footer, text="(Right side always shows 'Page X of Y')",
                      foreground="#777").grid(row=0, column=2, sticky="w",
                                              padx=(8, 0))

            nav = ttk.Frame(frm)
            nav.pack(fill="x", pady=(20, 0))
            ttk.Button(nav, text="Quit",
                       command=self._on_close_window).pack(side="right",
                                                           padx=(6, 0))
            ttk.Button(nav, text="Next >",
                       command=self._goto_step2).pack(side="right")

            add_brand_label(self.root)

        def _add_files(self):
            paths = filedialog.askopenfilenames(
                title="Select Excel files",
                filetypes=[("Excel workbooks",
                            "*.xlsx;*.xlsm;*.xlsb;*.xls"),
                           ("All files", "*.*")]
            )
            for p in paths:
                if p not in self.files:
                    self.files.append(p)
                    self.files_box.insert(tk.END, p)

        def _remove_files(self):
            sel = list(self.files_box.curselection())
            sel.reverse()
            for i in sel:
                del self.files[i]
                self.files_box.delete(i)

        def _clear_files(self):
            self.files.clear()
            self.files_box.delete(0, tk.END)

        def _goto_step2(self):
            if not self.files:
                messagebox.showwarning("No files",
                                       "Please add at least one Excel file.")
                return
            if not self.companion:
                # A companion run's file list belongs to the bundle job,
                # not to the user's standalone history.
                self.app_config["last_files"] = list(self.files)
            self.app_config["left_footer"] = self.left_footer_var.get()
            save_app_config(self.app_config)
            self._build_step2()

        # ------------------------------------------------------------------
        # STEP 2 - Sequential per-workbook wizard
        # ------------------------------------------------------------------

        def _build_step2(self):
            """Run auto-detect on all workbooks, then show first one."""
            self._clear()
            self.root.title("Excel Quick Print - Step 2 of 3")

            self.current_wb_index = 0

            # Loading frame while analysing
            self._loading_frame = ttk.Frame(self.root, padding=30)
            self._loading_frame.pack(fill="both", expand=True)

            ttk.Label(self._loading_frame,
                      text="Step 2.  Analysing workbooks...",
                      font=("Calibri", 13, "bold")).pack(anchor="w",
                                                         pady=(0, 12))
            self.status = tk.StringVar(value="Starting...")
            ttk.Label(self._loading_frame, textvariable=self.status,
                      foreground="#0045BF",
                      font=("Consolas", 10)).pack(anchor="w")

            self.root.after(50, self._auto_detect_all)

        def _auto_detect_all(self):
            try:
                # Quit any existing Excel session before spawning a new one.
                # Without this, navigating Step 2 -> Back to Step 1 -> Next
                # leaks an EXCEL.EXE process on every round trip.
                if self.excel is not None:
                    try:
                        self.excel.quit()
                    except Exception:
                        pass
                    self.excel = None

                self.excel = ExcelSession(visible=False)

                # If any macro-capable workbooks are in the selection AND
                # we couldn't disable macros, refuse to open them.
                # Opening an .xlsm/.xlsb/.xls with macros enabled via COM
                # would execute Workbook_Open / Auto_Open code without
                # the user's consent. .xls and .xlsb both support
                # macros - earlier rounds only guarded .xlsm. The file
                # picker offers an "All files" option so .xls/.xlsb
                # files can reach this point.
                MACRO_CAPABLE_EXTS = (".xlsm", ".xlsb", ".xls")
                macro_files = [p for p in self.files
                               if p.lower().endswith(MACRO_CAPABLE_EXTS)]
                if macro_files and not self.excel.macros_disabled:
                    self._loading_frame.destroy()
                    names = "\n".join("  - " + os.path.basename(p)
                                      for p in macro_files)
                    messagebox.showerror(
                        "Macro security",
                        f"The following macro-capable workbooks are in "
                        f"your selection:\n\n{names}\n\n"
                        f"Excel's AutomationSecurity could not be set "
                        f"to disable macros. To prevent any macro from "
                        f"running automatically, the tool will not "
                        f"open these files.\n\n"
                        f"Re-save them as .xlsx if they don't need "
                        f"macros, or open them manually in Excel first.")
                    try:
                        self.excel.quit()
                        self.excel = None
                    except Exception:
                        pass
                    return

                # Track any corrupt sidecar configs we hit during analysis
                # so we can show one consolidated warning at the end.
                self._corrupt_configs = []

                # Preserve any prior workbook_data for files still in the
                # selection. This prevents the user losing all their manual
                # overrides by going back to Step 1 and forward again.
                prior = self.workbook_data
                self.workbook_data = {}

                for path in self.files:
                    self.status.set(f"Analysing  {os.path.basename(path)} ...")
                    # update_idletasks() processes label redraw only, not
                    # button clicks or window close - so the user can't
                    # cause reentrancy during a COM-heavy analysis loop.
                    # Note: an earlier approach used self.root.update() here,
                    # which is the same reentrancy hazard fixed in _log_line.
                    self.root.update_idletasks()
                    wb = self.excel.open(path)

                    if path in prior:
                        # Keep the user's in-memory overrides; just refresh
                        # the auto-detect bounds so reset-to-auto still works
                        existing = prior[path]
                        for cfg in existing:
                            try:
                                ws = wb.Sheets(cfg["name"])
                                lr, lc = detect_used_range(ws)
                                cfg["_detected_rows"] = lr
                                cfg["_detected_cols"] = lc
                            except Exception:
                                pass
                        self.workbook_data[path] = existing
                    else:
                        self.workbook_data[path] = \
                            self._analyse_workbook(wb, path)

                # If any sidecar configs had issues, show a dialog. Fatal
                # losses (config unreadable, lost everything) and partial
                # losses (some entries dropped, most preserved) are shown
                # separately so the user understands the actual impact.
                if self._corrupt_configs:
                    fatal = [(p, e) for p, e, sev in self._corrupt_configs
                             if sev == "fatal"]
                    partial = [(p, e) for p, e, sev in self._corrupt_configs
                               if sev == "partial"]
                    msg_parts = []
                    if fatal:
                        msg_parts.append("UNREADABLE (settings lost, "
                                         "auto-detected from scratch):")
                        msg_parts += [f"  - {os.path.basename(p)}: {e}"
                                      for p, e in fatal]
                    if partial:
                        if msg_parts:
                            msg_parts.append("")
                        msg_parts.append("PARTIAL (most settings loaded, "
                                         "some entries dropped):")
                        msg_parts += [f"  - {os.path.basename(p)}: {e}"
                                      for p, e in partial]
                    msg_parts.append("")
                    msg_parts.append("Review the affected tabs before "
                                     "exporting.")
                    messagebox.showwarning(
                        "Config file issue(s)",
                        "\n".join(msg_parts))

                # Tear down the loading frame and show the first workbook
                self._loading_frame.destroy()
                self._show_workbook(0)
            except Exception as e:
                tb = traceback.format_exc()
                messagebox.showerror("Analysis failed",
                                     f"Could not analyse workbooks.\n\n"
                                     f"{e}\n\n{tb}")
                try:
                    if self.excel is not None:
                        self.excel.quit()
                        self.excel = None
                except Exception:
                    pass

        def _analyse_workbook(self, wb, path):
            """Return list of tab dicts describing each visible sheet.

            Hidden / VeryHidden sheets (Dashboard, IPC History Audit, etc.)
            are excluded entirely - never appear in the configurator.
            """
            saved_cfg, load_err = load_workbook_config(path)
            if load_err:
                # Distinguish fully-failed load (saved_cfg is None,
                # everything is lost, auto-detected from scratch) from
                # partial load (saved_cfg is a dict, some entries
                # dropped, most settings preserved). Earlier code
                # treated both as "corrupt - settings lost" which was
                # misleading.
                severity = "fatal" if saved_cfg is None else "partial"
                self._corrupt_configs.append((path, load_err, severity))
            saved_tabs = {t["name"]: t for t in saved_cfg.get("tabs", [])} \
                if saved_cfg else {}

            # If the sidecar holds a saved per-workbook footer, seed
            # self.workbook_footers so the Step-3 prompt pre-fills it.
            # In-memory value (set during this session) takes priority
            # over the sidecar; sidecar takes priority over the global
            # default.
            if saved_cfg and isinstance(saved_cfg.get("left_footer"),
                                        str):
                if path not in self.workbook_footers:
                    self.workbook_footers[path] = saved_cfg["left_footer"]

            tabs = []
            # xlWorksheet = -4167. wb.Sheets includes chart sheets and
            # dialog sheets, which don't have UsedRange/PrintArea and would
            # crash _apply_tab_setup. Skip anything that's not a worksheet.
            xlWorksheet = -4167
            for i in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(i)
                try:
                    if ws.Type != xlWorksheet:
                        continue
                except Exception:
                    # If we can't read .Type, assume it's not safe
                    continue
                if ws.Visible != xlSheetVisible:
                    continue
                name = ws.Name
                # detect_print_setup returns the workbook's existing
                # PrintArea (if any) plus the bounding-box extents.
                # This is what fixes the "auto-detect shows row 64
                # when data ends at row 39" issue on IPC sub-tabs
                # whose UsedRange is polluted by formatted blank rows
                # below the actual data.
                existing_pa, last_row, last_col = detect_print_setup(ws)
                default_print_area = (
                    existing_pa
                    or f"$A$1:${col_letter(last_col)}${last_row}")

                if name in saved_tabs:
                    # Reuse saved config. Every required field must be
                    # backfilled - earlier rounds only defaulted some of
                    # them, so a saved config missing 'orientation' or
                    # 'print_area' would survive validation and then
                    # crash the UI on _row_view.
                    cfg = dict(saved_tabs[name])
                    cfg.setdefault("include", True)
                    cfg.setdefault("title_rows", "")
                    cfg.setdefault("splits", [])
                    # print_area: fall back to the workbook's existing
                    # PrintArea if set, otherwise the auto-detected
                    # rectangle. Saved value (from sidecar JSON) wins
                    # if present - that's the user's previous choice.
                    cfg.setdefault("print_area", default_print_area)
                    # orientation: fall back to width-based auto-detect
                    cfg.setdefault(
                        "orientation", auto_orientation(last_col))
                    # paper_size: fall back to width-based default
                    cfg.setdefault(
                        "paper_size",
                        xlA4 if last_col <= 16 else xlA3)
                    cfg["fit_height"] = sanitise_fit_height(
                        cfg.get("fit_height", 1))
                    cfg["fit_width"] = safe_int(cfg.get("fit_width", 1), 1)
                    # Coerce orientation to a valid value
                    if cfg["orientation"] not in (xlPortrait, xlLandscape):
                        cfg["orientation"] = auto_orientation(last_col)
                    # Coerce paper_size to a valid value
                    if cfg["paper_size"] not in PAPER_SIZE_LABELS:
                        cfg["paper_size"] = (xlA4 if last_col <= 16
                                             else xlA3)
                    cfg["_detected_rows"] = last_row
                    cfg["_detected_cols"] = last_col
                else:
                    cfg = auto_cfg(name, last_row, last_col,
                                   print_area=existing_pa)
                tabs.append(cfg)
            return tabs

        # ------------------------------------------------------------------
        # Per-workbook view (sequential)
        # ------------------------------------------------------------------

        def _show_workbook(self, index):
            """Render the current workbook view."""
            self._clear()
            self.root.title(f"Excel Quick Print - Step 2 of 3 "
                            f"(Workbook {index + 1} of {len(self.files)})")

            self.current_wb_index = index
            path = self.files[index]
            tabs = self.workbook_data[path]

            frm = ttk.Frame(self.root, padding=12)
            frm.pack(fill="both", expand=True)

            # Header
            hdr = ttk.Frame(frm)
            hdr.pack(fill="x", pady=(0, 6))
            ttk.Label(hdr,
                      text=f"Workbook {index + 1} of {len(self.files)}:",
                      foreground="#777").pack(side="left")
            ttk.Label(hdr, text=os.path.basename(path),
                      font=("Calibri", 12, "bold"),
                      foreground="#0045BF").pack(side="left", padx=(6, 0))

            ttk.Label(frm,
                      text="Multi-select rows (Ctrl/Shift-click) then use "
                           "the bulk toolbar. Single-row actions are on the "
                           "right.",
                      foreground="#555").pack(anchor="w", pady=(0, 8))

            # Bulk toolbar (top)
            self._build_bulk_toolbar(frm, path)

            # Body: treeview + right side single-row buttons
            body = ttk.Frame(frm)
            body.pack(fill="both", expand=True, pady=(8, 0))

            # Treeview
            tv_frame = ttk.Frame(body)
            tv_frame.pack(side="left", fill="both", expand=True,
                          padx=(0, 4))

            cols = ("name", "include", "area", "paper", "orient", "fit",
                    "titles", "splits")
            tv = ttk.Treeview(tv_frame, columns=cols, show="headings",
                              height=18, selectmode="extended")
            tv.heading("name", text="Tab")
            tv.heading("include", text="Print?")
            tv.heading("area", text="Print Area")
            tv.heading("paper", text="Paper")
            tv.heading("orient", text="Orientation")
            tv.heading("fit", text="Fit")
            tv.heading("titles", text="Repeat Rows")
            tv.heading("splits", text="Page Breaks")
            tv.column("name", width=180)
            tv.column("include", width=55, anchor="center")
            tv.column("area", width=160)
            tv.column("paper", width=60, anchor="center")
            tv.column("orient", width=85, anchor="center")
            tv.column("fit", width=150, anchor="center")
            tv.column("titles", width=85, anchor="center")
            tv.column("splits", width=100, anchor="center")
            tv.pack(side="left", fill="both", expand=True)

            sb = ttk.Scrollbar(tv_frame, orient="vertical",
                               command=tv.yview)
            sb.pack(side="right", fill="y")
            tv.configure(yscrollcommand=sb.set)

            # Tag styles
            tv.tag_configure("skip",
                             foreground="#999",
                             font=("Calibri", 9, "overstrike"))
            tv.tag_configure("normal", foreground="#000")

            # Populate
            for cfg in tabs:
                tv.insert("", tk.END, values=self._row_view(cfg),
                          tags=("normal" if cfg["include"] else "skip",))

            self.current_tv = tv

            # Right side - single-row actions
            right = ttk.Frame(body)
            right.pack(side="right", fill="y", padx=(8, 0))

            ttk.Label(right, text="Single row",
                      font=("Calibri", 10, "bold")).pack(anchor="w",
                                                         pady=(0, 4))

            self.btn_edit = ttk.Button(
                right, text="Edit selected...",
                command=lambda: self._edit_tab_dialog(path, tv))
            self.btn_edit.pack(fill="x", pady=2)

            self.btn_click = ttk.Button(
                right, text="Click in Excel\n(set range)",
                command=lambda: self._click_in_excel(path, tv))
            self.btn_click.pack(fill="x", pady=2)

            self.btn_preview = ttk.Button(
                right, text="Preview in Excel",
                command=lambda: self._preview_in_excel(path, tv))
            self.btn_preview.pack(fill="x", pady=2)

            # Selection event updates button states
            tv.bind("<<TreeviewSelect>>",
                    lambda e: self._update_single_row_buttons())
            self._update_single_row_buttons()

            # Navigation footer
            nav = ttk.Frame(frm)
            nav.pack(fill="x", pady=(12, 0))

            ttk.Button(nav, text="< Back to file list",
                       command=self._back_to_step1).pack(side="left")

            ttk.Button(nav, text="Save configs (all workbooks)",
                       command=self._save_workbook_configs).pack(
                           side="left", padx=8)

            # Right-side navigation
            if index < len(self.files) - 1:
                ttk.Button(nav, text="Next workbook >",
                           command=self._next_workbook).pack(side="right")
            else:
                ttk.Button(nav, text="Continue to Export >",
                           command=self._goto_step3).pack(side="right")

            if index > 0:
                ttk.Button(nav, text="< Previous workbook",
                           command=self._prev_workbook).pack(
                               side="right", padx=(0, 8))

            add_brand_label(self.root)

        def _build_bulk_toolbar(self, parent, path):
            """Bulk action toolbar - applies to all selected rows."""
            bar = ttk.LabelFrame(parent, text="Bulk actions (apply to all "
                                              "selected rows)",
                                 padding=6)
            bar.pack(fill="x")

            row1 = ttk.Frame(bar)
            row1.pack(fill="x")

            ttk.Button(row1, text="Select all",
                       command=lambda: self._bulk_select_all()).pack(
                           side="left", padx=2)
            ttk.Button(row1, text="Toggle Print",
                       command=lambda: self._bulk_toggle(path)).pack(
                           side="left", padx=2)
            ttk.Separator(row1, orient="vertical").pack(side="left",
                                                        fill="y", padx=6)
            ttk.Button(row1, text="Set range...",
                       command=lambda: self._bulk_set_range(path)).pack(
                           side="left", padx=2)
            ttk.Button(row1, text="Portrait",
                       command=lambda: self._bulk_set_orient(
                           path, xlPortrait)).pack(side="left", padx=2)
            ttk.Button(row1, text="Landscape",
                       command=lambda: self._bulk_set_orient(
                           path, xlLandscape)).pack(side="left", padx=2)
            ttk.Separator(row1, orient="vertical").pack(side="left",
                                                        fill="y", padx=6)
            ttk.Button(row1, text="A4",
                       command=lambda: self._bulk_set_paper(
                           path, xlA4)).pack(side="left", padx=2)
            ttk.Button(row1, text="A3",
                       command=lambda: self._bulk_set_paper(
                           path, xlA3)).pack(side="left", padx=2)
            ttk.Separator(row1, orient="vertical").pack(side="left",
                                                        fill="y", padx=6)
            ttk.Button(row1, text="Fit 1 page",
                       command=lambda: self._bulk_set_fit(
                           path, "one")).pack(side="left", padx=2)
            ttk.Button(row1, text="Fit width",
                       command=lambda: self._bulk_set_fit(
                           path, "width")).pack(side="left", padx=2)

            row2 = ttk.Frame(bar)
            row2.pack(fill="x", pady=(4, 0))

            ttk.Button(row2, text="Set page breaks...",
                       command=lambda: self._bulk_set_breaks(path)).pack(
                           side="left", padx=2)
            ttk.Button(row2, text="Set repeat rows...",
                       command=lambda: self._bulk_set_titles(path)).pack(
                           side="left", padx=2)
            ttk.Separator(row2, orient="vertical").pack(side="left",
                                                        fill="y", padx=6)
            ttk.Button(row2, text="Copy settings from tab...",
                       command=lambda: self._bulk_copy_from(path)).pack(
                           side="left", padx=2)
            ttk.Separator(row2, orient="vertical").pack(side="left",
                                                        fill="y", padx=6)
            ttk.Button(row2, text="Reset to auto-detect",
                       command=lambda: self._bulk_reset_auto(path)).pack(
                           side="left", padx=2)

        # ------------------------------------------------------------------
        # Row view helpers
        # ------------------------------------------------------------------

        def _row_view(self, cfg):
            orient = ("Portrait" if cfg["orientation"] == xlPortrait
                      else "Landscape")
            paper = PAPER_SIZE_LABELS.get(
                cfg.get("paper_size", xlA4), "A4")
            fit_w = safe_int(cfg.get("fit_width", 1), 1)
            fit_h = sanitise_fit_height(cfg.get("fit_height", 1))
            if fit_w == 1 and fit_h == 1:
                fit = "Fit 1 page"
            elif fit_w == 1 and fit_h == FIT_UNLIMITED:
                fit = "Fit width x unlimited"
            else:
                fit = f"{fit_w}w x {fit_h}h"
            splits = ", ".join(cfg.get("splits", [])) or "-"
            return (cfg["name"],
                    "Yes" if cfg["include"] else "No",
                    cfg["print_area"],
                    paper,
                    orient,
                    fit,
                    cfg["title_rows"] or "-",
                    splits)

        def _refresh_row(self, tv, item_id, cfg):
            tag = "normal" if cfg["include"] else "skip"
            tv.item(item_id, values=self._row_view(cfg), tags=(tag,))

        def _refresh_all_rows(self, path, tv):
            tabs = self.workbook_data[path]
            by_name = {t["name"]: t for t in tabs}
            for item_id in tv.get_children():
                try:
                    vals = tv.item(item_id, "values")
                    name = vals[0] if vals else None
                except Exception:
                    name = None
                if name and name in by_name:
                    self._refresh_row(tv, item_id, by_name[name])

        # ------------------------------------------------------------------
        # Selected row(s) accessors
        # ------------------------------------------------------------------

        def _get_selected_cfgs(self, path, tv):
            """Return list of (item_id, idx, cfg) for currently selected rows.

            Look up the cfg by the tab name stored in column 0, not by
            Treeview position. Tab names are unique within a workbook, so
            this stays correct even if column-click sorting is added later.
            """
            sel = tv.selection()
            if not sel:
                return []
            tabs = self.workbook_data[path]
            by_name = {t["name"]: (i, t) for i, t in enumerate(tabs)}
            result = []
            for item_id in sel:
                try:
                    vals = tv.item(item_id, "values")
                    name = vals[0] if vals else None
                except Exception:
                    name = None
                if name and name in by_name:
                    idx, cfg = by_name[name]
                    result.append((item_id, idx, cfg))
            return result

        def _need_selection(self, tv):
            if not tv.selection():
                messagebox.showinfo("Select rows",
                                    "Please select at least one tab in the "
                                    "list first.")
                return False
            return True

        def _update_single_row_buttons(self):
            """Enable single-row buttons only when exactly one row selected."""
            if self.current_tv is None:
                return
            sel = self.current_tv.selection()
            state = "normal" if len(sel) == 1 else "disabled"
            for btn in (self.btn_edit, self.btn_click, self.btn_preview):
                try:
                    btn.configure(state=state)
                except Exception:
                    pass

        # ------------------------------------------------------------------
        # Bulk actions
        # ------------------------------------------------------------------

        def _bulk_select_all(self):
            if self.current_tv is None:
                return
            self.current_tv.selection_set(self.current_tv.get_children())

        def _bulk_toggle(self, path):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            for _item_id, _idx, cfg in self._get_selected_cfgs(path, tv):
                cfg["include"] = not cfg["include"]
            self._refresh_all_rows(path, tv)

        def _bulk_set_range(self, path):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            sel = self._get_selected_cfgs(path, tv)

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Set print area for selected tabs")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("520x220")

            ttk.Label(dlg,
                      text=f"Applying to {len(sel)} selected tab(s).",
                      foreground="#0045BF",
                      font=("Calibri", 10, "bold")).pack(anchor="w",
                                                         padx=12,
                                                         pady=(12, 4))
            ttk.Label(dlg,
                      text="Enter the print area to apply to all selected "
                           "tabs.\n"
                           "Examples:\n"
                           "  $A$1:$S$46         (single rectangle)\n"
                           "  $A$1:$K$45,$M$1:$X$45   (two pages on same "
                           "sheet)",
                      justify="left").pack(anchor="w", padx=12, pady=4)
            default_area = (sel[0][2].get("print_area", "$A$1:$S$46")
                            if sel else "$A$1:$S$46")
            v = tk.StringVar(value=default_area)
            ttk.Entry(dlg, textvariable=v, width=55).pack(padx=12, pady=8,
                                                         anchor="w")

            def on_ok():
                addr = v.get().strip()
                if not addr:
                    return
                is_multi = "," in addr
                cleared_with_splits = []
                for _item_id, _idx, cfg in sel:
                    cfg["print_area"] = addr
                    if is_multi:
                        # Multi-area print area + vertical page-break
                        # splits is unsafe (unpredictable page count).
                        # Force fit-1-page and clear splits. Consistency
                        # with the Click-in-Excel path. Without this,
                        # bulk Set range would silently leave stale
                        # splits in place when multi-area was chosen.
                        cfg["fit_width"] = 1
                        cfg["fit_height"] = 1
                        if cfg.get("splits"):
                            cleared_with_splits.append(cfg["name"])
                            cfg["splits"] = []
                if cleared_with_splits:
                    self._safe_log(
                        f"Cleared page-break splits on "
                        f"{len(cleared_with_splits)} tab(s) because "
                        f"multi-area print area was set: "
                        f"{', '.join(cleared_with_splits)}")
                self._refresh_all_rows(path, tv)
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.pack(fill="x", padx=12, pady=8)
            ttk.Button(btns, text="OK",
                       command=on_ok).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

        def _bulk_set_orient(self, path, orient):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            for _item_id, _idx, cfg in self._get_selected_cfgs(path, tv):
                cfg["orientation"] = orient
            self._refresh_all_rows(path, tv)

        def _bulk_set_paper(self, path, paper):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            for _item_id, _idx, cfg in self._get_selected_cfgs(path, tv):
                cfg["paper_size"] = paper
            self._refresh_all_rows(path, tv)

        def _bulk_set_fit(self, path, mode):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            for _item_id, _idx, cfg in self._get_selected_cfgs(path, tv):
                if mode == "one":
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = 1
                else:
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = FIT_UNLIMITED
            self._refresh_all_rows(path, tv)

        def _bulk_set_breaks(self, path):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            sel = self._get_selected_cfgs(path, tv)

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Set vertical page breaks for selected tabs")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("520x260")

            ttk.Label(dlg,
                      text=f"Applying to {len(sel)} selected tab(s).",
                      foreground="#0045BF",
                      font=("Calibri", 10, "bold")).pack(anchor="w",
                                                         padx=12,
                                                         pady=(12, 4))
            ttk.Label(dlg,
                      text="Enter column letter(s) where a new page should "
                           "begin.\n"
                           "Examples:\n"
                           "  L      -> 2 pages: columns A:K then L:end\n"
                           "  L, R   -> 3 pages: A:K, L:Q, R:end\n"
                           "Leave blank to remove all vertical page breaks.",
                      justify="left").pack(anchor="w", padx=12, pady=4)

            current_splits = sel[0][2].get("splits", []) if sel else []
            v = tk.StringVar(value=", ".join(current_splits))
            ttk.Entry(dlg, textvariable=v, width=30).pack(padx=12, pady=8,
                                                         anchor="w")

            def on_ok():
                raw = v.get().strip()
                splits = [s.strip().upper() for s in raw.split(",")
                          if s.strip()]
                for _item_id, _idx, cfg in sel:
                    cfg["splits"] = list(splits)
                self._refresh_all_rows(path, tv)
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.pack(fill="x", padx=12, pady=8)
            ttk.Button(btns, text="OK",
                       command=on_ok).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

        def _bulk_set_titles(self, path):
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            sel = self._get_selected_cfgs(path, tv)

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Set repeat rows for selected tabs")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("480x200")

            ttk.Label(dlg,
                      text=f"Applying to {len(sel)} selected tab(s).",
                      foreground="#0045BF",
                      font=("Calibri", 10, "bold")).pack(anchor="w",
                                                         padx=12,
                                                         pady=(12, 4))
            ttk.Label(dlg,
                      text="Enter rows to repeat at top of each printed "
                           "page.\n"
                           "Example:  1:13   (rows 1 to 13)\n"
                           "Leave blank to clear.",
                      justify="left").pack(anchor="w", padx=12, pady=4)

            current_titles = sel[0][2].get("title_rows", "") if sel else ""
            v = tk.StringVar(value=current_titles)
            ttk.Entry(dlg, textvariable=v, width=20).pack(padx=12, pady=8,
                                                         anchor="w")

            def on_ok():
                raw = v.get().strip()
                for _item_id, _idx, cfg in sel:
                    cfg["title_rows"] = raw
                self._refresh_all_rows(path, tv)
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.pack(fill="x", padx=12, pady=8)
            ttk.Button(btns, text="OK",
                       command=on_ok).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

        def _bulk_copy_from(self, path):
            """Copy all settings from a chosen source tab to all selected."""
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            sel = self._get_selected_cfgs(path, tv)
            tabs = self.workbook_data[path]
            tab_names = [t["name"] for t in tabs]

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Copy settings from another tab")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("440x200")

            ttk.Label(dlg,
                      text=f"Copy settings INTO {len(sel)} selected tab(s) "
                           "FROM:",
                      font=("Calibri", 10, "bold")).pack(anchor="w",
                                                         padx=12,
                                                         pady=(12, 4))
            ttk.Label(dlg,
                      text="(Print area, orientation, fit, repeat rows and "
                           "page breaks will be copied. The Print? state is "
                           "NOT copied.)",
                      foreground="#777",
                      wraplength=400,
                      justify="left").pack(anchor="w", padx=12, pady=2)

            v = tk.StringVar(value=tab_names[0] if tab_names else "")
            cb = ttk.Combobox(dlg, textvariable=v, values=tab_names,
                              state="readonly", width=35)
            cb.pack(padx=12, pady=8, anchor="w")

            def on_ok():
                src_name = v.get()
                src = next((t for t in tabs if t["name"] == src_name), None)
                if not src:
                    dlg.destroy()
                    return
                for _item_id, _idx, cfg in sel:
                    if cfg["name"] == src_name:
                        continue
                    cfg["print_area"] = src["print_area"]
                    cfg["orientation"] = src["orientation"]
                    cfg["paper_size"] = src.get("paper_size", xlA4)
                    cfg["fit_width"] = src["fit_width"]
                    cfg["fit_height"] = src["fit_height"]
                    cfg["title_rows"] = src["title_rows"]
                    cfg["splits"] = list(src.get("splits", []))
                self._refresh_all_rows(path, tv)
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.pack(fill="x", padx=12, pady=8)
            ttk.Button(btns, text="OK",
                       command=on_ok).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

        def _bulk_reset_auto(self, path):
            """Reset selected rows back to auto-detected defaults."""
            tv = self.current_tv
            if not self._need_selection(tv):
                return
            sel = self._get_selected_cfgs(path, tv)
            if not messagebox.askyesno(
                "Confirm reset",
                f"Reset {len(sel)} selected tab(s) to auto-detected "
                "defaults?\n\nThis discards any manual changes to print "
                "area, orientation, fit, repeat rows and page breaks."):
                return
            for _item_id, _idx, cfg in sel:
                lr = cfg.get("_detected_rows", 1)
                lc = cfg.get("_detected_cols", 1)
                reset = auto_cfg(cfg["name"], lr, lc)
                # Preserve include state - reset is about layout, not whether
                # we print it
                reset["include"] = cfg["include"]
                cfg.update(reset)
            self._refresh_all_rows(path, tv)

        # ------------------------------------------------------------------
        # Single-row Edit dialog
        # ------------------------------------------------------------------

        def _edit_tab_dialog(self, path, tv):
            sel = self._get_selected_cfgs(path, tv)
            if len(sel) != 1:
                return
            item_id, _idx, cfg = sel[0]

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title(f"Edit settings - {cfg['name']}")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("500x460")

            pad = {"padx": 8, "pady": 4}

            ttk.Label(dlg, text=f"Tab: {cfg['name']}",
                      font=("Calibri", 11, "bold")).grid(
                          row=0, column=0, columnspan=2, sticky="w", **pad)

            include_var = tk.BooleanVar(value=cfg["include"])
            ttk.Checkbutton(dlg, text="Include in PDF print",
                            variable=include_var).grid(
                                row=1, column=0, columnspan=2, sticky="w",
                                **pad)

            ttk.Label(dlg, text="Print area (e.g. $A$1:$K$45 or "
                                "A1:K45,M1:X45):").grid(
                row=2, column=0, columnspan=2, sticky="w", **pad)
            area_var = tk.StringVar(value=cfg["print_area"])
            ttk.Entry(dlg, textvariable=area_var, width=55).grid(
                row=3, column=0, columnspan=2, sticky="we", **pad)

            ttk.Label(dlg, text="Orientation:").grid(row=4, column=0,
                                                    sticky="w", **pad)
            orient_var = tk.StringVar(
                value="Portrait" if cfg["orientation"] == xlPortrait
                else "Landscape")
            ttk.Combobox(dlg, textvariable=orient_var,
                         values=["Portrait", "Landscape"],
                         state="readonly", width=15).grid(row=4, column=1,
                                                          sticky="w", **pad)

            ttk.Label(dlg, text="Paper size:").grid(row=5, column=0,
                                                   sticky="w", **pad)
            paper_var = tk.StringVar(
                value=PAPER_SIZE_LABELS.get(
                    cfg.get("paper_size", xlA4), "A4"))
            ttk.Combobox(dlg, textvariable=paper_var,
                         values=list(PAPER_SIZES.keys()),
                         state="readonly", width=15).grid(row=5, column=1,
                                                          sticky="w", **pad)

            ttk.Label(dlg, text="Fit:").grid(row=6, column=0, sticky="w",
                                             **pad)
            fit_var = tk.StringVar(value=self._fit_label(cfg))
            ttk.Combobox(dlg, textvariable=fit_var,
                         values=["Fit 1 page",
                                 "Fit width x unlimited tall"],
                         state="readonly", width=30).grid(row=6, column=1,
                                                          sticky="w", **pad)

            ttk.Label(dlg, text="Repeat rows at top of each page "
                                "(e.g. 1:13):").grid(
                row=7, column=0, columnspan=2, sticky="w", **pad)
            title_var = tk.StringVar(value=cfg["title_rows"])
            ttk.Entry(dlg, textvariable=title_var, width=20).grid(
                row=8, column=0, sticky="w", **pad)

            ttk.Label(dlg, text="Vertical page breaks (column letters, "
                                "comma-separated):").grid(
                row=9, column=0, columnspan=2, sticky="w", **pad)
            splits_var = tk.StringVar(value=", ".join(cfg.get("splits", [])))
            ttk.Entry(dlg, textvariable=splits_var, width=20).grid(
                row=10, column=0, sticky="w", **pad)

            ttk.Label(dlg,
                      text="(e.g. L splits before column L -> 2 pages: "
                           "A:K and L:end)",
                      foreground="#777").grid(row=11, column=0, columnspan=2,
                                              sticky="w", **pad)

            def on_ok():
                include = include_var.get()
                area = area_var.get().strip()

                # Refuse an empty print area on an INCLUDED tab. Excel
                # treats PrintArea = "" as 'use the entire used range',
                # which on a tab with formatted blank cells can extend
                # to row 1000+ - producing silent wrong-content PDFs,
                # which is exactly what this tool is meant to prevent.
                #
                if include and not area:
                    messagebox.showerror(
                        "Print area required",
                        "Print area cannot be empty when the tab is "
                        "marked Print. An empty print area makes Excel "
                        "use the sheet's used range, which can include "
                        "formatted blank cells and produce a PDF with "
                        "the wrong content.\n\n"
                        "Either set a specific range (e.g. $A$1:$S$46) "
                        "or untick 'Include in PDF print'.")
                    return

                cfg["include"] = include
                cfg["print_area"] = area
                cfg["orientation"] = (xlPortrait
                                      if orient_var.get() == "Portrait"
                                      else xlLandscape)
                cfg["paper_size"] = PAPER_SIZES.get(paper_var.get(), xlA4)
                if fit_var.get() == "Fit 1 page":
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = 1
                else:
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = FIT_UNLIMITED
                cfg["title_rows"] = title_var.get().strip()
                raw = splits_var.get().strip()
                new_splits = [s.strip().upper()
                              for s in raw.split(",") if s.strip()]

                # If the print area is multi-area, force fit-1-page and
                # clear vertical splits (same rule as Click-in-Excel and
                # bulk Set range - consistency was missing here, flagged
                #).
                if "," in area:
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = 1
                    if new_splits:
                        messagebox.showwarning(
                            "Splits cleared",
                            "Vertical page break splits were entered "
                            "alongside a multi-area print area. The "
                            "combination produces an unpredictable "
                            "page count. Splits have been cleared - "
                            "each area in the print range will print "
                            "as its own page.")
                        new_splits = []
                cfg["splits"] = new_splits
                self._refresh_row(tv, item_id, cfg)
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.grid(row=20, column=0, columnspan=2, pady=12)
            ttk.Button(btns, text="OK", command=on_ok).pack(side="right",
                                                            padx=4)
            ttk.Button(btns, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

        def _fit_label(self, cfg):
            fit_h = sanitise_fit_height(cfg.get("fit_height", 1))
            if cfg.get("fit_width", 1) == 1 and fit_h == 1:
                return "Fit 1 page"
            return "Fit width x unlimited tall"

        # ------------------------------------------------------------------
        # Click in Excel (single row only)
        # ------------------------------------------------------------------

        # ------------------------------------------------------------------
        # Excel session liveness / reconnect
        #
        # Excel COM is fragile. The user can close the Excel window,
        # PrintPreview can hit a COM race, an add-in can crash, etc. Once
        # that happens, every subsequent call on self.excel.app raises an
        # RPC error and the export cascades into failure.
        #
        # _ensure_excel_alive() probes the session cheaply and rebuilds it
        # if dead, reopening all known workbooks. Settings live in
        # self.workbook_data (Python-side), so they survive Excel dying.
        # ------------------------------------------------------------------

        def _is_excel_alive(self):
            if self.excel is None:
                return False
            try:
                _ = self.excel.app.Visible  # cheap probe
                return True
            except Exception:
                return False

        def _ensure_excel_alive(self):
            """Return True if Excel is responsive, else rebuild the session.

            On reconnect, .xlsm files are only reopened if macros were
            successfully disabled in the new session. Without this guard,
            a Step-2 .xlsm refusal could be silently bypassed if Excel
            died mid-export and the reconnect rebuilt with a session that
            failed AutomationSecurity.
            """
            if self._is_excel_alive():
                return True
            try:
                self._safe_log("Excel session is dead. Reconnecting...")
            except Exception:
                pass
            # Tear down stale handle
            try:
                if self.excel is not None:
                    try:
                        self.excel.quit()
                    except Exception:
                        pass
            except Exception:
                pass
            self.excel = None
            # Spin up new session and re-open all known files
            try:
                self.excel = ExcelSession(visible=False)

                # Repeat the macro-security guard from Step 2 entry.
                # If macros could not be disabled in the new session AND
                # macro-capable files are in the selection, abort
                # reconnect. Guard covers all macro-capable formats
                # (.xlsm, .xlsb, .xls) since any of them can hold
                # auto-running VBA.
                MACRO_CAPABLE_EXTS = (".xlsm", ".xlsb", ".xls")
                macro_files = [p for p in self.files
                               if p.lower().endswith(MACRO_CAPABLE_EXTS)]
                if macro_files and not self.excel.macros_disabled:
                    self._safe_log(
                        "  Reconnect ABORTED: macro security could not "
                        "be disabled and macro-capable files are in "
                        "the selection. Restart the tool and re-run.")
                    try:
                        self.excel.quit()
                    except Exception:
                        pass
                    self.excel = None
                    return False

                for path in self.files:
                    try:
                        self.excel.open(path)
                    except Exception as e:
                        self._safe_log(f"  Could not reopen "
                                       f"{os.path.basename(path)}: {e}")
                return self._is_excel_alive()
            except Exception:
                return False

        def _safe_log(self, text):
            """Log to the Step-3 log if it exists, else stderr."""
            try:
                if hasattr(self, "log") and self.log.winfo_exists():
                    self._log_line(text)
                    return
            except Exception:
                pass
            try:
                print(text, file=sys.stderr)
            except Exception:
                pass

        def _activate_workbook_visible(self, path):
            if not self._ensure_excel_alive():
                raise RuntimeError(
                    "Excel could not be started. Close any orphaned EXCEL.EXE "
                    "processes in Task Manager and try again.")
            self.excel.app.Visible = True
            self.excel.app.ScreenUpdating = True
            target = os.path.abspath(path)
            for wb in self.excel.app.Workbooks:
                try:
                    if os.path.abspath(wb.FullName) == target:
                        wb.Activate()
                        return wb
                except Exception:
                    continue
            return self.excel.open(path)

        def _click_in_excel(self, path, tv):
            """Build a print area by capturing one or more Excel selections.

            Excel's print model treats a comma-separated print area as a
            non-contiguous area: each piece prints as its own page (subject
            to fit settings). This dialog lets the user build that list one
            piece at a time, which is far more reliable than asking them to
            Ctrl-click far-apart ranges in a single Excel selection.
            """
            sel = self._get_selected_cfgs(path, tv)
            if len(sel) != 1:
                return
            item_id, _idx, cfg = sel[0]

            try:
                wb = self._activate_workbook_visible(path)
                wb.Sheets(cfg["name"]).Activate()
            except Exception as e:
                messagebox.showerror("Cannot activate tab",
                                     f"Could not activate tab "
                                     f"'{cfg['name']}': {e}")
                return

            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Click in Excel - build print area")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.geometry("620x500")

            ttk.Label(dlg, text=f"Tab: {cfg['name']}",
                      font=("Calibri", 11, "bold")).pack(padx=12,
                                                         pady=(12, 4),
                                                         anchor="w")

            msg = (
                "How to use:\n"
                "  1. Switch to Excel.\n"
                "  2. Select a range on the tab.\n"
                "  3. Click 'Add this range' below. The range is added to "
                "the list.\n"
                "  4. Repeat steps 2-3 for any additional ranges (each "
                "becomes its own page).\n"
                "  5. Click 'Apply' when done.\n\n"
                "Important: each range becomes a separate page in the "
                "PDF, so the ranges should NOT overlap. If one range "
                "sits inside another, the inner block will be printed "
                "twice. The tool will warn you on Apply if this happens."
            )
            ttk.Label(dlg, text=msg, justify="left",
                      wraplength=580).pack(padx=12, pady=4, anchor="w")

            ttk.Label(dlg, text="Captured ranges (each = one page):",
                      font=("Calibri", 10, "bold")).pack(padx=12,
                                                         pady=(8, 2),
                                                         anchor="w")

            list_frame = ttk.Frame(dlg)
            list_frame.pack(fill="both", expand=True, padx=12, pady=4)

            ranges_box = tk.Listbox(list_frame, font=("Consolas", 10),
                                    selectmode="extended", height=6)
            ranges_box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(list_frame, orient="vertical",
                               command=ranges_box.yview)
            sb.pack(side="right", fill="y")
            ranges_box.configure(yscrollcommand=sb.set)

            def _strip_sheet_prefix(piece):
                """Remove SheetName! or [Book]Sheet! prefix from one
                address piece. Excel sometimes qualifies multi-area
                addresses per-piece, so this must run on each piece
                independently. Splitting only the first ! would be
                a bug for sheet names containing exclamation marks."""
                if "!" in piece:
                    return piece.rsplit("!", 1)[1]
                return piece

            # Pre-populate from current cfg (also stripping any sheet
            # prefix that might be there from a hand-edited config)
            existing = cfg["print_area"].strip()
            if existing:
                for piece in existing.split(","):
                    piece = _strip_sheet_prefix(piece.strip())
                    if piece:
                        ranges_box.insert(tk.END, piece)

            preview_var = tk.StringVar(value="")
            ttk.Label(dlg, text="Combined print area:",
                      font=("Calibri", 9)).pack(padx=12, pady=(8, 0),
                                                anchor="w")
            ttk.Label(dlg, textvariable=preview_var,
                      font=("Consolas", 10),
                      foreground="#0045BF",
                      wraplength=580).pack(padx=12, anchor="w")

            def refresh_preview():
                pieces = list(ranges_box.get(0, tk.END))
                preview_var.set(",".join(pieces) if pieces
                                else "(empty - no print area)")

            refresh_preview()

            def on_add():
                if not self._ensure_excel_alive():
                    messagebox.showerror(
                        "Excel unavailable",
                        "Lost connection to Excel. The dialog will close.")
                    dlg.destroy()
                    return
                # Verify the user is still on the right tab. If they
                # navigated to another sheet or workbook in Excel before
                # clicking Add, the selection address would silently be
                # applied to the wrong tab. Warn and refuse.
                try:
                    active_wb = self.excel.app.ActiveWorkbook
                    active_ws = self.excel.app.ActiveSheet
                    expected_wb = os.path.basename(path)
                    actual_wb = active_wb.Name if active_wb else "(none)"
                    actual_ws = active_ws.Name if active_ws else "(none)"
                    if actual_wb != expected_wb or actual_ws != cfg["name"]:
                        if not messagebox.askyesno(
                            "Wrong sheet active",
                            f"Excel is currently on:\n"
                            f"   Workbook: {actual_wb}\n"
                            f"   Sheet:    {actual_ws}\n\n"
                            f"You are configuring:\n"
                            f"   Workbook: {expected_wb}\n"
                            f"   Sheet:    {cfg['name']}\n\n"
                            f"Reading the selection now would apply the "
                            f"WRONG sheet's address to '{cfg['name']}'.\n\n"
                            f"Click NO to switch to the correct sheet in "
                            f"Excel first (recommended).\n"
                            f"Click YES only if you know what you're "
                            f"doing."):
                            return
                except Exception:
                    pass

                try:
                    rng = self.excel.app.Selection
                    addr = rng.Address
                except Exception as e:
                    messagebox.showerror("Could not read selection", str(e))
                    return

                # Selection.Address for a multi-area selection is
                # comma-separated. Strip a sheet/workbook prefix from
                # EACH piece - not just the first one. Splitting on the
                # first ! was a bug flagged
                for piece in addr.split(","):
                    piece = _strip_sheet_prefix(piece.strip())
                    if piece:
                        ranges_box.insert(tk.END, piece)
                refresh_preview()

            def on_remove():
                sel_idx = list(ranges_box.curselection())
                sel_idx.reverse()
                for i in sel_idx:
                    ranges_box.delete(i)
                refresh_preview()

            def on_clear():
                ranges_box.delete(0, tk.END)
                refresh_preview()

            def on_apply():
                pieces = list(ranges_box.get(0, tk.END))
                if not pieces:
                    messagebox.showinfo("Nothing to apply",
                                        "Add at least one range first, or "
                                        "click Cancel.")
                    return

                # First: check whether any captured pieces are
                # unparseable by our overlap detector (whole-row
                # references like $1:$5, whole-column like $A:$D,
                # named ranges, anything malformed). If so, overlap
                # cannot be verified for them - warn explicitly rather
                # than silently skip. Note:
                # silently skipping means a whole-column piece could
                # still produce duplicate content in the PDF without
                # any warning.
                if len(pieces) > 1:
                    unparseable = [p for p in pieces
                                   if parse_a1_piece(p) is None]
                    if unparseable:
                        bullet_list = "\n".join(
                            f"  - {p}" for p in unparseable)
                        proceed = messagebox.askokcancel(
                            "Cannot verify overlap",
                            "The following captured ranges are not in "
                            "the standard $A$1:$Z$99 form, so the "
                            "tool cannot check whether they overlap "
                            "with the others:\n\n"
                            + bullet_list +
                            "\n\nIf any of these covers an area that "
                            "another range also covers, that area "
                            "will print twice in the PDF.\n\n"
                            "OK = proceed anyway.\n"
                            "Cancel = go back and replace these with "
                            "standard cell ranges.")
                        if not proceed:
                            return

                # Detect overlapping pieces. Excel's PDF exporter
                # treats a multi-area print area as one page per area,
                # so overlapping or wholly-contained ranges produce
                # duplicate content in the PDF (the contained range is
                # printed once as part of the outer range and once
                # again on its own page). Concrete example: an N3:P5
                # piece sitting wholly inside an L3:T39 piece would
                # duplicate the N3:P5 block in the output.
                if len(pieces) > 1:
                    overlaps = find_overlapping_pieces(pieces)
                    if overlaps:
                        # Build a human-readable summary
                        lines = []
                        # Indices of pieces that are wholly contained
                        # in another - candidates for safe auto-removal
                        contained_indices = set()
                        for i, j, kind in overlaps:
                            pa = pieces[i]
                            pb = pieces[j]
                            if kind == "contained":
                                lines.append(
                                    f"  - {pb} is entirely inside {pa} "
                                    f"(would print twice)")
                                contained_indices.add(j)
                            elif kind == "contains":
                                lines.append(
                                    f"  - {pa} is entirely inside {pb} "
                                    f"(would print twice)")
                                contained_indices.add(i)
                            else:
                                lines.append(
                                    f"  - {pa} and {pb} overlap "
                                    f"partially (overlap region would "
                                    f"print twice)")

                        # Decide which action to offer. If every issue
                        # is a clean "contained" case we can offer a
                        # safe auto-fix. If any is partial, auto-fix
                        # would be lossy so we only offer Cancel /
                        # Proceed.
                        all_contained = all(
                            kind in ("contained", "contains")
                            for _i, _j, kind in overlaps)

                        if all_contained:
                            # Two-button dialog: "covering range" or
                            # "Cancel". An earlier version offered a
                            # third "Print both as separate pages"
                            # button on the assumption Excel would
                            # honour a contained area as a separate
                            # page. Excel's PDF
                            # exporter does NOT reliably do this for
                            # wholly-contained areas - it can collapse
                            # them silently. Rather than expose a
                            # button that may not deliver what its
                            # label promises, the dialog now only
                            # offers the reliable action ("keep the
                            # outer covering range") or Cancel.
                            ov_dlg = tk.Toplevel(dlg)
                            add_brand_label(ov_dlg)
                            ov_dlg.title("Overlapping ranges detected")
                            ov_dlg.transient(dlg)
                            ov_dlg.grab_set()
                            ov_dlg.resizable(False, False)

                            ttk.Label(
                                ov_dlg,
                                text="The following range(s) overlap. "
                                     "Excel's PDF exporter does not "
                                     "reliably produce a separate page "
                                     "for an area that sits entirely "
                                     "inside another area - it may "
                                     "collapse them into one page. "
                                     "To avoid a confusing or missing "
                                     "page, drop the inner range and "
                                     "keep only the outer covering "
                                     "range:",
                                wraplength=560,
                                justify="left"
                            ).pack(padx=14, pady=(14, 6), anchor="w")

                            ttk.Label(
                                ov_dlg,
                                text="\n".join(lines),
                                font=("Consolas", 10),
                                foreground="#0045BF",
                                justify="left"
                            ).pack(padx=14, pady=(0, 8), anchor="w")

                            ttk.Label(
                                ov_dlg,
                                text="If you genuinely want the inner "
                                     "block on its own page, cancel "
                                     "and split your workbook layout "
                                     "so the inner block sits in a "
                                     "place not covered by the outer "
                                     "range.",
                                wraplength=560,
                                justify="left",
                                foreground="#666"
                            ).pack(padx=14, pady=(0, 10), anchor="w")

                            # Choice slot for the button handlers to
                            # write into. Using a list to avoid Python
                            # closure-rebinding gotchas with a plain
                            # variable.
                            ov_choice = [None]

                            def _ov_cover():
                                ov_choice[0] = "cover"
                                ov_dlg.destroy()

                            def _ov_cancel():
                                ov_choice[0] = "cancel"
                                ov_dlg.destroy()

                            ov_btns = ttk.Frame(ov_dlg)
                            ov_btns.pack(fill="x", padx=14,
                                         pady=(4, 14))
                            ttk.Button(
                                ov_btns,
                                text="Keep only the covering range",
                                command=_ov_cover
                            ).pack(side="left", padx=4)
                            ttk.Button(
                                ov_btns, text="Cancel",
                                command=_ov_cancel
                            ).pack(side="right", padx=4)

                            # If the user closes the dialog via the X,
                            # treat as Cancel
                            ov_dlg.protocol("WM_DELETE_WINDOW",
                                            _ov_cancel)

                            # Centre on parent and wait
                            ov_dlg.update_idletasks()
                            dlg.wait_window(ov_dlg)

                            choice = ov_choice[0]
                            if choice in (None, "cancel"):
                                return
                            # Only choice that proceeds is "cover":
                            # drop contained pieces, keep covers
                            kept = [
                                p for k, p in enumerate(pieces)
                                if k not in contained_indices
                            ]
                            ranges_box.delete(0, tk.END)
                            for p in kept:
                                ranges_box.insert(tk.END, p)
                            refresh_preview()
                            pieces = kept

                        else:
                            # Partial overlap - cannot safely identify
                            # a single covering range, so only offer
                            # Proceed-anyway / Cancel.
                            proceed = messagebox.askokcancel(
                                "Overlapping ranges detected",
                                "The following range(s) overlap. Excel "
                                "treats each captured range as a "
                                "separate page in the PDF, so the "
                                "overlapping area(s) will appear "
                                "twice in the output:\n\n"
                                + "\n".join(lines) +
                                "\n\nOK = proceed as-is (the "
                                "overlapping area will print on "
                                "multiple pages).\n"
                                "Cancel = go back and edit manually.\n\n"
                                "Note: there is no single 'covering' "
                                "range to keep automatically, because "
                                "the ranges overlap only partially. "
                                "If you want a single combined block, "
                                "cancel, then capture one larger "
                                "range that contains everything you "
                                "need.")
                            if not proceed:
                                return

                addr = ",".join(pieces)
                cfg["print_area"] = addr
                if len(pieces) > 1:
                    # Multi-area = each piece is one page. Force fit-to-1
                    # and clear any existing vertical page-break splits -
                    # combining multi-area print area WITH vertical
                    # splits produces an unpredictable number of pages.
                    cfg["fit_width"] = 1
                    cfg["fit_height"] = 1
                    if cfg.get("splits"):
                        self._safe_log(
                            f"Cleared {len(cfg['splits'])} page-break "
                            f"split(s) on '{cfg['name']}' because "
                            f"multi-area print area was set.")
                        cfg["splits"] = []
                self._refresh_row(tv, item_id, cfg)
                dlg.destroy()

            btn_row1 = ttk.Frame(dlg)
            btn_row1.pack(fill="x", padx=12, pady=(8, 4))
            ttk.Button(btn_row1, text="Add this range from Excel",
                       command=on_add).pack(side="left", padx=2)
            ttk.Button(btn_row1, text="Remove selected",
                       command=on_remove).pack(side="left", padx=2)
            ttk.Button(btn_row1, text="Clear all",
                       command=on_clear).pack(side="left", padx=2)

            btn_row2 = ttk.Frame(dlg)
            btn_row2.pack(fill="x", padx=12, pady=(4, 12))
            ttk.Button(btn_row2, text="Apply",
                       command=on_apply).pack(side="right", padx=4)
            ttk.Button(btn_row2, text="Cancel",
                       command=dlg.destroy).pack(side="right", padx=4)

            dlg.lift()

        # ------------------------------------------------------------------
        # Preview in Excel
        # ------------------------------------------------------------------

        def _preview_in_excel(self, path, tv):
            sel = self._get_selected_cfgs(path, tv)
            if len(sel) != 1:
                return
            _item_id, _idx, cfg = sel[0]
            preview_warnings = []
            try:
                wb = self._activate_workbook_visible(path)
                ws = wb.Sheets(cfg["name"])
                ws.Activate()
                footer = self.app_config.get("left_footer", "")
                preview_warnings = self._apply_tab_setup(
                    ws, cfg, footer) or []
            except PrintCommunicationError as e:
                # Application-wide state is unsafe. Preview reaches
                # _apply_tab_setup by a different route from the export
                # loop, so without this branch it would return normally
                # and leave every later action - including the export -
                # writing page setup into a cache that is never
                # committed. Drop the session so the next action
                # reconnects to a clean Excel Application.
                try:
                    if self.excel is not None:
                        self.excel.quit()
                except Exception:
                    pass
                self.excel = None
                messagebox.showerror(
                    "Excel print state unrecoverable",
                    f"{e}\n\nThe Excel session has been closed. Reopen "
                    f"the workbook and try again before exporting.")
                return
            except Exception as e:
                messagebox.showerror("Preview failed",
                                     f"Could not apply settings to "
                                     f"'{cfg['name']}': {e}")
                return

            if preview_warnings:
                messagebox.showwarning(
                    "Settings applied with warnings",
                    "Settings applied for preview, but with the "
                    "following issue(s):\n\n"
                    + "\n".join(f"  - {w}" for w in preview_warnings)
                    + "\n\nThe PDF export will surface these in the log.")

            # ws.PrintPreview() is the proper modal preview, but Excel's COM
            # bridge frequently raises RPC_S_CALL_FAILED (-2147023170) on it,
            # and can occasionally take Excel down with it. Try it, and if
            # it fails, just leave Excel visible on the configured sheet so
            # the user can press Ctrl+F2 themselves.
            try:
                ws.PrintPreview()
            except Exception:
                if not self._is_excel_alive():
                    messagebox.showerror(
                        "Excel closed during preview",
                        "Excel closed unexpectedly during Print Preview.\n\n"
                        "Click OK and the tool will reconnect on the next "
                        "action.")
                    return
                messagebox.showinfo(
                    "Preview ready in Excel",
                    f"Settings applied to '{cfg['name']}'.\n\n"
                    "The native Print Preview call failed (Excel COM is "
                    "fragile with it), but Excel is now showing the tab "
                    "with the settings applied.\n\n"
                    "Press Ctrl+F2 in Excel, or File > Print, to see the "
                    "preview manually.")

        # ------------------------------------------------------------------
        # Save config to workbook
        # ------------------------------------------------------------------

        def _save_workbook_configs(self):
            failures = []
            successes = []
            for path, tabs in self.workbook_data.items():
                clean = []
                for t in tabs:
                    clean.append({k: v for k, v in t.items()
                                  if not k.startswith("_")})
                # Include the per-workbook footer if one is set. This
                # is persisted alongside the tab config so reopening
                # the same workbook next month pre-fills the same
                # footer rather than reverting to the global default.
                footer = self.workbook_footers.get(path)
                sidecar_data = {
                    "tabs": clean,
                    "saved_at": datetime.now().isoformat(),
                }
                if footer is not None:
                    sidecar_data["left_footer"] = footer
                ok, err = save_workbook_config(path, sidecar_data)
                if ok:
                    successes.append(path)
                else:
                    failures.append((path, err))

            if not failures:
                messagebox.showinfo(
                    "Saved",
                    f"Print configuration saved next to "
                    f"{len(successes)} workbook(s).\n\n"
                    f"Next month, opening the same files will auto-load "
                    f"these settings.")
                return

            # Note: an earlier approach showed 'Saved' even when writes
            # failed, which defeated the whole point of the sidecar.
            # Surface real failures clearly so the user knows their
            # settings did not persist.
            lines = [f"  - {os.path.basename(p)}: {err}"
                     for p, err in failures]
            messagebox.showerror(
                "Save failed",
                f"Saved {len(successes)} of "
                f"{len(successes) + len(failures)} workbook(s).\n\n"
                f"FAILED:\n" + "\n".join(lines) +
                "\n\nLikely causes: read-only folder, locked sidecar "
                "file open in another program, OneDrive sync conflict, "
                "or insufficient permissions.")

        # ------------------------------------------------------------------
        # Workbook navigation
        # ------------------------------------------------------------------

        def _next_workbook(self):
            if self.current_wb_index < len(self.files) - 1:
                self._show_workbook(self.current_wb_index + 1)

        def _prev_workbook(self):
            if self.current_wb_index > 0:
                self._show_workbook(self.current_wb_index - 1)

        def _back_to_step1(self):
            """Return to file list, quitting Excel along the way.

            An earlier approach just called _build_step1() which left the
            Excel session running in the background. If the user then
            closed the app from the Step 1 screen via a path other than
            the window-close handler, the Excel process would be
            orphaned.
            """
            try:
                if self.excel is not None:
                    self.excel.quit()
            except Exception:
                pass
            self.excel = None
            self._build_step1()

        # ------------------------------------------------------------------
        # STEP 3 - Apply and export
        # ------------------------------------------------------------------

        def _goto_step3(self):
            self._build_step3()

        def _build_step3(self):
            self._clear()
            self.root.title("Excel Quick Print - Step 3 of 3")

            frm = ttk.Frame(self.root, padding=20)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="Step 3.  Apply settings and export PDFs",
                      font=("Calibri", 13, "bold")).pack(anchor="w")

            ttk.Label(frm,
                      text="Excel's native PDF engine will be used. Same "
                           "quality as File > Save As > PDF, with your real "
                           "fonts.",
                      foreground="#555").pack(anchor="w", pady=(0, 14))

            # Save-after-export toggle. Default ON: the user almost always
            # wants their applied page setup persisted into the .xlsx so
            # the same settings are visible on opening the workbook next.
            # Off when they want a non-destructive print only.
            self.save_workbook_after_export = self.app_config.get(
                "save_workbook_after_export", True)
            # Companion mode: never write _rev.NN workbook copies into the
            # bundle folders (the bundle builder would pick them up as
            # further workbooks on the next run). The setting is not
            # persisted from here, so the standalone default is untouched.
            companion_no_copy = bool(self.companion
                                     and not self.companion["workbook_copy"])
            if companion_no_copy:
                self.save_workbook_after_export = False
            self._save_wb_var = tk.BooleanVar(
                value=self.save_workbook_after_export)

            def _on_save_toggle():
                self.save_workbook_after_export = self._save_wb_var.get()
                if companion_no_copy:
                    return
                self.app_config["save_workbook_after_export"] = \
                    self.save_workbook_after_export
                save_app_config(self.app_config)

            save_cb = ttk.Checkbutton(
                frm,
                text="Save applied page setup as a new _rev.NN workbook "
                     "copy (original is preserved)",
                variable=self._save_wb_var,
                command=_on_save_toggle)
            save_cb.pack(anchor="w", pady=(0, 10))
            if companion_no_copy:
                save_cb.configure(state="disabled")
                ttk.Label(frm, text="(disabled in companion mode - page "
                                    "setup is still saved to each "
                                    "workbook's .printcfg.json)",
                          foreground="#777").pack(anchor="w", pady=(0, 8))

            self.progress_var = tk.StringVar(value="Ready.")
            ttk.Label(frm, textvariable=self.progress_var,
                      font=("Consolas", 10),
                      foreground="#0045BF").pack(anchor="w", pady=8)

            self.log = tk.Text(frm, height=18, font=("Consolas", 9),
                               wrap="none", state="disabled")
            self.log.pack(fill="both", expand=True, pady=(4, 8))

            nav = ttk.Frame(frm)
            nav.pack(fill="x")
            self.back_btn = ttk.Button(
                nav, text="< Back",
                command=self._safe_back_from_step3)
            self.back_btn.pack(side="left")
            self.start_btn = ttk.Button(nav, text="Export to PDF",
                                        command=self._start_export)
            self.start_btn.pack(side="right")
            self.open_btn = ttk.Button(nav, text="Open output folder",
                                       command=self._open_output_folder,
                                       state="disabled")
            self.open_btn.pack(side="right", padx=8)
            # _exporting is initialised in __init__; do not reset it
            # here. An earlier approach reset it on every Step-3
            # rebuild, which was harmless in practice but contradicted
            # the brief and looked confusing.

            add_brand_label(self.root)

        def _on_close_window(self):
            if getattr(self, "_exporting", False):
                messagebox.showwarning(
                    "Export running",
                    "An export is in progress. Wait for it to finish.")
                return
            try:
                if self.excel is not None:
                    self.excel.quit()
                    self.excel = None
            except Exception:
                pass
            self.root.destroy()

        def _log_line(self, text):
            self.log.configure(state="normal")
            self.log.insert(tk.END, text + "\n")
            self.log.see(tk.END)
            self.log.configure(state="disabled")
            # update_idletasks() only processes redraws and idle handlers,
            # NOT pending button clicks or other input events. The earlier
            # version called self.root.update() which processed everything
            # and let the user click 'Back' mid-export, causing widget
            # destruction while COM work was still running. Subtle but
            # nasty reentrancy hazard - flagged
            self.root.update_idletasks()

        def _safe_back_from_step3(self):
            """Return to the per-workbook configurator from Step 3.

            Wrapped in explicit error capture because field testing
            reported "an error" when clicking Back after a failed
            export, without specifics. If _show_workbook raises (e.g.
            because some bit of state was clobbered during the export
            failure path), surfacing the exception in a dialog is
            strictly better than the alternative of a silent crash
            where the user has no idea what happened.
            """
            try:
                # Re-enable the buttons we disabled at export start,
                # in case some path missed it.
                try:
                    self.start_btn.configure(state="normal")
                except Exception:
                    pass
                self._show_workbook(self.current_wb_index)
            except Exception as e:
                tb = traceback.format_exc()
                # Cap the traceback so the dialog stays readable
                tb_short = tb if len(tb) < 2000 else tb[-2000:]
                messagebox.showerror(
                    "Could not return to configurator",
                    f"Going back to the configurator failed:\n\n"
                    f"{e}\n\n"
                    f"Details:\n{tb_short}\n\n"
                    f"As a workaround, close and reopen the tool. "
                    f"Your saved per-workbook settings (the "
                    f".printcfg.json files next to each workbook) "
                    f"are preserved.")

        def _collect_workbook_footers(self):
            """Show a single dialog listing all workbooks and let the
            user set the left-side PDF footer for each one. Returns
            True on OK, False on Cancel. Populates
            self.workbook_footers.

            Earlier the tool used one global footer for every workbook
            in the batch. User requirement: each workbook can have its
            own footer. We collect them all up front (rather than
            interrupting mid-export) so the export run itself stays
            uninterrupted.
            """
            dlg = tk.Toplevel(self.root)
            add_brand_label(dlg)
            dlg.title("Footer text per workbook")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.resizable(True, True)
            dlg.geometry("760x480")

            ttk.Label(
                dlg,
                text="Set the left-side PDF footer for each workbook. "
                     "The right side of every page always shows "
                     "'Page X of Y'.",
                wraplength=680, justify="left",
            ).pack(padx=14, pady=(14, 4), anchor="w")

            ttk.Label(
                dlg,
                text="Pre-filled with: per-workbook saved footer if "
                     "previously saved, otherwise the global default.\n"
                     "  Save and continue = persist each workbook's "
                     "footer into its .printcfg.json sidecar, so the "
                     "same workbook reopened later pre-fills its own "
                     "specific footer.\n"
                     "  OK (use for this run) = apply for this export. "
                     "The first row's footer becomes the global default "
                     "for new workbooks next session, but per-workbook "
                     "sidecars are NOT updated.",
                wraplength=680, justify="left",
                foreground="#666",
            ).pack(padx=14, pady=(0, 8), anchor="w")

            # Scrollable list of (filename, entry) pairs
            outer = ttk.Frame(dlg)
            outer.pack(fill="both", expand=True, padx=14, pady=4)

            canvas = tk.Canvas(outer, highlightthickness=0)
            scrollbar = ttk.Scrollbar(outer, orient="vertical",
                                      command=canvas.yview)
            inner = ttk.Frame(canvas)
            inner.bind(
                "<Configure>",
                lambda e: canvas.configure(
                    scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # Mouse-wheel scrolling. On Windows, tk.Canvas does not
            # bind <MouseWheel> by default, so the user can only scroll
            # via the scrollbar - which feels broken with 20+ workbooks
            # to set footers for. Bind on Enter / unbind on Leave so we
            # only capture wheel events while the cursor is over the
            # canvas (otherwise we'd hijack wheel events for any
            # widget under the cursor). Note: the
            # missing binding.
            def _on_mousewheel(event):
                # On Windows, event.delta is +-120 per notch
                canvas.yview_scroll(int(-event.delta / 120), "units")

            def _bind_wheel(_e):
                canvas.bind_all("<MouseWheel>", _on_mousewheel)

            def _unbind_wheel(_e):
                canvas.unbind_all("<MouseWheel>")

            canvas.bind("<Enter>", _bind_wheel)
            canvas.bind("<Leave>", _unbind_wheel)
            # Also clean up on dialog destruction
            dlg.bind("<Destroy>",
                     lambda _e: canvas.unbind_all("<MouseWheel>"),
                     add="+")

            default_footer = self.app_config.get(
                "left_footer", "")
            entry_vars = {}
            for i, path in enumerate(self.files):
                row = ttk.Frame(inner)
                row.pack(fill="x", pady=4)
                ttk.Label(row, text=os.path.basename(path),
                          width=50, anchor="w").pack(side="left")
                # Reuse a previously-set footer for this workbook if
                # one exists (handles back-and-forth navigation)
                initial = self.workbook_footers.get(path, default_footer)
                v = tk.StringVar(value=initial)
                entry_vars[path] = v
                ttk.Entry(row, textvariable=v, width=40).pack(
                    side="left", padx=8, fill="x", expand=True)

            # Buttons
            choice = {"ok": False, "saved": False}

            def _apply_all():
                """Apply the topmost entry to all rows - convenience."""
                if not self.files:
                    return
                first = entry_vars[self.files[0]].get()
                for p in self.files:
                    entry_vars[p].set(first)

            def _capture_to_state():
                """Pull the entries into self.workbook_footers and
                persist the first one as the new global default.
                Shared by both OK and Save-and-continue paths."""
                for p, v in entry_vars.items():
                    self.workbook_footers[p] = v.get()
                # Persist the first workbook's footer as the new
                # default so next session pre-fills sensibly
                if self.files:
                    self.app_config["left_footer"] = entry_vars[
                        self.files[0]].get()
                    save_app_config(self.app_config)

            def _on_ok():
                _capture_to_state()
                choice["ok"] = True
                dlg.destroy()

            def _on_save_and_continue():
                """Persist the per-workbook footers to each workbook's
                sidecar JSON so they survive across sessions, then
                continue with the export.

                The sidecar is what _analyse_workbook reads on next
                open. Writing the footer there means reopening the
                same workbook next month pre-fills the same footer
                rather than reverting to the global default.
                """
                _capture_to_state()
                # Save each workbook's footer into its sidecar. We
                # write the existing tab config too so the sidecar
                # stays internally consistent (no half-update where
                # the footer is fresh but the tabs are stale).
                save_failures = []
                for path, tabs in self.workbook_data.items():
                    clean = []
                    for t in tabs:
                        clean.append({k: v for k, v in t.items()
                                      if not k.startswith("_")})
                    sidecar = {
                        "tabs": clean,
                        "saved_at": datetime.now().isoformat(),
                        "left_footer": self.workbook_footers.get(
                            path, ""),
                    }
                    ok, err = save_workbook_config(path, sidecar)
                    if not ok:
                        save_failures.append((path, err))
                if save_failures:
                    lines = [f"  - {os.path.basename(p)}: {err}"
                             for p, err in save_failures]
                    messagebox.showwarning(
                        "Footer save partial",
                        f"Could not save footer to sidecar for "
                        f"{len(save_failures)} of "
                        f"{len(self.workbook_data)} workbook(s):"
                        f"\n\n" + "\n".join(lines) +
                        "\n\nThe footers will still apply for THIS "
                        "run. Causes: read-only folder, sidecar "
                        "file locked, OneDrive sync conflict.",
                        parent=dlg)
                choice["ok"] = True
                choice["saved"] = True
                dlg.destroy()

            def _on_cancel():
                dlg.destroy()

            btns = ttk.Frame(dlg)
            btns.pack(fill="x", padx=14, pady=(4, 14))
            ttk.Button(btns, text="Apply first row to all",
                       command=_apply_all).pack(side="left")
            ttk.Button(btns, text="Cancel",
                       command=_on_cancel).pack(side="right", padx=4)
            ttk.Button(btns, text="OK (use for this run)",
                       command=_on_ok).pack(side="right", padx=4)
            ttk.Button(btns,
                       text="Save and continue",
                       command=_on_save_and_continue).pack(
                           side="right", padx=4)

            dlg.protocol("WM_DELETE_WINDOW", _on_cancel)
            self.root.wait_window(dlg)
            return choice["ok"]

        def _start_export(self):
            # Collect per-workbook footers before we lock the UI for
            # export. If the user cancels, abort cleanly.
            if not self._collect_workbook_footers():
                return

            self._exporting = True
            self.start_btn.configure(state="disabled")
            self.back_btn.configure(state="disabled")
            self._output_dir = os.path.dirname(self.files[0])
            self.progress_var.set("Working...")
            self.root.update_idletasks()

            results = []
            try:
                if not self._ensure_excel_alive():
                    raise RuntimeError(
                        "Could not start Excel. Close any orphaned EXCEL.EXE "
                        "processes in Task Manager and try again.")
                self._log_line("Excel session ready.")

                try:
                    self.excel.app.Visible = False
                    self.excel.app.ScreenUpdating = False
                except Exception:
                    pass

                for path in self.files:
                    result = self._export_workbook(path)
                    results.append(result)

            except Exception as e:
                tb = traceback.format_exc()
                self.progress_var.set("ERROR.")
                self._log_line(f"\n[ERROR] {e}\n{tb}")
                if self.companion:
                    write_companion_result(self.companion, "error", results,
                                           error=str(e))
                messagebox.showerror("Export failed", str(e))
                self._exporting = False
                self.back_btn.configure(state="normal")
                try:
                    if self.excel is not None:
                        self.excel.quit()
                        self.excel = None
                except Exception:
                    pass
                return
            finally:
                self._exporting = False

            # Summarise results. Important: distinguish full success
            # from partial success and from total failure. Silent
            # partial success (e.g. "Done. PDFs created." while some
            # workbooks failed) would be the biggest hazard - the user
            # could ship a deliverable that's missing files without
            # knowing.
            ok_count = sum(1 for r in results if r["ok"])
            fail_count = len(results) - ok_count
            warn_count = sum(1 for r in results
                             if r["ok"] and r.get("warning"))
            self._log_line("")
            self._log_line("=" * 60)
            self._log_line(f"SUMMARY: {ok_count} succeeded, "
                           f"{fail_count} failed out of {len(results)}"
                           + (f"  ({warn_count} with warnings)"
                              if warn_count else ""))
            for r in results:
                base = os.path.basename(r["path"])
                if r["ok"]:
                    line = (f"  OK    {base}  ->  "
                            f"{os.path.basename(r['pdf_path'])}")
                    if r.get("warning"):
                        line += "  [warning]"
                    self._log_line(line)
                    if r.get("warning"):
                        self._log_line(f"        WARNING: {r['warning']}")
                else:
                    self._log_line(f"  FAIL  {base}  ({r['error']})")
                    # Surface PDF paths even on failure - the user
                    # needs to know if a (possibly correct) PDF is
                    # sitting somewhere and shouldn't be confused with
                    # a stale file at the expected name.
                    if r.get("pdf_path"):
                        self._log_line(
                            f"        Note: a PDF was produced at "
                            f"{r['pdf_path']} despite the failure. "
                            f"Verify before using.")
                    if r.get("recovery_pdf_path"):
                        self._log_line(
                            f"        Recovery copy at: "
                            f"{r['recovery_pdf_path']}")

            if fail_count == 0 and warn_count == 0:
                self.progress_var.set(f"Done. {ok_count} PDF(s) created.")
                self.open_btn.configure(state="normal")
            elif fail_count == 0 and warn_count > 0:
                self.progress_var.set(
                    f"Done with warnings: {ok_count} PDF(s) created, "
                    f"{warn_count} with issues. See log.")
                self.open_btn.configure(state="normal")
                messagebox.showwarning(
                    "Export completed with warnings",
                    f"{ok_count} PDF(s) created, {warn_count} with "
                    f"warnings (e.g. workbook could not be saved). "
                    f"Review the log.")
            elif ok_count == 0:
                self.progress_var.set(
                    f"FAILED. No PDFs were produced.")
                messagebox.showerror(
                    "Export failed",
                    "No PDFs were produced. See the log for details.")
            else:
                self.progress_var.set(
                    f"PARTIAL: {ok_count} OK, {fail_count} FAILED. "
                    f"See log.")
                self.open_btn.configure(state="normal")
                messagebox.showwarning(
                    "Partial export",
                    f"{ok_count} workbook(s) succeeded, "
                    f"{fail_count} failed. Review the log carefully - "
                    f"some workbooks were NOT exported.")

            # Re-enable Back so the user can go back to the configurator
            # if they want to retry. Keep Export disabled to make them go
            # back through the wizard explicitly.
            self.back_btn.configure(state="normal")

            # Companion mode: hand the results back to the caller now, so
            # the bundle build can continue as soon as this window closes.
            if self.companion:
                self._companion_results = results
                if write_companion_result(self.companion, "done", results):
                    self._log_line("")
                    self._log_line(
                        "Companion result written for the Document Bundle "
                        "Builder. Close this window to continue the bundle.")
                else:
                    self._log_line(
                        "WARNING: could not write the companion result "
                        f"file at {self.companion['result_path']}.")

            # Safety net. PrintCommunication is an Application-wide
            # setting, not a per-workbook one, so if a per-tab restore was
            # ever missed the user's own Excel session would be left with
            # page setup suspended. Force it back on before teardown.
            try:
                if self.excel is not None and self.excel.app is not None:
                    self.excel.app.PrintCommunication = True
            except Exception:
                pass

            # Tear down Excel session now that we're done
            try:
                if self.excel is not None:
                    self.excel.quit()
                    self.excel = None
            except Exception:
                pass

        # ------------------------------------------------------------------
        # Export one workbook
        # ------------------------------------------------------------------

        def _export_workbook(self, path):
            """Apply per-tab settings, hide non-included sheets, export.

            Returns a result dict:
                {
                    "path": str,
                    "ok": bool,            # True only if PDF was written
                                           # AND visibility restored OK
                    "pdf_path": str|None,  # set even when ok=False if a
                                           # PDF was produced (e.g. restore
                                           # failed) so user knows the file
                                           # exists
                    "skipped_tabs": [str], # tabs the user excluded
                    "failed_tabs": [(name, error)],  # tab setup failures
                    "setup_warnings": [(tab, warning)],  # non-fatal
                                           # page-setup issues per tab
                    "saved": bool|str,     # True / False / "skipped (...)"
                                           # / "error: ..."
                    "error": str|None,     # top-level failure reason
                    "warning": str|None,   # non-fatal aggregated warning
                                           # shown in summary (read-only,
                                           # save failure, setup warnings)
                    "read_only_warning": str|None,  # internal: stashed
                                           # when workbook opened read-only
                                           # so it can be folded into
                                           # `warning` later
                }

            FAILURE POLICY (fail-loud):
              - If ANY included tab fails _apply_tab_setup, abort the
                workbook. Do not export, do not save.
              - If hiding a non-included sheet fails, abort. Do not export.
              - If restoring visibility fails after export, mark the result
                failed even if the PDF wrote successfully (because the
                workbook is now in an unpredictable state and saving it
                would persist that).
              - wb.Save() runs ONLY if export succeeded AND visibility
                restored cleanly AND save is enabled AND workbook is not
                read-only.
            """
            result = {
                "path": path,
                "ok": False,
                "pdf_path": None,
                "skipped_tabs": [],
                "failed_tabs": [],
                "setup_warnings": [],  # non-fatal page-setup issues
                                        # (e.g. paper-size fallback, a
                                        # specific page-break couldn't
                                        # be added)
                "saved": False,
                "error": None,
                "warning": None,  # non-fatal aggregated warning shown
                                  # in summary (read-only, save failure,
                                  # setup warnings)
                "read_only_warning": None,  # internal staging field
                "recovery_pdf_path": None,  # set if PDF was produced
                                            # but could not be moved to
                                            # the target name (e.g.
                                            # target locked in Acrobat).
                                            # User must rescue manually.
            }

            self._log_line(f"\n=== {os.path.basename(path)} ===")

            # Phase-level timing. Recommended: measure-first
            # before any optimisation work, and the user has asked
            # whether the tool can be made faster. Rather than guess,
            # log how long each major phase takes so we have data to
            # point at. Negligible runtime overhead (4-5 time.time()
            # calls per workbook).
            import time as _time
            _t_phase_start = _time.time()
            _phase_log = []   # list of (phase_name, elapsed_seconds)

            def _phase_done(name):
                nonlocal _t_phase_start
                now = _time.time()
                elapsed = now - _t_phase_start
                _phase_log.append((name, elapsed))
                _t_phase_start = now

            # Reconnect if Excel died between workbooks
            if not self._ensure_excel_alive():
                result["error"] = "Excel session unavailable"
                self._log_line("  ABORT: Excel not available.")
                return result

            wb = None
            target = os.path.abspath(path)
            try:
                for w in self.excel.app.Workbooks:
                    try:
                        if os.path.abspath(w.FullName) == target:
                            wb = w
                            break
                    except Exception:
                        continue
            except Exception:
                wb = None
            if wb is None:
                try:
                    wb = self.excel.open(path)
                except Exception as e:
                    result["error"] = f"Could not open workbook: {e}"
                    self._log_line(f"  ABORT: could not open workbook: {e}")
                    return result
            _phase_done("workbook_open")

            # Detect read-only. Beyond "Save will fail", this also means
            # the on-disk file may be older than what's currently in
            # another Excel session - so the exported PDF could be stale
            # vs the user's unsaved edits.
            read_only = False
            try:
                read_only = bool(wb.ReadOnly)
            except Exception:
                pass
            if read_only:
                self._log_line(
                    "  WARNING: workbook is open as READ-ONLY. This usually "
                    "means it is already open in another Excel window. The "
                    "PDF will be made from the LAST SAVED version on disk - "
                    "any unsaved edits in your other Excel window will NOT "
                    "be in the PDF. Page-setup changes also cannot be saved "
                    "back. Close the other window and re-run if either "
                    "matters.")
                # Critical: stash this so the final summary flags it.
                # Earlier rounds only logged the read-only warning, which
                # meant a stale PDF would still be reported as "Done" in
                # the summary.
                result["read_only_warning"] = (
                    "workbook was open as read-only; PDF was made from "
                    "the last saved disk version - any unsaved edits in "
                    "another Excel window are NOT included")

            # Per-workbook footer (set in the Step-3 prompt). Falls
            # back to the global default if the workbook somehow
            # didn't get one (defensive).
            footer_left = self.workbook_footers.get(
                path,
                self.app_config.get("left_footer", ""))

            # Set author metadata BEFORE export so the PDF inherits it.
            # This is intentional and required (file-authorship-metadata
            # rule). If it fails we still proceed - the PDF will then
            # show whoever Excel's user is set to.
            try:
                wb.BuiltinDocumentProperties("Author").Value = "Alaa Elsayed"
                wb.BuiltinDocumentProperties("Last Author").Value = \
                    "Alaa Elsayed"
            except Exception as e:
                self._log_line(f"  Note: could not set author metadata: {e}")

            # PHASE 1: apply per-tab setup. Any failure here is fatal.
            included_names = set()
            xlWorksheet = -4167
            setup_warnings = []
            for cfg in self.workbook_data[path]:
                if not cfg["include"]:
                    self._log_line(f"  SKIP {cfg['name']}")
                    result["skipped_tabs"].append(cfg["name"])
                    continue
                try:
                    ws = wb.Sheets(cfg["name"])
                    # Refuse non-worksheets explicitly. Previously, if
                    # reading ws.Type itself raised, we proceeded -
                    # which defeated the guard. Safer to refuse.
                    try:
                        sheet_type = ws.Type
                    except Exception as e:
                        raise RuntimeError(
                            f"could not read sheet type "
                            f"(refusing to setup non-worksheets): {e}")
                    if sheet_type != xlWorksheet:
                        raise RuntimeError(
                            "not a worksheet (chart/dialog sheet)")
                    warnings_for_tab = self._apply_tab_setup(
                        ws, cfg, footer_left)
                    orient_l = ("P" if cfg["orientation"] == xlPortrait
                                else "L")
                    self._log_line(f"  OK   {cfg['name']}  "
                                   f"[{cfg['print_area']}, {orient_l}]")
                    if warnings_for_tab:
                        for w in warnings_for_tab:
                            self._log_line(f"       WARN {cfg['name']}: {w}")
                            setup_warnings.append(
                                (cfg["name"], w))
                    included_names.add(cfg["name"])
                except PrintCommunicationError as e:
                    # NOT a per-tab failure. PrintCommunication is
                    # Application-wide, so continuing to the next tab
                    # would configure it in a state where PageSetup
                    # writes are cached and never committed - producing a
                    # PDF that silently ignores the configuration. Log
                    # and re-raise so the run-level handler aborts the
                    # whole export and tears down the Excel session.
                    self._log_line(f"  FATAL {cfg['name']}: {e}")
                    self._log_line(
                        "  ABORT: stopping the entire export. Remaining "
                        "tabs and workbooks have NOT been processed.")
                    raise
                except Exception as e:
                    self._log_line(f"  FAIL {cfg['name']}: {e}")
                    result["failed_tabs"].append((cfg["name"], str(e)))

            if result["failed_tabs"]:
                msg = (f"{len(result['failed_tabs'])} tab(s) failed setup. "
                       f"ABORTING workbook (no partial PDF). Fix the failed "
                       f"tabs and re-run.")
                self._log_line(f"  ABORT: {msg}")
                result["error"] = msg
                return result

            # Stash setup warnings on the result so the summary can show
            # the workbook as "OK with warnings" rather than silent
            # success.
            if setup_warnings:
                result["setup_warnings"] = setup_warnings

            if not included_names:
                self._log_line("  (no tabs marked for printing)")
                result["error"] = "no tabs marked for printing"
                return result

            _phase_done("page_setup")

            # PHASE 2: activate this workbook
            try:
                wb.Activate()
            except Exception as e:
                result["error"] = f"could not activate workbook: {e}"
                self._log_line(f"  ABORT: {result['error']}")
                return result

            # PHASE 3: hide non-included sheets. Any hide failure aborts
            # the export (silently failing here would put excluded sheets
            # into the PDF).
            original_visibility = {}
            hide_failed = None
            try:
                for i in range(1, wb.Sheets.Count + 1):
                    ws = wb.Sheets(i)
                    name = ws.Name
                    vis = ws.Visible
                    original_visibility[name] = vis
                    if vis == xlSheetVisible and name not in included_names:
                        try:
                            ws.Visible = xlSheetHidden
                        except Exception as e:
                            hide_failed = (name, str(e))
                            break
            except Exception as e:
                hide_failed = ("?", str(e))

            if hide_failed:
                name, err = hide_failed
                # Restore before bailing out
                for n, vis in original_visibility.items():
                    try:
                        wb.Sheets(n).Visible = vis
                    except Exception:
                        pass
                msg = (f"could not hide sheet '{name}' before export: "
                       f"{err}. ABORTING (would otherwise include "
                       f"excluded sheets in the PDF).")
                self._log_line(f"  ABORT: {msg}")
                result["error"] = msg
                return result

            _phase_done("hide_sheets")

            # PHASE 4: export. Inside a try so we can guarantee restore.
            export_ok = False
            export_error = None
            restore_failed_names = []
            try:
                try:
                    first = next(c["name"] for c in self.workbook_data[path]
                                 if c["include"])
                    wb.Sheets(first).Activate()
                except Exception:
                    pass

                # PDF target. If a file with this name already exists
                # in the workbook's folder (e.g. last week's run of
                # this same IPC04), produce a revisioned name rather
                # than overwriting. Design rule: preserve
                # the existing PDF, write the new one alongside.
                desired_pdf = os.path.splitext(path)[0] + ".pdf"
                if self.companion and self.companion.get("output_dir"):
                    # Companion mode: PDF goes to the caller's folder.
                    desired_pdf = os.path.join(
                        self.companion["output_dir"],
                        os.path.basename(desired_pdf))
                out_pdf = find_revisioned_path(desired_pdf)
                if out_pdf is None:
                    result["error"] = (
                        "could not find an unused revision suffix for "
                        f"{desired_pdf} - too many existing revisions")
                    self._log_line(f"  ABORT: {result['error']}")
                    return result
                if out_pdf != desired_pdf:
                    self._log_line(
                        f"  Note: '{os.path.basename(desired_pdf)}' "
                        f"already exists, will write as "
                        f"'{os.path.basename(out_pdf)}' to preserve "
                        f"the existing file.")
                out_dir = os.path.dirname(out_pdf)

                # TEMP FILENAME DESIGN.
                #
                # We export to a temp file first and atomically promote
                # it to the target only on success. The temp file design
                # matters more than it looks:
                #
                # 1. SHORT, FIXED BASENAME (not derived from the
                # workbook name). Excel's PDF exporter has fragile
                # behaviour on long paths; the workbook basename
                # plus a long suffix plus a SharePoint-synced full
                # path can blow past Windows' practical path limits
                # and Excel will silently fail to write while still
                # returning success.
                # 2. The `~` prefix is a temp-file convention. It helps
                # reduce path-length and may reduce sync-agent
                # interference on OneDrive/SharePoint folders, but
                # we do NOT rely on OneDrive ignoring this file -
                # Microsoft's documentation does not bless a plain
                # `~` prefix as a safe "skip sync" pattern (only the
                # `~$` prefix used by Office lock files is
                # explicitly excluded). The real safety mechanism
                # is the os.replace step below: we never write
                # directly to the target name, so cross-process
                # visibility of the temp file is not a correctness
                # issue, only a performance one.
                # 3. PID + HHMMSS in the name avoids collision if two
                # runs happen in the same second on Windows builds
                # that reuse PIDs aggressively.
                # 4. Same folder as the workbook so `os.replace` to the
                # final name is atomic (cross-drive moves are
                # copy+delete).
                #
                # If this STILL produces "Excel reported success but no
                # PDF found" on a SharePoint-synced folder, the next
                # fallback is a two-step write:
                # 1. Export to a temp file in %TEMP% (a known-local
                # drive, no sync agent watching).
                # 2. shutil.copy2 from %TEMP% to a local temp name in
                # the workbook folder (cross-drive copy with
                # metadata preserved).
                # 3. os.replace within the workbook folder to promote
                # the local temp to the final target name.
                # That keeps the user-visible "final name appears"
                # transition atomic, while sidestepping any sync-agent
                # interference during the actual Excel write.
                # (shutil.move alone would NOT do this: across drives it
                # falls back to copy-then-delete and the rename step
                # would no longer be atomic.)
                tmp_pdf = os.path.join(
                    out_dir,
                    f"~ipc_export_{os.getpid()}_"
                    f"{datetime.now().strftime('%H%M%S')}.pdf")
                # Normalise to native Windows path. ExportAsFixedFormat
                # currently works with forward-slash paths from Tkinter
                # but the same COM path-handling quirk that broke
                # SaveAs could surface here on different Excel
                # versions. Belt-and-braces.
                tmp_pdf = os.path.normpath(tmp_pdf)
                out_pdf = os.path.normpath(out_pdf)
                self._log_line(f"  -> Exporting PDF: "
                               f"{os.path.basename(out_pdf)}")

                # Pre-emptively clear any stale tmp file from a previous
                # crashed run (HHMMSS makes collision unlikely, but
                # belt-and-braces).
                if os.path.exists(tmp_pdf):
                    try:
                        os.remove(tmp_pdf)
                    except Exception as _e:
                        self._log_line(
                            f"  Note: could not remove stale tmp file "
                            f"{tmp_pdf}: {_e}")

                # Phase A: Excel writes to tmp_pdf
                import time as _time
                excel_export_ok = False
                try:
                    wb.ExportAsFixedFormat(
                        Type=xlTypePDF,
                        Filename=tmp_pdf,
                        Quality=xlQualityStandard,
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False
                    )
                    # Poll for the file to appear. Antivirus scans,
                    # OneDrive/SharePoint sync, and disk buffer flushing
                    # can all delay file visibility for up to a few
                    # seconds after Excel returns. 5-second deadline is
                    # generous enough for cloud-synced folders without
                    # being so long that a genuine failure feels frozen.
                    # 50ms interval keeps the common-case overhead low
                    # (file typically appears in under a second).
                    _deadline = _time.time() + 5
                    while (not os.path.exists(tmp_pdf)
                           and _time.time() < _deadline):
                        _time.sleep(0.05)
                    if not os.path.exists(tmp_pdf):
                        # Surface the EXACT path Excel was asked to
                        # write, plus the most likely causes.
                        raise RuntimeError(
                            f"Excel reported success but no PDF was "
                            f"found after 5 seconds at:\n  {tmp_pdf}\n"
                            f"(path length {len(tmp_pdf)} chars)\n\n"
                            f"Likely causes:\n"
                            f"  - Antivirus quarantined the file as "
                            f"it was written\n"
                            f"  - OneDrive/SharePoint sync intercepted "
                            f"the write or rejected the temp filename\n"
                            f"  - Path too long for Excel's PDF "
                            f"exporter (try a shorter folder)\n"
                            f"  - The folder is read-only or the disk "
                            f"is full\n\n"
                            f"Check the workbook's folder for any PDF "
                            f"files Excel may have produced under a "
                            f"different name.")
                    excel_export_ok = True
                except Exception as e:
                    export_error = str(e)
                    self._log_line(f"  EXPORT FAILED: {e}")
                    # ExportAsFixedFormat itself failed - the tmp file
                    # is incomplete or junk. Delete it.
                    if os.path.exists(tmp_pdf):
                        try:
                            os.remove(tmp_pdf)
                        except Exception:
                            pass

                # Phase B: promote tmp to target. Excel may still have a
                # transient file handle just after export; retry briefly
                # before giving up. If it permanently fails, KEEP the
                # tmp file so the user has a recovery copy - earlier
                # code told the user to rename it manually and then
                # deleted it in cleanup.
                if excel_export_ok:
                    replaced = False
                    last_replace_err = None
                    for attempt in range(5):
                        try:
                            os.replace(tmp_pdf, out_pdf)
                            replaced = True
                            break
                        except Exception as e:
                            last_replace_err = e
                            _time.sleep(0.2)

                    if replaced:
                        export_ok = True
                        result["pdf_path"] = out_pdf
                        self._log_line(f"  Saved PDF: {out_pdf}")
                    else:
                        # Promotion permanently failed. KEEP the tmp PDF
                        # so the user has the freshly-produced output
                        # available for manual rescue. The stale file
                        # at out_pdf (if any) is left in place too -
                        # warn the user explicitly so they don't submit
                        # the wrong one.
                        export_error = (
                            f"PDF was produced at:\n  {tmp_pdf}\n"
                            f"but could not be moved to:\n  {out_pdf}\n"
                            f"after 5 attempts ({last_replace_err}). "
                            f"This usually means the target PDF is open "
                            f"in Acrobat or held by antivirus.\n\n"
                            f"DO NOT submit the file at the target "
                            f"path - it is the OLD PDF. Close the open "
                            f"PDF and either re-run, or manually rename "
                            f"the temp file to replace the target.")
                        self._log_line(f"  EXPORT FAILED (promotion): "
                                       f"{export_error}")
                        # Track the temp path so summary can surface it
                        result["recovery_pdf_path"] = tmp_pdf

            finally:
                # Restore original sheet visibility. Track any failures so
                # we can refuse to save the workbook (a partially-restored
                # state would otherwise be persisted).
                for name, vis in original_visibility.items():
                    try:
                        wb.Sheets(name).Visible = vis
                    except Exception as e:
                        restore_failed_names.append((name, str(e)))

            if not export_ok:
                result["error"] = export_error or "export failed"
                # Critical: do NOT save on export failure
                self._log_line("  Workbook NOT saved (export failed).")
                result["saved"] = False
                return result

            # If visibility restore failed, the PDF is fine but the
            # workbook is in an inconsistent state. The docstring
            # promises this is marked as failed: the result must
            # honour that contract even though a PDF was produced.
            # PDF path is kept in the result so the user knows a PDF
            # was produced and can decide whether to use it.
            if restore_failed_names:
                bad = ", ".join(n for n, _ in restore_failed_names)
                result["error"] = (
                    f"PDF was created but visibility of "
                    f"{len(restore_failed_names)} sheet(s) could not be "
                    f"restored: {bad}. Workbook NOT saved. Check the PDF "
                    f"manually before relying on it.")
                result["saved"] = "skipped (visibility restore failed)"
                self._log_line(
                    "  Workbook NOT saved (visibility restore failed).")
                self._log_line(
                    "  RESULT: FAILED. PDF exists but workbook state was "
                    "corrupted during export. Manual verification needed.")
                return result

            _phase_done("pdf_export")

            # PHASE 5: save. Only reached if export succeeded AND
            # visibility restored cleanly.
            #
            # Design rule: do NOT overwrite the source
            # workbook in-place. Save as a NEW file with a _rev.NN
            # suffix so the original on disk is preserved exactly as
            # it was opened. After this step the user has two files
            # side by side:
            #
            # IPC-04 - Progress Statement - Printable.xlsx (untouched)
            # IPC-04 - Progress Statement - Printable_rev.01.xlsx (new, with applied page-setup)
            #
            # Next run produces _rev.02, then _rev.03, etc. The user
            # gets a clean audit trail of every export and the
            # original source is never mutated by this tool.
            save_workbook = getattr(self, "save_workbook_after_export",
                                    True)
            save_failed_error = None
            if not save_workbook:
                result["saved"] = "skipped (per setting)"
                self._log_line("  Workbook NOT saved (per setting).")
            elif read_only:
                result["saved"] = "skipped (read-only)"
                self._log_line("  Workbook NOT saved (read-only).")
            else:
                # Find an unused _rev.NN path. The source workbook
                # always exists on disk (we opened from it), so this
                # will always return a revisioned name, never the
                # original path.
                wb_save_path = find_revisioned_path(path)
                if wb_save_path is None:
                    result["saved"] = (
                        "skipped (could not find unused revision "
                        "suffix)")
                    self._log_line(
                        "  WARNING: too many existing workbook "
                        "revisions in this folder (>9999); workbook "
                        "NOT saved. Page-setup changes will not "
                        "persist.")
                elif wb_save_path == path:
                    # Defensive: would only happen if `path` somehow
                    # didn't exist on disk by the time we got here.
                    # Shouldn't happen but handle it cleanly.
                    try:
                        wb.Save()
                        result["saved"] = True
                        self._log_line("  Workbook saved (in-place "
                                       "fallback - source file was "
                                       "not on disk).")
                    except Exception as e:
                        result["saved"] = f"error: {e}"
                        save_failed_error = str(e)
                        self._log_line(
                            f"  WARNING: could not save: {e}")
                else:
                    try:
                        # SaveAs requires the file format constant.
                        # xlOpenXMLWorkbook = 51 for .xlsx,
                        # xlOpenXMLWorkbookMacroEnabled = 52 for .xlsm,
                        # xlExcel12 = 50 for .xlsb,
                        # xlExcel8 = 56 for .xls (legacy 97-2003).
                        ext = os.path.splitext(wb_save_path)[1].lower()
                        if ext == ".xlsm":
                            fmt = 52
                        elif ext == ".xlsb":
                            fmt = 50
                        elif ext == ".xls":
                            fmt = 56
                        else:
                            fmt = 51  # xlsx default

                        # CRITICAL: normalise to native Windows
                        # backslash path before handing to Excel COM.
                        # Tkinter's askopenfilenames returns paths with
                        # forward slashes (e.g.
                        # "C:/Users/Alaa/Downloads/foo.xlsx"). Excel's
                        # SaveAs goes through COM marshalling and
                        # produces "Microsoft Excel cannot access the
                        # file 'C:\//Users/...'" errors with mangled
                        # path separators if it gets forward slashes.
                        # os.path.normpath fixes both: collapses any
                        # mixed separators and converts to backslashes
                        # on Windows.
                        wb_save_path_native = os.path.normpath(
                            wb_save_path)
                        wb.SaveAs(Filename=wb_save_path_native,
                                  FileFormat=fmt)
                        result["saved"] = True
                        result["workbook_saved_path"] = wb_save_path_native
                        self._log_line(
                            f"  Workbook saved as new file: "
                            f"{os.path.basename(wb_save_path_native)}")
                        self._log_line(
                            f"  Original preserved untouched: "
                            f"{os.path.basename(path)}")
                    except Exception as e:
                        result["saved"] = f"error: {e}"
                        save_failed_error = str(e)
                        # Diagnostic: log the resolved path we
                        # actually attempted, so future failures of
                        # this kind can be diagnosed from the log
                        # alone.
                        self._log_line(
                            f"  WARNING: could not save workbook to "
                            f"'{os.path.normpath(wb_save_path)}': {e}")

            # Save failure is NOT a PDF failure - the PDF is correct. But
            # the user should see it in the summary because next month's
            # auto-loaded page setup will be stale. Mark ok=True with a
            # warning attached.
            warning_parts = []
            if result.get("read_only_warning"):
                # Critical: this MUST appear in the summary, not just
                # in the log. The earlier code only logged the read-only
                # warning, so a stale-version PDF would be reported as
                # "Done" in the final summary - undermining the workflow
                # protection the warning is meant to provide.
                warning_parts.append(result["read_only_warning"])
            if save_failed_error:
                warning_parts.append(
                    f"workbook save failed ({save_failed_error}); "
                    f"page-setup changes will not be loaded next time")
            if result.get("setup_warnings"):
                warning_parts.append(
                    f"{len(result['setup_warnings'])} page-setup issue(s) "
                    f"during tab configuration; see WARN lines in log")
            if warning_parts:
                result["warning"] = ". ".join(warning_parts)

            result["ok"] = True

            # Finalise phase timing - includes the save phase
            _phase_done("save")

            # Emit timing summary. The total is the sum of phases,
            # which should match wall-clock time. Helps the user (and
            # any future review) understand where time is actually
            # spent - guesses are unreliable.
            total = sum(t for _, t in _phase_log)
            self._log_line(
                f"  Timing: "
                + ", ".join(f"{name}={t:.2f}s"
                            for name, t in _phase_log)
                + f" (total {total:.2f}s)")
            result["phase_timings"] = list(_phase_log)

            return result

        # ------------------------------------------------------------------
        # Apply one tab's PageSetup
        # ------------------------------------------------------------------

        def _apply_tab_setup(self, ws, cfg, footer_left):
            """Apply print setup to one worksheet from cfg.

            Returns a list of warning strings for non-fatal issues
            (e.g. paper-size fallback, page-break add failure). Fatal
            issues (PrintArea, Orientation, Zoom, Fit) raise to the
            caller so the workbook is aborted.

            An earlier approach swallowed all PageSetup failures, which
            defeated the fail-loud policy at the export layer. Caught

            Order of operations matters. The key bug we are working around:
            Excel raises 'Unable to set the FitToPagesTall property' if any
            of these are wrong:
              - Zoom is True (must be False before any Fit* setter)
              - FitToPagesTall is being set to 0 (must be 1..9999)

            We sanitise fit_height up front and always set Zoom = False
            first.

            We only clear *manual vertical* page breaks. Clearing all
            breaks via ResetAllPageBreaks() would also wipe horizontal
            breaks the user may have set manually in the workbook (e.g.
            between sections), and since the workbook is saved at the end
            of export, that destruction would become permanent. Automatic
            vertical breaks are left alone because they cannot be removed
            - Excel regenerates them on every repagination - and because
            FitToPagesWide, not the break collection, governs the output.

            PageSetup writes are batched under
            Application.PrintCommunication = False. Properties whose
            failure policy depends on knowing which property failed
            (PaperSize, PrintTitleRows) are set outside that batch.
            """
            warnings_out = []
            ps = ws.PageSetup

            # 1. Clear existing MANUAL vertical page breaks. Horizontal
            # breaks set by the user are preserved, and automatic
            # vertical breaks are deliberately left alone.
            #
            # PERFORMANCE ROOT CAUSE (fixed here).
            # VPageBreaks contains manual AND automatic breaks. Automatic
            # breaks cannot be removed - Excel regenerates them on every
            # repagination, and at this point print area and fit-to-width
            # have not been applied yet, so the sheet is still wider than
            # one page and automatic breaks always exist.
            #
            # The previous implementation deleted index 1 in a while loop
            # until Count reached 0. On any sheet carrying automatic
            # breaks that condition is never satisfiable, so the loop ran
            # its full 200-iteration guard on every tab. Every Count read
            # and every Delete forces a repagination through the printer
            # driver, so the guard - intended only as an infinite-loop
            # safety net - became the dominant cost of the entire export.
            # Measured on D-18 IPC 06: page_setup = 876.68s of an 896.92s
            # run across 73 tabs (~12s per tab, 97.7% of total runtime).
            #
            # New behaviour: one bounded read pass to record the COLUMN of
            # each manual break, then remove each by column reference.
            # Removing via Range.PageBreak is index-independent, so
            # repagination triggered by one deletion cannot invalidate the
            # positions still to be processed - which walking the
            # collection by index would.
            #
            # Automatic breaks are irrelevant to the output: pagination is
            # governed by FitToPagesWide, set at step 3 below.
            try:
                vpbs = ws.VPageBreaks
                # Read-only pass. Cap is a defensive bound only; it does
                # not loop, so a large collection costs one pass, not 200.
                # Truncation is reported - silently ignoring manual breaks
                # beyond the cap would contradict the fail-visible policy.
                break_count = vpbs.Count
                total_breaks = min(break_count, 500)
                if break_count > total_breaks:
                    warnings_out.append(
                        f"vertical page breaks on '{ws.Name}': inspected "
                        f"only {total_breaks} of {break_count}; manual "
                        f"breaks beyond the limit may remain")
                manual_cols = []
                unreadable = 0
                for i in range(1, total_breaks + 1):
                    try:
                        pb = vpbs(i)
                        if pb.Type == xlPageBreakManual:
                            manual_cols.append(pb.Location.Column)
                    except Exception:
                        unreadable += 1

                failed_clear = 0
                for col in manual_cols:
                    try:
                        ws.Columns(col).PageBreak = xlPageBreakNone
                    except Exception:
                        failed_clear += 1

                if unreadable or failed_clear:
                    warnings_out.append(
                        f"vertical page breaks on '{ws.Name}': "
                        f"{unreadable} could not be inspected, "
                        f"{failed_clear} could not be cleared. Pagination "
                        f"is still governed by fit-to-width, so the PDF "
                        f"layout is unaffected unless a stale manual "
                        f"break survived")
            except Exception as e:
                warnings_out.append(
                    f"could not clear vertical page breaks on "
                    f"'{ws.Name}': {e}")

            # 2. Paper size (per-tab). Deliberately kept OUTSIDE the
            # batched block below. Excel is not documented to guarantee
            # when a cached PageSetup value is validated - it may raise at
            # assignment or when communication is re-enabled, and in the
            # latter case we could no longer tell WHICH property was
            # rejected. Paper size has a real fallback (A3 -> A4 is a 50%
            # area difference) that depends on accurate attribution, so it
            # is worth one driver round trip per tab.
            paper = cfg.get("paper_size", xlA4)
            paper_label = PAPER_SIZE_LABELS.get(paper, str(paper))
            try:
                ps.PaperSize = paper
            except Exception as e:
                try:
                    ps.PaperSize = xlA4
                    warnings_out.append(
                        f"paper size {paper_label} rejected on "
                        f"'{ws.Name}', fell back to A4: {e}")
                except Exception as e2:
                    # Both rejected; let downstream see this as a hard
                    # failure by re-raising
                    raise RuntimeError(
                        f"paper size could not be set ({paper_label} "
                        f"and A4 both failed): {e2}")

            # 3. Batched PageSetup writes.
            #
            # PERFORMANCE. Every PageSetup assignment is a synchronous
            # round trip to the default printer driver, and each one
            # repaginates the sheet. Excel 2010+ exposes
            # Application.PrintCommunication to suspend that: writes are
            # queued and flushed once when it is set back to True. This
            # collapses roughly a dozen driver round trips per tab into
            # one.
            #
            # Microsoft documents only that False caches PageSetup
            # commands and True commits them. It does NOT guarantee that
            # reads fail while suspended, nor that an invalid value
            # necessarily raises at the commit rather than at assignment.
            # Both are observed behaviours, not contract. The design here
            # is therefore conservative rather than reliant on either:
            #   - Nothing in this block reads a PageSetup property back.
            #   - Properties whose failure policy needs the error
            #     attributed to that specific property (PaperSize,
            #     PrintTitleRows) are set outside the batch, so their
            #     fallbacks work regardless of when validation happens.
            #   - PrintArea stays inside. Its policy is already raise and
            #     abort the workbook, so the abort happens either way;
            #     only message specificity is at risk.
            #
            # PrintCommunication is absent before Excel 2010, so its
            # absence is tolerated and the block simply runs unbatched.
            app = ws.Application

            fit_w = safe_int(cfg.get("fit_width", 1), 1)
            fit_h = sanitise_fit_height(cfg.get("fit_height", 1))

            # Margins in points, computed in Python. Previously six
            # Application.InchesToPoints calls per tab, each a COM round
            # trip, to apply a conversion fixed at 72 points per inch.
            margin_lr = 0.3 * POINTS_PER_INCH
            margin_tb = 0.5 * POINTS_PER_INCH
            margin_hf = 0.3 * POINTS_PER_INCH

            # Excel footer syntax uses & as a control character; a
            # literal & must be doubled.
            safe_footer_left = (footer_left or "").replace("&", "&&")

            print_comm_suspended = False
            try:
                app.PrintCommunication = False
                print_comm_suspended = True
            except Exception:
                # Excel 2007 or earlier, or the property is unavailable.
                # Not fatal - the writes below simply run at the old cost.
                pass

            try:
                ps.Orientation = cfg["orientation"]

                # Zoom must be False BEFORE any Fit* setter, or Excel
                # raises 'Unable to set the FitToPagesTall property'.
                ps.Zoom = False

                # Fit-to-pages. Always sanitised: 0/None/legacy/garbage
                # -> safe values. FitToPagesTall = 0 being rejected by
                # Excel was the root cause of ~80% per-item-tab failures
                # earlier in development.
                ps.FitToPagesWide = fit_w
                ps.FitToPagesTall = fit_h

                # Print area. Hard requirement - if this fails the PDF
                # would print the wrong content, so we let it raise.
                ps.PrintArea = cfg["print_area"]

                ps.CenterHorizontally = True
                ps.CenterVertically = False

                ps.LeftMargin = margin_lr
                ps.RightMargin = margin_lr
                ps.TopMargin = margin_tb
                ps.BottomMargin = margin_tb
                ps.HeaderMargin = margin_hf
                ps.FooterMargin = margin_hf

                ps.LeftHeader = ""
                ps.CenterHeader = ""
                ps.RightHeader = ""
                ps.LeftFooter = f'&"Calibri,Italic"&9{safe_footer_left}'
                ps.CenterFooter = ""
                ps.RightFooter = '&"Calibri,Regular"&9Page &P of &N'
            finally:
                # Must always be restored. PrintCommunication is an
                # Application-wide setting, so leaving it False would
                # corrupt page setup for every later tab, every later
                # workbook, and the user's own Excel session.
                #
                # Two failure modes must be told apart, and BOTH are
                # fatal - they differ only in blast radius:
                #
                #   a) Communication cannot be restored at all. The Excel
                #      Application is unusable for page setup from here
                #      on, so the whole run must stop.
                #      -> PrintCommunicationError
                #
                #   b) Communication IS restored on the retry, but the
                #      first commit reported an error. That error came
                #      from Excel rejecting one of the cached PageSetup
                #      values (PrintArea, orientation, fit, margins,
                #      footers). Only this workbook is affected, but its
                #      page setup is not what was configured.
                #      -> RuntimeError, handled as an ordinary tab
                #         failure, so the workbook aborts with no PDF.
                #
                # An earlier revision discarded the first error whenever
                # the retry succeeded, on the mistaken reasoning that the
                # underlying property error would still propagate from
                # the try body. It cannot: if Excel cached the value
                # rather than validating it at assignment, the try body
                # never saw an error, and the only report of the
                # rejection is the one raised here. Swallowing it would
                # have produced a PDF whose pagination silently ignored
                # the configuration. There is no reliable way to show
                # that a first-attempt error was harmless, so it is
                # always treated as fatal to the workbook.
                if print_comm_suspended:
                    first_commit_error = None
                    restored = False

                    for _attempt in range(2):
                        try:
                            app.PrintCommunication = True
                            restored = True
                            break
                        except Exception as e:
                            if first_commit_error is None:
                                first_commit_error = e

                    if not restored:
                        raise PrintCommunicationError(
                            f"could not restore Excel PrintCommunication "
                            f"after configuring '{ws.Name}'. Every "
                            f"subsequent page-setup change in this Excel "
                            f"session would be cached and never applied, "
                            f"so the run has been stopped. Close Excel "
                            f"completely and re-run: {first_commit_error}"
                        ) from first_commit_error

                    if first_commit_error is not None:
                        raise RuntimeError(
                            f"Excel reported an error while committing "
                            f"page setup for '{ws.Name}'; one or more "
                            f"cached values were rejected, so the "
                            f"configured layout was not fully applied. "
                            f"The workbook has been aborted rather than "
                            f"produce a PDF that does not match the "
                            f"configuration: {first_commit_error}"
                        ) from first_commit_error

            # 4. Print titles (repeat rows). Also kept outside the batched
            # block: its failure policy is "warn and continue", which
            # needs the error attributed to this specific property. Set
            # after PrintArea, matching the previous order, because Excel
            # validates repeat rows against the print area.
            if cfg.get("title_rows"):
                rows = cfg["title_rows"]
                if not rows.startswith("$"):
                    parts = rows.split(":")
                    if len(parts) == 2:
                        rows = f"${parts[0]}:${parts[1]}"
                try:
                    ps.PrintTitleRows = rows
                except Exception as e:
                    warnings_out.append(
                        f"repeat rows '{cfg['title_rows']}' rejected on "
                        f"'{ws.Name}' - PDF will use whatever was "
                        f"previously set: {e}")
            else:
                try:
                    ps.PrintTitleRows = ""
                except Exception as e:
                    warnings_out.append(
                        f"could not clear repeat rows on "
                        f"'{ws.Name}': {e}")

            # 5. Vertical page breaks (after print area is set). Adding a
            # break is a worksheet operation, not a PageSetup property,
            # so it must run with PrintCommunication restored. If a
            # specific break fails to add, the PDF won't have it where
            # the UI promised - surface as a warning.
            failed_breaks = []
            for letter in cfg.get("splits", []):
                col = col_index(letter)
                if col > 0:
                    try:
                        break_cell = ws.Cells(1, col)
                        ws.VPageBreaks.Add(break_cell)
                    except Exception as e:
                        failed_breaks.append((letter, str(e)))
                else:
                    failed_breaks.append((letter, "invalid column letter"))
            if failed_breaks:
                detail = ", ".join(f"{l} ({e})" for l, e in failed_breaks)
                warnings_out.append(
                    f"page break(s) could not be added on '{ws.Name}': "
                    f"{detail}")

            return warnings_out

        # ------------------------------------------------------------------
        # Output folder
        # ------------------------------------------------------------------

        def _open_output_folder(self):
            if hasattr(self, "_output_dir"):
                try:
                    os.startfile(self._output_dir)
                except Exception as e:
                    messagebox.showerror("Cannot open", str(e))

        # ------------------------------------------------------------------
        # Helpers
        # ------------------------------------------------------------------

        def _clear(self):
            for w in self.root.winfo_children():
                w.destroy()
            # The brand label is a child widget of self.root and gets
            # destroyed above. The marker attribute survives, though,
            # which would make add_brand_label() skip on the next
            # rebuild and leave the wizard pages without the label.
            # Reset the marker here so each rebuild re-adds the label.
            #
            try:
                if hasattr(self.root, "_brand_label_added"):
                    delattr(self.root, "_brand_label_added")
            except Exception:
                pass

        def run(self):
            self.root.mainloop()


    # -----------------------------------------------------------------------
    # Main
    # -----------------------------------------------------------------------

    def main():
        app = App()
        app.run()


    if __name__ == "__main__":
        main()


except Exception as _e:
    if COMPANION:
        write_companion_result(COMPANION, "error",
                               error=f"startup failure: {_e}")
    _fatal_error("Startup failure",
                 f"{_e}\n\n{traceback.format_exc()}")
