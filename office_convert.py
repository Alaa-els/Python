"""
Document Bundle Builder - Office conversion stage
=================================================

Optional stage that converts Word and Excel files found in the bundle
folder tree to PDF, so they can be bundled in their folder position like
any other PDF.

- Word (.doc/.docx/.docm): headless export through an isolated Microsoft
  Word instance (``DispatchEx``), read-only, macros disabled, links not
  updated. The source document is never saved.
- Excel (.xls/.xlsx/.xlsm/.xlsb): delegated to Excel Quick Print running
  as a *companion* process. The full wizard is used (worksheet selection,
  print areas, orientation, scaling, page breaks, repeat rows, preview and
  the per-workbook ``.printcfg.json`` settings), and the resulting PDFs
  are handed back through a small JSON result file. Excel Quick Print
  stays usable standalone; the companion mode only pre-loads the file
  list, redirects the PDF output folder and suppresses workbook copies.

Converted PDFs are written to ``<base>/_bundle_converted/<mirror of the
source folder path>/<stem>.pdf`` and are re-used on later runs while they
are newer than their source. The folder is excluded from the bundle walk;
the engine maps each source document to its converted PDF through the
"overlay" (source path -> PDF path) returned by ``run_conversion_stage``.

Only ``win32com`` / ``pythoncom`` (pywin32) and an installed Office are
needed, and only on Windows, and only when conversion is switched on.
Everything else in the bundle builder works without them.
"""

import os
import sys
import json
import time
import shutil
import subprocess
import importlib
from datetime import datetime

from bundle_engine import STAGING_DIRNAME, is_excluded_folder, norm_path

WORD_EXTS = (".doc", ".docx", ".docm", ".rtf")
EXCEL_EXTS = (".xls", ".xlsx", ".xlsm", ".xlsb")
OFFICE_EXTS = WORD_EXTS + EXCEL_EXTS

COMPANION_PROTOCOL = 1
COMPANION_RESULT_NAME = "excel_quick_print_result.json"
EXCEL_TOOL_CANDIDATES = ("Excel_Quick_Print.pyw", "excel_quick_print.pyw",
                         "Excel_Quick_Print.py", "excel_quick_print.py")

# Word constants
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# A password that no real document will have. Supplying it makes Word
# raise an error for protected documents instead of showing a modal
# prompt inside an invisible application (which would hang the build).
DUMMY_PASSWORD = "\x01bundle-builder-no-prompt\x02"


class ConversionCancelled(Exception):
    pass


class OfficeDocument:
    __slots__ = ("path", "kind", "target", "status", "error", "warning")

    def __init__(self, path, kind, target):
        self.path = path
        self.kind = kind          # "word" | "excel"
        self.target = target      # planned PDF path in the staging folder
        self.status = "pending"   # pending | converted | reused | failed | cancelled | skipped
        self.error = None
        self.warning = None

    @property
    def name(self):
        return os.path.basename(self.path)

    def as_dict(self):
        return {"path": self.path, "kind": self.kind, "target": self.target,
                "status": self.status, "error": self.error, "warning": self.warning}


class ConversionOutcome:
    def __init__(self):
        self.documents = []
        self.cancelled = False
        self.messages = []

    @property
    def overlay(self):
        """source path (normalised) -> converted PDF path, successes only."""
        return {norm_path(d.path): d.target for d in self.documents
                if d.status in ("converted", "reused") and os.path.isfile(d.target)}

    @property
    def failures(self):
        return [d for d in self.documents if d.status in ("failed", "cancelled")]

    def summary(self):
        counts = {}
        for d in self.documents:
            counts[d.status] = counts.get(d.status, 0) + 1
        return ", ".join(f"{v} {k}" for k, v in sorted(counts.items())) or "nothing to convert"


# ---------------------------------------------------------------------------
# Discovery and staging
# ---------------------------------------------------------------------------

def staging_root(base_dir):
    return os.path.join(base_dir, STAGING_DIRNAME)


def staging_pdf_path(base_dir, source_path):
    """Mirror the source folder path under the staging folder."""
    rel = os.path.relpath(os.path.dirname(source_path), base_dir)
    rel = "" if rel == "." else rel
    stem = os.path.splitext(os.path.basename(source_path))[0]
    return os.path.join(staging_root(base_dir), rel, stem + ".pdf")


def classify(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in WORD_EXTS:
        return "word"
    if ext in EXCEL_EXTS:
        return "excel"
    return None


def discover_office_documents(base_dir, kinds=("word", "excel")):
    """Office files in the tree, in walk order. Lock files (~$), the
    staging folder and other tool-owned folders are ignored."""
    found = []
    for cur, dirs, files in os.walk(base_dir):
        dirs[:] = sorted(d for d in dirs if not is_excluded_folder(d))
        for f in sorted(files):
            if f.startswith("~$"):
                continue
            kind = classify(f)
            if kind and kind in kinds:
                full = os.path.join(cur, f)
                found.append(OfficeDocument(full, kind, staging_pdf_path(base_dir, full)))
    return found


def planned_overlay(base_dir, kinds=("word", "excel")):
    """Overlay of every Office document to its *planned* PDF, used before
    conversion so folder-policy dialogs see the documents as content."""
    return {norm_path(d.path): d.target for d in discover_office_documents(base_dir, kinds)}


def is_up_to_date(source, target):
    try:
        return (os.path.isfile(target) and os.path.getsize(target) > 0
                and os.path.getmtime(target) >= os.path.getmtime(source))
    except OSError:
        return False


def looks_encrypted(path):
    """True when an Office Open XML file is actually an OLE container,
    which is how password-protected .docx/.xlsx files are stored.
    None when it cannot be told from the header (legacy .doc/.xls are
    always OLE containers)."""
    ext = os.path.splitext(path)[1].lower()
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return None
    if ext in (".docx", ".docm", ".xlsx", ".xlsm"):
        if head.startswith(b"PK"):
            return False
        if head.startswith(b"\xD0\xCF\x11\xE0"):
            return True
        return None
    return None


def pywin32_available():
    for name in ("win32com.client", "pythoncom"):
        try:
            importlib.import_module(name)
        except Exception:
            return False
    return True


def office_support_status():
    """Human-readable availability summary for the GUI."""
    if sys.platform != "win32":
        return False, "Office conversion needs Windows with Microsoft Office installed."
    if not pywin32_available():
        return False, ("pywin32 is not installed. Run:  python -m pip install --user pywin32  "
                       "then restart the tool.")
    return True, "pywin32 found. Word/Excel will be started on demand."


def find_excel_quick_print(explicit_path=""):
    if explicit_path and os.path.isfile(explicit_path):
        return explicit_path
    here = os.path.dirname(os.path.abspath(__file__))
    for name in EXCEL_TOOL_CANDIDATES:
        cand = os.path.join(here, name)
        if os.path.isfile(cand):
            return cand
    return None


# ---------------------------------------------------------------------------
# Word backend
# ---------------------------------------------------------------------------

class WordConverter:
    """Headless Word export through an isolated instance.

    Use as a context manager; the instance is created lazily on the first
    conversion and only that instance is quit on exit."""

    def __init__(self, log_fn=None):
        self.log = log_fn or (lambda s: None)
        self.app = None
        self._com_initialised = False

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def _ensure_app(self):
        if self.app is not None:
            return
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        self._com_initialised = True
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = WD_ALERTS_NONE
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
            try:
                app.Options.UpdateLinksAtOpen = False
            except Exception:
                pass
            self.app = app
        except Exception:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialised = False
            raise

    def convert(self, source, target):
        """Return (ok, error_message)."""
        enc = looks_encrypted(source)
        if enc:
            return False, "password-protected document (encrypted container)"
        try:
            self._ensure_app()
        except Exception as e:
            return False, f"Microsoft Word could not be started: {e}"
        os.makedirs(os.path.dirname(target), exist_ok=True)
        tmp = target + f".tmp{os.getpid()}.pdf"
        doc = None
        try:
            doc = self.app.Documents.Open(
                FileName=os.path.normpath(os.path.abspath(source)),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=DUMMY_PASSWORD,
                Visible=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.normpath(tmp),
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                Range=WD_EXPORT_ALL_DOCUMENT,
                Item=WD_EXPORT_DOCUMENT_CONTENT,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
            deadline = time.time() + 5
            while not os.path.exists(tmp) and time.time() < deadline:
                time.sleep(0.05)
            if not os.path.exists(tmp):
                return False, "Word reported success but no PDF appeared"
            os.replace(tmp, target)
            return True, None
        except Exception as e:
            msg = str(e)
            low = msg.lower()
            if "password" in low:
                msg = "password-protected document"
            elif "locked" in low or "in use" in low or "being used" in low:
                msg = f"document is locked by another application: {msg}"
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            return False, msg
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass

    def close(self):
        if self.app is not None:
            try:
                self.app.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                try:
                    self.app.Quit()
                except Exception:
                    pass
            self.app = None
        if self._com_initialised:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialised = False


# ---------------------------------------------------------------------------
# Excel companion backend
# ---------------------------------------------------------------------------

def companion_command(tool_path, files, output_dir, result_path, python_exe=None,
                      title=None):
    """Command line for Excel Quick Print in companion mode."""
    exe = python_exe or sys.executable
    cmd = [exe, tool_path, "--companion-result", result_path, "--output-dir", output_dir,
           "--no-workbook-copy"]
    if title:
        cmd += ["--title", title]
    cmd.append("--")
    cmd += list(files)
    return cmd


def read_companion_result(result_path):
    """Parse the result JSON. Returns a dict with keys status, error,
    results (list of {path, ok, pdf_path, error, warning})."""
    with open(result_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict) or "status" not in data:
        raise ValueError("companion result file has an unexpected shape")
    data.setdefault("results", [])
    data.setdefault("error", None)
    return data


class ExcelCompanion:
    """Runs Excel Quick Print as a companion process for a set of
    workbooks and collects the produced PDFs."""

    def __init__(self, tool_path, log_fn=None, python_exe=None, poll_interval=0.25):
        self.tool_path = tool_path
        self.log = log_fn or (lambda s: None)
        self.python_exe = python_exe
        self.poll_interval = poll_interval

    def convert_many(self, docs, output_dir, cancel_event=None, title=None):
        """docs: list of OfficeDocument (kind excel). Sets status on each.
        Returns the raw result dict (or None when cancelled/failed)."""
        os.makedirs(output_dir, exist_ok=True)
        result_path = os.path.join(output_dir, COMPANION_RESULT_NAME)
        try:
            if os.path.exists(result_path):
                os.remove(result_path)
        except OSError:
            pass
        files = [d.path for d in docs]
        cmd = companion_command(self.tool_path, files, output_dir, result_path,
                                self.python_exe, title)
        self.log(f"  Launching Excel Quick Print for {len(files)} workbook(s)...")
        self.log("  Complete the wizard (sheets, print areas, preview) and export; "
                 "the PDFs come back into the bundle automatically.")
        try:
            proc = subprocess.Popen(cmd, cwd=os.path.dirname(self.tool_path) or None)
        except Exception as e:
            for d in docs:
                d.status, d.error = "failed", f"Excel Quick Print could not be started: {e}"
            return None
        while proc.poll() is None:
            if cancel_event is not None and cancel_event.is_set():
                self.log("  Cancel requested - closing Excel Quick Print.")
                try:
                    proc.terminate()
                except Exception:
                    pass
                for d in docs:
                    d.status, d.error = "cancelled", "cancelled by user"
                return None
            time.sleep(self.poll_interval)
        if not os.path.isfile(result_path):
            err = (f"Excel Quick Print closed without a result (exit code {proc.returncode}). "
                   f"The export was not run or the tool failed before writing "
                   f"{COMPANION_RESULT_NAME}.")
            for d in docs:
                d.status, d.error = "failed", err
            return None
        try:
            result = read_companion_result(result_path)
        except Exception as e:
            for d in docs:
                d.status, d.error = "failed", f"unreadable companion result: {e}"
            return None
        by_path = {norm_path(r.get("path", "")): r for r in result.get("results", [])}
        status = result.get("status")
        for d in docs:
            r = by_path.get(norm_path(d.path))
            if status == "cancelled":
                d.status = "cancelled"
                d.error = "Excel Quick Print was closed before exporting"
                continue
            if status == "error" and r is None:
                d.status, d.error = "failed", result.get("error") or "Excel Quick Print failed"
                continue
            if r is None:
                d.status = "failed"
                d.error = "workbook was removed from the Excel Quick Print file list"
                continue
            if r.get("ok") and r.get("pdf_path") and os.path.isfile(r["pdf_path"]):
                pdf = r["pdf_path"]
                if norm_path(pdf) != norm_path(d.target):
                    os.makedirs(os.path.dirname(d.target), exist_ok=True)
                    try:
                        if os.path.exists(d.target):
                            os.remove(d.target)
                        shutil.move(pdf, d.target)
                    except Exception as e:
                        d.status, d.error = "failed", f"could not place PDF in staging: {e}"
                        continue
                d.status = "converted"
                d.warning = r.get("warning")
            else:
                d.status = "failed"
                d.error = r.get("error") or "no PDF produced"
        return result


# ---------------------------------------------------------------------------
# Stage runner
# ---------------------------------------------------------------------------

def run_conversion_stage(base_dir, opts, log_fn=None, cancel_event=None,
                         word_backend=None, excel_backend=None):
    """Convert every Office document under base_dir.

    opts keys: office_reconvert (bool), excel_quick_print_path (str),
    office_kinds (tuple, default word+excel), claim_title.
    word_backend / excel_backend: objects providing WordConverter /
    ExcelCompanion interfaces (injectable for tests). When None, the real
    COM-backed implementations are used.

    Returns ConversionOutcome. Never raises for per-document problems;
    those are recorded on the documents."""
    log = log_fn or (lambda s: None)
    opts = opts or {}
    outcome = ConversionOutcome()
    kinds = tuple(opts.get("office_kinds") or ("word", "excel"))
    docs = discover_office_documents(base_dir, kinds)
    outcome.documents = docs
    if not docs:
        log("Office conversion: no Word/Excel documents found.")
        return outcome

    force = bool(opts.get("office_reconvert"))
    log(f"Office conversion: {len(docs)} document(s) found "
        f"({sum(1 for d in docs if d.kind == 'word')} Word, "
        f"{sum(1 for d in docs if d.kind == 'excel')} Excel).")
    os.makedirs(staging_root(base_dir), exist_ok=True)

    # Reuse up-to-date conversions
    todo = []
    for d in docs:
        if not force and is_up_to_date(d.path, d.target):
            d.status = "reused"
            log(f"  Reusing: {d.name} (PDF is newer than the source)")
        else:
            todo.append(d)

    def cancelled():
        return cancel_event is not None and cancel_event.is_set()

    # Availability
    if todo and word_backend is None and excel_backend is None:
        ok, reason = office_support_status()
        if not ok:
            for d in todo:
                d.status, d.error = "failed", reason
            log(f"  Office conversion unavailable: {reason}")
            outcome.messages.append(reason)
            return outcome

    # Word
    word_docs = [d for d in todo if d.kind == "word"]
    if word_docs:
        backend = word_backend if word_backend is not None else WordConverter(log)
        try:
            with backend as conv:
                for d in word_docs:
                    if cancelled():
                        d.status, d.error = "cancelled", "cancelled by user"
                        continue
                    log(f"  Word: {d.name}")
                    ok, err = conv.convert(d.path, d.target)
                    if ok:
                        d.status = "converted"
                    else:
                        d.status, d.error = "failed", err
                        log(f"    FAILED: {err}")
        except Exception as e:
            for d in word_docs:
                if d.status == "pending":
                    d.status, d.error = "failed", f"Word backend error: {e}"
            log(f"  Word conversion aborted: {e}")

    if cancelled():
        for d in todo:
            if d.status == "pending":
                d.status, d.error = "cancelled", "cancelled by user"
        outcome.cancelled = True
        return outcome

    # Excel
    excel_docs = [d for d in todo if d.kind == "excel"]
    if excel_docs:
        backend = excel_backend
        if backend is None:
            tool = find_excel_quick_print(opts.get("excel_quick_print_path", ""))
            if not tool:
                msg = ("Excel_Quick_Print.pyw was not found next to the bundle builder; "
                       "set its location in the Office documents panel.")
                for d in excel_docs:
                    d.status, d.error = "failed", msg
                log(f"  {msg}")
                outcome.messages.append(msg)
                backend = None
            else:
                backend = ExcelCompanion(tool, log)
        if backend is not None:
            for d in excel_docs:
                # Stale staging PDFs are ours; clear them so the companion
                # writes a clean <stem>.pdf rather than a _rev copy.
                try:
                    if os.path.exists(d.target):
                        os.remove(d.target)
                except OSError:
                    pass
            backend.convert_many(excel_docs, staging_root(base_dir), cancel_event,
                                 title=opts.get("claim_title"))
            for d in excel_docs:
                if d.status == "converted":
                    log(f"  Excel: {d.name} -> {os.path.basename(d.target)}"
                        + (f"  [warning: {d.warning}]" if d.warning else ""))
                elif d.status != "pending":
                    log(f"  Excel: {d.name} FAILED: {d.error}")

    outcome.cancelled = any(d.status == "cancelled" for d in docs) and cancelled()
    log(f"Office conversion finished: {outcome.summary()}.")
    return outcome


def write_conversion_report(base_dir, outcome):
    """Persist a small JSON report next to the staging folder for audit."""
    path = os.path.join(staging_root(base_dir), "conversion_report.json")
    try:
        os.makedirs(staging_root(base_dir), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"generated": datetime.now().isoformat(),
                       "documents": [d.as_dict() for d in outcome.documents]}, f, indent=2)
        return path
    except Exception:
        return None
