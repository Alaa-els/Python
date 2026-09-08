"""
Document Bundle Builder
=======================

Merges a folder tree of PDFs (and, optionally, Word/Excel files converted
to PDF) into one organised bundle with:
- Numbered sections (Appendix / Section / Drawing / Document / custom term)
- Truly optional separator pages: none, top-level only, all sections, or
  every document - chosen once, saved between runs
- Hierarchical, clickable Table of Contents and PDF outline bookmarks,
  each controlled independently of separators
- Per-page footers that can be switched off or moved to avoid title blocks
- Bundle styles: Claims (navy + gold), Minimal monochrome, Technical
  drawing register, Corporate report, Correspondence pack - plus the
  individual cover / separator / contents / palette / font pickers
- Optional OCR pass, image compression, cover page, custom order via
  _order.txt, persistent settings

Designed for QS / claims work (variation packs, EOT submissions, IPC
bundles) and for general document, drawing and correspondence bundles.

Files: this launcher, bundle_engine.py (build engine), office_convert.py
(Word/Excel stage) and Excel_Quick_Print.pyw (Excel wizard) must sit in the
same folder. Double-click to run without a console window. First run
auto-installs the required Python packages.
"""

import os
import sys
import json
import subprocess
import importlib
import threading
import traceback
from datetime import datetime


# ---------------------------------------------------------------------------
# Step 1: Last-resort error display
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
    # Step 2: Bootstrap dependencies
    # -----------------------------------------------------------------------

    REQUIRED_PACKAGES = {"fitz": "PyMuPDF", "fpdf": "fpdf2"}
    OPTIONAL_PACKAGES = {"PIL": "Pillow", "pytesseract": "pytesseract"}

    def _missing(import_name):
        try:
            importlib.import_module(import_name)
            return False
        except ImportError:
            return True

    def _bootstrap_packages():
        """Install required packages on first run with a live-progress window."""
        missing_required = [(imp, pip) for imp, pip in REQUIRED_PACKAGES.items()
                            if _missing(imp)]
        if not missing_required:
            return

        import tkinter as tk
        from tkinter import scrolledtext

        win = tk.Tk()
        win.title("Document Bundle Builder - First Run Setup")
        win.geometry("700x420")
        win.attributes("-topmost", True)
        win.after(2000, lambda: win.attributes("-topmost", False))
        pkg_list = ", ".join(p for _, p in missing_required)
        tk.Label(win, text="First-run setup", font=("Arial", 14, "bold")).pack(pady=(12, 4))
        tk.Label(win, text=(f"Installing required Python package(s): {pkg_list}\n"
                            "This is a one-time install into your user folder. No admin rights needed.\n"
                            "Live install output appears below. The window will close automatically when done."),
                 justify="center").pack(padx=12, pady=(0, 8))
        log_widget = scrolledtext.ScrolledText(win, height=15, font=("Consolas", 9), state="disabled")
        log_widget.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        status_var = tk.StringVar(value="Starting...")
        tk.Label(win, textvariable=status_var, font=("Arial", 9, "bold"), fg="#1F3864").pack(pady=(0, 8))
        state = {"failed": False, "error_pkg": "", "error_msg": "", "done": False}

        def append_log(text):
            log_widget.configure(state="normal")
            log_widget.insert("end", text)
            log_widget.see("end")
            log_widget.configure(state="disabled")

        def append_log_safe(text):
            win.after(0, lambda t=text: append_log(t))

        def set_status_safe(text):
            win.after(0, lambda t=text: status_var.set(t))

        def install_worker():
            import time as _time
            for imp_name, pip_name in missing_required:
                set_status_safe(f"Installing {pip_name}... (this can take a few minutes)")
                append_log_safe(f"\n>>> Installing {pip_name}\n")
                try:
                    cmd = [sys.executable, "-m", "pip", "install", "--user",
                           "--disable-pip-version-check", "--no-input", "--progress-bar", "off",
                           pip_name]
                    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                            text=True, bufsize=1,
                                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                    last_output = _time.time()
                    IDLE_TIMEOUT = 180
                    while True:
                        line = proc.stdout.readline()
                        if line == "" and proc.poll() is not None:
                            break
                        if line:
                            append_log_safe(line)
                            last_output = _time.time()
                        elif _time.time() - last_output > IDLE_TIMEOUT:
                            proc.kill()
                            state.update(failed=True, error_pkg=pip_name,
                                         error_msg=f"pip produced no output for {IDLE_TIMEOUT} seconds.")
                            append_log_safe(f"\n*** STALLED: killing pip after {IDLE_TIMEOUT}s of silence.\n")
                            break
                        else:
                            _time.sleep(0.1)
                    proc.stdout.close()
                    rc = proc.wait()
                    if state["failed"]:
                        break
                    if rc != 0:
                        state.update(failed=True, error_pkg=pip_name, error_msg=f"pip exited with code {rc}")
                        append_log_safe(f"\n*** FAILED: {pip_name} (exit code {rc})\n")
                        break
                    append_log_safe(f"\n--- OK: {pip_name} installed\n")
                except Exception as e:
                    state.update(failed=True, error_pkg=pip_name, error_msg=str(e))
                    append_log_safe(f"\n*** EXCEPTION: {e}\n")
                    break
            if not state["failed"]:
                set_status_safe("Verifying...")
                importlib.invalidate_caches()
                try:
                    import site
                    user_site = site.getusersitepackages()
                    if user_site and os.path.isdir(user_site):
                        site.addsitedir(user_site)
                except Exception:
                    pass
                still_missing = [imp for imp, _ in missing_required if _missing(imp)]
                if still_missing:
                    state.update(failed=True, error_pkg=", ".join(still_missing),
                                 error_msg="Installed but cannot be imported.")
                    append_log_safe(f"\n*** Import verification failed: {still_missing}\n")
                else:
                    append_log_safe("\nAll packages installed and verified.\n")
                    set_status_safe("Done. Closing in 2 seconds...")
            state["done"] = True
            if not state["failed"]:
                win.after(2000, win.destroy)
            else:
                set_status_safe("FAILED - see log above. Close this window to exit.")

        threading.Thread(target=install_worker, daemon=True).start()
        win.mainloop()

        if state["failed"]:
            _fatal_error(
                f"{state['error_pkg']} install failed",
                f"Could not install {state['error_pkg']}.\n\nReason: {state['error_msg']}\n\n"
                f"Try manually in CMD:\n    python -m pip install --user "
                f"{' '.join(p for _, p in missing_required)}")
        importlib.invalidate_caches()
        still_missing = [imp for imp, _ in missing_required if _missing(imp)]
        if still_missing:
            _fatal_error("Install verification failed",
                         f"Packages installed but not importable: {still_missing}\n"
                         "Restart the script. If it persists, install manually.")

    _bootstrap_packages()

    # -----------------------------------------------------------------------
    # Step 3: Imports
    # -----------------------------------------------------------------------

    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog

    _HERE = os.path.dirname(os.path.abspath(__file__))
    if _HERE not in sys.path:
        sys.path.insert(0, _HERE)

    import bundle_engine as be
    import office_convert as oc
    import bundle_organiser as bo

    OCR_AVAILABLE = be.ocr_available()

    # -----------------------------------------------------------------------
    # Step 4: Folder policy dialogs
    # -----------------------------------------------------------------------

    def _term():
        return be.CURRENT_TERMS.get("section_term", "Appendix")

    def show_discovery_popup(parent, counts, term):
        """Stage 1: summary popup. Returns 'automatic' | 'manual' | None."""
        dlg = tk.Toplevel(parent)
        dlg.title("Folder structure detected")
        dlg.geometry("560x340")
        dlg.transient(parent)
        dlg.grab_set()
        choice = {"value": None}
        tk.Label(dlg, text="Folder structure detected", font=("Arial", 12, "bold")).pack(pady=(12, 6))
        summary = (
            f"Your bundle contains:\n\n"
            f"   - {counts['top_pdfs']} top-level document(s) (each becomes a numbered {term})\n"
            f"   - {counts['top_folders']} top-level folder(s) (become subsequent {term}s)\n"
            f"   - {counts['total_subfolders']} subfolder(s) nested inside\n"
            f"   - {counts['total_pdfs']} document(s) in total\n\n"
            f"How should folders be handled?"
        )
        tk.Label(dlg, text=summary, justify="left", font=("Arial", 10)).pack(padx=20, pady=8, anchor="w")
        tk.Label(dlg, text="Automatic = one rule applied to all folders (faster)\n"
                           "Manual    = decide for each folder individually",
                 fg="#555", font=("Arial", 9), justify="left").pack(padx=20, anchor="w")
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=18)

        def pick(v):
            choice["value"] = v
            dlg.destroy()

        tk.Button(btn_frame, text="Automatic mode", width=18, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=lambda: pick("automatic")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Manual mode", width=18, font=("Arial", 10, "bold"),
                  command=lambda: pick("manual")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel", width=10, command=lambda: pick(None)).pack(side="left", padx=6)
        dlg.wait_window()
        return choice["value"]

    def show_automatic_options_dialog(parent, default_layout="preserve_skip_numeric"):
        """Stage 2A: layout for automatic mode. Returns dict or 'BACK' / None.
        Separator pages are a global choice on the main window now."""
        dlg = tk.Toplevel(parent)
        dlg.title("Automatic mode - choose layout")
        dlg.geometry("640x430")
        dlg.transient(parent)
        dlg.grab_set()
        result = {"value": None}
        tk.Label(dlg, text="Automatic mode - choose layout", font=("Arial", 12, "bold")).pack(pady=(10, 6))
        layout_frame = tk.LabelFrame(dlg, text="How should folder hierarchy appear?", padx=10, pady=8)
        layout_frame.pack(fill="x", padx=12, pady=6)
        var_layout = tk.StringVar(value=default_layout)
        layouts = [
            ("flatten", "FLATTEN",
             "All PDFs from all subfolders combined into the top-level section.\n"
             "TOC shows files only, no folder headings.\n"
             "Best for: equipment docs where folder names are admin scaffolding."),
            ("preserve", "PRESERVE STRUCTURE",
             "Folders and subfolders appear as headings in the TOC.\n"
             "Files listed under their parent folder.\n"
             "Best for: claim packs where folder names are meaningful categories."),
            ("preserve_skip_numeric", "PRESERVE STRUCTURE - SKIP NUMERIC",
             "Same as above, but folders named only with digits are flattened\n"
             "into their parent. Other named folders kept as headings."),
        ]
        for value, label, desc in layouts:
            f = tk.Frame(layout_frame)
            f.pack(fill="x", anchor="w", pady=2)
            tk.Radiobutton(f, text=label, variable=var_layout, value=value,
                           font=("Arial", 9, "bold")).pack(anchor="w")
            tk.Label(f, text=desc, fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        tk.Label(dlg, text="Separator pages, contents and bookmarks are set on the main window.",
                 fg="#555", font=("Arial", 9)).pack(padx=12, pady=(4, 0), anchor="w")
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=12)

        def ok():
            result["value"] = {"mode": "automatic", "layout": var_layout.get()}
            dlg.destroy()

        def back():
            result["value"] = "BACK"
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Back", width=10, command=back).pack(side="left", padx=6)
        dlg.wait_window()
        return result["value"]

    def _depth_combo(parent, max_subtree_depth, labels_fn):
        depth_options = [1]
        depth_labels = {1: labels_fn(1)}
        for n in range(2, max(max_subtree_depth, 1) + 2):
            depth_options.append(n)
            depth_labels[n] = labels_fn(n)
        depth_options.append(999)
        depth_labels[999] = labels_fn(999)
        var_depth = tk.IntVar(value=999)
        menu = ttk.Combobox(parent, values=[depth_labels[d] for d in depth_options],
                            state="readonly", width=32, font=("Arial", 9))
        menu.set(depth_labels[999])

        def on_change(_event):
            label = menu.get()
            for d, lbl in depth_labels.items():
                if lbl == label:
                    var_depth.set(d)
                    return

        menu.bind("<<ComboboxSelected>>", on_change)
        return menu, var_depth

    def show_manual_folder_dialog(parent, folder_rel, idx, total, pdf_count, sub_count,
                                  max_subtree_depth):
        """Stage 2B per-folder: returns (decision_dict, apply_to_all_remaining)."""
        dlg = tk.Toplevel(parent)
        dlg.title(f"Manual mode - folder {idx} of {total}")
        dlg.geometry("680x600")
        dlg.transient(parent)
        dlg.grab_set()
        result = {"decision": None, "all": False}
        tk.Label(dlg, text=f"Manual mode - folder {idx} of {total}",
                 font=("Arial", 12, "bold")).pack(pady=(10, 4))
        path_frame = tk.LabelFrame(dlg, text="Folder path", padx=8, pady=6)
        path_frame.pack(fill="x", padx=12, pady=4)
        tk.Label(path_frame, text=folder_rel, font=("Consolas", 10), wraplength=620,
                 justify="left").pack(anchor="w")
        depth_str = f"{pdf_count} document(s) directly in this folder, {sub_count} immediate subfolder(s)"
        if max_subtree_depth > 0:
            depth_str += f", subtree up to {max_subtree_depth} level(s) deep"
        tk.Label(dlg, text="Contains: " + depth_str, font=("Arial", 9)).pack(pady=(4, 6))
        opt_frame = tk.LabelFrame(dlg, text="How should this folder be handled?", padx=10, pady=8)
        opt_frame.pack(fill="x", padx=12, pady=4)
        var_action = tk.StringVar(value="show")
        show_row = tk.Frame(opt_frame)
        show_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(show_row, text="SHOW FILES", variable=var_action, value="show",
                       font=("Arial", 9, "bold")).pack(side="left", anchor="w")
        tk.Label(show_row, text="depth limit:", font=("Arial", 9)).pack(side="left", padx=(20, 4))

        def labels(n):
            if n == 1:
                return "1 (collapse - just this folder)"
            if n == 2:
                return "2 (one level below)"
            if n == 999:
                return "unlimited (everything)"
            return f"{n} ({n - 1} levels below)"

        depth_menu, var_depth = _depth_combo(show_row, max_subtree_depth, labels)
        depth_menu.pack(side="left")
        tk.Label(opt_frame, text="    1 = collapse (folder name only, files hidden in TOC).\n"
                                 "    2-N = show files/subfolders up to N levels deep.\n"
                                 "    unlimited = show every file at every depth.",
                 fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        tk.Radiobutton(opt_frame, text="FLATTEN INTO PARENT", variable=var_action, value="flatten",
                       font=("Arial", 9, "bold")).pack(anchor="w", pady=(4, 0))
        tk.Label(opt_frame, text="    No heading for this folder. Documents absorbed into parent's level.\n"
                                 "    Subfolders inside still ask their own questions.",
                 fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        tk.Radiobutton(opt_frame, text="SKIP", variable=var_action, value="skip",
                       font=("Arial", 9, "bold")).pack(anchor="w", pady=(4, 0))
        tk.Label(opt_frame, text="    Exclude this folder and its contents entirely.",
                 fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        var_all = tk.BooleanVar(value=False)
        tk.Checkbutton(dlg, text="Apply this choice to all remaining folders in this section",
                       variable=var_all, font=("Arial", 9)).pack(pady=(8, 4))
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=8)

        def ok():
            action = var_action.get()
            result["decision"] = ({"action": "show", "depth": var_depth.get()}
                                  if action == "show" else {"action": action})
            result["all"] = var_all.get()
            dlg.destroy()

        def cancel():
            result["decision"] = {"action": "cancel"}
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel build", width=14, command=cancel).pack(side="left", padx=6)
        dlg.wait_window()
        return result["decision"], result["all"]

    def show_appendix_level_dialog(parent, appendix_name, idx, total, pdf_count, sub_count,
                                   max_subtree_depth, term):
        """Per-section dialog: {"action": "show", "depth": N} | {"action": "go_folder"} | cancel."""
        dlg = tk.Toplevel(parent)
        dlg.title(f"{term} decision - {idx} of {total}")
        dlg.geometry("680x540")
        dlg.transient(parent)
        dlg.grab_set()
        result = {"decision": None}
        tk.Label(dlg, text=f"{term} decision - {idx} of {total}",
                 font=("Arial", 12, "bold")).pack(pady=(10, 4))
        path_frame = tk.LabelFrame(dlg, text=term, padx=8, pady=6)
        path_frame.pack(fill="x", padx=12, pady=4)
        tk.Label(path_frame, text=appendix_name, font=("Consolas", 10), wraplength=620).pack(anchor="w")
        info = f"Contains: {pdf_count} document(s) directly, {sub_count} immediate subfolder(s)"
        if max_subtree_depth > 0:
            info += f", subtree up to {max_subtree_depth} level(s) deep"
        tk.Label(dlg, text=info, font=("Arial", 9)).pack(pady=(4, 6))
        opt_frame = tk.LabelFrame(dlg, text=f"Default for this {term.lower()}", padx=10, pady=8)
        opt_frame.pack(fill="x", padx=12, pady=4)
        var_action = tk.StringVar(value="show")
        show_row = tk.Frame(opt_frame)
        show_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(show_row, text="SHOW ALL CONTENTS", variable=var_action, value="show",
                       font=("Arial", 9, "bold")).pack(side="left", anchor="w")
        tk.Label(show_row, text="depth limit:", font=("Arial", 9)).pack(side="left", padx=(20, 4))

        def labels(n):
            if n == 1:
                return f"1 ({term.lower()} line only - collapse all)"
            if n == 2:
                return "2 (top level only)"
            if n == 999:
                return "unlimited (every file at every depth)"
            return f"{n} ({n - 1} levels below {term.lower()})"

        depth_menu, var_depth = _depth_combo(show_row, max_subtree_depth, labels)
        depth_menu.pack(side="left")
        tk.Label(opt_frame, text=f"    Applies one rule to the whole {term.lower()}. No further popups.",
                 fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        tk.Radiobutton(opt_frame, text="GO FOLDER-BY-FOLDER", variable=var_action, value="go_folder",
                       font=("Arial", 9, "bold")).pack(anchor="w", pady=(4, 0))
        tk.Label(opt_frame, text=f"    Ask separately for each subfolder inside this {term.lower()}.\n"
                                 "    Useful when different branches need different treatment.",
                 fg="#555", font=("Arial", 9), justify="left").pack(anchor="w", padx=24)
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=12)

        def ok():
            action = var_action.get()
            result["decision"] = ({"action": "show", "depth": var_depth.get()}
                                  if action == "show" else {"action": "go_folder"})
            dlg.destroy()

        def cancel():
            result["decision"] = {"action": "cancel"}
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel build", width=14, command=cancel).pack(side="left", padx=6)
        dlg.wait_window()
        return result["decision"]

    def collect_manual_policy(parent, base_dir, overlay, term):
        """Run the manual per-section / per-folder dialogs. Returns policy or None."""
        _, appendix_folders = be.list_folder_items(base_dir, overlay)
        appendix_decisions = {}
        per_folder_decisions = {}
        for appx_idx, appendix_folder in enumerate(appendix_folders, start=1):
            appx_path = os.path.join(base_dir, appendix_folder)
            pdfs_here, subs_here = be.folder_summary(appx_path, overlay)
            app_decision = show_appendix_level_dialog(
                parent, appendix_folder, appx_idx, len(appendix_folders), pdfs_here, subs_here,
                be.max_depth_below(appx_path, overlay), term)
            if app_decision is None or app_decision.get("action") == "cancel":
                return None
            if app_decision["action"] == "show":
                appendix_decisions[appendix_folder] = app_decision
                continue
            sub_folders_in_appx = []

            def walk_subs(path, rel, depth):
                _, subs = be.list_folder_items(path, overlay)
                for nm in subs:
                    full = os.path.join(path, nm)
                    sub_rel = os.path.join(rel, nm) if rel else nm
                    if depth >= 1:
                        full_rel = f"{appendix_folder} / {sub_rel.replace(os.sep, ' / ')}"
                        sub_folders_in_appx.append((full, full_rel, nm))
                    walk_subs(full, sub_rel, depth + 1)

            walk_subs(appx_path, "", 1)
            bulk_decision = None
            for sub_idx, (sub_full, sub_rel_full, _nm) in enumerate(sub_folders_in_appx, start=1):
                if bulk_decision is not None:
                    per_folder_decisions[sub_rel_full] = bulk_decision
                    continue
                sub_pdfs, sub_subs = be.folder_summary(sub_full, overlay)
                decision, apply_all = show_manual_folder_dialog(
                    parent, sub_rel_full, sub_idx, len(sub_folders_in_appx), sub_pdfs, sub_subs,
                    be.max_depth_below(sub_full, overlay))
                if decision is None or decision.get("action") == "cancel":
                    return None
                per_folder_decisions[sub_rel_full] = decision
                if apply_all:
                    bulk_decision = decision
        return {"mode": "manual", "decisions": per_folder_decisions,
                "appendix_decisions": appendix_decisions}

    # -----------------------------------------------------------------------
    # Step 5: GUI
    # -----------------------------------------------------------------------

    NAVY, GOLD, GREY = "#051641", "#FFC425", "#d9d9d9"
    PALETTE_HEX = {k: "#%02x%02x%02x" % v["accent"] for k, v in be.PALETTES.items()}
    INK_HEX = {k: "#%02x%02x%02x" % v["ink"] for k, v in be.PALETTES.items()}

    class BundleBuilderApp:
        def __init__(self, root):
            self.root = root
            self.root.title("Document Bundle Builder")
            self.root.geometry("1020x1040")
            self.root.minsize(920, 760)
            self.settings = be.load_settings()
            self.cancel_event = threading.Event()
            self._building = False
            self._create_pending = False
            self._organiser = None
            self._build_ui()
            self._load_settings_into_ui()

        # ---------------- UI construction ----------------
        def _build_ui(self):
            pad = {"padx": 8, "pady": 4}
            tk.Label(self.root, text="QS Alaa Elsayed", font=("Arial", 8), fg="#666666",
                     anchor="e").pack(side="bottom", fill="x", padx=10, pady=(0, 4))
            act = tk.Frame(self.root)
            act.pack(side="bottom", fill="x", **pad)
            self.btn_run = tk.Button(act, text="Create Bundle", command=lambda: self._on_run(False),
                                     font=("Arial", 10, "bold"), bg="#1F3864", fg="white",
                                     padx=20, pady=6)
            self.btn_run.pack(side="left")
            self.btn_organise = tk.Button(act, text="Organise pages & Preview...",
                                          command=lambda: self._on_run(True),
                                          font=("Arial", 10, "bold"), padx=16, pady=6)
            self.btn_organise.pack(side="left", padx=8)
            self.btn_cancel = tk.Button(act, text="Cancel", command=self._on_cancel, padx=16,
                                        pady=6, state="disabled")
            self.btn_cancel.pack(side="left", padx=8)
            tk.Button(act, text="Preview style", command=self._on_preview, padx=16,
                      pady=6).pack(side="left", padx=8)
            tk.Button(act, text="Quit", command=self._on_quit, padx=20, pady=6).pack(side="right")
            self.txt_status = scrolledtext.ScrolledText(self.root, height=9, font=("Consolas", 9),
                                                        state="disabled")
            self.txt_status.pack(side="bottom", fill="x", **pad)
            tk.Label(self.root, text="Status / Progress:", font=("Arial", 9, "bold")).pack(
                side="bottom", fill="x", padx=8)

            container = tk.Frame(self.root)
            container.pack(side="top", fill="both", expand=True)
            self._opts_canvas = tk.Canvas(container, highlightthickness=0)
            vsb = tk.Scrollbar(container, orient="vertical", command=self._opts_canvas.yview)
            self._opts_canvas.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self._opts_canvas.pack(side="left", fill="both", expand=True)
            body = tk.Frame(self._opts_canvas)
            body_id = self._opts_canvas.create_window((0, 0), window=body, anchor="nw")

            def _sync_scroll(event=None):
                self._opts_canvas.configure(scrollregion=self._opts_canvas.bbox("all"))
                self._opts_canvas.itemconfigure(body_id, width=self._opts_canvas.winfo_width())

            body.bind("<Configure>", _sync_scroll)
            self._opts_canvas.bind("<Configure>", _sync_scroll)

            def _on_wheel(event):
                w = self.root.winfo_containing(event.x_root, event.y_root)
                while w is not None:
                    if w is self._opts_canvas:
                        self._opts_canvas.yview_scroll(int(-event.delta / 120), "units")
                        break
                    w = getattr(w, "master", None)

            self.root.bind_all("<MouseWheel>", _on_wheel)

            # Base directory
            dir_frame = tk.LabelFrame(body, text="Base Directory")
            dir_frame.pack(fill="x", **pad)
            row = tk.Frame(dir_frame)
            row.pack(fill="x", padx=6, pady=6)
            self.entry_dir = tk.Entry(row, font=("Consolas", 9))
            self.entry_dir.pack(side="left", fill="x", expand=True, padx=(0, 6))
            tk.Button(row, text="Browse...", command=self._pick_dir).pack(side="left")

            title_frame = tk.Frame(body)
            title_frame.pack(fill="x", **pad)
            tk.Label(title_frame, text="Bundle title:", font=("Arial", 9, "bold")).pack(side="left", padx=(8, 6))
            self.entry_claim_title = tk.Entry(title_frame, font=("Arial", 10))
            self.entry_claim_title.pack(side="left", fill="x", expand=True, padx=(0, 8))

            # Page extraction / output mode
            ext_frame = tk.LabelFrame(body, text="Pages to take from each file")
            ext_frame.pack(fill="x", **pad)
            erow = tk.Frame(ext_frame)
            erow.pack(fill="x", padx=6, pady=2)
            self.var_page_mode = tk.StringVar(value="all")
            for val in be.PAGE_SELECTION_MODES:
                tk.Radiobutton(erow, text=be.PAGE_SELECTION_LABELS[val], variable=self.var_page_mode,
                               value=val, command=self._toggle_page_mode).pack(side="left", padx=3)
            tk.Label(erow, text="N:", font=("Arial", 9)).pack(side="left", padx=(8, 2))
            self.spin_first_n = tk.Spinbox(erow, from_=1, to=999, width=4)
            self.spin_first_n.pack(side="left")
            erow2 = tk.Frame(ext_frame)
            erow2.pack(fill="x", padx=6, pady=2)
            tk.Label(erow2, text="Custom ranges (one per line, 'file name.pdf: 1-3,5'; a line "
                                 "'default: 1-2' applies to unlisted files):",
                     font=("Arial", 8), fg="#555").pack(anchor="w")
            self.txt_ranges = tk.Text(erow2, height=3, font=("Consolas", 9))
            self.txt_ranges.pack(fill="x")
            erow3 = tk.Frame(ext_frame)
            erow3.pack(fill="x", padx=6, pady=2)
            tk.Label(erow3, text="Output:", font=("Arial", 9)).pack(side="left")
            self.var_output_mode = tk.StringVar(value="bundle")
            tk.Radiobutton(erow3, text="One bundle (folders become sections; contents / separators as set below)",
                           variable=self.var_output_mode, value="bundle").pack(side="left", padx=3)
            tk.Radiobutton(erow3, text="Separate '<Folder>#Combined_FirstPages.pdf' per folder (like the scripts)",
                           variable=self.var_output_mode, value="per_folder").pack(side="left", padx=3)
            self.var_recursive = tk.BooleanVar(value=True)
            tk.Checkbutton(erow3, text="include subfolders", variable=self.var_recursive).pack(side="left", padx=8)
            tk.Label(ext_frame, text="'First page' always means the first SOURCE page of each file, never a "
                                     "generated cover or separator. Files named '*#Combined*' and earlier "
                                     "bundle outputs are never used as inputs and are never overwritten "
                                     "(a _rev.NN name is used). Order: _order.txt first, then natural "
                                     "file-name order (Dwg 2 before Dwg 10).",
                     fg="#555", font=("Arial", 8), justify="left", wraplength=940).pack(anchor="w", padx=8, pady=(0, 4))

            # Terminology
            term_frame = tk.LabelFrame(body, text="Section terminology")
            term_frame.pack(fill="x", **pad)
            trow = tk.Frame(term_frame)
            trow.pack(fill="x", padx=6, pady=4)
            tk.Label(trow, text="Sections are called:", font=("Arial", 9)).pack(side="left")
            self.var_term = tk.StringVar(value="Appendix")
            self.cmb_term = ttk.Combobox(trow, textvariable=self.var_term, state="readonly", width=12,
                                         values=list(be.SECTION_TERMS) + ["Custom"])
            self.cmb_term.pack(side="left", padx=6)
            self.cmb_term.bind("<<ComboboxSelected>>", lambda e: self._toggle_custom_term())
            self.entry_term_custom = tk.Entry(trow, font=("Arial", 9), width=16)
            self.entry_term_custom.pack(side="left", padx=(0, 12))
            tk.Label(trow, text="Numbering:", font=("Arial", 9)).pack(side="left")
            self.var_numbering = tk.StringVar(value="alphabetic")
            for val, lbl in (("alphabetic", "A, B, C"), ("numeric", "01, 02, 03"), ("none", "None")):
                tk.Radiobutton(trow, text=lbl, variable=self.var_numbering, value=val).pack(side="left", padx=3)
            tk.Label(term_frame, text="Used on separator pages, in the contents, bookmarks and footers "
                                      "(e.g. 'Appendix A', 'Drawing 01', 'Section 3').",
                     fg="#555", font=("Arial", 8)).pack(anchor="w", padx=8, pady=(0, 4))

            # Separator pages
            sep_frame = tk.LabelFrame(body, text="Separator pages")
            sep_frame.pack(fill="x", **pad)
            self.var_sep_scope = tk.StringVar(value="all_sections")
            srow = tk.Frame(sep_frame)
            srow.pack(fill="x", padx=6, pady=2)
            for val in be.SEPARATOR_SCOPES:
                tk.Radiobutton(srow, text=be.SEPARATOR_SCOPE_LABELS[val], variable=self.var_sep_scope,
                               value=val, font=("Arial", 9)).pack(side="left", padx=4)
            tk.Label(sep_frame, text="Contents entries and bookmarks are kept whatever you choose; "
                                     "without a separator they open the first page of the section.",
                     fg="#555", font=("Arial", 8)).pack(anchor="w", padx=8, pady=(0, 4))

            # Contents / bookmarks / options
            opt_frame = tk.LabelFrame(body, text="Contents, bookmarks and options")
            opt_frame.pack(fill="x", **pad)
            r = tk.Frame(opt_frame)
            r.pack(fill="x", padx=6, pady=2)
            self.var_toc = tk.BooleanVar(value=True)
            self.var_toc_docs = tk.BooleanVar(value=True)
            self.var_bookmarks = tk.BooleanVar(value=True)
            self.var_bm_docs = tk.BooleanVar(value=True)
            tk.Checkbutton(r, text="Table of Contents (clickable, multi-page safe)", variable=self.var_toc).pack(side="left")
            tk.Checkbutton(r, text="list documents (not just sections)", variable=self.var_toc_docs).pack(side="left", padx=(6, 18))
            tk.Checkbutton(r, text="PDF bookmarks", variable=self.var_bookmarks).pack(side="left")
            tk.Checkbutton(r, text="include documents", variable=self.var_bm_docs).pack(side="left", padx=6)
            r2 = tk.Frame(opt_frame)
            r2.pack(fill="x", padx=6, pady=2)
            self.var_manual = tk.BooleanVar(value=False)
            tk.Checkbutton(r2, text="Manual mode (prompt for a description of each section)",
                           variable=self.var_manual).pack(side="left")
            self.var_ocr = tk.BooleanVar(value=False)
            ocr_text = "Run OCR on scanned pages (English + Arabic)" + (
                "" if OCR_AVAILABLE else " - NOT AVAILABLE (Pillow + pytesseract required)")
            ocr_cb = tk.Checkbutton(r2, text=ocr_text, variable=self.var_ocr)
            ocr_cb.pack(side="left", padx=(18, 0))
            if not OCR_AVAILABLE:
                ocr_cb.configure(state="disabled")
            r3 = tk.Frame(opt_frame)
            r3.pack(fill="x", padx=6, pady=2)
            tk.Label(r3, text="Compression:", font=("Arial", 9)).pack(side="left")
            self.var_compression = tk.StringVar(value="high")
            for val, lbl in (("high", "High quality (vector drawings untouched)"), ("balanced", "Balanced"),
                             ("small", "Smallest file")):
                tk.Radiobutton(r3, text=lbl, variable=self.var_compression, value=val).pack(side="left", padx=4)
            tk.Label(r3, text="Default folder layout:", font=("Arial", 9)).pack(side="left", padx=(18, 4))
            self.var_layout = tk.StringVar(value="preserve_skip_numeric")
            ttk.Combobox(r3, textvariable=self.var_layout, state="readonly", width=22,
                         values=["preserve", "preserve_skip_numeric", "flatten"]).pack(side="left")

            # Footer
            foot_frame = tk.LabelFrame(body, text="Page footers")
            foot_frame.pack(fill="x", **pad)
            fr = tk.Frame(foot_frame)
            fr.pack(fill="x", padx=6, pady=2)
            self.var_footer = tk.BooleanVar(value=True)
            tk.Checkbutton(fr, text="Add section / page-of-pages footers", variable=self.var_footer).pack(side="left")
            tk.Label(fr, text="Position:", font=("Arial", 9)).pack(side="left", padx=(18, 4))
            self.var_footer_pos = tk.StringVar(value="bottom")
            for val, lbl in (("bottom", "Bottom edge"), ("top", "Top edge")):
                tk.Radiobutton(fr, text=lbl, variable=self.var_footer_pos, value=val).pack(side="left", padx=2)
            tk.Label(fr, text="Layout:", font=("Arial", 9)).pack(side="left", padx=(18, 4))
            self.var_footer_layout = tk.StringVar(value="split")
            for val, lbl in (("split", "Left + right"), ("left", "Left"), ("centre", "Centre"), ("right", "Right")):
                tk.Radiobutton(fr, text=lbl, variable=self.var_footer_layout, value=val).pack(side="left", padx=2)
            tk.Label(fr, text="Inset (mm):", font=("Arial", 9)).pack(side="left", padx=(18, 4))
            self.entry_footer_inset = tk.Entry(fr, width=5, font=("Arial", 9))
            self.entry_footer_inset.pack(side="left")
            tk.Label(foot_frame, text="Drawings: switch footers off, or use 'Top edge' / 'Left' so the "
                                      "bottom-right title block is not overprinted. Page size, rotation "
                                      "and vector content are always preserved.",
                     fg="#555", font=("Arial", 8)).pack(anchor="w", padx=8, pady=(0, 4))

            # Cover page
            cover_frame = tk.LabelFrame(body, text="Cover Page (optional)")
            cover_frame.pack(fill="x", **pad)
            self.var_cover = tk.BooleanVar(value=False)
            tk.Checkbutton(cover_frame, text="Include cover page", variable=self.var_cover,
                           command=self._toggle_cover_fields).pack(anchor="w", padx=6)
            self.cover_fields = tk.Frame(cover_frame)
            self.cover_fields.pack(fill="x", padx=6, pady=4)
            self.cover_entries = {}
            for i, (label, key) in enumerate([
                ("Document type", "cover_doc_type"), ("Project", "cover_project"),
                ("Contract Reference", "cover_contract_ref"), ("Prepared by", "cover_prepared_by"),
                ("Prepared for", "cover_prepared_for"), ("Revision", "cover_revision"),
            ]):
                tk.Label(self.cover_fields, text=f"{label}:", font=("Arial", 9), width=18,
                         anchor="w").grid(row=i, column=0, padx=2, pady=2, sticky="w")
                e = tk.Entry(self.cover_fields, font=("Arial", 9))
                e.grid(row=i, column=1, padx=2, pady=2, sticky="ew")
                self.cover_entries[key] = e
            self.cover_fields.columnconfigure(1, weight=1)

            # Bundle style
            style_frame = tk.LabelFrame(body, text="Bundle Style")
            style_frame.pack(fill="x", **pad)
            prow = tk.Frame(style_frame)
            prow.pack(fill="x", padx=6, pady=3)
            tk.Label(prow, text="Preset:", font=("Arial", 9), width=12, anchor="w").pack(side="left")
            self.var_preset = tk.StringVar(value="claims_navy")
            self._preset_names = {v["name"]: k for k, v in be.STYLE_PRESETS.items()}
            self.cmb_preset = ttk.Combobox(prow, state="readonly", width=34,
                                           values=list(self._preset_names))
            self.cmb_preset.pack(side="left")
            tk.Button(prow, text="Apply preset", command=self._apply_preset).pack(side="left", padx=8)
            tk.Label(prow, text="(sets the pickers below, the terminology and footer defaults; "
                                "everything stays editable)", fg="#555", font=("Arial", 8)).pack(side="left")

            self.var_style_cover = tk.StringVar(value="top_rule")
            self.var_style_palette = tk.StringVar(value="navy_gold")
            self.var_style_font = tk.StringVar(value="aptos")
            self.var_style_separator = tk.StringVar(value="top_rule")
            self.var_style_toc = tk.StringVar(value="ruled")
            self.var_toc_density = tk.StringVar(value="standard")
            self._style_refreshers = []

            def ink():
                return INK_HEX.get(self.var_style_palette.get(), NAVY)

            def accent():
                return PALETTE_HEX.get(self.var_style_palette.get(), GOLD)

            def draw_cover(cv, value, _acc):
                ink_c, acc = ink(), accent()
                if value == "top_rule":
                    cv.create_rectangle(0, 0, 64, 3, fill=ink_c, outline="")
                    cv.create_rectangle(8, 16, 44, 20, fill=ink_c, outline="")
                    cv.create_rectangle(8, 22, 36, 26, fill=ink_c, outline="")
                    cv.create_rectangle(8, 30, 20, 32, fill=acc, outline="")
                    for k, w in enumerate((30, 24, 27)):
                        cv.create_rectangle(8, 62 + k * 7, 8 + w, 64 + k * 7, fill=GREY, outline="")
                elif value == "classical":
                    cv.create_rectangle(16, 20, 48, 24, fill=ink_c, outline="")
                    cv.create_rectangle(20, 26, 44, 30, fill=ink_c, outline="")
                    cv.create_rectangle(28, 34, 36, 35, fill=ink_c, outline="")
                    for k, w in enumerate((28, 22)):
                        cv.create_rectangle(32 - w // 2, 64 + k * 7, 32 + w // 2, 66 + k * 7, fill=GREY, outline="")
                elif value == "minimal":
                    cv.create_rectangle(8, 28, 20, 29, fill=GREY, outline="")
                    cv.create_rectangle(8, 33, 44, 37, fill=ink_c, outline="")
                    cv.create_rectangle(8, 40, 30, 42, fill=GREY, outline="")
                    cv.create_line(8, 62, 56, 62, fill=GREY)
                    for k, w in enumerate((30, 24)):
                        cv.create_rectangle(8, 65 + k * 6, 8 + w, 67 + k * 6, fill=GREY, outline="")
                elif value == "block":
                    cv.create_rectangle(0, 0, 64, 40, fill=ink_c, outline="")
                    cv.create_rectangle(8, 14, 40, 18, fill="white", outline="")
                    cv.create_rectangle(8, 22, 30, 25, fill="white", outline="")
                    cv.create_rectangle(8, 40, 18, 42, fill=acc, outline="")
                    for k, w in enumerate((30, 24)):
                        cv.create_rectangle(8, 60 + k * 7, 8 + w, 62 + k * 7, fill=GREY, outline="")
                else:  # register
                    cv.create_rectangle(0, 0, 64, 10, fill=ink_c, outline="")
                    cv.create_rectangle(0, 10, 64, 11, fill=acc, outline="")
                    cv.create_rectangle(8, 18, 44, 22, fill=ink_c, outline="")
                    for rr in range(2):
                        for cc in range(2):
                            cv.create_rectangle(8 + cc * 24, 30 + rr * 10, 32 + cc * 24, 40 + rr * 10, outline=GREY)
                    for k in range(3):
                        cv.create_line(8, 58 + k * 6, 56, 58 + k * 6, fill=GREY)

            def draw_palette(cv, value, _acc):
                cv.create_rectangle(8, 14, 56, 44, fill=INK_HEX[value], outline="")
                cv.create_rectangle(8, 50, 56, 66, fill=PALETTE_HEX[value], outline="")

            def draw_separator(cv, value, _acc):
                ink_c, acc = ink(), accent()
                if value == "top_rule":
                    cv.create_rectangle(0, 0, 64, 3, fill=ink_c, outline="")
                    cv.create_rectangle(8, 34, 32, 37, fill=ink_c, outline="")
                    cv.create_rectangle(8, 41, 44, 45, fill=ink_c, outline="")
                    cv.create_rectangle(8, 49, 20, 51, fill=acc, outline="")
                elif value == "plaque":
                    cv.create_rectangle(12, 26, 52, 62, outline=ink_c)
                    cv.create_rectangle(20, 34, 44, 37, fill=ink_c, outline="")
                    cv.create_rectangle(18, 42, 46, 45, fill=ink_c, outline="")
                    cv.create_rectangle(22, 49, 42, 52, fill=ink_c, outline="")
                elif value == "minimal":
                    cv.create_rectangle(24, 36, 40, 37, fill=GREY, outline="")
                    cv.create_rectangle(16, 41, 48, 45, fill=ink_c, outline="")
                elif value == "band":
                    cv.create_rectangle(0, 0, 64, 14, fill=ink_c, outline="")
                    cv.create_rectangle(0, 14, 64, 15, fill=acc, outline="")
                    cv.create_rectangle(8, 5, 30, 8, fill="white", outline="")
                    cv.create_rectangle(8, 26, 44, 30, fill=ink_c, outline="")
                    cv.create_line(8, 70, 56, 70, fill=GREY)
                else:  # block
                    cv.create_rectangle(0, 0, 64, 30, fill=ink_c, outline="")
                    cv.create_rectangle(8, 12, 40, 16, fill="white", outline="")
                    cv.create_rectangle(8, 20, 28, 23, fill="white", outline="")
                    cv.create_rectangle(8, 30, 18, 32, fill=acc, outline="")

            def draw_toc(cv, value, _acc):
                ink_c, acc = ink(), accent()
                cv.create_rectangle(8, 10, 28, 14, fill=ink_c, outline="")
                if value != "plain":
                    cv.create_rectangle(8, 16, 20, 18, fill=acc, outline="")
                if value == "grid":
                    cv.create_line(8, 26, 56, 26, fill=ink_c)
                for k in range(3):
                    y = 30 + k * 16
                    if value == "ruled":
                        cv.create_rectangle(8, y, 14, y + 6, fill=ink_c, outline="")
                    elif value == "grid":
                        cv.create_rectangle(8, y + 1, 13, y + 4, fill=ink_c, outline="")
                    cv.create_rectangle(17, y + 1, 38, y + 4, fill=ink_c, outline="")
                    cv.create_rectangle(50, y + 1, 56, y + 4, fill=GREY, outline="")
                    if value in ("ruled", "grid"):
                        cv.create_line(8, y + 9, 56, y + 9, fill=ink_c if value == "ruled" else GREY)
                    elif value == "dotted":
                        for dx in range(40, 49, 3):
                            cv.create_rectangle(dx, y + 4, dx + 1, y + 5, fill=GREY, outline="")

            def add_style_row(label, var, options, draw_fn, height=78):
                row = tk.Frame(style_frame)
                row.pack(fill="x", padx=6, pady=3, anchor="w")
                tk.Label(row, text=label, font=("Arial", 9), width=12, anchor="w").pack(side="left")
                canvases = []

                def refresh():
                    for value, cv in canvases:
                        cv.delete("all")
                        draw_fn(cv, value, None)
                        sel = var.get() == value
                        cv.configure(highlightbackground="#0045BF" if sel else "#c8c8c8",
                                     highlightthickness=2 if sel else 1)

                for value, caption in options:
                    holder = tk.Frame(row)
                    holder.pack(side="left", padx=5)
                    cv = tk.Canvas(holder, width=64, height=height, bg="white", highlightthickness=1,
                                   highlightbackground="#c8c8c8", cursor="hand2")
                    cv.pack()
                    cv.bind("<Button-1>", lambda e, v=value: (var.set(v), self._refresh_style_thumbs()))
                    tk.Label(holder, text=caption, font=("Arial", 8)).pack()
                    canvases.append((value, cv))
                self._style_refreshers.append(refresh)
                refresh()

            add_style_row("Cover:", self.var_style_cover, list(be.COVER_STYLES.items()), draw_cover)
            add_style_row("Palette:", self.var_style_palette,
                          [(k, v["name"]) for k, v in be.PALETTES.items()], draw_palette, height=72)
            add_style_row("Separator:", self.var_style_separator, list(be.SEPARATOR_STYLES.items()),
                          draw_separator)
            add_style_row("Contents:", self.var_style_toc, list(be.TOC_STYLES.items()), draw_toc)
            font_row = tk.Frame(style_frame)
            font_row.pack(fill="x", padx=6, pady=3, anchor="w")
            tk.Label(font_row, text="Font:", font=("Arial", 9), width=12, anchor="w").pack(side="left")
            tk.Radiobutton(font_row, text="Aptos (brand, falls back to Segoe UI)", variable=self.var_style_font,
                           value="aptos").pack(side="left", padx=4)
            tk.Radiobutton(font_row, text="Times New Roman (court bundle)", variable=self.var_style_font,
                           value="times").pack(side="left", padx=4)
            tk.Label(font_row, text="Contents density:", font=("Arial", 9)).pack(side="left", padx=(18, 4))
            for val, lbl in (("compact", "Compact"), ("standard", "Standard"), ("spacious", "Spacious")):
                tk.Radiobutton(font_row, text=lbl, variable=self.var_toc_density, value=val).pack(side="left", padx=2)

            # Office documents
            off_frame = tk.LabelFrame(body, text="Word / Excel documents (optional)")
            off_frame.pack(fill="x", **pad)
            orow = tk.Frame(off_frame)
            orow.pack(fill="x", padx=6, pady=2)
            self.var_office = tk.BooleanVar(value=False)
            tk.Checkbutton(orow, text="Convert Word (.doc/.docx) and Excel (.xls/.xlsx/.xlsm) files found in the "
                                      "folder tree and bundle them in place", variable=self.var_office).pack(side="left")
            self.var_office_reconvert = tk.BooleanVar(value=False)
            tk.Checkbutton(orow, text="re-convert all", variable=self.var_office_reconvert).pack(side="left", padx=12)
            orow2 = tk.Frame(off_frame)
            orow2.pack(fill="x", padx=6, pady=2)
            tk.Label(orow2, text="Excel Quick Print:", font=("Arial", 9)).pack(side="left")
            self.entry_xlq = tk.Entry(orow2, font=("Consolas", 9))
            self.entry_xlq.pack(side="left", fill="x", expand=True, padx=6)
            tk.Button(orow2, text="Browse...", command=self._pick_xlq).pack(side="left")
            ok, reason = oc.office_support_status()
            found = oc.find_excel_quick_print()
            tk.Label(off_frame, text=(reason + ("  Excel Quick Print found: " + os.path.basename(found)
                                                if found else "  Excel Quick Print not found next to this script.")
                                      + "\nWord converts silently in an isolated Word instance (read-only, "
                                        "macros off, links not refreshed). Excel opens the Excel Quick Print "
                                        "wizard so you choose sheets, print areas and preview as usual. "
                                        "Converted PDFs are kept in '_bundle_converted' and reused while "
                                        "newer than the source. Source files are never modified."),
                     fg="#555" if ok else "#8a4a00", font=("Arial", 8), justify="left",
                     wraplength=940).pack(anchor="w", padx=8, pady=(0, 4))

        def _refresh_style_thumbs(self):
            for refresh in self._style_refreshers:
                refresh()

        def _toggle_page_mode(self):
            mode = self.var_page_mode.get()
            self.spin_first_n.configure(state="normal" if mode == "first_n" else "disabled")
            self.txt_ranges.configure(state="normal" if mode == "ranges" else "disabled")

        def _page_selection(self):
            mode = self.var_page_mode.get()
            sel = {"mode": mode}
            if mode == "first_n":
                try:
                    sel["n"] = max(1, int(self.spin_first_n.get()))
                except ValueError:
                    sel["n"] = 1
            if mode == "ranges":
                ranges, default = {}, ""
                for line in self.txt_ranges.get("1.0", "end").splitlines():
                    if ":" not in line:
                        continue
                    name, spec = line.rsplit(":", 1)
                    name, spec = name.strip(), spec.strip()
                    if name.lower() == "default":
                        default = spec
                    elif name:
                        ranges[name] = spec
                sel["ranges"] = ranges
                sel["default_range"] = default
            return sel

        def _toggle_custom_term(self):
            self.entry_term_custom.configure(state="normal" if self.var_term.get() == "Custom" else "disabled")

        def _toggle_cover_fields(self):
            state = "normal" if self.var_cover.get() else "disabled"
            for e in self.cover_entries.values():
                e.configure(state=state)

        def _apply_preset(self):
            key = self._preset_names.get(self.cmb_preset.get())
            if not key:
                return
            preset = be.STYLE_PRESETS[key]
            st = preset["style"]
            self.var_style_cover.set(st["cover"])
            self.var_style_palette.set(st["palette"])
            self.var_style_font.set(st["font"])
            self.var_style_separator.set(st["separator"])
            self.var_style_toc.set(st["toc"])
            self.var_toc_density.set(st.get("toc_density", "standard"))
            term = preset["terminology"]
            self.var_term.set(term["section_term"])
            self.var_numbering.set(term["numbering_style"])
            self._toggle_custom_term()
            dt = self.cover_entries["cover_doc_type"]
            state = dt.cget("state")
            dt.configure(state="normal")
            dt.delete(0, "end")
            dt.insert(0, term.get("cover_doc_type", ""))
            dt.configure(state=state)
            foot = preset.get("footer", {})
            self.var_footer.set(foot.get("footer_enabled", True))
            self.var_footer_pos.set(foot.get("footer_position", "bottom"))
            self.var_footer_layout.set(foot.get("footer_layout", "split"))
            self.var_preset.set(key)
            self._refresh_style_thumbs()

        # ---------------- settings ----------------
        def _load_settings_into_ui(self):
            s = self.settings
            self.entry_dir.insert(0, s.get("last_directory", ""))
            self.entry_claim_title.insert(0, s.get("claim_title", ""))
            self.var_term.set(s.get("section_term", "Appendix"))
            self.entry_term_custom.insert(0, s.get("section_term_custom", ""))
            self.var_numbering.set(s.get("numbering_style", "alphabetic"))
            self.var_sep_scope.set(s.get("separator_scope", "all_sections"))
            self.var_toc.set(s.get("generate_toc", True))
            self.var_toc_docs.set(s.get("toc_include_documents", True))
            self.var_bookmarks.set(s.get("generate_bookmarks", True))
            self.var_bm_docs.set(s.get("bookmark_include_documents", True))
            self.var_manual.set(s.get("manual_descriptions", False))
            self.var_ocr.set(s.get("do_ocr", False) and OCR_AVAILABLE)
            self.var_compression.set(s.get("compression", "high"))
            self.var_layout.set(s.get("folder_layout", "preserve_skip_numeric"))
            self.var_footer.set(s.get("footer_enabled", True))
            self.var_footer_pos.set(s.get("footer_position", "bottom"))
            self.var_footer_layout.set(s.get("footer_layout", "split"))
            self.entry_footer_inset.insert(0, str(s.get("footer_inset_mm", 6.35)))
            self.var_cover.set(s.get("include_cover", False))
            for key, entry in self.cover_entries.items():
                if key in ("cover_prepared_by", "cover_prepared_for"):
                    continue  # identity fields start blank every launch
                entry.insert(0, s.get(key, ""))
            self.var_style_cover.set(s.get("style_cover", "top_rule"))
            self.var_style_palette.set(s.get("style_palette", "navy_gold"))
            self.var_style_font.set(s.get("style_font", "aptos"))
            self.var_style_separator.set(s.get("style_separator", "top_rule"))
            self.var_style_toc.set(s.get("style_toc", "ruled"))
            self.var_toc_density.set(s.get("style_toc_density", "standard"))
            preset = s.get("style_preset", "claims_navy")
            if preset in be.STYLE_PRESETS:
                self.cmb_preset.set(be.STYLE_PRESETS[preset]["name"])
            self.var_office.set(s.get("convert_office", False))
            self.var_office_reconvert.set(s.get("office_reconvert", False))
            self.entry_xlq.insert(0, s.get("excel_quick_print_path", ""))
            self.var_page_mode.set(s.get("page_mode", "all"))
            self.spin_first_n.delete(0, "end")
            self.spin_first_n.insert(0, str(s.get("page_first_n", 1)))
            self.var_output_mode.set(s.get("output_mode", "bundle"))
            self.var_recursive.set(s.get("recursive", True))
            self._toggle_page_mode()
            self._refresh_style_thumbs()
            self._toggle_cover_fields()
            self._toggle_custom_term()

        def _save_current_settings(self):
            s = self.settings
            s["last_directory"] = self.entry_dir.get().strip()
            s["claim_title"] = self.entry_claim_title.get().strip()
            s["section_term"] = self.var_term.get()
            s["section_term_custom"] = self.entry_term_custom.get().strip()
            s["numbering_style"] = self.var_numbering.get()
            s["separator_scope"] = self.var_sep_scope.get()
            s["generate_toc"] = self.var_toc.get()
            s["toc_include_documents"] = self.var_toc_docs.get()
            s["generate_bookmarks"] = self.var_bookmarks.get()
            s["bookmark_include_documents"] = self.var_bm_docs.get()
            s["manual_descriptions"] = self.var_manual.get()
            s["do_ocr"] = self.var_ocr.get()
            s["compression"] = self.var_compression.get()
            s["folder_layout"] = self.var_layout.get()
            s["footer_enabled"] = self.var_footer.get()
            s["footer_position"] = self.var_footer_pos.get()
            s["footer_layout"] = self.var_footer_layout.get()
            s["footer_inset_mm"] = self._footer_inset()
            s["include_cover"] = self.var_cover.get()
            for key, entry in self.cover_entries.items():
                if key in ("cover_prepared_by", "cover_prepared_for"):
                    continue  # identity fields are never persisted
                s[key] = entry.get().strip()
            s["style_cover"] = self.var_style_cover.get()
            s["style_palette"] = self.var_style_palette.get()
            s["style_font"] = self.var_style_font.get()
            s["style_separator"] = self.var_style_separator.get()
            s["style_toc"] = self.var_style_toc.get()
            s["style_toc_density"] = self.var_toc_density.get()
            s["style_preset"] = self._preset_names.get(self.cmb_preset.get(), "claims_navy")
            s["convert_office"] = self.var_office.get()
            s["office_reconvert"] = self.var_office_reconvert.get()
            s["excel_quick_print_path"] = self.entry_xlq.get().strip()
            s["page_mode"] = self.var_page_mode.get()
            try:
                s["page_first_n"] = max(1, int(self.spin_first_n.get()))
            except ValueError:
                s["page_first_n"] = 1
            s["output_mode"] = self.var_output_mode.get()
            s["recursive"] = self.var_recursive.get()
            be.save_settings(s)

        def _footer_inset(self):
            try:
                v = float(self.entry_footer_inset.get().strip())
                return max(2.0, min(v, 60.0))
            except ValueError:
                return 6.35

        def _current_style(self):
            return {"cover": self.var_style_cover.get(), "palette": self.var_style_palette.get(),
                    "font": self.var_style_font.get(), "separator": self.var_style_separator.get(),
                    "toc": self.var_style_toc.get(), "toc_density": self.var_toc_density.get()}

        def _effective_term(self):
            return be.effective_term({"section_term": self.var_term.get(),
                                      "section_term_custom": self.entry_term_custom.get()})

        # ---------------- small handlers ----------------
        def _pick_dir(self):
            path = filedialog.askdirectory(title="Select base directory containing PDFs / folders")
            if path:
                self.entry_dir.delete(0, "end")
                self.entry_dir.insert(0, path)

        def _pick_xlq(self):
            path = filedialog.askopenfilename(title="Locate Excel_Quick_Print.pyw",
                                              filetypes=[("Python", "*.pyw;*.py"), ("All files", "*.*")])
            if path:
                self.entry_xlq.delete(0, "end")
                self.entry_xlq.insert(0, path)

        def _set_status(self, msg, append=False):
            self.txt_status.configure(state="normal")
            if not append:
                self.txt_status.delete("1.0", "end")
            self.txt_status.insert("end", msg + "\n")
            self.txt_status.see("end")
            self.txt_status.configure(state="disabled")

        def _log_safe(self, msg):
            self.root.after(0, lambda m=msg: self._set_status(m, append=True))

        def _ask_yesno_safe(self, title, message):
            """Ask from the worker thread; blocks the worker until answered."""
            done = threading.Event()
            answer = {"value": False}

            def ask():
                answer["value"] = messagebox.askyesno(title, message, parent=self.root)
                done.set()

            self.root.after(0, ask)
            done.wait()
            return answer["value"]

        def _on_quit(self):
            self._save_current_settings()
            self.root.quit()

        def _on_cancel(self):
            self.cancel_event.set()
            self._set_status("Cancel requested - finishing the current step...", append=True)

        def _on_preview(self):
            try:
                self._save_current_settings()
                meta = {"title": self.entry_claim_title.get().strip() or "Style preview",
                        "doc_type": self.cover_entries["cover_doc_type"].get().strip(),
                        "project": self.cover_entries["cover_project"].get().strip(),
                        "contract_ref": self.cover_entries["cover_contract_ref"].get().strip(),
                        "revision": self.cover_entries["cover_revision"].get().strip()}
                data = be.render_style_preview(self._current_style(), self._effective_term(),
                                               self.var_numbering.get(), meta)
                import tempfile
                path = os.path.join(tempfile.gettempdir(),
                                    f"bundle_style_preview_{datetime.now():%H%M%S}.pdf")
                with open(path, "wb") as f:
                    f.write(data)
                self._set_status(f"Style preview written: {path}", append=True)
                try:
                    os.startfile(path)
                except Exception:
                    messagebox.showinfo("Preview", f"Preview saved to:\n{path}")
            except Exception as e:
                messagebox.showerror("Preview failed", str(e))

        # ---------------- build ----------------
        def _on_run(self, organise=False):
            base_dir = self.entry_dir.get().strip()
            if not base_dir or not os.path.isdir(base_dir):
                messagebox.showwarning("Invalid directory", "Pick a valid base directory.")
                return
            term = self._effective_term()
            be.set_current_terms(term, self.var_numbering.get())
            per_folder = self.var_output_mode.get() == "per_folder"
            if per_folder and organise:
                messagebox.showinfo("Per-folder mode",
                                    "The page organiser works on a single bundle. Per-folder combined "
                                    "files are produced straight through the same pipeline with the "
                                    "chosen page mode. Switch the output to 'One bundle' to organise pages.")
                return

            if self.var_cover.get():
                for key, label in (("cover_prepared_by", "Prepared by"),
                                   ("cover_prepared_for", "Prepared for")):
                    if not self.cover_entries[key].get().strip():
                        val = simpledialog.askstring("Cover details",
                                                     f"{label} (leave blank to omit from the cover):",
                                                     parent=self.root)
                        if val and val.strip():
                            self.cover_entries[key].delete(0, "end")
                            self.cover_entries[key].insert(0, val.strip())

            claim_title = self.entry_claim_title.get().strip()
            if not claim_title:
                if not messagebox.askyesno("No bundle title",
                                           "Bundle title is empty - file will be named "
                                           "'Combined_Appendices.pdf'. Continue?"):
                    return

            # Office documents are counted as content before conversion so the
            # folder dialogs see them; the real overlay is built in the worker.
            convert_office = self.var_office.get()
            overlay = oc.planned_overlay(base_dir) if convert_office else {}
            if convert_office and not overlay:
                self._set_status("No Word/Excel documents found - conversion stage will be skipped.")

            counts = be.count_folder_tree(base_dir, overlay)
            policy = None
            if per_folder or not self.var_recursive.get():
                policy = {"mode": "automatic", "layout": "flatten", "decisions": {},
                          "appendix_decisions": {}}
            elif counts["top_folders"] == 0 and counts["total_subfolders"] == 0:
                policy = {"mode": "automatic", "layout": "preserve", "decisions": {},
                          "appendix_decisions": {}}
            else:
                while True:
                    stage1 = show_discovery_popup(self.root, counts, term)
                    if stage1 is None:
                        return
                    if stage1 == "automatic":
                        result = show_automatic_options_dialog(self.root, self.var_layout.get())
                        if result == "BACK":
                            continue
                        if result is None:
                            return
                        policy = {"mode": "automatic", "layout": result["layout"], "decisions": {},
                                  "appendix_decisions": {}}
                        self.var_layout.set(result["layout"])
                        break
                    policy = collect_manual_policy(self.root, base_dir, overlay, term)
                    if policy is None:
                        return
                    break

            manual_titles = {}
            if self.var_manual.get():
                docs, folders = be.list_folder_items(base_dir, overlay)
                items_for_prompt = [("file", name, stem) for name, stem, _ in docs] + \
                                   [("folder", f, f) for f in folders]
                for kind, name, default in items_for_prompt:
                    custom = simpledialog.askstring(f"Description for {kind}",
                                                    f"Description for '{name}'\n(leave blank to use '{default}'):",
                                                    initialvalue=default, parent=self.root)
                    if custom is None:
                        return
                    manual_titles[name] = custom.strip() or default

            self._save_current_settings()
            opts = {
                "base_directory": base_dir,
                "claim_title": claim_title,
                "section_term": term,
                "numbering_style": self.var_numbering.get(),
                "manual_descriptions": self.var_manual.get(),
                "manual_titles": manual_titles,
                "generate_toc": self.var_toc.get(),
                "toc_include_documents": self.var_toc_docs.get(),
                "generate_bookmarks": self.var_bookmarks.get(),
                "bookmark_include_documents": self.var_bm_docs.get(),
                "separator_scope": self.var_sep_scope.get(),
                "do_ocr": self.var_ocr.get(),
                "compression": self.var_compression.get(),
                "footer_enabled": self.var_footer.get(),
                "footer_position": self.var_footer_pos.get(),
                "footer_layout": self.var_footer_layout.get(),
                "footer_inset_mm": self._footer_inset(),
                "include_cover": self.var_cover.get(),
                "cover_subtitle": "",
                "cover_doc_type": self.cover_entries["cover_doc_type"].get().strip(),
                "cover_project": self.cover_entries["cover_project"].get().strip(),
                "cover_contract_ref": self.cover_entries["cover_contract_ref"].get().strip(),
                "cover_prepared_by": self.cover_entries["cover_prepared_by"].get().strip(),
                "cover_prepared_for": self.cover_entries["cover_prepared_for"].get().strip(),
                "cover_revision": self.cover_entries["cover_revision"].get().strip(),
                "cover_date": "",
                "style": self._current_style(),
                "policy": policy,
                "convert_office": convert_office,
                "office_reconvert": self.var_office_reconvert.get(),
                "excel_quick_print_path": self.entry_xlq.get().strip(),
                "page_selection": self._page_selection(),
                "direct_files_only": not self.var_recursive.get(),
                "recursive": self.var_recursive.get(),
                "per_folder": per_folder,
                "organise": organise,
            }
            policy_summary = "Folder mode: " + policy["mode"]
            if policy["mode"] == "automatic":
                policy_summary += f" / {policy['layout']}"
            confirm = (
                f"Base: {base_dir}\n"
                f"Title: {claim_title or '(default name)'}\n"
                f"Sections: {term} / numbering {opts['numbering_style']}\n"
                f"Separators: {be.SEPARATOR_SCOPE_LABELS[opts['separator_scope']]}\n"
                f"TOC: {'Yes' if opts['generate_toc'] else 'No'} | "
                f"Bookmarks: {'Yes' if opts['generate_bookmarks'] else 'No'} | "
                f"Footers: {'Yes' if opts['footer_enabled'] else 'No'}\n"
                f"OCR: {'Yes' if opts['do_ocr'] else 'No'} | Compression: {opts['compression']}\n"
                f"Cover page: {'Yes' if opts['include_cover'] else 'No'}\n"
                f"Office conversion: {'Yes' if convert_office else 'No'}\n"
                f"Pages: {be.PAGE_SELECTION_LABELS[opts['page_selection']['mode']]}"
                f"{' (N=' + str(opts['page_selection'].get('n', 1)) + ')' if opts['page_selection']['mode'] == 'first_n' else ''}\n"
                f"Output: {'separate #Combined_FirstPages.pdf per folder' if per_folder else 'one bundle'}"
                f"{'' if self.var_recursive.get() else ' (this folder only, no subfolders)'}\n"
                f"{policy_summary}\n\n"
                + ("Continue to the page organiser?" if organise else "Proceed?")
            )
            if not messagebox.askyesno("Confirm build", confirm):
                return

            self.cancel_event.clear()
            opts["cancel_event"] = self.cancel_event
            self._set_building(True)
            self._set_status("Starting..." if organise else "Starting bundle build...")
            threading.Thread(target=self._worker, args=(opts,), daemon=True).start()

        def _set_building(self, flag):
            self._building = flag
            state = "disabled" if flag else "normal"
            self.btn_run.configure(state=state)
            self.btn_organise.configure(state=state)
            self.btn_cancel.configure(state="normal" if flag else "disabled")

        def _worker(self, opts):
            log = self._log_safe
            try:
                if opts.get("convert_office"):
                    outcome = oc.run_conversion_stage(opts["base_directory"], opts, log,
                                                      self.cancel_event)
                    oc.write_conversion_report(opts["base_directory"], outcome)
                    if outcome.cancelled:
                        self._on_done(False, "Build cancelled during Office conversion.")
                        return
                    failures = outcome.failures
                    if failures:
                        names = "\n".join(f"  - {d.name}: {d.error}" for d in failures[:12])
                        if len(failures) > 12:
                            names += f"\n  ... and {len(failures) - 12} more"
                        if not self._ask_yesno_safe(
                                "Some documents were not converted",
                                f"{len(failures)} document(s) could not be converted and will be "
                                f"left out of the bundle:\n\n{names}\n\nContinue without them?"):
                            self._on_done(False, "Build stopped: unconverted documents.")
                            return
                    opts["overlay"] = outcome.overlay

                if opts.get("per_folder"):
                    results = be.build_per_folder_first_pages(opts["base_directory"], opts, log)
                    made = [r for r in results if r["output"]]
                    skipped = [(os.path.basename(r["folder"]), n, why)
                               for r in results for n, why in r["skipped"]]
                    lines = [f"  - {os.path.relpath(r['output'], opts['base_directory'])}: "
                             f"{r['info']['page_count']} page(s)" for r in made]
                    summary = (f"{len(made)} combined file(s) written under\n{opts['base_directory']}\n\n"
                               + "\n".join(lines[:25]) + ("\n  ..." if len(lines) > 25 else ""))
                    if skipped:
                        summary += "\n\nPARTIAL - skipped files:\n" + "\n".join(
                            f"  - {f}/{n}: {why}" for f, n, why in skipped[:20])
                    self._on_done(bool(made), summary if made else "No combined files were produced.",
                                  partial=bool(skipped), open_dir=opts["base_directory"])
                    return

                if opts.get("organise"):
                    log("Planning the bundle for the organiser...")
                    be.set_current_style(opts.get("style") or {})
                    be.set_current_terms(be.effective_term(opts), opts.get("numbering_style"))
                    opts["output_path"] = be.default_output_path(opts["base_directory"], opts["claim_title"])
                    plan = be.plan_bundle(opts["base_directory"], opts["policy"], opts,
                                          overlay=opts.get("overlay"), log_fn=log)
                    for w in plan.warnings:
                        log(f"  NOTE: {w}")
                    if not plan.sections:
                        self._on_done(False, "No PDFs or folders with PDFs found in the base directory.")
                        return
                    self.root.after(0, lambda: self._open_organiser(plan, opts))
                    return

                self._run_build(opts)
            except be.BundleCancelled:
                self._on_done(False, "Build cancelled.")
            except Exception:
                log(traceback.format_exc())
                self._on_done(False, "Unexpected error - see status panel.")

        def _run_build(self, opts):
            """Build (worker thread) and report totals."""
            log = self._log_safe
            try:
                output_path, plan, info = be.build_bundle_sync(opts, log)
            except be.BundleCancelled:
                self._on_done(False, "Build cancelled.")
                return
            except ValueError as e:
                self._on_done(False, str(e))
                return
            except Exception:
                log(traceback.format_exc())
                self._on_done(False, "Unexpected error - see status panel.")
                return
            summary = (f"{output_path}\n\n{info['page_count']} page(s): {info['sections']} section(s), "
                       f"{info['documents_with_pages']} document(s), {info['pages_included']} source "
                       f"page(s) included, {info['pages_excluded']} excluded.")
            skipped = info.get("skipped_documents") or []
            if skipped:
                summary += "\n\nPARTIAL OUTPUT - skipped files:\n" + "\n".join(
                    f"  - {n}: {why}" for n, why in skipped[:20])
                if len(skipped) > 20:
                    summary += f"\n  ... and {len(skipped) - 20} more (see status panel)"
            self._on_done(True, summary, partial=bool(skipped), open_dir=os.path.dirname(output_path))

        def _open_organiser(self, plan, opts):
            """UI thread: show the organiser; Create Bundle resumes the build."""
            self._set_status("Organiser open - arrange pages, then press Create Bundle.", append=True)

            def on_create(organised_plan):
                self._set_status("Creating bundle from the organised pages...", append=True)
                build_opts = dict(opts)
                build_opts["plan"] = organised_plan
                threading.Thread(target=self._run_build, args=(build_opts,), daemon=True).start()

            def on_preview(organised_plan):
                path = bo.preview_output_path()
                build_opts = dict(opts)
                build_opts["plan"] = be.snapshot_plan(organised_plan)
                build_opts["output_path"] = path
                self._set_status(f"Building preview: {path}", append=True)

                def run():
                    try:
                        out, _p, _i = be.build_bundle_sync(build_opts, self._log_safe)
                        self.root.after(0, lambda: self._open_file(out))
                    except Exception as e:
                        self._log_safe(f"Preview failed: {e}")

                threading.Thread(target=run, daemon=True).start()

            def on_create_wrapped(organised_plan):
                # Mark before the window closes so the poll below does not
                # treat the close as a cancellation.
                self._create_pending = True
                on_create(organised_plan)

            self._create_pending = False
            win = bo.OrganiserWindow(self.root, plan, opts, on_create_wrapped, on_preview)
            self._organiser = win

            def poll():
                if not win.winfo_exists():
                    if self._building and not self._create_pending:
                        self._set_building(False)
                        self._set_status("Organiser closed without creating a bundle.", append=True)
                    return
                self.root.after(300, poll)

            self.root.after(300, poll)

        def _open_file(self, path):
            try:
                os.startfile(path)
            except Exception:
                messagebox.showinfo("Preview", f"Preview saved to:\n{path}", parent=self.root)

        def _on_done(self, success, info, partial=False, open_dir=None):
            def _finish():
                self._set_building(False)
                self._create_pending = False
                if success:
                    title = "Build complete - PARTIAL (some files skipped)" if partial else "Build complete"
                    if messagebox.askyesno(title, f"{info}\n\nOpen the output folder?"):
                        try:
                            os.startfile(open_dir or os.path.dirname(info.splitlines()[0]))
                        except Exception:
                            pass
                else:
                    messagebox.showerror("Build failed", info)
            self.root.after(0, _finish)

    # Backwards-compatible alias
    ClaimsBundleApp = BundleBuilderApp

    def main():
        root = tk.Tk()
        BundleBuilderApp(root)
        root.mainloop()

    if __name__ == "__main__":
        main()

except SystemExit:
    raise
except Exception:
    err = traceback.format_exc()
    _fatal_error("Document Bundle Builder - Startup Error",
                 f"The application failed to start.\n\nDetails:\n\n{err}")
