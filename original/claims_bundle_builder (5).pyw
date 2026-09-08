"""
Claims Bundle Builder
=====================

Merges a folder tree of PDFs into one organised bundle with:
- Auto-numbered appendix separators (numeric or alphabetic)
- Hierarchical, clickable Table of Contents (multi-page safe)
- Optional OCR pass for scanned pages (English + Arabic)
- Per-page footers showing appendix label and page-of-pages
- Optional image compression (high / balanced / smallest)
- Optional cover page with project metadata
- Custom appendix order via _order.txt
- Persistent settings between runs

Designed for QS / claims work: variation packs, EOT submissions, IPC bundles.

Save this file as claims_bundle_builder.pyw and double-click to run without a
console window. First run auto-installs the required Python packages.
"""

import os
import re
import sys
import io
import json
import subprocess
import importlib
import threading
import traceback
from textwrap import wrap
from datetime import datetime


# ---------------------------------------------------------------------------
# Step 1: Last-resort error display (used if anything fails before main loop)
# ---------------------------------------------------------------------------

def _fatal_error(title, message):
    """Show a fatal error dialog and exit. Used for startup failures."""
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


# Wrap startup so any failure produces a popup
try:

    # -----------------------------------------------------------------------
    # Step 2: Bootstrap dependencies
    # -----------------------------------------------------------------------

    REQUIRED_PACKAGES = {
        # import_name : pip_name
        "fitz": "PyMuPDF",
        "fpdf": "fpdf2",
    }

    OPTIONAL_PACKAGES = {
        # import_name : pip_name (OCR support)
        "PIL": "Pillow",
        "pytesseract": "pytesseract",
    }

    def _missing(import_name):
        try:
            importlib.import_module(import_name)
            return False
        except ImportError:
            return True

    def _bootstrap_packages():
        """Install required packages on first run with a live-progress window.

        Shows a tkinter window with real-time pip output so the user sees what
        is happening and any error is visible (rather than the script silently
        dying when running as .pyw without a console).
        """
        missing_required = [
            (imp, pip) for imp, pip in REQUIRED_PACKAGES.items() if _missing(imp)
        ]

        if not missing_required:
            return

        import tkinter as tk
        from tkinter import scrolledtext

        # Build a live-progress window
        win = tk.Tk()
        win.title("Claims Bundle Builder - First Run Setup")
        win.geometry("700x420")

        # Force the window above other windows briefly so user sees it
        win.attributes("-topmost", True)
        win.after(2000, lambda: win.attributes("-topmost", False))

        pkg_list = ", ".join(p for _, p in missing_required)

        tk.Label(
            win,
            text="First-run setup",
            font=("Arial", 14, "bold"),
        ).pack(pady=(12, 4))

        tk.Label(
            win,
            text=(
                f"Installing required Python package(s): {pkg_list}\n"
                "This is a one-time install into your user folder. No admin rights needed.\n"
                "Live install output appears below. The window will close automatically when done."
            ),
            justify="center",
        ).pack(padx=12, pady=(0, 8))

        log_widget = scrolledtext.ScrolledText(
            win, height=15, font=("Consolas", 9), state="disabled"
        )
        log_widget.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        status_var = tk.StringVar(value="Starting...")
        status_lbl = tk.Label(win, textvariable=status_var, font=("Arial", 9, "bold"), fg="#1F3864")
        status_lbl.pack(pady=(0, 8))

        # Mutable state shared with the worker thread
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
            for imp_name, pip_name in missing_required:
                set_status_safe(f"Installing {pip_name}... (this can take a few minutes)")
                append_log_safe(f"\n>>> Installing {pip_name}\n")
                try:
                    cmd = [
                        sys.executable, "-m", "pip", "install",
                        "--user", "--disable-pip-version-check",
                        "--no-input", "--progress-bar", "off",
                        pip_name,
                    ]
                    # Stream stdout/stderr live to the log window. We use an
                    # idle-timeout (no output for N seconds) instead of a wall-clock
                    # timeout, because slow networks downloading 30+ MB packages
                    # can legitimately take 5-10 minutes.
                    proc = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        bufsize=1,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )

                    # Read line-by-line; do NOT impose a wall-clock timeout.
                    # pip prints download progress updates so we will know if
                    # it hangs (no new lines).
                    import time as _time
                    last_output = _time.time()
                    IDLE_TIMEOUT = 180  # 3 minutes of zero output = stuck

                    while True:
                        line = proc.stdout.readline()
                        if line == "" and proc.poll() is not None:
                            break
                        if line:
                            append_log_safe(line)
                            last_output = _time.time()
                        else:
                            # No new output - check idle time
                            if _time.time() - last_output > IDLE_TIMEOUT:
                                proc.kill()
                                state["failed"] = True
                                state["error_pkg"] = pip_name
                                state["error_msg"] = (
                                    f"pip produced no output for {IDLE_TIMEOUT} seconds. "
                                    f"This usually means the network connection stalled."
                                )
                                append_log_safe(
                                    f"\n*** STALLED: pip stopped producing output for {IDLE_TIMEOUT}s. Killing process.\n"
                                )
                                break
                            _time.sleep(0.1)

                    proc.stdout.close()
                    rc = proc.wait()

                    if state["failed"]:
                        break

                    if rc != 0:
                        state["failed"] = True
                        state["error_pkg"] = pip_name
                        state["error_msg"] = f"pip exited with code {rc}"
                        append_log_safe(f"\n*** FAILED: {pip_name} (exit code {rc})\n")
                        break
                    else:
                        append_log_safe(f"\n--- OK: {pip_name} installed\n")

                except Exception as e:
                    state["failed"] = True
                    state["error_pkg"] = pip_name
                    state["error_msg"] = str(e)
                    append_log_safe(f"\n*** EXCEPTION: {e}\n")
                    break

            # Verify imports
            if not state["failed"]:
                set_status_safe("Verifying...")
                importlib.invalidate_caches()
                still_missing = [imp for imp, _ in missing_required if _missing(imp)]
                if still_missing:
                    state["failed"] = True
                    state["error_pkg"] = ", ".join(still_missing)
                    state["error_msg"] = "Installed but cannot be imported."
                    append_log_safe(f"\n*** Import verification failed: {still_missing}\n")
                else:
                    append_log_safe("\nAll packages installed and verified.\n")
                    set_status_safe("Done. Closing in 2 seconds...")

            state["done"] = True
            if not state["failed"]:
                win.after(2000, win.destroy)
            else:
                set_status_safe("FAILED - see log above. Close this window to exit.")

        # Run install in a worker thread so the UI stays responsive
        import threading as _threading
        worker = _threading.Thread(target=install_worker, daemon=True)
        worker.start()

        win.mainloop()

        # Re-check after window closes
        if state["failed"]:
            _fatal_error(
                f"{state['error_pkg']} install failed",
                f"Could not install {state['error_pkg']}.\n\n"
                f"Reason: {state['error_msg']}\n\n"
                f"Possible causes:\n"
                f"- No internet access\n"
                f"- IT firewall blocking pypi.org\n"
                f"- Slow network (timeout after 10 minutes)\n\n"
                f"Try manually in CMD:\n"
                f"    python -m pip install --user {' '.join(p for _, p in missing_required)}"
            )

        # Final import sanity check
        importlib.invalidate_caches()
        still_missing = [imp for imp, _ in missing_required if _missing(imp)]
        if still_missing:
            _fatal_error(
                "Install verification failed",
                f"Packages installed but not importable: {still_missing}\n"
                "Restart the script. If it persists, install manually."
            )

    if sys.platform != "win32":
        # Allow Linux/Mac for testing the GUI but warn that OCR/Outlook integration is Windows-targeted
        pass

    _bootstrap_packages()

    # -----------------------------------------------------------------------
    # Step 3: Imports (now safe)
    # -----------------------------------------------------------------------

    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    import fitz  # PyMuPDF
    from fpdf import FPDF

    # OCR is optional — if Pillow / pytesseract aren't available, OCR is disabled
    try:
        from PIL import Image
        import pytesseract
        OCR_AVAILABLE = True
    except Exception:
        OCR_AVAILABLE = False

    # -----------------------------------------------------------------------
    # Step 4: Settings constants
    # -----------------------------------------------------------------------

    MM_TO_PT = 72 / 25.4
    LEFT_MARGIN_MM = 20
    RIGHT_MARGIN_MM = 20
    TOP_MARGIN_MM = 20
    BOTTOM_MARGIN_MM = 20

    FOOTER_OFFSET_PT = 18
    FOOTER_FONT_SIZE = 8
    FOOTER_GREY = (0, 0, 0)
    FOOTER_OPACITY = 0.8

    TOC_TITLE_FONT = 24
    TOC_SUBTITLE_FONT = 12
    TOC_TEXT_FONT = 12
    TOC_LINE_HEIGHT = 6

    SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".claims_bundle_builder.json")
    ORDER_FILE = "_order.txt"

    # -----------------------------------------------------------------------
    # Step 4b: Bundle style system (palette, fonts, page styles)
    # -----------------------------------------------------------------------

    PALETTES = {
        # ink = all text and structural rules; accent = decorative rules only,
        # never text (accents fade in black-and-white copies)
        "navy_gold": {"ink": (5, 22, 65), "accent": (255, 196, 37),
                      "muted": (90, 98, 114), "hair": (226, 226, 226),
                      "hair_strong": (201, 201, 201)},
        "navy_steel": {"ink": (5, 22, 65), "accent": (122, 134, 153),
                       "muted": (90, 98, 114), "hair": (226, 226, 226),
                       "hair_strong": (201, 201, 201)},
    }

    # Updated from the UI before each build
    CURRENT_STYLE = {
        "cover": "top_rule",        # top_rule | classical
        "palette": "navy_gold",     # navy_gold | navy_steel
        "font": "aptos",            # aptos | times
        "separator": "top_rule",    # top_rule | plaque
        "toc": "ruled",             # dotted | ruled
    }

    def style_palette():
        return PALETTES.get(CURRENT_STYLE.get("palette", "navy_gold"),
                            PALETTES["navy_gold"])

    _FONT_CACHE = {}

    def _find_font_file(names):
        """Locate a font file in the Windows system or per-user font folders."""
        dirs = [
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""),
                         "Microsoft", "Windows", "Fonts"),
        ]
        for d in dirs:
            if not d or not os.path.isdir(d):
                continue
            for n in names:
                path = os.path.join(d, n)
                if os.path.isfile(path):
                    return path
        return None

    def _resolve_font_paths(font_key):
        """Return {'regular': path, 'bold': path} or Nones for core fallback.

        aptos: Aptos (Microsoft 365 / Windows 11), falling back to Segoe UI.
        times: Times New Roman TTF for proper embedding, falling back to the
        FPDF core Times font.
        """
        if font_key in _FONT_CACHE:
            return _FONT_CACHE[font_key]
        if font_key == "times":
            paths = {"regular": _find_font_file(["times.ttf", "Times.ttf"]),
                     "bold": _find_font_file(["timesbd.ttf", "Timesbd.ttf"])}
        else:
            paths = {"regular": _find_font_file(
                         ["aptos.ttf", "Aptos.ttf", "aptos-regular.ttf",
                          "Aptos-Regular.ttf"]),
                     "bold": _find_font_file(
                         ["aptos-bold.ttf", "Aptos-Bold.ttf", "aptosb.ttf"])}
            if not paths["regular"]:
                paths = {"regular": _find_font_file(["segoeui.ttf"]),
                         "bold": _find_font_file(["segoeuib.ttf"])}
        if paths.get("regular") and not paths.get("bold"):
            paths["bold"] = paths["regular"]
        _FONT_CACHE[font_key] = paths
        return paths

    def new_styled_pdf():
        """FPDF instance with the selected style font registered.

        Returns (pdf, family). family is the registered TTF family or the
        core fallback (Helvetica for aptos, Times for times)."""
        pdf = FPDF(format="A4")
        font_key = CURRENT_STYLE.get("font", "aptos")
        core = "Times" if font_key == "times" else "Helvetica"
        family = core
        paths = _resolve_font_paths(font_key)
        if paths.get("regular"):
            try:
                try:
                    pdf.add_font("BundleFont", "", paths["regular"])
                    pdf.add_font("BundleFont", "B", paths["bold"])
                except TypeError:
                    pdf.add_font("BundleFont", "", paths["regular"], uni=True)
                    pdf.add_font("BundleFont", "B", paths["bold"], uni=True)
                family = "BundleFont"
            except Exception:
                family = core
        return pdf, family

    def _spaced(text):
        """Letter-space a kicker line (words become triple-spaced)."""
        return " ".join(text)


    # -----------------------------------------------------------------------
    # Step 5: Helpers (engine — preserved from original)
    # -----------------------------------------------------------------------

    def sanitize_text(text):
        """Strip characters FPDF cannot render."""
        if not text:
            return ""
        repl = {
            "\u2019": "'", "\u2018": "'", "\u201C": '"', "\u201D": '"',
            "\u2013": "-", "\u2014": "-", "\u2022": "-",
            "\u2026": "...", "\u00E9": "e", "\u00E8": "e",
            "\u00E1": "a", "\u00E0": "a", "\u00F6": "o", "\u00FC": "u",
            "\u00DF": "ss"
        }
        for o, n in repl.items():
            text = text.replace(o, n)
        return text

    def convert_to_style(number, style):
        """Convert appendix index to numeric (01) or alphabetic (A, B, ... AA)."""
        if style == "alphabetic":
            letters = ""
            n = number
            while n > 0:
                n, r = divmod(n - 1, 26)
                letters = chr(65 + r) + letters
            return letters
        return f"{number:02d}"

    def mm_rect(x0_mm, y0_mm, x1_mm, y1_mm):
        return fitz.Rect(x0_mm * MM_TO_PT, y0_mm * MM_TO_PT,
                         x1_mm * MM_TO_PT, y1_mm * MM_TO_PT)

    def get_script_dir():
        try:
            return os.path.dirname(os.path.abspath(__file__))
        except NameError:
            return os.getcwd()

    def _wrap_by_width(pdf, text, max_w):
        """Greedy wrap by measured width with hard-split for long words."""
        text = (text or "").strip()
        if not text:
            return [""]
        words = text.split()
        lines, cur = [], ""

        def fits(s):
            return pdf.get_string_width(s) <= max_w

        for w in words:
            cand = (cur + " " + w) if cur else w
            if fits(cand):
                cur = cand
                continue
            if cur:
                lines.append(cur)
                cur = w
            else:
                chunk = ""
                for ch in w:
                    if fits(chunk + ch):
                        chunk += ch
                    else:
                        if chunk:
                            lines.append(chunk)
                        chunk = ch
                cur = chunk
        if cur:
            lines.append(cur)
        return lines

    def _dot_fill(pdf, avail_w):
        dot_w = max(pdf.get_string_width("."), 0.1)
        return "." * max(int(avail_w / dot_w), 0)

    def create_separator(title_main, title_sub, main_size=24, sub_size=14, spacing=40):
        """Styled separator page in the selected style.

        Appendix-level calls pass title_main='Appendix X' with the description
        in title_sub; folder and per-file calls pass the name in title_main
        with an empty sub. main_size >= 24 marks appendix level."""
        pal = style_palette()
        title_main = sanitize_text(title_main)
        title_sub = sanitize_text(title_sub)

        if title_sub and title_main.lower().startswith("appendix"):
            kicker, headline, subline = title_main.upper(), title_sub, ""
        else:
            kicker, headline, subline = "", title_main, title_sub

        appendix_level = main_size >= 24
        pdf, fam = new_styled_pdf()
        pdf.set_auto_page_break(False)
        pdf.set_margins(LEFT_MARGIN_MM, TOP_MARGIN_MM, RIGHT_MARGIN_MM)
        pdf.add_page()
        effective_w = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM

        if CURRENT_STYLE.get("separator") == "plaque":
            head_size = 15 if appendix_level else 13
            frame_w, pad = 120.0, 10.0
            inner_w = frame_w - 2 * pad
            head_lh = head_size * 0.5

            # Measure content, then draw a frame that fits it exactly
            k_lines, s_lines = [], []
            h = pad
            if kicker:
                pdf.set_font(fam, "B", 10)
                k_lines = _wrap_by_width(pdf, _spaced(kicker), inner_w)
                h += len(k_lines) * 6 + 3
            pdf.set_font(fam, "B", head_size)
            h_lines = _wrap_by_width(pdf, headline, inner_w)
            h += len(h_lines) * head_lh
            if subline:
                pdf.set_font(fam, "", 10)
                s_lines = _wrap_by_width(pdf, subline, inner_w)
                h += 4 + len(s_lines) * 5.5
            h += pad

            x0 = (210 - frame_w) / 2
            y0 = max(TOP_MARGIN_MM, 132 - h / 2)
            pdf.set_draw_color(*pal["ink"])
            pdf.set_line_width(0.35)
            pdf.rect(x0, y0, frame_w, h)

            pdf.set_y(y0 + pad)
            if k_lines:
                pdf.set_font(fam, "B", 10)
                pdf.set_text_color(*pal["ink"])
                for line in k_lines:
                    pdf.set_x(x0 + pad)
                    pdf.multi_cell(inner_w, 6, line, 0, "C")
                pdf.ln(3)
            pdf.set_font(fam, "B", head_size)
            pdf.set_text_color(*pal["ink"])
            for line in h_lines:
                pdf.set_x(x0 + pad)
                pdf.multi_cell(inner_w, head_lh, line, 0, "C")
            if s_lines:
                pdf.ln(4)
                pdf.set_font(fam, "", 10)
                pdf.set_text_color(*pal["muted"])
                for line in s_lines:
                    pdf.set_x(x0 + pad)
                    pdf.multi_cell(inner_w, 5.5, line, 0, "C")
        else:
            # Top rule style
            head_size = 18 if appendix_level else 15
            head_lh = head_size * 0.5
            pdf.set_fill_color(*pal["ink"])
            pdf.rect(0, 0, 210, 3, "F")
            pdf.set_y(95)
            if kicker:
                pdf.set_font(fam, "B", 11)
                pdf.set_text_color(*pal["ink"])
                pdf.set_x(LEFT_MARGIN_MM)
                pdf.multi_cell(effective_w, 7, _spaced(kicker), 0, "L")
                pdf.ln(3)
            pdf.set_font(fam, "B", head_size)
            pdf.set_text_color(*pal["ink"])
            for line in _wrap_by_width(pdf, headline, effective_w):
                pdf.set_x(LEFT_MARGIN_MM)
                pdf.multi_cell(effective_w, head_lh, line, 0, "L")
            pdf.ln(3)
            pdf.set_fill_color(*pal["accent"])
            pdf.rect(LEFT_MARGIN_MM, pdf.get_y(), 22, 1.2, "F")
            if subline:
                pdf.ln(7)
                pdf.set_font(fam, "", 11)
                pdf.set_text_color(*pal["muted"])
                for line in _wrap_by_width(pdf, subline, effective_w):
                    pdf.set_x(LEFT_MARGIN_MM)
                    pdf.multi_cell(effective_w, 6.5, line, 0, "L")
        return bytes(pdf.output())

    def _cover_top_rule(pdf, fam, pal, doc_type, title, subtitle, project, ref, rows):
        X, W = 25.0, 160.0

        def block(text, size, style, colour, lh):
            pdf.set_font(fam, style, size)
            pdf.set_text_color(*colour)
            for line in _wrap_by_width(pdf, text, W):
                pdf.set_x(X)
                pdf.multi_cell(W, lh, line, 0, "L")

        pdf.set_fill_color(*pal["ink"])
        pdf.rect(0, 0, 210, 4, "F")

        pdf.set_y(52)
        block(_spaced(doc_type.upper()), 10, "B", pal["ink"], 6)
        pdf.ln(8)

        pdf.set_font(fam, "B", 30)
        if len(_wrap_by_width(pdf, title, W)) > 2:
            size, lh = 24, 11
        else:
            size, lh = 30, 13
        block(title, size, "B", pal["ink"], lh)

        pdf.ln(4)
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(X, pdf.get_y(), 22, 1.5, "F")
        pdf.ln(10)

        if project:
            block(project, 14, "B", pal["ink"], 7.5)
            pdf.ln(1.5)
        if ref:
            block(ref, 11, "", pal["muted"], 6)
        if subtitle:
            pdf.ln(4)
            block(subtitle, 12, "", pal["ink"], 6.5)

        # Details ledger anchored to the lower third
        label_w = 42.0
        y = 232.0
        pdf.set_draw_color(*pal["hair_strong"])
        pdf.set_line_width(0.3)
        pdf.line(X, y - 4, X + W, y - 4)
        pdf.set_draw_color(*pal["hair"])
        pdf.set_line_width(0.2)
        for lbl, val in rows:
            pdf.set_xy(X, y)
            pdf.set_font(fam, "B", 8.5)
            pdf.set_text_color(*pal["ink"])
            pdf.cell(label_w, 6, lbl, 0, 0, "L")
            pdf.set_font(fam, "", 11)
            v_lines = _wrap_by_width(pdf, val, W - label_w)
            for k, vl in enumerate(v_lines):
                pdf.set_xy(X + label_w, y + k * 5.5)
                pdf.multi_cell(W - label_w, 5.5, vl, 0, "L")
            y += max(1, len(v_lines)) * 5.5 + 2
            pdf.line(X, y - 1, X + W, y - 1)
            y += 2

    def _cover_classical(pdf, fam, pal, doc_type, title, subtitle, project, ref, rows):
        X, W = 25.0, 160.0

        def cblock(text, size, style, colour, lh):
            pdf.set_font(fam, style, size)
            pdf.set_text_color(*colour)
            for line in _wrap_by_width(pdf, text, W):
                pdf.set_x(X)
                pdf.multi_cell(W, lh, line, 0, "C")

        pdf.set_y(78)
        cblock(_spaced(doc_type.upper()), 9, "B", pal["ink"], 6)
        pdf.ln(8)

        pdf.set_font(fam, "B", 26)
        if len(_wrap_by_width(pdf, title, W)) > 2:
            size, lh = 22, 10
        else:
            size, lh = 26, 12
        cblock(title, size, "B", pal["ink"], lh)

        pdf.ln(4)
        pdf.set_fill_color(*pal["ink"])
        pdf.rect(105 - 8, pdf.get_y(), 16, 0.8, "F")
        pdf.ln(8)

        if project:
            cblock(project, 12, "", pal["ink"], 6.5)
        if ref:
            pdf.ln(1)
            cblock(ref, 10, "", pal["muted"], 5.5)
        if subtitle:
            pdf.ln(3)
            cblock(subtitle, 11, "", pal["ink"], 6)

        # Centred details block
        y = 228.0
        for lbl, val in rows:
            pdf.set_xy(X, y)
            pdf.set_font(fam, "B", 8)
            pdf.set_text_color(*pal["ink"])
            pdf.multi_cell(W, 4.5, lbl, 0, "C")
            pdf.set_font(fam, "", 11)
            for vl in _wrap_by_width(pdf, val, W):
                pdf.set_x(X)
                pdf.multi_cell(W, 5.5, vl, 0, "C")
            y = pdf.get_y() + 2.5

    def create_cover_page(meta):
        """Cover page in the selected style (top rule or classical centred).

        Light-field, low-ink composition. Accent colours are used for rules
        only; all text is set in the palette ink so black-and-white copies
        lose nothing."""
        pal = style_palette()
        pdf, fam = new_styled_pdf()
        pdf.set_auto_page_break(False)
        pdf.set_margins(0, 0, 0)
        pdf.add_page()

        title = sanitize_text(meta.get("title", "Claims Bundle"))
        subtitle = sanitize_text(meta.get("subtitle", ""))
        if subtitle.strip().lower() == title.strip().lower():
            subtitle = ""  # never repeat the title as its own subtitle
        doc_type = sanitize_text(meta.get("doc_type", "")).strip() or "Claims Bundle"
        project = sanitize_text(meta.get("project", ""))
        contract_ref = sanitize_text(meta.get("contract_ref", ""))
        date_val = sanitize_text(meta.get("date", "")).strip()
        if not date_val:
            date_val = datetime.now().strftime("%d %B %Y")
        rows = [
            ("PREPARED BY", sanitize_text(meta.get("prepared_by", ""))),
            ("PREPARED FOR", sanitize_text(meta.get("prepared_for", ""))),
            ("DATE", date_val),
            ("REVISION", sanitize_text(meta.get("revision", ""))),
        ]
        rows = [(lbl, val) for lbl, val in rows if val.strip()]

        if CURRENT_STYLE.get("cover") == "classical":
            _cover_classical(pdf, fam, pal, doc_type, title, subtitle,
                             project, contract_ref, rows)
        else:
            _cover_top_rule(pdf, fam, pal, doc_type, title, subtitle,
                            project, contract_ref, rows)
        return bytes(pdf.output())

    def _has_immediate_subfolders(folder_path):
        try:
            return any(
                os.path.isdir(os.path.join(folder_path, name))
                for name in os.listdir(folder_path)
            )
        except Exception:
            return False

    def _tree_has_pdf(folder_path):
        """True if this folder or any subfolder contains at least one PDF.
        Used to filter out empty folders and folders whose entire subtree
        has no PDFs - those should not become appendices or separators.
        """
        try:
            for _, _, files in os.walk(folder_path):
                if any(f.lower().endswith(".pdf") for f in files):
                    return True
        except Exception:
            pass
        return False

    def create_toc_page_hier(claim_title, toc_entries, max_level=5):
        """Generate the TOC PDF supporting up to `max_level` indent levels.

        toc_entries: flat list of dicts in render order:
            {
              "level": 1..max_level,
              "title": str,
              "page": 1-based page number target (or None for non-clickable headings),
              "is_heading": bool (True if this is a folder-only entry),
              "show_page": bool (whether to render the page number on this row)
            }

        Returns (pdf_bytes, link_positions, page_count).

        link_positions is a list of dicts:
          {'page_in_toc', 'y0_mm', 'y1_mm', 'target_zero'}
        """
        pal = style_palette()
        toc_ruled = CURRENT_STYLE.get("toc") == "ruled"
        pdf, fam = new_styled_pdf()
        pdf.set_auto_page_break(auto=True, margin=BOTTOM_MARGIN_MM)
        pdf.set_margins(LEFT_MARGIN_MM, TOP_MARGIN_MM, RIGHT_MARGIN_MM)
        pdf.add_page()

        # Header
        pdf.set_font(fam, "B", 22)
        pdf.set_text_color(*pal["ink"])
        pdf.set_x(LEFT_MARGIN_MM)
        pdf.cell(0, 12, "Contents", 0, 1, "L")
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(LEFT_MARGIN_MM, pdf.get_y() + 1, 22, 1.5, "F")
        pdf.ln(6)
        if claim_title:
            pdf.set_font(fam, "", 11)
            pdf.set_text_color(*pal["muted"])
            pdf.set_x(LEFT_MARGIN_MM)
            pdf.cell(0, 7, sanitize_text(claim_title), 0, 1, "L")
        pdf.ln(4)

        total_w = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM
        page_box_w = 18
        line_h = TOC_LINE_HEIGHT
        per_level_indent = 6  # mm of indent per level beyond level 1

        link_pos = []

        # Per-level visual style
        def style_for_level(lvl):
            # level 1 = appendix (bold, normal size)
            # level 2 = first subfolder heading (bold, slightly smaller)
            # level 3-4 = nested headings (regular, smaller)
            # level 5+ = leaves (regular, smallest)
            if lvl == 1:
                return ("B", TOC_TEXT_FONT)
            if lvl == 2:
                return ("B", TOC_TEXT_FONT - 1)
            if lvl == 3:
                return ("", TOC_TEXT_FONT - 1)
            return ("", TOC_TEXT_FONT - 1)

        for entry in toc_entries:
            lvl = max(1, min(entry.get("level", 1), max_level))
            title = sanitize_text(entry["title"])
            page_no = entry.get("page")
            show_page = entry.get("show_page", True) and (page_no is not None)

            indent_mm = (lvl - 1) * per_level_indent
            available_w = total_w - indent_mm

            font_style, font_size = style_for_level(lvl)
            pdf.set_font(fam, font_style, font_size)
            pdf.set_text_color(*(pal["ink"] if lvl <= 2 else pal["muted"]))

            # Ruled style: level-1 rows carry the appendix letter in a chip
            chip_label = None
            chip_w = 0.0
            if toc_ruled and lvl == 1:
                m = re.match(r"Appendix\s+([A-Za-z0-9]{1,3}):\s*(.*)", title)
                if m:
                    chip_label, title = m.group(1), m.group(2)
                    chip_w = 9.0

            max_text_w = available_w - chip_w - (page_box_w + 2 if show_page else 0)
            lines = _wrap_by_width(pdf, title, max_text_w)

            page_in_toc_top = pdf.page_no() - 1
            y_top = pdf.get_y()

            if chip_label:
                chip_h = line_h
                pdf.set_fill_color(*pal["ink"])
                pdf.rect(LEFT_MARGIN_MM + indent_mm, y_top + 0.4, 7, chip_h - 0.8, "F")
                pdf.set_font(fam, "B", font_size - 2)
                pdf.set_text_color(255, 255, 255)
                pdf.set_xy(LEFT_MARGIN_MM + indent_mm, y_top)
                pdf.cell(7, chip_h, chip_label, 0, 0, "C")
                pdf.set_font(fam, font_style, font_size)
                pdf.set_text_color(*pal["ink"])

            for i, line in enumerate(lines):
                last = (i == len(lines) - 1)
                pdf.set_x(LEFT_MARGIN_MM + indent_mm + chip_w)
                if last:
                    text_w = pdf.get_string_width(line)
                    if show_page:
                        fill_w = max(available_w - chip_w - text_w - page_box_w - 2, 0)
                        pdf.cell(text_w + 1, line_h, line)
                        if toc_ruled:
                            pdf.cell(fill_w, line_h, "")
                        else:
                            pdf.set_text_color(*pal["hair_strong"])
                            pdf.cell(fill_w, line_h, _dot_fill(pdf, fill_w))
                            pdf.set_text_color(*(pal["ink"] if lvl <= 2 else pal["muted"]))
                        pdf.cell(page_box_w, line_h, str(page_no), ln=1, align="R")
                    else:
                        pdf.cell(0, line_h, line, ln=1)
                else:
                    pdf.cell(0, line_h, line, ln=1)

            # Ruled style: hairline under each entry (strong for appendices)
            if toc_ruled:
                rule_y = pdf.get_y() + 0.4
                if rule_y < 297 - BOTTOM_MARGIN_MM:
                    if lvl == 1:
                        pdf.set_draw_color(*pal["ink"])
                        pdf.set_line_width(0.35)
                    else:
                        pdf.set_draw_color(*pal["hair"])
                        pdf.set_line_width(0.2)
                    pdf.line(LEFT_MARGIN_MM + indent_mm, rule_y,
                             210 - RIGHT_MARGIN_MM, rule_y)
                pdf.ln(1.5)

            page_in_toc_bot = pdf.page_no() - 1
            y_bottom = pdf.get_y()

            # Link target: only register a link if this entry has a page target
            if page_no is not None:
                target_zero = page_no - 1
                if page_in_toc_bot == page_in_toc_top:
                    link_pos.append({
                        "page_in_toc": page_in_toc_top,
                        "y0_mm": y_top,
                        "y1_mm": y_bottom,
                        "target_zero": target_zero,
                    })
                else:
                    # Split across page boundary into two rectangles
                    link_pos.append({
                        "page_in_toc": page_in_toc_top,
                        "y0_mm": y_top,
                        "y1_mm": 297 - BOTTOM_MARGIN_MM,
                        "target_zero": target_zero,
                    })
                    link_pos.append({
                        "page_in_toc": page_in_toc_bot,
                        "y0_mm": TOP_MARGIN_MM,
                        "y1_mm": y_bottom,
                        "target_zero": target_zero,
                    })

            # Small gap after appendix-level entries (not for nested ones)
            if lvl == 1:
                pdf.ln(1)

        toc_page_count = pdf.page_no()
        return bytes(pdf.output()), link_pos, toc_page_count

    def insert_pdf_clean(merger, src_path):
        """Three-tier insert with fallbacks. Returns pages inserted."""
        # Tier 1: Direct insert
        try:
            with fitz.open(src_path) as d:
                merger.insert_pdf(d)
                return d.page_count
        except Exception:
            pass

        # Tier 2: Mild rebuild
        try:
            with fitz.open(src_path) as d:
                data = d.tobytes(garbage=1)
            with fitz.open(stream=data, filetype="pdf") as fixed:
                merger.insert_pdf(fixed)
                return fixed.page_count
        except Exception:
            pass

        # Tier 3: Page-by-page
        pages = 0
        try:
            with fitz.open(src_path) as d:
                for i in range(d.page_count):
                    try:
                        merger.insert_pdf(d, from_page=i, to_page=i)
                        pages += 1
                    except Exception:
                        pass
        except Exception:
            return 0
        return pages

    def draw_footer_numbering(doc, page_idx, left_text, right_text):
        page = doc.load_page(page_idx)
        rect = page.rect
        y = rect.y1 - FOOTER_OFFSET_PT
        page.insert_text(
            fitz.Point(rect.x0 + 20, y),
            left_text,
            fontsize=FOOTER_FONT_SIZE,
            color=FOOTER_GREY,
            fill_opacity=FOOTER_OPACITY,
            overlay=True,
        )
        box = fitz.Rect(rect.x0 + 20, y - 10, rect.x1 - 20, y + 10)
        page.insert_textbox(
            box,
            right_text,
            fontsize=FOOTER_FONT_SIZE,
            color=FOOTER_GREY,
            fill_opacity=FOOTER_OPACITY,
            align=fitz.TEXT_ALIGN_RIGHT,
            overlay=True,
        )

    def ocr_page_if_needed(page, dpi=200, lang_hint="eng+ara"):
        try:
            if page.get_text("text").strip():
                return False
            if not OCR_AVAILABLE:
                return False
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                txt = pytesseract.image_to_string(img, lang=lang_hint)
            except Exception:
                txt = pytesseract.image_to_string(img)
            txt = txt.strip()
            if not txt:
                return False
            rect = page.rect
            box = fitz.Rect(rect.x0 + 36, rect.y0 + 36, rect.x1 - 36, rect.y1 - 36)
            page.insert_textbox(
                box,
                txt,
                fontsize=6,
                color=(1, 1, 1),
                fill_opacity=0.01,
                overlay=False,
            )
            return True
        except Exception:
            return False

    def _page_has_content(page):
        try:
            if page.get_text("text").strip():
                return True
            if page.get_images(full=True):
                return True
            if page.get_drawings():
                return True
            try:
                annots = list(page.annots() or [])
                if annots:
                    return True
            except Exception:
                pass
        except Exception:
            pass
        return False

    def _read_order_file(folder_path):
        """Read _order.txt if present. Returns list of names in desired order."""
        order_path = os.path.join(folder_path, ORDER_FILE)
        if not os.path.isfile(order_path):
            return None
        try:
            with open(order_path, "r", encoding="utf-8") as f:
                lines = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
            return lines
        except Exception:
            return None

    def _apply_custom_order(items, custom_order):
        """Reorder items: those listed in custom_order first (in that order),
        then anything not listed in alphabetical order at the end."""
        if not custom_order:
            return sorted(items)
        seen = set()
        ordered = []
        for name in custom_order:
            if name in items and name not in seen:
                ordered.append(name)
                seen.add(name)
        leftovers = sorted([x for x in items if x not in seen])
        return ordered + leftovers

    # ------------------------------------------------------------------
    # FOLDER POLICY MODEL
    # ------------------------------------------------------------------
    # The policy object controls how each folder/subfolder is rendered
    # in the bundle. Two top-level modes:
    #
    #   "automatic" - one rule for every folder
    #     options:
    #       layout: "flatten" | "preserve" | "preserve_skip_numeric"
    #       separators: "every_pdf" | "folder_only" | "none"
    #
    #   "manual" - per-folder dialog
    #     decisions: dict mapping folder relative-path -> "heading"|"flatten"|"skip"
    #     separators: same options as automatic
    #
    # The policy object is built once before the bundle build starts.
    # ------------------------------------------------------------------

    def _count_folder_tree(base_directory):
        """Walk the base directory and count folders, subfolders, PDFs.
        Returns dict: {top_pdfs, top_folders, total_subfolders, total_pdfs}.
        Subfolders/folders with no PDFs in their subtree are excluded.
        """
        counts = {
            "top_pdfs": 0,
            "top_folders": 0,
            "total_subfolders": 0,
            "total_pdfs": 0,
        }
        try:
            base_items = os.listdir(base_directory)
        except Exception:
            return counts

        for name in base_items:
            full = os.path.join(base_directory, name)
            if os.path.isfile(full) and name.lower().endswith(".pdf"):
                counts["top_pdfs"] += 1
                counts["total_pdfs"] += 1
            elif os.path.isdir(full) and _tree_has_pdf(full):
                counts["top_folders"] += 1
                # Walk subtree
                for cur, subdirs, files in os.walk(full):
                    counts["total_pdfs"] += sum(
                        1 for f in files if f.lower().endswith(".pdf")
                    )
                    for sd in subdirs:
                        if _tree_has_pdf(os.path.join(cur, sd)):
                            counts["total_subfolders"] += 1
        return counts

    def _collect_folders_for_manual(base_directory):
        """Return ordered list of folder rel-paths needing a manual decision.
        Top-level folders are skipped (they are always appendices). Only
        nested subfolders inside an appendix are returned, with full base-
        relative paths like "Backup / Contractor / Plant".
        Order = depth-first, parent before children.
        """
        result = []

        def walk(path, rel, depth):
            try:
                items = sorted(os.listdir(path))
            except Exception:
                return
            for name in items:
                full = os.path.join(path, name)
                if os.path.isdir(full) and _tree_has_pdf(full):
                    sub_rel = os.path.join(rel, name) if rel else name
                    # Skip depth 0 (top-level folders are appendices, not subfolders)
                    if depth >= 1:
                        result.append(sub_rel.replace(os.sep, " / "))
                    walk(full, sub_rel, depth + 1)

        walk(base_directory, "", 0)
        return result

    def show_discovery_popup(parent, counts):
        """Stage 1: summary popup. Returns 'automatic' | 'manual' | None."""
        dlg = tk.Toplevel(parent)
        dlg.title("Folder structure detected")
        dlg.geometry("520x320")
        dlg.transient(parent)
        dlg.grab_set()

        choice = {"value": None}

        tk.Label(dlg, text="Folder structure detected",
                 font=("Arial", 12, "bold")).pack(pady=(12, 6))

        summary = (
            f"Your bundle contains:\n\n"
            f"   - {counts['top_pdfs']} top-level PDFs (become Appendix A, B, ...)\n"
            f"   - {counts['top_folders']} top-level folders (become subsequent appendices)\n"
            f"   - {counts['total_subfolders']} subfolders nested inside\n"
            f"   - {counts['total_pdfs']} total PDF files\n\n"
            f"How should folders be handled?"
        )
        tk.Label(dlg, text=summary, justify="left", font=("Arial", 10)).pack(
            padx=20, pady=8, anchor="w"
        )

        legend = (
            "Automatic = one rule applied to all folders (faster)\n"
            "Manual    = decide for each folder individually"
        )
        tk.Label(dlg, text=legend, fg="#555", font=("Arial", 9),
                 justify="left").pack(padx=20, anchor="w")

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=18)

        def pick(v):
            choice["value"] = v
            dlg.destroy()

        tk.Button(btn_frame, text="Automatic mode", width=18,
                  bg="#1F3864", fg="white", font=("Arial", 10, "bold"),
                  command=lambda: pick("automatic")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Manual mode", width=18,
                  font=("Arial", 10, "bold"),
                  command=lambda: pick("manual")).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel", width=10,
                  command=lambda: pick(None)).pack(side="left", padx=6)

        dlg.wait_window()
        return choice["value"]

    def show_automatic_options_dialog(parent):
        """Stage 2A: layout + separator options for automatic mode.
        Returns dict {layout, separators} or None if cancelled."""
        dlg = tk.Toplevel(parent)
        dlg.title("Automatic mode - choose layout")
        dlg.geometry("640x560")
        dlg.transient(parent)
        dlg.grab_set()

        result = {"value": None}

        tk.Label(dlg, text="Automatic mode - choose layout",
                 font=("Arial", 12, "bold")).pack(pady=(10, 6))

        # Layout options
        layout_frame = tk.LabelFrame(dlg, text="How should folder hierarchy appear?",
                                     padx=10, pady=8)
        layout_frame.pack(fill="x", padx=12, pady=6)

        var_layout = tk.StringVar(value="preserve_skip_numeric")
        layouts = [
            ("flatten", "FLATTEN",
             "All PDFs from all subfolders combined into the top-level appendix.\n"
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
            tk.Label(f, text=desc, fg="#555", font=("Arial", 9),
                     justify="left").pack(anchor="w", padx=24)

        # Separator options
        sep_frame = tk.LabelFrame(dlg, text="Separator pages", padx=10, pady=8)
        sep_frame.pack(fill="x", padx=12, pady=6)

        var_sep = tk.StringVar(value="folder_only")
        seps = [
            ("every_pdf", "Before every PDF (adds many pages)"),
            ("folder_only", "Before each folder/subfolder only (recommended)"),
            ("none", "None (rely on TOC + bookmarks)"),
        ]
        for value, label in seps:
            tk.Radiobutton(sep_frame, text=label, variable=var_sep,
                           value=value, font=("Arial", 9)).pack(anchor="w")

        # Buttons
        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=12)

        def ok():
            result["value"] = {
                "mode": "automatic",
                "layout": var_layout.get(),
                "separators": var_sep.get(),
            }
            dlg.destroy()

        def back():
            result["value"] = "BACK"
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Back", width=10, command=back).pack(side="left", padx=6)

        dlg.wait_window()
        return result["value"]

    def show_manual_separators_dialog(parent):
        """Stage 2B-pre: ask separator preference once for manual mode."""
        dlg = tk.Toplevel(parent)
        dlg.title("Manual mode - separator pages")
        dlg.geometry("520x260")
        dlg.transient(parent)
        dlg.grab_set()

        result = {"value": None}

        tk.Label(dlg, text="Separator pages",
                 font=("Arial", 12, "bold")).pack(pady=(12, 6))
        tk.Label(dlg,
                 text="Before asking you to decide each folder, choose one\n"
                      "separator-page rule that applies across the whole bundle.",
                 justify="center").pack(padx=12, pady=4)

        var_sep = tk.StringVar(value="folder_only")
        for value, label in [
            ("every_pdf", "Before every PDF (adds many pages)"),
            ("folder_only", "Before each folder/subfolder only (recommended)"),
            ("none", "None (rely on TOC + bookmarks)"),
        ]:
            tk.Radiobutton(dlg, text=label, variable=var_sep, value=value,
                           font=("Arial", 9)).pack(anchor="w", padx=30, pady=2)

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=12)

        def ok():
            result["value"] = var_sep.get()
            dlg.destroy()

        def back():
            result["value"] = "BACK"
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Back", width=10, command=back).pack(side="left", padx=6)

        dlg.wait_window()
        return result["value"]

    def show_manual_folder_dialog(parent, folder_rel, idx, total,
                                  pdf_count, sub_count, max_subtree_depth):
        """Stage 2B per-folder: returns (decision_dict, apply_to_all_remaining).

        decision_dict has shape:
          {"action": "show", "depth": int}     - show files, depth-limited
          {"action": "flatten"}                - absorb into parent
          {"action": "skip"}                   - exclude entirely

        depth = 1 means collapse to this folder only (one TOC line);
        depth = 2..N means show up to N levels below this folder;
        depth = 999 means unlimited / show everything.

        max_subtree_depth = max levels below this folder that actually exist;
        the dropdown caps at this value to keep choices meaningful.
        """
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
        tk.Label(path_frame, text=folder_rel, font=("Consolas", 10),
                 wraplength=620, justify="left").pack(anchor="w")

        depth_str = (
            f"{pdf_count} PDF(s) directly in this folder, "
            f"{sub_count} immediate subfolder(s)"
        )
        if max_subtree_depth > 0:
            depth_str += f", subtree up to {max_subtree_depth} level(s) deep"
        tk.Label(dlg, text="Contains: " + depth_str,
                 font=("Arial", 9)).pack(pady=(4, 6))

        opt_frame = tk.LabelFrame(dlg, text="How should this folder be handled?",
                                  padx=10, pady=8)
        opt_frame.pack(fill="x", padx=12, pady=4)

        var_action = tk.StringVar(value="show")

        # Option 1: SHOW with depth dropdown
        show_row = tk.Frame(opt_frame)
        show_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(show_row, text="SHOW FILES",
                       variable=var_action, value="show",
                       font=("Arial", 9, "bold")).pack(side="left", anchor="w")
        tk.Label(show_row, text="depth limit:",
                 font=("Arial", 9)).pack(side="left", padx=(20, 4))

        # Build depth options: 1 = collapse, 2..max+1 = limit, 999 = unlimited
        depth_options = []
        depth_labels = {}
        depth_options.append(1)
        depth_labels[1] = "1 (collapse - just this folder)"
        actual_max = max(max_subtree_depth, 1)
        for n in range(2, actual_max + 2):
            depth_options.append(n)
            if n == 2:
                depth_labels[n] = "2 (one level below)"
            else:
                depth_labels[n] = f"{n} ({n - 1} levels below)"
        depth_options.append(999)
        depth_labels[999] = "unlimited (everything)"

        # Default depth: unlimited (matches old "show everything" behaviour)
        var_depth = tk.IntVar(value=999)
        depth_menu = ttk.Combobox(
            show_row,
            values=[depth_labels[d] for d in depth_options],
            state="readonly",
            width=30,
            font=("Arial", 9),
        )
        depth_menu.set(depth_labels[999])
        depth_menu.pack(side="left")

        def on_depth_change(_event):
            label = depth_menu.get()
            for d, lbl in depth_labels.items():
                if lbl == label:
                    var_depth.set(d)
                    return
        depth_menu.bind("<<ComboboxSelected>>", on_depth_change)

        tk.Label(opt_frame,
                 text="    1 = collapse (folder name only, files hidden in TOC).\n"
                      "    2-N = show files/subfolders up to N levels deep.\n"
                      "    unlimited = show every file at every depth.",
                 fg="#555", font=("Arial", 9), justify="left").pack(
            anchor="w", padx=24
        )

        # Option 2: FLATTEN
        flatten_row = tk.Frame(opt_frame)
        flatten_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(flatten_row, text="FLATTEN INTO PARENT",
                       variable=var_action, value="flatten",
                       font=("Arial", 9, "bold")).pack(anchor="w")
        tk.Label(opt_frame,
                 text="    No heading for this folder. PDFs absorbed into parent's level.\n"
                      "    Subfolders inside still ask their own questions.",
                 fg="#555", font=("Arial", 9), justify="left").pack(
            anchor="w", padx=24
        )

        # Option 3: SKIP
        skip_row = tk.Frame(opt_frame)
        skip_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(skip_row, text="SKIP",
                       variable=var_action, value="skip",
                       font=("Arial", 9, "bold")).pack(anchor="w")
        tk.Label(opt_frame,
                 text="    Exclude this folder and its contents entirely.",
                 fg="#555", font=("Arial", 9), justify="left").pack(
            anchor="w", padx=24
        )

        var_all = tk.BooleanVar(value=False)
        tk.Checkbutton(
            dlg,
            text="Apply this choice to all remaining folders in this appendix",
            variable=var_all, font=("Arial", 9)
        ).pack(pady=(8, 4))

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=8)

        def ok():
            action = var_action.get()
            if action == "show":
                result["decision"] = {"action": "show", "depth": var_depth.get()}
            else:
                result["decision"] = {"action": action}
            result["all"] = var_all.get()
            dlg.destroy()

        def cancel():
            result["decision"] = {"action": "cancel"}
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel build", width=14,
                  command=cancel).pack(side="left", padx=6)

        dlg.wait_window()
        return result["decision"], result["all"]

    def show_appendix_level_dialog(parent, appendix_name, idx, total,
                                   pdf_count, sub_count, max_subtree_depth):
        """Per-appendix dialog: returns ({"action": str, "depth": int} | None, cancelled).

        Actions:
          "show"             - show all files in appendix, depth-limited
          "go_folder"        - per-subfolder dialogs (existing manual mode)
        """
        dlg = tk.Toplevel(parent)
        dlg.title(f"Appendix decision - {idx} of {total}")
        dlg.geometry("680x540")
        dlg.transient(parent)
        dlg.grab_set()

        result = {"decision": None}

        tk.Label(dlg, text=f"Appendix decision - {idx} of {total}",
                 font=("Arial", 12, "bold")).pack(pady=(10, 4))

        path_frame = tk.LabelFrame(dlg, text="Appendix", padx=8, pady=6)
        path_frame.pack(fill="x", padx=12, pady=4)
        tk.Label(path_frame, text=appendix_name, font=("Consolas", 10),
                 wraplength=620).pack(anchor="w")

        info = (
            f"Contains: {pdf_count} PDF(s) directly, "
            f"{sub_count} immediate subfolder(s)"
        )
        if max_subtree_depth > 0:
            info += f", subtree up to {max_subtree_depth} level(s) deep"
        tk.Label(dlg, text=info, font=("Arial", 9)).pack(pady=(4, 6))

        opt_frame = tk.LabelFrame(dlg, text="Default for this appendix",
                                  padx=10, pady=8)
        opt_frame.pack(fill="x", padx=12, pady=4)

        var_action = tk.StringVar(value="show")

        # SHOW with depth dropdown
        show_row = tk.Frame(opt_frame)
        show_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(show_row, text="SHOW ALL CONTENTS",
                       variable=var_action, value="show",
                       font=("Arial", 9, "bold")).pack(side="left", anchor="w")
        tk.Label(show_row, text="depth limit:",
                 font=("Arial", 9)).pack(side="left", padx=(20, 4))

        depth_options = [1]
        depth_labels = {1: "1 (appendix line only - collapse all)"}
        actual_max = max(max_subtree_depth, 1)
        for n in range(2, actual_max + 2):
            if n == 2:
                depth_labels[n] = "2 (top level only)"
            else:
                depth_labels[n] = f"{n} ({n - 1} levels below appendix)"
            depth_options.append(n)
        depth_options.append(999)
        depth_labels[999] = "unlimited (every file at every depth)"

        var_depth = tk.IntVar(value=999)
        depth_menu = ttk.Combobox(
            show_row,
            values=[depth_labels[d] for d in depth_options],
            state="readonly",
            width=32,
            font=("Arial", 9),
        )
        depth_menu.set(depth_labels[999])
        depth_menu.pack(side="left")

        def on_depth_change(_event):
            label = depth_menu.get()
            for d, lbl in depth_labels.items():
                if lbl == label:
                    var_depth.set(d)
                    return
        depth_menu.bind("<<ComboboxSelected>>", on_depth_change)

        tk.Label(opt_frame,
                 text="    Applies one rule to the whole appendix. No further popups.",
                 fg="#555", font=("Arial", 9), justify="left").pack(
            anchor="w", padx=24
        )

        # GO FOLDER-BY-FOLDER
        go_row = tk.Frame(opt_frame)
        go_row.pack(fill="x", pady=4, anchor="w")
        tk.Radiobutton(go_row, text="GO FOLDER-BY-FOLDER",
                       variable=var_action, value="go_folder",
                       font=("Arial", 9, "bold")).pack(anchor="w")
        tk.Label(opt_frame,
                 text="    Ask separately for each subfolder inside this appendix.\n"
                      "    Useful when different branches need different treatment.",
                 fg="#555", font=("Arial", 9), justify="left").pack(
            anchor="w", padx=24
        )

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=12)

        def ok():
            action = var_action.get()
            if action == "show":
                result["decision"] = {"action": "show", "depth": var_depth.get()}
            else:
                result["decision"] = {"action": "go_folder"}
            dlg.destroy()

        def cancel():
            result["decision"] = {"action": "cancel"}
            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=12, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), command=ok).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel build", width=14,
                  command=cancel).pack(side="left", padx=6)

        dlg.wait_window()
        return result["decision"]

    def _max_depth_below(folder_path):
        """Return the maximum depth of subfolders below folder_path that have PDFs.
        depth = 0 means no subfolders with PDFs."""
        max_d = 0

        def walk(path, d):
            nonlocal max_d
            try:
                items = os.listdir(path)
            except Exception:
                return
            for name in items:
                full = os.path.join(path, name)
                if os.path.isdir(full) and _tree_has_pdf(full):
                    if d + 1 > max_d:
                        max_d = d + 1
                    walk(full, d + 1)

        walk(folder_path, 0)
        return max_d

    def _is_numeric_name(name):
        """True if folder name is purely digits (e.g. '23', '012')."""
        return name.isdigit()

    def _resolve_folder_decision(folder_rel, folder_name, policy, current_depth_remaining):
        """Resolve action + remaining depth budget for a folder.

        Returns dict:
          {"action": "show", "depth": N}  - show files, depth-limited
          {"action": "flatten"}           - absorb into parent
          {"action": "skip"}              - exclude entirely

        current_depth_remaining is the depth budget inherited from an enclosing
        appendix-level "show all" decision; pass 999 for unlimited or None to
        mean "no inherited budget, look up explicit decision".
        """
        # Manual per-folder override always wins
        if policy["mode"] == "manual":
            decisions = policy.get("decisions", {})
            if folder_rel in decisions:
                d = decisions[folder_rel]
                if isinstance(d, dict):
                    return d
                # legacy string keys (heading / heading_collapsed / flatten / skip)
                if d == "flatten":
                    return {"action": "flatten"}
                if d == "skip":
                    return {"action": "skip"}
                if d == "heading_collapsed":
                    return {"action": "show", "depth": 1}
                return {"action": "show", "depth": 999}
            # No explicit decision - inherit from parent appendix
            if current_depth_remaining is not None:
                return {"action": "show", "depth": current_depth_remaining}
            return {"action": "show", "depth": 999}

        # Automatic mode: layout maps to action
        layout = policy["layout"]
        if layout == "flatten":
            return {"action": "flatten"}
        if layout == "preserve":
            return {"action": "show", "depth": 999}
        if layout == "preserve_skip_numeric":
            if _is_numeric_name(folder_name):
                return {"action": "flatten"}
            return {"action": "show", "depth": 999}
        return {"action": "show", "depth": 999}

    # Backward-compatible wrapper kept in case any code path still calls it
    def _resolve_folder_action(folder_rel, folder_name, policy):
        d = _resolve_folder_decision(folder_rel, folder_name, policy, None)
        if d["action"] == "show" and d.get("depth", 999) == 1:
            return "heading_collapsed"
        if d["action"] == "show":
            return "heading"
        return d["action"]

    def _process_collapsed_subtree(folder_path, merger, current_page,
                                   appendix_idx, file_records, rel_prefix="",
                                   level=2, log_fn=None):
        """Walk a folder subtree, inserting all PDFs into the bundle but
        creating only bookmark-visible records (no TOC entries).
        Used when a folder is marked 'heading_collapsed' - the user sees one
        TOC entry for the whole subtree but the side-panel bookmarks still
        show every PDF for searchability.
        """
        try:
            items = os.listdir(folder_path)
        except Exception:
            return current_page

        custom_order = _read_order_file(folder_path)
        all_pdfs = [
            f for f in items
            if f.lower().endswith(".pdf")
            and os.path.isfile(os.path.join(folder_path, f))
        ]
        all_subs = [
            f for f in items
            if os.path.isdir(os.path.join(folder_path, f))
            and _tree_has_pdf(os.path.join(folder_path, f))
        ]
        pdf_files = _apply_custom_order(all_pdfs, custom_order)
        subfolders = _apply_custom_order(all_subs, custom_order)

        # PDFs - inserted into bundle, bookmark record only (not TOC)
        for pdf in pdf_files:
            start = current_page
            display_name = os.path.splitext(pdf)[0]
            if log_fn:
                log_fn(f"  [collapsed-{level}] Inserting: {pdf}")
            pages = insert_pdf_clean(merger, os.path.join(folder_path, pdf))
            current_page += pages
            end = current_page - 1
            file_records.append({
                "appendix_idx": appendix_idx,
                "file": pdf,
                "display": display_name,
                "level": level,
                "start": start,
                "end": end,
                "toc_visible": False,
                "bookmark_visible": True,
            })

        # Subfolders - recurse with bookmark-only records, no TOC heading
        for sub in subfolders:
            sub_rel = os.path.join(rel_prefix, sub) if rel_prefix else sub
            heading_start = current_page
            file_records.append({
                "appendix_idx": appendix_idx,
                "file": None,
                "display": sub,
                "level": level,
                "start": heading_start,
                "end": heading_start,
                "is_heading": True,
                "toc_visible": False,
                "bookmark_visible": True,
            })
            heading_record_idx = len(file_records) - 1
            current_page = _process_collapsed_subtree(
                os.path.join(folder_path, sub), merger, current_page,
                appendix_idx, file_records,
                rel_prefix=sub_rel,
                level=level + 1,
                log_fn=log_fn,
            )
            file_records[heading_record_idx]["end"] = current_page - 1

        return current_page

    def process_folder_contents(folder_path, merger, current_page, appendix_idx,
                                file_records, policy, rel_prefix="",
                                appendix_root="", level=2,
                                inherited_depth=None, log_fn=None):
        """Policy-driven recursive folder processor.

        appendix_root: the base-relative path of the top-level appendix folder
            (e.g. "Backup"). Combined with rel_prefix to produce the full
            base-relative key used for manual policy lookups.

        level = the TOC nesting level for entries created here:
            level 1 = appendix heading itself (not used here)
            level 2 = first-level subfolders inside an appendix
            level 3, 4, 5 = deeper nesting

        inherited_depth = remaining depth budget passed down from an enclosing
            appendix-level "show all - depth N" decision. None means no inherited
            budget. 1 means stop here (collapse). 999 means unlimited.

        file_records will be populated with:
            { appendix_idx, file, display, level, start, end, toc_visible, bookmark_visible }
        Folder-heading records have file=None and represent a TOC heading.
        """
        try:
            items = os.listdir(folder_path)
        except Exception:
            return current_page

        custom_order = _read_order_file(folder_path)
        all_pdfs = [
            f for f in items
            if f.lower().endswith(".pdf")
            and os.path.isfile(os.path.join(folder_path, f))
        ]
        all_subs = [
            f for f in items
            if os.path.isdir(os.path.join(folder_path, f))
            and _tree_has_pdf(os.path.join(folder_path, f))
        ]
        pdf_files = _apply_custom_order(all_pdfs, custom_order)
        subfolders = _apply_custom_order(all_subs, custom_order)

        sep_mode = policy.get("separators", "folder_only")

        # PDFs at this level
        for pdf in pdf_files:
            start = current_page
            display_name = os.path.splitext(pdf)[0]

            if sep_mode == "every_pdf":
                sep = create_separator(display_name, "", 20, 12)
                with fitz.open(stream=sep, filetype="pdf") as sdoc:
                    merger.insert_pdf(sdoc)
                    current_page += 1

            if log_fn:
                log_fn(f"  [{level}] Inserting: {pdf}")

            pages = insert_pdf_clean(merger, os.path.join(folder_path, pdf))
            current_page += pages
            end = current_page - 1

            file_records.append({
                "appendix_idx": appendix_idx,
                "file": pdf,
                "display": display_name,
                "level": level,
                "start": start,
                "end": end,
                "toc_visible": True,
                "bookmark_visible": True,
            })

        # Subfolders
        for sub in subfolders:
            sub_rel = os.path.join(rel_prefix, sub) if rel_prefix else sub
            sub_rel_display = sub_rel.replace(os.sep, " / ")

            # Build the full base-relative key for policy lookup. Manual decisions
            # were collected with keys like "Backup / Contractor / Plant" relative
            # to the base directory, so we must include the appendix root.
            if appendix_root:
                policy_key = f"{appendix_root} / {sub_rel_display}" if sub_rel_display else appendix_root
            else:
                policy_key = sub_rel_display

            decision = _resolve_folder_decision(
                policy_key, sub, policy, inherited_depth
            )
            action = decision["action"]

            if action == "skip":
                if log_fn:
                    log_fn(f"  [skip] {policy_key}")
                continue

            if action == "flatten":
                if log_fn:
                    log_fn(f"  [flatten] {policy_key}")
                # Pass the SAME inherited_depth (flatten doesn't consume a level)
                current_page = process_folder_contents(
                    os.path.join(folder_path, sub), merger, current_page,
                    appendix_idx, file_records, policy,
                    rel_prefix=sub_rel,
                    appendix_root=appendix_root,
                    level=level,
                    inherited_depth=inherited_depth,
                    log_fn=log_fn,
                )
                continue

            # action == "show": render this folder and propagate depth budget
            depth_here = decision.get("depth", 999)

            heading_start = current_page

            if sep_mode in ("every_pdf", "folder_only"):
                sep = create_separator(sub, "", 20, 12)
                with fitz.open(stream=sep, filetype="pdf") as sdoc:
                    merger.insert_pdf(sdoc)
                    current_page += 1

            if depth_here <= 1:
                # COLLAPSE at this folder: one TOC heading, no per-PDF rows.
                # Bookmarks still get full subtree.
                if log_fn:
                    log_fn(f"  [{level}-collapsed depth={depth_here}] {policy_key}")

                heading_record_idx = len(file_records)
                file_records.append({
                    "appendix_idx": appendix_idx,
                    "file": None,
                    "display": sub,
                    "level": level,
                    "start": heading_start,
                    "end": heading_start,
                    "is_heading": True,
                    "toc_visible": True,
                    "bookmark_visible": True,
                })

                current_page = _process_collapsed_subtree(
                    os.path.join(folder_path, sub), merger, current_page,
                    appendix_idx, file_records,
                    rel_prefix=sub_rel,
                    level=level + 1,
                    log_fn=log_fn,
                )
                file_records[heading_record_idx]["end"] = current_page - 1
                continue

            # depth_here > 1: show heading AND descend, decrementing depth
            if log_fn:
                log_fn(f"  [{level}-heading depth={depth_here}] {policy_key}")

            heading_record_idx = len(file_records)
            file_records.append({
                "appendix_idx": appendix_idx,
                "file": None,
                "display": sub,
                "level": level,
                "start": heading_start,
                "end": heading_start,
                "is_heading": True,
                "toc_visible": True,
                "bookmark_visible": True,
            })

            # Compute next inherited depth: 999 stays 999, finite value decrements
            next_depth = 999 if depth_here >= 999 else depth_here - 1

            current_page = process_folder_contents(
                os.path.join(folder_path, sub), merger, current_page,
                appendix_idx, file_records, policy,
                rel_prefix=sub_rel,
                appendix_root=appendix_root,
                level=level + 1,
                inherited_depth=next_depth,
                log_fn=log_fn,
            )

            file_records[heading_record_idx]["end"] = current_page - 1

        return current_page

    def compress_document(doc, mode="high", log_fn=None):
        """Apply compression based on mode."""
        if mode == "high":
            return  # No compression

        if mode == "balanced":
            # Moderate downsampling: cap at 2400 px on longest edge, JPEG quality 80
            max_edge_target = 2400
            jpeg_quality = 80
        else:  # small
            max_edge_target = 1500
            jpeg_quality = 60

        if not OCR_AVAILABLE:  # Pillow comes with OCR deps
            if log_fn:
                log_fn("  Skipping compression (Pillow not available).")
            return

        try:
            from PIL import Image as PILImage
        except Exception:
            return

        compressed = 0
        for page_num in range(doc.page_count):
            try:
                page = doc.load_page(page_num)
                imgs = page.get_images(full=True)
                for img_info in imgs:
                    xref = img_info[0]
                    try:
                        img = doc.extract_image(xref)
                        imgdata = img["image"]
                        pil = PILImage.open(io.BytesIO(imgdata))
                        max_edge = max(pil.size)
                        scale = min(1.0, max_edge_target / max_edge) if max_edge > 0 else 1.0
                        if scale < 1.0:
                            new_size = (int(pil.size[0] * scale), int(pil.size[1] * scale))
                            pil = pil.resize(new_size, PILImage.LANCZOS)
                        buff = io.BytesIO()
                        pil = pil.convert("RGB")
                        pil.save(buff, format="JPEG", optimize=True, quality=jpeg_quality)
                        new_stream = buff.getvalue()
                        if len(new_stream) < len(imgdata):
                            try:
                                doc.update_stream(xref, new_stream)
                                compressed += 1
                            except Exception:
                                pass
                    except Exception:
                        continue
            except Exception:
                continue
        if log_fn:
            log_fn(f"  Compressed {compressed} images ({mode} mode).")

    # -----------------------------------------------------------------------
    # Step 6: Settings persistence
    # -----------------------------------------------------------------------

    DEFAULT_SETTINGS = {
        "last_directory": "",
        "claim_title": "",
        "numbering_style": "alphabetic",
        "manual_descriptions": False,
        "generate_toc": True,
        "do_ocr": False,
        "compression": "high",
        "include_cover": False,
        "cover_project": "",
        "cover_contract_ref": "",
        "cover_revision": "Rev 00",
        "cover_doc_type": "Claims Bundle",
        "style_cover": "top_rule",
        "style_palette": "navy_gold",
        "style_font": "aptos",
        "style_separator": "top_rule",
        "style_toc": "ruled",
    }

    def load_settings():
        if not os.path.isfile(SETTINGS_FILE):
            return dict(DEFAULT_SETTINGS)
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            merged = dict(DEFAULT_SETTINGS)
            merged.update({k: v for k, v in data.items() if k in DEFAULT_SETTINGS})
            return merged
        except Exception:
            return dict(DEFAULT_SETTINGS)

    def save_settings(settings):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Step 7: Build engine (called from worker thread)
    # -----------------------------------------------------------------------

    def build_bundle(opts, log_fn, done_fn):
        """Run the full bundle build. opts is a dict, log_fn(str) logs progress."""
        try:
            base_directory = opts["base_directory"]
            claim_title = opts["claim_title"]
            numbering_style = opts["numbering_style"]
            is_manual = opts["manual_descriptions"]
            generate_toc = opts["generate_toc"]
            do_ocr = opts["do_ocr"]
            comp_mode = opts["compression"]
            include_cover = opts["include_cover"]

            # Apply the selected bundle style for this build
            CURRENT_STYLE.update(opts.get("style") or {})

            if do_ocr and not OCR_AVAILABLE:
                log_fn("WARNING: OCR libraries not available. Proceeding without OCR.")
                do_ocr = False

            out_name = sanitize_text(claim_title) if claim_title else "Combined_Appendices"
            output_path = os.path.join(base_directory, f"{out_name}.pdf")

            log_fn(f"Output target: {output_path}")
            log_fn("Discovering inputs...")

            base_items_all = os.listdir(base_directory)
            base_custom_order = _read_order_file(base_directory)
            if base_custom_order:
                log_fn(f"  Found _order.txt - using custom order")

            base_pdfs_all = [
                f for f in base_items_all
                if f.lower().endswith(".pdf")
                and os.path.isfile(os.path.join(base_directory, f))
                and f != os.path.basename(output_path)
            ]
            # Only include folders whose subtree contains at least one PDF.
            # Empty folders or folder trees with no PDFs are excluded entirely
            # - they get no appendix label, no separator page, no TOC entry.
            base_folders_all = [
                f for f in base_items_all
                if os.path.isdir(os.path.join(base_directory, f))
                and _tree_has_pdf(os.path.join(base_directory, f))
            ]

            base_pdfs = _apply_custom_order(base_pdfs_all, base_custom_order)
            base_folders = _apply_custom_order(base_folders_all, base_custom_order)

            log_fn(f"  {len(base_pdfs)} top-level PDF(s)")
            log_fn(f"  {len(base_folders)} top-level folder(s)")

            if not base_pdfs and not base_folders:
                log_fn("ERROR: No PDFs or folders found in base directory.")
                done_fn(False, "No content to merge.")
                return

            merger = fitz.open()
            current_page = 1

            appendices = []
            files_lvl2 = []
            appendix_meta = []

            # Top-level PDFs as appendices
            for pdf_name in base_pdfs:
                name_no_ext = os.path.splitext(pdf_name)[0]
                app_idx = len(appendices) + 1
                app_label = convert_to_style(app_idx, numbering_style)
                description = opts["manual_titles"].get(pdf_name, name_no_ext) if is_manual else name_no_ext

                log_fn(f"Appendix {app_label}: {pdf_name}")
                start_page_app = current_page
                sep = create_separator(f"Appendix {app_label}", description)
                with fitz.open(stream=sep, filetype="pdf") as sdoc:
                    merger.insert_pdf(sdoc)
                    current_page += 1

                file_start = current_page
                pages = insert_pdf_clean(merger, os.path.join(base_directory, pdf_name))
                current_page += pages
                file_end = current_page - 1
                files_lvl2.append({
                    "appendix_idx": len(appendices),
                    "file": pdf_name,
                    "display": name_no_ext,
                    "level": 2,
                    "start": file_start, "end": file_end,
                })
                appendices.append({
                    "label": app_label, "title": description,
                    "start": start_page_app, "end": current_page - 1,
                })
                appendix_meta.append({"include_children": False})

            # Top-level folders as appendices
            for folder in base_folders:
                folder_path = os.path.join(base_directory, folder)

                app_idx = len(appendices) + 1
                app_label = convert_to_style(app_idx, numbering_style)
                description = opts["manual_titles"].get(folder, folder) if is_manual else folder

                log_fn(f"Appendix {app_label}: {folder} (folder)")
                start_page_app = current_page
                sep = create_separator(f"Appendix {app_label}", description)
                with fitz.open(stream=sep, filetype="pdf") as sdoc:
                    merger.insert_pdf(sdoc)
                    current_page += 1

                # Determine inherited depth budget for this appendix.
                # Manual mode may have stored an appendix-level decision.
                appendix_decisions = opts["policy"].get("appendix_decisions", {})
                app_decision = appendix_decisions.get(folder)
                use_collapsed_walker = False
                inherited_depth = None

                if app_decision and app_decision.get("action") == "show":
                    depth_at_appendix = app_decision.get("depth", 999)
                    if depth_at_appendix <= 1:
                        # Whole appendix collapses to one TOC line
                        use_collapsed_walker = True
                    else:
                        # Pass remaining depth budget down. Depth N at appendix
                        # means N levels visible BELOW the appendix (which is
                        # level 1). So children at level 2 use depth = N - 1
                        # downstream.
                        if depth_at_appendix >= 999:
                            inherited_depth = 999
                        else:
                            inherited_depth = depth_at_appendix - 1
                # Otherwise (no decision, or "go_folder"): per-subfolder
                # decisions take effect, no inherited depth.

                if use_collapsed_walker:
                    # Whole appendix collapses: insert all PDFs under this folder
                    # but produce no TOC entries at all under the appendix line.
                    log_fn(f"  [appendix collapsed] {folder}")
                    current_page = _process_collapsed_subtree(
                        folder_path, merger, current_page, len(appendices),
                        files_lvl2, rel_prefix="", level=2, log_fn=log_fn,
                    )
                else:
                    current_page = process_folder_contents(
                        folder_path, merger, current_page, len(appendices),
                        files_lvl2, opts["policy"], rel_prefix="",
                        appendix_root=folder, level=2,
                        inherited_depth=inherited_depth,
                        log_fn=log_fn,
                    )
                appendices.append({
                    "label": app_label, "title": description,
                    "start": start_page_app, "end": current_page - 1,
                })
                # Folder appendices always show their PDF children in the TOC -
                # this makes deep folder structures navigable.
                appendix_meta.append({"include_children": True})

            # Optional cover page (prepended)
            cover_page_count = 0
            if include_cover:
                log_fn("Generating cover page...")
                cover_meta = {
                    "title": claim_title or "Claims Bundle",
                    "subtitle": opts.get("cover_subtitle", ""),
                    "doc_type": opts.get("cover_doc_type", ""),
                    "project": opts.get("cover_project", ""),
                    "contract_ref": opts.get("cover_contract_ref", ""),
                    "prepared_by": opts.get("cover_prepared_by", ""),
                    "prepared_for": opts.get("cover_prepared_for", ""),
                    "date": opts.get("cover_date", ""),
                    "revision": opts.get("cover_revision", ""),
                }
                cover_bytes = create_cover_page(cover_meta)
                with fitz.open(stream=cover_bytes, filetype="pdf") as cover_doc:
                    cover_page_count = cover_doc.page_count
                    full_doc_with_cover = fitz.open()
                    full_doc_with_cover.insert_pdf(cover_doc)
                    full_doc_with_cover.insert_pdf(merger)
                merger.close()
                merger = full_doc_with_cover

                # Shift recorded page numbers
                for a in appendices:
                    a["start"] += cover_page_count
                    a["end"] += cover_page_count
                for ch in files_lvl2:
                    ch["start"] += cover_page_count
                    ch["end"] += cover_page_count

            # TOC insertion
            toc_page_count = 0
            if generate_toc and appendices:
                log_fn("Building Table of Contents...")

                # Build TOC entries using UNSHIFTED page numbers (pre-TOC).
                # The links will compensate by adding toc_page_count when inserting.
                # But we ALSO want the TOC text to display the FINAL page numbers,
                # so we use a two-pass approach:
                #   Pass 1: discover TOC page count using unshifted numbers
                #   Pass 2: rebuild with shifted numbers for display
                # The link target_zero in pass 2 entries is already the FINAL 0-based page index.

                def build_toc_entries(toc_offset):
                    """Build flat TOC entry list with level field.
                    Each entry: {level, title, page, show_page}.
                    Level 1 = appendix; Level 2-5 = nested folder/file rows.
                    """
                    entries = []
                    for i, a in enumerate(appendices):
                        shifted_start = a["start"] + toc_offset
                        shifted_end = a["end"] + toc_offset
                        rng = f"[pp {shifted_start}-{shifted_end}]"
                        entries.append({
                            "level": 1,
                            "title": f"Appendix {a['label']}: {a['title']}  {rng}",
                            "page": shifted_start,
                            "show_page": True,
                        })
                        if not appendix_meta[i].get("include_children", False):
                            continue
                        for ch in [f for f in files_lvl2 if f["appendix_idx"] == i]:
                            # Skip records that are bookmark-only (collapsed subtree)
                            if not ch.get("toc_visible", True):
                                continue
                            ch_level = ch.get("level", 2)
                            if ch.get("is_heading"):
                                # Folder heading row (clickable to its first page)
                                entries.append({
                                    "level": ch_level,
                                    "title": ch.get("display", "(folder)"),
                                    "page": ch["start"] + toc_offset,
                                    "show_page": True,
                                })
                            else:
                                # File leaf row
                                title = ch.get("display") or ch.get("file") or ""
                                if title.lower().endswith(".pdf"):
                                    title = title[:-4]
                                entries.append({
                                    "level": ch_level,
                                    "title": title,
                                    "page": ch["start"] + toc_offset,
                                    "show_page": True,
                                })
                    return entries

                # First pass: build with offset=0 to discover TOC page count
                toc_entries = build_toc_entries(0)
                _, _, toc_page_count = create_toc_page_hier(claim_title, toc_entries)

                # Second pass: rebuild with the actual offset to get correct page numbers
                toc_entries = build_toc_entries(toc_page_count)
                toc_bytes, link_pos, toc_page_count_check = create_toc_page_hier(
                    claim_title, toc_entries
                )

                # Edge case: if the second pass produced a different page count
                # (very rare, only if shifted page numbers cross a digit boundary
                # that wraps a line), do a third pass.
                if toc_page_count_check != toc_page_count:
                    toc_page_count = toc_page_count_check
                    toc_entries = build_toc_entries(toc_page_count)
                    toc_bytes, link_pos, toc_page_count = create_toc_page_hier(
                        claim_title, toc_entries
                    )

                log_fn(f"  TOC: {toc_page_count} page(s), {len(link_pos)} link(s)")

                # Merge: cover (if any) + TOC + body
                full_doc = fitz.open()
                if cover_page_count > 0:
                    full_doc.insert_pdf(merger, from_page=0, to_page=cover_page_count - 1)
                    with fitz.open(stream=toc_bytes, filetype="pdf") as toc_doc:
                        full_doc.insert_pdf(toc_doc)
                    full_doc.insert_pdf(merger, from_page=cover_page_count,
                                        to_page=merger.page_count - 1)
                else:
                    with fitz.open(stream=toc_bytes, filetype="pdf") as toc_doc:
                        full_doc.insert_pdf(toc_doc)
                    full_doc.insert_pdf(merger)
                merger.close()

                # CRITICAL: round-trip the document via tobytes() to settle internal
                # state. Adding links directly to a freshly-merged page causes
                # PyMuPDF to malform the first link's rectangle (it ends up
                # covering the whole page). Re-opening from bytes resets the
                # page state cleanly.
                merged_bytes = full_doc.tobytes()
                full_doc.close()
                merger = fitz.open(stream=merged_bytes, filetype="pdf")

                # Now add links to the TOC pages of the re-opened document.
                # Per official PyMuPDF docs: "most annotation updates require
                # reloading the page via page = doc.reload_page(page)". This
                # is the documented fix for the issue where the first link on
                # a page becomes a giant rectangle covering everything.
                toc_start_idx = cover_page_count  # 0-based index where TOC begins
                left_mm = LEFT_MARGIN_MM
                width_mm = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM

                for entry in link_pos:
                    page_in_toc = entry["page_in_toc"]
                    abs_toc_page_idx = toc_start_idx + page_in_toc
                    if abs_toc_page_idx >= merger.page_count:
                        continue
                    target_zero = entry["target_zero"]
                    if target_zero < 0 or target_zero >= merger.page_count:
                        continue

                    # Load fresh page object, insert link, then reload to flush
                    # PyMuPDF's internal annotation cache before next iteration.
                    toc_page = merger.load_page(abs_toc_page_idx)
                    rect = mm_rect(
                        left_mm, entry["y0_mm"],
                        left_mm + width_mm, entry["y1_mm"]
                    )
                    try:
                        toc_page.insert_link({
                            "from": rect,
                            "kind": fitz.LINK_GOTO,
                            "page": target_zero,
                        })
                        # CRITICAL: reload the page after insert_link so the
                        # next iteration sees the updated annotation tree.
                        merger.reload_page(toc_page)
                    except Exception as link_err:
                        log_fn(f"  Skipped link (target {target_zero}): {link_err}")

                # Bookmarks (PDF side-panel outline)
                outline = []
                if cover_page_count > 0:
                    outline.append((1, "Cover", 1))
                outline.append((1, "Table of Contents", cover_page_count + 1))
                for i, a in enumerate(appendices):
                    a["start"] += toc_page_count
                    a["end"] += toc_page_count
                    outline.append((1, f"Appendix {a['label']}: {a['title']}", a["start"] + 1))
                    if not appendix_meta[i].get("include_children", False):
                        continue
                    for ch in [x for x in files_lvl2 if x["appendix_idx"] == i]:
                        ch["start"] += toc_page_count
                        ch["end"] += toc_page_count
                        # Bookmark visibility is independent of TOC visibility:
                        # collapsed subtrees still get bookmark entries so the
                        # side panel shows everything for searchability.
                        if not ch.get("bookmark_visible", True):
                            continue
                        title = ch.get("display") or ch.get("file") or ""
                        if isinstance(title, str) and title.lower().endswith(".pdf"):
                            title = title[:-4]
                        # Use level from record; PDF outline accepts up to ~6 levels
                        bm_level = max(2, min(ch.get("level", 2), 6))
                        outline.append((bm_level, title, ch["start"] + 1))
                merger.set_toc(outline)

            # OCR
            if do_ocr:
                log_fn("Running OCR on pages without text...")
                total = merger.page_count
                ocr_count = 0
                for pno in range(total):
                    try:
                        page = merger.load_page(pno)
                        if ocr_page_if_needed(page):
                            ocr_count += 1
                        if pno % 10 == 0:
                            log_fn(f"  OCR scanning page {pno + 1} of {total}")
                    except Exception as e:
                        log_fn(f"  OCR error on page {pno + 1}: {e}")
                log_fn(f"  OCR added text to {ocr_count} page(s).")

            # Trim trailing blank pages
            log_fn("Trimming trailing blank pages...")
            trimmed = 0
            while merger.page_count > 0:
                last_idx = merger.page_count - 1
                last_page = merger.load_page(last_idx)
                if _page_has_content(last_page):
                    break
                merger.delete_page(last_idx)
                trimmed += 1
            if trimmed:
                log_fn(f"  Removed {trimmed} blank trailing page(s).")

            # Footers
            log_fn("Adding appendix footers...")
            for a in appendices:
                total_pages = a["end"] - a["start"] + 1
                right_label = f"Appendix {a['label']}: {a['title']}"
                for idx, p in enumerate(range(a["start"] - 1, a["end"])):
                    if 0 <= p < merger.page_count:
                        left = f"App {a['label']} - P{idx + 1:03d} / {total_pages:03d}"
                        try:
                            draw_footer_numbering(merger, p, left, right_label)
                        except Exception as e:
                            log_fn(f"  Footer skip page {p + 1}: {e}")

            # Compression
            if comp_mode != "high":
                log_fn(f"Applying {comp_mode} compression...")
                compress_document(merger, mode=comp_mode, log_fn=log_fn)

            # PDF metadata
            try:
                merger.set_metadata({
                    "title": claim_title or "Claims Bundle",
                    "author": opts.get("cover_prepared_by", "") or "",
                    "subject": opts.get("cover_project", ""),
                    "keywords": "Claim Bundle, Appendices",
                })
            except Exception:
                pass

            # Save
            log_fn("Saving final PDF...")
            try:
                merger.save(output_path, garbage=4, deflate=True)
                merger.close()
                log_fn(f"DONE. Bundle saved: {output_path}")
                done_fn(True, output_path)
            except Exception as e:
                log_fn(f"ERROR during save: {e}")
                done_fn(False, str(e))

        except Exception:
            log_fn("UNEXPECTED ERROR:")
            log_fn(traceback.format_exc())
            done_fn(False, "Unexpected error - see status panel.")

    # -----------------------------------------------------------------------
    # Step 8: GUI
    # -----------------------------------------------------------------------

    class ClaimsBundleApp:
        def __init__(self, root):
            self.root = root
            self.root.title("Claims Bundle Builder")
            self.root.geometry("980x1020")
            self.root.minsize(900, 760)

            self.settings = load_settings()
            self._build_ui()
            self._load_settings_into_ui()

        def _build_ui(self):
            pad = {"padx": 8, "pady": 4}

            # --- Fixed bottom bar: credit, action buttons, status panel ---
            # Packed side="bottom" first so nothing above can push them
            # off-screen regardless of display height.
            tk.Label(self.root, text="QS Alaa Elsayed", font=("Arial", 8),
                     fg="#666666", anchor="e").pack(side="bottom", fill="x",
                                                    padx=10, pady=(0, 4))

            act = tk.Frame(self.root)
            act.pack(side="bottom", fill="x", **pad)
            self.btn_run = tk.Button(
                act, text="Build Bundle", command=self._on_run,
                font=("Arial", 10, "bold"),
                bg="#1F3864", fg="white", padx=20, pady=6,
            )
            self.btn_run.pack(side="left")
            tk.Button(act, text="Quit", command=self._on_quit, padx=20, pady=6).pack(side="right")

            self.txt_status = scrolledtext.ScrolledText(
                self.root, height=9, font=("Consolas", 9), state="disabled"
            )
            self.txt_status.pack(side="bottom", fill="x", **pad)
            tk.Label(self.root, text="Status / Progress:",
                     font=("Arial", 9, "bold")).pack(side="bottom", fill="x", padx=8)

            # --- Scrollable options area (everything else lives in `body`) ---
            container = tk.Frame(self.root)
            container.pack(side="top", fill="both", expand=True)
            self._opts_canvas = tk.Canvas(container, highlightthickness=0)
            vsb = tk.Scrollbar(container, orient="vertical",
                               command=self._opts_canvas.yview)
            self._opts_canvas.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            self._opts_canvas.pack(side="left", fill="both", expand=True)
            body = tk.Frame(self._opts_canvas)
            body_id = self._opts_canvas.create_window((0, 0), window=body,
                                                      anchor="nw")

            def _sync_scroll(event=None):
                self._opts_canvas.configure(
                    scrollregion=self._opts_canvas.bbox("all"))
                self._opts_canvas.itemconfigure(
                    body_id, width=self._opts_canvas.winfo_width())

            body.bind("<Configure>", _sync_scroll)
            self._opts_canvas.bind("<Configure>", _sync_scroll)

            def _on_wheel(event):
                # Scroll only when the pointer is over the options pane
                w = self.root.winfo_containing(event.x_root, event.y_root)
                while w is not None:
                    if w is self._opts_canvas:
                        self._opts_canvas.yview_scroll(
                            int(-event.delta / 120), "units")
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

            # Claim title
            title_frame = tk.Frame(body)
            title_frame.pack(fill="x", **pad)
            tk.Label(title_frame, text="Claim title:", font=("Arial", 9, "bold")).pack(side="left", padx=(8, 6))
            self.entry_claim_title = tk.Entry(title_frame, font=("Arial", 10))
            self.entry_claim_title.pack(side="left", fill="x", expand=True, padx=(0, 8))

            # Options row
            opt_frame = tk.LabelFrame(body, text="Options")
            opt_frame.pack(fill="x", **pad)

            row1 = tk.Frame(opt_frame)
            row1.pack(fill="x", padx=6, pady=4)
            tk.Label(row1, text="Numbering:", font=("Arial", 9)).pack(side="left")
            self.var_numbering = tk.StringVar(value="alphabetic")
            tk.Radiobutton(row1, text="Alphabetic (A, B, C...)", variable=self.var_numbering, value="alphabetic").pack(side="left", padx=4)
            tk.Radiobutton(row1, text="Numeric (01, 02, 03...)", variable=self.var_numbering, value="numeric").pack(side="left", padx=4)

            row2 = tk.Frame(opt_frame)
            row2.pack(fill="x", padx=6, pady=4)
            self.var_toc = tk.BooleanVar(value=True)
            tk.Checkbutton(row2, text="Generate Table of Contents (clickable, multi-page safe)", variable=self.var_toc).pack(side="left")

            row3 = tk.Frame(opt_frame)
            row3.pack(fill="x", padx=6, pady=4)
            self.var_manual = tk.BooleanVar(value=False)
            tk.Checkbutton(row3, text="Manual mode (prompt for description of each appendix)", variable=self.var_manual).pack(side="left")

            row4 = tk.Frame(opt_frame)
            row4.pack(fill="x", padx=6, pady=4)
            self.var_ocr = tk.BooleanVar(value=False)
            ocr_text = "Run OCR on scanned pages (English + Arabic)" + ("" if OCR_AVAILABLE else " - NOT AVAILABLE (Pillow + pytesseract required)")
            ocr_cb = tk.Checkbutton(row4, text=ocr_text, variable=self.var_ocr)
            ocr_cb.pack(side="left")
            if not OCR_AVAILABLE:
                ocr_cb.configure(state="disabled")

            row5 = tk.Frame(opt_frame)
            row5.pack(fill="x", padx=6, pady=4)
            tk.Label(row5, text="Compression:", font=("Arial", 9)).pack(side="left")
            self.var_compression = tk.StringVar(value="high")
            tk.Radiobutton(row5, text="High Quality (largest)", variable=self.var_compression, value="high").pack(side="left", padx=4)
            tk.Radiobutton(row5, text="Balanced", variable=self.var_compression, value="balanced").pack(side="left", padx=4)
            tk.Radiobutton(row5, text="Smallest File", variable=self.var_compression, value="small").pack(side="left", padx=4)

            # Cover page
            cover_frame = tk.LabelFrame(body, text="Cover Page (optional)")
            cover_frame.pack(fill="x", **pad)

            self.var_cover = tk.BooleanVar(value=False)
            tk.Checkbutton(cover_frame, text="Include cover page", variable=self.var_cover,
                           command=self._toggle_cover_fields).pack(anchor="w", padx=6)

            self.cover_fields = tk.Frame(cover_frame)
            self.cover_fields.pack(fill="x", padx=6, pady=4)
            self.cover_entries = {}
            cover_fields_def = [
                ("Document type", "cover_doc_type"),
                ("Project", "cover_project"),
                ("Contract Reference", "cover_contract_ref"),
                ("Prepared by", "cover_prepared_by"),
                ("Prepared for", "cover_prepared_for"),
                ("Revision", "cover_revision"),
            ]
            for i, (label, key) in enumerate(cover_fields_def):
                tk.Label(self.cover_fields, text=f"{label}:", font=("Arial", 9), width=18, anchor="w").grid(row=i, column=0, padx=2, pady=2, sticky="w")
                e = tk.Entry(self.cover_fields, font=("Arial", 9))
                e.grid(row=i, column=1, padx=2, pady=2, sticky="ew")
                self.cover_entries[key] = e
            self.cover_fields.columnconfigure(1, weight=1)

            # Bundle style (this-or-this pickers with thumbnail previews)
            style_frame = tk.LabelFrame(body, text="Bundle Style")
            style_frame.pack(fill="x", **pad)

            self.var_style_cover = tk.StringVar(value="top_rule")
            self.var_style_palette = tk.StringVar(value="navy_gold")
            self.var_style_font = tk.StringVar(value="aptos")
            self.var_style_separator = tk.StringVar(value="top_rule")
            self.var_style_toc = tk.StringVar(value="ruled")
            self._style_refreshers = []

            NAVY, GOLD, STEEL, GREY = "#051641", "#FFC425", "#7A8699", "#d9d9d9"

            def draw_cover(cv, value, accent):
                cv.create_rectangle(0, 0, 64, 3, fill=NAVY, outline="")
                if value == "top_rule":
                    cv.create_rectangle(8, 16, 44, 20, fill=NAVY, outline="")
                    cv.create_rectangle(8, 22, 36, 26, fill=NAVY, outline="")
                    cv.create_rectangle(8, 30, 20, 32, fill=accent, outline="")
                    for k, w in enumerate((30, 24, 27)):
                        cv.create_rectangle(8, 62 + k * 7, 8 + w, 64 + k * 7, fill=GREY, outline="")
                else:
                    cv.create_rectangle(16, 20, 48, 24, fill=NAVY, outline="")
                    cv.create_rectangle(20, 26, 44, 30, fill=NAVY, outline="")
                    cv.create_rectangle(28, 34, 36, 35, fill=NAVY, outline="")
                    for k, w in enumerate((28, 22)):
                        cv.create_rectangle(32 - w // 2, 64 + k * 7, 32 + w // 2, 66 + k * 7, fill=GREY, outline="")

            def draw_palette(cv, value, accent):
                cv.create_rectangle(8, 14, 56, 44, fill=NAVY, outline="")
                cv.create_rectangle(8, 50, 56, 66, fill=accent, outline="")

            def draw_separator(cv, value, accent):
                if value == "top_rule":
                    cv.create_rectangle(0, 0, 64, 3, fill=NAVY, outline="")
                    cv.create_rectangle(8, 34, 32, 37, fill=NAVY, outline="")
                    cv.create_rectangle(8, 41, 44, 45, fill=NAVY, outline="")
                    cv.create_rectangle(8, 49, 20, 51, fill=accent, outline="")
                else:
                    cv.create_rectangle(12, 26, 52, 62, outline=NAVY)
                    cv.create_rectangle(20, 34, 44, 37, fill=NAVY, outline="")
                    cv.create_rectangle(18, 42, 46, 45, fill=NAVY, outline="")
                    cv.create_rectangle(22, 49, 42, 52, fill=NAVY, outline="")

            def draw_toc(cv, value, accent):
                cv.create_rectangle(8, 10, 28, 14, fill=NAVY, outline="")
                cv.create_rectangle(8, 16, 20, 18, fill=accent, outline="")
                for k in range(3):
                    y = 30 + k * 16
                    cv.create_rectangle(8, y, 14, y + 6, fill=NAVY, outline="")
                    cv.create_rectangle(17, y + 1, 38, y + 4, fill=NAVY, outline="")
                    cv.create_rectangle(50, y + 1, 56, y + 4, fill=GREY, outline="")
                    if value == "ruled":
                        cv.create_line(8, y + 9, 56, y + 9, fill=NAVY)
                    else:
                        for dx in range(40, 49, 3):
                            cv.create_rectangle(dx, y + 4, dx + 1, y + 5, fill=GREY, outline="")

            def add_style_row(label, var, options, draw_fn, height=78):
                row = tk.Frame(style_frame)
                row.pack(fill="x", padx=6, pady=3, anchor="w")
                tk.Label(row, text=label, font=("Arial", 9), width=12,
                         anchor="w").pack(side="left")
                canvases = []

                def refresh():
                    for value, cv in canvases:
                        sel = var.get() == value
                        cv.configure(highlightbackground="#0045BF" if sel else "#c8c8c8",
                                     highlightthickness=2 if sel else 1)

                for value, caption in options:
                    holder = tk.Frame(row)
                    holder.pack(side="left", padx=6)
                    cv = tk.Canvas(holder, width=64, height=height, bg="white",
                                   highlightthickness=1,
                                   highlightbackground="#c8c8c8", cursor="hand2")
                    cv.pack()
                    draw_fn(cv, value, GOLD if value != "navy_steel" else STEEL)
                    cv.bind("<Button-1>", lambda e, v=value: (var.set(v), refresh()))
                    tk.Label(holder, text=caption, font=("Arial", 8)).pack()
                    canvases.append((value, cv))
                refresh()
                self._style_refreshers.append(refresh)

            add_style_row("Cover:", self.var_style_cover,
                          [("top_rule", "Top rule"), ("classical", "Classical centred")],
                          draw_cover)
            add_style_row("Palette:", self.var_style_palette,
                          [("navy_gold", "Navy + gold"), ("navy_steel", "Navy + steel")],
                          draw_palette, height=72)
            font_row = tk.Frame(style_frame)
            font_row.pack(fill="x", padx=6, pady=3, anchor="w")
            tk.Label(font_row, text="Font:", font=("Arial", 9), width=12,
                     anchor="w").pack(side="left")
            tk.Radiobutton(font_row, text="Aptos (brand, falls back to Segoe UI)",
                           variable=self.var_style_font, value="aptos").pack(side="left", padx=4)
            tk.Radiobutton(font_row, text="Times New Roman (court bundle)",
                           variable=self.var_style_font, value="times").pack(side="left", padx=4)
            add_style_row("Separator:", self.var_style_separator,
                          [("top_rule", "Top rule"), ("plaque", "Plaque frame")],
                          draw_separator)
            add_style_row("Contents:", self.var_style_toc,
                          [("ruled", "Ruled ledger"), ("dotted", "Dotted leaders")],
                          draw_toc)


        def _toggle_cover_fields(self):
            state = "normal" if self.var_cover.get() else "disabled"
            for e in self.cover_entries.values():
                e.configure(state=state)

        def _load_settings_into_ui(self):
            s = self.settings
            self.entry_dir.insert(0, s.get("last_directory", ""))
            self.entry_claim_title.insert(0, s.get("claim_title", ""))
            self.var_numbering.set(s.get("numbering_style", "alphabetic"))
            self.var_toc.set(s.get("generate_toc", True))
            self.var_manual.set(s.get("manual_descriptions", False))
            self.var_ocr.set(s.get("do_ocr", False) and OCR_AVAILABLE)
            self.var_compression.set(s.get("compression", "high"))
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
            for refresh in self._style_refreshers:
                refresh()
            self._toggle_cover_fields()

        def _save_current_settings(self):
            self.settings["last_directory"] = self.entry_dir.get().strip()
            self.settings["claim_title"] = self.entry_claim_title.get().strip()
            self.settings["numbering_style"] = self.var_numbering.get()
            self.settings["generate_toc"] = self.var_toc.get()
            self.settings["manual_descriptions"] = self.var_manual.get()
            self.settings["do_ocr"] = self.var_ocr.get()
            self.settings["compression"] = self.var_compression.get()
            self.settings["include_cover"] = self.var_cover.get()
            for key, entry in self.cover_entries.items():
                if key in ("cover_prepared_by", "cover_prepared_for"):
                    continue  # identity fields are never persisted
                self.settings[key] = entry.get().strip()
            self.settings["style_cover"] = self.var_style_cover.get()
            self.settings["style_palette"] = self.var_style_palette.get()
            self.settings["style_font"] = self.var_style_font.get()
            self.settings["style_separator"] = self.var_style_separator.get()
            self.settings["style_toc"] = self.var_style_toc.get()
            save_settings(self.settings)

        def _pick_dir(self):
            path = filedialog.askdirectory(title="Select base directory containing PDFs / folders")
            if path:
                self.entry_dir.delete(0, "end")
                self.entry_dir.insert(0, path)

        def _set_status(self, msg, append=False):
            self.txt_status.configure(state="normal")
            if not append:
                self.txt_status.delete("1.0", "end")
            self.txt_status.insert("end", msg + "\n")
            self.txt_status.see("end")
            self.txt_status.configure(state="disabled")

        def _log_safe(self, msg):
            self.root.after(0, lambda m=msg: self._set_status(m, append=True))

        def _enable_run_safe(self):
            self.root.after(0, lambda: self.btn_run.configure(state="normal"))

        def _on_quit(self):
            self._save_current_settings()
            self.root.quit()

        def _on_run(self):
            base_dir = self.entry_dir.get().strip()
            if not base_dir or not os.path.isdir(base_dir):
                messagebox.showwarning("Invalid directory", "Pick a valid base directory.")
                return

            # Cover identity is never embedded as a default - if the cover is
            # on and these are blank, ask the user now.
            if self.var_cover.get():
                from tkinter import simpledialog
                for key, label in (("cover_prepared_by", "Prepared by"),
                                   ("cover_prepared_for", "Prepared for")):
                    if not self.cover_entries[key].get().strip():
                        val = simpledialog.askstring(
                            "Cover details",
                            f"{label} (leave blank to omit from the cover):",
                            parent=self.root)
                        if val and val.strip():
                            self.cover_entries[key].delete(0, "end")
                            self.cover_entries[key].insert(0, val.strip())

            claim_title = self.entry_claim_title.get().strip()
            if not claim_title:
                if not messagebox.askyesno("No claim title", "Claim title is empty - file will be named 'Combined_Appendices.pdf'. Continue?"):
                    return

            # ---------- Folder policy resolution ----------
            counts = _count_folder_tree(base_dir)
            policy = None

            if counts["top_folders"] == 0 and counts["total_subfolders"] == 0:
                # No folders at all - bypass policy dialogs entirely
                policy = {
                    "mode": "automatic",
                    "layout": "preserve",
                    "separators": "folder_only",
                    "decisions": {},
                    "appendix_decisions": {},
                }
            else:
                while True:
                    stage1 = show_discovery_popup(self.root, counts)
                    if stage1 is None:
                        return  # cancelled
                    if stage1 == "automatic":
                        result = show_automatic_options_dialog(self.root)
                        if result == "BACK":
                            continue
                        if result is None:
                            return
                        policy = {
                            "mode": "automatic",
                            "layout": result["layout"],
                            "separators": result["separators"],
                            "decisions": {},
                            "appendix_decisions": {},
                        }
                        break
                    if stage1 == "manual":
                        sep_choice = show_manual_separators_dialog(self.root)
                        if sep_choice == "BACK":
                            continue
                        if sep_choice is None:
                            return

                        # Discover top-level appendix folders (in same order
                        # as build_bundle iterates them) and walk each one.
                        try:
                            base_items_all = sorted(os.listdir(base_dir))
                        except Exception:
                            base_items_all = []
                        appendix_folders = [
                            f for f in base_items_all
                            if os.path.isdir(os.path.join(base_dir, f))
                            and _tree_has_pdf(os.path.join(base_dir, f))
                        ]

                        appendix_decisions = {}  # {folder_name: {"action", "depth"}}
                        per_folder_decisions = {}  # {full_rel_path: decision_dict}
                        cancelled = False

                        for appx_idx, appendix_folder in enumerate(appendix_folders, start=1):
                            appx_path = os.path.join(base_dir, appendix_folder)
                            try:
                                items = os.listdir(appx_path)
                                pdfs_here = sum(
                                    1 for f in items
                                    if f.lower().endswith(".pdf")
                                    and os.path.isfile(os.path.join(appx_path, f))
                                )
                                subs_here = sum(
                                    1 for f in items
                                    if os.path.isdir(os.path.join(appx_path, f))
                                    and _tree_has_pdf(os.path.join(appx_path, f))
                                )
                            except Exception:
                                pdfs_here, subs_here = 0, 0

                            max_depth_here = _max_depth_below(appx_path)

                            app_decision = show_appendix_level_dialog(
                                self.root,
                                appendix_folder,
                                appx_idx,
                                len(appendix_folders),
                                pdfs_here,
                                subs_here,
                                max_depth_here,
                            )

                            if app_decision is None or app_decision.get("action") == "cancel":
                                cancelled = True
                                break

                            if app_decision["action"] == "show":
                                # Single rule applies to the whole appendix
                                appendix_decisions[appendix_folder] = app_decision
                                continue

                            # action == "go_folder": ask for each subfolder in
                            # this appendix subtree.
                            sub_folders_in_appx = []

                            def walk_subs(path, rel, depth):
                                try:
                                    items = sorted(os.listdir(path))
                                except Exception:
                                    return
                                for nm in items:
                                    full = os.path.join(path, nm)
                                    if os.path.isdir(full) and _tree_has_pdf(full):
                                        sub_rel = os.path.join(rel, nm) if rel else nm
                                        if depth >= 1:
                                            # Store base-relative (incl. appendix name)
                                            full_rel = (
                                                f"{appendix_folder} / "
                                                f"{sub_rel.replace(os.sep, ' / ')}"
                                            )
                                            sub_folders_in_appx.append((full, full_rel, nm))
                                        walk_subs(full, sub_rel, depth + 1)

                            walk_subs(appx_path, "", 0)

                            bulk_decision = None
                            for sub_idx, (sub_full, sub_rel_full, sub_name) in enumerate(
                                sub_folders_in_appx, start=1
                            ):
                                if bulk_decision is not None:
                                    per_folder_decisions[sub_rel_full] = bulk_decision
                                    continue

                                try:
                                    sub_items = os.listdir(sub_full)
                                    sub_pdfs = sum(
                                        1 for f in sub_items
                                        if f.lower().endswith(".pdf")
                                        and os.path.isfile(os.path.join(sub_full, f))
                                    )
                                    sub_subs = sum(
                                        1 for f in sub_items
                                        if os.path.isdir(os.path.join(sub_full, f))
                                        and _tree_has_pdf(os.path.join(sub_full, f))
                                    )
                                except Exception:
                                    sub_pdfs, sub_subs = 0, 0

                                sub_max_depth = _max_depth_below(sub_full)

                                decision, apply_all = show_manual_folder_dialog(
                                    self.root,
                                    sub_rel_full,
                                    sub_idx,
                                    len(sub_folders_in_appx),
                                    sub_pdfs,
                                    sub_subs,
                                    sub_max_depth,
                                )

                                if decision is None or decision.get("action") == "cancel":
                                    cancelled = True
                                    break

                                per_folder_decisions[sub_rel_full] = decision
                                if apply_all:
                                    bulk_decision = decision

                            if cancelled:
                                break

                        if cancelled:
                            return

                        policy = {
                            "mode": "manual",
                            "separators": sep_choice,
                            "decisions": per_folder_decisions,
                            "appendix_decisions": appendix_decisions,
                        }
                        break

            # ---------- Manual TITLES (existing description prompt) ----------
            manual_titles = {}
            if self.var_manual.get():
                items_for_prompt = []
                try:
                    base_items = sorted(os.listdir(base_dir))
                    for f in base_items:
                        full = os.path.join(base_dir, f)
                        if os.path.isfile(full) and f.lower().endswith(".pdf"):
                            items_for_prompt.append(("file", f))
                        elif os.path.isdir(full) and _tree_has_pdf(full):
                            items_for_prompt.append(("folder", f))
                except Exception:
                    pass

                from tkinter import simpledialog
                for kind, name in items_for_prompt:
                    default = os.path.splitext(name)[0] if kind == "file" else name
                    custom = simpledialog.askstring(
                        f"Description for {kind}",
                        f"Description for '{name}'\n(leave blank to use '{default}'):",
                        initialvalue=default,
                    )
                    if custom is None:
                        return
                    if custom.strip():
                        manual_titles[name] = custom.strip()
                    else:
                        manual_titles[name] = default

            self._save_current_settings()

            opts = {
                "base_directory": base_dir,
                "claim_title": claim_title,
                "numbering_style": self.var_numbering.get(),
                "manual_descriptions": self.var_manual.get(),
                "manual_titles": manual_titles,
                "generate_toc": self.var_toc.get(),
                "do_ocr": self.var_ocr.get(),
                "compression": self.var_compression.get(),
                "include_cover": self.var_cover.get(),
                "cover_subtitle": "",
                "cover_doc_type": self.cover_entries["cover_doc_type"].get().strip(),
                "style": {
                    "cover": self.var_style_cover.get(),
                    "palette": self.var_style_palette.get(),
                    "font": self.var_style_font.get(),
                    "separator": self.var_style_separator.get(),
                    "toc": self.var_style_toc.get(),
                },
                "cover_project": self.cover_entries["cover_project"].get().strip(),
                "cover_contract_ref": self.cover_entries["cover_contract_ref"].get().strip(),
                "cover_prepared_by": self.cover_entries["cover_prepared_by"].get().strip(),
                "cover_prepared_for": self.cover_entries["cover_prepared_for"].get().strip(),
                "cover_revision": self.cover_entries["cover_revision"].get().strip(),
                "cover_date": "",
                "policy": policy,
            }

            policy_summary = "Folder mode: " + policy["mode"]
            if policy["mode"] == "automatic":
                policy_summary += f" / {policy['layout']}"
            policy_summary += f" / separators: {policy['separators']}"

            confirm = (
                f"Base: {base_dir}\n"
                f"Title: {claim_title or '(default name)'}\n"
                f"Numbering: {opts['numbering_style']}\n"
                f"TOC: {'Yes' if opts['generate_toc'] else 'No'} | "
                f"OCR: {'Yes' if opts['do_ocr'] else 'No'} | "
                f"Compression: {opts['compression']}\n"
                f"Cover page: {'Yes' if opts['include_cover'] else 'No'}\n"
                f"{policy_summary}\n\n"
                "Proceed?"
            )
            if not messagebox.askyesno("Confirm build", confirm):
                return

            self.btn_run.configure(state="disabled")
            self._set_status("Starting bundle build...")

            t = threading.Thread(
                target=build_bundle,
                args=(opts, self._log_safe, self._on_done),
                daemon=True,
            )
            t.start()

        def _on_done(self, success, info):
            def _finish():
                self.btn_run.configure(state="normal")
                if success:
                    if messagebox.askyesno(
                        "Build complete",
                        f"Bundle saved successfully:\n\n{info}\n\nOpen the output folder?",
                    ):
                        try:
                            os.startfile(os.path.dirname(info))
                        except Exception:
                            pass
                else:
                    messagebox.showerror("Build failed", info)
            self.root.after(0, _finish)

    # -----------------------------------------------------------------------
    # Step 9: Main
    # -----------------------------------------------------------------------

    def main():
        root = tk.Tk()
        ClaimsBundleApp(root)
        root.mainloop()

    if __name__ == "__main__":
        main()

except SystemExit:
    raise
except Exception:
    err = traceback.format_exc()
    _fatal_error(
        "Claims Bundle Builder - Startup Error",
        f"The application failed to start.\n\nDetails:\n\n{err}"
    )
