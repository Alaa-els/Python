"""
Document Bundle Builder - engine
================================

GUI-free build engine used by ``document_bundle_builder.pyw`` and by the
regression tests. It merges a folder tree of PDFs (plus any PDFs produced
by the optional Office conversion stage) into one organised bundle with:

- Optional separator pages, controlled by ONE global scope:
    none            - no separator pages anywhere
    top_level       - top-level sections only (the classic "Appendix A" page)
    all_sections    - every section, top-level and nested folder headings
    every_document  - every section plus a page before every document
- A clickable, multi-page-safe Table of Contents whose entries and PDF
  outline bookmarks are independent of separators: with no separator the
  link and the bookmark target the first actual content page.
- Configurable section terminology (Appendix / Section / Drawing / ...)
- Several coherent bundle styles (cover, separator, contents, palette)
- Per-page footers that can be switched off or repositioned so they do
  not overprint drawing title blocks; footers are placed upright on
  rotated pages and original page sizes / vector content are preserved.
- Optional OCR, compression, cover page, custom order via _order.txt

The build is split into three phases:

1. ``plan_bundle``       walk the folder tree and apply the folder policy,
                         producing a tree of ``BundleNode`` objects
                         (sections and documents). Nothing is rendered.
2. ``assemble_body``     insert separators and documents in plan order and
                         record 1-based page numbers on every node.
3. ``finalise_bundle``   prepend cover and TOC, shift page numbers, add
                         links, bookmarks, footers, OCR, compression.

Only this module and PyMuPDF / fpdf2 are needed to build a bundle; the
GUI module wraps it and adds dialogs.
"""

import os
import re
import io
import copy
import json
import threading
import traceback
from collections import OrderedDict
from datetime import datetime

import fitz  # PyMuPDF
from fpdf import FPDF

ENGINE_VERSION = "6.1"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MM_TO_PT = 72 / 25.4
LEFT_MARGIN_MM = 20
RIGHT_MARGIN_MM = 20
TOP_MARGIN_MM = 20
BOTTOM_MARGIN_MM = 20

FOOTER_FONT_SIZE = 8
FOOTER_GREY = (0, 0, 0)
FOOTER_OPACITY = 0.8

TOC_TEXT_FONT = 12

# Settings live in the same file as the original Claims Bundle Builder so
# existing users keep their saved defaults.
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".claims_bundle_builder.json")
ORDER_FILE = "_order.txt"

# Folders the tree walk never descends into (tool-owned staging folders).
EXCLUDED_FOLDER_PREFIXES = ("_bundle_",)
STAGING_DIRNAME = "_bundle_converted"
# File-name marker used by the per-folder first-pages outputs (and by the
# owner's earlier scripts). Files carrying it are never used as inputs.
COMBINED_MARKER = "#Combined"
FIRST_PAGES_SUFFIX = "#Combined_FirstPages.pdf"
PDF_CREATOR_TAG = "Document Bundle Builder"

PAGE_SELECTION_MODES = ("all", "first", "first_n", "ranges")
PAGE_SELECTION_LABELS = {
    "all": "All pages",
    "first": "First page of each file (WIR preset)",
    "first_n": "First N pages of each file",
    "ranges": "Custom page ranges per file",
}

SEPARATOR_SCOPES = ("none", "top_level", "all_sections", "every_document")
SEPARATOR_SCOPE_LABELS = {
    "none": "No separator pages anywhere",
    "top_level": "Top-level sections only",
    "all_sections": "All sections (top-level and nested folders)",
    "every_document": "Every section and every document",
}
# Names used by the original folder-policy dialogs. "none" is NOT mapped:
# the old "none" still produced top-level separators; the new "none" means
# none anywhere, which is what the label always promised.
LEGACY_SEPARATOR_MAP = {"every_pdf": "every_document", "folder_only": "all_sections"}

NUMBERING_STYLES = ("alphabetic", "numeric", "none")

SECTION_TERMS = ("Appendix", "Section", "Drawing", "Document", "Part",
                 "Exhibit", "Schedule", "Tab")
TERM_SHORT = {"Appendix": "App", "Section": "Sec", "Drawing": "Dwg",
              "Document": "Doc", "Exhibit": "Exh", "Schedule": "Sch"}

TOC_DENSITIES = ("compact", "standard", "spacious")

FOOTER_POSITIONS = ("bottom", "top")
FOOTER_LAYOUTS = ("split", "left", "centre", "right")


class BundleCancelled(Exception):
    """Raised inside the build when the cancel event is set."""


# ---------------------------------------------------------------------------
# Style system
# ---------------------------------------------------------------------------

PALETTES = {
    # ink = all text and structural rules; accent = decorative rules only,
    # never text (accents fade in black-and-white copies)
    "navy_gold": {"name": "Navy + gold", "ink": (5, 22, 65), "accent": (255, 196, 37),
                  "muted": (90, 98, 114), "hair": (226, 226, 226),
                  "hair_strong": (201, 201, 201)},
    "navy_steel": {"name": "Navy + steel", "ink": (5, 22, 65), "accent": (122, 134, 153),
                   "muted": (90, 98, 114), "hair": (226, 226, 226),
                   "hair_strong": (201, 201, 201)},
    "mono": {"name": "Monochrome", "ink": (20, 20, 20), "accent": (90, 90, 90),
             "muted": (110, 110, 110), "hair": (222, 222, 222),
             "hair_strong": (190, 190, 190)},
    "graphite_orange": {"name": "Graphite + orange", "ink": (40, 44, 52),
                        "accent": (232, 118, 36), "muted": (105, 110, 120),
                        "hair": (224, 224, 224), "hair_strong": (196, 196, 196)},
    "teal_slate": {"name": "Teal + slate", "ink": (18, 62, 78), "accent": (0, 150, 136),
                   "muted": (96, 110, 118), "hair": (224, 228, 230),
                   "hair_strong": (196, 204, 208)},
    "burgundy_cream": {"name": "Burgundy + cream", "ink": (92, 26, 38),
                       "accent": (196, 160, 96), "muted": (110, 96, 96),
                       "hair": (230, 224, 220), "hair_strong": (206, 196, 190)},
}

COVER_STYLES = {
    "top_rule": "Top rule", "classical": "Classical centred",
    "minimal": "Minimal", "block": "Colour block", "register": "Register grid",
}
SEPARATOR_STYLES = {
    "top_rule": "Top rule", "plaque": "Plaque frame",
    "minimal": "Minimal", "band": "Title band", "block": "Colour block",
}
TOC_STYLES = {
    "ruled": "Ruled ledger", "dotted": "Dotted leaders",
    "plain": "Plain", "grid": "Register grid",
}
FONT_STYLES = {"aptos": "Aptos (falls back to Segoe UI / Helvetica)",
               "times": "Times New Roman"}

# Presets apply a coherent combination. They are a convenience layered on
# top of the individual pickers, which stay fully editable afterwards.
STYLE_PRESETS = {
    "claims_navy": {
        "name": "Claims bundle (navy + gold)",
        "style": {"cover": "top_rule", "palette": "navy_gold", "font": "aptos",
                  "separator": "top_rule", "toc": "ruled", "toc_density": "standard"},
        "terminology": {"section_term": "Appendix", "numbering_style": "alphabetic",
                        "cover_doc_type": "Claims Bundle"},
        "footer": {"footer_enabled": True, "footer_position": "bottom",
                   "footer_layout": "split"},
    },
    "minimal_mono": {
        "name": "Minimal monochrome",
        "style": {"cover": "minimal", "palette": "mono", "font": "aptos",
                  "separator": "minimal", "toc": "plain", "toc_density": "compact"},
        "terminology": {"section_term": "Section", "numbering_style": "numeric",
                        "cover_doc_type": "Document Bundle"},
        "footer": {"footer_enabled": True, "footer_position": "bottom",
                   "footer_layout": "split"},
    },
    "drawing_register": {
        "name": "Technical drawing register",
        "style": {"cover": "register", "palette": "graphite_orange", "font": "aptos",
                  "separator": "band", "toc": "grid", "toc_density": "compact"},
        "terminology": {"section_term": "Drawing", "numbering_style": "numeric",
                        "cover_doc_type": "Drawing Register"},
        # Drawings carry their own title blocks: footers off by default so
        # nothing is overprinted. Switch on and move to the top if needed.
        "footer": {"footer_enabled": False, "footer_position": "top",
                   "footer_layout": "left"},
    },
    "corporate_report": {
        "name": "Corporate report",
        "style": {"cover": "block", "palette": "teal_slate", "font": "aptos",
                  "separator": "block", "toc": "ruled", "toc_density": "spacious"},
        "terminology": {"section_term": "Section", "numbering_style": "numeric",
                        "cover_doc_type": "Report"},
        "footer": {"footer_enabled": True, "footer_position": "bottom",
                   "footer_layout": "split"},
    },
    "correspondence_pack": {
        "name": "Correspondence pack",
        "style": {"cover": "classical", "palette": "burgundy_cream", "font": "times",
                  "separator": "plaque", "toc": "dotted", "toc_density": "standard"},
        "terminology": {"section_term": "Document", "numbering_style": "numeric",
                        "cover_doc_type": "Correspondence Pack"},
        "footer": {"footer_enabled": True, "footer_position": "bottom",
                   "footer_layout": "split"},
    },
}

# Updated from the UI / opts before each build
CURRENT_STYLE = {
    "cover": "top_rule",
    "palette": "navy_gold",
    "font": "aptos",
    "separator": "top_rule",
    "toc": "ruled",
    "toc_density": "standard",
}

# Terminology in effect for the current build
CURRENT_TERMS = {"section_term": "Appendix", "numbering_style": "alphabetic"}


def style_palette():
    return PALETTES.get(CURRENT_STYLE.get("palette", "navy_gold"), PALETTES["navy_gold"])


def set_current_style(style):
    """Merge a style dict into CURRENT_STYLE, ignoring unknown values."""
    style = style or {}
    for key, table in (("cover", COVER_STYLES), ("separator", SEPARATOR_STYLES),
                       ("toc", TOC_STYLES), ("palette", PALETTES),
                       ("font", FONT_STYLES)):
        v = style.get(key)
        if v in table:
            CURRENT_STYLE[key] = v
    d = style.get("toc_density")
    if d in TOC_DENSITIES:
        CURRENT_STYLE["toc_density"] = d


def set_current_terms(section_term=None, numbering_style=None):
    if section_term:
        CURRENT_TERMS["section_term"] = str(section_term).strip() or "Appendix"
    if numbering_style in NUMBERING_STYLES:
        CURRENT_TERMS["numbering_style"] = numbering_style


def term_short(term=None):
    """Short form used in footers ("App A - P001 / 012")."""
    term = term or CURRENT_TERMS["section_term"]
    return TERM_SHORT.get(term, term if len(term) <= 5 else term[:4] + ".")


def section_display(label, title, term=None):
    """'Appendix A: Title' / 'Drawing 01: Title' / 'Title' when unnumbered."""
    term = term or CURRENT_TERMS["section_term"]
    if label:
        return f"{term} {label}: {title}"
    return title


_FONT_CACHE = {}


def _find_font_file(names):
    """Locate a font file in the Windows system or per-user font folders."""
    dirs = [
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
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
    if font_key in _FONT_CACHE:
        return _FONT_CACHE[font_key]
    if font_key == "times":
        paths = {"regular": _find_font_file(["times.ttf", "Times.ttf"]),
                 "bold": _find_font_file(["timesbd.ttf", "Timesbd.ttf"])}
    else:
        paths = {"regular": _find_font_file(["aptos.ttf", "Aptos.ttf", "aptos-regular.ttf",
                                             "Aptos-Regular.ttf"]),
                 "bold": _find_font_file(["aptos-bold.ttf", "Aptos-Bold.ttf", "aptosb.ttf"])}
        if not paths["regular"]:
            paths = {"regular": _find_font_file(["segoeui.ttf"]),
                     "bold": _find_font_file(["segoeuib.ttf"])}
    if paths.get("regular") and not paths.get("bold"):
        paths["bold"] = paths["regular"]
    _FONT_CACHE[font_key] = paths
    return paths


class StyledPDF(FPDF):
    """FPDF with the selected style font registered and text sanitised for
    the core fonts when no TTF is available."""

    def __init__(self, orientation="P", fmt="A4"):
        super().__init__(orientation=orientation, format=fmt)
        font_key = CURRENT_STYLE.get("font", "aptos")
        core = "Times" if font_key == "times" else "Helvetica"
        self.family = core
        self.core_font = True
        paths = _resolve_font_paths(font_key)
        if paths.get("regular"):
            try:
                try:
                    self.add_font("BundleFont", "", paths["regular"])
                    self.add_font("BundleFont", "B", paths["bold"])
                except TypeError:
                    self.add_font("BundleFont", "", paths["regular"], uni=True)
                    self.add_font("BundleFont", "B", paths["bold"], uni=True)
                self.family = "BundleFont"
                self.core_font = False
            except Exception:
                self.family = core
                self.core_font = True

    def t(self, text):
        """Text made safe for the active font."""
        text = sanitize_text(text)
        if self.core_font:
            text = text.encode("latin-1", "replace").decode("latin-1")
        return text

    def font(self, style="", size=10):
        self.set_font(self.family, style, size)


def new_styled_pdf():
    """Backwards-compatible helper returning (pdf, family)."""
    pdf = StyledPDF()
    return pdf, pdf.family


def _spaced(text):
    """Letter-space a kicker line."""
    return " ".join(text)


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def sanitize_text(text):
    """Normalise punctuation that core PDF fonts cannot render."""
    if not text:
        return ""
    repl = {
        "\u2019": "'", "\u2018": "'", "\u201C": '"', "\u201D": '"',
        "\u2013": "-", "\u2014": "-", "\u2022": "-", "\u2026": "...",
    }
    for o, n in repl.items():
        text = text.replace(o, n)
    return text


def convert_to_style(number, style):
    """Section index to label: numeric (01), alphabetic (A..Z, AA) or none."""
    if style == "none":
        return ""
    if style == "alphabetic":
        letters = ""
        n = number
        while n > 0:
            n, r = divmod(n - 1, 26)
            letters = chr(65 + r) + letters
        return letters
    return f"{number:02d}"


def mm_rect(x0_mm, y0_mm, x1_mm, y1_mm):
    return fitz.Rect(x0_mm * MM_TO_PT, y0_mm * MM_TO_PT, x1_mm * MM_TO_PT, y1_mm * MM_TO_PT)


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


def _lines(pdf, text, x, w, lh, align="L", colour=None):
    if colour:
        pdf.set_text_color(*colour)
    for line in _wrap_by_width(pdf, pdf.t(text), w):
        pdf.set_x(x)
        pdf.multi_cell(w, lh, line, 0, align)


# ---------------------------------------------------------------------------
# Separator pages
# ---------------------------------------------------------------------------

def create_separator_page(headline, subline="", kicker="", top_level=True):
    """Styled separator page.

    Top-level sections pass kicker='APPENDIX A' (the term + label), the
    section title as headline and nothing else. Nested folder headings and
    per-document pages pass the name as headline with an empty kicker.
    """
    pal = style_palette()
    pdf = StyledPDF()
    pdf.set_auto_page_break(False)
    pdf.set_margins(LEFT_MARGIN_MM, TOP_MARGIN_MM, RIGHT_MARGIN_MM)
    pdf.add_page()
    W = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM
    style = CURRENT_STYLE.get("separator", "top_rule")
    headline = pdf.t(headline)
    subline = pdf.t(subline)
    kicker = pdf.t(kicker).upper()

    if style == "plaque":
        head_size = 15 if top_level else 13
        frame_w, pad = 120.0, 10.0
        inner_w = frame_w - 2 * pad
        head_lh = head_size * 0.5
        k_lines, s_lines = [], []
        h = pad
        if kicker:
            pdf.font("B", 10)
            k_lines = _wrap_by_width(pdf, _spaced(kicker), inner_w)
            h += len(k_lines) * 6 + 3
        pdf.font("B", head_size)
        h_lines = _wrap_by_width(pdf, headline, inner_w)
        h += len(h_lines) * head_lh
        if subline:
            pdf.font("", 10)
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
            pdf.font("B", 10)
            pdf.set_text_color(*pal["ink"])
            for line in k_lines:
                pdf.set_x(x0 + pad)
                pdf.multi_cell(inner_w, 6, line, 0, "C")
            pdf.ln(3)
        pdf.font("B", head_size)
        pdf.set_text_color(*pal["ink"])
        for line in h_lines:
            pdf.set_x(x0 + pad)
            pdf.multi_cell(inner_w, head_lh, line, 0, "C")
        if s_lines:
            pdf.ln(4)
            pdf.font("", 10)
            pdf.set_text_color(*pal["muted"])
            for line in s_lines:
                pdf.set_x(x0 + pad)
                pdf.multi_cell(inner_w, 5.5, line, 0, "C")

    elif style == "minimal":
        head_size = 16 if top_level else 14
        pdf.set_y(118)
        if kicker:
            pdf.font("", 9)
            _lines(pdf, _spaced(kicker), LEFT_MARGIN_MM, W, 6, "C", pal["muted"])
            pdf.ln(4)
        pdf.font("B", head_size)
        _lines(pdf, headline, LEFT_MARGIN_MM, W, head_size * 0.5, "C", pal["ink"])
        if subline:
            pdf.ln(5)
            pdf.font("", 10)
            _lines(pdf, subline, LEFT_MARGIN_MM, W, 6, "C", pal["muted"])

    elif style == "band":
        # Full-width ink band carrying the kicker; headline below it.
        pdf.set_fill_color(*pal["ink"])
        pdf.rect(0, 0, 210, 34, "F")
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(0, 34, 210, 1.5, "F")
        pdf.set_y(12)
        pdf.font("B", 11)
        band_text = _spaced(kicker) if kicker else _spaced("SECTION" if not top_level else "")
        if band_text.strip():
            _lines(pdf, band_text, LEFT_MARGIN_MM, W, 7, "L", (255, 255, 255))
        pdf.set_y(60)
        head_size = 20 if top_level else 16
        pdf.font("B", head_size)
        _lines(pdf, headline, LEFT_MARGIN_MM, W, head_size * 0.5, "L", pal["ink"])
        if subline:
            pdf.ln(4)
            pdf.font("", 11)
            _lines(pdf, subline, LEFT_MARGIN_MM, W, 6.5, "L", pal["muted"])
        pdf.set_draw_color(*pal["hair_strong"])
        pdf.set_line_width(0.3)
        pdf.line(LEFT_MARGIN_MM, 297 - 25, 210 - RIGHT_MARGIN_MM, 297 - 25)

    elif style == "block":
        # Colour block over the top third; white text inside it.
        pdf.set_fill_color(*pal["ink"])
        pdf.rect(0, 0, 210, 105, "F")
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(LEFT_MARGIN_MM, 105, 22, 1.5, "F")
        pdf.set_y(48)
        if kicker:
            pdf.font("B", 10)
            _lines(pdf, _spaced(kicker), LEFT_MARGIN_MM, W, 6, "L", (255, 255, 255))
            pdf.ln(4)
        head_size = 22 if top_level else 17
        pdf.font("B", head_size)
        _lines(pdf, headline, LEFT_MARGIN_MM, W, head_size * 0.5, "L", (255, 255, 255))
        if subline:
            pdf.set_y(118)
            pdf.font("", 11)
            _lines(pdf, subline, LEFT_MARGIN_MM, W, 6.5, "L", pal["muted"])

    else:  # top_rule
        head_size = 18 if top_level else 15
        head_lh = head_size * 0.5
        pdf.set_fill_color(*pal["ink"])
        pdf.rect(0, 0, 210, 3, "F")
        pdf.set_y(95)
        if kicker:
            pdf.font("B", 11)
            _lines(pdf, _spaced(kicker), LEFT_MARGIN_MM, W, 7, "L", pal["ink"])
            pdf.ln(3)
        pdf.font("B", head_size)
        _lines(pdf, headline, LEFT_MARGIN_MM, W, head_lh, "L", pal["ink"])
        pdf.ln(3)
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(LEFT_MARGIN_MM, pdf.get_y(), 22, 1.2, "F")
        if subline:
            pdf.ln(7)
            pdf.font("", 11)
            _lines(pdf, subline, LEFT_MARGIN_MM, W, 6.5, "L", pal["muted"])
    return bytes(pdf.output())


def create_separator(title_main, title_sub, main_size=24, sub_size=14, spacing=40):
    """Original signature kept for compatibility. main_size >= 24 marks a
    top-level call of the form ('Appendix A', description)."""
    top_level = main_size >= 24
    if title_sub and top_level:
        return create_separator_page(title_sub, "", kicker=title_main, top_level=True)
    return create_separator_page(title_main, title_sub, top_level=top_level)


# ---------------------------------------------------------------------------
# Cover pages
# ---------------------------------------------------------------------------

def _cover_top_rule(pdf, pal, doc_type, title, subtitle, project, ref, rows, sections):
    X, W = 25.0, 160.0
    pdf.set_fill_color(*pal["ink"])
    pdf.rect(0, 0, 210, 4, "F")
    pdf.set_y(52)
    pdf.font("B", 10)
    _lines(pdf, _spaced(doc_type.upper()), X, W, 6, "L", pal["ink"])
    pdf.ln(8)
    pdf.font("B", 30)
    if len(_wrap_by_width(pdf, title, W)) > 2:
        size, lh = 24, 11
    else:
        size, lh = 30, 13
    pdf.font("B", size)
    _lines(pdf, title, X, W, lh, "L", pal["ink"])
    pdf.ln(4)
    pdf.set_fill_color(*pal["accent"])
    pdf.rect(X, pdf.get_y(), 22, 1.5, "F")
    pdf.ln(10)
    if project:
        pdf.font("B", 14)
        _lines(pdf, project, X, W, 7.5, "L", pal["ink"])
        pdf.ln(1.5)
    if ref:
        pdf.font("", 11)
        _lines(pdf, ref, X, W, 6, "L", pal["muted"])
    if subtitle:
        pdf.ln(4)
        pdf.font("", 12)
        _lines(pdf, subtitle, X, W, 6.5, "L", pal["ink"])
    _ledger(pdf, pal, rows, X, W, 232.0)


def _ledger(pdf, pal, rows, X, W, y, label_w=42.0):
    pdf.set_draw_color(*pal["hair_strong"])
    pdf.set_line_width(0.3)
    pdf.line(X, y - 4, X + W, y - 4)
    pdf.set_draw_color(*pal["hair"])
    pdf.set_line_width(0.2)
    for lbl, val in rows:
        pdf.set_xy(X, y)
        pdf.font("B", 8.5)
        pdf.set_text_color(*pal["ink"])
        pdf.cell(label_w, 6, pdf.t(lbl), 0, align="L")
        pdf.font("", 11)
        v_lines = _wrap_by_width(pdf, pdf.t(val), W - label_w)
        for k, vl in enumerate(v_lines):
            pdf.set_xy(X + label_w, y + k * 5.5)
            pdf.multi_cell(W - label_w, 5.5, vl, 0, "L")
        y += max(1, len(v_lines)) * 5.5 + 2
        pdf.line(X, y - 1, X + W, y - 1)
        y += 2


def _cover_classical(pdf, pal, doc_type, title, subtitle, project, ref, rows, sections):
    X, W = 25.0, 160.0
    pdf.set_y(78)
    pdf.font("B", 9)
    _lines(pdf, _spaced(doc_type.upper()), X, W, 6, "C", pal["ink"])
    pdf.ln(8)
    pdf.font("B", 26)
    if len(_wrap_by_width(pdf, title, W)) > 2:
        size, lh = 22, 10
    else:
        size, lh = 26, 12
    pdf.font("B", size)
    _lines(pdf, title, X, W, lh, "C", pal["ink"])
    pdf.ln(4)
    pdf.set_fill_color(*pal["ink"])
    pdf.rect(105 - 8, pdf.get_y(), 16, 0.8, "F")
    pdf.ln(8)
    if project:
        pdf.font("", 12)
        _lines(pdf, project, X, W, 6.5, "C", pal["ink"])
    if ref:
        pdf.ln(1)
        pdf.font("", 10)
        _lines(pdf, ref, X, W, 5.5, "C", pal["muted"])
    if subtitle:
        pdf.ln(3)
        pdf.font("", 11)
        _lines(pdf, subtitle, X, W, 6, "C", pal["ink"])
    y = 228.0
    for lbl, val in rows:
        pdf.set_xy(X, y)
        pdf.font("B", 8)
        pdf.set_text_color(*pal["ink"])
        pdf.multi_cell(W, 4.5, pdf.t(lbl), 0, "C")
        pdf.font("", 11)
        for vl in _wrap_by_width(pdf, pdf.t(val), W):
            pdf.set_x(X)
            pdf.multi_cell(W, 5.5, vl, 0, "C")
        y = pdf.get_y() + 2.5


def _cover_minimal(pdf, pal, doc_type, title, subtitle, project, ref, rows, sections):
    X, W = 25.0, 160.0
    pdf.set_y(96)
    pdf.font("", 9)
    _lines(pdf, _spaced(doc_type.upper()), X, W, 6, "L", pal["muted"])
    pdf.ln(6)
    pdf.font("B", 24)
    _lines(pdf, title, X, W, 11, "L", pal["ink"])
    if project:
        pdf.ln(4)
        pdf.font("", 13)
        _lines(pdf, project, X, W, 7, "L", pal["ink"])
    if ref:
        pdf.ln(1)
        pdf.font("", 10)
        _lines(pdf, ref, X, W, 6, "L", pal["muted"])
    if subtitle:
        pdf.ln(3)
        pdf.font("", 11)
        _lines(pdf, subtitle, X, W, 6, "L", pal["ink"])
    pdf.set_draw_color(*pal["hair_strong"])
    pdf.set_line_width(0.25)
    y = 236.0
    pdf.line(X, y - 4, X + W, y - 4)
    for lbl, val in rows:
        pdf.set_xy(X, y)
        pdf.font("", 8.5)
        pdf.set_text_color(*pal["muted"])
        pdf.cell(42, 6, pdf.t(lbl), 0, align="L")
        pdf.font("", 10.5)
        pdf.set_text_color(*pal["ink"])
        v_lines = _wrap_by_width(pdf, pdf.t(val), W - 42)
        for k, vl in enumerate(v_lines):
            pdf.set_xy(X + 42, y + k * 5.5)
            pdf.multi_cell(W - 42, 5.5, vl, 0, "L")
        y += max(1, len(v_lines)) * 5.5 + 2


def _cover_block(pdf, pal, doc_type, title, subtitle, project, ref, rows, sections):
    X, W = 25.0, 160.0
    pdf.set_fill_color(*pal["ink"])
    pdf.rect(0, 0, 210, 135, "F")
    pdf.set_fill_color(*pal["accent"])
    pdf.rect(X, 135, 30, 2, "F")
    pdf.set_y(48)
    pdf.font("B", 10)
    _lines(pdf, _spaced(doc_type.upper()), X, W, 6, "L", (255, 255, 255))
    pdf.ln(8)
    pdf.font("B", 28)
    if len(_wrap_by_width(pdf, title, W)) > 2:
        pdf.font("B", 22)
        lh = 10
    else:
        lh = 12.5
    _lines(pdf, title, X, W, lh, "L", (255, 255, 255))
    if subtitle:
        pdf.ln(4)
        pdf.font("", 12)
        _lines(pdf, subtitle, X, W, 6.5, "L", (235, 235, 235))
    pdf.set_y(150)
    if project:
        pdf.font("B", 14)
        _lines(pdf, project, X, W, 7.5, "L", pal["ink"])
        pdf.ln(1.5)
    if ref:
        pdf.font("", 11)
        _lines(pdf, ref, X, W, 6, "L", pal["muted"])
    _ledger(pdf, pal, rows, X, W, 232.0)


def _cover_register(pdf, pal, doc_type, title, subtitle, project, ref, rows, sections):
    """Title plus a boxed metadata grid and a short register of the
    top-level sections (labels and titles only; no page numbers)."""
    X, W = 20.0, 170.0
    pdf.set_fill_color(*pal["ink"])
    pdf.rect(0, 0, 210, 26, "F")
    pdf.set_fill_color(*pal["accent"])
    pdf.rect(0, 26, 210, 1.5, "F")
    pdf.set_y(9)
    pdf.font("B", 10)
    _lines(pdf, _spaced(doc_type.upper()), X, W, 7, "L", (255, 255, 255))
    pdf.set_y(42)
    pdf.font("B", 22)
    _lines(pdf, title, X, W, 10, "L", pal["ink"])
    if project:
        pdf.ln(2)
        pdf.font("", 13)
        _lines(pdf, project, X, W, 7, "L", pal["ink"])
    if ref:
        pdf.font("", 10)
        _lines(pdf, ref, X, W, 6, "L", pal["muted"])
    if subtitle:
        pdf.ln(2)
        pdf.font("", 11)
        _lines(pdf, subtitle, X, W, 6, "L", pal["ink"])
    # Metadata grid: two columns of boxed label/value cells
    y = max(pdf.get_y() + 8, 92)
    pdf.set_draw_color(*pal["hair_strong"])
    pdf.set_line_width(0.3)
    cell_w, cell_h = W / 2, 14
    for i, (lbl, val) in enumerate(rows):
        cx = X + (i % 2) * cell_w
        cy = y + (i // 2) * cell_h
        pdf.rect(cx, cy, cell_w, cell_h)
        pdf.set_xy(cx + 2, cy + 1.5)
        pdf.font("B", 7.5)
        pdf.set_text_color(*pal["muted"])
        pdf.cell(cell_w - 4, 4, pdf.t(lbl), 0, align="L")
        pdf.set_xy(cx + 2, cy + 6)
        pdf.font("", 10.5)
        pdf.set_text_color(*pal["ink"])
        v = _wrap_by_width(pdf, pdf.t(val), cell_w - 4)
        pdf.cell(cell_w - 4, 6, v[0] if v else "", 0, align="L")
    y += ((len(rows) + 1) // 2) * cell_h + 12
    # Register of top-level sections
    if sections:
        pdf.set_xy(X, y)
        pdf.font("B", 9)
        pdf.set_text_color(*pal["ink"])
        pdf.cell(W, 6, pdf.t(_spaced("CONTENTS")), 0, new_x="LMARGIN", new_y="NEXT")
        pdf.set_draw_color(*pal["ink"])
        pdf.set_line_width(0.4)
        y = pdf.get_y() + 0.5
        pdf.line(X, y, X + W, y)
        y += 1.5
        max_rows = int((297 - 22 - y) / 6.2)
        shown = sections[:max_rows]
        pdf.set_draw_color(*pal["hair"])
        pdf.set_line_width(0.2)
        for label, stitle in shown:
            pdf.set_xy(X, y)
            pdf.font("B", 9)
            pdf.set_text_color(*pal["ink"])
            pdf.cell(18, 6, pdf.t(label or ""), 0, align="L")
            pdf.font("", 9)
            t = _wrap_by_width(pdf, pdf.t(stitle), W - 18)
            pdf.cell(W - 18, 6, t[0] if t else "", 0, align="L")
            y += 6.2
            pdf.line(X, y - 0.4, X + W, y - 0.4)
        if len(sections) > len(shown):
            pdf.set_xy(X, y)
            pdf.font("", 8.5)
            pdf.set_text_color(*pal["muted"])
            pdf.cell(W, 6, pdf.t(f"... and {len(sections) - len(shown)} more (see Contents)"),
                     0, align="L")


COVER_RENDERERS = {
    "top_rule": _cover_top_rule, "classical": _cover_classical,
    "minimal": _cover_minimal, "block": _cover_block, "register": _cover_register,
}


def create_cover_page(meta):
    """Cover page in the selected style.

    meta keys: title, subtitle, doc_type, project, contract_ref, prepared_by,
    prepared_for, date, revision, sections (list of (label, title))."""
    pal = style_palette()
    pdf = StyledPDF()
    pdf.set_auto_page_break(False)
    pdf.set_margins(0, 0, 0)
    pdf.add_page()

    title = pdf.t(meta.get("title") or "Document Bundle")
    subtitle = pdf.t(meta.get("subtitle", ""))
    if subtitle.strip().lower() == title.strip().lower():
        subtitle = ""
    doc_type = pdf.t(meta.get("doc_type", "")).strip() or "Document Bundle"
    project = pdf.t(meta.get("project", ""))
    contract_ref = pdf.t(meta.get("contract_ref", ""))
    date_val = pdf.t(meta.get("date", "")).strip()
    if not date_val:
        date_val = datetime.now().strftime("%d %B %Y")
    rows = [
        ("PREPARED BY", pdf.t(meta.get("prepared_by", ""))),
        ("PREPARED FOR", pdf.t(meta.get("prepared_for", ""))),
        ("DATE", date_val),
        ("REVISION", pdf.t(meta.get("revision", ""))),
    ]
    rows = [(lbl, val) for lbl, val in rows if val.strip()]
    sections = meta.get("sections") or []
    renderer = COVER_RENDERERS.get(CURRENT_STYLE.get("cover"), _cover_top_rule)
    renderer(pdf, pal, doc_type, title, subtitle, project, contract_ref, rows, sections)
    return bytes(pdf.output())


# ---------------------------------------------------------------------------
# Table of contents
# ---------------------------------------------------------------------------

def _toc_metrics():
    density = CURRENT_STYLE.get("toc_density", "standard")
    if density == "compact":
        return {"line_h": 5.0, "gap": 0.6, "top_gap": 0.4, "font": TOC_TEXT_FONT - 1.5,
                "indent": 5, "title_gap": 3}
    if density == "spacious":
        return {"line_h": 7.5, "gap": 2.6, "top_gap": 2.0, "font": TOC_TEXT_FONT,
                "indent": 7, "title_gap": 6}
    return {"line_h": 6.0, "gap": 1.5, "top_gap": 1.0, "font": TOC_TEXT_FONT,
            "indent": 6, "title_gap": 4}


def create_toc_page_hier(claim_title, toc_entries, max_level=5, heading="Contents"):
    """Generate the TOC PDF supporting up to `max_level` indent levels.

    toc_entries: flat list of dicts in render order:
        {"level": 1..max_level, "title": str, "label": str|None,
         "page": 1-based target page or None, "show_page": bool,
         "range_text": str|None}

    Returns (pdf_bytes, link_positions, page_count). link_positions is a
    list of {'page_in_toc', 'y0_mm', 'y1_mm', 'target_zero'}.
    """
    pal = style_palette()
    toc_style = CURRENT_STYLE.get("toc", "ruled")
    m = _toc_metrics()
    pdf = StyledPDF()
    pdf.set_auto_page_break(auto=True, margin=BOTTOM_MARGIN_MM)
    pdf.set_margins(LEFT_MARGIN_MM, TOP_MARGIN_MM, RIGHT_MARGIN_MM)
    pdf.add_page()

    total_w = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM
    page_box_w = 18
    line_h = m["line_h"]
    per_level_indent = m["indent"]
    grid = toc_style == "grid"
    label_col_w = 16.0 if grid else 0.0

    # Header
    pdf.font("B", 22)
    pdf.set_text_color(*pal["ink"])
    pdf.set_x(LEFT_MARGIN_MM)
    pdf.cell(0, 12, pdf.t(heading), 0, new_x="LMARGIN", new_y="NEXT", align="L")
    if toc_style != "plain":
        pdf.set_fill_color(*pal["accent"])
        pdf.rect(LEFT_MARGIN_MM, pdf.get_y() + 1, 22, 1.5, "F")
    pdf.ln(6)
    if claim_title:
        pdf.font("", 11)
        pdf.set_text_color(*pal["muted"])
        pdf.set_x(LEFT_MARGIN_MM)
        pdf.cell(0, 7, pdf.t(claim_title), 0, new_x="LMARGIN", new_y="NEXT", align="L")
    pdf.ln(m["title_gap"])

    def grid_header():
        pdf.font("B", 8.5)
        pdf.set_text_color(*pal["muted"])
        pdf.set_x(LEFT_MARGIN_MM)
        pdf.cell(label_col_w, line_h, pdf.t("No."), 0, align="L")
        pdf.cell(total_w - label_col_w - page_box_w, line_h, pdf.t("Title"), 0, align="L")
        pdf.cell(page_box_w, line_h, pdf.t("Page"), 0, new_x="LMARGIN", new_y="NEXT", align="R")
        y = pdf.get_y()
        pdf.set_draw_color(*pal["ink"])
        pdf.set_line_width(0.4)
        pdf.line(LEFT_MARGIN_MM, y, 210 - RIGHT_MARGIN_MM, y)
        pdf.ln(0.8)

    if grid:
        grid_header()
    link_pos = []

    def style_for_level(lvl):
        base = m["font"]
        if lvl == 1:
            return ("B", base)
        if lvl == 2:
            return ("B", base - 1)
        return ("", base - 1)

    for entry in toc_entries:
        lvl = max(1, min(entry.get("level", 1), max_level))
        title = pdf.t(entry.get("title", ""))
        label = entry.get("label")
        range_text = entry.get("range_text")
        page_no = entry.get("page")
        show_page = entry.get("show_page", True) and (page_no is not None)

        indent_mm = (lvl - 1) * per_level_indent
        font_style, font_size = style_for_level(lvl)
        pdf.font(font_style, font_size)
        pdf.set_text_color(*(pal["ink"] if lvl <= 2 else pal["muted"]))

        chip_label, chip_w = None, 0.0
        if toc_style == "ruled" and lvl == 1 and label:
            chip_label, chip_w = label, 9.0 if len(label) <= 2 else 12.0
        elif lvl == 1 and label and not grid:
            title = f"{CURRENT_TERMS['section_term']} {label}: {title}"
        if range_text and lvl == 1 and not grid:
            title = f"{title}  {range_text}"

        # Grid style: label column, then title column (indented per level)
        x_text = LEFT_MARGIN_MM + label_col_w + indent_mm + chip_w
        available_w = total_w - label_col_w - indent_mm
        max_text_w = available_w - chip_w - (page_box_w + 2 if show_page else 0)
        lines = _wrap_by_width(pdf, title, max_text_w)

        # Keep an entry on one page when it is short enough to fit.
        needed = len(lines) * line_h + m["gap"]
        if pdf.get_y() + needed > 297 - BOTTOM_MARGIN_MM and needed < 40:
            pdf.add_page()
            if grid:
                grid_header()

        page_in_toc_top = pdf.page_no() - 1
        y_top = pdf.get_y()

        if grid and lvl == 1:
            pdf.set_xy(LEFT_MARGIN_MM, y_top)
            pdf.font("B", font_size)
            pdf.cell(label_col_w, line_h, pdf.t(label or ""), 0, align="L")
            pdf.font(font_style, font_size)

        if chip_label:
            chip_box = chip_w - 2
            pdf.set_fill_color(*pal["ink"])
            pdf.rect(LEFT_MARGIN_MM + indent_mm, y_top + 0.4, chip_box, line_h - 0.8, "F")
            pdf.font("B", font_size - 2)
            pdf.set_text_color(255, 255, 255)
            pdf.set_xy(LEFT_MARGIN_MM + indent_mm, y_top)
            pdf.cell(chip_box, line_h, pdf.t(chip_label), 0, align="C")
            pdf.font(font_style, font_size)
            pdf.set_text_color(*pal["ink"])

        for i, line in enumerate(lines):
            last = (i == len(lines) - 1)
            pdf.set_x(x_text)
            if last and show_page:
                text_w = pdf.get_string_width(line)
                fill_w = max(available_w - chip_w - text_w - page_box_w - 2, 0)
                pdf.cell(text_w + 1, line_h, line)
                if toc_style == "dotted":
                    pdf.set_text_color(*pal["hair_strong"])
                    pdf.cell(fill_w, line_h, _dot_fill(pdf, fill_w))
                    pdf.set_text_color(*(pal["ink"] if lvl <= 2 else pal["muted"]))
                else:
                    pdf.cell(fill_w, line_h, "")
                pdf.cell(page_box_w, line_h, str(page_no), new_x="LMARGIN", new_y="NEXT", align="R")
            else:
                pdf.cell(0, line_h, line, new_x="LMARGIN", new_y="NEXT")

        if toc_style in ("ruled", "grid"):
            rule_y = pdf.get_y() + 0.4
            if rule_y < 297 - BOTTOM_MARGIN_MM:
                if lvl == 1 and toc_style == "ruled":
                    pdf.set_draw_color(*pal["ink"])
                    pdf.set_line_width(0.35)
                else:
                    pdf.set_draw_color(*pal["hair"])
                    pdf.set_line_width(0.2)
                x_rule = LEFT_MARGIN_MM if grid else LEFT_MARGIN_MM + indent_mm
                pdf.line(x_rule, rule_y, 210 - RIGHT_MARGIN_MM, rule_y)
            pdf.ln(m["gap"])
        else:
            pdf.ln(m["gap"] * 0.6)

        page_in_toc_bot = pdf.page_no() - 1
        y_bottom = pdf.get_y()

        if page_no is not None:
            target_zero = page_no - 1
            if page_in_toc_bot == page_in_toc_top:
                link_pos.append({"page_in_toc": page_in_toc_top, "y0_mm": y_top,
                                 "y1_mm": y_bottom, "target_zero": target_zero})
            else:
                link_pos.append({"page_in_toc": page_in_toc_top, "y0_mm": y_top,
                                 "y1_mm": 297 - BOTTOM_MARGIN_MM, "target_zero": target_zero})
                link_pos.append({"page_in_toc": page_in_toc_bot, "y0_mm": TOP_MARGIN_MM,
                                 "y1_mm": y_bottom, "target_zero": target_zero})

        if lvl == 1:
            pdf.ln(m["top_gap"])

    toc_page_count = pdf.page_no()
    return bytes(pdf.output()), link_pos, toc_page_count


# ---------------------------------------------------------------------------
# PDF insertion / page helpers
# ---------------------------------------------------------------------------

def inspect_pdf(path):
    """Preflight a PDF without loading page content.

    Returns {"pages": int|None, "reason": str|None, "previous_output": bool}.
    reason is set when the file must be skipped (encrypted, corrupt, empty,
    or a bundle produced earlier by this tool)."""
    info = {"pages": None, "reason": None, "previous_output": False}
    try:
        with fitz.open(path) as d:
            if d.needs_pass:
                info["reason"] = "password-protected (encrypted) PDF"
                return info
            if d.page_count == 0:
                info["reason"] = "PDF has no pages"
                return info
            try:
                meta = d.metadata or {}
                if (meta.get("creator") or "") == PDF_CREATOR_TAG:
                    info["previous_output"] = True
                    info["reason"] = "output of a previous bundle run"
            except Exception:
                pass
            info["pages"] = d.page_count
    except Exception as e:
        info["reason"] = f"cannot be opened ({str(e)[:80]})"
    return info


def count_pdf_pages(path):
    """Page count of a usable PDF, or None if it cannot be bundled."""
    info = inspect_pdf(path)
    return None if info["reason"] else info["pages"]


def open_source_pdf(path):
    """Open a source PDF for page-level insertion, repairing if needed.
    Returns a fitz.Document (caller closes) or raises."""
    try:
        d = fitz.open(path)
        _ = d.page_count
        return d
    except Exception:
        with fitz.open(path) as d:
            data = d.tobytes(garbage=1)
        return fitz.open(stream=data, filetype="pdf")


def parse_page_ranges(spec, page_count):
    """'1-3, 5, 9-' -> sorted unique 0-based page indices.

    Raises ValueError with a readable message for anything invalid or out of
    range. An empty spec means all pages."""
    spec = (spec or "").strip()
    if not spec:
        return list(range(page_count))
    out = set()
    for piece in spec.split(","):
        piece = piece.strip()
        if not piece:
            continue
        m = re.fullmatch(r"(\d+)?\s*(-)?\s*(\d+)?", piece)
        if not m or (m.group(1) is None and m.group(3) is None):
            raise ValueError(f"'{piece}' is not a page or range")
        if m.group(2) is None:
            a = b = int(m.group(1))
        else:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(3)) if m.group(3) else page_count
        if a < 1 or b < 1:
            raise ValueError(f"'{piece}': pages start at 1")
        if a > b:
            raise ValueError(f"'{piece}': start is after end")
        if b > page_count:
            raise ValueError(f"'{piece}': beyond the last page ({page_count})")
        out.update(range(a - 1, b))
    if not out:
        raise ValueError("no pages selected")
    return sorted(out)


def insert_pdf_clean(merger, src_path):
    """Three-tier insert with fallbacks. Returns pages inserted.

    insert_pdf keeps each page's MediaBox/CropBox, /Rotate entry and all
    vector content untouched, so A0/A1/A3 drawings keep their size,
    orientation and line quality."""
    try:
        with fitz.open(src_path) as d:
            merger.insert_pdf(d)
            return d.page_count
    except Exception:
        pass
    try:
        with fitz.open(src_path) as d:
            data = d.tobytes(garbage=1)
        with fitz.open(stream=data, filetype="pdf") as fixed:
            merger.insert_pdf(fixed)
            return fixed.page_count
    except Exception:
        pass
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


def _insert_bytes(merger, pdf_bytes):
    with fitz.open(stream=pdf_bytes, filetype="pdf") as d:
        merger.insert_pdf(d)
        return d.page_count


def draw_footer(doc, page_idx, left_text, right_text, position="bottom",
                layout="split", inset_mm=6.35, font_size=FOOTER_FONT_SIZE):
    """Draw a footer (or header) line on one page.

    Coordinates are computed in the page's displayed (rotated) rectangle and
    mapped back through the derotation matrix, so text is upright and sits
    along the displayed bottom (or top) edge on rotated drawings too.
    """
    page = doc.load_page(page_idx)
    rect = page.rect
    inset = inset_mm * MM_TO_PT
    side = 20.0
    if position == "top":
        y = rect.y0 + inset + font_size
    else:
        y = rect.y1 - inset
    rot = page.rotation
    mat = page.derotation_matrix
    band = fitz.Rect(rect.x0 + side, y - font_size - 3, rect.x1 - side, y + 3)
    kwargs = dict(fontsize=font_size, color=FOOTER_GREY, fill_opacity=FOOTER_OPACITY,
                  overlay=True, rotate=rot)
    if layout == "split":
        page.insert_text(fitz.Point(rect.x0 + side, y) * mat, left_text, **kwargs)
        page.insert_textbox(band * mat, right_text, align=fitz.TEXT_ALIGN_RIGHT, **kwargs)
        return
    combined = f"{left_text}   |   {right_text}" if (left_text and right_text) \
        else (left_text or right_text)
    align = {"left": fitz.TEXT_ALIGN_LEFT, "centre": fitz.TEXT_ALIGN_CENTER,
             "right": fitz.TEXT_ALIGN_RIGHT}.get(layout, fitz.TEXT_ALIGN_LEFT)
    page.insert_textbox(band * mat, combined, align=align, **kwargs)


def draw_footer_numbering(doc, page_idx, left_text, right_text):
    """Original signature kept for compatibility."""
    draw_footer(doc, page_idx, left_text, right_text)


def ocr_page_if_needed(page, dpi=200, lang_hint="eng+ara"):
    try:
        from PIL import Image
        import pytesseract
    except Exception:
        return False
    try:
        if page.get_text("text").strip():
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
        page.insert_textbox(box, txt, fontsize=6, color=(1, 1, 1),
                            fill_opacity=0.01, overlay=False)
        return True
    except Exception:
        return False


def ocr_available():
    try:
        import PIL  # noqa: F401
        import pytesseract  # noqa: F401
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
            if list(page.annots() or []):
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def compress_document(doc, mode="high", log_fn=None):
    """Recompress raster images only. Vector content and page geometry are
    never touched, so drawings keep their line quality in every mode."""
    if mode == "high":
        return
    if mode == "balanced":
        max_edge_target, jpeg_quality = 2400, 80
    else:
        max_edge_target, jpeg_quality = 1500, 60
    try:
        from PIL import Image as PILImage
    except Exception:
        if log_fn:
            log_fn("  Skipping compression (Pillow not available).")
        return
    compressed = 0
    for page_num in range(doc.page_count):
        try:
            page = doc.load_page(page_num)
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                try:
                    img = doc.extract_image(xref)
                    imgdata = img["image"]
                    pil = PILImage.open(io.BytesIO(imgdata))
                    max_edge = max(pil.size)
                    scale = min(1.0, max_edge_target / max_edge) if max_edge > 0 else 1.0
                    if scale < 1.0:
                        pil = pil.resize((int(pil.size[0] * scale), int(pil.size[1] * scale)),
                                         PILImage.LANCZOS)
                    buff = io.BytesIO()
                    pil.convert("RGB").save(buff, format="JPEG", optimize=True, quality=jpeg_quality)
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


# ---------------------------------------------------------------------------
# Folder discovery
# ---------------------------------------------------------------------------

def norm_path(path):
    return os.path.normcase(os.path.abspath(path))


def is_excluded_folder(name):
    return name.startswith(EXCLUDED_FOLDER_PREFIXES)


def _read_order_file(folder_path):
    order_path = os.path.join(folder_path, ORDER_FILE)
    if not os.path.isfile(order_path):
        return None
    try:
        with open(order_path, "r", encoding="utf-8") as f:
            return [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    except Exception:
        return None


_NAT_RE = re.compile(r"(\d+)")


def natural_key(name):
    """Deterministic natural sort key: case-insensitive, digit runs compared
    numerically ("Dwg 2" before "Dwg 10"). Ties broken by the raw string so
    the order is total."""
    parts = _NAT_RE.split(name)
    key = []
    for i, part in enumerate(parts):
        if i % 2 == 1:
            key.append((0, int(part), part))
        else:
            key.append((1, 0, part.casefold()))
    return (key, name)


def _apply_custom_order(items, custom_order):
    """Listed names first (in _order.txt order), then the rest in natural
    file-name order."""
    if not custom_order:
        return sorted(items, key=natural_key)
    seen, ordered = set(), []
    for name in custom_order:
        if name in items and name not in seen:
            ordered.append(name)
            seen.add(name)
    return ordered + sorted((x for x in items if x not in seen), key=natural_key)


def is_previous_output_name(name):
    return COMBINED_MARKER in name


def list_folder_items(folder_path, overlay=None, exclude_files=()):
    """Return (documents, subfolders) for one folder.

    documents: list of (order_name, display_title, pdf_path) in bundle order.
      Real PDFs are listed by file name; Office documents that appear in
      ``overlay`` (source path -> converted PDF path) are listed under their
      original name so _order.txt and alphabetical ordering work unchanged.
    subfolders: folder names (in bundle order) whose subtree has content.
    """
    overlay = overlay or {}
    excluded = {norm_path(p) for p in exclude_files}
    try:
        items = os.listdir(folder_path)
    except Exception:
        return [], []
    docs = {}
    subs = []
    for name in items:
        full = os.path.join(folder_path, name)
        if os.path.isfile(full):
            if name.startswith("~$"):
                continue
            if name.lower().endswith(".pdf"):
                if norm_path(full) in excluded or is_previous_output_name(name):
                    continue
                docs[name] = full
            else:
                conv = overlay.get(norm_path(full))
                if conv:
                    docs[name] = conv
        elif os.path.isdir(full):
            if is_excluded_folder(name):
                continue
            if tree_has_content(full, overlay):
                subs.append(name)
    custom_order = _read_order_file(folder_path)
    doc_names = _apply_custom_order(list(docs), custom_order)
    sub_names = _apply_custom_order(subs, custom_order)
    documents = [(n, os.path.splitext(n)[0], docs[n]) for n in doc_names]
    return documents, sub_names


def tree_has_content(folder_path, overlay=None):
    """True if the folder or any subfolder holds a PDF or an overlay
    document. Empty folders never become sections."""
    overlay = overlay or {}
    try:
        for cur, dirs, files in os.walk(folder_path):
            dirs[:] = [d for d in dirs if not is_excluded_folder(d)]
            for f in files:
                if f.startswith("~$") or is_previous_output_name(f):
                    continue
                if f.lower().endswith(".pdf"):
                    return True
                if overlay and norm_path(os.path.join(cur, f)) in overlay:
                    return True
    except Exception:
        pass
    return False


def _tree_has_pdf(folder_path):
    return tree_has_content(folder_path)


def count_folder_tree(base_directory, overlay=None):
    """Counts for the discovery popup: {top_pdfs, top_folders,
    total_subfolders, total_pdfs}."""
    counts = {"top_pdfs": 0, "top_folders": 0, "total_subfolders": 0, "total_pdfs": 0}
    docs, subs = list_folder_items(base_directory, overlay)
    counts["top_pdfs"] = len(docs)
    counts["total_pdfs"] = len(docs)
    counts["top_folders"] = len(subs)

    def walk(path):
        d, s = list_folder_items(path, overlay)
        counts["total_pdfs"] += len(d)
        counts["total_subfolders"] += len(s)
        for name in s:
            walk(os.path.join(path, name))

    for name in subs:
        walk(os.path.join(base_directory, name))
    return counts


def _count_folder_tree(base_directory):
    return count_folder_tree(base_directory)


def max_depth_below(folder_path, overlay=None):
    """Maximum depth of content-bearing subfolders below folder_path."""
    max_d = 0

    def walk(path, d):
        nonlocal max_d
        _, subs = list_folder_items(path, overlay)
        for name in subs:
            max_d = max(max_d, d + 1)
            walk(os.path.join(path, name), d + 1)

    walk(folder_path, 0)
    return max_d


def _max_depth_below(folder_path):
    return max_depth_below(folder_path)


def folder_summary(path, overlay=None):
    """(direct document count, immediate content subfolder count)."""
    docs, subs = list_folder_items(path, overlay)
    return len(docs), len(subs)


def _is_numeric_name(name):
    return name.isdigit()


# ---------------------------------------------------------------------------
# Folder policy
# ---------------------------------------------------------------------------
# policy = {
#   "mode": "automatic" | "manual",
#   "layout": "flatten" | "preserve" | "preserve_skip_numeric"   (automatic)
#   "decisions": {"Top / Sub / Sub": {"action": "show", "depth": N} | {"action": "flatten"} | {"action": "skip"}}
#   "appendix_decisions": {"Top": {"action": "show", "depth": N} | {"action": "go_folder"}}
#   "separators": legacy per-run value (optional; opts["separator_scope"] wins)
# }

def resolve_folder_decision(folder_rel, folder_name, policy, current_depth_remaining):
    if policy.get("mode") == "manual":
        decisions = policy.get("decisions", {})
        if folder_rel in decisions:
            d = decisions[folder_rel]
            if isinstance(d, dict):
                return d
            if d == "flatten":
                return {"action": "flatten"}
            if d == "skip":
                return {"action": "skip"}
            if d == "heading_collapsed":
                return {"action": "show", "depth": 1}
            return {"action": "show", "depth": 999}
        if current_depth_remaining is not None:
            return {"action": "show", "depth": current_depth_remaining}
        return {"action": "show", "depth": 999}
    layout = policy.get("layout", "preserve")
    if layout == "flatten":
        return {"action": "flatten"}
    if layout == "preserve_skip_numeric" and _is_numeric_name(folder_name):
        return {"action": "flatten"}
    return {"action": "show", "depth": 999}


_resolve_folder_decision = resolve_folder_decision


def normalise_separator_scope(value, fallback="all_sections"):
    if value in SEPARATOR_SCOPES:
        return value
    if value in LEGACY_SEPARATOR_MAP:
        return LEGACY_SEPARATOR_MAP[value]
    return fallback


# ---------------------------------------------------------------------------
# Bundle plan
# ---------------------------------------------------------------------------

class BundleNode:
    """A section (folder / top-level entry) or a document in the plan."""

    __slots__ = ("kind", "title", "level", "path", "source_name", "label", "children",
                 "toc", "bookmark", "separator", "root_document", "collapsed_child",
                 "start", "end", "sep_page", "content_start", "pages", "expected_pages",
                 "policy_key", "failed", "node_id", "page_refs")

    def __init__(self, kind, title, level, path=None, source_name=None):
        self.kind = kind
        self.title = title
        self.level = level
        self.path = path
        self.source_name = source_name or (os.path.basename(path) if path else title)
        self.label = ""
        self.children = []
        self.toc = True
        self.bookmark = True
        self.separator = False
        self.root_document = False
        self.collapsed_child = False
        self.start = None
        self.end = None
        self.sep_page = None
        self.content_start = None
        self.pages = 0
        self.expected_pages = None
        self.policy_key = None
        self.failed = False
        self.node_id = None
        # Page manifest for documents: list of {"src": int, "rot": int, "inc": bool}
        # in bundle order. None means "all pages, as-is".
        self.page_refs = None

    def included_pages(self):
        """Page refs that will be bundled, in order."""
        if self.page_refs is None:
            n = self.expected_pages or 0
            return [{"src": i, "rot": 0, "inc": True} for i in range(n)]
        return [r for r in self.page_refs if r.get("inc", True)]

    def walk(self):
        yield self
        for c in self.children:
            yield from c.walk()

    def documents(self):
        return [n for n in self.walk() if n.kind == "document"]

    def display(self, term=None):
        if self.kind == "section" and self.level == 1:
            return section_display(self.label, self.title, term)
        return self.title

    def as_dict(self):
        return {
            "kind": self.kind, "title": self.title, "level": self.level, "label": self.label,
            "path": self.path, "toc": self.toc, "bookmark": self.bookmark,
            "separator": self.separator, "start": self.start, "end": self.end,
            "sep_page": self.sep_page, "content_start": self.content_start,
            "node_id": self.node_id, "source_name": self.source_name,
            "expected_pages": self.expected_pages, "page_refs": self.page_refs,
            "children": [c.as_dict() for c in self.children],
        }


class BundlePlan:
    def __init__(self):
        self.sections = []
        self.warnings = []
        self.skipped = []
        self.skipped_documents = []   # [(file name, reason)] - problems worth flagging
        self.previous_outputs = []    # file names recognised as earlier bundle outputs
        self.selection = {"mode": "all"}
        self.numbering = "alphabetic"

    def walk(self):
        for s in self.sections:
            yield from s.walk()

    def documents(self):
        return [n for n in self.walk() if n.kind == "document"]

    def summary_lines(self, term=None):
        out = []
        for n in self.walk():
            pad = "  " * (n.level - 1)
            flag = "S" if n.separator else "-"
            out.append(f"{pad}[{flag}] {n.kind}: {n.display(term)}")
        return out


def _plan_folder(folder_path, parent, policy, top_name, rel_prefix, level,
                 inherited_depth, overlay, exclude_files, plan, log_fn, collapsed=False):
    """Populate parent.children from folder_path according to policy."""
    documents, subfolders = list_folder_items(folder_path, overlay, exclude_files)

    for order_name, title, pdf_path in documents:
        node = BundleNode("document", title, level, path=pdf_path, source_name=order_name)
        node.collapsed_child = collapsed
        if collapsed:
            node.toc = False
        parent.children.append(node)

    for sub in subfolders:
        sub_path = os.path.join(folder_path, sub)
        sub_rel = os.path.join(rel_prefix, sub) if rel_prefix else sub
        sub_rel_display = sub_rel.replace(os.sep, " / ")
        policy_key = f"{top_name} / {sub_rel_display}" if top_name else sub_rel_display

        if collapsed:
            node = BundleNode("section", sub, level)
            node.toc = False
            node.collapsed_child = True
            node.policy_key = policy_key
            _plan_folder(sub_path, node, policy, top_name, sub_rel, level + 1, None,
                         overlay, exclude_files, plan, log_fn, collapsed=True)
            if node.documents():
                parent.children.append(node)
            continue

        decision = resolve_folder_decision(policy_key, sub, policy, inherited_depth)
        action = decision.get("action", "show")
        if action == "skip":
            plan.skipped.append(policy_key)
            if log_fn:
                log_fn(f"  [skip] {policy_key}")
            continue
        if action == "flatten":
            if log_fn:
                log_fn(f"  [flatten] {policy_key}")
            _plan_folder(sub_path, parent, policy, top_name, sub_rel, level, inherited_depth,
                         overlay, exclude_files, plan, log_fn)
            continue

        depth_here = decision.get("depth", 999)
        node = BundleNode("section", sub, level)
        node.policy_key = policy_key
        if depth_here <= 1:
            if log_fn:
                log_fn(f"  [{level}-collapsed] {policy_key}")
            _plan_folder(sub_path, node, policy, top_name, sub_rel, level + 1, None,
                         overlay, exclude_files, plan, log_fn, collapsed=True)
        else:
            next_depth = 999 if depth_here >= 999 else depth_here - 1
            if log_fn:
                log_fn(f"  [{level}-heading depth={depth_here}] {policy_key}")
            _plan_folder(sub_path, node, policy, top_name, sub_rel, level + 1, next_depth,
                         overlay, exclude_files, plan, log_fn)
        if node.documents():
            parent.children.append(node)
        else:
            plan.warnings.append(f"Folder '{policy_key}' has no documents after applying "
                                 f"the folder policy - no section created.")


def plan_bundle(base_directory, policy, opts=None, overlay=None, log_fn=None):
    """Build the bundle plan. See module docstring.

    opts keys used: numbering_style, section_term, manual_titles,
    manual_descriptions, separator_scope, toc_include_documents,
    bookmark_include_documents, output_path (excluded from inputs).
    """
    opts = opts or {}
    policy = policy or {"mode": "automatic", "layout": "preserve"}
    plan = BundlePlan()
    overlay = overlay or {}
    exclude_files = [opts["output_path"]] if opts.get("output_path") else []

    numbering = opts.get("numbering_style") or "alphabetic"
    if numbering not in NUMBERING_STYLES:
        numbering = "alphabetic"
    plan.numbering = numbering
    is_manual = bool(opts.get("manual_descriptions"))
    manual_titles = opts.get("manual_titles") or {}
    appendix_decisions = policy.get("appendix_decisions") or {}

    documents, folders = list_folder_items(base_directory, overlay, exclude_files)
    if opts.get("direct_files_only"):
        folders = []

    # Top-level documents: each becomes its own numbered section
    for order_name, title, pdf_path in documents:
        description = manual_titles.get(order_name, title) if is_manual else title
        sec = BundleNode("section", description, 1)
        sec.root_document = True
        doc = BundleNode("document", title, 2, path=pdf_path, source_name=order_name)
        doc.toc = False  # the section row already represents this document
        sec.children.append(doc)
        plan.sections.append(sec)

    # Top-level folders
    for folder in folders:
        folder_path = os.path.join(base_directory, folder)
        description = manual_titles.get(folder, folder) if is_manual else folder
        sec = BundleNode("section", description, 1)
        sec.policy_key = folder
        app_decision = appendix_decisions.get(folder)
        inherited_depth = None
        collapsed = False
        if app_decision and app_decision.get("action") == "show":
            depth_at = app_decision.get("depth", 999)
            if depth_at <= 1:
                collapsed = True
            else:
                inherited_depth = 999 if depth_at >= 999 else depth_at - 1
        if log_fn:
            log_fn(f"Planning: {folder}" + (" [collapsed]" if collapsed else ""))
        _plan_folder(folder_path, sec, policy, folder, "", 2, inherited_depth, overlay,
                     exclude_files, plan, log_fn, collapsed=collapsed)
        if sec.documents():
            plan.sections.append(sec)
        else:
            plan.warnings.append(f"Top-level folder '{folder}' has no documents after "
                                 f"applying the folder policy - it is left out.")

    # Preflight: drop documents that cannot be opened at all
    for sec in list(plan.sections):
        _preflight_section(sec, plan)
    plan.sections = [s for s in plan.sections if s.documents()]

    # Labels are assigned after pruning so numbering has no gaps
    for i, sec in enumerate(plan.sections, start=1):
        sec.label = convert_to_style(i, numbering)

    scope = normalise_separator_scope(
        opts.get("separator_scope") or policy.get("separators"), "all_sections")
    apply_separator_scope(plan, scope)
    apply_visibility(plan, toc_documents=opts.get("toc_include_documents", True),
                     bookmark_documents=opts.get("bookmark_include_documents", True))
    assign_node_ids(plan)
    apply_page_selection(plan, opts.get("page_selection"))
    return plan


def assign_node_ids(plan):
    for i, node in enumerate(plan.walk(), start=1):
        node.node_id = i


# ---------------------------------------------------------------------------
# Page manifest (selection / order / rotation / exclusion)
# ---------------------------------------------------------------------------

def apply_page_selection(plan, selection=None):
    """Initialise page_refs on every document from a selection spec.

    selection: {"mode": "all"|"first"|"first_n"|"ranges", "n": int,
                "ranges": {source file name or path: "1-3,5"},
                "default_range": "..."}
    Invalid custom ranges are recorded as warnings and that document falls
    back to all pages. Returns the plan."""
    selection = dict(selection or {"mode": "all"})
    mode = selection.get("mode", "all")
    if mode not in PAGE_SELECTION_MODES:
        mode = "all"
    selection["mode"] = mode
    plan.selection = selection
    n = max(1, int(selection.get("n", 1) or 1))
    ranges = selection.get("ranges") or {}
    for doc in plan.documents():
        total = doc.expected_pages or 0
        if mode == "all":
            idx = list(range(total))
        elif mode == "first":
            idx = [0] if total else []
        elif mode == "first_n":
            idx = list(range(min(n, total)))
        else:
            spec = (ranges.get(doc.source_name) or ranges.get(doc.path)
                    or (ranges.get(norm_path(doc.path)) if doc.path else None)
                    or selection.get("default_range") or "")
            try:
                idx = parse_page_ranges(spec, total)
            except ValueError as e:
                plan.warnings.append(f"Page range for '{doc.source_name}' ignored ({e}); "
                                     f"all pages used.")
                idx = list(range(total))
        doc.page_refs = [{"src": i, "rot": 0, "inc": True} for i in idx]
    return plan


def find_parent(plan, node):
    for cand in plan.walk():
        if node in cand.children:
            return cand
    return None


def find_node(plan, node_id):
    for n in plan.walk():
        if n.node_id == node_id:
            return n
    return None


def move_page(doc, index, delta):
    """Move one page ref within its document by delta positions. Pages never
    leave their document (clamped at the document's ends). Returns new index."""
    refs = doc.page_refs
    if refs is None or not (0 <= index < len(refs)):
        return index
    new_index = max(0, min(len(refs) - 1, index + delta))
    ref = refs.pop(index)
    refs.insert(new_index, ref)
    return new_index


def move_pages(doc, indices, delta):
    """Move a set of page refs as a block (relative order kept), clamped to
    the document. Returns the new indices."""
    refs = doc.page_refs
    if refs is None or not indices:
        return []
    indices = sorted(i for i in set(indices) if 0 <= i < len(refs))
    if not indices:
        return []
    block = [refs[i] for i in indices]
    rest = [r for i, r in enumerate(refs) if i not in set(indices)]
    first = indices[0]
    insert_at = max(0, min(len(rest), first + delta))
    doc.page_refs = rest[:insert_at] + block + rest[insert_at:]
    return list(range(insert_at, insert_at + len(block)))


def rotate_pages(doc, indices, degrees):
    if doc.page_refs is None:
        return
    for i in indices:
        if 0 <= i < len(doc.page_refs):
            doc.page_refs[i]["rot"] = (doc.page_refs[i].get("rot", 0) + degrees) % 360


def set_pages_included(doc, indices, included):
    if doc.page_refs is None:
        return
    for i in indices:
        if 0 <= i < len(doc.page_refs):
            doc.page_refs[i]["inc"] = bool(included)


def move_node(plan, node, delta):
    """Reorder a section or document among its siblings. Returns new index."""
    parent = find_parent(plan, node)
    siblings = plan.sections if parent is None else parent.children
    if node not in siblings:
        return -1
    i = siblings.index(node)
    j = max(0, min(len(siblings) - 1, i + delta))
    siblings.pop(i)
    siblings.insert(j, node)
    return j


def snapshot_plan(plan):
    """Deep copy used for undo / reset."""
    return copy.deepcopy(plan)


def prune_excluded(plan):
    """Drop documents with no included pages and sections left empty.
    Returns the list of dropped document names."""
    dropped = []

    def prune(section):
        for child in list(section.children):
            if child.kind == "document":
                if not child.included_pages():
                    dropped.append(child.source_name)
                    section.children.remove(child)
            else:
                prune(child)
                if not child.documents():
                    section.children.remove(child)

    for sec in list(plan.sections):
        prune(sec)
    plan.sections = [s for s in plan.sections if s.documents()]
    return dropped


def relabel_sections(plan, numbering_style=None):
    """Re-number top-level sections after reordering or pruning."""
    numbering = numbering_style or getattr(plan, "numbering", None) or "alphabetic"
    for i, sec in enumerate(plan.sections, start=1):
        sec.label = convert_to_style(i, numbering)


def plan_to_manifest(plan):
    """JSON-serialisable manifest of the current selection/order/rotation."""
    return {"version": 1, "selection": plan.selection,
            "sections": [s.as_dict() for s in plan.sections]}


def manifest_totals(plan):
    docs = plan.documents()
    inc = sum(len(d.included_pages()) for d in docs)
    total = sum(len(d.page_refs) if d.page_refs is not None else (d.expected_pages or 0)
                for d in docs)
    rotated = sum(1 for d in docs for r in (d.page_refs or []) if r.get("rot"))
    return {"documents": len(docs), "documents_with_pages": sum(1 for d in docs if d.included_pages()),
            "pages_included": inc, "pages_excluded": total - inc, "pages_rotated": rotated,
            "sections": len(plan.sections)}


# ---------------------------------------------------------------------------
# Thumbnails (raster previews only - the bundle keeps native page content)
# ---------------------------------------------------------------------------

class ThumbnailCache:
    """Bounded LRU cache of rendered page thumbnails.

    get() returns (ppm_bytes, width, height) suitable for tk.PhotoImage.
    Rendering is serialised with a lock so it can run from a worker thread
    while the GUI stays responsive; the caller decides when to stop."""

    def __init__(self, max_items=400, max_open_docs=6):
        self.max_items = max_items
        self.max_open_docs = max_open_docs
        self._items = OrderedDict()
        self._docs = OrderedDict()
        self.lock = threading.Lock()

    def _doc(self, path):
        d = self._docs.get(path)
        if d is not None:
            self._docs.move_to_end(path)
            return d
        d = open_source_pdf(path)
        self._docs[path] = d
        while len(self._docs) > self.max_open_docs:
            _, old = self._docs.popitem(last=False)
            try:
                old.close()
            except Exception:
                pass
        return d

    def get(self, path, page_index, rot_delta=0, max_px=160):
        key = (norm_path(path), page_index, rot_delta % 360, max_px)
        with self.lock:
            hit = self._items.get(key)
            if hit is not None:
                self._items.move_to_end(key)
                return hit
            doc = self._doc(path)
            page = doc.load_page(page_index)
            rect = page.rect
            scale = max_px / max(rect.width, rect.height, 1)
            mat = fitz.Matrix(scale, scale).prerotate(rot_delta % 360)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            data = (pix.tobytes("ppm"), pix.width, pix.height)
            self._items[key] = data
            while len(self._items) > self.max_items:
                self._items.popitem(last=False)
            return data

    def page_size(self, path, page_index):
        with self.lock:
            page = self._doc(path).load_page(page_index)
            return page.rect.width, page.rect.height, page.rotation

    def clear(self):
        with self.lock:
            self._items.clear()
            for d in self._docs.values():
                try:
                    d.close()
                except Exception:
                    pass
            self._docs.clear()


def _preflight_section(section, plan):
    for child in list(section.children):
        if child.kind == "document":
            info = inspect_pdf(child.path)
            if info["previous_output"]:
                plan.previous_outputs.append(child.source_name)
                plan.warnings.append(f"Ignored '{child.source_name}': {info['reason']}.")
                section.children.remove(child)
            elif info["reason"]:
                plan.skipped_documents.append((child.source_name, info["reason"]))
                plan.warnings.append(f"Skipped '{child.source_name}': {info['reason']}.")
                section.children.remove(child)
            else:
                child.expected_pages = info["pages"]
        else:
            _preflight_section(child, plan)
            if not child.documents():
                section.children.remove(child)


def apply_separator_scope(plan, scope):
    """Set node.separator from the global scope. Collapsed subtrees never
    receive internal separators (they are one TOC line by choice)."""
    scope = normalise_separator_scope(scope)
    for node in plan.walk():
        node.separator = False
        if scope == "none" or node.collapsed_child:
            continue
        if node.kind == "section":
            if node.level == 1:
                node.separator = True
            elif scope in ("all_sections", "every_document"):
                node.separator = True
        elif scope == "every_document":
            # A top-level document already has its section separator.
            node.separator = True
    for sec in plan.sections:
        if sec.root_document:
            for c in sec.children:
                c.separator = False
    return plan


def apply_visibility(plan, toc_documents=True, bookmark_documents=True):
    for node in plan.walk():
        if node.kind == "document":
            if not toc_documents:
                node.toc = False
            node.bookmark = bool(bookmark_documents)
    return plan


# ---------------------------------------------------------------------------
# Assembly
# ---------------------------------------------------------------------------

def _check_cancel(opts):
    ev = opts.get("cancel_event") if opts else None
    if ev is not None and ev.is_set():
        raise BundleCancelled()


def assemble_body(plan, opts=None, log_fn=None):
    """Insert separators and documents in plan order into a new document.

    Sets start / end / sep_page / content_start (1-based, body-relative)
    on every node. Returns the fitz.Document."""
    opts = opts or {}
    term = CURRENT_TERMS["section_term"]
    body = fitz.open()
    state = {"page": 1}

    def emit_section(node):
        _check_cancel(opts)
        node.start = state["page"]
        if node.separator:
            if node.level == 1:
                kicker = f"{term} {node.label}" if node.label else term
                sep = create_separator_page(node.title, "", kicker=kicker, top_level=True)
            else:
                sep = create_separator_page(node.title, "", top_level=False)
            state["page"] += _insert_bytes(body, sep)
            node.sep_page = node.start
        for child in node.children:
            if child.kind == "section":
                emit_section(child)
            else:
                emit_document(child)
        node.end = state["page"] - 1
        node.content_start = _first_content_page(node)

    def emit_document(node):
        _check_cancel(opts)
        node.start = state["page"]
        sep_inserted = False
        if node.separator:
            sep = create_separator_page(node.title, "", top_level=False)
            state["page"] += _insert_bytes(body, sep)
            node.sep_page = node.start
            sep_inserted = True
        if log_fn:
            log_fn(f"  Inserting: {node.source_name}")
        pages = _insert_document_pages(body, node)
        if pages == 0:
            node.failed = True
            if log_fn:
                log_fn(f"  WARNING: no pages could be read from {node.source_name}")
            if sep_inserted:
                body.delete_page(body.page_count - 1)
                state["page"] -= 1
                node.sep_page = None
            node.content_start = None
            node.end = state["page"] - 1
            node.pages = 0
            return
        wanted = len(node.included_pages())
        if wanted and pages != wanted and log_fn:
            log_fn(f"  WARNING: {node.source_name}: inserted {pages} of {wanted} selected page(s)")
        node.content_start = state["page"]
        state["page"] += pages
        node.pages = pages
        node.end = state["page"] - 1

    prune_excluded(plan)
    relabel_sections(plan)
    for sec in plan.sections:
        emit_section(sec)

    # Drop documents that produced nothing, then sections left empty
    for sec in list(plan.sections):
        _drop_failed(sec)
    plan.sections = [s for s in plan.sections if s.documents() and s.end >= s.start]
    return body


def _insert_document_pages(body, node):
    """Insert the document's manifest pages (or all pages) into body.
    Returns the number of pages inserted."""
    if node.page_refs is None:
        return insert_pdf_clean(body, node.path)
    refs = node.included_pages()
    if not refs:
        return 0
    # Fast path: complete, in-order, unrotated selection
    if (node.expected_pages and len(refs) == node.expected_pages
            and all(r["src"] == i and not r.get("rot") for i, r in enumerate(refs))):
        return insert_pdf_clean(body, node.path)
    inserted = 0
    try:
        src = open_source_pdf(node.path)
    except Exception:
        return 0
    try:
        for r in refs:
            i = r["src"]
            if not (0 <= i < src.page_count):
                continue
            try:
                body.insert_pdf(src, from_page=i, to_page=i)
            except Exception:
                continue
            rot = r.get("rot", 0) % 360
            if rot:
                new_page = body.load_page(body.page_count - 1)
                new_page.set_rotation((new_page.rotation + rot) % 360)
            inserted += 1
    finally:
        src.close()
    return inserted


def _drop_failed(section):
    for child in list(section.children):
        if child.kind == "document":
            if child.failed:
                section.children.remove(child)
        else:
            _drop_failed(child)
            if not child.documents():
                section.children.remove(child)


def _first_content_page(node):
    for d in node.documents():
        if d.content_start:
            return d.content_start
    return None


def shift_pages(plan, offset):
    if not offset:
        return
    for node in plan.walk():
        for attr in ("start", "end", "sep_page", "content_start"):
            v = getattr(node, attr)
            if v is not None:
                setattr(node, attr, v + offset)


def build_toc_entries(plan, offset=0, term=None):
    """Flat TOC entries from the plan. Page numbers are body page + offset."""
    entries = []
    for sec in plan.sections:
        entries.append({
            "level": 1, "label": sec.label, "title": sec.title,
            "page": sec.start + offset, "show_page": True,
            "range_text": f"[pp {sec.start + offset}-{sec.end + offset}]",
        })
        if sec.root_document:
            continue
        for node in sec.walk():
            if node is sec or not node.toc:
                continue
            entries.append({"level": min(node.level, 5), "label": None, "title": node.title,
                            "page": node.start + offset, "show_page": True})
    return entries


def build_outline(plan, cover_pages=0, toc_pages=0, term=None):
    """PDF outline (bookmarks) as [(level, title, page_1based), ...]."""
    outline = []
    if cover_pages:
        outline.append((1, "Cover", 1))
    if toc_pages:
        outline.append((1, "Table of Contents", cover_pages + 1))
    for sec in plan.sections:
        outline.append((1, sec.display(term), sec.start))
        if sec.root_document:
            continue
        prev_level = 1
        for node in sec.walk():
            if node is sec or not node.bookmark:
                continue
            lvl = max(2, min(node.level, prev_level + 1, 6))
            outline.append((lvl, node.title, node.start))
            prev_level = lvl
    return outline


def _apply_toc_links(doc, link_pos, toc_start_idx, log_fn=None):
    """Add GOTO links to the TOC pages of a freshly re-opened document."""
    left_mm = LEFT_MARGIN_MM
    width_mm = 210 - LEFT_MARGIN_MM - RIGHT_MARGIN_MM
    added = 0
    for entry in link_pos:
        abs_idx = toc_start_idx + entry["page_in_toc"]
        if abs_idx >= doc.page_count:
            continue
        target_zero = entry["target_zero"]
        if target_zero < 0 or target_zero >= doc.page_count:
            if log_fn:
                log_fn(f"  Skipped link (target page {target_zero + 1} out of range)")
            continue
        toc_page = doc.load_page(abs_idx)
        rect = mm_rect(left_mm, entry["y0_mm"], left_mm + width_mm, entry["y1_mm"])
        try:
            toc_page.insert_link({"from": rect, "kind": fitz.LINK_GOTO, "page": target_zero})
            doc.reload_page(toc_page)
            added += 1
        except Exception as link_err:
            if log_fn:
                log_fn(f"  Skipped link (target {target_zero}): {link_err}")
    return added


def finalise_bundle(plan, body, opts, log_fn=None):
    """Prepend cover and TOC, add links, bookmarks, footers, OCR, compression.
    Returns (doc, info) where info has page counts."""
    log = log_fn or (lambda s: None)
    term = CURRENT_TERMS["section_term"]
    claim_title = opts.get("claim_title", "")
    generate_toc = bool(opts.get("generate_toc", True)) and bool(plan.sections)
    generate_bookmarks = bool(opts.get("generate_bookmarks", True))
    include_cover = bool(opts.get("include_cover", False))
    info = {"cover_pages": 0, "toc_pages": 0, "links": 0}

    # Cover
    cover_bytes = None
    cover_pages = 0
    if include_cover:
        log("Generating cover page...")
        cover_meta = {
            "title": claim_title or opts.get("cover_doc_type") or "Document Bundle",
            "subtitle": opts.get("cover_subtitle", ""),
            "doc_type": opts.get("cover_doc_type", ""),
            "project": opts.get("cover_project", ""),
            "contract_ref": opts.get("cover_contract_ref", ""),
            "prepared_by": opts.get("cover_prepared_by", ""),
            "prepared_for": opts.get("cover_prepared_for", ""),
            "date": opts.get("cover_date", ""),
            "revision": opts.get("cover_revision", ""),
            "sections": [(s.label, s.title) for s in plan.sections],
        }
        cover_bytes = create_cover_page(cover_meta)
        with fitz.open(stream=cover_bytes, filetype="pdf") as cd:
            cover_pages = cd.page_count

    # TOC: iterate until the page count is stable (numbers can wrap lines)
    toc_bytes, link_pos, toc_pages = None, [], 0
    if generate_toc:
        log("Building Table of Contents...")
        toc_pages = 0
        for _ in range(6):
            entries = build_toc_entries(plan, cover_pages + toc_pages, term)
            toc_bytes, link_pos, new_count = create_toc_page_hier(claim_title, entries)
            if new_count == toc_pages:
                break
            toc_pages = new_count
        else:
            log("  WARNING: TOC page count did not settle; links may be offset.")
        log(f"  TOC: {toc_pages} page(s), {len(link_pos)} link(s)")

    # Final document: cover + toc + body
    full = fitz.open()
    if cover_bytes:
        _insert_bytes(full, cover_bytes)
    if toc_bytes:
        _insert_bytes(full, toc_bytes)
    full.insert_pdf(body)
    body.close()
    offset = cover_pages + toc_pages
    shift_pages(plan, offset)

    # Round-trip via bytes so link insertion works on settled page objects
    merged_bytes = full.tobytes()
    full.close()
    doc = fitz.open(stream=merged_bytes, filetype="pdf")

    if toc_bytes:
        info["links"] = _apply_toc_links(doc, link_pos, cover_pages, log)

    if generate_bookmarks:
        try:
            doc.set_toc(build_outline(plan, cover_pages, toc_pages, term))
        except Exception as e:
            log(f"  WARNING: bookmarks could not be written: {e}")

    # OCR
    if opts.get("do_ocr"):
        if not ocr_available():
            log("WARNING: OCR libraries not available. Proceeding without OCR.")
        else:
            log("Running OCR on pages without text...")
            ocr_count = 0
            for pno in range(doc.page_count):
                _check_cancel(opts)
                try:
                    if ocr_page_if_needed(doc.load_page(pno)):
                        ocr_count += 1
                    if pno % 10 == 0:
                        log(f"  OCR scanning page {pno + 1} of {doc.page_count}")
                except Exception as e:
                    log(f"  OCR error on page {pno + 1}: {e}")
            log(f"  OCR added text to {ocr_count} page(s).")

    # Trim trailing blank pages (never below the last recorded content page)
    if opts.get("trim_trailing_blank", True):
        trimmed = 0
        last_content = max((n.end for n in plan.walk() if n.end), default=0)
        while doc.page_count > max(last_content, 1):
            last_page = doc.load_page(doc.page_count - 1)
            if _page_has_content(last_page):
                break
            doc.delete_page(doc.page_count - 1)
            trimmed += 1
        if trimmed:
            log(f"  Removed {trimmed} blank trailing page(s).")

    # Footers
    if opts.get("footer_enabled", True) and plan.sections:
        log("Adding section footers...")
        position = opts.get("footer_position", "bottom")
        layout = opts.get("footer_layout", "split")
        inset = float(opts.get("footer_inset_mm", 6.35) or 6.35)
        short = term_short(term)
        for sec in plan.sections:
            total_pages = sec.end - sec.start + 1
            right_label = sec.display(term)
            left_id = f"{short} {sec.label}" if sec.label else sec.title[:40]
            for idx, p in enumerate(range(sec.start - 1, sec.end)):
                if 0 <= p < doc.page_count:
                    left = f"{left_id} - P{idx + 1:03d} / {total_pages:03d}"
                    try:
                        draw_footer(doc, p, left, right_label, position, layout, inset)
                    except Exception as e:
                        log(f"  Footer skip page {p + 1}: {e}")

    comp_mode = opts.get("compression", "high")
    if comp_mode != "high":
        log(f"Applying {comp_mode} compression...")
        compress_document(doc, mode=comp_mode, log_fn=log)

    try:
        doc.set_metadata({
            "title": claim_title or opts.get("cover_doc_type") or "Document Bundle",
            "author": opts.get("cover_prepared_by", "") or "",
            "subject": opts.get("cover_project", ""),
            "keywords": f"Document Bundle, {term}s",
            "creator": PDF_CREATOR_TAG,
        })
    except Exception:
        pass

    info["cover_pages"] = cover_pages
    info["toc_pages"] = toc_pages
    info["page_count"] = doc.page_count
    info.update(manifest_totals(plan))
    info["skipped_documents"] = list(plan.skipped_documents)
    return doc, info


# ---------------------------------------------------------------------------
# Settings persistence
# ---------------------------------------------------------------------------

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
    # New in 6.0 - defaults reproduce the previous behaviour
    "style_preset": "claims_navy",
    "style_toc_density": "standard",
    "section_term": "Appendix",
    "section_term_custom": "",
    "separator_scope": "all_sections",
    "generate_bookmarks": True,
    "toc_include_documents": True,
    "bookmark_include_documents": True,
    "footer_enabled": True,
    "footer_position": "bottom",
    "footer_layout": "split",
    "footer_inset_mm": 6.35,
    "convert_office": False,
    "office_reconvert": False,
    "excel_quick_print_path": "",
    "folder_layout": "preserve_skip_numeric",
    # Page extraction / output mode (6.1)
    "page_mode": "all",
    "page_first_n": 1,
    "output_mode": "bundle",
    "recursive": True,
}


def load_settings(path=None):
    path = path or SETTINGS_FILE
    if not os.path.isfile(path):
        return dict(DEFAULT_SETTINGS)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(DEFAULT_SETTINGS)
        merged.update({k: v for k, v in data.items() if k in DEFAULT_SETTINGS})
        merged["separator_scope"] = normalise_separator_scope(merged.get("separator_scope"))
        if merged.get("numbering_style") not in NUMBERING_STYLES:
            merged["numbering_style"] = "alphabetic"
        return merged
    except Exception:
        return dict(DEFAULT_SETTINGS)


def save_settings(settings, path=None):
    path = path or SETTINGS_FILE
    try:
        tmp = f"{path}.tmp.{os.getpid()}"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
        os.replace(tmp, path)
        return True
    except Exception:
        return False


def effective_term(settings_or_opts):
    term = settings_or_opts.get("section_term") or "Appendix"
    if term == "Custom":
        term = (settings_or_opts.get("section_term_custom") or "").strip() or "Section"
    return term


# ---------------------------------------------------------------------------
# Build entry points
# ---------------------------------------------------------------------------

def default_output_path(base_directory, claim_title):
    out_name = sanitize_text(claim_title) if claim_title else "Combined_Appendices"
    out_name = re.sub(r'[\\/:*?"<>|]+', "-", out_name).strip() or "Combined_Appendices"
    return os.path.join(base_directory, f"{out_name}.pdf")


def build_bundle_sync(opts, log_fn=None):
    """Run the whole build synchronously. Returns (output_path, plan, info).

    Raises BundleCancelled if opts["cancel_event"] is set, ValueError when
    there is nothing to merge, and any other exception from the build."""
    log = log_fn or (lambda s: None)
    base_directory = opts["base_directory"]
    claim_title = opts.get("claim_title", "")

    set_current_style({
        **(opts.get("style") or {}),
    })
    set_current_terms(effective_term(opts), opts.get("numbering_style") or "alphabetic")

    output_path = opts.get("output_path") or default_output_path(base_directory, claim_title)
    opts = dict(opts)
    opts["output_path"] = output_path

    log(f"Output target: {output_path}")
    plan = opts.get("plan")
    if plan is None:
        log("Discovering inputs...")
        policy = opts.get("policy") or {"mode": "automatic",
                                        "layout": opts.get("folder_layout", "preserve_skip_numeric")}
        plan = plan_bundle(base_directory, policy, opts, overlay=opts.get("overlay"), log_fn=log)
        for w in plan.warnings:
            log(f"  NOTE: {w}")
    else:
        log("Using the organised page manifest.")
    if not plan.sections:
        raise ValueError("No PDFs or folders with PDFs found in the base directory.")
    if opts.get("page_selection"):
        log(f"  Page selection: {PAGE_SELECTION_LABELS.get(plan.selection.get('mode'), 'all')}")
    log(f"  {len(plan.sections)} top-level section(s), {len(plan.documents())} document(s)")
    scope = normalise_separator_scope(opts.get("separator_scope") or policy.get("separators"))
    log(f"  Separators: {SEPARATOR_SCOPE_LABELS[scope]}")

    log("Assembling documents...")
    body = assemble_body(plan, opts, log)
    if not plan.sections:
        body.close()
        raise ValueError("No pages could be read from the input documents.")
    doc, info = finalise_bundle(plan, body, opts, log)

    log("Saving final PDF...")
    try:
        doc.save(output_path, garbage=4, deflate=True)
    finally:
        doc.close()
    log(f"Totals: {info['sections']} section(s), {info['documents_with_pages']} document(s), "
        f"{info['pages_included']} source page(s) included, {info['pages_excluded']} excluded, "
        f"{info['page_count']} page(s) in the bundle.")
    if info["skipped_documents"]:
        log(f"PARTIAL OUTPUT: {len(info['skipped_documents'])} document(s) were skipped:")
        for name, reason in info["skipped_documents"]:
            log(f"  - {name}: {reason}")
    log(f"DONE. Bundle saved: {output_path}")
    return output_path, plan, info


def clean_folder_name(name):
    """Folder name made safe for the per-folder output file (brackets to '-')."""
    return "".join("-" if c in "()[]{}<>" else c for c in name)


def revisioned_path(desired_path, max_revisions=99):
    """Same path if free, otherwise <name>_rev.NN<ext> (never overwrites)."""
    if not os.path.exists(desired_path):
        return desired_path
    base, ext = os.path.splitext(desired_path)
    n = 1
    while True:
        cand = f"{base}_rev.{n:02d}{ext}" if n <= max_revisions else f"{base}_rev.{n}{ext}"
        if not os.path.exists(cand):
            return cand
        n += 1


def build_per_folder_first_pages(base_directory, opts, log_fn=None):
    """Behaviour of the owner's Combine-FirstPages scripts on the shared
    pipeline: for the chosen folder (and every subfolder when recursive) take
    the selected pages (default: first page) of each DIRECT PDF and write
    <FolderName>#Combined_FirstPages.pdf into that folder.

    Previous outputs (names containing '#Combined') are never inputs and are
    never overwritten: a _rev.NN name is used when the output already
    exists. Returns a list of result dicts per folder."""
    log = log_fn or (lambda s: None)
    recursive = opts.get("recursive", True)
    folders = [base_directory]
    if recursive:
        for cur, dirs, _files in os.walk(base_directory):
            dirs[:] = sorted((d for d in dirs if not is_excluded_folder(d)), key=natural_key)
            for d in dirs:
                folders.append(os.path.join(cur, d))
    selection = opts.get("page_selection") or {"mode": "first"}
    results = []
    for folder in folders:
        _check_cancel(opts)
        name = os.path.basename(os.path.normpath(folder)) or "Root"
        docs, _ = list_folder_items(folder, opts.get("overlay"))
        if not docs:
            log(f"=== {name}: no PDFs to process, skipped.")
            results.append({"folder": folder, "output": None, "skipped": [], "info": None,
                            "reason": "no PDFs"})
            continue
        output = revisioned_path(os.path.join(folder, clean_folder_name(name) + FIRST_PAGES_SUFFIX))
        log(f"=== {name}: {len(docs)} PDF(s) -> {os.path.basename(output)}")
        sub_opts = dict(opts)
        sub_opts.update({
            "base_directory": folder, "output_path": output, "direct_files_only": True,
            "page_selection": selection, "plan": None,
            "policy": {"mode": "automatic", "layout": "flatten"},
            "claim_title": opts.get("claim_title") or clean_folder_name(name),
        })
        # Plain output like the owner's scripts: no cover, contents, separators
        # or footers; bookmarks (one per file) are kept for navigation.
        sub_opts.update({"separator_scope": "none", "generate_toc": False,
                         "generate_bookmarks": True, "footer_enabled": False,
                         "include_cover": False})
        try:
            out, plan, info = build_bundle_sync(sub_opts, log)
            results.append({"folder": folder, "output": out, "info": info,
                            "skipped": list(plan.skipped_documents), "reason": None})
        except BundleCancelled:
            raise
        except Exception as e:
            log(f"  FAILED for {name}: {e}")
            results.append({"folder": folder, "output": None, "info": None, "skipped": [],
                            "reason": str(e)})
    made = sum(1 for r in results if r["output"])
    skipped_docs = sum(len(r["skipped"]) for r in results)
    log(f"Per-folder run finished: {made} combined file(s) written, "
        f"{len(results) - made} folder(s) without output, {skipped_docs} document(s) skipped.")
    return results


def build_bundle(opts, log_fn, done_fn):
    """Thread-friendly wrapper: done_fn(success, info_string)."""
    try:
        output_path, plan, info = build_bundle_sync(opts, log_fn)
        done_fn(True, output_path)
    except BundleCancelled:
        log_fn("Build cancelled.")
        done_fn(False, "Build cancelled.")
    except ValueError as e:
        log_fn(f"ERROR: {e}")
        done_fn(False, str(e))
    except Exception:
        log_fn("UNEXPECTED ERROR:")
        log_fn(traceback.format_exc())
        done_fn(False, "Unexpected error - see status panel.")


# ---------------------------------------------------------------------------
# Style preview
# ---------------------------------------------------------------------------

def render_style_preview(style, section_term="Appendix", numbering_style="alphabetic",
                         cover_meta=None):
    """Cover + top-level separator + nested separator + sample contents page
    in the given style. Returns PDF bytes."""
    saved_style, saved_terms = dict(CURRENT_STYLE), dict(CURRENT_TERMS)
    try:
        set_current_style(style)
        set_current_terms(section_term, numbering_style)
        meta = dict(cover_meta or {})
        meta.setdefault("title", "Style preview")
        meta.setdefault("doc_type", "Document Bundle")
        meta.setdefault("project", "Project name")
        meta.setdefault("prepared_by", "Prepared by")
        meta.setdefault("revision", "Rev 00")
        labels = [convert_to_style(i, numbering_style) for i in (1, 2, 3)]
        meta["sections"] = [(labels[0], "Notice of claim"), (labels[1], "Correspondence"),
                            (labels[2], "Programme and records")]
        out = fitz.open()
        _insert_bytes(out, create_cover_page(meta))
        kicker = f"{section_term} {labels[0]}" if labels[0] else section_term
        _insert_bytes(out, create_separator_page("Notice of claim", "", kicker=kicker,
                                                 top_level=True))
        _insert_bytes(out, create_separator_page("Contractor correspondence", "",
                                                 top_level=False))
        entries = [
            {"level": 1, "label": labels[0], "title": "Notice of claim", "page": 3,
             "range_text": "[pp 3-8]"},
            {"level": 2, "title": "Letter ref 001", "page": 4},
            {"level": 2, "title": "Letter ref 002", "page": 6},
            {"level": 1, "label": labels[1], "title": "Correspondence", "page": 9,
             "range_text": "[pp 9-20]"},
            {"level": 2, "title": "Contractor correspondence", "page": 10},
            {"level": 3, "title": "Email 2025-01-04", "page": 11},
            {"level": 1, "label": labels[2], "title": "Programme and records", "page": 21,
             "range_text": "[pp 21-40]"},
        ]
        toc_bytes, _, _ = create_toc_page_hier("Style preview", entries)
        _insert_bytes(out, toc_bytes)
        data = out.tobytes()
        out.close()
        return data
    finally:
        CURRENT_STYLE.update(saved_style)
        CURRENT_TERMS.update(saved_terms)
