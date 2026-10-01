"""Validate an Excel workbook against the workbook-readability standard.

A linter, not an authority. It catches mechanical defects - banned phrases,
corrupted words, codes carrying meaning, named-range policy, tab architecture,
unlabelled zero rates and missing first-tab content. It cannot judge whether
the writing is clear.

Usage
    python validate_workbook.py "workbook.xlsx"
    python validate_workbook.py "workbook.xlsx" --rules rules.json --json report.json

Exit codes
    0  clean, or warnings only
    1  one or more errors
    2  could not read the workbook

Note on formula errors: openpyxl does not calculate. This check reads cached
results, so it only works on a workbook already recalculated by Excel or
LibreOffice. If any formulas have no cached value the check says so, rather than
reporting a clean result it has not earned.
"""

import argparse
import json
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    sys.stderr.write("openpyxl is required:  pip install openpyxl\n")
    sys.exit(2)

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RULES = os.path.join(HERE, "..", "rules.json")

# Genuine external workbook references look like '[Book1.xlsx]Sheet1'!A1 or [1]Sheet1!A1.
# Structured table references - Table[Column] - are not external and must not be flagged.
DOCUMENT_CONTROL_TABS = {"project info", "cover", "appendix cover"}

# Only an explicit workbook filename counts. [2026] is a table year heading, not a link.
EXTERNAL_REF = re.compile(r"\[[^\]]+\.xls[xmb]?\][^!]*!", re.IGNORECASE)


class Finding:
    __slots__ = ("severity", "rule", "location", "detail")

    def __init__(self, severity, rule, location, detail):
        self.severity, self.rule, self.location, self.detail = severity, rule, location, detail

    def as_dict(self):
        return {"severity": self.severity, "rule": self.rule,
                "location": self.location, "detail": self.detail}


def text_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip():
                yield cell


def visible(wb):
    return [ws for ws in wb.worksheets if ws.sheet_state == "visible"]


def iter_defined_names(wb):
    """Both scopes. A sheet-scoped name defeats the policy if only workbook scope is read."""
    for name in wb.defined_names.keys():
        yield "workbook", name
    for ws in wb.worksheets:
        for name in getattr(ws, "defined_names", {}).keys():
            yield ws.title, name


# ------------------------------------------------------------------ checks

def check_banned_phrases(wb, rules, out):
    """Tactical language is scanned everywhere - a hidden tab can be unhidden and forwarded.

    The other groups are communication failures, so they are scanned on visible sheets only.
    """
    cfg = rules["banned_phrases"]
    everywhere = set(cfg.get("all_sheet_groups", ["tactics"]))
    vis = {ws.title for ws in visible(wb)}
    for ws in wb.worksheets:
        on_visible = ws.title in vis
        for cell in text_cells(ws):
            low = cell.value.lower()
            for group, phrases in cfg["groups"].items():
                if group not in everywhere and not on_visible:
                    continue
                for p in phrases:
                    if p.lower() in low:
                        i = low.find(p.lower())
                        where = f"{ws.title}!{cell.coordinate}" + ("" if on_visible else " (hidden tab)")
                        out.append(Finding(cfg["severity"], f"banned_phrases/{group}", where,
                                           f'"{p}" in: ...{cell.value[max(0,i-35):i+len(p)+35].strip()}...'))
                        break


def check_corrupted_words(wb, rules, out):
    cfg = rules["corrupted_words"]
    pats = [re.compile(p) for p in cfg["patterns"]]
    for ws in wb.worksheets:
        for cell in text_cells(ws):
            for pat in pats:
                m = pat.search(cell.value)
                if m:
                    out.append(Finding(cfg["severity"], "corrupted_words",
                                       f"{ws.title}!{cell.coordinate}",
                                       f'"{m.group(0)}" - find-and-replace damage'))
                    break


def check_internal_codes(wb, rules, out, reader_tabs):
    """A code is a defect only when it carries the meaning.

    Flagged: a code opening a cell that then continues as narrative ("R2 the
    Contractor's sections..."), and unambiguous project shorthand.
    Not flagged: "Drawing grid: A1", "Reinforcement bar R10", or a bare "R2"
    standing alone in a reference column.
    """
    cfg = rules["internal_codes"]
    sev = cfg["severity"]
    lead = re.compile(r"^\s*(" + "|".join(cfg["leading_code_patterns"]) + r")\b[\s:.,-]*(.*)$")
    shorthand = [re.compile(p, re.IGNORECASE) for p in cfg["unambiguous_shorthand"]]
    contextual = [re.compile(p, re.IGNORECASE) for p in cfg.get("contextual_patterns", [])]
    minw = cfg["min_following_words"]

    for ws in wb.worksheets:
        if ws.title not in reader_tabs:
            continue
        for cell in text_cells(ws):
            v = cell.value.strip()
            m = lead.match(v)
            if m and len(m.group(2).split()) >= minw:
                out.append(Finding(sev, "internal_codes/leads_sentence",
                                   f"{ws.title}!{cell.coordinate}",
                                   f'"{m.group(1)}" opens a narrative cell: {v[:80]}'))
                continue
            hit = False
            for pat in contextual:
                cm = pat.search(v)
                if cm:
                    out.append(Finding(sev, "internal_codes/carries_meaning",
                                       f"{ws.title}!{cell.coordinate}",
                                       f'"{cm.group(0).strip()}" - the code is doing the explaining: {v[:70]}'))
                    hit = True
                    break
            if hit:
                continue
            for pat in shorthand:
                sm = pat.search(v)
                if sm:
                    out.append(Finding(sev, "internal_codes/shorthand",
                                       f"{ws.title}!{cell.coordinate}",
                                       f'"{sm.group(0)}" - state the matter instead'))
                    break


def check_named_ranges(wb, rules, out):
    cfg = rules["named_ranges"]
    sev = cfg["severity"]
    allowed = set(cfg["allowed_exact"])
    excluded = set(cfg.get("excluded_from_count", []))

    names = [(scope, n) for scope, n in iter_defined_names(wb) if n not in excluded]
    for scope, n in names:
        # An approved name only counts at workbook scope. A sheet-scoped PROJECT_NAME is not
        # the document-control name the Cover links to - it is a calculation wearing its badge.
        if scope == "workbook" and n in allowed:
            continue
        where = f"name: {n}" + ("" if scope == "workbook" else f"  (sheet-scoped on {scope})")
        if n in allowed:
            out.append(Finding(sev, "named_ranges/approved_name_wrong_scope", where,
                               "an approved document-control name must be workbook-scoped - "
                               "a sheet-scoped copy can hide a calculation while appearing compliant"))
            continue
        out.append(Finding(sev, "named_ranges/not_document_control", where,
                           "not on the approved document-control list - replace with an explicit sheet "
                           "reference, or add it to allowed_exact if a workbook feature genuinely requires it"))
    if len(names) > cfg["max_total"]:
        out.append(Finding(sev, "named_ranges/count", "workbook",
                           f"{len(names)} defined names, limit {cfg['max_total']}"))


def check_tab_architecture(wb, rules, out):
    cfg = rules["tab_architecture"]
    sev = cfg["severity"]
    vis = visible(wb)
    titles = [ws.title for ws in vis]

    if len(vis) > cfg["max_visible_tabs"]:
        out.append(Finding(sev, "tab_architecture/too_many_visible", "workbook",
                           f"{len(vis)} visible tabs, target {cfg['max_visible_tabs']}: {', '.join(titles)}"))




def check_first_tab(wb, rules, out, first_tab):
    """Structural checks only.

    Whether the tab answers purpose, result, basis and action cannot be tested
    mechanically. Keyword matching fails - "No total has been calculated" contains
    "total". Counting cells fails too - three labels beside three numbers is a valid
    summary. Both were removed rather than left giving false confidence. This is the
    human acceptance test.
    """
    cfg = rules["first_tab"]
    sev = cfg["severity"]
    if first_tab is None:
        out.append(Finding("error", "first_tab/missing", "workbook",
                           "no substantive visible tab - the workbook carries only "
                           "document-control tabs and answers nothing"))
        return

    if not any(c.value is not None for row in first_tab.iter_rows() for c in row):
        out.append(Finding(sev, "first_tab/empty", first_tab.title, "the substantive tab is empty"))

    for cell in text_cells(first_tab):
        t = cell.value.strip()
        if t.startswith("=") or len(t) < cfg["block_min_chars"]:
            continue
        w = len(t.split())
        if w > cfg["warn_words_per_block"]:
            out.append(Finding("info", "first_tab/long_block",
                               f"{first_tab.title}!{cell.coordinate}",
                               f"{w} words - check whether it is doing two jobs"))


def _header_matchers(keywords):
    """Whole-word matchers. 'separately' must not register as a 'rate' heading."""
    pats = []
    for kw in keywords:
        esc = re.escape(kw)
        tail = r"\b" if kw[-1].isalnum() else ""
        pats.append(re.compile(r"\b" + esc + tail, re.IGNORECASE))
    return pats


def _contiguous_blocks(cols):
    """Group column numbers into runs. A blank header cell separates two tables."""
    blocks, run = [], []
    for c in sorted(cols):
        if run and c == run[-1] + 1:
            run.append(c)
        else:
            if run:
                blocks.append(run)
            run = [c]
    if run:
        blocks.append(run)
    return blocks


def check_status_labelling(wb, rules, out):
    """A zero in a rate column, with no status against it.

    Table boundaries are respected in three ways: header rows are found anywhere on the
    sheet, not only near the top; a header row is split into blocks at blank header cells,
    so a neighbouring table cannot lend its status column; and a table ends at the row
    before the next rate header in the same column, so stacked tables are not scanned into
    one another or reported twice.

    Recognises conventional vertical rate tables whose block carries at least two headed
    columns. Transposed tables, and a lone rate heading with no companion column, are not
    detected - see rules.json scope_note.
    """
    cfg = rules["status_labelling"]
    sev = cfg["severity"]
    rate_pats = _header_matchers(cfg["rate_header_keywords"])
    stat_pats = _header_matchers(cfg["status_header_keywords"])
    vocab = [s.lower() for s in cfg["status_vocabulary"]]
    linked = [re.compile(p, re.IGNORECASE) for p in cfg.get("rate_linked_status_patterns", [])]
    max_words = cfg.get("header_max_words", 6)
    min_block = cfg.get("min_block_columns", 2)

    for ws in wb.worksheets:
        # 1. every header row on the sheet, not just the first few
        headers = []          # (row, rate_cols, status_cols, blocks)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            filled, rate_here, stat_here = [], [], []
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                text = cell.value.strip()
                if not text:
                    continue
                filled.append(cell.column)
                if len(text.split()) > max_words:        # a sentence is not a heading
                    continue
                if any(p.search(text) for p in rate_pats):
                    rate_here.append(cell.column)
                if any(p.search(text) for p in stat_pats):
                    stat_here.append(cell.column)
            if rate_here:
                headers.append((row[0].row, rate_here, stat_here, _contiguous_blocks(filled)))

        if not headers:
            continue

        # 2. where each table ends: the row before the next header carrying a rate in the same column
        rows_by_col = {}
        for hrow, rate_cols, _, _ in headers:
            for c in rate_cols:
                rows_by_col.setdefault(c, []).append(hrow)

        seen = set()
        for hrow, rate_cols, stat_cols, blocks in headers:
            for col in rate_cols:
                block = next((b for b in blocks if col in b), None)
                if block is None or len(block) < min_block:
                    continue                              # a lone heading is not a table
                # 3. status only from inside this block
                in_block = [c for c in stat_cols if c in block and c != col]
                status_col = min(in_block, key=lambda c: abs(c - col)) if in_block else None
                later = [r for r in rows_by_col[col] if r > hrow]
                last_row = (min(later) - 1) if later else ws.max_row

                for r in range(hrow + 1, last_row + 1):
                    cell = ws.cell(row=r, column=col)
                    if not isinstance(cell.value, (int, float)) or cell.value != 0:
                        continue
                    key = (ws.title, cell.coordinate)
                    if key in seen:
                        continue
                    labelled = False
                    if status_col is not None:
                        v = ws.cell(row=r, column=status_col).value
                        if isinstance(v, str) and any(s in v.lower() for s in vocab):
                            labelled = True
                    if not labelled:
                        # 4. description-based status, inside this block only
                        for c2 in block:
                            v = ws.cell(row=r, column=c2).value
                            if isinstance(v, str) and any(p.search(v) for p in linked):
                                labelled = True
                                break
                    if not labelled:
                        seen.add(key)
                        detail = ("zero in a rate column with no status in its bound status cell"
                                  if status_col is not None else
                                  "zero in a rate column and this table has no status or note column")
                        out.append(Finding(sev, "status_labelling/unlabelled_zero_rate",
                                           f"{ws.title}!{cell.coordinate}",
                                           detail + " - state 'Not established' or similar"))


def check_integrity(wb_f, wb_v, rules, out):
    cfg = rules["integrity"]
    sev = cfg["severity"]

    if cfg.get("no_external_links"):
        try:
            for link in (wb_f._external_links or []):
                out.append(Finding(sev, "integrity/external_link", "workbook",
                                   f"external workbook link: {getattr(link, 'file_link', link)}"))
        except AttributeError:
            pass
        for ws in wb_f.worksheets:
            for cell in text_cells(ws):
                if cell.value.startswith("=") and EXTERNAL_REF.search(cell.value):
                    out.append(Finding(sev, "integrity/external_link",
                                       f"{ws.title}!{cell.coordinate}", cell.value[:70]))

    if cfg.get("no_formula_errors"):
        errvals = {e.upper() for e in cfg.get("error_values", [])}
        formulas = uncached = 0
        for ws in wb_f.worksheets:
            for cell in text_cells(ws):
                if cell.value.startswith("="):
                    formulas += 1
                    if wb_v[ws.title][cell.coordinate].value is None:
                        uncached += 1
        for ws in wb_v.worksheets:
            for cell in text_cells(ws):
                if cell.value.strip().upper() in errvals:
                    out.append(Finding(sev, "integrity/formula_error",
                                       f"{ws.title}!{cell.coordinate}", cell.value.strip()))
        if uncached:
            out.append(Finding("info", "integrity/formula_results_partial", "workbook",
                               f"{uncached} of {formulas} formulas have no cached result - those were "
                               "not checked. Recalculate in Excel or LibreOffice for a complete check"))

    if cfg.get("metadata_author_required"):
        creator = (wb_f.properties.creator or "").strip()
        if not creator or creator.lower() in ("openpyxl", "python", "unknown"):
            out.append(Finding(sev, "integrity/metadata_author", "docProps",
                               f"author is '{creator}' - set a real author"))


# ------------------------------------------------------------------ report

def report(findings, path, json_out=None):
    errors = [f for f in findings if f.severity == "error"]
    warns = [f for f in findings if f.severity == "warning"]
    infos = [f for f in findings if f.severity == "info"]

    print(f"\nWorkbook readability check - {os.path.basename(path)}")
    print("=" * 72)

    by_rule = {}
    for f in findings:
        by_rule.setdefault(f.rule, []).append(f)
    order = {"error": 0, "warning": 1, "info": 2}
    for rule in sorted(by_rule, key=lambda r: (order[by_rule[r][0].severity], r)):
        g = by_rule[rule]
        print(f"\n[{g[0].severity.upper()}] {rule}  ({len(g)})")
        for f in g[:10]:
            print(f"    {f.location:<42} {f.detail[:104]}")
        if len(g) > 10:
            print(f"    ... and {len(g) - 10} more")

    print("\n" + "=" * 72)
    print(f"{len(errors)} error(s), {len(warns)} warning(s), {len(infos)} note(s)")
    if not errors and not warns:
        print("PASS - mechanical checks clean.")
    print("Warnings are prompts to look, not orders to change.")
    print("Clarity is not machine-checkable. Read the first tab aloud before issue.\n")

    if json_out:
        with open(json_out, "w", encoding="utf-8") as fh:
            json.dump({"workbook": os.path.basename(path), "errors": len(errors),
                       "warnings": len(warns), "notes": len(infos),
                       "findings": [f.as_dict() for f in findings]},
                      fh, indent=2, ensure_ascii=False)
        print(f"JSON report written to {json_out}\n")

    return 1 if errors else 0


def main():
    ap = argparse.ArgumentParser(description="Validate a workbook against the readability standard.")
    ap.add_argument("workbook")
    ap.add_argument("--rules", default=DEFAULT_RULES)
    ap.add_argument("--json", dest="json_out", default=None)
    ap.add_argument("--max-visible-tabs", type=int, default=None,
                    help="Override the visible-tab warning threshold. Use for IPC workbooks "
                         "carrying one tab per BOQ item.")
    ap.add_argument("--reader-tabs", default=None,
                    help="Comma-separated reader-facing tabs. Default: the first visible tab, "
                         "plus any tab named Summary or Issues...")
    args = ap.parse_args()

    if not os.path.exists(args.workbook):
        sys.stderr.write(f"not found: {args.workbook}\n")
        return 2
    with open(args.rules, encoding="utf-8") as fh:
        rules = json.load(fh)
    if args.max_visible_tabs is not None:
        rules["tab_architecture"]["max_visible_tabs"] = args.max_visible_tabs
    try:
        wb_f = openpyxl.load_workbook(args.workbook)
        wb_v = openpyxl.load_workbook(args.workbook, data_only=True)
    except Exception as exc:
        sys.stderr.write(f"could not open workbook: {exc}\n")
        return 2

    # The first SUBSTANTIVE visible tab. Document-control covers may precede it - and there
    # may be several, since wtp-excel-format creates Project Info, Cover and Appendix Cover,
    # all three of which are visible while a workbook is being prepared for PDF issue.
    # Exact titles only: a tab called "Cover Note - Summary" is substantive.
    vis = visible(wb_v)
    first_tab = None
    for ws in vis:
        if ws.title.strip().lower() in DOCUMENT_CONTROL_TABS:
            continue
        first_tab = ws
        break

    if args.reader_tabs:
        reader_tabs = {t.strip() for t in args.reader_tabs.split(",")}
    else:
        reader_tabs = {ws.title for ws in wb_v.worksheets
                       if ws.title.strip().lower().startswith(("summary", "issues"))}
        if first_tab is not None:
            reader_tabs.add(first_tab.title)

    out = []
    check_banned_phrases(wb_v, rules, out)
    check_corrupted_words(wb_v, rules, out)
    check_internal_codes(wb_v, rules, out, reader_tabs)
    check_named_ranges(wb_f, rules, out)
    check_tab_architecture(wb_f, rules, out)
    check_first_tab(wb_v, rules, out, first_tab)
    check_status_labelling(wb_v, rules, out)
    check_integrity(wb_f, wb_v, rules, out)

    return report(out, args.workbook, args.json_out)


if __name__ == "__main__":
    sys.exit(main())
