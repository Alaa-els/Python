# Document Bundle Builder - change report

Branch: `claude/claims-bundle-builder-improvements-d9bo9m`
Sources: `original/claims_bundle_builder (5).pyw`, `original/Excel_Quick_Print (9).pyw` (unchanged copies of the files supplied).

## Deliverables

| File | Status |
|---|---|
| `document_bundle_builder.pyw` | New launcher/GUI replacing `claims_bundle_builder.pyw` (bootstrap, dialogs and manual workflow carried over) |
| `bundle_engine.py` | New GUI-free engine (all build logic, testable without tkinter) |
| `bundle_organiser.py` | New pre-build page organiser window |
| `office_convert.py` | New Word/Excel conversion stage and Excel Quick Print companion adapter |
| `Excel_Quick_Print.pyw` | v10: companion mode added; standalone behaviour unchanged |
| `tests/` | 59 regression tests (56 engine/manifest/office, 3 GUI smoke) |
| `README.md`, `requirements-dev.txt` | Deployment notes |

The old single-file layout is split so that the engine can be tested. The five
runtime files must sit in the same folder; the `.pyw` files remain
double-click launchers and the first run still installs PyMuPDF/fpdf2.

## Priority 1 - separator pages, contents, bookmarks

Root cause in the original: `build_bundle` inserted a top-level separator for
every appendix unconditionally (lines ~2107 and ~2139), and the
`policy.separators` value only affected nested folders and files. "None"
therefore still produced one separator per appendix, and the bookmarks were
only written when the TOC was on.

Fix: the build is now plan-then-render. `plan_bundle` walks the tree and
applies the folder policy to a tree of `BundleNode` objects (sections and
documents); `apply_separator_scope` sets `node.separator` from ONE global
scope; `assemble_body` records the start page of every node *before* any
separator is inserted, so a link or bookmark targets the separator when it
exists and the first real content page when it does not.

Global scope (saved in settings as `separator_scope`, default
`all_sections`, which reproduces the old "folder_only" output):

| Scope | Pages inserted |
|---|---|
| `none` | none anywhere |
| `top_level` | one per top-level section (old "Appendix A" page) |
| `all_sections` | top-level and every nested folder heading |
| `every_document` | all sections plus one before every document (a root PDF gets only its section page) |

Legacy names `every_pdf`/`folder_only` are still accepted. Collapsed
subtrees never receive internal separators (unchanged).

Independent controls: `generate_toc`, `toc_include_documents`,
`generate_bookmarks`, `bookmark_include_documents`. Bookmarks are written
even when the TOC is off. Sections with no documents (empty folders, all
children skipped, unreadable files) are pruned before labels are assigned,
so there are no phantom pages, no orphan labels and no gaps. The TOC is
rebuilt until its page count settles (multi-page offsets), and links are
verified by tests against the actual target page text.

Manual mode is preserved: the per-section / per-folder dialogs are the same;
the separator question moved from the dialogs to the main window.

## Priority 2 - general Document Bundle Builder

- Terminology: Appendix / Section / Drawing / Document / Part / Exhibit /
  Schedule / Tab / Custom, numbering alphabetic / numeric / none. Used on
  separators, contents, bookmarks, footers ("Dwg 01 - P003 / 012").
- Presets (in addition to the existing pickers, all still editable):
  Claims bundle (navy + gold, the previous defaults), Minimal monochrome,
  Technical drawing register, Corporate report, Correspondence pack.
- New designs: covers `minimal`, `block`, `register` (metadata grid plus a
  register of the top-level sections); separators `minimal`, `band`, `block`;
  contents `plain`, `grid`; palettes `mono`, `graphite_orange`, `teal_slate`,
  `burgundy_cream`; contents density compact / standard / spacious.
- "Preview style" renders a real cover, two separators and a sample
  contents page to a PDF and opens it.
- Footers: on/off, bottom or top edge, split / left / centre / right, inset
  in mm. Text is placed through the page's derotation matrix so it is upright
  on rotated drawings (the original code dropped footers on rotated pages).
  Page size, rotation and vector content are never altered; compression
  only recompresses raster images.
- No drawing metadata is extracted; the Drawing preset only changes labels.

## Priority 3 - Word/Excel conversion stage

`office_convert.run_conversion_stage` runs before planning when enabled:

- Discovery of `.doc/.docx/.docm/.rtf/.xls/.xlsx/.xlsm/.xlsb` in the tree
  (lock files and tool folders ignored). Converted PDFs go to
  `<base>/_bundle_converted/<mirrored path>/<stem>.pdf` and are reused while
  newer than the source ("re-convert all" forces it). Source files are never
  written.
- Word: isolated `DispatchEx("Word.Application")`, read-only, macros forced
  off, links not updated, dummy `PasswordDocument` so protected files raise
  instead of prompting; encrypted OOXML detected from the file header.
  Only the instance the tool created is quit.
- Excel: Excel Quick Print is launched as a companion process with the
  workbook list, an output folder and a result file. The full wizard runs
  (sheets, print areas, orientation, scaling, page breaks, repeat rows,
  preview, `.printcfg.json` sidecars). Companion mode only pre-loads the
  files, redirects the PDF output, disables the `_rev.NN` workbook copy and
  writes `excel_quick_print_result.json` (also on close or startup failure,
  via `atexit`). `Workbooks.Open` now passes a dummy `Password` so protected
  workbooks fail cleanly instead of blocking an invisible Excel.
- Failures (no Office, no pywin32, password-protected, locked, tool missing,
  cancelled) are reported per document; the GUI asks whether to continue
  without them. PDF-only use works without pywin32.
- The engine sees converted files through an "overlay" (source path to PDF),
  so they keep their folder position, `_order.txt` name and title.

## Page organiser and extraction modes (second request)

- Extraction modes: all pages (default), first page of each file (WIR
  preset), first N pages, custom ranges per file (validated; invalid ranges
  are reported and that file falls back to all pages). "First page" is the
  first *source* page, never a generated page.
- Output modes: one bundle (folders become sections, contents/separators as
  set) or one `<Folder>#Combined_FirstPages.pdf` per folder as in the owner's
  scripts (recursive toggle; bracket characters in folder names become `-`;
  plain output with one bookmark per file). Files whose name contains
  `#Combined` and PDFs whose Creator metadata marks them as an earlier bundle
  output are never inputs; existing outputs are never overwritten
  (`_rev.NN`). Both modes run through the same plan/assemble pipeline.
- Ordering: `_order.txt` first, then deterministic natural order
  (`Dwg 2` before `Dwg 10`, case-insensitive). Previously plain ASCII sort.
- Organiser (`Organise pages & Preview...`): thumbnail grid in bundle order
  with file name, source page number and bundle order; generated cover /
  contents / separator pages shown as grey "GENERATED PAGE" placeholders;
  click / Ctrl / Shift selection; move earlier/later (within the document,
  clamped), rotate 90, exclude/restore, undo (40 steps), reset, zoom, page
  preview at 900 px, "Preview bundle (PDF)" (full build to a temp file),
  structure tree with move up/down and exclude/restore for whole documents
  or sections, and Create Bundle. Thumbnails render lazily for the visible
  rows on a worker thread through a bounded LRU cache (400 thumbnails, 6
  open documents) and stop when the window closes. Documents left with no
  pages are pruned, sections re-numbered, and contents/bookmarks target the
  first retained page of each document (tested).
- Not supported, by design: cross-document page interleaving (pages cannot
  leave their document) and drag-and-drop (button moves are exact and
  undoable). Per-folder output mode goes straight through the pipeline
  without the organiser.
- Totals are logged and shown on completion; when files were skipped the
  dialog is titled "PARTIAL" and lists each file with its reason.

## Behaviour changes to be aware of

- Old `separators: none` produced top-level pages; new `none` produces none.
  Choose `top_level` for the old result. Saved default is `all_sections`.
- Natural file-name ordering replaces ASCII ordering where no `_order.txt`
  exists.
- Previous bundle outputs found in the base folder are no longer re-bundled.
- Per-folder first-pages outputs are always plain (no cover/TOC/separators).

## Tests

```
python -m pytest tests                       # 56 engine/manifest/office tests
xvfb-run -a python3.12 -m pytest tests/test_gui_smoke.py   # 3 GUI smoke tests (needs tkinter + display)
```

Run here: 56 passed (Python 3.11, PyMuPDF 1.28.2, fpdf2 2.8.8) and 3 passed
(Python 3.12 under Xvfb). Coverage includes: page counts and link targets
for every separator scope with and without cover; multi-page TOC offsets;
bookmarks without TOC; collapsed / flattened / skipped / empty folders;
manual titles and `_order.txt`; corrupt / encrypted / empty inputs;
terminology; every preset; footer positions on A0 and rotated pages;
settings migration; overlay bundling; extraction modes; manifest move /
rotate / exclude / undo; per-folder first-pages incl. rerun protection;
thumbnail cache bounds; conversion stage with fake Word/Excel backends,
cancellation, missing Office and the companion subprocess contract.

## Not verifiable in this environment (Windows / Office)

The Linux container has no Microsoft Office or pywin32, so the COM paths
were written against the documented APIs and the existing Excel tool's
patterns but were not executed. Please run this checklist on Windows:

1. Double-click `document_bundle_builder.pyw` on a machine without PyMuPDF:
   first-run installer window appears and completes.
2. PDF-only tree, conversion unticked, no pywin32: builds normally.
3. Tick conversion with pywin32 missing: panel shows the install hint; build
   reports every Office file as failed and asks to continue.
4. Word: `.docx`, legacy `.doc`, password-protected `.docx`, a document open
   in another Word window. Expect PDFs in `_bundle_converted`, the protected
   one reported, no source modified, no WINWORD.EXE left running.
5. Excel: two workbooks in different folders with the same file name, one
   with a saved `.printcfg.json`. Expect the wizard to open pre-loaded with a
   companion banner, sidecar settings applied, no `_rev.NN` workbook copies,
   PDFs placed in their mirrored staging folders, and the bundle continuing
   after the wizard closes. Also close the wizard without exporting: the
   bundle must report the workbooks as cancelled, not hang.
6. Excel standalone: run `Excel_Quick_Print.pyw` without arguments; last-files
   list, footer, `_rev.NN` copy option and export behave as in v9.
7. A password-protected workbook in the companion run: expect an open error
   in the log, not an Excel prompt.
8. Fonts: Aptos/Segoe/Times embedding on Windows (Linux used core fonts).
9. `os.startfile` "Open output folder" / previews open correctly.

## Suggested future additions (not implemented)

1. Editable document register: per-document title aliases, revision and date
   fields shown in the contents and on separators (the manifest already
   carries per-document nodes, so this is a data and UI addition).
2. Saved reusable bundle manifests (`plan_to_manifest` exists; add save/load
   and re-application to a refreshed tree).
3. Duplicate detection (same content hash or same drawing number/revision)
   with a preflight warning.
4. Preflight report before build: page sizes, rotations, encrypted or
   damaged files, expected page total, estimated output size.
5. Drawing register extraction from title blocks (OCR/text-zone based) to
   populate drawing number, title and revision; explicitly not attempted.
6. Outlook `.msg`/`.eml` import (print-to-PDF via Outlook COM or a parser)
   feeding the same conversion stage.
7. Drag-and-drop in the organiser once a robust virtualised implementation
   is available; keyboard shortcuts for the existing operations.
8. Per-section footer/label overrides (e.g. drawings without footers inside
   a correspondence bundle with footers).
