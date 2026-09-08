# Document Bundle Builder (formerly Claims Bundle Builder)

Merges a folder tree of PDFs (and, optionally, Word/Excel files converted to
PDF) into one organised, bookmarked bundle with a clickable table of contents.

## Files to deploy (keep in one folder)

| File | Purpose |
|---|---|
| `document_bundle_builder.pyw` | Double-click launcher and GUI (first run installs PyMuPDF and fpdf2) |
| `bundle_engine.py` | Build engine (planning, page manifest, cover/TOC/separators, footers) |
| `bundle_organiser.py` | Pre-build page organiser window |
| `office_convert.py` | Word/Excel conversion stage and Excel Quick Print companion adapter |
| `Excel_Quick_Print.pyw` | Excel Quick Print v10 (standalone wizard, also used in companion mode) |

`original/` holds the two source files exactly as received. `tests/` holds the
regression tests (`python -m pytest tests`; the GUI tests need a display).

Settings are stored in `~/.claims_bundle_builder.json` (same file as before, new
keys added with backward-compatible defaults). Converted Office files are kept
in `<base>/_bundle_converted/` and reused while newer than their source.

## Workflow

1. Pick the base directory and a bundle title.
2. Choose which pages to take from each file (all / first page / first N /
   custom ranges) and whether to produce one bundle or one
   `<Folder>#Combined_FirstPages.pdf` per folder.
3. Set terminology (Appendix / Section / Drawing / ...), separator scope,
   contents / bookmarks, footers, cover and style (or apply a preset).
4. **Create Bundle** builds directly; **Organise pages & Preview...** opens
   the page organiser first (select, move, rotate, exclude, undo, preview)
   and its **Create Bundle** button builds exactly the organised manifest.

See `CHANGES.md` for the full change report, the Windows/Office test checklist
and the list of suggested future additions.
