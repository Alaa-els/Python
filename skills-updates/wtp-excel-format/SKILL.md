---
name: wtp-excel-format
description: Apply WT Partnership brand formatting to Excel workbooks (.xlsx, .xlsm) for Alaa Elsayed and any WTP-branded deliverable - cost assessments, IPC workbooks, comparison sheets, schedules, trackers, change management schedules, BOQ extracts, payment certificates, variation assessments, Engineer's Assessment Trackers, RFQ comparisons, claim bundles, any .xlsx/.xlsm output for issue. Triggers on phrases like "create the workbook", "build the spreadsheet", "format this Excel", "WTP format", "branded workbook", "apply our format". The skill enforces the strict WT brand colour palette (#0045BF blue, #FFC425 yellow, #EDF0F2 grey-blue), PT Sans typography, named cell styles, Cover and Appendix Cover tabs with linked textboxes from a Project Info source-of-truth tab, the legacyDrawingHF letterhead on portrait content tabs, table styling patterns, page setup conventions, and clean authorship metadata. Always runs alongside the file-authorship-metadata skill.
---

# WTP Excel Format

This skill produces Excel workbooks that match the WT Partnership brand standard issued by WTP India (`3.1.1 Excel guide.xlsm`), as interpreted for Alaa Elsayed's MBU Saudi Arabia work on Qiddiya D-18.

The skill replaces ad-hoc styling with a single repeatable pipeline. **Apply this skill to every workbook produced or modified for Alaa unless he explicitly says "no formatting" or "raw data only".**

## Step 0 — Standing exception (Alaa Elsayed)

For Alaa's account, defaults are:
- **Author / Last Modified By**: `Alaa Elsayed`
- **Company**: `WT Partnership`
- **Region**: KSA (MBU business unit)
- **Letterhead**: WTP Australia version (the only letterhead available until WTP issues a KSA-localised version). Use this and flag the Australian footer in the response.

Do not re-ask these unless he changes them.

## Step 1 — Brand palette (STRICT)

Only these colours appear in the workbook. Do not introduce new colours, shades, or accent variants. The yellow is reserved — never use it for body text, table fills, or anything other than the WT logo position and the cost-forecast highlight.

| Slot | Hex | Permitted use |
|---|---|---|
| White | `#FFFFFF` | Text on dark fills, page background |
| Black | `#000000` | Body text, document titles |
| Light grey-blue | `#EDF0F2` | Description column fill, section banner row fill |
| **WT Royal Blue** | `#0045BF` | Primary accent: heading text, table header fill, group banner fill, hyperlinks, subheading text |
| **WT Yellow** | `#FFC425` | RESERVED: WT logo only on cover/appendix; cost-forecast column border highlight; client-name accent on cover textbox |
| Deep Navy | `#051641` | Appendix cover background image only |
| Border grey 25% | `#BFBFBF` | Cell border tint (for description column borders) |
| Border grey 15% | `#D9D9D9` | Cell border tint (for inner table borders) |

If any other colour is requested, push back and remap to one of the above. Teal, dark teal, navy variants, and any greys outside the two listed are NOT in the WT signature.

## Step 2 — Font

PT Sans throughout. No exceptions. If PT Sans is unavailable on the recipient's machine, Excel will substitute (typically Calibri). Don't use Calibri / Arial / Aptos in the styles.xml — set `name="PT Sans"` everywhere.

## Step 3 — Named cell styles to register

Every workbook registers these styles (idempotent — check `wb.named_styles` first):

| Style name | Font | Size | Weight | Colour | Fill | Border | Alignment |
|---|---|---|---|---|---|---|---|
| `WT Title` | PT Sans | 22 | Bold | Black | none | none | Vert centre |
| `WT Heading 1` | PT Sans | 13 | Bold | Black | none | none | Vert centre |
| `WT Heading 2` | PT Sans | 13 | Bold | Royal Blue | none | none | Vert centre |
| `WT Heading 3` | PT Sans | 11 | Bold | Royal Blue | none | none | Vert centre |
| `WT Heading 4` | PT Sans | 10 | Bold | Black | none | none | Vert centre |
| `WT Headline` | PT Sans | 11 | Bold | Royal Blue | none | none | Left, top |
| `WT Table Heading` | PT Sans | 10 | Bold | White | Royal Blue | thin white all sides | Centre, wrap |
| `WT Table Heading 2` | PT Sans | 10 | Bold | Royal Blue | none | medium blue bottom | Centre, wrap |
| `WT Table Heading 3` | PT Sans | 10 | Bold | Royal Blue | Light grey-blue | thin grey 25% all sides | Centre, wrap |

Use `scripts/wtp_styles.py:register_named_styles(wb)` for the canonical implementation.

## Step 4 — Project Info tab (source of truth)

Every workbook has a `Project Info` tab as the first sheet, **visible**, with the yellow tab colour. Structure:

- Row 2: `PROJECT INFORMATION` styled `WT Title`
- Row 4: italic instruction note in blue (`#0045BF`)
- Rows 6-17: editable fields (label in column B, value in column C with `#EDF0F2` fill and grey 25% border)
- Rows 18-20: helper formulas (auto-computed, italic)

Standard fields (in this order, rows 6-17):
1. Project Name
2. Client
3. Address
4. Contract Reference
5. Contractor
6. Engineer
7. Cost Consultant
8. Document Title
9. Document Date
10. Document Reference
11. Appendix Letter
12. Appendix Description

Helper formulas (rows 18-20):
- C18: `=C13 & " | For works completed to " & C14` (Cover Subtitle)
- C19: `="Reference: " & C15` (Cover Reference)
- C20: `="Appendix " & C16` (Appendix Title)

Define workbook-level named ranges for each field: `PROJECT_NAME`, `CLIENT`, `ADDRESS`, `CONTRACT_REF`, `CONTRACTOR`, `ENGINEER`, `COST_CONSULT`, `DOC_TITLE`, `DOC_DATE`, `DOC_REF`, `APP_LETTER`, `APP_DESC` — pointing to the absolute cell references on Project Info.

Use `scripts/wtp_styles.py:build_project_info_tab(wb, fields_dict)`.

## Step 5 — Cover and Appendix Cover tabs

Both tabs are **hidden** (regular `Hide`, not `Very Hidden`) so they only appear in PDF export when the user explicitly unhides them. Both use the WT brand image as a full-bleed background and overlay textboxes containing pre-computed Project Info values.

### Cover tab
- Page setup: A4 portrait, all margins 0, fitToPage
- Background: `assets/cover_bg.jpg` anchored absolute (0, 0), size 7559675 × 10691814 EMU (full A4)
- Column widths matching WTP guide cover: A=12.27, B=11.18, C=9.54, D=11.54, E=11.45, F=8.82, G=14.54, H=14.18, I=1.45
- Six textboxes overlaid via direct XML injection (openpyxl strips raw `<sp>` elements on roundtrip):
  1. Project Name — 24pt, white, anchor B7:G9
  2. Client — 18pt, yellow, anchor B10:G11
  3. Address — 12pt, white, anchor B12:G13
  4. Contract Reference — 10pt, white, anchor B14:G15
  5. Document Subtitle (helper) — 12pt bold, white, anchor B17:G18
  6. Reference (helper) — 10pt, white, anchor B19:G19

### Appendix Cover tab
- Same page setup
- Background: `assets/appendix_bg.jpg`
- Three textboxes:
  1. "Appendix [letter]" — 36pt bold, white, anchor B7:G9
  2. Description — 14pt, yellow, anchor B11:G12
  3. Reference — 10pt, white, anchor B14:G14

### Textbox specification — IMPORTANT

Each textbox embeds **the actual current value from Project Info** as static text (not the placeholder word "Project Name"), AND keeps a `textlink` attribute pointing to the source cell. This dual approach is necessary because:
- Excel's `textlink` attribute is unreliable on externally-created files (textboxes show the static text on first open, ignoring textlink)
- Pre-computing values into the static text guarantees correct display immediately
- Keeping `textlink` enables the user to "activate" dynamic linking later by clicking each textbox in Excel and saving

Use `scripts/wtp_textboxes.py:inject_cover_textboxes(unpacked_dir, project_info_dict)`.

## Step 6 — Letterhead (legacyDrawingHF)

Apply `assets/letterhead.png` as a `legacyDrawingHF` first-page header on **portrait content tabs only**. Skip landscape tabs (the letterhead is portrait-oriented and would render rotated).

The mechanism:
1. Add `image[N].png` to `xl/media/`
2. Add `vmlDrawing[N].vml` to `xl/drawings/` referencing the image at full A4 size (595.5pt × 842.5pt)
3. Add the vmlDrawing rels file pointing to the media image
4. Add a relationship to the sheet's `_rels/sheet[N].xml.rels` with type `vmlDrawingHF`
5. Inject into the worksheet XML before `</worksheet>`:
   ```xml
   <headerFooter differentFirst="1">
     <firstHeader>&L&G</firstHeader>
   </headerFooter>
   <legacyDrawingHF xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdN"/>
   ```
6. Update `[Content_Types].xml` to include `Default Extension="png"` and `Default Extension="vml"` if not present

Use `scripts/wtp_letterhead.py:inject_letterhead(unpacked_dir, portrait_tab_names)`.

**Flag the Australian footer.** The bundled `letterhead.png` shows "WTP Australia Pty Ltd ACN 605 212 182 ABN 69 605 212 182". For real KSA-issued work, this image must be replaced with a WTP MBU artwork. Always mention this in the response when the workbook is for external issue.

## Step 7 — Page setup conventions

Per content tab:
- A4 (paperSize 9), portrait or landscape per the data shape
- `fitToWidth=1`, `fitToHeight=0` for portrait; `fitToPage=True` in sheet properties
- Hidden gridlines (`showGridLines=False`)
- Margins:
  - Portrait: L 0.354" / R 0.0" / T 1.323" / B 0.433" / header 0.315" / footer 0.197"
  - Landscape: L 0.354" / R 0.354" / T 0.5" / B 0.433" / header 0.315" / footer 0.197"

The 1.323" top margin on portrait tabs leaves space for the printed letterhead image.

## Step 8 — Table styling patterns

Apply per table on each content tab. Detection heuristics:
- **Title row** (usually row 1): apply `WT Title`, override font to 18pt for fit
- **Subtitle rows** (rows 2-4): apply `WT Heading 2` then `WT Heading 3`
- **Group banner row** (e.g. row 5 with multi-column groupings like "SAMA Initial Submission"): direct format — bold 11pt white text on royal blue fill, centred
- **Column header row** (typically row 6): apply `WT Table Heading` style (white on royal blue, white inner borders, wrap text, height 30-32pt)
- **Section subheading rows** (where column A or B contains all-uppercase text > 3 chars): merge across all columns, apply blue text on `#EDF0F2` fill, bold 11pt, left-aligned with indent 1, row height 20-22pt
- **Description column** (typically C): light grey-blue fill `#EDF0F2` on data rows, skip section banner rows
- **Cost forecast column** (the column representing WTP's recommended cost — typically named "Cost Amount", "Cost Assessment Amount", "Total Cost", or the variance column on bid comparisons): apply medium **YELLOW** (`#FFC425`) borders left/right around every data cell, skip section banner rows, with medium yellow top border on the first non-banner cell (header) and medium yellow bottom border on the last non-banner cell

Use `scripts/wtp_styles.py:apply_table_pattern(ws, header_row, description_col, forecast_col, banner_check_col)`.


## Step 8a — Navigation links and print budget (added 01-Oct-2026, from RFP-027 Rev 02)

- **Internal hyperlinks** carry the reading path on assessment workbooks: Royal Blue `#0045BF`, single underline, on the cell's existing font and size; stored as in-workbook locations (`'Build-Up'!A9`), never as file targets. LibreOffice recalculation rewrites them as external links with the local path: `contractor-proposal-assessment/scripts/finalize_xlsx.py` restores them in Stage B.
- **Row 5 of each working tab** holds the navigation line in 9 pt italic blue: "Back to 'Assessment'" in column B and "Forward to 'Programme'" (or the next tab) in the last column. Item total rows carry "Return to 'Assessment' item n" right-aligned in the last column.
- **Comparison columns** (previous revision, submitted programme) are shaded `#EDF0F2`; the column actually used keeps the yellow `#FFC425` left and right borders.
- **Print:** one page wide, as many pages tall as needed, on every tab. Never fit a tab to one page tall if that scales below about 45 per cent. Repeat the three title rows (`print_title_rows = '1:3'`); freeze below them. Place manual page breaks from the actual row heights (LibreOffice breaks on row heights, not text), budget about 820 pt a landscape page, repeat the table header after each break, never leave a header at the foot of a page or a single row at the top of the next.
- **Cover drawings:** openpyxl drops the cover images and text boxes on save. When a delivered file must be edited with openpyxl, copy the cover sheet XML, rels, drawings and media back from the source afterwards (`contractor-proposal-assessment/scripts/restore_drawings.py`).

## Step 9 — Tab order, visibility, naming

Final tab order:
1. `Project Info` (visible, yellow tab colour, active on open)
2. `Cover` (hidden)
3. `Appendix Cover` (hidden)
4. Working content tabs in logical workflow order
5. `Data` or reference tabs at the end

Set `wb.active = wb.sheetnames.index('Project Info')`.

## Step 10 — Build pipeline (CRITICAL — order matters)

openpyxl does not preserve raw XML on roundtrip. Specifically:
- Direct-injected `<sp>` text boxes are stripped
- Custom drawing XML extensions are stripped
- Some VML constructs are stripped

Therefore the build pipeline must be:

```
Stage A (openpyxl):  builds workbook with cells, named styles, page setup,
                     Project Info tab, Cover/Appendix shells with images,
                     defined names. Saves as v_a.xlsx.

Stage B (raw XML):   unpacks v_a.xlsx. Injects:
                       - Cover textboxes (with pre-computed Project Info values)
                       - Appendix textboxes
                       - legacyDrawingHF letterhead on portrait tabs
                       - WT theme XML (replaces theme1.xml)
                       - Clean authorship metadata
                     Repacks as v_final.xlsx.

Stage C (verify):    Run scripts/verify_metadata.py from file-authorship-metadata
                     skill. Confirm 0 AI/library traces.
                     Confirm textbox count and textlink references survived.
```

**Never run openpyxl save after Stage B.** Any further modifications must be done via direct XML editing in the unpacked tree, then repack.

Use the orchestrator `scripts/build_wtp_workbook.py:build(source_xlsx, project_info, output_path)` which handles all stages.

## Step 11 — File authorship & metadata

Always invoke the file-authorship-metadata skill in parallel. After Stage B:
- `docProps/core.xml`: creator and lastModifiedBy = `Alaa Elsayed`; created/modified = current UTC timestamp; title and subject populated
- `docProps/app.xml`: Application = `Microsoft Excel`; Company = `WT Partnership`; AppVersion = `16.0300`; clear AI/library traces
- Remove or sanitise any `docProps/custom.xml` containing tool references

Run `python /mnt/skills/user/file-authorship-metadata/scripts/verify_metadata.py <output_path>` and confirm exit code 0 before presenting.

## Step 12 — Pre-issue confirmation in response

After delivering the file, the response must include:
1. Single-line metadata confirmation
2. Cover/Appendix tab status (hidden by default — user must unhide for PDF export)
3. Print procedure (Active Sheets, not Entire Workbook)
4. Australian-letterhead flag if relevant
5. Outstanding sparring points (any styling decisions made on interpretation rather than explicit instruction)

## Bundled assets

```
assets/
  cover_bg.jpg          1654×2339 px, 200 dpi - WTP cover background
  appendix_bg.jpg       1654×2339 px, 200 dpi - WTP appendix cover background
  letterhead.png        794×1123 px - WTP first-page letterhead (Australian, FLAG TO USER)
  wt_theme.xml          OOXML theme1.xml with WT Custom 20 colour scheme + PT Sans
```

## Bundled scripts

```
scripts/
  wtp_styles.py         Named cell styles, page setup, table patterns, Project Info tab
  wtp_textboxes.py      Direct XML injection for Cover/Appendix textboxes with textlink
  wtp_letterhead.py     Direct XML injection for legacyDrawingHF on portrait tabs
  wtp_finalize.py       Theme XML injection + metadata cleanup
  build_wtp_workbook.py End-to-end orchestrator
```

## When NOT to apply this skill

- User explicitly says "raw data only" or "no formatting"
- The workbook is a non-WTP deliverable (e.g., personal calculation, third-party template that must remain unchanged)
- User is asking for a quick calculation / scratch sheet not for issue
- User asks for CSV output

When uncertain, ask: "Apply WTP format?" with a clear yes/no option.

## Anti-patterns to avoid

- **Do not** use cell fills to draw the brand cover composition. Use the actual `cover_bg.jpg` image with textbox overlays.
- **Do not** use deep navy `#051641` for table headers — use Royal Blue `#0045BF`.
- **Do not** use teal, dark teal, or any non-listed colour. The signature is tight on purpose.
- **Do not** apply the cost-forecast border in blue — it must be **yellow** to stand out from the blue-dominated headings.
- **Do not** save with openpyxl after Stage B — it strips the textboxes and letterhead.
- **Do not** ship without verifying metadata is clean.
- **Do not** describe the WTP cover image in cell-fill terms in the response — refer to the actual image.

## Reference: where this came from

Brand source: `WTP India 3.1.1 Excel guide.xlsm` (Shivani Saw, May 2026), forwarded by Ayo Ajayi to Alberto / Matthew / Christopher Legg. Interpretation refined through iterative review with Alaa over the QPAC Laydown Area Access Roads workbook (May 2026). Key style decisions:
- Royal blue (not deep navy) for table headers — explicit Alaa correction
- Subheading text in royal blue — explicit Alaa correction
- Description column light grey-blue fill — explicit Alaa correction
- Cost-forecast border in **yellow** (not blue) — explicit Alaa correction
- Cell-based covers REJECTED — must use brand images with textbox overlays
- Cleaned palette to remove unused theme colours — explicit Alaa direction
