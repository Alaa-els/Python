"""
WTP Excel Format - Style helpers.

Provides:
  WT (palette dict)
  FONT
  register_named_styles(wb)
  build_project_info_tab(wb, fields=None, position=0) -> dict of pre-computed values
  define_named_ranges(wb)
  apply_wtp_page_setup(ws, orientation='portrait')
  apply_table_pattern(ws, header_row, description_col=None, forecast_col=None,
                      banner_check_col='A', last_data_row=None)
  add_brand_image_cover_shell(wb, name, image_path, position)
"""
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from copy import copy
import os

# ------------------------------------------------------------
# Brand palette (STRICT - do not extend without Alaa's approval)
# ------------------------------------------------------------
WT = {
    'white':     'FFFFFF',
    'black':     '000000',
    'lt2':       'EDF0F2',  # description column, section banner fill
    'blue':      '0045BF',  # accent1 - PRIMARY brand colour
    'yellow':    'FFC425',  # accent6 - WT logo, cost-forecast border highlight
    'deep_navy': '051641',  # accent5 - appendix bg image only
    'grey_25':   'BFBFBF',
    'grey_15':   'D9D9D9',
}
FONT = 'PT Sans'

# Border sides used in styling
_thin_white = Side(style='thin', color=WT['white'])
_thin_grey25 = Side(style='thin', color=WT['grey_25'])
_thin_grey15 = Side(style='thin', color=WT['grey_15'])
_medium_blue = Side(style='medium', color=WT['blue'])
_medium_yellow = Side(style='medium', color=WT['yellow'])


# ------------------------------------------------------------
# Named cell styles
# ------------------------------------------------------------
def _make_styles():
    styles = []
    s = NamedStyle(name='WT Title')
    s.font = Font(name=FONT, size=22, bold=True, color=WT['black'])
    s.alignment = Alignment(vertical='center')
    styles.append(s)

    s = NamedStyle(name='WT Heading 1')
    s.font = Font(name=FONT, size=13, bold=True, color=WT['black'])
    s.alignment = Alignment(vertical='center')
    styles.append(s)

    s = NamedStyle(name='WT Heading 2')
    s.font = Font(name=FONT, size=13, bold=True, color=WT['blue'])
    s.alignment = Alignment(vertical='center')
    styles.append(s)

    s = NamedStyle(name='WT Heading 3')
    s.font = Font(name=FONT, size=11, bold=True, color=WT['blue'])
    s.alignment = Alignment(vertical='center')
    styles.append(s)

    s = NamedStyle(name='WT Heading 4')
    s.font = Font(name=FONT, size=10, bold=True, color=WT['black'])
    s.alignment = Alignment(vertical='center')
    styles.append(s)

    s = NamedStyle(name='WT Headline')
    s.font = Font(name=FONT, size=11, bold=True, color=WT['blue'])
    s.alignment = Alignment(horizontal='left', vertical='top')
    styles.append(s)

    s = NamedStyle(name='WT Table Heading')
    s.font = Font(name=FONT, size=10, bold=True, color=WT['white'])
    s.fill = PatternFill('solid', fgColor=WT['blue'])
    s.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    s.border = Border(left=_thin_white, right=_thin_white,
                      top=_thin_white, bottom=_thin_white)
    styles.append(s)

    s = NamedStyle(name='WT Table Heading 2')
    s.font = Font(name=FONT, size=10, bold=True, color=WT['blue'])
    s.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    s.border = Border(bottom=Side(style='medium', color=WT['blue']))
    styles.append(s)

    s = NamedStyle(name='WT Table Heading 3')
    s.font = Font(name=FONT, size=10, bold=True, color=WT['blue'])
    s.fill = PatternFill('solid', fgColor=WT['lt2'])
    s.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    s.border = Border(left=_thin_grey25, right=_thin_grey25,
                      top=_thin_grey25, bottom=_thin_grey25)
    styles.append(s)
    return styles


def register_named_styles(wb):
    """Idempotent registration of all WT named cell styles."""
    for s in _make_styles():
        if s.name not in wb.named_styles:
            wb.add_named_style(s)


# ------------------------------------------------------------
# Project Info tab
# ------------------------------------------------------------
DEFAULT_FIELDS = [
    ('Project Name',         ''),
    ('Client',               ''),
    ('Address',              ''),
    ('Contract Reference',   ''),
    ('Contractor',           ''),
    ('Engineer',             ''),
    ('Cost Consultant',      'WT Partnership'),
    ('Document Title',       ''),
    ('Document Date',        ''),
    ('Document Reference',   ''),
    ('Appendix Letter',      'A'),
    ('Appendix Description', ''),
]

NAMED_RANGE_KEYS = [
    'PROJECT_NAME', 'CLIENT', 'ADDRESS', 'CONTRACT_REF',
    'CONTRACTOR', 'ENGINEER', 'COST_CONSULT', 'DOC_TITLE',
    'DOC_DATE', 'DOC_REF', 'APP_LETTER', 'APP_DESC',
]

def build_project_info_tab(wb, fields=None, position=0):
    """
    Build Project Info tab. fields is a list of (label, value) tuples.
    Returns dict mapping cell coords to pre-computed values for textbox embedding.
    """
    if fields is None:
        fields = DEFAULT_FIELDS

    if 'Project Info' in wb.sheetnames:
        del wb['Project Info']
    pi = wb.create_sheet('Project Info', position)
    pi.sheet_view.showGridLines = False

    # Title
    pi['B2'] = 'PROJECT INFORMATION'
    pi['B2'].style = 'WT Title'
    pi.row_dimensions[2].height = 36

    pi['B4'] = ('Edit values in column C. Cover and Appendix tabs display these values. '
                "After editing, click each Cover textbox and type the formula "
                "='Project Info'!$C$N in the formula bar to make it dynamic in Excel.")
    pi['B4'].font = Font(name=FONT, size=10, italic=True, color=WT['blue'])
    pi['B4'].alignment = Alignment(wrap_text=True)
    pi.row_dimensions[4].height = 30

    # Editable fields
    for i, (label, default) in enumerate(fields):
        r = 6 + i
        pi.cell(r, 2, label).font = Font(name=FONT, size=10, bold=True, color=WT['black'])
        pi.cell(r, 2).alignment = Alignment(vertical='center')
        pi.cell(r, 3, default).font = Font(name=FONT, size=10, color=WT['black'])
        pi.cell(r, 3).fill = PatternFill('solid', fgColor=WT['lt2'])
        pi.cell(r, 3).alignment = Alignment(vertical='center', wrap_text=True)
        pi.cell(r, 3).border = Border(
            left=_thin_grey25, right=_thin_grey25,
            top=_thin_grey25, bottom=_thin_grey25,
        )
        pi.row_dimensions[r].height = 22

    # Helper formulas (rows 18-20)
    helpers = [
        (18, 'Cover Subtitle (auto)',  '=C13 & " | For works completed to " & C14'),
        (19, 'Cover Reference (auto)', '="Reference: " & C15'),
        (20, 'Appendix Title (auto)',  '="Appendix " & C16'),
    ]
    for r, label, formula in helpers:
        pi.cell(r, 2, label).font = Font(name=FONT, size=10, bold=True, italic=True, color=WT['blue'])
        pi.cell(r, 2).alignment = Alignment(vertical='center')
        pi.cell(r, 3, formula).font = Font(name=FONT, size=10, italic=True, color=WT['black'])
        pi.cell(r, 3).fill = PatternFill('solid', fgColor=WT['lt2'])
        pi.cell(r, 3).alignment = Alignment(vertical='center', wrap_text=True)
        pi.cell(r, 3).border = Border(
            left=_thin_grey25, right=_thin_grey25,
            top=_thin_grey25, bottom=_thin_grey25,
        )
        pi.row_dimensions[r].height = 22

    pi.column_dimensions['A'].width = 2
    pi.column_dimensions['B'].width = 28
    pi.column_dimensions['C'].width = 75
    pi.sheet_properties.tabColor = WT['yellow']

    # Compute pre-computed values dict for textbox embedding
    pi_values = {}
    for i, (_label, default) in enumerate(fields):
        r = 6 + i
        pi_values[f'C{r}'] = default
    # Compute helpers
    pi_values['C18'] = f"{pi_values.get('C13', '')} | For works completed to {pi_values.get('C14', '')}"
    pi_values['C19'] = f"Reference: {pi_values.get('C15', '')}"
    pi_values['C20'] = f"Appendix {pi_values.get('C16', '')}"

    return pi_values


def define_named_ranges(wb):
    """Define workbook-level named ranges for Project Info cells."""
    ranges = {
        'PROJECT_NAME':  "'Project Info'!$C$6",
        'CLIENT':        "'Project Info'!$C$7",
        'ADDRESS':       "'Project Info'!$C$8",
        'CONTRACT_REF':  "'Project Info'!$C$9",
        'CONTRACTOR':    "'Project Info'!$C$10",
        'ENGINEER':      "'Project Info'!$C$11",
        'COST_CONSULT':  "'Project Info'!$C$12",
        'DOC_TITLE':     "'Project Info'!$C$13",
        'DOC_DATE':      "'Project Info'!$C$14",
        'DOC_REF':       "'Project Info'!$C$15",
        'APP_LETTER':    "'Project Info'!$C$16",
        'APP_DESC':      "'Project Info'!$C$17",
    }
    for name, ref in ranges.items():
        if name in wb.defined_names:
            del wb.defined_names[name]
        wb.defined_names[name] = DefinedName(name=name, attr_text=ref)


# ------------------------------------------------------------
# Page setup
# ------------------------------------------------------------
WTP_MARGINS_PORTRAIT = PageMargins(left=0.354, right=0.0, top=1.323, bottom=0.433,
                                   header=0.315, footer=0.197)
WTP_MARGINS_LANDSCAPE = PageMargins(left=0.354, right=0.354, top=0.5, bottom=0.433,
                                    header=0.315, footer=0.197)

def apply_wtp_page_setup(ws, orientation='portrait'):
    """A4, hidden gridlines, fitToWidth, WTP margins."""
    is_landscape = orientation == 'landscape'
    ws.page_setup.orientation = 'landscape' if is_landscape else 'portrait'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = WTP_MARGINS_LANDSCAPE if is_landscape else WTP_MARGINS_PORTRAIT
    ws.sheet_view.showGridLines = False


# ------------------------------------------------------------
# Brand cover shells
# ------------------------------------------------------------
def add_brand_image_cover_shell(wb, name, image_path, position):
    """
    Add Cover or Appendix Cover sheet with full-bleed brand image background.
    Textboxes injected separately via wtp_textboxes.py at Stage B.
    """
    if name in wb.sheetnames:
        del wb[name]
    sh = wb.create_sheet(name, position)
    sh.sheet_view.showGridLines = False
    sh.page_setup.orientation = 'portrait'
    sh.page_setup.paperSize = 9
    sh.page_margins = PageMargins(left=0, right=0, top=0, bottom=0, header=0, footer=0)
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 1
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    widths = {'A':12.27,'B':11.18,'C':9.54,'D':11.54,'E':11.45,
              'F':8.82,'G':14.54,'H':14.18,'I':1.45}
    for col, w in widths.items():
        sh.column_dimensions[col].width = w

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Brand image missing: {image_path}")
    img = XLImage(image_path)
    img.anchor = AbsoluteAnchor(
        pos=XDRPoint2D(x=0, y=0),
        ext=XDRPositiveSize2D(cx=7559675, cy=10691814),
    )
    sh.add_image(img)
    sh.print_area = 'A1:I54'
    return sh


# ------------------------------------------------------------
# Table patterns (apply to existing data)
# ------------------------------------------------------------
def _is_section_banner(ws, row, banner_check_col):
    val = ws[f'{banner_check_col}{row}'].value
    return (val and isinstance(val, str) and len(val) > 3 and val.isupper())


def fill_description_column(ws, col_letter, start_row, end_row, banner_check_col='A'):
    """Apply lt2 grey-blue fill to description column data cells."""
    fill = PatternFill('solid', fgColor=WT['lt2'])
    for r in range(start_row, end_row + 1):
        if not _is_section_banner(ws, r, banner_check_col):
            ws[f'{col_letter}{r}'].fill = copy(fill)


def highlight_forecast_column(ws, col_letter, start_row, end_row, banner_check_col='A'):
    """Add medium YELLOW (#FFC425) borders bracketing the cost forecast column."""
    for r in range(start_row, end_row + 1):
        if _is_section_banner(ws, r, banner_check_col):
            continue
        cell = ws[f'{col_letter}{r}']
        existing = cell.border
        cell.border = Border(
            left=_medium_yellow, right=_medium_yellow,
            top=copy(existing.top) if existing.top.style else _thin_grey15,
            bottom=copy(existing.bottom) if existing.bottom.style else _thin_grey15,
        )
    # Yellow top on first non-banner cell
    for r in range(start_row, end_row + 1):
        if not _is_section_banner(ws, r, banner_check_col):
            cell = ws[f'{col_letter}{r}']
            cell.border = Border(left=_medium_yellow, right=_medium_yellow,
                                 top=_medium_yellow,
                                 bottom=copy(cell.border.bottom))
            break
    # Yellow bottom on last non-banner cell
    for r in range(end_row, start_row - 1, -1):
        if not _is_section_banner(ws, r, banner_check_col):
            cell = ws[f'{col_letter}{r}']
            cell.border = Border(left=_medium_yellow, right=_medium_yellow,
                                 top=copy(cell.border.top),
                                 bottom=_medium_yellow)
            break


def style_section_banners(ws, max_col, start_row, end_row, banner_check_col='A'):
    """Merge banner rows across columns and apply blue text on grey fill."""
    for r in range(start_row, end_row + 1):
        if _is_section_banner(ws, r, banner_check_col):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
            cell = ws.cell(r, 1)
            cell.font = Font(name=FONT, size=11, bold=True, color=WT['blue'])
            cell.fill = PatternFill('solid', fgColor=WT['lt2'])
            cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            ws.row_dimensions[r].height = 22


def apply_table_pattern(ws, header_row, max_col, description_col=None,
                        forecast_col=None, banner_check_col='A',
                        last_data_row=None):
    """
    Apply WTP table pattern in one call.
    header_row: row index of column headers (gets WT Table Heading style)
    max_col: number of columns to apply patterns across (1-indexed)
    description_col: column letter to fill with grey (e.g. 'C')
    forecast_col: column letter to bracket with yellow borders (e.g. 'Q', 'I')
    banner_check_col: column to scan for uppercase section banners
    last_data_row: end of data range (defaults to ws.max_row)
    """
    if last_data_row is None:
        last_data_row = ws.max_row

    # Header row
    for c in range(1, max_col + 1):
        cell = ws.cell(header_row, c)
        if cell.value is not None or c <= max_col:
            cell.style = 'WT Table Heading'
    ws.row_dimensions[header_row].height = 32

    # Section banners
    style_section_banners(ws, max_col, header_row + 1, last_data_row, banner_check_col)

    # Description column
    if description_col:
        fill_description_column(ws, description_col, header_row + 1,
                                last_data_row, banner_check_col)

    # Forecast column (yellow bracket)
    if forecast_col:
        highlight_forecast_column(ws, forecast_col, header_row,
                                  last_data_row, banner_check_col)
