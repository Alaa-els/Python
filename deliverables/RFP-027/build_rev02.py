"""Rev 02 build - re-basing the RFP-027 assessment to the Contractor's programme of 29-Sep-2026.
Stage A: openpyxl on Alaa's master (Rev 01, 17-Sep-2026). Stage B (separate): recalc, zip-level metadata."""
import copy, datetime as dt, re, sys
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.pagebreak import Break
import math
PAGE = 820.0   # printable height of one landscape page at the fit-to-width scale, in points (measured on the rendered file)
def est(text, width, sz=10):
    cpl = max(8, width * 1.15 * 10 / sz)
    return math.ceil(len(text) / cpl) * (sz * 1.28) + 4

def autofit(ws, min_row=1, max_row=None):
    """Raise a row's height where wrapped text would not fit (PT Sans 10 pt: about 1.15 characters per width unit, 12.8 pt per line)."""
    merged = {}
    for rg in ws.merged_cells.ranges:
        c1, r1, c2, r2 = rg.bounds
        if r1 == r2:
            merged[(r1, c1)] = sum((ws.column_dimensions[get_column_letter(c)].width or 8.66) for c in range(c1, c2 + 1))
    for row in ws.iter_rows(min_row=min_row, max_row=max_row):
        for c in row:
            if not isinstance(c.value, str) or c.value.startswith('=') or not c.value.strip(): continue
            if not (c.alignment and c.alignment.wrap_text): continue
            width = merged.get((c.row, c.column), ws.column_dimensions[get_column_letter(c.column)].width or 8.66)
            sz = (c.font.sz or 10)
            cpl = max(8, width * 1.15 * 10 / sz)
            lines = sum(max(1, math.ceil(len(p) / cpl)) for p in c.value.split('\n'))
            need = lines * (sz * 1.28) + 4
            cur = ws.row_dimensions[c.row].height
            if cur is None or need > cur:
                ws.row_dimensions[c.row].height = round(need, 1)
SRC = 'src/master_rev01.xlsx'
OUT = 'stageA.xlsx'
REV = 'Rev 02'
DOCDATE = '30-Sep-2026'
XER = 'QC05958-BSL01THF-TSE.xer'
WK = '"0000100"'   # Friday non-working (Mon..Sun string)

wb = openpyxl.load_workbook(SRC)
D = dt.date

def cp(src, dst):
    dst._style = copy.copy(src._style)

def shift_refs(wb, sheet, at, n):
    """After inserting n rows at row 'at' on sheet, bump every formula reference to that sheet's rows >= at."""
    def fix(formula, same_sheet):
        def rep(m):
            col, dollar, row = m.group(2), m.group(3), int(m.group(4))
            if row >= at:
                row += n
            return f"{m.group(1)}{col}{dollar}{row}"
        # references qualified with the sheet name
        pat = re.compile(r"('" + re.escape(sheet) + r"'!\$?)([A-Z]{1,3})(\$?)(\d+)")
        formula = pat.sub(rep, formula)
        if same_sheet:
            # unqualified refs on the same sheet (avoid touching qualified ones already handled)
            def rep2(m):
                if m.group(1):   # preceded by ! or a letter -> qualified/other, skip
                    return m.group(0)
                col, dollar, row = m.group(2), m.group(3), int(m.group(4))
                if row >= at:
                    row += n
                return f"{col}{dollar}{row}"
            formula = re.sub(r"(['!A-Za-z])?\$?\b([A-Z]{1,3})(\$?)(\d+)\b", lambda m: m.group(0) if m.group(1) else rep2(m), formula)
        return formula
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('=') and (sheet in c.value or ws.title == sheet):
                    c.value = fix(c.value, ws.title == sheet)

def insert_rows_keep_styles(ws, at, n, style_row):
    """Insert n rows at 'at', re-merge ranges below, copy styles from style_row (pre-insert index)."""
    # move merged ranges
    merged = [str(r) for r in ws.merged_cells.ranges]
    for r in merged:
        ws.merged_cells.remove(r)
    ws.insert_rows(at, n)
    for r in merged:
        c1, r1, c2, r2 = openpyxl.utils.cell.range_boundaries(r)
        if r1 >= at:
            r1 += n; r2 += n
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    # row heights below shift
    heights = {k: v.height for k, v in ws.row_dimensions.items() if v.height and k >= at}
    for k in sorted(heights, reverse=True):
        ws.row_dimensions[k + n].height = heights[k]
        ws.row_dimensions[k].height = None
    src_row = style_row + n if style_row >= at else style_row
    for i in range(n):
        for col in range(1, ws.max_column + 1):
            cp(ws.cell(src_row, col), ws.cell(at + i, col))

# ---------------------------------------------------------------- Project Info
pi = wb['Project Info']
pi['C13'] = f"Cost Assessment of the Contractor's Proposal - TSE Irrigation Storage Tanks (RFP-027) - {REV}"
pi['C14'] = DOCDATE

# ---------------------------------------------------------------- Build-Up: insert row for 7.23 (FAT) at 122
bu = wb['Build-Up']
shift_refs(wb, 'Build-Up', 122, 1)
insert_rows_keep_styles(bu, 122, 1, 121)
bu['A122'] = '7.23'
bu['B122'] = 'Third-party factory acceptance test of the GRP panels - inspector days and travel to the factory'
bu['C122'] = 'item'
bu['D122'] = 0
bu['E122'] = 0
bu['F122'] = '=D122*E122'
bu['G122'] = ("Not assessed. The Contractor is arranging a third-party factory acceptance test (PQD, ITP, procedure and inspector CV "
              "submitted 26 to 28-Sep-2026 per its deliverables list of 29-Sep-2026). RFP-027 Scope of Works 5.3 requires hydrostatic "
              "factory test reports among the handover documents; that is the manufacturer's own test record, within the supplier's supply, "
              "and does not require a third-party inspector. The line is carried at nil until an Engineer's instruction or an approved ITP "
              "requiring third-party inspection is evidenced; the Contractor has not claimed it separately")
bu.row_dimensions[122].height = 78
# ---------------------------------------------------------------- Build-Up: insert rows for 7.24 (Tank 1 operating fill) and 7.25 (spray disinfection) at 123-124
shift_refs(wb, 'Build-Up', 123, 2)
insert_rows_keep_styles(bu, 123, 2, 121)
bu['A123'] = '7.24'
bu['B123'] = 'Tank 1 operating volume for the integrated commissioning - tankered fill to about 1 m depth after the test water has gone to Tank 2'
bu['C123'] = 'm3'
bu['D123'] = 1020   # replaced by the Programme link in Section 5
bu['E123'] = '=E100'
bu['F123'] = '=D123*E123'
bu['G123'] = ("Quantity: 34 m x 30 m x 1.0 m = 1,020 m3, an assessed operating depth for the witnessed pumping demonstration from Tank 1 (RFP Scope of Works 5.2, "
              "integrated system commissioning) - the Engineer decides the depth; Tank 1 is empty after its test water is pumped to Tank 2. Filled by tanker on 09 and "
              "10-Dec-2026 while Tank 2 is on its hold ('Programme' tab, window P20b), attended by the approved histogram helpers (duty H11), so no extra labour and no "
              "programme effect. A network refill after the tie-in is not assumed: no source, cost or timing has been confirmed by the Employer. If the Engineer "
              "requires Tank 1 at full test level for the commissioning, see sensitivity S8 on the 'Programme' tab. Rate: as 7.1 - needs confirmation of the water source")
bu['A124'] = '7.25'
bu['B124'] = 'Disinfection of the tank surfaces above the test water line (walls above 3.7 m, roof underside, hatches) - spray or swab application, both tanks'
bu['C124'] = 'tank'
bu['D124'] = 2
bu['E124'] = 2000
bu['F124'] = '=D124*E124'
bu['G124'] = ("Quantity: 2 tanks. The chlorinated test water (7.6, 7.7) disinfects only the surfaces it touches, up to the 3.7 m test level; the surfaces above it are not "
              "covered by that method and need a separate spray or swab application (AWWA C652 surface-application method). Tank 2 on 03 and 05-Dec-2026 before the "
              "transfer, Tank 1 on 08-Dec after it is emptied. Rate: assessment allowance per tank - 2 people x 2 days, sprayer hire, hypochlorite solution and "
              "confined-space attendance - needs confirmation by the Contractor's method statement; if the histogram helpers do the work, SAR 600 per tank is "
              "already in 5.10 and falls away here")
bu.row_dimensions[123].height = 105
bu.row_dimensions[124].height = 92
bu['B125'] = 'TOTAL ITEM 7 - both tanks'
bu['F125'] = '=SUM(F100:F124)'
bu['F126'] = '=F125/2'
# ---------------------------------------------------------------- Build-Up: insert line 2.5 (method statements and plans) at 41
shift_refs(wb, 'Build-Up', 41, 1)
insert_rows_keep_styles(bu, 41, 1, 40)
bu['A41'] = '2.5'
bu['B41'] = "Method statements, risk assessments and project plans - preparation by the Contractor's engineer (tank installation, pipe installation and tie-in method statements; HSE, CEMP, quality, execution and mobilisation plans; ITP)"
bu['C41'] = 'day'
bu['D41'] = 6
bu['E41'] = 1000
bu['F41'] = '=D41*E41'
bu['G41'] = ("Quantity: 6 engineer-days, an assessed allowance for the 20 preparation and approval activities in the programme (QCD18TSEMOB1040 to QCD18TSEMOB1230); "
             "the dismantling method statement is inside the Item 3 quotation, and the supplier's own installation documents are within Item 6 where it provides them "
             "(Assumption 9 is silent on documents). Administration of the submittals is 2.4; the site engineer (1.1) and HSE officer (1.2) review and implement them. "
             "Rate: assessed allowance (Riyadh market, Sep-2026) - needs confirmation by the Contractor")
bu.row_dimensions[41].height = 80
bu['F42'] = '=SUM(F37:F41)'

# ---------------------------------------------------------------- Programme tab
pg = wb.create_sheet('Programme', index=wb.sheetnames.index('Build-Up Comparison') + 1)
widths = {'A': 6, 'B': 46, 'C': 12, 'D': 12, 'E': 11, 'F': 14, 'G': 20, 'H': 15, 'I': 14, 'J': 52}
for k, v in widths.items():
    pg.column_dimensions[k].width = v
pg.sheet_view.showGridLines = False
pg.page_setup.orientation = 'landscape'
pg.page_setup.paperSize = 9
pg.page_setup.fitToWidth = 1
pg.page_setup.fitToHeight = 0
pg.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
pg.page_margins = copy.copy(wb['Assessment'].page_margins)
pg.print_title_rows = None

S_TITLE, S_SUB, S_SUB2, S_INTRO = bu['A1'], bu['A2'], bu['A3'], bu['A4']
S_BANNER, S_HDR = bu['A6'], bu['A8']
S_REF, S_DESC, S_UNIT, S_QTY, S_RATE, S_AMT, S_BASIS = bu['A9'], bu['B9'], bu['C9'], bu['D9'], bu['E9'], bu['F9'], bu['G9']
S_TOTLBL, S_TOTAMT = bu['B32'], bu['F32']
S_NOTE = bu['A7']

def banner(ws, r, text, ncol=10):
    ws.cell(r, 1, text)
    for c in range(1, ncol + 1):
        cp(S_BANNER, ws.cell(r, c))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    ws.row_dimensions[r].height = 21

def para(ws, r, text, ncol=10, height=None, bold=False):
    ws.cell(r, 1, text)
    for c in range(1, ncol + 1):
        cp(S_NOTE, ws.cell(r, c))
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    if bold:
        f = copy.copy(ws.cell(r, 1).font); f = Font(name=f.name, sz=f.sz, b=True, color=f.color); ws.cell(r, 1).font = f
    mw = sum((ws.column_dimensions[get_column_letter(c)].width or 8.66) for c in range(1, ncol + 1))
    cpl = max(8, mw * 1.15)
    need = sum(max(1, math.ceil(len(p) / cpl)) for p in text.split('\n')) * 12.8 + 6
    ws.row_dimensions[r].height = max(height or 15, need)

def header(ws, r, labels, merge_ij=False):
    for i, t in enumerate(labels, 1):
        c = ws.cell(r, i, t); cp(S_HDR, c)
    if merge_ij:
        cp(S_HDR, ws.cell(r, 10)); ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    need = max(est(t, 0.8 * (ws.column_dimensions[get_column_letter(i)].width or 8.66)) for i, t in enumerate(labels, 1) if t)   # bold headers wrap on word boundaries
    ws.row_dimensions[r].height = max(32, min(70, need)); PG['hdr'] = ws.row_dimensions[r].height

class _PG(dict):
    """'used' is the height already on the current Programme page: title rows plus every row since the last break, after wrapping."""
    def __getitem__(self, k):
        if k == 'used':
            lb = max([b.id for b in pg.row_breaks.brk], default=0)
            start, end = lb + 1, globals()['r']
            if end > start: autofit(pg, start, end - 1)
            return (58.0 if lb else 0.0) + sum((pg.row_dimensions[i].height or 14.4) for i in range(start, end))
        return dict.__getitem__(self, k)
PG = _PG(hdr=32.0)
def pg_add(h): pass
def pg_break(ws, r):
    ws.row_breaks.append(Break(id=r - 1)); PG['used'] = 58.0   # the repeated title rows 1 to 3
def pg_break_if(ws, r, need):
    if PG['used'] + need > PAGE: pg_break(ws, r)
HDR3 = ['Ref', 'Programme window', 'Start - SAMA Submitted Programme', 'Finish - SAMA Submitted Programme', 'Days - SAMA Submitted Programme', 'Start - Assessed', 'Finish - Assessed', 'Days - Assessed (used)', 'Where the dates come from and why']
HDR4 = ['Ref', "'Build-Up' line", 'Unit', 'Rate (SAR)', 'Qty - Rev 01 (superseded)', 'Qty - SAMA Submitted Programme (comparison)', 'Qty - Assessed (used on the Build-Up)', 'Amount - SAMA Submitted Programme (comparison, SAR)', 'Amount - Assessed (SAR)', 'Where the quantity comes from']
HDR5 = ['No.', 'Activity (as named in the programme)', 'Start', 'Finish', 'Working days', 'Cost loaded (SAR, incl. 5% OHP)', 'Activity ID', 'Calendar', 'Total float (days)', 'Programme section (as named)']
def hdr4(ws, r):
    header(ws, r, HDR4[:9]); ws.cell(r, 10, HDR4[9]); cp(S_HDR, ws.cell(r, 10))

def datecell(c):
    cp(S_QTY, c); c.number_format = 'dd-mmm-yyyy'; c.alignment = Alignment(horizontal='center', vertical='center')

def numcell(c, fmt='#,##0'):
    cp(S_QTY, c); c.number_format = fmt

pg['A1'] = "THE CONTRACTOR'S PROGRAMME - ACTIVITY DATA AND DURATION BASES"; cp(S_TITLE, pg['A1']); pg.row_dimensions[1].height = 25.5
pg['A2'] = 'TSE Irrigation Storage Tanks and Associated Pipeworks (RFP-027) - Contract QPMO-410-CT-05958'; cp(S_SUB, pg['A2']); pg.row_dimensions[2].height = 17.4
pg['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'; cp(S_SUB2, pg['A3'])
para(pg, 4, ("How this tab works\n"
    "• 'XER WBS' tab: SAMA's programme as received. Section 1: programme status. Section 2: working calendar.\n"
    "• Section 3: programme windows - SAMA Submitted Programme dates beside the Assessed dates (same erection dates, one-fill sequential testing).\n"
    "• Section 4: the approved manpower histogram week by week, helper duties day by day, skilled people by activity, plant days, site staff by phase.\n"
    "• Section 5: every 'Build-Up' quantity taken from this tab. Section 6: every 'Build-Up' line linked to its activities.\n"
    "• Blue figures are links: column A opens the 'Build-Up' line; a date or count opens its source row.\n"
    "• Units: working days exclude Fridays; calendar days count every day; a month is 30.4 calendar days; a person-day is one person for one working day. Amounts exclude VAT."), height=84)

# Section 1 - status
banner(pg, 6, '1. STATUS OF THE PROGRAMME AND OF THE INSTRUCTION')
para(pg, 7, (f"Programme\n"
    f"• Baseline 'QC05958-BSL01THF-TSE-Final' ({XER}): data date 01-Jul-2026, 81 activities, cost-loaded to SAR 8,110,296.62 including 5 per cent Overhead and Profit.\n"
    "• Submitted to the Engineer 12-Sep-2026; issued to the Cost Consultant 29-Sep-2026 with S-curve, cash flow, manpower histogram and two-week look-ahead (data date 28-Sep-2026).\n"
    "• Status 30-Sep-2026: under the Engineer's approval - acceptable with minor comments, procurement schedule required; not approved. Used as evidence of SAMA's intended sequence and durations, not as an agreed basis."), height=70)
para(pg, 8, ("Instruction\n"
    "• SAMA has been instructed to proceed and is on site (existing tank dismantled; Tank 1 base panels being received). Reference and date not yet supplied; to be recorded on 'Project Info'.\n"
    "• The instruction affects entitlement and certification, not the assessed value."), height=44)
para(pg, 9, ("How the programme is used\n"
    "• Site period: as submitted; its end is fixed by the pump-room readiness date 03-Dec-2026 and the tie-in, testing and demobilisation after it.\n"
    "• Erection windows and durations: as submitted; labour and plant priced for those durations (Section 4). Testing: the Engineer's one-fill basis (email 30-Aug-2026) fitted to the submitted dates.\n"
    "• SAMA's slippage (Tank 1 base panels forecast 30-Sep against 12-Sep programmed) is its own delivery risk and adds nothing. Dates after 28-Sep-2026 are forecasts."), height=57)

# Section 2 - calendar
banner(pg, 11, '2. WORKING CALENDAR')
header(pg, 12, ['Ref', 'Calendar item', 'Value', '', '', '', '', '', 'Source'], merge_ij=True)
pg.merge_cells('C12:H12')
pg['A13'] = 'C1'; pg['B13'] = 'Working week for site activities'; pg['C13'] = '6 days, Friday non-working, 10 hours a day'
pg['I13'] = f"Calendar 'QIC 6 Days 10 Hrs' in {XER}, applied to every site activity; the approval activities use 'QIC 5 Days 10 Hrs' and the milestones 'QIC 7 Days 10 Hrs'"
pg['A14'] = 'C2'; pg['B14'] = 'Non-working day within the works period'; pg['C14'] = D(2026, 9, 23)
pg['I14'] = f"Calendar exception in {XER} (Saudi National Day). Every working-day formula on this tab excludes Fridays and this date; the calendar tables as imported are on the 'XER WBS' tab"
pg['A15'] = 'C3'; pg['B15'] = 'Calendar days per month used to convert periods to months'; pg['C15'] = 30.4
pg['I15'] = 'Average month (365 days / 12); months are rounded to one decimal place'
for r in (13, 14, 15):
    cp(S_REF, pg.cell(r, 1)); cp(S_DESC, pg.cell(r, 2)); cp(S_BASIS, pg.cell(r, 9)); cp(S_BASIS, pg.cell(r, 10))
    for c in range(3, 9): cp(S_UNIT, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    pg.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    pg.cell(r, 3).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    pg.row_dimensions[r].height = 30
pg['C14'].number_format = 'dd-mmm-yyyy'
pg['C15'].number_format = '0.0'
HOL = '$C$14'
MON = '$C$15'

# Section 5 activity table is written first so Section 3 can reference its cells. Decide its start row now.
SEC3_START = 17
# we will lay Section 3 in rows 17..~62, Section 4 after; Section 5 after that. Compute positions after building 3 and 4.

# --- activity data (from the .xer, sorted by start date as in the brief's Appendix A)
def parse_xer(path):
    tables = {}; cur = None
    for line in open(path, encoding='latin-1'):
        p = line.rstrip('\n').split('\t')
        if p[0] == '%T': cur = p[1]; tables[cur] = {'F': None, 'R': []}
        elif p[0] == '%F': tables[cur]['F'] = p[1:]
        elif p[0] == '%R': tables[cur]['R'].append(dict(zip(tables[cur]['F'], p[1:])))
    return tables
T = parse_xer('src/prog.xer')
cal = {r['clndr_id']: r['clndr_name'] for r in T['CALENDAR']['R']}
wbs = {r['wbs_id']: r for r in T['PROJWBS']['R']}
def wbs_path(wid):
    parts = []
    while wid in wbs and wbs[wid]['parent_wbs_id'] in wbs:
        parts.append(wbs[wid]['wbs_name']); wid = wbs[wid]['parent_wbs_id']
    return ' / '.join(reversed(parts))
cost = {}
for r in T['TASKRSRC']['R']:
    cost[r['task_id']] = cost.get(r['task_id'], 0) + float(r['target_cost'])
acts = []
for r in T['TASK']['R']:
    s = dt.datetime.strptime(r['target_start_date'], '%Y-%m-%d %H:%M').date()
    f = dt.datetime.strptime(r['target_end_date'], '%Y-%m-%d %H:%M').date()
    dayhrs = float({c['clndr_id']: c['day_hr_cnt'] for c in T['CALENDAR']['R']}[r['clndr_id']])
    acts.append(dict(id=r['task_code'], name=r['task_name'], wd=int(float(r['target_drtn_hr_cnt'])) // 10, hrs=float(r['target_drtn_hr_cnt']), dayhrs=dayhrs,
                     type={'TT_Task': 'Task', 'TT_Mile': 'Start milestone', 'TT_FinMile': 'Finish milestone'}.get(r['task_type'], r['task_type']),
                     status={'TK_NotStart': 'Not started', 'TK_Active': 'In progress', 'TK_Complete': 'Complete'}.get(r['status_code'], r['status_code']),
                     start=s, finish=f, cal=cal[r['clndr_id']], sec=wbs_path(r['wbs_id']), cost=round(cost.get(r['task_id'], 0), 2),
                     tf=float(r['total_float_hr_cnt'])))
acts.sort(key=lambda a: (a['start'], a['id']))
assert len(acts) == 81 and abs(sum(a['cost'] for a in acts) - 8110296.62) < 0.05

# --- Section 3 layout
banner(pg, 17, '3. PROGRAMME WINDOWS - THE SUBMITTED PROGRAMME, AND THE ASSESSMENT ALLOWANCE BUILT ON ITS ERECTION DATES WITH SEQUENTIAL TESTING')
para(pg, 18, ("Two sets of dates\n"
    "• Columns C to E: SAMA Submitted Programme, including its parallel testing of both tanks.\n"
    "• Columns F to H (yellow border): Assessed - the same erection dates, with the Engineer's one-fill sequential testing and the demobilisation fitted in. These are the dates used.\n"
    "• Every source date is looked up on 'XER WBS' by activity ID; click a blue date to open it. People and plant for these windows: Section 4."), height=58)
header(pg, 19, HDR3, merge_ij=True)
PG['used'] = sum((pg.row_dimensions[i].height or 14.4) for i in range(1, 20))

# activity lookup by id -> row in section 5 (filled later); we reference dates by cell so record ids
ACT_ROW = {}   # id -> row number in section 5

# Section 5 will start at row SEC5_START; compute after Section 3/4 sized. We'll pre-assign: Section 3 rows 20..47, Section 4 rows 50..~96, Section 5 from 100.
import datetime as _dt
# ================================================================ 'XER WBS' source tab - raw import of the programme file
wx = wb.create_sheet('XER WBS', index=wb.sheetnames.index('Programme') + 1)
for k, v in {'A': 22, 'B': 46, 'C': 44, 'D': 10, 'E': 11, 'F': 16, 'G': 11, 'H': 9, 'I': 10, 'J': 12, 'K': 12, 'L': 10, 'M': 16}.items():
    wx.column_dimensions[k].width = v
wx.sheet_view.showGridLines = False
wx.page_setup.orientation = 'landscape'; wx.page_setup.paperSize = 9; wx.page_setup.fitToWidth = 1; wx.page_setup.fitToHeight = 0
wx.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
wx.page_margins = copy.copy(wb['Assessment'].page_margins)
wx.sheet_properties.tabColor = '0045BF'
proj = T['PROJECT']['R'][0]
hdr_line = open('src/prog.xer', encoding='latin-1').readline().rstrip('\n').split('\t')
wx['A1'] = "PROGRAMME SOURCE DATA - IMPORTED FROM THE CONTRACTOR'S PRIMAVERA FILE"; cp(S_TITLE, wx['A1']); wx.row_dimensions[1].height = 25.5
wx['A2'] = 'TSE Irrigation Storage Tanks and Associated Pipeworks (RFP-027) - Contract QPMO-410-CT-05958'; cp(S_SUB, wx['A2']); wx.row_dimensions[2].height = 17.4
wx['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'; cp(S_SUB2, wx['A3'])
def xpara(r_, text, h):
    wx.cell(r_, 1, text)
    for c in range(1, 14): cp(S_NOTE, wx.cell(r_, c))
    wx.merge_cells(start_row=r_, start_column=1, end_row=r_, end_column=13); wx.row_dimensions[r_].height = h
xpara(4, (f"Source file {XER}, Primavera P6 export version {hdr_line[1]} dated {hdr_line[2]}, project '{proj['proj_short_name']}', data date "
          f"{dt.datetime.strptime(proj['last_recalc_date'], '%Y-%m-%d %H:%M').strftime('%d-%b-%Y')}, received from the Contractor on 29-Sep-2026 and submitted to the "
          "Engineer on 12-Sep-2026; under the Engineer's approval (acceptable with minor comments, procurement schedule required), not approved. "
          "Everything on this tab is copied from the file as it was received: identifiers, names and spellings are the Contractor's. Columns marked "
          "'source' hold the file's values; 'calculated' columns convert hours to working days on each activity's own calendar. The 'Programme' tab "
          "reads every date and duration it uses from this tab by activity ID; nothing on this tab is an assessment."), 57)
def xbanner(r_, text):
    wx.cell(r_, 1, text)
    for c in range(1, 14): cp(S_BANNER, wx.cell(r_, c))
    wx.merge_cells(start_row=r_, start_column=1, end_row=r_, end_column=13); wx.row_dimensions[r_].height = 21
def xheader(r_, labels):
    for i, t in enumerate(labels, 1):
        wx.cell(r_, i, t); cp(S_HDR, wx.cell(r_, i))
    wx.row_dimensions[r_].height = 32
def xcell(r_, c, v, style=None, fmt=None, wrap=False):
    cell = wx.cell(r_, c, v); cp(style or S_UNIT, cell)
    if fmt: cell.number_format = fmt
    if wrap: cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    return cell
# --- calendars
xr = 6
xbanner(xr, '1. CALENDARS (source table CALENDAR)'); xr += 1
xheader(xr, ['Calendar ID', 'Calendar name', 'Working days', 'Hours a day', 'Hours a week', 'Non-working exceptions in the works period (Jul to Dec 2026)']); wx.merge_cells(start_row=xr, start_column=6, end_row=xr, end_column=13); xr += 1
CAL_ROW = {}
for c in T['CALENDAR']['R']:
    exc = [dt.date(1899, 12, 30) + dt.timedelta(int(e)) for e in re.findall(r'd\|(\d+)', c['clndr_data'])]
    exc = sorted(d for d in exc if dt.date(2026, 7, 1) <= d <= dt.date(2026, 12, 31))
    days = 'Sat to Thu (Friday off)' if c['week_hr_cnt'] == '60' else ('Sat to Wed (Friday and Thursday off)' if c['week_hr_cnt'] == '50' else ('7 days' if c['week_hr_cnt'] == '70' else 'Mon to Fri'))
    xcell(xr, 1, c['clndr_id']); xcell(xr, 2, c['clndr_name'], S_DESC); xcell(xr, 3, days, wrap=True); xcell(xr, 4, float(c['day_hr_cnt']), fmt='0'); xcell(xr, 5, float(c['week_hr_cnt']), fmt='0')
    xcell(xr, 6, ', '.join(d.strftime('%d-%b-%Y') for d in exc) or 'none', S_BASIS); wx.merge_cells(start_row=xr, start_column=6, end_row=xr, end_column=13)
    wx.row_dimensions[xr].height = 18; CAL_ROW[c['clndr_id']] = xr; xr += 1
XCAL_HOL = f"'XER WBS'!$F${CAL_ROW['11622']}"   # informational; the Programme tab's calendar cell C14 holds the date used in formulas
xr += 1
# --- WBS hierarchy
xbanner(xr, '2. WBS HIERARCHY (source table PROJWBS)'); xr += 1
xheader(xr, ['WBS ID', 'WBS name (source)', 'Parent WBS ID', 'Level', 'Full path']); wx.merge_cells(start_row=xr, start_column=5, end_row=xr, end_column=13); xr += 1
def wbs_level(wid):
    n = 0
    while wid in wbs and wbs[wid]['parent_wbs_id'] in wbs: n += 1; wid = wbs[wid]['parent_wbs_id']
    return n
order = sorted(T['PROJWBS']['R'], key=lambda w: (wbs_path(w['wbs_id']) or ''))
for w in order:
    lvl = wbs_level(w['wbs_id'])
    xcell(xr, 1, w['wbs_id']); c = xcell(xr, 2, ('    ' * lvl) + w['wbs_name'], S_DESC); xcell(xr, 3, w['parent_wbs_id'] if w['parent_wbs_id'] in wbs else '(project)'); xcell(xr, 4, lvl, fmt='0')
    xcell(xr, 5, wbs_path(w['wbs_id']) or w['wbs_name'], S_BASIS); wx.merge_cells(start_row=xr, start_column=5, end_row=xr, end_column=13)
    wx.row_dimensions[xr].height = 16; xr += 1
xr += 1
# --- activities
xbanner(xr, '3. ACTIVITIES (source table TASK, cost from TASKRSRC) - sorted by source start date'); xr += 1
xheader(xr, ['Activity ID (source)', 'Activity name (source)', 'WBS path', 'Type', 'Status', 'Calendar', 'Duration, hours (source)', 'Hours a day', 'Working days (calculated)', 'Start (source)', 'Finish (source)', 'Total float, hours (source)', 'Cost loaded, SAR (source)'])
XA_HDR = xr; xr += 1
XA0 = xr
XROW = {}
for a in acts:
    XROW[a['id']] = xr
    xcell(xr, 1, a['id'], wrap=False).alignment = Alignment(horizontal='left', vertical='center')
    xcell(xr, 2, a['name'], S_DESC); xcell(xr, 3, a['sec'], S_BASIS)
    xcell(xr, 4, a['type']); xcell(xr, 5, a['status']); xcell(xr, 6, a['cal'], wrap=True)
    xcell(xr, 7, a['hrs'], fmt='#,##0'); xcell(xr, 8, a['dayhrs'], fmt='0')
    xcell(xr, 9, f"=IF(H{xr}=0,0,G{xr}/H{xr})", fmt='#,##0'); xcell(xr, 10, a['start'], fmt='dd-mmm-yyyy'); xcell(xr, 11, a['finish'], fmt='dd-mmm-yyyy')
    xcell(xr, 12, a['tf'], fmt='#,##0'); xcell(xr, 13, a['cost'], S_AMT)
    wx.row_dimensions[xr].height = 30 if len(a['sec']) > 60 or len(a['name']) > 55 else 18
    xr += 1
XA1 = xr - 1
wx.cell(xr, 2, 'Total cost loaded (source)'); cp(S_TOTLBL, wx.cell(xr, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12): cp(S_TOTLBL, wx.cell(xr, c))
wx.cell(xr, 13, f"=SUM(M{XA0}:M{XA1})"); cp(S_TOTAMT, wx.cell(xr, 13)); wx.row_dimensions[xr].height = 19.5
xr += 2
# --- relationships
xbanner(xr, '4. RELATIONSHIPS (source table TASKPRED)'); xr += 1
xheader(xr, ['Successor activity ID', 'Successor name', 'Predecessor activity ID', 'Type', 'Lag, hours (source)', 'Lag, working days (calculated)', 'Predecessor name']); wx.merge_cells(start_row=xr, start_column=7, end_row=xr, end_column=13); xr += 1
id_by_task = {r_['task_id']: r_['task_code'] for r_ in T['TASK']['R']}
name_by_code = {a['id']: a['name'] for a in acts}
XR0 = xr
for rel in sorted(T['TASKPRED']['R'], key=lambda x: (id_by_task[x['task_id']], id_by_task[x['pred_task_id']])):
    su, pr = id_by_task[rel['task_id']], id_by_task[rel['pred_task_id']]
    xcell(xr, 1, su).alignment = Alignment(horizontal='left', vertical='center'); xcell(xr, 2, name_by_code[su], S_DESC)
    xcell(xr, 3, pr).alignment = Alignment(horizontal='left', vertical='center'); xcell(xr, 4, rel['pred_type'].replace('PR_', ''))
    xcell(xr, 5, float(rel['lag_hr_cnt']), fmt='#,##0'); xcell(xr, 6, f"=E{xr}/10", fmt='#,##0')
    xcell(xr, 7, name_by_code[pr], S_BASIS); wx.merge_cells(start_row=xr, start_column=7, end_row=xr, end_column=13)
    wx.row_dimensions[xr].height = 16; xr += 1
wx.freeze_panes = 'A5'
wx.print_title_rows = '1:3'
XID = f"'XER WBS'!$A${XA0}:$A${XA1}"
def _x(col, idx): return f"INDEX('XER WBS'!${col}${XA0}:${col}${XA1},MATCH(\"{idx}\",{XID},0))"
def AS(idx): return _x('J', idx)     # source start date, looked up by activity ID
def AF(idx): return _x('K', idx)     # source finish date
def AWD(idx): return _x('I', idx)    # working days calculated from source hours

def W(start, n):  # finish after n working days inclusive of start
    return f"WORKDAY.INTL({start},{n}-1,{WK},{HOL})"
def NEXT(cell): return f"WORKDAY.INTL({cell},1,{WK},{HOL})"
def WD(a, b): return f"NETWORKDAYS.INTL({a},{b},{WK},{HOL})"

rows3 = []  # (ref, window, A_start, A_finish, A_days, B_start, B_finish, B_days, derivation, kind)  kind: 'wd' or 'cd' or 'ms'
r = 20
def add3(ref, window, As, Af, Ad, Bs, Bf, Bd, deriv, kind='wd', height=None):
    global r
    height = max(height or 30, est(window, 46), est(deriv, 66))
    if PG['used'] + height > PAGE:
        old = r
        pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(PG['hdr'])
        # the caller built its same-row references against the row before the repeated header was inserted
        fix = lambda v: re.sub(r'(?<![A-Z$])([C-H])' + str(old) + r'(?!\d)', lambda m: m.group(1) + str(r), v) if isinstance(v, str) else v
        As, Af, Ad, Bs, Bf, Bd = map(fix, (As, Af, Ad, Bs, Bf, Bd))
    pg_add(height)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, window); cp(S_DESC, pg.cell(r, 2))
    for col, v in ((3, As), (4, Af), (6, Bs), (7, Bf)):
        c = pg.cell(r, col, v); datecell(c)
        if v in (None, ''): c.value = '-'
    for col, v in ((5, Ad), (8, Bd)):
        c = pg.cell(r, col, v); numcell(c, '#,##0' if kind != 'mo' else '0.0'); c.alignment = Alignment(horizontal='right', vertical='center')
        if v in (None, ''): c.value = '-'
    pg.cell(r, 9, deriv); cp(S_BASIS, pg.cell(r, 9)); cp(S_BASIS, pg.cell(r, 10))
    pg.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    pg.row_dimensions[r].height = height
    rows3.append(r); r += 1
    return r - 1

# P1 mobilisation
def same(r_): return (f"=C{r_}", f"=D{r_}", f"=E{r_}")
p1 = add3('P1', 'Mobilisation and site establishment (calendar days)', f"={AS('QCD18TSEMOB1240')}", f"={AF('QCD18TSEMOB1250')}", f"=D{r}-C{r}+1",
          *same(r), f"Activities QCD18TSEMOB1240 and QCD18TSEMOB1250, 22 to 26-Aug-2026, {XER}. Mobilisation has taken place (the existing tank is dismantled)", 'cd')
p2 = add3('P2', 'Dismantling of the existing damaged Tank-1 (calendar days)', f"={AS('QCD18TSECONDSM1020')}", f"={AF('QCD18TSECONDSM1030')}", f"=D{r}-C{r}+1",
          *same(r), f"Activities QCD18TSECONDSM1020 to 1030, 27-Aug to 12-Sep-2026, {XER}. Complete per the look-ahead of 29-Sep-2026 (data date 28-Sep-2026); the completion date is not recorded, so the programmed dates are kept. Dismantling, segregation and loading are within the Item 3 quotation", 'cd')
p3 = add3('P3', 'Tank 1 - base panels (working days)', f"={AS('QCD18TSECONT1INS1020')}", f"={AF('QCD18TSECONT1INS1020')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1020, 12-Sep to 05-Oct-2026, 20 working days, linked to the partial delivery QCD18TSEPRC1120 (08-Sep to 03-Oct-2026, start 3 days after, finish 2 days after). Taken as submitted. The identical work on Tank 2 (QCD18TSECONT2INS1020) is programmed at 6 working days - a productivity cross-check only (Section 4, S5). The look-ahead forecasts the start as 30-Sep-2026, delivery-driven, the Contractor's risk")
p4 = add3('P4', 'Tank 1 - wall panels (working days)', f"={AS('QCD18TSECONT1INS1030')}", f"={AF('QCD18TSECONT1INS1030')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1030, 06 to 24-Oct-2026, 16 working days, linked to the wall-panel delivery QCD18TSEPRC1170 (start 2 days after, finish 4 days after). Taken as submitted; Tank 2 walls (QCD18TSECONT2INS1030) 12 working days, cross-check only")
p5 = add3('P5', 'Tank 1 - bracing (working days)', f"={AS('QCD18TSECONT1INS1060')}", f"={AF('QCD18TSECONT1INS1060')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1060, 08 to 29-Oct-2026, 19 working days, starting 2 days after the walls and finishing 5 days after them. Taken as submitted; Tank 2 bracing 15 working days, cross-check only")
p6 = add3('P6', 'Tank 1 - roof supports and ladder (working days)', f"={AS('QCD18TSECONT1INS1040')}", f"={AF('QCD18TSECONT1INS1040')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1040, 31-Oct to 04-Nov-2026, 5 working days, the same as Tank 2")
p7 = add3('P7', 'Tank 1 - roof panels (working days)', f"={AS('QCD18TSECONT1INS1050')}", f"={AF('QCD18TSECONT1INS1050')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1050, 05 to 11-Nov-2026, 6 working days, the same as Tank 2")
p8 = add3('P8', 'Tank 1 - mechanical works: nozzles, level transmitter, pipes and fittings (working days)', f"={AS('QCD18TSECONT1MW2050')}", f"={AF('QCD18TSECONT1MW2055')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activities QCD18TSECONT1MW2050, 2060 and 2055, 09 to 16-Nov-2026. Nozzles and internals are within the supplier's scope (Item 6); the level transmitter and the pipes and fittings are priced supplied and installed under Item 8")
p9 = add3('P9', 'Tank 2 - base panels (working days)', f"={AS('QCD18TSECONT2INS1020')}", f"={AF('QCD18TSECONT2INS1020')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT2INS1020, 22 to 28-Oct-2026, 6 working days, starting 2 working days after the Tank 2 delivery begins (QCD18TSEPRC1210, 20-Oct-2026); from here the two tanks are erected on two fronts at once until 17-Nov-2026")
p10 = add3('P10', 'Tank 2 - wall panels (working days)', f"={AS('QCD18TSECONT2INS1030')}", f"={AF('QCD18TSECONT2INS1030')}", f"={WD(f'C{r}', f'D{r}')}", *same(r), "Activity QCD18TSECONT2INS1030, 29-Oct to 11-Nov-2026, 12 working days")
p11 = add3('P11', 'Tank 2 - bracing (working days)', f"={AS('QCD18TSECONT2INS1060')}", f"={AF('QCD18TSECONT2INS1060')}", f"={WD(f'C{r}', f'D{r}')}", *same(r), "Activity QCD18TSECONT2INS1060, 01 to 17-Nov-2026, 15 working days")
p12 = add3('P12', 'Tank 2 - roof supports and ladder (working days)', f"={AS('QCD18TSECONT2INS1040')}", f"={AF('QCD18TSECONT2INS1040')}", f"={WD(f'C{r}', f'D{r}')}", *same(r), "Activity QCD18TSECONT2INS1040, 18 to 23-Nov-2026, 5 working days")
p13 = add3('P13', 'Tank 2 - roof panels (working days)', f"={AS('QCD18TSECONT2INS1050')}", f"={AF('QCD18TSECONT2INS1050')}", f"={WD(f'C{r}', f'D{r}')}", *same(r), "Activity QCD18TSECONT2INS1050, 24 to 30-Nov-2026, 6 working days")
p14 = add3('P14', 'Tank 2 - mechanical works: nozzles, level transmitter, pipes and fittings (working days)', f"={AS('QCD18TSECONT2MW2030')}", f"={AF('QCD18TSECONT2MW2020')}", f"={WD(f'C{r}', f'D{r}')}", *same(r), "Activities QCD18TSECONT2MW2030, 2040 and 2020, 28-Nov to 03-Dec-2026")
p15 = add3('P15', "Milestone - SAJCO readiness for the tie-in connections (external)", f"={AS('QCD18TSECONIF2050')}", f"={AF('QCD18TSECONIF2050')}", '-',
           f"=C{r}", f"=D{r}", '-', "Activity QCD18TSECONIF2050, 03-Dec-2026, zero duration. An interface milestone outside the Contractor's control, taken as submitted as submitted and carried; it, not the tank erection, fixes the earliest tie-in date", 'ms')
p16 = add3('P16', 'Tie-in connections with the existing pump room (working days)', f"={AS('QCD18TSECONTC2040')}", f"={AF('QCD18TSECONTC2040')}", f"={WD(f'C{r}', f'D{r}')}",
           *same(r), "Activity QCD18TSECONTC2040, 05 to 08-Dec-2026, 4 working days, following the readiness milestone")
p17 = add3('P17', 'Testing and commissioning as programmed - both tanks in parallel (working days)', f"={AS('QCD18TSECONTCT12050')}", f"={AF('QCD18TSECONTCT22030')}", f"={WD(f'C{r}', f'D{r}')}",
           '-', '-', '-', "As submitted only: activities QCD18TSECONTCT12050 and QCD18TSECONTCT22030, 09 to 16-Dec-2026, 7 working days each, in parallel, with two simultaneous fills. Not carried: the Engineer's (KEO) email of 30-Aug-2026 states that installation and testing will not be in parallel and that one tank is filled and the water re-used for the second. The sequential, task-based sequence the assessment uses is at P18 to P22")
p18 = add3('P18', 'Hydrostatic test - Tank 1, before the tie-in (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'D{p8}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONTCT12050')}",
           "Carried: starts the working day after the Tank 1 mechanical works (submitted date) and takes the 7 working days the Contractor programmed per tank. Elapsed sequence assumed within those 7 days: tankered filling about 5 working days, the 24-hour hold (unattended apart from level readings), inspection of joints and nozzles and records 1 day. The fill rate is an assumption, not a measured throughput: 3,774 m3 in 5 days needs about 750 m3 a day, for example two 30 m3 tankers on about 12 round trips each, which depends on the water source the Employer has yet to confirm; if the source is further away the elapsed time lengthens but the attendance does not. Attendance: the QA/QC inspector (1.3, monthly) and the commissioning engineer for one day at the end of the hold (7.13); the supplier's leak-test supervision is within Item 6. It sits inside the 15 working days of float the programme gives Tank 1 (17 to 24-Nov-2026). Conditional on an Engineer-approved method: internals flushed and nozzles blind-flanged (RFP Scope of Works work packages 4 and 5), the water retained in Tank 1 until Tank 2 is ready, losses at 'Build-Up' line 7.5", 96)
p19 = add3('P19', 'Transfer of the test water from Tank 1 to Tank 2 (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'MAX(G{p18},G{p14})')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "Starts the working day after both the Tank 1 test has passed (P18) and Tank 2 is ready to receive water (mechanical works complete, P14, interior flushed). Pumped tank to tank through temporary hoses with the outlet valves isolated; the tie-in is not needed for the transfer. 3 working days is an assessed assumption pending the Contractor's method statement: 3,774 m3 at about 130 m3 an hour over 10-hour shifts, a 150 mm self-priming diesel pump against a low head (adjacent tanks at one level, about 4 m static plus hose friction) - 'Build-Up' lines 7.2 to 7.4")
p20 = add3('P20', 'Hydrostatic test - Tank 2, after the transfer (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'G{p19}')}", f"={W(f'F{r}', f'H{r}')}", 4,
           "Elapsed 4 working days (08, 09, 10 and 12-Dec; 11-Dec is a Friday): top-up to test level on 08-Dec, the 24-hour hold on 09-Dec (unattended apart from level readings), inspection on 10-Dec, sampling and inspection on 12-Dec - 3 attended days out of 4 (Section 4, dated check H8). With the 3 transfer days the Tank 2 sequence is 7 working days from the start of the transfer, the programme's per-tank figure. Disinfection relies on the dosed water brought over (see the water sequence below the derived periods) - a provisional method; acceptance rests on the 12-Dec samples")
p20b = add3('P20b', 'Tank 1 operating fill for the commissioning demonstration - tankered, about 1 m depth (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'MAX(G{p19},G{p16})')}", f"={W(f'F{r}', f'H{r}')}", 2,
           "Assessment allowance: Tank 1 is empty after the transfer (P19), and the integrated commissioning (P21b) needs both tanks connected and holding water. About 1,020 m3 "
           "(1 m depth) is tankered in on the two working days after the tie-in and the transfer, while Tank 2 is on its hold - off the critical path, attended by the approved "
           "helpers (Section 4, duty H11). The depth the Engineer requires is not established; a full fill is sensitivity S8. Water from the network after the tie-in is not "
           "assumed: no source has been confirmed. Not in the submitted programme, which fills both tanks in full for parallel testing", 'wd', 80)
p21 = add3('P21', 'Component and subsystem checks after the tie-in - instruments, nozzles, valves, ladders; tank-pump-network interfaces (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'G{p16}')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "RFP Scope of Works 5.2, component testing and subsystem validation. Follows the tie-in and may overlap the Tank 2 hydrostatic test, because these checks do not need both tanks in service. 3 working days is an assessed assumption, not a programme figure: the Contractor's programme has no separate activity for this stage (its two 7-working-day 'Testing & Commissioning' activities cover the hydrostatic test and the commissioning of each tank together). Staffing at 'Build-Up' lines 7.13 to 7.18")
p21b = add3('P21b', 'Integrated system commissioning - full operational demonstration and witness testing (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'MAX(G{p16},G{p20},G{p21},G{p20b})')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "RFP Scope of Works 5.2, integrated system commissioning. Starts the working day after the last of the tie-in (P16), the Tank 2 hydrostatic test (P20) and the component checks (P21): both tanks must have passed and be connected before full operation is demonstrated. Pumping is demonstrated from Tank 2, which holds the dechlorinated test water; Tank 1 is empty after the transfer and is demonstrated on valves and instruments only (water sequence below; S8 if the Engineer requires more). Release also needs the Engineer's acceptance of the Tank 2 samples. 3 working days is an assessed assumption on the same footing as P21")
p22 = add3('P22', 'Completion of testing and commissioning - both tanks', f"={AF('QCD18TSEOMS1040')}", f"=C{r}", '-', f"=G{p21b}", f"=F{r}", '-',
           "As submitted: completion milestones QCD18TSEOMS1040 and 1050, 16-Dec-2026, with parallel testing. Carried: end of the integrated system commissioning, conditional on the Tank 1 test preceding the tie-in (P18); if it cannot, see sensitivity S2", 'ms')
p23 = add3('P23', 'Demobilisation, as-built drawings and close-out documents (working days)', f"={AS('QCD18TSEDMOB1020')}", f"={AF('QCD18TSEDMOB1020')}", f"={WD(f'C{r}', f'D{r}')}",
           f"={NEXT(f'G{p22}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}", "Activities QCD18TSEDMOB1020 and QCD18TSEDMOB3020, 17 to 24-Dec-2026, 7 working days in parallel. B follows the completion of testing (P22)")
# derived periods
if PG['used'] + 19.5 + 120 > PAGE:
    pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(PG['hdr'])
pg_add(19.5)
pg.cell(r, 2, 'Derived periods used in Sections 4 and 5'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5
r += 1
d1 = add3('D1', 'Site period - mobilisation start to demobilisation finish (calendar days)', f"=C{p1}", f"=D{p23}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p23}", f"=G{r}-F{r}+1",
          "The continuous site establishment runs for this period: welfare cabins, WC, water tank and deliveries, the 30 kVA welfare generator, the site pick-up, workforce transport and the watchman, and the site staff whose duties cover the demobilisation week (Section 4, roles). As submitted: the programme as submitted, 22-Aug to 24-Dec-2026. Carried: the same erection dates with the sequential testing fitted in, conditional on the early Tank 1 test (P18). The site engineer's 3 close-out days at 'Build-Up' line 1.22 fall after this period and are intentionally off-site visits, not a second allowance", 'cd')
d2 = add3('D2', 'Works period - mobilisation start to completion of testing and commissioning (calendar days)', f"=C{p1}", f"=D{p22}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p22}", f"=G{r}-F{r}+1",
          "The daytime 100 kVA works generator runs for this period; no works power is needed during demobilisation, when the welfare generator alone continues", 'cd')
d3 = add3('D3', 'Erection window - Tank 1 base panels start to Tank 2 mechanical finish (calendar days)', f"=C{p3}", f"=D{p14}", f"=D{r}-C{r}+1", *same(r),
          "Power tools and lighting towers are hired by the month while erection is in progress, 12-Sep to 03-Dec-2026 as submitted; the same as submitted and carried", 'cd')
d6 = add3('D6', 'Access equipment window - Tank 1 walls start to Tank 2 roof finish (calendar days)', f"=C{p4}", f"=D{p13}", f"=D{r}-C{r}+1", *same(r),
          "Mobile access towers and podium steps are hired by the month for the period the walls, bracing and roofs are worked on; rounded up to whole months at 'Build-Up' lines 5.4 and 5.6", 'cd')
d7 = add3('D7', 'Storage containers window - first panels on site to last roof panel installed (calendar days)', f"={AS('QCD18TSEPRC1120')}", f"=D{p13}", f"=D{r}-C{r}+1", *same(r),
          "First delivery 08-Sep-2026 to the Tank 2 roof finish 30-Nov-2026. Panels have to be stored from the first delivery whatever the erection pace; storage is priced by the container-month, so the delivery staging adds no separate cost here", 'cd')
# --- water sequence (provisional method)
if PG['used'] + 120 > PAGE:
    pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(PG['hdr'])
pg_add(120)
para(pg, r, ("Provisional commissioning method, both tanks - for SAMA's method statement and the Engineer's acceptance; not verified\n"
    "• Hydrostatic test water: Tank 1 filled by tanker to 3.7 m, 3,774 m3 (7.1), 24-hour hold, inspected (P18). Pumped to Tank 2 once Tank 2 is complete (P19, 3 days); Tank 1 is then empty, so the tanks are never full together. Tank 2 topped up (7.5), held, inspected, sampled (P20).\n"
    "• Disinfection: the test water is dosed with hypochlorite in Tank 1 (7.6, 7.7), circulated, sampled (7.10, 7.11) and boosted in Tank 2. That covers wetted surfaces only; surfaces above the test level are sprayed separately (7.25) - Tank 2 before it receives water, Tank 1 after it is emptied.\n"
    "• Functional commissioning: component checks after the tie-in (P21). The integrated demonstration (P21b) needs both tanks connected and holding water: Tank 2 keeps the test water (dechlorinated, 7.9); Tank 1 receives a 1,020 m3 tanker fill on 09 to 10-Dec (P20b, 7.24) within the approved helpers.\n"
    "• Not assumed: a network refill after the tie-in (no source or cost confirmed); a Tank 1 demonstration on valves and instruments alone (not shown to meet RFP Scope of Works 5.2).\n"
    "• Release: integrated commissioning starts after the Tank 2 hydrostatic pass, sample acceptance and the tie-in; a laboratory turnaround over one working day moves P21b day for day. Water kept as first stock if the Employer instructs, otherwise dechlorinated and discharged (7.9).\n"
    "• Not established: water source and price, disinfection method, depth required in Tank 1 (full fill = S8), disposal route. Item 7 is an Assessed allowance until the method statement is accepted."), height=170); r += 1
# --- sensitivity block
if PG['used'] + 19.5 + 70 > PAGE:
    pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(PG['hdr'])
pg_add(19.5)
pg.cell(r, 2, 'Sensitivity - not carried; same durations and dependencies as P18 to P23'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5
r += 1
def sens(ref, name, cs, cf, cd, deriv, height=None):
    return add3(ref, name, cs, cf, cd, '-', '-', '-', deriv, 'cd', height)
b_t1t_f = W(NEXT(f'G{p16}'), 7); b_xf_f = W(NEXT(b_t1t_f), 3); b_t2t_f = W(NEXT(b_xf_f), 4); b_cc_f = W(NEXT(f'G{p16}'), 3); b_of_f = W(NEXT(b_xf_f), 2)
b_ic_f = W(NEXT(f'MAX(G{p16},{b_t2t_f},{b_cc_f},{b_of_f})'), 3); b_dm_f = W(NEXT(b_ic_f), 7)
s2 = sens('S2', 'Carried basis if the Tank 1 test cannot precede the tie-in - Tank 1 test after the tie-in, transfer, Tank 2 test, integrated commissioning, demobilisation (calendar days from mobilisation)', f"=F{p1}", f"={b_dm_f}", f"=D{r}-C{r}+1",
          "Tank 1 test 09 to 16-Dec-2026; transfer 17 to 20-Dec; Tank 2 test 21 to 24-Dec; component checks 09 to 12-Dec and the Tank 1 operating fill 21 to 22-Dec in parallel; integrated commissioning 26 to 28-Dec-2026; demobilisation 29-Dec-2026 to 05-Jan-2027 (Fridays excluded). The additional site days against D1 are shown at S3", 60)
s3 = add3('S3', 'S2 - complete incremental amount, excluding Overhead and Profit (first column: additional site days; third column: SAR)', f"=E{s2}-H{d1}", '-', 'SENS_AMT', '-', '-', '-',
          "Every Section 5 line re-evaluated with the S2 dates by its own formula and rounding: the site-period lines and site staff ('Build-Up' 1.1 to 1.8, 1.10, 1.12, 1.13, 5.8 and 5.9) at the S2 days in months; welfare-water deliveries (1.9) at two per rounded-up week; the works generator (5.7) to the S2 end of commissioning. The erection-window, plant, helper, access, container and Item 7 lines do not move because only the testing tail moves. Not carried: it applies only if the Engineer does not accept the early Tank 1 test", 60)
pg.cell(s3, 3).number_format = '#,##0'; pg.cell(s3, 5).number_format = '#,##0.00'
s4 = add3('S4', 'S2 - Overhead and Profit at 5 per cent on S3 (first column) and incremental amount including it (third column), SAR', f"=ROUND(E{s3}*0.05,2)", '-', '-', '-', '-', '-',
          "Applied once at the rate on the 'Assessment' tab. Nothing in S2 to S4 is carried into the assessment", 30)
pg.cell(s4, 3).number_format = '#,##0.00'; pg.cell(s4, 5).number_format = '#,##0.00'; pg.cell(s4, 5).value = f"=E{s3}+C{s4}"
SEC3_END = r - 1

# ================================================================ Section 4: resource bridge on the approved histogram
r += 1
pg_break_if(pg, r, 21 + 3 * 60 + 83 + 32 + 60)
pg_add(21 + 3 * 60)
banner(pg, r, "4. RESOURCE BRIDGE - THE APPROVED MANPOWER HISTOGRAM RECONCILED TO THE PRICED LINES"); r += 1
para(pg, r, ("The approved histogram - what the figures are\n"
    "• Source: SAMA's weekly direct-manpower histogram (workbook 'SAMA-D18 -THS-TSE - Cost S-Curves Cash Flow Manpower Histogram.xlsx', sheet 'MP-HST-WK'), approved: 221 person-weeks - 136 skilled, 85 helpers; peak 35 in the week ending 30-Oct-2026. Reproduced below unchanged.\n"
    "• Columns E to G are people deployed in that week (a weekly headcount): 8 in W1 means 8 people through the week. Person-days (column H) = people x the actual working days of that week (6, or 5 in the week containing 23-Sep-2026). The totals row gives person-weeks.\n"
    "• Assumption: if the histogram gives a peak day rather than the people through the week, the person-days are overstated. Approval of the histogram does not say whose people they are, approve any rate or make the labour payable by itself."), height=84); r += 1
para(pg, r, ("Whose people, and where they are paid\n"
    "• Skilled: the tank supplier installs the tanks (Assumption 9 on the 'Build-Up' tab), so the erection crews are the supplier's, inside the Item 6 price. Dismantling crew: inside the Item 3 quotation. Pipework and tie-in fitters: inside the Item 8 installed rates. None is priced again; the exceptions are marked in the skilled table (K1, K5).\n"
    "• Helpers: the supplier excludes helpers, so all 85 approved helper person-weeks are SAMA's and are priced once at 'Build-Up' 5.10 (the dismantling weeks provisionally, S7). Duties after 11-Dec-2026 are priced at 7.8; days needing more helpers than approved are priced as 'Additional helpers assessed' below.\n"
    "• Specialists (commissioning engineer, technicians, electrician, calibration, tie-in fitters, pipework testing crew) are not helpers; priced in Items 7 and 8. No productivity adjustment is made to the approved figures."), height=84); r += 1
HDRB = ['Ref', 'Week ending (histogram week) - phase', 'Erection activities in progress (count; click for the first, IDs in the note)', 'Working days in week', 'Approved total (people in the week)', 'Approved skilled (people in the week)', 'Approved helpers (people in the week)', 'Helper person-days (people x working days)', 'Helper person-days priced at 5.10', 'Whose people and where paid; activities counted']
def hdrb(ws, r_):
    header(ws, r_, HDRB[:9]); ws.cell(r_, 10, HDRB[9]); cp(S_HDR, ws.cell(r_, 10))
pg_break_if(pg, r, 32 + 18 * 38 + 2 * 42)   # the weekly histogram table stays on one page
hdrb(pg, r); pg_add(PG['hdr']); r += 1
XJ = f"'XER WBS'!$J${XA0}:$J${XA1}"; XK = f"'XER WBS'!$K${XA0}:$K${XA1}"; XC = f"'XER WBS'!$C${XA0}:$C${XA1}"
HISTW = [
 ('28-Aug', 8, 5, 3, 'Mobilisation; dismantling starts', 'Skilled: mobilisation 22 to 26-Aug then dismantling from 27-Aug - allocation between Contractor mobilisation and the Item 3 crew unresolved. Helpers: site set-up; provisionally priced (S7)'),
 ('04-Sep', 14, 9, 5, 'Dismantling', 'Skilled: dismantling crew, within Item 3 (recorded). Helpers: attendance - provisionally priced, may be within the Item 3 crew (S7)'),
 ('11-Sep', 13, 8, 5, 'Dismantling ends; survey; first base-panel batch', 'Skilled: dismantling crew, within Item 3 (recorded). Helpers: attendance and first-batch offloading - provisionally priced (S7)'),
 ('18-Sep', 12, 8, 4, 'Single front - Tank 1 base panels (programmed)', 'Skilled: tank erection crew, within Item 6 (recorded scope). Helpers: panel handling'),
 ('25-Sep', 11, 7, 4, 'Single front - Tank 1 base (5-day week, 23-Sep off)', 'As above'),
 ('02-Oct', 10, 6, 4, 'Single front - Tank 1 base', 'As above'),
 ('09-Oct', 10, 6, 4, 'Single front - Tank 1 walls and bracing; wall-panel batches', 'As above; helpers offload the wall-panel batches within the same days'),
 ('16-Oct', 8, 4, 4, 'Single front - Tank 1 walls and bracing', 'As above'),
 ('23-Oct', 14, 8, 6, 'Tank 2 kit expected 20-Oct; second front opens 22-Oct', 'Skilled: erection crews, Item 6. Helpers: Tank 1 front plus offloading and the Tank 2 base'),
 ('30-Oct', 35, 23, 12, 'Two fronts - Tank 1 bracing and roof supports; Tank 2 base and walls (peak)', 'Skilled: two erection crews, Item 6. Helpers: two gangs'),
 ('06-Nov', 26, 16, 10, 'Two fronts - Tank 1 roof; Tank 2 walls and bracing', 'As above'),
 ('13-Nov', 19, 11, 8, 'Two fronts - Tank 1 roof and nozzles; Tank 2 bracing; pipe delivery', 'Skilled: erection crews (Item 6) and, from 12-Nov, pipework fitters (Item 8 rates) - split unresolved, both priced elsewhere. Helpers: two gangs, pipe offloading'),
 ('20-Nov', 16, 10, 6, 'Single front - Tank 2 roof supports; Tank 1 hydrostatic test', 'Skilled: erection crew (Item 6) or fitters (Item 8) - unresolved, both priced elsewhere. Helpers: Tank 2 gang; Tank 1 fill attendance'),
 ('27-Nov', 13, 8, 5, 'Single front - Tank 2 roof panels', 'As above'),
 ('04-Dec', 7, 4, 3, 'Tank 2 nozzles; readiness milestone 03-Dec', 'Skilled: nozzle fitters (Item 6) and pipework (Item 8) - unresolved, both priced elsewhere. Helpers: nozzle attendance; Tank 1 disinfection'),
 ('11-Dec', 5, 3, 2, 'Tie-in; transfer; Tank 2 test; disinfection', 'Skilled: tie-in fitters within 8.14 and 8.15. Helpers: transfer 05 to 07-Dec and Tank 2 top-up and hold attendance - priced here, so line 7.4 is nil'),
 ('18-Dec', 0, 0, 0, 'Demobilisation (carried basis from 16-Dec)', 'No approved labour: Tank 2 sampling 12-Dec and discharge 16 to 17-Dec (7.8) and commissioning specialists (7.13 to 7.17) priced on their own lines'),
 ('25-Dec', 0, 0, 0, 'Demobilisation ends 23-Dec (Assessed basis)', 'No approved labour: demobilisation clean-up is line 1.16'),
]
b_first = r
wk_end = _dt.date(2026, 8, 28)
for i, (lab, tot_, sk, hp, phase, note) in enumerate(HISTW):
    h = max(30, est(note, 52), est(phase, 40))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); hdrb(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(h)
    pg.cell(r, 1, f'W{i + 1}'); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, f'{lab}-2026 - {phase}'); cp(S_DESC, pg.cell(r, 2))
    we = wk_end; ws_ = wk_end - _dt.timedelta(6)
    pg.cell(r, 3, f'=SUMPRODUCT(({XJ}<=DATE({we.year},{we.month},{we.day}))*({XK}>=DATE({ws_.year},{ws_.month},{ws_.day}))*ISNUMBER(SEARCH("Tank Installation",{XC})))'); numcell(pg.cell(r, 3), '#,##0')
    pg.cell(r, 4, f'=NETWORKDAYS.INTL(DATE({ws_.year},{ws_.month},{ws_.day}),DATE({we.year},{we.month},{we.day}),{WK},{HOL})'); numcell(pg.cell(r, 4), '#,##0')
    pg.cell(r, 5, tot_); numcell(pg.cell(r, 5), '#,##0'); pg.cell(r, 6, sk); numcell(pg.cell(r, 6), '#,##0'); pg.cell(r, 7, hp); numcell(pg.cell(r, 7), '#,##0')
    pg.cell(r, 8, f"=G{r}*D{r}"); numcell(pg.cell(r, 8), '#,##0')
    pg.cell(r, 9, f"=H{r}"); numcell(pg.cell(r, 9), '#,##0')
    pg.cell(r, 9).border = Border(left=Side(style='medium', color='FFC425'), right=Side(style='medium', color='FFC425'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    live_ = [a['id'] for a in acts if 'Tank Installation' in a['sec'] and a['start'] <= we and a['finish'] >= ws_]
    note = note + (' Counted: ' + ', '.join(live_) if live_ else '')
    h = max(h, est(note, 52))
    pg.cell(r, 10, note); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = h
    r += 1; wk_end += _dt.timedelta(7)
b_last = r - 1
def btot(label, cells, note, h=None):
    global r
    if PG['used'] + (h or 42) > PAGE:
        pg_break(pg, r); hdrb(pg, r); r += 1; pg_add(PG['hdr'])
    pg.cell(r, 2, label); cp(S_TOTLBL, pg.cell(r, 2))
    for c in range(1, 11):
        if c != 2: cp(S_TOTLBL, pg.cell(r, c))
    for col, f in cells.items():
        pg.cell(r, col, f); cp(S_TOTAMT, pg.cell(r, col)); pg.cell(r, col).number_format = '#,##0.0' if col == 9 and 'month' in label else '#,##0'
    pg.cell(r, 10, note); cp(S_BASIS, pg.cell(r, 10))
    hh = max(19.5, est(note, 52), h or 0); pg.row_dimensions[r].height = hh; pg_add(hh)
    r += 1; return r - 1
bt1 = btot('Approved histogram totals (man-weeks) and helper man-days', {5: f"=SUM(E{b_first}:E{b_last})", 6: f"=SUM(F{b_first}:F{b_last})", 7: f"=SUM(G{b_first}:G{b_last})", 8: f"=SUM(H{b_first}:H{b_last})", 9: f"=SUM(I{b_first}:I{b_last})"},
           "221 man-weeks: 136 skilled (within Items 3, 6 and 8 on their recorded scopes, not priced again) and 85 helpers (priced). Man-days use the working days of each week")
bt2 = btot("Helper man-months at 26 working days - 'Build-Up' line 5.10 quantity (approved man-days plus the departure in the dated check)", {9: f"=ROUND((I{bt1}+DEPARTURE)/26,1)"}, "Conversion for the man-month rate on the 'Build-Up' tab (6-day week); the departure is the shortfall against the approved deployment shown in the dated check")
# --- helper duties by week against the approved capacity, W13 onwards (man-days per week, overlapping days shown per duty)
wk = {int(str(pg.cell(rr, 1).value)[1:]): rr for rr in range(b_first, b_last + 1) if str(pg.cell(rr, 1).value)[:1] == 'W' and str(pg.cell(rr, 1).value)[1:].isdigit()}
assert len(wk) == len(HISTW), wk
r += 1
pg_break_if(pg, r, 83 + 32 + 3 * 30)
para(pg, r, ("Dated check - do the helper duties fit inside the approved helpers?\n"
    "• Each helper duty from 14-Nov-2026 to demobilisation is placed on its working days (XER dates and the fitted tests, Section 3); person-days = people x days, split by histogram week.\n"
    "• Demand is compared with the approved helper person-days of the same week. Days after 11-Dec-2026 have no approved people and are priced at 7.8. Weeks W1 to W12 cannot be checked duty by duty from the documents received."), height=57); r += 1
HDRD = ['Ref', 'Helper duty and its working days', 'W13 (14 to 20-Nov) person-days', 'W14 (21 to 27-Nov) person-days', 'W15 (28-Nov to 04-Dec) person-days', 'W16 (05 to 10-Dec) person-days', 'After 11-Dec person-days', 'Total person-days', 'Priced at', 'People x days, and basis']
def hdrd(ws, r_):
    header(ws, r_, HDRD[:9]); ws.cell(r_, 10, HDRD[9]); cp(S_HDR, ws.cell(r_, 10))
hdrd(pg, r); pg_add(PG['hdr']); r += 1
DUTY = {}
def dadd(ref, duty, w13, w14, w15, w16, after, priced, basis):
    global r
    hh = max(30, est(duty, 46), est(basis, 52))
    if PG['used'] + hh > PAGE:
        pg_break(pg, r); hdrd(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(hh)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, duty); cp(S_DESC, pg.cell(r, 2))
    for col, v in ((3, w13), (4, w14), (5, w15), (6, w16), (7, after)):
        c = pg.cell(r, col, v); numcell(c, '#,##0')
    pg.cell(r, 8, f"=SUM(C{r}:G{r})"); numcell(pg.cell(r, 8), '#,##0')
    pg.cell(r, 9, priced); cp(S_UNIT, pg.cell(r, 9)); pg.cell(r, 9).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    pg.cell(r, 10, basis); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = hh; DUTY[ref] = r; r += 1
    return r - 1
dadd('H1', 'Tank 2 bracing - gang attendance, 14 to 17-Nov (QCD18TSECONT2INS1060 to its finish)', 12, 0, 0, 0, 0, '5.10', '3 people x 4 working days (14, 15, 16, 17-Nov) - assessed gang')
dadd('H2', 'Tank 2 roof supports and ladder, 18 to 23-Nov (QCD18TSECONT2INS1040)', 6, 9, 0, 0, 0, '5.10', '3 people: 18, 19-Nov in W13; 21, 22, 23-Nov in W14 (20-Nov is a Friday)')
dadd('H3', 'Tank 1 hydrostatic test P18, 17 to 24-Nov: tankered fill 17 to 22-Nov; hold 23-Nov; inspection 24-Nov', 6, 4, 0, 0, 0, '5.10', '2 people on the 5 fill days (17, 18, 19-Nov in W13; 21, 22-Nov in W14). The hold and the inspection are attended by the QA/QC inspector (1.3), the commissioning engineer (7.13) and the supplier - no helper')
dadd('H4', 'Tank 2 roof panels, 24 to 30-Nov (QCD18TSECONT2INS1050) - 5 people on every day, the gang the Contractor deployed for the Tank 1 roof', 0, 15, 15, 0, 0, '5.10', '5 people x 6 days: 24, 25, 26-Nov (W14); 28, 29, 30-Nov (W15). Not reduced to the approved W15 figure - see the departure row')
dadd('H5', 'Tank 2 nozzles and internals attendance, 28-Nov to 02-Dec (QCD18TSECONT2MW2030)', 0, 0, 5, 0, 0, '5.10', "1 person x 5 working days hoisting spools and flanges for the supplier's fitters (about 3 to 4 hours a day, rounded to a person) - assessed")
dadd('H6', 'Tank 1 test water: dosing and circulation 01 to 02-Dec (2 people), contact time unattended, sampling 03-Dec (1 person)', 0, 0, 5, 0, 0, '5.10', '2 x 2 + 1 = 5 man-days. Method: provisional, see the water sequence in Section 3; chemicals at 7.6')
dadd('H7', 'Transfer Tank 1 to Tank 2, P19, 05 to 07-Dec - pump and hose attendance', 0, 0, 0, 6, 0, '5.10', "2 people x 3 days; so 'Build-Up' line 7.4 is nil")
dadd('H8', 'Tank 2 test P20: top-up 08-Dec (1), hold 09-Dec (none), inspection 10-Dec (1), sampling and inspection 12-Dec (2) - 4 working days elapsed (11-Dec is a Friday), 3 attended', 0, 0, 0, 2, 2, '5.10 / 7.8', '08 and 10-Dec within W16; 12-Dec is after the histogram ends and is priced at 7.8')
dadd('H9', 'Dechlorination of the Tank 2 water on 13-Dec before the pumping demonstration (1 person); the water is then kept in Tank 2 as first stock, not discharged', 0, 0, 0, 0, 1, '7.8', 'After the histogram; chemicals at 7.9. Provisional method, see the water sequence')
dadd('H11', 'Tank 1 operating fill by tanker, 09 to 10-Dec (P20b) - 1 person attending the tanker discharge while Tank 2 is on its hold', 0, 0, 0, 2, 0, '5.10', "1 person x 2 working days within the approved W16 people (2 a day): 09-Dec with nobody else, 10-Dec with the H8 inspection - see the daily check. Water at 'Build-Up' 7.24")
dadd('H10', 'Integrated commissioning 13 to 15-Dec (P21b) and demobilisation 16 to 23-Dec (P23)', 0, 0, 0, 0, 0, '7.13 to 7.17; 1.16', 'Specialists only (engineer, technicians, electrician, calibration); demobilisation clean-up 4 x 3 days at 1.16; no helpers')
if PG['used'] + 200 > PAGE:
    pg_break(pg, r); hdrd(pg, r); r += 1
dem = r
pg.cell(r, 2, 'Demand - helper man-days by week'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
for col in range(3, 9):
    L_ = get_column_letter(col); pg.cell(r, col, f"=SUM({L_}{DUTY['H1']}:{L_}{DUTY['H10']})"); cp(S_TOTAMT, pg.cell(r, col)); pg.cell(r, col).number_format = '#,##0'
pg.cell(r, 10, "The 'after 11-Dec' column is carried to 'Build-Up' line 7.8"); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 19.5; pg_add(19.5); r += 1
capr = r
pg.cell(r, 2, 'Available - approved helper man-days of the week (Section 4 weekly table)'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
for col, wkno in ((3, 13), (4, 14), (5, 15), (6, 16)):
    pg.cell(r, col, f"=H{wk[wkno]}"); cp(S_TOTAMT, pg.cell(r, col)); pg.cell(r, col).number_format = '#,##0'
pg.cell(r, 7, 0); cp(S_TOTAMT, pg.cell(r, 7)); pg.cell(r, 8, f"=SUM(C{r}:G{r})"); cp(S_TOTAMT, pg.cell(r, 8)); pg.cell(r, 8).number_format = '#,##0'
pg.cell(r, 10, 'People x working days of the week: W13 6 x 6, W14 5 x 6, W15 3 x 6, W16 2 x 6 (05 to 10-Dec; 11-Dec is a Friday)'); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 30; pg_add(30); r += 1
chkr = r
pg.cell(r, 2, 'Spare (available less demand) - must not be negative'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
for col in range(3, 9):
    L_ = get_column_letter(col); pg.cell(r, col, f"={L_}{capr}-{L_}{dem}"); cp(S_TOTAMT, pg.cell(r, col)); pg.cell(r, col).number_format = '#,##0;-#,##0'
pg.cell(r, 10, "Spare is unused approved capacity; it is not deducted - the approved people are paid as approved. A negative week is a shortfall, priced as 'Additional helpers assessed' below; the negative 'after' figure is labour outside the histogram, priced at 7.8"); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 30; pg_add(30); r += 1
depr = r
pg.cell(r, 2, "Additional helpers assessed - person-days above the approved people on the days concerned (daily check), priced at 5.10"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 8, 0); cp(S_TOTAMT, pg.cell(r, 8)); pg.cell(r, 8).number_format = '#,##0'
pg.cell(r, 10, "28, 29 and 30-Nov: needed 6 (roof panels 5 + nozzle attendance 1), approved 3; additional 3 a day x 3 days = 9 person-days, added to 5.10. Counted day by day - spare people on other days cannot cover a peak. The 5-person roof gang is an assessment assumption (the gang SAMA used on the Tank 1 roof), not a measured productivity"); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 57; pg_add(57); r += 1
AFTER_CELL = f"G{dem}"
DEP_CELL = f"H{depr}"
# --- daily peak-demand check, W13 to W16
r += 1
pg_break_if(pg, r, 30 + 32 + 4 * 16)
HDRY = ['Ref', 'Working day - duties in progress (from the rows above)', 'Helpers needed', 'Approved a day', 'Spare', '', '', '', '', 'Note']
def hdry(ws, r_):
    header(ws, r_, HDRY[:9]); ws.cell(r_, 10, HDRY[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=6, end_row=r_, end_column=9)
para(pg, r, ("Daily check - a week can balance while a single day does not\n"
    "• Each working day lists the duties running that day and the helpers they need, against the approved people of that week. Spare = unused approved people, not a deduction.\n"
    "• A day needing more people than approved is a shortfall; the shortfalls are added and priced as 'Additional helpers assessed' (5.10). No duty is trimmed to fit."), height=44); r += 1
hdry(pg, r); pg_add(PG['hdr']); r += 1
DAILY = [
 ('14-Nov-2026', 'H1 bracing 3', 3, 13, ''), ('15-Nov-2026', 'H1 bracing 3', 3, 13, ''), ('16-Nov-2026', 'H1 bracing 3', 3, 13, ''),
 ('17-Nov-2026', 'H1 bracing 3; H3 fill 2', 5, 13, ''), ('18-Nov-2026', 'H2 roof supports 3; H3 fill 2', 5, 13, ''), ('19-Nov-2026', 'H2 roof supports 3; H3 fill 2', 5, 13, ''),
 ('21-Nov-2026', 'H2 roof supports 3; H3 fill 2', 5, 14, ''), ('22-Nov-2026', 'H2 roof supports 3; H3 fill 2', 5, 14, ''), ('23-Nov-2026', 'H2 roof supports 3; Tank 1 hold - none', 3, 14, ''),
 ('24-Nov-2026', 'H4 roof panels 5; Tank 1 inspection by specialists', 5, 14, ''), ('25-Nov-2026', 'H4 roof panels 5', 5, 14, ''), ('26-Nov-2026', 'H4 roof panels 5', 5, 14, ''),
 ('28-Nov-2026', 'H4 roof panels 5; H5 nozzles 1', 6, 15, 'Above the approved 3: departure'), ('29-Nov-2026', 'H4 roof panels 5; H5 nozzles 1', 6, 15, 'Above the approved 3: departure'), ('30-Nov-2026', 'H4 roof panels 5; H5 nozzles 1', 6, 15, 'Above the approved 3: departure'),
 ('01-Dec-2026', 'H5 nozzles 1; H6 dosing 2', 3, 15, ''), ('02-Dec-2026', 'H5 nozzles 1; H6 dosing 2', 3, 15, ''), ('03-Dec-2026', 'H6 sampling 1', 1, 15, ''),
 ('05-Dec-2026', 'H7 transfer 2', 2, 16, ''), ('06-Dec-2026', 'H7 transfer 2', 2, 16, ''), ('07-Dec-2026', 'H7 transfer 2', 2, 16, ''),
 ('08-Dec-2026', 'H8 top-up 1', 1, 16, ''), ('09-Dec-2026', 'Tank 2 hold - none; H11 Tank 1 operating fill 1', 1, 16, ''), ('10-Dec-2026', 'H8 inspection 1; H11 Tank 1 operating fill 1', 2, 16, 'Equal to the approved 2: no departure'),
]
d_first = r
for day, duties, need, wkno, note in DAILY:
    if PG['used'] + 16 > PAGE:
        pg_break(pg, r); hdry(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(16)
    pg.cell(r, 1, day[:6]); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, f'{day} - {duties}'); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, need); numcell(pg.cell(r, 3), '#,##0'); pg.cell(r, 4, f"=G{wk[wkno]}"); numcell(pg.cell(r, 4), '#,##0'); pg.cell(r, 5, f"=D{r}-C{r}"); numcell(pg.cell(r, 5), '#,##0;-#,##0')
    for c in (6, 7, 8, 9): cp(S_UNIT, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
    pg.cell(r, 10, note); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 16; r += 1
d_last = r - 1
if PG['used'] + 30 > PAGE:
    pg_break(pg, r); hdry(pg, r); r += 1
pg.cell(r, 2, 'Additional helpers assessed - sum of the daily shortfalls, taken to the weekly table and to 5.10'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 6, 7, 8, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 5, f'=-SUMIF(E{d_first}:E{d_last},"<0")'); cp(S_TOTAMT, pg.cell(r, 5)); pg.cell(r, 5).number_format = '#,##0'
pg.cell(r, 10, "28, 29 and 30-Nov: 3 people a day above the approved deployment"); cp(S_BASIS, pg.cell(r, 10)); pg.row_dimensions[r].height = 19.5; pg_add(19.5)
pg.cell(depr, 8).value = f"=E{r}"
r += 1

pg.cell(bt2, 9).value = pg.cell(bt2, 9).value.replace('DEPARTURE', DEP_CELL)
# --- skilled people: allocation by weeks and activities in progress
r += 1
pg_break_if(pg, r, 32 + 3 * 30)
HDRK = ['Ref', 'Skilled people - weeks and activities in progress', 'Weeks', '', 'Person-weeks', 'Whose people, and where paid', '', '', '', 'Basis and what is unresolved']
def hdrk(ws, r_):
    header(ws, r_, HDRK[:9]); ws.cell(r_, 10, HDRK[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=3, end_row=r_, end_column=4); ws.merge_cells(start_row=r_, start_column=6, end_row=r_, end_column=9)
hdrk(pg, r); pg_add(PG['hdr']); r += 1
SK = {}
def kadd(ref, grp, weeks, mw, alloc, status):
    global r
    hh = max(30, est(grp, 46), est(alloc, 52), est(status, 52))
    if PG['used'] + hh > PAGE:
        pg_break(pg, r); hdrk(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(hh)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, grp); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, weeks); cp(S_UNIT, pg.cell(r, 3)); cp(S_UNIT, pg.cell(r, 4)); pg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    pg.cell(r, 5, mw); numcell(pg.cell(r, 5), '#,##0')
    pg.cell(r, 6, alloc); cp(S_BASIS, pg.cell(r, 6))
    for c in (7, 8, 9): cp(S_BASIS, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
    pg.cell(r, 10, status); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = hh; SK[ref] = r; r += 1
kadd('K1', 'Mobilisation (22 to 26-Aug) and dismantling start (27-Aug)', 'W1', f"=F{wk[1]}", "SAMA's mobilisation riggers, or the Item 3 dismantling crew", "Unresolved: if SAMA's own people, they are in no priced line (bounded at S5)")
kadd('K2', 'Dismantling in progress (QCD18TSECONDSM1020 to 1030)', 'W2 to W3', f"=F{wk[2]}+F{wk[3]}", 'Dismantling crew, inside the Item 3 quotation (Al Mousa S04647)', 'Provisional: quotation scope as recorded, not attached')
kadd('K3', 'Tank 1 erection only (QCD18TSECONT1INS1020 to 1060)', 'W4 to W9', f"=SUM(F{wk[4]}:F{wk[9]})", "Supplier's erection crew, inside Item 6", "The supplier installs (Assumption 9); provisional until the Al Mousa offer is produced")
kadd('K4', 'Two erection fronts (Tank 1 bracing, roof; Tank 2 base, walls, bracing)', 'W10 to W11', f"=F{wk[10]}+F{wk[11]}", "Two supplier erection crews, inside Item 6", 'As K3')
kadd('K5', 'Erection, mechanical works and pipework installation overlapping (QCD18TSECONT1MW2055 from 12-Nov; QCD18TSECONT2MW2020)', 'W12 to W15', f"=SUM(F{wk[12]}:F{wk[15]})", "Supplier's erection crews (Item 6) and pipework fitters (Item 8 installed rates)", 'Split between Items 6 and 8 unresolved; both prices include their own labour, so nothing extra either way')
kadd('K6', 'Tie-in, transfer and tests (QCD18TSECONTC2040, P19, P20)', 'W16', f"=F{wk[16]}", 'Tie-in fitters (8.14, 8.15); test supervision (Item 6)', 'As K5')
btk = btot("Skilled man-weeks allocated (equals the approved 136); none priced separately", {5: f"=SUM(E{SK['K1']}:E{SK['K6']})"}, "If any group proves to be the Contractor's own people outside the recorded scopes, the effect is bounded at S5 and Item 6 would need re-basing to a supply-only price")
# --- plant, formula-linked to the XER activity IDs
r += 1
if PG['used'] + 32 + 4 * 45 > PAGE:
    pg_break(pg, r)
HDRP = ['Ref', 'Plant and basis', 'From (XER)', 'To (XER)', 'Working days', 'Hire days carried', '', '', '', 'Derivation from the XER activities; priced item']
def hdrp(ws, r_):
    header(ws, r_, HDRP[:9]); ws.cell(r_, 10, HDRP[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=7, end_row=r_, end_column=9)
hdrp(pg, r); pg_add(PG['hdr']); r += 1
PL = {}
def padd(ref, name, frm, to, wd, hire, note, h=None):
    global r
    hh = max(30, est(name, 46), est(note, 52), h or 0)
    if PG['used'] + hh > PAGE:
        old_r = r
        pg_break(pg, r); hdrp(pg, r); r += 1; pg_add(PG['hdr'])
        fix = lambda v: re.sub(r'(?<![A-Z$])([C-H])' + str(old_r) + r'(?!\d)', lambda m: m.group(1) + str(r), v) if isinstance(v, str) else v
        frm, to, wd, hire = map(fix, (frm, to, wd, hire))
    pg_add(hh)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, name); cp(S_DESC, pg.cell(r, 2))
    for col, v in ((3, frm), (4, to)):
        c = pg.cell(r, col, v); datecell(c)
        if v in (None, ''): c.value = '-'
    for col, v in ((5, wd), (6, hire)):
        c = pg.cell(r, col, v); numcell(c, '#,##0')
        if v in (None, ''): c.value = '-'
    for c in (7, 8, 9): cp(S_UNIT, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    pg.cell(r, 10, note); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = hh; PL[ref] = r; r += 1
    return r - 1
pl1 = padd('T1', 'Telehandler - continuous daily hire while panels are received or placed: first delivery to the Tank 2 walls finish', f"={AS('QCD18TSEPRC1120')}", f"={AF('QCD18TSECONT2INS1030')}",
           f"={WD(f'C{r}', f'D{r}')}", f"=E{r}", "QCD18TSEPRC1120 start to QCD18TSECONT2INS1030 finish: deliveries, base and wall panels of both tanks and the Tank 1 roof follow one another without a gap, so one unit is on hire every working day. One unit is a provisional utilisation assumption pending the plant schedule (S6)")
pl2 = padd('T2', 'Telehandler - Tank 2 roof supports (half the days) and roof panels', f"={AS('QCD18TSECONT2INS1040')}", f"={AF('QCD18TSECONT2INS1050')}",
           f"={WD(f'C{r}', f'D{r}')}", f"=ROUND({AWD('QCD18TSECONT2INS1040')}/2,0)+{AWD('QCD18TSECONT2INS1050')}", "QCD18TSECONT2INS1040 and QCD18TSECONT2INS1050; the Tank 2 bracing days between T1 and T2 need no telehandler")
plt = padd('T3', "Telehandler hire days carried - 'Build-Up' line 5.2", '-', '-', '-', f"=F{pl1}+F{pl2}", "Day-rate hire on working days on site; mobilisation within the day rate")
pb1 = padd('B1', 'Boom truck - alternate days through the walls-and-bracing window of each tank (upper tiers, bracing members)', f"={AS('QCD18TSECONT1INS1030')}", f"={AF('QCD18TSECONT2INS1060')}",
           f"={WD(AS('QCD18TSECONT1INS1030'), AF('QCD18TSECONT1INS1060'))}+{WD(AS('QCD18TSECONT2INS1030'), AF('QCD18TSECONT2INS1060'))}", f"=ROUND({WD(AS('QCD18TSECONT1INS1030'), AF('QCD18TSECONT1INS1060'))}/2,0)+ROUND({WD(AS('QCD18TSECONT2INS1030'), AF('QCD18TSECONT2INS1060'))}/2,0)",
           "QCD18TSECONT1INS1030 start to QCD18TSECONT1INS1060 finish, and the same for Tank 2; half the days - an assessed assumption for intermittent lifting")
pb2 = padd('B2', 'Boom truck - every day of the roof supports and roof panels of both tanks', f"={AS('QCD18TSECONT1INS1040')}", f"={AF('QCD18TSECONT2INS1050')}",
           f"={AWD('QCD18TSECONT1INS1040')}+{AWD('QCD18TSECONT1INS1050')}+{AWD('QCD18TSECONT2INS1040')}+{AWD('QCD18TSECONT2INS1050')}", f"=E{r}", "QCD18TSECONT1INS1040, QCD18TSECONT1INS1050, QCD18TSECONT2INS1040, QCD18TSECONT2INS1050")
pb3 = padd('B3', 'Less days on which the Tank 1 roof (B2) coincides with a Tank 2 alternate lifting day (B1) - one unit', f"=MAX({AS('QCD18TSECONT1INS1040')},{AS('QCD18TSECONT2INS1030')})", f"=MIN({AF('QCD18TSECONT1INS1050')},{AF('QCD18TSECONT2INS1060')})",
           f"={WD(f'C{r}', f'D{r}')}", f"=-ROUND(E{r}/2,0)", "Overlap of the two windows from the XER dates; half of it is already counted in B1")
pb4 = padd('B4', 'Boom truck - offloading on delivery days outside the lifting days above', '-', '-', 12, f"=E{r}", "Assessed assumption: 4 Tank 1 base batches, 2 of the 4 wall-panel batches, 4 of the 8 Tank 2 loads and 2 pipe loads fall on days with no other lift; truck-arrival days are not in any document received")
pbt = padd('B5', "Boom truck hire days carried - 'Build-Up' line 5.1", '-', '-', '-', f"=F{pb1}+F{pb2}+F{pb3}+F{pb4}", "Day-rate hire on working days on site")

# site staff roles
if PG['used'] + 32 + 5 * 40 > PAGE:
    pg_break(pg, r)
HDRR = ['Ref', 'Site staff role and phase duties', 'From', 'To', 'Calendar days', 'Months', '', '', '', 'Coverage and source; priced item']
header(pg, r, HDRR[:9]); pg.cell(r, 10, HDRR[9]); cp(S_HDR, pg.cell(r, 10)); pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9); pg_add(PG['hdr']); r += 1
ROLE = {}
def radd(ref, role, frm, to, src):
    global r
    h = max(30, est(role, 46), est(src, 52))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); header(pg, r, HDRR[:9]); pg.cell(r, 10, HDRR[9]); cp(S_HDR, pg.cell(r, 10)); r += 1; pg_add(PG['hdr'])
    pg_add(h)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, role); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, frm); datecell(pg.cell(r, 3)); pg.cell(r, 4, to); datecell(pg.cell(r, 4))
    pg.cell(r, 5, f"=D{r}-C{r}+1"); numcell(pg.cell(r, 5), '#,##0')
    pg.cell(r, 6, f"=ROUND(E{r}/{MON},1)"); numcell(pg.cell(r, 6), '0.0')
    pg.cell(r, 6).border = Border(left=Side(style='medium', color='FFC425'), right=Side(style='medium', color='FFC425'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    for c in (7, 8, 9): cp(S_UNIT, pg.cell(r, c)); pg.cell(r, c).value = None
    pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    pg.cell(r, 10, src); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = h
    ROLE[ref] = r; r += 1
radd('R1', 'Site engineer - by phase: mobilisation and permits (Aug); dismantling supervision and handover records (Sep); deliveries, inspection requests and two erection fronts (Sep to Nov); tie-in, tests and commissioning coordination (Nov to Dec); demobilisation and close-out (Dec)', f"=F{d1}", f"=G{d1}",
     "Dedicated for the site period D1: every phase has a duty that needs a responsible engineer present on a site outside the D-18 boundary, and no phase is quiet enough to share the role. Priced at 1.1")
radd('R2', 'HSE officer - by phase: lifting and demolition controls (dismantling); working at height, lifting and two fronts (erection); confined-space entry and water handling (tests); lifting and traffic (demobilisation)', f"=F{d1}", f"=G{d1}",
     "Dedicated for the site period D1: each phase carries a high-risk activity under the project HSE plan (RFP Scope of Works section 4, item 7), and the quietest phase, testing and commissioning, still has confined-space work. Priced at 1.2")
radd('R3', 'QA/QC inspector - material inspection requests, erection inspections and bolt-torque verification, hydrostatic test witnessing and records', f"=F{p2}", f"=G{p22}",
     "From the start of dismantling to completion of testing and commissioning; not needed during mobilisation or demobilisation. Priced at 1.3")
radd('R4', "Works foreman - the Contractor's own labour: helpers, deliveries, scaffold and access, temporary works, pipework crew, demobilisation", f"=F{d1}", f"=G{d1}",
     "Full-time for the site period D1: Contractor labour is on site from mobilisation to demobilisation. Priced at 1.4")
radd('R5', 'Storekeeper / timekeeper - receipt of panels, accessories and pipework, storage and issue to the erection fronts, attendance records', f"={AS('QCD18TSEPRC1120')}", f"=G{p22}",
     "From the first delivery (08-Sep-2026) to completion of testing; not needed before deliveries or during demobilisation. Priced at 1.5")

# ================================================================ Section 5: Build-Up lines
r += 1
if PG['used'] + 21 + 44 + 32 + 60 > PAGE:
    pg_break(pg, r)
pg_add(21 + 44 + 32)
banner(pg, r, "5. QUANTITIES TAKEN FROM THE PROGRAMME - THE FIGURES USED ON THE 'BUILD-UP' TAB, WITH THE SUBMITTED PROGRAMME FOR COMPARISON"); r += 1
para(pg, r, ("How to read this table\n"
    "• Column G (yellow border): the quantity used on the 'Build-Up' tab. Click column A to open the line; click the figure in G to open the row it is built from (column J names it).\n"
    "• Grey columns: comparison only - the Rev 01 quantity, and the quantity on the SAMA Submitted Programme exactly as submitted (parallel testing, every submitted day paid). Lines not listed are unchanged from Rev 01."), height=58); r += 1
hdr4(pg, r)
HDR4ROW = r; r += 1

def MO(days_cell): return f"ROUND({days_cell}/{MON},1)"

LINES = []  # (ref, build-up row, qtyA formula, qtyB formula, derivation)
def L(ref, burow, qA, qB, deriv, prev=None):
    LINES.append((ref, burow, qA, qB, deriv, prev))

wm_A, wm_B = f"E{d2}", f"H{d2}"       # works period days
sp_A, sp_B = f"E{d1}", f"H{d1}"       # site period days
er_A, er_B = f"E{d3}", f"H{d3}"
ac_A, ac_B = f"E{d6}", f"H{d6}"
ct_A, ct_B = f"E{d7}", f"H{d7}"
for ref, row_, role in (('1.1', 9, 'R1'), ('1.2', 10, 'R2'), ('1.3', 11, 'R3'), ('1.4', 12, 'R4'), ('1.5', 13, 'R5')):
    L(ref, row_, f"={MO(sp_A)}", f"=F{ROLE[role]}", f"Section 4, role {role}: phase coverage stated there. As submitted: the whole submitted site period")
for ref, row_ in (('1.6', 14), ('1.7', 15), ('1.8', 16)):
    L(ref, row_, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.9', 17, f"=2*ROUNDUP({sp_A}/7,0)", f"=2*ROUNDUP({sp_B}/7,0)", "Two 10 m3 deliveries a week for the weeks of the site period D1, rounded up to whole weeks")
L('1.10', 18, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.11', 19, f"=2*{MO(ct_A)}", f"=2*{MO(ct_B)}", "2 No. containers for the storage window D7 in months")
L('1.12', 20, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months, transport being needed while the workforce demobilises; the manpower histogram shows labour on site for 16 weeks, 28-Aug to 11-Dec-2026, inside this period")
L('1.13', 21, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months; still subject to whether the Employer's security covers the lower plateau")
L('5.1', 75, f"=F{pbt}", f"=F{pbt}", "Section 4, boom truck hire days B5, built from the XER activity dates; the same as submitted")
L('5.2', 76, f"=F{plt}", f"=F{plt}", "Section 4, telehandler hire days T3, built from the XER activity dates; the same as submitted")
L('5.4', 78, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "2 No. towers for the access window D6, rounded up to whole hire months")
L('5.6', 80, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "As 5.4")
L('5.7', 81, f"={MO(wm_A)}", f"={MO(wm_B)}", "Works period D2 in months - daytime works power to completion of commissioning; none during demobilisation")
L('5.8', 82, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months - welfare power runs continuously to the end of demobilisation")
L('5.9', 83, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months")
L('5.10', 84, f"=I{bt2}", f"=I{bt2}", "Section 4, the approved histogram helper man-days converted at 26 working days a month; the same as submitted")
L('5.11', 85, f"={MO(er_A)}", f"={MO(er_B)}", "Erection window D3 in months")
L('5.14', 88, f"=2*{MO(er_A)}", f"=2*{MO(er_B)}", "2 No. lighting towers for the erection window D3 in months")
L('7.1', 101, "=2*3774", "=3774", "As submitted: both tanks filled at once for parallel testing (7,548 m3, RFP Scope of Works section 1). Carried: one fill, water re-used for the second tank - the Engineer's (KEO) email of 30-Aug-2026")
L('7.2', 102, "=0", f"=H{p19}+H{p20}", "Transfer pump: not needed as submitted; carried for the transfer P19 and the Tank 2 test P20 (top-up and hold)")
L('7.3', 103, "=0", "=1", "Transfer hoses: not needed as submitted; 1 week carried")
L('7.4', 104, "=0", "=0", "Nil as submitted and carried: the transfer labour is within the approved histogram helpers priced at 5.10 (Section 4, week W16)")
L('7.5', 105, "=ROUND(2*3774*0.1,0)", "=ROUND(3774*0.1,0)", "Top-up at 10 per cent of the water filled: of two fills on the submitted programme, of one fill in the assessment allowance (retention in Tank 1 between P18 and P19)")
L('7.13', 113, f"=E{p17}+1", f"=H{p21}+H{p21b}+2", "As submitted: the programmed 7-working-day parallel testing and commissioning activity plus one day at the hold. Carried: component checks P21 (3 days) and integrated commissioning P21b (3 days), both assessed assumptions, plus one day at each hydrostatic test hold; the fills and holds themselves are supervised by the QA/QC inspector (line 1.3) and the supplier's leak-test supervision within Item 6")
L('7.14', 114, f"=2*E{p17}", f"=2*(H{p21}+H{p21b})", "2 No. technicians: as submitted for the 7-day programmed activity; allowed for the component checks P21 and the integrated commissioning P21b (6 days)")
L('7.16', 116, f"=E{p16}+1", f"=H{p16}+1", "Tie-in P16 working days plus one day of integrated commissioning, both bases")
L('7.24', 124, "=0", "=34*30*1", "Assessment allowance: 34 m x 30 m x 1.0 m operating depth in Tank 1 for the commissioning demonstration, tankered on 09 and 10-Dec-2026 (window P20b). Nil on the submitted programme, which fills both tanks in full for parallel testing")
L('7.22', 122, "=4", "=4", "Pump and hose set standing by through the two 24-hour holds and two days of contingency between the Tank 1 test and the transfer; tankers are working, not standing by, during the fill. The same on A, where two simultaneous fills need the same standby")

first4 = r
item_rows = {}
for ref, burow, qA, qB, deriv, prev in LINES:
    item = ref.split('.')[0]
    if item not in item_rows:
        item_rows[item] = []
    h4 = max(30, est(deriv, 52), est(str(bu.cell(burow, 2).value), 46))
    if PG['used'] + h4 > PAGE:
        pg_break(pg, r); hdr4(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(h4)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, f"='Build-Up'!B{burow}"); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, f"='Build-Up'!C{burow}"); cp(S_UNIT, pg.cell(r, 3))
    pg.cell(r, 4, f"='Build-Up'!E{burow}"); cp(S_RATE, pg.cell(r, 4))
    prev_val = bu.cell(burow, 4).value
    pg.cell(r, 5, prev_val); numcell(pg.cell(r, 5), '#,##0.00')
    pg.cell(r, 6, qA); numcell(pg.cell(r, 6), '#,##0.00')
    pg.cell(r, 7, qB); numcell(pg.cell(r, 7), '#,##0.00')
    pg.cell(r, 8, f"=ROUND(F{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 8))
    pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9))
    pg.cell(r, 10, deriv); cp(S_BASIS, pg.cell(r, 10))
    pg.cell(r, 7).border = Border(left=Side(style='medium', color='FFC425'), right=Side(style='medium', color='FFC425'),
                                  top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    pg.row_dimensions[r].height = h4
    item_rows[item].append(r)
    bu.cell(burow, 4).value = f"=Programme!$G${r}"
    r += 1
last4 = r - 1
if PG['used'] + 4 * 19.5 + 30 > PAGE:
    pg_break(pg, r); hdr4(pg, r); r += 1; pg_add(PG['hdr'])
# sensitivity S3 amount
d1_rows = [rr for rr in range(first4, last4 + 1) if pg.cell(rr, 1).value in ('1.1', '1.2', '1.4', '1.6', '1.7', '1.8', '1.10', '1.12', '1.13', '5.8', '5.9', '5.7', '1.3', '1.5')]
dl_row = [rr for rr in range(first4, last4 + 1) if pg.cell(rr, 1).value == '1.9'][0]
def _days(rr):
    ref = pg.cell(rr, 1).value
    if ref == '5.7': return f"({b_ic_f}-F{p1}+1)"
    if ref == '1.3': return f"({b_ic_f}-F{p2}+1)"
    if ref == '1.5': return f"({b_ic_f}-{AS('QCD18TSEPRC1120')}+1)"
    return f"(H{d1}+C{s3})"
pg.cell(s3, 5).value = ('=' + '+'.join(f"D{rr}*(ROUND({_days(rr)}/{MON},1)-G{rr})" for rr in d1_rows)
                        + f"+D{dl_row}*(2*ROUNDUP((H{d1}+C{s3})/7,0)-G{dl_row})")
cp(S_AMT, pg.cell(s3, 5)); pg.cell(s3, 5).number_format = '#,##0.00'
sub = {}
if PG['used'] + 5 * 30 > PAGE:
    pg_break(pg, r); hdr4(pg, r); r += 1
for item, label in (('1', 'Item 1 - lines listed above'), ('5', 'Item 5 - lines listed above'), ('7', 'Item 7 - lines listed above')):
    rows_ = item_rows[item]
    pg.cell(r, 2, label); cp(S_TOTLBL, pg.cell(r, 2))
    for c in (1, 3, 4, 5, 6, 7, 10): cp(S_TOTLBL, pg.cell(r, c))
    pg.cell(r, 8, '=' + '+'.join(f"H{x}" for x in rows_)); cp(S_TOTAMT, pg.cell(r, 8))
    pg.cell(r, 9, '=' + '+'.join(f"I{x}" for x in rows_)); cp(S_TOTAMT, pg.cell(r, 9))
    pg.cell(r, 10, f"Sum of the Item {item} lines in this table only; the full item total is on the 'Build-Up' tab"); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = 19.5
    sub[item] = r; r += 1
pg.cell(r, 2, 'All lines listed above'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 8, f"=SUM(H{first4}:H{last4})"); cp(S_TOTAMT, pg.cell(r, 8))
pg.cell(r, 9, f"=SUM(I{first4}:I{last4})"); cp(S_TOTAMT, pg.cell(r, 9))
pg.cell(r, 10, 'Difference between the two bases before Overhead and Profit'); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 19.5
ALL4 = r; r += 1
pg.cell(r, 2, 'Assessment total on each basis, VAT excluded, including 5 per cent Overhead and Profit'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 8, f"=ROUND((Assessment!$K$19-I{ALL4}+H{ALL4})*1.05,2)"); cp(S_TOTAMT, pg.cell(r, 8))
pg.cell(r, 9, "=Assessment!$K$21"); cp(S_TOTAMT, pg.cell(r, 9))
pg.cell(r, 10, "Carried = the 'Assessment' tab total (row 21). 'As submitted' swaps in the listed lines' as-submitted amounts and adds the 5 per cent once"); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 30
TOT4 = r; r += 1
for c in (1, 3, 4, 5, 6, 7, 10): pass
if PG['used'] + 19.5 + 4 * 68 > PAGE:
    pg_break(pg, r); hdr4(pg, r); r += 1
pg.cell(r, 2, "Sensitivities on the resource matrix - not carried"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5; r += 1
pg.cell(r, 1, 'S5'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, "Upper bound if all 136 approved skilled man-weeks were the Contractor's own people outside the recorded Item 3, 6 and 8 scopes (gross labour, before any re-basing of Item 6 to a supply-only price)"); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'man-day'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E51"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, f"=E{btk}*6"); numcell(pg.cell(r, 7), '#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9))
pg.cell(r, 10, "At the rigger day rate assessed at 'Build-Up' c.2; a bound only, not carried: the Contractor's answer to 'Build-Up Comparison' Section 7, item 10 decides, and Item 6 would fall if the supplier's price proved to be supply only, contrary to the scope described in 'Build-Up' Assumption 9"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 44; r += 1
pg.cell(r, 1, 'S7'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, "Reduction if the helpers approved in the dismantling weeks W1 to W3 (13 man-weeks) prove to be within the Item 3 dismantling crew rather than the Contractor's own attendance"); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'man-month'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E84"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, f"=-ROUND((H{wk[1]}+H{wk[2]}+H{wk[3]})/26,1)"); numcell(pg.cell(r, 7), '#,##0.00;-#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9)); pg.cell(r, 9).number_format = '#,##0.00;-#,##0.00'
pg.cell(r, 10, "Excluding Overhead and Profit; not carried. Those helpers are priced provisionally at 5.10 because the approved histogram lists them as the Contractor's direct manpower; the Al Mousa quotation scope (not attached) decides"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 44; r += 1
pg.cell(r, 1, 'S8'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, "Tank 1 filled to the full 3.7 m test level for the commissioning instead of the 1 m operating fill at 7.24 (2,754 m3 more by tanker, two further attended days), if the Engineer requires it and no network source is confirmed"); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'm3'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E101"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, "='Build-Up'!D101-'Build-Up'!D124"); numcell(pg.cell(r, 7), '#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)+2*'Build-Up'!E122"); cp(S_AMT, pg.cell(r, 9))
pg.cell(r, 10, "Excluding Overhead and Profit; not included in the assessment. The provisional commissioning method in Section 3 explains the 1 m operating fill and why Tank 1 is otherwise empty at commissioning"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 44; r += 1
pg.cell(r, 1, 'S6'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, 'A second telehandler for the 11 concurrent days (Tank 1 roof with Tank 2 walls, 29-Oct to 11-Nov-2026), if one unit proves insufficient'); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'day'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E76"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, 11); numcell(pg.cell(r, 7), '#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9))
pg.cell(r, 10, "Excluding Overhead and Profit. To be decided on the Contractor's plant schedule and method statement"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 30; r += 1
# ================================================================ Section 6: WBS link register - every priced 'Build-Up' line
# (ref, Build-Up row, WBS or site service, what it buys / when / who, resource or hire basis, elsewhere?)
SITE = 'Site service - whole site period D1 (QCD18TSEMOB1240 to QCD18TSEDMOB1020)'
REG = [
 ('1.1', 9, 'Site service - all phases; role R1', 'Management of the separate tank site from mobilisation to demobilisation; Contractor', 'Dedicated, monthly, site period', 'No - supplier supervises its own erection crew within Item 6'),
 ('1.2', 10, 'Site service - all phases; role R2', 'HSE control of dismantling, erection, testing and demobilisation; Contractor', 'Dedicated, monthly, site period', 'No'),
 ('1.3', 11, 'QCD18TSECONDSM1020 to QCD18TSECONTCT22030; role R3', 'Material inspections, erection inspections, test witnessing; Contractor', 'Dedicated, monthly, dismantling start to completion of testing', 'No'),
 ('1.4', 12, 'Site service - all phases; role R4', "Supervision of the Contractor's own labour and plant; Contractor", 'Dedicated, monthly, site period', 'No - not the supplier\'s erection foreman'),
 ('1.5', 13, 'QCD18TSEPRC1120 to QCD18TSECONTCT22030; role R5', 'Receipt, storage and issue of materials; Contractor', 'Dedicated, monthly, first delivery to completion of testing', 'No'),
 ('1.6', 14, SITE, 'Site office; Contractor', 'Monthly hire, site period', 'No'),
 ('1.7', 15, SITE, 'Welfare cabin; Contractor', 'Monthly hire, site period', 'No'),
 ('1.8', 16, SITE, 'WC unit; Contractor', 'Monthly hire with servicing, site period', 'No'),
 ('1.9', 17, SITE, 'Welfare and general site water, no mains supply; Contractor', 'Quantity: two 10 m3 deliveries a week, weeks rounded up', 'No - test water is 7.1, 7.5 and 8.17'),
 ('1.10', 18, SITE, 'Site water storage tank; Contractor', 'Monthly hire, site period', 'No'),
 ('1.11', 19, 'QCD18TSEPRC1120 to QCD18TSECONT2INS1050 (window D7)', 'Secure storage of accessories, bolts and small items from the first delivery to the last roof panel; Contractor', 'Container-months, 2 No.', 'No - panels stay on pallets in the open'),
 ('1.12', 20, SITE, 'Daily workforce bus to the separate site; Contractor', 'Monthly, site period', 'No'),
 ('1.13', 21, SITE, "Night security pending confirmation of the Employer's cover; Contractor", 'Monthly, site period; provisional', 'No'),
 ('1.14', 22, 'QCD18TSEMOB1240 and QCD18TSEMOB1250 (22 to 26-Aug-2026)', 'Delivery of cabins, containers and small plant; Contractor', 'Quantity: 5 flatbed trips, assessed', "No - the supplier's own mobilisation is within Item 6; plant mobilisation is within the day rates at 5.1 and 5.2"),
 ('1.15', 23, 'QCD18TSEDMOB1020 (17 to 24-Dec-2026 programmed)', 'Removal of the same; Contractor', 'Quantity: 5 flatbed trips, assessed', 'No'),
 ('1.16', 24, 'QCD18TSEDMOB1020', 'Site clean-up at demobilisation; Contractor labour', 'Attendance: 4 No. x 3 days', 'No - the approved histogram carries no people after 11-Dec-2026'),
 ('1.17', 25, 'Lump-sum deliverable - insurance endorsement, before mobilisation', "Extension of the Contract insurances to a site outside the Site; Contractor's insurer", 'Fixed: 0.2 per cent of the works value, to be substantiated', 'No'),
 ('1.18', 26, 'Site service - peak Contractor headcount (matrix M3, M5, M9, R1 to R5)', "PPE for the Contractor's own people; Contractor", 'Quantity: 20 sets', "No - the supplier's crews wear their own"),
 ('1.19', 27, 'Site service - HSE plan (RFP Scope of Works section 4, item 7)', 'Safety signage and notice boards; Contractor', 'Quantity: 12 No.', 'No'),
 ('1.20', 28, 'Site service - access route to the tank site', 'Work-zone traffic signs; Contractor', 'Quantity: 12 No. at the RFP-021 agreed rate', 'No'),
 ('1.21', 29, 'Site service - work-area protection', 'Temporary Jersey barriers; Contractor', 'Hire: 100 Lm, to be confirmed against the layout', 'No'),
 ('1.22', 30, 'After QCD18TSEDMOB1020 - completion inspection', "Engineer's completion inspection, correction list and warranties; Contractor site engineer, off-site visits", 'Attendance: 3 days', 'No - the monthly period at 1.1 ends at demobilisation'),
 ('1.23', 31, 'QCD18TSEDMOB3020 - close-out documents', 'Collation of the close-out file; Contractor document controller', 'Attendance: 4 days', 'No - test records and as-builts are 7.20 and 7.21'),
 ('2.1', 37, 'QCD18TSEENG1240, QCD18TSEENG1250 and the submittal activities QCD18TSEMOB1040 to 1230', "Coordination of the supplier's design, method statements, ITPs, shop-drawing resubmissions; Contractor", 'Attendance: 2 months, shared with the main Contract', 'No - tank structural design is the supplier\'s within Item 6'),
 ('2.2', 38, 'QCD18TSEENG1240 and the pipework shop drawings (Code C, resubmission 28 to 30-Sep-2026)', 'Shop drawings and submittal drawings; Contractor', 'Attendance: 1.5 months', 'No - as-built mark-ups are 7.21'),
 ('2.3', 39, 'QCD18TSECONSL1050 (07 to 13-Sep-2026) and after QCD18TSECONT1INS1060 / QCD18TSECONT2INS1060', 'Foundation level verification, then dimensional and verticality survey of each erected tank; Contractor surveyor', 'Attendance: 3 + 1.5 + 1.5 crew-days', 'No'),
 ('2.4', 40, 'QCD18TSEMOB1040 to QCD18TSEMOB1230 (19 submittal activities) and the MAR / MIR activities', 'Submittal administration via Aconex; Contractor', 'Attendance: 8 days', 'No'),
 ('2.5', 41, 'QCD18TSEMOB1040 to QCD18TSEMOB1230 (20 submittal and approval activities)', "Method statements, risk assessments and project plans; Contractor's engineer", 'Quantity: 6 engineer-days assessed', 'No - submittal administration is 2.4; the dismantling method statement is in Item 3'),
 ('3.1', 47, 'QCD18TSECONDSM1020 to 1030 (27-Aug to 12-Sep-2026, complete)', 'Dismantling, segregation and loading of the damaged tank; Al Mousa under quotation S04647 for the Contractor', 'Fixed: lowest of three quotations, recorded scope', 'No - haulage is 4.1; crane and crew within the quotation on its recorded scope (offer not attached), so no plant at Item 5; matrix M2 carries Contractor attendance only'),
 ('c.1 to c.14', 50, 'First-principles check of 3.1 - information only', 'Not carried', 'Not carried', 'Not carried'),
 ('4.1', 69, 'After QCD18TSECONDSM1030 - handover to the Employer', "Haulage of the dismantled materials to the Employer's local handover point; Contractor", 'Quantity: 14 loads, assessed; location unconfirmed', 'No - loading is within 3.1'),
 ('5.1', 75, 'Section 4 plant rows B1 to B5 (QCD18TSECONT1INS1030 to QCD18TSECONT2INS1050)', 'Boom truck for offloading and lifting wall tiers, bracing, roof supports and roof panels; Contractor for the supplier', 'Hire: working days on site, one unit', "No - the supplier's own lifting is excluded from its offer"),
 ('5.2', 76, 'Section 4 plant rows T1 to T3 (QCD18TSEPRC1120 to QCD18TSECONT2INS1050)', 'Telehandler moving pallets from the trucks to storage and into the tank footprint; Contractor for the supplier', 'Hire: working days on site, one unit, provisional', 'No'),
 ('5.3', 77, 'QCD18TSECONT1INS1030 to QCD18TSECONT2INS1050 (window D6)', 'Perimeter scaffold and edge protection for wall, bracing and roof work; Contractor (supplier condition)', 'Quantity: 1,024 m2 supplied, erected, 3-month hire, dismantled', 'No'),
 ('5.4', 78, 'Window D6', 'Mobile access towers; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.5', 79, 'Window D6', 'Relocation and re-inspection of the towers; Contractor', 'Quantity: 4 No.', 'No'),
 ('5.6', 80, 'Window D6', 'Podium steps; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.7', 81, 'QCD18TSEMOB1240 to QCD18TSECONTCT22030 (window D2)', "Daytime works power for the supplier's tools and the Contractor's works; Contractor (supplier condition)", 'Hire with fuel: monthly, works period', 'No - welfare power is 5.8'),
 ('5.8', 82, SITE, 'Continuous welfare power; Contractor', 'Hire with fuel: monthly, site period', 'No'),
 ('5.9', 83, SITE, 'Site pick-up; Contractor', 'Hire: monthly, site period', 'No - workforce bus is 1.12; plant is 5.1 and 5.2'),
 ('5.10', 84, 'Section 4 weekly bridge W1 to W16 (approved histogram)', "Contractor's helpers for every approved week: mobilisation, dismantling attendance, offloading, panel handling, transfer and disinfection", 'Attendance: approved helper man-days converted to man-months', 'No - erectors within Item 6; 7.4 nil; 7.8 Tank 2 only; clean-up 1.16'),
 ('5.11', 85, 'QCD18TSECONT1INS1020 to QCD18TSECONT2MW2020 (window D3)', "Power tools for the Contractor's own works; Contractor", 'Hire: monthly, erection window', "No - the supplier's erection tools are within Item 6"),
 ('5.12', 86, 'Within Item 6 on the described scope (MNT-AY-486 item 6; Assumption 9)', 'Sealant supplied and applied by the supplier', 'Nil', 'Yes - in Item 6'),
 ('5.13', 87, 'QCD18TSEPRC1120 to QCD18TSECONT2INS1050 (storage window D7)', 'Storage and handling consumables; Contractor - the supplier excludes storage and shifting (Assumption 9)', 'Quantity: 2 tanks', 'No - fixings are in Item 6; containers 1.11; plant 5.1 and 5.2'),
 ('5.14', 88, 'Window D3', 'Lighting towers for the erection fronts and the work area at dusk; Contractor', 'Hire: 2 No., monthly, erection window', 'No'),
 ('5.15', 89, 'Site service - power distribution from 5.7 and 5.8', 'Distribution boards and cabling; Contractor', 'Quantity: 2 sets', 'No'),
 ('6.1', 95, 'QCD18TSEPRC1060 to QCD18TSEPRC1230, QCD18TSECONT1INS1020 to QCD18TSECONT2MW2030', 'Design, manufacture, delivery duty paid, erection, sealing, bracing, nozzles and internals of both tanks; Al Mousa / Stalwart', 'Fixed: quotation per tank, insulated, provisional; supply-and-install scope as described (Assumption 9)', 'No - the Contractor provides helpers, unloading, storage, plant, scaffold and power (Items 1 and 5)'),
 ('7.1', 101, 'P18 (17 to 24-Nov-2026, assessment allowance)', 'Tankered water for the first fill; Contractor', 'Quantity: 3,774 m3, one fill', 'No'),
 ('7.2', 102, 'P19 and P20', 'Transfer pump; Contractor', 'Hire: transfer plus test days', 'No'),
 ('7.3', 103, 'P19', 'Transfer hoses and fittings; Contractor', 'Hire: 1 week', 'No'),
 ('7.4', 104, 'P19 - within the approved histogram week W16', 'Pump attendance during the transfer; Contractor helpers', 'Nil - priced at 5.10', 'Yes - 5.10'),
 ('7.5', 105, 'P18 to P20', 'Top-up for losses and test level; Contractor', 'Quantity: 10 per cent of one fill', 'No'),
 ('7.6', 106, 'P18 and P20 (AWWA C652)', 'Disinfection chemicals; Contractor', 'Quantity: 800 kg', 'No'),
 ('7.7', 107, 'P18 and P20', 'Dosing equipment; Contractor', 'Hire: 2 weeks', 'No'),
 ('7.8', 108, 'Dated check H8 and H9 (12-Dec; 16 to 17-Dec, after the histogram)', 'Tank 2 inspection, sampling, dechlorination and discharge; Contractor helpers', 'Attendance: man-days after 11-Dec from the dated check', 'The Tank 1 dosing (H6) is within 5.10'),
 ('7.9', 109, 'After P20', 'Dechlorination for discharge; Contractor', 'Quantity: 400 kg', 'No'),
 ('7.10', 110, 'P18 and P20', 'Sampling and transport; Contractor', 'Quantity: 2 tanks', 'No'),
 ('7.11', 111, 'P18 and P20', 'Laboratory water-quality tests; third-party laboratory', 'Quantity: 6 samples', "No - excluded by both tank suppliers"),
 ('7.12', 112, 'None - no third-party inspection in the RFP', 'Not assessed', 'Not assessed', 'Engineer witness only (RFP 5.1, 5.2)'),
 ('7.13', 113, 'P18, P20, P21, P21b', 'Commissioning engineer at the test holds, the component checks and the integrated commissioning; Contractor', 'Attendance: 8 days', "No - the supplier's leak-test supervision is within Item 6"),
 ('7.14', 114, 'P21 and P21b', 'MEP technicians on the pump and network interfaces; Contractor', 'Attendance: 2 No. x 6 days', 'No'),
 ('7.15', 115, 'P21 (after QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040)', 'Electrical connection and readout checks of the level instruments; Contractor', 'Attendance: 5 days', 'No - the cabling is supplied and installed at 8.11'),
 ('7.16', 116, 'P16 and P21b', 'Coordination of the tie-in and the witnessed demonstration with the Employer and Engineer; Contractor', 'Attendance: tie-in days plus 1', 'No'),
 ('7.17', 117, 'P21', 'Calibration of the instruments; technician', 'Attendance: 4 days', 'No'),
 ('7.18', 118, 'P21', 'Certified calibrations; laboratory', 'Quantity: 6 No.', 'No'),
 ('7.19', 119, 'P18 to P21b', 'Calibrated gauges and data logger; Contractor', 'Hire: 2 weeks', 'No'),
 ('7.20', 120, 'P18 to P23', 'Test records and ITP / WIR close-out; Contractor clerk', 'Attendance: 1 month', 'No - the document controller at 1.23 collates the close-out file'),
 ('7.21', 121, 'QCD18TSEDMOB3020', 'As-built mark-ups; Contractor draughtsman', 'Attendance: 0.5 month', 'No - shop drawings are 2.2'),
 ('7.22', 122, 'P18 and P20 holds', 'Pump and hose set standing by through the holds; Contractor', 'Hire: 4 days', 'No - tankers are paid at 7.1'),
 ('7.23', 123, 'Deliverables list 26 to 28-Sep-2026 (PQD, ITP, procedure, inspector CV)', 'Third-party factory acceptance test - not required by the RFP', 'Nil pending evidence', 'Manufacturer test reports are within Item 6'),
 ('7.24', 124, 'P20b (09 to 10-Dec-2026), after P19 and P16', 'Tankered operating water in Tank 1 for the pumping demonstration; Contractor', 'Quantity: 1,020 m3 (1 m depth), Section 5', 'No - the first fill is 7.1; attendance is within 5.10 (duty H11)'),
 ('7.25', 125, 'Tank 2 on 03 and 05-Dec (before P19), Tank 1 on 08-Dec (after P19)', 'Spray disinfection of the surfaces above the test water line; specialist crew within the rate', 'Quantity: 2 tanks', 'No - the chlorinated water at 7.6 and 7.7 covers only the wetted surfaces'),
 ('8.1', 132, 'QCD18TSEPRC1160 to 1250 (procurement), QCD18TSECONT1MW2055 and QCD18TSECONT2MW2020 (installation)', 'Main pipework above DN150 to the RFP specification, supplied and installed; Contractor', 'Quantity: 200 m assessed, take-off pending', 'No'),
 ('8.2', 133, 'As 8.1', 'Small-bore uPVC pipework; Contractor', 'Quantity: 60 m assessed', 'No'),
 ('8.3', 134, 'As 8.1', 'Butterfly isolation valves to the RFP; Contractor', 'Quantity: 10 No. assessed', 'No'),
 ('8.4', 135, 'As 8.1', 'Dismantling joints; Contractor', 'Quantity: 4 No.', 'No'),
 ('8.5', 136, 'QCD18TSECONT1MW2050 and QCD18TSECONT2MW2030', 'Blind flanges on spare nozzles; Contractor', 'Quantity: 8 No.', "No - the nozzles themselves are the supplier's"),
 ('8.6', 137, 'As 8.1', 'Flange sets, gaskets and bolting; Contractor', 'Quantity: 2 tank-sets', 'No'),
 ('8.7', 138, 'As 8.1', 'Pipe supports; Contractor', 'Quantity: 40 No.', 'No'),
 ('8.8', 139, 'As 8.1', 'Anchor and thrust blocks; Contractor', 'Quantity: 8 No.', 'No'),
 ('8.9', 140, 'QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040', 'Float-and-tape level indicators; Contractor', 'Quantity: 2 No.', "No - the supplier includes only a tube-type indicator (MNT-AY-486 item 5); the RFP float-and-tape unit with local readout is priced here"),
 ('8.10', 141, 'As 8.9', 'Level transmitters; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.11', 142, 'As 8.9', 'Instrument cabling and conduit; Contractor', 'Quantity: 160 m assessed', 'No - connection checks are 7.15'),
 ('8.12', 143, 'QCD18TSECONTC2040', 'Pressure gauge assemblies; Contractor', 'Quantity: 2 sets', 'No'),
 ('8.13', 144, 'QCD18TSECONTC2040', 'Sample taps; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.14', 145, 'QCD18TSECONTC2040 after QCD18TSECONIF2050 (03-Dec-2026)', 'Tie-ins to the networks after the external readiness milestone; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.15', 146, 'QCD18TSECONTC2040', 'Tie-in coordination and out-of-hours working; Contractor', 'Quantity: 2 No.', 'No - the coordination engineer at 7.16 is the Contractor\'s attendance at the demonstration'),
 ('8.16', 147, 'After QCD18TSECONT2MW2020, before P21', 'Pipework test pump and manifold; Contractor', 'Hire: 2 weeks', 'No - tank testing is Item 7'),
 ('8.17', 148, 'As 8.16', 'Pipework test and flushing water; Contractor', 'Quantity: 300 m3', 'No'),
 ('8.18', 149, 'As 8.16', 'Pipework testing crew; Contractor', 'Attendance: 2 No. x 12 days', 'No - not in 5.10 (matrix M14)'),
 ('8.19', 150, 'As 8.16', 'Test records and certificates; Contractor', 'Quantity: 2 systems', 'No'),
 ('8.20', 151, 'As 8.1', 'Colour banding; Contractor', 'Quantity: 260 m', 'No'),
 ('8.21', 152, 'As 8.1', 'Tags, arrows and labels; Contractor', 'Quantity: 70 No.', 'No'),
 ('a.1 to a.5', 155, "The Contractor's procured materials - information only", 'Not carried', 'Not carried', 'Not carried'),
]
r += 1
pg_break(pg, r)
pg_add(21 + 60 + 32)
banner(pg, r, "6. LINK REGISTER - EVERY 'BUILD-UP' LINE TO ITS PROGRAMME ACTIVITIES OR SITE SERVICE"); r += 1
para(pg, r, ("Link register\n"
    "• One row per 'Build-Up' line: the activities or site service it serves; what it buys, when and who provides it; the basis; whether paid elsewhere; the Assessed amount linked from the 'Build-Up' tab.\n"
    "• Whole-site lines are shown as site services. A resource serving several activities is charged once, where it is priced."), height=48); r += 1
HDRG = ['Ref', "'Build-Up' line (as described there)", 'Programme activities or site service', '', 'What it buys, when it can happen, who provides it', '', 'Basis: fixed, quantity, attendance or hire', '', 'Paid elsewhere?', 'Assessed (SAR)']
def hdrg(ws, r_):
    header(ws, r_, HDRG[:9]); ws.cell(r_, 10, HDRG[9]); cp(S_HDR, ws.cell(r_, 10))
    ws.merge_cells(start_row=r_, start_column=3, end_row=r_, end_column=4); ws.merge_cells(start_row=r_, start_column=5, end_row=r_, end_column=6); ws.merge_cells(start_row=r_, start_column=7, end_row=r_, end_column=8)
hdrg(pg, r); r += 1
g_first = r
for ref, burow, wbs, what, basis, elsewhere in REG:
    h = max(30, est(wbs, 24), est(what, 23), est(basis, 34), est(elsewhere, 14), est(str(bu.cell(burow, 2).value), 46))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); hdrg(pg, r); r += 1; pg_add(PG['hdr'])
    pg_add(h)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, f"='Build-Up'!B{burow}" if '.' in ref and ' to ' not in ref else str(bu.cell(burow, 2).value)); cp(S_DESC, pg.cell(r, 2))
    for col, txt in ((3, wbs), (5, what), (7, basis), (9, elsewhere)):
        pg.cell(r, col, txt); cp(S_BASIS, pg.cell(r, col)); cp(S_BASIS, pg.cell(r, col + 1) if col < 9 else pg.cell(r, col))
    for c1 in (3, 5, 7):
        pg.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c1 + 1)
    amt = f"='Build-Up'!F{burow}" if ('.' in ref and ' to ' not in ref) else '-'
    pg.cell(r, 10, amt); cp(S_AMT, pg.cell(r, 10))
    pg.row_dimensions[r].height = h
    r += 1
g_last = r - 1
if PG['used'] + 55 > PAGE:
    pg_break(pg, r); hdrg(pg, r); r += 1
pg.cell(r, 2, "Total of the priced lines above, before Overhead and Profit"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 10, f"=SUM(J{g_first}:J{g_last})+'Build-Up'!F95"); cp(S_TOTAMT, pg.cell(r, 10))
pg.cell(r, 9, "Equals the 'Assessment' tab subtotal, row 19 (Item 6 counted for the second tank)"); cp(S_BASIS, pg.cell(r, 9))
pg.row_dimensions[r].height = 30
GTOT = r; r += 1
pg.cell(r, 2, "Check: difference to 'Assessment'!K19 (must be nil)"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 10, f"=ROUND(J{GTOT}-Assessment!$K$19,2)"); cp(S_TOTAMT, pg.cell(r, 10))
pg.row_dimensions[r].height = 19.5
r += 1
pg.print_title_rows = '1:3'
pg.freeze_panes = 'A4'
# merge the section-3 header 'Value' cell row 12 handled; ensure row 4/7/8/9/18 paragraphs merge to I only (J unused there)

# ---------------------------------------------------------------- Build-Up text updates
bu['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'
bu['A4'] = ("How to read this tab: each item is built up as quantity x rate with its basis beside it; the item total goes to column J of the 'Assessment' "
            "tab. A quantity that is a formula comes from the 'Programme' tab, where the cell it points to names the programme activities it was built from - "
            "click it to trace. 'Assumption' in the basis means an assessed market rate or quantity the Contractor has not yet substantiated. All amounts exclude VAT.")
bu.row_dimensions[4].height = 55
bu['A7'] = ("The tank site is outside the D-18 site boundary, so site management, welfare, water, security, insurance and close-out are priced as a separate "
            "site. How long: the Contractor's programme has the site occupied from mobilisation on 22-Aug-2026 to the end of demobilisation, about 4 months "
            "('Programme' tab, period D1); the end is fixed by the pump-room readiness date of 03-Dec-2026 and the tie-in, testing and demobilisation after it. "
            "Each staff role is priced only for the phases in which it is needed ('Programme' tab, roles R1 to R5). Design staff are shared with the main Contract "
            "(Item 2). Power is in Item 5. The tank supplier's price includes its own crew's mobilisation; the foundations and steel base frames exist. The "
            "programme is under the Engineer's approval, not agreed.")
bu.row_dimensions[7].height = 96
SRC_A = "Programme' tab, Section 3"
def gset(row_, text): bu.cell(row_, 7).value = text
gset(9, "Dedicated for the site period, mobilisation to the end of demobilisation, with the phase duties set out on the 'Programme' tab, Section 4, role R1. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(10, "Dedicated for the site period with the phase duties at 'Programme' tab, Section 4, role R2. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(11, "From the start of dismantling to completion of testing and commissioning only ('Programme' tab, Section 4, role R3); not during mobilisation or demobilisation. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(12, "Dedicated for the site period: the Contractor's own labour is on site from mobilisation to demobilisation ('Programme' tab, Section 4, role R4). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(13, "From the first delivery to completion of testing only ('Programme' tab, Section 4, role R5). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
for row_ in (14, 15, 16, 18):
    gset(row_, f"Site period D1 on the '{SRC_A} (activity QCD18TSEMOB1250, 22-Aug-2026, to activity QCD18TSEDMOB1020, 24-Dec-2026). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(17, f"Two 10 m3 deliveries a week for the weeks of the site period D1 on the '{SRC_A}; no mains water at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B19'] = 'Storage containers, 2 No. for the panel storage window (RFP Scope of Works, work package 1 - pre-construction and setup: safe storage of panels)'
gset(19, f"2 No. containers for the storage window D7 on the '{SRC_A}, first delivery to the Tank 2 roof finish. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(20, f"Separate site outside the D-18 boundary; site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(21, f"Site period D1 on the '{SRC_A}. Included pending confirmation whether the Employer's security covers the lower plateau; delete if it does. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(22, "5 trips retained: the programme gives no trip count (activity QCD18TSEMOB1240, mobilisation, 5 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(23, "5 trips retained: the programme gives no trip count (activity QCD18TSEDMOB1020, demobilisation, 7 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D30'] = 3
gset(30, "3 days after demobilisation: the Engineer's completion inspection, the correction list and the warranties collation. Reduced from 8 days because the site engineer's monthly period at 1.1 now runs to the end of demobilisation and the programmed close-out activity (QCD18TSEDMOB3020, 17 to 24-Dec-2026) falls inside it. Test records, ITP/WIR close-out and as-built drawings are in Item 7 (7.20 and 7.21), not here. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B30'] = "Close-out at the tank site after demobilisation - completion inspection, correction list and warranties collation - site engineer"
gset(37, "2 months retained: the programme shows shop drawings 07 to 25-Aug-2026 (activities QCD18TSEENG1240 and 1250) with resubmissions to 29-Sep-2026 and the pipework shop drawings still in preparation, consistent with the period assessed. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(38, "1.5 months retained on the same programme evidence as 2.1. As-built drawings at close-out are line 7.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['A73'] = ("What Item 5 pays for: what the tank supplier excludes and the Contractor must provide - helpers and labourers, unloading and shifting of the "
             "materials, storage and handling, forklift, crane and scaffolding, and power for the installation (supplier scope as described, Assumption 9 below) - "
             "plus general site plant. Erection itself, with the sealant and the fixings, is in the supplier's supply-and-install price (Item 6). Helpers: the approved manpower histogram's helper row, priced once ('Programme' "
             "tab, Section 4). Plant days: counted from the programme activity dates on the same tab; one boom truck and one telehandler is an assumption pending "
             "the Contractor's plant schedule. Power: the works generator by day until commissioning ends, the welfare generator around the clock until "
             "demobilisation ends - no site power is available (Engineer's email of 30-Aug-2026, which also notes the Contractor's method uses forklifts, cranes "
             "and pallet jacks).")
bu.row_dimensions[73].height = 96
gset(75, "Hire days B5 on the 'Programme' tab, Section 4, built from the XER activity dates: alternate days through each tank's walls-and-bracing window, every roof-support and roof-panel day, less the overlap of the two tanks, plus offloading days outside those. The Contractor's methodology names forklifts, cranes and pallet jacks (Engineer's (KEO) email of 30-Aug-2026); no plant schedule has been submitted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(76, "Hire days T3 on the 'Programme' tab, Section 4, built from the XER activity dates: continuous from the first delivery to the Tank 2 walls finish, plus the Tank 2 roof days; one unit is a provisional utilisation assumption. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(77, f"2 tanks x 128 m perimeter x 4.0 m height; the rate includes a 3-month hire, which covers the access window D6 on the '{SRC_A} on either basis. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B78'] = 'Mobile aluminium access towers, 2 No. for the access window, whole hire months'
gset(78, f"2 No. for the access window D6 on the '{SRC_A}, rounded up to whole hire months. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B80'] = 'Podium steps, 2 No. for the access window, whole hire months'
gset(80, "As 5.4. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(81, f"Works period D2 on the '{SRC_A}, to completion of commissioning. Power is excluded by both tank suppliers; no Site power expected at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(82, f"Site period D1 on the '{SRC_A}; air conditioning and lighting run around the clock at a separate site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(83, f"Site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B84'] = "Contractor's helpers - the approved manpower histogram helper row (85 man-weeks), man-months"
gset(84, "The approved manpower histogram's helper row, 85 man-weeks, as reproduced week by week on the 'Programme' tab, Section 4, converted to man-days on each week's working days and to man-months at 26 working days; every helper man-day priced once here, so the transfer labour 7.4 is nil and 7.8 carries only the Tank 2 disinfection after the histogram ends. The dismantling-week helpers are priced provisionally (S7). Skilled people are within Items 3, 6 and 8 on their recorded or described scopes ('Build-Up' Assumption 9) and are not priced. The Contractor's histogram helper row (85 man-weeks) is reconciled week by week there; it is not adopted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B85'] = "Power and hand tools, slings and lifting tackle for the Contractor's helpers - unloading, shifting and attendance"
gset(85, f"The supplier installs with its own erection tools (Assumption 9); the Contractor's helpers need their own tools and tackle for the unloading and shifting the supplier excludes. Erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B88'] = 'Mobile lighting towers, 2 No. for the erection window'
gset(88, f"2 No. for the erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['A99'] = ("What Item 7 pays for, in three parts. (1) Hydrostatic test water: one fill of 3,774 m3 tankered into Tank 1 (7.1), tested, then pumped across to Tank 2 "
             "(7.2, 7.3) and topped up (7.5) - the Engineer's email of 30-Aug-2026 allows the water to be re-used; the Contractor's programme fills both tanks at once, "
             "shown for comparison on the 'Programme' tab, Section 5, and not included. (2) Disinfection: chlorinating the test water (7.6, 7.7) disinfects the "
             "wetted surfaces only; the surfaces above the 3.7 m test level are covered by a separate spray application (7.25), and the water is dechlorinated before "
             "any discharge (7.9); samples and laboratory tests at 7.10 and 7.11. (3) Functional commissioning: component checks after the tie-in and the integrated "
             "demonstration with both tanks connected and holding water (7.13 to 7.19); Tank 2 holds the test water and Tank 1 is given an operating fill of about "
             "1,020 m3 by tanker (7.24), because a network refill after the tie-in has not been confirmed as available or free and a demonstration of Tank 1 on its "
             "valves and instruments alone is not established as meeting the integrated commissioning scope (RFP Scope of Works 5.2). Dates: 'Programme' tab, "
             "windows P18 to P22 - Tank 1 test 17 to 24-Nov, transfer 05 to 07-Dec, Tank 2 test 08 to 12-Dec, Tank 1 operating fill 09 to 10-Dec, integrated "
             "commissioning 13 to 15-Dec-2026. The whole method is provisional until the Contractor's method statement is accepted by the Engineer: the water "
             "source, the disinfection method, the depth required in Tank 1 and the release of the water (kept as first stock or discharged) are all unconfirmed. "
             "Not included: the supplier's own leak-test supervision (Item 6), re-testing after a failed test (the Contractor's obligation under RFP Scope of "
             "Works 5.1) and pipework testing (Item 8).")
bu.row_dimensions[99].height = 150
gset(101, "Quantity: 3,774 m3, the volume at the 3.7 m test level (RFP Scope of Works section 1), one fill only - the same water is pumped to Tank 2 (Engineer's email of 30-Aug-2026); water for testing is excluded by the supplier (MNT-AY-486 exclusion 5, Assumption 9). Rate: tankered water at SAR 6.00/m3, assessment allowance - needs confirmation of the source; a network source confirmed by the Employer would replace it")
gset(102, "The transfer P19 and the Tank 2 test P20 on the 'Programme' tab (7 days); pump duty an assessed assumption pending the Contractor's method statement. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(104, "2 No. x 3 days, the transfer window P19 on the 'Programme' tab. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(105, "Quantity: 10 per cent of the first fill, for losses while the water waits in Tank 1 and for topping Tank 2 up to its test level - assessment allowance. Rate: as 7.1")
gset(113, "Component checks and integrated commissioning (P21 and P21b on the 'Programme' tab) plus one day at each hydrostatic test hold; reduced from 15 days because the fills and holds are supervised by the QA/QC inspector (1.3) and the supplier (Item 6). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(114, "2 No. for the component checks and the integrated commissioning (P21 and P21b on the 'Programme' tab); reduced from 2 No. x 10 days. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(116, "Tie-in connections (P16 on the 'Programme' tab, 4 working days) plus one day of integrated commissioning. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(120, "1 month retained: test records run from the first hydrostatic test to the close-out (P18 to P23 on the 'Programme' tab, about five weeks on either basis). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(122, "Pump and hose set standing by through the two 24-hour holds and two days of contingency between the Tank 1 test and the transfer; tankers work during the fill and are paid at 7.1, not here. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
# item-level corrections from the resource and scope review
bu['D26'] = 20
gset(26, "Peak Contractor headcount from the resource matrix ('Programme' tab, Section 4): two helper gangs and the offloading gang (up to 14), five site staff and the pipework crew - about 20 sets; the supplier's crews wear their own. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D39'] = 6
gset(39, "Survey and levelling of the existing foundation (activity QCD18TSECONSL1050, 07 to 13-Sep-2026, cost-loaded by the Contractor at SAR 5,000.00) assessed at 3 crew-days for an existing base frame, plus a dimensional and verticality survey of each erected tank, 1.5 crew-days each (RFP Scope of Works, work package 3). Not setting out: the steel base frames are already installed on concrete supports (site photograph, Aug-2026). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D40'] = 8
gset(40, "The programme carries 19 deliverable and submittal activities plus the material approval and inspection requests; 8 days of document control, doubled from the previous revision on that evidence. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D86'] = 0
gset(86, "Included in the supplier's price (Assumption 9): the leak sealant is supplied by the tank supplier (quotation MNT-AY-486, specification item 6) and its application at every joint is the supplier's installation work (RFP Scope of Works work package 2). No Contractor sealant tools or consumables are needed, so nil. Previously 2 tanks at SAR 1,500.00")
bu['B87'] = "Panel storage and handling consumables - timber packers, covers, strapping and cleaning materials for the Contractor's storage and shifting of the tank materials"
bu['D87'] = 2
bu['E87'] = 1500
gset(87, "Quantity: 2 tanks. The bolts, nuts, washers and tie rods are supplied and fixed by the tank supplier (MNT-AY-486, items 3 and 4; Assumption 9) and are not priced here. Storing the materials until the site is ready and moving them to the installation area are the client's, so the Contractor's (MNT-AY-486 note, page 2): the consumables for that handling are a fair Contractor cost, priced once here; the containers are 1.11 and the plant 5.1 and 5.2. Rate: assessed allowance - needs confirmation by the Contractor. Previously 'fixings, touch-up and miscellaneous consumables', 2 tanks at SAR 2,250.00")
bu['D112'] = 0
gset(112, "Not assessed: RFP Scope of Works 5.1 and 5.2 require the tests to be witnessed by the Engineer, not inspected by a third party, and the third-party factory acceptance test is dealt with at 7.23. Previously 2 visits at SAR 2,400.00")
bu['D104'] = 0
gset(104, "Nil: the transfer labour is within the approved manpower histogram, whose helper man-days are all priced at 5.10 ('Programme' tab, Section 4, week W16); priced once. Previously 2 No. x 3 days")
bu['D108'] = f"=Programme!{AFTER_CELL}"
gset(108, "Helper man-days after the approved histogram ends on 11-Dec-2026 ('Programme' tab, Section 4, dated check): Tank 2 sampling and inspection on 12-Dec and dechlorination in the tank on 13-Dec. The dosing of the test water in Tank 1 (01 to 03-Dec) falls within the approved weeks priced at 5.10; the water is kept in Tank 2 as first stock, so there is no discharge. Previously 2 No. x 10 days")
bu['B108'] = 'Sampling and dechlorination labour after the approved histogram weeks'
# Item 8
bu['A130'] = ("External pipework, valves, fittings and instruments from the tanks to the tie-in points (RFP Scope of Works, piping requirements); not in the tank "
              "supplier's price. Priced to the RFP specification: uPVC below DN150, GRP or ductile iron above, butterfly isolation valves. The Contractor is "
              "buying different materials (HDPE pipe, a gate valve, a motorised valve); these are not approved and are shown after line 8.21 for information "
              "only. No quantities exist yet - the pipework shop drawings were rejected and no take-off has been received - so the quantities are assessed "
              "from the RFP sketch and are provisional. Rates include installation. Lines 8.16 to 8.19 test the pipework only; tank testing is Item 7.")
bu.row_dimensions[130].height = 96
bu['B132'] = 'Main pipework above DN150 - GRP or ductile iron with internal lining, with fittings, supplied and installed (RFP General Piping Requirements)'
gset(132, "Assessed run lengths for inlet, outlet and overflow of 2 tanks - take-off required. The Contractor is procuring HDPE pipe of 355 mm and 315 mm outside diameter (nominal size subject to the pipe standard and SDR and to the Engineer's confirmation): see the alternative after line 8.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B133'] = 'Small-bore pipework below DN150 - uPVC Schedule 40, supplied and installed'
gset(133, "Assessed - take-off required. The Contractor is procuring uPVC pipe of 160 mm and 110 mm outside diameter (nominal size subject to the pipe standard and to the Engineer's confirmation); whether the 160 mm pipe falls below or at the DN150 boundary of the specification depends on the pipe standard and is to be confirmed by the Engineer. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B134'] = 'Resilient-seated butterfly isolation valves DN150-DN300, lever-operated up to DN200 and gearbox-operated above, installed'
gset(134, "Assessed count - take-off required; the procurement tracker lists 2 valves (a gate valve DN300 and a motorised butterfly valve DN355), neither to the specified type or operation - see the alternative after line 8.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(141, "Excluded from the tank supplier scope; the programme shows level transmitters installed on both tanks (activities QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(145, "Tie-in connections with the existing pump room after the readiness milestone of 03-Dec-2026 (activity QCD18TSECONTC2040, 4 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
# Item 8 alternative block: insert 8 rows at 151 (before ASSUMPTIONS at 152 after the earlier shift)
ASSUMP = 155
insert_rows_keep_styles(bu, ASSUMP - 1, 9, 132)   # rows 153..161 new; assumptions now at 163
for rr in range(154, 163):
    for c in range(1, 8): bu.cell(rr, c).value = None
bu['B154'] = "Contractor's procured materials in place of lines 8.1 to 8.3 - shown for information, not carried"
cp(bu['B49'], bu['B154'])
for c in (1, 3, 4, 5, 6, 7): cp(bu.cell(49, c), bu.cell(154, c))
bu.row_dimensions[154].height = 18
alt = [
    ('a.1', 'HDPE PE100 pipe, 355 mm and 315 mm outside diameter (nominal size subject to the pipe standard and SDR and to the Engineer\'s confirmation), with HDPE fittings, flange adaptors, reducer, elbow and tee, supplied and installed - in place of line 8.1', 'm', "=D132", 280,
     "Same assessed length as 8.1 (no take-off). Union Pipe Industry supply per the procurement tracker of 29-Sep-2026; PQD approved Code B 16-Sep-2026, material approval request under preparation - not approved against the GRP or ductile iron specification. The tracker gives outside diameters, not nominal sizes; the nominal size, pressure class and SDR are to be confirmed on the approved shop drawings and by the Engineer. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.2', 'uPVC pipe, 160 mm and 110 mm outside diameter (nominal size subject to the pipe standard and to the Engineer\'s confirmation), with uPVC elbows, supplied and installed - in place of line 8.2', 'm', "=D133", "=E133",
     "Same assessed length and rate as 8.2: the material conforms below DN150. Al Muneef supply; PQD submitted 15-Sep-2026, material approval request under preparation"),
    ('a.3', 'Gate valve DN300, installed - in place of one valve at line 8.3', 'No', 1, 6500,
     "Saudi Pipe Systems supply. Does not conform: the RFP General Piping Requirements call for resilient-seated butterfly isolation valves. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.4', 'Motorised butterfly valve DN355 as listed by the Contractor, installed, actuator included, power and control supply excluded - in place of one valve at line 8.3', 'No', 1, 14000,
     "Saudi Pipe Systems supply. Does not conform: the RFP requires gearbox operation above DN200; a motorised valve needs a power and control supply that is in no party's scope. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.5', 'Remaining valves at line 8.3 to the RFP specification - 8 No. retained', 'No', "=D134-2", "=E134",
     "The tracker lists only 2 valves; the count at 8.3 stays assessed until the pipework shop drawings are approved and measured"),
]
rr = 155
for ref, desc, unit, q, rate, basis in alt:
    bu.cell(rr, 1, ref); bu.cell(rr, 2, desc); bu.cell(rr, 3, unit); bu.cell(rr, 4, q); bu.cell(rr, 5, rate); bu.cell(rr, 6, f"=D{rr}*E{rr}"); bu.cell(rr, 7, basis)
    bu.row_dimensions[rr].height = max(39, 13 * (len(basis) // 70 + 1))
    rr += 1
bu.cell(rr, 2, 'Alternative total for lines 8.1 to 8.3 on the procured materials - information only'); cp(bu['B64'], bu.cell(rr, 2))
bu.cell(rr, 6, f"=SUM(F155:F{rr - 1})"); cp(bu['F64'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(64, c), bu.cell(rr, c))
rr += 1
bu.cell(rr, 2, 'Assessed amount of lines 8.1 to 8.3 to the RFP specification, included in the assessment'); cp(bu['B64'], bu.cell(rr, 2))
bu.cell(rr, 6, "=F132+F133+F134"); cp(bu['F64'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(64, c), bu.cell(rr, c))
rr += 1
bu.cell(rr, 2, "Difference - not carried: the substitution is not approved"); cp(bu['B64'], bu.cell(rr, 2))
bu.cell(rr, 6, f"=F{rr - 1}-F{rr - 2}"); cp(bu['F64'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(64, c), bu.cell(rr, c))
bu.row_dimensions[rr].height = 30
assert rr == 162
# assumptions block now rows 161..169
A0 = 164
assert str(bu.cell(A0, 1).value).startswith('ASSUMPTIONS'), bu.cell(A0, 1).value
bu.cell(A0 + 2, 1).value = ("2. Rates marked 'assumption' are assessed Riyadh market rates (Sep-2026): crane 50 t SAR 2,500/day; boom truck SAR 1,200/day; telehandler SAR 900/day; "
    "scaffolding SAR 55/m2; generator 100 kVA SAR 11,000/month with fuel; general labour SAR 150/day; rigger SAR 250/day; site engineer SAR 12,000/month; HSE and QA/QC "
    "SAR 10,000/month; foreman SAR 8,000/month; test water SAR 6.00/m3 tankered; site water SAR 350 per 10 m3 delivery; generator 30 kVA SAR 6,000/month continuous with fuel; "
    "local haulage SAR 450/load; HDPE PE100 pipe 355 mm outside diameter supplied and installed SAR 280/m; gate valve DN300 SAR 6,500; motorised butterfly valve DN355 as listed SAR 14,000 excluding power and control.")
bu.row_dimensions[A0 + 2].height = 50
bu.cell(A0 + 3, 1).value = ("3. Programme: every period on this tab comes from the Contractor's programme of 29-Sep-2026, imported as received on the 'XER WBS' tab and "
    "worked through on the 'Programme' tab. It is under the Engineer's approval, not agreed. Site period 22-Aug-2026 to the end of demobilisation, fixed by the "
    "pump-room readiness date of 03-Dec-2026 and what follows it. The staged tank deliveries (7 to 9 weeks for the first tank, 12 to 14 for the second, delivered "
    "duty paid from the UAE) are the Contractor's risk: no standby or prolongation is priced. Dates after 28-Sep-2026 are forecasts.")
bu.row_dimensions[A0 + 3].height = 63
SC = ("Tank supplier scope - user-authorised assumption. The split of work between the tank supplier and the Contractor is taken as described in the Al Muhaideb "
      "National Tanks (National Factory for Fiberglass) quotation MNT-AY-486 dated 24-Jun-2026 with its tank and foundation drawings dated 06-May-2026, received "
      "30-Sep-2026. Included by the supplier (page 1, 'Supply, Installation and Testing'; specifications 1 to 13): manufacture, installation and testing, the "
      "non-toxic leak sealant (item 6), all internal bolts, nuts, washers and tie rods in stainless steel 316 and external fixings in hot-dip galvanised steel "
      "(items 3 and 4), level indicator tube, screened air vent, lockable manhole, internal and external ladders, roof panel supports, inlet, outlet, overflow "
      "and drain openings and the galvanised steel skid base (items 5 to 13). Excluded by the supplier and therefore the Contractor's (exclusions 1 to 6, pages 1 "
      "and 2): civil works and foundations; power for the installation; helpers and labourers; unloading and shifting of material at the installation location; "
      "all pipe connections, piping, float switches, fittings, flanges, valves, water for testing and any third-party testing or inspection; forklift, crane, "
      "elevator and scaffolding if required. The client stores the materials until the site is ready and transports them to the installation area by its own "
      "means (note, page 2). Mismatch stated plainly: this is not the adopted Al Mousa S04488 / Stalwart offer. It is a different supplier, a different tank "
      "(3,000 m3 gross, 30 (15+15) m x 25 m x 4 m, 3 No., non-insulated, with partition), addressed to WTB, dated June 2026 with 15 days' validity, priced at "
      "SAR 1,575,000 per tank net of the discount, VAT excluded (SAR 525 per gross m3). Its price is not adopted; its scope split is applied on instruction as the "
      "assessment assumption for Items 5, 6, 7 and 8 until the Al Mousa offer and its conditions are produced ('Build-Up Comparison' tab, Section 7, item 14).")
bu.cell(A0 + 9, 1).value = '9. ' + SC
cp(bu.cell(A0 + 8, 1), bu.cell(A0 + 9, 1))
for c in range(2, 8): cp(bu.cell(A0 + 8, c), bu.cell(A0 + 9, c))
bu.merge_cells(start_row=A0 + 9, start_column=1, end_row=A0 + 9, end_column=7)
bu.row_dimensions[A0 + 9].height = 150
bu.cell(A0 + 6, 1).value = ("6. The 5 per cent Overhead and Profit on the 'Assessment' tab is head-office overhead and profit only; every site cost is in the items. The "
    "Contractor's cost loading of its programme shows how it will apply for payment, not what things cost ('Build-Up Comparison' tab, Section 8).")
bu.row_dimensions[A0 + 6].height = 37.8

# ---------------------------------------------------------------- Assessment tab
asm = wb['Assessment']
asm['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'
asm['A4'] = ('="Purpose: to value the Contractor\'s revised proposal ref. SAMACO-RRFP-000001 dated 13-Aug-2026 for the two TSE storage tanks, the dismantling of the damaged '
             'tank, pipework and testing. Result: proposal SAR " & TEXT(F21,"#,##0.00") & ", assessed SAR " & TEXT(K21,"#,##0.00") & ", both excluding VAT (row 21). How: supplier '
             'quotations where the Contractor obtained them (Items 3 and 6); everything else built up from quantities and rates (\'Build-Up\' tab), with every time period '
             'taken from the Contractor\'s programme (\'XER WBS\' tab as received, \'Programme\' tab worked through). The programme is under the Engineer\'s approval, not '
             'agreed. Status: preliminary; items marked Provisional await confirmations. Rev 02 replaces Rev 01 dated 17-Sep-2026 (SAR " & TEXT(6921685.64,"#,##0.00") & ")."')
asm.row_dimensions[4].height = 57
asm['L8'] = ("Separate site outside the D-18 boundary: staff, welfare, water, security, insurance, mobilisation, demobilisation and close-out are priced for the "
             "site period in the Contractor's programme (22-Aug-2026 to the end of demobilisation, about 4 months), each staff role only for the phases it is "
             "needed. See 'Build-Up' Item 1 and 'Programme' tab Sections 3 and 4.")
asm['L14'] = ("Plant, access, power, storage handling and helpers that the tank supplier excludes (supply-and-install scope as described, 'Build-Up' Assumption 9); erection, sealant and fixings are in the tank price. Helpers: the "
              "approved manpower histogram, priced once. Plant days: from the programme activity dates. No plant schedule has been submitted. See 'Build-Up' Item 5 "
              "and 'Programme' tab Section 4.")
asm['L17'] = ("Hydrostatic test on one fill re-used for the second tank (Engineer's email of 30-Aug-2026), disinfection of the wetted surfaces by the test water "
              "and of the surfaces above it by spraying, and an operating fill of Tank 1 for the integrated demonstration - no network refill is assumed. The method is "
              "provisional until the Contractor's method statement is accepted by the Engineer. The Contractor's parallel testing is shown for comparison, not included. "
              "Third-party factory test at nil pending evidence that it is required. See 'Build-Up' Item 7 and the 'Programme' tab, windows P18 to P22.")
asm['L18'] = ("Priced to the RFP specification. The Contractor's different materials (HDPE pipe, gate valve, motorised valve) are not approved and are shown for "
              "information. No quantities yet: shop drawings rejected, take-off awaited. See 'Build-Up' Item 8.")
asm['L12'] = ("Lowest of the Contractor's three quotations, adopted for dismantling, segregation and loading only on the scope recorded in the previous revision; the quotation itself is not attached to this revision and its scope split is not yet confirmed in writing. Haulage and handover are Item 4. See 'Build-Up' Item 3.")
asm['N17'] = 'Provisional - commissioning method, water source and disinfection to be accepted'
bu['A93'] = ("The Contractor's adopted lowest quotation is taken at net supplier cost; the 5 per cent Overhead and Profit is applied once on the 'Assessment' tab. The quotation is for an insulated tank. The Engineer's (KEO) email of 30-Aug-2026: thermal insulation is not required, so a lower non-insulated price is expected and has been requested; the rate below will change when it is received. Scope: supply, installation and testing on the split described in Assumption 9 - the Contractor provides helpers, unloading and shifting, storage, plant, scaffolding, power, piping and test water. Market indication only, not adopted: the Al Muhaideb quotation MNT-AY-486 of 24-Jun-2026 prices a 3,000 m3 non-insulated tank at SAR 1,575,000 net (SAR 525 per gross m3) against the adopted insulated SAR 677 per m3 - a different supplier, size, count and date, so it supports the expectation of a lower non-insulated price without fixing one.")
bu.row_dimensions[93].height = 84
asm['A22'] = ("Contractor columns are as submitted. Assessed rates (column J) come from the 'Build-Up' tab; their time periods come from the Contractor's programme "
              "('XER WBS' tab as received, 'Programme' tab worked through). 'Provisional' means a confirmation is still outstanding ('Build-Up Comparison' tab, "
              "Section 7). The programme is under the Engineer's approval, not agreed; the Contractor has been instructed and is on site, the instruction reference "
              "not yet supplied.")
asm.row_dimensions[22].height = 40
asm['A23'] = (f"This assessment is preliminary. It establishes a reasonable commercial provision on the information available at {DOCDATE} and does not constitute agreement of the final Variation value. "
              "Readiness for external issue is conditional: the testing and commissioning method (Item 7) has not been accepted by the Engineer, and the Contractor's original supplier offers, the labour ownership behind the approved manpower histogram and the pipework take-off remain outstanding; until then the items marked Provisional are allowances, not agreed values.")
asm.row_dimensions[23].height = 42
asm['A24'] = ('="Changes from Rev 01 dated 17-Sep-2026 (SAR 6,921,685.64): time periods taken from the Contractor\'s programme instead of an assumed 3 months (Items 1 and 5); '
              'helpers from the approved manpower histogram, priced once (Item 5); testing re-sequenced to the Engineer\'s one-fill basis with staff reconciled to it (Item 7); '
              'lines within the supplier\'s scope or not required removed (5.12, 7.4, 7.12) and 5.13 re-scoped as the Contractor\'s storage and handling consumables, on the supplier scope as described (\'Build-Up\' Assumption 9); PPE, survey and document control increased (1.18, 2.3, 2.4); method statements and plans (2.5), Tank 1 operating fill and spray disinfection (7.24, 7.25) added; '
              'Item 8 priced to the RFP specification with the Contractor\'s materials shown alongside. No existing rate has moved. Net effect: SAR " & TEXT(K21-6921685.64,"#,##0.00;-#,##0.00") & " (" & TEXT((K21-6921685.64)/6921685.64,"0.0%;-0.0%") & ")."')
def _split_literals(f, n=240):
    """Excel limits a string literal inside a formula to 255 characters: split long literals with &."""
    out = []; i = 0
    for m in re.finditer(r'"((?:[^"]|"")*)"', f):
        out.append(f[i:m.start()]); lit = m.group(1)
        parts = [lit[k:k + n] for k in range(0, len(lit), n)] or ['']
        out.append('&'.join('"' + q + '"' for q in parts)); i = m.end()
    out.append(f[i:]); return ''.join(out)
asm['A24'] = _split_literals(asm['A24'].value); asm['A4'] = _split_literals(asm['A4'].value)
asm.row_dimensions[24].height = 60

# ---------------------------------------------------------------- Build-Up Comparison
bc = wb['Build-Up Comparison']
bc['H48'] = "Helper-days assessed phase by phase on the 'Programme' tab, Section 4, converted to man-months ('Build-Up' 5.10)."
bc['H41'] = "Assessed as a telehandler for the hire days on the 'Programme' tab, Section 4 ('Build-Up' 5.2): the existing base frames and 4 m panel height do not call for a mobile crane. The Contractor's plant schedule is awaited. No capacity, units, period or rate stated."
bc['H42'] = "Assessed as a boom truck for the hire days on the 'Programme' tab, Section 4 ('Build-Up' 5.1). No units, period or rate stated."
bc['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'
bc.freeze_panes = 'A4'
bc['A73'] = ("The substantiation request was issued by WT Partnership to the Contractor via Aconex on 25-Aug-2026 under Sub-Clauses 13.3.1 and 12.3.2 of the Conditions of "
             "Contract. A follow-up clarification request was prepared after the Engineer's (KEO) email of 30-Aug-2026; its issue is to be confirmed before it is relied on.")
for rr in (76, 77, 78):
    bc.cell(rr, 3).value = "WT Partnership Aconex substantiation request dated 25-Aug-2026; WT Partnership follow-up clarification request (issue to be confirmed)."
bc['H74'] = f'Status at {DOCDATE}'
bc['H78'] = 'Baseline programme (data date 01-Jul-2026), S-curve, cash flow, manpower histogram and look-ahead received 29-Sep-2026; under the Engineer\'s approval. The impact on the Key Dates and Time for Completion is not stated'
bc['H83'] = 'Not submitted; the programme is cost-loaded but carries no plant resources'
bc['H25'] = ("Mobilisation is more than transport, and each part is assessed once, elsewhere: plant and cabin transport at 'Build-Up' 1.14 (demobilisation 1.15, against line 1.8 here); method statements, risk assessments and project plans at 2.5; submittal administration at 2.4; the dedicated site engineer and HSE officer, who run the inductions, permits, progress meetings and reporting, at 1.1 and 1.2; design coordination at 2.1. Head-office attendance at progress meetings is in the 5 per cent Overhead and Profit. The tank installer's own mobilisation is inside the supplier price (Item 6, Assumption 9). No trips or rates were stated by the Contractor.")
bc.row_dimensions[25].height = 96
bc.row_dimensions[78].height = 40
# new requests rows 84-87 (insert 4 rows before conclusion at 85 -> conclusion moves to 89)
insert_rows_keep_styles(bc, 84, 5, 83)
new_req = [
    (10, 'Whose labour the direct manpower histogram (221 man-weeks, 28-Aug to 11-Dec-2026) represents: whether the tank supplier\'s erection crews, priced within Item 6, are included', 'This assessment.', 'Not yet requested'),
    (11, 'Quantity take-off for Item 8 from the pipework shop drawings once approved (returned Code C; resubmission 28 to 30-Sep-2026 per the look-ahead)', 'This assessment.', 'Not yet requested'),
    (12, 'Material approval requests for the HDPE main-line pipe, the gate valve DN300 and the motorised butterfly valve DN355 against the RFP General Piping Requirements, and the power and control supply for the motorised valve', 'This assessment.', 'Not yet requested'),
    (13, 'Basis on which a third-party factory acceptance test is being arranged (PQD, ITP, procedure and inspector CV submitted 26 to 28-Sep-2026): whether instructed by the Engineer or the Contractor\'s own quality plan', 'This assessment.', 'Not yet requested'),
    (14, 'Copy of the adopted Al Mousa S04488 / Stalwart SS-07-26-1516 offer with its conditions, to confirm the supply-and-install scope split applied meanwhile from the Al Muhaideb quotation MNT-AY-486 (\'Build-Up\' Assumption 9: helpers, unloading and shifting, plant, scaffolding, power, storage, piping and test water excluded by the supplier; sealant and fixings included)', 'This assessment.', 'Not yet requested'),
]
for i, (n, req, src, st) in enumerate(new_req):
    rr = 84 + i
    bc.cell(rr, 1).value = n; bc.cell(rr, 2).value = req; bc.cell(rr, 3).value = src; bc.cell(rr, 8).value = st
    bc.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=7)
    bc.row_dimensions[rr].height = 40
# ensure rows 84-87 have merged C:G styles like row 83 (copied). Conclusion now at 89.
assert str(bc['A90'].value).startswith('="Conclusion')
# Section 8 after conclusion
r = 92
bc.cell(r, 1, "8.  THE CONTRACTOR'S PROGRAMME DOCUMENTS OF 29-SEP-2026 AS EVIDENCE");
for c in range(1, 9): cp(bc.cell(72, c), bc.cell(r, c))
bc.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8); bc.row_dimensions[r].height = 21; r += 1
def bcpara(text, height):
    global r
    bc.cell(r, 1, text)
    for c in range(1, 9): cp(bc.cell(73, c), bc.cell(r, c))
    bc.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8); bc.row_dimensions[r].height = height; r += 1
bcpara(("The programme is cost-loaded to the proposal value: the Contractor has spread SAR 8,110,296.62 including Overhead and Profit across its activities for "
        "progress measurement. That spread is the Contractor's own allocation of the sums claimed; it is evidence of how the Contractor intends to apply for payment "
        "against each activity, not evidence of the cost of any activity or of where any part of a lump sum sits. It is recorded here against the claimed and assessed "
        "amounts for that reason only."), 44)
# 8a table
hdr = ['Item', 'Programme section (as named)', 'Activities', '', 'Claimed by the Contractor (SAR)', 'Loaded in the programme, incl. 5% OHP (SAR)', 'Loaded, net of 5% OHP (SAR)', 'Note']
for i, t in enumerate(hdr, 1):
    bc.cell(r, i, t); cp(bc.cell(24, i), bc.cell(r, i))
bc.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
bc.row_dimensions[r].height = 44; r += 1
def L8(item, sec, acts_, claimed, loaded, note, h=30):
    global r
    bc.cell(r, 1, item); bc.cell(r, 2, sec); bc.cell(r, 3, acts_); bc.cell(r, 5, claimed); bc.cell(r, 6, loaded); bc.cell(r, 7, f"=ROUND(F{r}/1.05,2)"); bc.cell(r, 8, note)
    for c in range(1, 9): cp(bc.cell(25, c), bc.cell(r, c))
    for c in (5, 6, 7):
        bc.cell(r, c).number_format = '#,##0.00'; bc.cell(r, c).alignment = Alignment(horizontal='right', vertical='center')
    bc.cell(r, 3).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    bc.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    bc.row_dimensions[r].height = h; r += 1
L8('1', 'Mobilization (manpower and equipment; welfare facilities; project deliverables and submittals) and Demobilization and Closeout', 'QCD18TSEMOB1040 to 1250; QCD18TSEDMOB1020', '=Assessment!$F$8', 714000, 'Equals Item 1 x 1.05 exactly. Site establishment and welfare are loaded on the 5-day mobilisation activities (22 to 26-Aug-2026), so the Contractor will apply for the whole of Item 1 except SAR 50,000.00 on mobilisation', 44)
L8('2', 'Engineering - Shop Drawings', 'QCD18TSEENG1240 and 1250', '=Assessment!$F$10', 89250, 'Equals Item 2 x 1.05 exactly')
L8('3, 4', 'Construction - Dismantling of Existing TSE Tank; Survey and Levelling', 'QCD18TSECONDSM1020 to 1040; QCD18TSECONSL1050', '=Assessment!$F$12+Assessment!$F$13', 85575, 'Equals Items 3 and 4 x 1.05 exactly')
L8('5 to 8', 'Procurement (long-lead and short-lead items), Construction - TSE Tank 01 and 02, Testing and Commissioning, Tie-in Interface', 'QCD18TSEPRC1060 to 1210; QCD18TSECONT1 and CONT2 series; QCD18TSECONTCT12050 and CT22030; QCD18TSECONTC2040', '=Assessment!$F$14+Assessment!$F$16+Assessment!$F$17+Assessment!$F$18', 7221471.62, 'Equals Items 5 to 8 x 1.05 exactly. Item 5 (temporary works) is not loaded on any mobilisation activity: it is spread into the tank erection activities, which is consistent with plant and attendance being consumed by erection, not by the whole site period', 57)
L8('7', 'Within Items 5 to 8 above: Construction - Testing and Commissioning (both tanks)', 'QCD18TSECONTCT12050 and CT22030', '=Assessment!$F$17', 200305.97, "The Contractor's allocation to testing and commissioning is less than half of its Item 7 claim; the balance of Item 7 is not identifiable within the loading of the erection and tie-in activities. Recorded as the Contractor's own valuation of the testing activity, not as a cost", 57)
bc.cell(r, 2, 'Total loaded'); bc.cell(r, 6, f"=SUM(F{r-5}:F{r-2})"); bc.cell(r, 7, f"=ROUND(F{r}/1.05,2)"); bc.cell(r, 8, "Equals the proposal of SAR 8,110,296.62 including Overhead and Profit")
for c in range(1, 9): cp(bc.cell(34, c), bc.cell(r, c))
bc.cell(r, 6).number_format = '#,##0.00'; bc.cell(r, 7).number_format = '#,##0.00'
bc.row_dimensions[r].height = 21.75; r += 2
# 8b cash flow
bcpara(("Cash flow as submitted (workbook 'SAMA-D18 -THS-TSE - Cost S-Curves Cash Flow Manpower Histogram.xlsx', sheet 'Cash Flow - MN'). August carries the "
        "purchase order for the tanks (SAR 348,110.34) and the start of manufacturing: value the Contractor expects to draw before anything is installed. The "
        "payment stance for these works is no advance, no payment on account and no material on site, with tank works paid on installation and successful testing; "
        "the cash flow is recorded, not adopted, and the point is for certification, not for this assessment."), 57)
for i, t in enumerate(['', 'Month', 'Monthly (SAR)', '', 'Cumulative (SAR)', '', '', 'Note'], 1):
    bc.cell(r, i, t); cp(bc.cell(24, i), bc.cell(r, i))
bc.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4); bc.row_dimensions[r].height = 21; r += 1
cf = [('Jul-2026', 12035.96, 'Submittals'), ('Aug-2026', 1470052.21, 'Mobilisation, welfare, purchase order and manufacturing'), ('Sep-2026', 1005870.64, 'Dismantling, base-panel delivery and Tank 1 base'),
      ('Oct-2026', 2701160.67, 'Deliveries and erection of both tanks'), ('Nov-2026', 2417025.29, 'Erection and mechanical works'), ('Dec-2026', 504151.86, 'Tie-in, testing, demobilisation')]
cf0 = r
for i, (m, v, n) in enumerate(cf):
    bc.cell(r, 2, m); bc.cell(r, 3, v); bc.cell(r, 5, f"=SUM($C${cf0}:C{r})"); bc.cell(r, 8, n)
    for c in range(1, 9): cp(bc.cell(25, c), bc.cell(r, c))
    bc.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    for c in (3, 5): bc.cell(r, c).number_format = '#,##0.00'; bc.cell(r, c).alignment = Alignment(horizontal='right', vertical='center')
    bc.row_dimensions[r].height = 18; r += 1
bc.cell(r, 2, 'Total'); bc.cell(r, 3, f"=SUM(C{cf0}:C{r-1})"); bc.cell(r, 8, 'Equals the proposal within one halala of rounding')
for c in range(1, 9): cp(bc.cell(34, c), bc.cell(r, c))
bc.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4); bc.cell(r, 3).number_format = '#,##0.00'
bc.row_dimensions[r].height = 21.75; r += 2
# 8c manpower
bcpara(("Direct manpower histogram as submitted (same workbook, sheet 'MP-HST-WK'): 16 weeks of labour on site, week ending 28-Aug to week ending 11-Dec-2026; "
        "221 man-weeks in total, 136 skilled and 85 helpers; peak 35 in the week ending 30-Oct-2026 when the Tank 1 walls and bracing overlap the Tank 2 base. "
        "The workbook does not say whose labour it is. The tank supplier's price under Item 6 includes installation and requires 4 to 6 helpers from the Contractor; "
        "if the histogram includes the supplier's erection crews it cannot support the Contractor's Items 1, 5 or 7, and if it is the Contractor's own labour, 136 "
        "skilled man-weeks on a job whose erection is subcontracted are unexplained. Until the Contractor answers (Section 7, item 10), the histogram is used only as "
        "evidence of the 16-week labour window, which sits inside the works period allowed."), 70)
for i, t in enumerate(['', 'Week ending', 'Total', 'Skilled', 'Helper', '', '', 'Note'], 1):
    bc.cell(r, i, t); cp(bc.cell(24, i), bc.cell(r, i))
bc.row_dimensions[r].height = 21; r += 1
hist = [('28-Aug-2026', 8, 5, 3, 'Dismantling'), ('04-Sep-2026', 14, 9, 5, ''), ('11-Sep-2026', 13, 8, 5, ''), ('18-Sep-2026', 12, 8, 4, ''), ('25-Sep-2026', 11, 7, 4, ''),
        ('02-Oct-2026', 10, 6, 4, 'Tank 1 base panels'), ('09-Oct-2026', 10, 6, 4, ''), ('16-Oct-2026', 8, 4, 4, ''), ('23-Oct-2026', 14, 8, 6, ''), ('30-Oct-2026', 35, 23, 12, 'Peak: Tank 1 walls and bracing with Tank 2 base'),
        ('06-Nov-2026', 26, 16, 10, ''), ('13-Nov-2026', 19, 11, 8, ''), ('20-Nov-2026', 16, 10, 6, ''), ('27-Nov-2026', 13, 8, 5, ''), ('04-Dec-2026', 7, 4, 3, 'Mechanical works and tie-in'), ('11-Dec-2026', 5, 3, 2, 'Testing')]
h0 = r
for wk, t, s, h, n in hist:
    bc.cell(r, 2, wk); bc.cell(r, 3, t); bc.cell(r, 4, s); bc.cell(r, 5, h); bc.cell(r, 8, n)
    for c in range(1, 9): cp(bc.cell(25, c), bc.cell(r, c))
    for c in (3, 4, 5): bc.cell(r, c).number_format = '#,##0'; bc.cell(r, c).alignment = Alignment(horizontal='right', vertical='center')
    bc.row_dimensions[r].height = 16; r += 1
bc.cell(r, 2, 'Total man-weeks'); bc.cell(r, 3, f"=SUM(C{h0}:C{r-1})"); bc.cell(r, 4, f"=SUM(D{h0}:D{r-1})"); bc.cell(r, 5, f"=SUM(E{h0}:E{r-1})"); bc.cell(r, 8, 'Whose labour: to be confirmed by the Contractor')
for c in range(1, 9): cp(bc.cell(34, c), bc.cell(r, c))
for c in (3, 4, 5): bc.cell(r, c).number_format = '#,##0'
bc.row_dimensions[r].height = 21.75

# ---------------------------------------------------------------- Programme tab: print setup and tab order
# ================================================================ Navigation and visual pass: labels, links, freeze panes, print titles
LINK_BLUE = '0563C1'
def linkify(cell, target, tip=None):
    """Internal hyperlink on a cell, keeping its font but coloured and underlined."""
    cell.hyperlink = target
    f = copy.copy(cell.font)
    cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, color=LINK_BLUE, u='single')
    if tip: cell.comment = None
def navlink(ws, row, col, text, target, ncol=None):
    c = ws.cell(row, col, text); cp(S_NOTE, c)
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
    c.font = Font(name='PT Sans', sz=9, i=True, color=LINK_BLUE, u='single')
    c.hyperlink = target
    if ncol: ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=ncol)

# --- plain-language phrase pass over every text cell (not formulas)
ASM_ITEM_ROW = {8: 1, 10: 2, 12: 3, 13: 4, 14: 5, 16: 6, 17: 7, 18: 8}
PHRASES = [
 (re.compile(r"carried to 'Assessment'!J(\d+)"), lambda m: f"goes to 'Assessment' item {ASM_ITEM_ROW[int(m.group(1))]}, column J"),
 ("man-days", "person-days"), ("man-day", "person-day"), ("Man-days", "Person-days"), ("man-weeks", "person-weeks"), ("man-week", "person-week"), ("man-months", "person-months"), ("man-month", "person-month"),
 ("the departure", "the additional helpers assessed"), ("as the departure", "as 'Additional helpers assessed'"), ("Departure", "Additional helpers assessed"), ("departure", "additional helpers assessed"),
 ("assessment allowance", "Assessed basis"), ("Assessment allowance", "Assessed basis"), ("assessment-allowance", "Assessed-basis"),
 ("submitted programme", "SAMA Submitted Programme"), ("Submitted programme", "SAMA Submitted Programme"), ("submitted-programme", "SAMA Submitted Programme"),
 ("the programme as submitted", "the SAMA Submitted Programme"), ("programme exactly as submitted", "SAMA Submitted Programme exactly as submitted"),
 ("'As submitted'", "'SAMA Submitted Programme'"), ("As submitted:", "SAMA Submitted Programme:"), ("As submitted only:", "SAMA Submitted Programme only:"),
 ("Contractor columns are as submitted", "The SAMA Submitted Cost columns are as submitted"),
 ("Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation", "Rate: assessed allowance (Riyadh market, Sep-2026) - needs confirmation by the Contractor"),
 ("Assessed market rate, Riyadh, Sep-2026 - assumption", "Rate: assessed allowance (Riyadh market, Sep-2026) - needs confirmation"),
 ("- assumption pending the Contractor's substantiation", "- needs confirmation by the Contractor"),
 ("Provisionally not assessed, to avoid duplication:", "Included in the supplier's price (provisional, to avoid double counting):"),
 ("Nil: the transfer labour is within the approved manpower histogram", "Included in line 5.10 (approved histogram helpers): the transfer labour is within the approved manpower histogram"),
 ("Carried basis if", "Assessment allowance if"), ("carried basis", "assessment allowance"), ("Carried basis", "Assessment allowance"),
 ("the carried quantity", "the assessment-allowance quantity"), ("Carried:", "Assessment allowance:"), ("Carried =", "Assessment allowance ="),
 ("as-submitted amounts", "submitted-programme amounts"), ("as submitted and carried", "in the SAMA Submitted Programme and in the Assessed basis"),
 ("as submitted as submitted and carried", "in the SAMA Submitted Programme and in the Assessed basis"),
 ("not carried into the assessment", "not included in the assessment"), ("Not carried", "Not included in the assessment"), ("not carried", "not included in the assessment"),
 ("is carried at nil", "is included at nil"), ("hire days carried", "hire days used"), ("Hire days carried", "Hire days used"),
 ("the departure carried to", "the departure taken to"), ("is carried to 'Build-Up'", "goes to 'Build-Up'"), ("carried to 'Build-Up'", "taken to 'Build-Up'"),
 ("the approved figures are carried", "the approved figures are used"), ("carried for the transfer", "allowed for the transfer"), ("1 week carried", "1 week allowed"),
 ("Nil as submitted and carried", "Nil on both bases"), ("Carried for", "Allowed for"), ("carried into", "included in"),
 ("is carried", "is used"), ("are carried", "are used"), ("carried", "used"), ("Carried", "Used"),
]
def phrase_pass(ws):
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not isinstance(v, str) or v.startswith('=') or not v.strip(): continue
            new = v
            for a, b in PHRASES:
                new = a.sub(b, new) if hasattr(a, 'sub') else new.replace(a, b)
            if new != v: c.value = new
for _n in ('Assessment', 'Build-Up', 'Build-Up Comparison', 'Programme'):
    phrase_pass(wb[_n])

# --- Build-Up: quantity-source prefix, units, how-to text, freeze, print titles
bu['A4'] = ("How to read this tab: each line is quantity x rate = amount, with the reason beside it ('Basis': what the line buys, how many, how long, the rate and "
            "where the quantity comes from). Item totals go to column J of the 'Assessment' tab; the 'Return' link on each total row goes back there. A quantity in blue "
            "is a link to the 'Programme' tab row it is calculated from - click it. Units: a 'month' is 30.4 calendar days; a 'day' is one calendar day of hire or "
            "attendance; a 'person-day' is one person for one working day; working days exclude Fridays. 'Needs confirmation' marks an assessed allowance the "
            "Contractor has not yet substantiated; 'Provisional' marks a scope or price awaiting a document. All amounts exclude VAT.")
bu.row_dimensions[4].height = 84
for row in bu.iter_rows(min_row=9, max_row=bu.max_row):
    d, g = row[3], row[6]
    if isinstance(d.value, str) and d.value.startswith('=Programme!'):
        m = re.search(r'\$?([A-Z])\$?(\d+)', d.value.split('!')[1])
        linkify(d, f"#'Programme'!{'A' if pg.cell(int(m.group(2)), 1).value else 'B'}{m.group(2)}")
bu.freeze_panes = 'A4'
bu.column_dimensions['C'].width = 12
bu.print_title_rows = '1:3'
# item banners and totals
BU_ITEM_BANNER, BU_ITEM_TOTAL = {}, {}
for row in bu.iter_rows(min_row=5, max_row=bu.max_row, max_col=7):
    a, b = row[0].value, row[1].value
    if isinstance(a, str) and re.match(r'ITEM (\d) -', a):
        BU_ITEM_BANNER[int(re.match(r'ITEM (\d) -', a).group(1))] = row[0].row
    if isinstance(b, str) and (b.startswith('TOTAL ITEM') or b.startswith('RATE PER TANK') or b.startswith('Rate per tank')):
        m = re.search(r"item (\d)", b)
        if m: BU_ITEM_TOTAL[int(m.group(1))] = row[1].row
for item, arow in ASM_ITEM_ROW.items():
    pass
ASM_ROW_OF_ITEM = {v: k for k, v in ASM_ITEM_ROW.items()}
for item, trow in BU_ITEM_TOTAL.items():
    linkify(bu.cell(trow, 2), f"#'Assessment'!A{ASM_ROW_OF_ITEM[item]}")
    navlink(bu, trow, 7, f"Return to 'Assessment' item {item}", f"#'Assessment'!A{ASM_ROW_OF_ITEM[item]}")
    bu.cell(trow, 7).alignment = Alignment(horizontal='right', vertical='center')
navlink(bu, 5, 2, "Back to 'Assessment'", "#'Assessment'!A1")
navlink(bu, 5, 7, "Forward to 'Programme' (dates, people and plant behind the quantities)", "#'Programme'!A1")
bu.row_dimensions[5].height = 15

# --- Assessment: item links, notes
for arow, item in ASM_ITEM_ROW.items():
    linkify(asm.cell(arow, 1), f"#'Build-Up'!A{BU_ITEM_BANNER[item]}")
    linkify(asm.cell(arow, 10), f"#'Build-Up'!B{BU_ITEM_TOTAL[item]}")
asm['A22'] = ("Contractor columns are as submitted. Assessed rates (column J) come from the 'Build-Up' tab: click an item number or a blue rate to open its build-up; on "
              "the 'Build-Up' tab a blue quantity opens the 'Programme' tab row it is calculated from, and on the 'Programme' tab a blue date opens the Contractor's "
              "activity on the 'XER WBS' tab. 'Provisional' means a confirmation is still outstanding ('Build-Up Comparison' tab, Section 7). The programme is under "
              "the Engineer's approval, not agreed; the Contractor has been instructed and is on site, the instruction reference not yet supplied.")
asm.row_dimensions[22].height = 54
asm.print_title_rows = '1:6'
asm['C5'] = 'SAMA SUBMITTED COST - SAMACO-RRFP-000001, 13-AUG-2026'
asm['H5'] = 'ASSESSED'
asm['L6'] = 'Assessed - reason'

# --- Programme: links from every formula that looks up an activity ID, and from Section 5 / 6 references
def first_id(text):
    m = re.search(r'"(QCD18TSE[A-Z0-9]+)"', text) or re.search(r'\b(QCD18TSE[A-Z0-9]+)\b', text)
    return m.group(1) if m else None
for row in pg.iter_rows(min_row=6, max_row=pg.max_row, max_col=10):
    for c in row:
        v = c.value
        if isinstance(v, str) and v.startswith('=') and 'MATCH("QCD18TSE' in v:
            aid = first_id(v)
            if aid in XROW: linkify(c, f"#'XER WBS'!A{XROW[aid]}")
        elif isinstance(v, str) and v.startswith('=SUMPRODUCT') and "'XER WBS'" in v:
            ds = [dt.date(*map(int, m)) for m in re.findall(r'DATE\((\d+),(\d+),(\d+)\)', v)]
            if len(ds) == 2:
                w_end, w_start = ds
                live = [a for a in acts if 'Tank Installation' in a['sec'] and a['start'] <= w_end and a['finish'] >= w_start]
                if live: linkify(c, f"#'XER WBS'!A{XROW[live[0]['id']]}")
# Section 5 rows: column A -> Build-Up line, column G -> source row on this tab
for rr in range(first4, last4 + 1):
    b = pg.cell(rr, 2).value
    if isinstance(b, str) and b.startswith("='Build-Up'!B"):
        linkify(pg.cell(rr, 1), f"#'Build-Up'!A{b.split('B')[-1]}")
    g = pg.cell(rr, 7).value
    if isinstance(g, str) and g.startswith('='):
        m = re.search(r'(?<![A-Z$])\$?([C-I])\$?(\d+)', g)
        if m and int(m.group(2)) < first4 and (pg.cell(int(m.group(2)), 1).value or pg.cell(int(m.group(2)), 2).value):
            linkify(pg.cell(rr, 7), f"#'Programme'!{'A' if pg.cell(int(m.group(2)), 1).value else 'B'}{m.group(2)}")
# Section 6 register: column A -> Build-Up line; column C -> XER activity
for rr in range(g_first, g_last + 1):
    b = pg.cell(rr, 2).value
    if isinstance(b, str) and b.startswith("='Build-Up'!B"):
        linkify(pg.cell(rr, 1), f"#'Build-Up'!A{b.split('B')[-1]}")
    cc = pg.cell(rr, 3).value
    if isinstance(cc, str):
        aid = first_id(cc)
        if aid in XROW: linkify(pg.cell(rr, 3), f"#'XER WBS'!A{XROW[aid]}")
# Section 4 roles and plant: dates that copy a Section 3 row
for rr in range(SEC3_END + 1, first4):
    for col in (3, 4):
        v = pg.cell(rr, col).value
        if isinstance(v, str) and re.fullmatch(r'=[C-H]\d+', v) and int(v[2:]) < SEC3_END:
            linkify(pg.cell(rr, col), f"#'Programme'!A{v[2:]}")
navlink(pg, 5, 2, "Back to 'Build-Up'", "#'Build-Up'!A1")
navlink(pg, 5, 10, "Forward to 'XER WBS' (the Contractor's programme as received)   |   Back to 'Assessment'", "#'XER WBS'!A1")
pg.row_dimensions[5].height = 15
# Section 5 comparison columns shaded
GREY = PatternFill('solid', fgColor='EDF0F2')
for rr in range(first4, last4 + 1):
    for col in (5, 6, 8):
        if pg.cell(rr, col).value not in (None, ''): pg.cell(rr, col).fill = GREY

# --- generic pass: any remaining formula that reads another tab or another row gets a link to its first source
REF_RE = re.compile(r"(?:'([^']+)'!|(?<![A-Za-z_])([A-Za-z][A-Za-z ]*?)!)?\$?([A-Z]{1,2})\$?(\d+)(?![\d(])")
def first_source(formula, own, own_row):
    for m in REF_RE.finditer(formula):
        sheet = m.group(1) or m.group(2) or own
        end = m.end()
        if formula[end:end + 1] == ':' or formula[m.start() - 1:m.start()] == ':':   # a range: skip
            continue
        if sheet not in wb.sheetnames: continue
        col, row = m.group(3), int(m.group(4))
        if sheet == own and row == own_row: continue
        if sheet == own and m.group(0).count('$') == 2: continue   # fixed calendar constants ($C$14, $C$15) are not a source
        return sheet, col, row
    return None
def label_col(ws, row, col):
    if ws.cell(row, 1).value not in (None, ''): return 'A'
    if ws.cell(row, 2).value not in (None, ''): return 'B'
    return col
added = 0
for ws in (asm, bu, pg, wb['Build-Up Comparison']):
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if c.hyperlink or not isinstance(v, str) or not v.startswith('=') or v.startswith('="'): continue
            src = first_source(v, ws.title, c.row)
            if not src: continue
            sheet, col, r_ = src
            if sheet == ws.title and r_ == c.row: continue
            tgt_ws = wb[sheet]
            if tgt_ws.cell(r_, 1).value in (None, '') and tgt_ws.cell(r_, 2).value in (None, '') and tgt_ws.cell(r_, openpyxl.utils.column_index_from_string(col)).value in (None, ''): continue
            linkify(c, f"#'{sheet}'!{label_col(tgt_ws, r_, col)}{r_}"); added += 1
print('generic links added', added)
# --- XER WBS: navigation row
navlink(wx, 5, 2, "Back to 'Programme'", "#'Programme'!A1")
navlink(wx, 5, 13, "Back to 'Assessment'", "#'Assessment'!A1")
wx.row_dimensions[5].height = 15


for name in ('Assessment', 'Build-Up', 'Build-Up Comparison', 'Programme'):
    autofit(wb[name])
wb.active = wb.sheetnames.index('Project Info')
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = (ws.title == 'Project Info')
wb.save(OUT)
print('saved', OUT, 'Programme rows: sec3', SEC3_START, '-', SEC3_END, 'sec4 hdr', HDR4ROW, 'tot', TOT4)
