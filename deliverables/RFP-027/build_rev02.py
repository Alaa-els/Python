"""Rev 02 build - re-basing the RFP-027 assessment to the Contractor's programme of 29-Sep-2026.
Stage A: openpyxl on Alaa's master (Rev 01, 17-Sep-2026). Stage B (separate): recalc, zip-level metadata."""
import copy, datetime as dt, re, sys
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.pagebreak import Break
import math
PAGE = 930.0   # printable height of one landscape page at the fit-to-width scale, in points (measured on the rendered file)
def est(text, width, sz=10):
    cpl = max(8, width * 1.15 * 10 / sz)
    return math.ceil(len(text) / cpl) * (sz * 1.28) + 4

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
bu['B123'] = 'TOTAL ITEM 7 - both tanks'
bu['F123'] = '=SUM(F100:F122)'
bu['F124'] = '=F123/2'
bu.row_dimensions[122].height = 78

# ---------------------------------------------------------------- Programme tab
pg = wb.create_sheet('Programme', index=wb.sheetnames.index('Build-Up Comparison') + 1)
widths = {'A': 6, 'B': 46, 'C': 12, 'D': 12, 'E': 9, 'F': 14, 'G': 20, 'H': 14, 'I': 14, 'J': 52}
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
    ws.row_dimensions[r].height = height or max(15, 13 * (len(text) // 210 + 1))

def header(ws, r, labels, merge_ij=False):
    for i, t in enumerate(labels, 1):
        c = ws.cell(r, i, t); cp(S_HDR, c)
    if merge_ij:
        cp(S_HDR, ws.cell(r, 10)); ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    ws.row_dimensions[r].height = 32

PG = {'used': 0.0}
def pg_add(h): PG['used'] += h
def pg_break(ws, r):
    ws.row_breaks.append(Break(id=r - 1)); PG['used'] = 58.0   # the repeated title rows 1 to 3
HDR3 = ['Ref', 'Window', 'Start - Basis A', 'Finish - Basis A', 'Days - A', 'Start - Basis B', 'Finish - Basis B', 'Days - B', 'Derivation and source']
HDR4 = ['Ref', "'Build-Up' line", 'Unit', 'Rate (SAR)', 'Qty - previous revision', 'Qty - Basis A', 'Qty - Basis B (carried)', 'Amount - Basis A (SAR)', 'Amount - Basis B (SAR)', 'Derivation']
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
para(pg, 4, ("How to read this tab: Section 1 records the status of the programme and of the instruction. Section 2 is the working calendar "
             "taken from the programme file. Section 3 sets the programme's windows as submitted beside the basis carried, every date derived by "
             "formula, by activity ID, from the 'XER WBS' tab, which holds the programme file as received. Section 4 is the resource bridge: the approved "
             "manpower histogram week by week, reconciled to the priced lines, plant days built from the activity dates, and site staff by phase. Section 5 lists each 'Build-Up' line whose quantity is taken from this tab, with the programme exactly as "
             "submitted beside the basis carried; the carried quantity is the one linked into column D of the 'Build-Up' tab. Section 6 links every 'Build-Up' line to its activities or site service. All amounts exclude VAT."), height=70)

# Section 1 - status
banner(pg, 6, '1. STATUS OF THE PROGRAMME AND OF THE INSTRUCTION')
para(pg, 7, (f"Programme: baseline schedule 'QC05958-BSL01THF-TSE-Final' ({XER}, data date 01-Jul-2026, 81 activities, cost-loaded to the "
             "proposal value of SAR 8,110,296.62 including 5 per cent Overhead and Profit), submitted by the Contractor to the Engineer on "
             "12-Sep-2026 and issued to the Cost Consultant on 29-Sep-2026 with a cost S-curve, a monthly cash flow, a weekly direct-manpower "
             "histogram and a two-week look-ahead (data date 28-Sep-2026). Status at 30-Sep-2026: under the Engineer's approval process - "
             "returned as acceptable with minor comments, a procurement schedule being required; not approved. The programme is therefore "
             "evidence of the Contractor's intended sequence and durations, not an agreed basis; every line on the 'Build-Up' tab that relies "
             "on it says so."), height=70)
para(pg, 8, ("Instruction: the Contractor has been instructed to proceed and is on site (the existing tank is dismantled and the Tank 1 base "
             "panels are being received). The instruction reference and date have not been supplied to the Cost Consultant and are to be "
             "recorded on the 'Project Info' tab when received. RFP-027 dated 28-Jun-2026 stated that it was not itself an instruction to "
             "proceed. The instruction affects entitlement and the certification route, not the assessed value."), height=44)
para(pg, 9, ("Use made of the programme: (i) the site period is taken as submitted, its end being governed by the external milestone 'SAJCO "
             "Readiness for Tie-In Connections' on 03-Dec-2026 and the tie-in, sequential testing, commissioning and demobilisation that follow it; "
             "(ii) the erection windows and their durations are taken as submitted, and the labour and plant that attend them are assessed for those "
             "durations (Section 4); (iii) testing is carried on the Engineer's basis of "
             "30-Aug-2026 (one tank filled, water re-used for the second), fitted to the submitted dates. Slippage against the baseline shown in the "
             "look-ahead (Tank 1 base panels forecast to start 30-Sep-2026 against 12-Sep-2026 programmed) is delivery-driven and adds nothing to "
             "the assessment. Dates after the look-ahead data date of 28-Sep-2026 are forecasts, not actuals."), height=83)

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
banner(pg, 17, '3. PROGRAMME WINDOWS - A: AS SUBMITTED; B: AS CARRIED (SUBMITTED ERECTION DATES, SEQUENTIAL TESTING)')
para(pg, 18, ("Column A reproduces the submitted programme, including its parallel testing. Column B is the basis carried: the erection dates exactly as "
              "submitted - no activity is compressed - with the Engineer's sequential testing fitted after the tank each test needs, and demobilisation "
              "following completion. Both columns are derived by formula from the 'XER WBS' tab by activity ID and the calendar in Section 2. The "
              "resources that each window carries are assessed in Section 4; the site period is set by "
              "the external readiness milestone of 03-Dec-2026 and the tie-in, testing, commissioning and demobilisation that follow it."), height=70)
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
    height = max(height or 30, est(window, 46), est(deriv, 114))
    if PG['used'] + height > PAGE:
        old = r
        pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(32)
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
           f"=C{r}", f"=D{r}", '-', "Activity QCD18TSECONIF2050, 03-Dec-2026, zero duration. An interface milestone outside the Contractor's control, taken as submitted on both bases; it, not the tank erection, fixes the earliest tie-in date", 'ms')
p16 = add3('P16', 'Tie-in connections with the existing pump room (working days)', f"={AS('QCD18TSECONTC2040')}", f"={AF('QCD18TSECONTC2040')}", f"={WD(f'C{r}', f'D{r}')}",
           *same(r), "Activity QCD18TSECONTC2040, 05 to 08-Dec-2026, 4 working days, following the readiness milestone")
p17 = add3('P17', 'Testing and commissioning as programmed - both tanks in parallel (working days)', f"={AS('QCD18TSECONTCT12050')}", f"={AF('QCD18TSECONTCT22030')}", f"={WD(f'C{r}', f'D{r}')}",
           '-', '-', '-', "Basis A only: activities QCD18TSECONTCT12050 and QCD18TSECONTCT22030, 09 to 16-Dec-2026, 7 working days each, in parallel, with two simultaneous fills. Not carried: the Engineer's (KEO) email of 30-Aug-2026 states that installation and testing will not be in parallel and that one tank is filled and the water re-used for the second. The sequential, task-based sequence carried is at P18 to P22")
p18 = add3('P18', 'Hydrostatic test - Tank 1, before the tie-in (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'D{p8}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONTCT12050')}",
           "Carried basis: starts the working day after the Tank 1 mechanical works (submitted date) and takes the 7 working days the Contractor programmed per tank. Elapsed sequence assumed within those 7 days: tankered filling about 5 working days, the 24-hour hold (unattended apart from level readings), inspection of joints and nozzles and records 1 day. The fill rate is an assumption, not a measured throughput: 3,774 m3 in 5 days needs about 750 m3 a day, for example two 30 m3 tankers on about 12 round trips each, which depends on the water source the Employer has yet to confirm; if the source is further away the elapsed time lengthens but the attendance does not. Attendance: the QA/QC inspector (1.3, monthly) and the commissioning engineer for one day at the end of the hold (7.13); the supplier's leak-test supervision is within Item 6. It sits inside the 15 working days of float the programme gives Tank 1 (17 to 24-Nov-2026). Conditional on an Engineer-approved method: internals flushed and nozzles blind-flanged (RFP Scope of Works work packages 4 and 5), the water retained in Tank 1 until Tank 2 is ready, losses at 'Build-Up' line 7.5", 96)
p19 = add3('P19', 'Transfer of the test water from Tank 1 to Tank 2 (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'MAX(G{p18},G{p14})')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "Starts the working day after both the Tank 1 test has passed (P18) and Tank 2 is ready to receive water (mechanical works complete, P14, interior flushed). Pumped tank to tank through temporary hoses with the outlet valves isolated; the tie-in is not needed for the transfer. 3 working days is an assessed assumption pending the Contractor's method statement: 3,774 m3 at about 130 m3 an hour over 10-hour shifts, a 150 mm self-priming diesel pump against a low head (adjacent tanks at one level, about 4 m static plus hose friction) - 'Build-Up' lines 7.2 to 7.4")
p20 = add3('P20', 'Hydrostatic test - Tank 2, after the transfer (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'G{p19}')}", f"={W(f'F{r}', f'H{r}')}", 4,
           "Elapsed: the Tank 2 test sequence is 7 working days from the start of the transfer, the same as the programme's per-tank figure - 3 days of transfer (P19, attended by the transfer labour at 7.4), then 4 days here: top-up to test level (1), the 24-hour hold (1, unattended apart from level readings), inspection of joints and nozzles and records (2). Attendance as for P18. Assessed assumption")
p21 = add3('P21', 'Component and subsystem checks after the tie-in - instruments, nozzles, valves, ladders; tank-pump-network interfaces (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'G{p16}')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "RFP Scope of Works 5.2, component testing and subsystem validation. Follows the tie-in and may overlap the Tank 2 hydrostatic test, because these checks do not need both tanks in service. 3 working days is an assessed assumption, not a programme figure: the Contractor's programme has no separate activity for this stage (its two 7-working-day 'Testing & Commissioning' activities cover the hydrostatic test and the commissioning of each tank together). Staffing at 'Build-Up' lines 7.13 to 7.18")
p21b = add3('P21b', 'Integrated system commissioning - full operational demonstration and witness testing (working days)', f"=G{p17}", f"=G{p17}", f"=G{p17}",
           f"={NEXT(f'MAX(G{p16},G{p20},G{p21})')}", f"={W(f'F{r}', f'H{r}')}", 3,
           "RFP Scope of Works 5.2, integrated system commissioning. Starts the working day after the last of the tie-in (P16), the Tank 2 hydrostatic test (P20) and the component checks (P21): both tanks must have passed and be connected before full operation is demonstrated. 3 working days (pump interface run, network demonstration, witnessed test and records) is an assessed assumption on the same footing as P21")
p22 = add3('P22', 'Completion of testing and commissioning - both tanks', f"={AF('QCD18TSEOMS1040')}", f"=C{r}", '-', f"=G{p21b}", f"=F{r}", '-',
           "A: completion milestones QCD18TSEOMS1040 and 1050, 16-Dec-2026, with parallel testing. B: end of the integrated system commissioning, conditional on the Tank 1 test preceding the tie-in (P18); if it cannot, see sensitivity S2", 'ms')
p23 = add3('P23', 'Demobilisation, as-built drawings and close-out documents (working days)', f"={AS('QCD18TSEDMOB1020')}", f"={AF('QCD18TSEDMOB1020')}", f"={WD(f'C{r}', f'D{r}')}",
           f"={NEXT(f'G{p22}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}", "Activities QCD18TSEDMOB1020 and QCD18TSEDMOB3020, 17 to 24-Dec-2026, 7 working days in parallel. B follows the completion of testing (P22)")
# derived periods
if PG['used'] + 19.5 + 120 > PAGE:
    pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(32)
pg_add(19.5)
pg.cell(r, 2, 'Derived periods used in Sections 4 and 5'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5
r += 1
d1 = add3('D1', 'Site period - mobilisation start to demobilisation finish (calendar days)', f"=C{p1}", f"=D{p23}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p23}", f"=G{r}-F{r}+1",
          "The continuous site establishment runs for this period: welfare cabins, WC, water tank and deliveries, the 30 kVA welfare generator, the site pick-up, workforce transport and the watchman, and the site staff whose duties cover the demobilisation week (Section 4, roles). A: the programme as submitted, 22-Aug to 24-Dec-2026. B: the same erection dates with the sequential testing carried, conditional on the early Tank 1 test (P18). The site engineer's 3 close-out days at 'Build-Up' line 1.22 fall after this period and are intentionally off-site visits, not a second allowance", 'cd')
d2 = add3('D2', 'Works period - mobilisation start to completion of testing and commissioning (calendar days)', f"=C{p1}", f"=D{p22}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p22}", f"=G{r}-F{r}+1",
          "The daytime 100 kVA works generator runs for this period; no works power is needed during demobilisation, when the welfare generator alone continues", 'cd')
d3 = add3('D3', 'Erection window - Tank 1 base panels start to Tank 2 mechanical finish (calendar days)', f"=C{p3}", f"=D{p14}", f"=D{r}-C{r}+1", *same(r),
          "Power tools and lighting towers are hired by the month while erection is in progress, 12-Sep to 03-Dec-2026 as submitted; the same on both bases", 'cd')
d6 = add3('D6', 'Access equipment window - Tank 1 walls start to Tank 2 roof finish (calendar days)', f"=C{p4}", f"=D{p13}", f"=D{r}-C{r}+1", *same(r),
          "Mobile access towers and podium steps are hired by the month for the period the walls, bracing and roofs are worked on; rounded up to whole months at 'Build-Up' lines 5.4 and 5.6", 'cd')
d7 = add3('D7', 'Storage containers window - first panels on site to last roof panel installed (calendar days)', f"={AS('QCD18TSEPRC1120')}", f"=D{p13}", f"=D{r}-C{r}+1", *same(r),
          "First delivery 08-Sep-2026 to the Tank 2 roof finish 30-Nov-2026. Panels have to be stored from the first delivery whatever the erection pace; storage is priced by the container-month, so the delivery staging adds no separate cost here", 'cd')
# --- sensitivity block
if PG['used'] + 19.5 + 70 > PAGE:
    pg_break(pg, r); header(pg, r, HDR3, merge_ij=True); r += 1; pg_add(32)
pg_add(19.5)
pg.cell(r, 2, 'Sensitivity - not carried; same durations and dependencies as P18 to P23'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5
r += 1
def sens(ref, name, cs, cf, cd, deriv, height=None):
    return add3(ref, name, cs, cf, cd, '-', '-', '-', deriv, 'cd', height)
b_t1t_f = W(NEXT(f'G{p16}'), 7); b_xf_f = W(NEXT(b_t1t_f), 3); b_t2t_f = W(NEXT(b_xf_f), 4); b_cc_f = W(NEXT(f'G{p16}'), 3)
b_ic_f = W(NEXT(f'MAX(G{p16},{b_t2t_f},{b_cc_f})'), 3); b_dm_f = W(NEXT(b_ic_f), 7)
s2 = sens('S2', 'Carried basis if the Tank 1 test cannot precede the tie-in - Tank 1 test after the tie-in, transfer, Tank 2 test, integrated commissioning, demobilisation (calendar days from mobilisation)', f"=F{p1}", f"={b_dm_f}", f"=D{r}-C{r}+1",
          "Tank 1 test 09 to 16-Dec-2026; transfer 17 to 20-Dec; Tank 2 test 21 to 24-Dec; component checks 09 to 12-Dec in parallel; integrated commissioning 26 to 28-Dec-2026; demobilisation 29-Dec-2026 to 05-Jan-2027 (Fridays excluded). The additional site days against D1 are shown at S3", 60)
s3 = add3('S3', 'S2 - complete incremental amount, excluding Overhead and Profit (first column: additional site days; third column: SAR)', f"=E{s2}-H{d1}", '-', 'SENS_AMT', '-', '-', '-',
          "Every Section 5 line re-evaluated with the S2 dates by its own formula and rounding: the site-period lines and site staff ('Build-Up' 1.1 to 1.8, 1.10, 1.12, 1.13, 5.8 and 5.9) at the S2 days in months; welfare-water deliveries (1.9) at two per rounded-up week; the works generator (5.7) to the S2 end of commissioning. The erection-window, plant, helper, access, container and Item 7 lines do not move because only the testing tail moves. Not carried: it applies only if the Engineer does not accept the early Tank 1 test", 60)
pg.cell(s3, 3).number_format = '#,##0'; pg.cell(s3, 5).number_format = '#,##0.00'
s4 = add3('S4', 'S2 - Overhead and Profit at 5 per cent on S3 (first column) and incremental amount including it (third column), SAR', f"=ROUND(E{s3}*0.05,2)", '-', '-', '-', '-', '-',
          "Applied once at the rate on the 'Assessment' tab. Nothing in S2 to S4 is carried into the assessment", 30)
pg.cell(s4, 3).number_format = '#,##0.00'; pg.cell(s4, 5).number_format = '#,##0.00'; pg.cell(s4, 5).value = f"=E{s3}+C{s4}"
SEC3_END = r - 1

# ================================================================ Section 4: resource bridge on the approved histogram
r += 1
pg_break(pg, r)
pg_add(21 + 3 * 60)
banner(pg, r, "4. RESOURCE BRIDGE - THE APPROVED MANPOWER HISTOGRAM RECONCILED TO THE PRICED LINES"); r += 1
para(pg, r, ("Approved baseline. The Contractor's weekly direct-manpower histogram (workbook 'SAMA-D18 -THS-TSE - Cost S-Curves Cash Flow Manpower Histogram.xlsx', "
             "sheet 'MP-HST-WK': 221 man-weeks, 136 skilled and 85 helpers, peak 35 in the week ending 30-Oct-2026) has been approved and is taken as the "
             "resource baseline. Its weekly figures are reproduced below unchanged. Approval of the histogram does not identify whose people they are, "
             "approve any rate, or make the labour payable by itself, and it does not extend to the programme, which remains under approval."), height=57); r += 1
para(pg, r, ("Reading the figures. The source sheet gives 'Total', 'Skilled Labor' and 'Helper' per week without saying whether each is the average or the peak "
             "daily deployment; it is read here as the number of people deployed through the week (an assumption - if it is a peak-day count the man-days are "
             "overstated). A man-week is one person for the working days of that week on the 'QIC 6 Days 10 Hrs' calendar: 6 days, or 5 in the week "
             "containing 23-Sep-2026; man-days are people multiplied by that count, not a flat 6. Days are counted inclusively (finish minus start plus one) "
             "and months at 30.4 calendar days, rounded to one decimal place."), height=57); r += 1
para(pg, r, ("Skilled people. The histogram does not say who employs them. They are allocated below to the activities in progress in each week - dismantling "
             "(Item 3 quotation, recorded scope), tank erection (Item 6, recorded supply-and-install scope), pipework and tie-in (Item 8 supplied-and-installed "
             "rates) - and none is priced again, because each of those scopes already carries its own labour. Where a week's skilled people could belong to "
             "more than one scope, or to the Contractor's own mobilisation, the allocation is marked unresolved; the cost effect if they prove to be the "
             "Contractor's own people outside those scopes is bounded at S5 ('Build-Up Comparison' tab, Section 7, item 10)."), height=57); r += 1
para(pg, r, ("Helpers. Every approved helper man-day is priced once at 'Build-Up' line 5.10, provisionally including the dismantling weeks, where the helpers "
             "may instead be within the subcontractor's crew (S7). Duties that fall inside the approved weeks and within the approved capacity are not "
             "priced again (transfer labour 7.4); duties after the histogram ends on 11-Dec-2026, or that do not fit the approved capacity, are priced on "
             "their own lines (Tank 2 disinfection at 7.8, clean-up at 1.16). The dated check is in the table after the weekly figures. Specialists "
             "(commissioning engineer, technicians, electrician, calibration, tie-in fitters, pipework testing crew) are not helpers and are priced in Items 7 "
             "and 8. No productivity or crew-sharing adjustment is made to the approved figures; whether each week's helpers were fully occupied cannot be "
             "verified from the documents received and is left unresolved."), height=83); r += 1
HDRB = ['Ref', 'Week ending (histogram week) - phase', 'Erection activities in progress (XER)', 'Working days in week', 'Approved total', 'Approved skilled', 'Approved helpers', 'Helper man-days (people x days)', "Helper man-days priced at 5.10", 'Skilled people - where already paid; helper duties; adjustment']
def hdrb(ws, r_):
    header(ws, r_, HDRB[:9]); ws.cell(r_, 10, HDRB[9]); cp(S_HDR, ws.cell(r_, 10))
hdrb(pg, r); r += 1
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
 ('18-Dec', 0, 0, 0, 'Demobilisation (carried basis from 16-Dec)', 'No approved labour: Tank 2 inspection 12-Dec, disinfection 13 to 17-Dec (7.8) and commissioning specialists (7.13 to 7.17) priced on their own lines'),
 ('25-Dec', 0, 0, 0, 'Demobilisation ends 23-Dec (carried)', 'No approved labour: demobilisation clean-up is line 1.16'),
]
b_first = r
wk_end = _dt.date(2026, 8, 28)
for i, (lab, tot_, sk, hp, phase, note) in enumerate(HISTW):
    h = max(30, est(note, 52), est(phase, 40))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); hdrb(pg, r); r += 1; pg_add(32)
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
    pg.cell(r, 10, note); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = h
    r += 1; wk_end += _dt.timedelta(7)
b_last = r - 1
def btot(label, cells, note, h=None):
    global r
    if PG['used'] + 19.5 > PAGE:
        pg_break(pg, r); hdrb(pg, r); r += 1; pg_add(32)
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
bt2 = btot("Helper man-months at 26 working days - 'Build-Up' line 5.10 quantity", {9: f"=ROUND(I{bt1}/26,1)"}, "Conversion for the man-month rate on the 'Build-Up' tab (6-day week)")
# --- dated helper duties against approved capacity, W13 to W18
r += 1
if PG['used'] + 32 + 9 * 32 > PAGE:
    pg_break(pg, r)
HDRD = ['Ref', 'Helper duty (dated from the XER activities and the fitted tests)', 'From', 'To', 'People', 'Man-days', 'Approved helper man-days in the same week(s)', 'Treatment', '', 'Basis']
def hdrd(ws, r_):
    header(ws, r_, HDRD[:9]); ws.cell(r_, 10, HDRD[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=8, end_row=r_, end_column=9)
hdrd(pg, r); pg_add(32); r += 1
DUTY = {}
def dadd(ref, duty, frm, to, ppl, md, cap, treat, basis):
    global r
    hh = max(30, est(duty, 46), est(basis, 52), est(treat, 26))
    if PG['used'] + hh > PAGE:
        pg_break(pg, r); hdrd(pg, r); r += 1; pg_add(32)
    pg_add(hh)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, duty); cp(S_DESC, pg.cell(r, 2))
    for col, v in ((3, frm), (4, to)):
        c = pg.cell(r, col, v); datecell(c)
        if v in (None, ''): c.value = '-'
    for col, v in ((5, ppl), (6, md), (7, cap)):
        c = pg.cell(r, col, v); numcell(c, '#,##0')
        if v in (None, ''): c.value = '-'
    pg.cell(r, 8, treat); cp(S_BASIS, pg.cell(r, 8)); cp(S_BASIS, pg.cell(r, 9)); pg.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
    pg.cell(r, 10, basis); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = hh; DUTY[ref] = r; r += 1
    return r - 1
wk = {i + 1: b_first + i for i in range(len(HISTW))}
dadd('H1', 'Tank 2 roof supports and ladder - gang attendance', f"={AS('QCD18TSECONT2INS1040')}", f"={AF('QCD18TSECONT2INS1040')}", 3, f"=E{r}*{AWD('QCD18TSECONT2INS1040')}", f"=H{wk[13]}", 'Within the approved week W13 (36 man-days)', 'Gang of 3 on supports and ladder - assessed; approved W13 helpers 6')
dadd('H2', 'Tank 1 hydrostatic test - tanker and level attendance during the fill and hold (P18)', f"=F{p18}", f"=G{p18}", 2, f"=E{r}*H{p18}", f"=H{wk[13]}+H{wk[14]}", 'Within the approved weeks W13 and W14', 'Assessed 2 people; specialists (QA/QC 1.3, commissioning engineer 7.13) are not helpers')
dadd('H3', 'Tank 2 roof panels - gang attendance', f"={AS('QCD18TSECONT2INS1050')}", f"={AF('QCD18TSECONT2INS1050')}", 5, f"=E{r}*{AWD('QCD18TSECONT2INS1050')}", f"=H{wk[14]}", 'Within the approved week W14 (30 man-days) - fully used', "Gang of 5 - the supplier's condition read at its upper range for roof panels")
dadd('H4', 'Tank 1 disinfection and flushing after its test (AWWA C652), before the transfer', f"={NEXT(f'G{p18}')}", f"={W(NEXT(f'G{p18}'), 5)}", 2, f"=E{r}*5", f"=H{wk[14]}+H{wk[15]}", 'W14 is fully used by H3; fits within W15 (18 man-days) alongside H5', "Assessed 2 people for 5 working days; the chemicals are 7.6 and 7.9")
dadd('H5', 'Tank 2 nozzles and internals - attendance', f"={AS('QCD18TSECONT2MW2030')}", f"={AF('QCD18TSECONT2MW2030')}", 2, f"=E{r}*{AWD('QCD18TSECONT2MW2030')}", f"=H{wk[15]}", 'Within the approved week W15', 'Assessed 2 people')
dadd('H6', 'Transfer of the test water, Tank 1 to Tank 2 (P19) - pump and hose attendance', f"=F{p19}", f"=G{p19}", 2, f"=E{r}*H{p19}", f"=H{wk[16]}", "Within the approved week W16 (12 man-days) - so 'Build-Up' line 7.4 is nil", 'Assessed 2 people for the transfer days')
dadd('H7', 'Tank 2 top-up, hold and inspection attendance (P20) to 11-Dec', f"=F{p20}", f"=MIN(G{p20},DATE(2026,12,11))", 1, f"=E{r}*{WD(f'C{r}', f'D{r}')}", f"=H{wk[16]}", 'Within the approved week W16 with H6 (6 + 4 = 10 of 12)', 'Assessed 1 person')
dadd('H8', 'Tank 2 inspection on 12-Dec and disinfection and flushing 13 to 17-Dec (after the histogram ends 11-Dec)', f"=DATE(2026,12,12)", f"={W('DATE(2026,12,13)', 5)}", 2, f"=E{r}*5", 0, "Outside the approved weeks: priced at 'Build-Up' line 7.8 (10 man-days)", 'Assessed 2 people for 5 working days, in parallel with the integrated commissioning P21b')
dadd('H9', 'Integrated commissioning 13 to 15-Dec and demobilisation 16 to 23-Dec', f"=F{p21b}", f"=G{p23}", '-', '-', 0, "Specialists at 7.13 to 7.17 (engineer, technicians, electrician, calibration); clean-up 4 x 3 days at 1.16; no helpers", 'Outside the approved weeks; nothing added beyond the existing lines')
btd = btot("Helper man-days needed from the fitted tests and the XER activities, W13 to W18, against the approved 114 in W13 to W16", {6: f"=SUM(F{DUTY['H1']}:F{DUTY['H8']})"}, "Weeks W1 to W12 (offloading and panel handling on one then two fronts) cannot be checked duty by duty from the documents received - the panel counts, loads and crew method are not stated - and are left unresolved; the approved figures are carried there without adjustment")
# --- skilled people: allocation by weeks and activities in progress
r += 1
if PG['used'] + 32 + 6 * 40 > PAGE:
    pg_break(pg, r)
HDRK = ['Ref', 'Skilled people - weeks and activities in progress', 'Weeks', '', 'Man-weeks', 'Allocated to', '', '', '', 'Status of the allocation']
def hdrk(ws, r_):
    header(ws, r_, HDRK[:9]); ws.cell(r_, 10, HDRK[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=3, end_row=r_, end_column=4); ws.merge_cells(start_row=r_, start_column=6, end_row=r_, end_column=9)
hdrk(pg, r); pg_add(32); r += 1
SK = {}
def kadd(ref, grp, weeks, mw, alloc, status):
    global r
    hh = max(30, est(grp, 46), est(alloc, 52), est(status, 52))
    if PG['used'] + hh > PAGE:
        pg_break(pg, r); hdrk(pg, r); r += 1; pg_add(32)
    pg_add(hh)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, grp); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, weeks); cp(S_UNIT, pg.cell(r, 3)); cp(S_UNIT, pg.cell(r, 4)); pg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    pg.cell(r, 5, mw); numcell(pg.cell(r, 5), '#,##0')
    pg.cell(r, 6, alloc); cp(S_BASIS, pg.cell(r, 6))
    for c in (7, 8, 9): cp(S_BASIS, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=6, end_row=r, end_column=9)
    pg.cell(r, 10, status); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = hh; SK[ref] = r; r += 1
kadd('K1', 'Mobilisation (22 to 26-Aug) and dismantling start (27-Aug)', 'W1', f"=F{wk[1]}", "Contractor's mobilisation riggers, or the Item 3 dismantling crew", 'Unresolved - split unknown; if Contractor\'s own mobilisation labour, it is not in any priced line (S5 bounds it)')
kadd('K2', 'Dismantling in progress (QCD18TSECONDSM1020 to 1030)', 'W2 to W3', f"=F{wk[2]}+F{wk[3]}", 'Dismantling crew within the Al Mousa quotation S04647 (Item 3, recorded scope)', 'Provisional on the recorded scope; the quotation is not attached')
kadd('K3', 'Tank 1 erection only (QCD18TSECONT1INS1020 to 1060)', 'W4 to W9', f"=SUM(F{wk[4]}:F{wk[9]})", "Tank erection crew within the supplier's supply-and-install price (Item 6, recorded scope)", 'Provisional on the recorded scope; the offer is not attached')
kadd('K4', 'Two erection fronts (Tank 1 bracing, roof; Tank 2 base, walls, bracing)', 'W10 to W11', f"=F{wk[10]}+F{wk[11]}", 'Two erection crews within Item 6', 'As K3')
kadd('K5', 'Erection, mechanical works and pipework installation overlapping (QCD18TSECONT1MW2055 from 12-Nov; QCD18TSECONT2MW2020)', 'W12 to W15', f"=SUM(F{wk[12]}:F{wk[15]})", 'Erection crews (Item 6) and pipework fitters (Item 8 supplied-and-installed rates)', 'Split between Items 6 and 8 unresolved; both scopes carry their own labour, so no separate price either way')
kadd('K6', 'Tie-in, transfer and tests (QCD18TSECONTC2040, P19, P20)', 'W16', f"=F{wk[16]}", 'Tie-in fitters within 8.14 and 8.15; test supervision within Item 6', 'As K5')
btk = btot("Skilled man-weeks allocated (equals the approved 136); none priced separately", {5: f"=SUM(E{SK['K1']}:E{SK['K6']})"}, "If any group proves to be the Contractor's own people outside the recorded scopes, the effect is bounded at S5 and Item 6 would need re-basing to a supply-only price")
# --- plant, formula-linked to the XER activity IDs
r += 1
if PG['used'] + 32 + 4 * 45 > PAGE:
    pg_break(pg, r)
HDRP = ['Ref', 'Plant and basis', 'From (XER)', 'To (XER)', 'Working days', 'Hire days carried', '', '', '', 'Derivation from the XER activities; priced item']
def hdrp(ws, r_):
    header(ws, r_, HDRP[:9]); ws.cell(r_, 10, HDRP[9]); cp(S_HDR, ws.cell(r_, 10)); ws.merge_cells(start_row=r_, start_column=7, end_row=r_, end_column=9)
hdrp(pg, r); pg_add(32); r += 1
PL = {}
def padd(ref, name, frm, to, wd, hire, note, h=None):
    global r
    hh = max(30, est(name, 46), est(note, 52), h or 0)
    if PG['used'] + hh > PAGE:
        pg_break(pg, r); hdrp(pg, r); r += 1; pg_add(32)
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
header(pg, r, HDRR[:9]); pg.cell(r, 10, HDRR[9]); cp(S_HDR, pg.cell(r, 10)); pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9); pg_add(32); r += 1
ROLE = {}
def radd(ref, role, frm, to, src):
    global r
    h = max(30, est(role, 46), est(src, 52))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); header(pg, r, HDRR[:9]); pg.cell(r, 10, HDRR[9]); cp(S_HDR, pg.cell(r, 10)); r += 1; pg_add(32)
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
banner(pg, r, "5. 'BUILD-UP' LINES WHOSE QUANTITY IS TAKEN FROM THIS TAB - THE PROGRAMME AS SUBMITTED BESIDE THE BASIS CARRIED"); r += 1
para(pg, r, ("Rates are those on the 'Build-Up' tab. Column G (carried) is linked into column D of the 'Build-Up' tab; column F (Basis A) shows what the "
             "same rate would give on the programme exactly as submitted - parallel testing, every submitted day paid, the Contractor's own helper "
             "histogram - and is a comparison, not a target. Lines not listed here are unchanged from the previous revision because the programme does "
             "not inform them (mobilisation trips, design resources, scaffold area, consumables and pipework quantities); their basis is stated on the "
             "'Build-Up' tab."), height=44); r += 1
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
    L(ref, row_, f"={MO(sp_A)}", f"=F{ROLE[role]}", f"Section 4, role {role}: phase coverage stated there. A: the whole submitted site period")
for ref, row_ in (('1.6', 14), ('1.7', 15), ('1.8', 16)):
    L(ref, row_, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.9', 17, f"=2*ROUNDUP({sp_A}/7,0)", f"=2*ROUNDUP({sp_B}/7,0)", "Two 10 m3 deliveries a week for the weeks of the site period D1, rounded up to whole weeks")
L('1.10', 18, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.11', 19, f"=2*{MO(ct_A)}", f"=2*{MO(ct_B)}", "2 No. containers for the storage window D7 in months")
L('1.12', 20, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months, transport being needed while the workforce demobilises; the manpower histogram shows labour on site for 16 weeks, 28-Aug to 11-Dec-2026, inside this period")
L('1.13', 21, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months; still subject to whether the Employer's security covers the lower plateau")
L('5.1', 74, f"=F{pbt}", f"=F{pbt}", "Section 4, boom truck hire days B5, built from the XER activity dates; the same on A")
L('5.2', 75, f"=F{plt}", f"=F{plt}", "Section 4, telehandler hire days T3, built from the XER activity dates; the same on A")
L('5.4', 77, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "2 No. towers for the access window D6, rounded up to whole hire months")
L('5.6', 79, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "As 5.4")
L('5.7', 80, f"={MO(wm_A)}", f"={MO(wm_B)}", "Works period D2 in months - daytime works power to completion of commissioning; none during demobilisation")
L('5.8', 81, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months - welfare power runs continuously to the end of demobilisation")
L('5.9', 82, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months")
L('5.10', 83, f"=I{bt2}", f"=I{bt2}", "Section 4, the approved histogram helper man-days converted at 26 working days a month; the same on A")
L('5.11', 84, f"={MO(er_A)}", f"={MO(er_B)}", "Erection window D3 in months")
L('5.14', 87, f"=2*{MO(er_A)}", f"=2*{MO(er_B)}", "2 No. lighting towers for the erection window D3 in months")
L('7.1', 100, "=2*3774", "=3774", "A: both tanks filled at once for parallel testing (7,548 m3, RFP Scope of Works section 1). Carried: one fill, water re-used for the second tank - the Engineer's (KEO) email of 30-Aug-2026")
L('7.2', 101, "=0", f"=H{p19}+H{p20}", "Transfer pump: not needed on A; carried for the transfer P19 and the Tank 2 test P20 (top-up and hold)")
L('7.3', 102, "=0", "=1", "Transfer hoses: not needed on A; 1 week carried")
L('7.4', 103, "=0", "=0", "Nil on both bases: the transfer labour is within the approved histogram helpers priced at 5.10 (Section 4, week W16)")
L('7.5', 104, "=ROUND(2*3774*0.1,0)", "=ROUND(3774*0.1,0)", "Top-up at 10 per cent of the water filled: of two fills on A, of one fill carried (retention in Tank 1 between P18 and P19)")
L('7.13', 112, f"=E{p17}+1", f"=H{p21}+H{p21b}+2", "A: the programmed 7-working-day parallel testing and commissioning activity plus one day at the hold. Carried: component checks P21 (3 days) and integrated commissioning P21b (3 days), both assessed assumptions, plus one day at each hydrostatic test hold; the fills and holds themselves are supervised by the QA/QC inspector (line 1.3) and the supplier's leak-test supervision within Item 6")
L('7.14', 113, f"=2*E{p17}", f"=2*(H{p21}+H{p21b})", "2 No. technicians: A for the 7-day programmed activity; carried for the component checks P21 and the integrated commissioning P21b (6 days)")
L('7.16', 115, f"=E{p16}+1", f"=H{p16}+1", "Tie-in P16 working days plus one day of integrated commissioning, both bases")
L('7.22', 121, "=4", "=4", "Pump and hose set standing by through the two 24-hour holds and two days of contingency between the Tank 1 test and the transfer; tankers are working, not standing by, during the fill. The same on A, where two simultaneous fills need the same standby")

first4 = r
item_rows = {}
for ref, burow, qA, qB, deriv, prev in LINES:
    item = ref.split('.')[0]
    if item not in item_rows:
        item_rows[item] = []
    h4 = max(30, est(deriv, 52), est(str(bu.cell(burow, 2).value), 46))
    if PG['used'] + h4 > PAGE:
        pg_break(pg, r); hdr4(pg, r); r += 1; pg_add(32)
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
    pg_break(pg, r); hdr4(pg, r); r += 1; pg_add(32)
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
pg.cell(r, 10, "The carried basis is the 'Assessment' tab total (row 21). Basis A replaces the listed lines with their Basis A amounts and applies the Overhead and Profit once, as on the 'Assessment' tab"); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 30
TOT4 = r; r += 1
for c in (1, 3, 4, 5, 6, 7, 10): pass
pg.cell(r, 2, "Sensitivities on the resource matrix - not carried"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5; r += 1
pg.cell(r, 1, 'S5'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, "Upper bound if all 136 approved skilled man-weeks were the Contractor's own people outside the recorded Item 3, 6 and 8 scopes (gross labour, before any re-basing of Item 6 to a supply-only price)"); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'man-day'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E50"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, f"=E{btk}*6"); numcell(pg.cell(r, 7), '#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9))
pg.cell(r, 10, "At the rigger day rate assessed at 'Build-Up' c.2; a bound only, not carried: the Contractor's answer to 'Build-Up Comparison' Section 7, item 10 decides, and Item 6 would fall if it is supply only"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 44; r += 1
pg.cell(r, 1, 'S7'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, "Reduction if the helpers approved in the dismantling weeks W1 to W3 (13 man-weeks) prove to be within the Item 3 dismantling crew rather than the Contractor's own attendance"); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'man-month'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E83"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, f"=-ROUND((H{wk[1]}+H{wk[2]}+H{wk[3]})/26,1)"); numcell(pg.cell(r, 7), '#,##0.00;-#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9)); pg.cell(r, 9).number_format = '#,##0.00;-#,##0.00'
pg.cell(r, 10, "Excluding Overhead and Profit; not carried. Those helpers are priced provisionally at 5.10 because the approved histogram lists them as the Contractor's direct manpower; the Al Mousa quotation scope (not attached) decides"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 44; r += 1
pg.cell(r, 1, 'S6'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, 'A second telehandler for the 11 concurrent days (Tank 1 roof with Tank 2 walls, 29-Oct to 11-Nov-2026), if one unit proves insufficient'); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'day'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E75"); cp(S_RATE, pg.cell(r, 4))
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
 ('3.1', 46, 'QCD18TSECONDSM1020 to 1030 (27-Aug to 12-Sep-2026, complete)', 'Dismantling, segregation and loading of the damaged tank; Al Mousa under quotation S04647 for the Contractor', 'Fixed: lowest of three quotations, recorded scope', 'No - haulage is 4.1; crane and crew within the quotation on its recorded scope (offer not attached), so no plant at Item 5; matrix M2 carries Contractor attendance only'),
 ('c.1 to c.14', 49, 'First-principles check of 3.1 - information only', 'Not carried', 'Not carried', 'Not carried'),
 ('4.1', 68, 'After QCD18TSECONDSM1030 - handover to the Employer', "Haulage of the dismantled materials to the Employer's local handover point; Contractor", 'Quantity: 14 loads, assessed; location unconfirmed', 'No - loading is within 3.1'),
 ('5.1', 74, 'Section 4 plant rows B1 to B5 (QCD18TSECONT1INS1030 to QCD18TSECONT2INS1050)', 'Boom truck for offloading and lifting wall tiers, bracing, roof supports and roof panels; Contractor for the supplier', 'Hire: working days on site, one unit', "No - the supplier's own lifting is excluded from its offer"),
 ('5.2', 75, 'Section 4 plant rows T1 to T3 (QCD18TSEPRC1120 to QCD18TSECONT2INS1050)', 'Telehandler moving pallets from the trucks to storage and into the tank footprint; Contractor for the supplier', 'Hire: working days on site, one unit, provisional', 'No'),
 ('5.3', 76, 'QCD18TSECONT1INS1030 to QCD18TSECONT2INS1050 (window D6)', 'Perimeter scaffold and edge protection for wall, bracing and roof work; Contractor (supplier condition)', 'Quantity: 1,024 m2 supplied, erected, 3-month hire, dismantled', 'No'),
 ('5.4', 77, 'Window D6', 'Mobile access towers; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.5', 78, 'Window D6', 'Relocation and re-inspection of the towers; Contractor', 'Quantity: 4 No.', 'No'),
 ('5.6', 79, 'Window D6', 'Podium steps; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.7', 80, 'QCD18TSEMOB1240 to QCD18TSECONTCT22030 (window D2)', "Daytime works power for the supplier's tools and the Contractor's works; Contractor (supplier condition)", 'Hire with fuel: monthly, works period', 'No - welfare power is 5.8'),
 ('5.8', 81, SITE, 'Continuous welfare power; Contractor', 'Hire with fuel: monthly, site period', 'No'),
 ('5.9', 82, SITE, 'Site pick-up; Contractor', 'Hire: monthly, site period', 'No - workforce bus is 1.12; plant is 5.1 and 5.2'),
 ('5.10', 83, 'Section 4 weekly bridge W1 to W16 (approved histogram)', "Contractor's helpers for every approved week: mobilisation, dismantling attendance, offloading, panel handling, transfer and disinfection", 'Attendance: approved helper man-days converted to man-months', 'No - erectors within Item 6; 7.4 nil; 7.8 Tank 2 only; clean-up 1.16'),
 ('5.11', 84, 'QCD18TSECONT1INS1020 to QCD18TSECONT2MW2020 (window D3)', "Power tools for the Contractor's own works; Contractor", 'Hire: monthly, erection window', "No - the supplier's erection tools are within Item 6"),
 ('5.12', 85, 'Within Item 6 on its recorded scope (RFP work package 2)', 'Sealant application - supplier erection work', 'Provisionally not assessed', 'Yes - Item 6, to confirm'),
 ('5.13', 86, 'Within Item 6 on its recorded scope', 'Fixings and touch-up - supplier supply and erection', 'Provisionally not assessed', 'Yes - Item 6, to confirm'),
 ('5.14', 87, 'Window D3', 'Lighting towers for the erection fronts and the work area at dusk; Contractor', 'Hire: 2 No., monthly, erection window', 'No'),
 ('5.15', 88, 'Site service - power distribution from 5.7 and 5.8', 'Distribution boards and cabling; Contractor', 'Quantity: 2 sets', 'No'),
 ('6.1', 94, 'QCD18TSEPRC1060 to QCD18TSEPRC1230, QCD18TSECONT1INS1020 to QCD18TSECONT2MW2030', 'Design, manufacture, delivery duty paid, erection, sealing, bracing, nozzles and internals of both tanks; Al Mousa / Stalwart', 'Fixed: quotation per tank, insulated, provisional; erection scope as recorded, offer not attached', 'No - the Contractor provides offloading, scaffold, storage, power and helpers (Items 1 and 5)'),
 ('7.1', 100, 'P18 (17 to 24-Nov-2026 carried)', 'Tankered water for the first fill; Contractor', 'Quantity: 3,774 m3, one fill', 'No'),
 ('7.2', 101, 'P19 and P20', 'Transfer pump; Contractor', 'Hire: transfer plus test days', 'No'),
 ('7.3', 102, 'P19', 'Transfer hoses and fittings; Contractor', 'Hire: 1 week', 'No'),
 ('7.4', 103, 'P19 - within the approved histogram week W16', 'Pump attendance during the transfer; Contractor helpers', 'Nil - priced at 5.10', 'Yes - 5.10'),
 ('7.5', 104, 'P18 to P20', 'Top-up for losses and test level; Contractor', 'Quantity: 10 per cent of one fill', 'No'),
 ('7.6', 105, 'P18 and P20 (AWWA C652)', 'Disinfection chemicals; Contractor', 'Quantity: 800 kg', 'No'),
 ('7.7', 106, 'P18 and P20', 'Dosing equipment; Contractor', 'Hire: 2 weeks', 'No'),
 ('7.8', 107, 'Duties H4 (Tank 1, within W15) and H8 (Tank 2, 12 to 17-Dec, after the histogram)', 'Disinfection and flushing labour; Contractor helpers', 'Attendance: 2 No. x 5 days for Tank 2 only', 'Tank 1 part within 5.10'),
 ('7.9', 108, 'After P20', 'Dechlorination for discharge; Contractor', 'Quantity: 400 kg', 'No'),
 ('7.10', 109, 'P18 and P20', 'Sampling and transport; Contractor', 'Quantity: 2 tanks', 'No'),
 ('7.11', 110, 'P18 and P20', 'Laboratory water-quality tests; third-party laboratory', 'Quantity: 6 samples', "No - excluded by both tank suppliers"),
 ('7.12', 111, 'None - no third-party inspection in the RFP', 'Not assessed', 'Not assessed', 'Engineer witness only (RFP 5.1, 5.2)'),
 ('7.13', 112, 'P18, P20, P21, P21b', 'Commissioning engineer at the test holds, the component checks and the integrated commissioning; Contractor', 'Attendance: 8 days', "No - the supplier's leak-test supervision is within Item 6"),
 ('7.14', 113, 'P21 and P21b', 'MEP technicians on the pump and network interfaces; Contractor', 'Attendance: 2 No. x 6 days', 'No'),
 ('7.15', 114, 'P21 (after QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040)', 'Electrical connection and readout checks of the level instruments; Contractor', 'Attendance: 5 days', 'No - the cabling is supplied and installed at 8.11'),
 ('7.16', 115, 'P16 and P21b', 'Coordination of the tie-in and the witnessed demonstration with the Employer and Engineer; Contractor', 'Attendance: tie-in days plus 1', 'No'),
 ('7.17', 116, 'P21', 'Calibration of the instruments; technician', 'Attendance: 4 days', 'No'),
 ('7.18', 117, 'P21', 'Certified calibrations; laboratory', 'Quantity: 6 No.', 'No'),
 ('7.19', 118, 'P18 to P21b', 'Calibrated gauges and data logger; Contractor', 'Hire: 2 weeks', 'No'),
 ('7.20', 119, 'P18 to P23', 'Test records and ITP / WIR close-out; Contractor clerk', 'Attendance: 1 month', 'No - the document controller at 1.23 collates the close-out file'),
 ('7.21', 120, 'QCD18TSEDMOB3020', 'As-built mark-ups; Contractor draughtsman', 'Attendance: 0.5 month', 'No - shop drawings are 2.2'),
 ('7.22', 121, 'P18 and P20 holds', 'Pump and hose set standing by through the holds; Contractor', 'Hire: 4 days', 'No - tankers are paid at 7.1'),
 ('7.23', 122, 'Deliverables list 26 to 28-Sep-2026 (PQD, ITP, procedure, inspector CV)', 'Third-party factory acceptance test - not required by the RFP', 'Nil pending evidence', 'Manufacturer test reports are within Item 6'),
 ('8.1', 129, 'QCD18TSEPRC1160 to 1250 (procurement), QCD18TSECONT1MW2055 and QCD18TSECONT2MW2020 (installation)', 'Main pipework above DN150 to the RFP specification, supplied and installed; Contractor', 'Quantity: 200 m assessed, take-off pending', 'No'),
 ('8.2', 130, 'As 8.1', 'Small-bore uPVC pipework; Contractor', 'Quantity: 60 m assessed', 'No'),
 ('8.3', 131, 'As 8.1', 'Butterfly isolation valves to the RFP; Contractor', 'Quantity: 10 No. assessed', 'No'),
 ('8.4', 132, 'As 8.1', 'Dismantling joints; Contractor', 'Quantity: 4 No.', 'No'),
 ('8.5', 133, 'QCD18TSECONT1MW2050 and QCD18TSECONT2MW2030', 'Blind flanges on spare nozzles; Contractor', 'Quantity: 8 No.', "No - the nozzles themselves are the supplier's"),
 ('8.6', 134, 'As 8.1', 'Flange sets, gaskets and bolting; Contractor', 'Quantity: 2 tank-sets', 'No'),
 ('8.7', 135, 'As 8.1', 'Pipe supports; Contractor', 'Quantity: 40 No.', 'No'),
 ('8.8', 136, 'As 8.1', 'Anchor and thrust blocks; Contractor', 'Quantity: 8 No.', 'No'),
 ('8.9', 137, 'QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040', 'Float-and-tape level indicators; Contractor', 'Quantity: 2 No.', "No - excluded from the tank supply"),
 ('8.10', 138, 'As 8.9', 'Level transmitters; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.11', 139, 'As 8.9', 'Instrument cabling and conduit; Contractor', 'Quantity: 160 m assessed', 'No - connection checks are 7.15'),
 ('8.12', 140, 'QCD18TSECONTC2040', 'Pressure gauge assemblies; Contractor', 'Quantity: 2 sets', 'No'),
 ('8.13', 141, 'QCD18TSECONTC2040', 'Sample taps; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.14', 142, 'QCD18TSECONTC2040 after QCD18TSECONIF2050 (03-Dec-2026)', 'Tie-ins to the networks after the external readiness milestone; Contractor', 'Quantity: 2 No.', 'No'),
 ('8.15', 143, 'QCD18TSECONTC2040', 'Tie-in coordination and out-of-hours working; Contractor', 'Quantity: 2 No.', 'No - the coordination engineer at 7.16 is the Contractor\'s attendance at the demonstration'),
 ('8.16', 144, 'After QCD18TSECONT2MW2020, before P21', 'Pipework test pump and manifold; Contractor', 'Hire: 2 weeks', 'No - tank testing is Item 7'),
 ('8.17', 145, 'As 8.16', 'Pipework test and flushing water; Contractor', 'Quantity: 300 m3', 'No'),
 ('8.18', 146, 'As 8.16', 'Pipework testing crew; Contractor', 'Attendance: 2 No. x 12 days', 'No - not in 5.10 (matrix M14)'),
 ('8.19', 147, 'As 8.16', 'Test records and certificates; Contractor', 'Quantity: 2 systems', 'No'),
 ('8.20', 148, 'As 8.1', 'Colour banding; Contractor', 'Quantity: 260 m', 'No'),
 ('8.21', 149, 'As 8.1', 'Tags, arrows and labels; Contractor', 'Quantity: 70 No.', 'No'),
 ('a.1 to a.5', 152, "The Contractor's procured materials - information only", 'Not carried', 'Not carried', 'Not carried'),
]
r += 1
pg_break(pg, r)
pg_add(21 + 60 + 32)
banner(pg, r, "6. LINK REGISTER - EVERY 'BUILD-UP' LINE TO ITS PROGRAMME ACTIVITIES OR SITE SERVICE"); r += 1
para(pg, r, ("One row for each line on the 'Build-Up' tab: the activities or the site service it serves, what the money buys and when the work can happen, who "
             "performs or provides it, the basis (fixed sum, measured quantity, attendance or hire), whether the same work is paid elsewhere, and the assessed "
             "amount linked from the 'Build-Up' tab. Lines that serve the whole site are shown as site services rather than tied to an activity they do not "
             "belong to. A resource that serves several activities is charged once, on the row where it is priced."), height=60); r += 1
HDRG = ['Ref', "'Build-Up' line (as described there)", 'Programme activities or site service', '', 'What it buys, when it can happen, who provides it', '', 'Basis: fixed, quantity, attendance or hire', '', 'Paid elsewhere?', 'Assessed (SAR)']
def hdrg(ws, r_):
    header(ws, r_, HDRG[:9]); ws.cell(r_, 10, HDRG[9]); cp(S_HDR, ws.cell(r_, 10))
    ws.merge_cells(start_row=r_, start_column=3, end_row=r_, end_column=4); ws.merge_cells(start_row=r_, start_column=5, end_row=r_, end_column=6); ws.merge_cells(start_row=r_, start_column=7, end_row=r_, end_column=8)
hdrg(pg, r); r += 1
g_first = r
for ref, burow, wbs, what, basis, elsewhere in REG:
    h = max(30, est(wbs, 24), est(what, 34), est(basis, 26), est(elsewhere, 14), est(str(bu.cell(burow, 2).value), 46))
    if PG['used'] + h > PAGE:
        pg_break(pg, r); hdrg(pg, r); r += 1; pg_add(32)
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
pg.cell(r, 2, "Total of the priced lines above, before Overhead and Profit"); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 6, 7, 8, 9): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 10, f"=SUM(J{g_first}:J{g_last})+'Build-Up'!F94"); cp(S_TOTAMT, pg.cell(r, 10))
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
bu['A4'] = ("How to read this tab: each item on the 'Assessment' tab whose rate is assessed rather than quoted is built up here as quantity x rate, "
            "with its basis beside it. Item totals carry to column J of the 'Assessment' tab. Quantities shown in column D as links to the "
            "'Programme' tab are derived from the Contractor's programme of 29-Sep-2026 on the basis carried there (Basis B); the same lines on "
            "the programme exactly as submitted (Basis A) are on the 'Programme' tab, Section 4. Rates marked 'assumption' are working assumptions "
            "pending the Contractor's substantiation. All amounts exclude VAT.")
bu.row_dimensions[4].height = 55
bu['A7'] = ("The tank site is in the lower plateau, outside the D-18 site boundary. Site management, welfare, water, security, insurance cover and "
            "close-out are therefore assessed as a separate establishment. Periods are taken from the Contractor's programme of 29-Sep-2026 "
            "('Programme' tab, Section 3): site staff, transport and facilities for the site period D1, mobilisation to the end of demobilisation, "
            "with the post-demobilisation completion inspection priced at 1.22 and 1.23. The programme is "
            "under the Engineer's approval and is not an agreed basis; the end of the site period is set by the external readiness milestone of "
            "03-Dec-2026 and the tie-in, sequential testing, commissioning and demobilisation that follow. Site staff are assessed role by role with their phase "
            "duties on the 'Programme' tab, Section 4. Design and "
            "engineering resources are shared with the main Contract (Item 2). Power is in Item 5. The tank supplier's price includes the "
            "installer's own mobilisation; the foundations and steel base frames are existing.")
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
bu['A72'] = ("Covers what the tank supplier's offer (Stalwart technical offer SS-07-26-1516 dated 03-Aug-2026, with Al Mousa quotation S04488) requires the "
             "Contractor to provide free of cost: offloading and hoisting, scaffolding, storage and handling, power, and 4-6 helpers; plus general site support. "
             "On the recorded scope of that offer (not attached to this revision; to be confirmed against the original) Item 6 is supply and installation, so "
             "GRP erectors are not priced here. Boundary: Item 1 is management, welfare and mobilisation; Item 5 is execution plant, access, power and "
             "supplier attendance. On the 'Programme' tab, Section 4, the helpers are the approved manpower histogram's helper row, priced once, and the "
             "plant days are built from the XER activity dates. One boom truck and one telehandler are a provisional "
             "utilisation assumption pending the Contractor's plant schedule. Power: the works generator by day to completion of commissioning, the welfare "
             "generator continuously to the end of demobilisation. The Engineer's (KEO) email of 30-Aug-2026: suitable Site power is unlikely, so a generator "
             "is allowed; the Contractor's methodology uses forklifts, cranes and pallet jacks.")
bu.row_dimensions[72].height = 96
gset(74, "Hire days B5 on the 'Programme' tab, Section 4, built from the XER activity dates: alternate days through each tank's walls-and-bracing window, every roof-support and roof-panel day, less the overlap of the two tanks, plus offloading days outside those. The Contractor's methodology names forklifts, cranes and pallet jacks (Engineer's (KEO) email of 30-Aug-2026); no plant schedule has been submitted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(75, "Hire days T3 on the 'Programme' tab, Section 4, built from the XER activity dates: continuous from the first delivery to the Tank 2 walls finish, plus the Tank 2 roof days; one unit is a provisional utilisation assumption. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(76, f"2 tanks x 128 m perimeter x 4.0 m height; the rate includes a 3-month hire, which covers the access window D6 on the '{SRC_A} on either basis. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B77'] = 'Mobile aluminium access towers, 2 No. for the access window, whole hire months'
gset(77, f"2 No. for the access window D6 on the '{SRC_A}, rounded up to whole hire months. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B79'] = 'Podium steps, 2 No. for the access window, whole hire months'
gset(79, "As 5.4. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(80, f"Works period D2 on the '{SRC_A}, to completion of commissioning. Power is excluded by both tank suppliers; no Site power expected at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(81, f"Site period D1 on the '{SRC_A}; air conditioning and lighting run around the clock at a separate site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(82, f"Site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B83'] = "Contractor's helpers - the approved manpower histogram helper row (85 man-weeks), man-months"
gset(83, "The approved manpower histogram's helper row, 85 man-weeks, as reproduced week by week on the 'Programme' tab, Section 4, converted to man-days on each week's working days and to man-months at 26 working days; every helper man-day priced once here, so the transfer labour 7.4 is nil and 7.8 carries only the Tank 2 disinfection after the histogram ends. The dismantling-week helpers are priced provisionally (S7). Skilled people are within Items 3, 6 and 8 on their recorded scopes and are not priced. The Contractor's histogram helper row (85 man-weeks) is reconciled week by week there; it is not adopted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(84, f"Erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B87'] = 'Mobile lighting towers, 2 No. for the erection window'
gset(87, f"2 No. for the erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['A98'] = ("The Engineer's (KEO) email of 30-Aug-2026: installation and testing will not be in parallel; one tank is filled and the water is re-used for "
             "the second by pumping across. This basis is carried. The Contractor's programme tests both tanks in parallel on 09 to 16-Dec-2026, which "
             "would need 7,548 m3 of water at once; that basis is shown beside this one on the 'Programme' tab, Section 4 (lines 7.1 to 7.5 and 7.22) "
             "and is not carried. The sequential test is fitted to the submitted dates ('Programme' tab, P18 to P22): Tank 1 is tested before the tie-in, "
             "within its programme float, on the conditional assumption of an Engineer-approved method with the water retained in Tank 1; the water is "
             "transferred once Tank 1 has passed and Tank 2 is ready; component checks follow the tie-in; integrated commissioning follows the last of "
             "the tie-in, the Tank 2 test and the component checks. Each test is 7 elapsed working days (fill or transfer, hold, inspection); the "
             "attendance priced is set out line by line. The Contractor arranges the water; the source is being checked by the Employer, so tankered supply "
             "is assumed. Excludes the leak-test supervision already in the supplier's price (Item 6) and any re-testing after an unsatisfactory "
             "test, which is the Contractor's obligation under RFP Scope of Works 5.1. External pipework testing is Item 8. Scope: RFP Scope of "
             "Works 5.1 to 5.3.")
bu.row_dimensions[98].height = 96
gset(100, "3,774 m3 effective at 3.7 m water level (RFP Scope of Works section 1); one fill, the water re-used for Tank 2. Tankered at SAR 6.00/m3 - assumption; falls away if a network fill is confirmed")
gset(101, "The transfer P19 and the Tank 2 test P20 on the 'Programme' tab (7 days); pump duty an assessed assumption pending the Contractor's method statement. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(103, "2 No. x 3 days, the transfer window P19 on the 'Programme' tab. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(104, "10 per cent of one fill, covering losses while the water is retained in Tank 1 between its test and the transfer to Tank 2 - assumption")
gset(112, "Component checks and integrated commissioning (P21 and P21b on the 'Programme' tab) plus one day at each hydrostatic test hold; reduced from 15 days because the fills and holds are supervised by the QA/QC inspector (1.3) and the supplier (Item 6). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(113, "2 No. for the component checks and the integrated commissioning (P21 and P21b on the 'Programme' tab); reduced from 2 No. x 10 days. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(115, "Tie-in connections (P16 on the 'Programme' tab, 4 working days) plus one day of integrated commissioning. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(119, "1 month retained: test records run from the first hydrostatic test to the close-out (P18 to P23 on the 'Programme' tab, about five weeks on either basis). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(121, "Pump and hose set standing by through the two 24-hour holds and two days of contingency between the Tank 1 test and the transfer; tankers work during the fill and are paid at 7.1, not here. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
# item-level corrections from the resource and scope review
bu['D26'] = 20
gset(26, "Peak Contractor headcount from the resource matrix ('Programme' tab, Section 4): two helper gangs and the offloading gang (up to 14), five site staff and the pipework crew - about 20 sets; the supplier's crews wear their own. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D39'] = 6
gset(39, "Survey and levelling of the existing foundation (activity QCD18TSECONSL1050, 07 to 13-Sep-2026, cost-loaded by the Contractor at SAR 5,000.00) assessed at 3 crew-days for an existing base frame, plus a dimensional and verticality survey of each erected tank, 1.5 crew-days each (RFP Scope of Works, work package 3). Not setting out: the steel base frames are already installed on concrete supports (site photograph, Aug-2026). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D40'] = 8
gset(40, "The programme carries 19 deliverable and submittal activities plus the material approval and inspection requests; 8 days of document control, doubled from the previous revision on that evidence. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D85'] = 0
gset(85, "Provisionally not assessed, to avoid duplication: sealant application at every joint is the erection work of RFP Scope of Works work package 2, performed by the supplier under Item 6 on the recorded scope of its offer, whose conditions ask the Contractor for offloading, scaffolding, storage, power and helpers only. To be confirmed against the original offer. Previously 2 tanks at SAR 1,500.00")
bu['D86'] = 0
gset(86, "Provisionally not assessed, to avoid duplication: bolts, nuts, washers and tie rods are supplied and fixed by the tank supplier (Item 6, recorded scope); no Contractor fixing or touch-up work is identified. To be confirmed against the original offer. Previously 2 tanks at SAR 2,250.00")
bu['D111'] = 0
gset(111, "Not assessed: RFP Scope of Works 5.1 and 5.2 require the tests to be witnessed by the Engineer, not inspected by a third party, and the third-party factory acceptance test is dealt with at 7.23. Previously 2 visits at SAR 2,400.00")
bu['D103'] = 0
gset(103, "Nil: the transfer labour is within the approved manpower histogram, whose helper man-days are all priced at 5.10 ('Programme' tab, Section 4, week W16); priced once. Previously 2 No. x 3 days")
bu['D107'] = 10
gset(107, "Tank 2 disinfection and flushing, 2 No. x 5 days from 13-Dec-2026, after the approved histogram ends on 11-Dec ('Programme' tab, Section 4, duty H8). The Tank 1 disinfection (H4) falls within the approved helper weeks priced at 5.10 and is not repeated. Previously 2 No. x 10 days")
# Item 8
bu['A127'] = ("External pipework, valves, fittings and instrumentation connecting the two tanks to the confirmed tie-in points (RFP Scope of Works, Piping "
              "Connections and General Piping Requirements); excluded from both tank suppliers' scopes. Priced to the RFP specification: uPVC Schedule 40 "
              "below DN150; GRP or ductile iron with internal lining above DN150; resilient-seated butterfly isolation valves, lever-operated up to DN200 "
              "and gearbox-operated above. The Contractor's procurement tracker of 29-Sep-2026 shows different materials being procured (HDPE main "
              "line, a gate valve and a motorised butterfly valve); those substitutions are not approved and are shown after line 8.21, not carried. "
              "No quantities exist: the pipework shop drawings were returned Code C and are being resubmitted, and no take-off has been received, so "
              "the quantities remain assessed from the RFP schematic and are provisional. Unit rates are supplied and installed, inclusive of labour. "
              "Lines 8.16 to 8.19 test the pipework only; tank testing is Item 7.")
bu.row_dimensions[127].height = 96
bu['B129'] = 'Main pipework above DN150 - GRP or ductile iron with internal lining, with fittings, supplied and installed (RFP General Piping Requirements)'
gset(129, "Assessed run lengths for inlet, outlet and overflow of 2 tanks - take-off required. The Contractor is procuring HDPE pipe of 355 mm and 315 mm outside diameter (nominal size subject to the pipe standard and SDR and to the Engineer's confirmation): see the alternative after line 8.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B130'] = 'Small-bore pipework below DN150 - uPVC Schedule 40, supplied and installed'
gset(130, "Assessed - take-off required. The Contractor is procuring uPVC pipe of 160 mm and 110 mm outside diameter (nominal size subject to the pipe standard and to the Engineer's confirmation); whether the 160 mm pipe falls below or at the DN150 boundary of the specification depends on the pipe standard and is to be confirmed by the Engineer. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B131'] = 'Resilient-seated butterfly isolation valves DN150-DN300, lever-operated up to DN200 and gearbox-operated above, installed'
gset(131, "Assessed count - take-off required; the procurement tracker lists 2 valves (a gate valve DN300 and a motorised butterfly valve DN355), neither to the specified type or operation - see the alternative after line 8.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(138, "Excluded from the tank supplier scope; the programme shows level transmitters installed on both tanks (activities QCD18TSECONT1MW2060 and QCD18TSECONT2MW2040). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(142, "Tie-in connections with the existing pump room after the readiness milestone of 03-Dec-2026 (activity QCD18TSECONTC2040, 4 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
# Item 8 alternative block: insert 8 rows at 151 (before ASSUMPTIONS at 152 after the earlier shift)
ASSUMP = 152
insert_rows_keep_styles(bu, ASSUMP - 1, 9, 129)   # rows 151..159 new; assumptions now at 161
for rr in range(151, 160):
    for c in range(1, 8): bu.cell(rr, c).value = None
bu['B151'] = "Contractor's procured materials in place of lines 8.1 to 8.3 - shown for information, not carried"
cp(bu['B48'], bu['B151'])
for c in (1, 3, 4, 5, 6, 7): cp(bu.cell(48, c), bu.cell(151, c))
bu.row_dimensions[151].height = 18
alt = [
    ('a.1', 'HDPE PE100 pipe, 355 mm and 315 mm outside diameter (nominal size subject to the pipe standard and SDR and to the Engineer\'s confirmation), with HDPE fittings, flange adaptors, reducer, elbow and tee, supplied and installed - in place of line 8.1', 'm', "=D129", 280,
     "Same assessed length as 8.1 (no take-off). Union Pipe Industry supply per the procurement tracker of 29-Sep-2026; PQD approved Code B 16-Sep-2026, material approval request under preparation - not approved against the GRP or ductile iron specification. The tracker gives outside diameters, not nominal sizes; the nominal size, pressure class and SDR are to be confirmed on the approved shop drawings and by the Engineer. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.2', 'uPVC pipe, 160 mm and 110 mm outside diameter (nominal size subject to the pipe standard and to the Engineer\'s confirmation), with uPVC elbows, supplied and installed - in place of line 8.2', 'm', "=D130", "=E130",
     "Same assessed length and rate as 8.2: the material conforms below DN150. Al Muneef supply; PQD submitted 15-Sep-2026, material approval request under preparation"),
    ('a.3', 'Gate valve DN300, installed - in place of one valve at line 8.3', 'No', 1, 6500,
     "Saudi Pipe Systems supply. Does not conform: the RFP General Piping Requirements call for resilient-seated butterfly isolation valves. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.4', 'Motorised butterfly valve DN355 as listed by the Contractor, installed, actuator included, power and control supply excluded - in place of one valve at line 8.3', 'No', 1, 14000,
     "Saudi Pipe Systems supply. Does not conform: the RFP requires gearbox operation above DN200; a motorised valve needs a power and control supply that is in no party's scope. Assessed market rate, Riyadh, Sep-2026 - assumption"),
    ('a.5', 'Remaining valves at line 8.3 to the RFP specification - 8 No. retained', 'No', "=D131-2", "=E131",
     "The tracker lists only 2 valves; the count at 8.3 stays assessed until the pipework shop drawings are approved and measured"),
]
rr = 152
for ref, desc, unit, q, rate, basis in alt:
    bu.cell(rr, 1, ref); bu.cell(rr, 2, desc); bu.cell(rr, 3, unit); bu.cell(rr, 4, q); bu.cell(rr, 5, rate); bu.cell(rr, 6, f"=D{rr}*E{rr}"); bu.cell(rr, 7, basis)
    bu.row_dimensions[rr].height = max(39, 13 * (len(basis) // 70 + 1))
    rr += 1
bu.cell(rr, 2, 'Alternative total for lines 8.1 to 8.3 on the procured materials - information only'); cp(bu['B63'], bu.cell(rr, 2))
bu.cell(rr, 6, f"=SUM(F152:F{rr - 1})"); cp(bu['F63'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(63, c), bu.cell(rr, c))
rr += 1
bu.cell(rr, 2, 'Assessed amount of lines 8.1 to 8.3 to the RFP specification, carried'); cp(bu['B63'], bu.cell(rr, 2))
bu.cell(rr, 6, "=F129+F130+F131"); cp(bu['F63'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(63, c), bu.cell(rr, c))
rr += 1
bu.cell(rr, 2, "Difference - not carried: the substitution is not approved"); cp(bu['B63'], bu.cell(rr, 2))
bu.cell(rr, 6, f"=F{rr - 1}-F{rr - 2}"); cp(bu['F63'], bu.cell(rr, 6))
for c in (1, 3, 4, 5, 7): cp(bu.cell(63, c), bu.cell(rr, c))
bu.row_dimensions[rr].height = 30
assert rr == 159
# assumptions block now rows 161..169
A0 = 161
assert str(bu.cell(A0, 1).value).startswith('ASSUMPTIONS'), bu.cell(A0, 1).value
bu.cell(A0 + 2, 1).value = ("2. Rates marked 'assumption' are assessed Riyadh market rates (Sep-2026): crane 50 t SAR 2,500/day; boom truck SAR 1,200/day; telehandler SAR 900/day; "
    "scaffolding SAR 55/m2; generator 100 kVA SAR 11,000/month with fuel; general labour SAR 150/day; rigger SAR 250/day; site engineer SAR 12,000/month; HSE and QA/QC "
    "SAR 10,000/month; foreman SAR 8,000/month; test water SAR 6.00/m3 tankered; site water SAR 350 per 10 m3 delivery; generator 30 kVA SAR 6,000/month continuous with fuel; "
    "local haulage SAR 450/load; HDPE PE100 pipe 355 mm outside diameter supplied and installed SAR 280/m; gate valve DN300 SAR 6,500; motorised butterfly valve DN355 as listed SAR 14,000 excluding power and control.")
bu.row_dimensions[A0 + 2].height = 50
bu.cell(A0 + 3, 1).value = ("3. Programme: the Contractor's baseline programme QC05958-BSL01THF-TSE.xer (data date 01-Jul-2026), issued 29-Sep-2026 with a cost S-curve, cash flow, "
    "manpower histogram and two-week look-ahead, is the source of every period on this tab, as set out on the 'Programme' tab. It is under the Engineer's approval "
    "(acceptable with minor comments, a procurement schedule required) and is not an agreed basis. Site period D1 on the 'Programme' tab, from mobilisation on 22-Aug-2026 to the end of "
    "demobilisation, its end set by the external readiness milestone of 03-Dec-2026 and the tie-in, sequential testing, commissioning and demobilisation that follow; "
    "labour, plant and site staff assessed phase by phase in the resource matrix on that tab. Supplier delivery terms are 7-9 weeks for the first tank and 12-14 weeks for the second, delivered duty paid from "
    "the UAE. The staged delivery is the Contractor's procurement risk: no standby, idle time or prolongation arising from it is assessed. Dates after the look-ahead data "
    "date of 28-Sep-2026 are forecasts.")
bu.row_dimensions[A0 + 3].height = 63
bu.cell(A0 + 6, 1).value = ("6. The 5 per cent Overhead and Profit on the 'Assessment' tab covers head-office overheads and profit only; every site-specific cost is priced in the items. "
    "The Contractor's cost loading of its programme spreads the proposal including that percentage across activities for progress measurement; it is an allocation, "
    "not evidence of cost, and is recorded on the 'Build-Up Comparison' tab, Section 8.")
bu.row_dimensions[A0 + 6].height = 37.8

# ---------------------------------------------------------------- Assessment tab
asm = wb['Assessment']
asm['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'
asm['A4'] = ('="Purpose: to value the Contractor\'s revised proposal ref. SAMACO-RRFP-000001 dated 13-Aug-2026 for the 2 No. TSE irrigation storage tanks, dismantling of damaged Tank-1, '
             'associated pipework and testing. Result: the proposal of SAR " & TEXT(F21,"#,##0.00") & " is assessed at SAR " & TEXT(K21,"#,##0.00") & ", both excluding VAT (row 21). '
             'Basis: supplier quotations where the Contractor obtained them (Items 3 and 6), first-principles build-ups elsewhere (\'Build-Up\' tab), with every time-related period '
             'taken from the Contractor\'s programme of 29-Sep-2026 (\'Programme\' tab) - a programme under the Engineer\'s approval, not an agreed basis. Status: preliminary - items '
             'marked Provisional depend on confirmations still outstanding. Rev 02 replaces Rev 01 dated 17-Sep-2026 (SAR " & TEXT(6921685.64,"#,##0.00") & ")."')
asm.row_dimensions[4].height = 57
asm['L8'] = ("Separate site establishment: the tank site is in the lower plateau, outside the D-18 boundary, so site staff, welfare, water, security, insurance endorsement, "
             "mobilisation, demobilisation and close-out are assessed for the site period derived from the Contractor's programme of 29-Sep-2026 ('Programme' tab, D1): "
             "mobilisation on 22-Aug-2026 to the end of demobilisation, the end being set by the external readiness milestone of 03-Dec-2026 and the tie-in, "
             "sequential testing, commissioning and demobilisation that follow; site staff are assessed role by role by phase ('Programme' tab, Section 4). Design resources are shared with the main Contract. See 'Build-Up' "
             "Item 1, 'Programme' Sections 3 to 5 and 'Build-Up Comparison' Section 3.")
asm['L14'] = ("Plant, access, power and supplier attendance that the tank quotation excludes; the supplier's price is supply and installation, so GRP erectors "
              "are not priced (recorded scope of the offer, to be confirmed against the original). Helpers and plant are assessed phase by phase on the "
              "'Programme' tab, Section 4: helpers from the approved manpower histogram (85 man-weeks, priced once, with the transfer labour line "
              "set to nil and the disinfection labour limited to the Tank 2 work after the histogram ends), plant days from the XER activity dates. No plant schedule has been submitted. See 'Build-Up' Item 5, "
              "'Programme' Sections 3 to 5 and 'Build-Up Comparison' Section 4.")
asm['L17'] = ("One tank filled and the water re-used for the second, per the Engineer's (KEO) email of 30-Aug-2026; the Contractor's programme tests both tanks in "
              "parallel, which is shown beside this basis on the 'Programme' tab and not carried; the sequential test is fitted to the programme there, with Tank 1 tested early "
              "on a conditional method assumption and integrated commissioning after both tanks have passed and the tie-in is complete. "
              "A third-party factory acceptance test the Contractor is arranging is carried at nil (line 7.23) pending evidence that the Engineer requires it. Excludes the "
              "supplier's leak-test supervision (Item 6) and re-testing, which is the Contractor's obligation. Water source to be confirmed. See 'Build-Up' Item 7.")
asm['L18'] = ("Provisional allowance priced to the RFP specification (GRP or ductile iron above DN150, uPVC below, butterfly isolation valves). The Contractor's "
              "procurement tracker of 29-Sep-2026 shows HDPE main-line pipe, a gate valve and a motorised butterfly valve being procured; the substitutions are not "
              "approved and are shown for information after line 8.21. No quantities exist: the pipework shop drawings were returned Code C and no take-off has been "
              "received. See 'Build-Up' Item 8.")
asm['L12'] = ("Lowest of the Contractor's three quotations, adopted for dismantling, segregation and loading only on the scope recorded in the previous revision; the quotation itself is not attached to this revision and its scope split is not yet confirmed in writing. Haulage and handover are Item 4. See 'Build-Up' Item 3.")
asm['N17'] = 'Provisional - water source and testing basis'
asm['A22'] = ("Contractor columns are as submitted and unchanged. Assessed rates in column J are calculated on the 'Build-Up' tab; quantities are those submitted. "
              "Status 'Provisional' marks an item that depends on a confirmation still outstanding; the list is on the 'Build-Up Comparison' tab, Section 7. The "
              "Contractor's programme of 29-Sep-2026 is under the Engineer's approval and is not an agreed basis; the Contractor has been instructed and is on site, "
              "the instruction reference not having been supplied to the Cost Consultant.")
asm.row_dimensions[22].height = 40
asm['A23'] = f"This assessment is preliminary. It establishes a reasonable commercial provision on the information available at {DOCDATE} and does not constitute agreement of the final Variation value."
asm['A24'] = ('="Changes from Rev 01 dated 17-Sep-2026 (SAR 6,921,685.64): every time-related period re-based from an assumed 3-month site period to the Contractor\'s programme of '
              '29-Sep-2026 - site staff by role and phase and facilities for the site period to the end of demobilisation (Item 1 and the generators in Item 5), helpers from the approved manpower histogram priced once and plant from the XER activity dates, access equipment on the erection windows (Item 5); every programme period read by formula from the new \'XER WBS\' source tab; Item 7 re-sequenced with the Engineer\'s sequential basis carried, commissioning staff '
              'reconciled to that sequence and a nil line added for the '
              'third-party factory test; sealant tools and fixings provisionally excluded as within the recorded Item 6 scope and third-party inspection visits removed as not required (5.12, 5.13, 7.12); PPE, survey and document control increased on the programme evidence (1.18, 2.3, 2.4); every line linked to its activities or site service on the \'Programme\' tab; '
              'Item 8 re-described to the RFP specification with the Contractor\'s procured materials shown alongside; the Contractor\'s cost '
              'loading, cash flow and manpower histogram recorded as evidence on the \'Build-Up Comparison\' tab. No rate has moved. Net effect: SAR " & TEXT(K21-6921685.64,"#,##0.00;-#,##0.00") & " (" & TEXT((K21-6921685.64)/6921685.64,"0.0%;-0.0%") & ")."')
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
bc.row_dimensions[78].height = 40
# new requests rows 84-87 (insert 4 rows before conclusion at 85 -> conclusion moves to 89)
insert_rows_keep_styles(bc, 84, 4, 83)
new_req = [
    (10, 'Whose labour the direct manpower histogram (221 man-weeks, 28-Aug to 11-Dec-2026) represents: whether the tank supplier\'s erection crews, priced within Item 6, are included', 'This assessment.', 'Not yet requested'),
    (11, 'Quantity take-off for Item 8 from the pipework shop drawings once approved (returned Code C; resubmission 28 to 30-Sep-2026 per the look-ahead)', 'This assessment.', 'Not yet requested'),
    (12, 'Material approval requests for the HDPE main-line pipe, the gate valve DN300 and the motorised butterfly valve DN355 against the RFP General Piping Requirements, and the power and control supply for the motorised valve', 'This assessment.', 'Not yet requested'),
    (13, 'Basis on which a third-party factory acceptance test is being arranged (PQD, ITP, procedure and inspector CV submitted 26 to 28-Sep-2026): whether instructed by the Engineer or the Contractor\'s own quality plan', 'This assessment.', 'Not yet requested'),
]
for i, (n, req, src, st) in enumerate(new_req):
    rr = 84 + i
    bc.cell(rr, 1).value = n; bc.cell(rr, 2).value = req; bc.cell(rr, 3).value = src; bc.cell(rr, 8).value = st
    bc.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=7)
    bc.row_dimensions[rr].height = 40
# ensure rows 84-87 have merged C:G styles like row 83 (copied). Conclusion now at 89.
assert str(bc['A89'].value).startswith('="Conclusion')
# Section 8 after conclusion
r = 91
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
        "evidence of the 16-week labour window, which sits inside the works period carried."), 70)
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
import math
def autofit(ws):
    """Raise a row's height where wrapped text would not fit (PT Sans 10 pt: about 1.15 characters per width unit, 12.8 pt per line)."""
    merged = {}
    for rg in ws.merged_cells.ranges:
        c1, r1, c2, r2 = rg.bounds
        if r1 == r2:
            merged[(r1, c1)] = sum((ws.column_dimensions[get_column_letter(c)].width or 8.66) for c in range(c1, c2 + 1))
    for row in ws.iter_rows():
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
for name in ('Assessment', 'Build-Up', 'Build-Up Comparison', 'Programme'):
    autofit(wb[name])
wb.active = wb.sheetnames.index('Project Info')
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = (ws.title == 'Project Info')
wb.save(OUT)
print('saved', OUT, 'Programme rows: sec3', SEC3_START, '-', SEC3_END, 'sec4 hdr', HDR4ROW, 'tot', TOT4)
