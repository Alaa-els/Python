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
             "formula from the activity data in Section 7. Section 4 is the resource matrix: for each phase, what material is on site, which "
             "activities are active and what they depend on, whose crew does the work, which Contractor helpers, plant and site staff are needed, "
             "and where each is priced. Section 5 lists each 'Build-Up' line whose quantity is taken from this tab, with the programme exactly as "
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
             "(ii) the erection windows are taken as submitted, and the labour and plant inside them are assessed on the days the work needs, the "
             "balance being time paced by the Contractor's staged deliveries (Section 4); (iii) testing is carried on the Engineer's basis of "
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
pg['I14'] = f"Calendar exception in {XER} (Saudi National Day). Every working-day formula on this tab excludes Fridays and this date"
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
    acts.append(dict(id=r['task_code'], name=r['task_name'], wd=int(float(r['target_drtn_hr_cnt'])) // 10,
                     start=s, finish=f, cal=cal[r['clndr_id']], sec=wbs_path(r['wbs_id']), cost=round(cost.get(r['task_id'], 0), 2),
                     tf=int(float(r['total_float_hr_cnt'])) // 10))
acts.sort(key=lambda a: (a['start'], a['id']))
assert len(acts) == 81 and abs(sum(a['cost'] for a in acts) - 8110296.62) < 0.05

# --- Section 3 layout
banner(pg, 17, '3. PROGRAMME WINDOWS - A: AS SUBMITTED; B: AS CARRIED (SUBMITTED ERECTION DATES, SEQUENTIAL TESTING)')
para(pg, 18, ("Column A reproduces the submitted programme, including its parallel testing. Column B is the basis carried: the erection dates exactly as "
              "submitted - no activity is compressed - with the Engineer's sequential testing fitted after the tank each test needs, and demobilisation "
              "following completion. Both columns are derived by formula from the activity data in Section 6 and the calendar in Section 2. The "
              "resources that each window carries, and how much of each window is productive, are assessed in Section 4; the site period is set by "
              "the external readiness milestone of 03-Dec-2026 and the tie-in, testing, commissioning and demobilisation that follow it."), height=70)
header(pg, 19, HDR3, merge_ij=True)
PG['used'] = sum((pg.row_dimensions[i].height or 14.4) for i in range(1, 20))

# activity lookup by id -> row in section 5 (filled later); we reference dates by cell so record ids
ACT_ROW = {}   # id -> row number in section 5

# Section 5 will start at row SEC5_START; compute after Section 3/4 sized. We'll pre-assign: Section 3 rows 20..47, Section 4 rows 50..~96, Section 5 from 100.
SEC5_START = 330
SEC5_HDR = SEC5_START + 2
SEC5_FIRST = int((PAGE - 58 - 21 - 44 - 32) // 18)   # activities on the section's first page
SEC5_PER = int((PAGE - 58 - 32) // 18)             # activities per continuation page under a repeated header
row_ = SEC5_HDR + 1; n_on_page = 0; cap = SEC5_FIRST
for a in acts:
    if n_on_page == cap:
        row_ += 1; n_on_page = 0; cap = SEC5_PER   # a repeated header row
    ACT_ROW[a['id']] = row_; row_ += 1; n_on_page += 1
def AS(idx): return f"$C${ACT_ROW[idx]}"   # start cell of an activity in section 5
def AF(idx): return f"$D${ACT_ROW[idx]}"   # finish cell
def AWD(idx): return f"$E${ACT_ROW[idx]}"

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
          "Activity QCD18TSECONT1INS1020, 12-Sep to 05-Oct-2026, 20 working days, paced by the partial delivery QCD18TSEPRC1120 (08-Sep to 03-Oct-2026). The identical work on Tank 2 (QCD18TSECONT2INS1020) is programmed at 6 working days: the difference is the delivery-staging idle time deducted in Section 4, not a change of dates. The look-ahead forecasts the actual start as 30-Sep-2026, delivery-driven, the Contractor's risk")
p4 = add3('P4', 'Tank 1 - wall panels (working days)', f"={AS('QCD18TSECONT1INS1030')}", f"={AF('QCD18TSECONT1INS1030')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1030, 06 to 24-Oct-2026, 16 working days, paced by the wall-panel delivery QCD18TSEPRC1170; Tank 2 walls (QCD18TSECONT2INS1030) 12 working days - difference deducted as idle in Section 4")
p5 = add3('P5', 'Tank 1 - bracing (working days)', f"={AS('QCD18TSECONT1INS1060')}", f"={AF('QCD18TSECONT1INS1060')}", f"={WD(f'C{r}', f'D{r}')}", *same(r),
          "Activity QCD18TSECONT1INS1060, 08 to 29-Oct-2026, 19 working days; Tank 2 bracing (QCD18TSECONT2INS1060) 15 working days - difference deducted as idle in Section 4")
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
          "Power tools and lighting towers are hired by the month while erection is in progress, 12-Sep to 03-Dec-2026 as submitted. The window is the same on both bases; the idle time inside it is deducted from labour and plant days in Section 4, not from monthly hire", 'cd')
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

# ================================================================ Section 4: resource matrix
r += 1
pg_break(pg, r)
pg_add(21 + 110 + 44)
banner(pg, r, "4. RESOURCE MATRIX - LABOUR, PLANT AND SITE STAFF BY PHASE, RECONCILED TO THE 'BUILD-UP' TAB"); r += 1
para(pg, r, ("Supplier scope. Item 6 is supply and installation: the Contractor's own Item 6 description reads 'Supply and install', and the adopted "
             "Al Mousa quotation S04488 with the Stalwart technical offer SS-07-26-1516 dated 03-Aug-2026 is recorded in the previous revision as "
             "including erection, with the Contractor to provide offloading and hoisting, scaffolding, storage and handling, power and 4 to 6 "
             "non-skilled labourers for material handling (the exact clause is to be quoted from the proposal pack, which is not attached to this "
             "revision; Professional Bayan's PB/Q/0086/2026 was ex-works supply only and is not adopted). GRP erectors are therefore within Item 6 "
             "and are not priced below. The histogram does not identify employers or trades; its skilled row (peak 23, 136 man-weeks) is either the "
             "supplier's erectors, already in Item 6, or Contractor labour whose role is unexplained - neither reading supports a further allowance "
             "until the Contractor answers ('Build-Up Comparison' tab, Section 7, item 10). Material availability. Only Tank 1 material is on site "
             "from the first partial delivery (08-Sep-2026 programmed; first batch received 28-Sep-2026, remaining base material due 05-Oct-2026 per "
             "the look-ahead) until the Tank 2 kit arrives (20 to 21-Oct-2026 per the tracker; 20-Oct to 04-Nov-2026 programmed): one erection "
             "front only, one helper gang. The Tank 2 workfront opens on 22-Oct-2026 and the two fronts run together until the Tank 1 nozzles finish "
             "on 14-Nov-2026; in that interval a second gang is needed only for the Tank 2 panel handling, the Tank 1 front by then being in "
             "bracing, roof and nozzle stages. Walls and bracing are one front: the programme links bracing to the walls (start 2 working days "
             "after, finish 5 working days after), tier by tier, so they carry one gang, not two. Helper-days are assessed on the productive days "
             "each task needs (the Contractor's own Tank 2 durations for identical work, the only productivity evidence, crew size assumed equal - "
             "an assumption); the balance on the submitted Tank 1 durations is time paced by the staged deliveries, to which the programme links "
             "the base and wall activities finish-to-finish (QCD18TSECONT1INS1020 to QCD18TSEPRC1120, QCD18TSECONT1INS1030 to QCD18TSEPRC1170), "
             "shown but not assessed: the Contractor's procurement risk, and its quantum unproven (crew size and redeployment unknown). Conversions: "
             "helpers to man-months at 26 working days a month (6-day week); plant by the working day on site, mobilisation within the day rate; "
             "monthly hire at 30.4 calendar days a month, rounded to one decimal place, or rounded up where the hire is by the month. Nothing "
             "priced elsewhere is repeated: dismantling labour and crane (Item 3), pipework, level transmitters and their fixing (Item 8), transfer, "
             "disinfection and commissioning labour (Item 7), site clean-up (1.16)."), height=210); r += 1
HDRM = ['Ref', 'Phase, material availability, activities and window as submitted', 'Working days - submitted', 'Working days - task-based', 'Helpers (No.)', 'Helper-days assessed', 'Helper-days delivery-paced, not assessed', 'Boom truck days', 'Telehandler days', 'Predecessors, crew ownership, source or assessed assumption; priced item']
def hdrm(ws, r_):
    header(ws, r_, HDRM[:9]); ws.cell(r_, 10, HDRM[9]); cp(S_HDR, ws.cell(r_, 10))
hdrm(pg, r); r += 1
M = {}
def madd(ref, phase, sub_, prod_, helpers, boom, tele, src, height=None):
    global r
    h = max(30, est(phase, 46), est(src, 52), height or 0)
    if PG['used'] + h > PAGE:
        old = r
        pg_break(pg, r); hdrm(pg, r); r += 1; pg_add(32)
        fix = lambda v: re.sub(r'(?<![A-Z$])([C-I])' + str(old) + r'(?!\d)', lambda m: m.group(1) + str(r), v) if isinstance(v, str) else v
        sub_, prod_, helpers, boom, tele = map(fix, (sub_, prod_, helpers, boom, tele))
    pg_add(h)
    pg.cell(r, 1, ref); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, phase); cp(S_DESC, pg.cell(r, 2))
    for col, v, fmt in ((3, sub_, '#,##0'), (4, prod_, '#,##0'), (5, helpers, '#,##0'), (8, boom, '#,##0.0'), (9, tele, '#,##0.0')):
        c = pg.cell(r, col, v); numcell(c, fmt)
        if v in (None, ''): c.value = '-'
    pg.cell(r, 6, f"=D{r}*E{r}"); numcell(pg.cell(r, 6), '#,##0')
    pg.cell(r, 7, f"=(C{r}-D{r})*E{r}"); numcell(pg.cell(r, 7), '#,##0')
    pg.cell(r, 10, src); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = h
    M[ref] = r; r += 1
    return r - 1
madd('M1', 'Mobilisation and site establishment, 22 to 26-Aug-2026 (P1)', f"={WD(f'C{p1}', f'D{p1}')}", f"=C{r}", 2, 0, 0,
     "2 helpers setting out cabins, storage and barriers - assessed assumption; the flatbed trips are 'Build-Up' 1.14. Priced at 5.10")
madd('M2', 'Dismantling of the existing tank, 27-Aug to 12-Sep-2026 (P2) - attendance only', f"={WD(f'C{p2}', f'D{p2}')}", f"=C{r}", 2, 0, 0,
     "Dismantling, segregation and loading, with the crane, are within the Al Mousa quotation S04647 (Item 3). The histogram shows 8 to 14 on site in these weeks (skilled 5 to 9): the dismantling crew. 2 Contractor helpers for barricading, permits and housekeeping - assessed assumption. Priced at 5.10")
madd('M3', 'Receipt and offloading of deliveries - Tank 1 base panels 08-Sep to 03-Oct (partial, batches), Tank 1 wall panels 04 to 19-Oct, Tank 2 complete kit 20-Oct to 04-Nov, pipes and valves 08 to 11-Nov-2026', 18, f"=C{r}", 4, f"=D{r}", f"=D{r}",
     "Delivered duty paid, offloaded by the Contractor (supplier condition). 18 truck-arrival days is an assessed assumption (4 + 4 + 8 + 2; the programme gives windows, the tracker gives batch dates, neither gives loads) - to be replaced by the delivery schedule and material inspection requests. Boom truck offloads, telehandler moves pallets to storage, 4 helpers additional to the erection gang on those days. Priced at 5.10, 5.1, 5.2", 60)
madd('M4', 'SINGLE FRONT - Tank 1 base panels, 12-Sep to 05-Oct-2026 (P3), only Tank 1 material on site; predecessors: base-panel delivery (start 3 days after, finish 2 days after it), dismantling and survey complete', f"=E{p3}", f"={AWD('QCD18TSECONT2INS1020')}", 5, 0, f"=D{r}",
     "Supplier's erection crew (Item 6). One Contractor gang of 5 helpers feeding panels - the supplier's 4 to 6 taken at the midpoint for a single front, an assessed assumption (the quotation does not say per front or in total). Task-based days: the Contractor's Tank 2 base, 6 working days, for the same 34 m x 30 m base (about 1,020 panels if 1 m x 1 m - assumption). The 14-day balance is paced by the batched delivery (finish-to-finish link) and is not assessed. Telehandler moves pallets into the footprint; no lifting plant for base panels. Priced at 5.10, 5.2", 70)
madd('M5', 'SINGLE FRONT - Tank 1 wall panels and bracing, 06 to 29-Oct-2026 (P4 and P5, one front: bracing follows the walls tier by tier); predecessors: base complete, wall-panel delivery (walls start 2 days after, finish 4 days after it)', f"={WD(f'C{p4}', f'D{p5}')}", f"={WD(f'C{p10}', f'D{p11}')}", 5, f"=ROUND(D{r}/2,0)", f"=E{p10}",
     "Supplier's crew erects walls and bracing together (Item 6); one gang of 5 helpers serves both tasks - not 5 plus 3. Task-based days: the Tank 2 walls-and-bracing envelope, 17 working days (29-Oct to 17-Nov-2026); the 4-day balance is delivery-paced and not assessed. Telehandler on the wall-panel days (12); boom truck on alternate days for the upper tiers and the long galvanised bracing members - assessed assumption. Priced at 5.10, 5.1, 5.2", 70)
madd('M6', 'Tank 1 roof supports and ladder, 31-Oct to 04-Nov-2026 (P6); predecessor: bracing complete', f"=E{p6}", f"=C{r}", 3, f"=D{r}", f"=ROUND(D{r}/2,0)", "TWO FRONTS from 22-Oct-2026: the Tank 1 gang reduces to 3 helpers (supports and ladder, little panel handling) while the second gang works the Tank 2 base and walls (M10, M11). Boom truck daily, telehandler half the days. Priced at 5.10, 5.1, 5.2")
madd('M7', 'Tank 1 roof panels, 05 to 11-Nov-2026 (P7); predecessor: roof supports complete', f"=E{p7}", f"=C{r}", 5, f"=D{r}", f"=D{r}", "Tank 1 gang back to 5 helpers for the roof panels (about 1,020 panels lifted to 4 m); boom truck and telehandler daily, concurrent with the Tank 2 walls (M11) - see the plant concurrency deduction. Priced at 5.10, 5.1, 5.2")
madd('M8', 'Tank 1 nozzles and internals, 09 to 14-Nov-2026 (within P8); predecessor: roof panels (finish 2 days after)', f"={AWD('QCD18TSECONT1MW2050')}", f"=C{r}", 2, 0, 0,
     "2 helpers attending the supplier's fitters. The level transmitter and the pipes and fittings in P8 are Item 8 supplied-and-installed lines and carry no helpers here. Priced at 5.10")
madd('M9', 'TWO FRONTS - Tank 2 base panels, 22 to 28-Oct-2026 (P9); predecessors: Tank 2 kit delivery (start 2 days after its start), foundation survey', f"=E{p9}", f"=C{r}", 5, 0, f"=D{r}",
     "Second front opens when the Tank 2 kit is on site (20 to 21-Oct-2026 per the tracker). A second gang of 5 is needed because the Tank 1 gang is still on the Tank 1 walls and bracing (M5, to 29-Oct); a shared gang would hold Tank 2 until 30-Oct and, with the same durations, push the Tank 2 mechanical completion past the 03-Dec-2026 readiness date. The Contractor's own helper row shows 12 in this week. Priced at 5.10, 5.2")
madd('M10', 'Tank 2 wall panels and bracing, 29-Oct to 17-Nov-2026 (P10 and P11, one front)', f"={WD(f'C{p10}', f'D{p11}')}", f"=C{r}", 5, f"=ROUND(E{p10}/2,0)+ROUND((D{r}-E{p10})/2,0)", f"=E{p10}", "As M5 for the second gang; submitted durations are the task-based durations. Priced at 5.10, 5.1, 5.2")
madd('M11', 'Tank 2 roof supports and ladder, 18 to 23-Nov-2026 (P12) - single front again from 15-Nov-2026', f"=E{p12}", f"=C{r}", 3, f"=D{r}", f"=ROUND(D{r}/2,0)", "As M6; the Tank 1 front is finished, so one gang remains. Priced at 5.10, 5.1, 5.2")
madd('M12', 'Tank 2 roof panels, 24 to 30-Nov-2026 (P13)', f"=E{p13}", f"=C{r}", 5, f"=D{r}", f"=D{r}", "As M7. Priced at 5.10, 5.1, 5.2")
madd('M13', 'Tank 2 nozzles and internals, 28-Nov to 02-Dec-2026 (within P14)', f"={AWD('QCD18TSECONT2MW2030')}", f"=C{r}", 2, 0, 0, "As M8. Priced at 5.10")
madd('M14', 'Tie-in, hydrostatic tests, transfer and commissioning, 17-Nov to 15-Dec-2026 (P16, P18 to P21b)', f"=H{p16}+H{p18}+H{p19}+H{p20}+H{p21}+H{p21b}", f"=C{r}", 0, 0, 0,
     "No helpers here: the transfer labour (7.4), disinfection and flushing labour (7.8), commissioning staff (7.13 to 7.17) and pipework testing crew (8.18) are priced in Items 7 and 8; the tie-in is 8.14 and 8.15. Tankers, pump and hoses are 7.1 to 7.3 and 7.22")
madd('M15', 'Demobilisation and site clean-up, 16 to 23-Dec-2026 (P23)', f"=H{p23}", f"=C{r}", 0, 0, 0,
     "Site clean-up, 4 No. x 3 days, is 'Build-Up' line 1.16 and the flatbed trips are 1.15; nothing is assessed here")
m_first, m_last = M['M1'], M['M15']
# totals
def mtot(label, f_cells, src, bold=True):
    global r
    if PG['used'] + 19.5 > PAGE:
        pg_break(pg, r); hdrm(pg, r); r += 1; pg_add(32)
    pg_add(19.5)
    pg.cell(r, 2, label); cp(S_TOTLBL, pg.cell(r, 2))
    for c in range(1, 11):
        if c != 2: cp(S_TOTLBL, pg.cell(r, c))
    for col, f in f_cells.items():
        pg.cell(r, col, f); cp(S_TOTAMT, pg.cell(r, col)); pg.cell(r, col).number_format = '#,##0.0' if col in (8, 9) else '#,##0'
    pg.cell(r, 10, src); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = max(19.5, est(src, 52))
    M[label] = r; r += 1
    return r - 1
mt3 = mtot("Helper-days assessed, task-based - carried to 'Build-Up' line 5.10 (net figure)", {6: f"=SUM(F{m_first}:F{m_last})"}, "M1 to M13; nothing in this total depends on the delivery-paced balance")
mt2 = mtot('Helper-days on the submitted Tank 1 durations beyond the task-based days - delivery-paced, not assessed', {7: f"=SUM(G{m_first}:G{m_last})"}, "M4 and M5: paced by the batched Tank 1 deliveries to which the programme links those activities finish-to-finish. Shown so that the Contractor's position can be seen; the quantum of any waiting cost is unproven (crew size and redeployment unknown) and it is the Contractor's procurement risk. Sensitivity: at the helper rate it is worth the man-months shown at S5 in Section 5")
mt1 = mtot('Helper-days on the submitted durations, gross', {6: f"=F{mt3}+G{mt2}"}, "Task-based plus delivery-paced: what the helpers would cost if every submitted day were paid")
mt4 = mtot("Helper man-months at 26 working days - 'Build-Up' line 5.10 quantity", {6: f"=ROUND(F{mt3}/26,1)"}, "Conversion for the man-month rate on the 'Build-Up' tab (6-day week, 26 working days a month)")
mt5 = mtot('Histogram as submitted: helpers 85 man-weeks (510 man-days), skilled 136 man-weeks, peak 35 in the week ending 30-Oct-2026', {6: "=85*6"}, "The Contractor's own helper row, for comparison only; it was not used to set any figure above. The week-by-week reconciliation follows the plant rows. Employers and trades are not identified in the histogram")
# plant union computed on dated activity-days (placement rule stated)
import datetime as _dt
def _wdays(a, b):
    out = []; d = a
    while d <= b:
        if d.weekday() != 4 and d != _dt.date(2026, 9, 23): out.append(d)
        d += _dt.timedelta(1)
    return out
def _first(a, b, n): return _wdays(a, b)[:n]
def _every(a, b, n, step=None):
    days = _wdays(a, b); return [days[round(i * (len(days) - 1) / max(1, n - 1))] for i in range(n)] if n > 1 else days[:1]
_D = _dt.date
deliv = (_every(_D(2026, 9, 8), _D(2026, 10, 3), 4, 6) + _every(_D(2026, 10, 4), _D(2026, 10, 19), 4, 4)
         + _every(_D(2026, 10, 20), _D(2026, 11, 4), 8, 2) + _every(_D(2026, 11, 8), _D(2026, 11, 11), 2, 2))
assert len(deliv) == 18
def plant_sets(t1_prod):
    t1b = _wdays(_D(2026, 9, 12), _D(2026, 10, 5)); t1w = _wdays(_D(2026, 10, 6), _D(2026, 10, 24)); t1br = _wdays(_D(2026, 10, 8), _D(2026, 10, 29))
    if t1_prod: t1b, t1w, t1br = t1b[:6], t1w[:12], t1br[:15]
    t1rs = _wdays(_D(2026, 10, 31), _D(2026, 11, 4)); t1r = _wdays(_D(2026, 11, 5), _D(2026, 11, 11))
    t2b = _wdays(_D(2026, 10, 22), _D(2026, 10, 28)); t2w = _wdays(_D(2026, 10, 29), _D(2026, 11, 11)); t2br = _wdays(_D(2026, 11, 1), _D(2026, 11, 17))
    t2rs = _wdays(_D(2026, 11, 18), _D(2026, 11, 23)); t2r = _wdays(_D(2026, 11, 24), _D(2026, 11, 30))
    half = lambda L: L[::2]
    boom = [deliv, half(t1w), half(t1br), t1rs, t1r, half(t2w), half(t2br), t2rs, t2r]
    tele = [deliv, t1b, t1w, half(t1rs), t1r, t2b, t2w, half(t2rs), t2r]
    return sum(map(len, boom)), len(set().union(*map(set, boom))), sum(map(len, tele)), len(set().union(*map(set, tele)))
bsum_B, bun_B, tsum_B, tun_B = plant_sets(True)
bsum_A, bun_A, tsum_A, tun_A = plant_sets(False)
mt6 = mtot('Plant-days summed by activity, M3 to M14', {8: f"=SUM(H{m_first}:H{m_last})", 9: f"=SUM(I{m_first}:I{m_last})"}, "Before allowing for one unit serving two concurrent activities")
mt7 = mtot('Less days on which one unit serves two concurrent activities', {8: bsum_B - bun_B, 9: tsum_B - tun_B},
           "Counted on dated activity-days: the task-based days placed at the start of each submitted Tank 1 window, Tank 2 as programmed, the boom truck on alternate wall and bracing days, delivery days spread through their windows (4, 4, 8 and 2). One unit of each is a provisional utilisation assumption, not a capacity finding: panel dimensions and count, panels per pallet, lift weights, unloading cycle times and hours are not in any document received. The concurrent demand is 11 working days (Tank 1 roof with Tank 2 walls, 29-Oct to 11-Nov-2026); a second telehandler for those days is the sensitivity S6 in Section 5. To be confirmed by the Contractor's plant schedule and method statement")
mt8 = mtot("Plant hire days assessed - carried to 'Build-Up' lines 5.1 (boom truck) and 5.2 (telehandler)", {8: f"=H{mt6}-H{mt7}", 9: f"=I{mt6}-I{mt7}"}, "Day-rate hire on working days on site; idle days between staged deliveries are not paid")
# --- weekly reconciliation of the assessed helper-days against the histogram helper row (dated placement rule as above)
def _place():
    days = {}
    def put(dates, n):
        for d in dates: days[d] = days.get(d, 0) + n
    put(_wdays(_D(2026, 8, 22), _D(2026, 8, 26)), 2)               # M1
    put(_wdays(_D(2026, 8, 27), _D(2026, 9, 12)), 2)               # M2
    put(deliv, 4)                                                    # M3
    put(_wdays(_D(2026, 9, 12), _D(2026, 10, 5))[:6], 5)            # M4
    put(_wdays(_D(2026, 10, 6), _D(2026, 10, 29))[:17], 5)          # M5
    put(_wdays(_D(2026, 10, 31), _D(2026, 11, 4)), 3)               # M6
    put(_wdays(_D(2026, 11, 5), _D(2026, 11, 11)), 5)               # M7
    put(_wdays(_D(2026, 11, 9), _D(2026, 11, 14)), 2)               # M8
    put(_wdays(_D(2026, 10, 22), _D(2026, 10, 28)), 5)              # M9
    put(_wdays(_D(2026, 10, 29), _D(2026, 11, 17)), 5)              # M10
    put(_wdays(_D(2026, 11, 18), _D(2026, 11, 23)), 3)              # M11
    put(_wdays(_D(2026, 11, 24), _D(2026, 11, 30)), 5)              # M12
    put(_wdays(_D(2026, 11, 28), _D(2026, 12, 2)), 2)               # M13
    return days
_placed = _place()
HIST = [('28-Aug', 3, 'Mobilisation; dismantling'), ('04-Sep', 5, 'Dismantling (Item 3 crew); attendance'), ('11-Sep', 5, 'Dismantling; survey; first base-panel batch'), ('18-Sep', 4, 'Single front: Tank 1 base panels as batches arrive'),
        ('25-Sep', 4, 'Single front: Tank 1 base (actual first batch 28-Sep-2026)'), ('02-Oct', 4, 'Single front: Tank 1 base'), ('09-Oct', 4, 'Single front: Tank 1 walls and bracing; wall-panel batches'), ('16-Oct', 4, 'Single front: Tank 1 walls and bracing'),
        ('23-Oct', 6, 'Tank 2 kit arrives 20-Oct; second front opens 22-Oct'), ('30-Oct', 12, 'Two fronts: Tank 1 bracing and roof supports; Tank 2 base and walls'), ('06-Nov', 10, 'Two fronts: Tank 1 roof; Tank 2 walls and bracing'), ('13-Nov', 8, 'Two fronts: Tank 1 roof and nozzles; Tank 2 bracing; pipe delivery'),
        ('20-Nov', 6, 'Single front: Tank 2 roof supports'), ('27-Nov', 5, 'Single front: Tank 2 roof panels; Tank 1 test'), ('04-Dec', 3, 'Tank 2 nozzles; readiness milestone'), ('11-Dec', 2, 'Tie-in; transfer; Tank 2 test')]
if PG['used'] + 32 + 17 * 18 > PAGE:
    pg_break(pg, r)
HDRW = ['Ref', 'Week ending (histogram weeks) - activities in the assessed allocation', 'Histogram helpers (No.)', 'Histogram helper-days', 'Assessed helper-days', 'Assessed helpers a day', '', '', '', 'Note']
header(pg, r, HDRW[:9]); pg.cell(r, 10, HDRW[9]); cp(S_HDR, pg.cell(r, 10)); pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9); pg_add(32); r += 1
w_first = r
wk_end = _D(2026, 8, 28)
tot_placed = 0
for i, (lab, hh, note) in enumerate(HIST):
    wk = [wk_end - _dt.timedelta(k) for k in range(7)]
    md = sum(_placed.get(d, 0) for d in wk); tot_placed += md
    wd_in = sum(1 for d in wk if d.weekday() != 4 and d != _D(2026, 9, 23))
    pg.cell(r, 1, f'W{i + 1}'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, f'{lab}-2026 - {note}'); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, hh); numcell(pg.cell(r, 3), '#,##0'); pg.cell(r, 4, f"=C{r}*6"); numcell(pg.cell(r, 4), '#,##0')
    pg.cell(r, 5, md); numcell(pg.cell(r, 5), '#,##0'); pg.cell(r, 6, f"=ROUND(E{r}/{wd_in},1)"); numcell(pg.cell(r, 6), '0.0')
    for c in (7, 8, 9): cp(S_UNIT, pg.cell(r, c))
    pg.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    pg.cell(r, 10, ''); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = 30 if len(note) > 45 else 18; pg_add(pg.row_dimensions[r].height)
    r += 1; wk_end += _dt.timedelta(7)
w_last = r - 1
pg.cell(r, 2, 'Totals - histogram helper row against the assessed allocation'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 4, f"=SUM(D{w_first}:D{w_last})"); cp(S_TOTAMT, pg.cell(r, 4)); pg.cell(r, 4).number_format = '#,##0'
pg.cell(r, 5, f"=SUM(E{w_first}:E{w_last})"); cp(S_TOTAMT, pg.cell(r, 5)); pg.cell(r, 5).number_format = '#,##0'
pg.cell(r, 10, "The assessed column places the M1 to M13 helper-days on dates by the rule stated at the plant concurrency row (task-based days at the start of each window) and sums them by histogram week; its total equals the carried helper-days above. Differences by week: the Contractor's row is lower in the single-front Tank 1 weeks than the task-based allocation (its gang was working at the pace of its own deliveries) and carries helpers for the dismantling weeks whose work is within the Item 3 quotation"); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 66; pg_add(66); WTOT = r; r += 1
assert tot_placed == 10 + 28 + 72 + 30 + 85 + 15 + 30 + 10 + 30 + 85 + 15 + 30 + 10, tot_placed
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
L('5.1', 74, bun_A, f"=H{mt8}", f"Section 4, plant hire days. A: the same activity-days on the submitted durations without the idle deduction ({bun_A} days)")
L('5.2', 75, tun_A, f"=I{mt8}", f"As 5.1 ({tun_A} days on A)")
L('5.4', 77, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "2 No. towers for the access window D6, rounded up to whole hire months")
L('5.6', 79, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "As 5.4")
L('5.7', 80, f"={MO(wm_A)}", f"={MO(wm_B)}", "Works period D2 in months - daytime works power to completion of commissioning; none during demobilisation")
L('5.8', 81, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months - welfare power runs continuously to the end of demobilisation")
L('5.9', 82, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months")
L('5.10', 83, "=ROUND(85*6/26,1)", f"=F{mt4}", "Section 4, helper man-months: task-based helper-days M1 to M13 at 26 working days a month; the delivery-paced balance is not included. A: the Contractor's histogram helper row, 85 man-weeks, converted at 26 working days a month")
L('5.11', 84, f"={MO(er_A)}", f"={MO(er_B)}", "Erection window D3 in months")
L('5.14', 87, f"=2*{MO(er_A)}", f"=2*{MO(er_B)}", "2 No. lighting towers for the erection window D3 in months")
L('7.1', 100, "=2*3774", "=3774", "A: both tanks filled at once for parallel testing (7,548 m3, RFP Scope of Works section 1). Carried: one fill, water re-used for the second tank - the Engineer's (KEO) email of 30-Aug-2026")
L('7.2', 101, "=0", f"=H{p19}+H{p20}", "Transfer pump: not needed on A; carried for the transfer P19 and the Tank 2 test P20 (top-up and hold)")
L('7.3', 102, "=0", "=1", "Transfer hoses: not needed on A; 1 week carried")
L('7.4', 103, "=0", f"=2*H{p19}", "Transfer labour: not needed on A; 2 No. for the transfer P19")
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
pg.cell(r, 1, 'S5'); cp(S_REF, pg.cell(r, 1)); pg.cell(r, 2, 'Delivery-paced helper-days on the submitted Tank 1 durations (Section 4, not assessed), if they were paid'); cp(S_DESC, pg.cell(r, 2))
pg.cell(r, 3, 'man-month'); cp(S_UNIT, pg.cell(r, 3)); pg.cell(r, 4, "='Build-Up'!E83"); cp(S_RATE, pg.cell(r, 4))
pg.cell(r, 7, f"=ROUND(G{mt2}/26,1)"); numcell(pg.cell(r, 7), '#,##0.00'); pg.cell(r, 9, f"=ROUND(G{r}*D{r},2)"); cp(S_AMT, pg.cell(r, 9))
pg.cell(r, 10, "Excluding Overhead and Profit. Would apply only if the Employer were to bear the Contractor's staged-delivery pace, which the assessment does not accept"); cp(S_BASIS, pg.cell(r, 10))
for c in (5, 6, 8): cp(S_QTY, pg.cell(r, c)); pg.cell(r, c).value = '-'
pg.row_dimensions[r].height = 30; r += 1
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
 ('1.16', 24, 'QCD18TSEDMOB1020', 'Site clean-up at demobilisation; Contractor labour', 'Attendance: 4 No. x 3 days', 'No - the demobilisation helpers are not in 5.10 (matrix M15)'),
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
 ('3.1', 46, 'QCD18TSECONDSM1020 to 1030 (27-Aug to 12-Sep-2026, complete)', 'Dismantling, segregation and loading of the damaged tank; Al Mousa under quotation S04647 for the Contractor', 'Fixed: lowest of three quotations', 'No - haulage is 4.1; the crane and crew are within the quotation, so no plant or helpers at Item 5 (matrix M2 carries attendance only)'),
 ('c.1 to c.14', 49, 'First-principles check of 3.1 - information only', 'Not carried', 'Not carried', 'Not carried'),
 ('4.1', 68, 'After QCD18TSECONDSM1030 - handover to the Employer', "Haulage of the dismantled materials to the Employer's local handover point; Contractor", 'Quantity: 14 loads, assessed; location unconfirmed', 'No - loading is within 3.1'),
 ('5.1', 74, 'Matrix M3, M5 to M7, M10 to M12 (plant rows)', 'Boom truck for offloading and lifting wall tiers, bracing, roof supports and roof panels; Contractor for the supplier', 'Hire: working days on site, one unit', "No - the supplier's own lifting is excluded from its offer"),
 ('5.2', 75, 'Matrix M3 to M5, M7, M9 to M12 (plant rows)', 'Telehandler moving pallets from the trucks to storage and into the tank footprint; Contractor for the supplier', 'Hire: working days on site, one unit, provisional', 'No'),
 ('5.3', 76, 'QCD18TSECONT1INS1030 to QCD18TSECONT2INS1050 (window D6)', 'Perimeter scaffold and edge protection for wall, bracing and roof work; Contractor (supplier condition)', 'Quantity: 1,024 m2 supplied, erected, 3-month hire, dismantled', 'No'),
 ('5.4', 77, 'Window D6', 'Mobile access towers; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.5', 78, 'Window D6', 'Relocation and re-inspection of the towers; Contractor', 'Quantity: 4 No.', 'No'),
 ('5.6', 79, 'Window D6', 'Podium steps; Contractor', 'Hire: 2 No., whole months', 'No'),
 ('5.7', 80, 'QCD18TSEMOB1240 to QCD18TSECONTCT22030 (window D2)', "Daytime works power for the supplier's tools and the Contractor's works; Contractor (supplier condition)", 'Hire with fuel: monthly, works period', 'No - welfare power is 5.8'),
 ('5.8', 81, SITE, 'Continuous welfare power; Contractor', 'Hire with fuel: monthly, site period', 'No'),
 ('5.9', 82, SITE, 'Site pick-up; Contractor', 'Hire: monthly, site period', 'No - workforce bus is 1.12; plant is 5.1 and 5.2'),
 ('5.10', 83, 'Matrix M1 to M13', "Contractor's helpers: mobilisation, dismantling attendance, offloading, panel handling on one front then two, nozzle attendance", 'Attendance: task-based helper-days converted to man-months', 'No - erectors within Item 6; transfer, disinfection and testing labour in Item 7; clean-up 1.16'),
 ('5.11', 84, 'QCD18TSECONT1INS1020 to QCD18TSECONT2MW2020 (window D3)', "Power tools for the Contractor's own works; Contractor", 'Hire: monthly, erection window', "No - the supplier's erection tools are within Item 6"),
 ('5.12', 85, 'Within Item 6 (RFP Scope of Works work package 2)', 'Sealant application - supplier erection work', 'Not assessed', 'Yes - Item 6'),
 ('5.13', 86, 'Within Item 6', 'Fixings and touch-up - supplier supply and erection', 'Not assessed', 'Yes - Item 6'),
 ('5.14', 87, 'Window D3', 'Lighting towers for the erection fronts and the work area at dusk; Contractor', 'Hire: 2 No., monthly, erection window', 'No'),
 ('5.15', 88, 'Site service - power distribution from 5.7 and 5.8', 'Distribution boards and cabling; Contractor', 'Quantity: 2 sets', 'No'),
 ('6.1', 94, 'QCD18TSEPRC1060 to QCD18TSEPRC1230, QCD18TSECONT1INS1020 to QCD18TSECONT2MW2030', 'Design, manufacture, delivery duty paid, erection, sealing, bracing, nozzles and internals of both tanks; Al Mousa / Stalwart', 'Fixed: quotation per tank, insulated, provisional', 'No - the Contractor provides offloading, scaffold, storage, power and helpers (Items 1 and 5)'),
 ('7.1', 100, 'P18 (17 to 24-Nov-2026 carried)', 'Tankered water for the first fill; Contractor', 'Quantity: 3,774 m3, one fill', 'No'),
 ('7.2', 101, 'P19 and P20', 'Transfer pump; Contractor', 'Hire: transfer plus test days', 'No'),
 ('7.3', 102, 'P19', 'Transfer hoses and fittings; Contractor', 'Hire: 1 week', 'No'),
 ('7.4', 103, 'P19', 'Pump attendance during the transfer; Contractor labour', 'Attendance: 2 No. for the transfer days', 'No - not in 5.10 (matrix M14)'),
 ('7.5', 104, 'P18 to P20', 'Top-up for losses and test level; Contractor', 'Quantity: 10 per cent of one fill', 'No'),
 ('7.6', 105, 'P18 and P20 (AWWA C652)', 'Disinfection chemicals; Contractor', 'Quantity: 800 kg', 'No'),
 ('7.7', 106, 'P18 and P20', 'Dosing equipment; Contractor', 'Hire: 2 weeks', 'No'),
 ('7.8', 107, 'P18 and P20', 'Disinfection and flushing labour; Contractor', 'Attendance: 2 No. x 10 days', 'No - not in 5.10'),
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
assert r < SEC5_START - 1, r
pg.print_title_rows = '1:3'

# --- Section 5 activity table
r = SEC5_START
pg_break(pg, r)
banner(pg, r, f"7. ACTIVITY DATA - {XER}, PROJECT 'QC05958-BSL01THF-TSE-FINAL', DATA DATE 01-JUL-2026"); r += 1
para(pg, r, ("All 81 activities as submitted, sorted by start date; names and sections are reproduced as they appear in the programme file. "
             "Working days are on each activity's own calendar: '6-day, 10 h' is the calendar 'QIC 6 Days 10 Hrs', '5-day, 10 h' is 'QIC 5 Days 10 Hrs' and '7-day, 10 h' is 'QIC 7 Days 10 Hrs' or 'QIC 7D/10H'. 'Cost loaded' is the Contractor's allocation of its proposal of SAR 8,110,296.62 "
             "including 5 per cent Overhead and Profit to the single cost resource 'TSE Tank Cost'; it is an allocation for progress "
             "measurement, not evidence of the cost of any activity."), height=44); r += 1
header(pg, r, ['No.', 'Activity (as named in the programme)', 'Start', 'Finish', 'Working days', 'Cost loaded (SAR, incl. 5% OHP)', 'Activity ID', 'Calendar', 'Total float (days)', 'Programme section (as named)'])
assert r == SEC5_HDR
r += 1
CALSHORT = {'QIC 6 Days 10 Hrs': '6-day, 10 h', 'QIC 5 Days 10 Hrs': '5-day, 10 h', 'QIC 7 Days 10 Hrs': '7-day, 10 h', 'QIC 7D/10H': '7-day, 10 h'}
for i, a in enumerate(acts, 1):
    if r != ACT_ROW[a['id']]:
        pg_break(pg, r); header(pg, r, HDR5); r += 1
    assert r == ACT_ROW[a['id']]
    pg.cell(r, 1, i); cp(S_REF, pg.cell(r, 1))
    pg.cell(r, 2, a['name']); cp(S_DESC, pg.cell(r, 2))
    pg.cell(r, 3, a['start']); datecell(pg.cell(r, 3))
    pg.cell(r, 4, a['finish']); datecell(pg.cell(r, 4))
    pg.cell(r, 5, a['wd']); numcell(pg.cell(r, 5))
    pg.cell(r, 6, a['cost']); cp(S_AMT, pg.cell(r, 6))
    pg.cell(r, 7, a['id']); cp(S_UNIT, pg.cell(r, 7)); pg.cell(r, 7).alignment = Alignment(horizontal='left', vertical='center')
    pg.cell(r, 8, CALSHORT[a['cal']]); cp(S_UNIT, pg.cell(r, 8))
    pg.cell(r, 9, a['tf']); numcell(pg.cell(r, 9))
    pg.cell(r, 10, a['sec']); cp(S_BASIS, pg.cell(r, 10))
    pg.row_dimensions[r].height = 18
    r += 1
pg.cell(r, 2, 'Total cost loaded'); cp(S_TOTLBL, pg.cell(r, 2))
for c in (1, 3, 4, 5, 7, 8, 9, 10): cp(S_TOTLBL, pg.cell(r, c))
pg.cell(r, 6, f"=SUM(F{SEC5_HDR + 1}:F{r - 1})"); cp(S_TOTAMT, pg.cell(r, 6))
pg.cell(r, 10, "Equals the Contractor's proposal of SAR 8,110,296.62 including Overhead and Profit"); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 19.5
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
             "The supplier's price under Item 6 is supply and installation, so GRP erectors are not priced here. Boundary: Item 1 is management, welfare and "
             "mobilisation; Item 5 is execution plant, access, power and supplier attendance. Helper and plant days are assessed phase by phase on the "
             "'Programme' tab, Section 4: one helper gang while only Tank 1 material is on site, a second gang only while the two erection fronts run "
             "together (22-Oct to 14-Nov-2026), walls and bracing as one front, and the days the work needs rather than the days the staged deliveries "
             "stretched the Tank 1 activities to; that balance is shown there and not assessed. One boom truck and one telehandler are a provisional "
             "utilisation assumption pending the Contractor's plant schedule. Power: the works generator by day to completion of commissioning, the welfare "
             "generator continuously to the end of demobilisation. The Engineer's (KEO) email of 30-Aug-2026: suitable Site power is unlikely, so a generator "
             "is allowed; the Contractor's methodology uses forklifts, cranes and pallet jacks.")
bu.row_dimensions[72].height = 96
gset(74, "Hire days from the 'Programme' tab, Section 4 (plant rows): offloading days plus the wall, bracing, roof-support and roof lifting days of both tanks, one unit serving concurrent activities. The Contractor's methodology names forklifts, cranes and pallet jacks (Engineer's (KEO) email of 30-Aug-2026); no plant schedule has been submitted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(75, "Hire days from the 'Programme' tab, Section 4: offloading, base, wall and roof panel days of both tanks, one unit serving concurrent activities - a provisional utilisation assumption. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(76, f"2 tanks x 128 m perimeter x 4.0 m height; the rate includes a 3-month hire, which covers the access window D6 on the '{SRC_A} on either basis. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B77'] = 'Mobile aluminium access towers, 2 No. for the access window, whole hire months'
gset(77, f"2 No. for the access window D6 on the '{SRC_A}, rounded up to whole hire months. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B79'] = 'Podium steps, 2 No. for the access window, whole hire months'
gset(79, "As 5.4. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(80, f"Works period D2 on the '{SRC_A}, to completion of commissioning. Power is excluded by both tank suppliers; no Site power expected at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(81, f"Site period D1 on the '{SRC_A}; air conditioning and lighting run around the clock at a separate site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(82, f"Site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B83'] = "Contractor's helpers - attendance to the tank supplier, offloading and panel handling, man-months from the resource matrix"
gset(83, "Supplier condition: 4-6 non-skilled labourers for material handling. Helper-days assessed phase by phase on the 'Programme' tab, Section 4 (one gang on the single front, a second gang only while two fronts run, offloading days, mobilisation and dismantling attendance), converted at 26 working days a month. The Contractor's histogram helper row (85 man-weeks) is reconciled week by week there; it is not adopted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
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
gset(85, "Not assessed: sealing tape and sealant application at every joint is the erection work in RFP Scope of Works work package 2, which the supplier performs under Item 6 with its own tools and consumables; the supplier's conditions ask the Contractor for offloading, scaffolding, storage, power and helpers only. Previously 2 tanks at SAR 1,500.00")
bu['D86'] = 0
gset(86, "Not assessed: bolts, nuts, washers and tie rods are supplied by the tank supplier and fixed by its erection crew (Item 6); no Contractor fixing or touch-up work is identified. Previously 2 tanks at SAR 2,250.00")
bu['D111'] = 0
gset(111, "Not assessed: RFP Scope of Works 5.1 and 5.2 require the tests to be witnessed by the Engineer, not inspected by a third party, and the third-party factory acceptance test is dealt with at 7.23. Previously 2 visits at SAR 2,400.00")
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
              "are not priced. Helpers and plant are assessed phase by phase on the 'Programme' tab, Section 4, from material availability and activity "
              "dependencies: one helper gang while only Tank 1 material is on site, a second only while two fronts run (22-Oct to 14-Nov-2026), the days "
              "the work needs rather than the days the staged deliveries stretched Tank 1 to. No plant schedule has been submitted. See 'Build-Up' Item 5, "
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
asm['N17'] = 'Provisional - water source and testing basis'
asm['A22'] = ("Contractor columns are as submitted and unchanged. Assessed rates in column J are calculated on the 'Build-Up' tab; quantities are those submitted. "
              "Status 'Provisional' marks an item that depends on a confirmation still outstanding; the list is on the 'Build-Up Comparison' tab, Section 7. The "
              "Contractor's programme of 29-Sep-2026 is under the Engineer's approval and is not an agreed basis; the Contractor has been instructed and is on site, "
              "the instruction reference not having been supplied to the Cost Consultant.")
asm.row_dimensions[22].height = 40
asm['A23'] = f"This assessment is preliminary. It establishes a reasonable commercial provision on the information available at {DOCDATE} and does not constitute agreement of the final Variation value."
asm['A24'] = ('="Changes from Rev 01 dated 17-Sep-2026 (SAR 6,921,685.64): every time-related period re-based from an assumed 3-month site period to the Contractor\'s programme of '
              '29-Sep-2026 - site staff by role and phase and facilities for the site period to the end of demobilisation (Item 1 and the generators in Item 5), helpers and plant from a phase-by-phase resource matrix built on material availability and activity dependencies, access equipment on the erection windows (Item 5); Item 7 re-sequenced with the Engineer\'s sequential basis carried, commissioning staff '
              'reconciled to that sequence and a nil line added for the '
              'third-party factory test; sealant tools, fixings and third-party inspection visits removed as within Item 6 or not required (5.12, 5.13, 7.12); PPE, survey and document control increased on the programme evidence (1.18, 2.3, 2.4); every line linked to its activities or site service on the \'Programme\' tab; '
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
print('saved', OUT, 'Programme rows: sec3', SEC3_START, '-', SEC3_END, 'sec4 hdr', HDR4ROW, 'tot', TOT4, 'sec5', SEC5_START)
