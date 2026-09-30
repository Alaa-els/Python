"""Rev 02 build - re-basing the RFP-027 assessment to the Contractor's programme of 29-Sep-2026.
Stage A: openpyxl on Alaa's master (Rev 01, 17-Sep-2026). Stage B (separate): recalc, zip-level metadata."""
import copy, datetime as dt, re, sys
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

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

def datecell(c):
    cp(S_QTY, c); c.number_format = 'dd-mmm-yyyy'; c.alignment = Alignment(horizontal='center', vertical='center')

def numcell(c, fmt='#,##0'):
    cp(S_QTY, c); c.number_format = fmt

pg['A1'] = "THE CONTRACTOR'S PROGRAMME - ACTIVITY DATA AND DURATION BASES"; cp(S_TITLE, pg['A1']); pg.row_dimensions[1].height = 25.5
pg['A2'] = 'TSE Irrigation Storage Tanks and Associated Pipeworks (RFP-027) - Contract QPMO-410-CT-05958'; cp(S_SUB, pg['A2']); pg.row_dimensions[2].height = 17.4
pg['A3'] = f'Contractor: SAMA Construction   |   Engineer: KEO   |   Cost Consultant: WT Partnership   |   {REV}, {DOCDATE}'; cp(S_SUB2, pg['A3'])
para(pg, 4, ("How to read this tab: Section 1 records the status of the programme and of the instruction. Section 2 is the working calendar "
             "taken from the programme file. Section 3 sets the programme's windows as submitted (Basis A) beside the basis carried in the "
             "assessment (Basis B), every date derived by formula from the activity data in Section 5 and the calendar in Section 2. "
             "Section 4 lists each 'Build-Up' line whose quantity is taken from this tab, with its quantity and amount on both bases; "
             "the Basis B quantity is the one linked into column D of the 'Build-Up' tab. All amounts exclude VAT."), height=57)

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
para(pg, 9, ("Use made of the programme: (i) the site period, which is governed by the external milestone 'SAJCO Readiness for Tie-In "
             "Connections' on 03-Dec-2026 and the tie-in, commissioning and demobilisation that follow it, is taken as submitted; (ii) the "
             "erection windows, which the submitted programme stretches to match staged panel deliveries, are re-derived on an efficient-working "
             "basis because delivery staging is the Contractor's procurement risk and buys no plant, labour or hire time; (iii) testing is "
             "carried on the Engineer's basis of 30-Aug-2026 (one tank filled, water re-used for the second) and shown to fit inside the "
             "submitted completion dates. Slippage against the baseline shown in the look-ahead (Tank 1 base panels forecast to start "
             "30-Sep-2026 against 12-Sep-2026 programmed) is delivery-driven and adds nothing to the assessment. Dates after the look-ahead "
             "data date of 28-Sep-2026 are forecasts, not actuals."), height=83)

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
banner(pg, 17, '3. DURATION BASES - A: THE PROGRAMME AS SUBMITTED; B: THE BASIS CARRIED IN THE ASSESSMENT')
para(pg, 18, ("Basis A reproduces the submitted programme. Basis B is the basis carried: the site period as submitted, because its end is fixed "
              "by the external readiness milestone of 03-Dec-2026 and the tie-in, commissioning and demobilisation that follow, which the "
              "Contractor's delivery staging consumes but does not extend; and an efficient-working erection sequence for the work-consumed "
              "resources (plant, attendance labour, access equipment, lighting, containers), in which Tank 1 takes the durations the Contractor "
              "programmed for the identical work on Tank 2, the Tank 2 base panels start when the Tank 1 wall panels are complete (the submitted "
              "programme itself starts Tank 2 two working days before the Tank 1 walls finish), and the activity logic is the programme's own "
              "(bracing starts 2 working days after the walls and finishes 5 working days after them; the roof follows the bracing; the "
              "mechanical works finish 3 working days after the roof). The efficient sequence is an illustrative counterfactual used only to "
              "size those resources; it is not a forecast of the Contractor's progress. Working days are on the calendar in Section 2."), height=96)
header(pg, 19, ['Ref', 'Window', 'Start - Basis A', 'Finish - Basis A', 'Days - A', 'Start - Basis B', 'Finish - Basis B', 'Days - B', 'Derivation and source'], merge_ij=True)

# activity lookup by id -> row in section 5 (filled later); we reference dates by cell so record ids
ACT_ROW = {}   # id -> row number in section 5

# Section 5 will start at row SEC5_START; compute after Section 3/4 sized. We'll pre-assign: Section 3 rows 20..47, Section 4 rows 50..~96, Section 5 from 100.
SEC5_START = 100
SEC5_HDR = SEC5_START + 2
for i, a in enumerate(acts):
    ACT_ROW[a['id']] = SEC5_HDR + 1 + i
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
    pg.row_dimensions[r].height = height or max(30, 13 * (len(deriv) // 95 + 1))
    rows3.append(r); r += 1
    return r - 1

# P1 mobilisation
p1 = add3('P1', 'Mobilisation and site establishment (calendar days)', f"={AS('QCD18TSEMOB1240')}", f"={AF('QCD18TSEMOB1250')}", f"=D{r}-C{r}+1",
          f"=C{r}", f"=D{r}", f"=G{r}-F{r}+1", f"Activities QCD18TSEMOB1240 and QCD18TSEMOB1250, 22 to 26-Aug-2026, {XER}. Same on both bases; mobilisation has taken place (the existing tank is dismantled)", 'cd')
p2 = add3('P2', 'Dismantling of the existing damaged Tank-1 (calendar days)', f"={AS('QCD18TSECONDSM1020')}", f"={AF('QCD18TSECONDSM1030')}", f"=D{r}-C{r}+1",
          f"=C{r}", f"=D{r}", f"=G{r}-F{r}+1", f"Activities QCD18TSECONDSM1020 to 1030, 27-Aug to 12-Sep-2026, {XER}. Complete per the look-ahead of 29-Sep-2026 (data date 28-Sep-2026); the completion date is not recorded, so the programmed dates are kept", 'cd')
p3 = add3('P3', 'Tank 1 - base panels (working days)', f"={AS('QCD18TSECONT1INS1020')}", f"={AF('QCD18TSECONT1INS1020')}", f"={WD(f'C{r}', f'D{r}')}",
          f"=C{r}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1020')}",
          f"A: activity QCD18TSECONT1INS1020, 12-Sep to 05-Oct-2026 (20 working days, paced by the partial delivery QCD18TSEPRC1120, 08-Sep to 03-Oct-2026). B: the same start; duration as the Contractor programmed for the identical work on Tank 2 (activity QCD18TSECONT2INS1020, 6 working days). The look-ahead forecasts the actual start as 30-Sep-2026: delivery-driven, the Contractor's risk")
p4 = add3('P4', 'Tank 1 - wall panels (working days)', f"={AS('QCD18TSECONT1INS1030')}", f"={AF('QCD18TSECONT1INS1030')}", f"={WD(f'C{r}', f'D{r}')}",
          f"={NEXT(f'G{p3}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1030')}",
          "A: activity QCD18TSECONT1INS1030, 06 to 24-Oct-2026 (16 working days, paced by the wall-panel delivery QCD18TSEPRC1170). B: follows the base panels (finish-to-start, as programmed); duration as Tank 2 (QCD18TSECONT2INS1030, 12 working days)")
p5 = add3('P5', 'Tank 1 - bracing (working days)', f"={AS('QCD18TSECONT1INS1060')}", f"={AF('QCD18TSECONT1INS1060')}", f"={WD(f'C{r}', f'D{r}')}",
          f"=WORKDAY.INTL(F{p4},2,{WK},{HOL})", f"=WORKDAY.INTL(G{p4},5,{WK},{HOL})", f"={WD(f'F{r}', f'G{r}')}",
          "A: activity QCD18TSECONT1INS1060, 08 to 29-Oct-2026. B: the programme's own logic for the Tank 2 bracing (QCD18TSECONT2INS1060): starts 2 working days after the walls start and finishes 5 working days after the walls finish")
p6 = add3('P6', 'Tank 1 - roof supports and ladder (working days)', f"={AS('QCD18TSECONT1INS1040')}", f"={AF('QCD18TSECONT1INS1040')}", f"={WD(f'C{r}', f'D{r}')}",
          f"={NEXT(f'G{p5}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT1INS1040')}",
          "A: activity QCD18TSECONT1INS1040, 31-Oct to 04-Nov-2026. B: follows the bracing (finish-to-start, as programmed); 5 working days on both tanks")
p7 = add3('P7', 'Tank 1 - roof panels (working days)', f"={AS('QCD18TSECONT1INS1050')}", f"={AF('QCD18TSECONT1INS1050')}", f"={WD(f'C{r}', f'D{r}')}",
          f"={NEXT(f'G{p6}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT1INS1050')}",
          "A: activity QCD18TSECONT1INS1050, 05 to 11-Nov-2026. B: follows the roof supports; 6 working days on both tanks")
p8 = add3('P8', 'Tank 1 - mechanical works: nozzles, level transmitter, pipes and fittings (working days)', f"={AS('QCD18TSECONT1MW2050')}", f"={AF('QCD18TSECONT1MW2055')}", f"={WD(f'C{r}', f'D{r}')}",
          f"=WORKDAY.INTL(G{p7},-2,{WK},{HOL})", f"=WORKDAY.INTL(G{p7},3,{WK},{HOL})", f"={WD(f'F{r}', f'G{r}')}",
          "A: activities QCD18TSECONT1MW2050, 2060 and 2055, 09 to 16-Nov-2026. B: the programme's own logic for Tank 2 (nozzles finish 2 working days after the roof; pipes and fittings finish 3 working days after the roof)")
p9 = add3('P9', 'Tank 2 - base panels (working days)', f"={AS('QCD18TSECONT2INS1020')}", f"={AF('QCD18TSECONT2INS1020')}", f"={WD(f'C{r}', f'D{r}')}",
          f"={NEXT(f'G{p4}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1020')}",
          "A: activity QCD18TSECONT2INS1020, 22 to 28-Oct-2026, started 2 working days after the Tank 2 delivery begins (QCD18TSEPRC1210, 20-Oct-2026) - delivery-driven. B: assumption - the erection crew moves to the Tank 2 base when the Tank 1 walls are complete, consistent with the submitted overlap (Tank 2 base 22-Oct against Tank 1 walls finishing 24-Oct) and with the histogram peak in the week ending 30-Oct-2026 when the two tanks overlap. Materials assumed available: delivery staging is the Contractor's risk", height=70)
p10 = add3('P10', 'Tank 2 - wall panels (working days)', f"={AS('QCD18TSECONT2INS1030')}", f"={AF('QCD18TSECONT2INS1030')}", f"={WD(f'C{r}', f'D{r}')}",
           f"={NEXT(f'G{p9}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1030')}", "Activity QCD18TSECONT2INS1030, 29-Oct to 11-Nov-2026, 12 working days; B follows the Tank 2 base panels")
p11 = add3('P11', 'Tank 2 - bracing (working days)', f"={AS('QCD18TSECONT2INS1060')}", f"={AF('QCD18TSECONT2INS1060')}", f"={WD(f'C{r}', f'D{r}')}",
           f"=WORKDAY.INTL(F{p10},2,{WK},{HOL})", f"=WORKDAY.INTL(G{p10},5,{WK},{HOL})", f"={WD(f'F{r}', f'G{r}')}", "Activity QCD18TSECONT2INS1060, 01 to 17-Nov-2026; B applies the same logic to the B wall dates")
p12 = add3('P12', 'Tank 2 - roof supports and ladder (working days)', f"={AS('QCD18TSECONT2INS1040')}", f"={AF('QCD18TSECONT2INS1040')}", f"={WD(f'C{r}', f'D{r}')}",
           f"={NEXT(f'G{p11}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1040')}", "Activity QCD18TSECONT2INS1040, 18 to 23-Nov-2026, 5 working days")
p13 = add3('P13', 'Tank 2 - roof panels (working days)', f"={AS('QCD18TSECONT2INS1050')}", f"={AF('QCD18TSECONT2INS1050')}", f"={WD(f'C{r}', f'D{r}')}",
           f"={NEXT(f'G{p12}')}", f"={W(f'F{r}', f'H{r}')}", f"={AWD('QCD18TSECONT2INS1050')}", "Activity QCD18TSECONT2INS1050, 24 to 30-Nov-2026, 6 working days")
p14 = add3('P14', 'Tank 2 - mechanical works: nozzles, level transmitter, pipes and fittings (working days)', f"={AS('QCD18TSECONT2MW2030')}", f"={AF('QCD18TSECONT2MW2020')}", f"={WD(f'C{r}', f'D{r}')}",
           f"=WORKDAY.INTL(G{p13},-2,{WK},{HOL})", f"=WORKDAY.INTL(G{p13},3,{WK},{HOL})", f"={WD(f'F{r}', f'G{r}')}", "Activities QCD18TSECONT2MW2030, 2040 and 2020, 28-Nov to 03-Dec-2026; B applies the same logic to the B roof date")
p15 = add3('P15', "Milestone - SAJCO readiness for the tie-in connections (external)", f"={AS('QCD18TSECONIF2050')}", f"={AF('QCD18TSECONIF2050')}", '-',
           f"=C{r}", f"=D{r}", '-', "Activity QCD18TSECONIF2050, 03-Dec-2026, zero duration. An interface milestone outside the Contractor's control; not advanced on Basis B because no evidence supports an earlier readiness. It, not the tank erection, fixes the earliest tie-in date", 'ms')
p16 = add3('P16', 'Tie-in connections with the existing pump room (working days)', f"={AS('QCD18TSECONTC2040')}", f"={AF('QCD18TSECONTC2040')}", f"={WD(f'C{r}', f'D{r}')}",
           f"=C{r}", f"=D{r}", f"=E{r}", "Activity QCD18TSECONTC2040, 05 to 08-Dec-2026, 4 working days, following the readiness milestone. Same on both bases")
p17 = add3('P17', 'Testing and commissioning as programmed - both tanks in parallel (working days)', f"={AS('QCD18TSECONTCT12050')}", f"={AF('QCD18TSECONTCT22030')}", f"={WD(f'C{r}', f'D{r}')}",
           '-', '-', '-', "Activities QCD18TSECONTCT12050 and QCD18TSECONTCT22030, 09 to 16-Dec-2026, 7 working days each, in parallel. Not carried: the Engineer's (KEO) email of 30-Aug-2026 states that installation and testing will not be in parallel and that one tank is filled and the water re-used for the second. The sequential fit is at P18 to P21")
# sequential testing fitted: A columns = fitted to the submitted erection dates; B columns = fitted to the efficient chain
p18 = add3('P18', 'Hydrostatic test - Tank 1, before the tie-in (working days)', f"={NEXT(f'D{p8}')}", f"={W(f'C{r}', f'E{r}')}", f"={AWD('QCD18TSECONTCT12050')}",
           f"={NEXT(f'G{p8}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}",
           "Fitted on both bases: starts the working day after the Tank 1 mechanical works and takes the 7 working days the Contractor programmed per tank. On the submitted dates it sits inside the 15 working days of float the programme gives Tank 1 (17 to 24-Nov-2026), so it does not delay anything. Assumptions: Tank 1 internals flushed and nozzles blind-flanged (RFP Scope of Works work packages 4 and 5); fill by tankered supply pending the Employer's confirmation of a local source; the water is retained in Tank 1 until Tank 2 is ready - losses are the 10 per cent top-up at 'Build-Up' line 7.5", height=70)
p19 = add3('P19', 'Transfer of the test water from Tank 1 to Tank 2 (working days)', f"={NEXT(f'MAX(D{p18},D{p14})')}", f"={W(f'C{r}', f'E{r}')}", 3,
           f"={NEXT(f'MAX(G{p18},G{p14})')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}",
           "Starts the working day after both the Tank 1 test has passed (P18) and Tank 2 is ready to receive water (mechanical works complete, P14, and interior flushed - RFP Scope of Works work package 4). Pumped tank to tank through temporary hoses with the tank outlet valves isolated; the tie-in to the pump room is not needed for the transfer. 3 working days is an assessed assumption pending the Contractor's method statement: 3,774 m3 at about 130 m3 per hour over 10-hour shifts, a 150 mm self-priming diesel pump against a low head (adjacent tanks at the same level, about 4 m static plus hose friction) - 'Build-Up' lines 7.2 to 7.4")
p20 = add3('P20', 'Hydrostatic test - Tank 2 (working days)', f"={NEXT(f'D{p19}')}", f"={W(f'C{r}', f'E{r}')}", f"={AWD('QCD18TSECONTCT22030')}",
           f"={NEXT(f'G{p19}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}",
           "Follows the transfer; 7 working days as programmed per tank. On the submitted dates 08 to 15-Dec-2026, ending within the programmed testing window")
p21 = add3('P21', 'Component and subsystem checks after the tie-in - instruments, nozzles, valves, ladders; tank-pump-network interfaces (working days)', f"={NEXT(f'D{p16}')}", f"={W(f'C{r}', f'E{r}')}", 3,
           f"={NEXT(f'G{p16}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}",
           "RFP Scope of Works 5.2, component testing and subsystem validation. Follows the tie-in on both bases and may overlap the Tank 2 hydrostatic test, because these checks do not need both tanks in service. 3 working days is an assessed assumption: the Contractor's programme has no separate activity for this stage (its 7-working-day 'Testing & Commissioning' activities cover the hydrostatic test and commissioning of each tank together), and programme durations are in any case not of themselves payable")
p21b = add3('P21b', 'Integrated system commissioning - full operational demonstration and witness testing (working days)', f"={NEXT(f'MAX(D{p16},D{p20},D{p21})')}", f"={W(f'C{r}', f'E{r}')}", 3,
           f"={NEXT(f'MAX(G{p16},G{p20},G{p21})')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}",
           "RFP Scope of Works 5.2, integrated system commissioning. Starts the working day after the last of the tie-in (P16), the Tank 2 hydrostatic test (P20) and the component checks (P21): both tanks must have passed and be connected before full operation is demonstrated. 3 working days is an assessed assumption on the same footing as P21. On the submitted erection dates the Tank 2 test ends 15-Dec-2026, so completion falls after the Contractor's programmed date of 16-Dec-2026; that is a consequence of testing sequentially, as the Engineer requires, and is shown, not forced either way")
p22 = add3('P22', 'Completion of testing and commissioning - both tanks', f"=D{p21b}", f"=C{r}", '-', f"=G{p21b}", f"=F{r}", '-',
           "End of the integrated system commissioning. The Contractor's programme has completion milestones QCD18TSEOMS1040 and 1050 on 16-Dec-2026 with parallel testing", 'ms')
p23 = add3('P23', 'Demobilisation, as-built drawings and close-out documents (working days)', f"={NEXT(f'D{p22}')}", f"={W(f'C{r}', f'E{r}')}", f"={AWD('QCD18TSEDMOB1020')}",
           f"={NEXT(f'G{p22}')}", f"={W(f'F{r}', f'H{r}')}", f"=E{r}", "Activities QCD18TSEDMOB1020 and QCD18TSEDMOB3020, programmed 17 to 24-Dec-2026, 7 working days in parallel. On both bases re-timed to follow the completion of sequential testing (P22); the programme as submitted, with parallel testing, ends on 24-Dec-2026")
# derived periods
pg.cell(r, 2, 'Derived periods used on the \'Build-Up\' tab'); cp(S_TOTLBL, pg.cell(r, 2))
for c in range(1, 11):
    if c != 2: cp(S_TOTLBL, pg.cell(r, c))
pg.row_dimensions[r].height = 19.5
r += 1
d1 = add3('D1', 'Site period - mobilisation start to demobilisation finish (calendar days)', f"=C{p1}", f"=D{p23}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p23}", f"=G{r}-F{r}+1",
          "Site staff, workforce transport, welfare cabins, WC, water tank and deliveries, the continuous 30 kVA welfare generator, the site pick-up and the watchman run for this period, which includes the demobilisation week so that supervision, HSE and welfare cover the demobilisation. On both bases the end is set by the external readiness milestone of 03-Dec-2026 and the tie-in, sequential testing, commissioning and demobilisation that follow it; the bases differ only in how the Tank 2 test sits against the tie-in. The programme as submitted, with parallel testing, gives 22-Aug to 24-Dec-2026, 125 days", 'cd')
d2 = add3('D2', 'Works period - mobilisation start to completion of testing and commissioning (calendar days)', f"=C{p1}", f"=D{p22}", f"=D{r}-C{r}+1", f"=F{p1}", f"=G{p22}", f"=G{r}-F{r}+1",
          "The daytime 100 kVA works generator runs for this period; no works power is needed during demobilisation, when the welfare generator alone continues", 'cd')
d3 = add3('D3', 'Erection window - Tank 1 base panels start to Tank 2 mechanical finish (calendar days)', f"=C{p3}", f"=D{p14}", f"=D{r}-C{r}+1", f"=F{p3}", f"=G{p14}", f"=G{r}-F{r}+1",
          "Supplier-attendance helpers, power tools and lighting towers are needed while erection is in progress. A: 12-Sep to 03-Dec-2026, stretched by staged deliveries. B: the efficient sequence", 'cd')
d4 = add3('D4', 'Panel-handling plant, segment 1 - first panels on site to the later of the Tank 1 roof and the Tank 2 walls (working days)', f"={AS('QCD18TSEPRC1120')}", f"=MAX(D{p7},D{p10})", f"={WD(f'C{r}', f'D{r}')}", f"=F{p3}", f"=MAX(G{p7},G{p10})", f"={WD(f'F{r}', f'G{r}')}",
          "Boom truck and telehandler days: one hired unit of each serves both adjacent tanks, so overlapping activities are counted once. The panel-handling activities are the deliveries and the base, wall, roof-support and roof installations; bracing and mechanical works need none. Segment 1 is the continuous union of those activities: on A, from the first delivery (QCD18TSEPRC1120, 08-Sep-2026) through the Tank 1 base and walls, the Tank 2 delivery, base and walls and the Tank 1 roof, without a gap (checked activity by activity in Section 5); on B, from the Tank 1 base start through the Tank 2 walls, the Tank 1 roof falling inside. No delivery is charged on B: materials are assumed available")
d5 = add3('D5', 'Panel-handling plant, segment 2 - Tank 2 roof supports and roof panels (working days)', f"=C{p12}", f"=D{p13}", f"={WD(f'C{r}', f'D{r}')}", f"=F{p12}", f"=G{p13}", f"={WD(f'F{r}', f'G{r}')}",
          "Second segment after the Tank 2 bracing; the bracing days between the two segments are not charged")
d6 = add3('D6', 'Access equipment window - Tank 1 walls start to Tank 2 roof finish (calendar days)', f"=C{p4}", f"=D{p13}", f"=D{r}-C{r}+1", f"=F{p4}", f"=G{p13}", f"=G{r}-F{r}+1",
          "Mobile access towers and podium steps are hired by the month for the period the walls, bracing and roofs are worked on; rounded up to whole months at 'Build-Up' lines 5.4 and 5.6", 'cd')
d7 = add3('D7', 'Storage containers window - first panels on site to last roof panel installed (calendar days)', f"=C{d4}", f"=D{p13}", f"=D{r}-C{r}+1", f"=F{p3}", f"=G{p13}", f"=G{r}-F{r}+1",
          "A: first delivery 08-Sep-2026 to Tank 2 roof finish 30-Nov-2026. B: efficient sequence; storage between staged deliveries is the Contractor's risk", 'cd')
SEC3_END = r - 1

# --- Section 4: re-based lines
r += 1
banner(pg, r, "4. 'BUILD-UP' LINES WHOSE QUANTITY IS TAKEN FROM THIS TAB - BOTH BASES SIDE BY SIDE"); r += 1
para(pg, r, ("Rates are those on the 'Build-Up' tab. Column G (Basis B) is linked into column D of the 'Build-Up' tab and is the quantity carried; "
             "column F (Basis A) shows what the same rate would give on the programme exactly as submitted. Lines not listed here are unchanged "
             "from the previous revision because the programme does not inform them (mobilisation trips, design resources, scaffold area, "
             "consumables, commissioning staff and pipework quantities); their basis is stated on the 'Build-Up' tab."), height=44); r += 1
header(pg, r, ['Ref', "'Build-Up' line", 'Unit', 'Rate (SAR)', 'Qty - previous revision', 'Qty - Basis A', 'Qty - Basis B (carried)', 'Amount - Basis A (SAR)', 'Amount - Basis B (SAR)'])
pg.cell(r, 10, 'Derivation'); cp(S_HDR, pg.cell(r, 10))
HDR4 = r; r += 1

# widen the merged text ranges of section headers to J? Keep A:I merges; column J only used in section 4 - acceptable.
def MO(days_cell): return f"ROUND({days_cell}/{MON},1)"

LINES = []  # (ref, build-up row, unit, qtyA formula, qtyB formula, derivation)
def L(ref, burow, qA, qB, deriv, prev=None):
    LINES.append((ref, burow, qA, qB, deriv, prev))

wm_A, wm_B = f"E{d2}", f"H{d2}"       # works period days
sp_A, sp_B = f"E{d1}", f"H{d1}"       # site period days
er_A, er_B = f"E{d3}", f"H{d3}"
pl_A, pl_B = f"(E{d4}+E{d5})", f"(H{d4}+H{d5})"
ac_A, ac_B = f"E{d6}", f"H{d6}"
ct_A, ct_B = f"E{d7}", f"H{d7}"
for ref, row_ in (('1.1', 9), ('1.2', 10), ('1.3', 11), ('1.4', 12), ('1.5', 13)):
    L(ref, row_, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months, including the demobilisation week; the post-demobilisation completion inspection is line 1.22")
for ref, row_ in (('1.6', 14), ('1.7', 15), ('1.8', 16)):
    L(ref, row_, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.9', 17, f"=2*ROUNDUP({sp_A}/7,0)", f"=2*ROUNDUP({sp_B}/7,0)", "Two 10 m3 deliveries a week for the weeks of the site period D1, rounded up to whole weeks")
L('1.10', 18, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months (facilities run to the end of demobilisation)")
L('1.11', 19, f"=2*{MO(ct_A)}", f"=2*{MO(ct_B)}", "2 No. containers for the storage window D7 in months")
L('1.12', 20, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months, transport being needed while the workforce demobilises; the manpower histogram shows labour on site for 16 weeks, 28-Aug to 11-Dec-2026, inside this period")
L('1.13', 21, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months; still subject to whether the Employer's security covers the lower plateau")
L('5.1', 74, f"={pl_A}", f"={pl_B}", "Panel-handling windows D4 and D5, working days")
L('5.2', 75, f"={pl_A}", f"={pl_B}", "As 5.1")
L('5.4', 77, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "2 No. towers for the access window D6, rounded up to whole hire months")
L('5.6', 79, f"=2*ROUNDUP({ac_A}/{MON},0)", f"=2*ROUNDUP({ac_B}/{MON},0)", "As 5.4")
L('5.7', 80, f"={MO(wm_A)}", f"={MO(wm_B)}", "Works period D2 in months - daytime works power to completion of commissioning; none during demobilisation")
L('5.8', 81, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months - welfare power runs continuously to the end of demobilisation")
L('5.9', 82, f"={MO(sp_A)}", f"={MO(sp_B)}", "Site period D1 in months")
L('5.10', 83, f"=5*{MO(er_A)}", f"=5*{MO(er_B)}", "5 No. helpers for the erection window D3 in months. The histogram's 85 helper man-weeks (19.6 man-months) are not adopted: the workbook does not say whether the supplier's erection crews, whose cost is inside the Item 6 price, are included")
L('5.11', 84, f"={MO(er_A)}", f"={MO(er_B)}", "Erection window D3 in months")
L('5.14', 87, f"=2*{MO(er_A)}", f"=2*{MO(er_B)}", "2 No. lighting towers for the erection window D3 in months")
# Item 7: A = programme's parallel testing; B = Engineer's sequential basis (carried)
L('7.1', 100, "=2*3774", "=3774", "A: both tanks filled at once for parallel testing (7,548 m3, RFP Scope of Works section 1). B: one fill, water re-used for the second tank - the Engineer's (KEO) email of 30-Aug-2026")
L('7.2', 101, "=0", "=7", "Transfer pump: not needed on A; 7 days on B (P19 transfer plus set-up and hold)")
L('7.3', 102, "=0", "=1", "Transfer hoses: not needed on A; 1 week on B")
L('7.4', 103, "=0", "=6", "Transfer labour: not needed on A; 2 No. x 3 days on B (P19)")
L('7.5', 104, "=ROUND(2*3774*0.1,0)", "=ROUND(3774*0.1,0)", "Top-up at 10 per cent of the water filled: of two fills on A, of one fill on B (retention in Tank 1 between P18 and P19)")
L('7.13', 112, f"=E{p21}+E{p21b}+2", f"=H{p21}+H{p21b}+2", "Component checks P21 and integrated commissioning P21b, plus one day at each hydrostatic test hold; the fills and holds themselves are supervised by the QA/QC inspector (line 1.3) and the supplier's leak-test supervision within Item 6")
L('7.14', 113, f"=2*(E{p21}+E{p21b})", f"=2*(H{p21}+H{p21b})", "2 No. technicians for the component checks P21 and the integrated commissioning P21b")
L('7.16', 115, f"=E{p16}+1", f"=H{p16}+1", "Tie-in P16 working days plus one day of integrated commissioning")
L('7.22', 121, "=8", "=8", "Tanker and pump standby: 3 fill days and 1 hold day for Tank 1, 3 transfer days and 1 hold day for Tank 2; the same on both bases because the programme's parallel testing would need the same standby for two simultaneous fills")

first4 = r
item_rows = {}
for ref, burow, qA, qB, deriv, prev in LINES:
    item = ref.split('.')[0]
    if item not in item_rows:
        item_rows[item] = []
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
    pg.row_dimensions[r].height = max(30, 13 * (len(deriv) // 80 + 1))
    item_rows[item].append(r)
    # link Build-Up quantity to Basis B
    bu.cell(burow, 4).value = f"=Programme!$G${r}"
    r += 1
last4 = r - 1
# item subtotals and totals on each basis
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
pg.cell(r, 10, "Basis B is the 'Assessment' tab total (row 21). Basis A replaces the listed lines with their Basis A amounts and applies the Overhead and Profit once, as on the 'Assessment' tab"); cp(S_BASIS, pg.cell(r, 10))
pg.row_dimensions[r].height = 30
TOT4 = r; r += 1
assert r < SEC5_START - 1, r

# --- Section 5 activity table
r = SEC5_START
banner(pg, r, f"5. ACTIVITY DATA - {XER}, PROJECT 'QC05958-BSL01THF-TSE-FINAL', DATA DATE 01-JUL-2026"); r += 1
para(pg, r, ("All 81 activities as submitted, sorted by start date; names and sections are reproduced as they appear in the programme file. "
             "Working days are on each activity's own calendar: '6-day, 10 h' is the calendar 'QIC 6 Days 10 Hrs', '5-day, 10 h' is 'QIC 5 Days 10 Hrs' and '7-day, 10 h' is 'QIC 7 Days 10 Hrs' or 'QIC 7D/10H'. 'Cost loaded' is the Contractor's allocation of its proposal of SAR 8,110,296.62 "
             "including 5 per cent Overhead and Profit to the single cost resource 'TSE Tank Cost'; it is an allocation for progress "
             "measurement, not evidence of the cost of any activity."), height=44); r += 1
header(pg, r, ['No.', 'Activity (as named in the programme)', 'Start', 'Finish', 'Working days', 'Cost loaded (SAR, incl. 5% OHP)', 'Activity ID', 'Calendar', 'Total float (days)', 'Programme section (as named)'])
assert r == SEC5_HDR
r += 1
CALSHORT = {'QIC 6 Days 10 Hrs': '6-day, 10 h', 'QIC 5 Days 10 Hrs': '5-day, 10 h', 'QIC 7 Days 10 Hrs': '7-day, 10 h', 'QIC 7D/10H': '7-day, 10 h'}
for i, a in enumerate(acts, 1):
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
            "03-Dec-2026 and the tie-in, sequential testing, commissioning and demobilisation that follow, not by the Contractor's staged deliveries. Design and "
            "engineering resources are shared with the main Contract (Item 2). Power is in Item 5. The tank supplier's price includes the "
            "installer's own mobilisation; the foundations and steel base frames are existing.")
bu.row_dimensions[7].height = 96
SRC_A = "Programme' tab, Section 3"
def gset(row_, text): bu.cell(row_, 7).value = text
gset(9, f"Separate site; full-time for the site period D1 on the '{SRC_A} (activity QCD18TSEMOB1240, 22-Aug-2026, to the end of demobilisation, activity QCD18TSEDMOB1020). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(10, f"Site period D1 on the '{SRC_A}, covering the demobilisation week. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(11, f"Separate site; full-time for the site period D1 on the '{SRC_A}; witnesses the hydrostatic tests. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(12, f"Site period D1 on the '{SRC_A}, covering the demobilisation week. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(13, f"Separate site; materials receipt and attendance records for the site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
for row_ in (14, 15, 16, 18):
    gset(row_, f"Site period D1 on the '{SRC_A} (activity QCD18TSEMOB1250, 22-Aug-2026, to activity QCD18TSEDMOB1020, 24-Dec-2026). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(17, f"Two 10 m3 deliveries a week for the weeks of the site period D1 on the '{SRC_A}; no mains water at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B19'] = 'Storage containers, 2 No. for the panel storage window (RFP Scope of Works, work package 1 - pre-construction and setup: safe storage of panels)'
gset(19, f"2 No. containers for the storage window D7 on the '{SRC_A} (efficient sequence, Tank 1 base start to Tank 2 roof finish). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(20, f"Separate site outside the D-18 boundary; site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(21, f"Site period D1 on the '{SRC_A}. Included pending confirmation whether the Employer's security covers the lower plateau; delete if it does. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(22, "5 trips retained: the programme gives no trip count (activity QCD18TSEMOB1240, mobilisation, 5 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(23, "5 trips retained: the programme gives no trip count (activity QCD18TSEDMOB1020, demobilisation, 7 working days). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['D30'] = 3
gset(30, "3 days after demobilisation: the Engineer's completion inspection, the correction list and the warranties collation. Reduced from 8 days because the site engineer's monthly period at 1.1 now runs to the end of demobilisation and the programmed close-out activity (QCD18TSEDMOB3020, 17 to 24-Dec-2026) falls inside it. Test records, ITP/WIR close-out and as-built drawings are in Item 7 (7.20 and 7.21), not here. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B30'] = "Close-out at the tank site after demobilisation - completion inspection, correction list and warranties collation - site engineer"
gset(37, "2 months retained: the programme shows shop drawings 07 to 25-Aug-2026 (activities QCD18TSEENG1240 and 1250) with resubmissions to 29-Sep-2026 and the pipework shop drawings still in preparation, consistent with the period assessed. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(38, "1.5 months retained on the same programme evidence as 2.1. As-built drawings at close-out are line 7.21. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['A72'] = ("Covers what the tank supplier's offer (Stalwart technical offer SS-07-26-1516 dated 03-Aug-2026) requires the Contractor to provide free "
             "of cost: offloading and hoisting, scaffolding, storage and handling, power, and 4-6 helpers; plus general site support. Boundary: Item 1 "
             "is management, welfare and mobilisation; Item 5 is execution plant, access, power and supplier attendance. Plant, attendance and access "
             "periods are taken from the erection windows on the 'Programme' tab on the efficient-working basis (Basis B): the submitted programme "
             "stretches the Tank 1 erection to match staged panel deliveries (base panels 20 working days against 6 for the identical work on Tank 2), "
             "and standby between deliveries is the Contractor's risk and is not included. One boom truck and one telehandler serve both adjacent tanks. "
             "Power: the works generator by day to completion of commissioning, the welfare generator continuously to the end of demobilisation. The Engineer's (KEO) email of "
             "30-Aug-2026: suitable Site power is unlikely, so a generator is allowed; the Contractor's methodology uses forklifts, cranes and pallet jacks.")
bu.row_dimensions[72].height = 96
gset(74, f"Panel-handling windows D4 and D5 on the '{SRC_A} (Basis B); the same windows on the programme as submitted give the Basis A days on that tab. The Contractor's methodology names forklifts, cranes and pallet jacks (Engineer's (KEO) email of 30-Aug-2026); no plant schedule has been submitted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(75, "As 5.1. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(76, f"2 tanks x 128 m perimeter x 4.0 m height; the rate includes a 3-month hire, which covers the access window D6 on the '{SRC_A} on either basis. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B77'] = 'Mobile aluminium access towers, 2 No. for the access window, whole hire months'
gset(77, f"2 No. for the access window D6 on the '{SRC_A}, rounded up to whole hire months. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B79'] = 'Podium steps, 2 No. for the access window, whole hire months'
gset(79, "As 5.4. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(80, f"Works period D2 on the '{SRC_A}, to completion of commissioning. Power is excluded by both tank suppliers; no Site power expected at the tank site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(81, f"Site period D1 on the '{SRC_A}; air conditioning and lighting run around the clock at a separate site. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(82, f"Site period D1 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B83'] = 'Attendance labour for the tank supplier - 5 No. helpers for the erection window'
gset(83, f"Supplier condition: 4-6 non-skilled labourers for material handling; erection window D3 on the '{SRC_A} (Basis B). The Contractor's histogram shows 85 helper man-weeks but does not say whether the supplier's own erection crews are included, so it is not adopted. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(84, f"Erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['B87'] = 'Mobile lighting towers, 2 No. for the erection window'
gset(87, f"2 No. for the erection window D3 on the '{SRC_A}. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
bu['A98'] = ("The Engineer's (KEO) email of 30-Aug-2026: installation and testing will not be in parallel; one tank is filled and the water is re-used for "
             "the second by pumping across. This basis is carried. The Contractor's programme tests both tanks in parallel on 09 to 16-Dec-2026, which "
             "would need 7,548 m3 of water at once; that basis is shown beside this one on the 'Programme' tab, Section 4 (lines 7.1 to 7.5 and 7.22) "
             "and is not carried. The sequential test is fitted to the programme ('Programme' tab, P18 to P22): Tank 1 is tested before the tie-in, "
             "within its programme float, on the conditional assumption of an Engineer-approved method with the water retained in Tank 1; the water is "
             "transferred once Tank 1 has passed and Tank 2 is ready; component checks follow the tie-in; integrated commissioning follows the last of "
             "the tie-in, the Tank 2 test and the component checks. The Contractor arranges the water; the source is being checked by the Employer, so tankered supply "
             "is assumed. Excludes the leak-test supervision already in the supplier's price (Item 6) and any re-testing after an unsatisfactory "
             "test, which is the Contractor's obligation under RFP Scope of Works 5.1. External pipework testing is Item 8. Scope: RFP Scope of "
             "Works 5.1 to 5.3.")
bu.row_dimensions[98].height = 96
gset(100, "3,774 m3 effective at 3.7 m water level (RFP Scope of Works section 1); one fill, the water re-used for Tank 2. Tankered at SAR 6.00/m3 - assumption; falls away if a network fill is confirmed")
gset(101, "7 days: the 3-working-day transfer at P19 on the 'Programme' tab plus set-up, priming and the test hold; pump duty an assessed assumption pending the Contractor's method statement. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(103, "2 No. x 3 days, the transfer window P19 on the 'Programme' tab. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(104, "10 per cent of one fill, covering losses while the water is retained in Tank 1 between its test and the transfer to Tank 2 - assumption")
gset(112, "Component checks and integrated commissioning (P21 and P21b on the 'Programme' tab) plus one day at each hydrostatic test hold; reduced from 15 days because the fills and holds are supervised by the QA/QC inspector (1.3) and the supplier (Item 6). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(113, "2 No. for the component checks and the integrated commissioning (P21 and P21b on the 'Programme' tab); reduced from 2 No. x 10 days. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(115, "Tie-in connections (P16 on the 'Programme' tab, 4 working days) plus one day of integrated commissioning. Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(119, "1 month retained: test records run from the first hydrostatic test to the close-out (P18 to P23 on the 'Programme' tab, about five weeks on either basis). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
gset(121, "3 fill days and 1 hold day for Tank 1, 3 transfer days and 1 hold day for Tank 2 (P18 to P20 on the 'Programme' tab). Assessed market rate, Riyadh, Sep-2026 - assumption pending the Contractor's substantiation")
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
    "work-consumed resources on an efficient-working sequence in which Tank 1 takes the durations programmed for Tank 2. Supplier delivery terms are 7-9 weeks for the first tank and 12-14 weeks for the second, delivered duty paid from "
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
             "sequential testing, commissioning and demobilisation that follow, not by the Contractor's staged deliveries. Design resources are shared with the main Contract. See 'Build-Up' "
             "Item 1, 'Programme' Sections 3 and 4 and 'Build-Up Comparison' Section 3.")
asm['L14'] = ("Plant, access, power and supplier attendance that the tank quotation excludes. Power for the whole site period; plant, attendance labour, access "
              "equipment and lighting for the erection windows on an efficient-working reading of the Contractor's programme, in which Tank 1 takes the durations "
              "programmed for the identical work on Tank 2 - the submitted programme stretches Tank 1 to match staged deliveries, which is the Contractor's risk. "
              "No plant schedule has been submitted. See 'Build-Up' Item 5, 'Programme' Sections 3 and 4 and 'Build-Up Comparison' Section 4.")
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
              '29-Sep-2026 - site staff and facilities for the site period to the end of demobilisation (Item 1 and the generators in Item 5), plant, attendance and access '
              'equipment on the efficient-working erection windows (Item 5); Item 7 re-sequenced with the Engineer\'s sequential basis carried, commissioning staff '
              'reconciled to that sequence and a nil line added for the '
              'third-party factory test; Item 8 re-described to the RFP specification with the Contractor\'s procured materials shown alongside; the Contractor\'s cost '
              'loading, cash flow and manpower histogram recorded as evidence on the \'Build-Up Comparison\' tab. No rate has moved. Net effect: SAR " & TEXT(K21-6921685.64,"#,##0.00;-#,##0.00") & " (" & TEXT((K21-6921685.64)/6921685.64,"0.0%;-0.0%") & ")."')
asm.row_dimensions[24].height = 54

# ---------------------------------------------------------------- Build-Up Comparison
bc = wb['Build-Up Comparison']
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
print('saved', OUT, 'Programme rows: sec3', SEC3_START, '-', SEC3_END, 'sec4 hdr', HDR4, 'tot', TOT4, 'sec5', SEC5_START)
