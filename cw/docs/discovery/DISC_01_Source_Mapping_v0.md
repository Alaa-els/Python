# DISC_01 - Source files: direct inspection and mapping to the workflow

- Date: 09-Sep-2026. Author: Alaa Elsayed. Work package: Tony Project discovery, first package.
- Basis: direct read-only inspection of the six files Tony attached to his email of 25-Aug-2026, as held in the session container (D-05 tier 1). Every workbook was opened twice, once for formulas and once for cached values; macros were extracted as text and not executed; the PDF was read page by page. Nothing was written back.
- Labels: [DIRECT] a fact read from a named workbook, sheet, cell or page. [PRIOR] an observation carried from CW_06 that this pass did not re-test. [UNKNOWN] not determinable from the files.
- Confidence: High = cell-level evidence; Medium = structure read, meaning inferred from labels; Low = inferred from names or prior notes only.
- No customer, person, rate or money figure from the files is reproduced here. Cell references and counts are structural.

## 1. Files, as inspected

| Ref | File as received | Type | Sheets | Formulas | External links | Defined names (broken) | Cached errors | Macros |
|---|---|---|---|---|---|---|---|---|
| F1 | 2026 WJL KPI Report Rev A (Latest).xlsx | company KPI tracker on a purchased template | 15 | ~3,800 | 8 targets, 22 names reference them | 1,571 (1,549) | #DIV/0! 371, #VALUE! 41, #REF! 27 | none |
| F2 | Business Development Monthly Report - May 2026.xlsx | BD pack | 14 | ~290 | none | 0 | #DIV/0! 5, #REF! 1 | none |
| F3 | Updated Commercial Project Report + Procurement Plan Report.xlsx | portfolio commercial report, CVRs, procurement planners, cashflow | 38 | ~10,500 | 10 targets, 33 names reference them | 1,582 (1,549) | #REF! 1,300, #DIV/0! 420, #VALUE! 26 | none |
| F4 | WJLxxx - AFP No (new template old dayworks rates).xlsx | blank application for payment template | 11 (1 hidden) | ~48,300 | none | 2 (2) | #DIV/0! 5,036 (blank-quantity percentages) | none |
| F5 | GCSxxx - Quotation Internal Copy - Rev 0.xlsm | estimating workbook, internal copy | 16 (1 hidden) | ~240,000 | none | 5 (2) | #REF! 2,924, #N/A 2,250, #VALUE! 8 | 3 modules, see 3.6 |
| F6 | Sinq - Variation Process.pdf | vendor walkthrough | 10 pages, image only | - | - | - | - | - |

- [DIRECT] F1 and F3 carry the same creation timestamp and the same 1,549 broken defined names. F3 was cloned from F1, or both from one parent. Seven of F1's eight external targets recur in F3. High.
- [DIRECT] External targets in F1 and F3 are workbooks on network drives (W:, J:, \\sdc.local, \\fs) and include a monthly commercial report, a finance sales and revenue forecast, an accounts-payable payments file, a procurement tracker, a CVR macro workbook dated 2019, and two quotation workbooks. These are the real upstream sources of the figures; none was supplied. High for their existence, Low for their content.
- [DIRECT] The pound sign, en dashes and emoji appear in cells; F1 CONTROL column H holds emoji icons from the template. Irrelevant to the build; noted for fixture hygiene.

## 2. Mapping: source sheet to workflow area and record

Workflow areas are the seven numbered in the 09-Sep-2026 brief. Records are named as in DISC_03.

| Area | Source sheet(s) | What the sheet holds | Records it feeds | Confidence |
|---|---|---|---|---|
| 1 Estimating history and rates | F5 Data Library (A:P): category, item, thickness, unit, install hours, material cost, wastage, description, supplier link, notes, last updated; plus a timber pricing calculator (M:P) | one rate row per material or activity with a date and a supplier | rate_record (type = tendered or quoted), supplier | High |
| 1 | F5 DPM Data Library (175 columns) and hidden Data Library (w Red): rate matrices by door size and specification, each block captioned with its supplier source and month | doorset component rates with source and date | rate_record with specification keys | Medium |
| 1 | F5 Quotation (136 columns): per-item pricing engine; columns AN:AX read the Data Library, AY:BT are manual overrides; CT:DE lump-sum buying budgets; DQ onward a cash-flow sales split | one estimate item with build-up by head, override flags, checks | estimate_item, cost_budget (tender) | High |
| 1 | F5 Tender Log (A:G): date, revision, reason, pricing tier, estimator, adjustments, comments | tender revision history | scheme_revision | High |
| 1 | F5 Letter, Terms and Conditions, Attendances (client / Gamma / N/A matrix), Prelim Breakdown (weekly gang build-up) | quotation outputs and inclusions/exclusions | estimate inclusions, contract_terms defaults, prelims cost_budget | High |
| 2 Scheme register, BOQ, cost codes, WBS | F4 Cumulative rows 8 to 272: WJL ID, item, description, quantity, unit, drawing, clarifications, rate, total, then a three-way split of value (fixing, specialist, manufacturing) each with percentage, value, progress and total | the subcontract order breakdown that every application and CVR reads | boq_item, allocation (value split by cost head) | High |
| 2 | F3 PP Template row 8: bill item, description, tender budget, target budget, variation budget, revised budget, lead time, order date, supply only / supply and fix, hours, internal / external manufacture, supplier, on order, delivered | one row per bill item joining budget, procurement and manufacture | boq_item, cost_budget, commitment, manufacture_item | High |
| 2 | F5 Handover and Budgets rows 9 to 23: quotation number, contractor, project, site, payment terms, lump sum flag, chain-of-custody flag, contacts, order value, costs by head, profit | scheme header and accepted baseline budgets | scheme, party, contract_terms, cost_budget (baseline) | High |
| 2 | Programme or WBS | [DIRECT] no programme data exists in any file; F3 PP tabs carry an "Insert link to Programme" placeholder and one link to an external programme tool | programme_activity | none in files; UNKNOWN source |
| 3 Costs, CVR, cash flow | F3 CVR template rows 9 to 48: sales block (order value, omissions, additions, variations, QS adjustments, contra charges), cost block by head (material purchases, subcontractor, production labour, site prelims, design, then the same heads under variation costs, plus PM), forecast / actual v budget with variances, payment terms, retention, discount | the per-job CVR | cost_budget, cost_actual, forecast, period_snapshot, deduction | High |
| 3 | F3 Template 2 rows 5 to 34: final account projection v interim valuation; measured works, variation account, valuation adjustments, less MCD / rebate; costs as purchases, plant, labour, subcontract invoices, adjustments; reconciled margin | the reconciliation between valuation and cost | period_snapshot, deduction (MCD, rebate) | High |
| 3 | F3 1 2026 PP row 4 (46 columns): job, PM, QS, status, customer, manufacturing tender / variation / actual hours, material budget v spend, forecast final value, variation account, forecast final cost, GP forecast, tender GP, sales and costs remaining | the portfolio roll-up | forecast, scheme summary view | High |
| 3 | F3 2A Sales Turnover and 2b Cashflow Forecast: monthly sales and cash per job for two years | company cashflow | cashflow view | High for layout; the sheets are dead, see 3.3 |
| 3 | F1 ACTUALS, TARGETS, CONTROL: 20 metric slots by month; 0.4 Manufacturing: per-job budget v spend and tender v actual hours; 06 Overheads (four copies): overhead categories and staff counts | company KPIs and overhead pool | company KPI view, overhead allocation basis | Medium |
| 3 | F4 EWN (hidden): early warning / delay notice form with issuer, delay estimate, supporting information | early warning record | early_warning | High |
| 4 Applications and VOs | F4 Cumulative rows 273 to 284: measured works, variation account, application total; previous month, movement, applied v certified, differences | application summary block | application, certificate | High |
| 4 | F4 Variation Account Summary row 12: VO, description, qty, unit, rate, amount, claim %, claim total, date submitted, client SI number, commercial status, agreement date, comments; mirrored client-certified columns O:Z; key of statuses at G6:G9 and AF2:AF9 | the VO register with both sides | variation, certificate line | High |
| 4 | F4 Variation Account Breakdown: 100 pre-built VO blocks (rows 14 to 913, nine rows each, total in column BF); build-up heads: site fixing hours, access, fixings; supplier materials; specialist materials; manufacturing materials by type, ironmongery, lacquer, misc, transport, labour hours; CAD; prelims; overhead recovery; profit margin; check columns BS:BT | VO cost build-up | variation_line, manufacture_estimate | High |
| 4 | F4 door, screen and ironmongery schedules (rows 15 to 3487): per-door specification columns then supply / fixing split and an application claim block; each schedule is split across two sheets ("2" holds the claim half) | line-level progress claim per scheduled item | boq_item (schedule lines), progress_record, application_line | High |
| 5 Manufacture | F5 Manuf. (Gamyba): in-house manufacturing budgets per item, total labour / milling hours and cost, materials, budget with mark-up | manufacture estimate per item | manufacture_estimate | High for layout; the sheet is broken, see 3.4 |
| 5 | F5 DPM (190 columns): per-door parameter model, referencing, opening, frame, leaf, apertures, screens, seals, architraves, hinges, provisions, budget rate | door parameters driving manufacture cost | manufacture_item parameters | High |
| 5 | F1 0.4 Manufacturing and F3 3 Manuf PP: tender hours, variation hours, actual hours cost, difference, expected v actual by month | manufacture actuals v tender | manufacture_actual, variance | High |
| 6 Suppliers and procurement | F5 Handover and Budgets rows 51 to 104: material purchasing, subcontract packages, internal manufacturing, each with budget, quote 1 / supplier, quote 2 / supplier | budget v quotes per package | supplier, quote, cost_budget | High |
| 6 | F3 per-job PP tabs and PP Template: procurement planner with status key (in design, released, on hold, on order, delivered, agreed, faulty, urgent, chase, quotes received, sample needed) | procurement tracking per bill item | commitment, delivery, status vocabulary | High |
| 6 | F2 Credit Check (113 columns): customer, registration, tier, monthly risk score and limit | customer credit | party attributes (customer side) | Medium |
| 7 Site evidence | F6 pages 2 to 10: login, project list, raise request with photos, sync to web, office review table, status change, transfer to variation, end-to-end summary | the capture flow | progress_record, inspection_request, evidence, variation | High as a reference; not a UI to copy |
| 7 | F4 Iron schedule 3 and Screens: application claim per scheduled item with comments | what a site claim looks like today | progress_record | Medium |
| Pipeline (optional) | F2 Orders In, Enquiries, Intake Forecast, Lost Projects, Legacy Declined, Customer Tier, Long Term Goals | pipeline and BD | out of Stage 1 scope; referenced by area 1 benchmarking only | High |

## 3. Reconciliation defects, located

### 3.1 The same figure disagrees between files (CW_06 X1) - CONFIRMED, located
- [DIRECT] F1 sheet "0.4 Manufacturing" and F3 sheet "3 Manuf PP" have the same layout: 159 of 162 text cells identical over A1:BS39. Of the populated cells, 60 differ, 40 of them numeric, in rows 7 to 19 and columns F, G, L, M, N and Q (total spend, difference, total tender hours, actual hours cost, hours difference, percentage difference). High.
- [DIRECT] Both sheets pull those figures from a third workbook through external links (F1 link [7], F3 link [8], both to a sheet named "2 Project Performance" in a monthly commercial report workbook). The two copies were refreshed at different times. F3 "3 Manuf PP" also has a broken formula at L21 and cached #REF! at L17, L21, Q21. Medium for the cause.
- Consequence for the model: cost actuals need one row per figure with an import date; reports read that row and never a copy.

### 3.2 The same project is keyed by hand in several places (X3) - CONFIRMED, quantified
- [DIRECT] Job numbers: F1 holds 17 distinct job ids, F3 44; all 17 of F1's appear in F3. F2 holds none: its six project-bearing tabs use free-text project names, 504 distinct strings, and an exact-text match finds only four names shared between two tabs (Intake Forecast and Turnover Forecast). The same scheme therefore recurs under different spellings, which is the re-keying defect itself. F4 and F5 are blank templates and hold no ids. High for counts, Medium for the "different spellings" reading.
- [DIRECT] F3 has one tab per job in three different layouts (26 to 53 columns; PP Template, Template 2, CVR template) and five multi-job tabs. Aggregation is by hand.
- Consequence: scheme code is the key on every table; project name is an attribute of one row.

### 3.3 Cashflow and turnover reporting is dead (X5) - CONFIRMED, located
- [DIRECT] F3 "2A. Sales Turnover": 813 cached #REF! values; 326 formulas reference a deleted sheet. "2b. Cashflow Forecast": 414 cached #REF!, 178 hidden rows. Both sheets are non-functional as received. High.
- [DIRECT] F1 "!Helper": 27 #REF!; F1 "0. KPI": 34 #DIV/0!, 41 #VALUE!, a broken formula at J51; ACTUALS and TARGETS: 178 and 156 #DIV/0! (ratios over empty months).

### 3.4 The estimating workbook's manufacturing tab is broken
- [DIRECT] F5 "Manuf. (Gamyba)": 2,924 cached #REF! values from 17,494 references to a deleted sheet, concentrated from column BD onward. The in-house manufacturing budget in this internal copy does not compute. High.
- [DIRECT] F5 "DPM": 2,250 #N/A from lookups into "DPM Data Library" (230,244 references); "Schedule": 8 #VALUE!. Medium on cause (unmatched keys).

### 3.5 Template state of F4
- [DIRECT] F4 is blank: Cumulative rows 9 to 272 carry 21 formula columns per line and no line data; door, screen and ironmongery schedules carry 1,600 #DIV/0! each in the claim columns because percentages divide by empty quantities. Two broken defined names. One hidden sheet (EWN). High.
- [DIRECT] The per-item value is split three ways (fixing, specialist, manufacturing) with separate progress percentages, and the application claim block (Y:AE) sits beside a client-certified block (AG:AL) with a differences column (AM). This is the applied v certified pattern the model must keep. High.

### 3.6 Macros (X10) - CLOSED by direct reading, not executed
- [DIRECT] Three modules. Module1 builds the client copy: converts every sheet to values, deletes red-flagged rows, columns and tabs, prefaced by an eight-point estimator checklist. Module2 toggles column groups and Module3 trims empty rows on a sheet named "Pricing Matrix" that no longer exists in this file (dead code). Sheet modules are empty. Nothing reads or writes data the workbench would hold. High.

### 3.7 Not found in any file - UNKNOWN
- Programme or WBS activities (only link placeholders).
- Actual cost transactions (only external links to an accounts-payable payments file and a finance forecast).
- Purchase orders, deliveries, supplier invoices (F3 PP status columns only).
- Inspection, defect or sign-off records of any kind.
- Any real application, VO or CVR of Gamma's under the new company name; every populated file is WJL.

### 3.8 Prior-only observations kept
- [PRIOR] CW_06 X2 (reporting lapsing: F1 actuals stop at Feb-2026): plausible from the dates read on the sheets but not re-measured here.
- [PRIOR] CW_06 X4 template debris (template placeholders on F1 Summary; four overhead tabs): confirmed by sheet list and CONTROL text, not itemised again.

## 4. F6 as a workflow reference

- [DIRECT] Nine steps: mobile login; project list with value, QS and PM; raise a request with title, description, up to 15 photos or videos; request syncs to a web table; office reviews; QS sets status (the dropdown shows Pricing, Declined, Submitted, Approved, Closed); claimable requests are transferred into the bill as the next variation with columns value, currently agreed, client proposed, difference, to action; end-to-end summary.
- Used for: the capture-to-variation flow, the idea that evidence carries across without re-entry, and the status vocabulary (standard UK usage). Not used for: screen layout, column set or wording (R-06).
- What F6 lacks and the brief requires: location, BOQ item, WBS activity and instruction links on the request; inspection and acceptance states; defects; the rule that no response is not acceptance.

## 5. What this supersedes

- Sections 2 to 4 of CW_06 (test data inventory) for file facts and defects. CW_06 stays as history; where it and this document disagree, this document governs (see CHANGELOG 09-Sep-2026).
