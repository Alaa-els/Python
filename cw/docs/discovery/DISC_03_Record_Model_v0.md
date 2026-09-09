# DISC_03 - Shared record model: entities, relationships, statuses, rules

- Date: 09-Sep-2026. Author: Alaa Elsayed. Basis: DISC_01 sources, DISC_02 coverage, CLAUDE.md verifiable rules.
- Plain-language model for the S01 and S02 walk-throughs; no SQL. Every table carries company, created_by, created_at, updated_by, updated_at and history. Imported rows carry source_file, source_tab, source_row, imported_at. Nothing is hard-deleted; rows are closed or superseded.
- "One row =" says what a row means. Names are working names; the walk-through may rename.

## 1. Organisation

| Table | One row = | Links | Notes |
|---|---|---|---|
| company | one legal entity using the workbench | - | display_name setting (Alaa currently calls the new company GEMMA; confirm) |
| user | one person who logs in | company | |
| role | one permission set (owner, editor, viewer, site) | company | site = phone capture only |
| party | one organisation on a scheme: customer or main contractor, supplier, subcontractor, consultant, employer | company | party_role per scheme |
| supplier | one party that can be sent an RFQ | party | trade, geography, capability, quality note, performance notes per scheme |

## 2. Scheme register

| Table | One row = | Links | Notes |
|---|---|---|---|
| scheme | one job the company prices or delivers | company, parties | code is stable and unique per company (format: question Q3); name, site, my_role, status (enquiry, tendered, live, final account, closed) |
| scheme_revision | one commercial baseline of a scheme | scheme | kind = tender (each Tender Log revision), baseline (the accepted order), change (each subsequent agreed change); date, reason, who; exactly one baseline per scheme |
| contract_terms | one term for one scheme | scheme | retention percent, payment days, certification days, MCD or rebate percent, discount, notice days, time bar, defects period, chain-of-custody flag; each with a clause reference |
| cost_code | one code in the company's chart of cost heads | company | tree: head (material purchases, specialist materials, subcontractors, production labour, site prelims, design, manufacturing materials, transport, project management, overhead recovery, profit) then sub-code |
| programme_activity | one WBS activity | scheme | code, name, planned start and finish; entered or imported from CSV; no programme tool link in Stage 1 |
| boq_item | one bill or activity-schedule item on one revision | scheme_revision | item id (F4 "WJL ID"), section, description, quantity, unit, rate, drawing, clarifications, provisional_sum flag, manufactured_in_house flag |
| allocation | one share of one boq_item to one cost_code and optionally one programme_activity | boq_item, cost_code, programme_activity | share as a percentage; the F4 fixing / specialist / manufacturing split is three allocations |

Rules: allocation shares for an item sum to 100 percent; quantity and value live on boq_item only, never on an allocation; a boq_item on a superseded revision is read-only.

## 3. Estimating and rates

| Table | One row = | Links | Notes |
|---|---|---|---|
| estimate | one priced tender revision | scheme_revision | estimator, pricing tier, status |
| estimate_item | one priced line with its build-up by cost head | estimate, boq_item | fixing labour and hours, materials, specialist, manufacture, prelims, design, overhead, profit; override flag per head (F5 Quotation AY:BT) |
| rate_record | one rate for one thing at one time | scheme (source), supplier (optional), cost_code | type = quoted, tendered, ordered, actual; date; revision; specification key; unit; quantity basis; region; inclusions and exclusions text; comparable flag with reason; source_file, tab, row |
| inclusion | one inclusion or exclusion statement on an estimate | estimate | from F5 Letter terms, Attendances and Prelim Breakdown comments |

Rules: a rate_record is created, never edited, when a quote is received (quoted), a tender is submitted (tendered), a commitment is placed (ordered) or an actual is matched (actual); benchmarking queries filter comparable = true and group by specification key; a non-comparable row is shown, never averaged.

## 4. Procurement

| Table | One row = | Links | Notes |
|---|---|---|---|
| package | one thing to buy on a scheme (material, subcontract, internal manufacture) | scheme_revision, cost_code | budget from F5 Handover rows 51 to 104; lead time, target order date (F3 PP planner) |
| rfq | one draft request for quotation | package, supplier | draft only; export to document; never sent from the workbench |
| rfq_line | one item on the RFQ | rfq, boq_item | quantity, unit, specification |
| quote | one supplier response | rfq, supplier | date, validity, inclusions and exclusions |
| quote_line | one priced line | quote, rfq_line | rate; creates a rate_record (quoted) |
| commitment | one purchase order or subcontract | package, supplier, scheme_revision | kind, number, date, total, payment terms; creates rate_record (ordered) per line |
| commitment_line | one line on a commitment | commitment, boq_item, cost_code | quantity, rate, amount |
| delivery | one delivery received | commitment | date, quantity, delivery note reference, evidence |
| supplier_invoice | one invoice from a supplier | commitment, supplier | number, date, amount, matched status |
| invoice_match | one match of an invoice to commitment lines and deliveries | supplier_invoice, commitment_line, delivery | difference and reason |

Statuses: package (budgeted, rfq drafted, quoted, selected, ordered, delivered, closed); commitment (draft, issued, part delivered, delivered, invoiced, paid, closed).

## 5. Cost ledger and periods

| Table | One row = | Links | Notes |
|---|---|---|---|
| cost_budget | one budget amount for one cost_code on one revision | scheme_revision, cost_code, boq_item (optional) | baseline from F5 Handover; variation budgets from agreed VOs |
| accrual | one cost expected but not yet invoiced | commitment_line or cost_code, period | reversed when the invoice is matched |
| cost_actual | one actual cost line from the accounts export | scheme, cost_code, supplier_invoice (optional), period | import date and source row; creates rate_record (actual) where a unit exists |
| cash_paid | one payment made to a supplier | supplier_invoice, period | date, amount |
| forecast | one forecast-to-complete for one cost_code or the final value | scheme, cost_code, period | who set it, when, basis |
| period | one reporting period for one scheme | scheme | month end; status open or closed |
| period_snapshot | the frozen set of report figures for one closed period | period | every CVR, application and cash-flow figure with its source row; reports for a closed period read the snapshot |
| risk | one commercial risk or opportunity | scheme, period | value, likelihood, owner, status |
| early_warning | one early warning or delay notice | scheme, period, instruction (optional) | issuer, delay estimate, supporting evidence (F4 EWN) |

Double-counting rules, each a test:
- Cost to date for a code = matched invoices + unmatched invoices + accruals; an accrual and its matched invoice never both count in the same period.
- A commitment counts as committed, never as cost, until an invoice or accrual exists against it.
- A cash_paid row reduces the creditor balance; it never adds to cost.
- A provisional sum's value is either the allowance or the replacement work, never both; replacement boq_items reference the provisional sum they draw down and the report nets them.
- A snapshot, once written, is not recalculated; a correction is a new row in the next period with a reason.

## 6. Income: applications, certificates, receipts, variations

| Table | One row = | Links | Notes |
|---|---|---|---|
| application | one application for payment for one period | scheme, period | number, date, previous cumulative, current, cumulative; retention and deductions computed from contract_terms; status (draft, submitted, assessed, certified, paid, closed) |
| application_line | one line of one application | application, boq_item or variation_line | quantity or percent claimed, previous, current, cumulative; progress basis = accepted inspection_requests (rule below) |
| deduction | one deduction on one application | application, contract_terms | kind (retention, MCD, rebate, discount, contra charge), basis, amount |
| certificate | one certification received against an application | application | date, certified per line, differences and comments |
| receipt | one payment received | certificate | date, amount, retention released |
| instruction | one instruction from the customer or main contractor | scheme, party | reference, date, kind (written, verbal confirmed, drawing issue), evidence |
| variation | one change to the scheme | scheme, instruction (optional) | code = scheme code + sequence; description; status (raised, pricing, submitted, agreed, declined, instructed only, closed); claim percent; dates submitted and agreed; client reference |
| variation_line | one line of a VO build-up | variation, cost_code, rate_record (optional) | quantity, rate, amount by head (site fixing hours, access, fixings, materials, specialist, manufacture, CAD, prelims, overhead, profit); progress percent |
| substantiation | one link from a VO to a piece of evidence | variation, evidence | why it substantiates |

Rules: applied, certified and received are three different rows, never one field; a VO with no instruction row shows as claimed; an agreed VO adds its lines to cost_budget (variation budget) once, on agreement.

## 7. Manufacture

| Table | One row = | Links | Notes |
|---|---|---|---|
| manufacture_item | one item made in-house | boq_item or variation_line | any joinery item: doorset, reception desk, cladding panel; DPM parameters optional |
| manufacture_estimate | the estimate for one manufacture_item | manufacture_item | materials by type, waste percent, labour hours and rate, design hours, sundries, shipping, storage, overhead allocation basis |
| manufacture_actual | actual cost lines for one manufacture_item | manufacture_item, cost_actual, period | same heads; hours from timesheets or imports |
| variance | one explained difference between estimate and actual | manufacture_item or scheme, period | head, amount, reason, who |

Rule: overhead allocation uses one company setting (per labour hour or percent of cost) and shows the basis on every figure.

## 8. Site records and evidence

| Table | One row = | Links | Notes |
|---|---|---|---|
| progress_record | one capture from site | scheme, location, boq_item, programme_activity, variation (optional), instruction (optional) | kind (progress, dayworks, diary, quality); quantity or percent; who, when; evidence |
| inspection_request | one request for inspection or sign-off | progress_record | state = internal_complete, submitted, acknowledged, accepted, rejected; each transition has actor, time and evidence |
| defect | one defect or rework item | inspection_request, boq_item | raised by, description, status (open, reworked, re-inspected, closed) |
| evidence | one file (photo, document, drawing mark-up) | any record via a link table | version, hash, taken_at, uploaded_by, location; role check on read |
| location | one place on the scheme | scheme | building, level, zone, room |

Rules: only accepted inspection_requests feed application_line progress; submitted and acknowledged show as claimed-not-accepted; a lack of response never advances the state; the state history is the audit record.

## 9. Cross-cutting

| Table | One row = | Notes |
|---|---|---|
| sources | one source document or import file | name, version, date, hash; every import references it |
| history | one change to any row | who, when, field, old, new |
| action | one thing someone must do | owner, due, origin record, status |
| setting | one configurable value | per company or per scheme; every threshold, format and layout choice |

## 10. Money and progress status vocabulary

- Money: budget, committed, accrued, actual, paid (cost side); applied, assessed, certified, received (income side); forecast.
- Progress: internal_complete, submitted, acknowledged, accepted, rejected.
- Variation: raised, pricing, submitted, agreed, declined, instructed_only, closed.
- Every exported figure shows its status label and its source row.
