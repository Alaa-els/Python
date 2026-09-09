# DISC_02 - Requirements coverage: the seven Stage 1 areas against D-14 and the current stages

- Date: 09-Sep-2026. Author: Alaa Elsayed. Basis: the management brief of 09-Sep-2026 (Alaa via Codex), docs/DECISIONS.md D-14, docs/STAGES.md as at 08-Sep-2026, DISC_01.
- Status: COVERED (an existing stage delivers it as specified), PARTIAL (a stage touches it; named gap), GAP (no stage), EXCLUDED-BY-D-14 (the 08-Sep scope line put it out; the 09-Sep brief puts it back; recorded as D-15).
- "Increment" refers to DISC_04. Nothing here is parked; every GAP has an increment.

## Area 1 - Estimating history and rate build-ups

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Import estimating history across schemes | EXCLUDED-BY-D-14 (estimating engine out) | none | The engine stays in Excel; its outputs (rates, build-ups, budgets) are imported per scheme. D-15. Increment 2 (structure on the fixture scheme), full history import when real quotation files are supplied |
| Rate build-ups preserved with source scheme, date, revision, specification, unit, quantity, region, supplier, inclusions / exclusions | GAP | S27 rate library holds rate, source, date only | rate_record gains the nine attributes (DISC_03). Increment 2 |
| Distinguish quoted, tendered, ordered, actual rates | GAP | none | rate_record.type with four values; ordered comes from commitments (Increment 3), actual from cost actuals (Increment 7). Increment 2 sets the type, later increments populate |
| Flag non-comparable records | GAP | none | comparable flag with reason. Increment 2 |
| Tender benchmarking across schemes | GAP | none | a query over rate_record by specification key; thin demo shows it on one scheme plus one seeded comparison row. Increment 2 |
| SQLite evaluated for ingestion and pilot | PARTIAL | D-04 has SQLite for dev and tests | Evaluation in DISC_04 section 5: adopted for Increments 1 to 7, single writer; PostgreSQL from Increment 8 |

## Area 2 - Scheme register, item ids, cost codes, WBS

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| One scheme register with stable scheme codes | PARTIAL | S04 project setup; projects table | scheme.code stable and unique per company; project name is an attribute; format to be confirmed (question Q3). Increment 1 |
| BOQ / activity schedule item ids | PARTIAL | S01 bill_lines, S04 import | boq_item.id per scheme revision; F4 "WJL ID" and F3 "Bill Item" are the sources. Increment 1 |
| Cost codes | PARTIAL | S01 bill_lines.cost_code (added 08-Sep) | cost_code becomes its own table (company chart) with allocations, not a column. Increment 1 |
| Programme WBS activity ids | GAP | none; no source data in the files | programme_activity table; manual entry or CSV import; no programme tool integration in Stage 1 (question Q2). Increment 1 (table), Increment 4 (site records link to it) |
| Explicit many-to-many allocations, no duplicated quantity or value | GAP | none | allocation rows (boq_item x cost_code x activity, share); a test asserts shares sum to 100 percent per item and that value is stored once. Increment 1 |
| Tender revision, accepted baseline, subsequent changes distinct | GAP | F5 Tender Log only | scheme_revision with kind = tender, baseline, change; boq_item and cost_budget hang off a revision. Increment 1 |

## Area 3 - Budgets, commitments, accruals, actuals, cash; reports

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Budgets by cost head including prelims | COVERED | S18, S19 | cost heads from F3 CVR template and F5 Handover; prelims is a head. Increment 7 (Increment 1 holds baseline budgets) |
| Procurement commitments | PARTIAL | S19 | commitments are purchase orders and subcontracts from Increment 3, not typed in. Increment 3 |
| Accruals | GAP | none | accrual rows with reversal on invoice match. Increment 7 |
| Actual costs | COVERED | S18, S19 (accounts export) | unchanged; source format unknown (question Q1). Increment 7 |
| Cash paid, distinct from actuals | GAP | S21 cash flow reads commitments only | cash_paid rows per supplier invoice. Increment 7 |
| Income: applications, certificates, receipts distinct | PARTIAL | S08 tracks applied, certified, paid | receipt becomes its own record with date and amount; retention and deductions separate. Increment 6 |
| Cost reports, CVRs | COVERED | S20 | CVR layout from F3 CVR template; reconciliation from Template 2. Increment 7 |
| Cash-flow forecasts | COVERED | S21, S22 | unchanged. Increment 7 |
| VOs, provisional sums, risks, early warnings in the reports | PARTIAL | VOs yes; the rest GAP | provisional_sum with replacement link; risk register; early_warning from the F4 EWN form. Increments 5 and 7 |
| Period snapshots preserved | GAP | none | period_snapshot freezes every report figure per period; reports read snapshots. Increment 7 |
| Prevent double counting across budgets, POs, invoices, accruals, payments | GAP | none | the cost ledger rules in DISC_03 section 5; tests per rule. Increment 7 |
| Prevent double counting of provisional sums and their replacement work | GAP | none | replacement work references the provisional sum it draws down; a test refuses value on both. Increment 5 |

## Area 4 - Applications and variation accounts

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Cumulative AFP / IPA from measured site progress, BOQ and VO accounts | COVERED | S07, S08, S14 | progress comes from site records (Increment 4) rather than typed percentages. Increment 6 |
| Unique scheme-linked VO codes | PARTIAL | S13 assigns a reference | code = scheme code plus sequence; format setting. Increment 5 |
| Instruction references, substantiation, status, progress on each VO | PARTIAL | S13, S14 | instruction record; substantiation = linked evidence rows; progress per VO line. Increment 5 |
| Submission, agreement, certification, payment tracked separately | COVERED | S08, S14 | unchanged; F4 VA Summary has both sides. Increments 5, 6 |
| Previous / current / cumulative | PARTIAL | S08 export | explicit on application and every line. Increment 6 |
| Retention and applicable deductions configurable | PARTIAL | S08 retention; contract_terms | deduction table (retention, MCD, rebate, discount, contra charge) with rules per scheme; F3 Template 2 and CVR template name them. Increment 6 |

## Area 5 - In-house joinery manufacture

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Item estimate and actual for materials, waste, hours, labour rate, design, sundries, shipping, storage, overhead allocation | EXCLUDED-BY-D-14 ("manufacturing optimisation" parked; estimating engine out) | none | Manufacture cost is not the estimating engine and not optimisation: it is cost capture per item. D-15 brings it in. manufacture_estimate and manufacture_actual per item with the heads from F4 VA Breakdown AB:AU and F5 Manuf. Increment 4 |
| Compare cost and value, explain variance | GAP | F1 0.4 and F3 3 Manuf PP do it by hand, and disagree | variance rows with a reason. Increment 4 (item), Increment 7 (scheme) |
| Bespoke reception desks, cladding, not doorsets alone | GAP | F5 DPM is doorset-only | manufacture_item is generic: any BOQ item or VO line with a "manufactured in-house" flag; DPM parameters are an optional attachment. Increment 4 |

## Area 6 - Suppliers, RFQs, quotes, orders, reconciliation

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Supplier and subcontractor directory with trade, geography, capability, performance | GAP | parties table holds project parties only | supplier table (a party role) with those attributes and a performance note per scheme. Increment 3 |
| Scheme item budgets for procurement | PARTIAL | S19 | budget per package from F5 Handover rows 51 to 104 and F3 PP planner. Increment 3 |
| RFQ preparation, draft only, no sending | GAP | none | rfq and rfq_line; export to a document; no email integration. Increment 3 |
| Like-for-like quote comparison | PARTIAL | S26 comparator (schedules and quotes) | comparator reads rfq_line and quote_line keyed by item; inclusions / exclusions compared. Increment 3 (first cut), S26 remains for schedule comparison |
| Selection, POs and subcontracts | GAP | none | commitment = purchase_order or subcontract with lines; feeds cost ledger and rate_record (ordered). Increment 3 |
| Delivery and invoice reconciliation | GAP | none | delivery and supplier_invoice rows matched to commitment lines; unmatched invoices flagged. Increment 7 (needs the accounts feed) |

## Area 7 - Phone site progress and quality capture

| Requirement | Status | Where today | Delta and increment |
|---|---|---|---|
| Low-effort phone capture of photos and documents | COVERED | S13 (PWA) for variations | generalised: progress_record on any item, not only a variation. Increment 4 |
| Linked to scheme, location, BOQ item, WBS activity, VO, instruction | PARTIAL | S13 links scheme and variation | all six links on progress_record. Increment 4 |
| Inspection and sign-off requests | PARTIAL | S28 sign-off record | inspection_request with the four states below. Increment 4 |
| Defects and rework | GAP | none | defect record linked to the inspection and the item. Increment 4 |
| Who / when, auditable | COVERED | history on every table | unchanged |
| Distinguish internal completion, submitted inspection, client acknowledgement, actual acceptance | GAP | none | four explicit states plus rejected; each with actor and time. Increment 4 |
| No response is not approval or payment entitlement | GAP | none | rule: only accepted feeds application progress; acknowledged and submitted show as claimed-not-accepted. Increment 4 (rule), Increment 6 (application) |
| Cloud evidence with project roles, version history, recoverable | PARTIAL | S03 roles; sources register; S05 backups | evidence rows with versions and hash; role check on read; restore test already in S05. Increment 4, Increment 8 for hosting |

## Consultant face

- D-01 keeps the consultant edition on the same records. The brief says consultant reuse remains useful and the contractor workflow drives Stage 1. Coverage: every increment's records are readable from the assessing side; consultant screens (S09, S10, S15, S32) are not on the Stage 1 critical path and are scheduled after Increment 8 unless Alaa's own use pulls one forward.

## Summary of deltas from D-14

| Delta | D-14 said | 09-Sep brief says | Recorded as |
|---|---|---|---|
| Estimating history and rate build-ups | estimating engine out | import its outputs and history; engine itself stays in Excel | D-15 line 1 |
| Manufacture cost per item | manufacturing parked | in scope as cost and variance capture | D-15 line 2 |
| Supplier directory, RFQ, PO, reconciliation | procurement not named | in scope, draft RFQs only | D-15 line 3 |
| Programme WBS and allocations | not named | in scope | D-15 line 4 |
| Accruals, cash paid, receipts, period snapshots, provisional sums, risks, early warnings | not named | in scope | D-15 line 5 |
| Four-state acceptance and no-response rule | not named | in scope | D-15 line 6 |
| Bookkeeping, sending RFQs, drawing reading, pipeline | out | still out (pipeline optional) | unchanged |
