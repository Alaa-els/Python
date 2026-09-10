# CW Stages

- One stage = one Claude Code session = one branch `stage/Snn` = one tag `Snn-done`. Gate = what must be true before the next stage opens. Hours are estimates from docs/discovery/DISC_05 (base case): A = agent active hours, H = human review and QA hours.
- Status: TODO, READY (stage file reviewed, W-10 as amended 09-Sep-2026), OPEN (session running), AUDIT, DONE, PARKED. Evidence status in the map: PROPOSED (planned), BUILT (exists on the branch), VERIFIED (reviewed by Codex or proven by a passing test).
- The stage file for each stage lives at docs/stages/Snn_<slug>.md and is written from docs/stages/_TEMPLATE.md before the session opens.
- Re-cut 09-Sep-2026 against docs/discovery/DISC_04 on Alaa's authorisation of incremental building under Codex management (D-15). Stages are grouped by increment; the previous 35-stage table of 08-Sep-2026 is superseded (map at the foot). Scope: D-15.

## Increment 1 - Spine and scheme register

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S00 | Repository bootstrap: Django project, settings from env, pytest, hygiene and docs tests, stage commands and auditor; test settings for both SQLite and PostgreSQL engines | - | `pytest` green, `manage.py check` green, tag S00-done | 2 / 1.5 | OPEN 10-Sep-2026 (Codex review as second review, W-10 line 09-Sep-2026; Mo's line blank) |
| S01 | Spine walk-through part 1 (docs only): DISC_03 sections 1, 2 and 9 with one example row per table | S00 | Alaa signs; Codex reviews | 1.5 / 2 | TODO |
| S02 | Spine walk-through part 2 (DISC_03 sections 3 to 8), then every model, migration, admin and history; constraints in the models; migrations run on both engines in the test suite | S01 | fresh SQLite and PostgreSQL both migrate; every table in admin; a change appears in history; a raw-SQL grep test passes | 3 / 2 | TODO |
| S03 | Login, roles (owner, editor, viewer, site), company membership, scheme membership, permission check on every view, home skeleton | S02 | two users on different schemes see different rows; cross-company access refused; a site role cannot open office screens | 2.5 / 1.5 | TODO |
| S04 | Scheme register with stable codes, revisions (tender, baseline, change), cost code chart, programme activities (typed or CSV), BOQ import mapped once, allocations; dump-and-load rehearsal SQLite to PostgreSQL on the fixture | S03 | fixture bill imports with cost codes; allocation shares sum to 100 percent per item and value is stored once; a superseded revision is read-only; the dump-load comparison test passes | 2 / 1 | TODO |

## Increment 2 - Estimating and rates

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S05 | Rate rules specification (docs only): the nine attributes, four rate types, specification key, comparable rules, inclusions | S04 | Codex reviews | 1 / 1 | TODO |
| S06 | estimate, estimate_item, rate_record, inclusion; import from a Data Library-shaped extract; benchmark query | S05 | fixture estimate loads; tendered and quoted rates for one specification key shown side by side; non-comparable row excluded from the average; every rate shows source scheme, date, revision | 3 / 2 | TODO |

## Increment 3 - Procurement

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S07 | Supplier directory (trade, geography, capability, performance notes), packages with budgets | S04 | a package carries its budget from the fixture handover; a supplier has the four attributes | 2 / 1.5 | TODO |
| S08 | RFQ draft and document export, quotes, quote lines, like-for-like comparison with inclusions | S07 | an RFQ exports and nothing is sent; two quotes compare by item; a quoted rate_record is created | 3 / 2 | TODO |
| S09 | Commitments (purchase order, subcontract) with lines; ordered rate_records; commitment shows as committed, never cost | S08 | selecting a quote creates a commitment; the ledger shows committed and zero cost | 2 / 1.5 | TODO |

## Increment 4 - Manufacture and site evidence

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S10 | Site flow specification (docs only): capture fields and six links, measured readings, verification, inspection states, evidence checks per contract, defects, acceptance never deemed, roles on site | S04; D-02 LOCKED (the Tony call) | Tony's site user and QS agree the flow; Codex reviews | 1.5 / 2 | TODO |
| S11 | Phone PWA: progress_record with photos and the six links; evidence with version and hash; offline form queue | S10 | from a phone browser a record with two photos is created against scheme, location, item, activity, VO and instruction | 3.5 / 2.5 | TODO |
| S12 | measured_progress derivation and internal verification; inspection_request with five states, actor and time per transition; configurable evidence_checks; defect and rework; acceptance never deemed | S11 | two records for one item and location supersede, never add, and cumulative installed never exceeds the item quantity; internal_complete to accepted with actor and time on each step; acceptance is refused without a client act; a rejection opens a defect | 3 / 2.5 | TODO |
| S13 | manufacture_item, manufacture_estimate, manufacture_actual, variance with reason; overhead allocation setting | S04 | a doorset and a non-doorset item each carry estimate, actual and an explained variance | 3 / 2 | TODO |

## Increment 5 - Variations

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S14 | instruction, variation with scheme-linked code, substantiation links to evidence | S12 | VO code is scheme code plus sequence; a VO with no instruction shows as claimed; substantiation lists records and photos | 2.5 / 2 | TODO |
| S15 | variation_line (selling build-up by F4 heads) and variation_cost_budget (internal, cost heads only) kept apart; four independent VO axes; provisional sum drawdown; early warning | S14, S06 | instruction, commercial, valuation and site completion move independently; work before price agreement is visible; the cost budget is seeded without overhead recovery or profit and enters the ledger once; a replacement item and its provisional sum never both carry value; an early warning links to a VO | 2.5 / 1.5 | TODO |
| S16 | VO statuses raised to closed, notice export in the company layout, ageing alarm | S15 | status flow enforced; export byte-identical on re-run; alarm fires on a stale fixture | 2 / 1.5 | TODO |

## Increment 6 - Applications

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S17 | application and application_line with previous, current, cumulative; line starts from measured_progress; evidence checks shown per line; QS override with reason; VO account carried in | S12, S16 | a blocking check stops the claim unless an override with a reason is recorded; a warning shows and lets the claim through; acceptance is never deemed; previous plus current equals cumulative on every line | 3 / 2 | TODO |
| S18 | deduction (retention, MCD, rebate, discount, contra charge) from contract_terms; certificate and receipt as separate rows | S17 | deductions computed from settings and shown as rows; certified and received entered separately; differences show | 2.5 / 2 | TODO |
| S19 | Application export in the F4-derived company layout | S18 | export matches the layout and re-runs byte-identical | 2.5 / 2 | TODO |

## Increment 7 - Cost ledger, periods and reports

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S20 | Cost rules specification (docs only): heads, budget, commitment, accrual, actual, cash paid, period close, the authority and reconciliation rules, the accounts export format | S09; Q1 answered for the real-data gate | Codex reviews; format named or a fixture format declared | 1.5 / 1.5 | TODO |
| S21 | cost_budget, accrual, cost_actual import mapped once (the authority), cash_paid, supplier_invoice and invoice_match as links not additions, credits, allocations; actual rate_records | S20, S09 | each authority and reconciliation rule (ledger v invoice, partial accrual reversal, partial payment and receipt, credit, allocation) has a failing-then-passing test; a re-import shows old and new in history | 3.5 / 2.5 | TODO |
| S22 | period, period_snapshot, forecast, risk register | S21 | closing a period writes a snapshot; a later change does not alter it; a stale forecast fires an alarm | 3 / 2 | TODO |
| S23 | CVR in the F3 CVR template heads with the Template 2 reconciliation; cash-flow forecast per scheme; demo scheme runs end to end | S22, S19 | CVR reconciles value to S19 and cost to the ledger; cash forecast reconciles to applications, receipts, commitments and terms; the demo passes | 3 / 2 | TODO |

## Increment 8 - Multi-user and hosting

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S24 | PostgreSQL cut-over with dump-load comparison, HTTPS, nightly backup, restore test, role-restricted evidence read | S23; hosting account | restore succeeds on a second machine; evidence read refused across companies; row counts and checksums match after cut-over | 3 / 3 | TODO |
| S25 | Company onboarding, plans, data export and deletion, terms of service, security review | S24; D-10 | second company onboards in ten minutes and sees nothing of the first; deletion proven | 2.5 / 2.5 | TODO |
| S26 | Guides per role, admin guide, support routine, run-book | S25 | a person not on the build follows the run-book to a running system | 1.5 / 1.5 | TODO |

## Increment 9 - Consultant face

| Stage | Deliverable | Depends on | Gate | A / H | Status |
|---|---|---|---|---|---|
| S27 | Application intake of any layout mapped once; line checks; evidence reference check | S19 | fixture and one anonymised consultant-style application load; every check fires on a seeded fault | 2.5 / 2 | TODO |
| S28 | Assessed and certified values with reasons; certificate export; assessment tracker | S27; D-06 for live use | certificate matches a hand-prepared one | 2.5 / 2 | TODO |
| S29 | Variation assessment: proposal intake, checks, assessed values, tracker | S16, S27 | fixture variation assessed; alarm fires | 2 / 1.5 | TODO |
| S30 | Consultant cost report to the Employer: certified to date, change register, forecast final cost, cash flow | S23, S28 | matches a manual report on the fixture set | 2 / 1.5 | TODO |

## Later, outside Stage 1

| Stage | Deliverable | Status |
|---|---|---|
| S31 | Email import to correspondence, matters, actions; digest (absorbs Outlook Workbench) | PARKED until Stage 1 closes |
| S32 | Claim pack and claim assessment bundle (absorbs Claims Bundle Builder) | PARKED until Stage 1 closes |
| S33 | Pipeline and tender log (F2, F5 Tender Log headings) | PARKED; only if D-08 opens the product track |

- Totals (base): 31 stages S00 to S30; agent active 75 hours; human review and QA 58 hours; ranges and calendar in DISC_05. The thin demonstration is S00 to S23.
- Sales gate: the results pack, price list and D-08 decision follow S23 as a docs-only package outside the stage numbering.
- Map from the 08-Sep-2026 table: old S00 to S04 are S00 to S04 (extended); old S05 is inside S24; old S06 to S11 are S17 to S19 and S27, S28; old S12 to S16 are S10 to S12, S14 to S16, S29; old S18 to S22 are S20 to S23; old S23 to S27 are S05, S06, S08 (comparator as quote comparison); old S28 is S11, S12; old S29, S30 are S31, S32; old S31, S32 are S23, S30; old S33, S34 are S24 to S26; old S35 is S33.
