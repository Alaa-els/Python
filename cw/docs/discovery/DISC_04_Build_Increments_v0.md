# DISC_04 - Build increments, acceptance gates, first demonstration, blockers

- Date: 09-Sep-2026. Author: Alaa Elsayed. Basis: DISC_01 to DISC_03; the stage protocol in docs/FOUNDATION.md.
- One increment is one or more stage sessions under W-01; increments are the reconciled plan, stages remain the unit of work. docs/STAGES.md was re-cut against this document on 09-Sep-2026 (Alaa's authorisation via Codex); stage numbers below refer to that re-cut.
- Two scopes are kept apart: the THIN DEMONSTRATION (one anonymised scheme end to end, every area touched, shallow) and FULL STAGE 1 (every area at working depth on real data). Neither parks a requested function.

## 1. The thin demonstration

- One anonymised scheme, fixture-derived: header and budgets from the F5 Handover structure, bill from the F4 Cumulative structure with the three-way value split, one procurement package, two manufactured items (one doorset, one non-doorset), a week of site records, one VO with a build-up in the F4 Variation Account Breakdown heads, one application, one closed period with a CVR and a cash forecast.
- Every figure in the fixture is invented and labelled sample; no figure, name or rate from the six files is used (D-05).
- Demonstrated in order: estimate and rates -> scheme register and allocations -> RFQ, quote comparison, order -> manufacture estimate and actual, site evidence and acceptance -> VO -> application -> cost ledger, CVR and cash forecast. The demonstration is complete when Increment 7 closes.

## 2. Increments, in order

| Inc | Scope | Stage sessions (est.) | Acceptance gate (each is a test unless marked) | Full Stage 1 depth beyond the demo |
|---|---|---|---|---|
| 0 | Discovery: this package | done | DISC_01 to DISC_04 on the branch; CHANGELOG entry; D-15 recorded | - |
| 1 | Spine and scheme register: bootstrap (S00 as written), walk-throughs for sections 1, 2 and 9 of DISC_03, models, login and roles, scheme with stable code, revisions, cost codes, programme activities, BOQ import mapped once, allocations | 5 (S00 to S04 as they stand, extended) | fresh SQLite migrates; the fixture bill imports with cost codes; allocation shares sum to 100 percent per item and value is stored once; two users on different schemes see different rows; a change appears in history; a superseded revision is read-only | real bills in Gamma's live layouts; programme CSV import |
| 2 | Estimating and rates: estimate, estimate_item, rate_record with the nine attributes and four types, inclusions, comparable flag, benchmark query | 2 | the fixture estimate loads from a Data Library-shaped extract; a tendered and a quoted rate for the same specification key are returned side by side; a non-comparable row appears but is excluded from the benchmark average; every rate shows source scheme, date and revision | history import across all of Tony's quotation workbooks (needs the files, Q4); region and supplier normalisation |
| 3 | Procurement: supplier directory, package budgets, RFQ draft and export, quotes, like-for-like comparison, commitment (PO or subcontract), commitment lines; ordered rate_records | 3 | an RFQ document exports and nothing is sent; two quotes compare by item with inclusions shown; selecting one creates a commitment; the commitment shows as committed and not as cost; an ordered rate_record exists | populated directory; performance notes; delivery and invoice matching (moves to Increment 7 with the accounts feed) |
| 4 | Manufacture and site evidence: manufacture_item, estimate and actual by head, variance; phone PWA capture of progress_record (measured, cumulative readings) with the six links, internal verification, measured_progress derivation, inspection_request with five states, defect, evidence with version and hash; configurable evidence_checks; acceptance never deemed | 4 | a doorset and a non-doorset item each carry an estimate and an actual with a variance reason; from a phone browser a record with two photos is created against a scheme, location, item, activity and instruction; the request moves internal_complete -> submitted -> acknowledged -> accepted with actor and time on each, and acceptance can only be recorded from an act of the client; two records for the same item and location supersede rather than add, so cumulative installed never exceeds the item quantity; a rejected request opens a defect | offline queue; role-restricted read of evidence in the hosted build (Increment 8) |
| 5 | Variations: instruction, variation with scheme-linked code, lines by head from the F4 breakdown heads, substantiation links, statuses, provisional sum drawdown, early warning | 3 | the VO code is scheme code plus sequence; a VO with no instruction shows as claimed; substantiation lists the site records and photos; instruction status, commercial status, valuation and site completion move independently; work recorded against a VO before its price is agreed is visible; the variation_cost_budget is seeded from cost heads only, never from selling lines with profit, and enters the ledger once; a replacement item and its provisional sum never both carry value; an early warning links to the VO | notice export in the company layout; ageing alarms |
| 6 | Applications: application with previous / current / cumulative on every line, progress from accepted inspection requests, VO account carried in, deductions from contract_terms, certificate and receipt as separate rows, export in the F4-derived layout | 3 | the fixture application's lines start from measured progress, show each evidence check's result, block on a failed blocking check unless a QS override with a reason is recorded, and never deem acceptance; previous plus current equals cumulative on every line; retention and MCD are computed from settings and shown as rows; certified and received are entered separately and differences show; the export re-runs byte-identical | Gamma's real application layout mapped once (needs the files, Q4) |
| 7 | Cost ledger, periods and reports: cost_budget, accrual, cost_actual import mapped once, cash_paid, invoice matching, forecast, period close and snapshot, risk register; CVR in the F3 CVR template heads with the Template 2 reconciliation; cash-flow forecast per scheme | 4 | the authority and reconciliation rules in DISC_03 section 5 (ledger row v invoice, partial accrual reversal, partial payment and receipt, credit, allocation, provisional sum, snapshot) each have a failing-then-passing test; closing a period writes a snapshot and a later change does not alter it; the CVR reconciles value to Increment 6 and cost to the ledger; the cash forecast reconciles to applications, receipts, commitments and payment terms; the demo scheme runs end to end | the real accounts export (Q1); company roll-up |
| 8 | Multi-user and hosting: PostgreSQL, HTTPS, backups and restore, role-restricted evidence, company onboarding, terms of service, run-book | 3 | restore succeeds on a second machine; a second company onboards and sees nothing of the first; evidence read is refused across companies | Gamma pilot on live data (D-05 tier 2 permission first) |
| 9 | Consultant face on the same records: intake, checks, assessment, certificate, cost report to the Employer | 4 | as S09, S10, S15, S32 today | Alaa's live use, subject to D-06 |

- Sessions: 31 stages S00 to S30 (STAGES.md as re-cut 09-Sep-2026). Hours are in DISC_05, which separates agent active hours from human review hours and withdraws the earlier 231 to 320 hour figure. Increments 1 to 7 are the demonstration: 24 stages.

## 3. How the existing stages map

| STAGES.md today | Increment | Change |
|---|---|---|
| S00 to S05 | 1 | S01 and S02 walk-throughs use DISC_03; S04 adds cost codes, allocations, revisions; S05 moves to Increment 8 |
| S06 to S11 (applications first) | 6 | applications come after site evidence and VOs, because the brief makes progress evidence and the VO account the inputs to the application. D-03 line added (09-Sep-2026): order within Stage 1 follows the data flow, not the earlier applications-first argument, which still holds for the consultant face |
| S12 to S16 (variations) | 4, 5 | capture is generalised to any site record; VO pricing follows |
| S17 sales gate | after Increment 7 | unchanged in substance; the demo is the results pack |
| S18 to S22 (costs, cash flow) | 7 | extended with accruals, cash paid, snapshots, provisional sums, risks |
| S23 to S27 (terms, rates, comparator) | 2, 3 | rate library becomes rate_record; comparator becomes quote comparison first |
| S28 to S30 (records, correspondence, claims) | 4 (records) | correspondence and claims stay as later stages after Increment 9 |
| S31, S32 (reports) | 7, 9 | |
| S33, S34 (platform) | 8 | |
| S35 (pipeline, optional) | unchanged, optional | |

## 4. First stage sessions to open (stage numbers per the 09-Sep-2026 re-cut of STAGES.md)

1. S00 as written (bootstrap). Blocked only by W-10: Mo's signature.
2. S01 walk-through, part 1, from DISC_03 sections 1, 2, 9. Blocked by D-06 and D-07 per the register (see section 6).
3. S02 walk-through, part 2, from DISC_03 sections 3 to 8, then models.
4. S04 extended: scheme register, revisions, cost codes, allocations, bill import. Gate as Increment 1.

## 5. Technology: no change proposed

- Django 5, HTMX, progressive web app for the phone: kept. Every requirement in the brief is a records-and-rules problem the stack already fits; the PWA handles camera, photos and an offline form queue without a native app.
- SQLite: adopted for Increments 1 to 7 (single writer, one scheme, imports, tests, Alaa's and Mo's machines). It does not solve multi-user cloud use: concurrent writes from site phones and office users need PostgreSQL, which is Increment 8 and already D-04.
- SQLite to PostgreSQL, how it is made safe (added 09-Sep-2026 after management review). The schema is Django models; migrations are engine-neutral and are run against BOTH engines in the test suite from S02 onward (a PostgreSQL service in the test run, or a local instance on the developer machine; the stage file names which). Constraints live in the models, not in the database by hand: unique keys (scheme code per company, VO code per scheme), check constraints (allocation share between 0 and 100, money not negative where the rule says so), foreign keys with PROTECT so a referenced row cannot vanish, and NOT NULL on every column the rules require. SQLite enforces these less strictly than PostgreSQL, so the test matrix is what proves them. No raw SQL and no SQLite-specific functions are allowed (a test greps for them). The cut-over is a dump from SQLite through the ORM (dumpdata) and a load into PostgreSQL (loaddata) followed by a row-count and checksum comparison per table; that comparison is a test in Increment 8, rehearsed in Increment 1 on the fixture so the path is exercised early, not at the end.
- Permissions and integrity are designed from Increment 1, not deferred: role and company checks on every view from S03; row ownership by company on every table from S02; the double-counting and status rules in DISC_03 are model-level validation, not report-time filters. "Few users now" changes the database engine, not the rules.
- Imports: openpyxl for Excel, csv for accounts exports; python-docx for the RFQ and notice documents. No new dependency until a stage file names it (W-09).
- Anything that would change this (a programme tool API, an accounts package API, a native app) is raised in DECISIONS.md before it is built.

## 6. Blockers and questions (only those that stop the next increment)

| Q | Question | Blocks | If unanswered |
|---|---|---|---|
| Q1 | Which accounts package produces the actual-cost and payments export, and in what column layout? The F1 and F3 external links point at an accounts-payable payments workbook and a finance forecast workbook; neither was supplied | Increment 7 real-data gate only | the demo uses a fixture export in a plain column layout; the mapping is redone once |
| Q2 | Is there a programme tool in use (one F3 tab links to a field-management product) and can it export activities as CSV? | WBS import; not Increment 1 | activities are typed in for the demo |
| Q3 | Scheme code format (the files use WJL### and GCS###) and the company display name (GEMMA or other) | Increment 1 settings | default: company prefix plus three digits; display name GEMMA, editable |
| Q4 | Real quotation workbooks, applications and VOs under the new company for full-depth Increments 2 and 6 | full Stage 1 depth, not the demo | the demo proceeds on fixtures |
| G1 | Register text, read exactly: D-06 "Alaa's WTP contract checked for IP, outside-work and data clauses; Consultant face on live WTP work only if permitted - OPEN - blocks S01". D-07 "Mo's ownership, hours and split in writing - OPEN - blocks S01". W-10 "Each stage file is reviewed by Alaa and Mo before its session opens; the reviewer signs the stage file header. A stage file is not READY until both names and dates are present". Substantive dependency remaining after Alaa's 09-Sep-2026 authorisation of incremental building under Codex management: D-06 governs Alaa's personal IP and employment exposure and the consultant face on live WTP data; it does not bear on building against fixtures, so it now blocks Increment 9 (consultant face on live work) and D-12 (legal vehicle), not S01. D-07 governs who owns what between Alaa and Mo; it blocks D-12 and any external release, not fixture-based building. W-10's purpose is a second review before a session opens; Codex review is that second review for now, recorded as such on each stage file; Mo's signature line stays blank until Mo signs and is never written by anyone else | nothing for Increments 1 to 7 on fixtures; Increment 8 external release and Increment 9 live use | dated lines under D-06, D-07 and W-10 record this (DECISIONS.md, 09-Sep-2026) |

- Not blockers, noted: the Tony call (D-02) blocks S06 under the current STAGES.md; under this plan the first customer-facing specification is Increment 4's site flow, and D-02 should block that instead. Recorded in the D-02 log line.
