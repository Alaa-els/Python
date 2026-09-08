# CCW - Idea Filter v0 (what serves Alaa, what is Gamma-only)

- Prepared: 05-Sep-2026 (Chat C01)
- Author: Alaa Elsayed
- Lens: does it serve Alaa's work as a cost consultant (Engineer / Employer side) now and in any future market, or the tools project Alaa runs with Mo? Gamma is a subcontractor - the mirror image of that work - so it is a test bed, not the customer.
- Verdicts: BUILD (serves Alaa), TWO-SIDED (same engine serves both; build Alaa's side first, Gamma's side is the mirror), PARK (Gamma-only or needs drawing AI, a mobile app or a data licence).

## 1. Tony's ideas, filtered

| # | Idea | Consultant's version of it | Verdict | Goes into |
|---|---|---|---|---|
| 1 | PQQ auto-fill | none | PARK | - |
| 2 | Estimating libraries | Rate library for assessment (contract, market, supplier rates with dates) - Alaa's Rate Library Browser already does the core | BUILD | W3 |
| 3 | Drawing-pack sorter | none in version 1 | PARK | - |
| 4 | Contract clause finder | Contract review is daily consultant work - retention, payment periods, notice periods, time bars, defects, damages | TWO-SIDED | W3 |
| 5 | Supplier quote comparison | Contractor proposal comparison - Alaa's comparison skill; supplier quotes are Gamma's use of the same engine | TWO-SIDED | W2 |
| 6 | Doorset sign-off photos | none | PARK | - |
| 7 | Variation capture and submission | Variation assessment - proposal intake, entitlement, rates, assessment, tracker; the capture form is Gamma's mirror | TWO-SIDED | W2 (assessment), W6 (mirror) |
| 8 | Claims preparation | Claims assessment - chronology, notices, bundle; Alaa's Claims Bundle Builder | TWO-SIDED | W4 |
| 9 | Record capture generally | none | PARK | - |
| 10 | Manufacturing optimiser | none | PARK | - |
| 11 | Invoice vs PO vs GRN | none for a consultant | PARK | - |
| 12a | Drawing register change detection | Checking contractor drawing issues against contract documents - scope-change evidence | LATER | W2 add-on |
| 12b | Drawing content diff | none in version 1 | PARK | - |
| 13 | Single source from the tender bill | The contract bill feeding certificates, variations and reports without re-keying - Alaa's ipc-template-generator already turns a bill into assessment tabs | BUILD | W0 spine |
| 14 | SME commercial dashboard | The consultant's monthly cost report to the Employer: certified to date, change register, forecast final cost, cashflow; Gamma's CVR is the mirror | BUILD | W5 |
| 15 | Leads pipeline from ABI | none | PARK | - |
| 16 | Harvest old mailbox addresses | Alaa's Outlook export tool does it | GOODWILL | 2 hours for Tony, outside the plan |
| 17 | Project inbox with action summary | Alaa's action system: emails to matters, actions, chronology, Telegram digest | BUILD | W4 |
| 18 | Doorset schedule checking | Same comparison engine as 5 | TWO-SIDED | W2 |
| 19 | Work done vs bill before QS sign-off | This is Alaa's job: inspection-request checking, application checking, certification | BUILD | W1 |
| 20 | Joinery take-offs | none in version 1 | PARK | - |

## 2. The six files, filtered as test data

| File | Use for Alaa | Verdict |
|---|---|---|
| F1 KPI report | none | PARK |
| F2 BD report | none | PARK |
| F3 Commercial project report and CVRs | The contractor's view of a project - useful later to test the Employer cost report against its mirror | LATER (W5) |
| F4 Application for Payment template | A real subcontractor application, variation summary and variation build-up to assess - primary test data | BUILD test data (W1, W2) |
| F5 Quotation workbook | Data Library rates with dates (W3 rate ageing); Letter and T&Cs (W3 terms); door and ironmongery schedules (W2 comparator); the estimating engine itself is parked | BUILD test data (W2, W3) |
| F6 Sinq PDF | Status flow reference for the variation mirror | reference only (W6) |

## 3. The earlier Gamma system modules, filtered

| Gamma module | Verdict | Becomes |
|---|---|---|
| M0 Spine with company-wide roles | Simplify: one or two users, audit kept, roles dropped to owner / editor / viewer | W0 |
| M1 Register and home | Keep: projects, submissions register, "your things today" | W0 / W5 |
| M2 Pipeline | PARK | - |
| M3 Order and handover | Keep terms capture and the terms locator; park budgets | W3 |
| M4 Procurement and rates | Keep comparator and rate library; park procurement plan | W2, W3 |
| M5 Site and variations | Keep the assessment side; capture form is the mirror | W2, W6 |
| M6 Application for payment | Keep as certification side | W1 |
| M7 Certification and cash | Keep: certified vs applied, retention, payment dates, notices | W1, W5 |
| M8 Costs and CVR | PARK (contractor's cost side) | - |
| M9 Reporting and alerts | Keep as Employer cost report and consultant alarms | W5 |

## 4. Alaa's existing tools and skills that plug straight in

| Existing | Plugs into | What changes |
|---|---|---|
| ipa-to-ipc-translation (D-18 rules) and ipc-template-generator | W1 | Rules move from code to per-contract settings so any project works, not only D-18 |
| WIR checking procedure | W1 | Becomes the evidence check on every application line |
| contractor-proposal-assessment (comparison sheets) | W2 | Reads from the database instead of from scratch each time |
| mass-grading-assessment (four-tab, zero hardcoded) | W2 | Pattern for any rate build-up assessment |
| Engineer's Assessment Tracker template | W2, W5 | Generated from the database, never typed |
| Rate Library Browser (NRM2, CESMM4, POMI, overlays) | W3 | Adds last-updated dates and ageing alarms |
| Outlook Workbench email export | W4 | Feeds matters, actions and chronology |
| Claims Bundle Builder | W4 | Builds bundles from the chronology and sources register |
| Action intelligence design (Telegram, matters and actions with ids) | W0, W4 | Becomes the actions spine of the workbench |
| workbook-readability, wtp-excel-format, file-authorship-metadata | every output | Unchanged |

## 5. What Gamma gets, and why Tony should still say yes

- The mirror: Gamma runs its own application or variation through the consultant's checker before sending it, and sees what a client's QS would strike out. Same code, different labels. Real value to Gamma at no extra build cost.
- The comparator for its doorset quotes and schedules.
- The contract terms locator for its subcontracts.
- Nothing else unless Gamma pays for it.
