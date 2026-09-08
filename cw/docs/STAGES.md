# CW Stages

- One stage = one Claude Code session = one branch `stage/Snn` = one tag `Snn-done`. Gate = what must be true before the next stage opens. Hours are Alaa and Mo combined with Claude, ranges not promises.
- Status: TODO, READY (stage file signed by both, W-10), OPEN (session running), AUDIT, DONE, PARKED.
- The stage file for each stage lives at docs/stages/Snn_<slug>.md and is written from docs/stages/_TEMPLATE.md before the session opens.
- Scope: the whole commercial process of a small contractor, from handover to cost report, on one database, with the consultant face on the same records (D-14, 08-Sep-2026). Phases 3 and 4 were added on that date; later phases renumbered (map at the foot of this file).

## Phase 0 - Spine, logins, roles (W0)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S00 | Repository bootstrap: Django project, settings, pytest, hygiene test, docs in place, commands and auditor | `pytest` green, `manage.py check` green, tag S00-done | 3-4 | TODO (Mo to sign, W-10) |
| S01 | Spine walk-through part 1 (docs only): companies, users, roles, project_members, projects with my_role, parties, contract_terms, obligations, bill_lines with cost_code, rates | Alaa and Mo sign the walk-through in the handover | 3-4 | TODO |
| S02 | Spine walk-through part 2 (submissions, applications, application_lines, variations, variation_lines, cost_budgets, cost_commitments, cost_actuals, forecasts, records, claims, correspondence, matters, actions, sources, history), then all models, migrations, admin, history | Fresh SQLite migrates; every table in admin; a change appears in history; model tests green | 6-8 | TODO |
| S03 | Login, roles, project membership, my_role switch, permission check on every view, home skeleton | Test: two users on different projects see different rows; cross-company access refused | 6-8 | TODO |
| S04 | Project setup screens, bill import mapped once with cost codes carried, submissions register, home "your things today" | F4-derived fixture bill imports; a project is created through the screens in a test | 8-12 | TODO |
| S05 | First server: PostgreSQL, HTTPS, nightly backup, restore test, run-book v0 | Restore from backup succeeds on a second machine | 5-7 | TODO |

## Phase 1 - Applications for payment (W2) - D-03 reversed 08-Sep-2026

- D-02 must be LOCKED before S06 opens. No module is specified before an external signal exists.

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S06 | W2 rules and layouts specification (docs only): what is checked on every line, what is a setting, what stays a judgement; the company layout template | Mo reviews; D-02 LOCKED; the F4 layout is mapped field by field | 4-5 | TODO |
| S07 | Contractor face part 1: progress entry per bill line and per schedule; totals reconcile to the bill; variation account held as one manually entered total line with a note and a settings flag (the Phase 2 seam) | Fixture application totals reconcile to the fixture bill; the seam is covered by a test that fails when S14 lands | 6-8 | TODO |
| S08 | Contractor face part 2: application export in the company layout; applied, certified, paid and retention tracking; certification and payment alarms | Export matches the F4 layout and re-runs byte-identical | 8-10 | TODO |
| S09 | Consultant face part 1: intake of any application layout mapped once; line checks against bill and previous certificate; evidence reference check | The fixture application and one anonymised consultant-style application both load; every check fires on a seeded fault | 6-8 | TODO |
| S10 | Consultant face part 2: assessed and certified values with reasons; certificate export; assessment tracker | Alaa certifies the fixture application in the tool and the certificate matches a hand-prepared one | 6-8 | TODO |
| S11 | Pilot at Gamma: two application cycles through both faces; fixes; measurements | Applied against certified tracked for real; D-10 done before this stage | 4-6 | TODO |

## Phase 2 - Variations, raised on the phone (W1)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S12 | W1 rules and screen-flow specification (docs only): project and cost code synced from the record, who raised it and on whose instruction, tag on creation, evidence (photos, drawing mark-ups, instruction emails) as rows in sources | Tony's site user and QS agree the flow | 4-5 | TODO |
| S13 | Contractor face part 1: raise a variation on a phone browser with photos, as a progressive web app on the Django stack (home-screen icon, camera, form works offline); statuses Raised, Pricing | Variation with two photos created from a phone browser session; status flow enforced | 6-8 | TODO |
| S14 | Contractor face part 2: pricing from bill and rates with overhead and profit settings; Submitted, Agreed, Declined, Closed; notice export; subcontractor pass-through; replaces the S07 manual variation-account line | The S07 seam test now passes against the real variation account; export byte-identical on re-run | 8-10 | TODO |
| S15 | Consultant face: proposal intake, checks, assessed values, tracker export; ageing alarms both faces | Alaa assesses the fixture variation in the tool; the alarm fires on a stale fixture | 10-14 | TODO |
| S16 | Pilot at Gamma: three real variations through both faces; fixes; measurements | Feedback and measurements recorded in the handover | 4-6 | TODO |

## Sales gate

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S17 | Results pack, price list, three contractor conversations, decision D-08 (docs only) | D-08 recorded LOCKED with the decision | 10-13 | TODO |

## Phase 3 - Costs and CVR (W8) - added 08-Sep-2026 (D-14)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S18 | Cost rules specification (docs only): cost heads (from the F3 CVR and F5 Handover and Budgets layouts), what a budget, a commitment and an actual are, where actuals come from (the accounts export, mapped once), how forecast to complete is set and by whom | Mo reviews; the accounts export format is named | 3-4 | TODO |
| S19 | Contractor face: budgets from handover per cost head; commitments (orders, subcontracts) entered or imported; actual costs imported from the accounts export mapped once; every row with source and history | A fixture job's budget, commitments and actuals load; a figure changed in a re-import shows old and new in history | 8-12 | TODO |
| S20 | Contractor face: CVR per job and internal cost report - value side from applications and variations, cost side from S19, forecast to complete, margin against tender; export in the company layout with a source beside every figure | The fixture CVR reconciles to its application and cost rows; byte-identical re-run | 8-12 | TODO |

## Phase 4 - Cash flow and forecasts (W9) - added 08-Sep-2026 (D-14)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S21 | Both faces: cash flow per job and company - money in from applications, certificates, payment terms and retention release; money out from commitments and their payment terms | The fixture set's cash flow reconciles to its applications and commitments | 6-8 | TODO |
| S22 | Contractor face: forecasts - final value, final cost and margin per job; turnover by month and order book for the company; every forecast row carries who set it and when | Forecast rows drive the company roll-up; a stale forecast fires an alarm | 5-7 | TODO |

## Phase 5 - Terms, obligations, rates, comparator (W3, W4)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S23 | Terms checklist per form and obligations rules (docs only) | Mo reviews | 3-4 | TODO |
| S24 | Terms locator: PDF to text with OCR fallback, page and clause references, settings filled, disclaimer | Two fixture contracts, two forms | 6-8 | TODO |
| S25 | Obligations calendar and alarms both faces | Alarms fire on fixture dates | 4-6 | TODO |
| S26 | Comparator: two or more schedules or quotes, mapping once, differences by key and attribute, readable report | Finds the seeded error in the fixture quotes | 10-14 | TODO |
| S27 | Rate library with dates, ageing alarm, use in pricing and assessment | Alarm fires; a variation priced from the library | 5-7 | TODO |

## Phase 6 - Records, correspondence, claims (W5)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S28 | Site diary and dayworks screens with photos; weekly progress record and client sign-off record from the same screens; records feed variations, claims and the weekly progress report | A week of fixture records entered through the screens in a test; one sign-off record exported | 10-14 | TODO |
| S29 | Email import to correspondence, matters, actions; Telegram and email digest (absorbs Outlook Workbench) | Fixture mailbox loads; digest produced | 10-14 | TODO |
| S30 | Claim pack (contractor) and claim assessment bundle (consultant) (absorbs Claims Bundle Builder) | One pack and one bundle from fixtures, byte-identical re-run | 8-12 | TODO |

## Phase 7 - Reports (W6)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S31 | Contractor dashboard and company KPI report: per job applied, certified, paid, retention, variations by status, cost to date, margin, cash flow, alarms; company turnover, gross profit, order book (the F1 headings, generated from the records) | Figures reconcile to the fixture set; Tony reads his position unaided | 12-16 | TODO |
| S32 | Consultant cost report: certified to date, change register, forecast final cost, cashflow | Matches a manual report on the fixture set | 8-12 | TODO |

## Phase 8 - Platform and handover (W7)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S33 | Company onboarding, plans, data export and deletion, security review, backups tested, terms of service | Second company onboarded in ten minutes; deletion proven | 10-14 | TODO |
| S34 | Guides per role, admin guide, support routine, run-book, both handover sets | Gamma runs a month unaided | 8-12 | TODO |

## Phase 9 - Pipeline and tender log - optional, only if D-08 opens the product track

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S35 | Enquiries, intake forecast, lost projects, tender log on the projects table (F2 and F5 Tender Log headings); feeds the order-book forecast in S22 | D-08 LOCKED with a yes; a fixture pipeline rolls into S22 | 6-10 | PARKED |

- Totals: 35 stages S00 to S34, 231-320 hours; S35 optional. MVP is S00 to S16: spine, applications both faces, variations both faces, two pilots.
- Calendar at 8-10 hours a week from 15-Sep-2026: MVP by mid-December 2026; S34 between March and June 2027. The relocation is July 2027; there is no slack after Phase 6 (R-14).
- Phase order reversed on 08-Sep-2026 (D-03). Applications carry the evidence and Alaa's existing tools; variations follow and replace the S07 seam.
- Renumbering on 08-Sep-2026 (D-14): S00 to S17 unchanged; old S18-S20 are now S23-S25; old S21-S22 are S26-S27; old S23-S25 are S28-S30; old S26-S27 are S31-S32; old S28-S29 are S33-S34. Decisions and risks that cite stages read through this map.
