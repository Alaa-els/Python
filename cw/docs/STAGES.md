# CW Stages

- One stage = one Claude Code session = one branch `stage/Snn` = one tag `Snn-done`. Gate = what must be true before the next stage opens. Hours are Alaa and Mo combined with Claude, ranges not promises.
- Status: TODO, READY (stage file signed by both, W-10), OPEN (session running), AUDIT, DONE, PARKED.
- The stage file for each stage lives at docs/stages/Snn_<slug>.md and is written from docs/stages/_TEMPLATE.md before the session opens.

## Phase 0 - Spine, logins, roles (W0)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S00 | Repository bootstrap: Django project, settings, pytest, hygiene test, docs in place, commands and auditor | `pytest` green, `manage.py check` green, tag S00-done | 3-4 | TODO |
| S01 | Spine walk-through part 1 (docs only): companies, users, roles, project_members, projects with my_role, parties, contract_terms, obligations, bill_lines, rates | Alaa and Mo sign the walk-through in the handover | 3-4 | TODO |
| S02 | Spine walk-through part 2, then all models, migrations, admin, history | Fresh SQLite migrates; every table in admin; a change appears in history; model tests green | 6-8 | TODO |
| S03 | Login, roles, project membership, my_role switch, permission check on every view, home skeleton | Test: two users on different projects see different rows; cross-company access refused | 6-8 | TODO |
| S04 | Project setup screens, bill import mapped once, submissions register, home "your things today" | F4-derived fixture bill imports; a project is created through the screens in a test | 8-12 | TODO |
| S05 | First server: PostgreSQL, HTTPS, nightly backup, restore test, run-book v0 | Restore from backup succeeds on a second machine | 5-7 | TODO |

## Phase 1 - Variations (W1)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S06 | W1 rules and screen-flow specification (docs only) | Tony's site user and QS agree the flow | 4-5 | TODO |
| S07 | Contractor face part 1: raise a variation on a phone browser with photos; statuses Raised, Pricing | Test: variation with two photos created from a browser session; status flow enforced | 6-8 | TODO |
| S08 | Contractor face part 2: pricing from bill and rates with overhead and profit settings; Submitted, Agreed, Declined, Closed; notice export in company layout; subcontractor pass-through | Export byte-identical on re-run; matches Gamma's F4 variation layout | 8-10 | TODO |
| S09 | Consultant face: proposal intake, checks, assessed values, tracker export; ageing alarms both faces | Alaa assesses the fixture variation in the tool; alarm fires on a stale fixture | 10-14 | TODO |
| S10 | Pilot at Gamma: three real variations through both faces; fixes; measurements | Feedback and measurements recorded in the handover; D-10 done before this stage | 4-6 | TODO |

## Phase 2 - Applications (W2)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S11 | W2 rules and layouts specification (docs only) | Mo reviews; D-18-style rules expressed as settings | 4-5 | TODO |
| S12 | Contractor face part 1: progress entry per line and per schedule; variation account carried in | Fixture application totals reconcile to the bill | 6-8 | TODO |
| S13 | Contractor face part 2: application export in the company layout; applied, certified, paid, retention tracking; alarms | Export matches Gamma's F4 layout; byte-identical re-run | 8-10 | TODO |
| S14 | Consultant face part 1: intake of any layout mapped once; line checks; evidence check | Fixture application and one anonymised consultant-style application both load; checks fire | 6-8 | TODO |
| S15 | Consultant face part 2: assessed and certified with reasons; certificate export; tracker | Alaa certifies the fixture application in the tool | 6-8 | TODO |
| S16 | Pilot: two application cycles at Gamma; measurements | Applied against certified tracked for real | 4-6 | TODO |

## Sales gate

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S17 | Results pack, price list, three contractor conversations, decision D-08 (docs only) | D-08 recorded LOCKED with the decision | 10-13 | TODO |

## Phase 3 - Terms and obligations (W3)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S18 | Terms checklist per form and obligations rules (docs only) | Mo reviews | 3-4 | TODO |
| S19 | Terms locator: PDF to text with OCR fallback, page and clause references, settings filled, disclaimer | Two fixture contracts, two forms | 6-8 | TODO |
| S20 | Obligations calendar and alarms both faces | Alarms fire on fixture dates | 4-6 | TODO |

## Phase 4 - Comparator and rates (W4)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S21 | Comparator: two or more schedules or quotes, mapping once, differences by key and attribute, readable report | Finds the seeded error in the fixture quotes | 10-14 | TODO |
| S22 | Rate library with dates, ageing alarm, use in pricing and assessment | Alarm fires; a variation priced from the library | 5-7 | TODO |

## Phase 5 - Records, correspondence, claims (W5)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S23 | Site diary and dayworks screens with photos; records feed variations and claims | A week of fixture records entered through the screens in a test | 10-14 | TODO |
| S24 | Email import to correspondence, matters, actions; Telegram and email digest (absorbs Outlook Workbench) | Fixture mailbox loads; digest produced | 10-14 | TODO |
| S25 | Claim pack (contractor) and claim assessment bundle (consultant) (absorbs Claims Bundle Builder) | One pack and one bundle from fixtures, byte-identical re-run | 8-12 | TODO |

## Phase 6 - Reports (W6)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S26 | Contractor dashboard: applied, certified, paid, retention, variations by status, cashflow, alarms | Figures reconcile to the fixture set; Tony reads his position unaided | 12-16 | TODO |
| S27 | Consultant cost report: certified to date, change register, forecast final cost, cashflow | Matches a manual report on the fixture set | 8-12 | TODO |

## Phase 7 - Platform and handover (W7)

| Stage | Deliverable | Gate | Hours | Status |
|---|---|---|---|---|
| S28 | Company onboarding, plans, data export and deletion, security review, backups tested, terms of service | Second company onboarded in ten minutes; deletion proven | 10-14 | TODO |
| S29 | Guides per role, admin guide, support routine, run-book, both handover sets | Gamma runs a month unaided | 8-12 | TODO |

- Totals: 30 stages, 200-277 hours. MVP is S00 to S16.
