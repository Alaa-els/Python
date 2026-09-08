# CCW - Consultant Commercial Workbench - Blueprint v0

- Prepared: 05-Sep-2026 (Chat C01)
- Author: Alaa Elsayed
- Purpose: define WHAT is being built so every chat builds the same thing. HOW and WHEN are in CCW_03_Plan; what was filtered out and why is in CCW_01_Idea_Filter.

## 1. Definition

- One database per consultant holding each project's contract terms, bill and rates, submissions received, applications and certificates, variations, claims, correspondence and actions, with tools that receive a contractor's submission, check it, assess it, certify or respond, and report - every figure traced to its source, every action tracked, every output fit to be forwarded.
- Owner: Alaa and Mo. First user: Alaa on live consultancy work. Test partner: Gamma, whose submissions are the mirror image of the work.

## 2. Who it serves

| User | Use |
|---|---|
| Alaa at WTP now | Receive, check, assess, certify, track, report - faster and traceable |
| Alaa in any future market | Contract-form settings per project, not code; works under any standard form |
| Alaa and Mo | Shared toolset, shared database when needed, reusable across employers and clients |
| Gamma (mirror) | Pre-check its own application or variation as a client's QS would; compare quotes and schedules; locate contract terms |

## 3. Design rules

- Every figure shows its source: record id, who entered it, when, and the file, tab and row if imported. "Explain this figure" lists its inputs.
- Money always carries a status: submitted, claimed, assessed, certified, agreed, instructed, paid, forecast.
- Every change logged: who, when, field, old, new. Nothing deleted; closed or superseded.
- Every output follows workbook-readability: first tab answers purpose, result, basis, action; neutral voice; no tactical commentary; no named ranges; no author-facing language. Assume every file reaches the other side.
- Contract-form agnostic: retention, payment period, certification period, notice periods, time bars, defects period, damages, discount live in a per-project settings table, never in code.
- Contractor statements are claims until substantiated or agreed; certified, instructed and jointly verified figures are labelled as what they are.
- No employer client data in the workbench during build and test. Test data is Gamma's (with permission) and anonymised.
- One tool, one job, run from CMD with an Excel or Word output first. Screens come later, and when they do: one question per screen, Next and Back, where you are and what is next.
- Alarms carry a reason and an action; thresholds sit in settings.

## 4. Spine (tables in plain words - detail in the walk-through chats)

| Table | One row = | Links to |
|---|---|---|
| projects | one contract Alaa is administering or assessing | parties |
| parties | one employer, engineer, contractor or subcontractor on a project | projects |
| contract_terms | one term for one project (retention percent, payment days, certification days, notice days, time bar, defects period, damages, discount) with clause reference | projects |
| bill_lines | one bill item with quantity, unit, rate, section | projects |
| rates | one rate (contract, market, supplier) with source and date | projects, bill_lines |
| submissions | anything received: application, variation proposal, claim, notice, letter, drawing issue - with date received, from, reference, file | projects, parties |
| applications | one application for one period: applied, previous certified, assessed, certified, retention, due dates | projects, submissions |
| application_lines | one line of one application: applied quantity and value, evidence reference, assessed, certified, reason | applications, bill_lines |
| variations | one variation: proposal, entitlement finding, assessed, agreed, instructed, status, dates | projects, submissions |
| variation_lines | one line of a variation build-up: submitted and assessed quantity, rate, amount, basis | variations, rates |
| claims | one claim: event, notices, time bar, assessed, status | projects, submissions |
| correspondence | one email or letter: from, to, date, subject, file, matter | projects, matters |
| matters | one issue that runs across correspondence and actions | projects |
| actions | one action: owner, due, origin (email, meeting, submission), status, chronology | matters, submissions |
| sources | one source document: file name, version, date, hash | everything |
| history | one change to any record | everything |

## 5. Modules

| Module | What it does | Absorbs | Test at Gamma |
|---|---|---|---|
| W0 Spine and register | Database, project setup, contract terms settings, bill and rates import, submissions register, actions, audit; "your things today" list | ipc-template-generator, action design | Gamma's contract, bill, F4 template |
| W1 Payment assessment | Application intake (any layout, mapped once), line checks against bill and previous certificate, evidence reference check, assessed and certified values, retention and due dates, certificate workbook, tracker | ipa-to-ipc rules as settings, WIR checking | Gamma's real applications, F4 |
| W2 Variation assessment | Proposal intake, comparison against contract rates and bill, quantity and rate checks, entitlement note, assessment workbook, Engineer's Assessment Tracker, ageing alarms; comparator for any two or more schedules or quotes | contractor-proposal-assessment, mass-grading pattern, EAT template | Gamma's variation build-ups, doorset quotes and schedules |
| W3 Contract terms and rates | Terms locator on a contract PDF (page and clause references, disclaimer), terms settings filled from it, rate library with dates and ageing alarms | Rate Library Browser | Gamma's subcontracts, F5 Data Library |
| W4 Correspondence, actions, claims | Outlook export to correspondence, matters, actions, chronology; Telegram and email digest; claims chronology and bundle from sources | Outlook Workbench, Claims Bundle Builder, action design | Gamma's project mailbox export |
| W5 Reporting | Employer cost report: certified to date, change register, forecast final cost, cashflow; tracker exports; alarms digest | readable workbook skills | Gamma's CVR as the mirror check |
| W6 Gamma mirror | Pre-check mode: run W1 and W2 rules on Gamma's own application or variation before it is sent; labels change, code does not | W1, W2 | Gamma users |
| W7 Screens (later) | Guided web screens for W0-W5 once the CMD tools are proven | - | - |

## 6. Alarms (consultant's, thresholds in settings)

| Alarm | Trigger | Action |
|---|---|---|
| Certification due | application received and certification period ends in N days | assess now |
| Payment due | certificate issued and payment date in N days, or passed | notify |
| Proposal ageing | variation proposal unassessed after N days | assess or request information |
| Time bar | claim notice period ends in N days | notify the parties |
| Evidence missing | application line with no inspection or record reference | request, or assess at zero with reason |
| Rate stale | rate older than N months used in an assessment | confirm |
| Gap needs reason | certified below applied with no reason recorded | record the reason |
| Action overdue | any action past due | owner |
| Retention release | release date within N days | check conditions |

## 7. Stack

| Layer | Now | Later |
|---|---|---|
| Database | SQLite, one file per consultant, zero cost | PostgreSQL when Alaa and Mo share one live database |
| Tools | Python from CMD, one tool per module | Django screens (W7) if the tools prove their worth |
| Outputs | Excel and Word through the existing readable-workbook skills | unchanged |
| Notifications | Telegram bot and email | unchanged |
| Code | Git repository owned by Alaa and Mo | unchanged |
| Cost | £0 | £5-15 / month if hosted |

## 8. Out of scope

- Everything marked PARK in CCW_01: estimating engine, procurement, manufacturing, pipeline, KPI tracker, credit checks, accounts matching, sign-off photos, drawing reading, take-offs, leads.
- Any feature whose only user is Gamma, unless Gamma pays for it.
- Company-wide role matrix; the workbench has owner, editor and viewer.
