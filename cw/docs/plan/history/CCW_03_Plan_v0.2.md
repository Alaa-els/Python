# CCW - Plan v0.2

- Prepared: 05-Sep-2026 (Chat C01, final framing; v0.1 added the proof and sales track, Section 10; v0.2 corrects file references only)
- Author: Alaa Elsayed
- Supersedes: every WJL_ and Gamma_ plan. Read with CCW_02_Blueprint (what), CCW_01_Idea_Filter (why), CCW_04 (terms with Gamma), CCW_05 (chat rules), CCW_06 (test data).

## 1. Purpose and success tests

- Purpose: build the Consultant Commercial Workbench for Alaa and Mo, module by module, each module used on Alaa's live work as soon as it passes its gate, and mirrored at Gamma for feedback.
- Month 2 (end Oct-2026): W1 payment assessment used on one live certification and on one Gamma application.
- Month 4 (end Dec-2026): W2 variation assessment and comparator in use; Gamma has run one variation through the mirror.
- Month 7 (end Mar-2027): W3, W4 and W5 in use; the monthly cost report comes from the database.
- Before Alaa's relocation (Jul-2027): everything runs on a fresh machine from the run-book, with contract-form settings proven on at least two forms.
- Sales test (after Phase 2): a demo to Tony that shows measured time saved and errors caught on Gamma's own submissions - Section 10.

## 2. Principles (locked unless Alaa reopens them)

- Serves Alaa first; Gamma is the mirror. A feature with no consultant use is not built unless paid for.
- One module at a time; each proven on live work before the next.
- CMD tools with Excel or Word outputs first; screens only after the tools prove their worth.
- Contract-form agnostic through settings; no market-specific code.
- No employer client data in the workbench; test with Gamma's data (with permission) and anonymised data.
- Existing tools and skills are absorbed, not rewritten.
- No SQL knowledge assumed: plain-language walk-through before any table; one real row per table.
- Review gates; Codex as independent check with its own brief; one deliverable per chat; handover every chat.
- Personal accounts, personal email, personal machine, own time.

## 3. Decision Log

| ID | Decision | Default | Owner | Status |
|---|---|---|---|---|
| D-01 | What is built and for whom | The workbench for Alaa and Mo; Gamma as test partner; Gamma-only features only if paid | - | LOCKED (Alaa, 05-Sep-2026) |
| D-02 | Terms with Gamma | Test-partner note in CCW_04, sent after the call | Alaa | OPEN |
| D-03 | Module order | W0 -> W1 -> W2 -> W3 -> W4 -> W5 -> W6 -> W7 | Alaa | DEFAULT - lock in C02 |
| D-04 | Test data | Gamma's real applications, variations, contracts, quotes with written permission; WJL templates; never WTP client data | - | LOCKED |
| D-05 | Stack | Python and SQLite now; PostgreSQL when shared; Excel and Word outputs; Telegram; Django later | Alaa, Mo | DEFAULT |
| D-06 | Intellectual property and employment | Check WTP contract for IP and outside-work clauses; own time, own machine; if WTP could claim IP on tools used at work, keep the workbench off WTP machines and use it only for outputs | Alaa | OPEN - blocks C03 |
| D-07 | Mo's role and ownership split | Equal owners; split of modules per Section 6 | Alaa, Mo | OPEN - agree in writing before C03 |
| D-08 | Gamma permission to hold extracts in a Claude Project | Ask on the call; until then only WJL templates and anonymised data | Alaa | OPEN |
| D-09 | Sinq | Not needed for the consultant side; note its price only if W6 grows into capture | - | DROPPED |
| D-10 | Contract forms to prove settings on | The forms Alaa works under now and the ones used in the market he is moving to | Alaa | OPEN - W3 |

## 4. Phases and steps

- Hours are Alaa and Mo combined with Claude, ranges not promises. Gate = what must be true before the next step.

### Phase 0 - Spine (W0)
| Step | What | Why | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|---|
| 0.1 | Close D-02, D-06, D-07, D-08; call with Tony using CCW_04 | Nothing built on open terms | Alaa | 2-3 | Decision Log updated | C02 |
| 0.2 | Walk-through part 1: projects, parties, contract_terms, bill_lines, rates, sources, history, actions, matters | No SQL knowledge assumed | Alaa + Mo | 3-4 | Both can explain every table | C03 |
| 0.3 | Walk-through part 2: submissions, applications, application_lines, variations, variation_lines, claims, correspondence; then the schema file and Codex brief | The build artefact | Mo runs | 3-4 | Schema loads; Codex closed | C04 |
| 0.4 | Project setup tool: create a project, enter terms, import a bill from Excel (mapped once), import rates; submissions register tool | First daily-use tools | Mo build, Alaa test | 8-12 | Gamma's F4 bill and a WTP-style bill both import | C05-C06 |
| 0.5 | Audit and sources: history on every change; sources register with file hash; "explain this figure" export | Traceability | Mo | 4-6 | A changed value shows who, when, old, new | C07 |

- Phase 0 total 20-29 hours, six chats.

### Phase 1 - Payment assessment (W1)
| Step | What | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|
| 1.1 | Rules specification: what is checked on every line (quantity against bill and previous, evidence reference, arithmetic, retention, discount, dates), what becomes a setting, what stays a judgement | Alaa | 3-4 | Mo reviews; D-18 rules expressed as settings | C08 |
| 1.2 | Application intake: read any application layout, mapping sheet once per contractor, load into applications and application_lines | Mo | 6-8 | Gamma's F4 and one anonymised WTP-style application load | C09 |
| 1.3 | Checks and assessment: run the rules, write assessed and certified values with reasons, flag missing evidence | Alaa logic, Mo build | 8-12 | One live certification reproduced | C10-C11 |
| 1.4 | Certificate workbook and tracker export in readable form; CMD commands; Codex brief | Mo | 6-8 | Passes workbook-readability; Codex closed | C12 |
| 1.5 | Gamma mirror run: Gamma's application pre-checked; feedback recorded | Alaa + Tony | 3-4 | Tony's feedback logged as changes | C13 |

- Phase 1 total 26-36 hours, six chats. Month 2 test sits here.

### Phase 2 - Variation assessment and comparator (W2)
| Step | What | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|
| 2.1 | Rules and layout specification: proposal intake fields, entitlement note, rate sources, tolerances, tracker columns | Alaa | 3-4 | Mo reviews | C14 |
| 2.2 | Comparator: any two or more schedules or quotes, mapping once, difference report by key and attribute | Mo | 8-12 | Finds one real error in Gamma's quotes; one WTP-style proposal comparison | C15-C16 |
| 2.3 | Assessment workbook from the database, Engineer's Assessment Tracker export, ageing alarms; Codex brief | Alaa logic, Mo build | 10-14 | One live assessment reproduced | C17-C18 |
| 2.4 | Gamma mirror run on one variation | Alaa + Tony | 2-3 | Feedback logged | C19 |

- Phase 2 total 23-33 hours, six chats. Month 4 test sits here.

### Phase 3 - Contract terms and rates (W3)
| Step | What | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|
| 3.1 | Terms checklist: the terms, their signal words, found / not found / ambiguous, disclaimer text; per-form variants (D-10) | Alaa | 3-4 | Mo reviews | C20 |
| 3.2 | Locator tool: PDF to text with OCR fallback, locate terms with page and clause, fill contract_terms settings; Codex brief | Alaa with Claude | 6-8 | Two real contracts, two forms | C21-C22 |
| 3.3 | Rate library: import Rate Library Browser data and Gamma's F5 Data Library with dates; ageing alarm | Mo | 4-6 | Alarm fires on stale rates | C23 |

- Phase 3 total 13-18 hours, four chats.

### Phase 4 - Correspondence, actions, claims (W4)
| Step | What | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|
| 4.1 | Outlook export into correspondence, matters, actions; chronology per matter; Telegram digest | Alaa (owns the tool), Mo | 10-14 | One live project's mailbox loaded; digest received | C24-C26 |
| 4.2 | Claims: events, notices, time-bar alarms; bundle built from sources and chronology | Alaa | 8-12 | One bundle produced from the database | C27-C28 |

- Phase 4 total 18-26 hours, five chats.

### Phase 5 - Reporting (W5) and mirror (W6)
| Step | What | Who | Hours | Gate | Chat |
|---|---|---|---|---|---|
| 5.1 | Employer cost report from the database: certified to date, change register, forecast final cost, cashflow, readable | Alaa spec, Mo build | 10-14 | Monthly report matches a manual one | C29-C31 |
| 5.2 | Mirror labels and a one-page guide for Gamma's pre-check use | Mo | 3-4 | Tony runs it unaided | C32 |
| 5.3 | Run-book, fresh-machine install, backups | Mo | 4-6 | Installs clean on a second machine | C33 |

- Phase 5 total 17-24 hours, five chats.

### Phase 6 - Screens (W7) - only if the tools have earned it
- Django login, project pages, guided intake screens. 40-60 hours, 8-10 chats. Decide after Phase 5.

## 5. Chat map summary

| Chats | Content | Attach at start |
|---|---|---|
| C02 | Call sheet from CCW_04; then Decision Log closed in Plan v1 | Handover, Plan v0.2, CCW_04, CCW_01 |
| C03-C07 | Spine walk-throughs, schema, setup tools, audit | Handover, Blueprint, schema when it exists, one bill extract |
| C08-C13 | Payment assessment | Handover, rules spec, schema, one application extract |
| C14-C19 | Variation assessment and comparator | Handover, rules spec, schema, one proposal or quote extract |
| C20-C23 | Terms locator and rates | Handover, checklist, one contract PDF |
| C24-C28 | Correspondence, actions, claims | Handover, Outlook export sample |
| C29-C33 | Reporting, mirror, run-book | Handover, schema, report layout |

- Rules: one deliverable per chat; a third round of fixes on one file means stop and hand over; legacy workbooks extracted once to CSV and never attached again; code lives in Git and the handover names the commit.

## 6. Split between Alaa and Mo (proposed - confirm in C02)

| Area | Alaa | Mo |
|---|---|---|
| Rules, checklists, assessment logic, readable outputs | Lead | Review |
| Walk-throughs and schema | Sign off | Sign off and run |
| Intake and mapping tools, comparator, imports, run-book | Specify and test | Build |
| Terms locator, correspondence and claims (Alaa's existing tools) | Build with Claude | Review |
| Gamma relationship and mirror runs | Lead | Attend |
| Codex briefs | Written by whoever built the piece; the other one reads the findings | - |

## 7. Effort and timeline

| Phase | Hours | Chats | Elapsed at 8-10 hours / week combined |
|---|---|---|---|
| 0 Spine | 20-29 | 6 | 3 weeks |
| 1 Payment | 26-36 | 6 | 3-4 weeks |
| 2 Variations and comparator | 23-33 | 6 | 3-4 weeks |
| 3 Terms and rates | 13-18 | 4 | 2 weeks |
| 4 Correspondence and claims | 18-26 | 5 | 2-3 weeks |
| 5 Reporting, mirror, run-book | 17-24 | 5 | 2-3 weeks |
| Total (without screens) | 117-166 | 32 | 15-19 weeks |
| 6 Screens (optional) | 40-60 | 8-10 | 5-6 weeks |

- Cash: £0. Everything runs on personal machines; Git free; Telegram free.
- If only Phases 0-2 happen (69-98 hours), Alaa still has payment and variation assessment on a database, which is most of the daily value.

## 8. Working rules for every chat (also in CCW_05)

- Bullets and tables only. British English. Hyphens, no em dashes. Expand acronyms on first use.
- One deliverable per chat, then handover naming the Git commit and the files to attach next.
- Walk-through before schema; rules specification before any assessment code; readable output before pilot.
- Every script: CMD commands plus a separate Codex review brief.
- Every output: workbook-readability, neutral voice, contractor statements as claims, defined terms capitalised, metadata author Alaa Elsayed.
- No client data. No WTP systems.

## 9. Risks and blind spots

| Risk | Impact | Mitigation |
|---|---|---|
| WTP claims IP over tools used at work (D-06) | Loss of the asset | Read the contract first; keep the workbench on the personal machine; only outputs go to work |
| Employer client data slips into test data | Confidentiality breach | D-04; anonymise; Gamma data only with written permission |
| D-18 rules hard-coded again | Works on one project only | Rules as per-project settings, tested on two forms (D-10) |
| Building screens too early | Months on a UI nobody needed | Phase 6 is decided after Phase 5 |
| Gamma feedback drives Gamma-only features | Scope drift | D-01; mirror is labels only |
| Mo's time or role unclear | Stalls | D-07 in writing |
| Relocation before the run-book exists | Tool dies with the laptop | Phase 5.3 fresh-machine install; Git |
| Existing tools rewritten instead of absorbed | Wasted hours | Principle; absorb by import, wrap, or call |
| Assessment outputs read as determinations | Professional risk | Status labels; neutral voice; working assumptions marked |

## 10. Proof and sales track (runs beside the build; costs almost nothing extra)

- Purpose: if Alaa decides to sell, the proof must already exist. Proof is measured, not described.

| Step | What | When | Hours |
|---|---|---|---|
| S.1 | Baseline at Gamma: how long its QS takes to prepare one application and one variation today; how many client queries or rejections the last three submissions drew; how many actions are chased by email | Ask on the C02 call; record in CCW_04 v1 | 1 |
| S.2 | Baseline for Alaa: time to certify one application and assess one variation by hand today | Before Phase 1 | 1 |
| S.3 | Measure after each mirror run: time, checks raised, errors caught, items with no evidence, actions created with owners | Phases 1 and 2 | 1 per run |
| S.4 | Demo script to Tony: one application and one variation of Gamma's go through the mirror live; show the source link on one figure, the history on one change, the actions list with owners, the certificate and tracker exports; then the before / after numbers | After Phase 2 (about month 4) | 3-4 |
| S.5 | Sales gate: if Tony wants Gamma-only features, quote them (former Gamma_06 quote structure: hours x rate, 40 / 30 / 30); if three subcontractors say yes at a stated monthly price, open a product track; otherwise stay a consultant tool with Gamma as mirror | Month 4-5 | 2 |

- The pitch, in Tony's terms: centralised (one database, one source per figure), accountable (every action has an owner and a due date, every change has a name), fewer mistakes (checks run on every line before anything is sent), less time (measured in S.3), and the client's QS sees a submission that already passed the checks the QS would run.
- Rule: the demo shows only what the consultant tools genuinely do. Promising Gamma-side capture screens, procurement or dashboards in the demo drags the build back to the parked list.
