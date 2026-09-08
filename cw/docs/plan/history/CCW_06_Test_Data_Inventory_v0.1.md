# CCW - Test Data Inventory v0.1 (Tony's legacy WJL / Gamma files)

- Prepared: 05-Sep-2026 (Chat C01; v0.1 removes questions settled by the locked decisions)
- Author: Alaa Elsayed
- Basis: the six files Tony shared. They are test data and reference layouts for the Consultant Commercial Workbench; CCW_01_Idea_Filter Section 2 says which are used (F4, F5, later F3) and which are parked (F1, F2). Sections 1-4 below are the file facts; the defects in Section 4 are things the workbench must not repeat. Earlier note: Tony's email (25-Aug-2026) shows the files are examples and test data for a list of tool ideas, not a brief to fix them. Sections 5 and 7 below are superseded by CCW_01_Idea_Filter_v0.md and are kept only as a record. Sections 1-4 and 6 remain valid.
- Purpose: tell Alaa and Mo what each file is, what data it holds, what is broken, and what the database must absorb from it.

## 1. Two businesses, one set of templates

- Files carry two brands: WJL Contracts Ltd (job numbers WJLxxx) and Gamma / GCS (quotation GCSxxx, "Gamma Prelims", "Gamma Attendances", Lithuanian "Gamyba" manufacturing tab).
- Settled: WJL is closed; Gamma is the test partner; the workbench is Alaa's and Mo's (CCW_03 D-01, D-04). The two brands only matter as test-data labels.

## 2. Inventory

| # | File (as received) | Owner (file metadata) | Last saved | What it is | Size / shape |
|---|---|---|---|---|---|
| F1 | 2026 WJL KPI Report Rev A (Latest).xlsx | Anthony Simms | 21-May-2026 | Monthly company KPI tracker built on a purchased HerDataStudio template | 15 tabs, 3 MB |
| F2 | Business Development Monthly Report - May 2026.xlsx | Jack Hodson | 13-May-2026 | BD pack: order intake, enquiries, intake forecast, turnover forecast, manufacturing hours, credit checks, lost projects, customer tiers, goals | 14 tabs |
| F3 | Updated Commercial Project Report + Procurement Plan Report.xlsx | Anthony Simms | 20-May-2026 | Portfolio commercial report: 2026/2025 project performance, manufacturing and design PP, per-job CVRs, per-job procurement plans, sales turnover, cashflow | 37 tabs, 1.7 MB |
| F4 | WJLxxx - AFP No (new template old dayworks rates).xlsx | Luke Brown (created by asimms 2018) | 23-Feb-2026 | Blank Application for Payment template: cumulative measure, variation account summary and breakdown, door / screen / ironmongery schedules, hidden Early Warning Notice | 11 tabs, 1.5 MB |
| F5 | GCSxxx - Quotation Internal Copy - Rev 0.xlsm | Anthony Simms | 17-Aug-2026 | Full estimating workbook: tender log, handover and budgets, letter, quotation, door / iron schedules, manufacturing, RFI, prelims, T&Cs, DPM, data library, attendances; contains macros (56 KB VBA, content not inspected) | 16 tabs, 6.6 MB |
| F6 | Sinq - Variation Process.pdf | Sinq (third-party product) | - | 9-slide walkthrough of a mobile-to-web variation workflow: site raises a VRF with photos, office reviews, QS converts VRF to a VO, VO upstreamed to client | 9 pages |

## 3. What each file holds (data the database must absorb)

### F1 - KPI Report
- KPIs tracked monthly: Turnover, Gross Profit, Net Profit, Orders In, Order Book, Wins and Issues, Manufacturing Hours, Job Costed Overheads (Design, Production Overtime, Production Costs, Commercial, Project Delivery), Head Office Overheads.
- Targets 2026: turnover £650k / month (£7.8m / yr), GP 22-23%, NP £21k / month, head office overheads £122k / month, order book aim 9 months.
- Actuals entered only for Jan-Feb 2026; ACTUALS tab has monthly rows from 2021 but mostly empty before 2024.
- 0.4 Manufacturing: per-job manufacturing hours (tender vs actual) and budget vs spend.
- Overheads breakdown by category with staff numbers by department.

### F2 - BD Monthly Report
- Order intake register by quarter from 2023 (project, customer, value, margin, dates, manufacturing hours secured).
- Enquiry KPIs by quarter (in / declined / priced / value priced).
- Intake forecast (project, customer, value, probability, margin, dates, hours) by quarter.
- Turnover forecast per project by month (rolling 12).
- Projected manufacturing hours per project per month.
- Credit checks per customer per month (risk score, credit limit).
- Lost projects register with reasons.
- Customer tier table with intake by year.
- Long-term goals by year.

### F3 - Commercial Project Report / Procurement Plan
- 2026 PP: per-job summary (job no., PM, QS, status, customer, manufacturing tender / variation / actual hours, material budgets vs spend, latest sales vs cost vs margin).
- 3 Manuf PP and 4 Design PP: per-job hours and budget trackers with monthly expected / actual.
- 2026 CVRs / 2025 CVRs: one CVR block per job (order value, omissions, additions, variations, QS adjustments, contra charges; cost budget vs forecast by category: material purchases, specialist materials, subcontractor, production labour, site prelims, design, manufacturing materials, transport, PM).
- One "PP" tab per job (578 PP, 576 PP ...): procurement planner per bill item (tender budget, target budget less 3%, variation budget, lead time, order date, supply only / supply and fix, internal / external manufacture, on order, delivered).
- 2A Sales Turnover and 2b Cashflow Forecast: monthly sales, GP and cashflow per job.

### F4 - AFP Template
- Cumulative: order breakdown per line (WJL ID, item, description, qty, unit, drawing, rate, total) split into Fixing / Specialist / Manufacturing value with progress %; summary block at rows 273-283 (measured works, variation account, application total; previous month, movement, applied vs certified).
- Variation Account Summary: VO no., description, qty, unit, rate, amount, claim %, date submitted, client SI no., commercial status (Instructed and agreed / Instructed only / Pending / Not required), agreement date; mirrored client-certified columns.
- Variation Account Breakdown: cost build-up per VO (site fixing hours, materials, specialist, manufacturing materials and labour hours, CAD, prelims, overhead recovery 21%, profit 5.5%).
- Door / Screen / Ironmongery schedules: 3,000+ rows each, 74-87 columns, with per-row claim % and totals.
- EWN (hidden): Early Warning / Delay Notice form.

### F5 - Quotation (estimating)
- Tender Log (revision history), Handover and Budgets (project info, contacts, budgets by cost head, handover notes).
- Quotation: 137-column pricing engine per item (site fixing, subcontract, materials, in-house manufacturing by sheet material, prelims, design, overhead recovery, profit) with checks and a cash-flow forecast block.
- Door Schedule, Iron Schedule, Schedule, Manuf. (Gamyba - Lithuanian headings), RFI, Prelim Breakdown, T&Cs, DPM (190-column door parameter model), Data Library (rates by category with supplier and last-updated date), Attendances.
- Macros present - purpose unknown. Irrelevant to the workbench: only the Data Library, Letter, T&Cs and schedules are read from this file, never the pricing engine.

### F6 - Sinq PDF
- Shows the workflow Anthony wants (or is evaluating): VRF raised on site with photos -> web review -> status (Pricing / Approved / Declined / Closed) -> convert to VO -> upstream to client.
- Settled: Sinq is dropped (CCW_03 D-09); the PDF is only a status-flow reference for the mirror.

## 4. Findings - what is broken (evidence)

| Ref | Finding | Evidence | Why it matters |
|---|---|---|---|
| X1 | Same figure lives in two files with different values | WJL556 total spend: £40,230.87 (F1 0.4 Manufacturing) vs £110,548.13 (F3 3 Manuf PP). WJL552: £45,698.35 vs £50,168.28. WJL551: £60,480.32 vs £61,498.97. Files saved one day apart. | Nobody can say which is true. This alone justifies one database. |
| X2 | Monthly reporting is lapsing | F1 actuals stop at Feb-2026; F2 tabs "Last Updated" range Mar-Apr 2026; file received Sep-2026. | Manual process is not being kept up - the tool must reduce entry, not add to it. |
| X3 | Same project re-keyed in 3-5 places | Oxford Uni Plot 2B appears in F2 Intake Forecast, Turnover Forecast, Projected Manufacturing Hours, Customer Tier; jobs re-typed in F1 and F3. | Triple entry = the source of X1. Project master list is the first table. |
| X4 | Template debris | F1 Summary still shows "AOV ($)", "Site traffic", "Mktg spend ($)"; CONTROL lists "Job Costed Overheads" five times; four overheads tabs (06, 06 (2), OLD 0.5, 05 Jan); cell note "this formula needs looking into". | The KPI file was never finished being adapted. Rebuild the report, do not patch it. |
| X5 | Broken links and errors | #REF! in F2 Turnover Forecast, F3 2A Sales Turnover and 2b Cashflow; #DIV/0! throughout F1, F3, F4. F3 "Link to CVR" column = False on every job. | Reports issued with errors in them. |
| X6 | One tab per job, inconsistent | F3 per-job PP tabs vary 26-53 columns; multi-job tabs ("547, 546, 544"); Template 2 vs PP Template. | Cannot aggregate; every new job is a manual copy. |
| X7 | Version confusion in file names | F4 named "new template old dayworks rates"; F3 "Updated ..."; F1 "Rev A (Latest)". | No revision control. |
| X8 | Schedules too big for Excel by hand | F4 door schedule 3,487 rows x 78 cols; F5 DPM 190 cols; F5 Quotation 137 cols. | Row-per-door data belongs in a table, with Excel as the view. |
| X9 | Two entities share templates | WJL vs Gamma / GCS branding across F4 and F5. | Entity must be a field, not a file. |
| X10 | Macros of unknown purpose | F5 contains 56 KB of VBA. | Risk of silent logic; must be documented before replacing. |

## 5. Pain points - SUPERSEDED by CCW_01 (kept for record)

| Ref | Hypothesis | Files | Confidence |
|---|---|---|---|
| P1 | Cannot trust the numbers - files disagree | F1, F3 | High (X1) |
| P2 | Monthly KPI and BD packs take too long and fall behind | F1, F2 | High (X2) |
| P3 | Variations raised on site are lost or slow to reach the AFP | F4, F6 | Medium (Sinq deck is the tell) |
| P4 | Estimating -> handover -> budgets -> CVR -> AFP is copy-paste, so budgets drift | F5, F3, F4 | Medium |
| P5 | No single project register (job, customer, PM, QS, status, dates, values) | all | High (X3) |
| P6 | AFP door / ironmongery schedules are painful to build and re-issue monthly | F4 | Medium |
| P7 | Wants a Sinq-style site app but cheaper or integrated | F6 | Unknown - ask |

## 6. Out of scope for version 1 (recommended)

- Rebuilding the estimating engine (F5). It is a working pricing system with macros; version 1 should consume its outputs (Handover and Budgets) and nothing more.
- Replacing the accounts system. Actual costs and turnover must come from wherever finance keeps them (Sage / Xero / other - VERIFY). The database stores what is imported, it does not do bookkeeping.
- Social media, meetings, long-term goals tabs in F2.

## 7. Questions for Anthony - SUPERSEDED by CCW_04 Section 6 (kept for record)

- Which entity or entities: WJL only, Gamma / GCS only, or both?
- Rank the seven pain points P1-P7. Which one, if fixed alone, saves the most time or money?
- Who updates F1, F2, F3 today, how many hours per month each, and where do actual costs and turnover come from?
- Sinq: quoted or subscribed? Price? What is missing from it?
- Who at WJL will own the tool after handover? Excel-only users, or anyone comfortable with a web app?
- What is the Microsoft 365 licence level (SharePoint / Power Apps / Power BI available)?
- May the files be stored in a Claude Project (confidentiality: rates, customer credit scores, margins)?
- What do the macros in F5 do?
