# CW Decisions Register

- One register, edited in place. A decision changes only by adding a dated line under it; the old line stays. Git history is the delta chain.
- Status: LOCKED (build on it), DEFAULT (build on it unless overturned before the stage that needs it), OPEN (do not build on it), DROPPED.
- Product decisions carry the CW numbering. Working-method decisions are W-nn. The CCW to CW remap is D-13.

## Product decisions

| ID | Decision | Status | Owner | Log |
|---|---|---|---|---|
| D-01 | Two editions on one engine: the Contractor edition is the product small contractors pay for; the Consultant edition is the same records seen from the assessing side, for Alaa's and Mo's own work | LOCKED | Alaa | 05-Sep-2026 first locked as a consultant-only tool with Gamma as unpaid test partner. 05-Sep-2026 REOPENED and re-locked in the same session on Alaa's instruction: the paying customers are small contractors, not consultancies, so the product is the contractor-facing edition and the consultant edition rides on the same records. The first form is superseded, not deleted; see D-13 |
| D-02 | Gamma as launch customer: free six-month pilot from MVP go-live, then half list for 24 months (CW_04) | OPEN | Alaa with Tony | needs the call; C02 not yet run as at 08-Sep-2026 |
| D-03 | Build order W0 to W7 as in STAGES.md: Variations before Applications | DEFAULT | Alaa | 05-Sep-2026 changed from Payment-first to Variations-first when D-01 was re-locked: the contractor's pain starts at capture on site, and the application module consumes variation output. Lock or overturn at S00 close |
| D-04 | Stack: Python 3.12, Django 5, PostgreSQL in use, SQLite for dev and tests, HTMX, openpyxl, python-docx, Telegram and email, one small server | DEFAULT | Alaa, Mo | 05-Sep-2026 changed from CMD tools with SQLite when D-01 was re-locked: a sold product needs logins, roles, audit and phone-browser capture, which Django provides and CMD tools cannot. Lock or overturn at S00 close |
| D-05 | Test data: Gamma's with written permission, WJL templates, anonymised fixtures; never employer client data | LOCKED | - | 05-Sep-2026 locked. Until written permission from Tony exists, only the blank F4 template and anonymised fixtures may be used |
| D-06 | Alaa's WTP contract checked for IP, outside-work and data clauses; Consultant face on live WTP work only if permitted | OPEN | Alaa | blocks S01 |
| D-07 | Mo's ownership, hours and split in writing | OPEN | Alaa, Mo | blocks S01 |
| D-08 | Sales gate after the MVP pilot: measured results plus three contractors saying yes at a stated price before S18 onward continues at full pace | OPEN | Alaa, Mo | decides at S17 |
| D-09 | Pricing per company per month, tiered by users; Sinq is the benchmark | OPEN | Alaa | record price by S17 |
| D-10 | Hosting, data protection, one-page terms of service before the pilot | OPEN | Alaa, Mo | before S10 |
| D-11 | Contract forms: UK subcontracts for the Contractor face; Alaa's current forms for the Consultant face; settings prove both | OPEN | Alaa | by S18 |
| D-12 | Product name and legal vehicle | OPEN | Alaa, Mo | at S17 |
| D-13 | The CW document set (CW_01 to CW_06, CW_Handover_C01) supersedes the CCW set in full. CCW files are history: they are not placed in docs/plan/; if kept, they go in docs/plan/history/ and are never cited as current | LOCKED | Alaa | 08-Sep-2026. ID remap, CCW to CW: test data D-04 becomes D-05; stack D-05 becomes D-04; Gamma permission D-08 becomes a line under D-05; the sales gate is the new D-08. Any CCW decision id in an old document resolves through this line |

## Working-method decisions

| ID | Decision | Status | Log |
|---|---|---|---|
| W-01 | One stage per session, one branch per stage, tag on close; the protocol in FOUNDATION.md section 2 | LOCKED | 08-Sep-2026 |
| W-02 | The auditor subagent has no write tools; findings are recorded, never fixed silently | LOCKED | 08-Sep-2026 |
| W-03 | Hygiene is a test (tests/test_hygiene.py): ASCII, no en or em dashes, British spellings list, double-quoted strings, no library or AI strings in files or metadata | LOCKED | 08-Sep-2026 |
| W-04 | Exports are deterministic: one code path, byte-identical re-run under a frozen clock, proven by test | LOCKED | 08-Sep-2026 |
| W-05 | data/legacy/ is gitignored; fixtures are anonymised extracts committed under tests/fixtures/ | LOCKED | 08-Sep-2026 |
| W-06 | Steer by editing the stage file before the session; a premise that diverges from the delivered state is recorded, and the delivered state governs | LOCKED | 08-Sep-2026 |
| W-07 | Every stage handover states what the other face still needs | LOCKED | 08-Sep-2026 |
| W-08 | Codex review brief per stage at handovers/Snn_codex_brief.md; Codex findings enter FINDINGS.md like audit findings | LOCKED | 08-Sep-2026 |
| W-09 | No dependency is added unless the stage file names it | LOCKED | 08-Sep-2026 |
| W-10 | Each stage file is reviewed by Alaa and Mo before its session opens; the reviewer signs the stage file header. A stage file is not READY until both names and dates are present | LOCKED | 08-Sep-2026 |
| W-11 | A LOCKED decision is never rewritten in place. It is reopened by adding a dated line under it that states the change and the reason; the superseded form stays visible. The D-01 line above is the correction of the only breach to date | LOCKED | 08-Sep-2026 |
