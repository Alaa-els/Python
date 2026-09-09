# CW Risk Register

- Each risk shapes a stage; do not build past it. Add rows; never delete.

| ID | Risk | Handling | Stage |
|---|---|---|---|
| R-01 | Product work eats the consultant edition | Every module's handover states the other face's needs (W-07); consultant face built in the same module | all |
| R-02 | Gamma never pays after the pilot | D-02 terms; sales gate D-08 at S17 | S17 |
| R-03 | WTP IP or data clauses | D-06 before S01; Consultant face on Gamma test data unless permitted | S01 |
| R-04 | Data protection ignored with real contractor data | D-10 before S10; backups and deletion on request (S28) | S10 |
| R-05 | Site staff do not use it | Guided screens; one-page guide; Tony's site user reviews the S06 flow (S12 after the 08-Sep-2026 renumbering) | S12 |
| R-06 | Sinq or another vendor is cheaper and good enough | D-09 benchmark; the two faces are the difference. 08-Sep-2026: the status vocabulary (Raised, Pricing, Submitted, Agreed, Declined, Closed) is standard UK commercial usage and may be used; Sinq's screen layouts, wording and table columns are not copied | S17 |
| R-07 | No slack before July 2027 | Sales gate decides whether S18 to S29 happen | S17 |
| R-08 | Hard-coding one company's layout or one project's rules | Layouts as templates, terms as settings; tests assert no constants | S02 onward |
| R-09 | Existing tools rewritten instead of absorbed | Stage files name the tool to absorb; audit checks | S24, S29, S30 (S19, S24, S25 before the 08-Sep-2026 renumbering) |
| R-10 | Two builders, one codebase | One stage per branch; merge only at close; no direct commits to main | all |
| R-11 | Context overload in a session | Stage size limit (FOUNDATION section 4); split in the stage file | all |
| R-12 | Fixture figures typed from memory drift from the source | Fixture-derived literals only (FOUNDATION section 3); audit checks | all |
| R-13 | Actual costs are wrong or late because they come from the accounts package | The workbench never does bookkeeping; actuals arrive as an export mapped once (S18 names the format); every actual carries its import date; the CVR shows the actuals date beside the total | S18, S19 |
| R-14 | The plan runs past the July 2027 relocation | 231-320 hours at 8-10 a week ends between March and June 2027; no slack after Phase 6; the sales gate (D-08) is the point where Phases 3 to 8 shrink to Alaa's own needs if the product track does not open | S17 |
| R-15 | Programme WBS has no source in any file; allocations to activities are typed and drift from the real programme | programme_activity accepts CSV import; DISC_04 Q2 asks for the programme tool; allocations carry who and when | Increment 1 onward |
| R-16 | Stage 1 scope (seven areas) is built shallow for the demo and mistaken for done | DISC_04 keeps the thin demonstration and full Stage 1 depth in separate columns; every increment's full-depth column names what is still owed | Increments 1 to 8 |
