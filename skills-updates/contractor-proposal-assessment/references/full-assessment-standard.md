# Full cost assessment workbook - the standard settled on RFP-027 Rev 02 (Sep to Oct 2026)

This is the layout and method Alaa accepted after eleven review rounds on the TSE Irrigation Storage Tanks assessment (Contract QPMO-410-CT-05958). Build every Mode 2 assessment to it unless he says otherwise. It governs content and method; `wtp-excel-format` governs appearance and `workbook-readability` governs language.

## 1. Tabs, in this order, and nothing else

| Tab | State | What it holds |
|---|---|---|
| Project Info | visible, first | the twelve document-control fields |
| Cover, Appendix Cover | hidden | brand image plus the linked text boxes |
| Assessment | visible | the item list: SAMA Submitted Cost block, Assessed block, variance block, subtotal, OHP, total, four short notes below |
| Build-Up | visible | one block per item: banner, short introduction, the lines (Ref, Description, Unit, Qty, Rate, Amount, Basis), item total, Return link |
| Build-Up Comparison | visible | the Contractor's own components against the assessment, item by item; outstanding-information register (Section 7); the Contractor's programme documents as evidence (Section 8) |
| Programme | visible | only when time or resources drive quantities: status, calendar, windows, resource bridge, quantities table, link register |
| XER WBS | visible | the Contractor's programme file imported as received, never edited |
| Superseded Proposal (Ref) | visible | the earlier submission, reference only |
| Data | hidden | the validation lists |

No other tab. Comparisons and sensitivities go inside the existing tabs as grouped supporting rows, never as a new tab.

## 2. The reading path, and the links that carry it

Assessment item -> Build-Up item (quantity x rate, reason) -> Programme row (dates, people, plant) -> XER WBS activity (source). Every step is a working internal hyperlink; every tab has a Return link.

- Assessment: item numbers and assessed rates link to the Build-Up banner and total row.
- Build-Up: every quantity that is a formula links to the Programme row it reads; the total row carries "Return to 'Assessment' item n"; row 5 carries "Back to 'Assessment'" and "Forward to 'Programme'".
- Programme: a reference in column A opens the Build-Up line; a date or count opens its source row; a lookup by activity ID opens that activity on XER WBS; a count of several activities links to the first and lists every counted ID in the note; a count of nil has no link.
- Generic rule for every other formula: a formula that reads another tab or another row links to its first source; totals over ranges, same-row arithmetic and the fixed calendar cells are not linked; the Assessment totals carry no same-tab links.
- Links are blue, underlined, on the existing font. Stored as in-workbook locations only. LibreOffice rewrites them as external file links on recalculation: run `scripts/finalize_xlsx.py`, which restores them and strips the local path.
- Audit with `scripts/audit_links.py`: it checks every link against the activity ID or first referenced cell of the formula it sits on, and lists formulas with a cross-reference but no link. A link on a multi-input formula is navigation to one source; say so in the tab's how-to note.

## 3. Labels and vocabulary

- Column blocks: **SAMA Submitted Cost** (or the Contractor's name), **Assessed**, **Variance**. Programme columns: **SAMA Submitted Programme** beside **Assessed**. Never "carried", "Basis A/B", "as submitted" as a label, "departure".
- Status words at the point of use: "assessed allowance", "included in the supplier's price", "not included", "provisional", "unresolved", "outstanding". Codes (D1, H8, K6, P20b, activity IDs) stay as reference beside a plain name, never alone.
- Units are explicit: working days (Fridays excluded), calendar days, months of 30.4 calendar days, person-days, person-weeks, person-months at 26 working days. Never "man-days".
- **Never invite the Contractor to check the assessment.** "Rate: assessed allowance (Riyadh market, Sep-2026)" is complete. Not "needs confirmation by the Contractor", not "pending the Contractor's substantiation". Where a document is genuinely outstanding, state it as a fact: "the Contractor's method statement is outstanding".
- No author-facing words: "user-authorised", "on instruction", "to be recorded on Project Info", "mismatch stated plainly".
- Figures in notes must match the cells they describe (dates, counts, months). A note that quotes a figure is checked against the cell on every rebuild.

## 4. Notes are bullets, not essays

Every note longer than about 110 characters is a heading line plus two to five bullets (one sentence each, no full stop at the end). This applies to the item introductions in column A, the Basis column, the Programme notes, the Comparison comments and the Assessment reasons. Rate lists and point lists split on their semicolons. Row heights are recalculated from the wrapped lines. `scripts/bulletise_notes.py` does the mechanical pass; the item introductions are written by hand with a heading and one bullet per part (what, why, dates, provisional, excluded).

## 5. Method rules that survived review

- **Programme first.** Every time-related quantity (months of staff and facilities, plant days, hire windows) is a formula from the Contractor's programme: dates looked up by activity ID on XER WBS, periods computed with WORKDAY.INTL / NETWORKDAYS.INTL on the programme calendar, months by ROUND(days/30.4, 1). Changing a source date on XER WBS must change the total. Validate that on a temporary copy before delivery.
- **The submitted programme is reproduced as submitted** beside the assessed sequence. The assessed sequence keeps the submitted erection dates and fits in only what the Engineer has directed (for RFP-027, one-fill sequential testing). Never a compressed or "efficient" schedule, never unsupported idle or productivity deductions. External milestones (SAJCO readiness) are never advanced.
- **Approved resources are the baseline.** An approved manpower histogram is reproduced week by week unchanged and priced once: people in the week x the actual working days of that week = person-days. Departures are shown day by day (a week can balance while a day does not) as "Additional helpers assessed: needed 6, approved 3, 3 a day x 3 days = 9 person-days", with the crew size named as an assessment assumption. Spare capacity is unused, not deducted.
- **Supplier scope decides ownership.** What the supplier's quotation includes (installation, sealant, fixings) is not priced again; what it excludes (helpers, unloading and shifting, plant, scaffolding, power, storage, piping, test water, third-party testing) is the Contractor's and is priced once. Record the quotation reference, page and item numbers in an assumption; state plainly when the quotation in hand is not the adopted offer.
- **One rate per resource across the workbook.** A labourer is the same day rate on every line; a person-month is that rate x 26. Water is one basis (the Contractor's own haulage rate when no source is confirmed). Reviewers find rate contradictions first.
- **Nothing is zeroed merely because it might be in a supplier price.** Fair Contractor costs are included once; the duplication question is answered on the evidence and recorded.
- **Include fair omissions when found**: waste removal, interior cleaning before disinfection, operating water for commissioning, spray disinfection above the water line, laboratory turnaround days, statutory hours for watchmen, measured scaffold areas. Each with a one-line reason and an assessed rate.
- **Sensitivities stay out of the total** and sit in a labelled supporting block on the tab they belong to (programme sensitivities under Section 3, resource sensitivities under Section 5).
- **Outstanding information** is a numbered register on the Comparison tab (Section 7) with request, source and status. New items from each review are appended (OHP basis, water disposal, programme slippage, the adopted offer's conditions).
- **Changes from the previous revision** are one note on the Assessment tab, computed from the cells (net effect as a formula), naming the lines that moved and why.
- **Material qualifications are never silently resolved**: a commissioning method covering both tanks is stated with release conditions; a network refill or a valves-only demonstration is not assumed; insulation or specification price differences stay visible until the re-quote arrives.

## 6. Build and validation pipeline

1. Build with a script from the previous master (openpyxl); never hand-edit a delivered file when a script exists. When the client edits a delivered file by hand, post-process that file instead of rebuilding, and preserve its drawings (`scripts/restore_drawings.py`).
2. Recalculate with LibreOffice (`xlsx/scripts/recalc.py`); zero errors is the gate.
3. `scripts/finalize_xlsx.py`: restore in-workbook links, clean metadata (Alaa Elsayed, WT Partnership), rezip with `[Content_Types].xml` first.
4. `file-authorship-metadata/scripts/verify_metadata.py`, `workbook-readability/scripts/validate_workbook.py`, `scripts/audit_links.py`.
5. Independent recomputation: every line qty x rate = amount; item sums = item totals; Assessment J = Build-Up totals; K = I x J; subtotal; OHP; total; register check nil; Comparison item totals foot to the Assessment.
6. Render to PDF and look at the pages: no orphan rows, no header at a page foot, no page over the print budget (track row heights per page in the build; LibreOffice places automatic breaks on row heights, so budget from actual heights, not estimates), no 4 pt print (one page wide, as many pages tall as needed).
7. Deliver the file, two or three rendered previews, the total, what changed, what is still open, and whether the file is fit for issue or provisional. Never overclaim issue readiness.

## 7. Decisions that stay with Alaa

Rate levels and quantities are his calls once an argument is stated: report the argument and the effect, apply upward corrections with a good argument, hold reductions until he decides. Three recurring decisions to surface at the end: the wording "approved" versus "submitted" for a histogram, the supplier re-quote on a specification change, and the revision date before issue.
