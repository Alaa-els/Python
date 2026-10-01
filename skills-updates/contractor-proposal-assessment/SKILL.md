---
name: contractor-proposal-assessment
description: "Generate contractor proposal assessment and comparison sheets in Excel for commercial review by a cost consultant or quantity surveyor. Use this skill whenever the user mentions: assess this contractor proposal, compare contractor submissions, create a comparison sheet, prepare a cost assessment, bid comparison, contractor variation proposal assessment, update the assessment sheet, revise the comparison workbook, contractor cost proposal review, rate assessment, signage comparison, or any request to compare contractor pricing across revisions or against an independent assessment. Also trigger when the user uploads a contractor's Excel or PDF pricing schedule and asks for review, analysis, or comparison. This skill has two modes: Mode 1 (Comparison Sheet Generator) is the primary deliverable; Mode 2 (Full Cost Assessment Workbook) is secondary. Always trigger — even for simple contractor pricing queries — because the structured output is always more useful than an ad-hoc response."
---

# Contractor Proposal Assessment Sheet Generator

## Overview

This skill generates structured Excel workbooks for comparing contractor cost proposals across revisions and against the cost consultant's independent assessment. It follows RICS best practice principles (NRM, Black Book, Valuing Change guidance) and standard QS commercial management methodology.

## Modes

### Mode 1: Comparison Sheet Generator (Primary)
Produces a single-sheet multi-version comparison of a contractor's proposal items, tracking changes across submissions and the cost consultant's assessment side by side.

### Mode 2: Full Cost Assessment Workbook (Secondary)
Produces a multi-tab workbook with summary, detailed analysis tabs, bid tabulations, recommendation sections, and supporting schedules. Use only when the user explicitly requests a full workbook or the scope warrants it.

---

## MANDATORY PRE-QUESTIONS

**Before generating any file, the skill MUST ask and wait for answers:**

1. Is there an earlier contractor submission or revision before the current one in hand?
2. Which versions are available for comparison?
   - Initial contractor submission
   - Latest contractor revision
   - Your assessment
   - Other (specify)
3. Which file or sheet should be used as the base?
4. Should amounts be copied as provided, or recalculated?
   - Default: Contractor blocks = hardcoded (copied as submitted); Assessment block = formula-driven (Qty × Rate)
5. Should all revisions be shown, or only selected ones?

**Do NOT proceed until the user has answered. Do NOT assume only one contractor version exists.**

---

## MODE 1: COMPARISON SHEET STRUCTURE

### Fixed Left Columns (always present)

| Column | Content |
|--------|---------|
| Ref | Item reference number — match rows by Ref first, then Description |
| Description | Item description — preserve contractor's wording exactly unless user asks to edit |

### Version Blocks (repeated for each version)

Each version block MUST contain ALL FIVE of these columns in this exact order:

| Column | Rules |
|--------|-------|
| Unit | Never assume the same unit across versions. Repeat in every block. |
| Qty | Repeat in every block. May differ between versions. |
| Rate | Repeat in every block. |
| Amount | Contractor blocks: hardcoded (copied as submitted). Assessment block: formula = Qty × Rate. |
| Remarks | Must come immediately after Amount. Preserve contractor wording exactly. |

### Version Block Configurations

**If only one contractor submission exists:**
- Block 1: Contractor Submission
- Block 2: My Assessment

**If two contractor versions exist:**
- Block 1: Initial Contractor Submission
- Block 2: Latest Contractor Revision
- Block 3: My Assessment

**If more than two contractor versions exist:**
- Ask user which versions to display, unless instructed to include all
- Each version gets its own complete block with all five columns

### Assessment Block — Additional Columns

The assessment block (always the rightmost version block) includes these extra columns after Remarks:

| Column | Content |
|--------|---------|
| Basis of Rate | Rate derivation hierarchy (use drop-down from Data sheet) |
| Status | Assessment status (use drop-down from Data sheet) |

### Analysis Columns (after all version blocks)

| Column | Content |
|--------|---------|
| Variance (Amount) | = Latest contractor amount − Assessment amount |
| Variance (%) | = Variance ÷ Assessment amount (formatted as percentage) |

### Summary Section (below line items)

Below the last item row, include:

| Row | Content |
|-----|---------|
| Total – Contractor Proposed | Sum of contractor's latest Amount column |
| Total – Assessed | Sum of assessment Amount column |
| Variance (Amount) | = Proposed − Assessed |
| Variance (%) | = Variance ÷ Assessed |
| OHP / Commercial Adjustment | Separate row if contractor applies OHP, discount, or preliminaries adjustment — never bury in rates |

---

## DIFFERENCE HIGHLIGHTING

- Any value that **differs** between versions: **RED FONT** (RGB: 192, 0, 0)
- Values **identical** across all versions: normal font colour (black)
- Check differences in: Unit, Qty, Rate, Amount, and Remarks
- If an item exists in one version but not another: leave missing cells blank; show existing differing entry in red
- **Red font only.** Do NOT use red fill unless the user specifically requests it.
- Do NOT highlight the assessment block cells red just because they differ from the contractor — red is for tracking changes between contractor versions. The assessment is expected to differ.

---

## DATA SHEET (Mandatory — included in every workbook)

### Purpose
Holds master lists for drop-down selections, linked via data validation to relevant columns across the workbook. Ensures consistency, supports audit, and aligns with RICS commercial management practice.

### Structure
Each list occupies a named column range with a header. Lists are arranged side by side, each starting at Row 2 with a bold header in Row 1. Leave blank columns between lists for readability.

### Required Lists

#### 1. Status (Column A)
Aligned with RICS-style assessment stages for cost reporting and commercial management:

| Status | Meaning |
|--------|---------|
| Submitted | Contractor's proposal received, not yet reviewed |
| Under Review | Assessment in progress |
| Query Raised | Clarification or substantiation requested from contractor |
| Revised Submission | Contractor has resubmitted following query |
| Assessed | Cost consultant's assessment completed |
| Agreed | Rate/amount agreed between parties |
| Not Agreed | Rate/amount not agreed — dispute or further negotiation required |
| Included in Valuation | Agreed item included in interim or final valuation |
| Excluded | Item excluded from valuation (e.g., not instructed, out of scope) |
| Closed | Item fully resolved and closed out |

#### 2. Basis of Rate (Column C)
Rate derivation hierarchy per RICS Valuing Change guidance and standard QS practice:

| Basis | Meaning |
|-------|---------|
| BOQ Identical | Rate taken directly from the contract BOQ for the same item |
| BOQ Similar | Rate derived from a similar BOQ item with fair allowance for difference |
| Pro Rata | BOQ rate adjusted proportionally for dimensional or specification change |
| Star Rate | New rate built up from first principles (labour, plant, materials, OHP) |
| Market Tested | Rate verified against independent supplier/subcontractor quotation |
| Dayworks | Valued on the basis of time, materials, and plant records |
| Provisional Sum | Expenditure against a defined or undefined provisional sum |
| Lump Sum | Fixed price for a defined scope |
| Not Yet Determined | Rate basis not yet established — pending information |

#### 3. Assessment Outcome (Column E)
Commercial decision on each item:

| Outcome | Meaning |
|---------|---------|
| Accepted | Contractor's rate accepted as submitted |
| Not Accepted | Contractor's rate rejected — assessment rate applied |
| Accepted with Qualification | Rate accepted subject to stated conditions |
| Substantiation Required | Contractor must provide supporting breakdown |
| Deferred | Decision deferred pending further information or instruction |

#### 4. Contractual Basis (Column G)
Reference type supporting the assessed rate:

| Type | Example |
|------|---------|
| BOQ Item | e.g., BOQ Item 20301.1 |
| Contract Clause | e.g., Clause 13.3 |
| Specification Section | e.g., Spec Section 2.03 |
| CESMM4 Class/Rule | e.g., Class E Coverage Rule C4 |
| NRM Rule | e.g., NRM2 Section 3.3 |
| Instruction Reference | e.g., EVI-0045 |
| Correspondence | e.g., Letter Ref. 00102-AJE-xxx |
| Drawing Reference | e.g., Dwg No. D18-AR-001 |
| Supplier Quotation | e.g., Sign World Quot. 6512 |
| None | No specific contractual basis identified |

#### 5. Submission Type (Column I)

| Type |
|------|
| Initial Submission |
| Revised Submission |
| Final Submission |
| Formal Claim |
| Budget Estimate |
| Cost Proposal |

#### 6. Revision Type (Column K)

| Type |
|------|
| Rate Change |
| Quantity Change |
| Scope Change |
| Unit Change |
| Description Change |
| New Item Added |
| Item Omitted |
| No Change |

### Data Validation Setup
- Apply data validation (List type) to the relevant columns in all other sheets, referencing the named ranges in the Data sheet
- Named ranges should follow the pattern: `List_Status`, `List_BasisOfRate`, `List_AssessmentOutcome`, `List_ContractualBasis`, `List_SubmissionType`, `List_RevisionType`
- Data sheet should be formatted clearly but not locked — allow user to add entries

---

## FORMATTING REQUIREMENTS

### General
- Font: Arial, 10pt throughout (headers may be 11–12pt)
- Professional, clean layout suitable for printing in landscape A3
- Grouped headers for each version block with distinct background colours
- Alternating row shading (light grey/white) for readability
- Thin borders on all data cells
- Freeze panes: freeze the header rows and the Ref/Description columns

### Version Block Headers
- Contractor blocks: mid-blue header (RGB: 68, 114, 196)
- Assessment block: dark green header (RGB: 84, 130, 53)
- Analysis columns: navy header (RGB: 31, 56, 100)

### Version Block Sub-headers (column names)
- White font on coloured background matching the block header

### Row Heights
- Auto-fit to content; minimum 22px
- Description cells: wrap text enabled

### Number Formats
- Quantities: `#,##0.00`
- Rates: `#,##0.00`
- Amounts: `#,##0.00`
- Percentages: `0.0%`

### Date Format
- dd-MMM-yyyy (e.g., 15-Mar-2026) — consistent with user's preferred format

---

## WORKFLOW RULES

1. **Read the contractor's file first.** Extract all items, preserving Ref, Description, Unit, Qty, Rate, Amount, and Remarks exactly as submitted.
2. **Match items across versions.** Match by Ref first, then by Description if Ref is absent or ambiguous.
3. **Populate the comparison sheet.** Place each version's data into its respective block.
4. **Apply difference highlighting.** Compare each field across contractor versions. Apply red font to any cell that differs from another version of the same field for the same item.
5. **Leave the assessment block empty** unless the user provides assessed rates, or instruct the user to fill it in.
6. **Generate the Data sheet** with all master lists and apply data validation to relevant columns.
7. **Add the summary section** below the last item.
8. **Add a header block** at the top with: Project name, Contract reference, Date, Prepared by, Version/Revision number (if known).

---

## MODE 2: FULL COST ASSESSMENT WORKBOOK

**Build to `references/full-assessment-standard.md`.** That file is the layout, reading path, vocabulary, method and validation pipeline Alaa accepted on RFP-027 Rev 02 (Sep to Oct 2026) after eleven review rounds. Read it in full before the first cell is written. In short:

- Tabs, in order: Project Info, Cover (hidden), Appendix Cover (hidden), Assessment, Build-Up, Build-Up Comparison, Programme (only when time or resources drive quantities), XER WBS (the programme as received), Superseded Proposal (Ref), Data (hidden). No other tab; comparisons and sensitivities are grouped supporting rows inside these.
- Reading path with working hyperlinks and Return links: Assessment item -> Build-Up line (quantity x rate, reason) -> Programme row (dates, people, plant) -> XER WBS activity. Every cross-tab or cross-row formula links to its first source; nil counts and same-tab totals carry no link.
- Labels: SAMA Submitted Cost, SAMA Submitted Programme, Assessed, Variance. Never "carried", "Basis A/B", "departure", "man-days". Codes only beside a plain name.
- Notes are a heading plus two to five bullets, never an essay. Never invite the Contractor to check the assessment ("needs confirmation by the Contractor" is banned); state outstanding documents as facts.
- Time-related quantities are formulas from the Contractor's programme; approved resources are the baseline, priced once, with day-by-day departures shown as "Additional helpers assessed"; supplier scope decides ownership; one rate per resource; fair omissions included; nothing zeroed on a guess.
- Pipeline: build script -> LibreOffice recalc (0 errors) -> `scripts/finalize_xlsx.py` -> metadata and readability validators -> `scripts/audit_links.py` -> independent recomputation of the whole chain -> rendered page check -> deliver with previews, total, open items and an honest issue-readiness statement.
- Scripts in `scripts/`: `finalize_xlsx.py`, `audit_links.py`, `bulletise_notes.py`, `restore_drawings.py` (for a client-edited file, which must be post-processed, not rebuilt).

### Assessment tab fields
- Basis of Rate and Status (drop-downs from Data); the Assessed reason in plain words beside each item; four notes below the total: reading guide, preliminary status and issue-readiness condition, changes from the previous revision (computed), and nothing else.

## IMPORTANT RULES — ALWAYS ENFORCE

1. Never assume only one contractor version exists. Always ask.
2. Never merge Unit, Qty, Rate, Amount, or Remarks across version blocks.
3. Never recalculate contractor's Amount unless explicitly asked.
4. Always calculate assessment Amount as formula (Qty × Rate).
5. Preserve contractor's wording in descriptions and remarks.
6. Red font for differences between contractor versions only — not between contractor and assessment.
7. Include the Data sheet in every workbook.
8. Apply data validation drop-downs from the Data sheet to relevant columns.
9. Include version dating in block headers where the submission date is known.
10. OHP, discount, or commercial adjustments must be separate rows — never buried in unit rates.

---

## REFERENCES

- RICS Valuing Change Guidance Note (1st Edition) — rate hierarchy and valuation principles
- RICS NRM 2: Detailed Measurement for Building Works — measurement rules, BOQ structure
- RICS Black Book (QS & Construction Practice Information) — cost reporting, change management
- CESMM4 — civil engineering measurement rules (Class-specific coverage, definition, measurement rules)
- Contract-specific BOQ Preambles — project-specific measurement and valuation amendments

---

*Skill version: 1.1 | Created: March 2026 | Mode 2 standard added 01-Oct-2026 from RFP-027 Rev 02*
