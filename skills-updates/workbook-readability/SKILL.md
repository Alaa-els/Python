---
name: workbook-readability
description: Make every Excel workbook built for Alaa Elsayed readable by someone who did not build it. Governs what the first visible tab must answer, how tabs are layered into conclusions / working / hidden helper, when named ranges are allowed, how figures are traced to their source, how uncertainty is labelled, and the neutral voice required in anything that may be issued to a client, Engineer or Contractor. Applies to every .xlsx or .xlsm deliverable - cost assessments, FFC forecasts, IPC workbooks, rate build-ups, comparison sheets, claim assessments, trackers, variation assessments - alongside wtp-excel-format (appearance) and file-authorship-metadata (metadata). Trigger on any request to build, rebuild, restructure, simplify or review a workbook, and on complaints that a sheet is hard to follow, too detailed, over-engineered, or reads like the author talking to himself.
---

# Workbook Readability

`wtp-excel-format` decides how a workbook looks. This decides whether anyone can understand it.

Governs structure and language only. Applying it must never move a number.

## The test

> Could a competent person who has never seen this file explain its conclusion to someone else after two minutes on the first tab?

## The seven instructions

**1. Confirm the purpose.** One sentence: what must this workbook help the reader decide? Hold it for the whole build.

**2. Decide one-off or recurring, before designing.** Ask if it is not obvious. A monthly IPC or quarterly FFC justifies machinery a one-off assessment does not. Do not build reusable apparatus for a single deliverable, and do not hand-build a recurring one.

**3. Make the first substantive visible tab answer the reader's questions.** A document-control cover may precede it. Purpose, result, the basis or qualification that changes how the result is read, and the action required. Use only the sections that workbook needs - see `references/first-tab-examples.md`.

**4. Keep the visible workbook small.** (A workbook carrying one tab per BOQ item, as `ipc-template-generator` produces, is the recognised exception.) Conclusions first, supporting calculation next, genuine helper tabs hidden. A tab that exists so the calculation can be general, rather than so a finding can be communicated, is a helper tab. Hide, never delete - audit evidence must survive.

**5. Calculate each material figure once.** Every other appearance references that cell. Restated figures are future contradictions: the next revision updates one and misses the other.

**6. Label uncertainty at the point of use.** Claimed, contractual, certified, verified, assessed, estimated, unestablished - stated where the figure appears, not only in a register. Zero must never silently mean "unknown".

**7. Remove author-facing language.** No model architecture, no internal strategy, no unexplained codes, no instructions describing how the workbook was built.

## Traceability

The failure to avoid: a reviewer clicks a headline figure, sees `=Qty_ShopBasisTotalRC`, and must open Name Manager to find out where it was worked out.

**Named ranges are for document-control fields and data-validation anchors only, and only at workbook scope.** The approved list is exact, and it is the union of what the other skills require: the twelve Project Info names created by `wtp-excel-format`, which the Cover and Appendix Cover textboxes link to, and the six `List_` validation anchors mandated by `contractor-proposal-assessment`. Never name a rate, an array, a calculated range or a headline total. Add a further exact name only when another workbook feature genuinely requires one - never a general prefix.

Three mechanisms instead:

1. The formula names its own location: `='Calculation Engine'!$AC$49`.
2. A source sentence sits beside every material figure, in words: *"Calculation Engine row 49; total of RW-01, RW-02 and RW-04 on Contractor sections"*.
3. `Ctrl` + `[` and Trace Precedents then work. Both are defeated by names.

A data dictionary is not an answer to a traceability complaint. It is a second document explaining why the first is hard to read. Remove the names.

## Labelling uncertainty

**Totals may combine components of different status where the composition is disclosed.** An FFC properly combines committed values, assessed variations, provisional sums and risk allowance. The requirement is not to prevent the total - it is to show what is in it and how firm each part is. What is never acceptable is a total whose components look equally certain when they are not.

**An incomplete total says so in its own total row**, not in a footnote: *"Minimum priced amount only - excludes unestablished ancillary rates"*.

**A working assumption is allowed; a determination is not.** A forecast may adopt a base case, labelled *"Working forecast assumption only - not an Engineer's determination"*. A contractual assessment must not select between competing readings while the question is open - show them side by side.

**Assumed inputs are marked where they are used**, with whose assumption it is. Passive voice next to a client's document reads as the client's assumption.

## Voice

Workbooks get forwarded. Assume every one reaches the other side.

| Instead of | Write |
|---|---|
| "WTP measures 1,817.61" | "Assessed: 1,817.61" |
| "our measure", "we assess" | "the assessment", "assessed" |
| "Rev 04 had it reversed" | "the previously issued assessment had it reversed" |
| "Concede early; it costs credibility to defend" | "The correction increases the assessed quantity" |
| "expect concession", "do not let this be revised away" | delete |

**No tactical or negotiating commentary in any workbook.** Strategy belongs in a separate private note.

**Attribute third-party figures correctly rather than labelling everything a claim.** A Contractor's unsubstantiated statement is a claim - "the Contractor claims", "as submitted". A previously certified amount, a contractual rate, an instructed quantity or a jointly verified measurement is not; describe it as what it is. Never restate a claimed figure as fact, and never correct a submitted figure in place - record it exactly and raise the error separately.

## Never

1. Architecture description in a visible cell - "engine", "parametric", "the standard for every future wall", "the only tab edited to add a wall".
2. Tactical or negotiating instructions.
3. A code as the only explanation. `RFP-021` beside "Temporary traffic signage" is clearer than removing it; `R2` alone, carrying the meaning, is not.
4. A total mixing complete and incomplete components without status in the same row.
5. A figure whose only trace is a named range.
6. An instruction telling the user to do something the workbook already does automatically.
7. **Unrestricted find and replace.** Match whole words or entire cell contents. Replacing "our" without whole-word matching turned *source* into *sThece* and *four* into *fThe* across 31 cells - real damage, D-18 rev 01.
8. Hand-editing a delivered file when a build script exists. The next rebuild discards it silently. Change the script.


## Notes, labels and links (added 01-Oct-2026, from RFP-027 Rev 02)

- **A note is a heading and two to five bullets, one sentence each.** An item introduction, a basis cell, a programme note, a comparison comment: all of them. Alaa rejected a 180-word paragraph three times before it became seven bullets. Rate lists split on their semicolons.
- **Never invite the other side to check the assessment.** "Rate: assessed allowance (Riyadh market, Sep-2026)" is complete. "Needs confirmation by the Contractor", "pending the Contractor's substantiation" and "to be confirmed by the Contractor" are banned. An outstanding document is stated as a fact: "the Contractor's method statement is outstanding".
- **Block labels name the party and the basis:** SAMA Submitted Cost, SAMA Submitted Programme, Assessed, Variance. Never "carried", "Basis A/B", "as submitted" as a column label, "departure" (write "Additional helpers assessed"), "man-days" (write person-days).
- **A code never stands alone.** D1, H8, P20b and an activity ID are reference detail beside a plain name.
- **Units are explicit at every figure:** working days, calendar days, months of 30.4 days, person-days, person-weeks, person-months at 26 working days; a weekly headcount is "people in the week", not a daily figure and not a total.
- **Figures quoted in notes are checked against the cells on every rebuild.** A note that says 24-Dec beside a quantity built to 26-Dec is a contradiction the reader finds first.
- **Navigation is hyperlinks, not instructions.** Every cross-tab or cross-row formula links to its first source; every tab has a Return link; a multi-input formula links to one source and the how-to note says so. `contractor-proposal-assessment/scripts/audit_links.py` checks every link against the formula it sits on.
- **Author-facing words are banned in visible cells:** "user-authorised", "on instruction", "to be recorded on", "mismatch stated plainly", "(issue to be confirmed)".

## Validate

```cmd
python scripts\validate_workbook.py "path\to\workbook.xlsx"
```

```cmd
python scripts\validate_workbook.py "path\to\workbook.xlsx" --json report.json
```

A linter, not an authority, and deliberately narrow. It catches mechanical defects. It cannot judge whether the tab answers purpose, result, basis and action - that check was removed rather than left giving false confidence. Warnings are prompts to look, not orders to change. Read the first tab aloud before issue.

Where a build script exists, correct the script and rebuild. Otherwise, correct a controlled copy of the workbook.

## Build note - complex or recurring builds only

For a multi-tab model, a recurring build, or anything with material assumptions, write a short internal note: purpose, one-off or recurring, material assumptions, incomplete figures, checks run, known limitations. Template: `references/build-note-template.md`.

**Internal only. Never issued with the workbook, never forwarded to the Engineer or Contractor.** It records build limitations, not commercial position - no tactical content, no "what an opposing QS would ask".

Skip it for a single-sheet comparison, a rate build-up or a small variation assessment unless it rests on a material assumption.

## Working with the other skills

This skill governs communication. It never overrides a skill that governs content or format.

| Skill | Interaction |
|---|---|
| `wtp-excel-format` | Appearance, palette, Project Info tab, Cover textboxes. Its twelve named ranges are on the approved list - they are mandatory, not a defect |
| `file-authorship-metadata` | Metadata. This skill only checks an author exists |
| `contractor-proposal-assessment` | Comparison structure and the six `List_` validation anchors, which are on the approved list |
| `mass-grading-assessment` | Four-tab structure with a Summary and no hardcoded values - already compliant |
| `ipa-to-ipc-translation`, `ipc-template-generator` | Per-item tabs exceed the visible-tab count; use `--max-visible-tabs`. Freezing a previous period's row as values is deliberate practice, not a duplicated figure |

Where a content skill mandates a structure, that structure wins. This skill then applies to everything it does not mandate: the first tab, the wording, the labelling of uncertainty, and the absence of author-facing language.

## Evidence base

Every rule was earned on D-18 RFP-015, where a workbook that passed a 43-check audit of its own arithmetic was rejected twice by its reader. It was built as a reusable engine when an assessment was asked for; 95 named ranges made formulas compact for the author and untraceable for the reviewer; one issue was explained on six tabs; internal codes led the narrative; helper tabs sat level with conclusions; tactical commentary sat in a file intended for issue; zero rates produced an apparently complete total.

None of it was an arithmetic error.

**The first version of this skill repeated the same failure** - a mandatory seven-section Summary, a 40-word minimum per paragraph, a mandatory eleven-block JSON manifest, and a validator that flagged `=SUM(Table[Column])` as an external link. It was reviewed and cut to about two thirds of its size, and the JSON manifest was dropped entirely. If this skill starts growing again, that is the symptom, not the cure.
