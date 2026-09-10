# DISC_05 - Effort, calendar and token estimates, with the usage ledger

- Date: 09-Sep-2026. Author: Alaa Elsayed. Basis: DISC_04 increments, docs/STAGES.md as re-cut 09-Sep-2026, the stage protocol in docs/FOUNDATION.md.
- Every figure here is an estimate with a stated basis. Nothing has been calibrated yet: the ledger in section 6 is filled after the first two working increments and these tables are then re-issued as v1.
- The interactive map (docs/maps/tony_project_map.html) computes the same totals from the same per-stage inputs; if the two ever differ, the map is wrong and this file governs.

## 1. What is being estimated, and what the old figure was

- Two kinds of hours are kept apart. AGENT ACTIVE hours: wall-clock time a Claude Code session is executing a stage (reading, writing, running tests, audit, close). HUMAN REVIEW / QA hours: Alaa and Mo only, reading the plan, reviewing the diff and handover, running the acceptance check, signing and answering questions. Codex's machine review time is a separate line (section 2) and is never counted as human hours. A third quantity, DEPENDENCY WAIT, is calendar time with no work: Mo's signature, the Tony call, an accounts export, a hosting account.
- The 08-Sep-2026 figure of 231 to 320 hours is withdrawn, not carried. It was a blended "Alaa and Mo combined with Claude" number for 35 stages of a different scope (it included correspondence, claims, KPI and pipeline stages and excluded estimating history, procurement and manufacture), and it assumed humans building in chat with Claude, typing and re-typing. The tables below use the D-15 scope, 31 stages, and the agent-executes / humans-review model that Alaa authorised on 09-Sep-2026.

## 2. Effort per increment (hours)

| Inc | Stages | Agent active low / base / high | Human review and QA low / base / high | Main uncertainty |
|---|---|---|---|---|
| 1 Spine and scheme register | S00 to S04 | 8 / 11 / 15 | 6 / 8 / 12 | first two stages calibrate everything; walk-through review is human-heavy |
| 2 Estimating and rates | S05, S06 | 3 / 4 / 6 | 2 / 3 / 5 | specification key design |
| 3 Procurement | S07 to S09 | 5 / 7 / 10 | 4 / 5 / 8 | RFQ document layout |
| 4 Manufacture and site evidence | S10 to S13 | 8 / 11 / 15 | 6 / 9 / 13 | phone PWA offline behaviour; site-user feedback needs the Tony call |
| 5 Variations | S14 to S16 | 5 / 7 / 10 | 4 / 5 / 8 | provisional sum rules |
| 6 Applications | S17 to S19 | 6 / 8 / 11 | 4 / 6 / 9 | F4-derived layout export |
| 7 Cost ledger and reports | S20 to S23 | 8 / 11 / 15 | 6 / 8 / 12 | double-counting rules; snapshot design |
| 8 Multi-user and hosting | S24 to S26 | 5 / 7 / 10 | 5 / 7 / 10 | hosting account and restore rehearsal are hands-on |
| 9 Consultant face | S27 to S30 | 7 / 9 / 12 | 5 / 7 / 10 | D-06 for live use |
| Thin demonstration (Inc 1 to 7) | S00 to S23 | 43 / 59 / 82 | 32 / 44 / 67 | |
| All increments (Inc 1 to 9) | S00 to S30 | 55 / 75 / 104 | 42 / 58 / 87 | |
| Full Stage 1 depth beyond the demo | real data tasks in DISC_04 section 2 | 10 / 15 / 25 | 10 / 15 / 25 | needs Q1 and Q4 answered |
| Full Stage 1 total | | 65 / 90 / 129 | 52 / 73 / 112 | |

- Per-stage basis: a build stage is 1.5 to 3.5 agent hours (plan, gate, build with tests, audit, close) and 1 to 2.5 human hours (plan approval, diff review, acceptance run, handover read). A docs-only stage is about half of each. These are the numbers to calibrate.
- Codex management overhead is a separate line, not folded into either column: one planning review per increment and one review per stage handover, about 40 reviews for Increments 1 to 9. It is machine time on another provider; its token cost is outside this session's measurement and is left editable in the map.

## 3. Calendar scenarios

- The binding constraint is human-attended time: human review hours plus session supervision (an agent session needs a person to open it, approve the plan and close it; assumed 0.3 hours per agent hour).
- Attended hours, base: demo 44 + 0.3 x 59 = 62; all increments 58 + 0.3 x 75 = 81; full Stage 1 73 + 0.3 x 90 = 100.

- Weeks are the estimate. The formula, used identically here and in the map: weeks = attended hours / hours per week; with waits, low = weeks + wait low, high = 1.15 x weeks + wait high, where the wait bands are demo 0.5 to 2.5 weeks, all increments 1.5 to 4, full Stage 1 1.5 to 5. Dates are indicative only, derived by adding those weeks to 15-Sep-2026; they are not delivery commitments.

| Scenario | Human hours per week | Demo (Inc 1 to 7) | All increments | Full Stage 1 |
|---|---|---|---|---|
| Low pace | 8 | 7.7 weeks; 8.2 to 11.4 with waits (indicative 11-Nov to 03-Dec-2026) | 10.1 weeks; 11.6 to 15.6 (05-Dec-2026 to 02-Jan-2027) | 12.5 weeks; 14.0 to 19.4 (22-Dec-2026 to 29-Jan-2027) |
| High pace | 20 | 3.1 weeks; 3.6 to 6.0 (10-Oct to 27-Oct-2026) | 4.0 weeks; 5.5 to 8.6 (24-Oct to 14-Nov-2026) | 5.0 weeks; 6.5 to 10.8 (31-Oct to 29-Nov-2026) |

- Dependency waits assumed: Mo's signature on S00 (0 to 2 weeks, in parallel with nothing else because S00 is first); the Tony call before S10 (1 to 3 weeks, in parallel with Increments 1 to 3); a hosting account before S24 (1 week); the accounts export and real Gamma files for full depth (unknown; only the full-depth column waits on them).
- Uncertainty: plus or minus 35 percent on every hour figure until the ledger has two increments in it. The high and low columns are that band, not a promise.

## 4. Token estimate, with the assumptions written down

Token consumption is estimated per stage session from disjoint categories, then multiplied by a session count. Every input token falls into exactly one of three: read from cache, written to cache, or sent uncached; output is the fourth. All figures are tokens of the main model; subagent audits are inside the same session. The retry factor and the cache behaviour are uncalibrated assumptions until the ledger has two increments in it.

| Input per session | Low | Base | High | Why |
|---|---|---|---|---|
| Tool turns per session | 25 | 40 | 60 | plan, gate, build, tests, audit, close |
| Cached context at mid-session (tokens) | 60,000 | 100,000 | 150,000 | CLAUDE.md, FOUNDATION, DECISIONS, stage file, DISC extracts, accumulated tool output |
| Cache-read input per session (turns x context) | 1,500,000 | 4,000,000 | 9,000,000 | every turn re-reads the cached prefix |
| Cache-write input per session | 70,000 | 120,000 | 180,000 | the context written once plus growth; charged at the write rate |
| Uncached input per session | 10,000 | 20,000 | 30,000 | tokens sent that are neither read from nor written to the cache (turn deltas at the end of a request) |
| Output per session | 15,000 | 25,000 | 40,000 | code, tests, documents, handover |
| Sessions, the 31 stages x retry factor | 31 (x1.0) | 40 (x1.3) | 56 (x1.8) | a retry is a stage that needs a second session after audit; uncalibrated |
| Sessions, full-depth work beyond the 31 stages | 6 | 10 | 16 | real-data imports, layouts, pilot fixes (DISC_04 section 2 right-hand column) |

| Total, tokens | 31 stages low / base / high | Full-depth low / base / high | Combined low / base / high |
|---|---|---|---|
| Cache-read input | 46.5 M / 160 M / 504 M | 9 M / 40 M / 144 M | 55.5 M / 200 M / 648 M |
| Cache-write input | 2.17 M / 4.80 M / 10.08 M | 0.42 M / 1.20 M / 2.88 M | 2.59 M / 6.00 M / 12.96 M |
| Uncached input | 0.31 M / 0.80 M / 1.68 M | 0.06 M / 0.20 M / 0.48 M | 0.37 M / 1.00 M / 2.16 M |
| Output | 0.47 M / 1.00 M / 2.24 M | 0.09 M / 0.25 M / 0.64 M | 0.56 M / 1.25 M / 2.88 M |

- The 31-stage column EXCLUDES full-depth work, the discovery already spent (two packages on 09-Sep-2026 and this S00 session) and Codex. The combined column adds full-depth work only.

- Context window occupancy is not billed usage; only what each request sends and receives is. Account-wide usage percentages are not this project's consumption; the ledger records this project only.

## 5. Two ways to express the cost

- Claude Max subscription (what is actually in use): no per-token bill. Consumption shows as the plan's usage meter over its rolling windows, and per the official guidance at support.claude.com/en/articles/11145838 (cited by Codex on 09-Sep-2026; the page is blocked by this session's network policy, so it was not read here) the limits are shared across Claude and Claude Code. The API-equivalent figures below are a scale, not a subscription bill. The ledger records sessions and observed tokens only. No extra usage, model change, new subscription or purchased capacity is assumed or requested.
- Hypothetical API dollar equivalent, for scale only. Rates verified on 09-Sep-2026 from platform.claude.com/docs/en/about-claude/pricing (USD per million tokens): Claude Fable 5.1 - input 10, 5-minute cache write 12.50, cache read 0.25, output 50; Claude Opus 5 - input 5, cache write 6.25, cache read 0.50, output 25; Claude Sonnet 5 - input 2, cache write 2.50, cache read 0.20, output 10. This session is configured as Claude Fable 5.1.

| Model, 5-minute cache | 31 stages low / base / high | Combined with full-depth low / base / high |
|---|---|---|---|
| Claude Fable 5.1 equivalent | 65 / 158 / 381 USD | 78 / 198 / 490 USD |
| Claude Opus 5 equivalent | 50 / 139 / 379 USD | 60 / 174 / 487 USD |

- Arithmetic (31 stages, base, Fable 5.1): 0.80 x 10 + 4.80 x 12.5 + 160 x 0.25 + 1.00 x 50 = 8 + 60 + 40 + 50 = 158. The map recomputes these from editable rates.
- One-hour cache sensitivity: if the writes were one-hour cache writes (2 x base input instead of 1.25 x), the 31-stage base rises to about 194 USD on Fable 5.1 and 157 USD on Opus 5. Which cache duration the session actually uses is not visible here; the ledger cannot resolve it either, so both are shown.
- Codex management overhead is a separate provider and a separate meter; it is not in these figures. The map carries an editable line for it.

## 6. Usage ledger (calibration record)

- One row per stage session. Observed figures come from what the session surface shows; where a figure is not visible it is recorded as "not shown", never estimated into the observed column.

| Date | Package or stage | Sessions | Tool turns (observed) | Tokens observed (input / cache write / cache read / output) | Notes |
|---|---|---|---|---|---|
| 09-Sep-2026 | Discovery, first package (DISC_01 to DISC_04) | 1 | about 30 | not shown for the main session; one read-only review subagent reported 62,721 tokens total for 4 tool uses | the container counter is a context budget, not billed usage |
| 09-Sep-2026 | Discovery, second package (this document, the map, STAGES re-cut) | 1 (same session) | about 15 | not shown | same |
| 10-Sep-2026 | S00 bootstrap (this session, continued) | 1 (same session) | about 12 | not shown | PostgreSQL 16.13 run locally for the two-engine check |
| | S01 | | | | to be filled |

- Calibration rule: after S00 to S06 close (two increments), compare observed turns and tokens per stage with section 4's base column and re-issue sections 2 to 5 as v1 with the measured ratios.
