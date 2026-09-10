# W1 - Spine walk-through, part 1: organisation, scheme register, cross-cutting

- Stage S01, 10-Sep-2026. Author: Alaa Elsayed. Basis: DISC_03 sections 1, 2 and 9; Codex direction of 10-Sep-2026. Documents only; S02 writes the models from this.
- Reading rule: one row is one thing. Every column has a meaning and a unit. Sample rows use the authorised defaults (company display name GEMMA, scheme code GEM-###); every name and figure is invented.
- Columns on EVERY table, not repeated below: id (system), company (link, see I-3), created_by, created_at, updated_by, updated_at, status (never hard-deleted; closed or superseded), and a history row for every change (section 9). Imported rows add source_file, source_tab, source_row, imported_at.

## 1. Organisation

| Table | One row = | Columns (meaning, unit) | Links | Sample row |
|---|---|---|---|---|
| company | one legal entity using the workbench | code (short, unique); display_name (text; default GEMMA, editable, Q3); currency (ISO 4217 code, default GBP); timezone (default Europe/London) | - | CO-01, GEMMA, GBP |
| user | one person who can log in | Django user (name, email, password hash); company (link); is_active | company | u-alaa, A. Sample, CO-01 |
| role | one named permission set | name in {owner, editor, viewer, site}; a fixed list, not a table users edit | - | editor |
| company_membership | one user's standing in one company | user; company; role; active from and to (dates) | user, company, role | u-alaa in CO-01 as owner |
| scheme_membership | one user's access to one scheme | user; scheme; role (may be narrower than the company role, never wider); active from and to | user, scheme, role | u-site1 on GEM-001 as site |
| party | one organisation on a scheme | name; kind in {customer, main contractor, employer, consultant, supplier, subcontractor}; registration number (text); address | company | Sample Main Contractor Ltd, main contractor |
| party_role | one party's role on one scheme | party; scheme; role text (for example "client", "employer's agent") | party, scheme | Sample Main Contractor Ltd is client on GEM-001 |
| supplier | the supplier attributes of a party that can be sent an RFQ | party; trade (text); geography (text or region code); capability (text); quality note (text) | party | Sample Joinery Supplies Ltd, ironmongery, South East |
| supplier_performance | one note on a supplier's performance on one scheme | supplier; scheme; date; rating (1 to 5); note | supplier, scheme | 4, "delivered on time, one short delivery" |

## 2. Scheme register

| Table | One row = | Columns (meaning, unit) | Links | Sample row |
|---|---|---|---|---|
| scheme | one job the company prices or delivers | code (stable, unique within company; format setting, default GEM-### with a company prefix and three digits); name; site address; my_role in {contractor, consultant}; currency (ISO code, inherited from company, editable); status in {enquiry, tendered, live, final account, closed} | company | GEM-001, "Sample Office Fit-out", contractor, GBP, live |
| scheme_revision | one commercial baseline of a scheme | scheme; number (0, 1, 2 ...); kind in {tender, baseline, change}; date; reason; authorised_by (user); superseded_by (link to the later revision, empty while current) | scheme | GEM-001 R1, baseline, "order accepted" |
| contract_terms | one term for one scheme | scheme; name (retention percent, payment days, certification days, MCD or rebate percent, discount percent, notice days, time bar days, defects period months, chain-of-custody flag); value (decimal or text by name); unit; clause reference (text) | scheme | GEM-001, retention percent, 5.00, percent, "clause 4.3" |
| cost_code | one code in the company's chart of cost heads | company; code (text, unique in company); name; parent (link to the head); level (1 = head, 2 = sub-code); kind in {cost, overhead recovery, profit}; active | company, cost_code (parent) | CC-410 "Manufacturing materials" under CC-400 "Manufacture" |
| programme_activity | one WBS activity | scheme; code (unique in scheme); name; planned start; planned finish (dates); parent (optional) | scheme | A-2310 "Level 2 reception joinery", 01-Oct to 21-Oct-2026 |
| location | one place on the scheme | scheme; code (unique in scheme); building; level; zone; room | scheme | L2-R04, Block A, Level 2, reception |
| boq_item | one bill or activity-schedule item on one revision | scheme_revision; item_ref (stable identity across revisions, see I-4); section; description; quantity (decimal, 3 places); unit (text); rate (money); value (money, = quantity x rate, stored once); drawing; clarifications; provisional_sum flag; manufactured_in_house flag; superseded_by (link) | scheme_revision | R1, 2.3.014, "Bespoke reception desk", 1.000, nr, 8,400.00, 8,400.00 |
| cost_allocation | one share of one boq_item's VALUE to one cost_code | boq_item; cost_code; share (percent, 2 places) | boq_item, cost_code | 2.3.014 -> CC-410 60.00; 2.3.014 -> CC-230 40.00 |
| wbs_allocation | one share of one boq_item's QUANTITY to one programme_activity at one location | boq_item; programme_activity; location (optional); quantity (decimal, same unit as the item) | boq_item, programme_activity, location | 2.3.014 -> A-2310 at L2-R04, 1.000 nr |

## 3. Cross-cutting (DISC_03 section 9)

| Table | One row = | Columns | Links | Sample row |
|---|---|---|---|---|
| sources | one source document or import file | company; scheme (empty only for company-level files); name; version; date; sha256 (hex); kind in {bill, quotation, drawing, instruction, email, photo, export}; storage path (company/scheme keyed) | company, scheme | GEM-001, "sample_bill_R1.xlsx", v1, sha256 ... |
| history | one change to any row | company; scheme (when the row belongs to one); table; row id; field; old; new; who; when; reason (optional) | - | boq_item 2.3.014 rate 8,200.00 -> 8,400.00 by u-alaa, "tender correction" |
| action | one thing someone must do | company; scheme; owner (user); due (date); origin (table and row); status in {open, done, dropped}; note | user, scheme | "confirm MCD percent with client QS", due 17-Sep-2026 |
| setting | one configurable value | company; scheme (empty = company-wide); key; value (text or decimal); unit; effective from | company, scheme | scheme_code_format = "{prefix}-{seq:03d}" |

## 4. Invariants (each becomes a model rule and a test at S02)

| Id | Invariant | Where enforced |
|---|---|---|
| I-1 | Cost allocation and WBS allocation are two tables for two questions. cost_allocation splits VALUE by cost_code and its shares sum to exactly 100.00 percent per item. wbs_allocation splits QUANTITY by activity and location and its quantities sum to no more than the item quantity. Neither is derived from the other. | cost_allocation, wbs_allocation |
| I-2 | Aggregate quantity never exceeds the authorised quantity and nothing is clamped. The authorised quantity is boq_item.quantity on the current (not superseded) revision. Sum of wbs_allocation.quantity per item <= authorised quantity. Later, sum across locations of the latest verified measured reading per (item, location) <= authorised quantity, and each reading <= that location's allocated quantity where one exists. A breach is a validation error that names the item, the location and the excess; the row is refused, never silently reduced. A genuine over-measure is recorded as a variation candidate, not as progress. | boq_item, wbs_allocation, measured_progress (S12) |
| I-3 | Membership and ownership agree. Every row carries company. Every foreign key must point at a row with the same company; a scheme-scoped row's scheme must belong to the same company. A user reads or writes a scheme row only with an active scheme_membership on that scheme, and a scheme_membership role is never wider than the user's company_membership role. Cross-company or cross-scheme access is refused at the model layer, not only in views. | model clean() on every table; one test per table |
| I-4 | Stable identity across revisions. boq_item.item_ref is the identity of an item within a scheme; each revision holds a version row; (scheme, revision, item_ref) is unique; a version on a superseded revision is read-only; downstream records link to the version row and copy item_ref, so a report can follow an item through R0, R1, R2. Renumbering an item is a new item_ref with a recorded link to the old one, never an edit in place. | boq_item, scheme_revision |
| I-5 | Money is decimal with an explicit currency. Every money column is a fixed-point decimal (14 digits, 2 places; quantities 3 places; rates 4 places where the bill has them), never a float. Every money-bearing row resolves a currency through its scheme; a company-level row uses the company currency. Arithmetic across currencies is refused; conversion is not in Stage 1. | all money columns |
| I-6 | Audit records and files cannot leak across schemes. history and sources carry company and scheme; a history or sources query is always filtered by the caller's memberships; file storage paths are keyed by company and scheme codes; a test creates two schemes in two companies, writes a change and a file to each, and proves that neither user can list, read or download the other's. | history, sources, evidence (S11) |
| I-7 | Revisions are ordered and single-headed. A scheme has exactly one baseline; change revisions follow it in number order; the current revision is the highest not superseded; superseding writes superseded_by on the old one and a history row on both. | scheme_revision |

## 5. Defaults kept, choices left open

| Item | Default in force | Open choice, owner |
|---|---|---|
| Company display name | GEMMA (setting, editable) | confirm name, Alaa (Q3) |
| Scheme code format | GEM-### (setting) | confirm prefix and width, Alaa (Q3) |
| Role set | owner, editor, viewer, site | add "reviewer" for a consultant party later, Alaa (Increment 9) |
| Currency | GBP on company; inherited by scheme | multi-currency is out of Stage 1 |
| Cost head chart | the DISC_03 section 2 list, seeded at S04 | Tony's own heads, after the call (D-02) |
| Programme source | typed or CSV | programme tool export, Q2 |
| Consultant party membership | not modelled in Stage 1 | how a consultant user sees a contractor's scheme, Increment 9 |

## 6. Questions only Alaa can answer (each with the default that applies meanwhile)

- Q-W1-1: Are quantities ever measured in more than one unit for the same item (for example nr and m2)? Default: one unit per item; a second unit is a second item.
- Q-W1-2: Should a scheme_membership be able to grant "site" to someone with no company_membership (a subcontractor's foreman)? Default: no; every user has a company_membership first.
- Q-W1-3: Does a change revision re-baseline every item, or only the items it touches? Default: only the items it touches; untouched items carry forward by item_ref without a new version row.

## 7. Review record

- Alaa: building authorised under Codex management (instruction of 10-Sep-2026); no separate signature required for this stage.
- Codex: review pending at the S01 close.
- Mo: not signed.
