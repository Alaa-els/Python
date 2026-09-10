# SETUP - from an empty folder to the first Claude Code session

## 1. Install Claude Code (once, personal machine)
- Windows PowerShell: `irm https://claude.ai/install.ps1 | iex`
- macOS, Linux, WSL: `curl -fsSL https://claude.ai/install.sh | bash`
- Check: `claude --version` prints a version and "(Claude Code)". First `claude` run asks you to log in with your personal Claude subscription (never a WTP account).
- Also install: Python 3.12, Git for Windows (recommended on Windows so Claude Code has a Bash tool).

## 2. Create the repository
- Make a folder outside any WTP or OneDrive path, for example `C:\dev\cw`.
- Unzip this skeleton into it. You should see CLAUDE.md, SETUP.md, docs/, handovers/, data/, tests/, .claude/, .gitignore.
- Superseded 10-Sep-2026 (D-13 line): the seven CW_ plan files were never placed and the discovery documents in docs/discovery/ are the authoritative Stage 1 definition; docs/plan/ holds only history/. The list below is kept as a record of what was intended:
  - CW_01_Scope_and_Filter_v0.md
  - CW_02_Blueprint_v0.md
  - CW_03_Plan_v0.md (reference only; docs/STAGES.md is the live schedule)
  - CW_04_Gamma_Launch_Customer_Terms_v0.md
  - CW_05_Project_Instructions_v0.md (reference only; CLAUDE.md replaces it inside the repo)
  - CW_06_Test_Data_Inventory_v0.md
  - CW_Handover_C01.md (history only)
- Copy Tony's files into `data/legacy/` (they stay out of git). Not the KPI or BD reports; they are parked.
- In the folder: `git init`, `git add .`, `git commit -m "S00: skeleton"`. Create a private remote (GitHub or similar) owned by you and Mo and push. Mo clones it.

## 3. Before the first session (blocks S01, not S00)
- D-06: read the WTP contract for IP, outside-work and data clauses.
- D-07: ownership, hours and split agreed with Mo in writing.
- Both of you: read CLAUDE.md and docs/FOUNDATION.md once. Mo signs docs/stages/S00_bootstrap.md (W-10).

## 4. First session (S00)
- Open a terminal in the folder, run `claude`.
- Type `/stage S00`. Follow the protocol: it restates, gates, plans; approve the plan; it builds; then `/audit S00`; then `/close S00`.
- Paste this as the first message if `/stage` is not yet listed under `/help` (the commands load from .claude/commands/ when the session starts in the repo root):

```
Open stage S00 of the Commercial Workbench. Read CLAUDE.md, docs/FOUNDATION.md section 2, docs/DECISIONS.md and docs/stages/S00_bootstrap.md. Restate the objective, the decisions it depends on with statuses, the tests in section 5 and what is out of scope. Run the pre-write gate as far as it exists (git status clean). Create branch stage/S00. Then enter plan mode and show the plan: files, tests first, commits in order. Write nothing until I approve.
```

## 5. Every later stage
- Alaa or Mo writes docs/stages/Snn_<slug>.md from _TEMPLATE.md, the other reviews and signs it (W-10), status READY.
- New terminal, `claude`, `/stage Snn` ... `/audit Snn` ... `/close Snn`. Close the terminal. One stage, one session.
- The claude.ai Project stays for non-code work only: the Tony call (C02), decisions, sales gate pack, stage-file drafting if you prefer chat for writing.

## 6. What is where
| Need | Place |
|---|---|
| The standing brief Claude Code reads every session | CLAUDE.md |
| The method | docs/FOUNDATION.md |
| Decisions | docs/DECISIONS.md |
| Schedule and gates | docs/STAGES.md |
| The stage being built | docs/stages/Snn_*.md |
| Product definition | docs/discovery/DISC_03_Record_Model_v0.md and DISC_04_Build_Increments_v0.md (the CW_ blueprint was never placed; D-13 line of 10-Sep-2026) |
| Findings and risks | docs/FINDINGS.md, docs/RISKS.md |
| What each stage delivered | handovers/Snn_handover.md |
| Tony's files (never committed) | data/legacy/ |
| Anonymised test data (committed) | tests/fixtures/ |
