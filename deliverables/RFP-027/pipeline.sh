#!/bin/bash
set -e
cd /tmp/claude-0/-home-user-Python/db2f0e90-a774-5426-831f-3c6b89e31583/scratchpad
S=/root/.claude/skills/synced/7afc0081-3f0d-400c-b5d3-c14839f2b769_af511efe-500f-42cb-a41a-8b5eb7f786e9
OUT="out/Qiddiya - TSE Irrigation Storage Tanks - Cost Assessment - Rev 02 - 30 Sep 2026.xlsx"
python3 build_rev02.py
cp stageA.xlsx stageA_calc.xlsx
timeout 300 python3 $S/xlsx/scripts/recalc.py stageA_calc.xlsx 240 | head -4
python3 finalize.py stageA_calc.xlsx "$OUT"
python3 $S/file-authorship-metadata/scripts/verify_metadata.py "$OUT" | tail -3
python3 $S/workbook-readability/scripts/validate_workbook.py "$OUT" 2>&1 | tail -4
mkdir -p render && cp "$OUT" render/r.xlsx && cd render && rm -f p*.png r.pdf
timeout 280 soffice --headless --norestore -env:UserInstallation=file:///tmp/claude-0/-home-user-Python/db2f0e90-a774-5426-831f-3c6b89e31583/scratchpad/lo --convert-to pdf --outdir . r.xlsx 2>&1 | tail -1
python3 - <<'PY'
import pymupdf
doc=pymupdf.open('r.pdf'); print('pages',len(doc))
for i,p in enumerate(doc):
    t=p.get_text()[:70].replace('\n',' | '); print(i, t)
    p.get_pixmap(dpi=110).save(f'p{i:02d}.png')
PY
