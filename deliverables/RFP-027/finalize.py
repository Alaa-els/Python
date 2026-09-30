"""Stage B: zip-level metadata clean of the recalculated workbook (never re-saved with openpyxl afterwards)."""
import sys, os, shutil, zipfile, tempfile
sys.path.insert(0, '/root/.claude/skills/synced/7afc0081-3f0d-400c-b5d3-c14839f2b769_af511efe-500f-42cb-a41a-8b5eb7f786e9/wtp-excel-format/scripts')
from wtp_finalize import clean_metadata
src, dst = sys.argv[1], sys.argv[2]
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(src) as z: z.extractall(tmp)
# hyperlinks: LibreOffice rewrites internal links as external file links with the local path; restore pure in-workbook locations
import re, glob
for sheet in glob.glob(os.path.join(tmp, 'xl', 'worksheets', 'sheet*.xml')):
    x = open(sheet, encoding='utf-8').read()
    rids = re.findall(r'<hyperlink [^>]*?r:id="(rId\d+)"[^>]*location="[^"]+"[^>]*/>', x)
    x2 = re.sub(r'<hyperlink ref="([^"]+)" r:id="rId\d+" location="([^"]+)"( display="[^"]*")?/>', r'<hyperlink ref="\1" location="\2"/>', x)
    x2 = re.sub(r'location="([^"]*)"', lambda m: 'location="' + m.group(1).replace('%20', ' ') + '"', x2)
    if x2 != x:
        open(sheet, 'w', encoding='utf-8').write(x2)
        rels = os.path.join(tmp, 'xl', 'worksheets', '_rels', os.path.basename(sheet) + '.rels')
        if os.path.exists(rels):
            rx = open(rels, encoding='utf-8').read()
            for rid in rids:
                rx = re.sub(r'<Relationship Id="' + rid + r'" [^>]*?TargetMode="External"/>', '', rx)
            open(rels, 'w', encoding='utf-8').write(rx)
            if '<Relationship ' not in rx: os.remove(rels)
        print('hyperlinks restored to internal form:', os.path.basename(sheet), len(rids))
clean_metadata(tmp, title='Qiddiya - TSE Irrigation Storage Tanks - Cost Assessment - Rev 02', subject='RFP-027, Contract QPMO-410-CT-05958 - cost assessment of the Contractor proposal SAMACO-RRFP-000001')
os.makedirs(os.path.dirname(dst), exist_ok=True)
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(tmp, '[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for f in files:
            p = os.path.join(root, f); a = os.path.relpath(p, tmp)
            if a != '[Content_Types].xml': z.write(p, a)
shutil.rmtree(tmp); print('finalised', dst)
