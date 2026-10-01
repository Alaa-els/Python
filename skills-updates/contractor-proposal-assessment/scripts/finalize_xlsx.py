"""Stage B for an assessment workbook, after LibreOffice recalculation and never followed by an openpyxl save.
Restores in-workbook hyperlinks (LibreOffice rewrites them as external links carrying the local path), cleans the
authorship metadata and rezips with [Content_Types].xml first.

usage: python finalize_xlsx.py recalculated.xlsx "final name.xlsx" --title "..." --subject "..."
"""
import sys, os, re, glob, shutil, zipfile, tempfile, argparse
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', 'wtp-excel-format', 'scripts'))
from wtp_finalize import clean_metadata
ap = argparse.ArgumentParser(); ap.add_argument('src'); ap.add_argument('dst'); ap.add_argument('--title', default=None); ap.add_argument('--subject', default=None)
ap.add_argument('--author', default='Alaa Elsayed'); ap.add_argument('--company', default='WT Partnership')
a = ap.parse_args()
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(a.src) as z: z.extractall(tmp)
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
            for rid in rids: rx = re.sub(r'<Relationship Id="' + rid + r'" [^>]*?TargetMode="External"/>', '', rx)
            open(rels, 'w', encoding='utf-8').write(rx)
            if '<Relationship ' not in rx: os.remove(rels)
clean_metadata(tmp, author=a.author, company=a.company, title=a.title, subject=a.subject)
os.makedirs(os.path.dirname(os.path.abspath(a.dst)), exist_ok=True)
with zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(tmp, '[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for f in files:
            p = os.path.join(root, f); n = os.path.relpath(p, tmp)
            if n != '[Content_Types].xml': z.write(p, n)
shutil.rmtree(tmp)
leak = [n for n in zipfile.ZipFile(a.dst).namelist() if n.endswith('.xml') and b'file:///' in zipfile.ZipFile(a.dst).read(n)]
print('finalised', a.dst, '| external-path leaks:', len(leak))
