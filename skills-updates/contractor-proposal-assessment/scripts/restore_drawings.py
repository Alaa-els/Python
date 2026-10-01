"""Copy the Cover and Appendix Cover letterhead drawings from the source workbook into the finalised one (openpyxl drops images)."""
import sys, zipfile, re, shutil, os, tempfile
src, dst = sys.argv[1], sys.argv[2]
zs = zipfile.ZipFile(src)
names = zs.namelist()
# find which sheet files are Cover / Appendix Cover in the source via workbook.xml + rels
wbx = zs.read('xl/workbook.xml').decode(); rels = zs.read('xl/_rels/workbook.xml.rels').decode()
sheet_rid = dict(re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx)) or {m[1]: m[0] for m in re.findall(r'<sheet [^>]*r:id="([^"]+)"[^>]*name="([^"]+)"', wbx)}
rid_target = dict(re.findall(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)) or {m[1]: m[0] for m in re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)}
cover_files = []
for name in ('Cover', 'Appendix Cover'):
    rid = sheet_rid.get(name); tgt = rid_target.get(rid, '')
    f = 'xl/' + tgt.lstrip('/').replace('xl/', '') if not tgt.startswith('/') else tgt.lstrip('/')
    cover_files.append(f)
tmp = tempfile.mkdtemp()
with zipfile.ZipFile(dst) as zd: zd.extractall(tmp)
added = []
for f in cover_files:
    base = os.path.basename(f)
    for part in (f, f"xl/worksheets/_rels/{base}.rels"):
        if part in names:
            os.makedirs(os.path.dirname(os.path.join(tmp, part)), exist_ok=True)
            open(os.path.join(tmp, part), 'wb').write(zs.read(part)); added.append(part)
for part in names:
    if part.startswith('xl/drawings/') or part.startswith('xl/media/'):
        os.makedirs(os.path.dirname(os.path.join(tmp, part)), exist_ok=True)
        open(os.path.join(tmp, part), 'wb').write(zs.read(part)); added.append(part)
ct = open(os.path.join(tmp, '[Content_Types].xml'), encoding='utf-8').read()
if 'Extension="jpeg"' not in ct: ct = ct.replace('</Types>', '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>')
for part in added:
    if part.startswith('xl/drawings/drawing') and part.endswith('.xml') and f'PartName="/{part}"' not in ct:
        ct = ct.replace('</Types>', f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
open(os.path.join(tmp, '[Content_Types].xml'), 'w', encoding='utf-8').write(ct)
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(tmp, '[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(tmp):
        for fn in files:
            p = os.path.join(root, fn); a = os.path.relpath(p, tmp)
            if a != '[Content_Types].xml': z.write(p, a)
shutil.rmtree(tmp); print('drawings restored:', len(added), 'parts; cover sheets', cover_files)
