"""
WTP Excel Format - Letterhead injection.

Injects the WTP first-page letterhead (image2.png equivalent) on portrait
content tabs using OOXML's legacyDrawingHF mechanism.

Skips landscape tabs (the letterhead artwork is portrait-oriented).

Public function:
  inject_letterhead(unpacked_dir, portrait_tab_names, letterhead_path)
"""
import os, re, shutil
import xml.etree.ElementTree as ET

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'

VML_TEMPLATE = '''<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"
  o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
  <v:stroke joinstyle="miter"/>
  <v:formulas>
   <v:f eqn="if lineDrawn pixelLineWidth 0"/>
   <v:f eqn="sum @0 1 0"/>
   <v:f eqn="sum 0 0 @1"/>
   <v:f eqn="prod @2 1 2"/>
   <v:f eqn="prod @3 21600 pixelWidth"/>
   <v:f eqn="prod @3 21600 pixelHeight"/>
   <v:f eqn="sum @0 0 1"/>
   <v:f eqn="prod @6 1 2"/>
   <v:f eqn="prod @7 21600 pixelWidth"/>
   <v:f eqn="sum @8 21600 0"/>
   <v:f eqn="prod @7 21600 pixelHeight"/>
   <v:f eqn="sum @10 21600 0"/>
  </v:formulas>
  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype><v:shape id="LH" o:spid="_x0000_s1025" type="#_x0000_t75"
  style='position:absolute;margin-left:0;margin-top:0;width:595.5pt;height:842.5pt;
  z-index:1'>
  <v:imagedata o:relid="rIdImg" o:title="Letterhead"/>
  <o:lock v:ext="edit" rotation="t"/>
 </v:shape><v:shape id="LHFIRST" o:spid="_x0000_s1026" type="#_x0000_t75"
  style='position:absolute;margin-left:0;margin-top:0;width:595.5pt;height:842.5pt;
  z-index:2'>
  <v:imagedata o:relid="rIdImg" o:title="Letterhead"/>
  <o:lock v:ext="edit" rotation="t"/>
 </v:shape></xml>'''


def _sheet_num(unpacked_dir, sheet_name):
    wbtree = ET.parse(os.path.join(unpacked_dir, 'xl', 'workbook.xml'))
    sheets = wbtree.findall(f'{NS_MAIN}sheets/{NS_MAIN}sheet')
    name_to_rid = {s.get('name'): s.get(f'{NS_REL}id') for s in sheets}
    rid = name_to_rid.get(sheet_name)
    if not rid:
        return None
    rt = ET.parse(os.path.join(unpacked_dir, 'xl', '_rels', 'workbook.xml.rels'))
    rid_to_target = {r.get('Id'): r.get('Target') for r in rt.findall(f'{NS_PKG}Relationship')}
    target = rid_to_target.get(rid, '')
    m = re.search(r'sheet(\d+)\.xml', target)
    return int(m.group(1)) if m else None


def _ensure_content_types(ct_path):
    """Add png and vml extensions to [Content_Types].xml if missing."""
    with open(ct_path) as f:
        ct = f.read()
    changed = False
    if 'Extension="png"' not in ct:
        ct = ct.replace(
            '<Default Extension="xml"',
            '<Default Extension="png" ContentType="image/png"/>'
            '<Default Extension="xml"'
        )
        changed = True
    if 'Extension="vml"' not in ct:
        ct = ct.replace(
            '<Default Extension="xml"',
            '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
            '<Default Extension="xml"'
        )
        changed = True
    if changed:
        with open(ct_path, 'w') as f:
            f.write(ct)


def inject_letterhead(unpacked_dir, portrait_tab_names, letterhead_path):
    """
    Inject the WTP letterhead as legacyDrawingHF on each portrait tab.

    portrait_tab_names: list of sheet names to apply letterhead to
    letterhead_path: path to letterhead PNG (e.g., assets/letterhead.png)
    """
    if not os.path.exists(letterhead_path):
        raise FileNotFoundError(f"Letterhead missing: {letterhead_path}")

    media_dir = os.path.join(unpacked_dir, 'xl', 'media')
    os.makedirs(media_dir, exist_ok=True)
    drawings_dir = os.path.join(unpacked_dir, 'xl', 'drawings')
    os.makedirs(drawings_dir, exist_ok=True)

    # Add letterhead image with fresh number
    used_nums = []
    for f in os.listdir(media_dir):
        m = re.match(r'image(\d+)\.', f)
        if m:
            used_nums.append(int(m.group(1)))
    next_img_num = max(used_nums) + 1 if used_nums else 1
    img_filename = f'image{next_img_num}.png'
    shutil.copy(letterhead_path, os.path.join(media_dir, img_filename))

    # Ensure content types include png and vml
    ct_path = os.path.join(unpacked_dir, '[Content_Types].xml')
    _ensure_content_types(ct_path)

    # Find next vmlDrawing number
    existing_vmls = [f for f in os.listdir(drawings_dir) if f.startswith('vmlDrawing')]
    next_vml_num = len(existing_vmls) + 1

    applied_count = 0
    for tab_name in portrait_tab_names:
        snum = _sheet_num(unpacked_dir, tab_name)
        if not snum:
            print(f"  [skip] sheet not found: {tab_name}")
            continue
        sheet_xml = os.path.join(unpacked_dir, 'xl', 'worksheets', f'sheet{snum}.xml')
        if not os.path.exists(sheet_xml):
            continue

        vml_filename = f'vmlDrawing{next_vml_num}.vml'
        vml_path = os.path.join(drawings_dir, vml_filename)
        with open(vml_path, 'w') as f:
            f.write(VML_TEMPLATE)

        vml_rels_dir = os.path.join(drawings_dir, '_rels')
        os.makedirs(vml_rels_dir, exist_ok=True)
        vml_rels_path = os.path.join(vml_rels_dir, f'{vml_filename}.rels')
        with open(vml_rels_path, 'w') as f:
            f.write(f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{img_filename}"/>
</Relationships>''')

        sheet_rels_path = os.path.join(unpacked_dir, 'xl', 'worksheets', '_rels',
                                       f'sheet{snum}.xml.rels')
        if os.path.exists(sheet_rels_path):
            with open(sheet_rels_path) as f:
                sr = f.read()
            rid_nums = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', sr)]
            next_rid = max(rid_nums) + 1 if rid_nums else 1
            new_rel = (f'<Relationship Id="rId{next_rid}" '
                       f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawingHF" '
                       f'Target="../drawings/{vml_filename}"/>')
            sr = sr.replace('</Relationships>', new_rel + '</Relationships>')
            with open(sheet_rels_path, 'w') as f:
                f.write(sr)
        else:
            next_rid = 1
            os.makedirs(os.path.dirname(sheet_rels_path), exist_ok=True)
            with open(sheet_rels_path, 'w') as f:
                f.write(f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId{next_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawingHF" Target="../drawings/{vml_filename}"/>
</Relationships>''')

        with open(sheet_xml) as f:
            sx = f.read()
        # Strip any pre-existing headerFooter / legacyDrawingHF
        sx = re.sub(r'<headerFooter[^/]*/>|<headerFooter[^>]*>.*?</headerFooter>', '', sx, flags=re.DOTALL)
        sx = re.sub(r'<legacyDrawingHF[^/]*/>', '', sx)

        inject = (f'<headerFooter differentFirst="1">'
                  f'<firstHeader>&amp;L&amp;G</firstHeader>'
                  f'</headerFooter>'
                  f'<legacyDrawingHF xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                  f'r:id="rId{next_rid}"/>')
        sx = sx.replace('</worksheet>', inject + '</worksheet>')
        with open(sheet_xml, 'w') as f:
            f.write(sx)

        next_vml_num += 1
        applied_count += 1
        print(f"  Letterhead applied to {tab_name} (sheet{snum})")

    return applied_count
