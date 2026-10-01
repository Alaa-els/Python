"""
WTP Excel Format - Textbox injection.

Injects Cover and Appendix Cover textboxes via direct XML editing of unpacked
drawing files. Each textbox gets:
  - Pre-computed value from Project Info as static text (visible immediately)
  - textlink attribute pointing to source cell (for forward dynamic linking)

Use AFTER openpyxl save and BEFORE final repack.

Public function:
  inject_cover_textboxes(unpacked_dir, project_info_values)
    where project_info_values is a dict with keys C6, C7, C8, C9, C13-C20 (strings)
"""
import os, re
import xml.etree.ElementTree as ET

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def _xml_escape(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def _make_textbox(shape_id, name, textlink_ref,
                  from_col, from_row, to_col, to_row,
                  font_size_pt, color_hex, bold, embedded_text):
    """Build a single text box XML element."""
    sz = int(font_size_pt * 100)
    b = '1' if bold else '0'
    A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    safe = _xml_escape(embedded_text)
    return (
        f'<twoCellAnchor editAs="oneCell">'
        f'<from><col>{from_col}</col><colOff>0</colOff>'
        f'<row>{from_row}</row><rowOff>0</rowOff></from>'
        f'<to><col>{to_col}</col><colOff>0</colOff>'
        f'<row>{to_row}</row><rowOff>0</rowOff></to>'
        f'<sp macro="" textlink="{textlink_ref}">'
        f'<nvSpPr><cNvPr id="{shape_id}" name="{name}"/>'
        f'<cNvSpPr txBox="1"/></nvSpPr>'
        f'<spPr>'
        f'<a:xfrm {A}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
        f'<a:prstGeom {A} prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill {A}/>'
        f'<a:ln {A} w="9525"><a:noFill/></a:ln>'
        f'</spPr>'
        f'<txBody>'
        f'<a:bodyPr {A} wrap="square" rtlCol="0" anchor="t"/>'
        f'<a:lstStyle {A}/>'
        f'<a:p {A}>'
        f'<a:r>'
        f'<a:rPr lang="en-GB" sz="{sz}" b="{b}">'
        f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
        f'<a:latin typeface="PT Sans"/>'
        f'<a:cs typeface="PT Sans"/>'
        f'</a:rPr>'
        f'<a:t>{safe}</a:t>'
        f'</a:r>'
        f'</a:p>'
        f'</txBody>'
        f'</sp>'
        f'<clientData/>'
        f'</twoCellAnchor>'
    )


def _find_drawing_for_sheet(unpacked_dir, sheet_name):
    """Resolve sheet name to its drawing file name (e.g., 'drawing1')."""
    wbtree = ET.parse(os.path.join(unpacked_dir, 'xl', 'workbook.xml'))
    sheets = wbtree.findall(f'{NS_MAIN}sheets/{NS_MAIN}sheet')
    name_to_rid = {s.get('name'): s.get(f'{NS_REL}id') for s in sheets}
    rid = name_to_rid.get(sheet_name)
    if not rid:
        return None
    rel_tree = ET.parse(os.path.join(unpacked_dir, 'xl', '_rels', 'workbook.xml.rels'))
    rid_to_target = {r.get('Id'): r.get('Target') for r in rel_tree.findall(f'{NS_PKG}Relationship')}
    target = rid_to_target.get(rid, '')
    m = re.search(r'sheet(\d+)\.xml', target)
    if not m:
        return None
    snum = m.group(1)
    rels_path = os.path.join(unpacked_dir, 'xl', 'worksheets', '_rels',
                             f'sheet{snum}.xml.rels')
    if not os.path.exists(rels_path):
        return None
    rt = ET.parse(rels_path)
    for r in rt.findall(f'{NS_PKG}Relationship'):
        t = r.get('Target', '')
        if 'drawing' in t and t.endswith('.xml'):
            mm = re.search(r'(drawing\d+)\.xml', t)
            if mm:
                return mm.group(1)
    return None


def _patch_drawing(unpacked_dir, drawing_name, boxes):
    """Append textboxes to the named drawing XML file."""
    p = os.path.join(unpacked_dir, 'xl', 'drawings', f'{drawing_name}.xml')
    with open(p) as f:
        c = f.read()
    if not c.endswith('</wsDr>'):
        raise ValueError(f"Unexpected drawing structure: {p}")
    body = c[:-len('</wsDr>')]
    for box in boxes:
        body += _make_textbox(*box)
    body += '</wsDr>'
    with open(p, 'w') as f:
        f.write(body)


# Default Cover textbox layout (positions match WTP guide cover image)
# (id, name, textlink, from_col, from_row, to_col, to_row, font_size, colour, bold, pi_key)
COVER_BOXES_SPEC = [
    (1001, 'CoverProjectName', "'Project Info'!$C$6",  1, 6, 7, 9,  24, 'FFFFFF', False, 'C6'),
    (1002, 'CoverClient',      "'Project Info'!$C$7",  1, 9, 7, 11, 18, 'FFC425', False, 'C7'),
    (1003, 'CoverAddress',     "'Project Info'!$C$8",  1, 11, 7, 13, 12, 'FFFFFF', False, 'C8'),
    (1004, 'CoverContractRef', "'Project Info'!$C$9",  1, 13, 7, 15, 10, 'FFFFFF', False, 'C9'),
    (1005, 'CoverSubtitle',    "'Project Info'!$C$18", 1, 16, 7, 18, 12, 'FFFFFF', True,  'C18'),
    (1006, 'CoverReference',   "'Project Info'!$C$19", 1, 18, 7, 19, 10, 'FFFFFF', False, 'C19'),
]
APPENDIX_BOXES_SPEC = [
    (2001, 'AppendixTitle',       "'Project Info'!$C$20", 1, 6, 7, 9,  36, 'FFFFFF', True,  'C20'),
    (2002, 'AppendixDescription', "'Project Info'!$C$17", 1, 10, 7, 12, 14, 'FFC425', False, 'C17'),
    (2003, 'AppendixReference',   "'Project Info'!$C$15", 1, 13, 7, 14, 10, 'FFFFFF', False, 'C15'),
]


def inject_cover_textboxes(unpacked_dir, project_info_values):
    """
    Inject textboxes into Cover and Appendix Cover drawing files.

    project_info_values: dict mapping cell coords (C6-C20) to string values
                          to embed as textbox static text.
    """
    cover_drawing = _find_drawing_for_sheet(unpacked_dir, 'Cover')
    appendix_drawing = _find_drawing_for_sheet(unpacked_dir, 'Appendix Cover')

    if cover_drawing:
        boxes = []
        for spec in COVER_BOXES_SPEC:
            *meta, pi_key = spec
            text = project_info_values.get(pi_key, '')
            boxes.append(tuple(meta) + (text,))
        _patch_drawing(unpacked_dir, cover_drawing, boxes)
        print(f"Injected {len(boxes)} textboxes into Cover ({cover_drawing}.xml)")

    if appendix_drawing:
        boxes = []
        for spec in APPENDIX_BOXES_SPEC:
            *meta, pi_key = spec
            text = project_info_values.get(pi_key, '')
            boxes.append(tuple(meta) + (text,))
        _patch_drawing(unpacked_dir, appendix_drawing, boxes)
        print(f"Injected {len(boxes)} textboxes into Appendix Cover ({appendix_drawing}.xml)")
