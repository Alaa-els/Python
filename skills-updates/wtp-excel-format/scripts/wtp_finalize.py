"""
WTP Excel Format - Finalisation.

Stage B finalisation:
  - Inject WT theme XML (replaces theme1.xml)
  - Clean docProps for Alaa Elsayed authorship
  - Strip AI/library traces

Public functions:
  inject_theme(unpacked_dir, theme_xml_path)
  clean_metadata(unpacked_dir, author='Alaa Elsayed', company='WT Partnership',
                 title=None, subject=None)
"""
import os, re, shutil, datetime


def inject_theme(unpacked_dir, theme_xml_path):
    """Replace xl/theme/theme1.xml with WT theme."""
    target = os.path.join(unpacked_dir, 'xl', 'theme', 'theme1.xml')
    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.copy(theme_xml_path, target)
    print("Injected WT theme")


def clean_metadata(unpacked_dir, author='Alaa Elsayed', company='WT Partnership',
                   title=None, subject=None):
    """Strip AI traces and set authorship in docProps/."""
    now = datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    docprops = os.path.join(unpacked_dir, 'docProps')
    os.makedirs(docprops, exist_ok=True)

    def _esc(s):
        return (s.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;').replace('"', '&quot;'))

    title_xml = f'<dc:title>{_esc(title)}</dc:title>' if title else ''
    subject_xml = f'<dc:subject>{_esc(subject)}</dc:subject>' if subject else ''
    author_esc = _esc(author)
    company_esc = _esc(company)

    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:creator>{author_esc}</dc:creator>'
        f'<cp:lastModifiedBy>{author_esc}</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        f'{title_xml}{subject_xml}'
        '</cp:coreProperties>'
    )
    with open(os.path.join(docprops, 'core.xml'), 'w', encoding='utf-8') as f:
        f.write(core_xml)

    app_path = os.path.join(docprops, 'app.xml')
    if os.path.exists(app_path):
        with open(app_path, 'r', encoding='utf-8') as f:
            app_xml = f.read()
        app_xml = re.sub(r'<Application>[^<]*</Application>',
                         '<Application>Microsoft Excel</Application>', app_xml)
        app_xml = re.sub(r'<AppVersion>[^<]*</AppVersion>',
                         '<AppVersion>16.0300</AppVersion>', app_xml)
        if '<Company>' in app_xml:
            app_xml = re.sub(r'<Company>[^<]*</Company>',
                             f'<Company>{company_esc}</Company>', app_xml)
        else:
            app_xml = app_xml.replace('</Properties>',
                                      f'<Company>{company_esc}</Company></Properties>')
        app_xml = re.sub(r'<Manager>[^<]*</Manager>', '<Manager></Manager>', app_xml)
        with open(app_path, 'w', encoding='utf-8') as f:
            f.write(app_xml)
    else:
        app_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Excel</Application>'
            '<DocSecurity>0</DocSecurity>'
            '<ScaleCrop>false</ScaleCrop>'
            f'<Company>{company_esc}</Company>'
            '<LinksUpToDate>false</LinksUpToDate>'
            '<SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged>'
            '<AppVersion>16.0300</AppVersion>'
            '</Properties>'
        )
        with open(app_path, 'w', encoding='utf-8') as f:
            f.write(app_xml)

    # Remove custom.xml if it contains AI/library traces
    custom_path = os.path.join(docprops, 'custom.xml')
    if os.path.exists(custom_path):
        with open(custom_path, 'r', encoding='utf-8') as f:
            c = f.read()
        if any(s.lower() in c.lower() for s in
               ['claude', 'anthropic', 'python', 'openpyxl', 'xlsxwriter',
                'reportlab', 'libre']):
            os.remove(custom_path)
            ct_path = os.path.join(unpacked_dir, '[Content_Types].xml')
            with open(ct_path, 'r', encoding='utf-8') as f:
                ct = f.read()
            ct = re.sub(r'<Override[^/]*custom\.xml[^/]*/>', '', ct)
            with open(ct_path, 'w', encoding='utf-8') as f:
                f.write(ct)
            rel_path = os.path.join(unpacked_dir, '_rels', '.rels')
            if os.path.exists(rel_path):
                with open(rel_path, 'r', encoding='utf-8') as f:
                    rels = f.read()
                rels = re.sub(r'<Relationship[^/]*custom\.xml[^/]*/>', '', rels)
                with open(rel_path, 'w', encoding='utf-8') as f:
                    f.write(rels)
    print(f"Metadata cleaned (author={author}, company={company})")
