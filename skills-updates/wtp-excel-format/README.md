# WTP Excel Format - Installation & Usage

## Installation

Place the entire `wtp-excel-format/` folder under your skills directory:

```
/mnt/skills/user/wtp-excel-format/
├── SKILL.md
├── README.md
├── assets/
│   ├── cover_bg.jpg
│   ├── appendix_bg.jpg
│   ├── letterhead.png
│   └── wt_theme.xml
└── scripts/
    ├── wtp_styles.py
    ├── wtp_textboxes.py
    ├── wtp_letterhead.py
    ├── wtp_finalize.py
    └── build_wtp_workbook.py
```

Once installed, the skill's description triggers Claude to apply it whenever you ask for an Excel deliverable.

## Manual usage example (Python)

```python
from openpyxl import load_workbook
import sys
sys.path.insert(0, '/mnt/skills/user/wtp-excel-format/scripts')

from wtp_styles import (
    register_named_styles, build_project_info_tab,
    define_named_ranges, apply_wtp_page_setup,
    add_brand_image_cover_shell, apply_table_pattern,
)
from build_wtp_workbook import finalise_workbook

ASSETS = '/mnt/skills/user/wtp-excel-format/assets'

# Stage A: openpyxl build
wb = load_workbook('your_source.xlsx')
register_named_styles(wb)

# Project Info tab with your project's fields
project_fields = [
    ('Project Name',       'Your Project Name'),
    ('Client',             'Your Client'),
    ('Address',            'Project address'),
    ('Contract Reference', 'Contract ref'),
    ('Contractor',         'Contractor name'),
    ('Engineer',           'Engineer name'),
    ('Cost Consultant',    'WT Partnership'),
    ('Document Title',     'Document title'),
    ('Document Date',      '5 May 2026'),
    ('Document Reference', 'WTP-XXX-NNN'),
    ('Appendix Letter',    'A'),
    ('Appendix Description', 'Backup calculations'),
]
pi_values = build_project_info_tab(wb, fields=project_fields)
define_named_ranges(wb)

# Cover and Appendix tabs (image background only - textboxes added in Stage B)
add_brand_image_cover_shell(wb, 'Cover', f'{ASSETS}/cover_bg.jpg', 1)
add_brand_image_cover_shell(wb, 'Appendix Cover', f'{ASSETS}/appendix_bg.jpg', 2)

# Apply page setup and table patterns to your content tabs
ws = wb['Your Sheet']
apply_wtp_page_setup(ws, orientation='portrait')
apply_table_pattern(
    ws,
    header_row=6,
    max_col=10,
    description_col='C',
    forecast_col='I',  # cost forecast column - bracketed in YELLOW
    banner_check_col='A',
)

# Hide cover tabs
wb['Cover'].sheet_state = 'hidden'
wb['Appendix Cover'].sheet_state = 'hidden'
wb.active = wb.sheetnames.index('Project Info')

wb.save('your_workbook_stage_a.xlsx')

# Stage B: textboxes + letterhead + theme + metadata
finalise_workbook(
    input_xlsx='your_workbook_stage_a.xlsx',
    output_xlsx='your_workbook_FINAL.xlsx',
    project_info_values=pi_values,
    portrait_tabs=['Your Sheet', 'Other Portrait Sheet'],
    assets_dir=ASSETS,
    title='Document title',
    subject='Document subject',
)
```

## What you get

- 9 named cell styles (`WT Title`, `WT Heading 1-4`, `WT Headline`, `WT Table Heading 1-3`)
- Strict WT brand palette (only the 8 listed hex colours)
- PT Sans typography
- Project Info source-of-truth tab (visible, yellow tab)
- Cover and Appendix Cover tabs (hidden, brand image + linked textboxes)
- Letterhead on portrait content tabs
- A4 page setup with WTP margins
- Clean metadata (Alaa Elsayed / WT Partnership)

## Critical notes

1. **Never run openpyxl save after Stage B** — it strips raw textbox/letterhead XML.
2. **The bundled letterhead has the WTP Australia footer baked in.** Replace the file in `assets/letterhead.png` with a KSA-localised version when WTP issues one.
3. **Cost forecast borders are YELLOW**, not blue — this is deliberate. Yellow stands out against the blue-dominated headings.
4. **Cover textboxes are pre-computed at build time.** To make them dynamically link in Excel: open each textbox, type `='Project Info'!$C$N` in the formula bar, save.
5. **Run the verify_metadata.py from file-authorship-metadata skill** after every build to confirm zero AI/library traces.

## Brand palette reference (memorise this)

| Use | Hex |
|---|---|
| Primary accent (heading text, header fill) | `#0045BF` (WT Royal Blue) |
| Cost forecast border highlight, WT logo accent | `#FFC425` (WT Yellow) |
| Section banner / description column fill | `#EDF0F2` (Light grey-blue) |
| Appendix cover background only | `#051641` (Deep Navy) |
| White text on dark fill | `#FFFFFF` |
| Body text | `#000000` |

No teal, no other blues, no other greys (apart from border tints `#BFBFBF` and `#D9D9D9`).
