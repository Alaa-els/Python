"""
WTP Excel Format - End-to-end orchestrator.

Three-stage pipeline:
  Stage A: openpyxl build (caller's responsibility - styles, cells, Project Info,
                           Cover/Appendix shells, defined names, page setup)
  Stage B: this script - unpack, inject textboxes, inject letterhead,
                         inject theme, clean metadata, repack

Usage:
  from build_wtp_workbook import finalise_workbook

  # After your openpyxl save:
  finalise_workbook(
      input_xlsx='your_workbook.xlsx',
      output_xlsx='your_workbook_FINAL.xlsx',
      project_info_values={'C6': 'Project Name', 'C7': 'Client', ...},
      portrait_tabs=['Sheet1', 'Sheet2'],  # tabs to receive the letterhead
      assets_dir='assets/',
      title='Document title',
      subject='Document subject',
  )

OR call as CLI:
  python build_wtp_workbook.py --input X.xlsx --output Y.xlsx --pi pi.json --portrait Sheet1,Sheet2
"""
import os, sys, shutil, zipfile, json, argparse

# Allow running both as module and script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from wtp_textboxes import inject_cover_textboxes
from wtp_letterhead import inject_letterhead
from wtp_finalize import inject_theme, clean_metadata


def finalise_workbook(input_xlsx, output_xlsx, project_info_values,
                      portrait_tabs, assets_dir,
                      author='Alaa Elsayed', company='WT Partnership',
                      title=None, subject=None,
                      tmp_dir=None):
    """
    Run Stage B finalisation on an openpyxl-saved workbook.

    project_info_values: dict mapping cell coords (e.g. 'C6', 'C7'...) to strings
                          that will be embedded as Cover/Appendix textbox text
    portrait_tabs: list of sheet names to receive the legacyDrawingHF letterhead
    assets_dir: path to bundle directory containing cover_bg.jpg, appendix_bg.jpg,
                letterhead.png, wt_theme.xml
    """
    if tmp_dir is None:
        tmp_dir = output_xlsx + '.unpack'
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)

    print(f"Stage B: unpacking {input_xlsx}")
    with zipfile.ZipFile(input_xlsx, 'r') as z:
        z.extractall(tmp_dir)

    print("Stage B.1: inject Cover/Appendix textboxes")
    inject_cover_textboxes(tmp_dir, project_info_values)

    print("Stage B.2: inject letterhead on portrait tabs")
    letterhead_path = os.path.join(assets_dir, 'letterhead.png')
    if os.path.exists(letterhead_path) and portrait_tabs:
        n = inject_letterhead(tmp_dir, portrait_tabs, letterhead_path)
        print(f"  applied to {n} tabs")
    else:
        print("  skipped (no letterhead asset or no portrait tabs)")

    print("Stage B.3: inject WT theme")
    theme_path = os.path.join(assets_dir, 'wt_theme.xml')
    if os.path.exists(theme_path):
        inject_theme(tmp_dir, theme_path)

    print("Stage B.4: clean metadata")
    clean_metadata(tmp_dir, author=author, company=company,
                   title=title, subject=subject)

    print(f"Stage B.5: repack to {output_xlsx}")
    if os.path.exists(output_xlsx):
        os.remove(output_xlsx)
    with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(tmp_dir):
            for file in files:
                full = os.path.join(root, file)
                arc = os.path.relpath(full, tmp_dir)
                zf.write(full, arc)

    shutil.rmtree(tmp_dir)
    print(f"\nDone: {output_xlsx} ({os.path.getsize(output_xlsx):,} bytes)")
    print("\nNext: run scripts/verify_metadata.py from file-authorship-metadata "
          "to confirm 0 traces.")


def _cli():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', required=True, help='Stage A xlsx file')
    parser.add_argument('--output', required=True, help='Final xlsx output path')
    parser.add_argument('--pi', required=True,
                        help='JSON file with Project Info values dict')
    parser.add_argument('--portrait', default='',
                        help='Comma-separated portrait tab names for letterhead')
    parser.add_argument('--assets', required=True, help='Path to assets directory')
    parser.add_argument('--author', default='Alaa Elsayed')
    parser.add_argument('--company', default='WT Partnership')
    parser.add_argument('--title', default=None)
    parser.add_argument('--subject', default=None)
    args = parser.parse_args()

    with open(args.pi) as f:
        pi = json.load(f)
    portrait = [t.strip() for t in args.portrait.split(',') if t.strip()]
    finalise_workbook(args.input, args.output, pi, portrait,
                      args.assets, args.author, args.company,
                      args.title, args.subject)


if __name__ == '__main__':
    _cli()
