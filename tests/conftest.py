import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import fitz  # noqa: E402


def make_pdf(path, name, pages=2, width=595, height=842, rotation=0, text_prefix=None):
    """Write a small PDF whose every page carries a unique marker line."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=width, height=height)
        if rotation:
            page.set_rotation(rotation)
        page.insert_text((72, 72), f"MARK {text_prefix or name} page {i + 1}", fontsize=14)
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def make_tree(tmp_path):
    """Factory: build a folder tree from {relative/path/file.pdf: pages}."""

    def _make(spec, base="bundle"):
        root = tmp_path / base
        root.mkdir(parents=True, exist_ok=True)
        for rel, pages in spec.items():
            full = root / rel
            if rel.endswith("/"):
                full.mkdir(parents=True, exist_ok=True)
                continue
            if isinstance(pages, dict):
                make_pdf(str(full), os.path.splitext(os.path.basename(rel))[0], **pages)
            else:
                make_pdf(str(full), os.path.splitext(os.path.basename(rel))[0], pages=pages)
        return str(root)

    return _make


def page_text(doc, index):
    return doc.load_page(index).get_text()


def marker_of(doc, index):
    """Return 'name page N' from the MARK line on a page, or None."""
    for line in page_text(doc, index).splitlines():
        if line.startswith("MARK "):
            return line[5:].strip()
    return None


def toc_links(doc, cover_pages, toc_pages):
    """All GOTO link targets (0-based) on the TOC pages, in reading order."""
    out = []
    for pno in range(cover_pages, cover_pages + toc_pages):
        links = doc.load_page(pno).get_links()
        links.sort(key=lambda l: (l["from"].y0, l["from"].x0))
        out.extend(l["page"] for l in links if l["kind"] == fitz.LINK_GOTO)
    return out
