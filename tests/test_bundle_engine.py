"""Generated-PDF regression tests for the bundle engine.

Every test builds a real bundle with PyMuPDF/fpdf2 and inspects the output:
page counts, link destinations, bookmarks, footers and page geometry.
"""
import os
import json
import re

import fitz
import pytest

import bundle_engine as be
from conftest import make_pdf, marker_of, page_text, toc_links

TREE = {
    "Root Letter.pdf": 1,
    "A Notice/N1.pdf": 2,
    "A Notice/Sub1/S1.pdf": 3,
    "A Notice/Sub1/Deep/D1.pdf": 1,
    "B Records/R1.pdf": 2,
    "Empty/": None,
}
# documents in bundle order with their page counts
DOC_ORDER = [("Root Letter", 1), ("N1", 2), ("S1", 3), ("D1", 1), ("R1", 2)]
CONTENT_PAGES = sum(p for _, p in DOC_ORDER)


def build(root, **extra):
    opts = {"base_directory": root, "claim_title": "Test Bundle", "generate_toc": True,
            "separator_scope": "none", "include_cover": False,
            "policy": {"mode": "automatic", "layout": "preserve"}}
    opts.update(extra)
    out, plan, info = be.build_bundle_sync(opts, lambda s: None)
    return out, plan, info


def markers(doc):
    return [marker_of(doc, i) for i in range(doc.page_count)]


# ---------------------------------------------------------------- separators

@pytest.mark.parametrize("scope,expected_seps", [
    ("none", 0),
    ("top_level", 3),          # Root Letter, A Notice, B Records
    ("all_sections", 5),       # + Sub1, Deep
    ("every_document", 9),     # + N1, S1, D1, R1 (root doc has no extra page)
])
def test_separator_scope_page_counts(make_tree, scope, expected_seps):
    root = make_tree(TREE)
    out, plan, info = build(root, separator_scope=scope)
    with fitz.open(out) as doc:
        assert info["toc_pages"] == 1
        assert doc.page_count == 1 + expected_seps + CONTENT_PAGES
        # every source page appears exactly once, in order
        found = [m for m in markers(doc) if m]
        expected = [f"{n} page {k + 1}" for n, p in DOC_ORDER for k in range(p)]
        assert found == expected


@pytest.mark.parametrize("scope", be.SEPARATOR_SCOPES)
@pytest.mark.parametrize("cover", [False, True])
def test_links_target_first_content_or_separator(make_tree, scope, cover):
    """Each TOC link must open the section's separator when it has one, and
    the first actual content page of the section when it has none."""
    root = make_tree(TREE)
    out, plan, info = build(root, separator_scope=scope, include_cover=cover,
                            cover_prepared_by="QS", cover_project="Proj")
    with fitz.open(out) as doc:
        offset = info["cover_pages"] + info["toc_pages"]
        assert info["cover_pages"] == (1 if cover else 0)
        links = toc_links(doc, info["cover_pages"], info["toc_pages"])
        entries = be.build_toc_entries(plan, 0)  # plan pages already final
        assert len(links) == len(entries)
        for link, node in zip(links, _toc_nodes(plan)):
            assert link == node.start - 1, f"{node.title}: link {link} vs start {node.start}"
            target_text = page_text(doc, link)
            if node.separator:
                assert node.title in target_text
                assert "MARK" not in target_text
            else:
                first_doc = node.documents()[0] if node.kind == "section" else node
                assert f"MARK {first_doc.title} page 1" in target_text
        # no phantom pages: every page is either a marker page or a generated page
        for i in range(offset, doc.page_count):
            assert marker_of(doc, i) or any(n.sep_page == i + 1 for n in plan.walk())


def _toc_nodes(plan):
    nodes = []
    for sec in plan.sections:
        nodes.append(sec)
        if sec.root_document:
            continue
        for n in sec.walk():
            if n is not sec and n.toc:
                nodes.append(n)
    return nodes


def test_no_separators_section_ranges_and_footers(make_tree):
    root = make_tree(TREE)
    out, plan, info = build(root, separator_scope="none")
    with fitz.open(out) as doc:
        a, b, c = plan.sections
        assert (a.start, a.end) == (2, 2)
        assert (b.start, b.end) == (3, 8)
        assert (c.start, c.end) == (9, 10)
        assert b.content_start == 3
        # footer text on the last page of Appendix B
        txt = page_text(doc, b.end - 1)
        assert "App B - P006 / 006" in txt
        assert "Appendix B: A Notice" in txt
        # TOC page carries no footer
        assert "App " not in page_text(doc, 0).split("Contents")[0]


def test_legacy_separator_names_still_work(make_tree):
    root = make_tree(TREE)
    out, plan, _ = build(root, separator_scope=None,
                         policy={"mode": "automatic", "layout": "preserve", "separators": "folder_only"})
    with fitz.open(out) as doc:
        assert doc.page_count == 1 + 5 + CONTENT_PAGES
    assert be.normalise_separator_scope("every_pdf") == "every_document"
    assert be.normalise_separator_scope("none") == "none"
    assert be.normalise_separator_scope("garbage", "top_level") == "top_level"


# ---------------------------------------------------------------- TOC / bookmarks

def test_bookmarks_independent_of_toc(make_tree):
    root = make_tree(TREE)
    out, plan, info = build(root, generate_toc=False, generate_bookmarks=True, separator_scope="none")
    with fitz.open(out) as doc:
        assert info["toc_pages"] == 0
        toc = doc.get_toc()
        titles = [t[1] for t in toc]
        assert "Appendix A: Root Letter" in titles and "Appendix B: A Notice" in titles
        by_title = {t[1]: t for t in toc}
        assert by_title["Appendix B: A Notice"][2] == 2   # first content page
        assert by_title["S1"][2] == 4
        assert by_title["Deep"][2] == 7
    out, plan, info = build(root, generate_toc=True, generate_bookmarks=False)
    with fitz.open(out) as doc:
        assert doc.get_toc() == []
        assert info["toc_pages"] == 1


def test_toc_and_bookmark_document_rows_toggle(make_tree):
    root = make_tree(TREE)
    out, plan, info = build(root, toc_include_documents=False, bookmark_include_documents=False)
    with fitz.open(out) as doc:
        text = page_text(doc, 0)
        assert "N1" not in text and "Sub1" in text and "Deep" in text
        titles = [t[1] for t in doc.get_toc()]
        assert "S1" not in titles and "Sub1" in titles
        # links only for sections
        assert len(toc_links(doc, 0, 1)) == 5


def test_multi_page_toc_offsets(make_tree):
    spec = {f"Docs/Item {i:03d}.pdf": 1 for i in range(1, 70)}
    spec["Zed.pdf"] = 1
    root = make_tree(spec)
    out, plan, info = build(root, separator_scope="none", include_cover=True,
                            cover_prepared_by="x")
    with fitz.open(out) as doc:
        assert info["toc_pages"] >= 2
        links = toc_links(doc, info["cover_pages"], info["toc_pages"])
        nodes = _toc_nodes(plan)
        assert len(links) == len(nodes)
        for link, node in zip(links, nodes):
            assert link == node.start - 1
            first_doc = node.documents()[0]
            assert f"MARK {first_doc.title} page 1" in page_text(doc, link)
        # displayed page numbers in the TOC text match actual positions
        zed = [n for n in plan.sections if n.title == "Zed"][0]
        toc_text = "".join(page_text(doc, p) for p in range(info["cover_pages"],
                                                             info["cover_pages"] + info["toc_pages"]))
        assert re.search(rf"Zed\s+\[pp {zed.start}-{zed.end}\]", toc_text)


# ---------------------------------------------------------------- policy / structure

def test_empty_and_skipped_folders_create_no_phantoms(make_tree):
    root = make_tree({"Keep/K1.pdf": 1, "Skipme/X.pdf": 1, "Skipme/Inner/Y.pdf": 1,
                      "Keep/Empty/": None, "Keep/AllSkipped/Z.pdf": 1})
    policy = {"mode": "manual", "decisions": {"Skipme / Inner": {"action": "skip"},
                                              "Keep / AllSkipped": {"action": "skip"}},
              "appendix_decisions": {}}
    out, plan, info = build(root, policy=policy, separator_scope="all_sections")
    titles = [s.title for s in plan.sections]
    assert titles == ["Keep", "Skipme"]
    assert plan.sections[0].label == "A" and plan.sections[1].label == "B"
    with fitz.open(out) as doc:
        # TOC + 2 top-level separators + 2 content pages, nothing for skipped/empty folders
        assert doc.page_count == 1 + 2 + 2
        assert "AllSkipped" not in page_text(doc, 0)
        assert "Inner" not in page_text(doc, 0)
    assert "Keep / AllSkipped" in plan.skipped and "Skipme / Inner" in plan.skipped


def test_flatten_and_numeric_skip_layouts(make_tree):
    root = make_tree({"Pack/01/a.pdf": 1, "Pack/01/b.pdf": 1, "Pack/Named/c.pdf": 1})
    out, plan, _ = build(root, policy={"mode": "automatic", "layout": "preserve_skip_numeric"})
    kinds = [(n.kind, n.title, n.level) for n in plan.sections[0].walk()]
    assert ("section", "01", 2) not in kinds
    assert ("section", "Named", 2) in kinds
    out, plan, _ = build(root, policy={"mode": "automatic", "layout": "flatten"})
    assert all(n.kind == "document" for n in plan.sections[0].children)
    assert [n.title for n in plan.sections[0].children] == ["a", "b", "c"]


def test_collapsed_folder_toc_vs_bookmarks(make_tree):
    root = make_tree({"App/Coll/x.pdf": 1, "App/Coll/Sub/y.pdf": 1, "App/z.pdf": 1})
    policy = {"mode": "manual", "decisions": {"App / Coll": {"action": "show", "depth": 1}},
              "appendix_decisions": {}}
    out, plan, info = build(root, policy=policy, separator_scope="all_sections")
    with fitz.open(out) as doc:
        toc_text = page_text(doc, 0)
        assert "Coll" in toc_text and " x" not in toc_text and "Sub" not in toc_text
        bm = [t[1] for t in doc.get_toc()]
        assert "x" in bm and "Sub" in bm and "y" in bm
        # collapsed subtree has no internal separators; Coll itself has one
        assert doc.page_count == 1 + 1 + 1 + 3


def test_manual_ordering_and_titles(make_tree):
    root = make_tree({"Zeta.pdf": 1, "Alpha.pdf": 1, "Folder/q.pdf": 1, "Folder/p.pdf": 1})
    with open(os.path.join(root, "_order.txt"), "w") as f:
        f.write("# custom\nZeta.pdf\nFolder\n")
    with open(os.path.join(root, "Folder", "_order.txt"), "w") as f:
        f.write("q.pdf\n")
    out, plan, _ = build(root, manual_descriptions=True,
                         manual_titles={"Zeta.pdf": "The Zeta letter", "Folder": "Folder desc"})
    # root files always precede folders (original behaviour); _order.txt orders within each group
    assert [s.title for s in plan.sections] == ["The Zeta letter", "Alpha", "Folder desc"]
    folder = plan.sections[2]
    assert [d.title for d in folder.children] == ["q", "p"]


def test_natural_order_and_previous_outputs_excluded(make_tree):
    root = make_tree({"Dwg 10.pdf": 1, "Dwg 2.pdf": 1, "dwg 1.pdf": 1, "old#Combined_FirstPages.pdf": 1})
    out, plan, _ = build(root)
    assert [s.title for s in plan.sections] == ["dwg 1", "Dwg 2", "Dwg 10"]
    # a second run must not pick up the first output
    out2, plan2, _ = build(root, claim_title="Second")
    assert [s.title for s in plan2.sections] == ["dwg 1", "Dwg 2", "Dwg 10"]
    assert "Test Bundle.pdf" in plan2.previous_outputs
    assert plan2.skipped_documents == []


def test_corrupt_and_encrypted_inputs_are_reported(make_tree):
    root = make_tree({"Good.pdf": 1})
    with open(os.path.join(root, "Broken.pdf"), "wb") as f:
        f.write(b"%PDF-1.4 garbage")
    enc = fitz.open()
    enc.new_page().insert_text((72, 72), "secret")
    enc.save(os.path.join(root, "Locked.pdf"), encryption=fitz.PDF_ENCRYPT_AES_256,
             user_pw="pw", owner_pw="pw")
    enc.close()
    with open(os.path.join(root, "Empty.pdf"), "wb") as f:
        f.write(b"")
    out, plan, info = build(root)
    names = sorted(n for n, _ in plan.skipped_documents)
    assert names == ["Broken.pdf", "Empty.pdf", "Locked.pdf"]
    reasons = dict(plan.skipped_documents)
    assert "encrypted" in reasons["Locked.pdf"]
    assert info["skipped_documents"] == plan.skipped_documents
    with fitz.open(out) as doc:
        assert doc.page_count == 2


def test_no_content_raises(make_tree):
    root = make_tree({"Empty/": None})
    with pytest.raises(ValueError):
        build(root)


# ---------------------------------------------------------------- terminology / styles

def test_terminology_and_numbering(make_tree):
    root = make_tree({"Plan.pdf": 1, "Elev.pdf": 1})
    out, plan, _ = build(root, section_term="Drawing", numbering_style="numeric",
                         separator_scope="top_level", style={"separator": "band", "toc": "grid"})
    with fitz.open(out) as doc:
        assert "Drawing 01: Elev" in [t[1] for t in doc.get_toc()]
        sep = page_text(doc, 1)
        assert re.search(r"D R A W I N G\s+0 1", sep)
        assert "Dwg 01 - P002 / 002" in page_text(doc, 2)   # separator page is P001
    out, plan, _ = build(root, section_term="Custom", section_term_custom="Bundle Item",
                         numbering_style="none")
    with fitz.open(out) as doc:
        assert plan.sections[0].label == ""
        assert [t[1] for t in doc.get_toc()][1] == "Elev"


@pytest.mark.parametrize("preset", list(be.STYLE_PRESETS))
def test_every_preset_renders_cover_separators_and_toc(make_tree, preset):
    root = make_tree({"One.pdf": 1, "Folder/Two.pdf": 1})
    p = be.STYLE_PRESETS[preset]
    out, plan, info = build(root, style=p["style"], separator_scope="all_sections", include_cover=True,
                            cover_prepared_by="QS", cover_project="Project X",
                            cover_doc_type=p["terminology"]["cover_doc_type"],
                            section_term=p["terminology"]["section_term"],
                            numbering_style=p["terminology"]["numbering_style"], **p["footer"])
    with fitz.open(out) as doc:
        assert info["cover_pages"] == 1 and info["toc_pages"] == 1
        assert "Project X" in page_text(doc, 0)
        assert doc.page_count == 2 + 2 + 2
    data = be.render_style_preview(p["style"], p["terminology"]["section_term"],
                                   p["terminology"]["numbering_style"])
    with fitz.open(stream=data, filetype="pdf") as prev:
        assert prev.page_count == 4


def test_toc_density_changes_height(make_tree):
    entries = [{"level": 1, "label": "A", "title": f"Item {i}", "page": i} for i in range(1, 140)]
    counts = {}
    for density in be.TOC_DENSITIES:
        be.set_current_style({"toc_density": density, "toc": "ruled"})
        _, _, counts[density] = be.create_toc_page_hier("t", entries)
    be.set_current_style({"toc_density": "standard"})
    assert counts["compact"] <= counts["standard"] <= counts["spacious"]
    assert counts["compact"] < counts["spacious"]


# ---------------------------------------------------------------- drawings / footers

def test_drawing_geometry_rotation_and_footer_positions(make_tree):
    a0 = {"pages": 1, "width": 3370, "height": 2384}
    rot = {"pages": 1, "width": 3370, "height": 2384, "rotation": 90}
    root = make_tree({"Drawings/A0 sheet.pdf": a0, "Drawings/Rotated sheet.pdf": rot,
                      "Drawings/A4 note.pdf": 1})
    out, plan, _ = build(root, separator_scope="none", footer_position="top", footer_layout="left",
                         footer_inset_mm=10)
    with fitz.open(out) as doc:
        # natural order: A0 sheet, A4 note, Rotated sheet
        p_a0 = doc.load_page(1)
        assert (round(p_a0.rect.width), round(p_a0.rect.height)) == (3370, 2384)
        assert p_a0.rotation == 0
        p_rot = doc.load_page(3)
        assert p_rot.rotation == 90
        assert (round(p_rot.rect.width), round(p_rot.rect.height)) == (2384, 3370)
        # footer text sits near the displayed top edge and upright on the rotated page
        for pno in (1, 3):
            page = doc.load_page(pno)
            lines = [l for b in page.get_text("dict")["blocks"] for l in b.get("lines", [])
                     if f"P00{pno}" in "".join(s["text"] for s in l["spans"])]
            assert lines, "footer missing"
            bbox = fitz.Rect(lines[0]["bbox"]) * page.rotation_matrix
            assert bbox.y1 < page.rect.height * 0.05
            assert tuple(round(v) for v in lines[0]["dir"]) in ((1, 0), (0, -1), (0, 1), (-1, 0))
        # vector content preserved: no images were introduced
        assert not doc.load_page(1).get_images()
    out, plan, _ = build(root, separator_scope="none", footer_enabled=False)
    with fitz.open(out) as doc:
        assert "P001" not in page_text(doc, 1)


def test_footer_bottom_and_centre_layout(make_tree):
    root = make_tree({"Doc.pdf": 1})
    out, plan, _ = build(root, footer_layout="centre", separator_scope="none")
    with fitz.open(out) as doc:
        page = doc.load_page(1)
        words = [w for w in page.get_text("words") if w[4] == "P001"]
        assert words and words[0][3] > page.rect.height * 0.9


def test_compression_leaves_vector_pages_alone(make_tree):
    root = make_tree({"Doc.pdf": 1})
    out, _, _ = build(root, compression="small")
    with fitz.open(out) as doc:
        assert "MARK Doc page 1" in page_text(doc, 1)
        assert not doc.load_page(1).get_images()


# ---------------------------------------------------------------- settings

def test_settings_round_trip_and_legacy_merge(tmp_path):
    path = str(tmp_path / "settings.json")
    legacy = {"claim_title": "Old", "numbering_style": "numeric", "style_toc": "dotted",
              "unknown_key": 1, "separator_scope": "folder_only"}
    with open(path, "w") as f:
        json.dump(legacy, f)
    s = be.load_settings(path)
    assert s["claim_title"] == "Old" and s["style_toc"] == "dotted"
    assert s["separator_scope"] == "all_sections"
    assert s["section_term"] == "Appendix" and s["footer_enabled"] is True
    assert "unknown_key" not in s
    s["separator_scope"] = "none"
    assert be.save_settings(s, path)
    assert be.load_settings(path)["separator_scope"] == "none"
    assert be.load_settings(str(tmp_path / "missing.json")) == be.DEFAULT_SETTINGS


# ---------------------------------------------------------------- office overlay

def test_overlay_documents_are_bundled_in_folder_position(make_tree, tmp_path):
    root = make_tree({"Corr/A letter.pdf": 1, "Corr/C letter.pdf": 1})
    src = os.path.join(root, "Corr", "B letter.docx")
    with open(src, "wb") as f:
        f.write(b"PK\x03\x04fake")
    conv = make_pdf(str(tmp_path / "staging" / "B letter.pdf"), "B letter", pages=2)
    overlay = {be.norm_path(src): conv}
    # without the overlay the .docx is invisible
    assert be.count_folder_tree(root)["total_pdfs"] == 2
    assert be.count_folder_tree(root, overlay)["total_pdfs"] == 3
    opts = {"base_directory": root, "claim_title": "Ov", "separator_scope": "none",
            "policy": {"mode": "automatic", "layout": "preserve"}, "overlay": overlay}
    out, plan, _ = be.build_bundle_sync(opts, lambda s: None)
    assert [d.title for d in plan.sections[0].children] == ["A letter", "B letter", "C letter"]
    with fitz.open(out) as doc:
        assert "MARK B letter page 1" in page_text(doc, 2)


def test_cancel_event_stops_build(make_tree):
    import threading
    root = make_tree({"a.pdf": 1})
    ev = threading.Event()
    ev.set()
    with pytest.raises(be.BundleCancelled):
        build(root, cancel_event=ev)
