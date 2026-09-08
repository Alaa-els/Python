"""Extraction modes, page manifest operations and per-folder first-page runs."""
import os

import fitz
import pytest

import bundle_engine as be
from conftest import make_pdf, marker_of, page_text, toc_links


def plan_for(root, **extra):
    opts = {"base_directory": root, "claim_title": "Manifest", "separator_scope": "none",
            "policy": {"mode": "automatic", "layout": "preserve"}}
    opts.update(extra)
    plan = be.plan_bundle(root, opts["policy"], opts)
    return plan, opts


def build_plan(plan, opts, log=None):
    opts = dict(opts)
    opts["plan"] = plan
    return be.build_bundle_sync(opts, log or (lambda s: None))


def doc_named(plan, name):
    return [d for d in plan.documents() if d.source_name == name][0]


def markers(path):
    with fitz.open(path) as doc:
        return [marker_of(doc, i) for i in range(doc.page_count)]


# ---------------------------------------------------------------- range parsing

def test_parse_page_ranges():
    assert be.parse_page_ranges("", 4) == [0, 1, 2, 3]
    assert be.parse_page_ranges("1-2, 4", 4) == [0, 1, 3]
    assert be.parse_page_ranges("3-", 5) == [2, 3, 4]
    assert be.parse_page_ranges("-2", 5) == [0, 1]
    for bad in ("0", "5", "3-2", "a", "1-9", ","):
        with pytest.raises(ValueError):
            be.parse_page_ranges(bad, 4)


# ---------------------------------------------------------------- extraction modes

@pytest.fixture
def tree(make_tree):
    return make_tree({"WIR 001.pdf": 3, "WIR 002.pdf": 1, "Sub/WIR 003.pdf": 4})


def test_mode_all_default(tree):
    plan, opts = plan_for(tree)
    assert plan.selection["mode"] == "all"
    out, plan, info = build_plan(plan, opts)
    assert info["pages_included"] == 8 and info["pages_excluded"] == 0


def test_mode_first_page_is_first_source_page(tree):
    plan, opts = plan_for(tree, page_selection={"mode": "first"}, separator_scope="every_document",
                          include_cover=True, cover_prepared_by="x")
    out, plan, info = build_plan(plan, opts)
    with fitz.open(out) as doc:
        found = [m for m in markers(out) if m]
        assert found == ["WIR 001 page 1", "WIR 002 page 1", "WIR 003 page 1"]
        # cover + toc + 3 top-level separators + Sub separator + 3 pages
        assert doc.page_count == 2 + 3 + 1 + 3
        assert info["pages_included"] == 3 and info["pages_excluded"] == 0
        for link, sec in zip(toc_links(doc, 1, 1)[:3], plan.sections):
            assert link == sec.start - 1


def test_mode_first_n(tree):
    plan, opts = plan_for(tree, page_selection={"mode": "first_n", "n": 2})
    out, plan, info = build_plan(plan, opts)
    assert [m for m in markers(out) if m] == ["WIR 001 page 1", "WIR 001 page 2", "WIR 002 page 1",
                                              "WIR 003 page 1", "WIR 003 page 2"]


def test_mode_ranges_with_default_and_invalid(tree):
    plan, opts = plan_for(tree, page_selection={"mode": "ranges",
                                                "ranges": {"WIR 001.pdf": "3,1", "WIR 003.pdf": "9"},
                                                "default_range": "1"})
    assert any("WIR 003.pdf" in w for w in plan.warnings)
    out, plan, info = build_plan(plan, opts)
    assert [m for m in markers(out) if m] == ["WIR 001 page 1", "WIR 001 page 3", "WIR 002 page 1",
                                              "WIR 003 page 1", "WIR 003 page 2", "WIR 003 page 3",
                                              "WIR 003 page 4"]


# ---------------------------------------------------------------- manifest operations

def test_move_rotate_exclude_and_targets(tree):
    plan, opts = plan_for(tree, generate_toc=True, separator_scope="none")
    w1 = doc_named(plan, "WIR 001.pdf")
    w3 = doc_named(plan, "WIR 003.pdf")
    assert be.move_pages(w1, [2], -5) == [0]            # clamped to the document start
    assert [r["src"] for r in w1.page_refs] == [2, 0, 1]
    be.rotate_pages(w1, [0], 90)
    be.rotate_pages(w1, [0], 90)
    be.set_pages_included(w1, [1], False)               # exclude original page 1
    be.set_pages_included(w3, [0, 1], False)            # first retained page of WIR 003 is page 3
    out, plan, info = build_plan(plan, opts)
    with fitz.open(out) as doc:
        assert [m for m in markers(out) if m] == ["WIR 001 page 3", "WIR 001 page 2", "WIR 002 page 1",
                                                  "WIR 003 page 3", "WIR 003 page 4"]
        assert doc.load_page(1).rotation == 180
        assert info["pages_excluded"] == 3 and info["pages_rotated"] == 1
        links = toc_links(doc, 0, 1)
        sub = plan.sections[2]
        # Sub section and its document link to WIR 003 page 3 (first retained)
        assert "MARK WIR 003 page 3" in page_text(doc, links[-1])
        assert sub.start == links[-2] + 1 and doc_named(plan, "WIR 003.pdf").start == sub.start
        bm = {t[1]: t[2] for t in doc.get_toc()}
        assert bm["WIR 003"] == sub.start


def test_excluding_whole_document_prunes_entry_and_relabels(tree):
    plan, opts = plan_for(tree, separator_scope="top_level")
    w2 = doc_named(plan, "WIR 002.pdf")
    be.set_pages_included(w2, [0], False)
    out, plan, info = build_plan(plan, opts)
    titles = [s.title for s in plan.sections]
    assert titles == ["WIR 001", "Sub"]
    assert [s.label for s in plan.sections] == ["A", "B"]
    with fitz.open(out) as doc:
        text = page_text(doc, 0)
        assert "WIR 002" not in text
        assert "Appendix B: Sub" in [t[1] for t in doc.get_toc()]
        assert doc.page_count == 1 + 2 + 3 + 4


def test_move_node_reorders_sections_and_documents(make_tree):
    root = make_tree({"F/a.pdf": 1, "F/b.pdf": 1, "F/c.pdf": 1, "G/x.pdf": 1})
    plan, opts = plan_for(root)
    f = plan.sections[0]
    assert be.move_node(plan, f.children[2], -2) == 0
    assert [d.title for d in f.children] == ["c", "a", "b"]
    assert be.move_node(plan, plan.sections[1], -1) == 0
    assert be.move_node(plan, plan.sections[0], -1) == 0   # clamped
    out, plan, _ = build_plan(plan, opts)
    assert [s.title for s in plan.sections] == ["G", "F"]
    assert [s.label for s in plan.sections] == ["A", "B"]
    assert [m for m in markers(out) if m] == ["x page 1", "c page 1", "a page 1", "b page 1"]


def test_snapshot_undo_and_reset(tree):
    plan, _ = plan_for(tree)
    snap = be.snapshot_plan(plan)
    w1 = doc_named(plan, "WIR 001.pdf")
    be.set_pages_included(w1, [0, 1, 2], False)
    assert be.manifest_totals(plan)["pages_excluded"] == 3
    assert be.manifest_totals(snap)["pages_excluded"] == 0
    be.apply_page_selection(plan, {"mode": "all"})
    assert be.manifest_totals(plan)["pages_excluded"] == 0
    manifest = be.plan_to_manifest(plan)
    assert manifest["version"] == 1 and len(manifest["sections"]) == 3


# ---------------------------------------------------------------- per-folder first pages

def test_per_folder_first_pages_like_the_scripts(make_tree):
    root = make_tree({"Villa 1/A (x).pdf": 2, "Villa 1/B.pdf": 3, "Villa 1/Inner [2]/C.pdf": 2,
                      "Villa 1/Old#Combined_FirstPages.pdf": 1, "Villa 1/NoPdf/": None,
                      "top.pdf": 2})
    with open(os.path.join(root, "Villa 1", "Bad.pdf"), "wb") as f:
        f.write(b"nope")
    log = []
    results = be.build_per_folder_first_pages(root, {"recursive": True}, log.append)
    outputs = {os.path.relpath(r["folder"], root): r for r in results}
    base = os.path.basename(root)
    assert outputs["."]["output"].endswith(f"{base}#Combined_FirstPages.pdf")
    assert outputs["Villa 1"]["output"].endswith("Villa 1#Combined_FirstPages.pdf")
    assert outputs[os.path.join("Villa 1", "Inner [2]")]["output"].endswith(
        "Inner -2-#Combined_FirstPages.pdf")
    assert outputs[os.path.join("Villa 1", "NoPdf")]["output"] is None
    assert markers(outputs["Villa 1"]["output"]) == ["A (x) page 1", "B page 1"]
    assert outputs["Villa 1"]["skipped"] == [("Bad.pdf", outputs["Villa 1"]["skipped"][0][1])]
    assert "cannot be opened" in outputs["Villa 1"]["skipped"][0][1]
    assert os.path.exists(os.path.join(root, "Villa 1", "Old#Combined_FirstPages.pdf"))
    # rerun: previous outputs are not inputs and are not overwritten
    results2 = be.build_per_folder_first_pages(root, {"recursive": False}, log.append)
    assert len(results2) == 1
    assert results2[0]["output"].endswith(f"{base}#Combined_FirstPages_rev.01.pdf")
    assert markers(results2[0]["output"]) == ["top page 1"]
    with fitz.open(outputs["Villa 1"]["output"]) as doc:
        assert doc.metadata["creator"] == be.PDF_CREATOR_TAG
        assert [t[1] for t in doc.get_toc()] == ["Appendix A: A (x)", "Appendix B: B"]


def test_combined_first_pages_bundle_with_grouping(make_tree):
    root = make_tree({"Villa 1/A.pdf": 2, "Villa 2/B.pdf": 3, "Villa 2/C.pdf": 1})
    plan, opts = plan_for(root, page_selection={"mode": "first"}, separator_scope="top_level",
                          generate_toc=True)
    out, plan, info = build_plan(plan, opts)
    with fitz.open(out) as doc:
        assert doc.page_count == 1 + 2 + 3
        assert [m for m in markers(out) if m] == ["A page 1", "B page 1", "C page 1"]
        assert [t[1] for t in doc.get_toc()][1:] == ["Appendix A: Villa 1", "A", "Appendix B: Villa 2",
                                                     "B", "C"]
    # non-recursive: direct files only
    plan, opts = plan_for(root, page_selection={"mode": "first"}, direct_files_only=True)
    assert plan.sections == []


# ---------------------------------------------------------------- thumbnails

def test_thumbnail_cache_bounds_and_aspect(make_tree):
    root = make_tree({"land.pdf": {"pages": 3, "width": 1190, "height": 842},
                      "port.pdf": {"pages": 2, "width": 595, "height": 842}})
    cache = be.ThumbnailCache(max_items=3, max_open_docs=1)
    ppm, w, h = cache.get(os.path.join(root, "land.pdf"), 0, 0, 160)
    assert w == 160 and h < w and ppm.startswith(b"P6")
    ppm, w, h = cache.get(os.path.join(root, "land.pdf"), 0, 90, 160)
    assert h == 160 and w < h
    for i in range(3):
        cache.get(os.path.join(root, "land.pdf"), i, 0, 120)
    cache.get(os.path.join(root, "port.pdf"), 1, 0, 120)
    assert len(cache._items) == 3 and len(cache._docs) == 1
    cache.clear()
    assert not cache._items and not cache._docs
