"""GUI smoke tests. They need tkinter and a display (run under xvfb-run on
Linux); they are skipped automatically otherwise.

The main window is driven through its real mainloop; modal dialogs are
replaced with deterministic stand-ins so the run cannot block."""
import os
import sys
import json
import importlib.machinery
import importlib.util

import pytest

try:
    import tkinter as tk  # noqa: F401
    _TK = True
except Exception:
    _TK = False

pytestmark = pytest.mark.skipif(not _TK or not os.environ.get("DISPLAY"),
                                reason="tkinter or DISPLAY not available")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_gui(tmp_home):
    os.makedirs(str(tmp_home), exist_ok=True)
    os.environ["HOME"] = str(tmp_home)
    os.environ["USERPROFILE"] = str(tmp_home)
    import bundle_engine as be
    be.SETTINGS_FILE = os.path.join(str(tmp_home), ".claims_bundle_builder.json")
    loader = importlib.machinery.SourceFileLoader("dbb_gui", os.path.join(ROOT, "document_bundle_builder.pyw"))
    spec = importlib.util.spec_from_loader("dbb_gui", loader)
    mod = importlib.util.module_from_spec(spec)
    sys.argv = ["document_bundle_builder.pyw"]
    sys.modules["dbb_gui"] = mod
    loader.exec_module(mod)
    return mod


def patch_dialogs(mod, answers=None):
    mb = mod.messagebox
    mb.askyesno = lambda *a, **k: True
    mb.showinfo = lambda *a, **k: None
    mb.showwarning = lambda *a, **k: None
    mb.showerror = lambda *a, **k: (_ for _ in ()).throw(AssertionError(f"error dialog: {a}"))
    mod.simpledialog.askstring = lambda *a, **k: "Tester"
    mod.show_discovery_popup = lambda parent, counts, term: "automatic"
    mod.show_automatic_options_dialog = lambda parent, default_layout="preserve": {
        "mode": "automatic", "layout": "preserve"}


def run_until(root, predicate, timeout_ms=60000):
    state = {"ok": False}

    def check(elapsed=0):
        if predicate():
            state["ok"] = True
            root.quit()
        elif elapsed > timeout_ms:
            root.quit()
        else:
            root.after(100, check, elapsed + 100)

    root.after(100, check)
    root.mainloop()
    return state["ok"]


@pytest.fixture
def gui(tmp_path):
    mod = load_gui(tmp_path / "home")
    patch_dialogs(mod)
    root = mod.tk.Tk()
    app = mod.BundleBuilderApp(root)
    yield mod, root, app
    try:
        root.destroy()
    except Exception:
        pass


def test_direct_build_with_preset_and_settings(gui, make_tree):
    mod, root, app = gui
    import fitz
    tree = make_tree({"Root A.pdf": 1, "Folder One/F1.pdf": 2, "Folder One/Sub/S1.pdf": 1})
    app.entry_dir.delete(0, "end")
    app.entry_dir.insert(0, tree)
    app.entry_claim_title.delete(0, "end")
    app.entry_claim_title.insert(0, "GUI Smoke")
    app.cmb_preset.set("Technical drawing register")
    app._apply_preset()
    assert app.var_term.get() == "Drawing" and app.var_style_toc.get() == "grid"
    assert not app.var_footer.get()
    app.var_sep_scope.set("none")
    app.var_cover.set(True)
    app._toggle_cover_fields()
    root.after(50, lambda: app._on_run(False))
    assert run_until(root, lambda: not app._building)
    out = os.path.join(tree, "GUI Smoke.pdf")
    assert os.path.exists(out)
    with fitz.open(out) as doc:
        assert doc.page_count == 6
        assert [t[1] for t in doc.get_toc()][2:] == ["Drawing 01: Root A", "Drawing 02: Folder One",
                                                     "F1", "Sub", "S1"]
        assert [l["page"] for l in doc[1].get_links()] == [2, 3, 3, 5, 5]
    with open(mod.be.SETTINGS_FILE) as f:
        saved = json.load(f)
    assert saved["section_term"] == "Drawing" and saved["separator_scope"] == "none"
    assert saved["style_preset"] == "drawing_register"


def test_organiser_flow(gui, make_tree):
    mod, root, app = gui
    import fitz
    import bundle_engine as be
    tree = make_tree({"Pack/Dwg 1.pdf": 3, "Pack/Dwg 2.pdf": {"pages": 1, "width": 1190, "height": 842},
                      "Note.pdf": 1})
    app.entry_dir.delete(0, "end")
    app.entry_dir.insert(0, tree)
    app.entry_claim_title.delete(0, "end")
    app.entry_claim_title.insert(0, "Organised")
    app.var_sep_scope.set("top_level")
    app.var_page_mode.set("all")
    root.after(50, lambda: app._on_run(True))
    assert run_until(root, lambda: app._organiser is not None and app._organiser.winfo_exists())
    org = app._organiser
    root.update()
    kinds = [e["kind"] for e in org._entries]
    # TOC placeholder + 2 separators + 5 source pages
    assert kinds.count("generated") == 3 and kinds.count("page") == 5
    assert org._entries[0]["label"] == "Table of Contents"
    # thumbnails arrive for visible tiles
    assert run_until(root, lambda: len(org._images) >= 5, timeout_ms=30000)
    # select Dwg 1 page 1 and page 3 (entries after placeholders)
    page_entries = [i for i, e in enumerate(org._entries) if e["kind"] == "page"]
    d1 = [i for i in page_entries if org._entries[i]["doc"].source_name == "Dwg 1.pdf"]
    org.select_entries([d1[0]])
    org._include_selected(False)
    org.select_entries([i for i in range(len(org._entries)) if org._entries[i]["kind"] == "page"
                        and org._entries[i]["doc"].source_name == "Dwg 1.pdf"][2:])
    org._move_selected(-5)
    org._rotate_selected(90)
    assert len(org._undo) == 3
    org._undo_last()                       # undo the rotation; selection is cleared
    assert len(org._undo) == 2
    doc1 = [d for d in org.plan.documents() if d.source_name == "Dwg 1.pdf"][0]
    assert [(r["src"], r["inc"], r["rot"]) for r in doc1.page_refs] == [(2, True, 0), (0, False, 0), (1, True, 0)]
    first = [i for i, e in enumerate(org._entries) if e["kind"] == "page" and e["doc"] is doc1][0]
    org.select_entries([first])
    org._rotate_selected(90)
    doc1 = [d for d in org.plan.documents() if d.source_name == "Dwg 1.pdf"][0]
    assert [(r["src"], r["inc"], r["rot"]) for r in doc1.page_refs] == [(2, True, 90), (0, False, 0), (1, True, 0)]
    # move the Note section first via the tree
    iid = [k for k, n in org._tree_nodes.items() if n.kind == "section" and n.title == "Note"][0]
    org.tree.selection_set(iid)
    org._move_node(-1)
    assert [s.title for s in org.plan.sections] == ["Note", "Pack"]
    org._on_create()
    assert run_until(root, lambda: not app._building)
    out = os.path.join(tree, "Organised.pdf")
    with fitz.open(out) as doc:
        # toc + 2 separators + 4 included pages
        assert doc.page_count == 1 + 2 + 4
        # toc(0) sep A(1) Note(2) sep B(3) Dwg1 p3 rotated(4) Dwg1 p2(5) Dwg2(6)
        assert doc.load_page(4).rotation == 90
        assert "MARK Dwg 1 page 3" in doc.load_page(4).get_text()
        assert [t[1] for t in doc.get_toc()][1:] == ["Appendix A: Note", "Appendix B: Pack", "Dwg 1", "Dwg 2"]
        assert [l["page"] for l in doc[0].get_links()] == [1, 3, 4, 6]


def test_per_folder_mode_from_gui(gui, make_tree):
    mod, root, app = gui
    import fitz
    tree = make_tree({"Villa 1/A.pdf": 2, "Villa 1/B.pdf": 2, "Villa 2/C.pdf": 3})
    app.entry_dir.delete(0, "end")
    app.entry_dir.insert(0, tree)
    app.var_output_mode.set("per_folder")
    app.var_page_mode.set("first")
    root.after(50, lambda: app._on_run(False))
    assert run_until(root, lambda: not app._building)
    v1 = os.path.join(tree, "Villa 1", "Villa 1#Combined_FirstPages.pdf")
    assert os.path.exists(v1)
    with fitz.open(v1) as doc:
        assert doc.page_count == 2
    assert not os.path.exists(os.path.join(tree, os.path.basename(tree) + "#Combined_FirstPages.pdf"))
