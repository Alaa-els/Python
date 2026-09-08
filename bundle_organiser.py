"""
Document Bundle Builder - pre-build page organiser
==================================================

A Toplevel window that shows every page the bundle will contain as a
thumbnail grid, in bundle order, and lets the user select, re-order,
rotate, exclude or restore pages before the bundle is created. It edits
the engine's page manifest (``BundleNode.page_refs``) only; nothing is
written until "Create Bundle" hands the organised plan back.

Scope: selection / order / rotation / exclusion. Not a PDF editor.

Design notes
- Pages stay inside their own document. "Move earlier / later" moves the
  selected pages within the document that owns them (clamped at its
  ends). Documents and sections are re-ordered with the structure tree
  on the left. Cross-document page interleaving is deliberately not
  supported: the TOC and bookmarks target the first retained page of each
  document, which only stays meaningful while pages keep their document.
- Generated pages (cover, contents, separator pages) are shown as grey
  labelled placeholders so they are never mistaken for source pages.
  Their final look is available through "Preview bundle".
- Thumbnails are rendered lazily on a worker thread for the visible part
  of the grid only, through a bounded LRU cache, and rendering stops as
  soon as the window closes. Thumbnails are raster previews; the bundle
  itself keeps the native page content.
- Drag-and-drop re-ordering is not implemented: with a virtualised canvas
  and multi-thousand-page drawing packs it was not robust enough to
  trust, and the button-based moves are exact and undoable.
"""

import os
import queue
import threading
import tempfile
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox

import bundle_engine as be

ZOOM_LEVELS = {"Small": 110, "Medium": 160, "Large": 240}
TILE_PAD = 12
LABEL_H = 34
GENERATED_BG = "#e9e9ef"
GENERATED_FG = "#5a5a6a"
EXCLUDED_BG = "#f6f6f6"
SELECT_OUTLINE = "#0045BF"
UNDO_LIMIT = 40


class OrganiserWindow(tk.Toplevel):
    """Organise pages of a BundlePlan. on_create(plan) is called when the
    user presses Create Bundle; on_preview(plan) when they ask for a full
    preview build (both receive the organised plan)."""

    def __init__(self, parent, plan, opts, on_create, on_preview=None, cache=None):
        super().__init__(parent)
        self.title("Organise pages - Document Bundle Builder")
        self.geometry("1240x820")
        self.minsize(900, 600)
        self.plan = plan
        self.original = be.snapshot_plan(plan)
        self.opts = opts
        self.on_create = on_create
        self.on_preview = on_preview
        self.cache = cache or be.ThumbnailCache(max_items=400)
        self.preview_cache = be.ThumbnailCache(max_items=6, max_open_docs=2)
        self._undo = []
        self._closing = False
        self._entries = []
        self._selected = set()
        self._anchor = None
        self._tiles = {}          # entry index -> dict of canvas item ids
        self._images = {}         # entry index -> PhotoImage (visible tiles only)
        self._pending = set()
        self._queue = queue.Queue()
        self._requests = queue.Queue()
        self._zoom = "Medium"
        self._build_ui()
        self._rebuild_entries()
        self._worker = threading.Thread(target=self._render_worker, daemon=True)
        self._worker.start()
        self.after(60, self._pump_results)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.transient(parent)

    # ------------------------------------------------------------ UI
    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(side="top", fill="x", padx=8, pady=6)
        tk.Label(top, text="Organise pages", font=("Arial", 12, "bold")).pack(side="left")
        tk.Label(top, text="  Select pages (click, Ctrl-click, Shift-click), then use the buttons. "
                           "Pages stay within their document; re-order documents and sections "
                           "with the tree on the left.",
                 fg="#555", font=("Arial", 9)).pack(side="left")

        bar = tk.Frame(self)
        bar.pack(side="top", fill="x", padx=8, pady=(0, 6))

        def btn(text, cmd, width=13):
            b = tk.Button(bar, text=text, command=cmd, width=width)
            b.pack(side="left", padx=2)
            return b

        btn("< Move earlier", lambda: self._move_selected(-1))
        btn("Move later >", lambda: self._move_selected(+1))
        btn("Rotate left", lambda: self._rotate_selected(-90), 11)
        btn("Rotate right", lambda: self._rotate_selected(+90), 11)
        btn("Exclude", lambda: self._include_selected(False), 9)
        btn("Restore", lambda: self._include_selected(True), 9)
        btn("Undo", self._undo_last, 7)
        btn("Reset", self._reset, 7)
        btn("Select all", self._select_all, 9)
        tk.Label(bar, text="  Zoom:").pack(side="left", padx=(12, 2))
        self.var_zoom = tk.StringVar(value=self._zoom)
        zoom = ttk.Combobox(bar, textvariable=self.var_zoom, values=list(ZOOM_LEVELS),
                            state="readonly", width=8)
        zoom.pack(side="left")
        zoom.bind("<<ComboboxSelected>>", lambda e: self._set_zoom(self.var_zoom.get()))
        tk.Button(bar, text="Preview page", command=self._preview_selected, width=12).pack(side="left", padx=(12, 2))
        if self.on_preview:
            tk.Button(bar, text="Preview bundle (PDF)", command=self._preview_bundle, width=18).pack(side="left", padx=2)

        bottom = tk.Frame(self)
        bottom.pack(side="bottom", fill="x", padx=8, pady=6)
        self.var_status = tk.StringVar(value="")
        tk.Label(bottom, textvariable=self.var_status, anchor="w", fg="#1F3864",
                 font=("Arial", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(bottom, text="Cancel", command=self._on_cancel, padx=14).pack(side="right", padx=4)
        tk.Button(bottom, text="Create Bundle", command=self._on_create, bg="#1F3864", fg="white",
                  font=("Arial", 10, "bold"), padx=18, pady=4).pack(side="right", padx=4)

        body = tk.PanedWindow(self, orient="horizontal", sashwidth=6)
        body.pack(side="top", fill="both", expand=True, padx=8, pady=(0, 4))

        left = tk.Frame(body, width=300)
        body.add(left, minsize=220)
        tk.Label(left, text="Structure (sections and documents)", font=("Arial", 9, "bold")).pack(anchor="w")
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        tb = tk.Frame(left)
        tb.pack(fill="x")
        tk.Button(tb, text="Move up", command=lambda: self._move_node(-1), width=10).pack(side="left", padx=2)
        tk.Button(tb, text="Move down", command=lambda: self._move_node(+1), width=10).pack(side="left", padx=2)
        tk.Button(tb, text="Exclude doc", command=lambda: self._include_node(False), width=10).pack(side="left", padx=2)
        tk.Button(tb, text="Restore doc", command=lambda: self._include_node(True), width=10).pack(side="left", padx=2)
        tk.Label(left, text="Generated pages (cover, contents, separators) are shown as grey "
                            "placeholders in the grid; they are not source pages.",
                 fg="#555", font=("Arial", 8), wraplength=280, justify="left").pack(anchor="w", pady=4)

        right = tk.Frame(body)
        body.add(right)
        self.canvas = tk.Canvas(right, bg="white", highlightthickness=0)
        vsb = tk.Scrollbar(right, orient="vertical", command=self._on_scrollbar)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._relayout())
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Control-Button-1>", lambda e: self._on_click(e, ctrl=True))
        self.canvas.bind("<Shift-Button-1>", lambda e: self._on_click(e, shift=True))
        self.canvas.bind("<Double-Button-1>", lambda e: self._preview_selected())
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Button-4>", lambda e: self._scroll_units(-3))
        self.canvas.bind("<Button-5>", lambda e: self._scroll_units(3))

    # ------------------------------------------------------------ entries
    def _rebuild_entries(self):
        """Flatten plan + options into the ordered list of grid entries."""
        self._entries = []
        self._tiles.clear()
        self._images.clear()
        self._pending.clear()
        self.canvas.delete("all")
        opts = self.opts
        term = be.CURRENT_TERMS.get("section_term", "Appendix")
        if opts.get("include_cover"):
            self._entries.append({"kind": "generated", "label": "Cover page", "sub": "generated"})
        if opts.get("generate_toc", True):
            self._entries.append({"kind": "generated", "label": "Table of Contents", "sub": "generated"})
        relabel = be.CURRENT_TERMS.get("numbering_style", "alphabetic")
        for i, sec in enumerate(self.plan.sections, start=1):
            sec.label = be.convert_to_style(i, relabel)
        for node in self.plan.walk():
            if node.kind == "section":
                if node.separator:
                    self._entries.append({"kind": "generated", "label": f"Separator: {node.display(term)}",
                                          "sub": "generated", "node": node})
            else:
                refs = node.page_refs if node.page_refs is not None else \
                    [{"src": k, "rot": 0, "inc": True} for k in range(node.expected_pages or 0)]
                for k, ref in enumerate(refs):
                    self._entries.append({"kind": "page", "doc": node, "ref_index": k, "ref": ref})
        self._selected = {i for i in self._selected if i < len(self._entries)
                          and self._entries[i]["kind"] == "page"}
        self._rebuild_tree()
        self._relayout()
        self._update_status()

    def _rebuild_tree(self):
        self.tree.delete(*self.tree.get_children())
        self._tree_nodes = {}

        def add(node, parent=""):
            if node.kind == "document":
                inc = len(node.included_pages())
                total = len(node.page_refs) if node.page_refs is not None else (node.expected_pages or 0)
                text = f"{node.source_name}  ({inc}/{total} pages)"
            else:
                text = node.display() if node.level == 1 else node.title
            iid = self.tree.insert(parent, "end", text=text, open=True)
            self._tree_nodes[iid] = node
            for c in node.children:
                add(c, iid)

        for sec in self.plan.sections:
            add(sec)

    # ------------------------------------------------------------ layout
    def _tile_size(self):
        px = ZOOM_LEVELS[self._zoom]
        return px, int(px * 1.25) + LABEL_H

    def _relayout(self):
        w = max(self.canvas.winfo_width(), 200)
        tw, th = self._tile_size()
        self._cols = max(1, (w - TILE_PAD) // (tw + TILE_PAD))
        rows = (len(self._entries) + self._cols - 1) // self._cols
        self.canvas.configure(scrollregion=(0, 0, w, rows * (th + TILE_PAD) + TILE_PAD))
        self.canvas.delete("all")
        self._tiles.clear()
        self._images.clear()
        self._draw_visible()

    def _entry_box(self, index):
        tw, th = self._tile_size()
        r, c = divmod(index, self._cols)
        x = TILE_PAD + c * (tw + TILE_PAD)
        y = TILE_PAD + r * (th + TILE_PAD)
        return x, y, tw, th

    def _visible_range(self):
        tw, th = self._tile_size()
        y0 = self.canvas.canvasy(0)
        y1 = y0 + self.canvas.winfo_height()
        first_row = max(0, int(y0 // (th + TILE_PAD)) - 1)
        last_row = int(y1 // (th + TILE_PAD)) + 2
        return first_row * self._cols, min(len(self._entries), (last_row + 1) * self._cols)

    def _draw_visible(self):
        lo, hi = self._visible_range()
        keep = set(range(lo, hi))
        for idx in list(self._tiles):
            if idx not in keep:
                for item in self._tiles[idx].values():
                    self.canvas.delete(item)
                del self._tiles[idx]
                self._images.pop(idx, None)
        for idx in range(lo, hi):
            if idx not in self._tiles:
                self._draw_tile(idx)
        self._request_visible(lo, hi)

    def _draw_tile(self, idx):
        e = self._entries[idx]
        x, y, tw, th = self._entry_box(idx)
        img_h = th - LABEL_H
        items = {}
        selected = idx in self._selected
        if e["kind"] == "generated":
            items["bg"] = self.canvas.create_rectangle(x, y, x + tw, y + img_h, fill=GENERATED_BG,
                                                       outline="#b8b8c8", dash=(4, 3))
            items["txt"] = self.canvas.create_text(x + tw / 2, y + img_h / 2, text=e["label"],
                                                   width=tw - 12, fill=GENERATED_FG,
                                                   font=("Arial", 9, "italic"), justify="center")
            items["lbl"] = self.canvas.create_text(x + tw / 2, y + img_h + 12, text="GENERATED PAGE",
                                                   fill=GENERATED_FG, font=("Arial", 8, "bold"))
            items["ord"] = self.canvas.create_text(x + tw / 2, y + img_h + 26, text=f"bundle order {idx + 1}",
                                                   fill="#888", font=("Arial", 8))
        else:
            ref = e["ref"]
            inc = ref.get("inc", True)
            items["bg"] = self.canvas.create_rectangle(
                x, y, x + tw, y + img_h, fill=EXCLUDED_BG if not inc else "white",
                outline=SELECT_OUTLINE if selected else "#c8c8c8", width=3 if selected else 1)
            items["txt"] = self.canvas.create_text(x + tw / 2, y + img_h / 2, text="rendering...",
                                                   fill="#aaa", font=("Arial", 8))
            name = e["doc"].source_name
            if len(name) > 30:
                name = name[:27] + "..."
            items["lbl"] = self.canvas.create_text(x + tw / 2, y + img_h + 10, text=name, width=tw,
                                                   fill="#222" if inc else "#999", font=("Arial", 8))
            extra = f"page {ref['src'] + 1}"
            if ref.get("rot"):
                extra += f"  rot {ref['rot']}"
            extra += "  EXCLUDED" if not inc else f"  order {idx + 1}"
            items["ord"] = self.canvas.create_text(x + tw / 2, y + img_h + 24, text=extra, width=tw,
                                                   fill="#a00" if not inc else "#666", font=("Arial", 8))
        self._tiles[idx] = items

    def _request_visible(self, lo, hi):
        for idx in range(lo, hi):
            e = self._entries[idx]
            if e["kind"] != "page" or idx in self._images or idx in self._pending:
                continue
            self._pending.add(idx)
            self._requests.put((idx, e["doc"].path, e["ref"]["src"], e["ref"].get("rot", 0),
                                ZOOM_LEVELS[self._zoom]))

    # ------------------------------------------------------------ rendering
    def _render_worker(self):
        while not self._closing:
            try:
                idx, path, src, rot, px = self._requests.get(timeout=0.2)
            except queue.Empty:
                continue
            if self._closing:
                break
            try:
                data = self.cache.get(path, src, rot, px)
                self._queue.put((idx, path, src, rot, px, data, None))
            except Exception as e:
                self._queue.put((idx, path, src, rot, px, None, str(e)))

    def _pump_results(self):
        if self._closing:
            return
        try:
            for _ in range(40):
                idx, path, src, rot, px, data, err = self._queue.get_nowait()
                self._pending.discard(idx)
                if idx >= len(self._entries) or idx not in self._tiles:
                    continue
                e = self._entries[idx]
                if (e["kind"] != "page" or e["doc"].path != path or e["ref"]["src"] != src
                        or e["ref"].get("rot", 0) != rot or px != ZOOM_LEVELS[self._zoom]):
                    continue
                x, y, tw, th = self._entry_box(idx)
                img_h = th - LABEL_H
                if data is None:
                    self.canvas.itemconfigure(self._tiles[idx]["txt"], text=f"cannot render\n{err[:40]}",
                                              fill="#a00")
                    continue
                ppm, w, h = data
                photo = tk.PhotoImage(data=ppm)
                self._images[idx] = photo
                self.canvas.itemconfigure(self._tiles[idx]["txt"], text="")
                item = self.canvas.create_image(x + tw / 2, y + img_h / 2, image=photo)
                self._tiles[idx]["img"] = item
                if not e["ref"].get("inc", True):
                    self._tiles[idx]["veil"] = self.canvas.create_rectangle(
                        x + 1, y + 1, x + tw - 1, y + img_h - 1, fill="white", stipple="gray50", outline="")
                self.canvas.tag_raise(self._tiles[idx]["bg"])
                self.canvas.itemconfigure(self._tiles[idx]["bg"], fill="")
        except queue.Empty:
            pass
        self.after(60, self._pump_results)

    # ------------------------------------------------------------ scrolling
    def _on_scrollbar(self, *args):
        self.canvas.yview(*args)
        self._draw_visible()

    def _scroll_units(self, n):
        self.canvas.yview_scroll(n, "units")
        self._draw_visible()

    def _on_wheel(self, event):
        self._scroll_units(int(-event.delta / 120) * 3)

    def _set_zoom(self, level):
        if level in ZOOM_LEVELS:
            self._zoom = level
            self._relayout()

    # ------------------------------------------------------------ selection
    def _index_at(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        tw, th = self._tile_size()
        c = int((x - TILE_PAD) // (tw + TILE_PAD))
        r = int((y - TILE_PAD) // (th + TILE_PAD))
        if c < 0 or c >= self._cols or r < 0:
            return None
        idx = r * self._cols + c
        if idx >= len(self._entries):
            return None
        bx, by, _, _ = self._entry_box(idx)
        if x > bx + tw or y > by + th:
            return None
        return idx

    def _on_click(self, event, ctrl=False, shift=False):
        idx = self._index_at(event)
        if idx is None or self._entries[idx]["kind"] != "page":
            if not ctrl and not shift:
                self._selected.clear()
                self._refresh_selection()
            return
        if shift and self._anchor is not None:
            lo, hi = sorted((self._anchor, idx))
            for i in range(lo, hi + 1):
                if self._entries[i]["kind"] == "page":
                    self._selected.add(i)
        elif ctrl:
            if idx in self._selected:
                self._selected.discard(idx)
            else:
                self._selected.add(idx)
            self._anchor = idx
        else:
            self._selected = {idx}
            self._anchor = idx
        self._refresh_selection()

    def _refresh_selection(self):
        for idx, items in self._tiles.items():
            if self._entries[idx]["kind"] != "page":
                continue
            sel = idx in self._selected
            self.canvas.itemconfigure(items["bg"], outline=SELECT_OUTLINE if sel else "#c8c8c8",
                                      width=3 if sel else 1)
        self._update_status()

    def _select_all(self):
        self._selected = {i for i, e in enumerate(self._entries) if e["kind"] == "page"}
        self._refresh_selection()

    def select_entries(self, indices):
        """Programmatic selection (used by tests)."""
        self._selected = {i for i in indices if 0 <= i < len(self._entries)
                          and self._entries[i]["kind"] == "page"}
        self._refresh_selection()

    def _selected_by_doc(self):
        by_doc = {}
        for idx in self._selected:
            e = self._entries[idx]
            by_doc.setdefault(id(e["doc"]), (e["doc"], []))[1].append(e["ref_index"])
        return list(by_doc.values())

    # ------------------------------------------------------------ operations
    def _push_undo(self):
        self._undo.append(be.snapshot_plan(self.plan))
        if len(self._undo) > UNDO_LIMIT:
            self._undo.pop(0)

    def _reselect_refs(self, refs):
        """After an operation, select the entries whose ref objects match."""
        wanted = {id(r) for r in refs}
        self._selected = {i for i, e in enumerate(self._entries)
                          if e["kind"] == "page" and id(e["ref"]) in wanted}

    def _move_selected(self, delta):
        groups = self._selected_by_doc()
        if not groups:
            return
        self._push_undo()
        moved_refs = []
        for doc, idxs in groups:
            new_idx = be.move_pages(doc, idxs, delta)
            moved_refs.extend(doc.page_refs[i] for i in new_idx)
        self._rebuild_entries()
        self._reselect_refs(moved_refs)
        self._refresh_selection()

    def _rotate_selected(self, degrees):
        groups = self._selected_by_doc()
        if not groups:
            return
        self._push_undo()
        refs = []
        for doc, idxs in groups:
            be.rotate_pages(doc, idxs, degrees)
            refs.extend(doc.page_refs[i] for i in idxs)
        self._rebuild_entries()
        self._reselect_refs(refs)
        self._refresh_selection()

    def _include_selected(self, included):
        groups = self._selected_by_doc()
        if not groups:
            return
        self._push_undo()
        refs = []
        for doc, idxs in groups:
            be.set_pages_included(doc, idxs, included)
            refs.extend(doc.page_refs[i] for i in idxs)
        self._rebuild_entries()
        self._reselect_refs(refs)
        self._refresh_selection()

    def _current_node(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self._tree_nodes.get(sel[0])

    def _move_node(self, delta):
        node = self._current_node()
        if node is None:
            return
        self._push_undo()
        be.move_node(self.plan, node, delta)
        self._rebuild_entries()
        for iid, n in self._tree_nodes.items():
            if n is node:
                self.tree.selection_set(iid)
                self.tree.see(iid)
                break

    def _include_node(self, included):
        node = self._current_node()
        if node is None:
            return
        self._push_undo()
        for doc in node.documents():
            if doc.page_refs is None:
                doc.page_refs = [{"src": k, "rot": 0, "inc": True} for k in range(doc.expected_pages or 0)]
            be.set_pages_included(doc, range(len(doc.page_refs)), included)
        self._rebuild_entries()

    def _on_tree_select(self, _event=None):
        node = self._current_node()
        if node is None:
            return
        docs = node.documents()
        if not docs:
            return
        first = docs[0]
        for idx, e in enumerate(self._entries):
            if e["kind"] == "page" and e["doc"] is first:
                _, y, _, th = self._entry_box(idx)
                total = self.canvas.bbox("all") or (0, 0, 0, 1)
                region = self.canvas.cget("scrollregion").split()
                height = float(region[3]) if len(region) == 4 else max(total[3], 1)
                self.canvas.yview_moveto(max(0.0, (y - TILE_PAD) / max(height, 1)))
                self._draw_visible()
                break

    def _undo_last(self):
        if not self._undo:
            return
        self.plan = self._undo.pop()
        self._selected.clear()
        self._rebuild_entries()

    def _reset(self):
        if not messagebox.askyesno("Reset", "Discard all organiser changes and return to the "
                                            "initial selection?", parent=self):
            return
        self._push_undo()
        self.plan = be.snapshot_plan(self.original)
        self._selected.clear()
        self._rebuild_entries()

    # ------------------------------------------------------------ previews
    def _preview_selected(self):
        if not self._selected:
            return
        idx = min(self._selected)
        e = self._entries[idx]
        try:
            ppm, w, h = self.preview_cache.get(e["doc"].path, e["ref"]["src"], e["ref"].get("rot", 0), 900)
        except Exception as ex:
            messagebox.showerror("Preview", f"Could not render page: {ex}", parent=self)
            return
        win = tk.Toplevel(self)
        win.title(f"{e['doc'].source_name} - page {e['ref']['src'] + 1}")
        photo = tk.PhotoImage(data=ppm)
        lbl = tk.Label(win, image=photo)
        lbl.image = photo
        lbl.pack(padx=6, pady=6)
        tk.Label(win, text="Preview only (raster). The bundle keeps the original page content.",
                 fg="#555", font=("Arial", 8)).pack(pady=(0, 6))
        win.transient(self)

    def _preview_bundle(self):
        if self.on_preview:
            self.on_preview(be.snapshot_plan(self.plan))

    # ------------------------------------------------------------ status / close
    def _update_status(self):
        t = be.manifest_totals(self.plan)
        self.var_status.set(
            f"{t['sections']} section(s), {t['documents_with_pages']} of {t['documents']} document(s) "
            f"with pages, {t['pages_included']} page(s) included, {t['pages_excluded']} excluded, "
            f"{t['pages_rotated']} rotated  |  {len(self._selected)} selected  |  "
            f"undo steps: {len(self._undo)}")

    def _on_create(self):
        t = be.manifest_totals(self.plan)
        if t["pages_included"] == 0:
            messagebox.showwarning("Nothing to bundle", "Every page is excluded.", parent=self)
            return
        plan = self.plan
        self._close()
        self.on_create(plan)

    def _on_cancel(self):
        self._close()

    def _close(self):
        self._closing = True
        try:
            self.cache.clear()
            self.preview_cache.clear()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


def preview_output_path():
    return os.path.join(tempfile.gettempdir(), f"bundle_preview_{datetime.now():%Y%m%d_%H%M%S}.pdf")
