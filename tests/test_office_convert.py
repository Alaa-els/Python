"""Office conversion stage with fake Word / Excel backends (no Office needed)."""
import os
import json
import threading

import fitz

import bundle_engine as be
import office_convert as oc
from conftest import make_pdf, marker_of


class FakeWord:
    def __init__(self, fail=(), log=None):
        self.fail = set(fail)
        self.converted = []
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.closed = True
        return False

    def convert(self, source, target):
        name = os.path.basename(source)
        if name in self.fail:
            return False, "password-protected document"
        make_pdf(target, os.path.splitext(name)[0], pages=2)
        self.converted.append(source)
        return True, None


class FakeExcel:
    """Mimics ExcelCompanion: writes PDFs into output_dir and a result file."""

    def __init__(self, status="done", fail=(), drop=()):
        self.status = status
        self.fail = set(fail)
        self.drop = set(drop)
        self.calls = []

    def convert_many(self, docs, output_dir, cancel_event=None, title=None):
        self.calls.append([d.path for d in docs])
        results = []
        for d in docs:
            name = os.path.basename(d.path)
            if name in self.drop:
                continue
            if name in self.fail:
                results.append({"path": d.path, "ok": False, "error": "no tabs marked for printing"})
                continue
            pdf = os.path.join(output_dir, os.path.splitext(name)[0] + ".pdf")
            make_pdf(pdf, os.path.splitext(name)[0], pages=1)
            results.append({"path": d.path, "ok": True, "pdf_path": pdf, "warning": None})
        result_path = os.path.join(output_dir, oc.COMPANION_RESULT_NAME)
        with open(result_path, "w") as f:
            json.dump({"protocol": 1, "status": self.status, "error": None, "results": results}, f)
        # Reuse the real result-mapping logic
        real = oc.ExcelCompanion.__new__(oc.ExcelCompanion)
        real.log = lambda s: None
        data = oc.read_companion_result(result_path)
        by_path = {be.norm_path(r["path"]): r for r in data["results"]}
        for d in docs:
            r = by_path.get(be.norm_path(d.path))
            if self.status == "cancelled":
                d.status, d.error = "cancelled", "closed before exporting"
            elif r is None:
                d.status, d.error = "failed", "removed from list"
            elif r.get("ok"):
                os.makedirs(os.path.dirname(d.target), exist_ok=True)
                os.replace(r["pdf_path"], d.target)
                d.status = "converted"
            else:
                d.status, d.error = "failed", r.get("error")
        return data


def office_tree(make_tree):
    root = make_tree({"Corr/01 letter.pdf": 1, "Deep/Inner/x.pdf": 1})
    for rel, head in (("Corr/02 memo.docx", b"PK\x03\x04"), ("Corr/03 costs.xlsx", b"PK\x03\x04"),
                      ("Deep/Inner/04 old.doc", b"\xD0\xCF\x11\xE0"), ("Corr/~$02 memo.docx", b"lock"),
                      ("Corr/05 locked.docx", b"\xD0\xCF\x11\xE0")):
        full = os.path.join(root, rel)
        with open(full, "wb") as f:
            f.write(head)
    return root


def test_discovery_and_staging_paths(make_tree):
    root = office_tree(make_tree)
    docs = oc.discover_office_documents(root)
    names = [d.name for d in docs]
    assert names == ["02 memo.docx", "03 costs.xlsx", "05 locked.docx", "04 old.doc"]
    memo = docs[0]
    assert memo.kind == "word" and docs[1].kind == "excel"
    assert memo.target == os.path.join(root, be.STAGING_DIRNAME, "Corr", "02 memo.pdf")
    assert oc.looks_encrypted(os.path.join(root, "Corr", "05 locked.docx")) is True
    assert oc.looks_encrypted(os.path.join(root, "Corr", "02 memo.docx")) is False
    assert oc.looks_encrypted(os.path.join(root, "Deep", "Inner", "04 old.doc")) is None
    planned = oc.planned_overlay(root)
    assert len(planned) == 4
    # staging folder never counts as bundle content
    os.makedirs(os.path.join(root, be.STAGING_DIRNAME), exist_ok=True)
    make_pdf(os.path.join(root, be.STAGING_DIRNAME, "stray.pdf"), "stray")
    assert be.count_folder_tree(root)["total_pdfs"] == 2


def test_conversion_stage_with_fakes_and_bundle(make_tree):
    root = office_tree(make_tree)
    log = []
    word = FakeWord(fail={"05 locked.docx"})
    excel = FakeExcel()
    outcome = oc.run_conversion_stage(root, {"claim_title": "T"}, log.append, None,
                                      word_backend=word, excel_backend=excel)
    status = {d.name: d.status for d in outcome.documents}
    assert status == {"02 memo.docx": "converted", "03 costs.xlsx": "converted",
                      "05 locked.docx": "failed", "04 old.doc": "converted"}
    assert word.closed and excel.calls == [[os.path.join(root, "Corr", "03 costs.xlsx")]]
    assert [d.name for d in outcome.failures] == ["05 locked.docx"]
    assert not outcome.cancelled
    overlay = outcome.overlay
    assert len(overlay) == 3
    assert os.path.isfile(os.path.join(root, be.STAGING_DIRNAME, "Corr", "03 costs.pdf"))
    report = oc.write_conversion_report(root, outcome)
    assert report and os.path.isfile(report)

    opts = {"base_directory": root, "claim_title": "T", "separator_scope": "none",
            "policy": {"mode": "automatic", "layout": "preserve"}, "overlay": overlay}
    out, plan, _ = be.build_bundle_sync(opts, lambda s: None)
    corr = plan.sections[0]
    assert [d.title for d in corr.children] == ["01 letter", "02 memo", "03 costs"]
    with fitz.open(out) as doc:
        assert marker_of(doc, 2) == "02 memo page 1"
        assert marker_of(doc, 4) == "03 costs page 1"

    # second run reuses up-to-date conversions
    # failed documents are retried; converted ones are reused
    word2, excel2 = FakeWord(fail={"05 locked.docx"}), FakeExcel()
    outcome2 = oc.run_conversion_stage(root, {}, log.append, None, word_backend=word2,
                                       excel_backend=excel2)
    assert {d.status for d in outcome2.documents if d.name != "05 locked.docx"} == {"reused"}
    assert word2.converted == [] and excel2.calls == []
    outcome3 = oc.run_conversion_stage(root, {"office_reconvert": True}, log.append, None,
                                       word_backend=FakeWord(), excel_backend=FakeExcel())
    assert sum(1 for d in outcome3.documents if d.status == "converted") == 4


def test_excel_cancelled_and_dropped_results(make_tree):
    root = office_tree(make_tree)
    outcome = oc.run_conversion_stage(root, {}, None, None, word_backend=FakeWord(),
                                      excel_backend=FakeExcel(status="cancelled"))
    xl = [d for d in outcome.documents if d.kind == "excel"][0]
    assert xl.status == "cancelled"
    outcome = oc.run_conversion_stage(root, {"office_reconvert": True}, None, None,
                                      word_backend=FakeWord(),
                                      excel_backend=FakeExcel(drop={"03 costs.xlsx"}))
    xl = [d for d in outcome.documents if d.kind == "excel"][0]
    assert xl.status == "failed" and "removed" in xl.error


def test_word_cancel_event(make_tree):
    root = office_tree(make_tree)
    ev = threading.Event()
    ev.set()
    outcome = oc.run_conversion_stage(root, {}, None, ev, word_backend=FakeWord(),
                                      excel_backend=FakeExcel())
    assert outcome.cancelled
    assert all(d.status == "cancelled" for d in outcome.documents)


def test_unavailable_office_reports_cleanly(make_tree, monkeypatch):
    root = office_tree(make_tree)
    monkeypatch.setattr(oc, "office_support_status", lambda: (False, "no Office here"))
    outcome = oc.run_conversion_stage(root, {}, None, None)
    assert all(d.status == "failed" and d.error == "no Office here" for d in outcome.documents)
    assert outcome.overlay == {}
    assert outcome.messages == ["no Office here"]


def test_companion_command_and_result_contract(tmp_path):
    cmd = oc.companion_command("C:/tools/Excel_Quick_Print.pyw", ["a.xlsx", "b.xlsm"], "out",
                               "out/r.json", python_exe="pythonw.exe", title="Job")
    assert cmd[:2] == ["pythonw.exe", "C:/tools/Excel_Quick_Print.pyw"]
    assert "--companion-result" in cmd and "--no-workbook-copy" in cmd and cmd[-3:] == ["--", "a.xlsx", "b.xlsm"]
    p = tmp_path / "r.json"
    p.write_text(json.dumps({"status": "done", "results": [{"path": "a.xlsx", "ok": True}]}))
    data = oc.read_companion_result(str(p))
    assert data["status"] == "done" and data["error"] is None
    p.write_text("[]")
    try:
        oc.read_companion_result(str(p))
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_real_companion_runner_with_stub_tool(make_tree, tmp_path):
    """Run ExcelCompanion against a tiny stand-in script that honours the
    same CLI contract, so the subprocess / result-file path is exercised."""
    root = office_tree(make_tree)
    stub = tmp_path / "stub_tool.py"
    stub.write_text(
        "import sys, json, os\n"
        "argv = sys.argv[1:]\n"
        "res = argv[argv.index('--companion-result') + 1]\n"
        "out = argv[argv.index('--output-dir') + 1]\n"
        "files = argv[argv.index('--') + 1:]\n"
        "import fitz\n"
        "results = []\n"
        "for f in files:\n"
        "    pdf = os.path.join(out, os.path.splitext(os.path.basename(f))[0] + '.pdf')\n"
        "    d = fitz.open(); d.new_page().insert_text((72, 72), 'MARK ' + os.path.basename(f)); d.save(pdf); d.close()\n"
        "    results.append({'path': f, 'ok': True, 'pdf_path': pdf})\n"
        "json.dump({'protocol': 1, 'status': 'done', 'results': results}, open(res, 'w'))\n")
    docs = [d for d in oc.discover_office_documents(root) if d.kind == "excel"]
    import sys
    comp = oc.ExcelCompanion(str(stub), python_exe=sys.executable, poll_interval=0.05)
    result = comp.convert_many(docs, oc.staging_root(root))
    assert result["status"] == "done"
    assert docs[0].status == "converted" and os.path.isfile(docs[0].target)
    # a tool that exits without writing a result is reported, not hung
    bad = tmp_path / "bad_tool.py"
    bad.write_text("import sys; sys.exit(3)\n")
    docs = [d for d in oc.discover_office_documents(root) if d.kind == "excel"]
    comp = oc.ExcelCompanion(str(bad), python_exe=sys.executable, poll_interval=0.05)
    assert comp.convert_many(docs, oc.staging_root(root)) is None
    assert docs[0].status == "failed" and "exit code 3" in docs[0].error
