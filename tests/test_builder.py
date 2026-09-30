import subprocess
from contextlib import contextmanager
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter

from certpdf.builder import Cancelled, build_pdfs
from certpdf.config import CERTIFICATES, PASS_FAIL, default_sheet_names
from certpdf.engines import EngineError, LibreOfficeEngine
from certpdf.workbook_info import read_sheet_names


class Recorder:
    def __init__(self):
        self.logs, self.statuses = [], {}

    def log(self, level, text):
        self.logs.append((level, text))

    def progress(self, done, total, text):
        assert 0 <= done <= total

    def file_status(self, index, state, text):
        self.statuses[index] = state


class FakeEngine:
    """Prints each tab as one blank page whose width identifies it."""
    name = "Fake"

    def __init__(self, fail_open=()):
        self.widths = {}
        self.fail_open = set(fail_open)

    @contextmanager
    def open(self, path):
        if Path(path).name in self.fail_open:
            raise EngineError("cannot open")
        engine, names = self, read_sheet_names(path)

        class Workbook:
            sheet_names = names

            def export_many(self, jobs):
                for sheet, pdf in jobs:
                    width = 200 + 10 * len(engine.widths)
                    engine.widths[(Path(path).stem, sheet)] = width
                    writer = PdfWriter()
                    writer.add_blank_page(width, 300)
                    writer.write(str(pdf))
                return {}

        yield Workbook()


def widths(pdf):
    return [float(page.mediabox.width) for page in PdfReader(str(pdf)).pages]


def outline_titles(pdf):
    def walk(items):
        for item in items:
            if isinstance(item, list):
                yield list(walk(item))
            else:
                yield item.title
    return list(walk(PdfReader(str(pdf)).outline))


def test_pages_are_grouped_per_device_in_order(workbook_factory, tmp_path):
    files = [workbook_factory("dev2.xlsx"), workbook_factory("dev1.xlsx")]
    engine = FakeEngine()
    out = {CERTIFICATES: tmp_path / "out" / "cert.pdf", PASS_FAIL: tmp_path / "out" / "pf.pdf"}

    result = build_pdfs(files, default_sheet_names(), out, engine, Recorder())

    w = engine.widths
    assert widths(out[CERTIFICATES]) == [w[("dev2", "ECG.certificate")], w[("dev2", "NIBP.certificate")],
                                         w[("dev1", "ECG.certificate")], w[("dev1", "NIBP.certificate")]]
    assert widths(out[PASS_FAIL]) == [w[("dev2", "ECG.pass-fail")], w[("dev2", "NIBP.pass-fail")],
                                      w[("dev1", "ECG.pass-fail")], w[("dev1", "NIBP.pass-fail")]]
    assert outline_titles(out[CERTIFICATES]) == [
        "dev2", ["ECG certificate", "NIBP certificate"], "dev1", ["ECG certificate", "NIBP certificate"]]
    assert result.outputs[CERTIFICATES].pages == 4
    assert result.problem_count == 0


def test_missing_tabs_and_unreadable_files_are_reported(workbook_factory, tmp_path):
    files = [
        workbook_factory("full.xlsx"),
        workbook_factory("no_nibp.xlsx", tabs=["ECG Certificate", "ECG Pass Fail"], hidden=()),
        workbook_factory("broken.xlsx"),
    ]
    out = {CERTIFICATES: tmp_path / "cert.pdf", PASS_FAIL: tmp_path / "pf.pdf"}
    recorder = Recorder()

    result = build_pdfs(files, default_sheet_names(), out, FakeEngine(fail_open={"broken.xlsx"}), recorder)

    assert len(widths(out[CERTIFICATES])) == 3   # full: 2, no_nibp: 1 (loosely matched name)
    assert result.workbooks[1].missing == ["NIBP certificate", "NIBP pass/fail"]
    assert result.workbooks[2].problems == ["cannot open"]
    assert recorder.statuses == {0: "done", 1: "warning", 2: "error"}
    assert any("NIBP.certificate" in text for level, text in recorder.logs if level == "warning")


def test_group_without_any_tab_is_not_written(workbook_factory, tmp_path):
    files = [workbook_factory("dev.xlsx", tabs=["ECG.certificate"], hidden=())]
    out = {CERTIFICATES: tmp_path / "cert.pdf", PASS_FAIL: tmp_path / "pf.pdf"}
    result = build_pdfs(files, default_sheet_names(), out, FakeEngine(), Recorder())
    assert result.outputs[PASS_FAIL] is None and not out[PASS_FAIL].exists()
    assert result.outputs[CERTIFICATES].tabs == 1


def test_custom_sheet_names(workbook_factory, tmp_path):
    files = [workbook_factory("dev.xlsx", tabs=["Cert A", "Cert B"], hidden=())]
    names = dict(default_sheet_names(), ecg_cert="Cert B", nibp_cert="Cert A")
    engine = FakeEngine()
    out = {CERTIFICATES: tmp_path / "cert.pdf"}
    build_pdfs(files, names, out, engine, Recorder())
    assert widths(out[CERTIFICATES]) == [engine.widths[("dev", "Cert B")], engine.widths[("dev", "Cert A")]]


def test_cancel(workbook_factory, tmp_path):
    import threading
    cancel = threading.Event()
    cancel.set()
    with pytest.raises(Cancelled):
        build_pdfs([workbook_factory("dev.xlsx")], default_sheet_names(),
                   {CERTIFICATES: tmp_path / "c.pdf"}, FakeEngine(), Recorder(), cancel)
    assert not (tmp_path / "c.pdf").exists()


# --------------------------------------------------------------------------
# End to end with a real LibreOffice, when one is installed
# --------------------------------------------------------------------------

soffice = LibreOfficeEngine.find()
needs_libreoffice = pytest.mark.skipif(soffice is None, reason="LibreOffice is not installed")


def page_texts(pdf):
    return [" ".join(page.extract_text().split()) for page in PdfReader(str(pdf)).pages]


@needs_libreoffice
def test_libreoffice_end_to_end(workbook_factory, tmp_path):
    dev_a = workbook_factory("Monitor A.xlsx", "MonA")
    dev_b = workbook_factory("monitor-b.xlsx", "MonB")
    # An old-style .xls made by converting an .xlsx.
    xlsx_c = workbook_factory("dev_c.xlsx", "MonC")
    xls_dir = tmp_path / "xls"
    subprocess.run([soffice, f"-env:UserInstallation={(tmp_path / 'prof').as_uri()}", "--headless",
                    "--convert-to", "xls", "--outdir", str(xls_dir), str(xlsx_c)],
                   check=True, capture_output=True, timeout=180)
    dev_c = xls_dir / "dev_c.xls"
    assert dev_c.exists()

    out = {CERTIFICATES: tmp_path / "Certificates.pdf", PASS_FAIL: tmp_path / "Pass-Fail.pdf"}
    engine = LibreOfficeEngine()
    engine.start()
    try:
        result = build_pdfs([dev_a, dev_b, dev_c], default_sheet_names(), out, engine, Recorder())
    finally:
        engine.stop()

    assert result.problem_count == 0
    texts = page_texts(out[CERTIFICATES])
    assert [t.split(" Check")[0] for t in texts] == [
        "MonA ECG.certificate", "MonA NIBP.certificate",
        "MonB ECG.certificate", "MonB NIBP.certificate",
        "MonC ECG.certificate", "MonC NIBP.certificate",
    ]
    texts = page_texts(out[PASS_FAIL])
    assert [t.split(" Check")[0] for t in texts] == [
        "MonA ECG.pass-fail", "MonA NIBP.pass-fail",
        "MonB ECG.pass-fail", "MonB NIBP.pass-fail",
        "MonC ECG.pass-fail", "MonC NIBP.pass-fail",
    ]
    assert engine._tmp is None  # temporary files were removed
