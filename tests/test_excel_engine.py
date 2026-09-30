"""The Excel engine, tested against a stand-in for Excel's COM objects."""

from pathlib import Path

import pytest

from certpdf.engines import EngineUnusable, ExcelEngine, _com_message

EXCEL_ERROR = -2146827284  # 0x800A03EC, Excel's generic "it didn't work"


def com_error():
    # Same shape as pywintypes.com_error: (hresult, text, excepinfo, argerror)
    return Exception(-2147352567, "Exception occurred.", (0, None, None, None, 0, EXCEL_ERROR), None)


class FakeRange:
    Value = None


class FakeSheet:
    Visible = -1

    def __init__(self, name, export_ok=True):
        self.Name, self.export_ok = name, export_ok
        self.exports = []

    def Range(self, address):
        return FakeRange()

    def ExportAsFixedFormat(self, kind, path, quality, props, ignore_print_areas):
        self.exports.append(path)
        if not self.export_ok:
            raise com_error()
        Path(path).write_bytes(b"%PDF-1.4 fake")

    def PrintOut(self, *args, **kwargs):
        raise AssertionError("the app must never print")

    def Copy(self, *args, **kwargs):
        raise AssertionError("not needed")


class FakeSheets(list):
    def __call__(self, key):
        if isinstance(key, int):
            return self[key - 1]
        return next(sheet for sheet in self if sheet.Name == key)


class FakeWorkbook:
    def __init__(self, sheets):
        self.Worksheets = FakeSheets(sheets)
        self.closed = False

    def Close(self, SaveChanges):
        self.closed = True


class FakeBooks:
    def __init__(self, export_ok=True):
        self.export_ok = export_ok
        self.opened, self.added = [], []

    def Add(self):
        book = FakeWorkbook([FakeSheet("Sheet1", self.export_ok)])
        self.added.append(book)
        return book

    def Open(self, path, **kwargs):
        book = FakeWorkbook([FakeSheet("ECG.certificate", self.export_ok),
                             FakeSheet("NIBP.certificate", self.export_ok)])
        self.opened.append((path, kwargs, book))
        return book


class FakeApp:
    def __init__(self, export_ok=True):
        self.Workbooks = FakeBooks(export_ok)


def make_engine(tmp_path, export_ok=True):
    engine = ExcelEngine()
    engine._app = FakeApp(export_ok)
    engine._tmp = tmp_path / "excel-tmp"
    engine._tmp.mkdir()
    return engine


def test_com_message_shows_the_error_code():
    assert _com_message(com_error()) == "Exception occurred. (code 0x800A03EC)"
    described = Exception(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Document not saved.",
                                                               None, 0, EXCEL_ERROR), None)
    assert _com_message(described) == "Document not saved. (code 0x800A03EC)"
    assert _com_message(ValueError("plain")) == "plain"


def test_pdf_check_passes_when_excel_can_export(tmp_path):
    engine = make_engine(tmp_path)
    engine.check_pdf_export()
    assert engine._app.Workbooks.added[0].closed


def test_pdf_check_detects_excel_that_cannot_export(tmp_path):
    # Like Excel 2007 without the "Save as PDF" add-in.
    engine = make_engine(tmp_path, export_ok=False)
    with pytest.raises(EngineUnusable, match="cannot save PDFs.*0x800A03EC"):
        engine.check_pdf_export()
    assert engine._app.Workbooks.added[0].closed


def test_export_uses_a_local_copy_and_never_prints(tmp_path):
    engine = make_engine(tmp_path)
    original = tmp_path / "share" / "F23-AGH046-0626.xlsx"
    original.parent.mkdir()
    original.write_bytes(b"workbook bytes")

    with engine.open(original) as workbook:
        opened_path, options, book = engine._app.Workbooks.opened[0]
        assert Path(opened_path) != original
        assert Path(opened_path).read_bytes() == b"workbook bytes"
        assert workbook.sheet_names == ["ECG.certificate", "NIBP.certificate"]
        errors = workbook.export_many([("ECG.certificate", tmp_path / "ecg.pdf"),
                                       ("NIBP.certificate", tmp_path / "nibp.pdf")])
    assert errors == {}
    assert (tmp_path / "ecg.pdf").exists() and (tmp_path / "nibp.pdf").exists()
    assert book.closed
    assert not Path(opened_path).exists()  # the copy is cleaned up
    assert original.read_bytes() == b"workbook bytes"


def test_failed_export_is_reported_per_tab(tmp_path):
    engine = make_engine(tmp_path, export_ok=False)
    source = tmp_path / "dev.xlsx"
    source.write_bytes(b"x")
    with engine.open(source) as workbook:
        errors = workbook.export_many([("ECG.certificate", tmp_path / "ecg.pdf")])
    assert errors == {"ECG.certificate": "Exception occurred. (code 0x800A03EC)"}


def test_unreadable_file(tmp_path):
    engine = make_engine(tmp_path)
    with pytest.raises(Exception, match="could not be read"):
        with engine.open(tmp_path / "missing.xlsx"):
            pass
