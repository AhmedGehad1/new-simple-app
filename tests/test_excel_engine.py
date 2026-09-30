"""The Excel engine's fallbacks, tested against a stand-in for Excel's COM objects."""

from pathlib import Path

import pytest

from certpdf.engines import EngineError, ExcelEngine, _com_message

EXCEL_ERROR = -2146827284  # 0x800A03EC, Excel's generic "it didn't work"


def com_error():
    # Same shape as pywintypes.com_error: (hresult, text, excepinfo, argerror)
    return Exception(-2147352567, "Exception occurred.", (0, None, None, None, 0, EXCEL_ERROR), None)


def write_pdf(path):
    Path(path).write_bytes(b"%PDF-1.4 fake")


class FakeWorkbook:
    def __init__(self, app, sheets=(), export_ok=True):
        self.app, self.export_ok, self.closed = app, export_ok, False
        self.Worksheets = FakeSheets(sheets)

    def ExportAsFixedFormat(self, kind, path, quality, props, ignore_print_areas):
        if not self.export_ok:
            raise com_error()
        write_pdf(path)

    def Close(self, SaveChanges):
        self.closed = True
        self.app.books.remove(self)


class FakeSheets(list):
    def __call__(self, name):
        return next(sheet for sheet in self if sheet.Name == name)


class FakeBooks:
    def __init__(self, app):
        self.app, self.opened = app, []

    @property
    def Count(self):
        return len(self.app.books)

    def __call__(self, index):
        return self.app.books[index - 1]

    def Open(self, path, **kwargs):
        self.opened.append((path, kwargs))
        book = FakeWorkbook(self.app, [FakeSheet(self.app, "ECG.certificate")])
        self.app.books.append(book)
        return book


class FakeApp:
    def __init__(self):
        self.books = []
        self.Workbooks = FakeBooks(self)


class FakeSheet:
    Visible = -1

    def __init__(self, app, name, export=True, copy=True, printer=True, ignore_area=True):
        self.app, self.Name = app, name
        self.ok = {"export": export, "copy": copy, "print": printer, "ignore": ignore_area}
        self.calls = []

    def ExportAsFixedFormat(self, kind, path, quality, props, ignore_print_areas):
        method = "ignore" if ignore_print_areas else "export"
        self.calls.append(method)
        if not self.ok[method]:
            raise com_error()
        write_pdf(path)

    def Copy(self):
        self.calls.append("copy")
        self.app.books.append(FakeWorkbook(self.app, export_ok=self.ok["copy"]))

    def PrintOut(self, ActivePrinter, PrintToFile, PrToFileName):
        self.calls.append("print")
        if not self.ok["print"]:
            raise com_error()
        write_pdf(PrToFileName)


@pytest.fixture
def engine(tmp_path):
    excel = ExcelEngine()
    excel._app = FakeApp()
    excel._tmp = tmp_path / "excel-tmp"
    excel._tmp.mkdir()
    excel.logs = []
    excel.log = lambda level, text: excel.logs.append((level, text))
    return excel


def test_com_message_shows_the_error_code():
    assert _com_message(com_error()) == "Exception occurred. (code 0x800A03EC)"
    described = Exception(-2147352567, "Exception occurred.", (0, "Microsoft Excel", "Document not saved.",
                                                               None, 0, EXCEL_ERROR), None)
    assert _com_message(described) == "Document not saved. (code 0x800A03EC)"
    assert _com_message(ValueError("plain")) == "plain"


def test_normal_export(engine, tmp_path):
    sheet = FakeSheet(engine._app, "ECG")
    engine.convert(sheet, tmp_path / "a.pdf")
    assert (tmp_path / "a.pdf").exists()
    assert sheet.calls == ["export"]
    assert engine.logs == []


def test_falls_back_to_the_pdf_printer_and_remembers_it(engine, tmp_path):
    first = FakeSheet(engine._app, "ECG", export=False, copy=False)
    engine.convert(first, tmp_path / "a.pdf")
    assert first.calls == ["export", "copy", "print"]
    assert (tmp_path / "a.pdf").exists()
    assert engine._app.books == []  # the temporary copy was closed
    assert engine.logs and "Microsoft Print to PDF" in engine.logs[0][1]

    second = FakeSheet(engine._app, "NIBP", export=False, copy=False)
    engine.convert(second, tmp_path / "b.pdf")
    assert second.calls == ["print"]  # goes straight to what worked


def test_export_from_copy(engine, tmp_path):
    sheet = FakeSheet(engine._app, "ECG", export=False)
    engine.convert(sheet, tmp_path / "a.pdf")
    assert sheet.calls == ["export", "copy"]
    assert (tmp_path / "a.pdf").exists()


def test_every_method_failing_is_reported(engine, tmp_path):
    sheet = FakeSheet(engine._app, "ECG", export=False, copy=False, printer=False, ignore_area=False)
    with pytest.raises(EngineError) as info:
        engine.convert(sheet, tmp_path / "a.pdf")
    message = str(info.value)
    for label, _ in engine.methods:
        assert label in message
    assert "0x800A03EC" in message
    assert not (tmp_path / "a.pdf").exists()


def test_open_uses_a_local_copy(engine, tmp_path):
    original = tmp_path / "share" / "F23-AGH046-0626.xlsx"
    original.parent.mkdir()
    original.write_bytes(b"workbook bytes")

    with engine.open(original) as workbook:
        opened_path, options = engine._app.Workbooks.opened[0]
        assert Path(opened_path) != original
        assert Path(opened_path).read_bytes() == b"workbook bytes"
        assert workbook.sheet_names == ["ECG.certificate"]
        errors = workbook.export_many([("ECG.certificate", tmp_path / "out.pdf")])
    assert errors == {}
    assert (tmp_path / "out.pdf").exists()
    assert not Path(opened_path).exists()  # the copy is cleaned up
    assert original.read_bytes() == b"workbook bytes"


def test_open_reports_unreadable_file(engine, tmp_path):
    with pytest.raises(EngineError, match="could not be read"):
        with engine.open(tmp_path / "missing.xlsx"):
            pass
