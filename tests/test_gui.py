"""Builds the real window (skipped when there is no display)."""

import pytest

tk = pytest.importorskip("tkinter")

from certpdf.config import Settings  # noqa: E402


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setattr(Settings, "path", classmethod(lambda cls: tmp_path / "settings.json"))
    try:
        from certpdf.gui import App
        window = App()
    except tk.TclError as exc:
        pytest.skip(f"no display: {exc}")
    yield window
    window.destroy()


def rows(app):
    return [app.tree.item(item, "values") for item in app.tree.get_children()]


def test_add_files_shows_found_tabs(app, workbook_factory, tmp_path):
    full = workbook_factory("dev10.xlsx")
    partial = workbook_factory("dev9.xlsx", tabs=["ECG Certificate", "NIBP.pass-fail"], hidden=())
    app.add_paths([full, partial, full])  # the duplicate is ignored
    app.update()

    assert [r[1] for r in rows(app)] == ["dev10.xlsx", "dev9.xlsx"]
    assert rows(app)[0][2:6] == ("✔", "✔", "✔", "✔")
    assert rows(app)[1][2:6] == ("✔", "✖", "✖", "✔")
    assert app.output_var.get() == str(tmp_path)

    app._sort()
    assert [r[1] for r in rows(app)] == ["dev9.xlsx", "dev10.xlsx"]

    app.sheet_vars["nibp_cert"].set("ECG Certificate")
    app._rematch()
    assert rows(app)[0][3] == "✔"


def test_move_and_remove(app, workbook_factory):
    paths = [workbook_factory(f"d{i}.xlsx") for i in range(3)]
    app.add_paths(paths)
    app.tree.selection_set(app.tree.get_children()[2])
    app._move(-1)
    assert [r[1] for r in rows(app)] == ["d0.xlsx", "d2.xlsx", "d1.xlsx"]
    app._remove_selected()
    assert [r[1] for r in rows(app)] == ["d0.xlsx", "d1.xlsx"]


def run_to_end(app):
    import time
    app._start()
    deadline = time.monotonic() + 30
    while app.worker is not None and time.monotonic() < deadline:
        app.update()
        time.sleep(0.02)
    assert app.worker is None


@pytest.fixture
def dialogs(monkeypatch):
    import certpdf.gui as gui
    shown = []
    for name in ("showinfo", "showwarning", "showerror", "askyesno"):
        monkeypatch.setattr(gui.messagebox, name,
                            lambda title, message, _n=name, **k: shown.append((_n, title, message)) or False)
    monkeypatch.setattr(gui, "open_path", lambda *a, **k: None)
    return shown


def test_excel_files_are_made_when_no_program_can_make_pdfs(app, workbook_factory, tmp_path, monkeypatch, dialogs):
    # Like a PC with only Excel 2007: Excel can't save PDFs and LibreOffice isn't installed.
    from contextlib import contextmanager

    import certpdf.gui as gui
    from certpdf.engines import EngineError

    @contextmanager
    def no_pdf_engine(choice, log):
        raise EngineError("Microsoft Excel: this version of Excel cannot save PDFs")
        yield

    monkeypatch.setattr(gui, "open_engine", no_pdf_engine)
    monkeypatch.setattr(gui, "available_engines", lambda: {"excel": True, "libreoffice": False})
    out = tmp_path / "out"
    app.add_paths([workbook_factory("dev1.xlsx"), workbook_factory("dev2.xlsx")])
    app.output_var.set(str(out))
    run_to_end(app)

    assert sorted(p.name for p in out.iterdir()) == ["Certificates.xlsx", "Pass-Fail Test Sheets.xlsx"]
    kind, title, message = dialogs[-1]
    assert kind == "askyesno" and title == "Files created"
    assert "PDF files were not made" in message and "LibreOffice" in message
    assert [r[-1] for r in rows(app)] == ["Done", "Done"]
    assert "2 files saved" in app.status_label.cget("text")


def test_tab_errors_in_the_pdf_step_are_reported(app, workbook_factory, tmp_path, monkeypatch, dialogs):
    from contextlib import contextmanager

    import certpdf.gui as gui
    from certpdf.workbook_info import read_sheet_names

    class FailingEngine:
        name = "Microsoft Excel"

        @contextmanager
        def open(self, path):
            class Workbook:
                sheet_names = read_sheet_names(path)

                def export_many(self, jobs):
                    return {sheet: "Exception occurred. (code 0x800A03EC)" for sheet, _ in jobs}
            yield Workbook()

    @contextmanager
    def fake_open_engine(choice, log):
        yield FailingEngine()

    monkeypatch.setattr(gui, "open_engine", fake_open_engine)
    monkeypatch.setattr(gui, "available_engines", lambda: {"excel": True, "libreoffice": False})
    out = tmp_path / "out"
    app.add_paths([workbook_factory("dev.xlsx")])
    app.output_var.set(str(out))
    run_to_end(app)

    assert sorted(p.name for p in out.iterdir()) == ["Certificates.xlsx", "Pass-Fail Test Sheets.xlsx"]
    assert rows(app)[0][-1] == "Failed – see the messages below"
    kind, title, message = dialogs[-1]
    assert kind == "showwarning" and "had problems: dev.xlsx" in message
    assert "0x800A03EC" in app.log_text.get("1.0", "end")


def test_only_excel_files(app, workbook_factory, tmp_path, dialogs):
    out = tmp_path / "out"
    app.make_pdf_var.set(False)
    app._update_engine_note()
    assert app.engine_note.cget("text") == ""
    app.add_paths([workbook_factory("dev.xlsx")])
    app.output_var.set(str(out))
    app.cert_name_var.set("Batch 7 certificates.pdf")  # a typed extension is dropped
    run_to_end(app)
    assert sorted(p.name for p in out.iterdir()) == ["Batch 7 certificates.xlsx", "Pass-Fail Test Sheets.xlsx"]
    assert dialogs[-1][0] == "showinfo"
