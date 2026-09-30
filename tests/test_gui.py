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


def test_conversion_failure_is_explained(app, workbook_factory, tmp_path, monkeypatch):
    import time
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
                    return {sheet: "Excel could not convert it. export: Exception occurred. (code 0x800A03EC)"
                            for sheet, _ in jobs}
            yield Workbook()

    @contextmanager
    def fake_open_engine(choice, log):
        yield FailingEngine()

    dialogs = []
    monkeypatch.setattr(gui, "open_engine", fake_open_engine)
    monkeypatch.setattr(gui, "available_engines", lambda: {"excel": True, "libreoffice": False})
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(gui.messagebox, "showerror", lambda title, message, **k: dialogs.append((title, message)))

    app.add_paths([workbook_factory("dev.xlsx")])
    app.output_var.set(str(tmp_path / "out"))
    app._start()
    deadline = time.monotonic() + 20
    while app.worker is not None and time.monotonic() < deadline:
        app.update()
        time.sleep(0.02)

    assert app.worker is None
    assert "could not be converted" in app.status_label.cget("text")
    assert dialogs and dialogs[0][0] == "The tabs could not be converted"
    assert "0x800A03EC" in dialogs[0][1]
    assert rows(app)[0][-1] == "Failed – see the messages below"
    assert app.engine_button.cget("text") == "Automatic (recommended)"
