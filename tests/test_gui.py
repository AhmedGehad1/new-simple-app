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


def test_creates_the_two_excel_files(app, workbook_factory, tmp_path, dialogs):
    out = tmp_path / "out"
    app.add_paths([workbook_factory("dev1.xlsx"), workbook_factory("dev2.xlsx")])
    app.output_var.set(str(out))
    app.cert_name_var.set("Batch 7 certificates.xlsx")  # a typed extension is dropped
    run_to_end(app)

    assert sorted(p.name for p in out.iterdir()) == ["Batch 7 certificates.xlsx", "Pass-Fail Test Sheets.xlsx"]
    kind, title, message = dialogs[-1]
    assert kind == "showinfo" and title == "Files created"
    assert [r[-1] for r in rows(app)] == ["Done", "Done"]
    assert "2 files saved" in app.status_label.cget("text")


def test_unreadable_file_is_reported(app, workbook_factory, tmp_path, dialogs):
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a workbook")
    out = tmp_path / "out"
    app.add_paths([workbook_factory("dev.xlsx"), broken])
    app.output_var.set(str(out))
    run_to_end(app)

    assert sorted(p.name for p in out.iterdir()) == ["Certificates.xlsx", "Pass-Fail Test Sheets.xlsx"]
    assert [r[-1] for r in rows(app)] == ["Done", "Failed – see the messages below"]
    kind, title, message = dialogs[-1]
    assert kind == "showwarning" and "had problems: broken.xlsx" in message
