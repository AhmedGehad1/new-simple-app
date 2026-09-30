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
