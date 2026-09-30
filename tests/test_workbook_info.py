import re
import zipfile

import pytest

from certpdf.workbook_info import (
    WorkbookReadError,
    is_excel_file,
    isolate_sheet,
    match_sheet,
    normalize,
    read_sheet_names,
)
from conftest import TABS


def test_normalize_ignores_case_and_punctuation():
    assert normalize("ECG.certificate") == normalize("ecg certificate") == normalize("ECG_Certificate")
    assert normalize("NIBP.pass-fail") == normalize("NIBP Pass/Fail")


@pytest.mark.parametrize("available, expected", [
    (["ECG.certificate"], "ECG.certificate"),
    (["Info", "ECG Certificate"], "ECG Certificate"),
    (["ecg-certificate (2)"], "ecg-certificate (2)"),
    (["ECG"], None),
    (["ECG certificate old", "ECG certificate new"], None),  # ambiguous
    ([], None),
])
def test_match_sheet(available, expected):
    assert match_sheet("ECG.certificate", available) == expected


def test_pass_fail_does_not_match_certificate():
    assert match_sheet("ECG.pass-fail", ["ECG.certificate", "NIBP.pass-fail"]) is None


def test_is_excel_file(tmp_path):
    assert is_excel_file(tmp_path / "a.XLSX")
    assert is_excel_file(tmp_path / "a.xls")
    assert not is_excel_file(tmp_path / "~$a.xlsx")
    assert not is_excel_file(tmp_path / "a.pdf")


def test_read_sheet_names(workbook_factory):
    assert read_sheet_names(workbook_factory("dev.xlsx")) == TABS


def test_read_sheet_names_with_special_characters(workbook_factory):
    tabs = ["R&D <x>", "Tom's \"tab\""]
    assert read_sheet_names(workbook_factory("dev.xlsx", tabs=tabs, hidden=())) == tabs


def test_read_sheet_names_rejects_non_workbook(tmp_path):
    fake = tmp_path / "fake.xlsx"
    fake.write_bytes(b"not a zip")
    with pytest.raises(WorkbookReadError):
        read_sheet_names(fake)


def _workbook_xml(path):
    with zipfile.ZipFile(path) as zf:
        return zf.read("xl/workbook.xml").decode()


def test_isolate_sheet_hides_every_other_tab(workbook_factory, tmp_path):
    src = workbook_factory("dev.xlsx")
    dst = tmp_path / "only.xlsx"
    isolate_sheet(src, dst, "NIBP.pass-fail")

    xml = _workbook_xml(dst)
    tags = re.findall(r"<sheet\s[^>]*>", xml)
    states = {re.search(r'name="([^"]*)"', t).group(1): re.search(r'state="([^"]*)"', t) for t in tags}
    assert [name for name, state in states.items() if state is None] == ["NIBP.pass-fail"]
    assert all(s.group(1) == "hidden" for s in states.values() if s is not None)
    assert 'activeTab="4"' in xml and 'firstSheet="4"' in xml
    assert read_sheet_names(dst) == TABS

    # Everything except the workbook part is untouched.
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(dst) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "xl/workbook.xml":
                assert a.read(name) == b.read(name)


def test_isolate_sheet_can_pick_a_hidden_tab(workbook_factory, tmp_path):
    src = workbook_factory("dev.xlsx")
    isolate_sheet(src, tmp_path / "only.xlsx", "Raw data")
    tag = re.search(r'<sheet\s[^>]*name="Raw data"[^>]*>', _workbook_xml(tmp_path / "only.xlsx")).group(0)
    assert "state=" not in tag


def test_isolate_sheet_unknown_tab(workbook_factory, tmp_path):
    with pytest.raises(WorkbookReadError):
        isolate_sheet(workbook_factory("dev.xlsx"), tmp_path / "x.xlsx", "Nope")
