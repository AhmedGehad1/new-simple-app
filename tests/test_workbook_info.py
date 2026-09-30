import pytest

from certpdf.workbook_info import (
    WorkbookReadError,
    is_excel_file,
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
