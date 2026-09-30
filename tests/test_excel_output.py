import re
import subprocess
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles import Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet.pagebreak import Break

from certpdf.config import CERTIFICATES, PASS_FAIL, default_sheet_names
from certpdf.engines import LibreOfficeEngine
from certpdf.excel_output import build_workbooks
from certpdf.stacked_sheet import DIGIT_PX, read_block, write_blocks
from certpdf.workbook_info import auto_height_rows

TABS = ["Device data", "ECG.certificate", "NIBP.certificate", "ECG.pass-fail test sheet", "NIBP.pass-fail sheet"]
FOOTER = "Code No MECL-TR-05"


class Recorder:
    def __init__(self):
        self.logs, self.statuses = [], {}

    def log(self, level, text):
        self.logs.append((level, text))

    def progress(self, done, total, text):
        assert 0 <= done <= total

    def file_status(self, index, state, text):
        self.statuses[index] = state


def make_template(path: Path, device: str, tabs=TABS, nibp_footer: str = "") -> Path:
    """A workbook shaped like the real templates: Arial 10 default font, merged cells, print setup,
    and ECG and NIBP tabs with different column widths and print scales."""
    wb = Workbook()
    arial = Font(name="Arial", sz=10)
    wb._fonts = IndexedList([arial])
    wb._named_styles["Normal"].font = arial
    wb.remove(wb.active)
    thin = Side(style="thin")
    for tab in tabs:
        nibp = "NIBP" in tab
        ws = wb.create_sheet(tab)
        ws["A1"] = f"{device} {tab}"
        ws["A1"].font = Font(name="Arial Narrow", sz=14, bold=True)
        ws.merge_cells("A1:D1")
        ws["A1"].border = Border(left=thin, top=thin, bottom=thin)
        ws["D1"].border = Border(right=thin, top=thin, bottom=thin)
        for row in range(3, 40):
            ws.cell(row, 1, f"Check {row}")
            ws.cell(row, 2, "PASS")
            ws.cell(row, 3, row * 1.5)
        widths = {"A": 6.9, "B": 9.1, "C": 11.4} if nibp else {"A": 13.7, "B": 4.4}
        for letter, width in widths.items():
            ws.column_dimensions[letter].width = width
        ws.row_dimensions[3].height = 30
        ws.print_area = "A1:D39"
        ws.page_setup.scale = 97 if nibp else 82
        if "pass-fail" in tab:
            ws.oddFooter.center.text = nibp_footer if nibp else FOOTER
            ws.row_breaks.append(Break(id=20))
        # "Hide these cells" rule like the real pass/fail sheets: number format ;;; and no fill,
        # which is how openpyxl reads Excel's patternType="none".
        hide = DifferentialStyle(numFmt=NumberFormat(numFmtId=167, formatCode=";;;"),
                                 fill=PatternFill(bgColor=Color(auto=True)))
        ws.conditional_formatting.add("B30:C35", Rule(type="expression", formula=['$A$1="x"'], dxf=hide))
        ws.conditional_formatting.add("B36:C37", FormulaRule(formula=['$A$1="y"'],
                                                             fill=PatternFill("solid", bgColor="FFFF0000")))
    wb.save(path)
    return path


@pytest.fixture
def template(tmp_path):
    def factory(name, device=None, **kwargs):
        return make_template(tmp_path / name, device or Path(name).stem, **kwargs)
    return factory


def build(files, tmp_path, **kwargs):
    outputs = {CERTIFICATES: tmp_path / "out" / "Certificates.xlsx",
               PASS_FAIL: tmp_path / "out" / "Pass-Fail Test Sheets.xlsx"}
    recorder = Recorder()
    result = build_workbooks(files, default_sheet_names(), outputs, recorder, **kwargs)
    return result, outputs, recorder


def titles_in_order(ws):
    """The A1 title of every stacked tab, top to bottom."""
    found = []
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("dev"):
                found.append((cell.row, cell.value))
    return [value for _, value in sorted(found)]


def test_one_sheet_per_file_with_everything_in_order(template, tmp_path):
    files = [template("dev2.xlsx"), template("dev1.xlsx")]
    result, outputs, recorder = build(files, tmp_path)

    certs = load_workbook(outputs[CERTIFICATES])
    assert certs.sheetnames == ["Certificates"]
    assert titles_in_order(certs.active) == ["dev2 ECG.certificate", "dev2 NIBP.certificate",
                                             "dev1 ECG.certificate", "dev1 NIBP.certificate"]
    passfail = load_workbook(outputs[PASS_FAIL])
    assert passfail.sheetnames == ["Pass-Fail test sheets"]
    assert titles_in_order(passfail.active) == ["dev2 ECG.pass-fail test sheet", "dev2 NIBP.pass-fail sheet",
                                                "dev1 ECG.pass-fail test sheet", "dev1 NIBP.pass-fail sheet"]
    assert result.outputs[CERTIFICATES].tabs == 4
    assert recorder.statuses == {0: "done", 1: "done"}


def test_page_break_after_every_tab(template, tmp_path):
    _, outputs, _ = build([template("dev1.xlsx"), template("dev2.xlsx")], tmp_path)
    ws = load_workbook(outputs[CERTIFICATES]).active
    # four tabs of 39 rows: breaks after rows 39, 78 and 117, none after the last one
    assert [b.id for b in ws.row_breaks.brk] == [39, 78, 117]
    assert ws.print_area == "'Certificates'!$A$1:$" + get_column_letter(ws.max_column) + "$156"


def two_blocks(tmp_path):
    path = make_template(tmp_path / "dev.xlsx", "dev")
    wb = load_workbook(path)
    return [read_block(wb["ECG.certificate"], "ecg"), read_block(wb["NIBP.certificate"], "nibp")]


def test_each_tab_keeps_its_own_column_widths(tmp_path):
    blocks = two_blocks(tmp_path)
    ws = Workbook().active
    write_blocks(ws, blocks)
    grid = [round(ws.column_dimensions[get_column_letter(i)].width * DIGIT_PX)
            for i in range(1, len(ws.column_dimensions) + 1)]
    for block in blocks:
        for (first, last), width in zip(block.columns, block.widths_px):
            assert sum(grid[first - 1:last]) == round(width * block.scaled)


def test_one_print_scale_with_enlarged_tabs(tmp_path):
    blocks = two_blocks(tmp_path)
    ws = Workbook().active
    notes = write_blocks(ws, blocks)
    assert ws.page_setup.scale == 82
    ecg, nibp = blocks
    assert nibp.scaled == pytest.approx(97 / 82)
    assert ws.cell(ecg.first_row, ecg.columns[0][0]).font.sz == 14
    assert ws.cell(nibp.first_row, nibp.columns[0][0]).font.sz == 16.5  # 14 pt at 97 %, printed at 82 %
    assert ws.row_dimensions[nibp.first_row + 2].height == pytest.approx(30 * 97 / 82, abs=0.25)
    assert any("82%" in note for note in notes)


def test_merged_cells_and_values(tmp_path):
    blocks = two_blocks(tmp_path)
    ws = Workbook().active
    write_blocks(ws, blocks)
    merged = {str(r) for r in ws.merged_cells.ranges}
    for block in blocks:
        top = block.first_row
        expected = f"{get_column_letter(block.columns[0][0])}{top}:{get_column_letter(block.columns[3][1])}{top}"
        assert expected in merged
        assert ws.cell(top + 2, block.columns[0][0]).value == "Check 3"
        assert ws.cell(top + 2, block.columns[2][0]).value == 4.5


def test_conditional_formats_point_at_the_moved_cells(tmp_path):
    blocks = two_blocks(tmp_path)
    ws = Workbook().active
    write_blocks(ws, blocks)
    nibp = blocks[1]
    a1 = f"${get_column_letter(nibp.columns[0][0])}${nibp.first_row}"
    formulas = [f for cf in ws.conditional_formatting for rule in cf.rules for f in rule.formula]
    assert f'{a1}="x"' in formulas


def test_no_fill_conditional_format_does_not_become_black(template, tmp_path):
    _, outputs, _ = build([template("dev.xlsx")], tmp_path)
    with zipfile.ZipFile(outputs[PASS_FAIL]) as zf:
        styles = zf.read("xl/styles.xml").decode()
    dxfs = re.findall(r"<dxf>.*?</dxf>", styles, re.S)
    hide = [d for d in dxfs if ";;;" in d]
    # The "no fill" rule keeps hiding the values but must not become a solid (black) fill...
    assert hide and "<fill>" not in hide[0]
    # ...while a real fill colour is kept.
    assert any("<fill>" in d and "FFFF0000" in d for d in dxfs)


def test_different_footers_become_a_line_on_their_own_pages(template, tmp_path):
    _, outputs, recorder = build([template("dev.xlsx")], tmp_path)
    ws = load_workbook(outputs[PASS_FAIL]).active
    assert not ws.oddFooter.center.text
    lines = [c.row for row in ws.iter_rows() for c in row if c.value == FOOTER]
    assert len(lines) == 2  # the ECG pass/fail tab prints on two pages (break after row 20)
    assert any("footer" in text for _, text in recorder.logs)


def test_identical_footers_stay_page_footers(template, tmp_path):
    _, outputs, _ = build([template("dev.xlsx", nibp_footer=FOOTER)], tmp_path)
    ws = load_workbook(outputs[PASS_FAIL]).active
    assert ws.oddFooter.center.text == FOOTER
    assert not [c for row in ws.iter_rows() for c in row if c.value == FOOTER]


def test_default_font_is_kept_so_column_widths_match(template, tmp_path):
    _, outputs, _ = build([template("dev.xlsx")], tmp_path)
    with zipfile.ZipFile(outputs[CERTIFICATES]) as zf:
        styles = zf.read("xl/styles.xml").decode()
    first_font = re.search(r"<fonts[^>]*>\s*<font>(.*?)</font>", styles, re.S).group(1)
    assert 'val="Arial"' in first_font and 'val="10"' in first_font


def test_missing_tabs_and_unreadable_files(template, tmp_path):
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"not a workbook")
    files = [template("full.xlsx"), template("no_nibp.xlsx", tabs=["ECG Certificate", "ECG Pass Fail"]), broken]
    result, outputs, recorder = build(files, tmp_path)

    assert result.outputs[CERTIFICATES].tabs == 3
    assert result.workbooks[1].missing == ["NIBP certificate", "NIBP pass/fail"]
    assert result.workbooks[2].problems and "could not be read" in result.workbooks[2].problems[0]
    assert recorder.statuses == {0: "done", 1: "warning", 2: "error"}


def test_group_without_tabs_is_not_saved(template, tmp_path):
    result, outputs, recorder = build([template("dev.xlsx", tabs=["ECG.certificate"])], tmp_path)
    assert result.outputs[PASS_FAIL] is None and not outputs[PASS_FAIL].exists()
    assert any("no Excel file saved" in text for level, text in recorder.logs if level == "warning")


def test_auto_height_rows(template, tmp_path):
    path = template("dev.xlsx")
    # openpyxl always writes customHeight; make row 3 look like a height Excel calculated itself.
    patched = tmp_path / "patched.xlsx"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(patched, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                data = re.sub(rb'(<row r="3"[^>]*?) customHeight="1"', rb"\1", data)
            zout.writestr(item, data)
    rows = auto_height_rows(patched, ["ECG.certificate", "NIBP.certificate"])
    assert rows == {"ECG.certificate": {3}, "NIBP.certificate": set()}


soffice = LibreOfficeEngine.find()


@pytest.mark.skipif(soffice is None, reason="LibreOffice is not installed")
def test_values_not_formulas(tmp_path):
    # LibreOffice saves the calculated values, like Excel does.
    wb = Workbook()
    wb.active.title = "Device data"
    wb["Device data"]["A1"] = "F23-AGH019-0626"
    cert = wb.create_sheet("ECG.certificate")
    cert["A1"] = "='Device data'!A1"
    cert["A2"] = "=6*7"
    raw = tmp_path / "raw.xlsx"
    wb.save(raw)
    calc_dir = tmp_path / "calc"
    subprocess.run([soffice, f"-env:UserInstallation={(tmp_path / 'profile').as_uri()}", "--headless",
                    "--convert-to", "xlsx", "--outdir", str(calc_dir), str(raw)],
                   check=True, capture_output=True, timeout=180)
    _, outputs, _ = build([calc_dir / "raw.xlsx"], tmp_path)

    values = [c.value for row in load_workbook(outputs[CERTIFICATES]).active.iter_rows() for c in row
              if c.value is not None]
    assert values == ["F23-AGH019-0626", 42]
