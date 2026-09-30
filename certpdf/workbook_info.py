"""Reading tab names from workbooks and matching them to the wanted tabs.

Everything here works on the file itself, so it needs neither Excel nor
LibreOffice. It is used to show which tabs were found as soon as files are
added.
"""

from __future__ import annotations

import html
import posixpath
import re
import zipfile
from pathlib import Path

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
OOXML_EXTENSIONS = (".xlsx", ".xlsm")


class WorkbookReadError(Exception):
    """The list of tabs could not be read from a workbook."""


def is_excel_file(path: Path) -> bool:
    # "~$Book.xlsx" files are Excel's lock files, not workbooks.
    return path.suffix.lower() in EXCEL_EXTENSIONS and not path.name.startswith("~$")


# --------------------------------------------------------------------------
# Matching
# --------------------------------------------------------------------------

def normalize(name: str) -> str:
    """'ECG.certificate', 'ECG Certificate' and 'ecg_certificate' all compare equal."""
    return re.sub(r"[^0-9a-z]", "", name.lower())


def match_sheet(wanted: str, available: list[str]) -> str | None:
    """Return the tab in *available* that best matches *wanted*, or None."""
    if wanted in available:
        return wanted
    target = normalize(wanted)
    if not target:
        return None
    for name in available:
        if normalize(name) == target:
            return name
    # Tolerate extra words around the name, e.g. "ECG.certificate (2)",
    # but only when that points at exactly one tab.
    loose = [name for name in available if target in normalize(name)]
    return loose[0] if len(loose) == 1 else None


def match_sheets(wanted: dict[str, str], available: list[str]) -> dict[str, str | None]:
    return {key: match_sheet(name, available) for key, name in wanted.items()}


# --------------------------------------------------------------------------
# Reading tab names
# --------------------------------------------------------------------------

_SHEET_TAG = re.compile(r"<(?:[\w.-]+:)?sheet\s[^>]*>")
_NAME_ATTR = re.compile(r"""\sname\s*=\s*(["'])(.*?)\1""", re.S)


def read_sheet_names(path: Path) -> list[str]:
    """Names of all tabs (worksheets) in the workbook, in tab order."""
    path = Path(path)
    suffix = path.suffix.lower()
    try:
        if suffix in OOXML_EXTENSIONS:
            with zipfile.ZipFile(path) as zf:
                return _sheet_names_from_xml(zf.read(_workbook_part(zf)).decode("utf-8"))
        if suffix == ".xls":
            return _xls_sheet_names(path)
    except WorkbookReadError:
        raise
    except zipfile.BadZipFile:
        raise WorkbookReadError("the file is password protected or damaged") from None
    except (OSError, KeyError, ValueError) as exc:
        raise WorkbookReadError(str(exc) or type(exc).__name__) from None
    raise WorkbookReadError(f"unsupported file type '{path.suffix}'")


def _xls_sheet_names(path: Path) -> list[str]:
    try:
        import xlrd
    except ImportError:
        raise WorkbookReadError("reading .xls files needs the 'xlrd' package") from None
    try:
        book = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:  # xlrd raises a variety of its own errors
        raise WorkbookReadError(str(exc) or "the file could not be read") from None
    try:
        return list(book.sheet_names())
    finally:
        book.release_resources()


def _workbook_part(zf: zipfile.ZipFile) -> str:
    """Path of the main workbook XML part inside an .xlsx/.xlsm package."""
    try:
        rels = zf.read("_rels/.rels").decode("utf-8")
    except KeyError:
        rels = ""
    for rel in re.findall(r"<(?:[\w.-]+:)?Relationship\s[^>]*>", rels):
        rel_type = re.search(r"""\sType\s*=\s*(["'])(.*?)\1""", rel)
        target = re.search(r"""\sTarget\s*=\s*(["'])(.*?)\1""", rel)
        if rel_type and target and rel_type.group(2).endswith("/officeDocument"):
            return posixpath.normpath(target.group(2).lstrip("/"))
    if "xl/workbook.xml" in zf.namelist():
        return "xl/workbook.xml"
    raise WorkbookReadError("this does not look like an Excel workbook")


def _sheet_names_from_xml(workbook_xml: str) -> list[str]:
    names = []
    for tag in _SHEET_TAG.findall(workbook_xml):
        match = _NAME_ATTR.search(tag)
        if match:
            names.append(html.unescape(match.group(2)))
    return names
