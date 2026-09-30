"""Combine the wanted tabs of every workbook into Excel files with ONE sheet each.

Certificates.xlsx has a single sheet with every device's ECG certificate and
NIBP certificate one after another (a page break between each); the
pass/fail file likewise. Works on the files themselves (openpyxl): it needs
neither Excel nor LibreOffice, and never prints. Formulas are replaced by the
values Excel last calculated, so the result doesn't depend on the originals.
"""

from __future__ import annotations

import os
import threading
from copy import copy
from pathlib import Path

from openpyxl import Workbook

from .builder import BuildResult, Cancelled, OutputError, OutputResult, Reporter, WorkbookResult
from .config import GROUP_TITLES, GROUPS, SHEET_SPECS
from .engines import EngineError, LibreOfficeEngine
from .sheet_copy import load_values, use_default_font
from .stacked_sheet import Block, read_block, write_blocks
from .workbook_info import match_sheets

ENGINE_NAME = "Excel files"
SHEET_TITLES = {"certificates": "Certificates", "passfail": "Pass-Fail test sheets"}
TAB_LABELS = {
    "ecg_cert": "ECG cert",
    "nibp_cert": "NIBP cert",
    "ecg_pf": "ECG pass-fail",
    "nibp_pf": "NIBP pass-fail",
}


class _Sources:
    """Opens source workbooks; old .xls files are converted with LibreOffice when it's installed."""

    def __init__(self) -> None:
        self._libreoffice: LibreOfficeEngine | None = None

    def load(self, path: Path):
        if path.suffix.lower() == ".xls":
            path = self._convert_xls(path)
        try:
            return load_values(path)
        except Exception as exc:
            raise EngineError(f"the file could not be read ({exc or type(exc).__name__})") from None

    def _convert_xls(self, path: Path) -> Path:
        if self._libreoffice is None:
            if not LibreOfficeEngine.is_available():
                raise EngineError("Excel files can only be made from .xlsx files; this is an old .xls file "
                                  "(install LibreOffice to convert it automatically)")
            self._libreoffice = LibreOfficeEngine()
            self._libreoffice.start()
        return self._libreoffice.convert_to_xlsx(path)

    def close(self) -> None:
        if self._libreoffice is not None:
            self._libreoffice.stop()
            self._libreoffice = None


def build_workbooks(
    files: list[Path],
    sheet_names: dict[str, str],
    outputs: dict[str, Path],
    reporter: Reporter,
    cancel: threading.Event | None = None,
) -> BuildResult:
    """Put the tabs of *files* one after another on one sheet per group in *outputs*.

    Tabs are ordered workbook by workbook, and inside each workbook in the
    order given by GROUPS (ECG first, then NIBP).
    """
    cancel = cancel or threading.Event()
    wanted = {spec.key: sheet_names.get(spec.key) or spec.default_name for spec in SHEET_SPECS}
    blocks: dict[str, list[Block]] = {group: [] for group in outputs}
    default_font = None
    results: list[WorkbookResult] = []
    total_steps = len(files) + 1
    sources = _Sources()

    reporter.log("info", f"Making Excel files from {len(files)} workbook(s).")
    try:
        for index, path in enumerate(files):
            if cancel.is_set():
                raise Cancelled()
            path = Path(path)
            reporter.progress(index, total_steps, f"Excel files: reading {path.name}  ({index + 1} of {len(files)})")
            reporter.file_status(index, "working", "Reading tabs…")
            result = WorkbookResult(path)
            results.append(result)
            try:
                source = sources.load(path)
            except EngineError as exc:
                result.problems.append(str(exc))
                reporter.log("error", f"{path.name}: {exc}")
                reporter.file_status(index, "error", "Failed – see the messages below")
                continue
            if default_font is None:
                default_font = copy(source._fonts[0])

            result.found = match_sheets(wanted, source.sheetnames)
            for key, actual in result.found.items():
                if actual is None:
                    reporter.log("warning", f"{path.name}: tab '{wanted[key]}' not found "
                                            f"(tabs in this file: {', '.join(source.sheetnames) or 'none'})")
            copied = []
            for group, keys in GROUPS.items():
                if group not in outputs:
                    continue
                for key in keys:
                    actual = result.found.get(key)
                    if actual is None:
                        continue
                    try:
                        blocks[group].append(read_block(source[actual], f"{path.stem} {TAB_LABELS[key]}"))
                    except Exception as exc:
                        message = f"tab '{actual}' could not be copied: {exc or type(exc).__name__}"
                        result.problems.append(message)
                        reporter.log("error", f"{path.name}: {message}")
                        continue
                    copied.append(actual)
            del source
            if copied:
                reporter.log("info", f"{path.name}: copied {', '.join(copied)}")
            if result.problems:
                reporter.file_status(index, "error", "Failed – see the messages below")
            elif result.missing:
                reporter.file_status(index, "warning", f"Done – {len(result.missing)} tab(s) missing")
            else:
                reporter.file_status(index, "done", "Done")

        if cancel.is_set():
            raise Cancelled()
        reporter.progress(len(files), total_steps, "Saving Excel files…")
        written: dict[str, OutputResult | None] = {}
        for group, out_path in outputs.items():
            written[group] = _save(group, blocks[group], default_font, Path(out_path), reporter)
    finally:
        sources.close()

    reporter.progress(total_steps, total_steps, "Excel files finished")
    return BuildResult(ENGINE_NAME, results, written)


def _save(group: str, blocks: list[Block], default_font, out_path: Path, reporter: Reporter) -> OutputResult | None:
    title = GROUP_TITLES[group]
    if not blocks:
        reporter.log("warning", f"{title}: no Excel file saved – none of the files had these tabs.")
        return None
    book = Workbook()
    if default_font is not None:
        use_default_font(book, default_font)  # column widths are measured in this font
    sheet = book.active
    sheet.title = SHEET_TITLES[group]
    try:
        notes = write_blocks(sheet, blocks)
    except Exception as exc:
        raise OutputError(f"Could not put the {title.lower()} together: {exc or type(exc).__name__}") from None
    for note in notes:
        reporter.log("info", f"{out_path.name}: {note}.")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    partial = out_path.with_name(out_path.name + ".part")
    try:
        book.save(partial)
        os.replace(partial, out_path)
    except PermissionError:
        _remove_quietly(partial)
        raise OutputError(f"Could not save '{out_path.name}'. If it is open in Excel, "
                          f"close it and try again.") from None
    except OSError as exc:
        _remove_quietly(partial)
        raise OutputError(f"Could not save '{out_path}': {exc.strerror or exc}") from None
    reporter.log("success", f"Saved {out_path.name}: {len(blocks)} tabs one after another on one sheet.")
    return OutputResult(out_path, pages=0, tabs=len(blocks))


def _remove_quietly(path: Path) -> None:
    try:
        os.remove(path)
    except OSError:
        pass
