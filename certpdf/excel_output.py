"""Combine the wanted tabs of every workbook into two Excel files.

Certificates.xlsx has two sheets, "ECG certificates" and "NIBP certificates",
each holding that certificate of every device one after another with a page
break between them; Pass-Fail Test Sheets.xlsx has "ECG reports" and "NIBP
reports". A sheet only holds copies of one template, so it keeps that
template's print scale, margins and footer and prints exactly like the
original tabs. Works on the files themselves (openpyxl): it needs
neither Excel nor LibreOffice, and never prints. Formulas are replaced by the
values Excel last calculated, so the result doesn't depend on the originals.
"""

from __future__ import annotations

import os
import threading
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from openpyxl import Workbook

from .config import GROUP_TITLES, GROUPS, SHEET_SPECS, SPEC_BY_KEY
from .libreoffice import ConversionError, LibreOffice
from .sheet_copy import load_values, use_default_font
from .stacked_sheet import Block, read_block, write_blocks
from .workbook_info import match_sheets



class Reporter(Protocol):
    def log(self, level: str, text: str) -> None: ...
    def progress(self, done: int, total: int, text: str) -> None: ...
    def file_status(self, index: int, state: str, text: str) -> None: ...


class Cancelled(Exception):
    pass


class OutputError(Exception):
    """An Excel file could not be written."""


class ReadError(Exception):
    """A source workbook could not be read."""


@dataclass
class WorkbookResult:
    path: Path
    found: dict[str, str | None] = field(default_factory=dict)   # key -> actual tab name
    problems: list[str] = field(default_factory=list)

    @property
    def missing(self) -> list[str]:
        return [SPEC_BY_KEY[key].label for key, name in self.found.items() if name is None]


@dataclass
class OutputResult:
    path: Path
    tabs: int


@dataclass
class BuildResult:
    workbooks: list[WorkbookResult]
    outputs: dict[str, OutputResult | None]   # group -> written file (None if nothing to write)

SHEET_TITLES = {"ecg_cert": "ECG certificates", "nibp_cert": "NIBP certificates",
                "ecg_pf": "ECG reports", "nibp_pf": "NIBP reports"}
TAB_LABELS = {
    "ecg_cert": "ECG cert",
    "nibp_cert": "NIBP cert",
    "ecg_pf": "ECG pass-fail",
    "nibp_pf": "NIBP pass-fail",
}


class _Sources:
    """Opens source workbooks; old .xls files are converted with LibreOffice when it's installed."""

    def __init__(self) -> None:
        self._libreoffice: LibreOffice | None = None

    def load(self, path: Path):
        try:
            if path.suffix.lower() == ".xls":
                path = self._convert_xls(path)
            return load_values(path)
        except ReadError:
            raise
        except Exception as exc:
            raise ReadError(f"the file could not be read ({exc or type(exc).__name__})") from None

    def _convert_xls(self, path: Path) -> Path:
        if self._libreoffice is None:
            if not LibreOffice.is_available():
                raise ReadError("this is an old .xls file. Open it in Excel and save it as .xlsx, "
                                "or install LibreOffice (free) to convert it automatically")
            self._libreoffice = LibreOffice()
            try:
                self._libreoffice.start()
            except ConversionError as exc:
                self._libreoffice = None
                raise ReadError(str(exc)) from None
        try:
            return self._libreoffice.convert_to_xlsx(path)
        except ConversionError as exc:
            raise ReadError(str(exc)) from None

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
    blocks: dict[str, list[Block]] = {spec.key: [] for spec in SHEET_SPECS}
    default_font = theme = None
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
            except ReadError as exc:
                result.problems.append(str(exc))
                reporter.log("error", f"{path.name}: {exc}")
                reporter.file_status(index, "error", "Failed – see the messages below")
                continue
            if default_font is None:
                default_font = copy(source._fonts[0])
                theme = getattr(source, "loaded_theme", None)  # theme colours and fonts of the templates

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
                        blocks[key].append(read_block(source[actual], f"{path.stem} {TAB_LABELS[key]}"))
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
            sheets = {key: blocks[key] for key in GROUPS[group] if blocks[key]}
            written[group] = _save(group, sheets, default_font, theme, Path(out_path), reporter)
    finally:
        sources.close()

    reporter.progress(total_steps, total_steps, "Excel files finished")
    return BuildResult(results, written)


def _save(group: str, sheets: dict[str, list[Block]], default_font, theme, out_path: Path,
          reporter: Reporter) -> OutputResult | None:
    title = GROUP_TITLES[group]
    if not sheets:
        reporter.log("warning", f"{title}: no Excel file saved – none of the files had these tabs.")
        return None
    book = Workbook()
    book.remove(book.active)
    if default_font is not None:
        use_default_font(book, default_font)  # column widths are measured in this font
    if theme:
        book.loaded_theme = theme
    for key, blocks in sheets.items():
        sheet = book.create_sheet(SHEET_TITLES[key])
        try:
            notes = write_blocks(sheet, blocks)
        except Exception as exc:
            raise OutputError(f"Could not put the {SHEET_TITLES[key]} together: {exc or type(exc).__name__}") from None
        for note in notes:
            reporter.log("info", f"{out_path.name}, {SHEET_TITLES[key]}: {note}.")
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
    count = sum(len(blocks) for blocks in sheets.values())
    reporter.log("success", f"Saved {out_path.name}: sheets {', '.join(SHEET_TITLES[k] for k in sheets)} "
                            f"({count} tabs in all).")
    return OutputResult(out_path, tabs=count)


def _remove_quietly(path: Path) -> None:
    try:
        os.remove(path)
    except OSError:
        pass
