"""Print the wanted tabs of every workbook and combine them into two PDFs."""

from __future__ import annotations

import os
import shutil
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

from pypdf import PdfWriter

from . import APP_NAME
from .config import GROUP_TITLES, GROUPS, SHEET_SPECS, SPEC_BY_KEY
from .engines import EngineError
from .workbook_info import match_sheets


class Reporter(Protocol):
    def log(self, level: str, text: str) -> None: ...
    def progress(self, done: int, total: int, text: str) -> None: ...
    def file_status(self, index: int, state: str, text: str) -> None: ...


class Cancelled(Exception):
    pass


class OutputError(Exception):
    """A combined PDF could not be written."""


@dataclass
class WorkbookResult:
    path: Path
    found: dict[str, str | None] = field(default_factory=dict)   # key -> actual tab name
    pdfs: dict[str, Path] = field(default_factory=dict)          # key -> printed tab
    problems: list[str] = field(default_factory=list)

    @property
    def missing(self) -> list[str]:
        return [SPEC_BY_KEY[key].label for key, name in self.found.items() if name is None]


@dataclass
class OutputResult:
    path: Path
    pages: int
    tabs: int


@dataclass
class BuildResult:
    engine_name: str
    workbooks: list[WorkbookResult]
    outputs: dict[str, OutputResult | None]   # group -> written PDF (None if nothing to write)

    @property
    def problem_count(self) -> int:
        return sum(len(w.problems) for w in self.workbooks)


def device_title(path: Path) -> str:
    return Path(path).stem


def build_pdfs(
    files: list[Path],
    sheet_names: dict[str, str],
    outputs: dict[str, Path],
    engine,
    reporter: Reporter,
    cancel: threading.Event | None = None,
) -> BuildResult:
    """Print the tabs of *files* with *engine* and write one PDF per group in *outputs*.

    Pages are ordered workbook by workbook, and inside each workbook in the
    order given by GROUPS (ECG first, then NIBP).
    """
    cancel = cancel or threading.Event()
    wanted = {spec.key: sheet_names.get(spec.key) or spec.default_name for spec in SHEET_SPECS}
    total_steps = len(files) + 1
    results: list[WorkbookResult] = []

    reporter.log("info", f"Using {engine.name} to convert {len(files)} workbook(s).")
    tmp_dir = Path(tempfile.mkdtemp(prefix="certpdf-"))
    try:
        for index, path in enumerate(files):
            if cancel.is_set():
                raise Cancelled()
            path = Path(path)
            reporter.progress(index, total_steps, f"Converting {path.name}  ({index + 1} of {len(files)})")
            reporter.file_status(index, "working", "Converting…")
            result = _print_workbook(engine, path, wanted, tmp_dir / f"{index:04d}", reporter)
            results.append(result)
            if result.problems:
                reporter.file_status(index, "error", "Partly done – see the messages below" if result.pdfs
                                     else "Failed – see the messages below")
            elif result.missing:
                reporter.file_status(index, "warning", f"Done – {len(result.missing)} tab(s) missing")
            else:
                reporter.file_status(index, "done", "Done")

        if cancel.is_set():
            raise Cancelled()
        reporter.progress(len(files), total_steps, "Combining PDFs…")
        written: dict[str, OutputResult | None] = {}
        for group, out_path in outputs.items():
            written[group] = _combine(group, results, Path(out_path), reporter)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    reporter.progress(total_steps, total_steps, "Finished")
    return BuildResult(engine.name, results, written)


def _print_workbook(engine, path: Path, wanted: dict[str, str], tmp_dir: Path, reporter: Reporter) -> WorkbookResult:
    result = WorkbookResult(path)
    tmp_dir.mkdir(parents=True, exist_ok=True)
    try:
        with engine.open(path) as workbook:
            result.found = match_sheets(wanted, workbook.sheet_names)
            for key, actual in result.found.items():
                if actual is None:
                    reporter.log("warning", f"{path.name}: tab '{wanted[key]}' not found "
                                            f"(tabs in this file: {', '.join(workbook.sheet_names) or 'none'})")
            jobs = {key: (actual, tmp_dir / f"{key}.pdf") for key, actual in result.found.items() if actual}
            errors = workbook.export_many(list(jobs.values())) if jobs else {}
    except EngineError as exc:
        result.problems.append(str(exc))
        reporter.log("error", f"{path.name}: {exc}")
        return result

    for key, (actual, pdf_path) in jobs.items():
        if actual in errors:
            message = f"tab '{actual}' could not be converted: {errors[actual]}"
            result.problems.append(message)
            reporter.log("error", f"{path.name}: {message}")
        else:
            result.pdfs[key] = pdf_path
    printed = ", ".join(jobs[key][0] for key in result.pdfs)
    if printed:
        reporter.log("info", f"{path.name}: converted {printed}")
    return result


def _combine(group: str, results: list[WorkbookResult], out_path: Path, reporter: Reporter) -> OutputResult | None:
    writer = PdfWriter()
    tabs = 0
    for result in results:
        parent = None
        for key in GROUPS[group]:
            pdf = result.pdfs.get(key)
            if pdf is None:
                continue
            start = len(writer.pages)
            writer.append(str(pdf), import_outline=False)
            if len(writer.pages) == start:
                continue
            # Bookmarks: one per workbook, with one entry per tab under it.
            if parent is None:
                parent = writer.add_outline_item(device_title(result.path), start)
            writer.add_outline_item(SPEC_BY_KEY[key].label, start, parent=parent)
            tabs += 1

    title = GROUP_TITLES[group]
    if tabs == 0:
        reporter.log("warning", f"{title}: nothing to save – none of the files had these tabs.")
        return None

    writer.add_metadata({"/Title": title, "/Creator": APP_NAME, "/Producer": APP_NAME})
    out_path.parent.mkdir(parents=True, exist_ok=True)
    partial = out_path.with_name(out_path.name + ".part")
    try:
        with open(partial, "wb") as handle:
            writer.write(handle)
        os.replace(partial, out_path)
    except PermissionError:
        _remove_quietly(partial)
        raise OutputError(f"Could not save '{out_path.name}'. If it is open in a PDF viewer, "
                          f"close it and try again.") from None
    except OSError as exc:
        _remove_quietly(partial)
        raise OutputError(f"Could not save '{out_path}': {exc.strerror or exc}") from None
    pages = len(writer.pages)
    reporter.log("success", f"Saved {out_path.name}: {tabs} tab(s), {pages} page(s).")
    return OutputResult(out_path, pages, tabs)


def _remove_quietly(path: Path) -> None:
    try:
        os.remove(path)
    except OSError:
        pass
