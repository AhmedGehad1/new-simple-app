"""Copy a worksheet, with its formatting and page setup, into another workbook.

Used for the Excel files and for making PDFs with LibreOffice. Values are the
ones Excel last calculated and saved, so the copies need neither the rest of
the original workbook nor any recalculation.
"""

from __future__ import annotations

import re
import warnings
from copy import copy, deepcopy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.worksheet.dimensions import ColumnDimension

A4 = 9  # used when a tab doesn't say which paper it is for

MAX_TAB_NAME = 31
_INVALID_TAB_CHARS = re.compile(r"[\[\]:*?/\\]")
_PAGE_SETUP_ATTRS = ("orientation", "paperSize", "scale", "fitToWidth", "fitToHeight", "firstPageNumber",
                     "useFirstPageNumber", "pageOrder", "blackAndWhite", "draft", "cellComments", "errors")
_VIEW_ATTRS = ("showGridLines", "zoomScale", "zoomScaleNormal", "zoomScalePageLayoutView", "view")


def tab_title(stem: str, label: str, used: set[str]) -> str:
    """A valid, unique tab name such as 'F23-AGH019-0626 ECG cert' (at most 31 characters)."""
    stem = _INVALID_TAB_CHARS.sub("-", stem).strip("' ")
    room = MAX_TAB_NAME - len(label) - 1
    base = f"{stem[:room].rstrip()} {label}" if stem else label
    title, number = base, 2
    while title.lower() in used:
        suffix = f" ({number})"
        title = base[:MAX_TAB_NAME - len(suffix)].rstrip() + suffix
        number += 1
    used.add(title.lower())
    return title


def _copy_style(source, target) -> None:
    target.font = copy(source.font)
    target.border = copy(source.border)
    target.fill = copy(source.fill)
    target.number_format = source.number_format
    target.protection = copy(source.protection)
    target.alignment = copy(source.alignment)


def adopt_default_font(target_book, source_book) -> None:
    """Give a new workbook the source's default font.

    Excel measures column widths in characters of the default font, so with a
    different one (a new workbook has Calibri 11, the templates Arial 10) every
    column changes width and pages break differently. Call before adding cells.
    """
    if len(target_book._fonts) != 1:
        return  # cells already use other fonts; changing the list would re-point them
    font = copy(source_book._fonts[0])
    target_book._fonts = IndexedList([font])
    target_book._named_styles["Normal"].font = copy(font)


def copy_sheet(source, target, auto_height_rows: set[int] = frozenset()) -> None:
    """Copy one worksheet into another workbook: values, formatting and page setup.

    Rows in *auto_height_rows* keep an automatic height instead of the height
    Excel calculated (see workbook_info.auto_height_rows).
    """
    # Merge first: merging resets the covered cells, whose borders are copied below.
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))

    for row in source.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            new = target.cell(row=cell.row, column=cell.column)
            if not isinstance(cell, MergedCell) and not isinstance(new, MergedCell) and cell.value is not None:
                new.value = cell.value
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    new.data_type = "s"  # text that only looks like a formula
            if cell.has_style:
                _copy_style(cell, new)

    for key, dim in source.column_dimensions.items():
        new = ColumnDimension(target, index=key, width=dim.width, bestFit=dim.bestFit, hidden=dim.hidden,
                              outlineLevel=dim.outlineLevel, collapsed=dim.collapsed,
                              min=dim.min, max=dim.max, customWidth=dim.customWidth)
        if dim.has_style:
            _copy_style(dim, new)
        target.column_dimensions[key] = new
    for index, dim in source.row_dimensions.items():
        new = target.row_dimensions[index]
        new.height = None if index in auto_height_rows else dim.height
        new.hidden = dim.hidden
        new.outlineLevel, new.collapsed = dim.outlineLevel, dim.collapsed
        if dim.has_style:
            _copy_style(dim, new)
    target.sheet_format = copy(source.sheet_format)

    # Printing: page setup, margins, header/footer, print area, page breaks
    for attr in _PAGE_SETUP_ATTRS:
        setattr(target.page_setup, attr, getattr(source.page_setup, attr))
    if target.page_setup.paperSize is None:
        target.page_setup.paperSize = A4
    target.sheet_properties.pageSetUpPr = copy(source.sheet_properties.pageSetUpPr)
    target.sheet_properties.tabColor = copy(source.sheet_properties.tabColor)
    target.page_margins = copy(source.page_margins)
    target.print_options = copy(source.print_options)
    target.HeaderFooter = deepcopy(source.HeaderFooter)
    if source.print_area:
        # Drop the sheet name: "'ECG.certificate'!$A$1:$J$51" -> "$A$1:$J$51"
        areas = re.findall(r"!([$A-Za-z0-9:]+)", source.print_area)
        if areas:
            target.print_area = areas
    if source.print_title_rows:
        target.print_title_rows = source.print_title_rows
    if source.print_title_cols:
        target.print_title_cols = source.print_title_cols
    for brk in source.row_breaks.brk:
        target.row_breaks.append(copy(brk))
    for brk in source.col_breaks.brk:
        target.col_breaks.append(copy(brk))

    for formatting in source.conditional_formatting:
        for rule in formatting.rules:
            new_rule = copy(rule)
            new_rule.dxf = _without_empty_fill(deepcopy(rule.dxf))
            target.conditional_formatting.add(str(formatting.sqref), new_rule)

    for attr in _VIEW_ATTRS:
        setattr(target.sheet_view, attr, getattr(source.sheet_view, attr))


def _without_empty_fill(dxf):
    """Drop a conditional format's "no fill" setting.

    openpyxl reads patternType="none" as "not set" and writes it back without
    it, and Excel then shows a *solid* fill in the automatic colour: black
    boxes. A fill without a pattern and without a real colour means "no fill".
    """
    fill = getattr(dxf, "fill", None)
    if isinstance(fill, PatternFill) and fill.patternType is None:
        colour = fill.bgColor
        if colour is None or colour.type == "auto" or (colour.type == "rgb" and colour.rgb == "00000000") \
                or (colour.type == "indexed" and colour.indexed in (64, 65)):
            dxf.fill = None
    return dxf


def load_values(path: Path):
    """Open a workbook with the values Excel saved (not the formulas)."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # openpyxl warns about features it doesn't need to read
        return load_workbook(path, data_only=True, rich_text=True, keep_links=False)
