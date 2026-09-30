"""Helpers for copying tabs out of the original workbooks.

Values are the ones Excel last calculated and saved, so the copies need
neither the rest of the original workbook nor any recalculation.
"""

from __future__ import annotations

import warnings
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles.fills import PatternFill
from openpyxl.utils.indexed_list import IndexedList

_PAGE_SETUP_ATTRS = ("orientation", "paperSize", "scale", "fitToWidth", "fitToHeight", "firstPageNumber",
                     "useFirstPageNumber", "pageOrder", "blackAndWhite", "draft", "cellComments", "errors")


def use_default_font(target_book, font) -> None:
    """Give a new workbook the source's default font.

    Excel measures column widths in characters of the default font, so with a
    different one (a new workbook has Calibri 11, the templates Arial 10) every
    column changes width and pages break differently. Call before adding cells.
    """
    if len(target_book._fonts) != 1:
        return  # cells already use other fonts; changing the list would re-point them
    target_book._fonts = IndexedList([copy(font)])
    target_book._named_styles["Normal"].font = copy(font)


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
