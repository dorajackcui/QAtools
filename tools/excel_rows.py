"""Coordinate-preserving Excel reads that never expand a declared rectangle.

The openpyxl 3.x internals used here are isolated deliberately: its public
read-only iterator pads missing rows/columns even after reset_dimensions().
Use its own parser so shared strings, formulas, dates and cell types retain
the same interpretation, but yield only stored cells and their real positions.
"""
from __future__ import annotations

from collections import defaultdict
from contextlib import closing

from openpyxl.cell.cell import MergedCell
from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.worksheet._reader import WorkSheetParser


def iter_stored_rows(worksheet, *, min_row=1, max_row=None, min_col=1,
                     max_col=None, include_empty=False):
    """Yield (original row number, {original column number: cell}).

    None is skipped by default; zero, False, empty text and formulas are values.
    include_empty also exposes stored style-only cells, never invented cells.
    Close the generator explicitly when stopping early on a read-only sheet.
    """
    def selected(cell):
        return (min_col <= cell.column and (max_col is None or cell.column <= max_col)
                and (include_empty or cell.value is not None))

    if hasattr(worksheet, "_cells"):
        rows = defaultdict(dict)
        for (row, column), cell in worksheet._cells.items():
            if row >= min_row and (max_row is None or row <= max_row) and selected(cell):
                rows[row][column] = cell
        for row in sorted(rows):
            yield row, rows[row]
        return

    with worksheet._get_source() as source:
        parser = WorkSheetParser(
            source, worksheet._shared_strings,
            data_only=worksheet.parent.data_only, epoch=worksheet.parent.epoch,
            date_formats=worksheet.parent._date_formats,
            timedelta_formats=worksheet.parent._timedelta_formats,
        )
        for row, parsed in parser.parse():
            if max_row is not None and row > max_row:
                break
            if row < min_row:
                continue
            cells = {}
            for data in parsed:
                cell = ReadOnlyCell(worksheet, **data)
                if selected(cell):
                    cells[cell.column] = cell
            if cells:
                yield row, cells


def stored_row(worksheet, row=1):
    """Read one row without leaking a partially consumed ZIP stream."""
    with closing(iter_stored_rows(worksheet, min_row=row, max_row=row)) as rows:
        return next(rows, (row, {}))[1]


def row_values(cells):
    """A positional tuple bounded by this row's actual values, not sheet size."""
    return tuple(cells[column].value if column in cells else None
                 for column in range(1, max(cells, default=0) + 1))


def discard_empty_cells(worksheet):
    """Drop inert stored blanks before saving; preserve all formatting/metadata.

    This does not delete rows or columns or move any surviving coordinates.
    Styled cells, comments, links and merged ranges belong to the original
    document and must survive, including in a hidden restoration template.
    """
    merged_anchors = {(area.min_row, area.min_col) for area in worksheet.merged_cells.ranges}
    empty = [key for key, cell in worksheet._cells.items()
             if key not in merged_anchors and not isinstance(cell, MergedCell) and cell.value is None
             and not cell.has_style and cell.comment is None and cell.hyperlink is None]
    for key in empty:
        del worksheet._cells[key]
    worksheet._current_row = max((row for row, _ in worksheet._cells), default=0)
