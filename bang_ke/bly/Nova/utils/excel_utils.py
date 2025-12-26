from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


def ensure_unmerged(ws, row, col):
    """
    col: chữ cột ('A', 'B', ...) hoặc số cột (1-based)
    """
    if isinstance(col, int):
        col_letter = get_column_letter(col)
    else:
        col_letter = col

    cell = ws[f"{col_letter}{row}"]

    if isinstance(cell, MergedCell):
        for merged_range in list(ws.merged_cells.ranges):
            if cell.coordinate in merged_range:
                ws.unmerge_cells(str(merged_range))
                break
