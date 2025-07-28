# formula_updater.py
"""Utility to refresh formulas referencing the first sheet after row reorder."""

import re
from typing import List, Tuple

from excel_com import ExcelCOM
from logger import get_logger


def _snapshot_first_sheet(sheet) -> List[Tuple]:
    """Capture values of all rows on the given sheet."""
    used = sheet.UsedRange
    last_row = used.Row + used.Rows.Count - 1
    last_col = used.Column + used.Columns.Count - 1
    snapshot = []
    for r in range(used.Row, last_row + 1):
        row_values = [sheet.Cells(r, c).Value for c in range(used.Column, last_col + 1)]
        snapshot.append(tuple(row_values))
    return snapshot


def _build_row_map(old: List[Tuple], new: List[Tuple]) -> dict:
    """Map old row numbers to new based on row content."""
    mapping = {}
    value_to_rows = {}
    for idx, row in enumerate(new):
        value_to_rows.setdefault(row, []).append(idx + 1)

    used = set()
    for idx, row in enumerate(old):
        options = value_to_rows.get(row, [])
        new_index = next((i for i in options if i not in used), None)
        if new_index:
            mapping[idx + 1] = new_index
            used.add(new_index)
    return mapping


def _adjust_formula(formula: str, sheet_name: str, row_map: dict) -> str:
    """Replace references to ``sheet_name`` using ``row_map``."""
    escaped = re.escape(sheet_name).replace("'", "''")
    pattern = re.compile(rf"((?:'{escaped}'|{escaped})!\$?[A-Za-z]{{1,3}}\$?)(\d+)")

    def repl(match):
        prefix = match.group(1)
        row_num = int(match.group(2))
        if row_num in row_map:
            return f"{prefix}{row_map[row_num]}"
        return match.group(0)

    return pattern.sub(repl, formula)


def update_formulas(workbook_path: str, old_snapshot: List[Tuple]):
    """Update formulas in ``workbook_path`` after rows on the first sheet moved."""
    logger = get_logger()
    with ExcelCOM() as excel:
        wb = excel.open_workbook(workbook_path)
        first_sheet = wb.Sheets(1)
        sheet_name = first_sheet.Name
        new_snapshot = _snapshot_first_sheet(first_sheet)
        row_map = _build_row_map(old_snapshot, new_snapshot)
        if not row_map:
            logger.info("No row changes detected. Nothing to update.")
            wb.Close(False)
            return

        for sheet in wb.Sheets:
            if sheet.Index == 1:
                continue
            used = sheet.UsedRange
            for r in range(1, used.Rows.Count + 1):
                for c in range(1, used.Columns.Count + 1):
                    cell = sheet.Cells(r, c)
                    if cell.HasFormula:
                        new_formula = _adjust_formula(cell.Formula, sheet_name, row_map)
                        if new_formula != cell.Formula:
                            cell.Formula = new_formula

        wb.Save()
        wb.Close(False)
        logger.info("Formulas updated successfully.")
