import re
from pathlib import Path
from datetime import datetime
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


def generate_unique_filename(output_dir: Path, base_name: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_name}_processed_{timestamp}.xlsx"

    counter = 1
    while (output_dir / filename).exists():
        filename = f"{base_name}_processed_{timestamp}_{counter}.xlsx"
        counter += 1

    return filename


def parse_excel_address(address: str) -> tuple:
    match = re.match(r'([A-Z]+)(\d+)', address)
    if match:
        col_str, row_str = match.groups()
        col = sum((ord(c) - ord('A') + 1) * (26 ** i)
                  for i, c in enumerate(reversed(col_str)))
        return int(row_str), col
    return None, None


def adjust_formula_references(formula: str, row_offset: int, col_offset: int) -> str:
    """Shift cell references in ``formula`` by the given offsets.

    This util relies on :class:`openpyxl.formula.translate.Translator` which
    mimics Excel's behaviour when copying formulas. By using ``A1`` as the
    origin cell and offsetting to the destination cell, relative references are
    adjusted while absolute references remain intact. Cross-sheet references are
    also handled by ``Translator``.
    """

    try:
        origin = "A1"
        dest_col = 1 + col_offset
        dest_row = 1 + row_offset
        dest = f"{get_column_letter(dest_col)}{dest_row}"
        return Translator(formula, origin=origin).translate_formula(dest)
    except Exception:
        # If translation fails for any reason, return the original formula
        return formula
