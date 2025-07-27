import re
from pathlib import Path
from datetime import datetime


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
    def replace_ref(match):
        ref = match.group(0)
        absolute_col = ref.startswith('$')
        parts = ref.split('$')

        if len(parts) == 3:  # $A$1
            return ref
        elif len(parts) == 2:  # $A1 or A$1
            if absolute_col:
                col_part = parts[1][0]
                row_part = parts[1][1:]
                row_num = int(row_part) + row_offset
                return f"${col_part}{row_num}"
            else:
                col_part = parts[0][0]
                row_part = parts[1]
                return f"{col_part}${row_part}"
        else:  # A1
            col_part = re.match(r'[A-Z]+', ref).group(0)
            row_part = re.match(r'\d+$', ref).group(0)
            row_num = int(row_part) + row_offset
            # Simple column offset not implemented for brevity
            return f"{col_part}{row_num}"

    pattern = r'(\$?[A-Z]+\$?\d+)'
    return re.sub(pattern, replace_ref, formula)