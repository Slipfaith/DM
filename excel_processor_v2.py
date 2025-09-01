import re
from logger import get_logger


class ExcelProcessorV2:
    def __init__(self, config):
        self.config = config
        self.logger = get_logger()
        self._progress_callback = None

    def set_progress_callback(self, callback):
        self._progress_callback = callback

    def can_process(self, sheet):
        used_range = sheet.UsedRange
        if not used_range:
            return False

        yellow_headers_count = 0

        for row in range(1, min(50, used_range.Rows.Count + 1)):
            if sheet.Rows(row).Hidden:
                continue
            for col in range(1, min(10, used_range.Columns.Count + 1)):
                cell = sheet.Cells(row, col)
                if cell.Interior.Color == self.config.header_color and self._normalize_value(cell.Value):
                    yellow_headers_count += 1
                    break

        return yellow_headers_count >= 2

    def process_sheet(self, sheet):
        self.logger.info(f"Processing sheet '{sheet.Name}' with V2 method")

        used_range = sheet.UsedRange
        blocks = self._find_all_blocks(sheet, used_range)

        if not blocks:
            self.logger.info("No data blocks found")
            return

        sheet.Application.ScreenUpdating = False
        sheet.Application.Calculation = -4135

        total_groups = sum(len(block['data_groups']) for block in blocks)
        processed_groups = 0

        for block in reversed(blocks):
            for group in reversed(block['data_groups']):
                self._duplicate_block_rows(sheet, block, used_range, group)
                processed_groups += 1

                if self._progress_callback:
                    self._progress_callback(processed_groups, total_groups)

        sheet.Application.Calculation = -4105
        sheet.Application.ScreenUpdating = True
        # After duplicating all blocks Excel may end up with copied header
        # rows at the bottom of the sheet. These duplicated headers serve no
        # purpose and confuse users when exporting the result. Remove any
        # stray header rows that appear after the first data block.
        self._remove_duplicate_headers(sheet)

        self.logger.info(f"Processed {len(blocks)} blocks")

    def _find_all_blocks(self, sheet, used_range):
        blocks = []
        current_row = 1
        last_row = used_range.Row + used_range.Rows.Count - 1
        cols_count = used_range.Columns.Count

        while current_row <= last_row:
            if sheet.Rows(current_row).Hidden:
                current_row += 1
                continue
            is_header = False
            for col in range(1, min(10, cols_count + 1)):
                cell = sheet.Cells(current_row, col)
                if cell.Interior.Color == self.config.header_color and self._normalize_value(cell.Value):
                    is_header = True
                    break

            if is_header:
                block = {
                    'header_row': current_row,
                    'data_groups': []
                }

                current_row += 1
                current_group = []

                while current_row <= last_row:
                    if sheet.Rows(current_row).Hidden:
                        current_row += 1
                        continue
                    is_next_header = False
                    for col in range(1, min(10, cols_count + 1)):
                        if sheet.Cells(current_row, col).Interior.Color == self.config.header_color and \
                                self._normalize_value(sheet.Cells(current_row, col).Value):
                            is_next_header = True
                            break

                    if is_next_header:
                        if current_group:
                            block['data_groups'].append(current_group)
                        break

                    has_data = False
                    for col in range(1, cols_count + 1):
                        cell = sheet.Cells(current_row, col)
                        if self._normalize_value(cell.Value) or cell.HasFormula:
                            has_data = True
                            break

                    if not has_data:
                        if current_group:
                            block['data_groups'].append(current_group)
                            current_group = []
                        current_row += 1
                        break
                    else:
                        current_group.append(current_row)
                        current_row += 1

                if current_group:
                    block['data_groups'].append(current_group)

                if block['data_groups']:
                    blocks.append(block)
            else:
                current_row += 1

        return blocks

    def _duplicate_block_rows(self, sheet, block, used_range, group):
        cols_count = used_range.Columns.Count

        if not group:
            return

        group_size = len(group)

        insert_row = group[-1] + 1
        for _ in range(group_size):
            sheet.Rows(insert_row).Insert(Shift=-4121)

        for i, source_row in enumerate(group):
            target_row = insert_row + i

            sheet.Rows(source_row).Copy()
            sheet.Rows(target_row).PasteSpecial(-4104)

            for col in range(1, cols_count + 1):
                cell = sheet.Cells(target_row, col)
                if cell.HasFormula:
                    formula = cell.Formula
                    if "LEN(" in formula.upper() or "ДЛСТР(" in formula.upper():
                        col_letter = sheet.Cells(1, col).Address.split("$")[1]
                        if i > 0:
                            ref_row = target_row - 1
                        else:
                            ref_row = target_row
                        formula = re.sub(
                            r'(LEN|ДЛСТР)\s*\([^)]+\)',
                            rf'\1({col_letter}{ref_row})',
                            formula,
                            flags=re.IGNORECASE
                        )
                        cell.Formula = formula

        self._copy_shapes_in_range(sheet, group[0], group[-1], insert_row)

        sheet.Application.CutCopyMode = False

    def _copy_shapes_in_range(self, sheet, start_row, end_row, target_start_row):
        try:
            for shape in sheet.Shapes:
                shape_row = shape.TopLeftCell.Row
                if start_row <= shape_row <= end_row:
                    row_offset = shape_row - start_row

                    shape.Copy()
                    sheet.Paste()

                    new_shape = sheet.Shapes(sheet.Shapes.Count)

                    target_cell = sheet.Cells(target_start_row + row_offset, shape.TopLeftCell.Column)
                    new_shape.Top = target_cell.Top + (shape.Top - shape.TopLeftCell.Top)
                    new_shape.Left = target_cell.Left + (shape.Left - shape.TopLeftCell.Left)
        except Exception as e:
            self.logger.warning(f"Error copying shapes: {e}")

    def _normalize_value(self, value):
        """Return a stripped string representation of a cell value."""
        if value is None:
            return ""
        return str(value).strip()

    def _remove_duplicate_headers(self, sheet):
        """Remove extra header rows accidentally copied to the bottom.

        The processing steps duplicate ranges of rows, which occasionally
        results in the original header row being copied as data to the end of
        the sheet. This helper scans rows below the first header and deletes
        any rows that exactly match the header's contents and color.
        """
        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1
        cols_count = used_range.Columns.Count

        header_row = None
        header_values = []

        # Locate first header row by color
        for row in range(1, last_row + 1):
            if sheet.Rows(row).Hidden:
                continue
            for col in range(1, cols_count + 1):
                cell = sheet.Cells(row, col)
                if cell.Interior.Color == self.config.header_color and self._normalize_value(cell.Value):
                    header_row = row
                    header_values = [
                        self._normalize_value(sheet.Cells(row, c).Value)
                        for c in range(1, cols_count + 1)
                    ]
                    break
            if header_row:
                break

        if not header_row:
            return

        # Walk from bottom upwards and remove duplicated headers
        for row in range(last_row, header_row, -1):
            if sheet.Rows(row).Hidden:
                continue
            is_header = True
            for col in range(1, cols_count + 1):
                cell = sheet.Cells(row, col)
                cell_value = self._normalize_value(cell.Value)
                if cell.Interior.Color != self.config.header_color or cell_value != header_values[col - 1]:
                    is_header = False
                    break

            if is_header:
                sheet.Rows(row).Delete()
