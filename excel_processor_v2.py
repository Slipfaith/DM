# excel_processor_v2.py

import re
from logger import get_logger


class ExcelProcessorV2:
    def __init__(self, config):
        self.config = config
        self.logger = get_logger()

    def can_process(self, sheet):
        """Check if sheet has yellow headers with data blocks structure"""
        used_range = sheet.UsedRange
        if not used_range:
            return False

        # Check only first 20 rows for performance
        for row in range(1, min(20, used_range.Rows.Count + 1)):
            # Check only first few columns for yellow header
            for col in range(1, min(10, used_range.Columns.Count + 1)):
                cell = sheet.Cells(row, col)
                if cell.Interior.Color == self.config.header_color and cell.Value:
                    # Found at least one yellow header, check if there's data after it
                    if row + 1 <= used_range.Rows.Count:
                        for check_col in range(1, used_range.Columns.Count + 1):
                            if sheet.Cells(row + 1, check_col).Value:
                                return True
                    break

        return False

    def process_sheet(self, sheet):
        self.logger.info(f"Processing sheet '{sheet.Name}' with V2 method")

        used_range = sheet.UsedRange
        blocks = self._find_all_blocks(sheet, used_range)

        if not blocks:
            self.logger.info("No data blocks found")
            return

        sheet.Application.ScreenUpdating = False
        sheet.Application.Calculation = -4135  # xlCalculationManual

        # Process blocks from bottom to top
        for block in reversed(blocks):
            self._duplicate_block_rows(sheet, block, used_range)

        sheet.Application.Calculation = -4105  # xlCalculationAutomatic
        sheet.Application.ScreenUpdating = True
        self.logger.info(f"Processed {len(blocks)} blocks")

    def _find_all_blocks(self, sheet, used_range):
        blocks = []
        current_row = 1
        last_row = used_range.Row + used_range.Rows.Count - 1
        cols_count = used_range.Columns.Count

        while current_row <= last_row:
            # Check if current row is yellow header
            is_header = False
            for col in range(1, min(10, cols_count + 1)):  # Check only first 10 cols for performance
                cell = sheet.Cells(current_row, col)
                if cell.Interior.Color == self.config.header_color and cell.Value:
                    is_header = True
                    break

            if is_header:
                block = {
                    'header_row': current_row,
                    'data_groups': []  # Groups of data rows (including formula rows)
                }

                current_row += 1
                current_group = []

                # Collect data rows until empty row or next header
                while current_row <= last_row:
                    # Check if row is yellow header
                    is_next_header = False
                    for col in range(1, min(10, cols_count + 1)):
                        if sheet.Cells(current_row, col).Interior.Color == self.config.header_color and sheet.Cells(
                                current_row, col).Value:
                            is_next_header = True
                            break

                    if is_next_header:
                        if current_group:
                            block['data_groups'].append(current_group)
                        break

                    # Check if row has any data
                    has_data = False
                    for col in range(1, cols_count + 1):
                        if sheet.Cells(current_row, col).Value or sheet.Cells(current_row, col).HasFormula:
                            has_data = True
                            break

                    if not has_data:
                        # Empty row - save current group if exists
                        if current_group:
                            block['data_groups'].append(current_group)
                            current_group = []
                        current_row += 1
                        break
                    else:
                        # Row with data or formulas
                        current_group.append(current_row)
                        current_row += 1

                # Save last group if exists
                if current_group:
                    block['data_groups'].append(current_group)

                if block['data_groups']:
                    blocks.append(block)
            else:
                current_row += 1

        return blocks

    def _duplicate_block_rows(self, sheet, block, used_range):
        cols_count = used_range.Columns.Count

        # Process data groups from bottom to top
        for group in reversed(block['data_groups']):
            if not group:
                continue

            # Determine group structure (data rows + formula rows)
            group_size = len(group)

            # Insert rows for the duplicate
            insert_row = group[-1] + 1
            for _ in range(group_size):
                sheet.Rows(insert_row).Insert(Shift=-4121)

            # Copy each row in the group
            for i, source_row in enumerate(group):
                target_row = insert_row + i

                # Copy row formatting and values
                sheet.Rows(source_row).Copy()
                sheet.Rows(target_row).PasteSpecial(-4104)  # xlPasteAll

                # Fix formulas in the duplicate
                for col in range(1, cols_count + 1):
                    cell = sheet.Cells(target_row, col)
                    if cell.HasFormula:
                        formula = cell.Formula
                        # Check if this is a LEN/ДЛСТР formula
                        if "LEN(" in formula.upper() or "ДЛСТР(" in formula.upper():
                            # For formula rows in duplicate, reference should point to duplicate data row
                            col_letter = sheet.Cells(1, col).Address.split("$")[1]
                            # Calculate which row to reference (the duplicate of the data row)
                            if i > 0:  # This is a formula row
                                ref_row = target_row - 1  # Reference the row above
                            else:  # This is the first row, reference itself
                                ref_row = target_row
                            formula = re.sub(
                                r'(LEN|ДЛСТР)\s*\([^)]+\)',
                                rf'\1({col_letter}{ref_row})',
                                formula,
                                flags=re.IGNORECASE
                            )
                            cell.Formula = formula

            # Copy shapes (pictures) that are in the original group range
            self._copy_shapes_in_range(sheet, group[0], group[-1], insert_row)

        sheet.Application.CutCopyMode = False

    def _copy_shapes_in_range(self, sheet, start_row, end_row, target_start_row):
        """Copy shapes (pictures) from source range to target range"""
        try:
            for shape in sheet.Shapes:
                shape_row = shape.TopLeftCell.Row
                if start_row <= shape_row <= end_row:
                    # Calculate offset
                    row_offset = shape_row - start_row

                    # Copy shape
                    shape.Copy()
                    sheet.Paste()

                    # Get the newly pasted shape (it's the last one)
                    new_shape = sheet.Shapes(sheet.Shapes.Count)

                    # Position the new shape
                    target_cell = sheet.Cells(target_start_row + row_offset, shape.TopLeftCell.Column)
                    new_shape.Top = target_cell.Top + (shape.Top - shape.TopLeftCell.Top)
                    new_shape.Left = target_cell.Left + (shape.Left - shape.TopLeftCell.Left)
        except Exception as e:
            self.logger.warning(f"Error copying shapes: {e}")