# validator.py
from logger import get_logger


class ExcelValidator:
    def __init__(self):
        self.logger = get_logger()
        self.errors = []
        self.warnings = []

    def validate_result(self, sheet, header_range, original_data_blocks, expected_structure):
        """Validate that the sheet was transformed correctly"""
        self.errors = []
        self.warnings = []

        self.logger.info("Starting validation...")

        # 1. Check header column range
        if not self._validate_column_range(sheet, header_range):
            return False

        # 2. Check expected structure
        if not self._validate_structure(sheet, header_range, expected_structure):
            return False

        # 3. Check formatting preservation
        if not self._validate_formatting(sheet, header_range, expected_structure):
            return False

        # 4. Check shapes/images
        if not self._validate_shapes(sheet, expected_structure):
            return False

        # 5. Check formulas
        if not self._validate_formulas(sheet, expected_structure):
            return False

        if self.errors:
            self.logger.error(f"Validation failed with {len(self.errors)} errors:")
            for error in self.errors:
                self.logger.error(f"  - {error}")
            return False

        if self.warnings:
            self.logger.warning(f"Validation completed with {len(self.warnings)} warnings:")
            for warning in self.warnings:
                self.logger.warning(f"  - {warning}")

        self.logger.info("Validation passed successfully!")
        return True

    def _validate_column_range(self, sheet, header_range):
        """Check that no data extends beyond original column range"""
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        used_range = sheet.UsedRange
        last_used_col = used_range.Column + used_range.Columns.Count - 1

        if last_used_col > header_end_col:
            self.errors.append(
                f"Data extends beyond original range. Expected max column: {header_end_col}, but found: {last_used_col}")

            # Find which rows have data beyond the range
            for row in range(1, used_range.Rows.Count + 1):
                for col in range(header_end_col + 1, last_used_col + 1):
                    if sheet.Cells(row, col).Value:
                        self.errors.append(
                            f"  Row {row}, Column {col} has unexpected data: '{sheet.Cells(row, col).Value}'")

            return False

        self.logger.debug(f"Column range validation passed. Data within columns {header_start_col}-{header_end_col}")
        return True

    def _validate_structure(self, sheet, header_range, expected_structure):
        """Validate the row structure matches expectations"""
        current_row = header_range.Row
        structure_index = 0

        self.logger.debug(f"Validating structure. Expected {len(expected_structure)} rows")

        for i, expected in enumerate(expected_structure):
            if current_row > sheet.UsedRange.Rows.Count:
                self.errors.append(
                    f"Structure incomplete. Expected {len(expected_structure)} rows but sheet ends at row {sheet.UsedRange.Rows.Count}")
                return False

            row_type = expected['type']

            if row_type == 'header':
                if not self._is_header_row(sheet, current_row, header_range):
                    self.errors.append(f"Row {current_row}: Expected header but found different content")
                    return False
                else:
                    self.logger.debug(f"Row {current_row}: Header validated")

            elif row_type == 'data':
                if not self._has_data(sheet, current_row, header_range):
                    self.errors.append(f"Row {current_row}: Expected data but found empty row")
                    return False
                else:
                    self.logger.debug(f"Row {current_row}: Data row validated")

            elif row_type == 'formula':
                if not self._has_formulas_or_values(sheet, current_row, header_range):
                    self.warnings.append(f"Row {current_row}: Expected formula row but found empty")
                else:
                    self.logger.debug(f"Row {current_row}: Formula row validated")

            elif row_type == 'empty':
                if self._has_data(sheet, current_row, header_range):
                    self.warnings.append(f"Row {current_row}: Expected empty but found data")
                else:
                    self.logger.debug(f"Row {current_row}: Empty row validated")

            current_row += 1

        return True

    def _validate_formatting(self, sheet, header_range, expected_structure):
        """Check that formatting is preserved in duplicated rows"""
        self.logger.debug("Validating formatting preservation...")

        # Group rows by their source
        data_rows = []
        duplicate_rows = []

        current_row = header_range.Row
        for expected in expected_structure:
            if expected['type'] == 'data' and 'source_block' in expected:
                if expected.get('is_duplicate'):
                    duplicate_rows.append((current_row, expected['source_block']))
                else:
                    data_rows.append((current_row, expected['source_block']))
            current_row += 1

        # Compare formatting between original and duplicate
        for (dup_row, block_index), (orig_row, _) in zip(duplicate_rows, data_rows):
            if not self._compare_row_formatting(sheet, orig_row, dup_row, header_range):
                self.errors.append(f"Formatting mismatch between row {orig_row} and duplicate row {dup_row}")
                return False

        self.logger.debug("Formatting validation passed")
        return True

    def _validate_shapes(self, sheet, expected_structure):
        """Validate that shapes are properly duplicated"""
        self.logger.debug("Validating shapes/images...")

        # Count shapes per row
        shapes_by_row = {}
        try:
            for shape in sheet.Shapes:
                row = self._get_shape_row(sheet, shape)
                if row not in shapes_by_row:
                    shapes_by_row[row] = []
                shapes_by_row[row].append(shape)
        except:
            self.logger.debug("No shapes found in sheet")
            return True

        # Check that duplicated rows have same number of shapes
        current_row = 1
        for expected in expected_structure:
            if expected['type'] == 'data' and 'has_shapes' in expected and expected['has_shapes']:
                orig_count = len(shapes_by_row.get(current_row, []))

                # Find corresponding duplicate row
                dup_row = current_row
                for j, exp in enumerate(expected_structure[expected_structure.index(expected):]):
                    if exp.get('is_duplicate') and exp.get('source_block') == expected.get('source_block'):
                        dup_row = current_row + j
                        break

                dup_count = len(shapes_by_row.get(dup_row, []))

                if orig_count != dup_count:
                    self.warnings.append(
                        f"Shape count mismatch: Row {current_row} has {orig_count} shapes, duplicate row {dup_row} has {dup_count}")

            current_row += 1

        return True

    def _validate_formulas(self, sheet, expected_structure):
        """Validate formula preservation and adjustment"""
        self.logger.debug("Validating formulas...")

        current_row = 1
        for expected in expected_structure:
            if expected['type'] == 'formula':
                # Check that formula row has at least one formula
                has_formula = False
                for col in range(1, sheet.UsedRange.Columns.Count + 1):
                    if sheet.Cells(current_row, col).HasFormula:
                        has_formula = True
                        formula = sheet.Cells(current_row, col).Formula
                        self.logger.debug(f"Row {current_row}, Col {col}: Formula found: {formula}")
                        break

                if not has_formula and expected.get('must_have_formula'):
                    self.warnings.append(f"Row {current_row}: Expected formula but none found")

            current_row += 1

        return True

    def _is_header_row(self, sheet, row, header_range):
        """Check if row matches header pattern"""
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        # Check first few cells to see if they match header values
        for col in range(header_start_col, min(header_start_col + 3, header_end_col + 1)):
            orig_value = header_range.Cells(1, col - header_start_col + 1).Value
            curr_value = sheet.Cells(row, col).Value
            if orig_value != curr_value:
                return False

        return True

    def _has_data(self, sheet, row, header_range):
        """Check if row has any data in the header column range"""
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        for col in range(header_start_col, header_end_col + 1):
            if sheet.Cells(row, col).Value:
                return True
        return False

    def _has_formulas_or_values(self, sheet, row, header_range):
        """Check if row has formulas or values"""
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        for col in range(header_start_col, header_end_col + 1):
            cell = sheet.Cells(row, col)
            if cell.HasFormula or cell.Value:
                return True
        return False

    def _compare_row_formatting(self, sheet, row1, row2, header_range):
        """Compare formatting between two rows"""
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        for col in range(header_start_col, header_end_col + 1):
            cell1 = sheet.Cells(row1, col)
            cell2 = sheet.Cells(row2, col)

            # Compare key formatting properties
            if (cell1.Interior.Color != cell2.Interior.Color or
                    cell1.Font.Color != cell2.Font.Color or
                    cell1.Font.Bold != cell2.Font.Bold or
                    cell1.Font.Size != cell2.Font.Size):
                self.logger.debug(f"Format mismatch at column {col}: "
                                  f"Colors: {cell1.Interior.Color} vs {cell2.Interior.Color}, "
                                  f"Font colors: {cell1.Font.Color} vs {cell2.Font.Color}")
                return False

        return True

    def _get_shape_row(self, sheet, shape):
        """Determine which row a shape belongs to"""
        shape_top = shape.Top

        for row in range(1, sheet.UsedRange.Rows.Count + 1):
            row_top = sheet.Cells(row, 1).Top
            row_bottom = sheet.Cells(row + 1, 1).Top

            if row_top <= shape_top < row_bottom:
                return row

        return -1

    def build_expected_structure(self, header_row, data_blocks):
        """Build the expected structure based on input data"""
        structure = []

        # Original header
        structure.append({'type': 'header', 'row_num': header_row})

        for i, block in enumerate(data_blocks):
            # Header for subsequent blocks
            if i > 0:
                structure.append({'type': 'header'})

            # Original data row
            structure.append({
                'type': 'data',
                'source_block': i,
                'is_duplicate': False,
                'has_shapes': block.get('has_shapes', False)
            })

            # Formula row if exists
            if 'formula_row' in block:
                structure.append({
                    'type': 'formula',
                    'source_block': i,
                    'must_have_formula': True
                })

            # Duplicate data row
            structure.append({
                'type': 'data',
                'source_block': i,
                'is_duplicate': True,
                'has_shapes': block.get('has_shapes', False)
            })

            # Duplicate formula row if exists
            if 'formula_row' in block:
                structure.append({
                    'type': 'formula',
                    'source_block': i,
                    'is_duplicate': True,
                    'must_have_formula': True
                })

            # Empty row
            structure.append({'type': 'empty'})

        return structure