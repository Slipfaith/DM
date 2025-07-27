# excel_processor.py
import os
from pathlib import Path
import shutil
from excel_com import ExcelCOM
from config import Config
from logger import get_logger
from validator import ExcelValidator
from utils import generate_unique_filename, parse_excel_address


class ExcelProcessor:
    def __init__(self, config: Config):
        self.config = config
        self.logger = get_logger()

    def process_file(self, filepath: str):
        self.logger.info(f"Starting processing: {filepath}")

        # Create output folder next to source file
        source_path = Path(filepath)
        output_folder = source_path.parent / "Deeva"
        output_folder.mkdir(exist_ok=True)

        # Output file with same name
        output_file = output_folder / source_path.name

        if not self.config.dry_run:
            # First copy the entire file
            self.logger.info(f"Copying file to: {output_file}")
            shutil.copy2(filepath, output_file)

            # Then process it
            with ExcelCOM() as excel:
                wb = excel.open_workbook(str(output_file))
                total_sheets = wb.Sheets.Count

                for sheet_index, sheet in enumerate(wb.Sheets, 1):
                    sheet_name = sheet.Name
                    self.logger.info(
                        f"Excel {source_path.name} - Sheet {sheet_index}/{total_sheets} '{sheet_name}' - searching for header...")

                    header_range = self._find_header(sheet)
                    if header_range:
                        self.logger.info(
                            f"Excel {source_path.name} - Sheet '{sheet_name}' - found header at row {header_range.Row}, duplicating rows...")
                        self._restructure_sheet(sheet, header_range)
                        self.logger.info(f"Excel {source_path.name} - Sheet '{sheet_name}' - Done.")
                    else:
                        self.logger.warning(
                            f"Excel {source_path.name} - Sheet '{sheet_name}' - no header found, skipping.")

                self.logger.info(f"Saving file...")
                wb.Save()
                wb.Close(False)

            self.logger.info(f"Successfully saved to: {output_file}")
        else:
            self.logger.info(f"[DRY RUN] Would save to: {output_file}")

    def _find_header(self, sheet):
        used_range = sheet.UsedRange
        self.logger.debug(f"Sheet UsedRange: {used_range.Address}")

        # Only search by color
        if self.config.header_color:
            self.logger.debug(f"Searching for header with color: {self.config.header_color}")
            for row in range(1, min(20, used_range.Rows.Count + 1)):
                for col in range(1, used_range.Columns.Count + 1):
                    cell = sheet.Cells(row, col)
                    cell_color = cell.Interior.Color
                    if cell_color == self.config.header_color:
                        self.logger.debug(f"Found colored cell at Row:{row}, Col:{col}, Color:{cell_color}")
                        # Found colored cell, now find the actual range of the header
                        return self._find_header_range(sheet, row)

        return None

        return None

    def _find_header_range(self, sheet, header_row):
        """Find the actual start and end columns of the header"""
        used_range = sheet.UsedRange

        # Find first non-empty cell in the row
        first_col = None
        last_col = None

        for col in range(1, used_range.Columns.Count + 1):
            cell_value = sheet.Cells(header_row, col).Value
            if cell_value is not None:
                if first_col is None:
                    first_col = col
                    self.logger.debug(f"Header starts at column {first_col}, value: '{cell_value}'")
                last_col = col

        if first_col and last_col:
            header_range = sheet.Range(sheet.Cells(header_row, first_col), sheet.Cells(header_row, last_col))
            self.logger.debug(f"Header range: {header_range.Address} (Row:{header_row}, Cols:{first_col}-{last_col})")
            return header_range

        self.logger.warning(f"No header range found in row {header_row}")
        return None

    def _has_formulas(self, sheet, row, start_col, end_col):
        """Check if row contains any formulas"""
        for col in range(start_col, end_col + 1):
            cell = sheet.Cells(row, col)
            if cell.HasFormula:
                return True
        return False

    def _get_shapes_in_row(self, sheet, row):
        """Get all shapes (images, etc.) that are in the specified row"""
        shapes_in_row = []
        try:
            for shape in sheet.Shapes:
                # Check if shape's top is within the row
                cell_top = sheet.Cells(row, 1).Top
                cell_bottom = sheet.Cells(row + 1, 1).Top

                if cell_top <= shape.Top < cell_bottom:
                    shapes_in_row.append({
                        'shape': shape,
                        'left': shape.Left,
                        'top': shape.Top,
                        'width': shape.Width,
                        'height': shape.Height,
                        'name': shape.Name
                    })
        except Exception as e:
            self.logger.warning(f"Could not process shapes: {e}")

        return shapes_in_row

    def _copy_shapes_for_row(self, all_shapes, sheet, source_row, target_row):
        """Copy all shapes from source row to target row"""
        try:
            source_top = sheet.Cells(source_row, 1).Top
            target_top = sheet.Cells(target_row, 1).Top
            offset = target_top - source_top

            shapes_copied = 0
            for shape_info in all_shapes:
                if shape_info['row'] == source_row:
                    # Find the original shape by name
                    for shape in sheet.Shapes:
                        if shape.Name == shape_info['name']:
                            # Copy the shape
                            shape.Copy()
                            sheet.Paste()

                            # Get the newly pasted shape (it's the last one)
                            new_shape = sheet.Shapes(sheet.Shapes.Count)

                            # Position it in the target row
                            new_shape.Top = shape_info['top'] + offset
                            new_shape.Left = shape_info['left']

                            shapes_copied += 1
                            self.logger.debug(
                                f"Copied shape '{shape_info['name']}' from row {source_row} to row {target_row}")
                            break

            if shapes_copied > 0:
                self.logger.info(f"Copied {shapes_copied} shapes to row {target_row}")

        except Exception as e:
            self.logger.warning(f"Could not copy shapes from row {source_row} to {target_row}: {e}")

    def _get_shape_row(self, sheet, shape):
        """Determine which row a shape belongs to based on its top position"""
        shape_top = shape.Top

        # Find the row that contains this shape
        for row in range(1, sheet.UsedRange.Rows.Count + 2):
            try:
                row_top = sheet.Cells(row, 1).Top
                if row + 1 <= sheet.UsedRange.Rows.Count + 1:
                    row_bottom = sheet.Cells(row + 1, 1).Top
                else:
                    row_bottom = row_top + sheet.Rows(row).Height

                if row_top <= shape_top < row_bottom:
                    return row
            except:
                continue

        return -1

    def _restructure_sheet(self, sheet, header_range):
        header_row = header_range.Row
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1

        # Find data blocks (data row + optional formula row)
        data_blocks = []
        row = header_row + 1

        while row <= last_row:
            # Check if current row has data in the header columns range
            row_has_data = False
            data_cells = []
            for col in range(header_start_col, header_end_col + 1):
                cell_value = sheet.Cells(row, col).Value
                if cell_value:
                    row_has_data = True
                    data_cells.append(f"Col{col}:'{cell_value}'")

            if row_has_data:
                self.logger.debug(f"Found data row {row}: {', '.join(data_cells[:3])}...")
                block = {'data_row': row}

                # Check if next row has formulas
                if row + 1 <= last_row and self._has_formulas(sheet, row + 1, header_start_col, header_end_col):
                    self.logger.debug(f"Row {row + 1} contains formulas")
                    block['formula_row'] = row + 1
                    row += 2
                else:
                    row += 1

                data_blocks.append(block)
            else:
                row += 1

        if not data_blocks:
            self.logger.info(f"No data rows found after header")
            return

        self.logger.info(f"Found {len(data_blocks)} data blocks to process")

        # Store all data and shapes before clearing
        stored_blocks = []
        all_shapes = []

        # First, store ALL shapes info before any modifications
        try:
            self.logger.debug(f"Total shapes in sheet: {sheet.Shapes.Count}")
            for shape in sheet.Shapes:
                shape_info = {
                    'left': shape.Left,
                    'top': shape.Top,
                    'width': shape.Width,
                    'height': shape.Height,
                    'name': shape.Name,
                    'type': shape.Type,
                    'row': self._get_shape_row(sheet, shape)
                }
                all_shapes.append(shape_info)
                self.logger.debug(f"Stored shape '{shape.Name}' at row {shape_info['row']}")
        except Exception as e:
            self.logger.warning(f"Error storing shapes: {e}")

        # Store column widths
        column_widths = {}
        for col in range(header_start_col, header_end_col + 1):
            column_widths[col] = sheet.Columns(col).ColumnWidth

        # Store header row height
        header_row_height = sheet.Rows(header_row).RowHeight

        for block in data_blocks:
            stored_data = {'data': [], 'formulas': None, 'shapes': [], 'formula_shapes': [],
                           'data_row_height': sheet.Rows(block['data_row']).RowHeight,
                           'original_row': block['data_row']}

            # Store data row
            for col in range(header_start_col, header_end_col + 1):
                cell = sheet.Cells(block['data_row'], col)
                stored_data['data'].append({
                    'value': cell.Value,
                    'formula': cell.Formula if cell.HasFormula else None,
                    'format': self._get_cell_format(cell)
                })

            # Note which shapes belong to this row (we'll copy them later)
            stored_data['shapes'] = [s for s in all_shapes if s['row'] == block['data_row']]

            # Store formula row if exists
            if 'formula_row' in block:
                stored_data['formula_row_height'] = sheet.Rows(block['formula_row']).RowHeight
                stored_data['formulas'] = []
                for col in range(header_start_col, header_end_col + 1):
                    cell = sheet.Cells(block['formula_row'], col)
                    stored_data['formulas'].append({
                        'value': cell.Value,
                        'formula': cell.Formula if cell.HasFormula else None,
                        'format': self._get_cell_format(cell)
                    })

                # Note which shapes belong to formula row
                stored_data['formula_shapes'] = [s for s in all_shapes if s['row'] == block['formula_row']]

            stored_blocks.append(stored_data)

        # DO NOT DELETE SHAPES - we'll move them instead
        # Clear only cell data after header
        if last_row > header_row:
            sheet.Rows(f"{header_row + 1}:{last_row + 100}").Delete()

        # Build new structure
        current_row = header_row + 1

        # Apply column widths
        for col, width in column_widths.items():
            sheet.Columns(col).ColumnWidth = width

        for i, stored_block in enumerate(stored_blocks):
            self.logger.info(f"Processing data block {i + 1}/{len(stored_blocks)}")

            # Skip header for first block (already exists)
            if i > 0:
                self.logger.debug(f"Copying header to row {current_row}")
                # Copy header - IMPORTANT: paste in the same column range
                header_range.Copy()
                dest_range = sheet.Range(sheet.Cells(current_row, header_start_col),
                                         sheet.Cells(current_row, header_end_col))
                dest_range.PasteSpecial(-4104)  # xlPasteAll
                sheet.Rows(current_row).RowHeight = header_row_height
                self.logger.debug(f"Header copied to {dest_range.Address}, height set to {header_row_height}")
                current_row += 1

            # Original data row
            data_row_num = current_row
            original_data_row_in_blocks = stored_block.get('original_row', data_blocks[i]['data_row'])
            self.logger.debug(f"Writing data row to row {current_row}")
            sheet.Rows(current_row).RowHeight = stored_block['data_row_height']
            for j, cell_data in enumerate(stored_block['data']):
                cell = sheet.Cells(current_row, header_start_col + j)
                if cell_data['formula']:
                    cell.Formula = cell_data['formula']
                    self.logger.debug(f"Cell {cell.Address}: Formula='{cell_data['formula']}'")
                else:
                    cell.Value = cell_data['value']
                    self.logger.debug(f"Cell {cell.Address}: Value='{cell_data['value']}'")
                self._apply_cell_format(cell, cell_data['format'])

            # Copy shapes to data row
            self._copy_shapes_for_row(all_shapes, sheet, original_data_row_in_blocks, current_row)
            current_row += 1

            # Original formula row if exists
            if stored_block['formulas']:
                original_formula_row = data_blocks[i].get('formula_row')
                sheet.Rows(current_row).RowHeight = stored_block.get('formula_row_height', 15)
                for j, cell_data in enumerate(stored_block['formulas']):
                    cell = sheet.Cells(current_row, header_start_col + j)
                    if cell_data['formula']:
                        cell.Formula = cell_data['formula']
                    else:
                        cell.Value = cell_data['value']
                    self._apply_cell_format(cell, cell_data['format'])

                # Copy shapes to formula row
                if original_formula_row:
                    self._copy_shapes_for_row(all_shapes, sheet, original_formula_row, current_row)
                current_row += 1

            # Duplicate data row
            duplicate_data_row = current_row
            sheet.Rows(current_row).RowHeight = stored_block['data_row_height']
            for j, cell_data in enumerate(stored_block['data']):
                cell = sheet.Cells(current_row, header_start_col + j)
                if cell_data['formula']:
                    cell.Formula = cell_data['formula']
                else:
                    cell.Value = cell_data['value']
                self._apply_cell_format(cell, cell_data['format'])

            # Copy shapes to duplicate data row
            self._copy_shapes_for_row(all_shapes, sheet, original_data_row_in_blocks, current_row)
            current_row += 1

            # Duplicate formula row if exists
            if stored_block['formulas']:
                sheet.Rows(current_row).RowHeight = stored_block.get('formula_row_height', 15)
                for j, cell_data in enumerate(stored_block['formulas']):
                    cell = sheet.Cells(current_row, header_start_col + j)
                    if cell_data['formula']:
                        cell.Formula = cell_data['formula']
                    else:
                        cell.Value = cell_data['value']
                    self._apply_cell_format(cell, cell_data['format'])

                # Copy shapes to duplicate formula row
                if original_formula_row:
                    self._copy_shapes_for_row(all_shapes, sheet, original_formula_row, current_row)
                current_row += 1

            # Empty row
            current_row += 1

        # Clear clipboard
        sheet.Application.CutCopyMode = False

        # Validate the result
        self.logger.info("Validating transformation result...")
        validator = ExcelValidator()

        # Build expected structure for validation
        validation_blocks = []
        for block in data_blocks:
            val_block = {'has_shapes': len(self._get_shapes_in_row(sheet, block['data_row'])) > 0}
            if 'formula_row' in block:
                val_block['formula_row'] = block['formula_row']
            validation_blocks.append(val_block)

        expected_structure = validator.build_expected_structure(header_row, validation_blocks)

        if validator.validate_result(sheet, header_range, data_blocks, expected_structure):
            self.logger.info("✓ Validation passed - transformation completed successfully")
        else:
            self.logger.error("✗ Validation failed - please check the log for details")
            # Continue anyway, user can check the result

    def _get_cell_format(self, cell):
        """Store cell formatting"""
        return {
            'interior_color': cell.Interior.Color,
            'font_color': cell.Font.Color,
            'font_bold': cell.Font.Bold,
            'font_size': cell.Font.Size,
            'font_name': cell.Font.Name,
            'number_format': cell.NumberFormat,
            'horizontal_alignment': cell.HorizontalAlignment,
            'vertical_alignment': cell.VerticalAlignment,
            'wrap_text': cell.WrapText,
            'borders': self._get_borders(cell)
        }

    def _get_borders(self, cell):
        """Store cell borders"""
        borders = {}
        for edge in [7, 8, 9, 10]:  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
            try:
                border = cell.Borders(edge)
                borders[edge] = {
                    'line_style': border.LineStyle,
                    'weight': border.Weight,
                    'color': border.Color
                }
            except:
                pass
        return borders

    def _apply_cell_format(self, cell, format_dict):
        """Apply stored formatting to cell"""
        try:
            cell.Interior.Color = format_dict['interior_color']
            cell.Font.Color = format_dict['font_color']
            cell.Font.Bold = format_dict['font_bold']
            cell.Font.Size = format_dict['font_size']
            cell.Font.Name = format_dict['font_name']
            cell.NumberFormat = format_dict['number_format']
            cell.HorizontalAlignment = format_dict['horizontal_alignment']
            cell.VerticalAlignment = format_dict['vertical_alignment']
            cell.WrapText = format_dict['wrap_text']

            self.logger.debug(f"Applied format to {cell.Address}: BgColor={format_dict['interior_color']}, "
                              f"FontColor={format_dict['font_color']}, Bold={format_dict['font_bold']}")

            # Apply borders
            for edge, border_format in format_dict.get('borders', {}).items():
                try:
                    border = cell.Borders(edge)
                    if border_format['line_style'] is not None:
                        border.LineStyle = border_format['line_style']
                        border.Weight = border_format['weight']
                        border.Color = border_format['color']
                except:
                    pass
        except Exception as e:
            self.logger.error(f"Failed to apply formatting to {cell.Address}: {e}")
            self.logger.debug(f"Format dict: {format_dict}")