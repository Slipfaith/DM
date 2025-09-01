import re
from pathlib import Path
import shutil
from excel_com import ExcelCOM
from excel_processor_v2 import ExcelProcessorV2
from config import Config
from logger import get_logger


class ExcelProcessor:
    def __init__(self, config: Config):
        self.config = config
        self.logger = get_logger()
        self.v2_processor = ExcelProcessorV2(config)
        self._sheet_progress_callback = None

    def set_sheet_progress_callback(self, callback):
        self._sheet_progress_callback = callback

    def process_file(self, filepath: str):
        self.logger.info(f"Starting processing: {filepath}")

        source_path = Path(filepath)
        output_folder = source_path.parent / "Deeva"
        output_folder.mkdir(exist_ok=True)

        output_file = output_folder / source_path.name

        if not self.config.dry_run:
            self.logger.info(f"Copying file to: {output_file}")
            shutil.copy2(filepath, output_file)

            with ExcelCOM() as excel:
                wb = excel.open_workbook(str(output_file))
                total_sheets = wb.Sheets.Count

                for sheet_index, sheet in enumerate(wb.Sheets, 1):
                    if hasattr(self, '_pause_stop_checker') and self._pause_stop_checker:
                        if not self._pause_stop_checker():
                            wb.Close(False)
                            raise Exception("Processing stopped by user")

                    sheet_name = sheet.Name
                    self.logger.info(f"Processing sheet {sheet_index}/{total_sheets}: '{sheet_name}'")

                    if self.v2_processor.can_process(sheet):
                        self.logger.info(f"Using V2 method for sheet '{sheet_name}'")

                        def v2_progress_callback(processed, total):
                            progress_msg = f"Sheet '{sheet_name}': processing group {processed}/{total}"
                            self.logger.info(progress_msg)

                        self.v2_processor.set_progress_callback(v2_progress_callback)
                        self.v2_processor.process_sheet(sheet)
                    else:
                        self.logger.info(f"Using V1 method for sheet '{sheet_name}'")
                        self._process_sheet_v1(sheet, source_path.name)

                    self.logger.info(f"Sheet '{sheet_name}' - Done.")

                    if self._sheet_progress_callback:
                        self._sheet_progress_callback(sheet_index, total_sheets)

                self.logger.info(f"Saving file...")
                wb.Save()
                wb.Close(False)

            self.logger.info(f"Successfully saved to: {output_file}")
        else:
            self.logger.info(f"[DRY RUN] Would save to: {output_file}")

    def _process_sheet_v1(self, sheet, filename):
        sheet_name = sheet.Name
        self.logger.info(f"Excel {filename} - Sheet '{sheet_name}' - searching for header...")

        header_range = self._find_header(sheet)
        if header_range:
            self.logger.info(
                f"Excel {filename} - Sheet '{sheet_name}' - found header at row {header_range.Row}, duplicating rows...")
            self._restructure_sheet(sheet, header_range)
        else:
            self.logger.warning(
                f"Excel {filename} - Sheet '{sheet_name}' - no header found, skipping.")

    def _find_header(self, sheet):
        used_range = sheet.UsedRange
        self.logger.debug(f"Sheet UsedRange: {used_range.Address}")

        if self.config.header_color:
            self.logger.debug(f"Searching for header with color: {self.config.header_color}")
            for row in range(1, min(20, used_range.Rows.Count + 1)):
                if sheet.Rows(row).Hidden:
                    continue
                for col in range(1, used_range.Columns.Count + 1):
                    cell = sheet.Cells(row, col)
                    cell_color = cell.Interior.Color
                    if cell_color == self.config.header_color:
                        self.logger.debug(f"Found colored cell at Row:{row}, Col:{col}, Color:{cell_color}")
                        return self._find_header_range(sheet, row)

        return None

    def _find_header_range(self, sheet, header_row):
        used_range = sheet.UsedRange

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
        if sheet.Rows(row).Hidden:
            return False
        for col in range(start_col, end_col + 1):
            cell = sheet.Cells(row, col)
            if cell.HasFormula:
                return True
        return False

    def _has_data_in_range(self, sheet, row, start_col, end_col):
        if sheet.Rows(row).Hidden:
            return False
        for col in range(start_col, end_col + 1):
            if sheet.Cells(row, col).Value:
                return True
        return False

    def _restructure_sheet(self, sheet, header_range):
        header_row = header_range.Row
        header_start_col = header_range.Column
        header_end_col = header_range.Column + header_range.Columns.Count - 1

        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1

        header_height = sheet.Rows(header_row).RowHeight

        data_blocks = []
        row = header_row + 1

        while row <= last_row:
            if sheet.Rows(row).Hidden:
                row += 1
                continue

            if self._has_data_in_range(sheet, row, header_start_col, header_end_col):
                block = {'start_row': row, 'end_row': row}

                next_row = row + 1
                if next_row <= last_row and not sheet.Rows(next_row).Hidden and \
                        self._has_formulas(sheet, next_row, header_start_col, header_end_col):
                    block['end_row'] = next_row
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

        sheet.Application.ScreenUpdating = False
        sheet.Application.Calculation = -4135

        for i in range(len(data_blocks) - 1, -1, -1):
            block = data_blocks[i]

            source_range = sheet.Range(f"{block['start_row']}:{block['end_row']}")
            source_range.Copy()

            insert_row = block['end_row'] + 1
            sheet.Rows(insert_row).Insert(Shift=-4121)

            for row_offset in range(block['end_row'] - block['start_row'] + 1):
                source_row = block['start_row'] + row_offset
                dup_row = insert_row + row_offset

                for col in range(header_start_col, header_end_col + 1):
                    source_cell = sheet.Cells(source_row, col)
                    dup_cell = sheet.Cells(dup_row, col)

                    if source_cell.HasFormula:
                        formula = source_cell.Formula
                        if "LEN(" in formula.upper() or "ДЛСТР(" in formula.upper():
                            col_letter = sheet.Cells(1, col).Address.split("$")[1]
                            above_row = dup_row - 1
                            formula = re.sub(
                                r'(LEN|ДЛСТР)\s*\([^)]+\)',
                                rf'\1({col_letter}{above_row})',
                                formula,
                                flags=re.IGNORECASE
                            )
                        dup_cell.Formula = formula

            self._copy_shapes_in_range(sheet, block['start_row'], block['end_row'], insert_row)

            empty_row = block['end_row'] + (block['end_row'] - block['start_row'] + 1) + 1
            sheet.Rows(empty_row).Insert(Shift=-4121)
            sheet.Rows(empty_row).Clear()
            sheet.Rows(empty_row).RowHeight = 15

            if i > 0:
                header_insert_row = block['start_row']
                sheet.Rows(header_insert_row).Insert(Shift=-4121)
                sheet.Rows(header_insert_row).Clear()

                header_range.Copy()
                target_range = sheet.Range(
                    sheet.Cells(header_insert_row, header_start_col),
                    sheet.Cells(header_insert_row, header_end_col)
                )
                target_range.PasteSpecial(-4104)

                sheet.Rows(header_insert_row).RowHeight = header_height

                self._copy_shapes_in_range(sheet, header_row, header_row, header_insert_row)

        sheet.Application.CutCopyMode = False
        sheet.Application.Calculation = -4105
        sheet.Application.ScreenUpdating = True

        self.logger.info("Restructuring completed")

    def _copy_shapes_in_range(self, sheet, start_row, end_row, target_start_row):
        try:
            target_end_row = target_start_row + (end_row - start_row)

            # If shapes already exist in the target range, skip copying to avoid duplicates
            for idx in range(1, sheet.Shapes.Count + 1):
                shape = sheet.Shapes(idx)
                shape_row = shape.TopLeftCell.Row
                if target_start_row <= shape_row <= target_end_row:
                    return

            shapes_count = sheet.Shapes.Count
            existing_positions = set()

            for idx in range(1, shapes_count + 1):
                shape = sheet.Shapes(idx)
                shape_row = shape.TopLeftCell.Row
                if start_row <= shape_row <= end_row and not sheet.Rows(shape_row).Hidden:
                    row_offset = shape_row - start_row
                    target_cell = sheet.Cells(target_start_row + row_offset, shape.TopLeftCell.Column)
                    new_top = target_cell.Top + (shape.Top - shape.TopLeftCell.Top)
                    new_left = target_cell.Left + (shape.Left - shape.TopLeftCell.Left)
                    pos_key = (round(new_left, 2), round(new_top, 2))
                    if pos_key in existing_positions:
                        continue
                    shape.Copy()
                    sheet.Paste()
                    new_shape = sheet.Shapes(sheet.Shapes.Count)
                    new_shape.Top = new_top
                    new_shape.Left = new_left
                    existing_positions.add(pos_key)
        except Exception as e:
            self.logger.warning(f"Error copying shapes: {e}")
