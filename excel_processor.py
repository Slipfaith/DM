# excel_processor.py

import re
from pathlib import Path
import shutil
from excel_com import ExcelCOM
from config import Config
from logger import get_logger


class ExcelProcessor:
    def __init__(self, config: Config):
        self.config = config
        self.logger = get_logger()

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
                    # Check for pause/stop before each sheet
                    if hasattr(self, '_pause_stop_checker') and self._pause_stop_checker:
                        if not self._pause_stop_checker():
                            wb.Close(False)
                            raise Exception("Processing stopped by user")

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

        if self.config.header_color:
            self.logger.debug(f"Searching for header with color: {self.config.header_color}")
            for row in range(1, min(20, used_range.Rows.Count + 1)):
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
        for col in range(start_col, end_col + 1):
            cell = sheet.Cells(row, col)
            if cell.HasFormula:
                return True
        return False

    def _has_data_in_range(self, sheet, row, start_col, end_col):
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

        data_blocks = []
        row = header_row + 1

        while row <= last_row:
            if self._has_data_in_range(sheet, row, header_start_col, header_end_col):
                block = {'start_row': row, 'end_row': row}

                if row + 1 <= last_row and self._has_formulas(sheet, row + 1, header_start_col, header_end_col):
                    block['end_row'] = row + 1
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
        sheet.Application.Calculation = -4135  # xlCalculationManual

        # Process from end to beginning
        for i in range(len(data_blocks) - 1, -1, -1):
            block = data_blocks[i]

            # Copy and insert duplicate
            source_range = sheet.Range(f"{block['start_row']}:{block['end_row']}")
            source_range.Copy()

            insert_row = block['end_row'] + 1
            sheet.Rows(insert_row).Insert(Shift=-4121)

            # Fix formulas in duplicated rows
            for row_offset in range(block['end_row'] - block['start_row'] + 1):
                source_row = block['start_row'] + row_offset
                dup_row = insert_row + row_offset

                for col in range(header_start_col, header_end_col + 1):
                    source_cell = sheet.Cells(source_row, col)
                    dup_cell = sheet.Cells(dup_row, col)

                    if source_cell.HasFormula:
                        formula = source_cell.Formula
                        # Fix LEN/ДЛСТР formulas to reference cell above
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

            # Add empty row after duplicate
            empty_row = block['end_row'] + (block['end_row'] - block['start_row'] + 1) + 1
            sheet.Rows(empty_row).Insert(Shift=-4121)
            sheet.Rows(empty_row).Clear()
            sheet.Rows(empty_row).RowHeight = 15

            # Add header before block (except for first block)
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

        sheet.Application.CutCopyMode = False
        sheet.Application.Calculation = -4105  # xlCalculationAutomatic
        sheet.Application.ScreenUpdating = True

        self.logger.info("Restructuring completed")