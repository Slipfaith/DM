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

        if self.config.header_color:
            for row in range(1, min(20, used_range.Rows.Count + 1)):
                if sheet.Rows(row).Hidden:
                    continue
                for col in range(1, used_range.Columns.Count + 1):
                    cell = sheet.Cells(row, col)
                    cell_color = cell.Interior.Color
                    if cell_color == self.config.header_color:
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
                last_col = col
        if first_col and last_col:
            return sheet.Range(sheet.Cells(header_row, first_col), sheet.Cells(header_row, last_col))
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

    def _map_shapes_by_row(self, sheet, min_row, max_row):
        row_to_shapes = {}
        for idx in range(1, sheet.Shapes.Count + 1):
            shape = sheet.Shapes(idx)
            shape_row = shape.TopLeftCell.Row
            if min_row <= shape_row <= max_row:
                row_to_shapes.setdefault(shape_row, []).append(shape)
        return row_to_shapes

    def _copy_shapes_for_row(self, sheet, shapes, target_row):
        for shape in shapes:
            col = shape.TopLeftCell.Column
            delta_top = shape.Top - shape.TopLeftCell.Top
            delta_left = shape.Left - shape.TopLeftCell.Left
            target_cell = sheet.Cells(target_row, col)
            new_shape = shape.Duplicate()
            new_shape.Top = target_cell.Top + delta_top
            new_shape.Left = target_cell.Left + delta_left

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

        sheet.Application.ScreenUpdating = False
        sheet.Application.Calculation = -4135

        # Привязка картинок к строкам
        shapes_map = self._map_shapes_by_row(sheet, header_row, last_row)

        for i in range(len(data_blocks) - 1, -1, -1):
            block = data_blocks[i]
            start_row, end_row = block['start_row'], block['end_row']
            source_range = sheet.Range(f"{start_row}:{end_row}")
            source_range.Copy()
            insert_row = end_row + 1
            sheet.Rows(insert_row).Insert(Shift=-4121)
            for row_offset in range(end_row - start_row + 1):
                src_row = start_row + row_offset
                dst_row = insert_row + row_offset
                for col in range(header_start_col, header_end_col + 1):
                    source_cell = sheet.Cells(src_row, col)
                    dup_cell = sheet.Cells(dst_row, col)
                    if source_cell.HasFormula:
                        formula = source_cell.Formula
                        if "LEN(" in formula.upper() or "ДЛСТР(" in formula.upper():
                            col_letter = sheet.Cells(1, col).Address.split("$")[1]
                            above_row = dst_row - 1
                            formula = re.sub(
                                r'(LEN|ДЛСТР)\s*\([^)]+\)',
                                rf'\1({col_letter}{above_row})',
                                formula,
                                flags=re.IGNORECASE
                            )
                        dup_cell.Formula = formula
                # копируем все картинки, которые были над src_row
                if src_row in shapes_map:
                    self._copy_shapes_for_row(sheet, shapes_map[src_row], dst_row)

            empty_row = end_row + (end_row - start_row + 1) + 1
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
                if header_row in shapes_map:
                    self._copy_shapes_for_row(sheet, shapes_map[header_row], header_insert_row)

        sheet.Application.CutCopyMode = False
        sheet.Application.Calculation = -4105
        sheet.Application.ScreenUpdating = True