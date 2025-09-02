import subprocess
from pathlib import Path
import shutil
from config import Config
from logger import get_logger


class ExcelProcessor:
    def __init__(self, config: Config):
        self.config = config
        self.logger = get_logger()
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
            script_path = Path(__file__).with_name("excel_processor.vbs")
            cmd = [
                "cscript",
                "//nologo",
                str(script_path),
                str(output_file),
                str(self.config.header_color),
            ]
            self.logger.info("Running VBScript processor")
            try:
                subprocess.run(cmd, check=True)
            except subprocess.CalledProcessError as exc:
                self.logger.error(f"VBScript failed: {exc}")
                raise
            finally:
                if self._sheet_progress_callback:
                    self._sheet_progress_callback(1, 1)
            self.logger.info(f"Successfully saved to: {output_file}")
        else:
            self.logger.info(f"[DRY RUN] Would save to: {output_file}")

