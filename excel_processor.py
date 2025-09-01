import subprocess
import shutil
from pathlib import Path
from config import Config
from logger import get_logger


class ExcelProcessor:
    """Handle Excel files via external VBScript."""

    def __init__(self, config: Config):
        self.config = config
        self.logger = get_logger()
        self._sheet_progress_callback = None
        self._pause_stop_checker = None

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

            if self._pause_stop_checker and not self._pause_stop_checker():
                raise Exception("Processing stopped by user")

            script_path = Path(__file__).with_name("excel_processor.vbs")
            args = [
                "cscript",
                "//NoLogo",
                str(script_path),
                str(output_file),
                str(self.config.header_color),
            ]

            try:
                subprocess.run(args, check=True)
                self.logger.info(f"Successfully saved to: {output_file}")
            except subprocess.CalledProcessError as e:
                self.logger.error(f"VBScript processing failed: {e}")
                raise
        else:
            self.logger.info(f"[DRY RUN] Would save to: {output_file}")
