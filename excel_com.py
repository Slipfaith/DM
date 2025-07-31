# excel_com.py
import win32com.client
import pythoncom

from logger import get_logger


class ExcelCOM:
    def __init__(self):
        self.app = None
        self.logger = get_logger()
        self._original_state = {}

    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch("Excel.Application")

        # Save original state
        self._original_state = {
            "Visible": self.app.Visible,
            "ScreenUpdating": self.app.ScreenUpdating,
            "DisplayAlerts": self.app.DisplayAlerts,
            "EnableEvents": self.app.EnableEvents
        }

        # Try to save Calculation state
        try:
            self._original_state["Calculation"] = self.app.Calculation
        except:
            pass

        # Optimize for processing
        self.app.Visible = False
        self.app.ScreenUpdating = False
        self.app.DisplayAlerts = False
        self.app.EnableEvents = False

        # Try to set manual calculation
        try:
            self.app.Calculation = -4135  # xlCalculationManual
        except:
            pass

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            # Restore original state
            for prop, value in self._original_state.items():
                try:
                    setattr(self.app, prop, value)
                except:
                    pass

            # Close all workbooks
            for wb in self.app.Workbooks:
                wb.Close(False)

            self.app.Quit()
        except Exception as e:
            self.logger.error(f"Error closing Excel: {e}")
        finally:
            pythoncom.CoUninitialize()

    def open_workbook(self, filepath):
        return self.app.Workbooks.Open(filepath)