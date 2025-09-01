from logger import get_logger


class ExcelProcessorV2:
    """Compatibility wrapper.

    The original Python implementation has been replaced by a VBScript
    counterpart (excel_processor.vbs) which is invoked from
    :mod:`excel_processor`. This class remains to satisfy imports but no
    longer provides any processing logic.
    """

    def __init__(self, config):
        self.config = config
        self.logger = get_logger()
        self._progress_callback = None

    def set_progress_callback(self, callback):
        self._progress_callback = callback

    def can_process(self, sheet):
        return True

    def process_sheet(self, sheet):
        # Processing is handled entirely by the VBScript.
        pass
