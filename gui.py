# gui.py
import os
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QListWidget, QTextEdit, QLabel,
                               QLineEdit, QSpinBox, QProgressBar, QGroupBox,
                               QGridLayout, QCheckBox, QFileDialog)
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from excel_processor import ExcelProcessor
from config import Config
from logger import setup_logger


class DragDropListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            files = [u.toLocalFile() for u in event.mimeData().urls()]
            excel_files = [f for f in files if f.endswith(('.xlsx', '.xlsm', '.xls'))]

            for file in excel_files:
                items = [self.item(i).text() for i in range(self.count())]
                if file not in items:
                    self.addItem(file)
        else:
            event.ignore()


class ProcessorThread(QThread):
    progress = Signal(int)
    log_message = Signal(str)
    finished = Signal(dict)
    sheet_progress = Signal(int, int)  # current sheet, total sheets

    def __init__(self, files, config):
        super().__init__()
        self.files = files
        self.config = config

    def run(self):
        results = {"success": 0, "failed": 0}
        total_files = len(self.files)

        for i, file in enumerate(self.files):
            try:
                self.log_message.emit(f"Processing: {file}")

                # Create processor with custom logging
                processor = ExcelProcessor(self.config)

                # Hook into processor's logger to emit messages
                import logging

                class GuiLogHandler(logging.Handler):
                    def __init__(self, thread):
                        super().__init__()
                        self.thread = thread

                    def emit(self, record):
                        msg = self.format(record)
                        self.thread.log_message.emit(msg)

                gui_handler = GuiLogHandler(self)
                gui_handler.setFormatter(logging.Formatter('%(message)s'))
                processor.logger.addHandler(gui_handler)

                processor.process_file(file)
                results["success"] += 1

            except Exception as e:
                self.log_message.emit(f"Error in {file}: {str(e)}")
                results["failed"] += 1

            self.progress.emit(int((i + 1) / total_files * 100))

        self.finished.emit(results)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.logger = setup_logger()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Excel Processor")
        self.setFixedSize(600, 450)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Drag & Drop area
        self.file_list = DragDropListWidget()
        self.file_list.setMaximumHeight(150)
        layout.addWidget(QLabel("Drag Excel files here:"))
        layout.addWidget(self.file_list)

        # Settings
        settings_group = QGroupBox("Settings")
        settings_layout = QGridLayout()

        settings_layout.addWidget(QLabel("Header Color:"), 0, 0)
        self.color_input = QSpinBox()
        self.color_input.setRange(0, 16777215)
        self.color_input.setValue(65535)
        settings_layout.addWidget(self.color_input, 0, 1)

        self.dry_run_check = QCheckBox("Dry Run")
        settings_layout.addWidget(self.dry_run_check, 1, 0)

        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        # Buttons
        btn_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process")
        self.process_btn.clicked.connect(self.process_files)
        self.clear_btn = QPushButton("Clear Files")
        self.clear_btn.clicked.connect(self.file_list.clear)
        btn_layout.addWidget(self.process_btn)
        btn_layout.addWidget(self.clear_btn)
        layout.addLayout(btn_layout)

        # Progress
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Log
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        layout.addWidget(QLabel("Log:"))
        layout.addWidget(self.log_text)

    def process_files(self):
        files = [self.file_list.item(i).text()
                 for i in range(self.file_list.count())]

        if not files:
            self.log_text.append("No files to process")
            return

        self.config.header_color = self.color_input.value()
        self.config.dry_run = self.dry_run_check.isChecked()

        self.process_btn.setEnabled(False)
        self.progress_bar.setValue(0)

        self.thread = ProcessorThread(files, self.config)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log_message.connect(self.log_text.append)
        self.thread.finished.connect(self.on_process_finished)
        self.thread.start()

    @Slot(dict)
    def on_process_finished(self, results):
        self.process_btn.setEnabled(True)
        self.log_text.append(f"\nCompleted: {results['success']} success, {results['failed']} failed")