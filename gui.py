# gui.py
import os
import webbrowser
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QListWidget, QTextEdit, QLabel,
                               QSpinBox, QProgressBar, QFileDialog, QFrame,
                               QListWidgetItem, QGroupBox, QMenuBar, QMenu,
                               QMessageBox)
from PySide6.QtCore import Qt, QThread, Signal, Slot, QUrl
from PySide6.QtGui import (QDragEnterEvent, QDropEvent, QAction, QDesktopServices)
from excel_processor import ExcelProcessor
from config import Config
from logger import setup_logger
from styles import MAIN_STYLE


class DragDropArea(QFrame):
    filesDropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setObjectName("dragDropArea")
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(20)

        self.text_label = QLabel("Drag & Drop Excel Files Here\nor Click to Browse")
        self.text_label.setObjectName("dragDropText")
        self.text_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.text_label)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Select Excel Files",
                "",
                "Excel Files (*.xlsx *.xlsm *.xls)"
            )
            if files:
                self.filesDropped.emit(files)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)

    def dragLeaveEvent(self, event):
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            files = []
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.endswith(('.xlsx', '.xlsm', '.xls')):
                    files.append(file_path)
            if files:
                self.filesDropped.emit(files)
        self.dragLeaveEvent(event)


class FileListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("fileList")


class ProcessorThread(QThread):
    progress = Signal(int)
    log_message = Signal(str)
    finished = Signal(dict)
    file_processing = Signal(str)
    sheet_progress = Signal(int, int)

    def __init__(self, files, config):
        super().__init__()
        self.files = files
        self.config = config
        self.total_sheets = 0
        self.processed_sheets = 0

    def count_sheets(self):
        from excel_com import ExcelCOM
        total = 0
        with ExcelCOM() as excel:
            for file in self.files:
                try:
                    wb = excel.open_workbook(file)
                    total += wb.Sheets.Count
                    wb.Close(False)
                except:
                    pass
        return total

    def run(self):
        results = {"success": 0, "failed": 0, "output_folder": None}

        self.total_sheets = self.count_sheets()
        self.sheet_progress.emit(0, self.total_sheets)

        for file in self.files:
            try:
                self.file_processing.emit(Path(file).name)
                processor = ExcelProcessor(self.config)

                import logging
                class GuiLogHandler(logging.Handler):
                    def __init__(self, thread):
                        super().__init__()
                        self.thread = thread

                    def emit(self, record):
                        msg = self.format(record)
                        self.thread.log_message.emit(msg)
                        if "Sheet" in msg and "Done." in msg:
                            self.thread.processed_sheets += 1
                            progress = int((self.thread.processed_sheets / self.thread.total_sheets) * 100)
                            self.thread.progress.emit(progress)
                            self.thread.sheet_progress.emit(self.thread.processed_sheets, self.thread.total_sheets)

                gui_handler = GuiLogHandler(self)
                gui_handler.setFormatter(logging.Formatter('%(message)s'))
                processor.logger.addHandler(gui_handler)

                processor.process_file(file)
                results["success"] += 1

                if not results["output_folder"]:
                    output_folder = Path(file).parent / "Deeva"
                    if output_folder.exists():
                        results["output_folder"] = str(output_folder)

            except Exception as e:
                self.log_message.emit(f"Error in {Path(file).name}: {str(e)}")
                results["failed"] += 1

        self.finished.emit(results)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.logger = setup_logger()
        self.files = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Excel Processor")
        self.setFixedSize(400, 300)
        self.setStyleSheet(MAIN_STYLE)

        self.create_menu()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.drop_area = DragDropArea()
        self.drop_area.filesDropped.connect(self.add_files)
        main_layout.addWidget(self.drop_area)

        content_widget = QWidget()
        content_widget.setObjectName("contentWidget")
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)

        lists_container = QWidget()
        lists_layout = QHBoxLayout(lists_container)
        lists_layout.setSpacing(20)

        loaded_group = QGroupBox("Loaded Files")
        loaded_group.setObjectName("fileGroup")
        loaded_layout = QVBoxLayout(loaded_group)
        self.loaded_list = FileListWidget()
        loaded_layout.addWidget(self.loaded_list)

        processed_group = QGroupBox("Processed Files")
        processed_group.setObjectName("fileGroup")
        processed_layout = QVBoxLayout(processed_group)
        self.processed_list = FileListWidget()
        self.processed_list.setEnabled(False)
        processed_layout.addWidget(self.processed_list)

        lists_layout.addWidget(loaded_group)
        lists_layout.addWidget(processed_group)

        content_layout.addWidget(lists_container)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        content_layout.addWidget(self.status_label)

        progress_container = QWidget()
        progress_layout = QVBoxLayout(progress_container)
        progress_layout.setSpacing(5)

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        progress_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setObjectName("progressLabel")
        self.progress_label.setAlignment(Qt.AlignCenter)
        progress_layout.addWidget(self.progress_label)

        progress_container.hide()
        self.progress_container = progress_container
        content_layout.addWidget(progress_container)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)

        settings_widget = QWidget()
        settings_layout = QHBoxLayout(settings_widget)
        settings_layout.setContentsMargins(0, 0, 0, 0)

        settings_layout.addWidget(QLabel("Header Color:"))
        self.color_input = QSpinBox()
        self.color_input.setObjectName("colorInput")
        self.color_input.setRange(0, 16777215)
        self.color_input.setValue(65535)
        settings_layout.addWidget(self.color_input)

        buttons_layout.addWidget(settings_widget)
        buttons_layout.addStretch()

        self.clear_btn = QPushButton("Clear Files")
        self.clear_btn.setObjectName("clearButton")
        self.clear_btn.clicked.connect(self.clear_files)

        self.process_btn = QPushButton("Process Files")
        self.process_btn.setObjectName("processButton")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)

        buttons_layout.addWidget(self.clear_btn)
        buttons_layout.addWidget(self.process_btn)

        content_layout.addLayout(buttons_layout)

        self.log_text = QTextEdit()
        self.log_text.setObjectName("logText")
        self.log_text.setReadOnly(True)
        content_layout.addWidget(self.log_text)

        self.summary_label = QLabel("")
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setTextFormat(Qt.RichText)
        self.summary_label.setOpenExternalLinks(False)
        self.summary_label.linkActivated.connect(self.open_folder)
        self.summary_label.hide()
        content_layout.addWidget(self.summary_label)

        content_widget.hide()
        main_layout.addWidget(content_widget)
        self.content_widget = content_widget

    def create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu('File')

        clear_action = QAction('Clear All', self)
        clear_action.triggered.connect(self.clear_files)
        file_menu.addAction(clear_action)

        file_menu.addSeparator()

        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        help_menu = menubar.addMenu('Help')

        update_action = QAction('Check for Updates', self)
        update_action.triggered.connect(self.check_updates)
        help_menu.addAction(update_action)

        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def add_files(self, new_files):
        for file in new_files:
            if file not in self.files:
                self.files.append(file)
                item = QListWidgetItem(Path(file).name)
                item.setToolTip(file)
                self.loaded_list.addItem(item)

        if self.files:
            self.drop_area.hide()
            self.content_widget.show()
            self.setFixedSize(800, 600)
            self.process_btn.setEnabled(True)
            self.status_label.setText(f"{len(self.files)} files loaded")

    def clear_files(self):
        self.files.clear()
        self.loaded_list.clear()
        self.processed_list.clear()
        self.log_text.clear()
        self.summary_label.hide()
        self.progress_label.setText("")
        self.drop_area.show()
        self.content_widget.hide()
        self.setFixedSize(400, 300)
        self.process_btn.setEnabled(False)
        self.status_label.setText("Ready")

    def process_files(self):
        if not self.files:
            return

        self.config.header_color = self.color_input.value()

        self.process_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.progress_container.show()
        self.progress_bar.setValue(0)
        self.processed_list.clear()
        self.summary_label.hide()

        self.thread = ProcessorThread(self.files, self.config)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log_message.connect(self.log_text.append)
        self.thread.finished.connect(self.on_process_finished)
        self.thread.file_processing.connect(self.on_file_processing)
        self.thread.sheet_progress.connect(self.on_sheet_progress)
        self.thread.start()

    def on_file_processing(self, filename):
        self.status_label.setText(f"Processing: {filename}")
        item = QListWidgetItem(filename)
        self.processed_list.addItem(item)

    def on_sheet_progress(self, processed, total):
        self.progress_label.setText(f"Sheets: {processed}/{total}")

    @Slot(dict)
    def on_process_finished(self, results):
        self.process_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.progress_container.hide()

        total = results['success'] + results['failed']
        summary = f"<b>Summary:</b><br>"
        summary += f"Total files: {total}<br>"
        summary += f"✓ Success: {results['success']}<br>"
        summary += f"✗ Failed: {results['failed']}<br>"

        if results['output_folder']:
            summary += f"<br>Output folder: <a href='{results['output_folder']}'>Open Deeva folder</a>"

        self.summary_label.setText(summary)
        self.summary_label.show()

        self.status_label.setText(f"Completed: {results['success']} success, {results['failed']} failed")

    def open_folder(self, link):
        QDesktopServices.openUrl(QUrl.fromLocalFile(link))

    def check_updates(self):
        webbrowser.open("https://github.com/yourusername/excel-processor/releases")

    def show_about(self):
        QMessageBox.about(self, "About", "Excel Processor v1.0\n\nA tool for duplicating Excel rows with headers")