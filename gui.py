# gui.py
import os
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QListWidget, QTextEdit, QLabel,
                               QSpinBox, QProgressBar, QFileDialog, QFrame,
                               QCheckBox, QListWidgetItem, QAbstractItemView)
from PySide6.QtCore import Qt, QThread, Signal, Slot, QPropertyAnimation, QRect
from PySide6.QtGui import (QDragEnterEvent, QDropEvent, QPalette, QColor,
                           QFont, QIcon, QPainter, QBrush, QPen)
from excel_processor import ExcelProcessor
from config import Config
from logger import setup_logger


class DragDropArea(QFrame):
    filesDropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #3498db;
                border-radius: 10px;
                background-color: #f8f9fa;
            }
            QFrame:hover {
                background-color: #e3f2fd;
                border-color: #2196f3;
            }
        """)
        self.setMinimumHeight(120)
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        self.icon_label = QLabel("📁")
        self.icon_label.setStyleSheet("font-size: 32px;")
        self.icon_label.setAlignment(Qt.AlignCenter)

        self.text_label = QLabel("Drop Excel files here or click to browse")
        self.text_label.setStyleSheet("color: #666; font-size: 14px;")
        self.text_label.setAlignment(Qt.AlignCenter)

        self.count_label = QLabel("")
        self.count_label.setStyleSheet("color: #2196f3; font-size: 12px; font-weight: bold;")
        self.count_label.setAlignment(Qt.AlignCenter)
        self.count_label.hide()

        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.count_label)

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
            self.setStyleSheet("""
                QFrame {
                    border: 2px solid #2196f3;
                    border-radius: 10px;
                    background-color: #e3f2fd;
                }
            """)

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #3498db;
                border-radius: 10px;
                background-color: #f8f9fa;
            }
            QFrame:hover {
                background-color: #e3f2fd;
                border-color: #2196f3;
            }
        """)

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

    def updateFileCount(self, count):
        if count > 0:
            self.count_label.setText(f"{count} file{'s' if count > 1 else ''} selected")
            self.count_label.show()
            self.text_label.setText("Drop more files or click to add")
        else:
            self.count_label.hide()
            self.text_label.setText("Drop Excel files here or click to browse")


class FileListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 5px;
                background-color: white;
                padding: 5px;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:hover {
                background-color: #f5f5f5;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
            }
        """)
        self.setMaximumHeight(150)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)


class ProcessorThread(QThread):
    progress = Signal(int)
    log_message = Signal(str)
    finished = Signal(dict)
    sheet_progress = Signal(int, int)

    def __init__(self, files, config):
        super().__init__()
        self.files = files
        self.config = config

    def run(self):
        results = {"success": 0, "failed": 0}
        total_files = len(self.files)

        for i, file in enumerate(self.files):
            try:
                self.log_message.emit(f"Processing: {Path(file).name}")
                processor = ExcelProcessor(self.config)

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
                self.log_message.emit(f"Error in {Path(file).name}: {str(e)}")
                results["failed"] += 1

            self.progress.emit(int((i + 1) / total_files * 100))

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
        self.setFixedSize(450, 550)

        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #2196f3;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
            QPushButton:pressed {
                background-color: #0d47a1;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
            QLabel {
                color: #333;
            }
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 5px;
                text-align: center;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #4caf50;
                border-radius: 4px;
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("Excel Processor")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1976d2;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        self.drop_area = DragDropArea()
        self.drop_area.filesDropped.connect(self.add_files)
        layout.addWidget(self.drop_area)

        self.file_list = FileListWidget()
        self.file_list.hide()
        layout.addWidget(self.file_list)

        settings_frame = QFrame()
        settings_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        settings_layout = QHBoxLayout(settings_frame)

        settings_layout.addWidget(QLabel("Header Color:"))
        self.color_input = QSpinBox()
        self.color_input.setRange(0, 16777215)
        self.color_input.setValue(65535)
        self.color_input.setStyleSheet("""
            QSpinBox {
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
        """)
        settings_layout.addWidget(self.color_input)

        settings_layout.addStretch()

        self.dry_run_check = QCheckBox("Dry Run")
        self.dry_run_check.setStyleSheet("QCheckBox { color: #666; }")
        settings_layout.addWidget(self.dry_run_check)

        layout.addWidget(settings_frame)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.clear_btn = QPushButton("Clear Files")
        self.clear_btn.clicked.connect(self.clear_files)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        self.clear_btn.hide()

        self.process_btn = QPushButton("Process Files")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.process_btn)
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 5px;
                background-color: white;
                padding: 5px;
                font-family: monospace;
                font-size: 12px;
            }
        """)
        self.log_text.hide()
        layout.addWidget(self.log_text)

        layout.addStretch()

    def add_files(self, new_files):
        for file in new_files:
            if file not in self.files:
                self.files.append(file)
                item = QListWidgetItem(Path(file).name)
                item.setToolTip(file)
                self.file_list.addItem(item)

        self.drop_area.updateFileCount(len(self.files))

        if self.files:
            self.file_list.show()
            self.clear_btn.show()
            self.process_btn.setEnabled(True)
            self.setFixedSize(450, 650)

    def clear_files(self):
        self.files.clear()
        self.file_list.clear()
        self.file_list.hide()
        self.clear_btn.hide()
        self.process_btn.setEnabled(False)
        self.drop_area.updateFileCount(0)
        self.log_text.hide()
        self.setFixedSize(450, 550)

    def process_files(self):
        if not self.files:
            return

        self.config.header_color = self.color_input.value()
        self.config.dry_run = self.dry_run_check.isChecked()

        self.process_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.progress_bar.show()
        self.log_text.show()
        self.log_text.clear()
        self.progress_bar.setValue(0)

        self.thread = ProcessorThread(self.files, self.config)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log_message.connect(self.log_text.append)
        self.thread.finished.connect(self.on_process_finished)
        self.thread.start()

    @Slot(dict)
    def on_process_finished(self, results):
        self.process_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.log_text.append(f"\n✓ Completed: {results['success']} success, {results['failed']} failed")