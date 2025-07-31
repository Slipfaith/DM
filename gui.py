# gui.py
import os
import webbrowser
import traceback
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QListWidget, QTextEdit, QLabel,
                               QSpinBox, QProgressBar, QFileDialog, QFrame,
                               QListWidgetItem, QGroupBox, QMenuBar, QMenu,
                               QMessageBox)
from PySide6.QtCore import Qt, QThread, Signal, Slot, QUrl, QTimer
from PySide6.QtGui import (QDragEnterEvent, QDropEvent, QAction, QActionGroup,
                           QDesktopServices, QIcon)
from excel_processor import ExcelProcessor
from config import Config
from logger import setup_logger
from styles import MAIN_STYLE, ICON_PATH
from updater import UpdateChecker, CURRENT_VERSION
from translations import tr, set_language
from error_dialog import ErrorReportDialog, FeedbackDialog
from settings_manager import settings_manager


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

        self.text_label = QLabel(tr('drag_drop_text'))
        self.text_label.setObjectName("dragDropText")
        self.text_label.setAlignment(Qt.AlignCenter)

        layout.addWidget(self.text_label)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                tr('select_excel_files'),
                "",
                tr('excel_files_filter')
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
        self.is_paused = False
        self.should_stop = False
        self.current_file_index = 0
        self._pause_lock = False
        self._last_error = None
        self._last_traceback = None

    def pause(self):
        self.is_paused = True
        self._pause_lock = True

    def resume(self):
        self.is_paused = False
        self._pause_lock = False

    def stop(self):
        self.should_stop = True
        self.is_paused = False
        self._pause_lock = False

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

    def check_pause_stop(self):
        if self.should_stop:
            return False

        while self.is_paused:
            self.msleep(100)
            if self.should_stop:
                return False

        return True

    def run(self):
        results = {"success": 0, "failed": 0, "output_folder": None}

        self.total_sheets = self.count_sheets()
        self.sheet_progress.emit(0, self.total_sheets)

        for i in range(self.current_file_index, len(self.files)):
            if not self.check_pause_stop():
                break

            file = self.files[i]
            self.current_file_index = i

            try:
                self.file_processing.emit(Path(file).name)

                # Create a custom processor that checks for pause/stop
                from excel_processor import ExcelProcessor
                processor = ExcelProcessor(self.config)

                # Inject pause/stop checker
                processor._pause_stop_checker = self.check_pause_stop

                import logging
                class GuiLogHandler(logging.Handler):
                    def __init__(self, thread):
                        super().__init__()
                        self.thread = thread

                    def emit(self, record):
                        msg = self.format(record)
                        self.thread.log_message.emit(msg)

                        # Check for sheet completion
                        if "Sheet" in msg and "Done." in msg:
                            self.thread.processed_sheets += 1
                            progress = int((self.thread.processed_sheets / self.thread.total_sheets) * 100)
                            self.thread.progress.emit(progress)
                            self.thread.sheet_progress.emit(self.thread.processed_sheets, self.thread.total_sheets)

                        # Check pause/stop after each sheet
                        if "searching for header" in msg or "Done." in msg:
                            if not self.thread.check_pause_stop():
                                raise Exception("Processing stopped by user")

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
                if "stopped by user" in str(e):
                    break
                else:
                    self.log_message.emit(f"Error in {Path(file).name}: {str(e)}")
                    results["failed"] += 1

                    if "stopped by user" not in str(e):
                        self._last_error = str(e)
                        self._last_traceback = traceback.format_exc()

        self.finished.emit(results)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.logger = setup_logger()
        self.files = []
        self.updater = UpdateChecker(self)
        self.init_ui()

        # Load saved language
        saved_lang = settings_manager.get_language()
        self.set_language(saved_lang)

    def init_ui(self):
        self.setWindowTitle(tr('app_title'))
        self.setFixedSize(400, 300)

        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.setStyleSheet(MAIN_STYLE)

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

        self.loaded_group = QGroupBox(tr('loaded_files'))
        self.loaded_group.setObjectName("fileGroup")
        loaded_layout = QVBoxLayout(self.loaded_group)
        self.loaded_list = FileListWidget()
        loaded_layout.addWidget(self.loaded_list)

        self.processed_group = QGroupBox(tr('processed_files'))
        self.processed_group.setObjectName("fileGroup")
        processed_layout = QVBoxLayout(self.processed_group)
        self.processed_list = FileListWidget()
        self.processed_list.setEnabled(False)
        processed_layout.addWidget(self.processed_list)

        lists_layout.addWidget(self.loaded_group)
        lists_layout.addWidget(self.processed_group)

        content_layout.addWidget(lists_container)

        self.status_label = QLabel(tr('ready'))
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

        buttons_layout.addStretch()

        self.clear_btn = QPushButton(tr('clear_files'))
        self.clear_btn.setObjectName("clearButton")
        self.clear_btn.clicked.connect(self.clear_files)

        self.process_btn = QPushButton(tr('process_files'))
        self.process_btn.setObjectName("processButton")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)

        self.pause_btn = QPushButton("⏸")
        self.pause_btn.setObjectName("pauseButton")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.hide()

        self.stop_btn = QPushButton("⏹")
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.hide()

        buttons_layout.addWidget(self.pause_btn)
        buttons_layout.addWidget(self.stop_btn)
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

        # Create the menu after all widgets so translation strings can
        # reference them without hitting missing attributes.
        self.create_menu()

        # Check for updates on startup
        QTimer.singleShot(1000, lambda: self.updater.check_for_updates(silent=True))

    def create_menu(self):
        menubar = self.menuBar()

        self.file_menu = menubar.addMenu('')

        self.clear_action = QAction('', self)
        self.clear_action.triggered.connect(self.clear_files)
        self.file_menu.addAction(self.clear_action)

        self.file_menu.addSeparator()

        self.exit_action = QAction('', self)
        self.exit_action.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_action)

        self.help_menu = menubar.addMenu('')

        self.update_action = QAction('', self)
        self.update_action.triggered.connect(self.check_updates)
        self.help_menu.addAction(self.update_action)

        self.about_action = QAction('', self)
        self.about_action.triggered.connect(self.show_about)
        self.help_menu.addAction(self.about_action)

        self.feedback_action = QAction('', self)
        self.feedback_action.triggered.connect(self.show_feedback_dialog)
        self.help_menu.addAction(self.feedback_action)

        # Language menu
        self.language_menu = menubar.addMenu('')
        lang_group = QActionGroup(self)
        self.lang_en = QAction(tr('lang_en'), self, checkable=True)
        self.lang_ru = QAction(tr('lang_ru'), self, checkable=True)
        lang_group.addAction(self.lang_en)
        lang_group.addAction(self.lang_ru)
        self.lang_en.triggered.connect(lambda: self.set_language('en'))
        self.lang_ru.triggered.connect(lambda: self.set_language('ru'))
        self.language_menu.addAction(self.lang_en)
        self.language_menu.addAction(self.lang_ru)

        self.apply_translations()

    def set_language(self, lang):
        set_language(lang)
        settings_manager.set_language(lang)
        # keep actions checked
        self.lang_en.setChecked(lang == 'en')
        self.lang_ru.setChecked(lang == 'ru')
        self.apply_translations()

    def apply_translations(self):
        self.setWindowTitle(tr('app_title'))
        self.file_menu.setTitle(tr('menu_file'))
        self.clear_action.setText(tr('menu_clear_all'))
        self.exit_action.setText(tr('menu_exit'))
        self.help_menu.setTitle(tr('menu_help'))
        self.update_action.setText(tr('menu_check_updates'))
        self.about_action.setText(tr('menu_about'))
        self.feedback_action.setText(tr('menu_contact_developer'))
        self.language_menu.setTitle(tr('menu_language'))
        self.lang_en.setText(tr('lang_en'))
        self.lang_ru.setText(tr('lang_ru'))

        self.drop_area.text_label.setText(tr('drag_drop_text'))
        self.loaded_group.setTitle(tr('loaded_files'))
        self.processed_group.setTitle(tr('processed_files'))
        self.clear_btn.setText(tr('clear_files'))
        self.process_btn.setText(tr('process_files'))
        if not self.files:
            self.status_label.setText(tr('ready'))

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
            self.status_label.setText(tr('files_loaded', count=len(self.files)))

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
        self.status_label.setText(tr('ready'))

    def process_files(self):
        if not self.files:
            return

        self.config.header_color = 65535  # Always use yellow

        self.process_btn.hide()
        self.clear_btn.hide()
        self.pause_btn.show()
        self.stop_btn.show()
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

    def toggle_pause(self):
        if self.thread and self.thread.isRunning():
            if self.thread.is_paused:
                self.thread.resume()
                self.pause_btn.setText("⏸")
                self.status_label.setText(tr('processing_resumed'))
                self.log_text.append(">>> " + tr('processing_resumed'))
            else:
                self.thread.pause()
                self.pause_btn.setText("▶")
                self.status_label.setText(tr('stopping'))
                self.log_text.append(">>> " + tr('stopping'))

    def stop_processing(self):
        if self.thread and self.thread.isRunning():
            reply = QMessageBox.question(
                self,
                tr('stop_processing_title'),
                tr('stop_processing_confirm'),
                QMessageBox.Yes | QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.thread.stop()
                self.status_label.setText(tr('processing_paused'))
                self.log_text.append(">>> " + tr('processing_paused'))

    def on_file_processing(self, filename):
        self.status_label.setText(tr('processing', filename=filename))
        item = QListWidgetItem(filename)
        self.processed_list.addItem(item)

    def on_sheet_progress(self, processed, total):
        self.progress_label.setText(tr('sheets_progress', processed=processed, total=total))

    @Slot(dict)
    def on_process_finished(self, results):
        self.process_btn.show()
        self.clear_btn.show()
        self.pause_btn.hide()
        self.stop_btn.hide()
        self.progress_container.hide()

        total = results['success'] + results['failed']
        summary = tr('summary')
        summary += tr('total_files', total=total)
        summary += tr('success', count=results['success'])
        summary += tr('failed', count=results['failed'])

        if results['output_folder']:
            summary += tr('output_folder', folder=results['output_folder'])

        self.summary_label.setText(summary)
        self.summary_label.show()

        if hasattr(self.thread, 'should_stop') and self.thread.should_stop:
            self.status_label.setText(tr('processing_stopped'))
            self.log_text.append(">>> " + tr('processing_stopped'))
        else:
            self.status_label.setText(
                tr('completed', success=results['success'], failed=results['failed'])
            )
            self.log_text.append(">>> " + tr('completed', success=results['success'], failed=results['failed']))

            if results['failed'] > 0 and hasattr(self.thread, '_last_error'):
                reply = QMessageBox.question(
                    self,
                    tr('error_occurred'),
                    tr('send_error_report_prompt'),
                    QMessageBox.Yes | QMessageBox.No
                )

                if reply == QMessageBox.Yes:
                    error_msg = self.thread._last_error
                    if hasattr(self.thread, '_last_traceback'):
                        error_msg = self.thread._last_traceback
                    dialog = ErrorReportDialog(self, error_msg)
                    dialog.exec()

    def open_folder(self, link):
        QDesktopServices.openUrl(QUrl.fromLocalFile(link))

    def check_updates(self):
        self.updater.check_for_updates()

    def show_about(self):
        QMessageBox.about(
            self,
            tr('about_title'),
            tr('about_text', version=CURRENT_VERSION)
        )

    def show_feedback_dialog(self):
        dialog = FeedbackDialog(self)
        dialog.exec()