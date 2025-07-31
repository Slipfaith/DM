# feedback_dialog.py
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit,
                               QLineEdit, QPushButton, QLabel, QCheckBox,
                               QProgressBar, QMessageBox, QGroupBox, QTabWidget,
                               QWidget, QSpacerItem, QSizePolicy)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QIcon, QFont
from telegram_reporter import get_telegram_reporter, send_user_feedback
from translations import tr
import traceback


class FeedbackSendThread(QThread):
    """Поток для отправки обратной связи"""
    finished = Signal(bool, str)  # success, message

    def __init__(self, feedback_text, contact_info, include_logs, is_error_report=False, error=None):
        super().__init__()
        self.feedback_text = feedback_text
        self.contact_info = contact_info
        self.include_logs = include_logs
        self.is_error_report = is_error_report
        self.error = error

    def run(self):
        try:
            reporter = get_telegram_reporter()

            if self.is_error_report and self.error:
                success = reporter.report_error(self.error, f"User reported error: {self.feedback_text}")
            else:
                success = reporter.send_feedback(self.feedback_text, self.contact_info, self.include_logs)

            if success:
                self.finished.emit(True, tr('feedback_sent_success'))
            else:
                self.finished.emit(False, tr('feedback_sent_failed'))

        except Exception as e:
            self.finished.emit(False, f"{tr('feedback_sent_error')}: {str(e)}")


class ErrorReportDialog(QDialog):
    """Диалог для отправки отчета об ошибке"""

    def __init__(self, parent=None, error=None, error_context=""):
        super().__init__(parent)
        self.error = error
        self.error_context = error_context
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(tr('error_report_title'))
        self.setModal(True)
        self.setFixedSize(500, 400)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Заголовок
        title_label = QLabel(tr('error_report_header'))
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        # Описание ошибки (если есть)
        if self.error:
            error_group = QGroupBox(tr('error_details'))
            error_layout = QVBoxLayout(error_group)

            error_text = QTextEdit()
            error_text.setReadOnly(True)
            error_text.setMaximumHeight(100)
            error_text.setPlainText(f"{type(self.error).__name__}: {str(self.error)}")
            error_layout.addWidget(error_text)

            layout.addWidget(error_group)

        # Поле для описания проблемы
        description_label = QLabel(tr('error_description_label'))
        layout.addWidget(description_label)

        self.description_text = QTextEdit()
        self.description_text.setPlaceholderText(tr('error_description_placeholder'))
        self.description_text.setMaximumHeight(120)
        layout.addWidget(self.description_text)

        # Контактная информация
        contact_label = QLabel(tr('contact_info_label'))
        layout.addWidget(contact_label)

        self.contact_edit = QLineEdit()
        self.contact_edit.setPlaceholderText(tr('contact_info_placeholder'))
        layout.addWidget(self.contact_edit)

        # Чекбокс для включения логов
        self.include_logs_cb = QCheckBox(tr('include_logs_checkbox'))
        self.include_logs_cb.setChecked(True)
        layout.addWidget(self.include_logs_cb)

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.send_button = QPushButton(tr('send_report'))
        self.send_button.clicked.connect(self.send_report)

        self.cancel_button = QPushButton(tr('cancel'))
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.send_button)

        layout.addLayout(button_layout)

    def send_report(self):
        description = self.description_text.toPlainText().strip()
        if not description:
            QMessageBox.warning(self, tr('warning'), tr('error_description_required'))
            return

        # Показываем прогресс
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Неопределенный прогресс
        self.send_button.setEnabled(False)
        self.cancel_button.setEnabled(False)

        # Запускаем отправку в отдельном потоке
        self.send_thread = FeedbackSendThread(
            description,
            self.contact_edit.text().strip(),
            self.include_logs_cb.isChecked(),
            is_error_report=True,
            error=self.error
        )
        self.send_thread.finished.connect(self.on_send_finished)
        self.send_thread.start()

    def on_send_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.send_button.setEnabled(True)
        self.cancel_button.setEnabled(True)

        if success:
            QMessageBox.information(self, tr('success'), message)
            self.accept()
        else:
            QMessageBox.warning(self, tr('error'), message)


class FeedbackDialog(QDialog):
    """Диалог для отправки обратной связи"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(tr('feedback_title'))
        self.setModal(True)
        self.setFixedSize(600, 500)

        layout = QVBoxLayout(self)

        # Табы
        tab_widget = QTabWidget()

        # Таб обратной связи
        feedback_tab = QWidget()
        self.setup_feedback_tab(feedback_tab)
        tab_widget.addTab(feedback_tab, tr('feedback_tab'))

        # Таб связи с разработчиком
        contact_tab = QWidget()
        self.setup_contact_tab(contact_tab)
        tab_widget.addTab(contact_tab, tr('contact_tab'))

        layout.addWidget(tab_widget)

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.test_button = QPushButton(tr('test_connection'))
        self.test_button.clicked.connect(self.test_connection)

        self.send_button = QPushButton(tr('send_feedback'))
        self.send_button.clicked.connect(self.send_feedback)

        self.cancel_button = QPushButton(tr('cancel'))
        self.cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(self.test_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.send_button)

        layout.addLayout(button_layout)

    def setup_feedback_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(15)

        # Заголовок
        title_label = QLabel(tr('feedback_header'))
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        # Тип обратной связи
        feedback_type_group = QGroupBox(tr('feedback_type'))
        feedback_type_layout = QVBoxLayout(feedback_type_group)

        self.bug_report_rb = QCheckBox(tr('bug_report'))
        self.feature_request_rb = QCheckBox(tr('feature_request'))
        self.general_feedback_rb = QCheckBox(tr('general_feedback'))

        feedback_type_layout.addWidget(self.bug_report_rb)
        feedback_type_layout.addWidget(self.feature_request_rb)
        feedback_type_layout.addWidget(self.general_feedback_rb)

        layout.addWidget(feedback_type_group)

        # Текст обратной связи
        feedback_label = QLabel(tr('feedback_text_label'))
        layout.addWidget(feedback_label)

        self.feedback_text = QTextEdit()
        self.feedback_text.setPlaceholderText(tr('feedback_text_placeholder'))
        layout.addWidget(self.feedback_text)

        # Контактная информация
        contact_label = QLabel(tr('contact_info_label'))
        layout.addWidget(contact_label)

        self.contact_edit = QLineEdit()
        self.contact_edit.setPlaceholderText(tr('contact_info_placeholder'))
        layout.addWidget(self.contact_edit)

        # Включить логи
        self.include_logs_cb = QCheckBox(tr('include_logs_checkbox'))
        layout.addWidget(self.include_logs_cb)

    def setup_contact_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(20)

        # Информация о разработчике
        dev_info = QLabel(tr('developer_info'))
        dev_info.setWordWrap(True)
        layout.addWidget(dev_info)

        # Быстрые сообщения
        quick_group = QGroupBox(tr('quick_messages'))
        quick_layout = QVBoxLayout(quick_group)

        self.question_btn = QPushButton(tr('quick_question'))
        self.question_btn.clicked.connect(lambda: self.set_quick_message(tr('quick_question_text')))

        self.bug_btn = QPushButton(tr('quick_bug_report'))
        self.bug_btn.clicked.connect(lambda: self.set_quick_message(tr('quick_bug_text')))

        self.thanks_btn = QPushButton(tr('quick_thanks'))
        self.thanks_btn.clicked.connect(lambda: self.set_quick_message(tr('quick_thanks_text')))

        quick_layout.addWidget(self.question_btn)
        quick_layout.addWidget(self.bug_btn)
        quick_layout.addWidget(self.thanks_btn)

        layout.addWidget(quick_group)

        layout.addStretch()

    def set_quick_message(self, message):
        """Установка быстрого сообщения"""
        self.feedback_text.setPlainText(message)

    def test_connection(self):
        """Тестирование соединения с Telegram"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.test_button.setEnabled(False)

        def test_finished():
            self.progress_bar.setVisible(False)
            self.test_button.setEnabled(True)

            try:
                reporter = get_telegram_reporter()
                success = reporter.test_connection()

                if success:
                    QMessageBox.information(self, tr('success'), tr('connection_test_success'))
                else:
                    QMessageBox.warning(self, tr('error'), tr('connection_test_failed'))
            except Exception as e:
                QMessageBox.critical(self, tr('error'), f"{tr('connection_test_error')}: {str(e)}")

        # Используем таймер для имитации асинхронности
        QTimer.singleShot(100, test_finished)

    def send_feedback(self):
        feedback_text = self.feedback_text.toPlainText().strip()
        if not feedback_text:
            QMessageBox.warning(self, tr('warning'), tr('feedback_text_required'))
            return

        # Добавляем тип обратной связи к тексту
        feedback_types = []
        if self.bug_report_rb.isChecked():
            feedback_types.append(tr('bug_report'))
        if self.feature_request_rb.isChecked():
            feedback_types.append(tr('feature_request'))
        if self.general_feedback_rb.isChecked():
            feedback_types.append(tr('general_feedback'))

        if feedback_types:
            feedback_text = f"[{', '.join(feedback_types)}]\n\n{feedback_text}"

        # Показываем прогресс
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.send_button.setEnabled(False)
        self.cancel_button.setEnabled(False)

        # Запускаем отправку в отдельном потоке
        self.send_thread = FeedbackSendThread(
            feedback_text,
            self.contact_edit.text().strip(),
            self.include_logs_cb.isChecked()
        )
        self.send_thread.finished.connect(self.on_send_finished)
        self.send_thread.start()

    def on_send_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.send_button.setEnabled(True)
        self.cancel_button.setEnabled(True)

        if success:
            QMessageBox.information(self, tr('success'), message)
            self.accept()
        else:
            QMessageBox.warning(self, tr('error'), message)


def show_error_report_dialog(parent=None, error=None, error_context=""):
    """Показать диалог отчета об ошибке"""
    dialog = ErrorReportDialog(parent, error, error_context)
    return dialog.exec()


def show_feedback_dialog(parent=None):
    """Показать диалог обратной связи"""
    dialog = FeedbackDialog(parent)
    return dialog.exec()