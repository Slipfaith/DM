# error_dialog.py
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QTextEdit, QPushButton, QGroupBox, QCheckBox,
                               QLineEdit, QMessageBox, QFileDialog,
                               QApplication, QScrollArea, QWidget)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData
from PySide6.QtGui import (QDragEnterEvent, QDropEvent, QKeySequence,
                           QShortcut, QImage, QPixmap, QPainter,
                           QBrush, QPen)
from telegram import TelegramReporter
from translations import tr

import tempfile
import os


class ImagePreviewDialog(QDialog):
    """Simple dialog to preview attached images."""

    def __init__(self, image_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr('image_preview'))

        layout = QVBoxLayout(self)
        label = QLabel()
        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            screen = QApplication.primaryScreen()
            if screen:
                screen_size = screen.availableGeometry().size()
                max_width = min(600, screen_size.width())
                max_height = min(600, screen_size.height())
            else:
                max_width = max_height = 600
            scaled = pixmap.scaled(max_width, max_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            label.setPixmap(scaled)
        layout.addWidget(label)



class ImageThumbnail(QLabel):
    removed = Signal(str)

    def __init__(self, image_path):
        super().__init__()
        self.image_path = image_path
        self.setFixedSize(60, 60)
        self.setCursor(Qt.PointingHandCursor)

        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            scaled = pixmap.scaled(60, 60, Qt.KeepAspectRatio, Qt.SmoothTransformation)

            # Create rounded corners
            rounded = QPixmap(60, 60)
            rounded.fill(Qt.transparent)

            painter = QPainter(rounded)
            painter.setRenderHint(QPainter.Antialiasing)
            painter.setBrush(QBrush(scaled))
            painter.setPen(QPen(Qt.NoPen))
            painter.drawRoundedRect(0, 0, 60, 60, 5, 5)

            # Draw X button
            painter.setPen(QPen(Qt.white, 2))
            painter.setBrush(QBrush(Qt.red))
            painter.drawEllipse(45, 0, 15, 15)
            painter.setPen(QPen(Qt.white, 2))
            painter.drawLine(50, 5, 55, 10)
            painter.drawLine(50, 10, 55, 5)
            painter.end()

            self.setPixmap(rounded)

        self.setStyleSheet("""
            QLabel {
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            QLabel:hover {
                border: 1px solid #999;
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Top-right corner is used as remove button
            if 45 <= event.pos().x() <= 60 and 0 <= event.pos().y() <= 15:
                self.removed.emit(self.image_path)
            else:
                preview = ImagePreviewDialog(self.image_path, self)
                preview.exec()


class DragDropTextEdit(QTextEdit):
    imagesDropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def insertFromMimeData(self, source: QMimeData):
        """Handle pasting images directly from the clipboard."""
        # Image data
        if source.hasImage():
            image = source.imageData()
            if isinstance(image, QImage) and not image.isNull():
                temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                image.save(temp_file.name, 'PNG')
                self.imagesDropped.emit([temp_file.name])
                return

        # File URLs
        if source.hasUrls():
            files = []
            for url in source.urls():
                file_path = url.toLocalFile()
                if file_path and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    files.append(file_path)
            if files:
                self.imagesDropped.emit(files)
                return

        # Text with file path
        if source.hasText():
            text = source.text()
            if text.startswith('file:///'):
                file_path = text.replace('file:///', '').replace('/', os.sep)
                if os.path.exists(file_path) and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    self.imagesDropped.emit([file_path])
                    return

        super().insertFromMimeData(source)

    def dragEnterEvent(self, event: QDragEnterEvent):
        mimeData = event.mimeData()
        if mimeData.hasUrls() or mimeData.hasImage() or mimeData.hasHtml():
            # Check if it contains image data
            has_image = False

            if mimeData.hasImage():
                has_image = True
            elif mimeData.hasUrls():
                for url in mimeData.urls():
                    file_path = url.toLocalFile()
                    if file_path and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        has_image = True
                        break
            elif mimeData.hasHtml() and '<img' in mimeData.html():
                has_image = True

            if has_image:
                event.acceptProposedAction()
            else:
                super().dragEnterEvent(event)
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        mimeData = event.mimeData()

        # Handle image data
        if mimeData.hasImage():
            image = mimeData.imageData()
            if isinstance(image, QImage) and not image.isNull():
                temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                image.save(temp_file.name, 'PNG')
                self.imagesDropped.emit([temp_file.name])
                event.acceptProposedAction()
                return

        # Handle file URLs
        if mimeData.hasUrls():
            files = []
            for url in mimeData.urls():
                file_path = url.toLocalFile()
                if file_path and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    files.append(file_path)
            if files:
                self.imagesDropped.emit(files)
                event.acceptProposedAction()
                return

        # Handle HTML with images
        if mimeData.hasHtml():
            html = mimeData.html()
            import re
            img_pattern = r'<img[^>]+src="([^"]+)"'
            matches = re.findall(img_pattern, html)
            files = []
            for match in matches:
                if match.startswith('file:///'):
                    file_path = match.replace('file:///', '').replace('%20', ' ')
                    if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                        files.append(file_path)
            if files:
                self.imagesDropped.emit(files)
                event.acceptProposedAction()
                return

        super().dropEvent(event)


class SendReportThread(QThread):
    finished = Signal(bool, str)

    def __init__(self, reporter, error_message, log_content, user_message, include_log, images):
        super().__init__()
        self.reporter = reporter
        self.error_message = error_message
        self.log_content = log_content if include_log else None
        self.user_message = user_message
        self.images = images

    def run(self):
        success, message = self.reporter.send_error_report(
            self.error_message,
            self.log_content,
            self.user_message,
            self.images
        )
        self.finished.emit(success, message)


class ErrorReportDialog(QDialog):
    def __init__(self, parent, error_message):
        super().__init__(parent)
        self.error_message = error_message
        self.reporter = TelegramReporter()
        self.attached_images = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(tr('error_report_title'))
        self.setFixedSize(500, 400)
        self.setModal(True)

        layout = QVBoxLayout(self)

        error_group = QGroupBox(tr('error_details'))
        error_layout = QVBoxLayout(error_group)

        self.error_text = QTextEdit()
        self.error_text.setPlainText(self.error_message)
        self.error_text.setReadOnly(True)
        self.error_text.setMaximumHeight(100)
        error_layout.addWidget(self.error_text)

        layout.addWidget(error_group)

        message_group = QGroupBox(tr('your_message'))
        message_layout = QVBoxLayout(message_group)

        self.message_text = DragDropTextEdit()
        self.message_text.setPlaceholderText(tr('describe_problem_drag'))
        self.message_text.imagesDropped.connect(self.add_images)
        message_layout.addWidget(self.message_text)

        layout.addWidget(message_group)

        self.include_log_check = QCheckBox(tr('include_log'))
        self.include_log_check.setChecked(True)
        layout.addWidget(self.include_log_check)

        attach_layout = QHBoxLayout()
        self.attach_btn = QPushButton(tr('attach_image'))
        self.attach_btn.clicked.connect(self.attach_image)
        attach_layout.addWidget(self.attach_btn)
        attach_layout.addStretch()
        layout.addLayout(attach_layout)

        # Thumbnails area
        self.thumbnails_scroll = QScrollArea()
        self.thumbnails_scroll.setMaximumHeight(80)
        self.thumbnails_scroll.setWidgetResizable(True)
        self.thumbnails_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.thumbnails_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.thumbnails_scroll.hide()

        self.thumbnails_widget = QWidget()
        self.thumbnails_layout = QHBoxLayout(self.thumbnails_widget)
        self.thumbnails_layout.setAlignment(Qt.AlignLeft)
        self.thumbnails_scroll.setWidget(self.thumbnails_widget)

        layout.addWidget(self.thumbnails_scroll)

        buttons_layout = QHBoxLayout()
        self.send_btn = QPushButton("📤")  # Send icon
        self.send_btn.setObjectName("telegramButton")
        self.send_btn.setFixedSize(32, 32)
        self.send_btn.clicked.connect(self.send_report)
        self.cancel_btn = QPushButton("❌")  # Cancel icon
        self.cancel_btn.setObjectName("telegramButton")
        self.cancel_btn.setFixedSize(32, 32)
        self.cancel_btn.clicked.connect(self.reject)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.send_btn)
        buttons_layout.addWidget(self.cancel_btn)
        layout.addLayout(buttons_layout)
        # Paste shortcut
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.message_text)
        self.paste_shortcut.activated.connect(self.paste_image)

    def send_report(self):
        user_message = self.message_text.toPlainText().strip()

        if not user_message:
            QMessageBox.warning(self, tr('warning'), tr('please_describe_problem'))
            return

        self.send_btn.setEnabled(False)
        self.send_btn.setText("⏳")

        log_content = None
        if self.include_log_check.isChecked():
            log_content = self.reporter.get_latest_log_content()

        self.thread = SendReportThread(
            self.reporter,
            self.error_message,
            log_content,
            user_message,
            self.include_log_check.isChecked(),
            self.attached_images
        )
        self.thread.finished.connect(self.on_send_finished)
        self.thread.start()

    def attach_image(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            tr('select_images'),
            "",
            "Images (*.png *.jpg *.jpeg *.gif *.bmp)"
        )
        self.add_images(files)

    def add_images(self, files):
        for file in files:
            if file not in self.attached_images and len(self.attached_images) < 5:
                self.attached_images.append(file)
                self.add_thumbnail(file)

        if len(self.attached_images) >= 5:
            QMessageBox.information(self, tr('info'), tr('max_attachments'))

        self.update_thumbnails_visibility()

    def add_thumbnail(self, image_path):
        thumbnail = ImageThumbnail(image_path)
        thumbnail.removed.connect(self.remove_image)
        self.thumbnails_layout.addWidget(thumbnail)

    def remove_image(self, image_path):
        if image_path in self.attached_images:
            self.attached_images.remove(image_path)

        for i in range(self.thumbnails_layout.count()):
            widget = self.thumbnails_layout.itemAt(i).widget()
            if isinstance(widget, ImageThumbnail) and widget.image_path == image_path:
                widget.deleteLater()
                break

        self.update_thumbnails_visibility()

    def update_thumbnails_visibility(self):
        self.thumbnails_scroll.setVisible(len(self.attached_images) > 0)

    def paste_image(self):
        clipboard = QApplication.clipboard()

        # First try to get image directly
        image = clipboard.image()
        if not image.isNull():
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            image.save(temp_file.name, 'PNG')
            self.add_images([temp_file.name])
            return

        # Then try pixmap
        pixmap = clipboard.pixmap()
        if not pixmap.isNull():
            temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            pixmap.save(temp_file.name, 'PNG')
            self.add_images([temp_file.name])
            return

        # Then check mimeData
        mimeData = clipboard.mimeData()

        # Image from mimeData
        if mimeData.hasImage():
            image_variant = mimeData.imageData()
            if image_variant:
                # Convert QVariant to QImage
                if isinstance(image_variant, QImage):
                    image = image_variant
                else:
                    image = QImage(image_variant)

                if not image.isNull():
                    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                    image.save(temp_file.name, 'PNG')
                    self.add_images([temp_file.name])
                    return

        # File URLs (copied files from explorer)
        if mimeData.hasUrls():
            files = []
            for url in mimeData.urls():
                file_path = url.toLocalFile()
                if file_path and os.path.exists(file_path) and file_path.lower().endswith(
                        ('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    files.append(file_path)
            if files:
                self.add_images(files)
                return

        # Plain text with file path
        if mimeData.hasText():
            text = mimeData.text()
            if text.startswith('file:///'):
                file_path = text.replace('file:///', '').replace('/', os.sep)
                if os.path.exists(file_path) and file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    self.add_images([file_path])
                    return

    def on_send_finished(self, success, message):
        self.send_btn.setEnabled(True)
        self.send_btn.setText("📤")

        if success:
            QMessageBox.information(self, tr('success'), message)
            self.accept()
        else:
            QMessageBox.warning(self, tr('error'), message)


class FeedbackDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.reporter = TelegramReporter()
        self.attached_images = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(tr('contact_developer'))
        self.setFixedSize(450, 350)
        self.setModal(True)

        layout = QVBoxLayout(self)

        message_group = QGroupBox(tr('your_message'))
        message_layout = QVBoxLayout(message_group)

        self.message_text = DragDropTextEdit()
        self.message_text.setPlaceholderText(tr('feedback_placeholder_drag'))
        self.message_text.imagesDropped.connect(self.add_images)
        message_layout.addWidget(self.message_text)

        layout.addWidget(message_group)

        email_layout = QHBoxLayout()
        email_layout.addWidget(QLabel(tr('email_optional')))
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("your@email.com")
        email_layout.addWidget(self.email_input)
        layout.addLayout(email_layout)

        attach_layout = QHBoxLayout()
        self.attach_btn = QPushButton(tr('attach_image'))
        self.attach_btn.clicked.connect(self.attach_image)
        attach_layout.addWidget(self.attach_btn)
        attach_layout.addStretch()
        layout.addLayout(attach_layout)

        # Thumbnails area
        self.thumbnails_scroll = QScrollArea()
        self.thumbnails_scroll.setMaximumHeight(80)
        self.thumbnails_scroll.setWidgetResizable(True)
        self.thumbnails_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.thumbnails_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.thumbnails_scroll.hide()

        self.thumbnails_widget = QWidget()
        self.thumbnails_layout = QHBoxLayout(self.thumbnails_widget)
        self.thumbnails_layout.setAlignment(Qt.AlignLeft)
        self.thumbnails_scroll.setWidget(self.thumbnails_widget)

        layout.addWidget(self.thumbnails_scroll)

        buttons_layout = QHBoxLayout()

        self.send_btn = QPushButton("📤")  # Send icon
        self.send_btn.setObjectName("telegramButton")
        self.send_btn.setFixedSize(32, 32)
        self.send_btn.clicked.connect(self.send_feedback)

        self.cancel_btn = QPushButton("❌")  # Cancel icon
        self.cancel_btn.setObjectName("telegramButton")
        self.cancel_btn.setFixedSize(32, 32)
        self.cancel_btn.clicked.connect(self.reject)

        buttons_layout.addStretch()
        buttons_layout.addWidget(self.send_btn)
        buttons_layout.addWidget(self.cancel_btn)

        layout.addLayout(buttons_layout)

        # Paste shortcut
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.message_text)
        self.paste_shortcut.activated.connect(self.paste_image)

    def send_feedback(self):
        message = self.message_text.toPlainText().strip()

        if not message:
            QMessageBox.warning(self, tr('warning'), tr('please_enter_message'))
            return

        self.send_btn.setEnabled(False)
        self.send_btn.setText("⏳")

        email = self.email_input.text().strip()
        success, result_message = self.reporter.send_feedback(
            message,
            email if email else None,
            self.attached_images
        )

        self.send_btn.setEnabled(True)
        self.send_btn.setText("📤")

        if success:
            QMessageBox.information(self, tr('success'), result_message)
            self.accept()
        else:
            QMessageBox.warning(self, tr('error'), result_message)

    def attach_image(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            tr('select_images'),
            "",
            "Images (*.png *.jpg *.jpeg *.gif *.bmp)"
        )
        self.add_images(files)

    def add_images(self, files):
        for file in files:
            if file not in self.attached_images and len(self.attached_images) < 5:
                self.attached_images.append(file)
                self.add_thumbnail(file)

        if len(self.attached_images) >= 5:
            QMessageBox.information(self, tr('info'), tr('max_attachments'))

        self.update_thumbnails_visibility()

    def add_thumbnail(self, image_path):
        thumbnail = ImageThumbnail(image_path)
        thumbnail.removed.connect(self.remove_image)
        self.thumbnails_layout.addWidget(thumbnail)

    def remove_image(self, image_path):
        if image_path in self.attached_images:
            self.attached_images.remove(image_path)

        for i in range(self.thumbnails_layout.count()):
            widget = self.thumbnails_layout.itemAt(i).widget()
            if isinstance(widget, ImageThumbnail) and widget.image_path == image_path:
                widget.deleteLater()
                break

        self.update_thumbnails_visibility()

    def update_thumbnails_visibility(self):
        self.thumbnails_scroll.setVisible(len(self.attached_images) > 0)

    def paste_image(self):
        clipboard = QApplication.clipboard()

        if clipboard.image():
            image = clipboard.image()
            if not image.isNull():
                temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                image.save(temp_file.name, 'PNG')
                self.add_images([temp_file.name])
        elif clipboard.mimeData().hasUrls():
            files = []
            for url in clipboard.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                    files.append(file_path)
            if files:
                self.add_images(files)