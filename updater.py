# updater.py
import json
import urllib.request
import urllib.error
import os
import sys
import subprocess
import tempfile
from packaging import version
from PySide6.QtWidgets import QMessageBox, QProgressDialog
from PySide6.QtCore import QThread, Signal, QTimer

CURRENT_VERSION = "1.1.0"
GITHUB_API_URL = "https://api.github.com/repos/Slipfaith/DM/releases/latest"


class DownloadThread(QThread):
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, url, save_path):
        super().__init__()
        self.url = url
        self.save_path = save_path

    def run(self):
        try:
            response = urllib.request.urlopen(self.url)
            total_size = int(response.headers.get('Content-Length', 0))

            downloaded = 0
            block_size = 8192

            with open(self.save_path, 'wb') as f:
                while True:
                    buffer = response.read(block_size)
                    if not buffer:
                        break

                    downloaded += len(buffer)
                    f.write(buffer)

                    if total_size > 0:
                        progress = int((downloaded / total_size) * 100)
                        self.progress.emit(progress)

            self.finished.emit(self.save_path)

        except Exception as e:
            self.error.emit(str(e))


class UpdateChecker:
    def __init__(self, parent=None):
        self.parent = parent
        self.download_thread = None

    def check_for_updates(self, silent=False):
        try:
            with urllib.request.urlopen(GITHUB_API_URL, timeout=5) as response:
                data = json.loads(response.read().decode())

            latest_version = data.get('tag_name', '').lstrip('v')
            if not latest_version:
                if not silent:
                    self._show_error("Could not determine latest version")
                return

            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                self._show_update_available(latest_version, data)
            elif not silent:
                self._show_no_updates()

        except Exception as e:
            if not silent:
                self._show_error(f"Error checking for updates: {str(e)}")

    def _show_update_available(self, latest_version, release_data):
        exe_asset = None
        for asset in release_data.get('assets', []):
            if asset['name'].endswith('.exe'):
                exe_asset = asset
                break

        if not exe_asset:
            self._show_error("No executable file found in the release")
            return

        msg = QMessageBox(self.parent)
        msg.setWindowTitle("Update Available")
        msg.setText(f"Version {latest_version} is available!")
        msg.setInformativeText(
            f"Current version: {CURRENT_VERSION}\n"
            f"New version: {latest_version}\n\n"
            "Would you like to download and install it?"
        )
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        if msg.exec() == QMessageBox.Yes:
            self._download_update(exe_asset)

    def _download_update(self, asset):
        download_url = asset['browser_download_url']
        temp_file = os.path.join(tempfile.gettempdir(), asset['name'])

        self.progress_dialog = QProgressDialog(
            "Downloading update...",
            "Cancel",
            0, 100,
            self.parent
        )
        self.progress_dialog.setWindowTitle("Updating")
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.show()

        self.download_thread = DownloadThread(download_url, temp_file)
        self.download_thread.progress.connect(self.progress_dialog.setValue)
        self.download_thread.finished.connect(self._on_download_finished)
        self.download_thread.error.connect(self._on_download_error)
        self.progress_dialog.canceled.connect(self.download_thread.terminate)
        self.download_thread.start()

    def _on_download_finished(self, file_path):
        self.progress_dialog.close()

        msg = QMessageBox(self.parent)
        msg.setWindowTitle("Update Downloaded")
        msg.setText("Update downloaded successfully!")
        msg.setInformativeText(
            "The application will now close to install the update.\n"
            "Please restart the application after installation."
        )
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()

        self._install_update(file_path)

    def _on_download_error(self, error):
        self.progress_dialog.close()
        self._show_error(f"Download failed: {error}")

    def _install_update(self, new_exe_path):
        current_exe = sys.executable

        if getattr(sys, 'frozen', False):
            update_script = f'''
import os
import time
import shutil

time.sleep(2)

old_exe = r"{current_exe}"
new_exe = r"{new_exe_path}"
backup_exe = old_exe + ".backup"

try:
    if os.path.exists(backup_exe):
        os.remove(backup_exe)
    os.rename(old_exe, backup_exe)
    shutil.move(new_exe, old_exe)
    os.remove(backup_exe)
    os.startfile(old_exe)
except Exception as e:
    print(f"Update failed: {{e}}")
    if os.path.exists(backup_exe):
        os.rename(backup_exe, old_exe)
'''

            script_path = os.path.join(tempfile.gettempdir(), "update_script.py")
            with open(script_path, 'w') as f:
                f.write(update_script)

            subprocess.Popen([sys.executable, script_path],
                             creationflags=subprocess.CREATE_NO_WINDOW)

            QTimer.singleShot(100, self.parent.close)
        else:
            self._show_error("Auto-update only works with compiled exe files")

    def _show_no_updates(self):
        QMessageBox.information(
            self.parent,
            "No Updates",
            f"You are using the latest version ({CURRENT_VERSION})"
        )

    def _show_error(self, error_message):
        QMessageBox.warning(
            self.parent,
            "Update Check Failed",
            error_message
        )