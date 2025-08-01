# updater.py
import json
import urllib.request
import urllib.error
import os
import sys
import subprocess
import tempfile
import hashlib
from packaging import version
from PySide6.QtWidgets import QMessageBox, QProgressDialog
from PySide6.QtCore import QThread, Signal, QTimer
from translations import tr

CURRENT_VERSION = "1.2.0"
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
        self.release_data = None
        self.current_asset = None

    def check_for_updates(self, silent=False):
        try:
            with urllib.request.urlopen(GITHUB_API_URL, timeout=5) as response:
                data = json.loads(response.read().decode())

            latest_version = data.get('tag_name', '').lstrip('v')
            if not latest_version:
                if not silent:
                    self._show_error(tr('could_not_determine'))
                return

            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                self._show_update_available(latest_version, data)
            elif not silent:
                self._show_no_updates()

        except Exception as e:
            if not silent:
                self._show_error(tr('error_checking_updates', error=str(e)))

    def _show_update_available(self, latest_version, release_data):
        exe_asset = None
        hash_asset = None

        for asset in release_data.get('assets', []):
            if asset['name'].endswith('.exe'):
                exe_asset = asset
            elif asset['name'].endswith('.exe.sha256'):
                hash_asset = asset

        if not exe_asset:
            self._show_error(tr('no_exe_error'))
            return

        if not hash_asset:
            reply = QMessageBox.warning(
                self.parent,
                tr('warning'),
                "No hash file found for verification.\nThis update may not be authentic.\n\nContinue anyway?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        msg = QMessageBox(self.parent)
        msg.setWindowTitle(tr('update_available_title'))
        msg.setText(tr('update_available_text', latest=latest_version))
        msg.setInformativeText(
            f"{tr('current_version', version=CURRENT_VERSION)}\n"
            f"{tr('new_version', version=latest_version)}\n\n"
            f"{tr('update_prompt')}"
        )
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        if msg.exec() == QMessageBox.Yes:
            self.release_data = release_data
            self.current_asset = exe_asset
            self._download_update(exe_asset)

    def _download_update(self, asset):
        download_url = asset['browser_download_url']
        temp_file = os.path.join(tempfile.gettempdir(), asset['name'])

        self.progress_dialog = QProgressDialog(
            tr('downloading_update'),
            tr('cancel'),
            0, 100,
            self.parent
        )
        self.progress_dialog.setWindowTitle(tr('updating'))
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

        # Try to download and verify hash
        hash_verified = False
        for asset in self.release_data.get('assets', []):
            if asset['name'] == self.current_asset['name'] + '.sha256':
                try:
                    with urllib.request.urlopen(asset['browser_download_url']) as response:
                        expected_hash = response.read().decode().strip().split()[0]

                    if self.verify_file_hash(file_path, expected_hash):
                        hash_verified = True
                except Exception as e:
                    QMessageBox.warning(
                        self.parent,
                        "Verification Warning",
                        f"Could not verify file integrity:\n{str(e)}\n\nProceed with caution!"
                    )
                break

        if not hash_verified:
            reply = QMessageBox.warning(
                self.parent,
                "No Verification",
                "File integrity could not be verified.\nInstall anyway?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                os.remove(file_path)
                return

        msg = QMessageBox(self.parent)
        msg.setWindowTitle(tr('update_downloaded_title'))
        msg.setText(tr('update_downloaded_text'))
        msg.setInformativeText(tr('update_restart_text'))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()

        self._install_update(file_path)

    def _on_download_error(self, error):
        self.progress_dialog.close()
        self._show_error(tr('download_failed', error=error))

    def verify_file_hash(self, file_path, expected_hash):
        """Verify file integrity using SHA256 hash"""
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)

        actual_hash = sha256_hash.hexdigest().lower()
        expected_hash = expected_hash.lower()

        if actual_hash != expected_hash:
            raise Exception(
                f"Hash mismatch! File may be corrupted or tampered.\n"
                f"Expected: {expected_hash}\n"
                f"Actual: {actual_hash}"
            )

        return True

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
            self._show_error(tr('auto_update_exe_only'))

    def _show_no_updates(self):
        QMessageBox.information(
            self.parent,
            tr('no_updates_title'),
            tr('no_updates_text', version=CURRENT_VERSION)
        )

    def _show_error(self, error_message):
        QMessageBox.warning(
            self.parent,
            tr('update_error_title'),
            error_message
        )