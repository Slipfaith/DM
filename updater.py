# updater.py
import json
import urllib.request
import urllib.error
import os
import sys
import subprocess
import tempfile
import hashlib
import types
from packaging import version
from PySide6.QtWidgets import QMessageBox, QProgressDialog
from PySide6.QtCore import QThread, Signal, QTimer
from translations import tr

CURRENT_VERSION = "1.2.1"
GITHUB_API_URL = "https://api.github.com/repos/Slipfaith/DM/releases/latest"

PUBLIC_KEY = """-----BEGIN PGP PUBLIC KEY BLOCK-----

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
0c8=
=ohJm
-----END PGP PUBLIC KEY BLOCK-----
"""


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
        self.signature_asset = None

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
        sig_asset = None

        for asset in release_data.get('assets', []):
            if asset['name'].endswith('.exe'):
                exe_asset = asset
            elif asset['name'].endswith('.exe.sha256'):
                hash_asset = asset
            elif asset['name'].endswith('.exe.asc'):
                sig_asset = asset

        if not exe_asset:
            self._show_error(tr('no_exe_error'))
            return

        if not sig_asset:
            self._show_error("No signature file found for verification.")
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
            self.signature_asset = sig_asset
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

        sig_path = file_path + '.asc'
        try:
            with urllib.request.urlopen(self.signature_asset['browser_download_url']) as response, open(sig_path, 'wb') as f:
                f.write(response.read())
            if not self.verify_signature(file_path, sig_path):
                raise Exception('Signature verification failed')
        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)
            if os.path.exists(sig_path):
                os.remove(sig_path)
            QMessageBox.critical(self.parent, "Signature Verification Failed", str(e))
            return
        finally:
            if os.path.exists(sig_path):
                os.remove(sig_path)

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

    def verify_signature(self, file_path, sig_path):
        try:
            try:
                import imghdr  # noqa: F401
            except ModuleNotFoundError:
                mod = types.ModuleType('imghdr')
                mod.what = lambda *args, **kwargs: None
                sys.modules['imghdr'] = mod
            from pgpy import PGPKey, PGPSignature
            pubkey, _ = PGPKey.from_blob(PUBLIC_KEY)
            sig = PGPSignature.from_file(sig_path)
            with open(file_path, 'rb') as f:
                data = f.read()
            if not pubkey.verify(data, sig):
                raise Exception('Invalid signature')
            return True
        except Exception as e:
            raise Exception(f'Signature verification failed: {e}')

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