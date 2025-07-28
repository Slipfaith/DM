# updater.py
import json
import urllib.request
import urllib.error
from packaging import version
import webbrowser
from PySide6.QtWidgets import QMessageBox

CURRENT_VERSION = "1.0.0"
GITHUB_API_URL = "https://api.github.com/repos/Slipfaith/DM/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/Slipfaith/DM/releases"


class UpdateChecker:
    def __init__(self, parent=None):
        self.parent = parent

    def check_for_updates(self):
        try:
            with urllib.request.urlopen(GITHUB_API_URL, timeout=5) as response:
                data = json.loads(response.read().decode())

            latest_version = data.get('tag_name', '').lstrip('v')
            if not latest_version:
                self._show_error("Could not determine latest version")
                return

            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                self._show_update_available(latest_version, data)
            else:
                self._show_no_updates()

        except urllib.error.URLError as e:
            self._show_error(f"Network error: {str(e)}")
        except Exception as e:
            self._show_error(f"Error checking for updates: {str(e)}")

    def _show_update_available(self, latest_version, release_data):
        release_notes = release_data.get('body', 'No release notes available')
        download_url = release_data.get('html_url', GITHUB_RELEASES_URL)

        msg = QMessageBox(self.parent)
        msg.setWindowTitle("Update Available")
        msg.setText(f"A new version is available!")
        msg.setInformativeText(
            f"Current version: {CURRENT_VERSION}\n"
            f"Latest version: {latest_version}\n\n"
            f"Release notes:\n{release_notes[:200]}..."
        )
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.button(QMessageBox.Yes).setText("Download")
        msg.button(QMessageBox.No).setText("Later")

        if msg.exec() == QMessageBox.Yes:
            webbrowser.open(download_url)

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
            f"Could not check for updates:\n{error_message}"
        )