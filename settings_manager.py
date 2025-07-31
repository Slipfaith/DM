# settings_manager.py
import json
from pathlib import Path


class SettingsManager:
    def __init__(self):
        self.settings_file = Path("settings.json")
        self.settings = self._load_settings()

    def _load_settings(self):
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def _save_settings(self):
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=2)
        except:
            pass

    def get(self, key, default=None):
        return self.settings.get(key, default)

    def set(self, key, value):
        self.settings[key] = value
        self._save_settings()

    def get_language(self):
        return self.get('language', 'en')

    def set_language(self, language):
        self.set('language', language)


settings_manager = SettingsManager()