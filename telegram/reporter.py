# telegram/reporter.py
import json
import urllib.request
import urllib.error
import urllib.parse
import os
import platform
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from .config import BOT_TOKEN, CHAT_ID, REPORT_COOLDOWN


class TelegramReporter:
    def __init__(self):
        self.last_report_time = None
        self.cache_file = Path("telegram_report_cache.json")
        self._load_cache()

    def _load_cache(self):
        if self.cache_file.exists():
            try:
                with open(self.cache_file, 'r') as f:
                    data = json.load(f)
                    last_time = data.get('last_report_time')
                    if last_time:
                        self.last_report_time = datetime.fromisoformat(last_time)
            except:
                pass

    def _save_cache(self):
        try:
            with open(self.cache_file, 'w') as f:
                json.dump({
                    'last_report_time': self.last_report_time.isoformat() if self.last_report_time else None
                }, f)
        except:
            pass

    def can_send_report(self):
        if not self.last_report_time:
            return True
        return datetime.now() - self.last_report_time > timedelta(seconds=REPORT_COOLDOWN)

    def send_error_report(self, error_message, log_content=None, user_message=None, images=None):
        if not self.can_send_report():
            remaining = REPORT_COOLDOWN - (datetime.now() - self.last_report_time).total_seconds()
            return False, f"Please wait {int(remaining / 60)} minutes before sending another report"

        try:
            system_info = {
                'platform': platform.platform(),
                'python_version': platform.python_version(),
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

            message = f"🚨 *Excel Processor Error Report*\n\n"
            message += f"*Time:* {system_info['timestamp']}\n"
            message += f"*System:* {system_info['platform']}\n"
            message += f"*Python:* {system_info['python_version']}\n\n"

            if user_message:
                message += f"*User Message:*\n{user_message}\n\n"

            message += f"*Error:*\n```\n{error_message[:1000]}\n```\n"

            if log_content:
                log_preview = log_content[-2000:] if len(log_content) > 2000 else log_content
                message += f"\n*Log (last 2000 chars):*\n```\n{log_preview}\n```"

            self._send_telegram_message(message)

            if images:
                for image_path in images:
                    try:
                        self._send_telegram_photo(image_path)
                    except:
                        pass

            self.last_report_time = datetime.now()
            self._save_cache()

            return True, "Report sent successfully"

        except Exception as e:
            return False, f"Failed to send report: {str(e)}"

    def send_feedback(self, user_message, email=None, images=None):
        if not self.can_send_report():
            remaining = REPORT_COOLDOWN - (datetime.now() - self.last_report_time).total_seconds()
            return False, f"Please wait {int(remaining / 60)} minutes before sending another message"

        try:
            message = f"💬 *Excel Processor Feedback*\n\n"
            message += f"*Time:* {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"

            if email:
                message += f"*Contact:* {email}\n"

            message += f"\n*Message:*\n{user_message}"

            self._send_telegram_message(message)

            if images:
                for image_path in images:
                    try:
                        self._send_telegram_photo(image_path)
                    except:
                        pass

            self.last_report_time = datetime.now()
            self._save_cache()

            return True, "Message sent successfully"

        except Exception as e:
            return False, f"Failed to send message: {str(e)}"

    def _send_telegram_message(self, text):
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"

        data = {
            'chat_id': CHAT_ID,
            'text': text,
            'parse_mode': 'Markdown'
        }

        data_encoded = urllib.parse.urlencode(data).encode('utf-8')
        req = urllib.request.Request(url, data=data_encoded)

        with urllib.request.urlopen(req, timeout=10) as response:
            result = json.loads(response.read().decode())
            if not result.get('ok'):
                raise Exception(f"Telegram API error: {result}")

    def _send_telegram_photo(self, image_path):
        url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"

        boundary = '----WebKitFormBoundary' + os.urandom(16).hex()

        with open(image_path, 'rb') as f:
            image_data = f.read()

        body = []
        body.append(f'------{boundary}')
        body.append('Content-Disposition: form-data; name="chat_id"')
        body.append('')
        body.append(str(CHAT_ID))
        body.append(f'------{boundary}')
        body.append(f'Content-Disposition: form-data; name="photo"; filename="{os.path.basename(image_path)}"')
        body.append('Content-Type: image/jpeg')
        body.append('')

        body_start = '\r\n'.join(body).encode('utf-8')
        body_end = f'\r\n------{boundary}--\r\n'.encode('utf-8')

        body_data = body_start + b'\r\n' + image_data + body_end

        req = urllib.request.Request(url)
        req.add_header('Content-Type', f'multipart/form-data; boundary=----{boundary}')
        req.add_header('Content-Length', str(len(body_data)))
        req.data = body_data

        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode())
            if not result.get('ok'):
                raise Exception(f"Telegram API error: {result}")

    def get_latest_log_content(self):
        log_dir = Path("logs")
        if not log_dir.exists():
            return None

        log_files = sorted(log_dir.glob("excel_processor_*.log"), key=os.path.getmtime, reverse=True)

        if log_files:
            try:
                with open(log_files[0], 'r', encoding='utf-8') as f:
                    return f.read()
            except:
                return None

        return None