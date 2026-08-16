import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CONFIG_PATH = BASE_DIR / "config.local.json"
PROFESSORS_PATH = DATA_DIR / "professors.json"
STUDENTS_PATH = DATA_DIR / "students.json"
CHAT_IDS_PATH = Path(os.getenv("CHAT_IDS_PATH", DATA_DIR / "chat_ids.json"))

START_MESSAGE = (
"با عرض سلام و خوش آمد خدمت اساتید گرامی \n"
"با انتخاب نام خود در جدول زیر نام دانشجویان اینترن مربوط به مطب شما نمایش داده می شود:"
)


def read_json(path, default):
    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except FileNotFoundError:
        return default
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Invalid JSON in {path}: {exc}") from exc


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        json.dump(value, file, ensure_ascii=False, indent=2)
        file.write("\n")


def load_config():
    config = read_json(CONFIG_PATH, {})
    token = os.getenv("BOT_TOKEN", config.get("bot_token", "")).strip()
    if not token:
        raise RuntimeError("Set BOT_TOKEN environment variable before running the bot.")
    config["bot_token"] = token
    config["api_base_url"] = os.getenv(
        "API_BASE_URL",
        config.get("api_base_url", "https://tapi.bale.ai/bot"),
    )
    config["poll_timeout_seconds"] = os.getenv(
        "POLL_TIMEOUT_SECONDS",
        config.get("poll_timeout_seconds", 25),
    )
    return config


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("ok\n".encode("utf-8"))

    def log_message(self, format, *args):
        return


def start_health_server():
    port = int(os.getenv("PORT", "8000"))
    server = ThreadingHTTPServer(("0.0.0.0", port), HealthHandler)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    print(f"Health server is listening on port {port}.")


class BaleBot:
    def __init__(self, config):
        token = config["bot_token"].strip()
        api_base_url = config.get("api_base_url", "https://tapi.bale.ai/bot").rstrip("/")
        self.base_url = f"{api_base_url}{token}"
        self.poll_timeout = int(config.get("poll_timeout_seconds", 25))
        self.offset = None

    def request(self, method, payload=None):
        url = f"{self.base_url}/{method}"
        data = None
        headers = {}

        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"

        request = urllib.request.Request(url, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(request, timeout=self.poll_timeout + 10) as response:
            body = response.read().decode("utf-8")
            return json.loads(body)

    def get_updates(self):
        payload = {
            "timeout": self.poll_timeout,
            "limit": 50,
        }
        if self.offset is not None:
            payload["offset"] = self.offset
        return self.request("getUpdates", payload)

    def send_message(self, chat_id, text, keyboard=None):
        payload = {
            "chat_id": chat_id,
            "text": text,
        }
        if keyboard:
            payload["reply_markup"] = keyboard
        return self.request("sendMessage", payload)


def get_professors():
    professors = read_json(PROFESSORS_PATH, [])
    return [str(name).strip() for name in professors if str(name).strip()]


def get_students():
    students = read_json(STUDENTS_PATH, {})
    return {
        str(professor).strip(): value
        for professor, value in students.items()
        if str(professor).strip()
    }


def professor_keyboard():
    professors = get_professors()
    rows = [[{"text": professor}] for professor in professors]
    return {
        "keyboard": rows,
        "resize_keyboard": True,
        "one_time_keyboard": False,
    }


def remember_chat_id(chat_id):
    chat_id = str(chat_id)
    chat_ids = [str(item) for item in read_json(CHAT_IDS_PATH, [])]
    if chat_id not in chat_ids:
        chat_ids.append(chat_id)
        write_json(CHAT_IDS_PATH, chat_ids)


def format_students(professor_name):
    students_by_professor = get_students()
    students = students_by_professor.get(professor_name, [])
    if not students:
        return f"برای {professor_name} هنوز دانشجویی ثبت نشده است."

    lines = [f"دانشجویان {professor_name}:", ""]
    for index, student in enumerate(students, start=1):
        name = str(student.get("name", "")).strip()
        student_id = str(student.get("student_id", "")).strip()
        lines.append(f"{index}. {name} - شماره دانشجویی: {student_id}")
    return "\n".join(lines)


def handle_message(bot, message):
    chat = message.get("chat") or {}
    chat_id = chat.get("id")
    text = str(message.get("text") or "").strip()

    if chat_id is None:
        return

    remember_chat_id(chat_id)

    professors = get_professors()
    if text in professors:
        bot.send_message(chat_id, format_students(text), keyboard=professor_keyboard())
        return

    bot.send_message(chat_id, START_MESSAGE, keyboard=professor_keyboard())


def run():
    start_health_server()
    bot = BaleBot(load_config())
    print("Ostad Yar bot is running. Press Ctrl+C to stop.")

    while True:
        try:
            response = bot.get_updates()
            for update in response.get("result", []):
                update_id = update.get("update_id")
                if update_id is not None:
                    bot.offset = int(update_id) + 1

                message = update.get("message")
                if message:
                    handle_message(bot, message)
        except KeyboardInterrupt:
            print("\nBot stopped.")
            break
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as exc:
            print(f"Network/API error: {exc}")
            time.sleep(5)
        except Exception as exc:
            print(f"Unexpected error: {exc}")
            time.sleep(5)


if __name__ == "__main__":
    run()
