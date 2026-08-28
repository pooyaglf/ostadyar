import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from threading import Thread

import requests
from openpyxl import load_workbook


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
CONFIG_PATH = BASE_DIR / "config.local.json"
PROFESSOR_PHONES_PATH = Path(os.getenv("PROFESSOR_PHONES_PATH", DATA_DIR / "professor_phones.json"))
CHAT_IDS_PATH = Path(os.getenv("CHAT_IDS_PATH", DATA_DIR / "chat_ids.json"))
REMINDERS_SENT_PATH = Path(os.getenv("REMINDERS_SENT_PATH", DATA_DIR / "reminders_sent.json"))

DEFAULT_SHEET_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1sDIbSkFHlgsqxrYZyK2diG53eh4LT09h/export?format=xlsx"
)

START_MESSAGE = (
    "با عرض سلام و خوش آمد خدمت اساتید گرامی\n"
    "لطفا شماره همراه خود را جهت مشاهده ی نام دانشجویان اینترن مربوط به مطب خود را وارد نمایید"
)

INVALID_PHONE_MESSAGE = "با عرض پوزش شماره ی وارد شده ثبت نشده است"
SCHEDULE_FOOTER = "برنامه کلاس های شما به شکل بالا هست و در روز کلاس برای شما یک پیام یاداوری ارسال خواهد شد"
VIEW_CLASSES_TEXT = "مشاهده برنامه کلاس‌ها"
REGISTERED_PHONE_MESSAGE = "شماره همراه شما قبلا با شماره {phone} ثبت شده است."
LOGIN_SUCCESS_MESSAGE = (
    "با عرض سلام خدمت {professor_name}\n"
    "ضمن تشکر از زحمات شما\n"
    "لیست دانشجویان ماه آینده ی شما به صورت زیر است"
)

PERSIAN_MONTHS = {
    "فروردین": 1,
    "اردیبهشت": 2,
    "خرداد": 3,
    "تیر": 4,
    "مرداد": 5,
    "شهریور": 6,
    "مهر": 7,
    "آبان": 8,
    "آذر": 9,
    "دی": 10,
    "بهمن": 11,
    "اسفند": 12,
}


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


def normalize_text(value):
    text = str(value or "").strip()
    text = text.replace("ي", "ی").replace("ك", "ک")
    return re.sub(r"\s+", " ", text)


def normalize_phone(value):
    digits = re.sub(r"\D+", "", str(value or ""))
    if digits.startswith("0098"):
        digits = "0" + digits[4:]
    elif digits.startswith("98") and len(digits) == 12:
        digits = "0" + digits[2:]
    return digits


def to_persian_digits(value):
    return str(value).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def gregorian_to_jalali(g_year, g_month, g_day):
    g_days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    j_days_in_month = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    gy = g_year - 1600
    gm = g_month - 1
    gd = g_day - 1
    g_day_no = 365 * gy + (gy + 3) // 4 - (gy + 99) // 100 + (gy + 399) // 400
    for index in range(gm):
        g_day_no += g_days_in_month[index]
    leap = (gy + 1600) % 4 == 0 and ((gy + 1600) % 100 != 0 or (gy + 1600) % 400 == 0)
    if gm > 1 and leap:
        g_day_no += 1
    g_day_no += gd
    j_day_no = g_day_no - 79
    j_np = j_day_no // 12053
    j_day_no %= 12053
    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461
    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365
    jm = 0
    while jm < 11 and j_day_no >= j_days_in_month[jm]:
        j_day_no -= j_days_in_month[jm]
        jm += 1
    return jy, jm + 1, j_day_no + 1


def parse_jalali_date(value):
    if value:
        year, month, day = [int(part) for part in value.split("-")]
        return year, month, day
    today = date.today()
    return gregorian_to_jalali(today.year, today.month, today.day)


def load_config():
    config = read_json(CONFIG_PATH, {})
    token = os.getenv("BOT_TOKEN", config.get("bot_token", "")).strip()
    if not token:
        raise RuntimeError("Set BOT_TOKEN environment variable before running the bot.")
    config["bot_token"] = token
    config["api_base_url"] = os.getenv("API_BASE_URL", config.get("api_base_url", "https://tapi.bale.ai/bot"))
    config["poll_timeout_seconds"] = os.getenv("POLL_TIMEOUT_SECONDS", config.get("poll_timeout_seconds", 25))
    config["sheet_export_url"] = os.getenv("SHEET_EXPORT_URL", config.get("sheet_export_url", DEFAULT_SHEET_EXPORT_URL))
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
    Thread(target=server.serve_forever, daemon=True).start()
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
            return json.loads(response.read().decode("utf-8"))

    def get_updates(self):
        payload = {"timeout": self.poll_timeout, "limit": 50}
        if self.offset is not None:
            payload["offset"] = self.offset
        return self.request("getUpdates", payload)

    def send_message(self, chat_id, text, keyboard=None):
        payload = {"chat_id": chat_id, "text": text}
        if keyboard:
            payload["reply_markup"] = keyboard
        return self.request("sendMessage", payload)


class SheetSchedule:
    def __init__(self, export_url):
        self.export_url = export_url
        self.loaded_at = 0
        self.cache_seconds = int(os.getenv("SHEET_CACHE_SECONDS", "300"))
        self.workbook = None

    def worksheet(self):
        now = time.time()
        if self.workbook is None or now - self.loaded_at > self.cache_seconds:
            response = requests.get(self.export_url, timeout=30)
            response.raise_for_status()
            self.workbook = load_workbook(BytesIO(response.content), data_only=True)
            self.loaded_at = now
        return self.workbook[self.workbook.sheetnames[0]]

    def month_info(self):
        ws = self.worksheet()
        title = " ".join(normalize_text(cell.value) for cell in ws[1] if cell.value)
        found_month = None
        for month_name, month_number in PERSIAN_MONTHS.items():
            if month_name in title:
                found_month = (month_name, month_number)
                break
        year_match = re.search(r"(\d{4})", title)
        year = int(year_match.group(1)) if year_match else None
        if found_month and year:
            return year, found_month[1], found_month[0]
        current_year, current_month, _ = parse_jalali_date(os.getenv("REMINDER_DATE"))
        month_name = next(name for name, number in PERSIAN_MONTHS.items() if number == current_month)
        return year or current_year, current_month, month_name

    def professor_columns(self):
        ws = self.worksheet()
        columns = {}
        for col in range(1, ws.max_column + 1):
            value = normalize_text(ws.cell(2, col).value)
            if value.startswith("دکتر"):
                columns[value] = col
        return columns

    def schedule_for_professor(self, professor_name):
        ws = self.worksheet()
        col = self.professor_columns().get(normalize_text(professor_name))
        if not col:
            return []
        classes = []
        for row in range(3, ws.max_row + 1):
            weekday = normalize_text(ws.cell(row, 1).value)
            day_number = ws.cell(row, 2).value
            student = normalize_text(ws.cell(row, col).value)
            if not weekday or not day_number or not student or student == "*":
                continue
            classes.append({"weekday": weekday, "day": int(day_number), "student": student})
        return classes

    def schedule_for_day(self, professor_name, jalali_day):
        return [item for item in self.schedule_for_professor(professor_name) if item["day"] == int(jalali_day)]


def professor_phone_map():
    raw = read_json(PROFESSOR_PHONES_PATH, {})
    return {normalize_phone(phone): normalize_text(name) for name, phone in raw.items() if normalize_phone(phone)}


def chat_registry():
    raw = read_json(CHAT_IDS_PATH, {})
    if isinstance(raw, list):
        return {str(chat_id): {"phone": None, "professor": None} for chat_id in raw}
    return {
        str(chat_id): {
            "phone": normalize_phone(value.get("phone")),
            "professor": normalize_text(value.get("professor")),
        }
        for chat_id, value in raw.items()
        if isinstance(value, dict)
    }


def save_chat_registry(registry):
    write_json(CHAT_IDS_PATH, registry)


def set_chat_professor(chat_id, phone, professor):
    registry = chat_registry()
    registry[str(chat_id)] = {"phone": normalize_phone(phone), "professor": normalize_text(professor)}
    save_chat_registry(registry)


def get_chat_professor(chat_id):
    return normalize_text(chat_registry().get(str(chat_id), {}).get("professor"))


def get_chat_registration(chat_id):
    item = chat_registry().get(str(chat_id), {})
    return {
        "phone": normalize_phone(item.get("phone")),
        "professor": normalize_text(item.get("professor")),
    }


def classes_keyboard():
    return {
        "keyboard": [[{"text": VIEW_CLASSES_TEXT}]],
        "resize_keyboard": True,
        "one_time_keyboard": False,
    }


def format_phone(phone):
    return to_persian_digits(normalize_phone(phone))


def format_schedule(professor_name, schedule, schedule_reader):
    year, _, month_name = schedule_reader.month_info()
    if not schedule:
        return f"برای {professor_name} در این ماه برنامه‌ای ثبت نشده است."
    lines = [f"برنامه کلاس های {professor_name}:", ""]
    for index, item in enumerate(schedule, start=1):
        lines.append(
            f"{index}. {item['weekday']} {to_persian_digits(item['day'])} {month_name} {to_persian_digits(year)}\n"
            f"   دانشجو: {item['student']}"
        )
    lines.extend(["", SCHEDULE_FOOTER])
    return "\n".join(lines)


def format_login_schedule(professor_name, schedule, schedule_reader):
    return "\n\n".join(
        [
            LOGIN_SUCCESS_MESSAGE.format(professor_name=professor_name),
            format_schedule(professor_name, schedule, schedule_reader),
        ]
    )


def format_reminder(professor_name, day_schedule, schedule_reader):
    year, _, month_name = schedule_reader.month_info()
    day = day_schedule[0]["day"]
    lines = [
        f"یادآوری برنامه امروز {professor_name}",
        f"{day_schedule[0]['weekday']} {to_persian_digits(day)} {month_name} {to_persian_digits(year)}",
        "",
    ]
    for index, item in enumerate(day_schedule, start=1):
        lines.append(f"{index}. دانشجو: {item['student']}")
    return "\n".join(lines)


def handle_message(bot, schedule_reader, message):
    chat_id = (message.get("chat") or {}).get("id")
    text = normalize_text(message.get("text"))
    if chat_id is None:
        return

    registration = get_chat_registration(chat_id)
    registered_phone = registration["phone"]
    registered_professor = registration["professor"]

    if text == "/start":
        if registered_phone and registered_professor:
            schedule = schedule_reader.schedule_for_professor(registered_professor)
            bot.send_message(
                chat_id,
                format_login_schedule(registered_professor, schedule, schedule_reader),
                keyboard=classes_keyboard(),
            )
            return
        bot.send_message(chat_id, START_MESSAGE)
        return

    if text == VIEW_CLASSES_TEXT:
        if registered_phone and registered_professor:
            schedule = schedule_reader.schedule_for_professor(registered_professor)
            bot.send_message(
                chat_id,
                format_schedule(registered_professor, schedule, schedule_reader),
                keyboard=classes_keyboard(),
            )
            return
        bot.send_message(chat_id, START_MESSAGE)
        return

    phone = normalize_phone(text)
    if registered_phone and registered_professor:
        if phone and phone != registered_phone:
            bot.send_message(
                chat_id,
                REGISTERED_PHONE_MESSAGE.format(phone=format_phone(registered_phone)),
                keyboard=classes_keyboard(),
            )
            return
        schedule = schedule_reader.schedule_for_professor(registered_professor)
        bot.send_message(
            chat_id,
            format_schedule(registered_professor, schedule, schedule_reader),
            keyboard=classes_keyboard(),
        )
        return

    phones = professor_phone_map()
    if phone in phones:
        professor = phones[phone]
        set_chat_professor(chat_id, phone, professor)
        schedule = schedule_reader.schedule_for_professor(professor)
        bot.send_message(
            chat_id,
            format_login_schedule(professor, schedule, schedule_reader),
            keyboard=classes_keyboard(),
        )
        return

    bot.send_message(chat_id, INVALID_PHONE_MESSAGE)


def reminder_key(chat_id, year, month, day):
    return f"{chat_id}:{year:04d}-{month:02d}-{day:02d}"


def send_due_reminders(bot, schedule_reader, target_date=None, dry_run=False):
    year, month, day = parse_jalali_date(target_date)
    sheet_year, sheet_month, _ = schedule_reader.month_info()
    if year != sheet_year or month != sheet_month:
        return []
    sent = read_json(REMINDERS_SENT_PATH, {})
    messages = []
    for chat_id, value in chat_registry().items():
        professor = normalize_text(value.get("professor"))
        if not professor:
            continue
        day_schedule = schedule_reader.schedule_for_day(professor, day)
        if not day_schedule:
            continue
        key = reminder_key(chat_id, year, month, day)
        if sent.get(key):
            continue
        text = format_reminder(professor, day_schedule, schedule_reader)
        messages.append({"chat_id": chat_id, "professor": professor, "text": text})
        if not dry_run:
            bot.send_message(chat_id, text)
            sent[key] = datetime.now().isoformat(timespec="seconds")
    if messages and not dry_run:
        write_json(REMINDERS_SENT_PATH, sent)
    return messages


def reminder_loop(bot, schedule_reader):
    hour = int(os.getenv("REMINDER_HOUR", "8"))
    while True:
        if datetime.now().hour >= hour:
            send_due_reminders(bot, schedule_reader)
        time.sleep(300)


def run():
    config = load_config()
    start_health_server()
    bot = BaleBot(config)
    schedule_reader = SheetSchedule(config["sheet_export_url"])
    Thread(target=reminder_loop, args=(bot, schedule_reader), daemon=True).start()
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
                    handle_message(bot, schedule_reader, message)
        except KeyboardInterrupt:
            print("\nBot stopped.")
            break
        except (urllib.error.URLError, urllib.error.HTTPError, requests.RequestException, TimeoutError) as exc:
            print(f"Network/API error: {exc}")
            time.sleep(5)
        except Exception as exc:
            print(f"Unexpected error: {exc}")
            time.sleep(5)


def preview_schedule(args):
    config = load_config()
    reader = SheetSchedule(config["sheet_export_url"])
    professor = professor_phone_map().get(normalize_phone(args.phone))
    if not professor:
        print(INVALID_PHONE_MESSAGE)
        return
    print(format_schedule(professor, reader.schedule_for_professor(professor), reader))


def preview_reminders(args):
    config = load_config()
    reader = SheetSchedule(config["sheet_export_url"])
    if args.phone:
        professor = professor_phone_map().get(normalize_phone(args.phone))
        if not professor:
            print(INVALID_PHONE_MESSAGE)
            return
        _, _, day = parse_jalali_date(args.date)
        day_schedule = reader.schedule_for_day(professor, day)
        if not day_schedule:
            print("No reminders found for this professor and date.")
            return
        print(format_reminder(professor, day_schedule, reader))
        return

    messages = send_due_reminders(None, reader, target_date=args.date, dry_run=True)
    if not messages:
        print("No reminders found for this date.")
        return
    for message in messages:
        print(f"CHAT_ID: {message['chat_id']}")
        print(message["text"])
        print("-" * 30)


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command")
    schedule_parser = subparsers.add_parser("preview-schedule")
    schedule_parser.add_argument("--phone", required=True)
    reminder_parser = subparsers.add_parser("test-reminders")
    reminder_parser.add_argument("--date", required=True, help="Jalali date, for example 1405-04-01")
    reminder_parser.add_argument("--phone", help="Optional professor phone for testing before Bale login")
    args = parser.parse_args()
    if args.command == "preview-schedule":
        preview_schedule(args)
    elif args.command == "test-reminders":
        preview_reminders(args)
    else:
        run()


if __name__ == "__main__":
    main()
