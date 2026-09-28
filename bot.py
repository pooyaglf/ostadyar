import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timezone as datetime_timezone, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from threading import Thread
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook
from openpyxl import load_workbook


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
CONFIG_PATH = BASE_DIR / "config.local.json"
ABSENCES_FALLBACK_PATH = Path(os.getenv("ABSENCES_FALLBACK_PATH", DATA_DIR / "absences_fallback.json"))
ABSENCES_XLSX_PATH = Path(os.getenv("ABSENCES_XLSX_PATH", DATA_DIR / "absences.xlsx"))
TEST_CHAT_REGISTRY_PATH = Path(os.getenv("TEST_CHAT_REGISTRY_PATH", BASE_DIR / "test_chat_registry.json"))

# Main operational settings. Change these values, then restart the bot.
# Leave these empty to read year/month from the class sheet title.
SCHEDULE_YEAR_OVERRIDE = ""
SCHEDULE_MONTH_OVERRIDE = ""
CURRENT_JALALI_DATE_OVERRIDE = ""
REMINDER_TIME = "09:00"
REMINDER_WINDOW_MINUTES = 15
ATTENDANCE_TIME = "21:00"
ATTENDANCE_WINDOW_MINUTES = 30
MONTHLY_SCHEDULE_SEND_AT = "1405/07/03 09:00"
SCHEDULER_INTERVAL_SECONDS = 15
BOT_TIMEZONE = "Asia/Tehran"
SHEET_ACTION_RETRY_ATTEMPTS = 4
SHEET_ACTION_RETRY_DELAY_SECONDS = 10
MANUAL_TEST_CLASSES = []

DEFAULT_SHEET_EXPORT_URL = (
    # Active class schedule sheet:
    # https://docs.google.com/spreadsheets/d/1jwQ-2k6zbOGLTgPjSpOXGvnlBRWm70Mk/edit
    "https://docs.google.com/spreadsheets/d/"
    "1jwQ-2k6zbOGLTgPjSpOXGvnlBRWm70Mk/export?format=xlsx"
)
DEFAULT_CONTACTS_EXPORT_URL = (
    # Professor phone sheet + absence sheet:
    # https://docs.google.com/spreadsheets/d/1P_wWkcMIpsUZYME8xCllQRvjfHSGCa0s/edit
    "https://docs.google.com/spreadsheets/d/"
    "1P_wWkcMIpsUZYME8xCllQRvjfHSGCa0s/export?format=xlsx"
)
DEFAULT_ABSENCE_WEBHOOK_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwnjF9SHN--LO5pyAhk6ToENYK09WFaTvtUu1UkMFrhKcS5x9O3Xxdb_uOxIClZ4Y2Qhw/exec"
)
DEFAULT_CHAT_REGISTRY_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1nXdlhmGj4lDsIK237I-I77j5VUpUjrsvwI3uCeOC7bU/export?format=xlsx"
)
DEFAULT_CHAT_REGISTRY_WEBHOOK_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbxK7dXQcYg0s0vOtmKt0SZT8M6t5dJ0pk-K3GPIi4ggfH_OMxtXjXNf18PhFveveffjQg/exec"
)
DEFAULT_PENDING_ATTENDANCE_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1pMrPJy4vSOWZJM4uG5F29v_wDMflLXuSm4hCzVMu1wg/export?format=xlsx"
)
DEFAULT_PENDING_ATTENDANCE_WEBHOOK_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwbcqnm9HaxmBIwoY0V4lfAHFZQ_wasGCfWL_QJh4vg3RmNLIZF1vSlHzQwzPfuymHk/exec"
)
DEFAULT_ATTENDANCE_LOG_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1x45n5nnYM1GP5E527CDx5yuZ46A_sVCcPgUqdtMhEro/export?format=xlsx"
)
DEFAULT_ATTENDANCE_LOG_WEBHOOK_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwJ0YFWJkVQhDTdvclLGnj0FNE_oCTYivJJnzgmuhQajBPcsQy5Ygf-TG8wER3XKYT0Bw/exec"
)
START_MESSAGE = (
    "با عرض سلام و احترام\n"
    "به ربات کلینیک ویژه ی دانشگاه علوم پزشکی اصفهان خوش آمدید.\n"
    "جهت ادامه ی فرایند و تکمیل اطلاعات، نقش خود را انتخاب کنید :"
)

PROFESSOR_INTRO_MESSAGE = (
    "با عرض سلام مجدد خدمت اساتید گرامی\n"
    "و ضمن تشکر از زحمات شما\n"
    "این ربات جهت مشاهده ی برنامه ی ماهانه و یادآوری روزانه ی اسامی اینترن های شما طراحی شده است .\n"
    "همچنین در این ربات امکان ثبت حضور و غیاب و ارزیابی عملکرد اینترن ها فراهم گردیده است .\n"
    "جهت احراز هویت و تکمیل اطلاعات خود بر روی دکمه ی اشتراک گذاری تلفن همراه کلیک بفرمایید ."
)

STUDENT_SECTION_MESSAGE = "این بخش در حال توسعه است"
INVALID_PHONE_MESSAGE = "با عرض معذرت شماره ی شما ثبت نشده است"
CONTACT_REQUIRED_MESSAGE = "لطفا شماره را به صورت دستی وارد نکنید و از دکمه اشتراک گذاری شماره همراه استفاده نمایید."
CONTACT_OWNER_MISMATCH_MESSAGE = "لطفا فقط شماره همراه حساب بله خودتان را با دکمه اشتراک گذاری ارسال نمایید."
PROCESSING_MESSAGE = "در حال پردازش اطلاعات، لطفا چند لحظه صبر بفرمایید..."
SHARE_CONTACT_TEXT = "📱 اشتراک‌گذاری شماره همراه"
ROLE_STUDENT_TEXT = "دانشجو"
ROLE_PROFESSOR_TEXT = "استاد"
RESELECT_ROLE_TEXT = "انتخاب مجدد نقش"
SCHEDULE_FOOTER = "برنامه ی کلاس های شما به شکل بالا است و در روز کلاس برای شما یک پیام یاداوری ارسال خواهد شد"
VIEW_CLASSES_TEXT = "برنامه ماهانه کلینیک ویژه من"
TODAY_ATTENDANCE_TEXT = "دریافت حضور و غیاب امروز"
ATTENDANCE_ABSENT_TEXT = "غیبت"
ATTENDANCE_WEAK_TEXT = "ضعیف"
ATTENDANCE_MEDIUM_TEXT = "متوسط"
ATTENDANCE_EXCELLENT_TEXT = "عالی"
REGISTERED_PHONE_MESSAGE = "شماره همراه شما قبلا با شماره {phone} ثبت شده است."
LOGIN_SUCCESS_MESSAGE = (
    "با عرض تشکر خدمت استاد گرامی {professor_name}\n"
    "لیست اینترن های ماه آینده ی شما به صورت زیر است :"
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

PERSIAN_WEEKDAYS = [
    "دوشنبه",
    "سه شنبه",
    "چهارشنبه",
    "پنج شنبه",
    "جمعه",
    "شنبه",
    "یکشنبه",
]

REMINDERS_SENT_THIS_RUN = set()
ATTENDANCE_SENT_THIS_RUN = set()
MONTHLY_SCHEDULE_SENT_THIS_RUN = set()
SHEET_ROW_CACHE = {}


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


def config_or_env(env_name, default):
    return os.getenv(env_name, default).strip()


def cached_sheet_rows(export_url, cache_seconds=10):
    now = time.time()
    cached = SHEET_ROW_CACHE.get(export_url)
    if cached and now - cached["loaded_at"] <= cache_seconds:
        return cached["rows"]

    response = requests.get(export_url, timeout=30)
    response.raise_for_status()
    workbook = load_workbook(BytesIO(response.content), data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    headers = [normalize_text(cell.value) for cell in worksheet[1]]
    rows = []
    for raw_row in worksheet.iter_rows(min_row=2, values_only=True):
        row = {}
        for index, header in enumerate(headers):
            if header:
                row[header] = raw_row[index] if index < len(raw_row) else None
        if any(value not in (None, "") for value in row.values()):
            rows.append(row)
    SHEET_ROW_CACHE[export_url] = {"loaded_at": now, "rows": rows}
    return rows


def invalidate_sheet_cache(export_url):
    SHEET_ROW_CACHE.pop(export_url, None)


def post_sheet_action(webhook_url, payload):
    response = requests.post(webhook_url, json=payload, timeout=30)
    response.raise_for_status()
    result = response.json()
    if not result.get("ok"):
        raise RuntimeError(result.get("error") or "Google Sheet webhook returned ok=false")
    return result


def post_sheet_action_with_retry(webhook_url, payload, attempts=None, delay_seconds=None):
    attempts = int(attempts or SHEET_ACTION_RETRY_ATTEMPTS)
    delay_seconds = int(delay_seconds or SHEET_ACTION_RETRY_DELAY_SECONDS)
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            return post_sheet_action(webhook_url, payload)
        except Exception as exc:
            last_error = exc
            if attempt >= attempts:
                break
            print(
                "Google Sheet write failed "
                f"(attempt {attempt}/{attempts}); retrying in {delay_seconds}s: {exc}"
            )
            time.sleep(delay_seconds)
    raise last_error


def normalize_text(value):
    text = str(value or "").strip()
    text = text.replace("ي", "ی").replace("ك", "ک")
    return re.sub(r"\s+", " ", text)


def normalize_digits(value):
    return str(value or "").translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789"))


def normalize_phone(value):
    if isinstance(value, float) and value.is_integer():
        value = str(int(value))
    elif isinstance(value, int):
        value = str(value)
    else:
        value = str(value or "").strip()
        if re.fullmatch(r"\d+\.0", value):
            value = value[:-2]
    value = normalize_digits(value)
    digits = re.sub(r"\D+", "", value)
    if digits.startswith("0098"):
        digits = "0" + digits[4:]
    elif digits.startswith("98") and len(digits) >= 12:
        digits = "0" + digits[2:12]
    elif digits.startswith("9") and len(digits) == 10:
        digits = "0" + digits
    return digits


def normalize_id(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def compact_text(value):
    return normalize_text(value).replace(" ", "")


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


def jalali_to_gregorian(j_year, j_month, j_day):
    j_days_in_month = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    g_days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    jy = j_year - 979
    jm = j_month - 1
    jd = j_day - 1

    j_day_no = 365 * jy + (jy // 33) * 8 + ((jy % 33) + 3) // 4
    for index in range(jm):
        j_day_no += j_days_in_month[index]
    j_day_no += jd

    g_day_no = j_day_no + 79
    gy = 1600 + 400 * (g_day_no // 146097)
    g_day_no %= 146097

    leap = True
    if g_day_no >= 36525:
        g_day_no -= 1
        gy += 100 * (g_day_no // 36524)
        g_day_no %= 36524
        if g_day_no >= 365:
            g_day_no += 1
        else:
            leap = False

    gy += 4 * (g_day_no // 1461)
    g_day_no %= 1461

    if g_day_no >= 366:
        leap = False
        g_day_no -= 1
        gy += g_day_no // 365
        g_day_no %= 365

    gm = 0
    while gm < 11:
        days_in_month = g_days_in_month[gm]
        if gm == 1 and leap:
            days_in_month += 1
        if g_day_no < days_in_month:
            break
        g_day_no -= days_in_month
        gm += 1

    return gy, gm + 1, g_day_no + 1


def jalali_weekday_name(j_year, j_month, j_day):
    g_year, g_month, g_day = jalali_to_gregorian(j_year, j_month, j_day)
    return PERSIAN_WEEKDAYS[date(g_year, g_month, g_day).weekday()]


def parse_jalali_date(value):
    if value:
        year, month, day = [int(part) for part in value.split("-")]
        return year, month, day
    today = date.today()
    return gregorian_to_jalali(today.year, today.month, today.day)


def config_value(config, env_name, key, default=None):
    value = os.getenv(env_name)
    if value is not None:
        return value
    return config.get(key, default)


def config_bool(config, env_name, key, default=False):
    value = config_value(config, env_name, key, default)
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def parse_time(value, default):
    value = str(value or default).strip()
    if ":" in value:
        hour, minute = value.split(":", 1)
        return int(hour), int(minute)
    return int(value), 0


def time_reached(now, value):
    hour, minute = parse_time(value, "00:00")
    return (now.hour, now.minute) >= (hour, minute)


def time_matches(now, value):
    hour, minute = parse_time(value, "00:00")
    return now.hour == hour and now.minute == minute


def time_is_in_window(now, start_value, window_minutes):
    hour, minute = parse_time(start_value, "00:00")
    start = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    end = start + timedelta(minutes=int(window_minutes))
    return start <= now < end


def get_bot_timezone(timezone_name):
    timezone_name = str(timezone_name or BOT_TIMEZONE).strip() or BOT_TIMEZONE
    try:
        return ZoneInfo(timezone_name)
    except Exception as exc:
        if timezone_name == "Asia/Tehran":
            print(f"Timezone data for {timezone_name} is unavailable, using fixed +03:30 offset: {exc}")
            return datetime_timezone(timedelta(hours=3, minutes=30), timezone_name)
        print(f"Timezone data for {timezone_name} is unavailable, using UTC: {exc}")
        return datetime_timezone.utc


def timezone_label(timezone):
    return getattr(timezone, "key", str(timezone))


def load_config():
    config = read_json(CONFIG_PATH, {})
    token = os.getenv("BOT_TOKEN", config.get("bot_token", "")).strip()
    if not token:
        raise RuntimeError("Set BOT_TOKEN environment variable before running the bot.")
    config["bot_token"] = token
    config["api_base_url"] = os.getenv("API_BASE_URL", config.get("api_base_url", "https://tapi.bale.ai/bot"))
    config["poll_timeout_seconds"] = os.getenv("POLL_TIMEOUT_SECONDS", config.get("poll_timeout_seconds", 25))
    config["sheet_export_url"] = os.getenv("SHEET_EXPORT_URL", config.get("sheet_export_url", DEFAULT_SHEET_EXPORT_URL))
    config["contacts_export_url"] = os.getenv(
        "CONTACTS_EXPORT_URL",
        config.get("contacts_export_url", DEFAULT_CONTACTS_EXPORT_URL),
    )
    config["absence_webhook_url"] = os.getenv(
        "ABSENCE_WEBHOOK_URL",
        config.get("absence_webhook_url", DEFAULT_ABSENCE_WEBHOOK_URL),
    )
    config["current_jalali_date"] = os.getenv("CURRENT_JALALI_DATE", CURRENT_JALALI_DATE_OVERRIDE)
    config["schedule_year_override"] = os.getenv("SCHEDULE_YEAR_OVERRIDE", str(SCHEDULE_YEAR_OVERRIDE))
    config["schedule_month_override"] = os.getenv("SCHEDULE_MONTH_OVERRIDE", str(SCHEDULE_MONTH_OVERRIDE))
    config["reminder_time"] = os.getenv("REMINDER_TIME", REMINDER_TIME)
    config["reminder_window_minutes"] = int(os.getenv("REMINDER_WINDOW_MINUTES", str(REMINDER_WINDOW_MINUTES)))
    config["attendance_time"] = os.getenv("ATTENDANCE_TIME", ATTENDANCE_TIME)
    config["attendance_window_minutes"] = int(os.getenv("ATTENDANCE_WINDOW_MINUTES", str(ATTENDANCE_WINDOW_MINUTES)))
    config["monthly_schedule_send_at"] = os.getenv("MONTHLY_SCHEDULE_SEND_AT", MONTHLY_SCHEDULE_SEND_AT)
    config["scheduler_interval_seconds"] = int(os.getenv("SCHEDULER_INTERVAL_SECONDS", str(SCHEDULER_INTERVAL_SECONDS)))
    config["bot_timezone"] = os.getenv("BOT_TIMEZONE", BOT_TIMEZONE)
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
        payload = {"chat_id": normalize_id(chat_id), "text": text}
        if keyboard:
            payload["reply_markup"] = keyboard
        return self.request("sendMessage", payload)

    def delete_message(self, chat_id, message_id):
        payload = {"chat_id": normalize_id(chat_id), "message_id": message_id}
        return self.request("deleteMessage", payload)

    def answer_callback_query(self, callback_query_id, text=None):
        payload = {"callback_query_id": callback_query_id}
        if text:
            payload["text"] = text
        return self.request("answerCallbackQuery", payload)


class WorkbookCache:
    def __init__(self, export_url):
        self.export_url = export_url
        self.loaded_at = 0
        self.cache_seconds = int(os.getenv("SHEET_CACHE_SECONDS", "300"))
        self.workbook = None

    def workbook_data(self):
        now = time.time()
        if self.workbook is None or now - self.loaded_at > self.cache_seconds:
            response = requests.get(self.export_url, timeout=30)
            response.raise_for_status()
            self.workbook = load_workbook(BytesIO(response.content), data_only=True)
            self.loaded_at = now
        return self.workbook


class SheetSchedule:
    def __init__(self, export_url, config=None):
        self.config = config or {}
        self.cache = WorkbookCache(export_url)

    def worksheet(self):
        workbook = self.cache.workbook_data()
        return workbook[workbook.sheetnames[0]]

    def month_info(self):
        year_override = config_value(self.config, "SCHEDULE_YEAR_OVERRIDE", "schedule_year_override")
        month_override = config_value(self.config, "SCHEDULE_MONTH_OVERRIDE", "schedule_month_override")
        if year_override and month_override:
            month_number = int(month_override)
            month_name = next(name for name, number in PERSIAN_MONTHS.items() if number == month_number)
            return int(year_override), month_number, month_name

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
        current_year, current_month, _ = parse_jalali_date(self.config.get("current_jalali_date"))
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
        professor_name = normalize_text(professor_name)
        schedule_year, schedule_month, _ = self.month_info()
        columns = self.professor_columns()
        col = columns.get(professor_name)
        if not col:
            compact_columns = {compact_text(name): col for name, col in columns.items()}
            col = compact_columns.get(compact_text(professor_name))
        if not col:
            return []
        classes = []
        for row in range(3, ws.max_row + 1):
            weekday = normalize_text(ws.cell(row, 1).value)
            day_number = ws.cell(row, 2).value
            student = normalize_text(ws.cell(row, col).value)
            if not weekday or not day_number or not student or student == "*":
                continue
            day_number = int(day_number)
            classes.append(
                {
                    "weekday": jalali_weekday_name(schedule_year, schedule_month, day_number),
                    "day": day_number,
                    "student": student,
                }
            )
        professor_key = compact_text(professor_name)
        for item in MANUAL_TEST_CLASSES:
            if compact_text(item.get("professor")) != professor_key:
                continue
            day_number = int(item["day"])
            classes.append(
                {
                    "weekday": jalali_weekday_name(schedule_year, schedule_month, day_number),
                    "day": day_number,
                    "student": normalize_text(item.get("student")),
                }
            )
        classes.sort(key=lambda item: int(item["day"]))
        return classes

    def schedule_for_day(self, professor_name, jalali_day):
        return [item for item in self.schedule_for_professor(professor_name) if item["day"] == int(jalali_day)]


class ContactDirectory:
    def __init__(self, export_url):
        self.cache = WorkbookCache(export_url)

    def professor_phone_map(self):
        workbook = self.cache.workbook_data()
        worksheet = workbook[workbook.sheetnames[0]]
        result = {}
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            professor = normalize_text(row[1] if len(row) > 1 else "")
            phone = normalize_phone(row[2] if len(row) > 2 else "")
            if professor and phone:
                result[phone] = professor
        return result


def normalize_registry_rows(rows):
    registry = {}
    for row in rows:
        chat_id = normalize_id(row.get("chat_id"))
        if not chat_id:
            continue
        registry[str(chat_id)] = {
            "phone": normalize_phone(row.get("phone")),
            "professor": normalize_text(row.get("professor")),
        }
    return registry


def test_chat_registry():
    value = read_json(TEST_CHAT_REGISTRY_PATH, [])
    if isinstance(value, dict):
        rows = [
            {"chat_id": chat_id, **details}
            for chat_id, details in value.items()
            if isinstance(details, dict)
        ]
    else:
        rows = value
    return normalize_registry_rows(rows)


def chat_registry(config=None):
    if (config or {}).get("test_mode"):
        return test_chat_registry()
    export_url = config_or_env("CHAT_REGISTRY_EXPORT_URL", DEFAULT_CHAT_REGISTRY_EXPORT_URL)
    return normalize_registry_rows(cached_sheet_rows(export_url))


def set_test_chat_professor(chat_id, phone, professor):
    rows = []
    current = test_chat_registry()
    current[normalize_id(chat_id)] = {
        "phone": normalize_phone(phone),
        "professor": normalize_text(professor),
    }
    for saved_chat_id, value in current.items():
        rows.append(
            {
                "chat_id": saved_chat_id,
                "phone": value.get("phone", ""),
                "professor": value.get("professor", ""),
            }
        )
    write_json(TEST_CHAT_REGISTRY_PATH, rows)


def set_chat_professor(chat_id, phone, professor, config=None):
    chat_id = normalize_id(chat_id)
    if (config or {}).get("test_mode"):
        set_test_chat_professor(chat_id, phone, professor)
        return
    export_url = config_or_env("CHAT_REGISTRY_EXPORT_URL", DEFAULT_CHAT_REGISTRY_EXPORT_URL)
    webhook_url = config_or_env("CHAT_REGISTRY_WEBHOOK_URL", DEFAULT_CHAT_REGISTRY_WEBHOOK_URL)
    post_sheet_action(
        webhook_url,
        {
            "action": "upsert_chat",
            "chat_id": chat_id,
            "phone": normalize_phone(phone),
            "professor": normalize_text(professor),
        },
    )
    invalidate_sheet_cache(export_url)


def get_chat_professor(chat_id, config=None):
    return normalize_text(chat_registry(config).get(normalize_id(chat_id), {}).get("professor"))


def get_chat_registration(chat_id, config=None):
    item = chat_registry(config).get(normalize_id(chat_id), {})
    return {
        "phone": normalize_phone(item.get("phone")),
        "professor": normalize_text(item.get("professor")),
    }


def filtered_chat_registry(config):
    return chat_registry(config)


def classes_keyboard(show_today_attendance=False):
    row = [{"text": VIEW_CLASSES_TEXT}]
    if show_today_attendance:
        row.append({"text": TODAY_ATTENDANCE_TEXT})
    return {
        "keyboard": [row],
        "resize_keyboard": True,
        "one_time_keyboard": False,
    }


def current_schedule_date(config):
    return parse_jalali_date(config.get("current_jalali_date"))


def professor_has_class_today(config, schedule_reader, professor):
    year, month, day = current_schedule_date(config)
    sheet_year, sheet_month, _ = schedule_reader.month_info()
    if year != sheet_year or month != sheet_month:
        return False
    return bool(schedule_reader.schedule_for_day(professor, day))


def professor_keyboard(config, schedule_reader, professor):
    return classes_keyboard(professor_has_class_today(config, schedule_reader, professor))


def role_keyboard():
    return {
        "keyboard": [[{"text": ROLE_STUDENT_TEXT}, {"text": ROLE_PROFESSOR_TEXT}]],
        "resize_keyboard": True,
        "one_time_keyboard": True,
    }


def reselect_role_keyboard():
    return {
        "keyboard": [[{"text": RESELECT_ROLE_TEXT}]],
        "resize_keyboard": True,
        "one_time_keyboard": False,
    }


def contact_keyboard():
    return {
        "keyboard": [[{"text": SHARE_CONTACT_TEXT, "request_contact": True}]],
        "resize_keyboard": True,
        "one_time_keyboard": True,
    }


def attendance_inline_keyboard(attendance_key):
    return {
        "inline_keyboard": [
            [
                {"text": ATTENDANCE_ABSENT_TEXT, "callback_data": f"att:absent:{attendance_key}"},
                {"text": ATTENDANCE_WEAK_TEXT, "callback_data": f"att:weak:{attendance_key}"},
            ],
            [
                {"text": ATTENDANCE_MEDIUM_TEXT, "callback_data": f"att:medium:{attendance_key}"},
                {"text": ATTENDANCE_EXCELLENT_TEXT, "callback_data": f"att:excellent:{attendance_key}"},
            ]
        ]
    }


def remove_keyboard():
    return {"remove_keyboard": True}


def format_phone(phone):
    return to_persian_digits(normalize_phone(phone))


def sent_message_id(response):
    result = response.get("result") if isinstance(response, dict) else None
    if isinstance(result, dict):
        return result.get("message_id")
    return None


def send_processing_message(bot, chat_id):
    return bot.send_message(chat_id, PROCESSING_MESSAGE)


def delete_processing_message(bot, chat_id, response):
    message_id = sent_message_id(response)
    if not message_id:
        return
    try:
        bot.delete_message(chat_id, message_id)
    except Exception as exc:
        print(f"Processing message delete failed: {exc}")


def format_schedule(professor_name, schedule, schedule_reader):
    year, _, month_name = schedule_reader.month_info()
    if not schedule:
        return f"برای {professor_name} در این ماه برنامه‌ای ثبت نشده است."
    lines = []
    for index, item in enumerate(schedule, start=1):
        lines.append(
            f"{index}. {item['weekday']} {to_persian_digits(item['day'])} {month_name} {to_persian_digits(year)}\n"
            f"   دانشجو: {item['student']}"
        )
        lines.append("")
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
        f"با عرض سلام , یادآوری برنامه امروز {professor_name}",
        f"{day_schedule[0]['weekday']} {to_persian_digits(day)} {month_name} {to_persian_digits(year)}",
        "",
    ]
    for index, item in enumerate(day_schedule, start=1):
        lines.append(f"{index}. دانشجو: {item['student']}")
    return "\n".join(lines)


def format_attendance_question(item):
    return (
        "با عرض سلام و خسته نباشید خدمت استاد گرامی\n"
        f"اینترن تاریخ {item['date']} شما دکتر {item['student']} بوده است .\n"
        "در صورت عدم حضور دکمه ی غیبت و در صورت حضور عملکرد ایشان را ارزیابی و انتخاب بفرمایید :"
    )


def pending_attendance():
    export_url = config_or_env("PENDING_ATTENDANCE_EXPORT_URL", DEFAULT_PENDING_ATTENDANCE_EXPORT_URL)
    pending = {}
    for row in cached_sheet_rows(export_url):
        chat_id = normalize_id(row.get("chat_id"))
        key = normalize_text(row.get("key"))
        if not chat_id or not key:
            continue
        pending.setdefault(str(chat_id), {"items": {}})
        pending[str(chat_id)]["items"][key] = {
            "professor": normalize_text(row.get("professor_name")),
            "student": normalize_text(row.get("student_name")),
            "date": normalize_text(row.get("date")),
            "key": key,
        }
    return pending


def pending_attendance_key_exists(chat_id, key):
    export_url = config_or_env("PENDING_ATTENDANCE_EXPORT_URL", DEFAULT_PENDING_ATTENDANCE_EXPORT_URL)
    invalidate_sheet_cache(export_url)
    state = pending_attendance().get(normalize_id(chat_id), {})
    return normalize_text(key) in state.get("items", {})


def config_value_or_env(config, env_name, config_name, default):
    return os.getenv(env_name, (config or {}).get(config_name, default)).strip()


def attendance_log(config=None):
    export_url = config_value_or_env(
        config,
        "ATTENDANCE_LOG_EXPORT_URL",
        "attendance_log_export_url",
        DEFAULT_ATTENDANCE_LOG_EXPORT_URL,
    )
    logged = set()
    try:
        rows = cached_sheet_rows(export_url)
    except Exception as exc:
        print(f"Attendance log read failed, continuing without sent-log cache: {exc}")
        return logged
    for row in rows:
        chat_id = normalize_id(row.get("chat_id"))
        date_text = normalize_text(row.get("date"))
        professor = normalize_text(row.get("professor"))
        student = normalize_text(row.get("student"))
        if chat_id and date_text and professor and student:
            logged.add(attendance_log_key(chat_id, date_text, professor, student))
    return logged


def attendance_log_key(chat_id, date_text, professor, student):
    return ":".join(
        [
            normalize_id(chat_id),
            normalize_text(date_text),
            compact_text(professor),
            compact_text(student),
        ]
    )


def upsert_attendance_log(config, chat_id, item, status="", sent_at="", answered_at=""):
    webhook_url = config_value_or_env(
        config,
        "ATTENDANCE_LOG_WEBHOOK_URL",
        "attendance_log_webhook_url",
        DEFAULT_ATTENDANCE_LOG_WEBHOOK_URL,
    )
    if not webhook_url:
        return False
    export_url = config_value_or_env(
        config,
        "ATTENDANCE_LOG_EXPORT_URL",
        "attendance_log_export_url",
        DEFAULT_ATTENDANCE_LOG_EXPORT_URL,
    )
    post_sheet_action_with_retry(
        webhook_url,
        {
            "action": "upsert_attendance_log",
            "date": item["date"],
            "professor": item["professor"],
            "student": item["student"],
            "chat_id": normalize_id(chat_id),
            "sent_at": sent_at,
            "answered_at": answered_at,
            "status": status,
        },
    )
    invalidate_sheet_cache(export_url)
    return True


def now_iso(config):
    return datetime.now(get_bot_timezone(config.get("bot_timezone", BOT_TIMEZONE))).isoformat(timespec="seconds")


def upsert_pending_attendance(chat_id, item):
    chat_id = normalize_id(chat_id)
    export_url = config_or_env("PENDING_ATTENDANCE_EXPORT_URL", DEFAULT_PENDING_ATTENDANCE_EXPORT_URL)
    webhook_url = config_or_env("PENDING_ATTENDANCE_WEBHOOK_URL", DEFAULT_PENDING_ATTENDANCE_WEBHOOK_URL)
    post_sheet_action(
        webhook_url,
        {
            "action": "upsert_pending",
            "chat_id": chat_id,
            "key": item["key"],
            "professor_name": item["professor"],
            "student_name": item["student"],
            "date": item["date"],
        },
    )
    invalidate_sheet_cache(export_url)


def delete_pending_attendance(chat_id, key):
    chat_id = normalize_id(chat_id)
    export_url = config_or_env("PENDING_ATTENDANCE_EXPORT_URL", DEFAULT_PENDING_ATTENDANCE_EXPORT_URL)
    webhook_url = config_or_env("PENDING_ATTENDANCE_WEBHOOK_URL", DEFAULT_PENDING_ATTENDANCE_WEBHOOK_URL)
    post_sheet_action(
        webhook_url,
        {
            "action": "delete_pending",
            "chat_id": chat_id,
            "key": key,
        },
    )
    invalidate_sheet_cache(export_url)


def local_absences():
    return read_json(ABSENCES_FALLBACK_PATH, [])


def append_local_absence(absence):
    items = local_absences()
    items.append(absence)
    write_json(ABSENCES_FALLBACK_PATH, items)


def append_absence_to_excel(absence):
    ABSENCES_XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    if ABSENCES_XLSX_PATH.exists():
        workbook = load_workbook(ABSENCES_XLSX_PATH)
        worksheet = workbook.active
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "غیبت دانشجویان"
        worksheet.append(["ردیف", "نام دانشجو", "نام استاد", "تاریخ", "وضعیت"])

    next_number = max(1, worksheet.max_row)
    worksheet.append(
        [
            next_number,
            absence["student"],
            absence["professor"],
            absence["date"],
            absence.get("status", ATTENDANCE_ABSENT_TEXT),
        ]
    )
    workbook.save(ABSENCES_XLSX_PATH)


def append_absence_to_webhook(config, absence):
    webhook_url = str(config.get("absence_webhook_url", "")).strip()
    if not webhook_url:
        return False
    response = requests.post(
        webhook_url,
        json={
            "student": absence["student"],
            "professor": absence["professor"],
            "date": absence["date"],
            "status": absence.get("status", ATTENDANCE_ABSENT_TEXT),
        },
        timeout=30,
    )
    response.raise_for_status()
    try:
        result = response.json()
    except ValueError as exc:
        preview = response.text[:500] if response.text else "<empty response>"
        raise RuntimeError(f"Absence webhook returned non-JSON response: {preview}") from exc
    if not result.get("ok"):
        raise RuntimeError(result.get("error") or "Absence webhook returned ok=false")
    return True


def record_absence(config, absence):
    try:
        if append_absence_to_webhook(config, absence):
            return
    except Exception as exc:
        print(f"Apps Script absence write failed: {exc}")

    try:
        append_absence_to_excel(absence)
        return
    except Exception as exc:
        print(f"Excel absence write failed: {exc}")
    append_local_absence(absence)


ATTENDANCE_STATUS_BY_CODE = {
    "absent": ATTENDANCE_ABSENT_TEXT,
    "weak": ATTENDANCE_WEAK_TEXT,
    "medium": ATTENDANCE_MEDIUM_TEXT,
    "excellent": ATTENDANCE_EXCELLENT_TEXT,
}


def complete_attendance_callback(bot, config, schedule_reader, chat_id, attendance_id, status):
    chat_id = normalize_id(chat_id)
    pending = pending_attendance()
    state = pending.get(chat_id, {})
    items = state.get("items", {})
    active = items.get(attendance_id)
    if not active:
        return False

    active["status"] = status
    record_absence(config, active)
    try:
        upsert_attendance_log(
            config,
            chat_id,
            active,
            status=status,
            answered_at=datetime.now(get_bot_timezone(config.get("bot_timezone", BOT_TIMEZONE))).isoformat(timespec="seconds"),
        )
    except Exception as exc:
        print(f"Attendance log answer write failed: {exc}")
    keyboard = professor_keyboard(config, schedule_reader, active["professor"])
    if status == ATTENDANCE_ABSENT_TEXT:
        bot.send_message(chat_id, "غیبت دانشجو ثبت شد با تشکر", keyboard=keyboard)
    else:
        bot.send_message(chat_id, "عملکرد اینترن شما ثبت شد با تشکر", keyboard=keyboard)

    items.pop(attendance_id, None)
    try:
        delete_pending_attendance(chat_id, attendance_id)
    except Exception as exc:
        print(f"Pending attendance sheet delete failed, using local JSON fallback: {exc}")
    if items:
        state["items"] = items
        pending[chat_id] = state
    else:
        pending.pop(chat_id, None)
    return True


def handle_callback_query(bot, config, schedule_reader, callback_query):
    data = str(callback_query.get("data") or "")
    message = callback_query.get("message") or {}
    chat_id = (message.get("chat") or {}).get("id")
    if not data.startswith("att:") or chat_id is None:
        return

    try:
        _, answer_code, attendance_id = data.split(":", 2)
    except ValueError:
        return
    status = ATTENDANCE_STATUS_BY_CODE.get(answer_code)
    if not status:
        return
    if complete_attendance_callback(bot, config, schedule_reader, chat_id, attendance_id, status):
        return
    registration = get_chat_registration(chat_id, config)
    bot.send_message(
        chat_id,
        "این مورد قبلا ثبت شده",
        keyboard=professor_keyboard(config, schedule_reader, registration["professor"]),
    )


def handle_message(bot, config, schedule_reader, contact_directory, message):
    chat_id = (message.get("chat") or {}).get("id")
    sender_id = normalize_id((message.get("from") or {}).get("id"))
    text = normalize_text(message.get("text"))
    contact = message.get("contact") or {}
    contact_phone = normalize_phone(contact.get("phone_number") or contact.get("phone"))
    contact_user_id = normalize_id(contact.get("user_id"))
    if chat_id is None:
        return

    registration = get_chat_registration(chat_id, config)
    registered_phone = registration["phone"]
    registered_professor = registration["professor"]

    if contact_phone and contact_user_id and sender_id and contact_user_id != sender_id:
        bot.send_message(chat_id, CONTACT_OWNER_MISMATCH_MESSAGE, keyboard=contact_keyboard())
        return

    if text in {"/start", RESELECT_ROLE_TEXT}:
        bot.send_message(chat_id, START_MESSAGE, keyboard=role_keyboard())
        return

    if text == ROLE_STUDENT_TEXT:
        bot.send_message(chat_id, STUDENT_SECTION_MESSAGE, keyboard=reselect_role_keyboard())
        return

    if text == ROLE_PROFESSOR_TEXT:
        bot.send_message(chat_id, PROFESSOR_INTRO_MESSAGE, keyboard=contact_keyboard())
        return

    if text == VIEW_CLASSES_TEXT:
        if registered_phone and registered_professor:
            processing_message = send_processing_message(bot, chat_id)
            schedule = schedule_reader.schedule_for_professor(registered_professor)
            delete_processing_message(bot, chat_id, processing_message)
            bot.send_message(
                chat_id,
                format_login_schedule(registered_professor, schedule, schedule_reader),
                keyboard=professor_keyboard(config, schedule_reader, registered_professor),
            )
            return
        bot.send_message(chat_id, START_MESSAGE, keyboard=role_keyboard())
        return

    if text == TODAY_ATTENDANCE_TEXT:
        if registered_phone and registered_professor:
            processing_message = send_processing_message(bot, chat_id)
            if professor_has_class_today(config, schedule_reader, registered_professor):
                messages = send_attendance_questions_for_professor(
                    bot,
                    config,
                    schedule_reader,
                    chat_id,
                    registered_professor,
                    target_date=config.get("current_jalali_date"),
                    send_mode="manual",
                )
                delete_processing_message(bot, chat_id, processing_message)
                if not messages:
                    bot.send_message(
                        chat_id,
                        "حضور و غیاب امروز قبلا ارسال شده است.",
                        keyboard=professor_keyboard(config, schedule_reader, registered_professor),
                    )
                return
            delete_processing_message(bot, chat_id, processing_message)
            bot.send_message(
                chat_id,
                "برای امروز کلاسی ثبت نشده است.",
                keyboard=professor_keyboard(config, schedule_reader, registered_professor),
            )
            return
        bot.send_message(chat_id, START_MESSAGE, keyboard=role_keyboard())
        return

    if registered_phone and registered_professor:
        if contact_phone and contact_phone != registered_phone:
            bot.send_message(
                chat_id,
                REGISTERED_PHONE_MESSAGE.format(phone=format_phone(registered_phone)),
                keyboard=professor_keyboard(config, schedule_reader, registered_professor),
            )
            return
        processing_message = send_processing_message(bot, chat_id)
        schedule = schedule_reader.schedule_for_professor(registered_professor)
        delete_processing_message(bot, chat_id, processing_message)
        bot.send_message(
            chat_id,
            format_login_schedule(registered_professor, schedule, schedule_reader),
            keyboard=professor_keyboard(config, schedule_reader, registered_professor),
        )
        return

    processing_message = send_processing_message(bot, chat_id) if contact_phone else None
    phones = contact_directory.professor_phone_map()
    if contact_phone in phones:
        professor = phones[contact_phone]
        set_chat_professor(chat_id, contact_phone, professor, config)
        schedule = schedule_reader.schedule_for_professor(professor)
        delete_processing_message(bot, chat_id, processing_message)
        bot.send_message(
            chat_id,
            format_login_schedule(professor, schedule, schedule_reader),
            keyboard=professor_keyboard(config, schedule_reader, professor),
        )
        return

    if text and not contact_phone:
        bot.send_message(chat_id, CONTACT_REQUIRED_MESSAGE, keyboard=role_keyboard())
        return

    delete_processing_message(bot, chat_id, processing_message)
    bot.send_message(chat_id, INVALID_PHONE_MESSAGE, keyboard=contact_keyboard())


def reminder_key(chat_id, year, month, day):
    return f"{chat_id}:{year:04d}-{month:02d}-{day:02d}"


def send_due_reminders(bot, config, schedule_reader, target_date=None, dry_run=False):
    year, month, day = parse_jalali_date(target_date)
    sheet_year, sheet_month, _ = schedule_reader.month_info()
    if year != sheet_year or month != sheet_month:
        return []
    messages = []
    for chat_id, value in filtered_chat_registry(config).items():
        professor = normalize_text(value.get("professor"))
        if not professor:
            continue
        day_schedule = schedule_reader.schedule_for_day(professor, day)
        if not day_schedule:
            continue
        key = reminder_key(chat_id, year, month, day)
        if key in REMINDERS_SENT_THIS_RUN:
            continue
        text = format_reminder(professor, day_schedule, schedule_reader)
        messages.append({"chat_id": chat_id, "professor": professor, "text": text})
        if not dry_run:
            bot.send_message(chat_id, text)
            REMINDERS_SENT_THIS_RUN.add(key)
    return messages


def parse_jalali_datetime(value):
    value = normalize_digits(value).strip()
    if not value:
        return None
    match = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,3})\s+(\d{1,2}):(\d{2})", value)
    if not match:
        raise ValueError(f"Invalid MONTHLY_SCHEDULE_SEND_AT value: {value}")
    year, month, day, hour, minute = (int(part) for part in match.groups())
    if not (1 <= month <= 12 and 1 <= day <= 31 and 0 <= hour <= 23 and 0 <= minute <= 59):
        raise ValueError(f"Invalid MONTHLY_SCHEDULE_SEND_AT value: {value}")
    return year, month, day, hour, minute


def current_jalali_from_datetime(now):
    return gregorian_to_jalali(now.year, now.month, now.day)


def monthly_schedule_key(chat_id, year, month, day, hour, minute):
    return f"{chat_id}:{year:04d}-{month:02d}-{day:02d}:{hour:02d}:{minute:02d}"


def send_due_monthly_schedules(bot, config, schedule_reader, send_at, now, dry_run=False):
    parsed = parse_jalali_datetime(send_at)
    if not parsed:
        return []
    year, month, day, hour, minute = parsed
    current_year, current_month, current_day = current_jalali_from_datetime(now)
    if (current_year, current_month, current_day, now.hour, now.minute) != (year, month, day, hour, minute):
        return []

    messages = []
    for chat_id, value in filtered_chat_registry(config).items():
        professor = normalize_text(value.get("professor"))
        if not professor:
            continue
        key = monthly_schedule_key(chat_id, year, month, day, hour, minute)
        if key in MONTHLY_SCHEDULE_SENT_THIS_RUN:
            continue
        schedule = schedule_reader.schedule_for_professor(professor)
        text = format_login_schedule(professor, schedule, schedule_reader)
        messages.append({"chat_id": chat_id, "professor": professor, "text": text})
        if not dry_run:
            bot.send_message(chat_id, text, keyboard=professor_keyboard(config, schedule_reader, professor))
            MONTHLY_SCHEDULE_SENT_THIS_RUN.add(key)
    return messages


def attendance_key(chat_id, year, month, day, student):
    digest = hashlib.sha1(compact_text(student).encode("utf-8")).hexdigest()[:12]
    return f"{chat_id}-{year:04d}{month:02d}{day:02d}-{digest}"


def jalali_date_text(year, month, day):
    month_text = to_persian_digits(f"{month:02d}")
    day_text = to_persian_digits(f"{day:02d}")
    return f"{to_persian_digits(year)}/{month_text}/{day_text}"


def send_attendance_questions_for_professor(
    bot,
    config,
    schedule_reader,
    chat_id,
    professor,
    target_date=None,
    pending=None,
    logged_attendance=None,
    send_mode="automatic",
    dry_run=False,
):
    year, month, day = parse_jalali_date(target_date)
    sheet_year, sheet_month, _ = schedule_reader.month_info()
    if year != sheet_year or month != sheet_month:
        return []

    chat_id = normalize_id(chat_id)
    professor = normalize_text(professor)
    pending = pending if pending is not None else pending_attendance()
    logged_attendance = logged_attendance if logged_attendance is not None else attendance_log(config)
    state = pending.get(chat_id, {})
    pending_items = state.get("items", {})
    day_schedule = schedule_reader.schedule_for_day(professor, day)
    messages = []

    for item in day_schedule:
        key = attendance_key(chat_id, year, month, day, item["student"])
        absence_item = {
            "professor": professor,
            "student": item["student"],
            "date": jalali_date_text(year, month, day),
            "key": key,
        }
        log_key = attendance_log_key(chat_id, absence_item["date"], professor, item["student"])
        if key in ATTENDANCE_SENT_THIS_RUN or key in pending_items or log_key in logged_attendance:
            continue
        if dry_run:
            messages.append({"chat_id": chat_id, **absence_item})
            continue
        pending_saved = False
        last_pending_error = None
        for attempt in range(1, SHEET_ACTION_RETRY_ATTEMPTS + 1):
            try:
                upsert_pending_attendance(chat_id, absence_item)
                pending_saved = True
                break
            except Exception as exc:
                last_pending_error = exc
                try:
                    if pending_attendance_key_exists(chat_id, key):
                        print(
                            "Pending attendance write returned an error, "
                            "but the key exists in the sheet; continuing."
                        )
                        pending_saved = True
                        break
                except Exception as verify_exc:
                    print(f"Pending attendance verification failed: {verify_exc}")
                if attempt < SHEET_ACTION_RETRY_ATTEMPTS:
                    print(
                        "Pending attendance sheet write failed "
                        f"(attempt {attempt}/{SHEET_ACTION_RETRY_ATTEMPTS}); "
                        f"retrying in {SHEET_ACTION_RETRY_DELAY_SECONDS}s: {exc}"
                    )
                    time.sleep(SHEET_ACTION_RETRY_DELAY_SECONDS)
        if not pending_saved:
            print(f"Pending attendance sheet write failed; attendance message was not sent: {last_pending_error}")
            continue
        pending_items[key] = absence_item
        try:
            bot.send_message(
                chat_id,
                format_attendance_question(absence_item),
                keyboard=attendance_inline_keyboard(key),
            )
        except Exception:
            pending_items.pop(key, None)
            try:
                delete_pending_attendance(chat_id, key)
            except Exception as exc:
                print(f"Pending attendance cleanup failed after Bale send failure: {exc}")
            raise
        ATTENDANCE_SENT_THIS_RUN.add(key)
        messages.append({"chat_id": chat_id, **absence_item})
        try:
            if upsert_attendance_log(
                config,
                chat_id,
                absence_item,
                sent_at=datetime.now(get_bot_timezone(config.get("bot_timezone", BOT_TIMEZONE))).isoformat(timespec="seconds"),
            ):
                logged_attendance.add(log_key)
        except Exception as exc:
            print(f"Attendance log sent write failed: {exc}")

    if pending_items:
        pending[chat_id] = {"items": pending_items}
    return messages


def send_due_attendance_questions(bot, config, schedule_reader, target_date=None, dry_run=False):
    year, month, day = parse_jalali_date(target_date)
    sheet_year, sheet_month, _ = schedule_reader.month_info()
    if year != sheet_year or month != sheet_month:
        return []

    pending = pending_attendance()
    logged_attendance = attendance_log(config)
    messages = []

    for chat_id, value in filtered_chat_registry(config).items():
        professor = normalize_text(value.get("professor"))
        if not professor:
            continue
        messages.extend(
            send_attendance_questions_for_professor(
                bot,
                config,
                schedule_reader,
                chat_id,
                professor,
                target_date=target_date,
                pending=pending,
                logged_attendance=logged_attendance,
                dry_run=dry_run,
            )
        )

    return messages


def reminder_loop(bot, config, schedule_reader):
    reminder_time = config.get("reminder_time", "09:00")
    reminder_window_minutes = int(config.get("reminder_window_minutes", REMINDER_WINDOW_MINUTES))
    attendance_time = config.get("attendance_time", "21:00")
    attendance_window_minutes = int(config.get("attendance_window_minutes", ATTENDANCE_WINDOW_MINUTES))
    monthly_schedule_send_at = config.get("monthly_schedule_send_at", MONTHLY_SCHEDULE_SEND_AT)
    target_date = config.get("current_jalali_date")
    interval_seconds = int(config.get("scheduler_interval_seconds", SCHEDULER_INTERVAL_SECONDS))
    timezone = get_bot_timezone(config.get("bot_timezone", BOT_TIMEZONE))
    print(
        "Scheduler is using "
        f"{timezone_label(timezone)}; now={datetime.now(timezone).strftime('%Y-%m-%d %H:%M:%S')}; "
        f"reminder={reminder_time}; reminder_window={reminder_window_minutes}m; "
        f"attendance={attendance_time}; "
        f"attendance_window={attendance_window_minutes}m; "
        f"monthly_schedule={monthly_schedule_send_at}; interval={interval_seconds}s"
    )
    while True:
        try:
            now = datetime.now(timezone)
            send_due_monthly_schedules(bot, config, schedule_reader, monthly_schedule_send_at, now)
            if time_is_in_window(now, reminder_time, reminder_window_minutes):
                send_due_reminders(bot, config, schedule_reader, target_date=target_date)
            if time_is_in_window(now, attendance_time, attendance_window_minutes):
                send_due_attendance_questions(bot, config, schedule_reader, target_date=target_date)
        except Exception as exc:
            print(f"Scheduler error: {exc}")
        time.sleep(interval_seconds)


def run(test_mode=False):
    config = load_config()
    config["test_mode"] = bool(test_mode)
    start_health_server()
    bot = BaleBot(config)
    schedule_reader = SheetSchedule(config["sheet_export_url"], config)
    contact_directory = ContactDirectory(config["contacts_export_url"])
    Thread(target=reminder_loop, args=(bot, config, schedule_reader), daemon=True).start()
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
                    handle_message(bot, config, schedule_reader, contact_directory, message)
                callback_query = update.get("callback_query")
                if callback_query:
                    handle_callback_query(bot, config, schedule_reader, callback_query)
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
    config["test_mode"] = bool(getattr(args, "test", False))
    reader = SheetSchedule(config["sheet_export_url"], config)
    professor = ContactDirectory(config["contacts_export_url"]).professor_phone_map().get(normalize_phone(args.phone))
    if not professor:
        print(INVALID_PHONE_MESSAGE)
        return
    print(format_schedule(professor, reader.schedule_for_professor(professor), reader))


def preview_reminders(args):
    config = load_config()
    config["test_mode"] = bool(getattr(args, "test", False))
    reader = SheetSchedule(config["sheet_export_url"], config)
    if args.phone:
        professor = ContactDirectory(config["contacts_export_url"]).professor_phone_map().get(normalize_phone(args.phone))
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

    messages = send_due_reminders(None, config, reader, target_date=args.date, dry_run=True)
    if not messages:
        print("No reminders found for this date.")
        return
    for message in messages:
        print(f"CHAT_ID: {message['chat_id']}")
        print(message["text"])
        print("-" * 30)


def preview_attendance(args):
    config = load_config()
    config["test_mode"] = bool(getattr(args, "test", False))
    reader = SheetSchedule(config["sheet_export_url"], config)
    messages = send_due_attendance_questions(None, config, reader, target_date=args.date, dry_run=True)
    if not messages:
        print("No attendance questions found for this date.")
        return
    for message in messages:
        print(f"CHAT_ID: {message['chat_id']}")
        print(format_attendance_question(message))
        print("-" * 30)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="Use local test_chat_registry.json instead of chat-id Google Sheet")
    subparsers = parser.add_subparsers(dest="command")
    schedule_parser = subparsers.add_parser("preview-schedule")
    schedule_parser.add_argument("--phone", required=True)
    reminder_parser = subparsers.add_parser("test-reminders")
    reminder_parser.add_argument("--date", required=True, help="Jalali date, for example 1405-04-01")
    reminder_parser.add_argument("--phone", help="Optional professor phone for testing before Bale login")
    attendance_parser = subparsers.add_parser("test-attendance")
    attendance_parser.add_argument("--date", required=True, help="Jalali date, for example 1405-04-01")
    args = parser.parse_args()
    if args.command == "preview-schedule":
        preview_schedule(args)
    elif args.command == "test-reminders":
        preview_reminders(args)
    elif args.command == "test-attendance":
        preview_attendance(args)
    else:
        run(test_mode=args.test)


if __name__ == "__main__":
    main()
