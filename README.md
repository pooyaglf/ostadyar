# Ostad Yar Bale Bot

Bot link: https://ble.ir/Ostad_YarBot

The bot uses the Google Sheet as the live schedule database:

```text
https://docs.google.com/spreadsheets/d/1jwQ-2k6zbOGLTgPjSpOXGvnlBRWm70Mk/export?format=xlsx
```

For local testing, this populated sheet is still used as the default schedule source. The empty production schedule sheet can be used later by setting `SHEET_EXPORT_URL`.

Professor phone numbers and absence records use this spreadsheet:

```text
https://docs.google.com/spreadsheets/d/1P_wWkcMIpsUZYME8xCllQRvjfHSGCa0s/export?format=xlsx
```

## Professor Flow

1. Professor sends `/start`.
2. Bot asks for mobile number.
3. Bot checks the number in the `شماره تماس اساتید` sheet.
4. If the number exists, the bot stores one professor phone against the Bale `chat_id` in `data/chat_ids.json`.
5. A chat ID cannot change to another professor phone. If another number is sent later, the bot rejects it and shows the saved number.
6. Bot reads the Google Sheet and sends only that professor's student schedule.
7. Empty cells and `*` cells are skipped.
8. After login, the bot shows one button for viewing the class schedule again.
9. On class day, the bot sends a morning reminder to that professor's saved chat ID at 9 AM.
10. At 9 PM, the bot asks whether the student attended class. If the answer is `خیر`, the absence is written to `data/absences.xlsx`.

## Run Locally

```powershell
cd C:\Users\Laptopkaran\Desktop\ostadyar
$env:BOT_TOKEN="YOUR_BALE_BOT_TOKEN"
python bot.py
```

For local testing you can also keep `config.local.json`, but do not commit it.

## Hamravesh Environment Variables

Required:

```text
BOT_TOKEN=your Bale bot token
PORT=8000
```

Optional:

```text
SHEET_EXPORT_URL=https://docs.google.com/spreadsheets/d/1jwQ-2k6zbOGLTgPjSpOXGvnlBRWm70Mk/export?format=xlsx
CONTACTS_EXPORT_URL=https://docs.google.com/spreadsheets/d/1P_wWkcMIpsUZYME8xCllQRvjfHSGCa0s/export?format=xlsx
API_BASE_URL=https://tapi.bale.ai/bot
POLL_TIMEOUT_SECONDS=25
SHEET_CACHE_SECONDS=300
DATA_DIR=/app/data
ABSENCE_WEBHOOK_URL=https://script.google.com/macros/s/AKfycbw6Mn9mGMjkUsQKaLRiY6MDV22Xc6jtWB4BPpzJo3vQk7rvr1wC6h-ZfQQwD89FECo/exec
BOT_TIMEZONE=Asia/Tehran
```

To change the schedule month/year and reminder times, edit the settings at the top of `bot.py`:

```python
SCHEDULE_YEAR_OVERRIDE = ""
SCHEDULE_MONTH_OVERRIDE = ""
CURRENT_JALALI_DATE_OVERRIDE = ""
REMINDER_TIME = "09:00"
ATTENDANCE_TIME = "21:00"
SCHEDULER_INTERVAL_SECONDS = 15
BOT_TIMEZONE = "Asia/Tehran"
```

Leave `SCHEDULE_YEAR_OVERRIDE` and `SCHEDULE_MONTH_OVERRIDE` empty to read month/year from the class sheet title. `REMINDER_TIME` controls the morning reminder. `ATTENDANCE_TIME` controls the attendance question. Use `HH:MM`, for example `14:35`. `BOT_TIMEZONE` makes those times independent of the server timezone.

Keep Hamravesh replicas at `1`, because the Bale polling bot must not run twice.

Absences are saved to the Apps Script webhook first:

```text
ABSENCE_WEBHOOK_URL
```

If that write fails, the bot saves locally in:

```text
data/absences.xlsx
```

The file uses this format:

```text
ردیف | نام دانشجو | نام استاد | تاریخ غیبت
```

The bot still contains fallback code for Google Sheet writes, but local Excel is the primary save path now.

## Test Schedule Parsing

Preview a professor schedule by phone number:

```powershell
python bot.py preview-schedule --phone 09133881014
```

## Test Reminder Logic

First, make sure `data/chat_ids.json` contains a chat mapped to a professor. This happens automatically after the professor enters a valid phone number in Bale.

Then preview reminders for a Jalali date:

```powershell
python bot.py test-reminders --date 1405-04-01
```

This command does not send messages. It only prints the reminder messages that would be sent.

You can also test one professor before they log in to Bale:

```powershell
python bot.py test-reminders --date 1405-04-06 --phone 09133881014
```

Preview the 9 PM attendance questions:

```powershell
python bot.py test-attendance --date 1405-04-06
```

## Deploy Update

After changing files:

```powershell
git add .
git commit -m "Use Google Sheet schedule database"
git push
```

Hamravesh will redeploy automatically if auto deploy is enabled. Otherwise, run a manual deploy from the Hamravesh panel.
