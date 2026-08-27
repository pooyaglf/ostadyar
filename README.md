# Ostad Yar Bale Bot

Bot link: https://ble.ir/Ostad_YarBot

The bot uses the Google Sheet as the live schedule database:

```text
https://docs.google.com/spreadsheets/d/1sDIbSkFHlgsqxrYZyK2diG53eh4LT09h/export?format=xlsx
```

## Professor Flow

1. Professor sends `/start`.
2. Bot asks for mobile number.
3. Bot checks the number in `data/professor_phones.json`.
4. If the number exists, the bot stores the professor against the Bale `chat_id` in `data/chat_ids.json`.
5. Bot reads the Google Sheet and sends only that professor's student schedule.
6. Empty cells and `*` cells are skipped.
7. On class day, the bot sends a morning reminder to that professor's saved chat ID.

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
SHEET_EXPORT_URL=https://docs.google.com/spreadsheets/d/1sDIbSkFHlgsqxrYZyK2diG53eh4LT09h/export?format=xlsx
API_BASE_URL=https://tapi.bale.ai/bot
POLL_TIMEOUT_SECONDS=25
SHEET_CACHE_SECONDS=300
REMINDER_HOUR=8
DATA_DIR=/app/data
```

Keep Hamravesh replicas at `1`, because the Bale polling bot must not run twice.

## Edit Professor Phone Numbers

Edit:

```text
data/professor_phones.json
```

The names must match the professor names in row 2 of the Google Sheet.

Example:

```json
{
  "دکتر امیر محمد آرمانیان": "09133881014"
}
```

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

## Deploy Update

After changing files:

```powershell
git add .
git commit -m "Use Google Sheet schedule database"
git push
```

Hamravesh will redeploy automatically if auto deploy is enabled. Otherwise, run a manual deploy from the Hamravesh panel.
