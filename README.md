# Ostad Yar Bale Bot

Bot link: https://ble.ir/Ostad_YarBot

## Run

```powershell
cd C:\Users\Laptopkaran\Desktop\ostadyar
$env:BOT_TOKEN="YOUR_BALE_BOT_TOKEN"
python bot.py
```

The bot uses long polling, so keep the terminal open while you want the bot to answer users.

## Deploy Environment Variables

Set these variables in Hamravesh:

```text
BOT_TOKEN=your Bale bot token
PORT=8000
```

Optional:

```text
API_BASE_URL=https://tapi.bale.ai/bot
POLL_TIMEOUT_SECONDS=25
CHAT_IDS_PATH=/app/data/chat_ids.json
```

For production, use persistent storage or a database for `CHAT_IDS_PATH` if you need chat IDs to survive redeploys.

## Docker

```bash
docker build -t ostadyar .
docker run --rm -e BOT_TOKEN="YOUR_BALE_BOT_TOKEN" -p 8000:8000 ostadyar
```

## Edit Professors

Edit `data/professors.json`.

Each professor in this file automatically becomes a button when the bot sends the start message. If you remove a professor from this file, that button disappears automatically the next time the bot sends the keyboard.

## Edit Students

Edit `data/students.json`.

Use this format:

```json
{
  "استاد بهرامیان": [
    {
      "name": "شایان اکبران",
      "student_id": "455555555"
    }
  ]
}
```

The professor names in `data/students.json` should match the names in `data/professors.json`.

## Chat IDs

`data/chat_ids.json` starts with:

```json
[
  "1581433567",
  "547772131"
]
```

Every user who sends a message or starts the bot is added automatically to this file.
