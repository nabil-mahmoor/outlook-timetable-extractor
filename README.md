# Outlook Timetable Extractor

A Python automation script that queries your Outlook mailbox, finds the latest timetable email, extracts your specific batch's timetable page from the PDF attachment, and saves it as an image — automatically.

---

## How It Works

1. Authenticates with your Outlook mailbox via Microsoft Graph API
2. Finds the latest email with "Timetable" in the subject
3. Downloads the PDF attachment
4. Scans each page for keywords unique to your batch and degree
5. Renders that page as a dated PNG image
6. Opens it automatically on your PC
7. Deletes the temporary PDF
8. *(Optional)* Sends the image to your phone via Telegram

---

## Project Structure

```
outlook-timetable-extractor/
├── .venv/                  # Virtual environment
├── output/                 # Saved timetable images land here
├── auth.py                 # OAuth2 authentication and token management
├── mail.py                 # Graph API queries — finding email and downloading attachment
├── pdf_handler.py          # PDF page detection and image rendering
├── config.py               # All constants and settings
├── main.py                 # Orchestrates the full workflow
├── notifier.py             # (Optional) Telegram delivery
├── run.bat                 # Windows batch file for Task Scheduler
├── token_cache.json        # Auto-generated on first login, do not commit
├── .env                    # Secret credentials, do not commit
├── .gitignore
└── pyproject.toml
```

---

## Prerequisites

- Python 3.13+
- [uv](https://github.com/astral-sh/uv) package manager
- A Microsoft Azure account (personal)
- A personal Outlook/Hotmail account with forwarding set up from your university email

---

## Setup

### 1. Clone and install dependencies

```bash
git clone https://github.com/yourusername/outlook-timetable-extractor.git
cd outlook-timetable-extractor
uv sync
```

### 2. Register an Azure App

1. Go to [portal.azure.com](https://portal.azure.com) and sign in with your **personal** Microsoft account
2. Search for **App registrations** and click **New registration**
3. Fill in the form:
   - **Name:** `outlook-timetable-extractor`
   - **Supported account types:** Any Entra ID tenant + Personal accounts
   - **Redirect URI:** Public client/native → `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Register** and copy your **Application (client) ID** and **Directory (tenant) ID**
5. Under **Authentication**, enable **Allow public client flows** → **Yes**
6. Under **API permissions**, add **Microsoft Graph → Delegated → Mail.Read**

### 3. Set up email forwarding

In your university Outlook (browser), go to **Settings → Mail → Forwarding** and forward all incoming mail to your personal Outlook/Hotmail address.

> ⚠️ Forwarding only applies to emails received **after** it is enabled. To test before the next timetable email arrives, manually forward a previous timetable email to your personal mailbox.

### 4. Configure credentials

Create a `.env` file in the project root:

```
APPLICATION_CLIENT_ID=your-client-id-here
DIRECTORY_TENANT_ID=your-tenant-id-here
```

### 5. Configure your keywords

Open `config.py` and update `KEYWORDS` with text that uniquely identifies your batch's page in the timetable PDF:

```python
KEYWORDS = ["Software Engineering", "Intake 41"]
```

---

## Usage

### Run manually

```bash
uv run main.py
```

On first run you will be prompted to authenticate via browser using the device code flow. After that, the script runs silently using the cached token.

### Run automatically with Task Scheduler

1. Update `run.bat` with the correct path to your project folder
2. Open **Task Scheduler** and create a new basic task
3. Set trigger to **Weekly** on your preferred day and time
4. Set action to run `run.bat`
5. Under **Settings**, check **"Run task as soon as possible after a scheduled start is missed"** — this ensures the script runs on the next startup if your PC was off during the scheduled time

---

## Output

Images are saved in the `output/` folder with dated filenames:

```
output/
├── timetable_2026-03-23.png
├── timetable_2026-03-27.png
└── ...
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `msal` | Microsoft OAuth2 authentication |
| `requests` | HTTP calls to Microsoft Graph API |
| `pymupdf` | PDF parsing and image rendering |
| `python-dotenv` | Loading credentials from `.env` |
| `python-telegram-bot` | *(Optional)* Sending image to Telegram |

---

## Notes

- `token_cache.json` is auto-created on first login and stores your refresh token. It is excluded from version control via `.gitignore`. Do not share it.
- The refresh token is valid for 90 days. After expiry you will be prompted to log in once more via the device code flow.
- The MuPDF warning `format error: No common ancestor in structure tree` is suppressed — it is a harmless quirk of how your university generates their PDFs and does not affect output.
- If the script prints `No timetable email found`, the email likely hasn't been sent yet for the week or forwarding was set up after the email arrived. You can manually forward a previous timetable email to your personal mailbox to test.
- The `$filter` and `$orderby` combination is not used in the Graph API query due to a known API limitation. Instead, up to 10 matching emails are fetched and sorted by date in Python.
- This script is built specifically to align with the workflow of my university — the email structure, PDF format, and timetable layout may differ elsewhere. But the approach is fully adaptable.

---

## Optional: Telegram Delivery

You can have the script send the timetable image directly to your phone via a Telegram bot.

### 1. Install the library

```bash
uv add python-telegram-bot
```

### 2. Create a Telegram bot

1. Open Telegram and search for **BotFather**
2. Send `/newbot` and follow the prompts to name your bot
3. BotFather will give you a **bot token** — copy it

### 3. Get your chat ID

1. Open a chat with your new bot and send it any message
2. Open this URL in your browser (replace with your actual token):
```
https://api.telegram.org/bot<your-token>/getUpdates
```
3. Find the `"chat"` object in the JSON response and copy the `"id"` value

### 4. Add credentials to `.env`

```
TELEGRAM_BOT_TOKEN=your-bot-token-here
TELEGRAM_CHAT_ID=your-chat-id-here
```

### 5. Add to `config.py`

```python
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
```

### 6. Create `notifier.py`

```python
import telegram
import asyncio
from pathlib import Path
from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID


async def _send(image_path, image_name):
    bot = telegram.Bot(token=TELEGRAM_BOT_TOKEN)
    with open(image_path, "rb") as img:
        await bot.send_photo(
            chat_id=TELEGRAM_CHAT_ID,
            photo=img,
            caption=image_name
        )


def send_timetable(image_path):
    asyncio.run(_send(image_path, Path(image_path).name))
    print("Timetable is sent to Telegram bot: Timetable Extractor")
```

### 7. Update `main.py`

Import and call `send_timetable()` after the image is saved:

```python
from notifier import send_timetable

# After save_page_as_image()
send_timetable(image_path)
```