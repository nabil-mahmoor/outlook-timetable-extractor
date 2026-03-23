# Outlook Timetable Extractor

A Python automation script that queries your Outlook mailbox, finds the latest timetable email, extracts your specific batch's timetable page from the PDF attachment, and saves it as an image — automatically.

---

## How It Works

1. Authenticates with your Outlook mailbox via Microsoft Graph API
2. Finds the latest email with "Timetable" in the subject
3. Downloads the PDF attachment
4. Scans each page for keywords unique to your batch and degree
5. Renders that page as a dated PNG image
6. Deletes the temporary PDF

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
3. Set trigger to **Weekly** on **Thursday**, **Friday**, and **Saturday** at your preferred time
4. Set action to run `run.bat`

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

---

## Notes

- `token_cache.json` is auto-created on first login and stores your refresh token. It is excluded from version control via `.gitignore`. Do not share it.
- The MuPDF warning `format error: No common ancestor in structure tree` is suppressed — it is a harmless quirk of how your university generates their PDFs and does not affect output.
- If the script prints `No timetable email found`, the email likely hasn't been sent yet for the week or forwarding was set up after the email arrived. You can manually forward a previous timetable email to your personal mailbox to test.
- The `$filter` and `$orderby` combination is not used in the Graph API query due to a known API limitation. Instead, up to 10 matching emails are fetched and sorted by date in Python.