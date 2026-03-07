# BK Forms Doc Parser

Monitors an inbox for BK intake form emails from Mercedes, parses unpaired `.docx` attachments using Gemini AI, and emails a CSV of client contact info (first name, last name, email, phone) to `info@andrewgriffinlawoffice.com`.

**Pairing logic:** If a client has both a `.docx` and `.pdf` attachment in the same email, the `.docx` is skipped. Only standalone `.docx` files are processed.

Run via Windows Task Scheduler on whatever interval you want (every 30 minutes is recommended). Each run checks for a new email — if none is found it exits immediately.

---

## Setup

### 1. Prerequisites

- Python 3.10 or newer
- A Google Cloud project with the **Gmail API** enabled
- A Gemini API key from [Google AI Studio](https://aistudio.google.com/app/apikey)

### 2. Clone the repo

```
git clone https://github.com/YOUR_USERNAME/docparser.git
cd docparser
```

### 3. Create the virtual environment and install dependencies

```
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

### 4. Set up Gmail API credentials

The script reads from the `info@andrewgriffinlawoffice.com` inbox.

1. Go to [console.cloud.google.com](https://console.cloud.google.com) signed in as `info@andrewgriffinlawoffice.com`
2. Create a project and enable the **Gmail API**
3. Go to **APIs & Services > Credentials > Create Credentials > OAuth 2.0 Client ID**
4. Application type: **Desktop app**
5. Download the JSON and save it as `credentials.json` in the project folder
6. The first time the script runs it will open a browser to authorize — after that `token.json` is saved and no browser is needed again

### 5. Create the .env file

Create a file named `.env` in the project folder with the following:

```
# Gmail (sending account)
SENDER_EMAIL=info@andrewgriffinlawoffice.com
APP_PASSWORD=xxxx xxxx xxxx xxxx

# Recipients for status notifications
NOTIFY_EMAILS=info@andrewgriffinlawoffice.com,kevinkakkary@gmail.com

# Gemini
GEMINI_API_KEY=your_gemini_api_key_here
```

To get an app password for `info@`:
1. Sign in at myaccount.google.com as `info@andrewgriffinlawoffice.com`
2. Security > 2-Step Verification > App passwords
3. Create one named "docparser" and paste it as `APP_PASSWORD`

### 6. Set up Task Scheduler (Windows)

1. Open **Task Scheduler** and click **Create Basic Task**
2. Set the trigger to **repeat every 30 minutes**
3. Action: **Start a program**
   - Program/script: `C:\path\to\docparser\.venv\Scripts\python.exe`
   - Arguments: `main.py`
   - Start in: `C:\path\to\docparser`
4. Check **"Run whether user is logged on or not"**

---

## Files

| File | Description |
|------|-------------|
| `main.py` | Main script |
| `.env` | Secrets and config (not committed) |
| `credentials.json` | Gmail OAuth credentials (not committed) |
| `token.json` | Gmail OAuth token, auto-generated on first run (not committed) |
| `last_processed_id.txt` | Tracks the last processed email to avoid duplicates (not committed) |
