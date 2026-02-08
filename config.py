import os

# Telegram Bot Token — set via environment variable or replace the default
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")

# Admin Telegram user IDs (numeric). Set via comma-separated env var.
_admin_ids_raw = os.getenv("ADMIN_IDS", "")
ADMIN_IDS: set[int] = set()
if _admin_ids_raw:
    for uid in _admin_ids_raw.split(","):
        uid = uid.strip()
        if uid.isdigit():
            ADMIN_IDS.add(int(uid))

# Output Excel file path
EXCEL_FILE = os.getenv("EXCEL_FILE", "results.xlsx")

# Questions file
QUESTIONS_FILE = os.getenv("QUESTIONS_FILE", "questions.json")
