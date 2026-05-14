import os
import json
import time
import re
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# ---------------- CONFIG ----------------
with open("config.json") as f:
    config = json.load(f)

SPREADSHEET_ID = config["spreadsheet_id"]
CREDENTIALS_FILE = config["credentials_file"]
BASE_PATH = config["base_path"]
WORKSHEET_NAME = config.get("worksheet_name", "ETIMSBackupFolders")

# ---------------- AUTH ----------------
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
client = gspread.authorize(creds)

# ---------------- SHEET ----------------
spreadsheet = client.open_by_key(SPREADSHEET_ID)

try:
    sheet = spreadsheet.worksheet(WORKSHEET_NAME)
except WorksheetNotFound:
    print(f"[INFO] Worksheet '{WORKSHEET_NAME}' not found. Creating it...")
    sheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows="1000", cols="3")
    sheet.append_row(["Station", "Anydesk Code", "Path"])

# ---------------- BASE FOLDER ----------------
os.makedirs(BASE_PATH, exist_ok=True)

# ---------------- SANITIZE FUNCTION ----------------
def sanitize_name(name):
    # Remove invalid Windows characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    return name.strip()

# ---------------- READ DATA ----------------
data = sheet.get_all_values()

updates = []
row_indices = []

# Skip header
for i, row in enumerate(data[1:], start=2):
    station = row[0].strip() if len(row) > 0 else ""
    path_cell = row[2].strip() if len(row) > 2 else ""

    if not station or path_cell:
        continue

    clean_name = sanitize_name(station)
    folder_path = os.path.join(BASE_PATH, clean_name)

    try:
        os.makedirs(folder_path, exist_ok=True)

        formatted_path = "/" + folder_path.replace("\\", "/")

        updates.append(formatted_path)
        row_indices.append(i)

        print(f"[READY] {station} -> {formatted_path}")

    except Exception as e:
        print(f"[ERROR] {station}: {e}")

# ---------------- WRITE BACK (RATE SAFE) ----------------
BATCH_SIZE = 20   # safe buffer under 60 writes/min

for idx, path in zip(row_indices, updates):
    try:
        sheet.update_cell(idx, 3, path)
        print(f"[WRITTEN] Row {idx}")

        time.sleep(1)  # 1 sec delay = ~60 writes/min safe

    except Exception as e:
        print(f"[WRITE ERROR] Row {idx}: {e}")
        print("[INFO] Sleeping 60s to recover from rate limit...")
        time.sleep(60)
