import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from Setup import *

# =============================
# Configuration
# =============================
WORKSHEET_NAME = "Parcel"
CSV_FILE = "scans.csv"
CSV_HEADERS = [
    "scan_index",
    "parcel_id",
    "sub_batch",
    "status",
    "driver",
    "warehouse",
    "segment",
    "storage",
    "driver_memo"
]

# =============================
# CSV → Google Sheet
# =============================
def load_csv_to_google_sheet():
    worksheet = google_sheet_api(worksheet_name=WORKSHEET_NAME)

    worksheet.clear()
    rows = []
    rows.append(CSV_HEADERS)  # header row for Google Sheet

    with open(CSV_FILE, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) == len(CSV_HEADERS):
                rows.append(row)

    worksheet.update(
        range_name="A1",
        values=rows
    )

    print(f"✅ Uploaded {len(rows) - 1} parcels to Google Sheet")

if __name__ == "__main__":
    load_csv_to_google_sheet()