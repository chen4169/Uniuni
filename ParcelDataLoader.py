from Setup import *
from Operation import *
import time
from datetime import datetime

# =============================
# Google api setup
# =============================
worksheet_name = 'Parcel'
sheet = google_sheet_api(worksheet_name=worksheet_name)

# =============================
# Read pending parcel IDs
# =============================
HEADERS = {
    "scan_index": 1,
    "parcel_id": 2,
    "sub_batch": 3,
    "status": 4,
    "driver_id": 5,
    "warehouse": 6,
    "segment": 7,
    "storage": 8,
    "driver_memo": 9,
}

def get_pending_parcels(sheet):
    """
    Return rows where status is empty or NO_STATUS
    """
    records = sheet.get_all_records()
    pending = []

    for idx, row in enumerate(records, start=2):
        parcel_id = str(row.get("parcel_id", "")).strip()
        status = str(row.get("status", "")).strip()

        if parcel_id and (status == "" or status == "NO_STATUS"):
            pending.append({
                "row": idx,
                "parcel_id": parcel_id
            })
    return pending

# =============================
# Selenium setup
# =============================
driver = init_chrome_driver(keep_open=False)
# Login to uniuni system
login_uniuni(driver)
open_edit_order(driver)


# =============================
# Colect parcel info
# =============================
def load_pending_parcels(sheet, driver):
    pending = get_pending_parcels(sheet)
    updates = []  # <-- collect here
    print("Searching parcels:", [item["parcel_id"] for item in pending])
    print(f"Found {len(pending)} parcels to process")

    for item in pending:
        row_num = item["row"]
        parcel_id = item["parcel_id"]

        try:
            info = search_parcel(driver, parcel_id)

            updates.append({
                "row": row_num,
                "values": [
                    info.get("sub_batch", ""),
                    info.get("status", ""),
                    info.get("driver_id", ""),
                    info.get("warehouse", ""),
                    info.get("segment", ""),
                    info.get("storage", ""),
                    info.get("driver_memo", "")
                ]
            })

        except Exception as e:
            print(f"[ERROR] Row {row_num}, parcel {parcel_id}: {e}")

    batch_write_parcel_info(sheet, updates)

# =============================
# Load parcel info to google sheet
# =============================
def batch_write_parcel_info(sheet, updates):
    """
    Perform a single batch update for multiple rows.
    """
    data = []

    for item in updates:
        row = item["row"]
        data.append({
            "range": f"C{row}:I{row}",
            "values": [item["values"]]
        })

    if data:
        sheet.batch_update(data)


load_pending_parcels(sheet, driver)
