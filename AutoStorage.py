from Setup import *
from Operation import *
import time
from datetime import datetime

# =============================
# Configuration
# =============================
worksheet_name = 'Storage Record'
driver_sheep = "66017"
sheet = google_sheet_api(worksheet_name=worksheet_name)

# =============================
# Read pending parcel IDs
# =============================
rows = sheet.get_all_values()
pending_ids = []

for row in rows[1:]:
    while len(row) < 5:
        row.append("")
    a, b, c, d, e = row[:5]
    if a and not b and not c and not d and not e:
        pending_ids.append(a)

print(f"Pending parcels: {pending_ids}")

# =============================
# Start Selenium
# =============================
driver = init_edge_driver()
open_new_tab(driver)
open_edit_order(driver)

results = []

# =============================
# Main loop
# =============================
for parcel_id in pending_ids:
    try:
        print(f"\n📦 Processing {parcel_id}")

        parcel_info = search_parcel(
            driver, parcel_id,
            get_driver=True,
            get_warehouse=False,
            get_sub_batch=False,
            get_status=False,
            get_segment=False,
            get_storage=True,
            get_driver_memo=False
        )

        driver_id = parcel_info.get("driver_id", "")
        storage_info = parcel_info.get("storage", "N/A")

        # ✅ Already stored → record and skip
        if storage_info != "N/A":
            print(f"ℹ️ Already stored: {storage_info}")
            results.append([parcel_id, storage_info])
            continue

        print(f"[STORE] {parcel_id}")

        # Refresh driver ID from UI
        current_driver_id = get_driver_id(driver)

        # Fix sheep driver if needed
        if isinstance(current_driver_id, str) and current_driver_id.startswith("45000"):
            print("🔧 Updating driver ID")
            update_driver_id(driver, driver_id=driver_sheep)
            time.sleep(3)

        # Send to storage
        storage_success = send_parcel_to_storage(driver)

        if not storage_success:
            print("❌ Failed to send to storage")
            results.append([parcel_id, "N/A"])
            continue

        # ✅ Re-check storage after transition
        updated_info = search_parcel(
            driver, parcel_id,
            get_driver=False,
            get_warehouse=False,
            get_sub_batch=False,
            get_status=False,
            get_segment=False,
            get_storage=True,
            get_driver_memo=False
        )

        final_storage = updated_info.get("storage", "N/A")
        print(f"✅ Stored: {final_storage}")
        results.append([parcel_id, final_storage])

    except Exception as e:
        print(f"❌ Store failed for {parcel_id}: {e}")
        results.append([parcel_id, "N/A"])

# =============================
# Write results to Google Sheet
# =============================
today_str = datetime.today().strftime("%m/%d/%Y")
all_rows = sheet.get_all_values()
result_map = {pid: info for pid, info in results}

for i, row in enumerate(all_rows[1:], start=2):
    parcel_id = row[0].strip() if row else ""
    if parcel_id in result_map:
        sheet.update_cell(i, 2, result_map[parcel_id])
        sheet.update_cell(i, 4, today_str)
        sheet.update_cell(i, 5, "1")