from Setup import *
from Operation import *
import time
from datetime import datetime
import pandas as pd
# =============================
# Configuration
# =============================
worksheet_name = 'SD'
sheet = google_sheet_api(worksheet_name=worksheet_name)

# =============================
# Read pending parcel IDs
# =============================
# load sheet into DataFrame
data = sheet.get_all_records()
df = pd.DataFrame(data)
# filter conditions
filtered_df = df[
    (df["Process Date"].astype(str).str.strip() == "") &
    (df["Process Type"].astype(str).str.strip() == "Second delivery")
]
# extract parcel IDs
parcel_id_list = filtered_df["Process Tracking Number"].tolist()
print(f"Pending parcels: {parcel_id_list}")

# =============================
# Start Selenium
# =============================
driver = init_chrome_driver(keep_open=False)
login_uniuni(driver)
open_edit_order(driver)
results = []

for parcel_id in parcel_id_list:
    try:
        print(f"\n📦 Processing {parcel_id}")

        parcel_info = search_parcel(
            driver, parcel_id,
            get_driver=False,
            get_warehouse=False,
            get_sub_batch=False,
            get_status=False,
            get_segment=False,
            get_storage=False,
            get_driver_memo=False,
            get_reschedule_delivery_sn=True
        )

        reschedule_delivery_sn = parcel_info.get("reschedule_delivery_sn", "")

        print(f"📝 Reschedule Delivery SN: {reschedule_delivery_sn}")
        results.append([parcel_id, reschedule_delivery_sn])

    except Exception as e:
        print(f"❌ Failed for {parcel_id}: {e}")
        results.append([parcel_id, ""])

for parcel_id, reschedule_delivery_sn in results:
    print(f"Updated {parcel_id} with Reschedule Delivery SN: {reschedule_delivery_sn}")
    

    parcel_info = search_parcel(driver, parcel_id=reschedule_delivery_sn,
                  get_driver=False, 
                  get_warehouse=False, 
                  get_sub_batch=False, 
                  get_status=True, 
                  get_segment=False, 
                  get_storage=False, 
                  get_driver_memo=False, 
                  get_reschedule_delivery_sn=False)
    print(parcel_info)
    
    status = parcel_info.get("status", "")
    print(f"Current status of {reschedule_delivery_sn}: {status}")
    if status == "190: ORDER_RECEIVED":
        try:
            print(f"✅ Parcel {reschedule_delivery_sn} status is ORDER_RECEIVED")
            open_operation_and_next_transition(driver)
            select_transition_option(driver, "GATEWAY_PROCESSING")
            click_submit_status(driver)
            time.sleep(3)
            open_next_transition_dropdown(driver)
            select_transition_option(driver, "GATEWAY_TRANSIT_DISPATCH")
            click_submit_status(driver)
            time.sleep(3)

            process_date = f"{datetime.now().month}/{datetime.now().day}/{datetime.now().year}"
            results.append([parcel_id, reschedule_delivery_sn, process_date])
        except:
            print(f"❌ Failed to update status for {reschedule_delivery_sn}")
            process_date = ""
            results.append([parcel_id, reschedule_delivery_sn, process_date])
        

print("second delivery status update completed.")

# =============================
# Update reschedule delivery SN in Google Sheet
# =============================
headers = sheet.row_values(1)
process_col = headers.index("Process Tracking Number") + 1
reschedule_col = headers.index("Reschedule Delivery SN") + 1
process_values = sheet.col_values(process_col)

for parcel_id, reschedule_delivery_sn in results:
    try:
        row_index = process_values.index(parcel_id) + 1

        sheet.update_cell(
            row_index,
            reschedule_col,
            reschedule_delivery_sn
        )

    except ValueError:
        print(f"Parcel ID not found: {parcel_id}")

print("Google Sheet update completed.")
