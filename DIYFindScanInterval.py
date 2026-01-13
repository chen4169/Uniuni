
from Setup import *
from Operation import *
from datetime import datetime
driver = init_chrome_driver()
login_uniuni(driver)
open_edit_order(driver)
wait_timeout: int = 10
wait = WebDriverWait(driver, wait_timeout)
#parcel_id="UUS5BN0567412254784"

import re
from datetime import datetime

def debug_check_scan_transit(all_texts, min_minutes=5):
    scanned_time = None
    transit_time = None

    time_pattern = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}")

    for i, text in enumerate(all_texts):
        text = text.strip()

        # ---- Find PARCEL_SCANNED time ----
        if text.upper().startswith("200: PARCEL_SCANNED"):
            if i + 1 < len(all_texts) and time_pattern.match(all_texts[i + 1]):
                scanned_time = datetime.strptime(
                    all_texts[i + 1], "%Y-%m-%d %H:%M:%S"
                )

        # ---- Find IN_TRANSIT time ----
        if text.upper().startswith("202: IN_TRANSIT"):
            if i + 1 < len(all_texts) and time_pattern.match(all_texts[i + 1]):
                transit_time = datetime.strptime(
                    all_texts[i + 1], "%Y-%m-%d %H:%M:%S"
                )

    # ---- Print what we found ----
    print("PARCEL_SCANNED time:", scanned_time)
    print("IN_TRANSIT time:", transit_time)

    # ---- Check existence ----
    if not scanned_time or not transit_time:
        print("❌ Missing one or both timestamps")
        return None, False, scanned_time, transit_time

    # ---- Calculate interval ----
    interval_minutes = abs(
        (transit_time - scanned_time).total_seconds()
    ) / 60

    print("Interval (minutes):", interval_minutes)

    # ---- Validate ----
    is_correct = interval_minutes >= min_minutes
    print("Is correct:", is_correct)

    return int(interval_minutes), is_correct, scanned_time, transit_time


def extract_all_timeline_texts(driver):
    """
    Extract all visible <p class='MuiTypography-body2'> texts
    from the parcel timeline.

    Returns:
        List[str]: ordered list of timeline texts
    """
    all_texts = []

    try:
        timeline_items = driver.find_elements(
            By.XPATH, "//li[contains(@class,'MuiTimelineItem-root')]"
        )
    except Exception as e:
        print("❌ Failed to find timeline items:", e)
        return all_texts

    for li in timeline_items:
        try:
            paper_div = li.find_element(
                By.XPATH,
                ".//div[contains(@class,'MuiTimelineContent-root')]"
                "//div[contains(@class,'MuiPaper-root')]"
            )

            p_tags = paper_div.find_elements(
                By.XPATH,
                ".//p[contains(@class,'MuiTypography-body2')]"
            )

            for p in p_tags:
                text = p.text.strip()
                if text:
                    all_texts.append(text)

        except Exception:
            # structure missing → skip safely
            continue

    return all_texts


sheet = google_sheet_api(worksheet_name="DIY")
rows = sheet.get_all_records()
# Build parcel_id list + row index map (1-based row numbers)
parcel_row_map = {}
parcel_id_list = []

for i, row in enumerate(rows, start=2):  # start=2 (row 1 is header)
    pid = str(row.get("parcel_id", "")).strip()
    if pid:
        parcel_id_list.append(pid)
        parcel_row_map[pid] = i

print("parcel_id_list:", parcel_id_list)
results = []  

for parcel_id in parcel_id_list:
    try:
        # ---- Step 1: Search input ----
        search_input = wait.until(
            EC.element_to_be_clickable((By.ID, "searchSN"))
        )
        search_input.send_keys(Keys.CONTROL, "a")
        search_input.send_keys(Keys.DELETE)
        search_input.send_keys(parcel_id, Keys.ENTER)
        # small wait for page to update
        time.sleep(2)

        # ---- Extract timeline ----
        all_texts = extract_all_timeline_texts(driver)
        # Print all collected <p> texts
        # for t in all_texts:
        #     print(t)
        # ---- Calculate interval ----
        interval_minutes, is_correct, scanned_time, transit_time = debug_check_scan_transit(all_texts)

        results.append((parcel_id, interval_minutes, is_correct, scanned_time, transit_time))

    except Exception as e:
        print(f"❌ Failed parcel {parcel_id}: {e}")

print("Results:", results)

# Headers in desired order
headers = [
    "parcel_id",
    "interval_minutes",
    "is_correct",
    "scanned_time",
    "transit_time",
]

# Helper to format datetime
def format_dt(dt):
    if isinstance(dt, datetime):
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return ""

# Build sheet values (first row = headers)
values = [headers]
for parcel_id, interval, is_correct, scanned_time, transit_time in results:
    values.append([
        parcel_id,
        interval,
        is_correct,
        format_dt(scanned_time),
        format_dt(transit_time),
    ])

# -------------------------------
# 3. Overwrite the sheet
# -------------------------------

sheet.clear()          # remove all existing data
sheet.update(values)   # write new data

print("Sheet updated successfully!")


# # ---- Step 1: Search input ----
# search_input = wait.until(
#     EC.element_to_be_clickable((By.ID, "searchSN"))
# )
# search_input.send_keys(Keys.CONTROL, "a")
# search_input.send_keys(Keys.DELETE)
# search_input.send_keys(parcel_id, Keys.ENTER)
# # small wait for page to update
# time.sleep(2)

# from selenium.webdriver.common.by import By

# all_texts = extract_all_timeline_texts(driver)

# # Print all collected <p> texts
# for t in all_texts:
#     print(t)


# interval_minutes, is_correct = debug_check_scan_transit(all_texts)




# sheet = google_sheet_api(worksheet_name="DIY")
# rows = sheet.get_all_records()
# # Build parcel_id list + row index map (1-based row numbers)
# parcel_row_map = {}
# parcel_id_list = []

# for i, row in enumerate(rows, start=2):  # start=2 (row 1 is header)
#     pid = str(row.get("parcel_id", "")).strip()
#     if pid:
#         parcel_id_list.append(pid)
#         parcel_row_map[pid] = i

# print("parcel_id_list:", parcel_id_list)
# results = []  # (row_index, interval_minutes, is_correct)

# for parcel_id in parcel_id_list:
#     try:
#         # ---- Search parcel (you already have this logic elsewhere) ----
#         search_input = wait.until(EC.element_to_be_clickable((By.ID, "searchSN")))
#         search_input.send_keys(Keys.CONTROL, "a")
#         search_input.send_keys(Keys.DELETE)
#         search_input.send_keys(parcel_id, Keys.ENTER)
#         time.sleep(3)

#         # ---- Extract timeline ----
#         all_texts = extract_all_timeline_texts(driver)

#         # ---- Calculate interval ----
#         interval_minutes, is_correct = debug_check_scan_transit(all_texts)

#         row_index = parcel_row_map[parcel_id]
#         results.append((row_index, interval_minutes, is_correct))

#     except Exception as e:
#         print(f"❌ Failed parcel {parcel_id}: {e}")


# # Find column indexes
# headers = sheet.row_values(1)
# interval_col = headers.index("interval_minutes") + 1
# correct_col = headers.index("is_correct") + 1

# # Prepare batch updates
# updates = []

# for row_index, interval, correct in results:
#     updates.append({
#         "range": f"{chr(64 + interval_col)}{row_index}",
#         "values": [[interval]]
#     })
#     updates.append({
#         "range": f"{chr(64 + correct_col)}{row_index}",
#         "values": [[str(correct)]]
#     })

# # Execute batch update ONCE
# sheet.batch_update(updates)