from datetime import timedelta
import datetime as dt
from Setup import init_chrome_driver, google_sheet_api
from Operation import *

# =============================
# Init
# =============================
driver = init_chrome_driver()
wait = WebDriverWait(driver, 20)
wait_time = 5

login_uniuni(driver)

# =============================
# Create blank sub-batch
# =============================
click_batch_management(driver)
time.sleep(wait_time)
click_operate_shortcut(driver)
time.sleep(wait_time)
create_blank_sub_batch(driver)
time.sleep(wait_time)
click_recent_7_days(driver)
time.sleep(wait_time)

sub_batches = get_recent_sub_batches(driver, limit=1)
sub_batch = sub_batches[0]
print("Sub batch written to googel sheet:", sub_batch)

# =============================
# Load & Save
# =============================
open_new_tab_and_login(driver, DEFAULT_WEBSITE, DEFAULT_USERNAME, DEFAULT_PASSWORD)

time.sleep(wait_time)
click_load(driver)
time.sleep(wait_time)
submit_batch_number(driver, sub_batch)
time.sleep(wait_time)
click_save(driver)

# =============================
# Dispatch name
# =============================
tomorrow = dt.datetime.today() + timedelta(days=1)
dispatch_name = tomorrow.strftime("%m-%d") + "deliver"
input_dispatch_name(driver, dispatch_name)
time.sleep(wait_time)
click_submit(driver)

# =============================
# Google Sheet
# =============================
sheet = google_sheet_api(worksheet_name="Batch")
today = datetime.today()
today_str = today.strftime("%Y-%m-%d")

cn_month = {
    1: "一月", 2: "二月", 3: "三月", 4: "四月",
    5: "五月", 6: "六月", 7: "七月", 8: "八月",
    9: "九月", 10: "十月", 11: "十一月", 12: "十二月"
}

sub_sheet_name = f"{cn_month[today.month]}{today.year}-NY"

# =============================
# Locate columns
# =============================
headers = sheet.row_values(1)

date_col = headers.index("Date") + 1
sub_sheet_col = headers.index("Sub Sheet") + 1
sub_batch_col = headers.index("Sub Batch") + 1

# =============================
# Find next empty row (based on Sub Batch column)
# =============================
next_row = len(sheet.col_values(sub_batch_col)) + 1

# =============================
# Write values (ONE CELL AT A TIME)
# =============================
sheet.update_cell(next_row, date_col, today_str)
sheet.update_cell(next_row, sub_sheet_col, sub_sheet_name)
sheet.update_cell(next_row, sub_batch_col, sub_batch)