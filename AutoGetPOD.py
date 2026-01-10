from datetime import datetime
from Setup import *
from Operation import *

dsp_chosen = None  # Example: "DHL" or ["DHL", "FedEx"]
POD_count = 50  # Number of PODs to download per driver

# ====================
# Get data from google sheet section
# ====================
sheet = google_sheet_api(worksheet_name="Delivery Record")
data = sheet.get_all_records()

# ====================
# Process data section
# ====================
def get_today_str():
    """Return today's date in M/D/YYYY format (no leading zeros)."""
    today = datetime.today()
    return f"{today.month}/{today.day}/{today.year}"

def get_today_delivery_record(data):
    """
    Find rows where 'Execution Date' equals today's date.
    Return a list of row dicts.
    """
    today_str = get_today_str()
    print(f"Today's date string: {today_str}")

    today_delivery_record = []

    for row in data:
        exec_date = row.get("Execution Date")

        if not exec_date:
            continue

        # Normalize Execution Date to string
        if isinstance(exec_date, datetime):
            exec_date_str = f"{exec_date.month}/{exec_date.day}/{exec_date.year}"
        else:
            exec_date_str = str(exec_date).strip()

        if exec_date_str == today_str:
            today_delivery_record.append(row)

    return today_delivery_record

def parse_allocation(allocation_str):
    """
    Extract driver id inside parentheses and join them with commas.
    """
    if not allocation_str:
        return ""

    values = re.findall(r"\((\d+)\)", allocation_str)
    return ",".join(values)

today_delivery_record = get_today_delivery_record(data)
valid_delivery_record = [
    row for row in today_delivery_record
    if str(row.get("Allocation", "")).strip()
]
if not valid_delivery_record:
    print("No valid delivery record for today.")
    exit()

print("today_delivery_record:", valid_delivery_record)

def get_driver_ids(records, dsp_chosen=None):
    """
    Get driver ids from Allocation.

    Args:
        records (list[dict]): today_delivery_record
        dsp_chosen (str | None): DSP name to filter. None = all DSPs.

    Returns:
        str: comma-separated driver ids
    """
    driver_set = set()

    for row in records:
        # Filter DSP if specified
        if dsp_chosen and row.get("DSP") != dsp_chosen:
            continue

        allocation = row.get("Allocation")
        if not allocation:
            continue

        driver_str = parse_allocation(allocation)
        if not driver_str:
            continue

        driver_set.update(driver_str.split(","))

    return ",".join(sorted(driver_set))


driver_ids = get_driver_ids(valid_delivery_record, dsp_chosen=dsp_chosen)
print("driver_ids:", driver_ids)

# =============================
# Chrome driver setup section
# =============================
driver = init_chrome_driver(keep_open=True)

# =============================
# website login section
# =============================
login_uniuni(driver)
time.sleep(3)  # Wait for login to complete

# =============================
# Download PODs for each driver ID section
# =============================
# Click the "Mgmt." button
mgmt_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//p[text()='Mgmt.']/parent::div/parent::div"))
)
mgmt_button.click()
time.sleep(1)
# Click the "POD Quality Management" button
pod_qm_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "menu-pod-quality-management"))
)
pod_qm_button.click()
time.sleep(1)
# Wait for the textarea to be present
textarea = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "pod_management_select_drivers_textfield"))
)
textarea.click()
textarea.send_keys(driver_ids)
time.sleep(1)


textbox = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "pod_management_count_numbers_textfield"))
)

textbox.click()
textbox.send_keys(Keys.CONTROL + "a")  # select all
textbox.send_keys(Keys.DELETE)
textbox.send_keys(POD_count)

time.sleep(1)  # Wait for the UI to update
# Click the Submit button
submit_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[./span[text()='Submit']]"))
)
submit_button.click()
time.sleep(10)  # Wait for the results to load

# =============================
# Generate POD download file section
# =============================
import os
from pathlib import Path

def get_latest_pod_file(download_dir):
    files = list(Path(download_dir).glob("POD Quality (*).xlsx"))
    if not files:
        raise FileNotFoundError("No POD Quality Excel file found")

    return max(files, key=lambda f: f.stat().st_mtime)

download_dir = r"C:\Users\Xince\Downloads"
pod_file = get_latest_pod_file(download_dir)
print("Using file:", pod_file)

# rondomly selecte 50 rows from the excel file
import pandas as pd

def sample_pod_rows(excel_path, sample_size=50):
    """
    Read POD Excel, clean invalid rows, and randomly sample POD rows.
    """
    df = pd.read_excel(excel_path)

    # --- 1. Keep only valid POD rows ---
    df = df[
        df["Driver Id"].notna() &
        df["Tracking Number"].notna() &
        df["POD Website"].notna()
    ].copy()

    if len(df) < sample_size:
        raise ValueError(
            f"Only {len(df)} valid POD rows available, need {sample_size}"
        )

    # --- 2. Fix dtypes explicitly ---
    df["Driver Id"] = df["Driver Id"].astype(int).astype(str)
    df["Tracking Number"] = df["Tracking Number"].astype(str)
    df["Photo Count"] = df["Photo Count"].astype(int)

    # Delivery Time can stay as datetime or string
    df["Delivery Time"] = df["Delivery Time"].astype(str)

    # --- 3. Random sample ---
    sampled_df = df.sample(n=sample_size)

    return sampled_df

sampled_df = sample_pod_rows(pod_file, sample_size=50)

# =============================
# Write sampled POD rows to Google Sheet section
# =============================
worksheet = google_sheet_api(worksheet_name="POD")

def prepare_for_gsheet(df):
    df = df.copy()
    df = df.fillna("")
    return df.astype(str)

def upload_to_google_sheet(df, worksheet, start_row=2):
    df = prepare_for_gsheet(df)
    rows = df.values.tolist()

    worksheet.update(
        range_name=f"A{start_row}",
        values=rows,
        value_input_option="RAW"
    )

upload_to_google_sheet(sampled_df, worksheet)