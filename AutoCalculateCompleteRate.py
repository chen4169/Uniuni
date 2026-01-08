
from datetime import datetime
from Setup import *
from Operation import *

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
 
# =============================
#  Selenium setup section
# =============================
driver = init_chrome_driver()
sleeptime = 3 
wait = WebDriverWait(driver, 10)

# =============================
# Uniuni login section
# =============================
login_uniuni(driver)


# =============================
# Get complete rate data section
# =============================
def open_driver_efficiency_monitor(driver, timeout=10):
    """
    Click Tools → Driver Efficiency Monitor
    """

    wait = WebDriverWait(driver, timeout)

    # 1️⃣ Click "Tools"
    tools_button = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[text()='Tools']/ancestor::div[contains(@class,'MuiPaper-root')]")
        )
    )
    tools_button.click()
    time.sleep(1)

    # 2️⃣ Click "Driver Efficiency Monitor"
    efficiency_button = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "menu-driver-efficiency-management")
        )
    )
    efficiency_button.click()
    time.sleep(1)

def get_complete_rate(driver, row, timeout=10):
    wait = WebDriverWait(driver, timeout)

    batch_number = row["Batch Number"]
    allocation_raw = row["Allocation"]

    # 1️⃣ Batch Number input
    batch_input = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "searchactionsection_selectedinput_inputtextfield")
        )
    )
    batch_input.click()
    batch_input.send_keys(Keys.CONTROL, "a")
    batch_input.send_keys(Keys.DELETE)
    batch_input.send_keys(batch_number)

    # 2️⃣ Driver ID input
    driver_input = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "searchactionsection_selecteddrivers_inputtextfield")
        )
    )
    driver_input.click()
    driver_input.send_keys(Keys.CONTROL, "a")
    driver_input.send_keys(Keys.DELETE)
    driver_input.send_keys(allocation_raw)

    # 3️⃣ Wait before search
    time.sleep(1)

    # 4️⃣ Click Search
    search_button = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "searchactionsection_search_button")
        )
    )
    search_button.click()

    time.sleep(1)
    table_data = extract_complete_rate_table(driver)
    return table_data
 
def extract_complete_rate_table(driver, timeout=10):
    wait = WebDriverWait(driver, timeout)

    # Wait until at least one row appears
    wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.MuiDataGrid-row[data-rowindex]")
        )
    )

    rows = driver.find_elements(
        By.CSS_SELECTOR, "div.MuiDataGrid-row[data-rowindex]"
    )

    results = []

    for row in rows:
        row_data = {}

        cells = row.find_elements(By.CSS_SELECTOR, "div[role='cell'][data-field]")

        for cell in cells:
            field = cell.get_attribute("data-field")
            value = cell.get_attribute("data-value")

            # Skip empty technical filler cells
            if field and value is not None:
                row_data[field] = value

        if row_data:
            results.append(row_data)

    return results

open_driver_efficiency_monitor(driver)
complete_rate = []

for row in valid_delivery_record:
    # Parse allocation first
    allocation_raw = row["Allocation"]
    row["Allocation"] = parse_allocation(allocation_raw)

    result_list = get_complete_rate(driver, row)  # usually a list of dicts
    for result in result_list:
        # Merge the original row info into result
        combined = {**row, **result}  # row info + result info, result keys overwrite if duplicated

        # Rename keys in the combined dict
        combined["driver_id"] = combined.pop("shipping_staff_id")
        combined["complete_rate"] = combined.pop("done_ratio")
        combined["first_to_current"] = combined.pop("first_idle_interval")
        combined["inactive_time"] = combined.pop("idle_interval")

        complete_rate.append(combined)

print("complete_rate: ", complete_rate)


# =============================
# Process complete rate data section
# =============================
import pandas as pd

def prepare_complete_rate_report(complete_rate_list):
    """
    Prepare group-level and driver-level complete rate DataFrames.
    
    Args:
        complete_rate_list (list of dict): driver-level data with keys like DSP, driver_id, 202, 211, 231, 232, 203, complete_rate
    
    Returns:
        tuple: (complete_rate_by_group, complete_rate_by_driver) as DataFrames
    """
    # Convert to DataFrame
    df = pd.DataFrame(complete_rate_list)
    
    # Make numeric columns
    for col in ["202","211","231","232","203","complete_rate"]:
        df[col] = pd.to_numeric(df[col])
    
    # Driver-level derived columns
    df["total_volumn"] = df[["202","211","231","232","203"]].sum(axis=1)
    df["return_volumn"] = df[["211","231","232"]].sum(axis=1)
    df["return_rate"] = df["return_volumn"] / df["total_volumn"]
    
    # Count drivers under 95% (used later for group-level)
    df["under_95_driver"] = df["complete_rate"] < 95
    
    # Group-level aggregation
    agg_cols = ["202","211","231","232","203"]
    df_group = df.groupby("DSP")[agg_cols].sum().reset_index()
    batch_per_dsp = df.groupby("DSP")["Batch Number"].apply(lambda x: ", ".join(map(str, x.unique()))).reset_index()
    df_group = df_group.merge(batch_per_dsp, on="DSP", how="left")
    df_group["total_volumn"] = df_group[["202","211","231","232","203"]].sum(axis=1)
    df_group["return_volumn"] = df_group[["211","231","232"]].sum(axis=1)
    df_group["return_rate"] = df_group["return_volumn"] / df_group["total_volumn"]
    df_group["complete_rate"] = df_group["203"] / df_group["total_volumn"]
    
    # Count under-95 drivers per DSP
    under_95_count = df[df["complete_rate"] < 95].groupby("DSP").size().to_dict()
    df_group["under_95_driver"] = df_group["DSP"].map(lambda x: under_95_count.get(x, 0))
    
    # Driver-level DataFrame: drop unnecessary columns
    drop_cols = ["Status", "Execution Date", "Scheduled Quantity", "Scan ID", "Allocation", "first_update", "first_to_current", "inactive_time", "under_95_driver", "time_id", "last_update", "Execution Date"]
    complete_rate_by_driver = df.drop(columns=[c for c in drop_cols if c in df.columns])
    complete_rate_by_driver.rename(columns={"complete_rate": "complete_rate"}, inplace=True)
    
    # Group-level DataFrame
    complete_rate_by_group = df_group
    
    return complete_rate_by_group, complete_rate_by_driver

complete_rate_by_group, complete_rate_by_driver = prepare_complete_rate_report(complete_rate)

print("Group-level data:")
print(complete_rate_by_group)
print("\nDriver-level data:")
print(complete_rate_by_driver)

# =============================
# Upload complete rate data to Google Sheets section
# =============================
target_sheet = google_sheet_api(worksheet_name='Complete Rate')
today_str = datetime.now().strftime("%m/%d/%Y-%H:%M")

# First, prepare data to write for the sheet
def prepare_sheet_data(df, date_str):
    df_to_write = df.copy()
    df_to_write['Date'] = date_str  # update Date column
    # Reorder columns to match your sheet headers
    columns_order = [
        'Date', 'Batch Number', 'DSP', '202', '211', '231', '232', '203',
        'total_volumn', 'return_rate', 'return_volumn', 'complete_rate',
        'under_95_driver', 'driver_id', 'Route'
    ]
    # Add missing columns if they don't exist in df
    for col in columns_order:
        if col not in df_to_write.columns:
            df_to_write[col] = ""
    df_to_write = df_to_write[columns_order]
    return df_to_write

# Prepare group and driver data
group_sheet_data = prepare_sheet_data(complete_rate_by_group, today_str)
driver_sheet_data = prepare_sheet_data(complete_rate_by_driver, today_str)

# Combine them into a single DataFrame to upload
sheet_data = pd.concat([group_sheet_data, driver_sheet_data], ignore_index=True)

# Convert DataFrame to list of lists (without headers, since headers already exist)
values_to_append = sheet_data.values.tolist()

# Append all rows to the sheet, at the newest row
target_sheet.append_rows(values_to_append, value_input_option='USER_ENTERED')


# =============================
# Send complete rate report to WeChat chats section
# =============================
dsp_to_chat = {
    "Speedy Sloth": "🚛Speedy Sloth 【1340】-BUF",
    "LogiPro": "🚛LogiPro【1279】- BUF",
    "Brothers Shipping": "🚛BS 【1276】- BUF",
    "KGM": "🚛KGM【1341】-BUF",
}

focus_wechat()

def send_group_complete_rate_report(df_group, df_driver, dsp_chat_map):
    """
    Send group-level complete rate report to corresponding WeChat chats,
    including driver-level summary.

    Args:
        df_group (DataFrame): complete_rate_by_group
        df_driver (DataFrame): complete_rate_by_driver
        dsp_chat_map (dict): mapping {DSP_name: WeChat_chat_name}
    """
    for _, row in df_group.iterrows():
        dsp_name = row['DSP']
        if dsp_name not in dsp_chat_map:
            print(f"No chat found for DSP {dsp_name}, skipping.")
            continue

        chat_name = dsp_chat_map[dsp_name]

        # Prepare the group report text
        now_str = datetime.now().strftime("%m/%d/%Y-%H:%M")
        report = (
            f"=== {now_str} Complete Rate Report ===\n"
            f"Batch Number: {row['Batch Number']}\n"
            f"DSP: {row['DSP']}\n"
            f"202: {row['202']}\n"
            f"211: {row['211']}\n"
            f"213: {row.get('213', 0)}\n"
            f"231: {row['231']}\n"
            f"232:  {row['232']}\n"
            f"Return volumn: {row['return_volumn']}\n"
            f"Complete Rate: {row['complete_rate']*100:.1f}%\n"
            f"Number of Driver Didn't Pass 95%: {row['under_95_driver']}\n"
            f"==================\n"
            f"=== Driver Level ===\n"
            f"Driver ID   Complete Rate   Return Rate\n"
        )

        # Append driver-level info for this DSP
        df_dsp_drivers = df_driver[df_driver['DSP'] == dsp_name]
        for _, drv in df_dsp_drivers.iterrows():
            report += (
                f"{drv['driver_id']}, "
                f"         {drv['complete_rate']:.1f}%, "
                f"         {drv['return_rate']*100:.1f}%\n"
            )

        # Open the chat
        open_chat(chat_name)
        chat_refresh()
        time.sleep(1)

        # Paste the report
        pyperclip.copy(report)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        print(f"Report sent to {chat_name}")
        time.sleep(1)


send_group_complete_rate_report(complete_rate_by_group, complete_rate_by_driver, dsp_to_chat)