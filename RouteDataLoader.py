from Setup import *
from Operation import *
from datetime import datetime

sheet = google_sheet_api(worksheet_name="Route")
# Get all values from column A
sub_batches = sheet.col_values(1)
# Remove header (row 1) and empty cells
sub_batch_list = [v for v in sub_batches[1:] if v.strip()]
headers = sheet.row_values(1)
route_number = headers[1:10]
scan_id = headers[10:]

#print(sub_batch_list)
#print(headers)
#print(route_number)

driver = init_chrome_driver(detach=False)
login_uniuni(driver)

qty_data = []

for batch_number in sub_batch_list:
    qty = [batch_number]  # <-- reset qty for this sub-batch
    try:
        click_load(driver)
        submit_batch_number(driver, batch_number=batch_number)
        click_button_a(driver)
        expand_buf_zone(driver)

        for route in route_number:
            value = get_route_red_value(driver, route)
            #print(f"route value: {value} appended for {route}")
            qty.append(value)

        click_driver_off(driver)
        for sid in scan_id:
            value = get_quantity_by_sid(driver, sid)
            #print(f"scan value: {value} appended for scan id {sid}")
            qty.append(value)

    except Exception as e:
        print(f"Error processing batch {batch_number}: {e}")

    qty_data.append(qty)  # append the row for this sub-batch

print(qty_data)

# =============================
# Write Data Into Google Sheet
# =============================
now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
start_col = 'B'  # start writing data from column B

# Get all existing sub batches from column A (skip header)
all_sub_batches = sheet.col_values(1)[1:]  # row 2 onward
all_sub_batches_lookup = {v: i + 2 for i, v in enumerate(all_sub_batches)}  # map batch_number -> row_index

num_headers = len(qty_data[0]) - 1  # first element is batch number, rest are data

for row in qty_data:
    batch_number = row[0]
    data_values = row[1:]

    # Replace 0 with "" if needed
    data_values_clean = ["" if v == 0 else v for v in data_values]

    # find row index in sheet
    row_index = all_sub_batches_lookup.get(batch_number)
    if not row_index:
        print(f"[Warning] Batch {batch_number} not found in column A. Skipping.")
        continue

    # Compute range for data (B → next columns)
    end_col_letter = chr(ord(start_col) + len(data_values_clean) - 1)
    range_name = f"{start_col}{row_index}:{end_col_letter}{row_index}"

    # Update the data
    sheet.update(values=[data_values_clean], range_name=range_name, value_input_option="RAW")

    # Write timestamp in the next column
    ts_col_letter = chr(ord(start_col) + len(data_values_clean))
    ts_range = f"{ts_col_letter}{row_index}"
    sheet.update(values=[[now_str]], range_name=ts_range, value_input_option="RAW")