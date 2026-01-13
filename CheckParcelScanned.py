import pandas as pd

# File paths
source_file = "C:\\Users\\Xince\\Downloads\\BUF_ROC司机退件 Office Use ONLY.xlsx"        # Excel file with 100+ subsheets
search_file = "C:\\Users\\Xince\\Downloads\\Speedy Inactivity penalities.xlsx"           # Excel file with parcel IDs to search
output_file = "Speedy Inactivity penalities Check.xlsx"   # Output file

# Load parcel IDs to search
search_df = pd.read_excel(search_file, sheet_name="Multi Search", usecols="C")
parcel_ids_to_search = search_df.iloc[:, 0].dropna().astype(str).tolist()  # ensure string

# Initialize a dictionary to store results
results = {pid: [] for pid in parcel_ids_to_search}

# Load the source file with all subsheets
xls = pd.ExcelFile(source_file)

for sheet_name in xls.sheet_names:
    sheet_df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Flatten all columns to string and search
    for pid in parcel_ids_to_search:
        if pid in sheet_df.astype(str).values:
            results[pid].append(sheet_name)

# Prepare results DataFrame
output_data = []
for pid in parcel_ids_to_search:
    # Join multiple sheet names with comma
    sheet_names = ", ".join(results[pid]) if results[pid] else ""
    output_data.append([pid, sheet_names])

output_df = pd.DataFrame(output_data, columns=["parcel_id", "scan_date"])

# Save to new Excel file
output_df.to_excel(output_file, index=False)

print(f"Scan complete! Results saved to '{output_file}'.")
