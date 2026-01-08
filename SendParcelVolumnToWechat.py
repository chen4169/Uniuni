import pygetwindow
import pyautogui
import pyperclip
import time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import keyboard

from Setup import *
from Operation import *

#pip install pygetwindow, pip install schedule, pip install pyautogui, pip install pyperclip, pip install keyboard

# config
route_to_chat = {
    "450001": "🚛Speedy Sloth 【1340】-BUF",
    "450002": "🚛LogiPro【1279】- BUF",
    "450003": "文件传输助手",
    "450004": "🚛BS 【1276】- BUF",
    "450005": "文件传输助手",
    "450006": "🚛BS 【1276】- BUF",
    "450007": "🚛BS 【1276】- BUF",
    "450008": "🚛KGM【1341】-BUF",
    "450009": "🚛BS 【1276】- BUF"
}

skip_today = ["450001", "450002", "450004", "450006", "450007", "450008", "450009"]

# =====================
# gspread setup section
# =====================
sheet = google_sheet_api(worksheet_name = "Route")

# =====================
# data processing section
# =====================
def validate_header(data, header):
    if not data or header not in data[0]:
        raise ValueError(f"Header {header} not found")

def get_non_zero_by_header(data, header):
    """
    data: list of dicts from get_all_records()
    header: column header as string, e.g. "450004"
    """
    result = []

    for row in data:
        value = row.get(header, 0)

        if value not in (0, "", None):
            result.append({
                "Sub Batch": row["Sub Batch"],
                header: value
            })

    return result

data = sheet.get_all_records()

print("Parcel volumn data: ", data)

# =====================
# send messages to WeChat section
# =====================
def send_route_messages_to_chats(data, route_to_chat, skip_routes=None):
    if skip_routes is None:
        skip_routes = []

    for route, chat_name in route_to_chat.items():
        if route in skip_routes:
            continue  # skip routes for today

        # Validate header
        validate_header(data, route)
        
        # Get non-zero rows
        non_zero_rows = get_non_zero_by_header(data, route)
        if not non_zero_rows:
            continue

        # Format message
        total = sum(row[route] for row in non_zero_rows)
        today = datetime.today().strftime("%m/%d")
        lines = [f"{today} {route} tomorrow volumn forecast: {total}", "------"]
        for row in non_zero_rows:
            lines.append(f"{row['Sub Batch']}: {row[route]}")
        message = "\n".join(lines)

        # Open chat and send message
        open_chat(chat_name)
        chat_refresh()
        time.sleep(1)
        pyperclip.copy(message)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)

focus_wechat()
send_route_messages_to_chats(data, route_to_chat, skip_routes=skip_today)
