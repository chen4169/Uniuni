
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pygetwindow
import pyautogui
import pyperclip
import time
import keyboard

# =====================================================
# Defaults (can be overridden when calling functions)
# =====================================================

# --- Selenium defaults ---
DEFAULT_CHROME_DRIVER_PATH = "chromedriver.exe"
DEFAULT_EDGE_DRIVER_PATH = "msedgedriver.exe"
DEFAULT_EDGE_DEBUGGER_ADDRESS = "127.0.0.1:9222"
DEFAULT_TARGET_URL = "https://dispatch.uniuni.com/main"

# --- Google Sheets defaults ---
DEFAULT_CREDS_PATH = "credentials.json"
DEFAULT_SHEET_KEY = "1w8AkWcq_Lm-IOLUcGaZIY77hI1I6CibkQVW4nFC30nM"
DEFAULT_WORKSHEET_NAME = ""

# =====================================================
# Selenium: Chrome
# =====================================================
def init_chrome_driver(
    chromedriver_path: str = DEFAULT_CHROME_DRIVER_PATH,
    keep_open: bool = True
):
    chrome_options = ChromeOptions()
    chrome_options.add_argument("--no-sandbox")
    #chrome_options.add_argument("--start-maximized")

    if keep_open:
        chrome_options.add_experimental_option("detach", True) # if keep_open is set to True when calling, keep the browser open after finished

    driver = webdriver.Chrome(
        service=ChromeService(chromedriver_path),
        options=chrome_options
    )
    return driver

# =====================================================
# Selenium: Edge (attach to existing browser)
# =====================================================
def init_edge_driver(
    edgedriver_path: str = DEFAULT_EDGE_DRIVER_PATH,
    debugger_address: str = DEFAULT_EDGE_DEBUGGER_ADDRESS
):
    edge_options = EdgeOptions()
    edge_options.add_experimental_option(
        "debuggerAddress", debugger_address
    )

    driver = webdriver.Edge(
        service=EdgeService(edgedriver_path),
        options=edge_options
    )
    return driver

def switch_to_target_tab(
    driver,
    target_url: str = DEFAULT_TARGET_URL
) -> bool:
    """
    Switch to the first tab whose URL starts with target_url.
    Returns True if found, False otherwise.
    """
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if driver.current_url.startswith(target_url):
            return True
    return False

def open_new_tab(driver, url: str = DEFAULT_TARGET_URL):
    """
    Open a new browser tab (Edge or Chrome) and navigate to the given URL.
    """
    driver.switch_to.new_window('tab')  # opens a new tab
    driver.get(url)                     # navigate to the URL

# =====================================================
# Google Sheets
# =====================================================
def google_sheet_api(
    creds_path: str = DEFAULT_CREDS_PATH,
    sheet_key: str = DEFAULT_SHEET_KEY,
    worksheet_name: str = DEFAULT_WORKSHEET_NAME
):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        creds_path, scope
    )
    client = gspread.authorize(creds)

    return client.open_by_key(sheet_key).worksheet(worksheet_name)

#
# WeChat setup
#
def open_chat(chat_name):
    """
    Focus WeChat chat by name.
    Assumes WeChat is already the active window. It only works if the target group is the first option after ctrl+f search
    """
    # Open search box
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(1)  # wait for search box to appear

    # Copy chat name to clipboard and paste
    pyperclip.copy(chat_name)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    # Press Enter to open the chat
    pyautogui.press('enter')
    time.sleep(1)  # wait for chat to open

def chat_refresh():
    """
    Refresh WeChat window by pressing the Windows key twice.
    """
    # Press Win key twice to refresh (simulate minimizing/maximizing)
    keyboard.press_and_release("windows")
    time.sleep(1)
    keyboard.press_and_release("windows")
    time.sleep(1)


def focus_wechat():
    windows = pygetwindow.getAllWindows()

    wechat = None
    for w in windows:
        if w.title.strip() == "WeChat":
            wechat = w
            break

    if not wechat:
        raise RuntimeError("Desktop WeChat window not found")

    if wechat.isMinimized:
        wechat.restore()
        time.sleep(1)

    wechat.activate()
    time.sleep(1)