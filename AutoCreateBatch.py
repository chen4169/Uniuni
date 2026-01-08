import time
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import pyperclip


# =============================
#  Selenium setup   
# =============================
username = "Libin_Chen"
password = "Xbaqs1221@uniuni"
website = "https://dispatch.uniuni.com/login"
sleeptime = 3 

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_experimental_option("detach", True)  # keep browser open after script ends
    driver = webdriver.Chrome(service=Service('chromedriver.exe'), options=chrome_options)
    return driver

driver = init_driver()
wait = WebDriverWait(driver, 20)
# =============================
# website login
# =============================
driver.get(website)
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
).send_keys(username)
driver.find_element(By.XPATH, "//input[@type='password']").send_keys(password)
driver.find_element(By.XPATH, "//button[.//span[text()='Login']]").click()
time.sleep(sleeptime)

# Click Batch Management button
batch_management_btn = wait.until(
    EC.element_to_be_clickable((By.ID, "menu-batch-management"))
)
batch_management_btn.click()

# Click operate shortcut (FAB button)
operate_shortcut_btn = wait.until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="operate shortcut"]'))
)
operate_shortcut_btn.click()

# Click create blank sub batch
create_blank_btn = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, '//button[.//span[text()="create blank sub batch"]]')
    )
)
create_blank_btn.click()
driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
driver.refresh()

# Click recent 7 days
recent_7_days_btn = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, '//button[.//span[text()="recent 7 days"]]')
    )
)
recent_7_days_btn.click()

# Collect sub batch numbers
sub_batch = []
elements = wait.until(
    EC.presence_of_all_elements_located(
        (By.XPATH, '//span[contains(text(), "BUSUB-")]')
    )
)

for el in elements:
    text = el.text.strip()
    if text:
        sub_batch.append(text)

sub_batch = sub_batch[0:1]  # Get the first sub batch
print(sub_batch)

# Open a new tab
driver.execute_script(f"window.open('{website}', '_blank');")
driver.switch_to.window(driver.window_handles[-1])
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
).send_keys(username)
driver.find_element(By.XPATH, "//input[@type='password']").send_keys(password)
driver.find_element(By.XPATH, "//button[.//span[text()='Login']]").click()
time.sleep(sleeptime)

# click Load button
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Load']]"))
).click()

# input batch number and submit
batch_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located(
        (By.XPATH, "//input[@placeholder='Batch Number (Please use , separate)']")
    )
)
batch_input.clear()
batch_input.send_keys(sub_batch)
driver.find_element(By.XPATH, "//button[.//span[text()='Submit']]").click()
time.sleep(2)  # wait for the list to load

# Click Save button
save_btn = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, '//button[.//span[text()="Save"]]')
    )
)
save_btn.click()

# Generate sub batch name based on today's date
today = datetime.today()
sub_batch_name = today.strftime("%m-%d") + "deliver"
print(sub_batch_name)

# Input the sub batch name
dispatch_input = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        '//label[text()="Dispatch History Name"]/following::input[1]'
    ))
)
dispatch_input.send_keys(sub_batch_name)

# Click Submit button
submit_btn = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, '//button[.//span[text()="Submit"]]')
    )
)
submit_btn.click()

# gspread setup
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'D:/Microsoft VS Code Project/credentials.json', scope)
client = gspread.authorize(creds)
sheet = client.open_by_url(
    'https://docs.google.com/spreadsheets/d/1w8AkWcq_Lm-IOLUcGaZIY77hI1I6CibkQVW4nFC30nM/edit#gid=839809763'
).worksheet('Batch')

today_str = datetime.today().strftime("%Y-%m-%d")
cn_month = {
    1: "一月", 2: "二月", 3: "三月", 4: "四月",
    5: "五月", 6: "六月", 7: "七月", 8: "八月",
    9: "九月", 10: "十月", 11: "十一月", 12: "十二月"
}

today = datetime.today()
sub_sheet_name = f"{cn_month[today.month]}{today.year}-NY"

# Find the next empty row in "Sub Batch" column and write the sub batch number
headers = sheet.row_values(1)
try:
    sub_batch_col = headers.index("Sub Batch") + 1  # gspread is 1-based
    date_col = headers.index("Date") + 1
    sub_sheet_col = headers.index("Sub Sheet") + 1
except ValueError:
    raise Exception('Header "Sub Batch" not found')

col_values = sheet.col_values(sub_batch_col)
next_row = len(col_values) + 1
# -------- write values --------
sheet.update_cell(next_row, date_col, today_str)
sheet.update_cell(next_row, sub_sheet_col, sub_sheet_name)
sheet.update_cell(next_row, sub_batch_col, sub_batch)