import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, ElementClickInterceptedException

DEFAULT_WAIT_TIMEOUT = 10
DEFAULT_SHORT_WAIT = 1
DEFAULT_LOGIN_WAIT_TIME = 3
DEFAULT_LOGIN_TIMEOUT = 10
DEFAULT_WEBSITE = "https://dispatch.uniuni.com/login"
DEFAULT_USERNAME = "Libin_Chen"
DEFAULT_PASSWORD = "Xbaqs1221@uniuni"

def login_uniuni(
    driver,
    website: str = DEFAULT_WEBSITE,
    username: str = DEFAULT_USERNAME,
    password: str = DEFAULT_PASSWORD,
    timeout: int = DEFAULT_LOGIN_TIMEOUT,
    post_login_wait: int = DEFAULT_LOGIN_WAIT_TIME
):
    """
    Log in to UniUni website using provided credentials.
    """

    # Open website
    driver.get(website)

    # Wait for username field and input credentials
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
    ).send_keys(username)

    driver.find_element(
        By.XPATH, "//input[@type='password']"
    ).send_keys(password)

    # Click login button
    driver.find_element(
        By.XPATH, "//button[.//span[text()='Login']]"
    ).click()

    # Wait for login to complete
    time.sleep(post_login_wait)

    return True

def open_new_tab_and_login(driver, website, username, password):
    driver.execute_script(f"window.open('{website}', '_blank');")
    driver.switch_to.window(driver.window_handles[-1])
    login_uniuni(driver, website, username, password)

# ===== Main Page Operation =====
def click_load(driver, timeout: int = DEFAULT_WAIT_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Load']]"))
    ).click()

def open_edit_order(
    driver,
    timeout: int = DEFAULT_WAIT_TIMEOUT
):
    """
    Click the 'Edit Order' button.
    """
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, "menu-edit-order"))
    ).click()

    return True

def click_batch_management(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, "menu-batch-management"))
    ).click()


# ===== Sub Batch Operation =====
def submit_batch_number(driver, batch_number: str, timeout: int = DEFAULT_WAIT_TIMEOUT):
    batch_input = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[@placeholder='Batch Number (Please use , separate)']")
        )
    )
    batch_input.clear()
    time.sleep(1)
    batch_input.send_keys(batch_number)
    time.sleep(1)

    driver.find_element(By.XPATH, "//button[.//span[text()='Submit']]").click()

def expand_buf_zone(driver, timeout: int = DEFAULT_WAIT_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//p[contains(text(),'BUF ZONE')]"))
    ).click()

    time.sleep(DEFAULT_SHORT_WAIT)

def get_route_red_value(driver, route: str, timeout: int = DEFAULT_WAIT_TIMEOUT) -> int:
    """
    Safely get the red chip value for a route.
    Returns 0 if element not found or value format is invalid.
    """
    try:
        # Wait for the route <p> element
        route_elem = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, f"//p[contains(text(),'{route}')]"))
        )

        # Find the red chip span
        span_elem = route_elem.find_element(
            By.XPATH,
            (
                "following-sibling::div//div[contains(@class,'MuiChip-root') "
                "and contains(@style,'background-color: red')]"
                "//span[contains(@class,'MuiChip-labelSmall')]"
            )
        )

        text = span_elem.text.strip()
        # Try to get number after '/'
        if '/' in text:
            return int(text.split('/')[1])
        else:
            return int(text)  # fallback if '/' not present

    except (NoSuchElementException, TimeoutException, ValueError, IndexError) as e:
        return 0

def click_driver_off(driver, timeout: int = DEFAULT_WAIT_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Driver Off']"))
    ).click()

    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Off']"))
    ).click()

    time.sleep(DEFAULT_SHORT_WAIT)

def get_quantity_by_sid(driver, sid: str, timeout: int = DEFAULT_WAIT_TIMEOUT) -> int:
    """
    Safely get the red chip quantity for a given scan ID.
    Returns 0 if element not found or value format is invalid.
    """
    try:
        # Wait for the accordion row
        row = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, f"accordion-row-{sid}"))
        )

        # Find the red chip span inside this row
        span_elem = row.find_element(
            By.XPATH,
            (
                ".//div[contains(@class,'MuiChip-root') "
                "and contains(@style,'background-color: red')]"
                "//span[contains(@class,'MuiChip-label')]"
            )
        )

        text = span_elem.text.strip()
        if '/' in text:
            return int(text.split('/')[1])
        else:
            return int(text)

    except (NoSuchElementException, TimeoutException, ValueError, IndexError) as e:
        return 0

def click_button_a(driver, timeout: int = DEFAULT_WAIT_TIMEOUT):
    """
    Click the 'A' button on the page.
    """
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[normalize-space()='A']")
        )
    ).click()
    return True

def input_dispatch_name(driver, name, timeout=20):
    dispatch_input = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((
            By.XPATH,
            '//label[text()="Dispatch History Name"]/following::input[1]'
        ))
    )
    time.sleep(1)
    dispatch_input.clear()
    time.sleep(1)
    dispatch_input.send_keys(name)
    time.sleep(1)

def click_save(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[.//span[text()="Save"]]')
        )
    ).click()

def click_submit(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[.//span[text()="Submit"]]')
        )
    ).click()


# ===== Edit Order Operation =====
def clear_search_input(driver, input_id: str = "searchSN", timeout: int = DEFAULT_WAIT_TIMEOUT):
    search_input = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, input_id))
    )
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(Keys.DELETE)

def get_driver_id(driver) -> str:
    try:
        return driver.find_element(
            By.XPATH,
            "//tr[th[normalize-space()='Driver ID']]/td[1]"
        ).text.strip()
    except NoSuchElementException:
        return "UNKNOWN"

def get_warehouse(driver) -> str:
    try:
        return driver.find_element(
            By.XPATH,
            "//tr[th[normalize-space()='Warehouse']]/td[1]"
        ).text.strip()
    except NoSuchElementException:
        return "UNKNOWN"

def get_status_list(driver) -> list[str]:
    try:
        items = driver.find_elements(
            By.XPATH, "//p[contains(@class,'MuiTypography-body2')]"
        )
        return [
            i.text.strip()
            for i in items
            if re.match(r"^\d+:\s*[A-Z_]+$", i.text.strip())
        ]
    except Exception:
        return []

def get_sub_batch(driver) -> str:
    try:
        elem = driver.find_element(
            By.XPATH,
            "//p[.//span[contains(text(),'Sub Batch')]]"
        )
        return elem.text.replace("Sub Batch:", "").strip()
    except NoSuchElementException:
        return "UNKNOWN"

def get_segment_info(driver) -> str:
    try:
        elem = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//span[normalize-space()='Segment:']/parent::p"
            ))
        )
        segment = elem.text.replace("Segment:", "").strip()
        return segment.split("-")[0]
    except Exception:
        return "N/A"

def get_storage_info(driver) -> str:
    try:
        elem = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[contains(@class,'MuiGrid-root')]/p"
            ))
        )
        return elem.text.split("Storage Info:")[1].strip()
    except Exception:
        return "N/A"

def parse_driver_memo(text: str) -> str:
    if not text:
        return "N/A"

    matches = re.findall(r"【([^】]+)】", text)
    if matches:
        parts = matches[0].split("-")
        return parts[1] if len(parts) >= 2 else matches[0]

    return text.split("/")[0].strip()

def search_parcel(
    driver,
    parcel_id: str,
    get_driver: bool = True,
    get_warehouse: bool = True,
    get_sub_batch: bool = True,
    get_status: bool = True,
    get_segment: bool = True,
    get_storage: bool = True,
    get_driver_memo: bool = True,
    wait_timeout: int = 10
) -> dict:
    """
    Search for a parcel by tracking number and return requested info as a dictionary.
    Only collects fields set to True.
    """

    wait = WebDriverWait(driver, wait_timeout)

    # ---- Step 1: Search input ----
    search_input = wait.until(
        EC.element_to_be_clickable((By.ID, "searchSN"))
    )
    search_input.send_keys(Keys.CONTROL, "a")
    search_input.send_keys(Keys.DELETE)
    search_input.send_keys(parcel_id, Keys.ENTER)

    # small wait for page to update
    time.sleep(2.5)

    # ---- Helper to safely get text ----
    def safe(xpath, default="N/A"):
        try:
            return driver.find_element(By.XPATH, xpath).text.strip()
        except:
            return default

    # ---- Step 2: Collect data ----
    data = {"parcel_id": parcel_id}

    if get_driver:
        data["driver_id"] = safe("//tr[th[normalize-space()='Driver ID']]/td[1]")

    if get_warehouse:
        data["warehouse"] = safe("//tr[th[normalize-space()='Warehouse']]/td[1]")

    if get_sub_batch:
        data["sub_batch"] = safe("//p[.//span[contains(text(),'Sub Batch')]]").replace("Sub Batch:", "").strip()

    if get_status:
        status = "NO_STATUS"
        try:
            items = driver.find_elements(By.XPATH, "//p[contains(@class,'MuiTypography-body2')]")
            valid = [i.text.strip() for i in items if re.match(r"^\d+:\s*[A-Z_]+$", i.text.strip())]
            if valid:
                status = valid[-1]
        except:
            pass
        data["status"] = status

    if get_segment:
        try:
            seg = driver.find_element(By.XPATH, "//span[normalize-space()='Segment:']/parent::p").text
            data["segment"] = seg.replace("Segment:", "").split("-")[0].strip()
        except:
            data["segment"] = "N/A"

    if get_storage:
        try:
            s = driver.find_element(By.XPATH, "//div[contains(@class,'MuiGrid-root')]/p").text
            data["storage"] = s.split("Storage Info:")[1].strip()
        except:
            data["storage"] = "N/A"

    if get_driver_memo:
        memo_text = safe("//tr[th[normalize-space()='Driver Memo']]/td[1]", "")
        data["driver_memo"] = parse_driver_memo(memo_text)

    return data

def update_driver_id(driver, driver_id, timeout=10):
    """
    Click the Change Driver icon for the driver row,
    enter the new driver ID, check the confirmation checkbox,
    and submit the form. Robust for Material-UI / React.
    """
    try:
        wait = WebDriverWait(driver, timeout)

        # 1️⃣ Find the row containing "Driver ID"
        driver_row = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//tr[th[contains(text(),'Driver ID')]]")
            )
        )

        # 2️⃣ Find the Change Driver button inside that row
        change_driver_btn = driver_row.find_element(
            By.XPATH,
            ".//button[contains(@class,'MuiIconButton-root') and contains(@class,'MuiIconButton-sizeSmall')]"
        )

        # 3️⃣ Scroll into view and click with ActionChains
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", change_driver_btn)
        time.sleep(0.3)
        actions = ActionChains(driver)
        actions.move_to_element(change_driver_btn).pause(0.3).click().perform()
        time.sleep(0.5)  # allow modal / input to appear

        # 4️⃣ Enter Driver ID
        input_box = wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//input[@type='number' and contains(@class,'MuiOutlinedInput-input')]")
            )
        )
        input_box.clear()
        input_box.send_keys(driver_id)
        time.sleep(0.2)

        # 5️⃣ Check the confirmation checkbox (JS click for MUI)
        checkbox = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='checkbox']")
            )
        )
        if not checkbox.is_selected():
            driver.execute_script("arguments[0].click();", checkbox)
        time.sleep(0.2)

        # 6️⃣ Click the LAST Submit button (dialog submit)
        time.sleep(1)

        submit_buttons = wait.until(
            EC.presence_of_all_elements_located(
                (By.XPATH, "//button[.//span[text()='Submit']]")
            )
        )

        submit_button = submit_buttons[-1]  # last one = active dialog

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_button)
        driver.execute_script("arguments[0].click();", submit_button)

        return True

    except (TimeoutException, StaleElementReferenceException) as e:
        print(f"[ERROR] Failed to change driver ID: {e}")
        return False

def open_operation_and_next_transition(driver, wait):
    # Click Operation accordion
    operation_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[text()='Operation']/ancestor::div[@role='button']")
        )
    )
    operation_btn.click()
    time.sleep(2)

    # Open Next Transition dropdown
    next_transition = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[@role='button' and @id='nextTransition']")
        )
    )
    next_transition.click()
    time.sleep(2)

def select_transition_option(driver, option_text):
    options = driver.find_elements(
        By.XPATH,
        f"//ul[@role='listbox']//option[normalize-space()='{option_text}']"
    )

    if not options:
        return False

    driver.execute_script("arguments[0].click();", options[0])
    time.sleep(2)
    return True

def submit_deliver_parcel_apt(driver, wait):
    # Click first visible Submit (dialog submit)
    submit_buttons = driver.find_elements(
        By.XPATH, "//button[.//span[text()='Submit']]"
    )

    for btn in submit_buttons:
        if btn.is_displayed():
            driver.execute_script("arguments[0].click();", btn)
            break

    time.sleep(3)

    # Open failed reason dropdown (MUI Select → ActionChains)
    dropdown = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "failed_reason_dialog_select_reason_textfield")
        )
    )
    time.sleep(3)

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", dropdown
    )
    time.sleep(3)

    ActionChains(driver) \
        .move_to_element(dropdown) \
        .pause(0.5) \
        .click() \
        .perform()

    time.sleep(3)

    # Select reason
    option = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//li[@role='option' and normalize-space()='Contact Failed and Inaccessible']"
            )
        )
    )
    driver.execute_script("arguments[0].click();", option)
    time.sleep(3)

    # Final submit (last Submit button = active dialog)
    submit_buttons = wait.until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//button[.//span[text()='Submit']]")
        )
    )
    time.sleep(3)

    final_submit = submit_buttons[-1]
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});", final_submit
    )
    driver.execute_script("arguments[0].click();", final_submit)
    time.sleep(3)

def send_parcel_to_storage(driver, timeout=10):
    wait = WebDriverWait(driver, timeout)

    try:
        open_operation_and_next_transition(driver, wait)

        # PART A: try SEND_PARCEL_TO_STORAGE directly
        if select_transition_option(driver, "SEND_PARCEL_TO_STORAGE"):
            final_submit = wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "nexttrasition_submit_timeout_button")
                )
            )
            driver.execute_script("arguments[0].click();", final_submit)
            time.sleep(3)
            return True

        # PART B: fallback → DELIVER_PARCEL_APT
        if not select_transition_option(driver, "DELIVER_PARCEL_APT"):
            return False

        submit_deliver_parcel_apt(driver, wait)
        time.sleep(3)

        # After DELIVER_PARCEL_APT, retry SEND_PARCEL_TO_STORAGE
        open_operation_and_next_transition(driver, wait)
        time.sleep(3)

        if not select_transition_option(driver, "SEND_PARCEL_TO_STORAGE"):
            return False

        final_submit = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "nexttrasition_submit_timeout_button")
            )
        )
        driver.execute_script("arguments[0].click();", final_submit)
        time.sleep(3)
        return True

    except Exception as e:
        print(f"[ERROR] send_parcel_to_storage failed: {e}")
        return False


# ===== Sub Batch Management Operation
def click_operate_shortcut(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'button[aria-label="operate shortcut"]')
        )
    ).click()

def create_blank_sub_batch(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[.//span[text()="create blank sub batch"]]')
        )
    ).click()

    # Close modal & refresh
    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
    driver.refresh()

def click_recent_7_days(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(
            (By.XPATH, '//button[.//span[text()="recent 7 days"]]')
        )
    ).click()

def get_recent_sub_batches(driver, limit=1, timeout=20):
    elements = WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, '//span[contains(text(), "BUSUB-")]')
        )
    )

    batches = [el.text.strip() for el in elements if el.text.strip()]
    return batches[:limit]






