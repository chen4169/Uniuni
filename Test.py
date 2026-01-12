from Setup import *
from Operation import *
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time

# -----------------------------
# Start browser
# -----------------------------
driver = init_edge_driver()
switch_to_target_tab(driver)

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time

wait = WebDriverWait(driver, 10)
driver_sheep = "66017"  
parcel_id = "UUS5BR3922435548357"

# Open parcel A detail page
parcel_info = search_parcel(
            driver, parcel_id,
            get_driver=True,
            get_warehouse=False,
            get_sub_batch=False,
            get_status=False,
            get_segment=False,
            get_storage=True,
            get_driver_memo=False
        )

print(parcel_info)

driver_id = parcel_info.get("driver_id")
storage_info = parcel_info.get("storage")

print(driver_id)
if isinstance(driver_id, str) and driver_id.startswith("45000"):
                update_driver_id(driver, driver_id=driver_sheep)
                time.sleep(3)

# Send parcel A to storage
success = send_parcel_to_storage(driver)

if success:
    print(f"✅ Parcel {parcel_id} sent to storage successfully")
else:
    print(f"❌ Failed to send parcel {parcel_id} to storage")





# print("STEP 1: Click Operation accordion")
# operation_btn = wait.until(
#     EC.element_to_be_clickable(
#         (By.XPATH, "//p[text()='Operation']/ancestor::div[@role='button']")
#     )
# )
# operation_btn.click()
# time.sleep(2)

# # ------------------------------

# print("STEP 2: Open Next Transition dropdown")
# next_transition = wait.until(
#     EC.element_to_be_clickable(
#         (By.XPATH, "//div[@role='button' and @id='nextTransition']")
#     )
# )
# next_transition.click()
# time.sleep(2)


# # ===========================
# # Part A
# # If SEND_PARCEL_TO_STORAGE option exist, then click it. If not, go to Part B directly
# # ===========================
# print("STEP 3: Look for SEND_PARCEL_TO_STORAGE option")
# send_to_storage_options = driver.find_elements(
#     By.XPATH,
#     "//ul[@role='listbox']//option[normalize-space()='SEND_PARCEL_TO_STORAGE']"
# )

# print(f"Found SEND_PARCEL_TO_STORAGE options: {len(send_to_storage_options)}")

# if send_to_storage_options:
#     print("STEP 4: Click SEND_PARCEL_TO_STORAGE")
#     driver.execute_script("arguments[0].click();", send_to_storage_options[0])
#     time.sleep(2)

#     print("STEP 5: Click final submit (timeout submit button)")
#     final_submit = wait.until(
#         EC.element_to_be_clickable(
#             (By.ID, "nexttrasition_submit_timeout_button")
#         )
#     )
#     driver.execute_script("arguments[0].click();", final_submit)
#     print("✅ DONE: Sent to storage directly")
    
# else:
#     print("SEND_PARCEL_TO_STORAGE not found, fallback path")

# # ===========================
# # Part B
# # Try to click DELIVER_PARCEL_APT option, after success doing this, go back to Part A.
# # If this option does not exist, end the function directly and return False
# # ===========================
# print("STEP 6: Click DELIVER_PARCEL_APT")
# deliver_options = driver.find_elements(
#     By.XPATH,
#     "//ul[@role='listbox']//option[normalize-space()='DELIVER_PARCEL_APT']"
# )

# if not deliver_options:
#     print("❌ DELIVER_PARCEL_APT not found, abort")

# driver.execute_script("arguments[0].click();", deliver_options[0])
# time.sleep(2)


# print("STEP 7: Submit DELIVER_PARCEL_APT")

# submit_buttons = driver.find_elements(
#     By.XPATH, "//button[.//span[text()='Submit']]"
# )

# for btn in submit_buttons:
#     if btn.is_displayed():
#         driver.execute_script("arguments[0].click();", btn)
#         break

# time.sleep(3)

# # Click dropdown to open options
# dropdown = wait.until(
#     EC.visibility_of_element_located(
#         (By.ID, "failed_reason_dialog_select_reason_textfield")
#     )
# )

# # Scroll into view (important for MUI)
# driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
# time.sleep(0.5)

# # Real mouse click
# ActionChains(driver) \
#     .move_to_element(dropdown) \
#     .pause(0.5) \
#     .click() \
#     .perform()

# time.sleep(1)

# # choose an option
# option = wait.until(
#     EC.presence_of_element_located(
#         (
#             By.XPATH,
#             "//li[@role='option' and normalize-space()='Contact Failed and Inaccessible']"
#         )
#     )
# )

# driver.execute_script("arguments[0].click();", option)
# time.sleep(1)

# # click submit
# submit_buttons = wait.until(
#             EC.presence_of_all_elements_located(
#                 (By.XPATH, "//button[.//span[text()='Submit']]")
#             )
#         )

# submit_button = submit_buttons[-1]  # last one = active dialog

# driver.execute_script("arguments[0].scrollIntoView({block:'center'});", submit_button)
# driver.execute_script("arguments[0].click();", submit_button)
