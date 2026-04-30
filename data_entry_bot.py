# =============================================================================
# data_entry_bot.py
# Robust Selenium-based web form automation script.
#
# Author:  [Your Name]
# Purpose: Automated data entry from data.xlsx into a web portal.
#
# IMPORTANT: This script is intended for use by authorized personnel only.
#            Always ensure you have permission to automate interactions
#            with the target web portal.
# =============================================================================

import sys
import time
import random
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

import config  # Import our configuration file


# =============================================================================
# LOGGING SETUP
# Logs to both the console AND a timestamped log file for review after the run.
# =============================================================================

def setup_logging() -> logging.Logger:
    """
    Configures logging to output to both the terminal and a log file.
    Returns a configured logger instance.
    """
    log_filename = f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger = logging.getLogger("DataEntryBot")
    logger.setLevel(logging.DEBUG)

    # Formatter: includes timestamp, log level, and message
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console handler — shows INFO and above in the terminal
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # File handler — captures DEBUG and above in the log file
    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info(f"Logging initialized. Log file: '{log_filename}'")
    return logger


# =============================================================================
# DATA INGESTION
# =============================================================================

def load_data(filepath: str, sheet_name) -> pd.DataFrame:
    """
    Loads data from an Excel file into a pandas DataFrame.

    Args:
        filepath:   Path to the .xlsx file.
        sheet_name: Sheet index (int) or sheet name (str).

    Returns:
        A pandas DataFrame containing all rows from the sheet.

    Raises:
        SystemExit: If the file is not found or cannot be parsed.
    """
    logger = logging.getLogger("DataEntryBot")
    path = Path(filepath)

    if not path.exists():
        logger.critical(f"Data file not found: '{filepath}'. Please check the path in config.py.")
        sys.exit(1)

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
        # Strip leading/trailing whitespace from all string values
        df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
        # Replace pandas NaN with empty string for safe form filling
        df = df.fillna("")
        logger.info(f"Successfully loaded {len(df)} rows from '{filepath}' (Sheet: '{sheet_name}').")
        return df
    except Exception as e:
        logger.critical(f"Failed to read Excel file '{filepath}': {e}")
        sys.exit(1)


# =============================================================================
# BROWSER SETUP
# =============================================================================

def create_driver() -> webdriver.Chrome:
    """
    Initializes and returns a visible Chrome WebDriver instance.

    Features:
    - Uses webdriver-manager to automatically handle ChromeDriver versioning.
    - Loads an existing Chrome user profile to preserve cookies/login sessions.
    - Disables automation flags to reduce bot-detection by the portal.

    Returns:
        A configured, running selenium WebDriver instance.
    """
    logger = logging.getLogger("DataEntryBot")
    options = Options()

    # --- Visible Mode (non-headless) ---
    # Comment out the next two lines if you ever need headless operation.
    # options.add_argument("--headless=new")
    # options.add_argument("--disable-gpu")

    # --- Use Existing Chrome Profile ---
    # This loads your saved cookies, bookmarks, and login sessions.
    options.add_argument(f"--user-data-dir={config.CHROME_USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={config.CHROME_PROFILE_NAME}")

    # --- Window & Display Settings ---
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")

    # --- Stability & Performance ---
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")

    # --- Reduce Bot Detection Signals ---
    # Removes "Chrome is being controlled by automated software" banner.
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features=AutomationControlled")

    # --- Automatic ChromeDriver Management ---
    # webdriver-manager downloads the correct driver version automatically.
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        # Inject script to mask navigator.webdriver property (anti-detection)
        driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        logger.info("Chrome WebDriver started successfully.")
        return driver
    except WebDriverException as e:
        logger.critical(f"Failed to start Chrome WebDriver: {e}")
        logger.critical("Ensure Google Chrome is installed and try again.")
        sys.exit(1)


# =============================================================================
# LOGIN FLOW
# =============================================================================

def handle_login(driver: webdriver.Chrome) -> None:
    """
    Placeholder login function. Navigates to the login page and pauses
    execution, allowing you to manually log in and complete any CAPTCHA.

    HOW TO EXTEND THIS:
    If the portal has a simple login without CAPTCHA, you can automate it here:
        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys("your_username")
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys("your_password")
        driver.find_element(By.ID, "login_button").click()

    Args:
        driver: The active Selenium WebDriver instance.
    """
    logger = logging.getLogger("DataEntryBot")
    logger.info(f"Navigating to login page: {config.LOGIN_URL}")
    driver.get(config.LOGIN_URL)

    # Give the page a moment to load before prompting
    time.sleep(2)

    print("\n" + "=" * 65)
    print("  ACTION REQUIRED: MANUAL LOGIN")
    print("=" * 65)
    print("  The browser has navigated to the login page.")
    print("  Please:")
    print("    1. Enter your username and password in the browser.")
    print("    2. Complete any CAPTCHA or 2-Factor Authentication.")
    print("    3. Confirm you are fully logged in and on the dashboard.")
    print("  Once you are ready, return to this window and press Enter.")
    print("=" * 65 + "\n")

    input("  >>> Press ENTER to begin the data entry loop... ")

    logger.info("User confirmed login. Proceeding with data entry.")


# =============================================================================
# ELEMENT INTERACTION HELPERS
# =============================================================================

def get_by_type(locator_type: str) -> By:
    """
    Converts a string locator type from config.py into a Selenium By constant.

    Args:
        locator_type: String like 'id', 'xpath', 'name', 'css'.

    Returns:
        The corresponding selenium.webdriver.common.by.By constant.
    """
    mapping = {
        "id":    By.ID,
        "xpath": By.XPATH,
        "name":  By.NAME,
        "css":   By.CSS_SELECTOR,
        "class": By.CLASS_NAME,
        "tag":   By.TAG_NAME,
    }
    result = mapping.get(locator_type.lower())
    if result is None:
        raise ValueError(
            f"Unknown locator type '{locator_type}'. "
            f"Choose from: {list(mapping.keys())}"
        )
    return result


def wait_and_find(
    driver: webdriver.Chrome,
    locator_type: str,
    locator_value: str,
    timeout: int = None,
) -> webdriver.remote.webelement.WebElement:
    """
    Waits until an element is visible on the page, then returns it.
    This replaces fragile time.sleep() calls for page loading.

    Args:
        driver:        The active WebDriver instance.
        locator_type:  String locator type ('id', 'xpath', etc.).
        locator_value: The actual locator string (e.g., 'input_email').
        timeout:       Seconds to wait before raising TimeoutException.
                       Defaults to config.EXPLICIT_WAIT_TIMEOUT.

    Returns:
        The found and visible WebElement.

    Raises:
        TimeoutException: If element is not found within the timeout.
    """
    if timeout is None:
        timeout = config.EXPLICIT_WAIT_TIMEOUT

    by = get_by_type(locator_type)
    wait = WebDriverWait(driver, timeout)
    element = wait.until(
        EC.visibility_of_element_located((by, locator_value)),
        message=f"Timed out waiting for element: ({locator_type}, '{locator_value}')"
    )
    return element


def fill_text_field(element, value: str) -> None:
    """
    Clears a text input field and types a new value into it.

    Args:
        element: The Selenium WebElement for the input field.
        value:   The string value to type into the field.
    """
    element.clear()
    element.send_keys(str(value))


def select_dropdown_value(element, value: str) -> None:
    """
    Selects an option from a <select> dropdown element.
    Tries matching by visible text first, then by value attribute.

    Args:
        element: The Selenium WebElement for the <select> dropdown.
        value:   The option to select (by visible text or value attribute).
    """
    select = Select(element)
    try:
        select.select_by_visible_text(str(value))
    except NoSuchElementException:
        # Fallback: try matching by the 'value' attribute of the <option> tag
        select.select_by_value(str(value))


def handle_checkbox(element, value: str) -> None:
    """
    Checks or unchecks a checkbox based on the Excel cell value.
    The checkbox will be CHECKED if value is truthy (e.g., 'yes', 'true',
    '1', 'x', 'checked'), and UNCHECKED otherwise.

    Args:
        element: The Selenium WebElement for the checkbox.
        value:   The string value from the Excel cell.
    """
    # Define what counts as "truthy" for a checkbox
    truthy_values = {"yes", "true", "1", "x", "checked", "on", "y"}
    should_be_checked = str(value).strip().lower() in truthy_values

    is_currently_checked = element.is_selected()

    # Only click if the current state doesn't match the desired state
    if should_be_checked and not is_currently_checked:
        element.click()
    elif not should_be_checked and is_currently_checked:
        element.click()


# =============================================================================
# CORE DATA ENTRY FUNCTION (FOR A SINGLE ROW)
# =============================================================================

def fill_form_row(
    driver: webdriver.Chrome,
    row: pd.Series,
    row_index: int,
) -> bool:
    """
    Navigates to the data entry form and fills in all fields for one data row.

    Args:
        driver:    The active WebDriver instance.
        row:       A pandas Series representing one row of data (one Excel row).
        row_index: The 0-based index of the row, used for logging.

    Returns:
        True if the form was submitted successfully, False otherwise.
    """
    logger = logging.getLogger("DataEntryBot")

    # --- Navigate to the form URL ---
    logger.debug(f"Row {row_index}: Navigating to form URL: {config.FORM_URL}")
    driver.get(config.FORM_URL)

    # --- Wait for the page to indicate it's ready ---
    # We wait for the FIRST field in our map to appear as a proxy for page load.
    first_column = list(config.FIELD_MAP.keys())[0]
    first_locator = config.FIELD_MAP[first_column]
    logger.debug(f"Row {row_index}: Waiting for page to load (checking for first field)...")
    wait_and_find(driver, first_locator[0], first_locator[1])
    logger.debug(f"Row {row_index}: Page loaded. Beginning field population.")

    # --- Iterate over each field defined in our mapping ---
    for excel_column, (locator_type, locator_value) in config.FIELD_MAP.items():

        # Get the value from the pandas row for this column
        cell_value = row.get(excel_column, "")

        # Log each field action at DEBUG level (visible in the log file)
        logger.debug(
            f"Row {row_index}: Filling '{excel_column}' = '{cell_value}' "
            f"via ({locator_type}, '{locator_value}')"
        )

        # Find the element with an explicit wait
        element = wait_and_find(driver, locator_type, locator_value)

        # Scroll element into view to prevent interactability issues
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(0.2)  # Tiny pause after scroll for rendering

        # Interact with the element based on its type
        if excel_column in config.DROPDOWN_COLUMNS:
            select_dropdown_value(element, cell_value)

        elif excel_column in config.CHECKBOX_COLUMNS:
            handle_checkbox(element, cell_value)

        else:
            # Default: treat as a standard text input
            fill_text_field(element, cell_value)

    # --- Submit the form ---
    logger.debug(f"Row {row_index}: All fields filled. Locating submit button...")
    submit_locator = config.SUBMIT_BUTTON_LOCATOR
    submit_button = wait_and_find(driver, submit_locator[0], submit_locator[1])
    submit_button.click()
    logger.debug(f"Row {row_index}: Submit button clicked.")

    # --- Wait for confirmation that submission was successful ---
    time.sleep(config.PAGE_LOAD_WAIT_AFTER_SUBMIT)  # Brief pause for server response
    success_locator = config.SUCCESS_MESSAGE_LOCATOR
    wait_and_find(
        driver,
        success_locator[0],
        success_locator[1],
        timeout=config.EXPLICIT_WAIT_TIMEOUT,
    )
    logger.debug(f"Row {row_index}: Success confirmation element found.")
    return True


# =============================================================================
# RESULTS TRACKER CLASS
# =============================================================================

class RunTracker:
    """
    Tracks the success/failure of each row processed during the run.
    Provides a summary report at the end.
    """

    def __init__(self, total_rows: int):
        self.total = total_rows
        self.succeeded: list[int] = []  # List of successful row indices
        self.failed: list[dict] = []    # List of dicts with row index and error info

    def record_success(self, row_index: int) -> None:
        self.succeeded.append(row_index)

    def record_failure(self, row_index: int, error_message: str) -> None:
        self.failed.append({"row_index": row_index, "error": error_message})

    def print_summary(self) -> None:
        """Prints a formatted summary table to the console."""
        logger = logging.getLogger("DataEntryBot")
        separator = "=" * 65

        logger.info(separator)
        logger.info("  RUN COMPLETE — FINAL SUMMARY")
        logger.info(separator)
        logger.info(f"  Total Rows Processed : {self.total}")
        logger.info(f"  Successful           : {len(self.succeeded)}")
        logger.info(f"  Failed               : {len(self.failed)}")
        logger.info(separator)

        if self.failed:
            logger.warning("  FAILED ROWS (review manually):")
            logger.warning(f"  {'Row Index':<12} {'Error'}")
            logger.warning(f"  {'-'*12} {'-'*48}")
            for item in self.failed:
                # Truncate long error messages for display
                err_display = item["error"][:50] + "..." if len(item["error"]) > 50 else item["error"]
                logger.warning(f"  {item['row_index']:<12} {err_display}")
        else:
            logger.info("  All rows processed successfully!")

        logger.info(separator)


# =============================================================================
# MAIN ORCHESTRATOR
# =============================================================================

def main():
    """
    Main entry point for the data entry automation script.
    Orchestrates: logging setup → data load → browser launch →
                  login → data entry loop → cleanup → summary.
    """
    # 1. Setup logging
    logger = setup_logging()
    logger.info("=" * 65)
    logger.info("  Data Entry Bot — Starting Up")
    logger.info("=" * 65)

    # 2. Load data from Excel
    df = load_data(config.DATA_FILE, config.SHEET_NAME)

    # Validate that all expected columns exist in the DataFrame
    missing_cols = [col for col in config.FIELD_MAP if col not in df.columns]
    if missing_cols:
        logger.critical(
            f"The following columns from config.py FIELD_MAP are MISSING "
            f"in '{config.DATA_FILE}': {missing_cols}"
        )
        logger.critical("Please check your column names. Exiting.")
        sys.exit(1)

    # 3. Initialize results tracker
    tracker = RunTracker(total_rows=len(df))

    # 4. Launch browser
    driver = create_driver()

    try:
        # 5. Handle login (manual + CAPTCHA pause)
        handle_login(driver)

        # 6. Main data entry loop
        logger.info(f"Starting data entry loop for {len(df)} rows...")
        logger.info("-" * 65)

        for index, row in df.iterrows():
            # Display a clear per-row header in the log for easy reading
            logger.info(f"Processing Row {index + 1}/{len(df)} (DataFrame Index: {index})...")

            try:
                # --- Core action: fill and submit the form for this row ---
                success = fill_form_row(driver, row, row_index=index)

                if success:
                    tracker.record_success(index)
                    logger.info(f"  ✔ Row {index + 1}: SUCCESS")

            # --- Granular exception handling ---

            except TimeoutException as e:
                # Element didn't appear within the wait timeout
                error_msg = f"TimeoutException — An element was not found in time. Detail: {e.msg}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)

            except NoSuchElementException as e:
                # Element locator is wrong or element doesn't exist on page
                error_msg = f"NoSuchElementException — Element not found. Detail: {e.msg}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)

            except ElementNotInteractableException as e:
                # Element found but is hidden, disabled, or covered by another element
                error_msg = f"ElementNotInteractableException — Cannot interact with element. Detail: {e.msg}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)

            except StaleElementReferenceException as e:
                # DOM was refreshed between finding and interacting with element
                error_msg = f"StaleElementReferenceException — Element no longer attached to DOM. Detail: {e.msg}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)

            except (ValueError, KeyError) as e:
                # Issues with data (e.g., dropdown value not in options list)
                error_msg = f"Data/Mapping Error — {type(e).__name__}: {e}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)

            except WebDriverException as e:
                # Broad Selenium error (e.g., browser crash, connection lost)
                error_msg = f"WebDriverException — Selenium-level error: {e.msg}"
                logger.error(f"  ✘ Row {index + 1}: FAILED — {error_msg}")
                tracker.record_failure(index, error_msg)
                # For browser-level errors, add extra wait before next row
                logger.warning("  Waiting 5 seconds after WebDriverException before continuing...")
                time.sleep(5)

            finally:
                # --- Rate Limiting ---
                # Runs after EVERY row (success or failure) to pace requests.
                delay = random.uniform(config.RATE_LIMIT_MIN, config.RATE_LIMIT_MAX)
                logger.debug(f"  Rate limiting: sleeping for {delay:.2f} seconds...")
                time.sleep(delay)

        # 7. Print final summary
        tracker.print_summary()

    except KeyboardInterrupt:
        # Graceful handling if the user presses Ctrl+C to stop the script
        logger.warning("\nScript interrupted by user (Ctrl+C). Generating partial summary...")
        tracker.print_summary()

    finally:
        # 8. Always close the browser, even if an unhandled error occurred
        logger.info("Closing browser and cleaning up...")
        try:
            driver.quit()
        except Exception:
            pass  # Ignore errors during cleanup
        logger.info("Browser closed. Script finished.")


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    main()
