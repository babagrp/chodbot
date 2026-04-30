# =============================================================================
# config.py
# Centralized configuration. Edit these values to match your specific portal.
# =============================================================================

# --- File Settings ---
DATA_FILE = "data.xlsx"
SHEET_NAME = 0  # Use 0 for the first sheet, or provide a sheet name string.

# --- URL Settings ---
LOGIN_URL = "https://your-portal.com/login"       # Replace with your login URL
FORM_URL  = "https://your-portal.com/data-entry"  # Replace with your form URL

# --- Browser Profile Settings ---
# To find your Chrome profile path:
#   Windows: C:\Users\<YourUser>\AppData\Local\Google\Chrome\User Data
#   macOS:   /Users/<YourUser>/Library/Application Support/Google/Chrome
#   Linux:   /home/<YourUser>/.config/google-chrome
CHROME_USER_DATA_DIR = r"C:\Users\YourUser\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE_NAME  = "Default"  # Usually "Default" or "Profile 1", "Profile 2", etc.

# --- Excel Column → HTML Element Mapping ---
# Key   = Exact column header name from your Excel file
# Value = The locator used to find the element on the web page
#
# Format for values:
#   ("id", "actual-html-id")
#   ("xpath", "//your/xpath/here")
#   ("name", "input-name-attribute")
#   ("css", "css-selector")

FIELD_MAP = {
    # --- Text Input Fields ---
    # "Excel Column Name": ("locator_type", "locator_value")
    "FullName":       ("id",    "input_full_name"),       # <-- REPLACE WITH REAL ID
    "EmailAddress":   ("id",    "input_email"),            # <-- REPLACE WITH REAL ID
    "PhoneNumber":    ("id",    "input_phone"),            # <-- REPLACE WITH REAL ID
    "Address":        ("xpath", "//input[@name='address']"), # <-- REPLACE WITH REAL XPATH

    # --- Dropdown Field ---
    # Uses Selenium's Select class. Value should be the element locator.
    "Category":       ("id",    "select_category"),       # <-- REPLACE WITH REAL ID

    # --- Checkbox Field ---
    # Script will CHECK the box if the Excel value is truthy (True, 'yes', 'x', 1, etc.)
    "IsActive":       ("id",    "checkbox_is_active"),    # <-- REPLACE WITH REAL ID

    # Add more fields here following the same pattern...
}

# Which columns from FIELD_MAP are DROPDOWNS? (list their Excel column names)
DROPDOWN_COLUMNS = ["Category"]

# Which columns from FIELD_MAP are CHECKBOXES? (list their Excel column names)
CHECKBOX_COLUMNS = ["IsActive"]

# --- Locators for Submit Button and Success Confirmation ---
SUBMIT_BUTTON_LOCATOR  = ("id",    "btn_submit")         # <-- REPLACE
SUCCESS_MESSAGE_LOCATOR = ("xpath", "//div[@class='alert-success']") # <-- REPLACE

# --- Timing Settings (seconds) ---
EXPLICIT_WAIT_TIMEOUT   = 15    # Max seconds to wait for an element to appear
RATE_LIMIT_MIN          = 1.5   # Minimum random sleep between submissions
RATE_LIMIT_MAX          = 3.5   # Maximum random sleep between submissions
PAGE_LOAD_WAIT_AFTER_SUBMIT = 2 # Fixed pause after clicking submit
