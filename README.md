Setup & Usage Guide

# ============================================================
# STEP 1: Create and activate a virtual environment (recommended)
# ============================================================
python -m venv venv

# Windows:
venv\Scripts\activate

# macOS / Linux:
source venv/bin/activate


# ============================================================
# STEP 2: Install all dependencies
# ============================================================
pip install -r requirements.txt


# ============================================================
# STEP 3: Configure the script
# ============================================================
# Edit config.py and replace ALL placeholder values:
#
#   LOGIN_URL             → Your portal's login page URL
#   FORM_URL              → Your portal's data entry form URL
#   CHROME_USER_DATA_DIR  → Your Chrome user data directory path
#   CHROME_PROFILE_NAME   → Your profile folder name ("Default", etc.)
#
#   FIELD_MAP             → Map your Excel column names to element locators
#   DROPDOWN_COLUMNS      → List which of your columns are dropdowns
#   CHECKBOX_COLUMNS      → List which of your columns are checkboxes
#
#   SUBMIT_BUTTON_LOCATOR  → Locator for the form's submit button
#   SUCCESS_MESSAGE_LOCATOR → Locator for the post-submit success element


# ============================================================
# STEP 4: Prepare your data file
# ============================================================
# Place data.xlsx in the same folder as data_entry_bot.py
# Ensure the column headers EXACTLY match the keys in FIELD_MAP


# ============================================================
# STEP 5: Run the script
# ============================================================
python data_entry_bot.py

# The browser will open → navigate to the login page
# → YOU log in manually and solve any CAPTCHA
# → Press Enter in the terminal
# → The script processes all rows automatically





Quick Reference: Finding Locators in Chrome

To find the correct ID/XPath for FIELD_MAP in config.py:

1. Open your form in Chrome
2. Right-click the input field → "Inspect"
3. In the DevTools HTML panel, look for:

   ID:    <input id="THIS_IS_THE_ID" ...>
           → Use: ("id", "THIS_IS_THE_ID")

   XPath: Right-click the element in DevTools
           → Copy → Copy XPath
           → Use: ("xpath", "//pasted/xpath/here")

   Name:  <input name="THIS_IS_THE_NAME" ...>
           → Use: ("name", "THIS_IS_THE_NAME")
