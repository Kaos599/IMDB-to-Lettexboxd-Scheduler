# IMDB-to-Lettexboxd-Scheduler

![Image](https://github.com/user-attachments/assets/5ccd3a74-2da4-4617-8a66-cde90ba825c9)

[![Python](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/selenium-%20-4682B4.svg)](https://www.selenium.dev/)
[![PyAutoGUI](https://img.shields.io/badge/pyautogui-0.9.53-brightgreen)](https://pyautogui.readthedocs.io/en/latest/)


This Python script automates the transfer of movie ratings from IMDb to Letterboxd. It addresses the common issue of manual transfer and simplifies the process for users.  This tool handles dynamic web elements and mimics human-like behavior to avoid bot detection, ensuring a smooth and reliable transfer.  Built using Selenium for web automation and PyAutoGUI for operating system-level interactions.

## Features

* **Automated Login:** Seamlessly logs in to both IMDb and Letterboxd using your credentials.
* **IMDb Export:**  Navigates to your IMDb ratings, exports them as a CSV, and handles the download process intelligently. Includes smart waiting with user-friendly messages and robust error handling.
* **Letterboxd Import:**  Uploads the CSV to Letterboxd, manages file dialogs automatically, and initiates the import, providing progress updates and error management.
* **Headless Mode:** Runs discreetly in the background without a visible browser window. (Configurable)
* **Automatic Cleanup:** Deletes the temporary CSV file after import, keeping your system tidy.
* **Human-like Behavior:**  Uses random delays and actions (scrolling, clicks) to minimize the risk of bot detection.
* **Scheduled Execution (Windows):**  Easily schedule automatic imports as a Windows task with customizable settings. Create, update or delete scheduled task.
* **Detailed Logging:** Provides comprehensive logs for easy debugging and monitoring.
* **Robust Error Handling:** Prevents unexpected crashes with informative error messages.

## Requirements

* **Python 3.7+:** Ensure you have Python 3.7 or a later version installed.
* **Libraries:** Install required libraries using `pip install selenium pyautogui pywin32`
* **Firefox WebDriver:** Download the `geckodriver` executable and add it to your system's PATH.  See [GeckoDriver Releases](https://github.com/mozilla/geckodriver/releases)
* **uBlock Origin:**  Install the uBlock Origin extension for Firefox (`ublock_origin-latest.xpi`). This helps in blocking ads and other unnecessary elements that may cause problems with the web scraping process.
* **Credentials:**  Your IMDb and Letterboxd email/password, and your IMDb user ID.


## Usage

1. **Configuration:**
   * Update the `Configuration Section` of the script with your IMDb email, password, Letterboxd email, password, IMDb User ID (find it in your IMDb profile URL), preferred download directory. Consider using environment variables for sensitive data like passwords and API keys. This enhances security and maintainability.
   * Adjust the `HEADLESS_MODE` and scheduled task settings as needed.
   * If the default download directory is not suitable, modify the `DOWNLOAD_DIR` variable.

2. **Running the Script:**
   * Execute the script.  On the first run, it creates a Windows scheduled task for daily execution.  Subsequent runs offer options to update or delete the task.  You can also choose to run the import immediately.

## Scheduling (Windows)

The script streamlines scheduled execution:

* **First Run:** Creates a scheduled task with the specified name (`TASK_NAME`) at the designated time.
* **Subsequent Runs:** Offers options to update the existing task or delete it.

## How It Works

1. **WebDriver Setup:** Initializes the Firefox WebDriver with uBlock Origin and download preferences.
2. **IMDb Export:** Logs into IMDb, navigates to your ratings, initiates the export, and monitors its status until download completion.
3. **Letterboxd Import:** Logs into Letterboxd, uploads the downloaded CSV, handles the matching process, and triggers the import.
4. **Task Scheduling:**  Manages the creation, update, or deletion of the Windows scheduled task.
5. **Cleanup:** Deletes the temporary CSV file.


## Troubleshooting

* **Dependencies:** Verify all required libraries are installed correctly using pip and are compatible with the script's version requirements. Ensure that the Firefox WebDriver is the correct version and is accessible in your system's PATH.
* **Scheduled Tasks:**  For issues with scheduled tasks, confirm the Python executable path (`PYTHON_EXECUTABLE`) and script path (`SCRIPT_PATH`) are accurate. If there are errors about missing libraries or permission issues, consider if your PYTHON_EXECUTABLE points to correct virtual environment if one is used. Try running the command used for task creation directly from the console and see if there are any errors outputted by the Task Scheduler itself.
* **Selenium Errors:**  If Selenium encounters problems, check the WebDriver version and PATH. Consider updating to the latest geckodriver. Also use explicit waits to wait for elements to load completely before attempting interaction, addressing potential timing issues with dynamic page content.  Inspect the website for any changes in structure using browser developer tools and adjust locators if needed.
* **Login Issues:** Double-check your IMDb and Letterboxd credentials.  If the website structure has changed, the login elements may need updating.  Implement `WebDriverWait` to handle dynamically loading login elements.
* **Import Failures:**  If the Letterboxd import fails, increase the timeout values for `WebDriverWait` to accommodate slower loading times.

## Disclaimer

This script is provided as-is. It depends on the structure of IMDb and Letterboxd, which may change, affecting compatibility.  Use it at your own risk.
