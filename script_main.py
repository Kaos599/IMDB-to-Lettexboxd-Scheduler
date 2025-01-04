"""IMDb to Letterboxd Ratings Importer

This script automates the transfer of movie ratings from IMDb to Letterboxd.
It handles dynamic web elements and mimics human behavior to avoid bot detection.
Utilizes Selenium for web automation and PyAutoGUI for OS-level interactions.
"""

import os
import time
import csv
import subprocess
import logging
import random
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


"""You can find your IMDb User ID by logging in to your IMDb account, navigating to your profile page, and then copying the string of characters from the URL. 
The URL of your profile page will look something like this: "https://www.imdb.com/user/urxxxxxxxxxxxx/". 
In this example, your IMDb User ID would be "urxxxxxxxxxxxx". 
Make sure to copy the part after /user/ and before the final /."""

# Configuration Section
# Replace placeholder values with your actual credentials and desired settings.
IMDB_EMAIL = "your_imdb_email"  # Your IMDb login email address.
IMDB_PASSWORD = "your_imdb_password"  # Your IMDb login password.
LETTERBOXD_EMAIL = "your_letterboxd_email"  # Your Letterboxd login email address.
LETTERBOXD_PASSWORD = "your_letterboxd_password"  # Your Letterboxd login password.
IMDB_USER_ID = "your_imdb_user_id"  # Your IMDb user ID, found on your profile page URL.
DOWNLOAD_DIR = r"downloads"  # The directory where downloaded files will be saved. Relative or absolute path.
HEADLESS_MODE = True  # Set to True to run the browser in the background, False to see the browser window.

# Scheduled Task Configuration (Windows Only)
TASK_NAME = "IMDb_to_Letterboxd_Task"  # The name of the scheduled task in Windows Task Scheduler.
PYTHON_EXECUTABLE = r"C:\Path\To\Your\python.exe"  # The full path to your Python executable.
SCRIPT_PATH = r"C:\Path\To\Your\Script.py"  # The full path to this Python script.
SCHEDULED_TIME = "10:00"  # The time of day to run the scheduled task (in 24-hour format, e.g., "14:30").

# Logging setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Helper Functions
def random_delay(min_delay: int = 2, max_delay: int = 5) -> None:
    """
    Introduce random delays to mimic human behavior.

    Args:
        min_delay: The minimum delay in seconds.
        max_delay: The maximum delay in seconds.
    """
    delay = random.uniform(min_delay, max_delay)
    time.sleep(delay)

def random_scroll(driver: webdriver.Firefox) -> None:
    """
    Perform random scrolls to mimic human behavior.

    Args:
        driver: The Selenium WebDriver instance.
    """
    for _ in range(random.randint(1, 3)):
        driver.execute_script(f"window.scrollBy(0, {random.randint(50, 200)})")
        random_delay(0.5, 1.5)

def random_clicks(driver: webdriver.Firefox) -> None:
    """
    Perform random clicks on the page to mimic human behavior.

    Args:
        driver: The Selenium WebDriver instance.
    """
    body = driver.find_element(By.TAG_NAME, "body")
    for _ in range(random.randint(0, 2)):
        try:
            width = driver.execute_script("return document.documentElement.clientWidth;")
            height = driver.execute_script("return document.documentElement.clientHeight;")
            x = random.randint(10, width - 10)
            y = random.randint(10, height - 10)
            action = webdriver.ActionChains(driver)
            action.move_by_offset(x, y).click().perform()
            random_delay(0.3, 0.8)
        except Exception as e:
            logging.error(f"Error during random click: {e}")

def setup_driver() -> webdriver.Firefox:
    """
    Setup Selenium WebDriver (Firefox) with specified options and extensions.

    Returns:
        The configured Firefox WebDriver instance.
    """
    extension_path = 'ublock_origin-1.61.2.xpi'  # Path to the uBlock Origin extension.
    options = webdriver.FirefoxOptions()
    profile = webdriver.FirefoxProfile()

    if HEADLESS_MODE:
        options.add_argument("--headless")

    # Configure download preferences to avoid prompts
    profile.set_preference("browser.download.folderList", 2)  # Use custom download directory
    profile.set_preference("browser.download.manager.showWhenStarting", False)  # Don't show download manager
    profile.set_preference("browser.download.dir", os.path.abspath(DOWNLOAD_DIR))  # Set download directory
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")  # Auto-save CSV files
    options.profile = profile
    driver = webdriver.Firefox(options=options)
    driver.install_addon(extension_path, temporary=True)  # Install uBlock Origin
    return driver

def login_to_imdb(driver: webdriver.Firefox) -> bool:
    """
    Log in to IMDb using the provided credentials.

    Args:
        driver: The Selenium WebDriver instance.

    Returns:
        True if login is successful, False otherwise.
    """
    logging.info("Navigating to IMDb login page...")
    driver.get("https://www.imdb.com/registration/signin")
    random_delay()

    try:
        # Find and click the IMDb sign-in button
        signin_imdb_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "list-group-item"))
        )
        signin_imdb_button.click()
        random_delay()

        # Enter email
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ap_email"))
        )
        email_input.send_keys(IMDB_EMAIL)
        random_delay()

        # Enter password
        password_input = driver.find_element(By.ID, "ap_password")
        password_input.send_keys(IMDB_PASSWORD)
        random_delay()

        # Submit login form
        signin_submit_button = driver.find_element(By.ID, "signInSubmit")
        signin_submit_button.click()
        random_delay()
        logging.info("Logged in to IMDb successfully.")
        return True
    except TimeoutException:
        logging.error("Login to IMDb failed: Could not find login elements.")
        return False

def navigate_to_imdb_ratings(driver: webdriver.Firefox) -> None:
    """
    Navigate to the user ratings page on IMDb.

    Args:
        driver: The Selenium WebDriver instance.
    """
    ratings_url = f"https://www.imdb.com/user/{IMDB_USER_ID}/ratings/"
    logging.info(f"Navigating to IMDb ratings page: {ratings_url}")
    driver.get(ratings_url)
    random_delay()

def initiate_imdb_export(driver: webdriver.Firefox) -> bool:
    """
    Initiate the export of IMDb ratings.

    Args:
        driver: The Selenium WebDriver instance.

    Returns:
        True if export initiation is successful, False otherwise.
    """
    logging.info("Initiating IMDb ratings export...")
    try:
        # Find and click the export button
        export_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Export')]"))
        )
        export_button.click()
        random_delay()
        logging.info("Export initiated.")
        return True
    except TimeoutException:
        logging.error("Failed to initiate IMDb export: Export button not found.")
        return False

def wait_for_download_completion(download_dir: str, timeout: int = 60) -> str:
    """
    Wait for a .csv file to complete downloading in the specified directory.

    Args:
        download_dir: Directory to monitor for downloads.
        timeout: Maximum time to wait in seconds.

    Returns:
        The path to the downloaded file.

    Raises:
        TimeoutError: If the download does not complete within the timeout.
    """
    logging.info(f"Waiting for download to complete in: {download_dir}")

    start_time = time.time()
    while time.time() - start_time < timeout:
        # Look for both .csv and .csv.part files
        files = os.listdir(download_dir)
        csv_files = [f for f in files if f.endswith('.csv')]
        part_files = [f for f in files if f.endswith('.csv.part')]

        # If we find a complete CSV file and no .part files, download is done
        if csv_files and not part_files:
            # Get the most recent CSV file
            latest_csv = max(
                [os.path.join(download_dir, f) for f in csv_files],
                key=os.path.getmtime
            )

            # Ensure file size has stabilized to confirm download completion
            size = os.path.getsize(latest_csv)
            time.sleep(1)  # Wait a second
            if size == os.path.getsize(latest_csv):
                logging.info(f"Download completed: {latest_csv}")
                return latest_csv

        time.sleep(0.5)  # Check every half second

    raise TimeoutError("Download did not complete within the specified timeout")

def monitor_imdb_export_status(driver: webdriver.Firefox) -> str or None:
    """
    Monitor the export status on IMDb and wait for the download to become ready.

    Args:
        driver: The Selenium WebDriver instance.

    Returns:
        The path to the downloaded CSV file if successful, None otherwise.
    """
    exports_url = "https://www.imdb.com/exports/"
    logging.info(f"Monitoring IMDb export status at: {exports_url}")
    driver.get(exports_url)
    random_delay()

    max_attempts = 30
    attempt = 0

    while attempt < max_attempts:
        try:
            # Wait for the list of exports to load
            list_content = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="list-page-mc-list-content"]'))
            )

            # Find all export items
            export_items = list_content.find_elements(By.CSS_SELECTOR, 'li[data-testid="user-ll-item"]')

            if export_items:
                first_item = export_items[0]
                try:  # Try to find the Ready button
                    ready_button = first_item.find_element(
                        By.CSS_SELECTOR,
                        'button[data-testid="export-status-button"].READY'
                    )

                    logging.info("Found Ready button. Initiating download...")
                    random_delay()
                    ready_button.click()

                    # Wait for download to complete
                    try:
                        downloaded_file = wait_for_download_completion(DOWNLOAD_DIR)
                        if downloaded_file:
                            return downloaded_file
                    except TimeoutError as e:
                        logging.error(f"Download timeout: {e}")
                        return None

                except NoSuchElementException:
                    print("Ready button not available yet, Trying again")

            else:  # Handle case where no exports are listed
                print("Ready button not available yet, Trying again")

            attempt += 1
            logging.info(f"Export not ready yet. Attempt {attempt}/{max_attempts}")
            random_delay(10, 15)
            driver.refresh()  # Refresh the page to check for updates
            random_delay()

        except Exception as e:
            logging.error(f"Error checking export status: {e}")
            attempt += 1
            random_delay(5, 10)
            driver.refresh()

    logging.error("Maximum attempts reached without finding a ready export")
    return None

def login_to_letterboxd(driver: webdriver.Firefox) -> bool:
    """
    Log in to Letterboxd using the provided credentials.

    Args:
        driver: The Selenium WebDriver instance.

    Returns:
        True if login is successful, False otherwise.
    """
    logging.info("Navigating to Letterboxd login page...")
    driver.get("https://letterboxd.com/settings/data/")
    random_delay()

    try:
        # Find and enter the username
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "field-username"))
        )
        logging.info("Found username input field.")
        username_input.send_keys(LETTERBOXD_EMAIL)
        random_delay(1, 2)

        # Find and enter the password
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "field-password"))
        )
        logging.info("Found password input field.")
        password_input.send_keys(LETTERBOXD_PASSWORD)
        random_delay(1, 2)

        # Find and click the login button
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
        )
        logging.info("Clicking login button.")
        login_button.click()
        random_delay(2, 3)

        # Verify successful login by checking for an element on the account settings page
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//h1[@class='title-hero -textleft js-hide-in-app' and text()='Account Settings']"))
        )

        logging.info("Successfully logged into Letterboxd")

        # Navigate to the import page after successful login
        driver.get("https://letterboxd.com/import/")
        logging.info("Navigated to Letterboxd import page.")

        return True

    except Exception as e:
        logging.exception(f"Failed to login to Letterboxd: {e}")
        return False

def get_latest_csv_from_downloads() -> str or None:
    """
    Get the path of the most recently downloaded CSV file from the downloads directory.

    Returns:
        The path to the latest CSV file, or None if no CSV files are found.
    """
    try:
        csv_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.csv')]
        if not csv_files:
            logging.error("No CSV files found in downloads directory")
            return None

        # Find the CSV file with the most recent modification time
        latest_csv = max([os.path.join(DOWNLOAD_DIR, f) for f in csv_files],
                            key=os.path.getmtime)
        logging.info(f"Found latest CSV file: {latest_csv}")
        return latest_csv
    except Exception as e:
        logging.error(f"Error finding latest CSV file: {e}")
        return None

def wait_for_element_with_text(driver: webdriver.Firefox, text: str, timeout: int = 10) -> webdriver.remote.webelement.WebElement or None:
    """
    Wait for an element containing specific text to appear.

    Args:
        driver: The Selenium WebDriver instance.
        text: The text to search for within the element.
        timeout: The maximum time to wait in seconds.

    Returns:
        The WebElement if found, None otherwise.
    """
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{text}')]"))
        )
    except TimeoutException:
        return None

def import_to_letterboxd(driver: webdriver.Firefox, csv_path: str) -> bool:
    """
    Import ratings to Letterboxd using the downloaded CSV file.

    Args:
        driver: The Selenium WebDriver instance.
        csv_path: The path to the downloaded IMDb ratings CSV file.

    Returns:
        True if the import is successful, False otherwise.
    """
    logging.info("Starting Letterboxd import process...")

    try:
        # Navigate to the IMDb import section on the import page
        driver.get("https://letterboxd.com/import/#imdb-import")
        random_delay(2, 3)

        # Check for and handle an intermittent "Continue" button
        try:
            continue_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, '_2Sbg_-vS') and contains(@class, '_1CPr_04v') and contains(@class, 'f5CXro_f') and contains(@class, '_9eRJjhym') and text()='Continue']"))
            )
            logging.info("Found the 'Continue' button. Clicking it.")
            continue_button.click()
            random_delay(1, 2)
        except TimeoutException:
            logging.info("The 'Continue' button was not found, proceeding with import.")
        except Exception as e:
            logging.error(f"Error while handling the 'Continue' button: {e}")

        # Find and click the link to select the IMDb CSV file
        select_file_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '#imdb-import')]"))
        )
        select_file_button.click()
        random_delay()

        # Locate the file input element and send the path to the CSV file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )
        file_input.send_keys(os.path.abspath(csv_path))
        logging.info("CSV file selected for upload")

        time.sleep(2)  # Give time for the file dialog to open
        pyautogui.press('enter')  # Simulate pressing Enter to close the file dialog
        time.sleep(1)
        pyautogui.press('enter')

        # Wait for the matching process to complete
        matching_complete = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Matching complete')]"))
        )
        logging.info("Film matching process completed")
        random_delay(2, 3)

        # Find and click the "Import Films" button
        import_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Import Films')]"))
        )
        import_button.click()
        logging.info("Started importing films")

        # Wait for the import completion message
        saved_films_element = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Saved')]"))
        )

        # Extract the number of saved films from the completion message
        saved_films_text = saved_films_element.text
        num_films = ''.join(filter(str.isdigit, saved_films_text))
        logging.info(f"Successfully imported {num_films} films to Letterboxd")

        return True
    except Exception as e:
        logging.error(f"Error during Letterboxd import: {e}")
        return False

def schedule_task(task_name: str, python_path: str, script_path: str, scheduled_time: str, task_action: str = "create") -> None:
    """
    Schedule or modify a task in Windows Task Scheduler.

    Args:
        task_name: The name of the scheduled task.
        python_path: The path to the Python executable.
        script_path: The path to the Python script.
        scheduled_time: The time to run the task (HH:MM).
        task_action: "create" to create, "update" to modify, "delete" to remove the task.
    """
    if task_action == "create":
        command = f'schtasks /create /tn "{task_name}" /tr "{python_path} {script_path}" /sc daily /st {scheduled_time} /f'
    elif task_action == "update":
        command = f'schtasks /change /tn "{task_name}" /tr "{python_path} {script_path}" /sc daily /st {scheduled_time} /f'
    elif task_action == "delete":
        command = f'schtasks /delete /tn "{task_name}" /f'  # Command to delete the task
    else:
        logging.error(f"Invalid task action: {task_action}")
        return

    try:
        subprocess.run(command, check=True, shell=True)  # Execute the schtasks command
        if task_action == "delete":
            print(f"Task '{task_name}' deleted successfully.")
        elif task_action == "create":
            print(f"Task '{task_name}' created successfully.")
        elif task_action == "update":
            print(f"Task '{task_name}' updated successfully.")

    except subprocess.CalledProcessError as e:
        print(f"Error scheduling task: {e}")
        logging.error(f"Error scheduling task: {e}")

def check_if_task_exists(task_name: str) -> bool:
    """
    Check if a scheduled task with the given name exists.

    Args:
        task_name: The name of the scheduled task to check.

    Returns:
        True if the task exists, False otherwise.
    """
    try:
        command = f'schtasks /query /tn "{task_name}"'
        subprocess.run(command, check=True, shell=True, stdout=subprocess.DEVNULL)  # Suppress output
        return True  # Task exists

    except subprocess.CalledProcessError:
        return False  # Task does not exist

def main():
    """
    Main function to execute the IMDb to Letterboxd ratings import process.
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)  # Ensure the download directory exists
    driver = setup_driver()
    downloaded_file = None  # Initialize to handle potential errors

    try:
        # Part 1: IMDb Export
        logging.info("Starting IMDb Export process...")
        if not login_to_imdb(driver):
            raise Exception("IMDb login failed")

        navigate_to_imdb_ratings(driver)

        if not initiate_imdb_export(driver):
            raise Exception("Failed to initiate IMDb export")

        downloaded_file = monitor_imdb_export_status(driver)
        if not downloaded_file:
            raise Exception("Failed to download IMDb export file")

        logging.info(f"IMDb ratings downloaded to: {downloaded_file}")

        # Part 2: Letterboxd Import
        logging.info("Starting Letterboxd Import process...")
        if not login_to_letterboxd(driver):
            raise Exception("Failed to log in to Letterboxd")

        if not import_to_letterboxd(driver, downloaded_file):
            raise Exception("Failed to import ratings to Letterboxd")

        logging.info("Successfully completed the entire import process!")

    except Exception as e:
        logging.error(f"An error occurred during the process: {e}")

    finally:
        logging.info("Closing the browser.")
        if downloaded_file and os.path.exists(downloaded_file):
            try:
                os.remove(downloaded_file)
                logging.info(f"Deleted downloaded file: {downloaded_file}")
                logging.info("Operation successful")
            except OSError as e:
                logging.error(f"Error deleting file: {e}")
        driver.quit()

if __name__ == "__main__":
    TASK_EXISTS = check_if_task_exists(TASK_NAME)  # Check if the task already exists

    if TASK_EXISTS:
        # If task already exists, ask the user to update or delete
        choice = input("A scheduled task with this name already exists. Do you want to update (u) or delete (d) it?: ").lower()
        if choice == 'u':
            schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME, task_action="update")
        elif choice == 'd':
            schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME, task_action="delete")

    else:
        schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME)  # Create if doesn't exist

    # Option to run the script directly after creating/modifying the scheduled task
    run_now = input("Do you want to run the script now? (y/n): ").lower()
    if run_now == 'y':
        main()