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

# Configuration
IMDB_EMAIL = "your_imdb_email"  # Replace with your IMDb email
IMDB_PASSWORD = "your_imdb_password" # Replace with your IMDb password
LETTERBOXD_EMAIL = "your_letterboxd_email" # Replace with your Letterboxd email
LETTERBOXD_PASSWORD = "your_letterboxd_password" # Replace with your Letterboxd password
IMDB_USER_ID = "your_imdb_user_id" # Replace with your IMDB user ID
DOWNLOAD_DIR = r"\downloads" # Path to download directory
HEADLESS_MODE = True  # Set to False for visible browser


# Scheduled Task Configuration
TASK_NAME = "IMDb_to_Letterboxd_Task"
PYTHON_EXECUTABLE = r"C:\Path\To\Your\python.exe" # Path to Python executable.
SCRIPT_PATH = r"C:\Path\To\Your\Script.py"  # Path to this script
SCHEDULED_TIME = "10:00"  # Scheduled time in 24-hour format (e.g., "14:30" for 2:30 PM)



# Helper Functions
# Logging setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Helper Functions

def random_delay(min_delay=2, max_delay=5):
    """Introduce random delays to mimic human behavior."""
    delay = random.uniform(min_delay, max_delay)
    time.sleep(delay)

def random_scroll(driver):
    """Perform random scrolls to mimic human behavior."""
    for _ in range(random.randint(1, 3)):
        driver.execute_script(f"window.scrollBy(0, {random.randint(50, 200)})")
        random_delay(0.5, 1.5)

def random_clicks(driver):
    """Perform random clicks on the page to mimic human behavior."""
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

def setup_driver():
    """Setup Selenium WebDriver (Firefox)."""

    extension_path = 'ublock_origin-1.61.2.xpi'
    options = webdriver.FirefoxOptions()
    profile = webdriver.FirefoxProfile()

    if HEADLESS_MODE:
        options.add_argument("--headless")

    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", os.path.abspath(DOWNLOAD_DIR))
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/csv")
    options.profile = profile
    driver = webdriver.Firefox(options=options)
    driver.install_addon(extension_path, temporary=True)
    return driver

def login_to_imdb(driver):
    """Log in to IMDb."""
    logging.info("Navigating to IMDb login page...")
    driver.get("https://www.imdb.com/registration/signin")
    random_delay()

    try:
        signin_imdb_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "list-group-item"))
        )
        signin_imdb_button.click()
        random_delay()

        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ap_email"))
        )
        email_input.send_keys(IMDB_EMAIL)
        random_delay()

        password_input = driver.find_element(By.ID, "ap_password")
        password_input.send_keys(IMDB_PASSWORD)
        random_delay()

        signin_submit_button = driver.find_element(By.ID, "signInSubmit")
        signin_submit_button.click()
        random_delay()
        logging.info("Logged in to IMDb successfully.")
        return True
    except TimeoutException:
        logging.error("Login to IMDb failed: Could not find login elements.")
        return False

def navigate_to_imdb_ratings(driver):
    """Navigate to the user ratings page on IMDb."""
    ratings_url = f"https://www.imdb.com/user/{IMDB_USER_ID}/ratings/"
    logging.info(f"Navigating to IMDb ratings page: {ratings_url}")
    driver.get(ratings_url)
    random_delay()

def initiate_imdb_export(driver):
    """Initiate the export of IMDb ratings."""
    logging.info("Initiating IMDb ratings export...")
    try:
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

def wait_for_download_completion(download_dir, timeout=60):
    """
    Wait for a .csv file to complete downloading in the specified directory.
    Returns the path to the downloaded file.

    Args:
        download_dir (str): Directory to monitor for downloads
        timeout (int): Maximum time to wait in seconds
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

            # Ensure file size has stabilized
            size = os.path.getsize(latest_csv)
            time.sleep(1)  # Wait a second
            if size == os.path.getsize(latest_csv):
                logging.info(f"Download completed: {latest_csv}")
                return latest_csv

        time.sleep(0.5)  # Check every half second

    raise TimeoutError("Download did not complete within the specified timeout")

def monitor_imdb_export_status(driver):
    """Monitor the export status and wait for download completion."""
    exports_url = "https://www.imdb.com/exports/"
    logging.info(f"Monitoring IMDb export status at: {exports_url}")
    driver.get(exports_url)
    random_delay()

    max_attempts = 30
    attempt = 0

    while attempt < max_attempts:
        try:
            list_content = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="list-page-mc-list-content"]'))
            )

            export_items = list_content.find_elements(By.CSS_SELECTOR, 'li[data-testid="user-ll-item"]')

            if export_items:
                first_item = export_items[0]
                try:  # Try to find the Ready button, handle exception if not found
                    ready_button = first_item.find_element(
                        By.CSS_SELECTOR,
                        'button[data-testid="export-status-button"].READY'
                    )

                    logging.info("Found Ready button. Initiating download...")
                    random_delay()
                    ready_button.click()

                    # Wait for download to complete and get the file path
                    try:
                        downloaded_file = wait_for_download_completion(DOWNLOAD_DIR)
                        if downloaded_file:
                            return downloaded_file
                    except TimeoutError as e:
                        logging.error(f"Download timeout: {e}")
                        return None

                except NoSuchElementException:
                    print("Ready button not available yet, Trying again")

            else:  # Handle empty export_items (no exports yet)
                print("Ready button not available yet, Trying again")


            attempt += 1
            logging.info(f"Export not ready yet. Attempt {attempt}/{max_attempts}")
            random_delay(10, 15)
            driver.refresh()
            random_delay()

        except Exception as e:
            logging.error(f"Error checking export status: {e}")
            attempt += 1
            random_delay(5, 10)
            driver.refresh()

    logging.error("Maximum attempts reached without finding a ready export")
    return None

def login_to_letterboxd(driver):
    """Log in to Letterboxd with enhanced error handling and debugging."""
    logging.info("Navigating to Letterboxd login page...")
    driver.get("https://letterboxd.com/settings/data/")
    random_delay()

    try:
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "field-username"))
        )
        logging.info("Found username input field.")
        username_input.send_keys(LETTERBOXD_EMAIL)
        random_delay(1, 2)

        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "field-password"))
        )
        logging.info("Found password input field.")
        password_input.send_keys(LETTERBOXD_PASSWORD)
        random_delay(1, 2)

        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
        )
        logging.info("Clicking login button.")
        login_button.click()
        random_delay(2, 3)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//h1[@class='title-hero -textleft js-hide-in-app' and text()='Account Settings']"))  # Check for profile after login
        )

        logging.info("Successfully logged into Letterboxd")

        # Navigate to the import page *after* successful login
        driver.get("https://letterboxd.com/import/")
        logging.info("Navigated to Letterboxd import page.")

        return True

    except Exception as e:
        logging.exception(f"Failed to login to Letterboxd: {e}") # More detailed error logging
        return False

def get_latest_csv_from_downloads():
    """Get the path of the most recently downloaded CSV file."""
    try:
        csv_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.csv')]
        if not csv_files:
            logging.error("No CSV files found in downloads directory")
            return None

        latest_csv = max([os.path.join(DOWNLOAD_DIR, f) for f in csv_files],
                            key=os.path.getmtime)
        logging.info(f"Found latest CSV file: {latest_csv}")
        return latest_csv
    except Exception as e:
        logging.error(f"Error finding latest CSV file: {e}")
        return None

def wait_for_element_with_text(driver, text, timeout=10):
    """Wait for an element containing specific text to appear."""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{text}')]"))
        )
    except TimeoutException:
        return None

def import_to_letterboxd(driver, csv_path):
    """Enhanced function to import ratings to Letterboxd with proper waiting and verification."""
    logging.info("Starting Letterboxd import process...")

    try:
        # Navigate to import page
        driver.get("https://letterboxd.com/import/#imdb-import")
        random_delay(2, 3)

        # Check for the intermittent "Continue" button and click it if present
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

        select_file_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '#imdb-import')]"))
        )
        select_file_button.click()
        random_delay()

        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )
        file_input.send_keys(os.path.abspath(csv_path))
        logging.info("CSV file selected for upload")

        time.sleep(2)  # Give time for the file dialog to open
        pyautogui.press('enter')  # Close the file dialog
        time.sleep(1)
        pyautogui.press('enter')

        # Wait for matching process to complete
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

        # Wait for import completion message
        saved_films_element = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Saved')]"))
        )

        # Extract number of saved films
        saved_films_text = saved_films_element.text
        num_films = ''.join(filter(str.isdigit, saved_films_text))
        logging.info(f"Successfully imported {num_films} films to Letterboxd")

        return True
    except Exception as e:
        logging.error(f"Error during Letterboxd import: {e}")
        return False

def schedule_task(task_name, python_path, script_path, scheduled_time, task_action="create"):

    if task_action == "create":
        command = f'schtasks /create /tn "{task_name}" /tr "{python_path} {script_path}" /sc daily /st {scheduled_time} /f'

    elif task_action == "update":
        command = f'schtasks /change /tn "{task_name}" /tr "{python_path} {script_path}" /sc daily /st {scheduled_time} /f'

    elif task_action == "delete":
        command = f'schtasks /delete /tn "{task_name}" /f' # Command to delete the task
        

    try:
        subprocess.run(command, check=True, shell=True)  # Run the command
        if task_action == "delete":
            print(f"Task '{task_name}' deleted successfully.")
        elif task_action == "create":
            print(f"Task '{task_name}' created successfully.")
        elif task_action == "update":
            print(f"Task '{task_name}' updated successfully.")

    except subprocess.CalledProcessError as e:
        print(f"Error scheduling task: {e}")
        logging.error(f"Error scheduling task: {e}")

def check_if_task_exists(task_name):
    """Checks if a scheduled task with the given name exists."""
    try:
        command = f'schtasks /query /tn "{task_name}"'
        subprocess.run(command, check=True, shell=True, stdout=subprocess.DEVNULL) # Suppress output
        return True # Task exists

    except subprocess.CalledProcessError:
        return False  # Task does not exist



def main():
    """Main function with enhanced download monitoring and file deletion."""
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
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
    TASK_EXISTS = check_if_task_exists(TASK_NAME) # Check if the task already exists

    if TASK_EXISTS:
        #If already exists then ask the user if update or delete the scheduled task.
        choice = input("A scheduled task with this name already exists. Do you want to update (u) or delete (d) it?: ").lower()
        if choice == 'u':
            schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME, task_action="update")
        elif choice == 'd':
            schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME, task_action="delete")

    else:
        schedule_task(TASK_NAME, PYTHON_EXECUTABLE, SCRIPT_PATH, SCHEDULED_TIME) # Create if doesn't exist

    #Option to run the script directly after creating/modifying the scheduled task
    run_now = input("Do you want to run the script now? (y/n): ").lower()
    if run_now == 'y':
        main()
