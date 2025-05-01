import requests
from bs4 import BeautifulSoup
import time
import re
import os
import glob

# Import Selenium libraries
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, ElementClickInterceptedException

def wait_for_file_download(download_dir, filename_pattern, timeout=120): # Increased timeout again
    """
    Waits for a file matching a pattern to appear in the download directory
    and checks if it's finished downloading.
    Returns the path of the found file or None if timeout occurs.
    """
    seconds = 0
    last_size = -1 # To detect if the file is still growing
    last_check_time = time.time() # To measure time since last size change
    initial_files = set(glob.glob(os.path.join(download_dir, filename_pattern))) # Get files before download attempt

    print(f"\nWaiting for file matching '{filename_pattern}' in {download_dir}...")

    while seconds < timeout:
        # List current files in the download directory matching the pattern
        current_files = set(glob.glob(os.path.join(download_dir, filename_pattern)))

        # Find new files that weren't there initially
        new_files = list(current_files - initial_files)

        # Filter out potential temporary files (like .crdownload for Chrome)
        valid_new_files = [f for f in new_files if not f.endswith(('.crdownload', '.part', '.tmp'))]

        if valid_new_files:
            # Find the most recently modified file among the valid new ones
            latest_file = max(valid_new_files, key=os.path.getmtime)

            # Check if the file size is stable (download finished)
            current_size = os.path.getsize(latest_file)
            # Check if size is greater than 0 and hasn't changed in the last 2 seconds
            if current_size > 0 and (time.time() - last_check_time) >= 2 and current_size == last_size:
                 print(f"\nFile download appears complete: {os.path.basename(latest_file)} ({current_size} bytes)")
                 # Give it a tiny bit more time just in case the file system is slow
                 time.sleep(1)
                 return latest_file
            elif current_size > last_size:
                 # File is growing
                 last_size = current_size
                 last_check_time = time.time() # Reset timer when file grows
                 print(f"Downloading... {os.path.basename(latest_file)} size: {current_size} bytes ({seconds}/{timeout}s)", end='\r')
            else: # current_size <= last_size but not stable yet or current_size == 0
                 # File size is same or decreased (shouldn't happen), or file is empty
                 last_size = current_size # Update size
                 print(f"Waiting for download to start/resume/stabilize... {os.path.basename(latest_file) if current_size > 0 else 'file not started'} size: {current_size} bytes ({seconds}/{timeout}s)", end='\r')


        else:
             # No new valid files found yet
             print(f"Waiting for file to appear... ({seconds}/{timeout}s)", end='\r')


        time.sleep(1)
        seconds += 1

    print("\nTimeout waiting for file download.")
    return None


def download_pdf_selenium(rut, driver, download_dir, num_click_attempts=5, click_delay=3):
    """
    Navigates to the provider page and attempts to click the PDF download link multiple times via JavaScript.

    Args:
        rut (str): The RUT (identifier number) of the provider.
        driver: The Selenium WebDriver instance.
        download_dir (str): The directory where the PDF should be downloaded.
        num_click_attempts (int): How many times to attempt clicking the link.
        click_delay (int): Delay in seconds between click attempts.


    Returns:
        str or None: The path to the downloaded and renamed PDF file, or None if failed.
    """
    base_url = "https://proveedor.mercadopublico.cl/ficha/comportamiento/"
    url = f"{base_url}{rut}"
    print(f"\nAttempting to download PDF for RUT: {rut} from {url}")

    # Define the expected final filename format
    expected_final_filename = f"ficha_{rut}.pdf"
    expected_final_filepath = os.path.join(download_dir, expected_final_filename)

    # Check if the file already exists
    if os.path.exists(expected_final_filepath):
        print(f"PDF for RUT {rut} already exists: {expected_final_filepath}. Skipping download.")
        return expected_final_filepath

    # Clean up any potentially incomplete previous downloads for this RUT
    # Clean up the target name format and any temp files that might match the RUT
    old_files_pattern = os.path.join(download_dir, f"ficha_{rut}.*") # Matches ficha_{rut}.pdf, .crdownload etc
    old_files = glob.glob(old_files_pattern)
    # Also clean up generic temp files that might exist if the site names files generically initially
    generic_temp_pattern = os.path.join(download_dir, "*.crdownload")
    old_files.extend(glob.glob(generic_temp_pattern))
    generic_temp_pattern = os.path.join(download_dir, "*.part")
    old_files.extend(glob.glob(generic_temp_pattern))


    for f in set(old_files): # Use set to avoid processing duplicates if patterns overlap
        try:
            os.remove(f)
            print(f"Removed potentially incomplete previous download or temp file: {f}")
        except OSError as e:
            print(f"Error removing old file {f}: {e}")


    try:
        # Navigate to the URL
        driver.get(url)

        # --- Wait for the PDF download link to be present/visible ---
        # We'll wait for presence first, then try clicking with JS
        pdf_link_xpath = "//a[text()='Descargar ficha en formato PDF']"
        wait = WebDriverWait(driver, 20) # Wait up to 20 seconds

        try:
             # Wait for the element to be visible
             pdf_link_element = wait.until(
                 EC.visibility_of_element_located((By.XPATH, pdf_link_xpath))
             )
             print("PDF download link found and is visible.")

        except TimeoutException:
            print(f"Timed out waiting for the PDF download link for RUT {rut}. Page might not have loaded correctly or provider not found.")
            # Check for "no existe" message as fallback
            soup_after_timeout = BeautifulSoup(driver.page_source, 'html.parser')
            if soup_after_timeout.find(string=re.compile(r'no existe|no encontrada|proveedor inválido', re.IGNORECASE)):
                 print(f"Provider with RUT {rut} might not exist or have a public ficha.")
            return None
        except NoSuchElementException:
            print(f"PDF download link element not found using XPath '{pdf_link_xpath}' for RUT {rut}.")
            # Check for provider not found message as fallback
            soup_after_select = BeautifulSoup(driver.page_source, 'html.parser')
            if soup_after_select.find(string=re.compile(r'no existe|no encontrada|proveedor inválido', re.IGNORECASE)):
                 print(f"Provider with RUT {rut} might not exist or have a public ficha.")
            return None


        # --- Attempt to click the download link multiple times using JavaScript ---
        print(f"Attempting to click the download link {num_click_attempts} times with {click_delay}s delay using JavaScript...")
        for i in range(num_click_attempts):
            print(f"Click attempt {i+1}/{num_click_attempts}...")
            try:
                driver.execute_script("arguments[0].click();", pdf_link_element)
                # Add a pause after each click attempt
                time.sleep(click_delay)
            except Exception as js_error:
                print(f"Error during JavaScript click attempt {i+1}: {js_error}")
                # Optionally break or continue based on error type

        print("Finished click attempts.")


        # --- Wait for the file to appear in the download directory ---
        # We'll wait for *any* PDF file (*.pdf) to show up and finish downloading
        downloaded_filepath = wait_for_file_download(download_dir, "*.pdf", timeout=120) # Increased timeout

        if downloaded_filepath:
            print(f"\nFile detected and finished: {downloaded_filepath}")
            # --- Rename the downloaded file to the expected format ---
            try:
                new_filepath = os.path.join(download_dir, expected_final_filename)
                # If the downloaded file is not already the target file name
                if os.path.abspath(downloaded_filepath) != os.path.abspath(new_filepath): # Use abspath for robust comparison
                    # Ensure the target path doesn't exist from a very old run that wasn't cleaned
                    if os.path.exists(new_filepath):
                        try:
                            os.remove(new_filepath)
                            print(f"Removed old file at target path: {new_filepath}")
                        except OSError as e:
                             print(f"Warning: Could not remove old file at target path {new_filepath}: {e}. Renaming might fail.")

                    os.rename(downloaded_filepath, new_filepath)
                    print(f"File renamed to: {new_filepath}")
                else:
                    print(f"File already has the expected name: {expected_final_filename}")

                return new_filepath # Return the path to the renamed file

            except OSError as e:
                print(f"Error renaming file {downloaded_filepath} to {new_filepath}: {e}")
                # Return None if renaming failed, as the file is not in the expected place/name
                return None # Or return downloaded_filepath if you want the original name

        else:
            print(f"\nNo PDF file detected in download directory matching '*.pdf' for RUT {rut} within timeout after click attempts.")
            # You might want to check browser logs for download errors if needed (advanced)
            # Example: driver.get_log('browser') after enabling logging preferences
            return None


    except WebDriverException as e:
        print(f"WebDriver error during PDF download for RUT {rut}: {e}")
        return None # Handle WebDriver specific errors
    except Exception as e:
        print(f"An unexpected error occurred during PDF download for RUT {rut}: {e}")
        return None


# --- Main execution ---

# Define your desired download directory
DOWNLOAD_DIRECTORY = os.path.join(os.getcwd(), "mercadopublico_pdfs")
# Create the directory if it doesn't exist
if not os.path.exists(DOWNLOAD_DIRECTORY):
    os.makedirs(DOWNLOAD_DIRECTORY)
    print(f"Created download directory: {DOWNLOAD_DIRECTORY}")
else:
    print(f"Using download directory: {DOWNLOAD_DIRECTORY}")


# Configure Selenium WebDriver (Chrome)
chrome_options = webdriver.ChromeOptions()
# >>>>>>>>>> COMMENT THIS LINE TO RUN IN NON-HEADLESS MODE FOR DEBUGGING <<<<<<<<<<
# Keep headless for server/background running after debugging
# chrome_options.add_argument("--headless=new")
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.24 Safari/537.36")

# --- Configure download preferences ---
prefs = {
    "download.default_directory": DOWNLOAD_DIRECTORY,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# --- Initialize WebDriver ---
try:
    # Using the recommended method (assuming chromedriver is in PATH)
    # If chromedriver is NOT in your PATH, provide the full path via Service as shown before
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver initialized successfully.")
except WebDriverException as e:
    print(f"Error initializing WebDriver: {e}")
    print("Please ensure you have Chrome installed and the correct chromedriver executable")
    print("is in your system's PATH or specified via the Service object.")
    print("Exiting.")
    exit()


# Your list of RUT numbers
rut_numbers = [
    "78.615.850-8", # Example RUT from the image
    # Add all the RUT numbers you want to process here
    # "12.345.678-9", # Example dummy RUT (replace with real ones)
    # "98.765.432-1", # Example dummy RUT (replace with real ones)
    # "..." Add more RUTs from your list
]

# Dictionary to store the results of downloads (optional, for tracking)
download_results = {}

# Iterate through the list and download PDF for each RUT
for rut in rut_numbers:
    # --- Download the PDF ---
    downloaded_file_path = download_pdf_selenium(rut, driver, DOWNLOAD_DIRECTORY, num_click_attempts=5, click_delay=3)
    download_results[rut] = downloaded_file_path

    # Add a small delay between processing each RUT
    time.sleep(3) # Sleep for 3 seconds between processing RUTs


# --- Close the WebDriver ---
driver.quit()
print("\nWebDriver closed.")


# --- Output the results of downloads ---
print("\n--- PDF Download Results ---")
for rut, file_path in download_results.items():
    if file_path and os.path.exists(file_path):
        print(f"RUT: {rut}, PDF Downloaded: {os.path.basename(file_path)}")
    elif file_path and not os.path.exists(file_path):
         print(f"RUT: {rut}, PDF Download Failed (File missing after reported download path - check manual download).")
    else:
        print(f"RUT: {rut}, PDF Download Failed (Could not initiate or find file - check browser window).")