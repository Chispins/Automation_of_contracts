import time
import re
import os
import glob
import sys  # Import sys for flushing output and exiting

# Import Selenium libraries
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, WebDriverException, NoSuchElementException,
    ElementClickInterceptedException
)


def wait_for_file_download(
    download_dir, filename_pattern="*.pdf", timeout=90, stability_duration_secs=3
):
    """
    Waits for a file matching a pattern to appear in the download directory
    and checks if it's finished downloading by monitoring its size stability.

    Args:
        download_dir (str): The directory where the file is expected.
        filename_pattern (str): The glob pattern to match files (e.g., "*.pdf").
        timeout (int): Maximum seconds to wait for the download to complete.
        stability_duration_secs (int): How long the file size must be stable
                                       to consider the download finished.

    Returns:
        str or None: The path of the found and stable file, or None if timeout
                     occurs.
    """
    seconds_elapsed = 0
    last_size = -1  # To detect if the file is still growing
    last_size_change_time = time.time()  # To measure time since last size change

    # Get initial files to ignore those already present before download attempt
    initial_files = set(glob.glob(os.path.join(download_dir, filename_pattern)))
    # Include common temporary extensions in initial check
    initial_files.update(glob.glob(os.path.join(download_dir, "*.crdownload")))
    initial_files.update(glob.glob(os.path.join(download_dir, "*.part")))
    initial_files.update(glob.glob(os.path.join(download_dir, "*.tmp")))

    print(
        f"\nWaiting for file matching '{filename_pattern}' in {download_dir}"
        f" (timeout: {timeout}s)..."
    )

    while seconds_elapsed < timeout:
        # List current files matching the pattern or temp extensions
        current_files = set(glob.glob(os.path.join(download_dir, filename_pattern)))
        current_files.update(glob.glob(os.path.join(download_dir, "*.crdownload")))
        current_files.update(glob.glob(os.path.join(download_dir, "*.part")))
        current_files.update(glob.glob(os.path.join(download_dir, "*.tmp")))

        # Find new files that weren't there initially
        new_files = list(current_files - initial_files)

        # Filter out *only* explicit temporary files
        potential_final_files = [
            f for f in new_files if not f.endswith(('.crdownload', '.part', '.tmp'))
        ]
        temp_files = [
            f for f in new_files if f.endswith(('.crdownload', '.part', '.tmp'))
        ]

        # Prioritize a potential final file if found
        file_to_check = None
        if potential_final_files:
            # Find the most recently modified among potential final files
            file_to_check = max(potential_final_files, key=os.path.getmtime)
        elif temp_files:
            # Find the most recently modified temp file if no potential file yet
            file_to_check = max(temp_files, key=os.path.getmtime)

        if file_to_check and os.path.exists(file_to_check):
            # Double check existence as files can disappear/be renamed by browser
            try:
                current_size = os.path.getsize(file_to_check)
            except OSError:
                 # File might have just been renamed/moved by the browser,
                 # continue waiting or check again.
                 # Let the loop continue and check for new/renamed file.
                 print(f"  Warning: File vanished during size check: {os.path.basename(file_to_check)}", end='\r', flush=True)
                 file_to_check = None # Reset to force re-check of directory
                 time.sleep(1) # Give browser a moment
                 seconds_elapsed += 1
                 continue # Skip rest of this iteration

            # Check if size is stable for stability_duration_secs
            if current_size > 0 and \
               (time.time() - last_size_change_time) >= stability_duration_secs and \
               current_size == last_size:
                print(
                    "\nDownload complete: "
                    f"{os.path.basename(file_to_check)} ({current_size} bytes)."
                )
                # Give file system a moment for final writes
                time.sleep(0.5)
                return file_to_check
            elif current_size > last_size:
                # File is growing, update size and time
                last_size = current_size
                last_size_change_time = time.time()
                print(
                    f"Downloading: {os.path.basename(file_to_check)} | "
                    f"Size: {current_size} bytes | Time: {seconds_elapsed}s",
                    end='\r', flush=True
                )
            elif current_size < last_size:
                 # Size decreased (rare, maybe resume or error), reset stability timer
                 last_size = current_size
                 last_size_change_time = time.time()
                 print(
                     f"Downloading (size decreased): {os.path.basename(file_to_check)} | "
                     f"Size: {current_size} bytes | Time: {seconds_elapsed}s",
                     end='\r', flush=True
                 )
            else:  # current_size == last_size but not stable yet, or current_size is 0
                # Size is same but not stable for long enough, or file is empty/stalled
                print(
                    f"Downloading: {os.path.basename(file_to_check)} | "
                    f"Size: {current_size} bytes | Time: {seconds_elapsed}s",
                    end='\r', flush=True
                )

        else:
            # No new potential or temp files found yet matching patterns
            print(
                f"Waiting for download to start... Time: {seconds_elapsed}s",
                end='\r', flush=True
            )

        time.sleep(1)
        seconds_elapsed += 1

    print("\nTimeout waiting for file download.")
    return None


def download_pdf_selenium(
    rut, driver, download_dir, expected_filename_suffix=".pdf",
    num_click_attempts=3, click_delay_secs=2
):
    """
    Navigates to the provider page and attempts to click the PDF download link
    via JavaScript for robustness. Waits for and renames the downloaded file.

    Args:
        rut (str): The RUT (identifier number) of the provider.
        driver: The Selenium WebDriver instance.
        download_dir (str): The directory where the PDF should be downloaded.
        expected_filename_suffix (str): The expected suffix of the downloaded file
                                        before renaming (e.g., ".pdf").
        num_click_attempts (int): How many times to attempt clicking the link.
        click_delay_secs (int): Delay in seconds between JS click attempts.

    Returns:
        str or None: The path to the downloaded and renamed PDF file, or None if failed.
    """
    base_url = "https://proveedor.mercadopublico.cl/ficha/comportamiento/"
    url = f"{base_url}{rut}"
    print(f"\n--- Processing RUT: {rut} ---")
    print(f"Navigating to {url}")

    # Define the expected final filename format
    expected_final_filename = f"ficha_{rut}.pdf"
    expected_final_filepath = os.path.join(download_dir, expected_final_filename)

    # Check if the file already exists
    if os.path.exists(expected_final_filepath):
        print(
            "File already exists: "
            f"{os.path.basename(expected_final_filepath)}. Skipping download."
        )
        return expected_final_filepath

    # Clean up any potentially incomplete previous downloads for this RUT.
    # This targets files matching the final name pattern (*.pdf, .crdownload etc.)
    # and also generic temp files that might not have the RUT in the name yet.
    cleanup_patterns = [
        os.path.join(download_dir, f"ficha_{rut}.*"),
        os.path.join(download_dir, "*.crdownload"),
        os.path.join(download_dir, "*.part"),
        os.path.join(download_dir, "*.tmp"),
    ]
    files_to_clean = set()
    for pattern in cleanup_patterns:
        files_to_clean.update(glob.glob(pattern))

    for file_path in files_to_clean:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"Cleaned up: {os.path.basename(file_path)}")
            except OSError as e:
                print(f"Warning: Could not remove old file {os.path.basename(file_path)}: {e}")

    try:
        # Navigate to the URL
        driver.get(url)

        # Wait for the PDF download link to be present/visible
        pdf_link_xpath = "//a[text()='Descargar ficha en formato PDF']"
        # Wait up to 15 seconds for the link to appear
        wait = WebDriverWait(driver, 15)

        try:
            # Wait for the element to be visible
            pdf_link_element = wait.until(
                EC.visibility_of_element_located((By.XPATH, pdf_link_xpath))
            )
            print("PDF download link found.")

        except (TimeoutException, NoSuchElementException):
            print(
                f"Timeout or element not found: PDF download link not found"
                f" or not visible for RUT {rut} within 15 seconds."
            )
            # Check for common "not found" messages if link wasn't found
            try:
                page_text = driver.find_element(By.TAG_NAME("body")).text.lower()
                if re.search(
                    r'no existe|no encontrada|proveedor inválido|página no encontrada',
                    page_text
                ):
                    print(
                        f"Info: Page content suggests provider {rut} might not"
                        f" exist or page is unavailable."
                    )
            except NoSuchElementException:
                 print("Info: Could not read page body to check for 'not found' message.")
            return None


        # --- Attempt to click the download link multiple times using JavaScript ---
        # JS click is often more reliable than .click() when dealing with
        # potential overlays or complex page interactions.
        print(
            "Attempting to initiate download "
            f"({num_click_attempts} JS clicks with {click_delay_secs}s delay)..."
        )
        successful_click = False
        for i in range(num_click_attempts):
            try:
                print(f"  Click attempt {i+1}/{num_click_attempts}...",
                      end='\r', flush=True)
                driver.execute_script("arguments[0].click();", pdf_link_element)
                successful_click = True  # Assume click initiated something
                # Delay briefly after the click attempt
                time.sleep(click_delay_secs)
            except Exception as js_error:
                print(f"\n  Warning: Error during JS click attempt {i+1}: {js_error}",
                      flush=True)
                time.sleep(1)  # Shorter delay on error before retrying

        if successful_click:
            print("\nInitiated download attempts.")
        else:
            print("\nWarning: Could not successfully execute JS click on the download link.")
            # If clicks failed, the download won't start. Exit here.
            return None

        # --- Wait for the file to appear and finish downloading ---
        # Wait for any file ending in the expected suffix to show up and finish
        download_timeout_secs = 90  # Allow up to 90 seconds for the actual download
        downloaded_filepath = wait_for_file_download(
            download_dir,
            filename_pattern=f"*{expected_filename_suffix}",
            timeout=download_timeout_secs,
            stability_duration_secs=3
        )

        if downloaded_filepath:
            print(f"Download detected: {os.path.basename(downloaded_filepath)}")
            # --- Rename the downloaded file to the expected format ---
            try:
                # Ensure the downloaded file is not already the target name (unlikely but safe)
                if os.path.abspath(downloaded_filepath) != os.path.abspath(expected_final_filepath):
                    # Before renaming, check if the target path exists (should be cleaned but double check)
                    if os.path.exists(expected_final_filepath):
                        try:
                            os.remove(expected_final_filepath)
                            print(
                                "Removed existing file at target path: "
                                f"{os.path.basename(expected_final_filepath)}"
                            )
                        except OSError as e:
                            print(
                                "Warning: Could not remove existing target file "
                                f"{os.path.basename(expected_final_filepath)}: {e}. "
                                "Renaming might fail."
                            )

                    os.rename(downloaded_filepath, expected_final_filepath)
                    print(f"File renamed to: {os.path.basename(expected_final_filepath)}")
                else:
                    print(f"File already has the expected name: {os.path.basename(expected_final_filepath)}")

                return expected_final_filepath  # Return the path to the final file

            except OSError as e:
                print(
                    "Error renaming file "
                    f"{os.path.basename(downloaded_filepath)} to "
                    f"{os.path.basename(expected_final_filepath)}: {e}"
                )
                return None  # Renaming failed

        else:
            print(
                "Error: No file matching "
                f"'*{expected_filename_suffix}' detected in download directory"
                f" for RUT {rut} within timeout after click attempts."
            )
            return None

    except WebDriverException as e:
        print(f"WebDriver error processing RUT {rut}: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while processing RUT {rut}: {e}")
        return None
    finally:
        # Add a small pause regardless of outcome before processing the next RUT
        time.sleep(2)


# --- Main execution ---

if __name__ == "__main__":
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

    # >>>>>>>>>> COMMENT THE LINE BELOW TO RUN IN NON-HEADLESS MODE FOR DEBUGGING <<<<<<<<<<
    # Headless mode runs the browser without a visible window
    # chrome_options.add_argument("--headless=new")
    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    # Standard arguments for headless/automation stability
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")  # Recommended window size
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        " AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.24 Safari/537.36"
    )
    chrome_options.add_argument("--disable-gpu")  # Often recommended for headless

    # --- Configure download preferences ---
    # These settings tell Chrome to download files automatically
    prefs = {
        "download.default_directory": DOWNLOAD_DIRECTORY,
        "download.prompt_for_download": False,  # Never ask where to save
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # Crucial for downloading PDFs
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # --- Initialize WebDriver ---
    driver = None  # Initialize driver variable outside try block
    try:
        # Using the recommended method (assuming chromedriver is in PATH)
        # If chromedriver is NOT in your PATH, provide the full path via Service
        # Example: service = Service('/path/to/chromedriver')
        # driver = webdriver.Chrome(service=service, options=chrome_options)
        driver = webdriver.Chrome(options=chrome_options)
        print("WebDriver initialized successfully.")
    except WebDriverException as e:
        print(f"Error initializing WebDriver: {e}")
        print(
            "Please ensure you have Chrome installed and the correct chromedriver "
            "executable is in your system's PATH or specified via the Service object."
        )
        print("Exiting.")
        sys.exit(1)  # Use sys.exit to stop the script

    # Your list of RUT numbers
    rut_numbers = [
        "8.046.185-2",  # Example RUT from the image
        "76.012.419-6",
        "78.612.480-7",
        # Add all the RUT numbers you want to process here
        # "12.345.678-9", # Example dummy RUT (replace with real ones)
    ]

    # Dictionary to store the results of downloads
    download_results = {}

    print("\n--- Starting PDF Download Process ---")

    # Iterate through the list and download PDF for each RUT
    for i, rut in enumerate(rut_numbers):
        print(f"\nProcessing RUT {i+1}/{len(rut_numbers)}: {rut}")
        try:
            # --- Download the PDF ---
            downloaded_file_path = download_pdf_selenium(
                rut,
                driver,
                DOWNLOAD_DIRECTORY,
                num_click_attempts=10,  # Try clicking 3 times
                click_delay_secs=2     # Wait 2 seconds between click attempts
            )
            download_results[rut] = downloaded_file_path

        except Exception as e:
            print(f"\nAn error occurred processing RUT {rut} in the main loop: {e}")
            download_results[rut] = None  # Mark as failed

    # --- Close the WebDriver ---
    if driver:
        driver.quit()
        print("\nWebDriver closed.")

    # --- Output the results of downloads ---
    print("\n--- PDF Download Summary ---")
    success_count = 0
    fail_count = 0
    for rut, file_path in download_results.items():
        if file_path and os.path.exists(file_path):
            print(f" ✅ {rut}: Downloaded -> {os.path.basename(file_path)}")
            success_count += 1
        else:
            print(f" ❌ {rut}: Failed to download.")
            fail_count += 1

    print(
        f"\nProcess finished. Successfully downloaded {success_count} PDFs, "
        f"failed {fail_count}."
    )