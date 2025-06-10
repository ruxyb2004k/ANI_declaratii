import os
import random
import time
import logging
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from bs4 import BeautifulSoup
import pandas as pd
import requests
from dotenv import load_dotenv
import urllib.parse
import sys

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

class DeclaratiiScraper:
    def __init__(self):
        self.base_url = "https://declaratii.integritate.eu/"
        self.setup_driver()
        self.all_data = []  # List to store all table data
        
    def setup_driver(self):
        """Set up the undetected Chrome WebDriver with appropriate options"""
        try:
            options = uc.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-software-rasterizer')
            
            # Set download preferences
            prefs = {
                "download.default_directory": os.path.abspath("downloads"),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)
            
            # Create undetected-chromedriver instance
            self.driver = uc.Chrome(options=options)
            self.driver.maximize_window()
            
        except Exception as e:
            logger.error(f"Error setting up Chrome driver: {str(e)}")
            logger.error("Please make sure Chrome is installed and up to date.")
            raise

    def get_names_from_excel(self, excel_file):
        """Read names from all sheets in the Excel file"""
        try:
            # Read all sheets
            excel = pd.ExcelFile(excel_file)
            all_names = []
            
            for sheet_name in excel.sheet_names:
                logger.info(f"Reading sheet: {sheet_name}")
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                # Check if 'Nume' column exists
                if 'Nume' not in df.columns:
                    logger.warning(f"No 'Nume' column found in sheet {sheet_name}")
                    continue
                
                # Get names, remove dashes and clean up
                names = df['Nume'].dropna().tolist()
                cleaned_names = [name.replace('-', ' ').strip() for name in names]
                all_names.extend(cleaned_names)
            
            logger.info(f"Found {len(all_names)} names across all sheets")
            return all_names
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return []

    def process_name(self, name):
        """Process a single name and download its declarations"""
        try:
            logger.info(f"\nProcessing name: {name}")
            results = self.search_person(name)
            
            if results is not None and not results.empty:
                logger.info(f"Found {len(results)} declarations for {name}")
                
                # Process all pages
                while True:
                    # Get all download buttons on current page
                    download_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button.mdc-button")
                    
                    # Process current page
                    for idx, row in results.iterrows():
                        if row['has_download'] and idx < len(download_buttons):
                            # Create filename
                            filename = f"{row['name'].replace(' ', '_')}_{row['date'].replace('.', '-')}_{row['declaration_type'].replace(' ', '_')}.pdf"
                            filename = urllib.parse.unquote(filename)  # Handle special characters
                            
                            # Add filename to row data
                            row_dict = row.to_dict()
                            row_dict['saved_filename'] = filename
                            
                            # Download the file
                            success, final_filename = self.download_file_from_button(download_buttons[idx], filename)
                            if success:
                                logger.info(f"Downloaded to {final_filename}")
                                row_dict['download_status'] = 'Success'
                                row_dict['saved_filename'] = final_filename  # Update with final filename
                            else:
                                logger.error(f"Failed to download {filename}")
                                row_dict['download_status'] = 'Failed'
                            
                            # Add to all_data
                            self.all_data.append(row_dict)
                            
                            # Wait between downloads
                            self.random_delay()
                        else:
                            logger.warning(f"No download button for {row['name']} on {row['date']}")
                            # Add to all_data even if no download button
                            row_dict = row.to_dict()
                            row_dict['saved_filename'] = 'N/A'
                            row_dict['download_status'] = 'No download button'
                            self.all_data.append(row_dict)
                    
                    # Check for next page button
                    try:
                        next_page_button = self.driver.find_element(By.CSS_SELECTOR, "button.mat-mdc-paginator-navigation-next")
                        if not next_page_button.is_enabled():
                            logger.info("No more pages to process")
                            break
                            
                        # Click next page
                        next_page_button.click()
                        logger.info("Moving to next page")
                        
                        # Wait for the new page to load
                        time.sleep(random.uniform(3, 5))
                        
                        # Get new results
                        results = self.extract_table_data()
                        if results is None or results.empty:
                            logger.info("No more results found")
                            break
                            
                    except NoSuchElementException:
                        logger.info("No pagination found - this is the only page")
                        break
                        
            else:
                logger.warning(f"No declarations found for {name}")
                
        except Exception as e:
            logger.error(f"Error processing name {name}: {str(e)}")
            # Take a screenshot for debugging
            try:
                self.driver.save_screenshot(f"error_{name.replace(' ', '_')}.png")
                logger.info(f"Error screenshot saved as 'error_{name.replace(' ', '_')}.png'")
            except:
                pass

    def random_delay(self):
        delay = random.uniform(4, 10)  # Increased delay between actions
        time.sleep(delay)

    def wait_for_element(self, by, value, timeout=20):
        """Wait for an element to be present and visible"""
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {value}")
            return None

    def wait_for_download(self, timeout=30):  # Reduced timeout
        """Wait for the download to complete and return the downloaded filename"""
        start_time = time.time()
        downloads_dir = os.path.abspath("downloads")
        initial_files = set(f for f in os.listdir(downloads_dir) if f.endswith('.pdf'))
        
        while time.time() - start_time < timeout:
            current_files = set(f for f in os.listdir(downloads_dir) if f.endswith('.pdf'))
            new_files = current_files - initial_files
            if new_files:
                # Return the path of the new file
                return os.path.join(downloads_dir, new_files.pop())
            time.sleep(0.5)  # Check more frequently
        return None

    def download_file_from_button(self, button, filename):
        """Download file by clicking the download button"""
        try:
            # Scroll the button into view
            self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
            time.sleep(2)  # Wait for scroll to complete
            
            # Click the download button
            button.click()
            logger.info(f"Clicked download button for {filename}")
            
            # Wait for the download to complete and get the downloaded file path
            downloaded_file = self.wait_for_download()
            if downloaded_file:
                # Handle duplicate filenames
                base_name = os.path.splitext(filename)[0]
                extension = os.path.splitext(filename)[1]
                counter = 1
                new_path = os.path.join(os.path.dirname(downloaded_file), filename)
                
                # If file exists, add a number to the filename
                while os.path.exists(new_path):
                    new_path = os.path.join(os.path.dirname(downloaded_file), f"{base_name}_{counter}{extension}")
                    counter += 1
                
                try:
                    os.rename(downloaded_file, new_path)
                    final_filename = os.path.basename(new_path)
                    logger.info(f"Successfully downloaded and renamed to {final_filename}")
                    return True, final_filename
                except Exception as e:
                    logger.error(f"Error renaming file: {str(e)}")
                    return False, None
            else:
                logger.error(f"Download timeout for {filename}")
                return False, None
                
        except Exception as e:
            logger.error(f"Error downloading file: {str(e)}")
            return False, None

    def wait_for_cloudflare(self, timeout=30):
        """Wait for Cloudflare verification to complete"""
        start_time = time.time()
        logger.info("Waiting for Cloudflare verification...")
        
        while time.time() - start_time < timeout:
            try:
                # Check if we're on the Cloudflare verification page
                if "challenge" in self.driver.current_url or "cloudflare" in self.driver.current_url.lower():
                    logger.info("Cloudflare verification detected. Please complete the verification manually.")
                    # Wait for the verification to complete
                    while time.time() - start_time < timeout:
                        if "challenge" not in self.driver.current_url and "cloudflare" not in self.driver.current_url.lower():
                            logger.info("Cloudflare verification completed!")
                            return True
                        time.sleep(1)
                else:
                    # If we're not on the verification page, we're good
                    return True
            except Exception as e:
                logger.error(f"Error checking Cloudflare status: {str(e)}")
                time.sleep(1)
        
        logger.error("Timeout waiting for Cloudflare verification")
        return False

    def search_person(self, name):
        """Search for a person by name"""
        try:
            logger.info(f"Navigating to {self.base_url}")
            self.driver.get(self.base_url)
            self.random_delay()
            
            # Wait for the page to load completely
            self.wait_for_element(By.TAG_NAME, "body")
            
            # Log the current URL and page title for debugging
            logger.info(f"Current URL: {self.driver.current_url}")
            logger.info(f"Page title: {self.driver.title}")
            
            # Try different possible selectors for the search input
            search_input = None
            possible_selectors = [
                (By.ID, "ssidLastName"),  # Primary selector - exact ID match
                (By.CSS_SELECTOR, "input.form-control[type='text']"),  # Class and type match
                (By.CSS_SELECTOR, "input[style*='width: 600px']"),  # Style attribute match
                (By.CSS_SELECTOR, "input[type='text'][maxlength='60']"),  # Type and maxlength match
                (By.CSS_SELECTOR, "input[type='text']"),  # Generic text input fallback
            ]
            
            for by, value in possible_selectors:
                try:
                    search_input = self.wait_for_element(by, value)
                    if search_input:
                        logger.info(f"Found search input with selector: {value}")
                        break
                except:
                    continue
            
            if not search_input:
                logger.error("Could not find search input field")
                return None
            
            # Enter the name with human-like typing
            logger.info(f"Entering search term: {name}")
            search_input.clear()
            for char in name:
                search_input.send_keys(char)
                time.sleep(random.uniform(0.1, 0.3))  # Random delay between keystrokes
            self.random_delay()
            
            # Try different possible selectors for the submit button
            submit_button = None
            button_selectors = [
                (By.CSS_SELECTOR, "button.btn.btn-success"),
                (By.CSS_SELECTOR, "button[class*='btn-success']"),
                (By.XPATH, "//button[contains(text(), 'Cautare')]"),
                (By.XPATH, "//button[contains(@class, 'btn-success')]"),
                (By.CSS_SELECTOR, "button[type='button']"),
                (By.CSS_SELECTOR, "button.btn")
            ]
            
            for by, value in button_selectors:
                try:
                    submit_button = self.wait_for_element(by, value)
                    if submit_button:
                        logger.info(f"Found submit button with selector: {value}")
                        break
                except:
                    continue
            
            if not submit_button:
                logger.error("Could not find submit button")
                return None
            
            # Submit the form
            logger.info("Clicking submit button")
            submit_button.click()
            
            # Wait for Cloudflare verification if needed
            if not self.wait_for_cloudflare():
                logger.error("Failed to pass Cloudflare verification")
                return None
            
            # Additional wait after verification
            time.sleep(5)
            
            return self.extract_table_data()
            
        except Exception as e:
            logger.error(f"Error during search: {str(e)}")
            # Take a screenshot for debugging
            try:
                self.driver.save_screenshot("error_screenshot.png")
                logger.info("Error screenshot saved as 'error_screenshot.png'")
            except:
                pass
            return None
            
    def extract_table_data(self):
        """Extract data from the results table"""
        try:
            # Wait for the table to be present
            table = self.wait_for_element(By.CSS_SELECTOR, "table.mat-mdc-table")
            if not table:
                logger.error("Could not find results table")
                return None
                
            self.random_delay()
            
            # Parse the table using BeautifulSoup
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            table_data = []
            
            # Find all rows (mat-row elements)
            rows = soup.find_all('mat-row')
            if not rows:
                logger.warning("No rows found in the table")
                return None
                
            for row in rows:
                cells = row.find_all('mat-cell')
                if cells:
                    # Extract data from each cell
                    row_data = {
                        'name': cells[0].text.strip(),
                        'institution': cells[1].text.strip(),
                        'position': cells[2].text.strip(),
                        'city': cells[3].text.strip(),
                        'county': cells[4].text.strip(),
                        'date': cells[5].text.strip(),
                        'declaration_type': cells[6].text.strip()
                    }
                    
                    # Find download button
                    download_button = cells[7].find('button')
                    if download_button:
                        row_data['has_download'] = True
                    else:
                        row_data['has_download'] = False
                    
                    table_data.append(row_data)
                    
            if not table_data:
                logger.warning("No data extracted from table")
                return None
                
            return pd.DataFrame(table_data)
            
        except Exception as e:
            logger.error(f"Error extracting table data: {str(e)}")
            return None
            
    def close(self):
        """Close the WebDriver"""
        if hasattr(self, 'driver'):
            self.driver.quit()

def main():
    # Check if Excel file is provided as argument
    if len(sys.argv) != 2:
        logger.error("Please provide the Excel file name as an argument.")
        logger.error("Usage: python scraper.py <excel_file_name>")
        return

    excel_file = sys.argv[1]
    if not os.path.exists(excel_file):
        logger.error(f"Excel file '{excel_file}' not found.")
        return

    scraper = DeclaratiiScraper()
    try:
        # Create downloads directory if it doesn't exist
        os.makedirs('downloads', exist_ok=True)
        
        # Get names from Excel file
        # excel_file = "Baza de date - Cautare ANI_short.xlsx"
        names = scraper.get_names_from_excel(excel_file)
        
        if not names:
            logger.error("No names found in Excel file")
            return
            
        # Process each name
        for name in names:
            scraper.process_name(name)
            # Add a longer delay between different people
            time.sleep(random.uniform(10, 15))
        
        # Save all collected data to Excel
        if scraper.all_data:
            df = pd.DataFrame(scraper.all_data)
            base_output_file = "all_declarations_data.xlsx"
            
            # Check if file exists and add timestamp if it does
            if os.path.exists(base_output_file):
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                output_file = f"all_declarations_data_{timestamp}.xlsx"
            else:
                output_file = base_output_file
                
            df.to_excel(output_file, index=False)
            logger.info(f"\nAll data saved to {output_file}")
        else:
            logger.warning("No data was collected to save")
            
    finally:
        scraper.close()

if __name__ == "__main__":
    main() 