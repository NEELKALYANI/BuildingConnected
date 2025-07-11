import requests
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import os
from datetime import datetime

# For automatic ChromeDriver management
try:
    WEBDRIVER_MANAGER_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False


class EmployeeDataExtractor:
    def __init__(self, debug_port=9222):
        self.debug_port = debug_port
        self.driver = None
        self.data = []
        
    def connect_to_existing_chrome(self):
        """Connect to existing Chrome debugging session"""
        try:
            # === Connect to existing Chrome session ===
            chrome_options = Options()
            chrome_options.debugger_address = f"127.0.0.1:{self.debug_port}"
            
            # Try to use WebDriver Manager for automatic ChromeDriver management
            if WEBDRIVER_MANAGER_AVAILABLE:
                try:
                    service = Service(ChromeDriverManager().install())
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                except Exception as wdm_error:
                    print(f"WebDriver Manager failed: {wdm_error}")
                    print("Falling back to default ChromeDriver...")
                    driver = webdriver.Chrome(options=chrome_options)
            else:
                driver = webdriver.Chrome(options=chrome_options)
            
            self.driver = driver
            print(" Successfully connected to Chrome debugging session")
            return True
        except Exception as e:
            print(f" Failed to connect to Chrome: {e}")
            print("\nPossible solutions:")
            print("1. Make sure Chrome is running with debugging enabled:")
            print(f'   chrome.exe --remote-debugging-port={self.debug_port} --user-data-dir="C:\\chrome_debug"')
            print("2. Update ChromeDriver to match your Chrome version")
            print("3. Install webdriver-manager: pip install webdriver-manager")
            return False
    
    def navigate_to_url(self, url):
        """Navigate to the specified URL"""
        try:
            self.driver.get(url)
            print(f" Navigated to: {url}")
            time.sleep(3)  # Wait for page to load fully
            return True
        except Exception as e:
            print(f" Failed to navigate to URL: {e}")
            return False
    
    def wait_for_element(self, xpath, timeout=10):
        """Wait for element to be present"""
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return element
        except TimeoutException:
            print(f"✗ Timeout waiting for element: {xpath}")
            return None
    
    def extract_employee_data(self):
        """Extract employee data from the page"""
        print("Starting data extraction...")
        
        # Full XPath for employee data
        xpaths = {
            'name': '/html/body/div[1]/div/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[1]/div/div[2]/div[1]/a',
            'designation': '/html/body/div[1]/div/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div',
            'email': '/html/body/div[1]/div/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[1]',
            'phone': '/html/body/div[1]/div/div[1]/div[1]/div[2]/div/div/div[2]/div[1]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[2]'
        }
        
        # Alternative CSS selectors 
        css_selectors = {
            'name': 'a[class*="userName"]',
            'designation': 'div[class*="title"]',
            'email': 'div[data-id="employee-email"]',
            'phone': 'div[data-id="employee-phone"]'
        }
        
        employee_rows = []
        
        # Method 1: Try to find by the specific XPath structure
        try:
            # Look for employee container divs
            containers = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'ReactVirtualized')]//div[contains(@role, 'row')]")
            print(f"Found {len(containers)} potential employee containers")
            
            for container in containers:
                try:
                    # Try to extract data from each container
                    employee_data = {}
                    
                    # Try different methods to extract name
                    name_element = None
                    try:
                        name_element = container.find_element(By.XPATH, ".//a[contains(@class, 'userName') or contains(@data-id, 'user-name')]")
                    except:
                        try:
                            name_element = container.find_element(By.XPATH, ".//a")
                        except:
                            pass
                    
                    if name_element:
                        employee_data['name'] = name_element.text.strip()
                    
                    # Try to extract designation
                    try:
                        designation_element = container.find_element(By.XPATH, ".//div[contains(@class, 'title')]")
                        employee_data['designation'] = designation_element.text.strip()
                    except:
                        employee_data['designation'] = 'N/A'
                    
                    # Try to extract email
                    try:
                        email_element = container.find_element(By.XPATH, ".//div[@data-id='employee-email']")
                        employee_data['email'] = email_element.text.strip()
                    except:
                        employee_data['email'] = 'N/A'
                    
                    # Try to extract phone
                    try:
                        phone_element = container.find_element(By.XPATH, ".//div[@data-id='employee-phone']")
                        employee_data['phone'] = phone_element.text.strip()
                    except:
                        employee_data['phone'] = 'N/A'
                    
                    # Only add if we found at least a name
                    if 'name' in employee_data and employee_data['name']:
                        employee_rows.append(employee_data)
                        print(f" Extracted: {employee_data['name']}")
                
                except Exception as e:
                    print(f" Error extracting from container: {e}")
                    continue
        
        except Exception as e:
            print(f" Error finding employee containers: {e}")
        
        # Method 2: Try the specific XPaths you provided
        if not employee_rows:
            print("Trying specific XPaths...")
            try:
                employee_data = {}
                
                for field, xpath in xpaths.items():
                    try:
                        element = self.driver.find_element(By.XPATH, xpath)
                        employee_data[field] = element.text.strip()
                        print(f" Found {field}: {employee_data[field]}")
                    except:
                        employee_data[field] = 'N/A'
                        print(f" Could not find {field}")
                
                if any(value != 'N/A' for value in employee_data.values()):
                    employee_rows.append(employee_data)
            
            except Exception as e:
                print(f" Error with specific XPaths: {e}")
        
        # Method 3: Try to find by visible text patterns
        if not employee_rows:
            print("Trying text pattern matching...")
            try:
                # Look for elements containing email patterns
                all_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), '@')]")
                for element in all_elements:
                    if '@' in element.text:
                        print(f"Found potential email: {element.text}")
                        
                # Look for phone number patterns
                phone_elements = self.driver.find_elements(By.XPATH, "//*[contains(text(), '+1') or contains(text(), '(')]")
                for element in phone_elements:
                    print(f"Found potential phone: {element.text}")
                        
            except Exception as e:
                print(f" Error with pattern matching: {e}")
        
        self.data = employee_rows
        print(f" Total employees extracted: {len(self.data)}")
        return self.data
    
    def save_to_excel(self, filename=None):
        """Save extracted data to Excel file"""
        if not self.data:
            print(" No data to save")
            return False
        
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"employee_data_{timestamp}.xlsx"
        
        try:
            df = pd.DataFrame(self.data)
            df.to_excel(filename, index=False)
            print(f" Data saved to: {filename}")
            print(f" Saved {len(self.data)} employee records")
            return True
        except Exception as e:
            print(f" Error saving to Excel: {e}")
            return False
    
    def print_extracted_data(self):
        """Print extracted data to console"""
        if not self.data:
            print("No data extracted")
            return
        
        print("\n" + "="*50)
        print("EXTRACTED EMPLOYEE DATA")
        print("="*50)
        
        for i, employee in enumerate(self.data, 1):
            print(f"\nEmployee {i}:")
            print(f"  Name: {employee.get('name', 'N/A')}")
            print(f"  Designation: {employee.get('designation', 'N/A')}")
            print(f"  Email: {employee.get('email', 'N/A')}")
            print(f"  Phone: {employee.get('phone', 'N/A')}")
    
    def close(self):
        """Close the driver connection"""
        if self.driver:
            # Don't quit the driver as it's an existing session
            print(" Disconnected from Chrome session")

def main():
    """Main execution function"""
    url = "https://app.buildingconnected.com/companies/5cf7cd58d8ee170033942880/offices/5cf7cd58d8ee170033942881/employees"
    
    extractor = EmployeeDataExtractor(debug_port=9222)
    
    try:
        # Connect to existing Chrome session
        if not extractor.connect_to_existing_chrome():
            print("Make sure Chrome is running with debugging enabled:")
            print('chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chrome_debug"')
            return
        
        # Navigate to the URL
        if not extractor.navigate_to_url(url):
            return
        
        # Wait a moment for the page to fully load
        print("Waiting for page to load completely...")
        time.sleep(5)
        
        # Extract employee data
        data = extractor.extract_employee_data()
        
        # Print extracted data
        extractor.print_extracted_data()
        
        # Save to Excel
        if data:
            extractor.save_to_excel()
        else:
            print("No data extracted. Please check the page structure.")
            
            # Print current page source for debugging
            print("\nCurrent page title:", extractor.driver.title)
            print("Current URL:", extractor.driver.current_url)
            
            # Try to find any elements that might contain employee data
            print("\nLooking for potential employee data elements...")
            try:
                # Look for any links that might be employee names
                links = extractor.driver.find_elements(By.TAG_NAME, "a")
                print(f"Found {len(links)} links on the page")
                
                # Look for any elements with email-like text
                all_text = extractor.driver.find_elements(By.XPATH, "//*[contains(text(), '@')]")
                print(f"Found {len(all_text)} elements containing '@' symbol")
                
                # Print some sample text content for debugging
                body_text = extractor.driver.find_element(By.TAG_NAME, "body").text
                print(f"Page contains {len(body_text)} characters of text")
                
            except Exception as e:
                print(f"Error during debugging: {e}")
    
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        extractor.close()

if __name__ == "__main__":
    print("Employee Data Extractor")
    print("=" * 30)
    
    # Check if required packages are installed
    try:
        if WEBDRIVER_MANAGER_AVAILABLE:
            print(" All required packages are available (including webdriver-manager)")
        else:
            print(" Required packages are available")
            print(" Note: Install webdriver-manager for automatic ChromeDriver management:")
            print("       pip install webdriver-manager")
    except ImportError as e:
        print(f" Missing required package: {e}")
        print("Please install with: pip install selenium pandas openpyxl webdriver-manager")
        exit(1)
    
    main()