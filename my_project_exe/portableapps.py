from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service  # Import Service class
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import time
import openpyxl
import selenium

def main():
# Path to the chromedriver
    chrome_driver_path = 'C:\\Users\\DELL\\OneDrive\\Desktop\\chromedriver-win64\\chromedriver.exe'

# Initialize Chrome WebDriver with options and Service
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")

# Use Service to set up the chromedriver path
    service = Service(executable_path=chrome_driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

# Navigate to PortableApps website
    driver.get('https://portableapps.com/')
    time.sleep(3)

# XPaths for the elements
    logo_xpath = '//a[@id="logo"]'
    login_xpath = "//ul[@id='main-menu-links']/li/a[text() = 'Login / Create Account']"
    download_xpath = "//ul[@id='main-menu-links']/li/a[text() = 'Download']"

# Check if the elements are present
    is_logo_present = len(driver.find_elements(By.XPATH, logo_xpath)) > 0
    is_login_xpath = len(driver.find_elements(By.XPATH, login_xpath)) > 0
    is_download_xpath = len(driver.find_elements(By.XPATH, download_xpath)) > 0

# Create Excel workbook and sheet
    file_path = "C:\\Users\\DELL\\OneDrive\\Desktop\\Websites.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Websites"

# Create header row
    sheet['A1'] = "Name of Website"
    sheet['B1'] = "URLs of Website"
    sheet['C1'] = "Aditya-pc"

# If all elements are present, write 'open', otherwise write 'block'
    if is_logo_present and is_login_xpath and is_download_xpath:
        print("All elements are present, adding 'open' status to Excel")
        sheet.append(["PortableApps", "https://portableapps.com/", "open"])
    else:
        print("One or more elements are not present, adding 'block' status to Excel")
        sheet.append(["PortableApps", "https://portableapps.com/", "block"])

# Save Excel file
    workbook.save(file_path)
    print(f"Excel file created successfully at {file_path}")

# Close browser
    driver.quit()
# Execute the script only if it's run directly
if __name__ == "__main__":
    main()
