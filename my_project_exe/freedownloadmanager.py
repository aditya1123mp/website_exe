from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
#from openpyxl import Workbook
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

   # Navigate to the website
   driver.get('https://www.freedownloadmanager.org/')
   time.sleep(3)

   # Define the XPath for the elements
   appsticklogoXPath = "//a[contains(text(),'Features')]"
   sidebarpolpularXPath = "//a[contains(text(),'Awards')]"
   sidebarprogrammingXPath = "//a[contains(text(),'Forum')]"

   # Check if the elements are present using find_elements (returns list)
   isAppStickLogoPresent = len(driver.find_elements(By.XPATH, appsticklogoXPath)) > 0
   isSidebarPopularPresent = len(driver.find_elements(By.XPATH, sidebarpolpularXPath)) > 0
   isSidebarProgrammingPresent = len(driver.find_elements(By.XPATH, sidebarprogrammingXPath)) > 0

   # Path to the existing Excel file
   file_path = "C:\\Users\\DELL\\OneDrive\\Desktop\\Websites.xlsx"

   # Open the existing Excel file
   workbook = load_workbook(file_path)
   sheet = workbook['Websites']

   # If all conditions are satisfied, write 'open', else write 'block'
   if isAppStickLogoPresent and isSidebarPopularPresent and isSidebarProgrammingPresent:
       print("All elements are present, adding 'open' status to Excel file")
       sheet.append(["freedownload_manager", "https://www.freedownloadmanager.org/", "open"])
   else:
       print("One or more elements are not present, adding 'block' status to Excel file")
       sheet.append(["freedownload_manager", "https://www.freedownloadmanager.org/", "block"])

   # Save changes to the Excel file
   workbook.save(file_path)
   print(f"Excel file updated successfully at {file_path}")

   # Close browser
   driver.quit()

# Execute the script only if it's run directly
if __name__ == "__main__":
    main()
