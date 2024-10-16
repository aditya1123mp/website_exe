import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# List of websites to check
websites = [
    {"site_name": "PortableApps", "site_url": "https://portableapps.com/"},
    {"site_name": "AppSticks", "site_url": "https://www.app-stick.com/"},
    {"site_name": "liberkey", "site_url": "https://www.liberkey.com/en.html"},
    {"site_name": "Pendriveapps", "site_url": "https://pendriveapps.com/"},
    {"site_name": "linuxliveusb", "site_url": "https://www.linuxliveusb.com/en/download"},
    {"site_name": "portablefreeware", "site_url": "https://www.portablefreeware.com/"},
    {"site_name": "Ssuitebladerunner", "site_url": "https://ssuitesoft.com/"},
    {"site_name": "softpedia", "site_url": "https://www.softpedia.com/"},
    {"site_name": "majorgeek", "site_url": "https://www.majorgeeks.com/"},
    {"site_name": "uptodown", "site_url": "https://en.uptodown.com/"},
    {"site_name": "techspot", "site_url": "https://www.techspot.com/"},
    {"site_name": "sourceforge", "site_url": "https://sourceforge.net/"},
    {"site_name": "filehorse", "site_url": "https://www.filehorse.com/"},
    {"site_name": "cnet", "site_url": "https://www.cnet.com/"},
    {"site_name": "downloadcrew", "site_url": "https://www.downloadcrew.com/"},
    {"site_name": "softradar", "site_url": "https://softradar.com/"},
    {"site_name": "freedownload_manager", "site_url": "https://www.freedownloadmanager.org/"}
]

# Create an Excel workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Website Status"

# Create header row
sheet.append(["site_name", "site_url", "status"])

# Set ChromeOptions to run the browser in the background
chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Disable WebDriver flag
chrome_options.add_argument("--window-position=-10000,-10000")  # Move the browser off-screen
# chrome_options.add_argument("--headless")  # Uncomment this to run in headless mode

# Initialize WebDriver using WebDriverManager for handling driver setup
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

try:
    # Iterate through each website and check the status
    for website in websites:
        try:
            # Navigate to the website
            driver.get(website['site_url'])

            # Get the title of the website
            page_title = driver.title

            # Check if the website is open by checking the page title
            if page_title:
                print(f"{website['site_name']} is open. Title: {page_title}")
                sheet.append([website['site_name'], website['site_url'], "open"])  # Write "open" in the status column
            else:
                print(f"{website['site_name']} did not open successfully.")
                sheet.append([website['site_name'], website['site_url'], "block"])  # Write "block" in the status column

        except Exception as e:
            print(f"Error opening {website['site_name']}: {e}")
            sheet.append([website['site_name'], website['site_url'], "block"])  # Write "block" in case of error

    # Save the Excel file
    excel_file = "C:/Users/DELL/OneDrive/Desktop/website_status.xlsx"
    workbook.save(excel_file)
    print(f"Excel file '{excel_file}' created successfully.")

except Exception as e:
    print(f"Error: {e}")

finally:
    # Close the browser
    driver.quit()
