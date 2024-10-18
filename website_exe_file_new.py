import requests
from bs4 import BeautifulSoup
import openpyxl
import time

# Create a new Excel file with columns: Name of Website, URLs of Website, Status
def create_excel_file(file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Website Status"

    # Create the header row
    sheet.append(["Name of Website", "URLs of Website", "Status"])

    # Save the Excel file
    workbook.save(file_path)

# Function to update the Excel file with website data
def update_excel_file(file_path, website_name, url, status):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Add the website details to the next available row
    sheet.append([website_name, url, status])

    # Save the updated Excel file
    workbook.save(file_path)

# Function to fetch website and return the status
def fetch_website(website_name, url, file_path, retries=3):
    delay = 5  # 5 seconds delay before retry
    retry_count = 0
    status = "block"  # Default status is 'block'

    while retry_count < retries:
        try:
            # Send request to the website with headers to avoid 403
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3",
                "Accept-Language": "en-US,en;q=0.5",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Connection": "keep-alive"
            }

            response = requests.get(url, headers=headers, timeout=30)

            # Check if the status code is 403 (forbidden)
            if response.status_code == 403:
                raise Exception("403 Forbidden Access")

            # Parse the document if the request is successful
            soup = BeautifulSoup(response.content, "html.parser")

            # Check if title is present
            title = soup.find("title")
            if title:
                print(f"Website: {url} - Title is present")
                status = "open"
            else:
                print(f"Website: {url} - Title is not present")
            break  # Exit loop if successful

        except Exception as e:
            print(f"Website: {url} - Could not open (Error: {str(e)})")
            retry_count += 1
            if retry_count < retries:
                print(f"Retrying ({retry_count}/{retries})...")
                time.sleep(delay)  # Wait before retrying
            else:
                print(f"Website: {url} - Failed after {retries} attempts.")

    # Update the Excel file with the status
    update_excel_file(file_path, website_name, url, status)

# List of websites to check (Name, URL)
websites = [
    ["PortableApps", "https://portableapps.com/"],
    ["App-Stick", "https://www.app-stick.com/"],
    ["LiberKey", "https://www.liberkey.com/en.html"],
    ["PenDriveApps", "https://pendriveapps.com/"],
    ["LinuxLiveUSB", "https://www.linuxliveusb.com/en/download"],
    ["PortableFreeware", "https://www.portablefreeware.com/"],
    ["SSuiteSoft", "https://ssuitesoft.com/ssuiteportable.htm"],
    ["Softpedia", "https://www.softpedia.com/"],
    ["MajorGeeks", "https://www.majorgeeks.com/"],
    ["Uptodown", "https://en.uptodown.com/"],
    ["TechSpot", "https://www.techspot.com/"],
    ["SourceForge", "https://sourceforge.net/"],
    ["FileHorse", "https://www.filehorse.com/"],
    ["CNET", "https://www.cnet.com/"],
    ["DownloadCrew", "https://www.downloadcrew.com/"],
    ["Softradar", "https://softradar.com/"],
    ["FreeDownloadManager", "https://www.freedownloadmanager.org/"]
]

# Path to save the Excel file
excel_file_path = "C:/Users/DELL/OneDrive/Desktop/Website_Status.xlsx"

# Create the Excel file initially
create_excel_file(excel_file_path)

# Loop through each website and fetch its status
for website in websites:
    name, url = website
    fetch_website(name, url, excel_file_path)

print("Process completed! The results are saved in the Excel file.")
