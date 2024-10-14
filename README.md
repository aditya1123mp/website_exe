Website XPath Checker

This Python project checks 17 different websites to determine if specific XPath elements are present on each site. It automates the process of scanning websites, recording the presence of certain elements, and generating an Excel file with the results.

 Features

- XPath Validation: The script checks if specific XPath elements exist on the target websites.
- Excel Report: After the script runs, an Excel file is generated with the status of each XPath for each website.
  - If the XPath is found, the status is marked as `Open`.
  - If the XPath is not found, the status is marked as `Block`.
- Cross-platform Compatibility: The project includes an executable (`.exe`) file, allowing the program to run on systems without Python or PyCharm installed.
  - Simply double-click the `.exe` file to execute the program and generate the Excel report.
  
 Prerequisites

If you're running the Python version of the project, you'll need the following installed on your system:

- Python 3.x
- Libraries listed in the `requirements.txt` file (you can install them using `pip install -r requirements.txt`)

If you're using the executable, no installations are required. You can run the `.exe` file directly on any Windows machine.

 Installation & Setup

 Running the Python Script

1. Clone the Repository:
   ```bash
   git clone https://github.com/your-username/your-repo.git
   cd your-repo
   ```

2. Install Dependencies:
   Install the required Python packages by running:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the Script:
   To start checking the websites for the given XPath:
   ```bash
   python fulltestcase.py
   ```

 Running the Executable File

1. Download the Executable:
   Navigate to the `dist` folder and locate the `fulltestcase.exe` file.

2. Run the Executable:
   Double-click the `fulltestcase.exe` file to run the program.

3. Output:
   Once the program finishes, an Excel file will be created in the same directory, containing the status of each website’s XPath check.

 Output

The Excel file generated will contain the following columns:

| Website URL | XPath Status |
|-------------|--------------|
| example.com | Open/Block   |
| website2.com| Open/Block   |

- Open: If the specified XPath is found on the website.
- Block: If the XPath is not found on the website.

 Packaging as an Executable

This project uses PyInstaller to convert the Python script into a standalone executable. This allows it to run on any Windows machine without requiring Python or other dependencies.

To package the project into an `.exe` file:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Create the executable:
   ```bash
   pyinstaller --onefile --hidden-import=selenium --hidden-import=openpyxl fulltestcase.py
   ```

The executable will be located in the `dist` directory after the build completes.
