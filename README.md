# Telework Updates Tracker

This Python script automates the process of collecting employee telework agreement files and storing relevant information in an Excel spreadsheet.

## Features

- Reads employee telework agreement files (PDF and DOCX) from structured directories.
- Extracts employee names and calculates expiry dates based on folder names.
- Saves the collected data into an Excel spreadsheet, avoiding duplicate entries.

## Requirements

To run this script, you need:

- Python 3.x
- Required Python packages:
  - `pandas`
  - `openpyxl`
  - `python-dotenv`

You can install the necessary packages using `pip3`:
```bash
pip3 install pandas openpyxl python-dotenv
```

Installation
Clone the repository (if applicable):
```bash
Copy code
git clone <repository-url>
cd <repository-directory>
```

Set up a virtual environment (optional but recommended):
```bash
Copy code
python3 -m venv venv
deactivate (when you are finished in venv)
source venv/bin/activate
```

Install the required packages:
Copy code
```bash
pip install pandas openpyxl python-dotenv
```
Create a .env file in the project directory to set the required environment variables:

Copy code
```bash
BASE_PATH=/path/to/employee/files
OUTPUT_FILE=/path/to/output/spreadsheet.xlsx
BASE_PATH: Directory containing employee folders organized by year-month.
OUTPUT_FILE: Path to the Excel file where the data will be saved.
```

Usage
Run the script with:
```bash
python3 teleworkupdates.py
```

The script will:
Load environment variables from the .env file.
List all employee files in the specified base path.
Extract employee names and calculate their telework agreement expiry dates.
Save the data into the specified Excel spreadsheet.
