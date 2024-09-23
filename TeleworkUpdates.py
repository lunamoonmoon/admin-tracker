import os
import pandas as pd
from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import datetime, timedelta

# Load environment variables from .env file
load_dotenv()

# Get paths from environment variables
base_path = os.getenv('BASE_PATH')
output_file = os.getenv('OUTPUT_FILE')

# Debug print statements
print(f"Base path: {base_path}")
print(f"Output file: {output_file}")

def list_employee_files(base_path):
    data = []
    for folder_name in os.listdir(base_path):
        year_month_path = os.path.join(base_path, folder_name)
        if os.path.isdir(year_month_path):
            try:
                year, month_name = folder_name.split('-')
                month = datetime.strptime(month_name.strip(), '%b').month
                expiry_date = datetime(int(year), month, 1) + timedelta(days=365)
                expiry_date_str = expiry_date.strftime('%Y-%m-%d')
            except ValueError:
                print(f"Folder name format unexpected: {folder_name}")
                continue
            
            for file_name in os.listdir(year_month_path):
                if "consideration" in file_name.lower():
                  print(f"Skipping file: {file_name}")
                  continue

                if file_name.endswith(('.pdf', '.docx')):
                    try:
                        employee_name = file_name.split('_')[1]  # Adjust as necessary
                        data.append({
                            'Employee Name': employee_name, 
                            'Year-Month': folder_name,
                            'Expiry Date': expiry_date_str
                        })
                    except IndexError:
                        print(f"Filename format unexpected: {file_name}")

    print(f"Collected {len(data)} entries.")
    return data

def save_to_spreadsheet(data, output_file):
    if not data:  # Check if there's any data to save
        print("No data to save.")
        return

    # Delete the output file if it exists to avoid BadZipFile errors
    if os.path.exists(output_file):
        os.remove(output_file)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df = pd.DataFrame(data)
        df.to_excel(writer, index=False)

    print(f'Spreadsheet saved as: {output_file}')

def main():
    employee_data = list_employee_files(base_path)
    print(f"Data to save: {employee_data}")  # Print the collected data
    save_to_spreadsheet(employee_data, output_file)

if __name__ == '__main__':
    main()
