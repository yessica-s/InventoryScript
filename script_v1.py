"""
Class to handle being inputted with the CSV file for monthly product sales 
and to be formatted into a new sheet of the inventory spreadsheet.
"""

import subprocess
import sys
import os
import uno  # import to format LibreOffice
import time  # Add this import at the top
# import openpyxl

def format_data() -> None:
    print("Formatting data...")
    # Your formatting logic here
    return

def check_path(file_path: str) -> bool:
    """Function to check if the file path is valid."""
    if os.path.exists(file_path):
        with open(file_path, "r"):
            return True
    return False
    

def start_libreoffice():
    """Start LibreOffice in headless mode."""
    command = [
        "libreoffice",
        "--headless",
        "--accept=socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    ]
    
    # Start LibreOffice
    process = subprocess.Popen(command)
    print("Started LibreOffice in headless mode.")
    time.sleep(2)  # Add a 2-second delay
    return process


def terminate_libreoffice():
    """Terminate any existing LibreOffice processes."""
    subprocess.run(["pkill", "soffice"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print("Terminated existing LibreOffice processes.")

def connect_to_libreoffice():
    """Connect to a running instance of LibreOffice."""
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    print("Connected to LibreOffice.")
    return context

def load_spreadsheet(context, spreadsheet_path):
    """Load the existing spreadsheet."""
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    print(f"Loading spreadsheet: {spreadsheet_path}")
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(spreadsheet_path)), "_blank", 0, ())

def create_new_sheet(sheet, new_sheet_name):
    """Create a new sheet in the spreadsheet."""
    try:
        # Create a new sheet
        new_sheet = sheet.Sheets.insertNewByName(new_sheet_name, sheet.Sheets.getCount())
        print(f"New sheet '{new_sheet_name}' created successfully.")
    except Exception as e:
        print(f"Error creating new sheet: {e}")

def close_spreadsheet(sheet):
    """Close the spreadsheet and save changes."""
    try:
        sheet.store()  # Save changes
        sheet.close()  # Close the sheet
    except Exception as e:
        print(f"Error closing spreadsheet: {e}")


def main():
    print("Starting script...")

    terminate_libreoffice() # terminate any existing libreoffice processes

    if len(sys.argv) != 2: 
        print("Usage error: python3 monthly_transfer_script.py <path_to_monthly_csv_file_data>")
        sys.exit(1)

    file_path = sys.argv[1] # Path to XLSX Data Download
    
    if not check_path(file_path):  # Check if the CSV file path is valid
        print(f"Invalid file path: {file_path}")
        sys.exit(1)

    spreadsheet_path = "/home/medigroup/Documents/Inventory.ods"  # Assign path to inventory spreadsheet
    sheet_name = "Monthly Sales" # New sheet to be created

    # start_libreoffice()
    
    context = connect_to_libreoffice()
    sheet = load_spreadsheet(context, spreadsheet_path)
    
    create_new_sheet(sheet, sheet_name)

    try:
        worksheet = sheet.getSheets().getByName(sheet_name)
        print(f"Accessed sheet: {sheet_name}")
    except Exception as e:
        print(f"Failed to open newly created sheet {sheet_name}")
        print(f"Error: {e}")
        return

    format_data()
    # close_spreadsheet(sheet)

if __name__ == "__main__":
    main()