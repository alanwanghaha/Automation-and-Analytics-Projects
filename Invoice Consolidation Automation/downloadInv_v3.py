# Download .txt Invoices by Batch#

import win32com.client
from win32com.client import Dispatch
from pathlib import Path
import pandas as pd
import re
import tkinter as tk
from tkinter import ttk
import os
import time
import ctypes



# Get the current user's desktop directory
DESKTOP_DIR = Path.home() /"OneDrive"/ "Desktop"

# Load the warehouse codes mapping from Excel file 'WarehouseCode.xlsx' 
current_script_path = Path(__file__).parent
WAREHOUSE_CODE_EXCEL = Path.home()/ "OneDrive" / "Invoice - General"/ "WarehouseCode.xlsx"

# Remind the user if WarehouseCode.xlsx is not found
if WAREHOUSE_CODE_EXCEL.exists():
    df_warehouse = pd.read_excel(WAREHOUSE_CODE_EXCEL)
else:
    print("Can't find WarehouseCode file at:", WAREHOUSE_CODE_EXCEL)

warehouse_dict = {str(code).strip(): name.strip() for code, name in zip(df_warehouse['Code'], df_warehouse['Description'])}

# To open outlook automatically if it's not already opened
def open_outlook():
    try:
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        useremail = outlook.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        return outlook, useremail
    
    except win32com.client.pywintypes.com_error as error:
        # Check for the specific E_ABORT error
        if error.hresult == -2147467260:  
            # Outlook might not be open, try to start Outlook
            print("Outlook is not opened. Attempting to start Outlook...")
            os.startfile("outlook")
            # Wait for Outlook to start
            print("Please wait, starting Outlook...")
            time.sleep(15)  # Wait 15 seconds for Outlook to open
            try:
                # Try to connect to Outlook again
                outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
                useremail = outlook.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                return outlook, useremail
            except Exception as e:
                # If it still fails, prompt the user to open Outlook manually
                ctypes.windll.user32.MessageBoxW(0, "Please open Outlook manually and restart the program.", "Outlook Not Opened", 1)
                raise e  # Re-raise the exception to exit the script
        else:
            # If it's a different COM error, re-raise it
            raise

# Now use the function to ensure Outlook is open
outlook, useremail = open_outlook()


# Find warehouse code from email content
def extract_warehouse_code(email_content):
    match = re.search(r"Destination: (\d+)", email_content)
    if match:
        #print('match.group(1):', match.group(1))
        return match.group(1).strip()
    return None

# Get warehouse code from Excel
def get_warehouse_name(code):
    #print('warehouse_dict.get(code): ', warehouse_dict.get(code))
    return warehouse_dict.get(code, 'Unknown')



def save_attachment_to_desktop(email, is_exception=False, selected_date=None):
    # Extracting the warehouse code
    warehouse_code = extract_warehouse_code(email.Body) 
    warehouse_name = get_warehouse_name(warehouse_code)

    # Determine the brand based on the first digit of the warehouse code
    brand_dict = {
        '2': 'Brand 1',
        '3': 'Brand 2',
        '5': 'Brand 3',
        '6': 'Brand 4',
        '8': 'Brand 5'
    }
    brand = brand_dict.get(warehouse_code[0], 'Unknown Brand')

    # if warehouse name is not found in the warehouse code sheet, name the folder as the warehouse code found in the email
    if warehouse_name == 'Unknown':
        return False

    # # Determine folder structure based on the selected date or the received date of the email
    foldername = 'Invoice'
    # if is_exception and selected_date:
    #     year, month, day = selected_date.strftime('%Y'), selected_date.strftime('%b'), selected_date.strftime('%d')
    # else:
    #     received_date = email.ReceivedTime
    #     year = str(received_date.year)
    #     month = received_date.strftime('%b')
    #     day = str(received_date.day).zfill(2)

    folder_path = DESKTOP_DIR / foldername / brand / warehouse_name
    folder_path.mkdir(parents=True, exist_ok=True)  # Create folders recursively

    if email.Attachments.Count > 0:
        for attachment in email.Attachments:
            local_path = folder_path / attachment.FileName
            attachment.SaveAsFile(str(local_path))
            print(local_path, ' is downloaded')

    return True


# Define a function to move processed emails to another Outlook folder
def move_email_to_folder(email, newFolderName):
    mailbox = outlook.Session.DefaultStore.GetRootFolder()
    dst_folder = None
    try:
        dst_folder = mailbox.Folders(newFolderName)
    except:
        # If the folder does not exist, create it
        dst_folder = mailbox.Folders.Add(newFolderName)

    email.Move(dst_folder)


# loop thru Outlook folder 'Inv'
def process_main_code(batch_numbers_set):
    inbox = outlook.Folders[useremail].Folders['Inbox'].Folders['Inv']
    # Dictionary to track the status of each batch number
    batch_numbers_status = {batch_number: False for batch_number in batch_numbers_set}

    # Collect all email items into a list
    email_list = [email for email in inbox.Items]

    for email in email_list:
        # Extract batch number from email
        match = re.search(r"JDA Batch Number: (\S+)", email.Body)
        if match:
            email_batch_number = match.group(1).strip()
            if email_batch_number in batch_numbers_set:
                warehouse_code = extract_warehouse_code(email.Body)
                warehouse_name = get_warehouse_name(warehouse_code)
                # Check if the warehouse code is found and process accordingly
                if warehouse_name != 'Unknown':
                    if save_attachment_to_desktop(email):
                        batch_numbers_status[email_batch_number] = True
                        move_email_to_folder(email, 'Downloaded Invoices')

    # Check for any batch numbers that were not processed
    for batch_number, processed in batch_numbers_status.items():
        if not processed:
            print(f"The Invoice for {batch_number} is not downloaded because: \n", 
                  '1. Batch# is not found in emails. \n',
                  '2. The warehouse code for ',batch_number, 'is not in the library [WarehouseCode.xlsx]')

def main():
    # Create the main GUI window
    root = tk.Tk()
    root.title("Download Invoice")

    def process_emails():
        batch_numbers = batch_number_entry.get("1.0", tk.END).split('\n')  # Split the input by newlines
        # Remove any empty strings and strip whitespace
        batch_numbers_set = {batch.strip() for batch in batch_numbers if batch.strip()}

        process_main_code(batch_numbers_set)

    # Label & Text Box for Batch Numbers
    batch_number_label = ttk.Label(root, text="Paste Batch Numbers (one per line):")
    batch_number_label.pack(padx=20, pady=5, anchor='w')

    # Use a Text widget instead of Entry to allow multiple lines
    batch_number_entry = tk.Text(root, height=15, width=50)
    batch_number_entry.pack(padx=20, pady=5, fill='both', expand=True)

    # Process button
    process_button = ttk.Button(root, text="Start Processing", command=process_emails)
    process_button.pack(padx=20, pady=20)

    root.mainloop()


if __name__ == "__main__":
    main()
