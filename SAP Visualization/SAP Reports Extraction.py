import subprocess
import win32com.client as win32
import time
from datetime import datetime, timedelta
import pyautogui
import threading
import os


# function to run SAP Macro Script
def run_vbs_script(script_path, Date):
    try:
        subprocess.run(['cscript.exe', script_path, Date], check=True)
        print("VBS script: ", script_path, "executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while executing the VBS script: {e}")

def close_excel_reports():
    try:
        subprocess.run(["taskkill","/f", "/im", "excel.exe"],check=True)
        print("Reports are closed successfully")
    except subprocess.CalledProcessError as e:
        print(f"Failed to close the reports: {e}")

# function to send email
def send_reports_via_outlook(recipients, copies, subject, body, attachments):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipients
    mail.CC = copies
    mail.Subject = subject
    mail.Body = body

    # Adding attachments
    for attachment_path in attachments:
        mail.Attachments.Add(attachment_path)
    mail.Send()
    print("Email sent successfully.")

# function to adjust volume to keep device awake
def caffeine(stop_event, times=3, interval=220):
    while not stop_event.is_set():
        for _ in range(times):
            if stop_event.is_set():
                break
            pyautogui.press('volumemute')
            time.sleep(1)
            pyautogui.press('volumemute')
            print('Volume adjusted')

            # Sleep for the specified interval or until the stop event is set
            stop_event.wait(interval)


# VBS Paths: 
VBS_WarehouseUtilization = r"C:\Users\OneDrive\Documents\SAP\SAP GUI\SAP Macro\WarehouseUtilization.vbs"
VBS_MB51 = r"C:\Users\OneDrive\Documents\SAP\SAP GUI\SAP Macro\MB51.vbs"
VBS_VL06 = r"C:\Users\OneDrive\Documents\SAP\SAP GUI\SAP Macro\VL06.vbs"
VBS_ZMRB = r"C:\Users\OneDrive\Documents\SAP\SAP GUI\SAP Macro\ZMRB.vbs"

# Define email details
recipients = "alang@outlook.com; Amy@outlook.com"
copies = 'Ben@outlook.com'
subject = 'SAP Reports'
body = 'Please find attached the SAP reports.'
attachments = [r"C:\Users\OneDrive\Documents\SAP\SAP GUI\0010.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\7164.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\9150.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\9170.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\LX03.xlsx",
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\MB51_1000_1020.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\MB51_1090.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\VL06.xlsx", 
            r"C:\Users\OneDrive\Documents\SAP\SAP GUI\ZMRB.xlsx", ]


def main():
    global stop_caffeine
    stop_caffeine = threading.Event()

    while True:
        current_time = datetime.now()
        Date = current_time.strftime("%m/%d/%Y")
        print(f"Date: {Date}")
        next_run = (current_time + timedelta(minutes=15)).replace(second=0, microsecond=0)
        next_run_minute = next_run.minute // 15 * 15
        next_run = next_run.replace(minute=next_run_minute)

        seconds_until_next_run = (next_run - current_time).total_seconds()
        print(f"Current time: {current_time}")
        print(f"Scheduled next run at: {next_run}, in {seconds_until_next_run} seconds.")

        # Start the caffeine thread to keep the device awake
        caffeine_thread = threading.Thread(target=caffeine, args=(stop_caffeine,))
        caffeine_thread.start()

        # Main thread sleeps until the next quarter-hour mark
        time.sleep(seconds_until_next_run)

        # Stop the caffeine thread after waking
        stop_caffeine.set()
        caffeine_thread.join()  # Wait for the caffeine thread to finish

        # Reset the event for the next iteration
        stop_caffeine.clear()        
        
        # Run the VBS scripts to download reports
        run_vbs_script(VBS_WarehouseUtilization, Date)
        run_vbs_script(VBS_MB51, Date)
        run_vbs_script(VBS_VL06, Date)
        run_vbs_script(VBS_ZMRB, Date)

        print("Loading...")
        time.sleep(15)

        # Close the reports
        close_excel_reports()
        print('Closing all Excel workbooks...')

        time.sleep(40)
        print("Preparing to send email...")
        time.sleep(30)

        # Send the reports via Outlook
        send_reports_via_outlook(recipients, copies, subject, body, attachments)
        print(f"Data will be extracted again in another 15 mins...")


try:
    main()
except KeyboardInterrupt:
    print("Exiting...")



