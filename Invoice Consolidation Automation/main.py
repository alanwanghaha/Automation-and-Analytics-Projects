import tkinter as tk
from PIL import Image, ImageTk
import detailsReport, downloadInv_v3
from pathlib import Path
import shutil

def run_downloadInv_v2():
    downloadInv_v3.main()

def run_details():
    detailsReport.main()


DESKTOP_DIR = Path.home() / "OneDrive" / "Desktop" / "Invoice"   # /"details report"

# Define the full path to the file
details_xlsx_path = DESKTOP_DIR / "details.xlsx"

# Check if 'details.xlsx' exists
if DESKTOP_DIR.exists():
    # If it exists, delete the file
    try:
        shutil.rmtree(DESKTOP_DIR)
        print(DESKTOP_DIR, 'is found and deleted')
    except PermissionError as e:
        print(f"Permission error: {e}")


app = tk.Tk()
app.title('MY Invoice Consolidation')
app.geometry('300x200')

# open the image and resize it
try:
    logo = Image.open('Schenker logo.png')
    logo = logo.resize((182, 30), Image.Resampling.LANCZOS)  
    banner = ImageTk.PhotoImage(logo)
except Exception as e:
    print(f"An error occurred: {e}")
    banner = None  # Set banner to None if there's an error

# If the banner was successfully created, pack it in a label
if banner:
    label = tk.Label(app, image=banner)
    label.pack(pady=15)

button_download = tk.Button(app, text='Download Invoice', command=run_downloadInv_v2)
button_download.pack(pady=15, padx=30)

button_details = tk.Button(app, text='Generate Details Report', command=run_details)
button_details.pack(pady=15, padx=30)

app.mainloop()
