# Automation and Analytics Projects

# 1. Invoice Consolidation Application

## Overview
A Tkinter-based PC application for outbound document consolidation. Eliminated operation time by 1 man hour daily.

## Features
- Download attachments from Outlook with the key codes keyed in by User.
- Extract required information from .txt files under the folder selected by the user into the required format.
- Data cleaning and consolidating Excel files under the folder selected by the user into one report.
- .bat created for the user to trigger the application
- Additonal .bat file for the user to download the required Python environment as a bypass of corporate cyber security restriction.

## Technologies Used
- Python
- Batch File
- win32com.client
- pathlib
- pandas
- tkinter
- Pillow
- openpyxl
- os


# 2. SAP Reports Extraction and Visualization

## Overview
A warehouse space utilization dashboard with inbound and outbound status. Data from an automation that triggers .vbs macros for SAP data extraction and sending the reports to the corporate network in 15-minute intervals with an annual saving and avoidance higher than S$72,000.

## Features
- Trigger 6 separate .vbs SAP GUI Script
- Capture the current date and pass it as an argument from Python Script to .vbs
- Caffeine function to prevent the device from sleeping during the interval
- Attach 14 SAP reports and send them to the corporation network via Outlook
---
Downstream:
- Automatically save the reports to SharePoint using Microsoft PowerAutomate Flow
- Data visualization on PowerBI



## Technologies Used
- Python
- Batch File
- VBScript
- PyAutoGUI
- Threading



# 3. SAP Data Entry Automation

## Overview
Warehouse admin staff will need to key in shipment information into SAP to print out documents for cargo releasing. These templates are to automate this process to increase accuracy and boost efficiency.

## Technologies Used
- SAP GUI Scripting
- VBA
- Windows API


# 4. Shipment Forecasting

## Overview
The cargo of this customer is split into two categories, main parts and subparts. New orders will only show the main parts. Thus, as their supply chain partner, we will need to estimate the number of subparts to forecast the volume of each shipment. Through analysis, we discovered the number of subparts of each main part is consistent for the same buyer buying the same type of machine. I created an automated data pipeline to capture all shipped-out cargos into a library, and this Python code is to refer to that library and adjusts the number of subparts.

With this automation, we increased the accuracy of the forecast from 80% to 95%, reducing freight and operations costs, and strengthening customer relationships.

## Technologies Used
- Python Pandas

