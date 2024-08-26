# Warehouse Operations Dashboard
### This is a auto-refreshing close-to-realtime dashboard displaying warehouse operations data such as:
- **Warehouse Space Utilization:** Mapping out the real-time inventory utilization rate for KPI reporting and operational planning.
- **Inbound Tracker:** Indicating the volume of incoming shipments matching with the actual shipment received.
- **Outbound Overview:** Tracking the outbound activities from picking to ship out to ensure order fulfilment.

## Solution
Tools used to automate the data pipeline as below:
- **SAP Visual Basic Script:** To pull out the required 14 reports from SAP automatically.
- **Python:** To trigger the SAP automation and send reports from LAM PC to Schenker network via Outlook.
- **Power Automate Flow:** To capture the reports from Outlook and save them in SharePoint which linked to Power BI.
- **Power BI:** Data analytics and visualization.




