# meraki-api-record-macaddr

Summary:

Purpose of this script is to go through all serial numbers in the spreadsheet, pull mac addresses of those devices, and 
record them in that spreadsheet.

Requirements:

1) Interpreter: Python 3.9.0+
2) Python Packages: requests, json, openpyxl
3) Excel Spreadsheet - .XLSX format
4) API support for the Organization is enabled in Meraki Dashboard. Admin has generated their custom API key.
5) Device licensing, and inventory for devices used below, have been claimed under Organization settings.

How to run:

1) Attached is a spreadsheet that can be used. Custom spreadsheet can be used, but column variables must be changed under
   PARAMETERS section.
2) Open claim_devices_to_network.py with your favorite text editor and edit PARAMETERS sections of the script:
    1) Lines 11-16 is mandatory.
    2) Line 21-23 are required if using custom spreadsheet.
3) Run python3 claim_devices_to_network.py in your terminal. Ensure spreadsheet file and python script are
   in the same folder/location.