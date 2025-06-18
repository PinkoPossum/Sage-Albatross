# Sage-Albatross

Python and powershell script. Powershell script will install python, python environment, and dependencies. Python script will perform a cisco audit on devices specified by a spreadsheet and output the audit into a dated spreadsheet.

## Automated Network Assistant

You give it a list of your Cisco switches and routers in an Excel file. The script then goes through that list one by one, logs into each device for you, and runs a set of basic diagnostic commands.

The information it gathered is put into a new, neatly organized Excel spreadsheet. As a helpful feature, the new spreadsheet is named "dd-mm-yy-audit" for easy organization, highlights any network ports that are down in red, so you can easily spot potential problems, and it flags any devices it couldn't log into for you to check manually.

## Automated Environment Setup

The PowerShell script is meant to simplify using the python script. The PowerShell Script will ensure Python is installed on your machine, creates a virtual environment for running the script, installs the required dependencies (ie openpyxl and netmiko), tells you how to enter the environment and work in it.

## Creating Your Device List

1. Open a spreadsheet program that can output in .xlsx format.
2. In cell A1, typ the header "Devices".
3. Starting in cell A2 list the IP address or domain names of your switches and routers.
4. Save the file as "devices.xlsx".

## To Use The Script

1. Navigate to your script folder
2. Hold "shift", right-click empty space in the folder
3. Select "Open PowerShell window here".
4. In the terminal type:
   > .\venv\Scripts\Activate.ps1
   > 
   > python cisco_audit.py
5. The script will have you input your SSH username and password, they are not stored for security reasons but you could modify the script to store them for you if you don't care.
6. After the audit a file named **dd-mm-yy-audit.xlsx** will appear in your script folder with the results of the audit.
