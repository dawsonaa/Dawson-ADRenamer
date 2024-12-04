# Dawson's ADRenamer

## Support
For assistance, please contact Dawson Adams (dawsonaa@ksu.edu).

All Campuses Device Naming Scheme KB: https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163

## Overview
This script is a comprehensive tool for renaming Active Directory (AD) computer objects. It includes functionalities to load and filter AD computer objects, rename them based on specified criteria, handle restarts, and log the operations performed. Additionally, it manages the script and log files' organization, ensuring they are stored in a consistent folder structure.

## Prerequisites
- Ensure you have the necessary permissions to run PowerShell scripts and interact with Active Directory.
- The Installer requires administrative privileges to create shortcuts in the Start Menu and on the Public Desktop.

## Installation
1. **Download DawsonADRenamerInstaller.exe**: Download the installer and run it (Will ask for admin).
2. **Parent Folder**: Script is located in C:\Users\CurrentUser\Dawson's ADRenamer\Dawson's ADRenamer.ps1
3. **Shortcuts**: Shortcuts gets added to C:\Users\CurrentUser\Desktop and C:\Users\CurrentUser\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer

## Running the Script
   - Windows search for "Dawson's ADRenamer" and open the shorcut. This shorcut uses Set-ExecutionPolicy to get around script warnings.
   - Alternatively you can right-click on the script file and select "Run with PowerShell". If you encounter permission issues, right-click the script file, select "Run as Administrator", and then run the script again.
   - The script will check if it is running with elevated privileges and restart itself with elevated privileges if necessary.

## Features

### Loading and Filtering
- Load and filter AD computer objects based on the last logon date (LoadAndFilterComputers function).

### Display and Selection
- Display a list of computers for selection (computerCheckedListBox and selectedComputersListBox).
- Allow users to specify new names for selected computers (UpdateAllListBoxes function and input text boxes).
- Validate new names against specified naming conventions (UpdateAllListBoxes function).
- Ensure new computer names are unique using a hash set (UpdateAllListBoxes function).

### Renaming and Logging
- Rename AD computer objects and log successful and failed renames (ApplyRenameButton click event handler).
- Check if computers are online before renaming (ApplyRenameButton click event handler).
- Check if users are logged on and handle restarts (ApplyRenameButton click event handler).
- Display progress and results of the renaming process (ApplyRenameButton click event handler).
- Track and display total time taken for operations (ApplyRenameButton click event handler).
- Display and handle invalid rename guidelines (ApplyRenameButton click event handler).

### Search and Filter
- Implement search functionality for filtering computers (searchBox).
- Select and set Organizational Unit (OU) for filtering (Select-OU function and refreshButton).
- Populate a TreeView with Organizational Units from Active Directory (Select-OU function).
- Handle OU selection and expand nodes to show child OUs (Select-OU function).

### Bulk Selection and Removal
- Implement Ctrl+A to mass select up to 500 items in the main list box (KeyDown event handler for computerCheckedListBox).
- Remove selected devices from the list (context menu item for "Remove selected device(s)").
- Remove all devices from the list (context menu item for "Remove all device(s)").

### Operation Summaries
- Display rename and restart operation summaries (ApplyRenameButton click event handler).
- Display the list of logged on users (ApplyRenameButton click event handler).

### Credentials
- Ensure credentials are retained across login attempts (login section at the beginning of the script).

### Custom Scroll Event
- Implement a custom scroll event to keep list box's top items index synced (CustomListBox class and Scroll event handlers).
- Synchronize color panels with their respective items and changelist groups to ensure visual consistency when scrolling (CustomListBox class and Scroll event handlers).

### Email Drafts
- Automatically generate email drafts for users of logged on devices (Show-EmailDrafts function).
- Allow users to specify a support link that is included in email drafts (supportLinkTextBox and Update-OutlookWebDraft function).
- Create an email draft for each selected device with the specified support link (Update-OutlookWebDraft function).
- Remove selected items from the list using right-click context menu (Show-EmailDrafts function with context menu).
- Display a form for users to select devices and create email drafts (Show-EmailDrafts function with threading).

### Logging and Export
- Create RESULTS and LOGS folders if they do not exist in the script directory (ApplyRenameButton click event handler).
- Export CSV file with renaming results to the RESULTS folder (ApplyRenameButton click event handler).
- Export log content to a .txt file in the LOGS folder (ApplyRenameButton click event handler).
- Include timestamp in the CSV file name for better organization (ApplyRenameButton click event handler).

### Power Automate Integration
- Convert CSV and log files to Base64 format (ApplyRenameButton click event handler).
- Send HTTP POST request to trigger a Power Automate flow to upload the CSV and log files to SharePoint (ApplyRenameButton click event handler).

### Department to String Conversion
- Convert department names to strings based on OU location, with truncation logic for specific naming patterns (Get-DepartmentString function).

### Select All Functionality
- Support Ctrl+A for select all functionality in text boxes (KeyDown event handler for text boxes).

## Log Management
- **Log Files Location**: Logs are stored in a folder named "Dawson's ADRenamer Logs" within the script's parent folder.
- **Log File Naming**: Logs are named with a timestamp for easy identification, e.g., "ADRenamer_Results_YY-MM-DD_HH-MMtt.csv".

## Troubleshooting
- **Permissions Issues**: Ensure you run the script with administrative privileges if you encounter any permission-related errors.
