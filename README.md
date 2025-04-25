# <img src="icon.ico" alt="Icon" height="32" style="vertical-align: middle;"> <span style="vertical-align: middle;"> Dawson's ADRenamer </span>  <img src="icon.ico" alt="Icon" height="32" style="vertical-align: middle;">

## By Dawson Adams (dawsonaa@ksu.edu) at Kansas State university </span>

## Overview
Dawson's ADRenamer is a comprehensive tool for renaming Active Directory (AD) computer objects. It includes functionalities to load and filter AD computer objects, rename them based on specified criteria, handle restarts, and log the operations performed. Additionally, it manages the script and log files' organization, ensuring they are stored in a consistent folder structure.

## Main GUI
<img src="images/MainGUI.png" alt="Main GUI Screenshot">

## Terminal
<img src="images/Terminal.png" alt="Terminal Screenshot">

## Select OU
<img src="images/SelectOU.png" alt="Select OU Screenshot">

## Email Generation GUI
<img src="images/EmailGUI.png" alt="Email Generation GUI Screenshot"> 

<img src="images/ExampleEmail.png" alt="Example Email Screenshot" style="width: 577px;">

## Logs Viewer
<img src="images/LogsViewer.png" alt="Logs Viewer Screenshot">

## Results CSV
<img src="images/ResultsCSV.png" alt="Results CSV Screenshot">

## Installation
1. **Download the Installer**: Download `DawsonADRenamerInstaller.exe` and run it (requires admin privileges).
2. **Script Location**: The script is located in `C:\Users\CurrentUser\Dawson's ADRenamer\Dawson's ADRenamer.ps1`.
3. **Shortcuts**: Shortcuts are added to `C:\Users\CurrentUser\Desktop` and `C:\Users\CurrentUser\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer`.

## Running the Script
- **Via Shortcut**: Search for "Dawson's ADRenamer" in Windows and open the shortcut. This shortcut uses `Set-ExecutionPolicy` to bypass script warnings.
- **Via PowerShell**: Right-click the script file and select "Run with PowerShell". If permission issues occur, right-click the script file, select "Run as Administrator", and then run the script again.
- The script will check if it is running with elevated privileges and restart itself with elevated privileges if necessary.

## Features

### Loading and Filtering
- Load and filter AD computer objects based on the last logon date.

### Display and Selection
- Display a list of computers for selection.
- Allow users to specify new names for selected computers.
- Validate new names against specified naming conventions.
- Ensure new computer names are unique using a hash set.

### Renaming and Logging
- Rename AD computer objects and log successful and failed renames.
- Check if computers are online before renaming.
- Check if users are logged on and handle restarts.
- Display progress and results of the renaming process.
- Track and display total time taken for operations.
- Display and handle invalid rename guidelines.

### Search and Filter
- Implement search functionality for filtering computers.
- Select and set Organizational Unit (OU) for filtering.
- Populate a TreeView with Organizational Units from Active Directory.
- Handle OU selection and expand nodes to show child OUs.

### Bulk Selection and Removal
- Remove selected devices from the list.
- Remove all devices from the list.

### Operation Summaries
- Display rename and restart operation summaries.
- Display the list of logged on users.

### Email Drafts
- Automatically generate email drafts for users of logged on devices.
- Allow users to specify a support link that is included in email drafts.
- Create an email draft for each selected device with the specified support link.
- Remove selected items from the list using the right-click context menu.
- Display a form for users to select devices and create email drafts.

### Logging and Export
- Create `RESULTS` and `LOGS` folders if they do not exist in the script directory.
- Export CSV file with renaming results to the `RESULTS` folder.
- Export log content to a .txt file in the `LOGS` folder.
- Include a timestamp in the CSV file name for better organization.

### Power Automate Integration
- Convert CSV and log files to Base64 format.
- Send HTTP POST request to trigger a Power Automate flow to upload the CSV and log files to SharePoint.

### Department to String Conversion
- Convert department names to strings based on OU location, with truncation logic for specific naming patterns.

### Log Management
- **Log Files Location**: Logs are stored in a folder named `Dawson's ADRenamer Logs` within the script's parent folder.
- **Log File Naming**: Logs are named with a timestamp for easy identification, e.g., `ADRenamer_Results_YY-MM-DD_HH-MMtt.csv`.

## Support
For assistance, please contact Dawson Adams at [dawsonaa@ksu.edu](mailto:dawsonaa@ksu.edu).

All Campuses Device Naming Scheme KB: [https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163](https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163)

## Troubleshooting
- **Permissions Issues**: Ensure you run the script with administrative privileges if you encounter any permission-related errors.