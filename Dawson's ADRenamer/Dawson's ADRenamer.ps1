<#
# Dawson's AD Computer Renamer 2.5.21
# Author: Dawson Adams (dawsonaa@ksu.edu)
# Version: 2.5.21
# Date: 5/21/2024
# This script provides a comprehensive tool for renaming Active Directory (AD) computer objects. It includes functionalities to:

## Loading and Filtering

- Load and filter AD computer objects based on the last logon date (LoadAndFilterComputers function).

## Display and Selection

- Display a list of computers for selection (computerCheckedListBox and selectedComputersListBox).
- Allow users to specify new names for selected computers (UpdateAllListBoxes function and input text boxes).
- Validate new names against specified naming conventions (UpdateAllListBoxes function).
- Ensure new computer names are unique using a hash set (UpdateAllListBoxes function).

## Renaming and Logging

- Rename AD computer objects and log successful and failed renames (ApplyRenameButton click event handler).
- Check if computers are online before renaming (ApplyRenameButton click event handler).
- Check if users are logged on and handle restarts (ApplyRenameButton click event handler).
- Display progress and results of the renaming process (ApplyRenameButton click event handler).
- Track and display total time taken for operations (ApplyRenameButton click event handler).
- Display and handle invalid rename guidelines (ApplyRenameButton click event handler).

## Custom Names

- Add custom names for selected computers (context menu item for "Set/Change Custom rename").
- Remove custom names for selected computers (context menu item for "Remove Custom rename").

## Search and Filter

- Implement search functionality for filtering computers (searchBox).
- Select and set Organizational Unit (OU) for filtering (Select-OU function and refreshButton).
- Populate a TreeView with Organizational Units from Active Directory (Select-OU function).
- Handle OU selection and expand nodes to show child OUs (Select-OU function).

## Synchronization and Context Menus

- Synchronize scrolling between list boxes (Scroll event handlers for list boxes).
- Enable context menu items based on selected and available items (contextMenu Opening event handler).

## Bulk Selection and Removal

- Implement Ctrl+A to mass select up to 500 items in the main list box (KeyDown event handler for computerCheckedListBox).
- Remove selected devices from the list (context menu item for "Remove selected device(s)").
- Remove all devices from the list (context menu item for "Remove all device(s)").

## Find and Replace

- Find and replace strings within selected computers (context menu item for "Find and Replace").

## Operation Summaries

- Display rename and restart operation summaries (ApplyRenameButton click event handler).
- Display the list of logged on users (ApplyRenameButton click event handler).

## Credentials

- Ensure credentials are retained across login attempts (login section at the beginning of the script).

## Custom Scroll Event

- Implement a custom scroll event to keep list box's top items index synced (CustomListBox class and Scroll event handlers).

## Email Drafts

- Automatically generate email drafts for users of logged on devices (Show-EmailDrafts function).
- Allow users to specify a support link that is included in email drafts (supportLinkTextBox and Update-OutlookWebDraft function).
- Create an email draft for each selected device with the specified support link (Update-OutlookWebDraft function).
- Remove selected items from the list using right-click context menu (Show-EmailDrafts function with context menu).
- Display a form for users to select devices and create email drafts (Show-EmailDrafts function with threading).

## Logging and Export

- Create `RESULTS` and `LOGS` folders if they do not exist in the script directory (ApplyRenameButton click event handler).
- Export CSV file with renaming results to the `RESULTS` folder (ApplyRenameButton click event handler).
- Export log content to a .txt file in the `LOGS` folder (ApplyRenameButton click event handler).
- Include timestamp in the CSV file name for better organization (ApplyRenameButton click event handler).

## Script Relocation

- Check if the current script's parent folder is named as the target folder.
- If not, create the target parent folder.
- Define the new script path in the target parent folder.
- Copy the current script to the new folder.
- Define the logs folder path and the new logs folder path in the target parent folder.
- Move the logs folder to the new location if it exists.
- Schedule deletion of the current script after copying.

## Power Automate Integration

- Convert CSV and log files to Base64 format (ApplyRenameButton click event handler).
- Send HTTP POST request to trigger a Power Automate flow to upload the CSV and log files to SharePoint (ApplyRenameButton click event handler).

# All Campuses Device Naming Scheme KB: https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163
#>

# IMPORTANT test3
# $false will run the applicatiion with dummy devices and will not connect to AD or ask for cred's
# $true will run the application with imported AD modules and will request credentials to be used for actual AD device name manipulation
$online = $false

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not $online) {
    # Dummy data for computers # OFFLINE
    # Dummy data for computers
    $dummyComputers = @(
        @{ Name = 'HL-CS-LOAN-1'; LastLogonDate = (Get-Date).AddDays(-10) },
        @{ Name = 'HL-CS-LOAN-2'; LastLogonDate = (Get-Date).AddDays(-20) },
        @{ Name = 'HL-CS-LOAN-3'; LastLogonDate = (Get-Date).AddDays(-30) },
        @{ Name = 'HL-CS-LOAN-4'; LastLogonDate = (Get-Date).AddDays(-10) },
        @{ Name = 'HL-CS-LOAN-5'; LastLogonDate = (Get-Date).AddDays(-20) },
        @{ Name = 'HL-CS-LOAN-6'; LastLogonDate = (Get-Date).AddDays(-30) },
        @{ Name = 'HL-CS-LOAN-7'; LastLogonDate = (Get-Date).AddDays(-10) },
        @{ Name = 'HL-CS-LOAN-8'; LastLogonDate = (Get-Date).AddDays(-20) },
        @{ Name = 'HL-CS-LOAN-9'; LastLogonDate = (Get-Date).AddDays(-30) },
        @{ Name = 'HL-CS-LOAN-10'; LastLogonDate = (Get-Date).AddDays(-200) }  # This one should be filtered out
    )

    # Dummy data for OUs # OFFLINE
    $dummyOUs = @(
        "OU=Sales,OU=DEPT,DC=users,DC=campus",
        "OU=IT,OU=DEPT,DC=users,DC=campus",
        "OU=HR,OU=DEPT,DC=users,DC=campus"
    )

    # Arrays of possible outcomes # OFFLINE
    $onlineStatuses = @("Online", "Online", "Online", "Offline")
    $restartOutcomes = @("Success", "Success", "Success", "Success", "Fail")
    $loggedOnUserss = @("User1", "User2", "User3", "User4", "User5", "User6", "User7", "none", "none", "none")
    # Arrays of possible rename outcomes
    $renameOutcomes = @(
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Success"; ReturnValue = 0 },
        @{ Result = "Fail"; ReturnValue = 1 },
        @{ Result = "Fail"; ReturnValue = 2 },
        @{ Result = "Fail"; ReturnValue = 3 }
    )

    # Function to get a random outcome from an array # OFFLINE
    function Get-RandomOutcome {
        param (
            [Parameter(Mandatory = $true)]
            [array]$outcomes
        )
        return $outcomes | Get-Random
    }

    $username = "dawsonaa" # OFFLINE
}
else {
    Import-Module ActiveDirectory
}

# Adds custom scroll event handling to keep listbox's top items index synced
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class CustomListBox : ListBox
{
    private const int WM_VSCROLL = 0x115;
    private const int WM_MOUSEWHEEL = 0x20A;

    public event ScrollEventHandler Scroll;

    public delegate void ScrollEventHandler(object sender, ScrollEventArgs e);

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);

        if (m.Msg == WM_VSCROLL || m.Msg == WM_MOUSEWHEEL)
        {
            OnScroll();
        }
    }

    protected void OnScroll()
    {
        if (Scroll != null)
        {
            Scroll(this, new ScrollEventArgs(ScrollEventType.EndScroll, 0));
        }
    }
}
"@ -Language CSharp -ReferencedAssemblies System.Windows.Forms

# Set text for version labels
$Version = "2.5.21"

# Set Character limits
$part0Max = 5 #Part 0 max characters (Set between 1-15, if above 15 it will still be 15)
$part1Max = 6 #Part 1 max characters (Set between 1-15, if above 15 it will still be 15)

# Below variables can be changed to fit deparment needs
$part0Name = "Part-0" # Related to original variables with part0 in the name
$part1Name = "Part-1" # Related to original variables with part1 in the name
$part2Name = "Part-2" # Related to original variables with part2 in the name
$part3Name = "Part-3" # Related to original variables with part3 in the name

# Initialize to false so while loop can exit when a successful login attempt is made
$connectionSuccess = $false
$errorMessage = ""  # Initialize the error message to an empty string
# $username = ""  # Initialize the username variable # ONLINE

# Function to create a README file
function Add-Readme {
    param (
        [string]$readmePath
    )

    $readmeContent = @"
# Dawson's ADRenamer 2.5.21

## Overview
This script is a comprehensive tool for renaming Active Directory (AD) computer objects. It includes functionalities to load and filter AD computer objects, rename them based on specified criteria, handle restarts, and log the operations performed. Additionally, it manages the script and log files' organization, ensuring they are stored in a consistent folder structure.

## Prerequisites
- Ensure you have the necessary permissions to run PowerShell scripts and interact with Active Directory.
- The script requires administrative privileges to create shortcuts in the Start Menu and on the Public Desktop.

## Installation
1. **Download the Script**: Save the script file to your desired location.
2. **Parent Folder**: Ensure the script is placed in a parent folder named \`Dawson's ADRenamer 2.5.21\`. If the script is not in a folder with this name, it will automatically create one and move itself and its logs to the newly created folder.

## Running the Script
1. **Run as Administrator**:
   - Right-click on the script file and select \`Run with PowerShell\`. If you encounter permission issues, right-click the script file, select \`Run as Administrator\`, and then run the script again.
   - The script will check if it is running with elevated privileges and restart itself with elevated privileges if necessary.

2. **Initial Setup**:
   - On the first run, if the script is not already in the correct parent folder (\`Dawson's ADRenamer 2.5.21\`), it will create this folder, move itself and its logs there, and then prompt you to rerun the script from the new location.

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

### Custom Names
- Add custom names for selected computers (context menu item for "Set/Change Custom rename").
- Remove custom names for selected computers (context menu item for "Remove Custom rename").

### Search and Filter
- Implement search functionality for filtering computers (searchBox).
- Select and set Organizational Unit (OU) for filtering (Select-OU function and refreshButton).
- Populate a TreeView with Organizational Units from Active Directory (Select-OU function).
- Handle OU selection and expand nodes to show child OUs (Select-OU function).

### Synchronization and Context Menus
- Synchronize scrolling between list boxes (Scroll event handlers for list boxes).
- Enable context menu items based on selected and available items (contextMenu Opening event handler).

### Bulk Selection and Removal
- Implement Ctrl+A to mass select up to 500 items in the main list box (KeyDown event handler for computerCheckedListBox).
- Remove selected devices from the list (context menu item for "Remove selected device(s)").
- Remove all devices from the list (context menu item for "Remove all device(s)").

### Find and Replace
- Find and replace strings within selected computers (context menu item for "Find and Replace").

### Operation Summaries
- Display rename and restart operation summaries (ApplyRenameButton click event handler).
- Display the list of logged on users (ApplyRenameButton click event handler).

### Credentials
- Ensure credentials are retained across login attempts (login section at the beginning of the script).

### Custom Scroll Event
- Implement a custom scroll event to keep list box's top items index synced (CustomListBox class and Scroll event handlers).

### Email Drafts
- Automatically generate email drafts for users of logged on devices (Show-EmailDrafts function).
- Allow users to specify a support link that is included in email drafts (supportLinkTextBox and Update-OutlookWebDraft function).
- Create an email draft for each selected device with the specified support link (Update-OutlookWebDraft function).
- Remove selected items from the list using right-click context menu (Show-EmailDrafts function with context menu).
- Display a form for users to select devices and create email drafts (Show-EmailDrafts function with threading).

### Logging and Export
- Create `RESULTS` and `LOGS` folders if they do not exist in the script directory (ApplyRenameButton click event handler).
- Export CSV file with renaming results to the `RESULTS` folder (ApplyRenameButton click event handler).
- Export log content to a .txt file in the `LOGS` folder (ApplyRenameButton click event handler).
- Include timestamp in the CSV file name for better organization (ApplyRenameButton click event handler).

### Script Relocation
- Check if the current script's parent folder is named as the target folder.
- If not, create the target parent folder.
- Define the new script path in the target parent folder.
- Copy the current script to the new folder.
- Define the logs folder path and the new logs folder path in the target parent folder.
- Move the logs folder to the new location if it exists.
- Schedule deletion of the current script after copying.

### Power Automate Integration
- Convert CSV and log files to Base64 format (ApplyRenameButton click event handler).
- Send HTTP POST request to trigger a Power Automate flow to upload the CSV and log files to SharePoint (ApplyRenameButton click event handler).

## Log Management
- **Log Files Location**: Logs are stored in a folder named \`Dawson's ADRenamer Logs\` within the script's parent folder.
- **Log File Naming**: Logs are named with a timestamp for easy identification, e.g., \`ADRenamer_Results_YY-MM-DD_HH-MMtt.csv\`.

## Troubleshooting
- **Permissions Issues**: Ensure you run the script with administrative privileges if you encounter any permission-related errors.

## Support
For further assistance, please contact Dawson Adams (dawsonaa@ksu.edu).

# All Campuses Device Naming Scheme KB: https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163
"@

    # Write the README content to the specified path
    $readmeContent | Out-File -FilePath $readmePath -Encoding utf8
}

# Function to ensure README file exists
function EnsureReadmeExists {
    $currentScriptDir = Split-Path -Parent $PSCommandPath
    $readmePath = Join-Path -Path $currentScriptDir -ChildPath "README.txt"
    
    if (-not (Test-Path -Path $readmePath)) {
        Add-Readme -readmePath $readmePath
        Write-Host "Created README at: $readmePath"
    }
}

# Function to create a duplicate of the script in a new folder and move logs
function DuplicateScriptAndMoveLogs {
    # Get the current script's full path
    $currentScriptPath = $PSCommandPath
    $currentScriptDir = Split-Path -Parent $currentScriptPath
    $currentScriptName = Split-Path -Leaf $currentScriptPath

    # Define the target parent folder name
    $targetParentFolderName = "Dawson's ADRenamer"
    $targetParentFolderPath = Join-Path -Path $currentScriptDir -ChildPath $targetParentFolderName

    # Check if the current script's parent folder is named as target
    if ((Split-Path -Leaf $currentScriptDir) -ne $targetParentFolderName) {
        # Create the target parent folder if it doesn't exist
        if (-not (Test-Path -Path $targetParentFolderPath)) {
            New-Item -Path $targetParentFolderPath -ItemType Directory | Out-Null
        }

        # Define the new script path
        $newScriptPath = Join-Path -Path $targetParentFolderPath -ChildPath $currentScriptName

        # Copy the current script to the new folder
        Copy-Item -Path $currentScriptPath -Destination $newScriptPath -Force

        # Define the logs folder path
        $logsFolderPath = Join-Path -Path $currentScriptDir -ChildPath "Dawson's ADRenamer Logs"
        $newLogsFolderPath = Join-Path -Path $targetParentFolderPath -ChildPath "Dawson's ADRenamer Logs"

        # Move the logs folder to the new location
        if (Test-Path -Path $logsFolderPath) {
            Move-Item -Path $logsFolderPath -Destination $newLogsFolderPath -Force
        }

        # Schedule deletion of the current script after copying
        Start-Job -ScriptBlock {
            Start-Sleep -Seconds 5
            Remove-Item -Path $using:currentScriptPath -Force
        } | Out-Null

        # Create README file in the new folder
        $readmePath = Join-Path -Path $targetParentFolderPath -ChildPath "README.txt"
        Add-Readme -readmePath $readmePath
        Write-Host "Created README at: $readmePath"

        # Display message to indicate completion of the operation
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("The script has been moved to the new location:`n$newScriptPath`nPlease run the script from the new location.", "Operation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        # Exit the current script
        exit
    }
}

# Call the function at the beginning of the script
DuplicateScriptAndMoveLogs

EnsureReadmeExists

if ($online) {
    # Present initial login window # ONLINE
    while (-not $connectionSuccess) {
        if ($errorMessage) {
            Write-Host $errorMessage -ForegroundColor Red
        }

        if ($username -ne "") {
            $cred = Get-Credential -Message "Please enter AD admin credentials" -UserName $username
        }
        else {
            $cred = Get-Credential -Message "Please enter AD admin credentials"
        }

        if (-not $cred) {
            exit
        }

        # Store the username to retain it across attempts
        $username = $cred.UserName

        try {
            # Test the AD connection with the provided credentials.
            $null = Get-ADDomain -Credential $cred

            # Set to Exit While loop
            $connectionSuccess = $true
        }
        catch {
            $errorMessage = "Invalid credentials or insufficient permissions. Please try again."
        }
    }
}
Write-Host "Sufficient Credentials Provided. Logged on as $username." -ForegroundColor Green

# Initialize a new hash set to store unique strings.
# This hash set will be used to ensure that new computer names are unique.
$hashSet = [System.Collections.Generic.HashSet[string]]::new()

# Updates all list boxes with computer names and their corresponding new names based on user-defined rules and custom names.
# This function performs the following steps:
# 1. Checks if any relevant checkboxes are checked.
# 2. Clears existing items and lists.
# 3. Iterates through the checked computer names and updates the selected computers list.
# 4. Checks for custom names and applies them if found.
# 5. Splits the computer names into parts and applies user-defined replacements for parts.
# 6. Calculates the remaining length for parts based on total length limits and applies truncation if necessary.
# 7. Swaps part0 and part1 if the swap checkbox is checked.
# 8. Constructs new names and validates them based on length and uniqueness constraints.
# 9. Adds valid and invalid names to the respective lists and updates the new names list box with appropriate labels.
# 10. Enables or disables the ApplyRenameButton based on the presence of valid names and checked conditions.
class Change {
    [string[]]$ComputerNames
    [string]$Part0
    [string]$Part1
    [string]$Part2
    [string]$Part3

    Change([string[]]$computerNames, [string]$part0, [string]$part1, [string]$part2, [string]$part3) {
        $this.ComputerNames = $computerNames
        $this.Part0 = $part0
        $this.Part1 = $part1
        $this.Part2 = $part2
        $this.Part3 = $part3
    }
}

$script:changesList = New-Object System.Collections.ArrayList
$script:newNamesList = @()

function UpdateAllListBoxes {
    Write-Host "UpdateAllListBoxes"
    Write-Host ""

    # Check if any relevant checkboxes are checked
    $anyCheckboxChecked = (-not $part1Input.ReadOnly) -or (-not $part2Input.ReadOnly) -or (-not $part3Input.ReadOnly)

    # Clear existing items and lists except for the new names list
    $hashSet.Clear()
    $script:newNamesListBox.Items.Clear()
    $script:validNamesList = @()
    $script:invalidNamesList = @()
    $newChangesList = New-Object System.Collections.ArrayList

    Write-Host "Checked items: $($script:checkedItems.Keys -join ', ')"

    # Process selected checked items
    foreach ($computerName in $script:selectedCheckedItems.Keys) {
        Write-Host "Processing computer: $computerName"
        
        $parts = $computerName -split '-'
        $part0 = $parts[0]
        $part1 = $parts[1]
        $part2 = if ($parts.Count -ge 3) { $parts[2] } else { $null }
        $part3 = if ($parts.Count -ge 4) { $parts[3..($parts.Count - 1)] -join '-' } else { $null }

        Write-Host "Initial parts: Part0: $part0, Part1: $part1, Part2: $part2, Part3: $part3"

        $part0InputValue = if (-not $part0Input.ReadOnly) { $part0Input.Text } else { $null }
        $part1InputValue = if (-not $part1Input.ReadOnly) { $part1Input.Text } else { $null }
        $part2InputValue = if (-not $part2Input.ReadOnly) { $part2Input.Text } else { $null }
        $part3InputValue = if (-not $part3Input.ReadOnly) { $part3Input.Text } else { $null }

        Write-Host "Input values: Part0: $part0InputValue, Part1: $part1InputValue, Part2: $part2InputValue, Part3: $part3InputValue"

        if ($part0InputValue) { $part0 = $part0InputValue }
        if ($part1InputValue) { $part1 = $part1InputValue }
        if ($part2InputValue) {
            $totalLengthForpart2 = 15 - ($part0.Length + $part1.Length + ($parts.Count - 1))  # parts.count -1 is for hyphens
            if ($totalLengthForpart2 -gt 0) {
                $part2 = $part2InputValue.Substring(0, [Math]::Min($part2InputValue.Length, $totalLengthForpart2))
            }
            else {
                $part2 = ""
            }
        }
        if ($part3InputValue) {
            $totalLengthForpart3 = 15 - ($part0.Length + $part1.Length + $part2.Length + ($parts.Count - 1))
            if ($totalLengthForpart3 -gt 0) {
                $part3 = $part3InputValue.Substring(0, [Math]::Min($part3InputValue.Length, $totalLengthForpart3))
            }
            else {
                $part3 = ""
            }
        }

        Write-Host "Updated parts: Part0: $part0, Part1: $part1, Part2: $part2, Part3: $part3"

        if ($part3) {
            $newName = "$part0-$part1-$part2-$part3"
        }
        elseif ($part2) {
            $newName = "$part0-$part1-$part2"
        }
        else {
            $newName = "$part0-$part1"
        }

        Write-Host "New name: $newName"

        if ($hashSet.Add($newName)) {
            $script:validNamesList += "$computerName -> $newName"
            if (-not ($script:newNamesList | Where-Object { $_.ComputerName -eq $computerName })) {
                $script:newNamesList += @{"ComputerName" = $computerName; "NewName" = $newName; "Custom" = $false }
            }

            # Check if an existing change matches
            $existingChange = $null
            foreach ($change in $script:changesList) {
                Write-Host "Comparing changes for $computerName..."
                Write-Host "Part0: '$($change.Part0)' vs '$part0InputValue'"
                Write-Host "Part1: '$($change.Part1)' vs '$part1InputValue'"
                Write-Host "Part2: '$($change.Part2)' vs '$part2InputValue'"
                Write-Host "Part3: '$($change.Part3)' vs '$part3InputValue'"

                $part0Comparison = ($change.Part0 -eq $part0InputValue -or ([string]::IsNullOrEmpty($change.Part0) -and [string]::IsNullOrEmpty($part0InputValue)))
                $part1Comparison = ($change.Part1 -eq $part1InputValue -or ([string]::IsNullOrEmpty($change.Part1) -and [string]::IsNullOrEmpty($part1InputValue)))
                $part2Comparison = ($change.Part2 -eq $part2InputValue -or ([string]::IsNullOrEmpty($change.Part2) -and [string]::IsNullOrEmpty($part2InputValue)))
                $part3Comparison = ($change.Part3 -eq $part3InputValue -or ([string]::IsNullOrEmpty($change.Part3) -and [string]::IsNullOrEmpty($part3InputValue)))

                Write-Host "Part0 Comparison: $part0Comparison"
                Write-Host "Part1 Comparison: $part1Comparison"
                Write-Host "Part2 Comparison: $part2Comparison"
                Write-Host "Part3 Comparison: $part3Comparison"

                if ($part0Comparison -and $part1Comparison -and $part2Comparison -and $part3Comparison) {
                    Write-Host "Found matching change for parts: Part0: $($change.Part0), Part1: $($change.Part1), Part2: $($change.Part2), Part3: $($change.Part3)"
                    $existingChange = $change
                    break
                }
            }

            # Remove the computer name from any previous change entries if they exist
            foreach ($change in $script:changesList) {
                Write-Host "Checking $($change.ComputerNames) for $computerName removal..."
                Write-Host "does it contain: " ($change.ComputerNames -contains $computerName)
                Write-Host "does change equal existing: " ($change -eq $existingChange)
                if ($change -ne $existingChange -and $change.ComputerNames -contains $computerName) {
                    Write-Host "Removing $computerName from previous change entry: Part0: $($change.Part0), Part1: $($change.Part1), Part2: $($change.Part2), Part3: $($change.Part3)"
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $computerName }
                }
            }

            if ($existingChange) {
                Write-Host "Merging with existing change for parts: Part0: $($existingChange.Part0), Part1: $($existingChange.Part1), Part2: $($existingChange.Part2), Part3: $($existingChange.Part3)"
                $existingChange.ComputerNames += $computerName
            }
            else {
                Write-Host "Creating new change entry for parts: Part0: $part0InputValue, Part1: $part1InputValue, Part2: $part2InputValue, Part3: $part3InputValue"
                $newChange = [Change]::new(@($computerName), $part0InputValue, $part1InputValue, $part2InputValue, $part3InputValue)
                $script:changesList.Add($newChange) | Out-Null
            }

        
        
        
        
        
        }
        else {
            $script:invalidNamesList += $computerName
        }
    }

    # Populate the newNamesListBox with items from the newNamesList
    foreach ($entry in $script:newNamesList) {
        $customSuffix = if ($entry.Custom) { " - Custom" } else { "" }
        $script:newNamesListBox.Items.Add($entry.NewName + $customSuffix) | Out-Null
    }

    # Enable or disable the ApplyRenameButton based on valid names count and checkbox states
    $applyRenameButton.Enabled = ($anyCheckboxChecked -and ($script:validNamesList.Count -gt 0)) -or ($script:customNamesList.Count -gt 0)

    # Print the changesList for debugging
    Write-Host "`nChanges List:"
    foreach ($change in $script:changesList) {
        Write-Host "Change Parts: Part0: $($change.Part0), Part1: $($change.Part1), Part2: $($change.Part2), Part3: $($change.Part3)"
        Write-Host "ComputerNames: $($change.ComputerNames -join ', ')"
    }
}


# Function for setting individual custom names
function Show-InputBox {
    param (
        [string]$message,
        [string]$title,
        [string]$defaultText
    )

    # Create the form
    $inputBoxForm = New-Object System.Windows.Forms.Form
    $inputBoxForm.Text = $title
    $inputBoxForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $inputBoxForm.Width = 400
    $inputBoxForm.Height = 150
    $inputBoxForm.MaximizeBox = $false
    $inputBoxForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create the label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 10
    $inputBoxForm.Controls.Add($label)

    # Create the text box
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Width = 360
    $textBox.Top = 50
    $textBox.Left = 10
    $textBox.MaxLength = 15
    $textBox.Text = $defaultText
    $textBox.add_KeyDown({
            param($s, $e)
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
                $textBox.SelectAll()
                $e.SuppressKeyPress = $true
                $e.Handled = $true
            }
        })
    $inputBoxForm.Controls.Add($textBox)

    # Create the OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Top = 80
    $okButton.Left = 220
    $okButton.Width = 75
    $okButton.Enabled = $false # Initially disabled to avoid renaming to the oldname
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputBoxForm.Controls.Add($okButton)

    # Create the Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Top = 80
    $cancelButton.Left = 300
    $cancelButton.Width = 75
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputBoxForm.Controls.Add($cancelButton)

    # Add OK and Cancel button to $form
    $inputBoxForm.AcceptButton = $okButton
    $inputBoxForm.CancelButton = $cancelButton

    # Add event handler for text changed event
    $textBox.Add_TextChanged({
            if ($textBox.Text -ne $defaultText -and $textBox.Text.Trim() -ne "") {
                $okButton.Enabled = $true
            }
            else {
                $okButton.Enabled = $false
            }
        })

    # Show the form
    $dialogResult = $inputBoxForm.ShowDialog()
    
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }
    else {
        return $null
    }
}

# Initialize the link for submitting a ticket
$defaultSupportTicketLink = "https://support.ksu.edu/TDClient/30/Portal/Requests/ServiceCatalog"

# Function to format usernames into email addresses
function ConvertTo-EmailAddress {
    param (
        [string]$username
    )

    # Remove "Users\" from the username
    $emailLocalPart = $username -replace "Users\\", ""
    # Append "@ksu.edu" to the local part
    $email = $emailLocalPart + "@ksu.edu"
    return $email
}

# Function to create an Outlook web draft email
function Update-OutlookWebDraft {
    param (
        [string]$oldName,
        [string]$newName,
        [string]$emailAddress,
        [string]$supportTicketLink
    )

    # Extract the username from the email address
    $username = ($emailAddress -split '@')[0]

    # Construct the email message
    $subject = "Action Required: Restart Your Device"
    $body = @"
Dear $username,

Your computer has been renamed from $oldName to $newName as part of a maintenance operation. To avoid device name syncing issues, please restart your device as soon as possible. If you face any issues, please contact IT support.

You can submit a ticket using the following link: $supportTicketLink

Best regards,
IT Support Team
"@

    # Construct the Outlook web URL for creating a draft
    $url = "https://outlook.office.com/mail/deeplink/compose?to=" + [System.Uri]::EscapeDataString($emailAddress) + "&subject=" + [System.Uri]::EscapeDataString($subject) + "&body=" + [System.Uri]::EscapeDataString($body)

    # Open the URL in the default browser
    Start-Process $url
    Write-Host "Draft email created for $emailAddress" -ForegroundColor Green
} # FIX - adds +'s sometimes. Had it do it on the first try and not on the second.

# Function to prompt user to create email drafts using three synchronized ListBox controls
function Show-EmailDrafts {
    param (
        [array]$loggedOnDevices
    )

    # Create a new form
    $emailForm = New-Object System.Windows.Forms.Form
    $emailForm.Text = "Devices to Create Email Drafts (Right Click to Remove)"
    $emailForm.Size = New-Object System.Drawing.Size(600, 420)
    $emailForm.MaximizeBox = $false
    $emailForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $emailForm.StartPosition = "CenterScreen"
    
    # Create labels for each ListBox
    $labelOldName = New-Object System.Windows.Forms.Label
    $labelOldName.Text = "Old Name"
    $labelOldName.Location = New-Object System.Drawing.Point(70, 15)
    $labelOldName.Size = New-Object System.Drawing.Size(180, 20)
    $emailForm.Controls.Add($labelOldName)

    $labelNewName = New-Object System.Windows.Forms.Label
    $labelNewName.Text = "New Name"
    $labelNewName.Location = New-Object System.Drawing.Point(260, 15)
    $labelNewName.Size = New-Object System.Drawing.Size(180, 20)
    $emailForm.Controls.Add($labelNewName)

    $labelUserName = New-Object System.Windows.Forms.Label
    $labelUserName.Text = "User Name"
    $labelUserName.Location = New-Object System.Drawing.Point(445, 15)
    $labelUserName.Size = New-Object System.Drawing.Size(180, 20)
    $emailForm.Controls.Add($labelUserName)

    # Create ListBox for OldName
    $listBoxOldName = [CustomListBox]::new()
    $listBoxOldName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxOldName.Location = New-Object System.Drawing.Point(10, 40)
    $listBoxOldName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    # Create ListBox for NewName
    $listBoxNewName = [CustomListBox]::new()
    $listBoxNewName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxNewName.Location = New-Object System.Drawing.Point(200, 40)
    $listBoxNewName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    # Create ListBox for UserName
    $listBoxUserName = [CustomListBox]::new()
    $listBoxUserName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxUserName.Location = New-Object System.Drawing.Point(390, 40)
    $listBoxUserName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    # Add devices to the ListBoxes
    foreach ($device in $loggedOnDevices) {
        $listBoxOldName.Items.Add($device.OldName)
        $listBoxNewName.Items.Add($device.NewName)
        $listBoxUserName.Items.Add($device.UserName)
    }

    $emailForm.Controls.Add($listBoxOldName)
    $emailForm.Controls.Add($listBoxNewName)
    $emailForm.Controls.Add($listBoxUserName)

    # Flag to prevent recursive selection change events
    $global:syncingSelection = $false

    # Sync ListBox selections
    $syncSelection = {
        param ($s, $e)
        if (-not $global:syncingSelection) {
            $global:syncingSelection = $true
            $selectedIndices = $s.SelectedIndices

            # Sync other list boxes
            if ($listBoxOldName -ne $s) {
                $listBoxOldName.ClearSelected()
                foreach ($index in $selectedIndices) {
                    $listBoxOldName.SetSelected($index, $true)
                }
            }
            if ($listBoxNewName -ne $s) {
                $listBoxNewName.ClearSelected()
                foreach ($index in $selectedIndices) {
                    $listBoxNewName.SetSelected($index, $true)
                }
            }
            if ($listBoxUserName -ne $s) {
                $listBoxUserName.ClearSelected()
                foreach ($index in $selectedIndices) {
                    $listBoxUserName.SetSelected($index, $true)
                }
            }
            $global:syncingSelection = $false
        }
    }

    $listBoxOldName.add_SelectedIndexChanged($syncSelection)
    $listBoxNewName.add_SelectedIndexChanged($syncSelection)
    $listBoxUserName.add_SelectedIndexChanged($syncSelection)

    # Sync ListBox scrolling
    $syncScroll = {
        param ($s, $e)
        if ($listBoxOldName.TopIndex -ne $s.TopIndex) {
            $listBoxOldName.TopIndex = $s.TopIndex
        }
        if ($listBoxNewName.TopIndex -ne $s.TopIndex) {
            $listBoxNewName.TopIndex = $s.TopIndex
        }
        if ($listBoxUserName.TopIndex -ne $s.TopIndex) {
            $listBoxUserName.TopIndex = $s.TopIndex
        }
    }

    $listBoxOldName.add_Scroll($syncScroll)
    $listBoxNewName.add_Scroll($syncScroll)
    $listBoxUserName.add_Scroll($syncScroll)

    # Create a context menu for the ListBoxes
    $contextMenu = New-Object System.Windows.Forms.ContextMenu
    $menuItemRemove = New-Object System.Windows.Forms.MenuItem "Remove"
    $menuItemRemove.Add_Click({
            $selectedIndices = $listBoxOldName.SelectedIndices
            for ($i = $selectedIndices.Count - 1; $i -ge 0; $i--) {
                $index = $selectedIndices[$i]
                $listBoxOldName.Items.RemoveAt($index)
                $listBoxNewName.Items.RemoveAt($index)
                $listBoxUserName.Items.RemoveAt($index)
            }
        })
    $contextMenu.MenuItems.Add($menuItemRemove)
    $listBoxOldName.ContextMenu = $contextMenu
    $listBoxNewName.ContextMenu = $contextMenu
    $listBoxUserName.ContextMenu = $contextMenu

    # Create a label for the support link
    $supportLinkLabel = New-Object System.Windows.Forms.Label
    $supportLinkLabel.Text = "Support Link:"
    $supportLinkLabel.Location = New-Object System.Drawing.Point(10, 340)
    $supportLinkLabel.Size = New-Object System.Drawing.Size(80, 20)
    $emailForm.Controls.Add($supportLinkLabel)

    # Create a textbox for the support link
    $supportLinkTextBox = New-Object System.Windows.Forms.TextBox
    $supportLinkTextBox.Text = $defaultSupportTicketLink
    $supportLinkTextBox.Location = New-Object System.Drawing.Point(90, 340) 
    $supportLinkTextBox.Size = New-Object System.Drawing.Size(340, 20)

    # Handle key down event to enable Ctrl+A functionality
    $supportLinkTextBox.add_KeyDown({
            param($s, $e)
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
                $supportLinkTextBox.SelectAll()
                $e.SuppressKeyPress = $true
                $e.Handled = $true
            }
        })
    $emailForm.Controls.Add($supportLinkTextBox)

    # Create a button to create drafts
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Create Email Drafts"
    $createButton.Size = New-Object System.Drawing.Size(120, 30)
    $createButton.Location = New-Object System.Drawing.Point(440, 340)
    $createButton.Add_Click({
            $supportTicketLink = $supportLinkTextBox.Text
            for ($i = 0; $i -lt $listBoxOldName.Items.Count; $i++) {
                $oldName = $listBoxOldName.Items[$i]
                $newName = $listBoxNewName.Items[$i]
                $userName = $listBoxUserName.Items[$i]
                $deviceInfo = $loggedOnDevices | Where-Object { $_.OldName -eq $oldName -and $_.NewName -eq $newName -and $_.UserName -eq $userName }
                if ($deviceInfo) {
                    $emailAddress = ConvertTo-EmailAddress $deviceInfo.UserName
                    Update-OutlookWebDraft -oldName $deviceInfo.OldName -newName $deviceInfo.NewName -emailAddress $emailAddress -supportTicketLink $supportTicketLink
                }
            }
            $emailForm.Close()
        })
    $emailForm.Controls.Add($createButton)

    # Show the form
    $emailForm.ShowDialog()
}

# Function to create dummy devices
function New-DummyDevices {
    param (
        [int]$count = 10  # Default to create 10 dummy devices
    )

    $dummyDevices = @()
    for ($i = 0; $i -lt $count; $i++) {
        $device = [PSCustomObject]@{
            OldName  = "HL-CS-old$i"
            NewName  = "HL-CS3-new$i"
            UserName = "USERS\user$i"
        }
        $dummyDevices += $device
    }
    return $dummyDevices
}
# Example usage: create 20 dummy devices
# $dummyDevices = New-DummyDevices -count 20
# Show-EmailDrafts -loggedOnDevices $dummyDevices | Out-Null

# Function to display a form with a TreeView control for selecting an Organizational Unit (OU)
function Select-OU {
    # Create and configure the form
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "Select Organizational Unit"
    $ouForm.Size = New-Object System.Drawing.Size(400, 600)
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $ouForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Add a handler for the FormClosing event to exit the script if the form is closed using the red X button
    $ouForm.Add_FormClosing({
            param($s, $e)
            if ($ouForm.DialogResult -eq [System.Windows.Forms.DialogResult]::None) {
                #Form closed with red X, Exit Script.
                [Environment]::Exit(0)
            }
        })

    # Create and configure the TreeView control
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Size = New-Object System.Drawing.Size(365, 500)
    $treeView.Location = New-Object System.Drawing.Point(10, 10)
    $treeView.Visible = $true

    # Add "OK"(selectedOU) button for OU selection
    $selectedOUButton = New-Object System.Windows.Forms.Button
    $selectedOUButton.Text = "No OU selected"
    $selectedOUButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $selectedOUButton.Enabled = $false
    $selectedOUButton.Size = New-Object System.Drawing.Size(75, 23)
    $selectedOUButton.Location = New-Object System.Drawing.Point(200, 520)
    $selectedOUButton.Add_Click({
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $ouForm.Close()
        })

    # Add "Cancel"(defaultOU) button for OU selection
    $defaultOUButton = New-Object System.Windows.Forms.Button
    $defaultOUButton.Text = "DC=users,DC=campus"
    $defaultOUButton.Size = New-Object System.Drawing.Size(75, 23)
    $defaultOUButton.Location = New-Object System.Drawing.Point(290, 520)
    $defaultOUButton.Add_Click({
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $ouForm.Close()
        })

    # Add the TreeView and buttons to the form
    $ouForm.Controls.Add($treeView)
    $ouForm.Controls.Add($selectedOUButton)
    $ouForm.Controls.Add($defaultOUButton)

    # Event handler for the NodeMouseClick event to handle node selection and expansion
    $treeView.Add_NodeMouseClick({
            param ($s, $e)
            $selectedNode = $e.Node
            if ($null -ne $selectedNode) {
                if ($selectedNode.Tag -match "DC=") {
                    $ouForm.Tag = [string]$selectedNode.Tag
                    $selectedOUButton.Text = $selectedNode.Tag
                    $selectedOUButton.Enabled = $true

                    # Expand the selected node to show its child OUs
                    $selectedNode.Nodes.Clear()
                    $selectedNode.Expand()

                    # Fetch and add child OUs to the expanded node
                    $childOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $selectedNode.Tag | Where-Object {
                        $_.DistinguishedName -match '^OU=[^,]+,OU=' + [regex]::Escape($selectedNode.Text) + ','
                    } | Sort-Object DistinguishedName

                    foreach ($childOU in $childOUs) {
                        $childNode = New-Object System.Windows.Forms.TreeNode
                        $childNode.Text = $childOU.Name
                        $childNode.Tag = $childOU.DistinguishedName
                        $selectedNode.Nodes.Add($childNode)
                    }
                }
                else {
                    $selectedOUButton.Text = "No OU selected"
                    $selectedOUButton.Enabled = $false
                }
            }
        })

    # Function to populate the TreeView with OUs from Active Directory
    function Update-TreeView {
        param ($treeView)

        # Fetch OUs directly under 'users.campus' and sort them by DistinguishedName
        Write-Host "Fetching OUs from AD..."
        $ous = Get-ADOrganizationalUnit -Filter * | Where-Object {
            $_.DistinguishedName -match '^OU=[^,]+,DC=users,DC=campus$'
        } | Sort-Object DistinguishedName

        # Build the tree structure by adding nodes for each OU
        $nodeHashTable = @{}
        foreach ($ou in $ous) {
            $node = New-Object System.Windows.Forms.TreeNode
            $node.Text = $ou.Name
            $node.Tag = $ou.DistinguishedName

            # Identify the parent DistinguishedName
            $parentDN = $ou.DistinguishedName -replace "^OU=[^,]+,", ""

            if ($parentDN -eq 'DC=users,DC=campus') {
                $treeView.Nodes.Add($node)
            } 

            $nodeHashTable[$ou.DistinguishedName] = $node
        }
    }

    # Populate the TreeView with initial OUs under 'users.campus'
    Update-TreeView -treeView $treeView | Out-Null

    # Show the form and wait for user input
    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:ouPath = $ouForm.Tag
    }
    else {
        $script:ouPath = 'DC=users,DC=campus'
    }
}

# Modified function to load and filter computers using dummy data # OFFLINE
function LoadAndFilterComputersOFFLINE {
    param (
        [System.Windows.Forms.CheckedListBox]$computerCheckedListBox
    )

    try {
        # Simulate OU selection
        $script:ouPath = 'OU=Workstations,OU=KSUL,OU=Dept,DC=users,DC=campus'
        Write-Host "Selected OU Path: $script:ouPath"  # Debug message

        # Disable the form while loading data to prevent user interaction
        $form.Enabled = $false
        Write-Host "Form disabled for loading..."
        Write-Host "Loading AD endpoints (offline mode with dummy data)..."

        # Initialize counters and timers for progress tracking
        $loadedCount = 0
        $deviceRefresh = 150
        $deviceTimer = 0

        # Define the cutoff date for filtering computers based on their last logon date
        $cutoffDate = (Get-Date).AddDays(-180)

        # Use dummy data instead of querying Active Directory
        $script:filteredComputers = $dummyComputers | Where-Object {
            $_.LastLogonDate -and
            [DateTime]::Parse($_.LastLogonDate) -ge $cutoffDate
        }

        $computerCount = $script:filteredComputers.Count
        $filteredOutCount = $dummyComputers.Count - $script:filteredComputers.Count

        # Populate the CheckedListBox with the filtered computer names
        $script:filteredComputers | ForEach-Object {
            $computerCheckedListBox.Items.Add($_.Name, $false) | Out-Null
            Write-Host "loaded:" $_.Name
            Write-Host ""
            $loadedCount++
            $deviceTimer++

            # Update the progress bar every 150 devices
            if ($deviceTimer -ge $deviceRefresh) {
                $deviceTimer = 0
                $progress = [math]::Round(($loadedCount / $computerCount) * 100)
                Write-Progress -Activity "Loading endpoints from AD (offline mode)..." -Status "$progress% Complete:" -PercentComplete $progress
            }
        }

        Write-Progress -Activity "Loading endpoints from AD (offline mode)..." -Completed
        Write-Host "Successfully loaded $computerCount endpoints" -ForegroundColor Green
        Write-Host "Filtered out $filteredOutCount endpoints due to 180 day offline exclusion"
        
        # Re-enable the form after loading is complete
        $form.Enabled = $true
        Write-Host "Form enabled."
        Write-Host ""
    }
    catch {
        Write-Progress -Activity "Loading endpoints from AD (offline mode)..." -Completed
        Write-Host "Error loading AD endpoints: $_" -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to close the window..."
    }
}
            

# Function to load and filter computer objects from Active Directory based on the selected Organizational Unit (OU)
# This function populates the provided CheckedListBox with computer names that have logged on within the last 180 days
# It filters out computers that have not logged in within the past 180 days.
function LoadAndFilterComputers {
    param (
        [System.Windows.Forms.CheckedListBox]$computerCheckedListBox
    )

    try {
        # Show the OU selection form
        Select-OU | Out-Null
        if (-not $script:ouPath) {
            Write-Host "No OU selected, using DC=users,DC=campus"
            $script:ouPath = 'DC=users,DC=campus'
            return
        }
        #example# $OUpath = 'OU=Workstations,OU=KSUL,OU=Dept,DC=users,DC=campus'

        Write-Host "Selected OU Path: $script:ouPath"  # Debug message

        # Disable the form while loading data to prevent user interaction
        $form.Enabled = $false
        Write-Host "Form disabled for loading..."
        Write-Host "Loading AD endpoints..."

        # Initialize counters and timers for progress tracking
        $loadedCount = 0
        $deviceRefresh = 150
        $deviceTimer = 0

        # Define the cutoff date for filtering computers based on their last logon date
        $cutoffDate = (Get-Date).AddDays(-180)

        # Query Active Directory for computers within the selected OU and retrieve their last logon date
        $computers = Get-ADComputer -Filter * -Properties LastLogonDate -SearchBase $script:ouPath

        # Populate the CheckedListBox with the filtered computer names
        $script:filteredComputers = $computers | Where-Object {
            $_.LastLogonDate -and
            [DateTime]::Parse($_.LastLogonDate) -ge $cutoffDate
        }

        $computerCount = $script:filteredComputers.Count
        $filteredOutCount = $computers.Count - $script:filteredComputers.Count

        $script:filteredComputers | ForEach-Object {
            $computerCheckedListBox.Items.Add($_.Name, $false) | Out-Null
            $loadedCount++
            $deviceTimer++
            Write-Host "loaded:" $_.Name
            Write-Host ""
            # Update the progress bar every 150 devices
            if ($deviceTimer -ge $deviceRefresh) {
                $deviceTimer = 0
                $progress = [math]::Round(($loadedCount / $computerCount) * 100)
                Write-Progress -Activity "Loading endpoints from AD..." -Status "$progress% Complete:" -PercentComplete $progress
            }
        }

        Write-Progress -Activity "Loading endpoints from AD..." -Completed
        Write-Host "Successfully loaded $computerCount endpoints" -ForegroundColor Green
        Write-Host "Filtered out $filteredOutCount endpoints due to 180 day offline exclusion"
        
        # Re-enable the form after loading is complete
        $form.Enabled = $true
        Write-Host "Form enabled."
        Write-Host ""
    }
    catch {
        Write-Progress -Activity "Loading endpoints from AD..." -Completed
        Write-Host "Error loading AD endpoints: $_" -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to close the window..."
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(785, 520)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'

if ($online) {
    $form.Text = "ONLINE - Dawson's AD Computer Renamer $Version"
    Write-Host 'You have started this application in ONLINE mode. Set the variable $online to $false for OFFLINE mode. (Line 104)' -ForegroundColor Yellow
}
else {
    $form.Text = "OFFLINE - Dawson's AD Computer Renamer $Version"
    Write-Host 'You have started this application in OFFLINE mode. Set the variable $online to $true for ONLINE mode. (Line 104)' -ForegroundColor Yellow
}

# Customize form appearance
$form.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240, 240) # Light grey background
$form.ForeColor = [System.Drawing.Color]::FromArgb(255, 105, 105, 105) # Dark grey text
$form.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold

# Set the icon for the form
$iconPath = Join-Path $PSScriptRoot "icon.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
}
else {
    Write-Host "icon not found at path $iconPath"
}

# Initialize script-scope variables
# $script:checkedItems = @{}
$script:invalidNamesList = @()
$script:validNamesList = @()
$script:customNamesList = @() 
$script:ouPath = 'DC=users,DC=campus'

# Create label to display current script version
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "Version $Version"
$versionLabel.Location = New-Object System.Drawing.Point(690, 460)
$versionLabel.AutoSize = $true
$form.Controls.Add($versionLabel)

# Create label to display author information
$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Text = "Author: Dawson Adams (dawsonaa@ksu.edu)"
$authorLabel.Location = New-Object System.Drawing.Point(10, 460)
$authorLabel.AutoSize = $true
$form.Controls.Add($authorLabel)

# Define the size for the list boxes
$listBoxWidth = 250
$listBoxHeight = 350

# Define the script-wide variables
$script:checkedItems = @{}
$script:selectedCheckedItems = @{}

# Function to sync checked items to computerCheckedListBox
function SyncCheckedItems {
    $computerCheckedListBox.Items.Clear()
    foreach ($item in $listBox.Items) {
        $computerCheckedListBox.Items.Add($item, $script:checkedItems.ContainsKey($item))
    }
}

# Function to sync selected checked items to selectedCheckedListBox
function SyncSelectedCheckedItems {
    $selectedCheckedListBox.Items.Clear()
    foreach ($item in $script:checkedItems.Keys) {
        $selectedCheckedListBox.Items.Add($item, $script:selectedCheckedItems.ContainsKey($item))
    }
}

# Create checked list box for computers
$computerCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$computerCheckedListBox.Location = New-Object System.Drawing.Point(10, 40)
$computerCheckedListBox.Size = New-Object System.Drawing.Size($listBoxWidth, $listBoxHeight)

# Handle the KeyDown event to detect Ctrl+A
$computerCheckedListBox.Add_KeyDown({
        param($s, $e)
    
        # Check if Ctrl+A is pressed
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            # Disable the form and controls to prevent interactions
            $form.Enabled = $false
            Write-Host "Ctrl+A pressed, starting mass selection, Form disabled for loading..."

            # Limit the number of items to select to 500
            $maxItemsToCheck = 500
            $itemsChecked = 0

            # Check all items in the CheckedListBox up to the limit
            for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
                if ($itemsChecked -ge $maxItemsToCheck) {
                    break
                }
                $computerCheckedListBox.SetItemChecked($i, $true)
                $currentItem = $computerCheckedListBox.Items[$i]
                $script:checkedItems[$currentItem] = $true
                $itemsChecked++
            }

            # Disable the form and controls to prevent interactions
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host ""

            # Prevent default action
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })


# Attach the event handler to the CheckedListBox

$form.Controls.Add($computerCheckedListBox)

# Create a new checked list box for displaying selected computers
$selectedCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$selectedCheckedListBox.Location = New-Object System.Drawing.Point(260, 40)
$selectedCheckedListBox.Size = New-Object System.Drawing.Size($listBoxWidth, ($listBoxHeight))

# Define the script-wide variable for selectedCheckedListBox
# $script:selectedCheckedItems = @{}

# Event handler for checking items in computerCheckedListBox
$computerCheckedListBox_ItemCheck = {
    param($s, $e)

    $item = $s.Items[$e.Index]
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        $script:checkedItems[$item] = $true
    }
    else {
        $script:checkedItems.Remove($item)
        $script:selectedCheckedItems.Remove($item)
    }
    SyncSelectedCheckedItems
}

# Event handler for checking items in selectedCheckedListBox
$selectedCheckedListBox_ItemCheck = {
    param($s, $e)

    $item = $s.Items[$e.Index]
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        $script:selectedCheckedItems[$item] = $true
    }
    else {
        $script:selectedCheckedItems.Remove($item)
        SyncSelectedCheckedItems
    }
}

# Subscribe to the ItemCheck events
$computerCheckedListBox.Add_ItemCheck($computerCheckedListBox_ItemCheck)
$selectedCheckedListBox.Add_ItemCheck($selectedCheckedListBox_ItemCheck)

# Initial sync
SyncCheckedItems
SyncSelectedCheckedItems



# Create the context menu for right-click actions
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Create menu context item for selecting specific devices to remove
$menuRemove = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemove.Text = "Remove selected device(s)"
$menuRemove.Add_Click({
        # Update the script-wide variable with currently checked items
        $global:selectedItems = @($selectedCheckedListBox.CheckedItems | ForEach-Object { $_ })

        if (!($selectedItems.Count -gt 0)) {
            Write-Host "No devices selected"
            return
        }
        Write-Host ""

        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            Write-Host "Selected device removed: $item"  # Outputs the names of selected devices to the console
        }

        Write-Host ""
        #UpdateAllListBoxes  # Call this function if it updates the UI based on changes
    })

# Create menu context item for removing all devices within the selectedComputersListBox
$menuRemoveAll = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemoveAll.Text = "Remove all device(s)"
$menuRemoveAll.Enabled = $false
$menuRemoveAll.Add_Click({
        $global:selectedItems = @($selectedCheckedListBox.Items | ForEach-Object { $_ })
        $form.Enabled = $false
        Write-Host "Form disabled for remove all"
        $script:customNamesList = @()
        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            #Write-Host "device removed: $item"  # Outputs the names of selected devices to the console
        }
        $form.Enabled = $true
        Write-Host "Form enabled"
        Write-Host ""

        # UpdateAllListBoxes  # Call this function if it updates the UI based on changes
    })

# Create context menu item for adding a custom name if one item in the selectedComputersListBox is selected
$menuAddCustomName = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuAddCustomName.Text = "Set/Change Custom rename"
$menuAddCustomName.Enabled = $false
$menuAddCustomName.Add_Click({
        $global:selectedItems = @($selectedCheckedListBox.CheckedItems | ForEach-Object { $_ })
        $tempList = $script:customNamesList
        $script:customNamesList = @()

        # Preserve existing items in customNamesList
        foreach ($tempItem in $tempList) {
            $isSelected = $false
            foreach ($selectedItem in $selectedItems) {
                if ($tempItem -match "^$selectedItem\s*->") {
                    $isSelected = $true
                    break
                }
            }
            if (-not $isSelected) {
                $script:customNamesList += $tempItem
            }
        }    

        foreach ($item in $selectedItems) {
            # Prompt for a custom name
            $customItem = Show-InputBox -message "Enter custom name for $item :" -title "Custom Name" -defaultText $item
    
            if ($customItem -and $customItem -ne "") {
                # Add the new custom name
                $script:customNamesList += "$item -> $customItem"
                # Write-Host "$item -> $customItem" # for debugging
            }
        }
        # UpdateAllListBoxes
    })

# Create context menu item for removing the custom names attached to selected items within the selectedComputersListBox
$menuRemoveCustomName = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemoveCustomName.Text = "Remove Custom rename"
$menuRemoveCustomName.Enabled = $false
$menuRemoveCustomName.Add_Click({
        $global:selectedItems = @($selectedCheckedListBox.CheckedItems | ForEach-Object { $_ })
        $tempList = $script:customNamesList
        $script:customNamesList = @()

        foreach ($tempItem in $tempList) {
            $isSelected = $false
            foreach ($selectedItem in $selectedItems) {
                if ($tempItem -match "^$selectedItem\s*->") {
                    $isSelected = $true
                    break
                }
            }
            if (-not $isSelected) {
                $script:customNamesList += $tempItem
            }
        }
        # UpdateAllListBoxes
    })

# Event handler for when the context menu is opening
$contextMenu.add_Opening({
        # Check if there are any items
        $global:selectedItems = @($selectedCheckedListBox.CheckedItems)
        $itemsInBox = @($selectedCheckedListBox.Items)

        if ($itemsInBox.Count -gt 0) {
            $menuRemoveAll.Enabled = $true
            $menuFindAndReplace.Enabled = $true
        }
        else {
            $menuRemoveAll.Enabled = $false
            $menuFindAndReplace.Enabled = $false
        }

        if ($selectedItems.Count -gt 0) {
            $menuRemove.Enabled = $true  # Enable the menu item if items are checked

            # Check if all checked items are in customNamesList
            $allInCustomNamesList = $true
            foreach ($item in $selectedItems) {
                if (-not ($script:customNamesList | Where-Object { $_ -match "^$item\s*->" })) {
                    $allInCustomNamesList = $false
                    break
                }
            }
        
            if ($allInCustomNamesList) {
                $menuRemoveCustomName.Enabled = $true
            }
            else {
                $menuRemoveCustomName.Enabled = $false
            }

            if ($selectedItems.Count -eq 1) {
                $menuAddCustomName.Enabled = $true
            }
            else {
                $menuAddCustomName.Enabled = $false
            }         
        }
        else {
            $menuRemove.Enabled = $false  # Disable the menu item if no items are checked
            $menuRemoveCustomName.Enabled = $false
        }
    })

# Add the right click menu options to the context menu
#$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemove) | Out-Null
#$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemoveAll) | Out-Null
#$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuAddCustomName) | Out-Null
#$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemoveCustomName) | Out-Null
#$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuFindAndReplace) | Out-Null

# Attach the context menu to the CheckedListBox
# $selectedCheckedListBox.ContextMenuStrip = $contextMenu

# Add the key down event handler to selectedComputersListBox
$selectedCheckedListBox.add_KeyDown({
        param ($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $menuFindAndReplace.PerformClick()
            $e.Handled = $true
        }
    })

$form.Controls.Add($selectedCheckedListBox)


<# Create a new list box for displaying selected computers
$selectedComputersListBox = New-Object CustomListBox
$selectedComputersListBox.Location = New-Object System.Drawing.Point(260, 40)
$selectedComputersListBox.Size = New-Object System.Drawing.Size($listBoxWidth, ($listBoxHeight))
$selectedComputersListBox.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended  # Set selection mode to allow multiple selections
$form.Controls.Add($selectedComputersListBox)

# Create the context menu for right-click actions
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Create menu context item for finding and replacing strings within selected computers
$menuFindAndReplace = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFindAndReplace.Text = "Find and Replace"
$menuFindAndReplace.Add_Click({
        $searchString = Show-InputBox -message "Enter the string to search for (max 15 chars):" -title "Find and Replace"
        if (-not $searchString) { return }

        $replaceString = Show-InputBox -message "Enter the string to replace with (max 15 chars):" -title "Find and Replace"
        if (-not $replaceString) { return }

        $listBoxItems = @($selectedComputersListBox.Items | ForEach-Object { $_ })
        $tempList = $script:customNamesList
        $script:customNamesList = @()

        # Preserve existing items in customNamesList that do not match the searchString
        foreach ($tempItem in $tempList) {
            $computerName, $newName = $tempItem -split ' -> '
            if ($newName -notmatch [regex]::Escape($searchString)) {
                $script:customNamesList += $tempItem
            }
        }

        foreach ($entry in $listBoxItems) {
            $computerName, $newName = $entry -split ' -> '

            # Write-Host "Checking $computerName for $searchString" # for debugging
            if ($computerName -match [regex]::Escape($searchString)) {
                # Write-Host "Original newName: $newName" # for debugging
                $newName = $computerName -replace [regex]::Escape($searchString), $replaceString
                # Write-Host "Newname after replace: $newName" # for debugging

                # Ensure the new name is still valid (max length 15 characters)
                if ($newName.Length -le 15) {
                    $script:customNamesList += "$computerName -> $newName"
                    # Write-Host "$computerName -> $newName" # for debugging
                }
                else {
                    Write-Host "$computerName ignored in find and replace, exceeds 15 characters"
                }
            }
            else {
                Write-Host "$computerName does not contain $searchString"
            }
        
        }
        UpdateAllListBoxes
    })

# Add the key down event handler to selectedComputersListBox
$selectedComputersListBox.add_KeyDown({
        param ($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $menuFindAndReplace.PerformClick()
            $e.Handled = $true
        }
    })

# Create menu context item for selecting specific devices to remove
$menuRemove = New-Object System.Windows.Forms.ToolStripMenuItem
$menuRemove.Text = "Remove selected device(s)"
$menuRemove.Add_Click({
        # Create a temporary array of selected items to avoid modifying the collection while iterating
        $selectedItems = @($selectedComputersListBox.SelectedItems | ForEach-Object { $_ })

        if (!($selectedItems.Count -gt 0)) {
            Write-Host "No devices selected"
            continue
        }
        Write-Host ""

        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            Write-Host "Selected device removed: $item"  # Outputs the names of selected devices to the console
        }

        Write-Host ""
        UpdateAllListBoxes  # Call this function if it updates the UI based on changes
    })

# create menu context item for removing all devices within the selectedComputersListBox
$menuRemoveAll = New-Object System.Windows.Forms.ToolStripMenuItem
$menuRemoveAll.Text = "Remove all device(s)"
$menuRemoveAll.Enabled = $false
$menuRemoveAll.Add_Click({
        $selectedItems = @($selectedComputersListBox.Items | ForEach-Object { $_ })
        $form.Enabled = $false
        Write-Host "Form disabled for remove all"
        $script:customNamesList = @()
        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            #Write-Host "device removed: $item"  # Outputs the names of selected devices to the console
        }
        $form.Enabled = $true
        Write-Host "Form enabled"
        Write-Host ""

    })

# Create context menu item for adding a custom name if one item in the selectedComputersListBox is selected
$menuAddCustomName = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAddCustomName.Text = "Set/Change Custom rename"
$menuAddCustomName.Enabled = $false
$menuAddCustomName.Add_Click({
        $selectedItems = @($selectedComputersListBox.SelectedItems | ForEach-Object { $_ })
        $tempList = $script:customNamesList
        $script:customNamesList = @()

        # Preserve existing items in customNamesList
        foreach ($tempItem in $tempList) {
            $isSelected = $false
            foreach ($selectedItem in $selectedItems) {
                if ($tempItem -match "^$selectedItem\s*->") {
                    $isSelected = $true
                    break
                }
            }
            if (-not $isSelected) {
                $script:customNamesList += $tempItem
            }
        }    

        foreach ($item in $selectedItems) {
            # Prompt for a custom name
            $customItem = Show-InputBox -message "Enter custom name for $item :" -title "Custom Name" -defaultText $item
        
            if ($customItem -and $customItem -ne "") {
                # Add the new custom name
                $script:customNamesList += "$item -> $customItem"
                # Write-Host "$item -> $customItem" # for debugging
            }
        }
        UpdateAllListBoxes
    })

# Create context menu item for removing the custom names attached to selected items within the selectedComputersListBox
$menuRemoveCustomName = New-Object System.Windows.Forms.ToolStripMenuItem
$menuRemoveCustomName.Text = "Remove Custom rename"
$menuRemoveCustomName.Enabled = $false
$menuRemoveCustomName.Add_Click({
        $selectedItems = @($selectedComputersListBox.SelectedItems | ForEach-Object { $_ })
        $tempList = $script:customNamesList
        $script:customNamesList = @()

        foreach ($tempItem in $tempList) {
            $isSelected = $false
            foreach ($selectedItem in $selectedItems) {
                if ($tempItem -match "^$selectedItem\s*->") {
                    $isSelected = $true
                    break
                }
            }
            if (-not $isSelected) {
                $script:customNamesList += $tempItem
            }
        }
        UpdateAllListBoxes
    })

# Event handler for when the context menu is opening
$contextMenu.Add_Opening({
        # Check if there are any selected items
        $selectedItems = @($selectedComputersListBox.SelectedItems)
        $itemsInBox = @($selectedComputersListBox.Items)

        if ($itemsInBox.Count -gt 0) {
            $menuRemoveAll.Enabled = $true
            $menuFindAndReplace.Enabled = $true
        }
        else {
            $menuRemoveAll.Enabled = $false
            $menuFindAndReplace.Enabled = $false
        }

        if ($selectedItems.Count -gt 0) {
            $menuRemove.Enabled = $true  # Enable the menu item if items are selected

            # Check if all selected items are in customNamesList
            $allInCustomNamesList = $true
            foreach ($item in $selectedItems) {
                if (-not ($script:customNamesList | Where-Object { $_ -match "^$item\s*->" })) {
                    $allInCustomNamesList = $false
                    break
                }
            }
            
            if ($allInCustomNamesList) {
                $menuRemoveCustomName.Enabled = $true
            }
            else {
                $menuRemoveCustomName.Enabled = $false
            }

            if ($selectedItems.Count -eq 1) {
                $menuAddCustomName.Enabled = $true
            }
            else {
                $menuAddCustomName.Enabled = $false
            }         
        }
        else {
            $menuRemove.Enabled = $false  # Disable the menu item if no items are selected
            $menuRemoveCustomName.Enabled = $false
        }
    })

# Add the right click menu options to the context menu
$contextMenu.Items.Add($menuRemove) | Out-Null
$contextMenu.Items.Add($menuRemoveAll) | Out-Null
$contextMenu.Items.Add($menuAddCustomName) | Out-Null
$contextMenu.Items.Add($menuRemoveCustomName) | Out-Null
$contextMenu.Items.Add($menuFindAndReplace) | Out-Null

# Attach the context menu to the ListBox
$selectedComputersListBox.ContextMenuStrip = $contextMenu
#>

# Create a list box for displaying proposed new names
$newNamesListBox = New-Object CustomListBox
$newNamesListBox.Location = New-Object System.Drawing.Point(510, 40)
$newNamesListBox.Size = New-Object System.Drawing.Size($listBoxWidth, $listBoxHeight)
# Override the selection behavior to prevent selection
$newNamesListBox.add_SelectedIndexChanged({
        $newNamesListBox.ClearSelected()
    })
$form.Controls.Add($newNamesListBox)

<# Create a label for the swapCheckBox
$swapCheckBoxLabel = New-Object System.Windows.Forms.Label
$swapCheckBoxLabel.Text = "Swap $part0Name and $part1Name"
$swapCheckBoxLabel.Location = New-Object System.Drawing.Point(30, 430)
$swapCheckBoxLabel.Size = New-Object System.Drawing.Size(200, 25)
$form.Controls.Add($swapCheckBoxLabel) #>

<# Synchronize the scrolling of the two list boxes # FIX
$selectedComputersListBox.add_Scroll({
        param($s, $e)
        $newNamesListBox.TopIndex = $selectedComputersListBox.TopIndex
    })

$newNamesListBox.add_Scroll({
        param($s, $e)
        $selectedComputersListBox.TopIndex = $newNamesListBox.TopIndex
    }) #>

# Search Text Box with Enter Key Event
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 10)
$searchBox.Size = New-Object System.Drawing.Size(180, 20)
$searchBox.ForeColor = [System.Drawing.Color]::Gray
$searchBox.Text = "Search for computer"
$searchBox.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $searchBox.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Clear placeholder text when the text box gains focus
$searchBox.Add_Enter({
        if ($this.Text -eq "Search for computer") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$searchBox.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Search for computer"
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })

# Handle the Enter key press for search
$searchBox.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $form.Enabled = $false
            Write-Host "Form disabled for search event, loading..."

            $e.SuppressKeyPress = $true  # Prevent sound on enter press
            $e.Handled = $true

            # Clear the checked list box
            $computerCheckedListBox.Items.Clear()

            # Filter computers
            $searchTerm = $searchBox.Text
            $filteredList = $script:filteredComputers | Where-Object { $_.Name -like "*$searchTerm*" }

            # Repopulate the checked list box with filtered computers and restore their checked state
            foreach ($computer in $filteredList) {
                $isChecked = $false
                if ($script:checkedItems.ContainsKey($computer.Name)) {
                    $isChecked = $script:checkedItems[$computer.Name]
                }
                $computerCheckedListBox.Items.Add($computer.Name, $isChecked)
            }
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host ""
        }
    })
$form.Controls.Add($searchBox)

# Add button to refresh or select a new OU to manage
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(195, 5)
$refreshButton.AutoSize = $true
$refreshButton.Text = 'Refresh / Change OU'
$refreshButton.Add_Click({
        $computerCheckedListBox.Items.Clear()
        $selectedCheckedListBox.Items.Clear()
        $newNamesListBox.Items.Clear()
        $script:checkedItems.Clear()
        # UpdateAllListBoxes

        if ($online) {
            LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox
        }
        else {
            LoadAndFilterComputersOFFLINE -computerCheckedListBox $computerCheckedListBox
        }

        # UpdateAllListBoxes
    })
$form.Controls.Add($refreshButton)

# Create label for selectedComputersListBox to show its filled with the original names
$beforeChangeLabel = New-Object System.Windows.Forms.Label
$beforeChangeLabel.Text = "Before Change"
$beforeChangeLabel.Location = New-Object System.Drawing.Point(340, 15)
$beforeChangeLabel.Size = New-Object System.Drawing.Size(110, 25)
$form.Controls.Add($beforeChangeLabel)

# Create label for newNamesListBox to show its filled with the manipulated names
$afterChangeLabel = New-Object System.Windows.Forms.Label
$afterChangeLabel.Text = "After Change"
$afterChangeLabel.Location = New-Object System.Drawing.Point(590, 15)
$afterChangeLabel.Size = New-Object System.Drawing.Size(110, 25)
$form.Controls.Add($afterChangeLabel)

function New-CustomTextBox {
    param (
        [string]$name,
        [string]$defaultText,
        [int]$x,
        [int]$y,
        [System.Drawing.Size]$size,
        [int]$maxLength
    )

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Name = $name
    $textBox.Location = New-Object System.Drawing.Point($x, $y)
    $textBox.Size = $size
    $textBox.ForeColor = [System.Drawing.Color]::Gray
    $textBox.BackColor = [System.Drawing.Color]::LightGray
    $textBox.Text = "Change $defaultText"
    $textBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $textBox.ReadOnly = $true
    $textBox.MaxLength = [Math]::Min(15, $maxLength)
    $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $textBox.Tag = $defaultText  # Store the default text in the Tag property

    $textBox.add_KeyDown({
            param($s, $e)
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
                $s.SelectAll()
                $e.SuppressKeyPress = $true
                $e.Handled = $true
            }
        })

    # MouseDown handler
    $textBox.add_MouseDown({
            param($s, $e)
            $defaultText = $s.Tag
            if ($s.ReadOnly) {
                $s.ReadOnly = $false
                $s.BackColor = [System.Drawing.Color]::White
                $s.ForeColor = [System.Drawing.Color]::Black
                $s.Focus()
                if ($s.Text -eq "Change $defaultText") {
                    $s.Text = ''
                }
            }
        })

    # Enter handler
    $textBox.Add_Enter({
            param($s, $e)
            $defaultText = $s.Tag
            if ($s.Text -eq "Change $defaultText") {
                $s.Text = ''
                $s.BackColor = [System.Drawing.Color]::White
                $s.ForeColor = [System.Drawing.Color]::Black
            }
        })

    # Leave handler
    $textBox.Add_Leave({
            param($s, $e)
            $defaultText = $s.Tag
            if ($s.Text -eq '') {
                $s.ReadOnly = $true
                $s.Text = "Change $defaultText"
                $s.ForeColor = [System.Drawing.Color]::Gray
                $s.BackColor = [System.Drawing.Color]::LightGray
            }
        })

    return $textBox
}

$textBoxSize = New-Object System.Drawing.Size(150, 20)

$gap = 40

# Calculate the total width occupied by the text boxes and their distances
$totalWidth = (4 * $textBoxSize.Width) + (3 * $gap) # 3 gaps between 4 text boxes, each gap is 20 pixels

# Determine the starting X-coordinate to center the group of text boxes
$startX = [Math]::Max(($form.ClientSize.Width - $totalWidth) / 2, 0)

# Create and add the text boxes, setting their X-coordinates based on the starting point
$part0Input = New-CustomTextBox -name "part0Input" -defaultText "part0Name" -x $startX -y 400 -size $textBoxSize -maxLength 15
$form.Controls.Add($part0Input)

$part1Input = New-CustomTextBox -name "part1Input" -defaultText "part1Name" -x ($startX + $textBoxSize.Width + $gap) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part1Input)

$part2Input = New-CustomTextBox -name "part2Input" -defaultText "part2Name" -x ($startX + 2 * ($textBoxSize.Width + $gap)) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part2Input)

$part3Input = New-CustomTextBox -name "part3Input" -defaultText "part3Name" -x ($startX + 3 * ($textBoxSize.Width + $gap)) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part3Input)

# Part Input and CheckBox event triggers
# $part0Input.Add_TextChanged({ UpdateAllListBoxes })
# $part1Input.Add_TextChanged({ UpdateAllListBoxes })
# $part2Input.Add_TextChanged({ UpdateAllListBoxes })
# $part3Input.Add_TextChanged({ UpdateAllListBoxes })

<# Create input text box for part-0 name manipulation
$part0Input = New-Object System.Windows.Forms.TextBox
$part0Input.Location = New-Object System.Drawing.Point(30, 400)
$part0Input.Size = $textBoxSize
$part0Input.ForeColor = [System.Drawing.Color]::Gray
$part0Input.BackColor = [System.Drawing.Color]::LightGray
$part0Input.Text = "Change $part0Name"
# $part0Input.Enabled = $false # Initially disabled
$part0Input.ReadOnly = $true
$part0Input.MaxLength = [Math]::Min(15, $part0Max)
$part0Input.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $part0Input.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Event handler for simulating enable/disable behavior
$part0Input.add_Click({
        if ($part0Input.ReadOnly) {
            $part0Input.ReadOnly = $false
            $part0Input.BackColor = [System.Drawing.Color]::White
            $part0Input.ForeColor = [System.Drawing.Color]::Black
            $part0Input.Focus()
            if ($part0Input.Text -eq "Change $part0Name") {
                $part0Input.Text = ''
            }
        }
    })

# Clear placeholder text when the text box gains focus
$part0Input.Add_Enter({
        if ($this.Text -eq "Change $part0Name") {
            $this.Text = ''
            $part0Input.BackColor = [System.Drawing.Color]::White
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$part0Input.Add_Leave({
        if ($part0Input.Text -eq '') {
            $part0Input.Text = "Change $part0Name"
            $part0Input.ForeColor = [System.Drawing.Color]::Gray
            $part0Input.BackColor = [System.Drawing.Color]::LightGray
            $part0Input.ReadOnly = $true
        }
    })
$form.Controls.Add($part0Input) #>


<# Create input text box for part-1 name manipulation
$part1Input = New-Object System.Windows.Forms.TextBox
$part1Input.Location = New-Object System.Drawing.Point(200, 400) # 160, 400
$part1Input.Size = $textBoxSize
$part1Input.ForeColor = [System.Drawing.Color]::Gray
$part1Input.Text = "Change $part1Name"
$part1Input.Enabled = $false  # Initially disabled
$part1Input.MaxLength = [Math]::Min(15, $part1Max)
$part1Input.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $part1Input.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Clear placeholder text when the text box gains focus
$part1Input.Add_Enter({
        if ($this.Text -eq "Change $part1Name") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$part1Input.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Change $part1Name"
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })
$form.Controls.Add($part1Input) #>

<# Create input text box for part-2 name manipulation
$part2Input = New-Object System.Windows.Forms.TextBox
$part2Input.Location = New-Object System.Drawing.Point(290, 400)
$part2Input.Size = $textBoxSize
$part2Input.ForeColor = [System.Drawing.Color]::Gray
$part2Input.Text = "Change $part2Name"
$part2Input.MaxLength = 15
$part2Input.Enabled = $false  # Initially disabled
$part2Input.Visible = $false # FIX
$part2Input.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $part2Input.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Clear placeholder text when the text box gains focus
$part2Input.Add_Enter({
        if ($this.Text -eq "Change $part2Name") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$part2Input.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Change $part2Name"
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })
$form.Controls.Add($part2Input) #>

<# Create input text box for part-3 name manipulation
$part3Input = New-Object System.Windows.Forms.TextBox
$part3Input.Location = New-Object System.Drawing.Point(420, 400)
$part3Input.Size = $textBoxSize
$part3Input.ForeColor = [System.Drawing.Color]::Gray
$part3Input.Text = "Change $part3Name"
$part3Input.MaxLength = 15
$part3Input.Enabled = $false  # Initially disabled
$part3Input.Visible = $false # FIX
$part3Input.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $part3Input.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Clear placeholder text when the text box gains focus
$part3Input.Add_Enter({
        if ($this.Text -eq "Change $part3Name") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$part3Input.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Change $part3Name"
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })
$form.Controls.Add($part3Input) #>
<#
<# Create checkbox for swapping department and part1 names
$swapCheckBox = New-Object System.Windows.Forms.CheckBox
$swapCheckBox.Location = New-Object System.Drawing.Point(10, 430)
$swapCheckBox.AutoSize = $true
$swapCheckBox.Visible = $false # HIDDEN
$swapCheckBox.Add_CheckedChanged({
        if ($swapCheckBox.Checked) {
            # FIX
            # $part0CheckBox.Enabled = $false
            # $part1CheckBox.Enabled = $false
        }
        else {
            # FIX
            # $part0CheckBox.Enabled = $true
            # $part1CheckBox.Enabled = $true
        }
        UpdateAllListBoxes
    })
$form.Controls.Add($swapCheckBox) #>

<# Create checkbox to enable/disable per-part 15 character limit
$maxCharacterCheckBox = New-Object System.Windows.Forms.CheckBox
$maxCharacterCheckBox.Location = New-Object System.Drawing.Point(270, 430)
$maxCharacterCheckBox.AutoSize = $true
$maxCharacterCheckBox.Add_CheckedChanged({
        if ($maxCharacterCheckBox.Checked) {
            $part0Input.MaxLength = 15
            $part1Input.MaxLength = 15
        }
        else {
            $part0Input.MaxLength = [Math]::Min(15, $part0Max)
            $part1Input.MaxLength = [Math]::Min(15, $part0Max)
        }
        UpdateAllListBoxes
    })
$form.Controls.Add($maxCharacterCheckBox) #>

<# Create label for maxCharacterCheckBox
$maxCharacterCheckBoxLabel = New-Object System.Windows.Forms.Label
$maxCharacterCheckBoxLabel.Text = "Set each part to a maximum of 15 characters"
$maxCharacterCheckBoxLabel.Location = New-Object System.Drawing.Point(290, 430)
$maxCharacterCheckBoxLabel.Size = New-Object System.Drawing.Size(200, 25)
$form.Controls.Add($maxCharacterCheckBoxLabel) #>

<# Create checkbox to enable/disable part-1 name manipulation
$part1CheckBox = New-Object System.Windows.Forms.CheckBox
$part1CheckBox.Location = New-Object System.Drawing.Point(140, 400)
$part1CheckBox.AutoSize = $true
$part1CheckBox.Visible = $false # HIDDEN
$part1CheckBox.Add_CheckedChanged({
        $part1Input.Enabled = $part1CheckBox.Checked

        # FIX
        if ($part1CheckBox.Checked -or $part0CheckBox.Checked) {
            $swapCheckBox.Enabled = $false
        }
        else {
            $swapCheckBox.Enabled = $true
        }
        

        if ($part1CheckBox.Checked -eq $false) {
            $part1Input.Text = "Change $part1Name"
        }
        else {
            $part1Input.Text = ""
            $part1Input.ForeColor = [System.Drawing.Color]::Black
        }

    })
$form.Controls.Add($part1CheckBox) #>

<# Create checkbox to enable/disable part-2 name manipulation
$part2CheckBox = New-Object System.Windows.Forms.CheckBox
$part2CheckBox.Location = New-Object System.Drawing.Point(270, 400)
$part2CheckBox.AutoSize = $true
$part2CheckBox.Visible = $false # HIDDEN
$part2CheckBox.Add_CheckedChanged({
        $part2Input.Enabled = $part2CheckBox.Checked

        # FIX
        if ($part1CheckBox.Checked -or $part0CheckBox.Checked) {
            $swapCheckBox.Enabled = $false
        }
        else {
            $swapCheckBox.Enabled = $true
        } 

        if ($part2CheckBox.Checked -eq $false) {
            $part2Input.Text = "Change $part2Name"
        }
        else {
            $part2Input.Text = ""
            $part2Input.ForeColor = [System.Drawing.Color]::Black
        }

    })
$form.Controls.Add($part2CheckBox) #>

<# Create checkbox to enable/disable part-3 name manipulation
$part3CheckBox = New-Object System.Windows.Forms.CheckBox
$part3CheckBox.Location = New-Object System.Drawing.Point(400, 400)
$part3CheckBox.AutoSize = $true
$part3CheckBox.Visible = $false # HIDDEN
$part3CheckBox.Add_CheckedChanged({
        $part3Input.Enabled = $part3CheckBox.Checked

        if ($part3CheckBox.Checked -eq $false) {
            $part3Input.Text = "Change $part3Name"
        }
        else {
            $part3Input.Text = ""
            $part3Input.ForeColor = [System.Drawing.Color]::Black
        }

    })
$form.Controls.Add($part3CheckBox) #>

<# Part Input and CheckBox event triggers
$part0Input.Add_TextChanged({ UpdateAllListBoxes })
$part1Input.Add_TextChanged({ UpdateAllListBoxes })
$part2Input.Add_TextChanged({ UpdateAllListBoxes })
$part3Input.Add_TextChanged({ UpdateAllListBoxes })


$part1CheckBox.Add_Click({ UpdateAllListBoxes })
$part2CheckBox.Add_Click({ UpdateAllListBoxes })
$part3CheckBox.Add_Click({ UpdateAllListBoxes })
$swapCheckBox.Add_Click({ UpdateAllListBoxes })
$maxCharacterCheckBox.Add_Click({ UpdateAllListBoxes }) #>

# Function to create styled buttons
function New-StyledButton {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width = 100,
        [int]$height = 35,
        [bool]$enabled = $true
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.Text = $text
    $button.Enabled = $enabled

    return $button
}
$commitChangesButton = New-StyledButton -text "Commit Changes" -x 480 -y 430 -width 100 -height 40 -enabled $true

# Event handler for clicking the Commit Changes button
$commitChangesButton.Add_Click({
        UpdateAllListBoxes
    })
$form.Controls.Add($commitChangesButton)


$applyRenameButton = New-StyledButton -text "Apply Rename" -x 580 -y 430 -width 100 -height 40 -enabled $false

<# Create and configure the 'Apply Rename' button
$applyRenameButton = New-Object System.Windows.Forms.Button
$applyRenameButton.Location = New-Object System.Drawing.Point(580, 430)
$applyRenameButton.Size = New-Object System.Drawing.Size(100, 35)
$applyRenameButton.Text = 'Apply Rename'
$applyRenameButton.Enabled = $false #>

<# ApplyRenameButton click event to start renaming process if user chooses
$applyRenameButton.Add_Click({
        # Create a string from the invalid names list
        if ($script:invalidNamesList.Count -gt 0) {
            $url = "https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163"
            $message = "The below invalid renames will be ignored:`n" + ($script:invalidNamesList -join "`n") + "`n`nDo you want to review the guidelines?" + "`n`n'Yes' = open guidelines, 'No' = continue, 'Cancel' = cancel rename"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Invalid Renaming schemes found", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
            if ($result -eq 'Yes') {
                Start-Process $url  # Open the guidelines URL
            }
            elseif ($result -eq 'Cancel') {
                return  # Exit if the user cancels
            }
        }

        # Prompt the user to confirm if they want to proceed with renaming
        $userResponse = [System.Windows.Forms.MessageBox]::Show(("`nDo you want to proceed with renaming? `n`n"), "Apply Rename", [System.Windows.Forms.MessageBoxButtons]::YesNo)

        # Initialize variables
        $successfulRenames = @()
        $failedRenames = @()
        $successfulRestarts = @()
        $failedRestarts = @()
        $loggedOnUsers = @()
        $loggedOnDevices = @() # Array to store offline devices and their users
        $totalTime = [System.TimeSpan]::Zero

        UpdateAllListBoxes # Update all list boxes

        #  If user confirms they want to proceed with renaming
        if ($userResponse -eq "Yes") {
            # Initialize the log variable
            $logContent = ""

            # Redefine Write-Host to also capture log content
            function Write-Host {
                param (
                    [Parameter(Mandatory = $true, Position = 0)]
                    [string] $Object,
                    [ConsoleColor] $ForegroundColor,
                    [ConsoleColor] $BackgroundColor
                )
                $script:logContent += $Object + "`n"
                if ($PSBoundParameters.ContainsKey('ForegroundColor') -and $PSBoundParameters.ContainsKey('BackgroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
                }
                elseif ($PSBoundParameters.ContainsKey('ForegroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -ForegroundColor $ForegroundColor
                }
                elseif ($PSBoundParameters.ContainsKey('BackgroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -BackgroundColor $BackgroundColor
                }
                else {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object
                }
            }
        
            $form.Enabled = $false
            Write-Host "Form disabled for rename operation, loading..."
            Write-Host " "
            Write-Host "Starting rename operation..."
            Write-Host " "

            # Iterate through the valid names list and perform renaming operations
            foreach ( $computerName in $script:invalidNamesList) {
                Write-Host "$computerName has been ignored due to invalid naming scheme" -ForegroundColor Red
            }

            # Check if new name is the same as old name
            foreach ($renameEntry in $script:validNamesList) {
                $individualTime = [System.TimeSpan]::Zero

                $oldName, $newName = $renameEntry -split ' -> '

                # Check if new name is the same as old name
                if ($oldName -eq $newName) {
                    Write-Host "New name for $oldName is the same as the old one. Ignoring this device." -ForegroundColor Yellow
                    Write-Host " "
                    continue
                }

                $checkOfflineTime = Measure-Command {
                    # Simulate checking if the computer is online
                    Write-Host "Checking if $oldName is online..."
                    $onlineStatus = Get-RandomOutcome -outcomes $onlineStatuses
                    if ($onlineStatus -eq "Offline") {
                        Write-Host "Computer $oldName is offline. Skipping rename." -ForegroundColor Red
                        Write-Host " "
                        $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                        continue
                    }
                    Write-Host "Computer $oldName is online." -ForegroundColor Green
                }

                # Output the time taken to check if computer is online
                Write-Host "Time taken to check if $oldName was online: $($checkOfflineTime.TotalSeconds) seconds" -ForegroundColor Blue
                $individualTime = $individualTime.Add($checkOfflineTime)

                Write-Host "Checking if $oldName was renamed successfully..."

                # Start timing rename operation
                $checkRenameTime = Measure-Command {
                    try {
                        $renameResult = Get-RandomOutcome -outcomes $renameOutcomes 
                        if ($renameResult.ReturnValue -ne 0) {
                            throw "Failed to rename the computer $oldName. Error code: $($renameResult.ReturnValue)"
                        }
                    }
                    catch {
                        Write-Host "Error during rename operation: $_" -ForegroundColor Red
                        $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                        continue # Skip to the next iteration of the loop
                    }
                }
            
                # Check if computer was successfully renamed
                if ($renameResult.ReturnValue -eq 0) {
                    Write-Host "Computer $oldName successfully renamed to $newName." -ForegroundColor Green

                    # Output the time taken to rename
                    Write-Host "Time taken to rename $oldName`: $($checkRenameTime.TotalSeconds) seconds" -ForegroundColor Blue
                
                    $individualTime = $individualTime.Add($checkRenameTime)
                    $successfulRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }


                    #$loggedOnUserss = @("User1", "User2")

                    # Start timing $loggedOnUser operation
                    $checkLoginTime = Measure-Command {
                        Write-Host "Checking if $oldName has a user logged on..."
                        $loggedOnUser = Get-RandomOutcome -outcomes $loggedOnUserss
                        #$loggedOnUser = ""
                    }

                    if ($loggedOnUser -eq "none") {
                        $loggedOnUser = $null
                    }

                    # Start timing restart operation
                    $checkRestartTime = Measure-Command {
                        if (-not $loggedOnUser) {
                            try {
                                Write-Host "Computer $oldName ($newName) has no users logged on." -ForegroundColor Green

                                # Output the time taken to check if user was logged on
                                Write-Host "Time taken to check $oldName for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                $individualTime = $individualTime.Add($checkLoginTime)

                                Write-Host "Checking if $oldName restarted successfully..."
                                $restartOutcome = Get-RandomOutcome -outcomes $restartOutcomes
                                if ($restartOutcome -eq "Success") {
                                    Write-Host "Computer $oldName ($newName) successfully restarted." -ForegroundColor Green
                                    $successfulRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                }
                                else {
                                    throw "Manual restart required."
                                }
                            }
                            catch {
                                Write-Host "Computer $oldName ($newName) attempted to restart and failed. Manual restart required." -ForegroundColor Red
                                $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            }
                        }
                        else {
                            Write-Host "Computer $oldName ($newName) has $loggedOnUser logged in. Manual restart required." -ForegroundColor Yellow
                        
                            # Output the time taken to check if user was logged on
                            Write-Host "Time taken to check $oldName ($newName) for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                            $individualTime = $individualTime.Add($checkLoginTime)
                        
                            $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            $loggedOnUsers += "$oldName`: $loggedOnUser"

                            # Collect offline device information
                            $loggedOnDevices += [PSCustomObject]@{
                                OldName  = $oldName
                                NewName  = $newName
                                UserName = $loggedOnUser
                            }
                        }
                    }
                    # Output the time taken to send restart
                    Write-Host "Time taken to send restart to $oldName`: $($checkRestartTime.TotalSeconds) seconds" -ForegroundColor Blue
                    $individualTime = $individualTime.Add($checkRestartTime)
                }
                else {
                    # Output the time taken to rename
                    Write-Host "Time taken to rename $oldName to $newName`: $($checkRenameTime.TotalSeconds) seconds" -ForegroundColor Blue
                    $individualTime = $individualTime.Add($checkRenameTime)

                    Write-Host "Failed to rename the computer $oldName to $newName. Error code: $($renameResult.ReturnValue)" -ForegroundColor Red
                    Write-Host " "
                    $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                }

                $totalTime = $totalTime.Add($individualTime)
                Write-Host ("Total time taken for $oldName to be renamed: {0:F2} seconds" -f $individualTime.TotalSeconds) -ForegroundColor Blue
                Write-Host " "
            }
        
            Write-Host "Rename operation completed." 

            # Output the total time taken for all operations in the appropriate format
            if ($totalTime.TotalMinutes -lt 1) {
                Write-Host ("Total time taken for all rename operations: {0:F2} seconds" -f $totalTime.TotalSeconds) -ForegroundColor Blue
            }
            elseif ($totalTime.TotalHours -lt 1) {
                Write-Host ("Total time taken for all rename operations: {0:F2} minutes" -f $totalTime.TotalMinutes) -ForegroundColor Blue
            }
            else {
                Write-Host ("Total time taken for all rename operations: {0:F2} hours" -f $totalTime.TotalHours) -ForegroundColor Blue
            }
            Write-Host " "

            # Determine the script directory
            $scriptDir = Split-Path -Parent $PSCommandPath

            # Define the RESULTS and LOGS folder paths
            $resultsFolderPath = Join-Path -Path $scriptDir -ChildPath "RESULTS"
            $logsFolderPath = Join-Path -Path $scriptDir -ChildPath "LOGS"

            # Create the RESULTS folder if it doesn't exist
            if (-not (Test-Path -Path $resultsFolderPath)) {
                New-Item -Path $resultsFolderPath -ItemType Directory | Out-Null
            }

            # Create the LOGS folder if it doesn't exist
            if (-not (Test-Path -Path $logsFolderPath)) {
                New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
            }
 
            # Create the CSV file
            $csvData = @()

            # Determine the maximum count manually
            $maxCount = $successfulRenames.Count
            if ($successfulRestarts.Count -gt $maxCount) { $maxCount = $successfulRestarts.Count }
            if ($failedRenames.Count -gt $maxCount) { $maxCount = $failedRenames.Count }
            if ($failedRestarts.Count -gt $maxCount) { $maxCount = $failedRestarts.Count }
            if ($loggedOnUsers.Count -gt $maxCount) { $maxCount = $loggedOnUsers.Count }

            # Initialize CSV data as a string
            $csvData = ""

            # Add main headers and sub-headers row as a string
            $headers = @(
                "Successful Renames,,Successful Restarts,,Failed Renames,,Failed Restarts,,Logged On Users,,,"
                "New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,Logged On User"
            ) -join "`r`n"

            # Add data rows
            for ($i = 0; $i -lt $maxCount; $i++) {
                $successfulRenameNew = if ($i -lt $successfulRenames.Count) { $successfulRenames[$i].NewName } else { "" }
                $successfulRenameOld = if ($i -lt $successfulRenames.Count) { $successfulRenames[$i].OldName } else { "" }
                $successfulRestartNew = if ($i -lt $successfulRestarts.Count) { $successfulRestarts[$i].NewName } else { "" }
                $successfulRestartOld = if ($i -lt $successfulRestarts.Count) { $successfulRestarts[$i].OldName } else { "" }
                $failedRenameNew = if ($i -lt $failedRenames.Count) { $failedRenames[$i].NewName } else { "" }
                $failedRenameOld = if ($i -lt $failedRenames.Count) { $failedRenames[$i].OldName } else { "" }
                $failedRestartNew = if ($i -lt $failedRestarts.Count) { $failedRestarts[$i].NewName } else { "" }
                $failedRestartOld = if ($i -lt $failedRestarts.Count) { $failedRestarts[$i].OldName } else { "" }
                $loggedOnUserNew = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].NewName } else { "" }
                $loggedOnUserOld = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].OldName } else { "" }
                $loggedOnUser = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].UserName } else { "" }

                $csvData += "$successfulRenameNew,$successfulRenameOld,$successfulRestartNew,$successfulRestartOld,$failedRenameNew,$failedRenameOld,$failedRestartNew,$failedRestartOld,$loggedOnUserNew,$loggedOnUserOld,$loggedOnUser`r`n"
            }

            # Combine headers and data
            $csvOutput = "$headers`r`n$csvData"

            # Get the current date and time in the desired format
            $dateTimeString = (Get-Date).ToString("yy-MM-dd_HH-mmtt")

            # Create the CSV file path with the date and time appended
            $csvFileName = "ADRenamer_Results_$dateTimeString"
            $csvFilePath = Join-Path -Path $resultsFolderPath -ChildPath "$csvFileName.csv"

            # Write the combined output to the CSV file
            $csvOutput | Out-File -FilePath $csvFilePath -Encoding utf8

            Write-Host "RESULTS CSV file created at $csvFilePath" -ForegroundColor Yellow
            Write-Host " "

            # Save the log content to a .txt file
            $logFileName = "ADRenamer_Log_$dateTimeString"
            $logFilePath = Join-Path -Path $logsFolderPath -ChildPath "$logFileName.txt"
            $script:logContent | Out-File -FilePath $logFilePath -Encoding utf8

            Write-Host "LOGS TXT file created at $logFilePath" -ForegroundColor Yellow
            Write-Host " "

            # Convert the CSV file content to Base64
            # $fileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
            # $base64Content = [Convert]::ToBase64String($fileContent)

            # Convert the CSV file content to Base64
            $csvFileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
            $csvBase64Content = [Convert]::ToBase64String($csvFileContent)

            # Convert the log file content to Base64
            $logFileContent = [System.IO.File]::ReadAllBytes($logFilePath)
            $logBase64Content = [Convert]::ToBase64String($logFileContent)

            # Define the HTTP trigger URL for the Power Automate flow
            $flowUrl = "https://prod-166.westus.logic.azure.com:443/workflows/5e172f6d92d24c6a995023362c53472f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7-7I6wW8ga9i3hSfzjP7-O_AFLNFmE-_cxCGt6g3f9A"

            # Prepare the body of the request
            $body = @{
                csvFileName    = "$csvFileName`-$username.csv"
                csvFileContent = $csvBase64Content
                logFileName    = "$logFileName`-$username.txt"
                logFileContent = $logBase64Content
            }

            # Convert the body to JSON
            $jsonBody = $body | ConvertTo-Json -Depth 3

            # Set the headers
            $headers = @{
                "Content-Type" = "application/json"
            }

            # Send the HTTP POST request to trigger the flow
            Invoke-RestMethod -Uri $flowUrl -Method Post -Headers $headers -Body $jsonBody

            Write-Host "Triggered Power Automate flow to upload the log files to SharePoint" -ForegroundColor Yellow
            Write-Host " "

            # Print the list of logged on users
            if ($loggedOnUsers.Count -gt 0) {
                Write-Host "Logged on users:" -ForegroundColor Yellow
                $loggedOnUsers | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
                Show-EmailDrafts -loggedOnDevices $loggedOnDevices | Out-Null
            }
            else {
                Write-Host "No users were logged on to the renamed computers." -ForegroundColor Green
            }
            Write-Host " "

            # Remove successfully renamed computers from the list of computers
            foreach ($renameEntry in $script:validNamesList) {
                $oldName, $newName = $renameEntry -split ' -> '
                $script:checkedItems.Remove($oldName)
    
                # Remove the old name from the filteredComputers
                $script:filteredComputers = $script:filteredComputers | Where-Object { $_.Name -ne $oldName }
    
                # Try to remove the old name from the computerCheckedListBox
                $index = $computerCheckedListBox.Items.IndexOf($oldName)
                if ($index -ge 0) {
                    $computerCheckedListBox.Items.RemoveAt($index)
                }
            }
    
            # Refresh the checked list box based on current search term
            $searchTerm = $searchBox.Text
            $filteredList = $script:filteredComputers | Where-Object { $_.Name -like "*$searchTerm*" }
    
            # Clear and repopulate the checked list box with filtered computers and restore their checked state
            $computerCheckedListBox.Items.Clear()
            foreach ($computer in $filteredList) {
                $isChecked = $false
                if ($script:checkedItems.ContainsKey($computer.Name)) {
                    $isChecked = $script:checkedItems[$computer.Name]
                }
                $computerCheckedListBox.Items.Add($computer.Name, $isChecked)
            }
            UpdateAllListBoxes
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host " "

        }
    })
$form.Controls.Add($applyRenameButton)
#>

    


# ApplyRenameButton click event to start renaming process if user chooses # ONLINE
$applyRenameButton.Add_Click({
        # Create a string from the invalid names list
        if ($script:invalidNamesList.Count -gt 0) {
            $url = "https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163"
            $message = "The below invalid renames will be ignored:`n" + ($script:invalidNamesList -join "`n") + "`n`nDo you want to review the guidelines?" + "`n`n'Yes' = open guidelines, 'No' = continue, 'Cancel' = cancel rename"
            $result = [System.Windows.Forms.MessageBox]::Show($message, "Invalid Renaming schemes found", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
            if ($result -eq 'Yes') {
                Start-Process $url  # Open the guidelines URL
            }
            elseif ($result -eq 'Cancel') {
                return  # Exit if the user cancels
            }
        }

        # Prompt the user to confirm if they want to proceed with renaming
        $userResponse = [System.Windows.Forms.MessageBox]::Show(("`nDo you want to proceed with renaming? `n`n"), "Apply Rename", [System.Windows.Forms.MessageBoxButtons]::YesNo)

        # Initialize variables
        $successfulRenames = @()
        $failedRenames = @()
        $successfulRestarts = @()
        $failedRestarts = @()
        $loggedOnUsers = @()
        $loggedOnDevices = @() # Array to store offline devices and their users
        $totalTime = [System.TimeSpan]::Zero

        # UpdateAllListBoxes # Update all list boxes

        #  If user confirms they want to proceed with renaming
        if ($userResponse -eq "Yes") {
            # Initialize the log variable
            $logContent = ""

            # Redefine Write-Host to also capture log content
            function Write-Host {
                param (
                    [Parameter(Mandatory = $true, Position = 0)]
                    [string] $Object,
                    [ConsoleColor] $ForegroundColor,
                    [ConsoleColor] $BackgroundColor
                )
                $script:logContent += $Object + "`n"
                if ($PSBoundParameters.ContainsKey('ForegroundColor') -and $PSBoundParameters.ContainsKey('BackgroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor
                }
                elseif ($PSBoundParameters.ContainsKey('ForegroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -ForegroundColor $ForegroundColor
                }
                elseif ($PSBoundParameters.ContainsKey('BackgroundColor')) {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object -BackgroundColor $BackgroundColor
                }
                else {
                    Microsoft.PowerShell.Utility\Write-Host -Object $Object
                }
            }
            
            $form.Enabled = $false
            Write-Host "Form disabled for rename operation, loading..."
            Write-Host " "
            Write-Host "Starting rename operation..."
            Write-Host " "

            # Iterate through the valid names list and perform renaming operations
            foreach ( $computerName in $script:invalidNamesList) {
                Write-Host "$computerName has been ignored due to invalid naming scheme" -ForegroundColor Red
            }
    
            # Check if new name is the same as old name
            foreach ($renameEntry in $script:validNamesList) {
                $individualTime = [System.TimeSpan]::Zero

                $oldName, $newName = $renameEntry -split ' -> '

                # Check if new name is the same as old name
                if ($oldName -eq $newName) {
                    Write-Host "New name for $oldName is the same as the old one. Ignoring this device." -ForegroundColor Yellow
                    Write-Host " "
                    continue
                }

                if ($online) {
                    $checkOfflineTime = Measure-Command {
                        # Check if the computer is online
                        Write-Host "Checking if $oldName is online..."
                        if (-not (Test-Connection -ComputerName $oldName -Count 1 -Quiet)) {
                            Write-Host "Computer $oldName is offline. Skipping rename." -ForegroundColor Red
                            Write-Host " "
                            $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            continue
                        }
                        Write-Host "Computer $oldName is online." -ForegroundColor Green
                    }
                }
                else {
                    # OFFLINE
                    $checkOfflineTime = Measure-Command {
                        # Simulate checking if the computer is online
                        Write-Host "Checking if $oldName is online..."
                        $onlineStatus = Get-RandomOutcome -outcomes $onlineStatuses
                        if ($onlineStatus -eq "Offline") {
                            Write-Host "Computer $oldName is offline. Skipping rename." -ForegroundColor Red
                            Write-Host " "
                            $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            continue
                        }
                        Write-Host "Computer $oldName is online." -ForegroundColor Green
                    }
                }

                # Output the time taken to check if computer is online
                Write-Host "Time taken to check if $oldName was online: $($checkOfflineTime.TotalSeconds) seconds" -ForegroundColor Blue
                $individualTime = $individualTime.Add($checkOfflineTime)
                
                if ($online) {
                    $testComp = Get-WmiObject Win32_ComputerSystem -ComputerName $oldName -Credential $cred
                }
                
                Write-Host "Checking if $oldName was renamed successfully..."

                if ($online) {
                    # Start timing rename operation
                    $checkRenameTime = Measure-Command {
                        try {
                            $password = $cred.GetNetworkCredential().Password
                            $username = $cred.GetNetworkCredential().UserName
                            $renameResult = $testComp.Rename($newName, $password, $username)
                            if ($renameResult.ReturnValue -ne 0) {
                                throw "Failed to rename the computer $oldName. Error code: $($renameResult.ReturnValue)"
                            }
                        }
                        catch {
                            Write-Host "Error during rename operation: $_" -ForegroundColor Red
                            $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            continue # Skip to the next iteration of the loop
                        }
                    }
                }
                else {
                    # OFFLINE
                    # Start timing rename operation
                    $checkRenameTime = Measure-Command {
                        try {
                            $renameResult = Get-RandomOutcome -outcomes $renameOutcomes 
                            if ($renameResult.ReturnValue -ne 0) {
                                throw "Failed to rename the computer $oldName. Error code: $($renameResult.ReturnValue)"
                            }
                        }
                        catch {
                            Write-Host "Error during rename operation: $_" -ForegroundColor Red
                            $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                            continue # Skip to the next iteration of the loop
                        }
                    }
                }
                
                # Check if computer was successfully renamed
                if ($renameResult.ReturnValue -eq 0) {
                    Write-Host "Computer $oldName successfully renamed to $newName." -ForegroundColor Green

                    # Output the time taken to rename
                    Write-Host "Time taken to rename $oldName`: $($checkRenameTime.TotalSeconds) seconds" -ForegroundColor Blue
                    
                    $individualTime = $individualTime.Add($checkRenameTime)
                    $successfulRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }

                    if ($online) {
                        # Start timing $loggedOnUser operation
                        $checkLoginTime = Measure-Command {
                            Write-Host "Checking if $oldname has a user logged on..."
                            $loggedOnUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $oldName -Credential $cred | Select-Object -ExpandProperty UserName
                        }
                    }
                    else {
                        # Start timing $loggedOnUser operation
                        $checkLoginTime = Measure-Command {
                            Write-Host "Checking if $oldName has a user logged on..."
                            $loggedOnUser = Get-RandomOutcome -outcomes $loggedOnUserss
                        }
                    
                        if ($loggedOnUser -eq "none") {
                            $loggedOnUser = $null
                        }
                    }

                    if ($online) {
                        # Start timing restart operation
                        $checkRestartTime = Measure-Command {
                            if (-not $loggedOnUser) {
                                try {
                                    Write-Host "Computer $oldname ($newName) has no users logged on." -ForegroundColor Green

                                    # Output the time taken to check if user was logged on
                                    Write-Host "Time taken to check $oldName for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                    $individualTime = $individualTime.Add($checkLoginTime)

                                    Write-Host "Checking if $oldName restarted successfully..."
                                    Restart-Computer -ComputerName $oldName -Credential $cred -Force
                                    Write-Host "Computer $oldName ($newName) successfully restarted." -ForegroundColor Green
                                    $successfulRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                }
                                catch {
                                    Write-Host "Computer $oldname ($newName) attempted to restart and failed. Manual restart required." -ForegroundColor Red
                                    $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                }
                            }
                            else {
                                Write-Host "Computer $oldname ($newName) has $loggedOnUser logged in. Manual restart required." -ForegroundColor Yellow #Need to add excel sheet creation to capture users logged into devices
                            
                                # Output the time taken to check if user was logged on
                                Write-Host "Time taken to check $oldname ($newName) has logged on users`: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                $individualTime = $individualTime.Add($checkLoginTime)
                            
                                $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                $loggedOnUsers += "$oldName`: $loggedOnUser"

                                # Collect offline device information
                                $loggedOnDevices += [PSCustomObject]@{
                                    OldName  = $oldName
                                    NewName  = $newName
                                    UserName = $loggedOnUser
                                }
                            }
                        }
                    }
                    else {
                        # Start timing restart operation
                        $checkRestartTime = Measure-Command {
                            if (-not $loggedOnUser) {
                                try {
                                    Write-Host "Computer $oldName ($newName) has no users logged on." -ForegroundColor Green
                    
                                    # Output the time taken to check if user was logged on
                                    Write-Host "Time taken to check $oldName for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                    $individualTime = $individualTime.Add($checkLoginTime)
                    
                                    Write-Host "Checking if $oldName restarted successfully..."
                                    $restartOutcome = Get-RandomOutcome -outcomes $restartOutcomes
                                    if ($restartOutcome -eq "Success") {
                                        Write-Host "Computer $oldName ($newName) successfully restarted." -ForegroundColor Green
                                        $successfulRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                    }
                                    else {
                                        throw "Manual restart required."
                                    }
                                }
                                catch {
                                    Write-Host "Computer $oldName ($newName) attempted to restart and failed. Manual restart required." -ForegroundColor Red
                                    $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                }
                            }
                            else {
                                Write-Host "Computer $oldName ($newName) has $loggedOnUser logged in. Manual restart required." -ForegroundColor Yellow
                                            
                                # Output the time taken to check if user was logged on
                                Write-Host "Time taken to check $oldName ($newName) for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                $individualTime = $individualTime.Add($checkLoginTime)
                                            
                                $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                $loggedOnUsers += "$oldName`: $loggedOnUser"
                    
                                # Collect offline device information
                                $loggedOnDevices += [PSCustomObject]@{
                                    OldName  = $oldName
                                    NewName  = $newName
                                    UserName = $loggedOnUser
                                }
                            }
                        }
                    }
                    # Output the time taken to send restart
                    Write-Host "Time taken to send restart to $oldname`: $($checkRestartTime.TotalSeconds) seconds" -ForegroundColor Blue
                    $individualTime = $individualTime.Add($checkRestartTime)
                }
                else {
                    # Output the time taken to rename
                    Write-Host "Time taken to rename $oldName to $newName`: $($checkRenameTime.TotalSeconds) seconds" -ForegroundColor Blue
                    $individualTime = $individualTime.Add($checkRenameTime)

                    Write-Host "Failed to rename the computer $oldName to $newName. Error code: $($renameResult.ReturnValue)" -ForegroundColor Red
                    Write-Host " "
                    $failedRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                }

                $totalTime = $totalTime.Add($individualTime)
                Write-Host ("Total time taken for $oldName to be renamed: {0:F2} seconds" -f $individualTime.TotalSeconds) -ForegroundColor Blue
                Write-Host " "
            }
            
            Write-Host "Rename operation completed." 

            # Output the total time taken for all operations in the appropriate format
            if ($totalTime.TotalMinutes -lt 1) {
                Write-Host ("Total time taken for all rename operations: {0:F2} seconds" -f $totalTime.TotalSeconds) -ForegroundColor Blue
            }
            elseif ($totalTime.TotalHours -lt 1) {
                Write-Host ("Total time taken for all rename operations: {0:F2} minutes" -f $totalTime.TotalMinutes) -ForegroundColor Blue
            }
            else {
                Write-Host ("Total time taken for all rename operations: {0:F2} hours" -f $totalTime.TotalHours) -ForegroundColor Blue
            }
            Write-Host " "

            # Determine the script directory
            $scriptDir = Split-Path -Parent $PSCommandPath

            # Define the RESULTS and LOGS folder paths
            $resultsFolderPath = Join-Path -Path $scriptDir -ChildPath "RESULTS"
            $logsFolderPath = Join-Path -Path $scriptDir -ChildPath "LOGS"

            # Create the RESULTS folder if it doesn't exist
            if (-not (Test-Path -Path $resultsFolderPath)) {
                New-Item -Path $resultsFolderPath -ItemType Directory | Out-Null
            }

            # Create the LOGS folder if it doesn't exist
            if (-not (Test-Path -Path $logsFolderPath)) {
                New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
            }
            
            # Create the CSV file
            $csvData = @()

            # Determine the maximum count manually
            $maxCount = $successfulRenames.Count
            if ($successfulRestarts.Count -gt $maxCount) { $maxCount = $successfulRestarts.Count }
            if ($failedRenames.Count -gt $maxCount) { $maxCount = $failedRenames.Count }
            if ($failedRestarts.Count -gt $maxCount) { $maxCount = $failedRestarts.Count }
            if ($loggedOnUsers.Count -gt $maxCount) { $maxCount = $loggedOnUsers.Count }

            # Initialize CSV data as a string
            $csvData = ""

            # Add main headers and sub-headers row as a string
            $headers = @(
                "Successful Renames,,Successful Restarts,,Failed Renames,,Failed Restarts,,Logged On Users,,,"
                "New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,Logged On User"
            ) -join "`r`n"

            # Add data rows
            for ($i = 0; $i -lt $maxCount; $i++) {
                $successfulRenameNew = if ($i -lt $successfulRenames.Count) { $successfulRenames[$i].NewName } else { "" }
                $successfulRenameOld = if ($i -lt $successfulRenames.Count) { $successfulRenames[$i].OldName } else { "" }
                $successfulRestartNew = if ($i -lt $successfulRestarts.Count) { $successfulRestarts[$i].NewName } else { "" }
                $successfulRestartOld = if ($i -lt $successfulRestarts.Count) { $successfulRestarts[$i].OldName } else { "" }
                $failedRenameNew = if ($i -lt $failedRenames.Count) { $failedRenames[$i].NewName } else { "" }
                $failedRenameOld = if ($i -lt $failedRenames.Count) { $failedRenames[$i].OldName } else { "" }
                $failedRestartNew = if ($i -lt $failedRestarts.Count) { $failedRestarts[$i].NewName } else { "" }
                $failedRestartOld = if ($i -lt $failedRestarts.Count) { $failedRestarts[$i].OldName } else { "" }
                $loggedOnUserNew = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].NewName } else { "" }
                $loggedOnUserOld = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].OldName } else { "" }
                $loggedOnUser = if ($i -lt $loggedOnDevices.Count) { $loggedOnDevices[$i].UserName } else { "" }

                $csvData += "$successfulRenameNew,$successfulRenameOld,$successfulRestartNew,$successfulRestartOld,$failedRenameNew,$failedRenameOld,$failedRestartNew,$failedRestartOld,$loggedOnUserNew,$loggedOnUserOld,$loggedOnUser`r`n"
            }

            # Combine headers and data
            $csvOutput = "$headers`r`n$csvData"

            # Get the current date and time in the desired format
            $dateTimeString = (Get-Date).ToString("yy-MM-dd_HH-mmtt")

            # Create the CSV file path with the date and time appended
            $csvFileName = "ADRenamer_Results_$dateTimeString"
            $csvFilePath = Join-Path -Path $resultsFolderPath -ChildPath "$csvFileName.csv"

            # Write the combined output to the CSV file
            $csvOutput | Out-File -FilePath $csvFilePath -Encoding utf8

            Write-Host "RESULTS CSV file created at $csvFilePath" -ForegroundColor Yellow
            Write-Host " "

            # Save the log content to a .txt file
            $logFileName = "ADRenamer_Log_$dateTimeString"
            $logFilePath = Join-Path -Path $logsFolderPath -ChildPath "$logFileName.txt"
            $script:logContent | Out-File -FilePath $logFilePath -Encoding utf8

            Write-Host "LOGS TXT file created at $logFilePath" -ForegroundColor Yellow
            Write-Host " "

            # Convert the CSV file content to Base64
            # $fileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
            # $base64Content = [Convert]::ToBase64String($fileContent)

            # Convert the CSV file content to Base64
            $csvFileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
            $csvBase64Content = [Convert]::ToBase64String($csvFileContent)

            # Convert the log file content to Base64
            $logFileContent = [System.IO.File]::ReadAllBytes($logFilePath)
            $logBase64Content = [Convert]::ToBase64String($logFileContent)

            # Define the HTTP trigger URL for the Power Automate flow
            $flowUrl = "https://prod-166.westus.logic.azure.com:443/workflows/5e172f6d92d24c6a995023362c53472f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7-7I6wW8ga9i3hSfzjP7-O_AFLNFmE-_cxCGt6g3f9A"

            # Prepare the body of the request
            $body = @{
                csvFileName    = "$csvFileName`-$username.csv"
                csvFileContent = $csvBase64Content
                logFileName    = "$logFileName`-$username.txt"
                logFileContent = $logBase64Content
            }

            # Convert the body to JSON
            $jsonBody = $body | ConvertTo-Json -Depth 3

            # Set the headers
            $headers = @{
                "Content-Type" = "application/json"
            }

            if ($online) {
                # Send the HTTP POST request to trigger the flow
                Invoke-RestMethod -Uri $flowUrl -Method Post -Headers $headers -Body $jsonBody
                Write-Host "Triggered Power Automate flow to upload the log files to SharePoint" -ForegroundColor Yellow
                Write-Host " "
            }
            else {
                # Send dummy write-host to emulate the real output
                Write-Host "Triggered Power Automate flow to upload the log files to SharePoint (OFFLINE - IGNORED)" -ForegroundColor Yellow
                Write-Host " "
            }
            
            # Print the list of logged on users
            if ($loggedOnUsers.Count -gt 0) {
                Write-Host "Logged on users:" -ForegroundColor Yellow
                $loggedOnUsers | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
                Show-EmailDrafts -loggedOnDevices $loggedOnDevices | Out-Null
            }
            else {
                Write-Host "No users were logged on to the renamed computers." -ForegroundColor Green
            }
            Write-Host " "

            # Remove successfully renamed computers from the list of computers
            foreach ($renameEntry in $script:validNamesList) {
                $oldName, $newName = $renameEntry -split ' -> '
                $script:checkedItems.Remove($oldName)
        
                # Remove the old name from the filteredComputers
                $script:filteredComputers = $script:filteredComputers | Where-Object { $_.Name -ne $oldName }
        
                # Try to remove the old name from the computerCheckedListBox
                $index = $computerCheckedListBox.Items.IndexOf($oldName)
                if ($index -ge 0) {
                    $computerCheckedListBox.Items.RemoveAt($index)
                }
            }
        
            # Refresh the checked list box based on current search term
            $searchTerm = $searchBox.Text
            $filteredList = $script:filteredComputers | Where-Object { $_.Name -like "*$searchTerm*" }
        
            # Clear and repopulate the checked list box with filtered computers and restore their checked state
            $computerCheckedListBox.Items.Clear()
            foreach ($computer in $filteredList) {
                $isChecked = $false
                if ($script:checkedItems.ContainsKey($computer.Name)) {
                    $isChecked = $script:checkedItems[$computer.Name]
                }
                $computerCheckedListBox.Items.Add($computer.Name, $isChecked)
            }
            # UpdateAllListBoxes
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host " "

        }
    })
$form.Controls.Add($applyRenameButton) 

# Call the function to load and filter computers
if ($online) {
    LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox | Out-Null
}
else {
    LoadAndFilterComputersOFFLINE -computerCheckedListBox $computerCheckedListBox | Out-Null
}
# Show the form
# $form.Add_Shown({ $form.Activate() })
$form.ShowDialog()
