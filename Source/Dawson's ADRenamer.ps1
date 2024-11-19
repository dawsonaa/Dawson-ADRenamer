<#
.SYNOPSIS
    A PowerShell script for renaming Active Directory computer objects based on user-defined rules and inputs.

.DESCRIPTION
    This script allows administrators to rename computer objects in Active Directory (AD) based on customizable patterns and input values.
    The script provides a graphical user interface (GUI) for selecting computers, specifying renaming patterns, and performing renaming operations.
    The script supports both online and offline modes, and handles various scenarios including checking for logged-on users, restarting computers,
    and logging the results of the renaming operations.

    Key functionalities include:
    - Loading and filtering computers from AD or an offline dataset.
    - Selecting computers and specifying renaming patterns.
    - Validating new names to ensure they meet naming conventions and length restrictions.
    - Grouping related changes by color for easier visualization.
    - Removing selected computers or groups of computers based on their assigned colors.
    - Generating logs and results, and triggering a Power Automate flow to upload these to SharePoint.
    - Converting department names to strings based on OU location, with truncation logic for specific naming patterns.
    - Supporting Ctrl+A for select all functionality in text boxes.
    - Dynamically updating context menu items based on the changes list.
    - Differentiating between online and offline modes for loading computer data and performing operations.

    Online Mode:
    - Queries AD for computer objects and their last logon dates.
    - Checks the online status of computers before renaming and restarting them.
    - Uses provided credentials for AD operations.

    Offline Mode:
    - Simulates computer data and online checks.
    - Does not perform actual AD operations but logs actions as if they were performed. (Other than sharepoint upload)

.PARAMETER None
    This script does not take any parameters.

.NOTES
    - The script uses a GUI built with Windows Forms.
    - The script includes functions for handling AD queries, user interactions, and renaming operations.
    - The script logs its operations and can upload results to SharePoint via a Power Automate flow.
    - The script is designed to handle edge cases such as duplicate names, invalid names, and offline computers.
    - The script dynamically updates context menu items based on the changes list.
    - The script supports converting department names to strings and truncating them based on specific patterns.
    - The script supports Ctrl+A for select all functionality in text boxes.

.EXAMPLE
    # Run the script
    .\Dawson's ADRenamer.ps1
#>
# All Campuses Device Naming Scheme KB: https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Version = "11.18.24"
$iconPath = Join-Path $PSScriptRoot "icon2.ico"

# COLOR PALETTE
$visualStudioBlue = [System.Drawing.Color]::FromArgb(78, 102, 221) # RGB values for Visual Studio blue
$kstateDarkPurple = [System.Drawing.Color]::FromArgb(64, 1, 111)
$kstateLightPurple = [System.Drawing.Color]::FromArgb(115, 25, 184)
$white = [System.Drawing.Color]::FromArgb(255, 255, 255)
$black = [System.Drawing.Color]::FromArgb(0, 0, 0)
$red = [System.Drawing.Color]::FromArgb(218, 25, 25)
$lightGray = [System.Drawing.Color]::FromArgb(45, 45, 45)
$grayBlue = [System.Drawing.Color]::FromArgb(75, 75, 140)
$catBlue = [System.Drawing.Color]::FromArgb(3, 106, 199)
$catPurple = [System.Drawing.Color]::FromArgb(170, 13, 206)
$catRed = [System.Drawing.Color]::FromArgb(255, 27, 82)
$catYellow = [System.Drawing.Color]::FromArgb(254, 162, 2)
$catLightYellow = [System.Drawing.Color]::FromArgb(255, 249, 227)
$catDark = [System.Drawing.Color]::FromArgb(1, 3, 32)

# Create a form for selecting Online, Offline, or Cancel
$modeSelectionForm = New-Object System.Windows.Forms.Form
$modeSelectionForm.Text = "Select ADR Mode"
$modeSelectionForm.Size = New-Object System.Drawing.Size(300,150)
$modeSelectionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$modeSelectionForm.MaximizeBox = $false
$modeSelectionForm.StartPosition = "CenterScreen"
$modeSelectionForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

$modeSelectionForm.BackColor = $lightGray
$modeSelectionForm.ForeColor = [System.Drawing.Color]::LightGray
$modeSelectionForm.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold

$label = New-Object System.Windows.Forms.Label
$label.Text = "Do you want to use ADRenamer in Online or Offline mode?"
$label.Size = New-Object System.Drawing.Size(280,30)
$label.Location = New-Object System.Drawing.Point(10,20)
$modeSelectionForm.Controls.Add($label)

$global:formClosedByButton = $false

$buttonOnline = New-Object System.Windows.Forms.Button
$buttonOnline.Text = "Online"
$buttonOnline.Location = New-Object System.Drawing.Point(10,70)
$buttonOnline.Size = New-Object System.Drawing.Size(75,30)
$buttonOnline.Add_Click({
    $global:choice = "Online"
    $global:formClosedByButton = $true
    $modeSelectionForm.Close()
})
$modeSelectionForm.Controls.Add($buttonOnline)

$buttonOffline = New-Object System.Windows.Forms.Button
$buttonOffline.Text = "Offline"
$buttonOffline.Location = New-Object System.Drawing.Point(100,70)
$buttonOffline.Size = New-Object System.Drawing.Size(75,30)
$buttonOffline.Add_Click({
    $global:choice = "Offline"
    $global:formClosedByButton = $true
    $modeSelectionForm.Close()
})
$modeSelectionForm.Controls.Add($buttonOffline)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Text = "Cancel"
$buttonCancel.Location = New-Object System.Drawing.Point(190,70)
$buttonCancel.Size = New-Object System.Drawing.Size(75,30)
$buttonCancel.Add_Click({
    $global:choice = "Cancel"
    $modeSelectionForm.Close()
})
$modeSelectionForm.Controls.Add($buttonCancel)

$modeSelectionForm.Add_FormClosing({
    param($sender, $e)
    if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and -not $global:formClosedByButton) {
        # Terminate the entire script because the form is closing due to the user pressing 'X'
        [Environment]::Exit(0)
    }
})

# Show the form
$modeSelectionForm.ShowDialog() | Out-Null

# Set the $online variable based on the selection or exit if canceled
switch ($global:choice) {
    "Online" {
        $online = $true
    }
    "Offline" {
        $online = $false
    }
    "Cancel" {
        exit
    }
}

if (-not $online) {
    # Dummy data for computers # OFFLINE
    function Add-DummyComputers {
        param (
            [int]$numberOfDevices = 10
        )
    
        # Function to generate a random date within the last year
        function Get-RandomDate {
            $randomDays = Get-Random -Minimum 1 -Maximum 200
            return (Get-Date).AddDays(-$randomDays)
        }
    
        # Generate dummy computers
        $dummyComputers = @()
        for ($i = 1; $i -le $numberOfDevices; $i++) {
            $dummyComputers += @{
                Name          = "HL-CS-$i"
                LastLogonDate = Get-RandomDate
            }
        }
    
        return $dummyComputers
    }
    
    # Call the function with the desired number of devices
    $numberOfDevices = 200
    $dummyComputers = Add-DummyComputers -numberOfDevices $numberOfDevices

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

if ($online) {
    # Present initial login window # ONLINE
    while (-not $connectionSuccess) {
        if ($errorMessage) {
            Write-Host $errorMessage -ForegroundColor Red
        }

        if ($username -ne "") {
            $cred = Get-Credential -Message "Please enter Active Directory admin credentials" -UserName $username
        }
        else {
            $cred = Get-Credential -Message "Please enter Active Directory admin credentials"
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
    Write-Host "ONLINE MODE - Sufficient Credentials Provided. Logged on as $username." -ForegroundColor Green
}else {
    Write-Host "OFFLINE MODE - No credentials are needed." -ForegroundColor Green
}

# Initialize a new hash set to store unique strings.
# This hash set will be used to ensure that new computer names are unique.
$hashSet = [System.Collections.Generic.HashSet[string]]::new()

# Change object to store different types of rename operations
class Change {
    [string[]]$ComputerNames
    [string]$Part0
    [string]$Part1
    [string]$Part2
    [string]$Part3
    [CustomColor]$GroupColor
    [bool[]]$Valid
    [string[]]$AttemptedNames
    [bool[]]$Duplicate

    Change([string[]]$computerNames, [string]$part0, [string]$part1, [string]$part2, [string]$part3, [CustomColor]$groupColor, [bool[]]$valid, [string[]]$attemptedNames, [bool[]]$duplicate) {
        $this.ComputerNames = $computerNames
        $this.Part0 = $part0
        $this.Part1 = $part1
        $this.Part2 = $part2
        $this.Part3 = $part3
        $this.GroupColor = $groupColor
        $this.Valid = $valid
        $this.AttemptedNames = $attemptedNames
        $this.Duplicate = $duplicate
    }
}

# Initialize
$script:changesList = New-Object System.Collections.ArrayList
$script:newNamesList = @()

# Define a custom class to represent RGB color
class CustomColor {
    [int]$R
    [int]$G
    [int]$B

    CustomColor([int]$r, [int]$g, [int]$b) {
        $this.R = $r
        $this.G = $g
        $this.B = $b
    }

    [string]ToString() {
        return "R: $($this.R), G: $($this.G), B: $($this.B)"
    }
}

# Define a list of unique colors for the items
$colors = @(
    [CustomColor]::new(243, 12, 122), # Vibrant Pink
    [CustomColor]::new(243, 120, 22), # Vibrant Orange
    #[CustomColor]::new(255, 223, 0), # Bright Yellow
    [CustomColor]::new(76, 175, 80), # Light Green
    [CustomColor]::new(0, 188, 212), # Cyan
    [CustomColor]::new(103, 58, 183)   # Deep Purple
)

# Initialize the global variable for the color index
$global:nextColorIndex = 0

# Function to get the department string
function Get-DepartmentString($deviceName, $part) {
    # Get the OU location of the device
    if ($online) {
        $device = Get-ADComputer -Identity $deviceName -Properties CanonicalName
        $ouLocation = $device.CanonicalName -replace "^CN=[^,]+,", ""

        # Extract the string directly after "Dept/"
        if ($ouLocation -match "Dept/([^/]+)") {
            $deptString = $matches[1]
        }
        else {
            $deptString = "deptnotfound"
        }
    }
    else {
        $ouLocation = "/Dept/OFFLN/Workstations/"
        if ($ouLocation -match "Dept/([^/]+)") {
            $deptString = $matches[1]
        }
        else {
            $deptString = "deptnotfound"
        }
    }

    # Apply truncation logic if part contains numbers before or after "dept"
    if ($part -match "(?i)(\d*)dept(\d*)") {
        $prefixLength = if ($matches[1] -and $matches[1] -ge 2 -and $matches[1] -le 5) { [int]::Parse($matches[1]) } else { $null }
        $suffixLength = if ($matches[2] -and $matches[2] -ge 2 -and $matches[2] -le 5) { [int]::Parse($matches[2]) } else { $null }

        if ($suffixLength) {
            # If the number is after "dept", truncate from the left
            return $deptString.Substring(0, [Math]::Min($deptString.Length, $suffixLength))
        }
        elseif ($prefixLength) {
            # If the number is before "dept", truncate from the right
            $startIndex = [Math]::Max(0, $deptString.Length - $prefixLength)
            return $deptString.Substring($startIndex, $prefixLength)
        }
    }

    return $deptString
}

<#
.SYNOPSIS
    Processes the committed changes for computer name renaming and updates the list boxes accordingly.

.DESCRIPTION
    This function processes the selected checked items, generates new names for each computer based on the input values and specified rules, 
    and updates the valid and invalid name lists. It also handles checking for duplicate names, updating the changes list, 
    and synchronizing the updated names with the relevant list boxes. It ensures that each computer name conforms to the maximum length of 15 characters,
    handles department-specific truncation logic, and updates the changes list by grouping related changes together based on their naming components.

.PARAMETER None
    This function does not take any parameters.

.NOTES
    - The function clears existing items and lists except for the new names list.
    - It tracks attempted names to identify duplicates.
    - It processes each selected checked item, splits the computer name into parts, and applies input values.
    - It checks if any part contains "dept" and replaces it with the department string.
    - It generates new names based on the parts and checks for validity and duplicates.
    - It updates the valid and invalid names lists accordingly.
    - It updates the changes list with the new names and synchronizes the updated names with the list boxes.
    - It handles color coding for different groups of changes.

.EXAMPLE
    ProcessCommittedChanges
    This command processes the committed changes for computer name renaming and updates the list boxes.

#>
function ProcessCommittedChanges {
    # Clear existing items and lists except for the new names list
    $hashSet.Clear()
    $script:newNamesListBox.Items.Clear()
    $script:validNamesList = @()
    $script:invalidNamesList = @()

    # Write-Host "Checked items: $($script:checkedItems.Keys -join ', ')"

    # Track attempted names to identify duplicates
    $attemptedNamesTracker = @{}

    # Process selected checked items
    foreach ($computerName in $script:selectedCheckedItems.Keys) {
        # Write-Host "Processing computer: $computerName"
        
        $parts = $computerName -split '-'
        $part0 = $parts[0]
        $part1 = $parts[1]
        $part2 = if ($parts.Count -ge 3) { $parts[2] } else { $null }
        $part3 = if ($parts.Count -ge 4) { $parts[3..($parts.Count - 1)] -join '-' } else { $null }

        # Write-Host "Initial parts: Part0: $part0, Part1: $part1, Part2: $part2, Part3: $part3" -ForegroundColor DarkBlue

        $part0InputValue = if (-not $part0Input.ReadOnly) { $part0Input.Text } else { $null }
        $part1InputValue = if (-not $part1Input.ReadOnly) { $part1Input.Text } else { $null }
        $part2InputValue = if (-not $part2Input.ReadOnly) { $part2Input.Text } else { $null }
        $part3InputValue = if (-not $part3Input.ReadOnly) { $part3Input.Text } else { $null }

        # Write-Host "Input values: Part0: $part0InputValue, Part1: $part1InputValue, Part2: $part2InputValue, Part3: $part3InputValue"

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

        # Check if any part contains "dept" (case insensitive) and replace with $deptString with truncation logic
        if ($part0 -match "(?i)dept") { $part0 = Get-DepartmentString $computerName $part0 }
        if ($part1 -match "(?i)dept") { $part1 = Get-DepartmentString $computerName $part1 }
        if ($part2 -match "(?i)dept") { $part2 = Get-DepartmentString $computerName $part2 }
        if ($part3 -match "(?i)dept") { $part3 = Get-DepartmentString $computerName $part3 }

        # Write-Host "Updated parts: Part0: $part0, Part1: $part1, Part2: $part2, Part3: $part3" -ForegroundColor DarkRed

        if ($part3) {
            $newName = "$part0-$part1-$part2-$part3"
        }
        elseif ($part2) {
            $newName = "$part0-$part1-$part2"
        }
        else {
            $newName = "$part0-$part1"
        }

        # Write-Host "New name: $newName"
        $isValid = $newName.Length -le 15
        $isDuplicate = $attemptedNamesTracker.ContainsKey($newName)
        
        if ($isValid -and -not $isDuplicate) {
            if ($hashSet.Add($newName)) {
                $script:validNamesList += "$computerName -> $newName"
                if (-not ($script:newNamesList | Where-Object { $_.ComputerName -eq $computerName })) {
                    $script:newNamesList += @{"ComputerName" = $computerName; "NewName" = $newName; "Custom" = $false }
                }
            }
        }
        else {
            $script:invalidNamesList += $computerName
        }

        # Add the new name to the attempted names list and track duplicates
        if (-not $attemptedNamesTracker.ContainsKey($newName)) {
            $attemptedNamesTracker[$newName] = 1
        }
        else {
            $attemptedNamesTracker[$newName]++
        }

        $attemptedNames = @($newName)
        $duplicate = @($isDuplicate)

        # Update the name map with the new attempted name
        UpdateNameMap -originalName $computerName -newName $newName

        # Check if an existing change matches
        $existingChange = $null
        foreach ($change in $script:changesList) {
            $part0Comparison = ($change.Part0 -eq $part0InputValue -or ([string]::IsNullOrEmpty($change.Part0) -and [string]::IsNullOrEmpty($part0InputValue)))
            $part1Comparison = ($change.Part1 -eq $part1InputValue -or ([string]::IsNullOrEmpty($change.Part1) -and [string]::IsNullOrEmpty($part1InputValue)))
            $part2Comparison = ($change.Part2 -eq $part2InputValue -or ([string]::IsNullOrEmpty($change.Part2) -and [string]::IsNullOrEmpty($part2InputValue)))
            $part3Comparison = ($change.Part3 -eq $part3InputValue -or ([string]::IsNullOrEmpty($change.Part3) -and [string]::IsNullOrEmpty($part3InputValue)))
            $validComparison = ($change.Valid -eq $isValid)

            if ($part0Comparison -and $part1Comparison -and $part2Comparison -and $part3Comparison -and $validComparison) {
                # Write-Host "Found matching change for parts: Part0: $($change.Part0), Part1: $($change.Part1), Part2: $($change.Part2), Part3: $($change.Part3), Valid: $($change.Valid)" -ForegroundColor DarkRed
                $existingChange = $change
                break
            }
        }

        # Create a temporary list to store changes that need to be removed
        $tempChangesToRemove = @()

        # Remove the computer name from any previous change entries if they exist
        foreach ($change in $script:changesList) {
            if ($change -ne $existingChange -and $change.ComputerNames -contains $computerName) {
                $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $computerName }

                # Mark the change for removal if no computer names are left
                if ($change.ComputerNames.Count -eq 0) {
                    $tempChangesToRemove += $change
                }
            }
        }

        # Remove the marked changes from the changesList
        foreach ($changeToRemove in $tempChangesToRemove) {
            $script:changesList.Remove($changeToRemove)
        }

        if ($existingChange) {
            if (-not ($existingChange.ComputerNames -contains $computerName)) {
                $existingChange.ComputerNames += $computerName
            }
            $existingChange.AttemptedNames += $newName
            $existingChange.Valid += $isValid
            $existingChange.Duplicate += $isDuplicate
        }
        else {
            # Assign a unique color to the new change
            $groupColor = if (-not $isValid) { [CustomColor]::new(255, 0, 0) } else { $colors[$global:nextColorIndex % $colors.Count] }
            $global:nextColorIndex++
            $newChange = [Change]::new(@($computerName), $part0InputValue, $part1InputValue, $part2InputValue, $part3InputValue, $groupColor, @($isValid), $attemptedNames, $duplicate)
            $script:changesList.Add($newChange) | Out-Null
        }
    }

    <# Print the changesList for debugging
    Write-Host "`nChanges List:" -ForegroundColor red
    foreach ($change in $script:changesList) {
        Write-Host "Change Parts: Part0: $($change.Part0), Part1: $($change.Part1), Part2: $($change.Part2), Part3: $($change.Part3), Valid: $($change.Valid)" -ForegroundColor DarkRed
        Write-Host "ComputerNames: $($change.ComputerNames -join ', ')"
        Write-Host "AttemptedNames: $($change.AttemptedNames -join ', ')"
        Write-Host "Duplicate: $($change.Duplicate -join ', ')"
    }

    # Update the colors in the selectedCheckedListBox and colorPanel
    UpdateColors #>
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

    # Helper function to replace '+' with '%20'
    function EncodeSpaces($string) {
        return [System.Uri]::EscapeDataString($string) -replace '\+', '%20'
    }

    # Construct the Outlook web URL for creating a draft
    $url = "https://outlook.office.com/mail/deeplink/compose?to=" + (EncodeSpaces($emailAddress)) + "&subject=" + (EncodeSpaces($subject)) + "&body=" + (EncodeSpaces($body))

    # Open the URL in the default browser
    Start-Process $url
    Write-Host "Draft email created for $emailAddress" -ForegroundColor Green
}


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

    $emailForm.BackColor = $catDark
    $emailForm.ForeColor = $white
    # $emailForm.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold
    
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

# Function to display a form with a TreeView control for selecting an Organizational Unit (OU)
function Select-OU {
    # Create and configure the form
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "Select Organizational Unit"
    $ouForm.Size = New-Object System.Drawing.Size(400, 590)
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $ouForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ouForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

    $ouForm.BackColor = $lightGray
    $ouForm.ForeColor = [System.Drawing.Color]::LightGray
    $ouForm.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold

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
    $treeView.BackColor = [System.Drawing.Color]::LightGray
    $treeView.Visible = $true

    # Add "OK"(selectedOU) button for OU selection
    $selectedOUButton = New-Object System.Windows.Forms.Button
    $selectedOUButton.Text = "No OU selected"
    $selectedOUButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $selectedOUButton.Enabled = $false
    $selectedOUButton.Size = New-Object System.Drawing.Size(190, 23)
    $selectedOUButton.Location = New-Object System.Drawing.Point(10, 520)
    $selectedOUButton.Add_Click({
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $ouForm.Close()
        })

    # Add "Cancel"(defaultOU) button for OU selection
    $defaultOUButton = New-Object System.Windows.Forms.Button
    $defaultOUButton.Text = "DC=users,DC=campus"
    $defaultOUButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $defaultOUButton.Size = New-Object System.Drawing.Size(165, 23)
    $defaultOUButton.Location = New-Object System.Drawing.Point(210, 520)
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

# Initialize the $script:filteredComputers variable as an empty array
$script:filteredComputers = @()

<#
.SYNOPSIS
    Loads and filters computer objects from Active Directory (AD) or offline data source and updates the provided CheckedListBox with the filtered computer names.

.DESCRIPTION
    This function loads computer objects from Active Directory (online mode) or from a simulated offline data source.
    It filters the computers based on their last logon date, excluding those that have been offline for more than 180 days.
    The filtered and sorted computer names are then added to the provided CheckedListBox.
    The function also includes progress tracking and user feedback during the loading process.

.PARAMETER computerCheckedListBox
    The CheckedListBox control to be updated with the filtered computer names.

.NOTES
    - The function disables the form while loading data to prevent user interaction.
    - It filters computers based on their last logon date, excluding those offline for more than 180 days.
    - It updates the CheckedListBox with the filtered and sorted computer names.
    - It includes progress tracking and user feedback.
    - It re-enables the form after loading is complete.
    - It handles both online and offline modes.

.EXAMPLE
    LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox
    This command loads and filters computer objects and updates the provided CheckedListBox with the filtered computer names.
#>
function LoadAndFilterComputers {
    param (
        [System.Windows.Forms.CheckedListBox]$computerCheckedListBox
    )

    try {

        if ($online) {
            # Show the OU selection form
            Select-OU | Out-Null
            if (-not $script:ouPath) {
                Write-Host "No OU selected, using DC=users,DC=campus"
                $script:ouPath = 'DC=users,DC=campus'
                return
            }
        }
        else {
            $script:ouPath = 'OU=Workstations,OU=KSUL,OU=Dept,DC=users,DC=campus'
        }

        Write-Host "Selected OU Path: $script:ouPath"  # Debug message

        # Disable the form while loading data to prevent user interaction
        $form.Enabled = $false
        Write-Host "Form disabled for loading..."

        if ($online) {
            Write-Host "Loading AD endpoints..."
        }
        else {
            Write-Host "Loaded 'endpoints'... OFFLINE"
        }

        # Initialize counters and timers for progress tracking
        $loadedCount = 0
        $deviceRefresh = 150
        $deviceTimer = 0

        # Define the cutoff date for filtering computers based on their last logon date
        $cutoffDate = (Get-Date).AddDays(-180)
        
        # Clear array
        $script:filteredComputers = @()

        if ($online) {
            # Query Active Directory for computers within the selected OU and retrieve their last logon date
            $computers = Get-ADComputer -Filter * -Properties LastLogonDate -SearchBase $script:ouPath

            # Filter and sort the computers alphanumerically
            $script:filteredComputers = $computers | Where-Object {
                $_.LastLogonDate -and
                [DateTime]::Parse($_.LastLogonDate) -ge $cutoffDate
            } | Sort-Object -Property Name

            $computerCount = $script:filteredComputers.Count
            $filteredOutCount = $computers.Count - $script:filteredComputers.Count

            $computerCheckedListBox.Items.Clear()

            $script:filteredComputers | ForEach-Object {
                $computerCheckedListBox.Items.Add($_.Name, $false) | Out-Null
                $loadedCount++
                $deviceTimer++
                # Write-Host "loaded:" $_.Name
                # Write-Host ""
                # Update the progress bar every 150 devices
                if ($deviceTimer -ge $deviceRefresh) {
                    $deviceTimer = 0
                    $progress = [math]::Round(($loadedCount / $computerCount) * 100)
                    Write-Progress -Activity "Loading endpoints from AD..." -Status "$progress% Complete:" -PercentComplete $progress
                }
            }
        }
        else {
            # Simulate offline data
            $script:filteredComputers = $dummyComputers | Where-Object {
                $_.LastLogonDate -and
                [DateTime]::Parse($_.LastLogonDate) -ge $cutoffDate
            } | Sort-Object -Property Name

            $computerCount = $script:filteredComputers.Count
            $filteredOutCount = $dummyComputers.Count - $script:filteredComputers.Count

            $computerCheckedListBox.Items.Clear()

            # Populate the CheckedListBox with the filtered and sorted computer names
            $script:filteredComputers | ForEach-Object {
                $computerCheckedListBox.Items.Add($_.Name, $false) | Out-Null
                $loadedCount++
                $deviceTimer++

                # Update the progress bar every 150 devices
                if ($deviceTimer -ge $deviceRefresh) {
                    $deviceTimer = 0
                    $progress = [math]::Round(($loadedCount / $computerCount) * 100)
                    Write-Progress -Activity "Loading endpoints from AD (offline mode)..." -Status "$progress% Complete:" -PercentComplete $progress
                }
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
        exit
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(830, 490) # 785, 520
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

$form.BackColor = $lightGray
$form.ForeColor = [System.Drawing.Color]::LightGray
$form.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold

# Make sure user knows what mode they are in
if ($online) {
    $form.Text = "ADRenamer - Online"
}
else {
    $form.Text = "ADRenamer - Offline"
}

# Initialize script-scope variables
$script:invalidNamesList = @()
$script:validNamesList = @()
$script:customNamesList = @() 
$script:ouPath = 'DC=users,DC=campus'

# Create label to display current script version
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "Version $Version"
$versionLabel.Location = New-Object System.Drawing.Point(700, 430)
$versionLabel.AutoSize = $true
$versionLabel.Cursor = [System.Windows.Forms.Cursors]::Hand  # Change cursor to hand to indicate it's clickable

# Add click event handler to open the URL
$versionLabel.Add_Click({
        Start-Process "https://github.com/dawsonaa/Dawson-ADRenamer"
    })

$form.Controls.Add($versionLabel)

# Create label to display author information
$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Text = "Author: Dawson Adams (dawsonaa@ksu.edu)"
$authorLabel.Location = New-Object System.Drawing.Point(10, 430)
$authorLabel.AutoSize = $true
$authorLabel.Cursor = [System.Windows.Forms.Cursors]::Hand  # Change cursor to hand to indicate it's clickable

# Add click event handler to open Microsoft Teams chat with the specified email
$authorLabel.Add_Click({
        Start-Process "msteams://teams.microsoft.com/l/chat/0/0?users=dawsonaa@ksu.edu"
    })

$form.Controls.Add($authorLabel)

# Define the size for the list boxes
$listBoxWidth = 250
$listBoxHeight = 350

# Define the script-wide variables
$script:checkedItems = @{}
$script:selectedCheckedItems = @{}

<#
.SYNOPSIS
    Updates and synchronizes the list boxes with the latest changes in computer names.

.DESCRIPTION
    This function updates the new names list box and the selected checked list box with the latest computer name changes.
    It processes both valid and invalid items, groups them by change groups, and assigns colors to each group. 
    The function ensures that the list boxes are updated and refreshed to reflect the current state of the changes.

.PARAMETER None
    This function does not take any parameters.

.NOTES
    - The function clears existing items and lists except for the new names list.
    - It groups invalid and valid items by their change groups.
    - It assigns colors to each group for visual differentiation.
    - It processes checked items not in the changes list separately.
    - It updates and refreshes the new names list box and the selected checked list box with the latest changes.

.EXAMPLE
    UpdateAndSyncListBoxes
    This command updates and synchronizes the list boxes with the latest changes in computer names.

#>
function UpdateAndSyncListBoxes {
    # Write-Host "Updating and Syncing ListBoxes" -ForegroundColor Cyan
    $script:newNamesListBox.Items.Clear()
    $selectedCheckedListBox.Items.Clear()
    $itemsToRemove = New-Object System.Collections.ArrayList

    # Clear script:selectedItems
    $script:selectedCheckedItems.Clear()

    # Create collections for invalid and valid items grouped by change group
    $groupedInvalidItems = @{}
    $groupedValidItems = @{}
    $groupColors = @{}

    # Process changesList and sort items into valid and invalid groups
    # Write-Host "Processing changesList..." -ForegroundColor Green
    $groupIndex = 0
    foreach ($change in $script:changesList) {
        $groupName = "$($change.Part0)-$($change.Part1)-$($change.Part2)-$($change.Part3)"
        if (-not $groupedInvalidItems.ContainsKey($groupName)) {
            $groupedInvalidItems[$groupName] = New-Object System.Collections.ArrayList
            $groupedValidItems[$groupName] = New-Object System.Collections.ArrayList
            $groupColors[$groupName] = $colors[$groupIndex % $colors.Count]
            $groupIndex++
        }
        $sortedComputerNames = $change.ComputerNames | Sort-Object
        foreach ($computerName in $sortedComputerNames) {
            $index = [array]::IndexOf($change.ComputerNames, $computerName)
            $isValid = $change.Valid[$index]
            $isDuplicate = $change.Duplicate[$index]
            if ($isValid -and -not $isDuplicate) {
                $groupedValidItems[$groupName].Add($computerName) | Out-Null
            }
            else {
                $groupedInvalidItems[$groupName].Add($computerName) | Out-Null
            }
        }
    }

    # Process checkedItems not in changesList
    # Write-Host "Processing checkedItems not in changesList..." -ForegroundColor Blue
    $nonChangeItems = New-Object System.Collections.ArrayList
    foreach ($item in $script:checkedItems.Keys) {
        $isInChangeList = $false
        foreach ($change in $script:changesList) {
            if ($change.ComputerNames -contains $item) {
                $isInChangeList = $true
                break
            }
        }
        if (-not $isInChangeList) {
            $nonChangeItems.Add($item) | Out-Null
        }
    }

    # Sort the non-change items alphanumerically
    # Write-Host "Sorting non-change items..." -ForegroundColor Magenta
    $sortedNonChangeItems = $nonChangeItems | Sort-Object

    # Update both ListBoxes
    # Write-Host "Updating ListBoxes..." -ForegroundColor Yellow
    $script:newNamesListBox.BeginUpdate()
    $selectedCheckedListBox.BeginUpdate()

    foreach ($group in $groupedInvalidItems.Keys) {
        # Set group color
        # $color = $groupColors[$group]
        # Write-Host "Group: $group, Color: $color" -ForegroundColor Green

        # Add invalid items with "- Invalid" or "- Duplicate" suffix
        foreach ($item in ($groupedInvalidItems[$group] | Sort-Object)) {
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $item }
            $index = [array]::IndexOf($change.ComputerNames, $item)
            $newName = if ($change.Duplicate[$index]) {
                $item + " - Duplicate"
            }
            else {
                $change.AttemptedNames[$index] + " - Invalid"
            }
            $script:newNamesListBox.Items.Add($newName) | Out-Null
            $selectedCheckedListBox.Items.Add($item) | Out-Null
        }

        # Add valid items
        foreach ($item in ($groupedValidItems[$group] | Sort-Object)) {
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $item }
            $index = [array]::IndexOf($change.ComputerNames, $item)
            $newName = $change.AttemptedNames[$index]
            $script:newNamesListBox.Items.Add($newName) | Out-Null
            $selectedCheckedListBox.Items.Add($item) | Out-Null
        }
    }

    # Add non-change items
    foreach ($item in $sortedNonChangeItems) {
        $script:newNamesListBox.Items.Add($item) | Out-Null
        $selectedCheckedListBox.Items.Add($item) | Out-Null
    }

    # Remove items marked for removal
    # Write-Host "Removing items marked for removal..." -ForegroundColor Red
    foreach ($item in $itemsToRemove) {
        # Write-Host "Removing $item from newNamesListBox" -ForegroundColor Red
        $script:newNamesListBox.Items.Remove($item)
    }

    $script:newNamesListBox.EndUpdate()
    $selectedCheckedListBox.EndUpdate()

    # Force a refresh of the ListBox controls
    $script:newNamesListBox.Refresh()
    $selectedCheckedListBox.Refresh()
}


# Create checked list box for computers
$computerCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$computerCheckedListBox.Location = New-Object System.Drawing.Point(10, 10)
$computerCheckedListBox.Size = New-Object System.Drawing.Size($listBoxWidth, $listBoxHeight)
$computerCheckedListBox.IntegralHeight = $false
$computerCheckedListBox.BackColor = [System.Drawing.Color]::LightGray
$computerCheckedListBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold) # Arial, 10pt, Bold

$script:computerCtrlA = 1

# Handle the KeyDown event to detect Ctrl+A
$computerCheckedListBox.Add_KeyDown({
        param($s, $e)

        # Check if Ctrl+A is pressed
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            # Disable the form and controls to prevent interactions
            $form.Enabled = $false
            Write-Host "Ctrl+A pressed, toggling selection state, Form disabled for loading..."

            # Limit the number of items to select/unselect to 500
            $maxItemsToCheck = 500
            $itemsProcessed = 0

            if ($script:computerCtrlA -eq 1) {
                # Select all items up to the limit
                for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
                    if ($itemsProcessed -ge $maxItemsToCheck) {
                        break
                    }
                    $computerCheckedListBox.SetItemChecked($i, $true)
                    $currentItem = $computerCheckedListBox.Items[$i]
                    $script:checkedItems[$currentItem] = $true
                    $itemsProcessed++
                }
                $script:computerCtrlA = 0  # Set to unselect next time
                Write-Host "All items selected"
            }
            else {
                # Unselect all items up to the limit
                for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
                    if ($itemsProcessed -ge $maxItemsToCheck) {
                        break
                    }
                    $computerCheckedListBox.SetItemChecked($i, $false)
                    $currentItem = $computerCheckedListBox.Items[$i]
                    $script:checkedItems.Remove($currentItem)
                    $newNamesListBox.Items.Clear()
                    $itemsProcessed++
                }
                $script:computerCtrlA = 1  # Set to select next time
                Write-Host "All items unselected"
            }

            # Update the selectedCheckedListBox with sorted items
            $selectedCheckedListBox.BeginUpdate()
            $selectedCheckedListBox.Items.Clear()
            $sortedCheckedItems = $script:checkedItems.Keys | Sort-Object
            foreach ($item in $sortedCheckedItems) {
                $selectedCheckedListBox.Items.Add($item, $script:checkedItems[$item]) | Out-Null
            }
            $selectedCheckedListBox.EndUpdate()

            # Enable the form and controls
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host ""

            # Prevent default action
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

# Initialize a dictionary to map the original computer name to the new attempted name
$script:originalToNewNameMap = @{}

# Function to update the dictionary when names change
function UpdateNameMap {
    param (
        [string]$originalName,
        [string]$newName
    )

    if ($newName) {
        $script:originalToNewNameMap[$originalName] = $newName
    }
    else {
        $script:originalToNewNameMap.Remove($originalName)
    }
}

# Handle the ItemCheck event to update script:checkedItems
$computerCheckedListBox.add_ItemCheck({
        param($s, $e)

        $item = $computerCheckedListBox.Items[$e.Index]

        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            # Add the item to script:checkedItems if checked
            if (-not $script:checkedItems.ContainsKey($item)) {
                $script:checkedItems[$item] = $true
                # Add the item directly to the selectedCheckedListBox
                $selectedCheckedListBox.Items.Add($item, $true)
                # Write-Host "Item added: $item" -ForegroundColor Green
            }
        }
        elseif ($e.NewValue -eq [System.Windows.Forms.CheckState]::Unchecked) {
            # Remove the item from script:checkedItems if unchecked
            if ($script:checkedItems.ContainsKey($item)) {
                $script:checkedItems.Remove($item)

                # Get the new name from the map
                $newName = $script:originalToNewNameMap[$item]

                # Temporarily disable updates to the list boxes
                $selectedCheckedListBox.BeginUpdate()
                $newNamesListBox.BeginUpdate()

                # Remove the item from the selectedCheckedListBox and newNamesListBox
                $selectedCheckedListBox.Items.Remove($item)
                if ($newName) {
                    $newNamesListBox.Items.Remove($newName)
                }

                # Re-enable updates to the list boxes
                $selectedCheckedListBox.EndUpdate()
                $newNamesListBox.EndUpdate()

                # Remove the item from the map
                $script:originalToNewNameMap.Remove($item)

                # Write-Host "Item removed: $item" -ForegroundColor Red
            }

            # Remove the item from changesList
            $tempChangesToRemove = @()
            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    # Mark the change for removal if no computer names are left
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }

            # Remove the marked changes from the changesList after iteration
            foreach ($changeToRemove in $tempChangesToRemove) {
                $script:changesList.Remove($changeToRemove)
            }
        }
    })


# Attach the event handler to the CheckedListBox

$form.Controls.Add($computerCheckedListBox)

# Create a new checked list box for displaying selected computers
$selectedCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$selectedCheckedListBox.Location = New-Object System.Drawing.Point(260, 10)
$selectedCheckedListBox.Size = New-Object System.Drawing.Size(($listBoxWidth + 20), ($listBoxHeight))
$selectedCheckedListBox.IntegralHeight = $false
$selectedCheckedListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawVariable
$selectedCheckedListBox.BackColor = [System.Drawing.Color]::LightGray

$script:selectedCtrlA = 1

# Handle the KeyDown event to implement Ctrl+A select all
$selectedCheckedListBox.add_KeyDown({
        param($s, $e)
        # Check if Ctrl+A was pressed
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            if ($script:selectedCtrlA -eq 0) {
                for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
                    $selectedCheckedListBox.SetItemChecked($i, $true)
                }
                $script:selectedCtrlA = 1
            }
            else {
                for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
                    $selectedCheckedListBox.SetItemChecked($i, $false)
                }
                $script:selectedCtrlA = 0
            }
            $e.Handled = $true
        }
    })

# Handle the ItemCheck event to update selectedCheckedItems
$selectedCheckedListBox.add_ItemCheck({
        param($s, $e)

        $item = $selectedCheckedListBox.Items[$e.Index]
    
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            # Add the item to script:selectedCheckedItems if checked
            if (-not $script:selectedCheckedItems.ContainsKey($item)) {
                $script:selectedCheckedItems[$item] = $true
                # Write-Host "Item added: $item" -ForegroundColor Green
            }
            $colorPanel3.Invalidate()
            $colorPanel.Invalidate()
            $colorPanel2.Invalidate()
        }
        elseif ($e.NewValue -eq [System.Windows.Forms.CheckState]::Unchecked) {
            # Remove the item from script:selectedCheckedItems if unchecked
            if ($script:selectedCheckedItems.ContainsKey($item)) {
                $script:selectedCheckedItems.Remove($item)
                # Write-Host "Item removed: $item" -ForegroundColor Red
            }
            $colorPanel3.Invalidate()
            $colorPanel.Invalidate()
            $colorPanel2.Invalidate()
        }

        if (($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) -and $selectedCheckedListBox.CheckedItems.Count -gt 0) {
            $commitChangesButton.Enabled = $true
        }
        else {
            $commitChangesButton.Enabled = $false
        }  
    })


# Handle the MeasureItem event to set the item height
$selectedCheckedListBox.add_MeasureItem({
        param ($s, $e)
        $e.ItemHeight = 20
    })

# Updates color panels with change
function UpdateColors {
    $selectedCheckedListBox.Invalidate()
    $colorPanel3.Invalidate()
    $colorPanel.Invalidate()
    $colorPanel2.Invalidate()
}

<#
.SYNOPSIS
    Updates the SelectedCheckedListBox with sorted items from the changes list and other checked items.

.DESCRIPTION
    This function updates the SelectedCheckedListBox by first adding items from the changes list, sorted alphanumerically within groups,
    and then adding items not in any changes list group, also sorted alphanumerically. It preserves the checked state of items during the update process.

.PARAMETER None
    This function does not take any parameters.

.NOTES
    - The function initializes two collections: one for items from the changes list and another for non-change items.
    - It sorts and combines these collections to update the SelectedCheckedListBox.
    - It preserves and restores the checked state of items during the update.
    - It ensures the list is updated without flicker by using BeginUpdate and EndUpdate methods.
    - It calls UpdateAndSyncListBoxes to synchronize the updated list with other relevant controls.

.EXAMPLE
    UpdateSelectedCheckedListBox
    This command updates the SelectedCheckedListBox with sorted items from the changes list and other checked items.
#>
function UpdateSelectedCheckedListBox {
    #Write-Host "UPDATESELECTED" -ForegroundColor Cyan
    $sortedItems = New-Object System.Collections.ArrayList
    $nonChangeItems = New-Object System.Collections.ArrayList

    # Add items from changesList first, sorted alphanumerically within groups
    foreach ($change in $script:changesList) {
        $sortedComputerNames = $change.ComputerNames | Sort-Object
        foreach ($computerName in $sortedComputerNames) {
            $sortedItems.Add($computerName) | Out-Null
        }
    }

    # Add items not in any changesList group
    foreach ($item in $script:checkedItems.Keys) {
        $isInChangeList = $false
        foreach ($change in $script:changesList) {
            if ($change.ComputerNames -contains $item) {
                $isInChangeList = $true
                break
            }
        }
        if (-not $isInChangeList) {
            # Write-Host "Adding non-change item: $item" -ForegroundColor Yellow
            $nonChangeItems.Add($item) | Out-Null
        }
        <#
        else {
            Write-Host "Item in change list, skipping: $item" -ForegroundColor Green
        } #>
    }

    # Sort the non-change items alphanumerically
    $sortedNonChangeItems = $nonChangeItems | Sort-Object

    # Debugging: Print non-change items before sorting
    <# Write-Host "`nNon-change items before sorting:"
    foreach ($item in $nonChangeItems) {
        Write-Host $item
    } #>

    # Debugging: Print non-change items after sorting
    <# Write-Host "`nNon-change items after sorting:"
    foreach ($item in $sortedNonChangeItems) {
        Write-Host $item
    } #>

    # Combine the sorted change items and sorted non-change items
    if ($sortedNonChangeItems.Count -gt 0) {
        $sortedItems.AddRange($sortedNonChangeItems)
    }

    # Preserve the checked state
    $checkedItems = @{}
    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        if ($selectedCheckedListBox.GetItemChecked($i)) {
            $checkedItems[$selectedCheckedListBox.Items[$i]] = $true
        }
    }

    # Update the CheckedListBox
    $selectedCheckedListBox.BeginUpdate()
    $selectedCheckedListBox.Items.Clear()
    foreach ($item in $sortedItems) {
        $selectedCheckedListBox.Items.Add($item) | Out-Null
    }
    $selectedCheckedListBox.EndUpdate()

    # Restore the checked state
    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        if ($checkedItems.ContainsKey($selectedCheckedListBox.Items[$i])) {
            $selectedCheckedListBox.SetItemChecked($i, $true)
        }
    }

    # Print the items in the selectedCheckedListBox
    <# Write-Host "`nSelectedCheckedListBox Items in Order:"
    foreach ($item in $selectedCheckedListBox.Items) {
        Write-Host $item
    } #>

    UpdateAndSyncListBoxes
}

# Handle the DrawItem event to customize item drawing
$selectedCheckedListBox.add_DrawItem({
        param ($s, $e)
        $index = $e.Index
        if ($index -lt 0) { return }

        $itemText = $selectedCheckedListBox.Items[$index]
        $change = $script:changesList | Where-Object { $_.ComputerNames -contains $itemText }
        $backgroundColor = if ($change) { $change.GroupColor } else { [System.Drawing.Color]::White }

        $e.Graphics.FillRectangle([System.Drawing.SolidBrush]::new($backgroundColor), $e.Bounds)
        $layoutRectangle = [System.Drawing.RectangleF]::new($e.Bounds.X, $e.Bounds.Y, $e.Bounds.Width, $e.Bounds.Height)
        $e.Graphics.DrawString($itemText, $e.Font, [System.Drawing.SystemBrushes]::WindowText, $layoutRectangle)

        if (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0) {
            $e.DrawFocusRectangle()
        }
    })

$colorPanelBack = $catDark

# Create a Panel to show the colors next to the CheckedListBox
$colorPanel3 = New-Object System.Windows.Forms.Panel
$colorPanel3.Location = New-Object System.Drawing.Point(240, 10) # 530, 40
$colorPanel3.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel3.AutoScroll = $true
$colorPanel3.BackColor = $lightGray

# Create a Panel to show the colors next to the CheckedListBox
$colorPanel = New-Object System.Windows.Forms.Panel
$colorPanel.Location = New-Object System.Drawing.Point(510, 10) # 260, 40
$colorPanel.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel.BackColor = $lightGray

# Create a Panel to show the colors next to the CheckedListBox
$colorPanel2 = New-Object System.Windows.Forms.Panel
$colorPanel2.Location = New-Object System.Drawing.Point(780, 10) # 530, 40
$colorPanel2.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel2.BackColor = $lightGray

# Handle the Paint event for the color panel
$colorPanel.add_Paint({
        param ($s, $e)
        $visibleItems = [Math]::Ceiling($selectedCheckedListBox.ClientRectangle.Height / $selectedCheckedListBox.ItemHeight)
        $firstVisibleIndex = [Math]::Ceiling($script:globalTopIndex)
        $y = 0
        for ($i = $firstVisibleIndex; $i -lt ($firstVisibleIndex + $visibleItems); $i++) {
            if ($i -ge $selectedCheckedListBox.Items.Count) { break }
            $itemText = $selectedCheckedListBox.Items[$i]
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $itemText }
            $backgroundColor = if ($change) { [System.Drawing.Color]::FromArgb($change.GroupColor.R, $change.GroupColor.G, $change.GroupColor.B) } else { $colorPanelBack }
            $e.Graphics.FillRectangle([System.Drawing.SolidBrush]::new($backgroundColor), 0, $y, $colorPanel.Width, $selectedCheckedListBox.ItemHeight)
            $y += $selectedCheckedListBox.ItemHeight
        }
    })

# Handle the Paint event for the color panel
$colorPanel2.add_Paint({
        param ($s, $e)
        $visibleItems = [Math]::Ceiling($selectedCheckedListBox.ClientRectangle.Height / $selectedCheckedListBox.ItemHeight)
        $firstVisibleIndex = [Math]::Ceiling($script:globalTopIndex)
        $y = 0
        for ($i = $firstVisibleIndex; $i -lt ($firstVisibleIndex + $visibleItems); $i++) {
            if ($i -ge $selectedCheckedListBox.Items.Count) { break }
            $itemText = $selectedCheckedListBox.Items[$i]
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $itemText }
            $backgroundColor = if ($change) { [System.Drawing.Color]::FromArgb($change.GroupColor.R, $change.GroupColor.G, $change.GroupColor.B) } else { $colorPanelBack }
            $e.Graphics.FillRectangle([System.Drawing.SolidBrush]::new($backgroundColor), 0, $y, $colorPanel2.Width, $selectedCheckedListBox.ItemHeight)
            $y += $selectedCheckedListBox.ItemHeight
        }
    })

# Handle the Paint event for the color panel
$colorPanel3.add_Paint({
        param ($s, $e)
        $visibleItems = [Math]::Ceiling($selectedCheckedListBox.ClientRectangle.Height / $selectedCheckedListBox.ItemHeight)
        $firstVisibleIndex = [Math]::Ceiling($script:globalTopIndex)
        $y = 0
        for ($i = $firstVisibleIndex; $i -lt ($firstVisibleIndex + $visibleItems); $i++) {
            if ($i -ge $selectedCheckedListBox.Items.Count) { break }
            $itemText = $selectedCheckedListBox.Items[$i]
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $itemText }
            $backgroundColor = if ($change) { [System.Drawing.Color]::FromArgb($change.GroupColor.R, $change.GroupColor.G, $change.GroupColor.B) } else { $colorPanelBack }
            $e.Graphics.FillRectangle([System.Drawing.SolidBrush]::new($backgroundColor), 0, $y, $colorPanel3.Width, $selectedCheckedListBox.ItemHeight)
            $y += $selectedCheckedListBox.ItemHeight
        }
    })

# Create a list box for displaying proposed new names
$newNamesListBox = New-Object System.Windows.Forms.ListBox
$newNamesListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawVariable
$newNamesListBox.Location = New-Object System.Drawing.Point(530, 10)
$newNamesListBox.Size = New-Object System.Drawing.Size(($listBoxWidth + 20), ($listBoxHeight))
$newNamesListBox.IntegralHeight = $false
$newNamesListBox.BackColor = [System.Drawing.Color]::LightGray

# Define the MeasureItem event handler
$measureItemHandler = {
    param (
        [object]$s,
        [System.Windows.Forms.MeasureItemEventArgs]$e
    )
    # Set the item height to a custom value (e.g., 30 pixels)
    $e.ItemHeight = 18
}

# Define the DrawItem event handler
$drawItemHandler = {
    param (
        [object]$s,
        [System.Windows.Forms.DrawItemEventArgs]$e
    )

    # Ensure the index is valid
    if ($e.Index -ge 0) {
        # Get the item text
        $itemText = $s.Items[$e.Index]

        # Draw the background
        $e.DrawBackground()

        # Draw the item text
        $textBrush = [System.Drawing.SolidBrush]::new($e.ForeColor)
        $pointF = [System.Drawing.PointF]::new($e.Bounds.X, $e.Bounds.Y)
        $e.Graphics.DrawString($itemText, $e.Font, $textBrush, $pointF)

        # Draw the focus rectangle if the ListBox has focus
        $e.DrawFocusRectangle()
    }
}

# Attach the event handlers
$newNamesListBox.add_MeasureItem($measureItemHandler)
$newNamesListBox.add_DrawItem($drawItemHandler)

# Override the selection behavior to prevent selection
$newNamesListBox.add_SelectedIndexChanged({
        $newNamesListBox.ClearSelected()
    })
$form.Controls.Add($newNamesListBox)

# Script-wide variable to store the current TopIndex
$script:globalTopIndex = 0

# Function to update list boxes based on the global TopIndex
function Update-ListBoxes {
    param ($topIndex)
    if ($topIndex -ge 0 -and $topIndex -lt $selectedCheckedListBox.Items.Count) {
        $selectedCheckedListBox.TopIndex = $topIndex
        $newNamesListBox.TopIndex = $topIndex
    }
}

# Disable default scrolling for the list boxes
$selectedCheckedListBox.add_MouseWheel({
        param ($s, $e)
        $e.Handled = $true
    })

$newNamesListBox.add_MouseWheel({
        param ($s, $e)
        $e.Handled = $true
    })

# Handle the MouseWheel event for the CheckedListBox to act as a scrollbar
$selectedCheckedListBox.add_MouseWheel({
        param ($s, $e)
        $delta = [math]::Sign($e.Delta)
        if ($delta -eq 1) {
            $script:globalTopIndex -= 1
        }
        elseif ($delta -eq -1) {
            $script:globalTopIndex += 1
        }
        $script:globalTopIndex = [Math]::Max(0, [Math]::Min($script:globalTopIndex, $selectedCheckedListBox.Items.Count - 19))
        Update-ListBoxes -topIndex $script:globalTopIndex
        $colorPanel3.Invalidate()
        $colorPanel.Invalidate()
        $colorPanel2.Invalidate()
    })

$newNamesListBox.add_MouseWheel({
        param ($s, $e)
        $delta = [math]::Sign($e.Delta)
        if ($delta -eq 1) {
            $script:globalTopIndex -= 1
        }
        elseif ($delta -eq -1) {
            $script:globalTopIndex += 1
        }
        $script:globalTopIndex = [Math]::Max(0, [Math]::Min($script:globalTopIndex, $newNamesListBox.Items.Count - 19))
        Update-ListBoxes -topIndex $script:globalTopIndex
        $colorPanel3.Invalidate()
        $colorPanel.Invalidate()
        $colorPanel2.Invalidate()
    })

# Handle the SelectedIndexChanged event to update the panel colors
$selectedCheckedListBox.add_SelectedIndexChanged({
        param ($s, $e)
        $colorPanel.Invalidate()
        $colorPanel2.Invalidate()
        $colorPanel3.Invalidate()
    })

# Create the context menu for right-click actions
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Create menu context item for selecting specific devices to remove
$menuRemove = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemove.Text = "Remove selected device(s)"
$menuRemove.Add_Click({
        # Update the script-wide variable with currently checked items
        $selectedItems = @($selectedCheckedListBox.CheckedItems | ForEach-Object { $_ })

        if (!($selectedItems.Count -gt 0)) {
            Write-Host "No devices selected"
            return
        }
        Write-Host ""

        # Create a temporary list to store changes that need to be removed
        $tempChangesToRemove = @()

        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            # Remove the item from the selectedCheckedListBox and newNamesListBox
            $selectedCheckedListBox.Items.Remove($item)
            $newName = $script:originalToNewNameMap[$item]
            if ($newName) {
                $newNamesListBox.Items.Remove($newName)
            }

            # Remove the item from changesList
            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }
        }

        # Remove the marked changes from the changesList after iteration
        foreach ($changeToRemove in $tempChangesToRemove) {
            $script:changesList.Remove($changeToRemove)
        }

        Write-Host "Selected device(s) removed: $($selectedItems -join ', ')"  # Outputs the names of selected devices to the console
        Write-Host ""
    })

# Create menu context item for removing all devices within the selectedCheckedListBox
$menuRemoveAll = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemoveAll.Text = "Remove all device(s)"
$menuRemoveAll.Enabled = $false
$menuRemoveAll.Add_Click({
        $selectedItems = @($selectedCheckedListBox.Items | ForEach-Object { $_ })
        $form.Enabled = $false
        Write-Host "Form disabled for remove all"
        $script:customNamesList = @()

        # Create a temporary list to store changes that need to be removed
        $tempChangesToRemove = @()

        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            # Try to uncheck the selected items from the computerCheckedListBox
            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            # Remove the item from the selectedCheckedListBox and newNamesListBox
            $selectedCheckedListBox.Items.Remove($item)
            $newName = $script:originalToNewNameMap[$item]
            if ($newName) {
                $newNamesListBox.Items.Remove($newName)
            }

            # Remove the item from changesList
            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }
        }

        # Remove the marked changes from the changesList after iteration
        foreach ($changeToRemove in $tempChangesToRemove) {
            $script:changesList.Remove($changeToRemove)
        }

        $form.Enabled = $true
        Write-Host "Form enabled"
        Write-Host "All devices removed from the list"
    })

# Event handler for when the context menu is opening
$contextMenu.add_Opening({
        # Check if there are any items
        $selectedItems = @($selectedCheckedListBox.CheckedItems)
        $itemsInBox = @($selectedCheckedListBox.Items)

        if ($itemsInBox.Count -gt 0) {
            $menuRemoveAll.Enabled = $true
        }
        else {
            $menuRemoveAll.Enabled = $false
        }

        if ($selectedItems.Count -gt 0) {
            $menuRemove.Enabled = $true  # Enable the menu item if items are checked

 
        }
        else {
            $menuRemove.Enabled = $false  # Disable the menu item if no items are checked
        }
    })

# Add the key down event handler to selectedCheckedListBox
$selectedCheckedListBox.add_KeyDown({
        param ($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $menuFindAndReplace.PerformClick()
            $e.Handled = $true
        }
    })


# Add the right click menu options to the context menu
$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemove) | Out-Null
$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemoveAll) | Out-Null

# Attach the context menu to the CheckedListBox
$selectedCheckedListBox.ContextMenuStrip = $contextMenu
$form.Controls.Add($selectedCheckedListBox)


$form.Controls.Add($colorPanel3)
$form.Controls.Add($colorPanel)
$form.Controls.Add($colorPanel2)

$colorPanel3.BringToFront()
$colorPanel.BringToFront()
$colorPanel2.BringToFront()

# Search Text Box with Enter Key Event
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 365)
$searchBox.Size = New-Object System.Drawing.Size(133, 20)
$searchBox.ForeColor = [System.Drawing.Color]::Gray
$searchBox.BackColor = [System.Drawing.Color]::LightGray
$searchBox.Text = "Search"
$searchBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
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
        if ($this.Text -eq "Search") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Restore placeholder text when the text box loses focus and is empty
$searchBox.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Search"
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

<#
.SYNOPSIS
    Creates a custom TextBox control with specified properties and behavior.

.DESCRIPTION
    This function creates a custom TextBox control with specified properties such as name, default text, location, size, and maximum length.
    The TextBox is initialized with a default appearance and behavior, including read-only mode, color settings, and event handlers for mouse and keyboard interactions.
    The TextBox becomes editable and changes color when clicked, and supports Ctrl+A for selecting all text.

.PARAMETER name
    The name of the TextBox control.

.PARAMETER defaultText
    The default text displayed in the TextBox.

.PARAMETER x
    The X-coordinate of the TextBox location.

.PARAMETER y
    The Y-coordinate of the TextBox location.

.PARAMETER size
    The size of the TextBox control.

.PARAMETER maxLength
    The maximum length of text that can be entered in the TextBox.

.NOTES
    - The TextBox is initialized in read-only mode with a gray background and text color.
    - When the TextBox is clicked, it becomes editable and changes its background and text color.
    - The TextBox supports Ctrl+A for selecting all text.
    - The default text is restored if the TextBox loses focus and no text is entered.

.EXAMPLE
    $textBox = New-CustomTextBox -name "exampleTextBox" -defaultText "Enter text here" -x 10 -y 10 -size (New-Object System.Drawing.Size(200, 20)) -maxLength 15
    This command creates a custom TextBox with the specified properties.
#>
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
    $textBox.Text = $defaultText
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
                if ($s.Text -eq $defaultText) {
                    $s.Text = ''
                }
            }
        })

    # Enter handler
    $textBox.Add_Enter({
            param($s, $e)
            $defaultText = $s.Tag
            if ($s.Text -eq $defaultText) {
                $s.ReadOnly = $false
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
                $s.Text = $defaultText
                $s.ForeColor = [System.Drawing.Color]::Gray
                $s.BackColor = [System.Drawing.Color]::LightGray
            }
        })

    return $textBox
}

# Handle the MouseDown event to set ReadOnly to true if clicked outside the TextBoxes
$form.add_MouseDown({
        param($s, $e)

        # Helper function to set TextBox to ReadOnly and reset text if empty
        function SetReadOnlyIfNotFocused($textBox) {
            $textBox.ReadOnly = $true
            if ($textBox.Text -eq '') {
                $textBox.Text = "$($textBox.Tag)"
                $textBox.ForeColor = [System.Drawing.Color]::Gray
                $textBox.BackColor = [System.Drawing.Color]::LightGray
                $textbox.Enabled = $false
                $textbox.Enabled = $true
            }
        }

        $searchBox.Enabled = $false
        $searchBox.Enabled = $true

        # Check each part#Input TextBox
        SetReadOnlyIfNotFocused $part0Input
        SetReadOnlyIfNotFocused $part1Input
        SetReadOnlyIfNotFocused $part2Input
        SetReadOnlyIfNotFocused $part3Input
    })

$textBoxSize = New-Object System.Drawing.Size(150, 20)
$gap = 35 # set space between bottom buttons

# Calculate the total width occupied by the text boxes and their distances
$totalWidth = (4 * $textBoxSize.Width) + (3 * $gap)

# Determine the starting X-coordinate to center the group of text boxes
$startX = [Math]::Max((($form.ClientSize.Width - $totalWidth) / 2) - 12, 0)

$script:part0DefaultText = "X-O-O-O"
$script:part1DefaultText = "O-X-O-O"
$script:part2DefaultText = "O-O-X-O"
$script:part3DefaultText = "O-O-O-X"

# Create and add the text boxes, setting their X-coordinates based on the starting point
$part0Input = New-CustomTextBox -name "part0Input" -defaultText $script:part0DefaultText -x $startX -y 400 -size $textBoxSize -maxLength 15
$form.Controls.Add($part0Input)

$part1Input = New-CustomTextBox -name "part1Input" -defaultText $script:part1DefaultText -x ($startX + $textBoxSize.Width + $gap) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part1Input)

$part2Input = New-CustomTextBox -name "part2Input" -defaultText $script:part2DefaultText -x ($startX + 2 * ($textBoxSize.Width + $gap)) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part2Input)

$part3Input = New-CustomTextBox -name "part3Input" -defaultText $script:part3DefaultText -x ($startX + 3 * ($textBoxSize.Width + $gap)) -y 400 -size $textBoxSize -maxLength 20
$form.Controls.Add($part3Input)

$part0Input.Add_TextChanged({
        if (($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) -and $selectedCheckedListBox.CheckedItems.Count -gt 0) {
            $commitChangesButton.Enabled = $true
        }
        else {
            $commitChangesButton.Enabled = $false
        } 
    })

$part1Input.Add_TextChanged({
        if (($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) -and $selectedCheckedListBox.CheckedItems.Count -gt 0) {
            $commitChangesButton.Enabled = $true
        }
        else {
            $commitChangesButton.Enabled = $false
        }
    })

$part2Input.Add_TextChanged({
        if (($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) -and $selectedCheckedListBox.CheckedItems.Count -gt 0) {
            $commitChangesButton.Enabled = $true
        }
        else {
            $commitChangesButton.Enabled = $false
        }
    })

$part3Input.Add_TextChanged({
        if ($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) {
            $commitChangesButton.Enabled = $true
        }
        else {
            $commitChangesButton.Enabled = $false
        }
    })

# Function to check conditions and enable/disable the commit button
function UpdateCommitChangesButton {
    if (($part0Input.Text -ne "" -and $part1Input.Text -ne "" -and $part2Input.Text -ne "" -and $part3Input.Text -ne "") -and 
        ($part0Input.ReadOnly -ne $true -or $part1Input.ReadOnly -ne $true -or $part2Input.ReadOnly -ne $true -or $part3Input.ReadOnly -ne $true) -and 
        $selectedCheckedListBox.CheckedItems.Count -gt 0) {
        $commitChangesButton.Enabled = $true
    }
    else {
        $commitChangesButton.Enabled = $false
    }
}

# Attach the UpdateCommitChangesButton function to the TextChanged events of the text boxes
$part0Input.Add_TextChanged({ UpdateCommitChangesButton })
$part1Input.Add_TextChanged({ UpdateCommitChangesButton })
$part2Input.Add_TextChanged({ UpdateCommitChangesButton })
$part3Input.Add_TextChanged({ UpdateCommitChangesButton })

<#
.SYNOPSIS
    Creates a styled Button control with specified properties.

.DESCRIPTION
    This function creates a styled Button control with specified properties such as text, location, size, and enabled state.
    The Button is initialized with the provided text and dimensions, and can be optionally enabled or disabled.

.PARAMETER text
    The text displayed on the Button control.

.PARAMETER x
    The X-coordinate of the Button location.

.PARAMETER y
    The Y-coordinate of the Button location.

.PARAMETER width
    The width of the Button control. Default is 100.

.PARAMETER height
    The height of the Button control. Default is 35.

.PARAMETER enabled
    The enabled state of the Button control. Default is $true.

.NOTES
    - The Button control is created with specified dimensions and text.
    - The Button's enabled state can be set using the enabled parameter.
    - The Button is styled with default appearance settings.

.EXAMPLE
    $button = New-StyledButton -text "Submit" -x 50 -y 100 -width 120 -height 40 -enabled $false
    This command creates a styled Button with the specified properties, and the Button is initially disabled.
#>
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
#$commitChangesButton = New-StyledButton -text "Commit Changes" -x 360 -y 10 -width 150 -height 25 -enabled $false
$commitChangesButton = New-StyledButton -text "Commit Changes" -x 260 -y 365 -width ($listBoxWidth + 2) -height 25 -enabled $false
$commitChangesButton.BackColor = $catPurple

# Event handler for clicking the Commit Changes button
$commitChangesButton.Add_Click({
        $form.Enabled = $false
        #$script:selectedItems.Clear()
        $script:selectedCtrlA = 1
        ProcessCommittedChanges
        UpdateAndSyncListBoxes

        $script:globalTopIndex = 0

        # Reset part#Input TextBoxes to default text and ReadOnly status
        function ResetTextBox($textBox, $defaultText) {
            $textBox.Text = $defaultText
            $textBox.ForeColor = [System.Drawing.Color]::Gray
            $textBox.BackColor = [System.Drawing.Color]::LightGray
            $textBox.ReadOnly = $true
        }

        ResetTextBox $part0Input $script:part0DefaultText
        ResetTextBox $part1Input $script:part1DefaultText
        ResetTextBox $part2Input $script:part2DefaultText
        ResetTextBox $part3Input $script:part3DefaultText


        # Set checked status for all selectedCheckedListBox items
        foreach ($index in 0..($selectedCheckedListBox.Items.Count - 1)) {
            $selectedCheckedListBox.SetItemChecked($index, $false)
        }

        $commitChangesButton.Enabled = $false
        $ApplyRenameButton.Enabled = $true
        $form.Enabled = $true
    })

$form.Controls.Add($commitChangesButton)

# Add button to refresh or select a new OU to manage
$loadButton = New-StyledButton -text "Load OU" -x 148 -y 365 -width 94 -height 25 -enabled $true
$loadButton.BackColor = $catBlue
$loadButton.Enabled = $online

<#
.SYNOPSIS
    Refreshes the computer lists and clears the current selections.

.DESCRIPTION
    This event handler function is triggered when the refresh button is clicked. It performs the following steps:
    - Clears the items in the computerCheckedListBox, selectedCheckedListBox, and newNamesListBox.
    - Clears the script-wide checkedItems hashtable.
    - Calls the LoadAndFilterComputers function to reload and filter the computers, updating the computerCheckedListBox with the refreshed data.

.PARAMETER None
    This event handler does not take any parameters.

.NOTES
    - The function ensures that all current selections and items are cleared before reloading the computer list.
    - The LoadAndFilterComputers function is responsible for repopulating the computerCheckedListBox with the updated data.
#>
$loadButton.Add_Click({
        $computerCheckedListBox.Items.Clear()
        $selectedCheckedListBox.Items.Clear()
        $newNamesListBox.Items.Clear()
        $script:checkedItems.Clear()

        LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox
    })
$form.Controls.Add($loadButton)

$applyRenameButton = New-StyledButton -text "Apply Rename" -x 530 -y 365 -width ($listBoxWidth + 2) -height 25 -enabled $false
$applyRenameButton.BackColor = $catRed
$applyRenameButton.ForeColor = $white

<#
.SYNOPSIS
    Initiates the computer renaming process when the "Apply Rename" button is clicked.

.DESCRIPTION
    This event handler function is triggered when the "Apply Rename" button is clicked. It performs the following steps:
    - Validates the list of invalid renames and prompts the user for action.
    - Prompts the user to confirm the renaming operation.
    - Iterates through the selected items and performs the renaming process.
    - Checks the online status of each computer and attempts to rename it.
    - Logs the results of the renaming operation and outputs the total time taken.
    - Generates CSV and log files for the renaming results and triggers a Power Automate flow to upload the files to SharePoint.
    - Updates the UI by removing successfully renamed computers from the list and refreshing the checked list box.

.PARAMETER None
    This event handler does not take any parameters.

.NOTES
    - The function handles both online and offline modes for the renaming process.
    - It logs the results of the renaming, restart, and user login checks.
    - It generates CSV and log files for the renaming results and uploads them to SharePoint.
    - The UI is updated to reflect the changes after the renaming process is completed.

#>
# ApplyRenameButton click event to start renaming process if user chooses
$applyRenameButton.Add_Click({
        $applyRenameButton.Enabled = $false

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

            # Iterate through the selectedCheckedListBox items and perform renaming operations
            foreach ($item in $selectedCheckedListBox.Items) {
                foreach ($change in $script:changesList) {
                    $index = [array]::IndexOf($change.ComputerNames, $item)
                    if ($index -ne -1) {
                        $oldName = $change.ComputerNames[$index]
                        $newName = $change.AttemptedNames[$index]
                        $isValid = $change.Valid[$index]
                        $isDuplicate = $change.Duplicate[$index]

                        if ($isDuplicate) {
                            Write-Host "$oldName has been ignored due to duplicate name" -ForegroundColor Magenta
                            Write-Host " "
                            continue
                        }

                        if (-not $isValid) {
                            Write-Host "$oldName has been ignored due to invalid naming scheme" -ForegroundColor Red
                            Write-Host " "
                            continue
                        }

                        # Check if new name is the same as old name
                        if ($oldName -eq $newName) {
                            Write-Host "New name for $oldName is the same as the old one. Ignoring this device." -ForegroundColor Yellow
                            Write-Host " "
                            continue
                        }

                        $individualTime = [System.TimeSpan]::Zero

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
                }
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

            # Iterate through the script:changesList items and output each change object with its connected devices
            foreach ($change in $script:changesList) {
                Write-Host "Change Object:" -ForegroundColor Green
                Write-Host ("Part0: {0}, Part1: {1}, Part2: {2}, Part3: {3}" -f $change.Part0, $change.Part1, $change.Part2, $change.Part3) -ForegroundColor Green
                foreach ($i in 0..($change.ComputerNames.Length - 1)) {
                    $computerName = $change.ComputerNames[$i]
                    $attemptedName = $change.AttemptedNames[$i]
                    $isValid = $change.Valid[$i]
                    $isDuplicate = $change.Duplicate[$i]
                    Write-Host ("{0}, attemptedname: {1}, valid: {2}, duplicate: {3}" -f $computerName, $attemptedName, $isValid, $isDuplicate) -ForegroundColor Blue
                }
                Write-Host " "
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
                try {
                    # Send the HTTP POST request to trigger the flow
                    Invoke-RestMethod -Uri $flowUrl -Method Post -Headers $headers -Body $jsonBody
                    Write-Host "Triggered Power Automate flow to upload the log files to SharePoint" -ForegroundColor Yellow
                    Write-Host " "
                }
                catch {
                    # Catch block to handle any exceptions from Invoke-RestMethod
                    Write-Host "Failed to trigger Power Automate flow: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host " "
                }
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

            # Clear all relevant lists
            $selectedCheckedListBox.Items.Clear()
            $newNamesListBox.Items.Clear()
            $script:changesList.Clear()
            $script:checkedItems.Clear()
            $script:validNamesList = @()
            $script:invalidNamesList = @()
            $script:logContent = ""

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
            $form.Enabled = $true
            Write-Host "Form enabled"
            Write-Host " "
        }
    })
$form.Controls.Add($applyRenameButton)

LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox | Out-Null

$form.ShowDialog()


