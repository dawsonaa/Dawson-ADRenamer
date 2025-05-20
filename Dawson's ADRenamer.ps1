#### Dawson's ADRenamer
<##> $Version = "24.12.04"
#### Author: Dawson Adams (dawsonaa@ksu.edu, https://github.com/dawsonaa)
#### Kansas State University
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$darkGray = [System.Drawing.Color]::FromArgb(45, 45, 45)
$lightGray = [System.Drawing.Color]::LightGray
$gray = [System.Drawing.Color]::Gray
$visualStudioBlue = [System.Drawing.Color]::FromArgb(78, 102, 221) # RGB values for Visual Studio blue
$kstateDarkPurple = [System.Drawing.Color]::FromArgb(64, 1, 111)
$kstateLightPurple = [System.Drawing.Color]::FromArgb(115, 25, 184)
$white = [System.Drawing.Color]::FromArgb(255, 255, 255)
$black = [System.Drawing.Color]::FromArgb(0, 0, 0)
$red = [System.Drawing.Color]::FromArgb(218, 25, 25)
$grayBlue = [System.Drawing.Color]::FromArgb(75, 75, 140)
$catBlue = [System.Drawing.Color]::FromArgb(3, 106, 199)
$catPurple = [System.Drawing.Color]::FromArgb(170, 13, 206)
$catRed = [System.Drawing.Color]::FromArgb(255, 27, 82)
$catYellow = [System.Drawing.Color]::FromArgb(254, 162, 2)
$catLightYellow = [System.Drawing.Color]::FromArgb(255, 249, 227)
$catDark = [System.Drawing.Color]::FromArgb(1, 3, 32)

$scriptDirectory = Split-Path -Parent $PSCommandPath
$settingsFilePath = Join-Path $scriptDirectory "settings.txt"
$logsFilePath = Join-Path $scriptDirectory "LOGS"

$iconPath = Join-Path $PSScriptRoot "icon.ico"
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
$renameGuideURL = "https://support.ksu.edu/TDClient/30/Portal/KB/ArticleDet?ID=1163"
$companyName = "KSU"

function LoadSettings {
    $settings = @{}
    if (Test-Path $settingsFilePath) {
        $lines = Get-Content $settingsFilePath
        foreach ($line in $lines) {
            if (-not [string]::IsNullOrWhiteSpace($line) -and -not $line.Trim().StartsWith("#")) {
                $parts = $line -split '=', 2
                if ($parts.Length -eq 2) {
                    $key = $parts[0].Trim()
                    $value = $parts[1].Trim()

                    if ($value -match '^(?i)(true|false)$') {
                        $value = [bool]::Parse($value)
                    } elseif ($value -match '^\d+(\.\d+)?$') {
                        $value = [double]$value
                    }

                    $settings[$key] = $value
                }
            }
        }
    } else {
        Write-Host "Settings file not found. Using default values." -ForegroundColor Yellow
    }
    return $settings
}

function Save-Settings {
    if ($settings -and $settings.Count -gt 0) {
        $lines = @()
        foreach ($key in $settings.Keys) {
            $lines += "$key=$($settings[$key])"
        }
        $lines | Set-Content $settingsFilePath
        Write-Host "Settings saved successfully." -ForegroundColor Cyan
    } else {
        Write-Host "No settings to save." -ForegroundColor Yellow
    }
}

function Apply-Style {
    param (
        [hashtable]$settings
    )

    if ($settings["style"] -eq 1) { # light
        $global:defaultFont = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $global:defaultBackColor = $global:catLightYellow
        $global:defaultForeColor = $global:black
        $global:defaultBoxBackColor = $global:lightGray
        $global:defaultBoxForeColor = $global:gray
        $global:defaultListForeColor = $global:black
    } elseif ($settings["style"] -eq 2) { # dark
        $global:defaultFont = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $global:defaultBackColor = $global:catDark
        $global:defaultForeColor = $global:white
        $global:defaultBoxBackColor = $global:catLightYellow
        $global:defaultBoxForeColor = $global:gray
        $global:defaultListForeColor = $global:black
    } else { # default
        $global:defaultFont = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $global:defaultBackColor = $global:darkGray
        $global:defaultForeColor = $global:white
        $global:defaultBoxBackColor = $global:lightGray
        $global:defaultBoxForeColor = $global:gray
        $global:defaultListForeColor = $global:black
    }
}

$settings = LoadSettings
Apply-Style -settings $settings | Out-Null

function Set-FormState {
    param (
        [Parameter(Mandatory = $true)]
        [bool]$IsEnabled,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form,
        [bool]$Loading = $true
    )

    $global:OverlayForm = $global:OverlayForm -as [System.Windows.Forms.Form]

    if ($IsEnabled) {
        $Form.Enabled = $true
        $Form.BringToFront()
        if ($global:OverlayForm) {
            $global:OverlayForm.Close()
            $global:OverlayForm.Dispose()
            $global:OverlayForm = $null
        }
    } else {
        if (-not $global:OverlayForm) {
            $global:OverlayForm = New-Object System.Windows.Forms.Form
            $global:OverlayForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
            $global:OverlayForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
            $global:OverlayForm.BackColor = $defaultListForeColor
            $global:OverlayForm.Opacity = 0.5
            $global:OverlayForm.ShowInTaskbar = $false
            $global:OverlayForm.Size = $Form.Size
            $global:OverlayForm.Location = $Form.Location

            if ($Loading) {
                $loadingLabel = New-Object System.Windows.Forms.Label
                $loadingLabel.Text = "Loading..."
                $loadingLabel.Font = New-Object System.Drawing.Font("Arial", 30, [System.Drawing.FontStyle]::Bold)
                $loadingLabel.ForeColor = $defaultForeColor
                $loadingLabel.BackColor = [System.Drawing.Color]::Transparent
                $loadingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
                $loadingLabel.AutoSize = $false
                $loadingLabel.Width = $global:OverlayForm.Width
                $loadingLabel.Height = 30
                $loadingLabel.Left = 0
                $loadingLabel.Top = [int]($global:OverlayForm.Height / 2 - $loadingLabel.Height / 2)
                $global:OverlayForm.Controls.Add($loadingLabel)
            }

            $global:OverlayForm.Show()
            $global:OverlayForm.Enabled = $false
        }
        $Form.Enabled = $false
    }
}

$formStartY = 30

if ($settings["ask"] -eq $true) {
    $modeSelectionForm = New-Object System.Windows.Forms.Form
    $modeSelectionForm.Text = "Dawson's ADRenamer"
    $modeSelectionForm.Size = New-Object System.Drawing.Size(290, 130)
    $modeSelectionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $modeSelectionForm.MaximizeBox = $false
    $modeSelectionForm.StartPosition = "CenterScreen"
    $modeSelectionForm.Icon = $icon

    $modeSelectionForm.BackColor = $defaultBackColor
    $modeSelectionForm.ForeColor = $defaultForeColor
    $modeSelectionForm.Font = $defaultFont

    $labelMode = New-Object System.Windows.Forms.Label
    $labelMode.Text = "Select ADRenamer Mode"
    $labelMode.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelMode.Size = New-Object System.Drawing.Size(280, 30)
    $labelMode.Location = New-Object System.Drawing.Point(0, 10)
    $modeSelectionForm.Controls.Add($labelMode)

    $global:formClosedByButton = $false

    $buttonOnline = New-Object System.Windows.Forms.Button
    $buttonOnline.Text = "Online"
    $buttonOnline.Location = New-Object System.Drawing.Point(10, 50)
    $buttonOnline.Size = New-Object System.Drawing.Size(75, 30)
    $buttonOnline.BackColor = $catBlue
    $buttonOnline.Add_Click({
        $global:choice = "Online"
        $global:formClosedByButton = $true
        $modeSelectionForm.Close()
    })
    $modeSelectionForm.Controls.Add($buttonOnline)

    $buttonOffline = New-Object System.Windows.Forms.Button
    $buttonOffline.Text = "Offline"
    $buttonOffline.Location = New-Object System.Drawing.Point(100, 50)
    $buttonOffline.Size = New-Object System.Drawing.Size(75, 30)
    $buttonOffline.BackColor = $defaultBoxBackColor
    $buttonOffline.ForeColor = $defaultListForeColor
    $buttonOffline.Add_Click({
        $global:choice = "Offline"
        $global:formClosedByButton = $true
        $modeSelectionForm.Close()
    })
    $modeSelectionForm.Controls.Add($buttonOffline)

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Location = New-Object System.Drawing.Point(190, 50)
    $buttonCancel.Size = New-Object System.Drawing.Size(75, 30)
    $buttonCancel.BackColor = $catRed
    $buttonCancel.Add_Click({
        $global:choice = "Cancel"
        $modeSelectionForm.Close()
    })
    $modeSelectionForm.Controls.Add($buttonCancel)

    $modeSelectionForm.Add_FormClosing({
        param($sender, $e)
        if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing -and -not $global:formClosedByButton) {
            [Environment]::Exit(0)
        }
    })

    $modeSelectionForm.ShowDialog() | Out-Null

    $invalidRenameForm = New-Object System.Windows.Forms.Form
    $invalidRenameForm.Text = ""
    $invalidRenameForm.Size = New-Object System.Drawing.Size(385, 295)
    $invalidRenameForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $invalidRenameForm.ControlBox = $false
    $invalidRenameForm.MaximizeBox = $false
    $invalidRenameForm.StartPosition = "CenterScreen"

    $invalidRenameForm.BackColor = $defaultBackColor
    $invalidRenameForm.ForeColor = $defaultForeColor
    $invalidRenameForm.Font = $defaultFont
    $invalidRenameForm.Icon = $icon

    $invalidRenameLabel = New-Object System.Windows.Forms.Label
    $invalidRenameLabel.Text = "The below invalid renames will not be completed"
    $invalidRenameLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $invalidRenameLabel.Location = New-Object System.Drawing.Point(10, 5)
    $invalidRenameLabel.Size = New-Object System.Drawing.Size(359, 20)
    $invalidRenameForm.Controls.Add($invalidRenameLabel)

    $listBoxInvalidNames = New-Object System.Windows.Forms.ListBox
    $listBoxInvalidNames.Size = New-Object System.Drawing.Size(359, 180)
    $listBoxInvalidNames.Location = New-Object System.Drawing.Point(10, 30)
    $listBoxInvalidNames.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
    $listBoxInvalidNames.BackColor = $defaultBoxBackColor
    $listBoxInvalidNames.ForeColor = $defaultListForeColor
    $invalidRenameForm.Controls.Add($listBoxInvalidNames)

    function RefreshInvalidNamesListBox {
        $listBoxInvalidNames.Items.Clear()

        foreach ($invalidName in $script:invalidNamesList) {
            $listBoxInvalidNames.Items.Add($invalidName)
        }
    }

    $global:formResult = $null

    $buttonOpenGuide = New-Object System.Windows.Forms.Button
    $buttonOpenGuide.Text = "Open $companyName Renaming Guidelines"
    $buttonOpenGuide.Location = New-Object System.Drawing.Point(10, 220)
    $buttonOpenGuide.Size = New-Object System.Drawing.Size(113, 60)
    $buttonOpenGuide.BackColor = $catPurple
    $buttonOpenGuide.Add_Click({
        $global:formResult = "Yes"
        $invalidRenameForm.Close()
    })
    $invalidRenameForm.Controls.Add($buttonOpenGuide)

    $buttonContinue = New-Object System.Windows.Forms.Button
    $buttonContinue.Text = "Continue"
    $buttonContinue.Location = New-Object System.Drawing.Point(133, 220)
    $buttonContinue.Size = New-Object System.Drawing.Size(113, 60)
    $buttonContinue.BackColor = $catBlue
    $buttonContinue.Add_Click({
        $global:formResult = "No"
        $invalidRenameForm.Close()
    })
    $invalidRenameForm.Controls.Add($buttonContinue)

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel Rename Operation"
    $buttonCancel.Location = New-Object System.Drawing.Point(256, 220)
    $buttonCancel.Size = New-Object System.Drawing.Size(113, 60)
    $buttonCancel.BackColor = $catRed
    $buttonCancel.Add_Click({
        $global:formResult = "Cancel"
        $invalidRenameForm.Close()
    })
    $invalidRenameForm.Controls.Add($buttonCancel)

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
}
else {
    $online = $settings["online"]
}

if ($online) {
    Import-Module ActiveDirectory
}
else {
    function Add-DummyComputers {
        param (
            [int]$numberOfDevices = 10
        )

        function Get-RandomDate {
            $randomDays = Get-Random -Minimum 1 -Maximum 200
            return (Get-Date).AddDays(-$randomDays)
        }

        $dummyComputers = @()
        for ($i = 1; $i -le $numberOfDevices; $i++) {
            $dummyComputers += @{
                Name          = "HL-CS-$i"
                LastLogonDate = Get-RandomDate
            }
        }

        return $dummyComputers
    }

    $numberOfDevices = 200
    $dummyComputers = Add-DummyComputers -numberOfDevices $numberOfDevices

    $dummyOUs = @(
        "OU=Sales,OU=DEPT,DC=users,DC=campus",
        "OU=IT,OU=DEPT,DC=users,DC=campus",
        "OU=HR,OU=DEPT,DC=users,DC=campus"
    )

    $onlineStatuses = @("Online", "Online", "Online", "Offline")
    $restartOutcomes = @("Success", "Success", "Success", "Success", "Fail")
    $loggedOnUserss = @("User1", "User2", "User3", "User4", "User5", "User6", "User7", "none", "none", "none")
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

    function Get-RandomOutcome {
        param (
            [Parameter(Mandatory = $true)]
            [array]$outcomes
        )
        return $outcomes | Get-Random
    }

    $username = "dawsonaa" # OFFLINE
}

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

        $username = $cred.UserName

        try {
            $null = Get-ADDomain -Credential $cred
            $connectionSuccess = $true
        }
        catch {
            $errorMessage = "Invalid credentials or insufficient permissions. Please try again."
        }
    }
    Write-Host "ONLINE MODE - Version $Version - Sufficient Credentials Provided. Logged on as $username." -ForegroundColor Green
}else {
    Write-Host "OFFLINE MODE - Version $Version - No credentials are needed." -ForegroundColor Green
}

$hashSet = [System.Collections.Generic.HashSet[string]]::new()

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

$script:changesList = New-Object System.Collections.ArrayList
$script:newNamesList = @()

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

$colors = @(
    [CustomColor]::new(243, 12, 122), # Vibrant Pink
    [CustomColor]::new(243, 120, 22), # Vibrant Orange
    [CustomColor]::new(76, 175, 80), # Light Green
    [CustomColor]::new(0, 188, 212), # Cyan
    [CustomColor]::new(103, 58, 183)   # Deep Purple
)

$global:nextColorIndex = 0

function Get-DepartmentString($deviceName, $part) {
    if ($online) {
        $device = Get-ADComputer -Identity $deviceName -Properties CanonicalName
        $ouLocation = $device.CanonicalName -replace "^CN=[^,]+,", ""

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

    if ($part -match "(?i)(\d*)dept(\d*)") {
        $prefixLength = if ($matches[1] -and $matches[1] -ge 2 -and $matches[1] -le 5) { [int]::Parse($matches[1]) } else { $null }
        $suffixLength = if ($matches[2] -and $matches[2] -ge 2 -and $matches[2] -le 5) { [int]::Parse($matches[2]) } else { $null }

        if ($suffixLength) {
            return $deptString.Substring(0, [Math]::Min($deptString.Length, $suffixLength))
        }
        elseif ($prefixLength) {
            $startIndex = [Math]::Max(0, $deptString.Length - $prefixLength)
            return $deptString.Substring($startIndex, $prefixLength)
        }
    }

    return $deptString
}

function ProcessCommittedChanges {
    $hashSet.Clear()
    $script:newNamesListBox.Items.Clear()
    $script:validNamesList = @()
    $script:invalidNamesList = @()

    $attemptedNamesTracker = @{}

    foreach ($computerName in $script:selectedCheckedItems.Keys) {
        $parts = $computerName -split '-'
        $part0 = $parts[0]
        $part1 = $parts[1]
        $part2 = if ($parts.Count -ge 3) { $parts[2] } else { $null }
        $part3 = if ($parts.Count -ge 4) { $parts[3..($parts.Count - 1)] -join '-' } else { $null }

        $part0InputValue = if (-not $part0Input.ReadOnly) { $part0Input.Text } else { $null }
        $part1InputValue = if (-not $part1Input.ReadOnly) { $part1Input.Text } else { $null }
        $part2InputValue = if (-not $part2Input.ReadOnly) { $part2Input.Text } else { $null }
        $part3InputValue = if (-not $part3Input.ReadOnly) { $part3Input.Text } else { $null }

        if ($part0InputValue) { $part0 = $part0InputValue }
        if ($part1InputValue) { $part1 = $part1InputValue }
        if ($part2InputValue) {
            $totalLengthForpart2 = 15 - ($part0.Length + $part1.Length + ($parts.Count - 1))
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

        if ($part0 -match "(?i)dept") { $part0 = Get-DepartmentString $computerName $part0 }
        if ($part1 -match "(?i)dept") { $part1 = Get-DepartmentString $computerName $part1 }
        if ($part2 -match "(?i)dept") { $part2 = Get-DepartmentString $computerName $part2 }
        if ($part3 -match "(?i)dept") { $part3 = Get-DepartmentString $computerName $part3 }

        if ($part3) {
            $newName = "$part0-$part1-$part2-$part3"
        }
        elseif ($part2) {
            $newName = "$part0-$part1-$part2"
        }
        else {
            $newName = "$part0-$part1"
        }

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

        if (-not $attemptedNamesTracker.ContainsKey($newName)) {
            $attemptedNamesTracker[$newName] = 1
        }
        else {
            $attemptedNamesTracker[$newName]++
        }

        $attemptedNames = @($newName)
        $duplicate = @($isDuplicate)

        UpdateNameMap -originalName $computerName -newName $newName

        $existingChange = $null
        foreach ($change in $script:changesList) {
            $part0Comparison = ($change.Part0 -eq $part0InputValue -or ([string]::IsNullOrEmpty($change.Part0) -and [string]::IsNullOrEmpty($part0InputValue)))
            $part1Comparison = ($change.Part1 -eq $part1InputValue -or ([string]::IsNullOrEmpty($change.Part1) -and [string]::IsNullOrEmpty($part1InputValue)))
            $part2Comparison = ($change.Part2 -eq $part2InputValue -or ([string]::IsNullOrEmpty($change.Part2) -and [string]::IsNullOrEmpty($part2InputValue)))
            $part3Comparison = ($change.Part3 -eq $part3InputValue -or ([string]::IsNullOrEmpty($change.Part3) -and [string]::IsNullOrEmpty($part3InputValue)))
            $validComparison = ($change.Valid -eq $isValid)

            if ($part0Comparison -and $part1Comparison -and $part2Comparison -and $part3Comparison -and $validComparison) {
                $existingChange = $change
                break
            }
        }

        $tempChangesToRemove = @()

        foreach ($change in $script:changesList) {
            if ($change -ne $existingChange -and $change.ComputerNames -contains $computerName) {
                $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $computerName }

                if ($change.ComputerNames.Count -eq 0) {
                    $tempChangesToRemove += $change
                }
            }
        }

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
            $groupColor = if (-not $isValid) { [CustomColor]::new(255, 0, 0) } else { $colors[$global:nextColorIndex % $colors.Count] }
            $global:nextColorIndex++
            $newChange = [Change]::new(@($computerName), $part0InputValue, $part1InputValue, $part2InputValue, $part3InputValue, $groupColor, @($isValid), $attemptedNames, $duplicate)
            $script:changesList.Add($newChange) | Out-Null
        }
    }
}

function ConvertTo-EmailAddress {
    param (
        [string]$username
    )

    $emailLocalPart = $username -replace "Users\\", ""
    $email = $emailLocalPart + "@ksu.edu"
    return $email
}

function Update-OutlookWebDraft {
    param (
        [string]$oldName,
        [string]$newName,
        [string]$emailAddress,
        [string]$emailSubject,
        [string]$emailBody
    )

    $username = ($emailAddress -split '@')[0]
    $subject = "Action Required: Restart Your Device"
    $body = @"
Dear $username,

Your computer has been renamed from $oldName to $newName as part of a maintenance operation. To avoid device name syncing issues, please restart your device as soon as possible. If you face any issues, please contact IT support.

You can submit a ticket using the following link: $supportTicketLink

Best regards,
IT Support Team
"@

    function EncodeURL($string) {
        $encodedString = [System.Uri]::EscapeDataString($string)
        return $encodedString -replace '\+', '%20'
    }

    $url = "https://outlook.office.com/mail/deeplink/compose?to=" + (EncodeURL($emailAddress)) +
            "&subject=" + (EncodeURL($emailSubject)) +
            "&body=" + (EncodeURL($emailBody))

    Start-Process $url
    Write-Host "Draft email created for $emailAddress" -ForegroundColor Green
}

function Show-EmailDrafts {
    param (
        [array]$loggedOnDevices
    )

    $emailForm = New-Object System.Windows.Forms.Form
    $emailForm.Text = "Generate Email Drafts"
    $emailForm.Size = New-Object System.Drawing.Size(600, 620)
    $emailForm.MaximizeBox = $false
    $emailForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $emailForm.StartPosition = "CenterScreen"
    $emailForm.Icon = $icon

    $emailForm.BackColor = $defaultBackColor
    $emailForm.ForeColor = $defaultForeColor
    $emailForm.Font = $defaultFont

    $emailSubjectTextBox = New-Object System.Windows.Forms.TextBox
    $emailSubjectTextBox.Text = "IT Support - Computer [oldName] renamed to [newName]"  # Default email subject
    $emailSubjectTextBox.Location = New-Object System.Drawing.Point(10, 340)
    $emailSubjectTextBox.Size = New-Object System.Drawing.Size(560, 20)
    $emailSubjectTextBox.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular)

    $emailSubjectTextBox.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $emailSubjectTextBox.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })
    $emailForm.Controls.Add($emailSubjectTextBox)

    $emailBodyTextBox = New-Object System.Windows.Forms.TextBox
    $emailBodyTextBox.Text = @"
Dear [Username],

Your computer has been renamed from [oldName] to [newName] as part of a maintenance operation. To avoid device name syncing issues, please restart your device as soon as possible. If you face any issues, please contact IT support.

Best regards,
IT Support Team
"@
    $emailBodyTextBox.Location = New-Object System.Drawing.Point(10, 370)
    $emailBodyTextBox.Size = New-Object System.Drawing.Size(560, 140)
    $emailBodyTextBox.Multiline = $true
    $emailBodyTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $emailBodyTextBox.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular)

    $emailBodyTextBox.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $emailBodyTextBox.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })
    $emailForm.Controls.Add($emailBodyTextBox)

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

    $listBoxOldName = [CustomListBox]::new()
    $listBoxOldName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxOldName.Location = New-Object System.Drawing.Point(10, 40)
    $listBoxOldName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    $listBoxNewName = [CustomListBox]::new()
    $listBoxNewName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxNewName.Location = New-Object System.Drawing.Point(200, 40)
    $listBoxNewName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    $listBoxUserName = [CustomListBox]::new()
    $listBoxUserName.Size = New-Object System.Drawing.Size(180, 300)
    $listBoxUserName.Location = New-Object System.Drawing.Point(390, 40)
    $listBoxUserName.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

    foreach ($device in $loggedOnDevices) {
        $listBoxOldName.Items.Add($device.OldName)
        $listBoxNewName.Items.Add($device.NewName)
        $listBoxUserName.Items.Add($device.UserName)
    }

    $emailForm.Controls.Add($listBoxOldName)
    $emailForm.Controls.Add($listBoxNewName)
    $emailForm.Controls.Add($listBoxUserName)

    $global:syncingSelection = $false

    $syncSelection = {
        param ($s, $e)
        if (-not $global:syncingSelection) {
            $global:syncingSelection = $true
            $selectedIndices = $s.SelectedIndices

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

    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Open Email Drafts"
    $createButton.Size = New-Object System.Drawing.Size(90, 45)
    $createButton.Location = New-Object System.Drawing.Point(480, 520)
    $createButton.Add_Click({
        $emailSubject = $emailSubjectTextBox.Text
        $emailBody = $emailBodyTextBox.Text

        for ($i = 0; $i -lt $listBoxOldName.Items.Count; $i++) {
            $oldName = $listBoxOldName.Items[$i]
            $newName = $listBoxNewName.Items[$i]
            $userName = $listBoxUserName.Items[$i]
            $deviceInfo = $loggedOnDevices | Where-Object { $_.OldName -eq $oldName -and $_.NewName -eq $newName -and $_.UserName -eq $userName }
            if ($deviceInfo) {
                $emailAddress = ConvertTo-EmailAddress $deviceInfo.UserName

                $customSubject = $emailSubject -replace '\[oldName\]', $oldName `
                                            -replace '\[newName\]', $newName `
                                            -replace '\[Username\]', $userName

                $customBody = $emailBody -replace '\[oldName\]', $oldName `
                                     -replace '\[newName\]', $newName `
                                     -replace '\[Username\]', $userName

                Update-OutlookWebDraft -oldName $deviceInfo.OldName -newName $deviceInfo.NewName -emailAddress $emailAddress -emailSubject $customSubject -emailBody $customBody
            }
        }
        $emailForm.Close()
    })
    $emailForm.Controls.Add($createButton)
    $emailForm.ShowDialog()
}

function Select-OU {
    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "Select Organizational Unit"
    $ouForm.Size = New-Object System.Drawing.Size(400, 590)
    $ouForm.MaximizeBox = $false
    $ouForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $ouForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ouForm.Icon = $icon

    $ouForm.BackColor = $defaultBackColor
    $ouForm.ForeColor = $defaultForeColor
    $ouForm.Font = $defaultFont
    $ouForm.Add_FormClosing({
            param($s, $e)
                [Environment]::Exit(0)
        })

    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Size = New-Object System.Drawing.Size(365, 500)
    $treeView.Location = New-Object System.Drawing.Point(10, 10)
    $treeView.BackColor = $defaultBoxBackColor
    $treeView.ForeColor = $defaultListForeColor
    $treeView.Visible = $true

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

    $defaultOUButton = New-Object System.Windows.Forms.Button
    $defaultOUButton.Text = "DC=users,DC=campus"
    $defaultOUButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $defaultOUButton.Size = New-Object System.Drawing.Size(165, 23)
    $defaultOUButton.Location = New-Object System.Drawing.Point(210, 520)
    $defaultOUButton.Add_Click({
            $ouForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $ouForm.Close()
        })

    $ouForm.Controls.Add($treeView)
    $ouForm.Controls.Add($selectedOUButton)
    $ouForm.Controls.Add($defaultOUButton)

    $treeView.Add_NodeMouseClick({
            param ($s, $e)
            $selectedNode = $e.Node
            if ($null -ne $selectedNode) {
                if ($selectedNode.Tag -match "DC=") {
                    $ouForm.Tag = [string]$selectedNode.Tag
                    $selectedOUButton.Text = $selectedNode.Tag
                    $selectedOUButton.Enabled = $true

                    $selectedNode.Nodes.Clear()
                    $selectedNode.Expand()

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

    function Update-TreeView {
        param ($treeView)

        Write-Host "Fetching OUs from AD..."
        $ous = Get-ADOrganizationalUnit -Filter * | Where-Object {
            $_.DistinguishedName -match '^OU=[^,]+,DC=users,DC=campus$'
        } | Sort-Object DistinguishedName

        $nodeHashTable = @{}
        foreach ($ou in $ous) {
            $node = New-Object System.Windows.Forms.TreeNode
            $node.Text = $ou.Name
            $node.Tag = $ou.DistinguishedName
            $parentDN = $ou.DistinguishedName -replace "^OU=[^,]+,", ""

            if ($parentDN -eq 'DC=users,DC=campus') {
                $treeView.Nodes.Add($node)
            }

            $nodeHashTable[$ou.DistinguishedName] = $node
        }
    }

    Update-TreeView -treeView $treeView | Out-Null

    if ($ouForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:ouPath = $ouForm.Tag
    }
    else {
        $script:ouPath = 'DC=users,DC=campus'
    }
}

$script:filteredComputers = @()

function LoadAndFilterComputers {
    param (
        [System.Windows.Forms.CheckedListBox]$computerCheckedListBox
    )

    try {

        if ($online) {
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

        Write-Host "Selected OU Path: $script:ouPath"

        Set-FormState -IsEnabled $false -Form $form

        if ($online) {
            Write-Host "Loading AD endpoints..."
        }
        else {
            Write-Host "Loaded 'endpoints'... OFFLINE"
        }

        $loadedCount = 0
        $deviceRefresh = 150
        $deviceTimer = 0

        $cutoffDate = (Get-Date).AddDays(-180)

        $script:filteredComputers = @()

        if ($online) {
            $computers = Get-ADComputer -Filter * -Properties LastLogonDate -SearchBase $script:ouPath

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
                if ($deviceTimer -ge $deviceRefresh) {
                    $deviceTimer = 0
                    $progress = [math]::Round(($loadedCount / $computerCount) * 100)
                    Write-Progress -Activity "Loading endpoints from AD..." -Status "$progress% Complete:" -PercentComplete $progress
                }
            }
        }
        else {
            $script:filteredComputers = $dummyComputers | Where-Object {
                $_.LastLogonDate -and
                [DateTime]::Parse($_.LastLogonDate) -ge $cutoffDate
            } | Sort-Object -Property Name

            $computerCount = $script:filteredComputers.Count
            $filteredOutCount = $dummyComputers.Count - $script:filteredComputers.Count

            $computerCheckedListBox.Items.Clear()

            $script:filteredComputers | ForEach-Object {
                $computerCheckedListBox.Items.Add($_.Name, $false) | Out-Null
                $loadedCount++
                $deviceTimer++

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

        Set-FormState -IsEnabled $true -Form $form
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

$form = New-Object System.Windows.Forms.Form
$form.Opacity = 1
$form.Size = New-Object System.Drawing.Size(805, 490) # 785, 520
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Icon = $icon

$form.BackColor = $defaultBackColor
$form.ForeColor = $defaultForeColor
$form.Font = $defaultFont

if ($online) {
    $form.Text = "ADRenamer $Version - Online"
}
else {
    $form.Text = "ADRenamer $Version - Offline"
}

function Add-MenuItemSeparator {
    param (
        [System.Windows.Forms.MenuStrip]$menuStrip,
        [string]$character = "I"
    )

    $separator = New-Object System.Windows.Forms.ToolStripMenuItem
    $separator.Text = $character
    $separator.Enabled = $false

    $separator.Add_MouseHover({
        # Do nothing on hover
    })

    $menuStrip.Items.Add($separator) | Out-Null
}

$rendererCode = @"
using System.Drawing;
using System.Windows.Forms;

public class CustomMenuStripRenderer : ToolStripProfessionalRenderer
{
    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    {
        // Skip rendering the background entirely for the separator (text = "I")
        if (e.Item.Text == "I" || e.Item.Text == " ")
        {
            return; // Do nothing
        }

        // Render other items as usual
        base.OnRenderMenuItemBackground(e);
    }

    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    {
        // Custom text rendering for separator
        if (e.Item.Text == "I")
        {
            e.TextColor = Color.Gray; // Set a custom color for the separator text
            e.Graphics.DrawString(e.Text, e.TextFont, Brushes.Gray, e.TextRectangle);
        }
        else
        {
            base.OnRenderItemText(e);
        }
    }
    public void doNothing() {} // hush the compiler
}
"@

Add-Type -TypeDefinition $rendererCode -Language CSharp -ReferencedAssemblies @(
    "System.Windows.Forms",
    "System.Drawing"
)

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Renderer = New-Object CustomMenuStripRenderer
$menuStrip.BackColor = $defaultBackColor
$menuStrip.ForeColor = $defaultForeColor
$fontStyle = [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic

$menuStrip.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

$settingsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$settingsMenu.Text = "Settings"
$settingsMenu.Add_Click({
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Edit Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(400, 270)
    $settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $settingsForm.MaximizeBox = $false
    $settingsForm.StartPosition = "CenterScreen"
    $settingsForm.BackColor = $defaultBackColor
    $settingsForm.ForeColor = $defaultForeColor
    $settingsForm.Font = $defaultFont

    $styleOptions = @("default", "dark", "light")
    $onlineOptions = @("true", "false")
    $askOptions = @("true", "false")

    $styleValueMap = @{
        "default" = 3
        "dark" = 2
        "light" = 1
    }
    $styleReverseMap = @{
        3 = "default"
        2 = "dark"
        1 = "light"
    }

    $settings = @{}

    $settingsFilePath = Join-Path $scriptDirectory "settings.txt"
    if (Test-Path $settingsFilePath) {
        $settingsLines = Get-Content $settingsFilePath
        foreach ($line in $settingsLines) {
            if (-not [string]::IsNullOrWhiteSpace($line) -and -not $line.Trim().StartsWith("#")) {
                $parts = $line -split '=', 2
                if ($parts.Length -eq 2) {
                    $key = $parts[0].Trim()
                    $value = $parts[1].Trim()
                    $settings[$key] = $value
                }
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Settings file not found. Creating a default settings file.", "Settings")
        $settings["online"] = "false"
        $settings["ask"] = "false"
        $settings["style"] = 3 # default style
        $settings | ForEach-Object { "$($_.Key)=$($_.Value)" } | Set-Content $settingsFilePath
    }

    $yPosition = 10
    $dropdowns = @{}

    foreach ($key in $settings.Keys) {
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, $yPosition)
        $label.Size = New-Object System.Drawing.Size(400, 20)
        $settingsForm.Controls.Add($label)

        $yPosition += 20

        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Location = New-Object System.Drawing.Point(10, $yPosition)
        $dropdown.Size = New-Object System.Drawing.Size(80, 20)
        $dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        switch ($key) {
            "style" {
                $label.Text = "Style"
                $dropdown.Items.AddRange($styleOptions)
                $currentStyle = $styleReverseMap[[int]$settings[$key]]
                $dropdown.SelectedItem = $currentStyle
            }
            "online" {
                $label.Text = "Online on Startup"
                $dropdown.Items.AddRange($onlineOptions)
                $dropdown.SelectedItem = $settings[$key]
                if ($dropdowns["ask"].SelectedItem -eq "true") {
                    $dropdown.Enabled = $false
                }
                else {
                    $dropdown.Enabled = $true
                }
            }
            "Ask" {
                $label.Text = "Ask Mode on Startup"
                $dropdown.Items.AddRange($askOptions)
                $dropdown.SelectedItem = $settings[$key]
                $dropdown.Add_SelectedIndexChanged({
                    if ($dropdowns["ask"].SelectedItem -eq "true") {
                        $dropdowns["online"].Enabled = $false
                    } else {
                        $dropdowns["online"].Enabled = $true
                    }
                })
            }
        }

        $dropdowns[$key] = $dropdown
        $settingsForm.Controls.Add($dropdown)

        $yPosition += 40
    }

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(150, $yPosition)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.BackColor = $catBlue
    $saveButton.ForeColor = $white
    $saveButton.Font = $defaultFont
    $saveButton.Add_Click({
        foreach ($key in $dropdowns.Keys) {
            switch ($key) {
                "style" {
                    $selectedStyle = $dropdowns[$key].SelectedItem
                    $settings[$key] = $styleValueMap[$selectedStyle]
                }
                default {
                    $settings[$key] = $dropdowns[$key].SelectedItem
                }
            }
        }

        Save-Settings
        Apply-Style -settings $settings

        [System.Windows.Forms.MessageBox]::Show("Settings saved successfully.", "Settings")
        $settingsForm.Close()

        $form.BackColor = $defaultBackColor
        $form.ForeColor = $defaultForeColor
        $form.Font = $defaultFont
        $menuStrip.BackColor = $defaultBackColor
        $menuStrip.ForeColor = $defaultForeColor
        $computerCheckedListBox.BackColor = $defaultBoxBackColor
        $computerCheckedListBox.ForeColor = $defaultListForeColor
        $selectedCheckedListBox.BackColor = $defaultBoxBackColor
        $selectedCheckedListBox.ForeColor = $defaultListForeColor
        $newNamesListBox.BackColor = $defaultBoxBackColor
        $newNamesListBox.ForeColor = $defaultListForeColor
        $searchBox.ForeColor = $defaultBoxForeColor
        $searchBox.BackColor = $defaultBoxBackColor

        $part0Input.BackColor = $defaultBoxBackColor
        $part0Input.ForeColor = $defaultBoxForeColor
        $part1Input.BackColor = $defaultBoxBackColor
        $part1Input.ForeColor = $defaultBoxForeColor
        $part2Input.BackColor = $defaultBoxBackColor
        $part2Input.ForeColor = $defaultBoxForeColor
        $part3Input.BackColor = $defaultBoxBackColor
        $part3Input.ForeColor = $defaultBoxForeColor

        $commitChangesButton.ForeColor = $defaultForeColor
        $applyRenameButton.ForeColor = $defaultForeColor

        $colorPanel3.BackColor = $defaultBackColor
        $colorPanel2.BackColor = $defaultBackColor
        $colorPanel.BackColor = $defaultBackColor
    })

    $settingsForm.Controls.Add($saveButton)

    $settingsForm.ShowDialog()
})

$githubMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$githubMenu.Text = "Github Repository"
$githubMenu.Add_Click({
    Start-Process "https://github.com/dawsonaa/Dawson-ADRenamer"
})

$downloadMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$downloadMenu.Text = "Download newest installer"
$downloadMenu.Add_Click({
    Start-Process "https://github.com/dawsonaa/Dawson-ADRenamer/raw/refs/heads/main/DawsonADRenamerInstaller.exe"
})

$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "View"

$viewLogs = New-Object System.Windows.Forms.ToolStripMenuItem
$viewLogs.Text = "Logs"
$viewLogs.Add_Click({
    $logsForm = New-Object System.Windows.Forms.Form
    $logsForm.Text = "ADRenamer Logs Viewer"
    $logsForm.Size = New-Object System.Drawing.Size(830, 400)
    $logsForm.Icon = $icon
    $logsForm.StartPosition = "CenterScreen"
    $logsForm.BackColor = $defaultBackColor
    $logsForm.ForeColor = $defaultForeColor

    $logsListBox = New-Object System.Windows.Forms.ListBox
    $logsListBox.Dock = [System.Windows.Forms.DockStyle]::Left
    $logsListBox.Width = 200

    $searchPanel = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $searchPanel.Height = 20

    $searchTextBox = New-Object System.Windows.Forms.TextBox
    $searchTextBox.Text = "Search"
    $searchTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $searchTextBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $searchTextBox.Margin = [System.Windows.Forms.Padding]::Empty
    $searchTextBox.ForeColor = $defaultBoxForeColor
    $searchTextBox.BackColor = $defaultBoxBackColor

    $searchTextBox.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $searchTextBox.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

    $searchTextBox.Add_Enter({
        if ($this.Text -eq "Search") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
            $this.BackColor = [System.Drawing.Color]::White
        }
    })

    $searchTextBox.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Search"
            $this.ForeColor = $defaultBoxForeColor
            $this.BackColor = $defaultBoxBackColor
        }
    })

    $searchTextBox.Add_Keydown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter)
        {
            $e.SuppressKeyPress = $true
            $e.Handled = $true

            if ($selectedFile = $logsListBox.SelectedItem)
            {
                $filePath = Join-Path $logsFilePath $selectedFile
                $logsContent = Get-Content -Path $filePath
                $searchTerm = $searchTextBox.Text
                if (-not $searchTerm)
                {
                    $logsTextBox.Lines = $logsContent
                    return
                }

                $matchingLines = $logsContent -match $searchTerm
                if (!$matchingLines)
                {
                    $logsTextBox.Text = "No results for $searchTerm"
                }
                else
                {
                    $logsTextBox.Text = ($matchingLines -join "`r`n")
                }
            }
            else
            {
                Write-Host "No log.txt selected"
            }
        }
    })

    $logsTextBox = New-Object System.Windows.Forms.TextBox
    $logsTextBox.Multiline = $true
    $logsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $logsTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $logsTextBox.ReadOnly = $true

    if (-Not (Test-Path $logsFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("No LOGS folder found in the current directory.")
        return
    }

    $first = 0
    Get-ChildItem -Path $logsFilePath -Filter "*.txt" | ForEach-Object {
        $logsListBox.Items.Add($_.Name)
        if ($first -eq 0){
            $logsListBox.SelectedItem = $_.Name
            $logsTextBox.Lines = Get-Content -Path (Join-Path $logsFilePath $logsListBox.SelectedItem)
            $logsTextBox.SelectionStart = 0
            $logsTextBox.SelectionLength = 0
            $first = 1
        }
    }

    $logsListBox.Add_MouseDoubleClick({
        $selectedFile = $logsListBox.SelectedItem
        if ($selectedFile) {
            $filePath = Join-Path $logsFilePath $selectedFile
            $logsTextBox.Lines = Get-Content -Path $filePath
        }
    })

    $searchPanel.Controls.Add($searchTextBox)

    $logsForm.Controls.Add($logsTextBox)
    $logsForm.Controls.Add($searchPanel)
    $logsForm.Controls.Add($logsListBox)

    $logsForm.ShowDialog()
})

$viewResults = New-Object System.Windows.Forms.ToolStripMenuItem
$viewResults.Text = "Results"
$viewResults.Add_Click({
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Open Results File"
    $resultsForm.Size = New-Object System.Drawing.Size(300, 400)
    $resultsForm.icon = $icon
    $resultsForm.StartPosition = "CenterScreen"
    $resultsForm.BackColor = $defaultBackColor
    $resultsForm.ForeColor = $defaultForeColor

    $resultsListBox = New-Object System.Windows.Forms.ListBox
    $resultsListBox.Dock = [System.Windows.Forms.DockStyle]::Fill

    $resultsFolder = Join-Path $scriptDirectory "RESULTS"
    Write-Host "Results Folder: $resultsFolder"

    if (-Not (Test-Path $resultsFolder)) {
        [System.Windows.Forms.MessageBox]::Show("No RESULTS folder found in the current directory.")
        return
    }

    $csvFiles = Get-ChildItem -Path $resultsFolder -Filter "*.csv"
    Write-Host "Files Found: $($csvFiles.Count)"
    $csvFiles | ForEach-Object {
        $resultsListBox.Items.Add($_.Name)
    }

    $resultsListBox.Add_MouseDoubleClick({
        $selectedFile = $resultsListBox.SelectedItem
        if ($selectedFile) {
            $filePath = Join-Path $resultsFolder $selectedFile
            Write-Host "Selected file: $filePath"

            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace((Get-Item $resultsFolder).FullName)
            $item = $folder.ParseName($selectedFile)
            $item.InvokeVerb("openas")
        } else {
            Write-Host "No file selected."
        }
    })

    $resultsForm.Controls.Add($resultsListBox)
    $resultsForm.ShowDialog()
})

$loadOU = New-Object System.Windows.Forms.ToolStripMenuItem
$loadOU.Text = "Load OU"
$loadOU.Add_Click({
    $computerCheckedListBox.Items.Clear()
    $selectedCheckedListBox.Items.Clear()
    $newNamesListBox.Items.Clear()
    $script:checkedItems.Clear()

    LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox
})

$contactMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$contactMenu.Text = "Contact Author: Dawson Adams (dawsonaa@ksu.edu)"
$contactMenu.Add_Click({
    Start-Process "msteams://teams.microsoft.com/l/chat/0/0?users=dawsonaa@ksu.edu"
})

$fileMenu.DropDownItems.Add($settingsMenu) | Out-Null
$fileMenu.DropDownItems.Add($downloadMenu) | Out-Null
$fileMenu.DropDownItems.Add($githubMenu) | Out-Null

$viewMenu.DropDownItems.Add($viewResults) | Out-Null
$viewMenu.DropDownItems.Add($viewLogs) | Out-Null

$menuStrip.Items.Add($fileMenu) | Out-Null
Add-MenuItemSeparator -menuStrip $menuStrip
$menuStrip.Items.Add($viewMenu) | Out-Null
Add-MenuItemSeparator -menuStrip $menuStrip
$menuStrip.Items.Add($loadOU) | Out-Null
Add-MenuItemSeparator -menuStrip $menuStrip

for ($i = 0; $i -lt 3; $i++){
    Add-MenuItemSeparator -menuStrip $menuStrip -character " "
}

$contactMenu.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]$fontStyle)
$menuStrip.Items.Add($contactMenu) | Out-Null

$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

$script:invalidNamesList = @()
$script:validNamesList = @()
$script:customNamesList = @() 
$script:ouPath = 'DC=users,DC=campus'
$script:checkedItems = @{}
$script:selectedCheckedItems = @{}

$listBoxWidth = 250
$listBoxHeight = 350

function UpdateAndSyncListBoxes {
    $script:newNamesListBox.Items.Clear()
    $selectedCheckedListBox.Items.Clear()
    $itemsToRemove = New-Object System.Collections.ArrayList

    $script:selectedCheckedItems.Clear()

    $groupedInvalidItems = @{}
    $groupedValidItems = @{}
    $groupColors = @{}

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

    $sortedNonChangeItems = $nonChangeItems | Sort-Object

    $script:newNamesListBox.BeginUpdate()
    $selectedCheckedListBox.BeginUpdate()

    foreach ($group in $groupedInvalidItems.Keys) {
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

        foreach ($item in ($groupedValidItems[$group] | Sort-Object)) {
            $change = $script:changesList | Where-Object { $_.ComputerNames -contains $item }
            $index = [array]::IndexOf($change.ComputerNames, $item)
            $newName = $change.AttemptedNames[$index]
            $script:newNamesListBox.Items.Add($newName) | Out-Null
            $selectedCheckedListBox.Items.Add($item) | Out-Null
        }
    }

    foreach ($item in $sortedNonChangeItems) {
        $script:newNamesListBox.Items.Add($item) | Out-Null
        $selectedCheckedListBox.Items.Add($item) | Out-Null
    }

    foreach ($item in $itemsToRemove) {
        $script:newNamesListBox.Items.Remove($item)
    }

    $script:newNamesListBox.EndUpdate()
    $selectedCheckedListBox.EndUpdate()
    $script:newNamesListBox.Refresh()
    $selectedCheckedListBox.Refresh()
}

$computerCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$computerCheckedListBox.Location = New-Object System.Drawing.Point(10, $formStartY)
$computerCheckedListBox.Size = New-Object System.Drawing.Size($listBoxWidth, $listBoxHeight)
$computerCheckedListBox.IntegralHeight = $false
$computerCheckedListBox.BackColor = $defaultBoxBackColor
$computerCheckedListBox.ForeColor = $defaultListForeColor

$script:computerCtrlA = 1

$contextMenu = New-Object System.Windows.Forms.ContextMenu

$menuItemSelectAll = New-Object System.Windows.Forms.MenuItem "Select All"
$menuItemSelectAll.Add_Click({
    for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
        $computerCheckedListBox.SetItemChecked($i, $true)
        $currentItem = $computerCheckedListBox.Items[$i]
        $script:checkedItems[$currentItem] = $true
    }
})

$menuItemUnselectAll = New-Object System.Windows.Forms.MenuItem "Unselect All"
$menuItemUnselectAll.Add_Click({
    for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
        $computerCheckedListBox.SetItemChecked($i, $false)
        $currentItem = $computerCheckedListBox.Items[$i]
        $script:checkedItems.Remove($currentItem)
    }
})

$contextMenu.MenuItems.Add($menuItemSelectAll)
$contextMenu.MenuItems.Add($menuItemUnselectAll)

$computerCheckedListBox.ContextMenu = $contextMenu

$computerCheckedListBox.Add_MouseDown({
    $allSelected = $true
    $allUnselected = $true

    for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
        if ($computerCheckedListBox.GetItemChecked($i)) {
            $allUnselected = $false
        } else {
            $allSelected = $false
        }
    }

    $menuItemSelectAll.Enabled = -not $allSelected
    $menuItemUnselectAll.Enabled = -not $allUnselected
})

$computerCheckedListBox.Add_KeyDown({
        param($s, $e)

        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            Set-FormState -IsEnabled $false -Form $form
            Write-Host "Ctrl+A pressed, toggling selection state, Form disabled for loading..."

            $maxItemsToCheck = 500
            $itemsProcessed = 0

            if ($script:computerCtrlA -eq 1) {
                for ($i = 0; $i -lt $computerCheckedListBox.Items.Count; $i++) {
                    if ($itemsProcessed -ge $maxItemsToCheck) {
                        break
                    }
                    $computerCheckedListBox.SetItemChecked($i, $true)
                    $currentItem = $computerCheckedListBox.Items[$i]
                    $script:checkedItems[$currentItem] = $true
                    $itemsProcessed++
                }
                $script:computerCtrlA = 0
                Write-Host "All items selected"
            }
            else {
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
                $script:computerCtrlA = 1
                Write-Host "All items unselected"
            }

            $selectedCheckedListBox.BeginUpdate()
            $selectedCheckedListBox.Items.Clear()
            $sortedCheckedItems = $script:checkedItems.Keys | Sort-Object
            foreach ($item in $sortedCheckedItems) {
                $selectedCheckedListBox.Items.Add($item, $script:checkedItems[$item]) | Out-Null
            }
            $selectedCheckedListBox.EndUpdate()

            Set-FormState -IsEnabled $true -Form $form
            Write-Host ""

            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

$script:originalToNewNameMap = @{}
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

$computerCheckedListBox.add_ItemCheck({
        param($s, $e)

        $item = $computerCheckedListBox.Items[$e.Index]

        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            if (-not $script:checkedItems.ContainsKey($item)) {
                $script:checkedItems[$item] = $true
                $selectedCheckedListBox.Items.Add($item, $true)
            }
        }
        elseif ($e.NewValue -eq [System.Windows.Forms.CheckState]::Unchecked) {
            if ($script:checkedItems.ContainsKey($item)) {
                $script:checkedItems.Remove($item)

                $newName = $script:originalToNewNameMap[$item]

                $selectedCheckedListBox.BeginUpdate()
                $newNamesListBox.BeginUpdate()

                $selectedCheckedListBox.Items.Remove($item)
                if ($newName) {
                    $newNamesListBox.Items.Remove($newName)
                }

                $selectedCheckedListBox.EndUpdate()
                $newNamesListBox.EndUpdate()

                $script:originalToNewNameMap.Remove($item)
            }

            $tempChangesToRemove = @()
            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }

            foreach ($changeToRemove in $tempChangesToRemove) {
                $script:changesList.Remove($changeToRemove)
            }
        }
    })

$form.Controls.Add($computerCheckedListBox)

$selectedCheckedListBox = New-Object System.Windows.Forms.CheckedListBox
$selectedCheckedListBox.Location = New-Object System.Drawing.Point(260, $formStartY)
$selectedCheckedListBox.Size = New-Object System.Drawing.Size(($listBoxWidth + 20), ($listBoxHeight))
$selectedCheckedListBox.IntegralHeight = $false
$selectedCheckedListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawVariable
$selectedCheckedListBox.BackColor = $defaultBoxBackColor
$selectedCheckedListBox.ForeColor = $defaultListForeColor

$script:selectedCtrlA = 1

$selectedCheckedListBox.add_KeyDown({
        param($s, $e)
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

$selectedCheckedListBox.add_ItemCheck({
        param($s, $e)

        $item = $selectedCheckedListBox.Items[$e.Index]
    
        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            if (-not $script:selectedCheckedItems.ContainsKey($item)) {
                $script:selectedCheckedItems[$item] = $true
            }
            $colorPanel3.Invalidate()
            $colorPanel.Invalidate()
            $colorPanel2.Invalidate()
        }
        elseif ($e.NewValue -eq [System.Windows.Forms.CheckState]::Unchecked) {
            if ($script:selectedCheckedItems.ContainsKey($item)) {
                $script:selectedCheckedItems.Remove($item)
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

$selectedCheckedListBox.add_MeasureItem({
        param ($s, $e)
        $e.ItemHeight = 20
    })

function UpdateSelectedCheckedListBox {
    $sortedItems = New-Object System.Collections.ArrayList
    $nonChangeItems = New-Object System.Collections.ArrayList

    foreach ($change in $script:changesList) {
        $sortedComputerNames = $change.ComputerNames | Sort-Object
        foreach ($computerName in $sortedComputerNames) {
            $sortedItems.Add($computerName) | Out-Null
        }
    }

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

    if ($sortedNonChangeItems.Count -gt 0) {
        $sortedItems.AddRange($sortedNonChangeItems)
    }

    $checkedItems = @{}
    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        if ($selectedCheckedListBox.GetItemChecked($i)) {
            $checkedItems[$selectedCheckedListBox.Items[$i]] = $true
        }
    }

    $selectedCheckedListBox.BeginUpdate()
    $selectedCheckedListBox.Items.Clear()
    foreach ($item in $sortedItems) {
        $selectedCheckedListBox.Items.Add($item) | Out-Null
    }
    $selectedCheckedListBox.EndUpdate()

    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        if ($checkedItems.ContainsKey($selectedCheckedListBox.Items[$i])) {
            $selectedCheckedListBox.SetItemChecked($i, $true)
        }
    }

    UpdateAndSyncListBoxes
}

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

$colorPanelBack = $defaultBackColor

$colorPanel3 = New-Object System.Windows.Forms.Panel
$colorPanel3.Location = New-Object System.Drawing.Point(240, $formStartY) # 530, 40
$colorPanel3.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel3.AutoScroll = $true
$colorPanel3.BackColor = $defaultBackColor

$colorPanel = New-Object System.Windows.Forms.Panel
$colorPanel.Location = New-Object System.Drawing.Point(510, $formStartY) # 260, 40
$colorPanel.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel.BackColor = $defaultBackColor

$colorPanel2 = New-Object System.Windows.Forms.Panel
$colorPanel2.Location = New-Object System.Drawing.Point(780, $formStartY) # 530, 40
$colorPanel2.Size = New-Object System.Drawing.Size(20, 350)
$colorPanel2.BackColor = $defaultBackColor

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

$newNamesListBox = New-Object System.Windows.Forms.ListBox
$newNamesListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawVariable
$newNamesListBox.Location = New-Object System.Drawing.Point(530, $formStartY)
$newNamesListBox.Size = New-Object System.Drawing.Size(($listBoxWidth + 20), ($listBoxHeight))
$newNamesListBox.IntegralHeight = $false
$newNamesListBox.BackColor = $defaultBoxBackColor
$newNamesListBox.ForeColor = $defaultListForeColor

$measureItemHandler = {
    param (
        [object]$s,
        [System.Windows.Forms.MeasureItemEventArgs]$e
    )
    $e.ItemHeight = 18
}

$drawItemHandler = {
    param (
        [object]$s,
        [System.Windows.Forms.DrawItemEventArgs]$e
    )

    if ($e.Index -ge 0) {
        $itemText = $s.Items[$e.Index]
        $e.DrawBackground()
        $textBrush = [System.Drawing.SolidBrush]::new($e.ForeColor)
        $pointF = [System.Drawing.PointF]::new($e.Bounds.X, $e.Bounds.Y)
        $e.Graphics.DrawString($itemText, $e.Font, $textBrush, $pointF)
        $e.DrawFocusRectangle()
    }
}

$newNamesListBox.add_MeasureItem($measureItemHandler)
$newNamesListBox.add_DrawItem($drawItemHandler)

$newNamesListBox.add_SelectedIndexChanged({
        $newNamesListBox.ClearSelected()
    })
$form.Controls.Add($newNamesListBox)

$script:globalTopIndex = 0

function Update-ListBoxes {
    param ($topIndex)
    if ($topIndex -ge 0 -and $topIndex -lt $selectedCheckedListBox.Items.Count) {
        $selectedCheckedListBox.TopIndex = $topIndex
        $newNamesListBox.TopIndex = $topIndex
    }
}

$selectedCheckedListBox.add_MouseWheel({
        param ($s, $e)
        $e.Handled = $true
    })

$newNamesListBox.add_MouseWheel({
        param ($s, $e)
        $e.Handled = $true
    })

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

$selectedCheckedListBox.add_SelectedIndexChanged({
        param ($s, $e)
        $colorPanel.Invalidate()
        $colorPanel2.Invalidate()
        $colorPanel3.Invalidate()
    })

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

$menuRemove = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemove.Text = "Remove selected device(s)"
$menuRemove.Add_Click({
        $selectedItems = @($selectedCheckedListBox.CheckedItems | ForEach-Object { $_ })

        if (!($selectedItems.Count -gt 0)) {
            Write-Host "No devices selected"
            return
        }
        Write-Host ""

        $tempChangesToRemove = @()
        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            $selectedCheckedListBox.Items.Remove($item)
            $newName = $script:originalToNewNameMap[$item]
            if ($newName) {
                $newNamesListBox.Items.Remove($newName)
            }

            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }
        }

        foreach ($changeToRemove in $tempChangesToRemove) {
            $script:changesList.Remove($changeToRemove)
        }

        Write-Host "Selected device(s) removed: $($selectedItems -join ', ')"
        Write-Host ""
    })

$menuRemoveAll = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuRemoveAll.Text = "Remove all device(s)"
$menuRemoveAll.Enabled = $false
$menuRemoveAll.Add_Click({
        $selectedItems = @($selectedCheckedListBox.Items | ForEach-Object { $_ })
        Set-FormState -IsEnabled $false -Form $form
        $script:customNamesList = @()

        $tempChangesToRemove = @()

        foreach ($item in $selectedItems) {
            $script:checkedItems.Remove($item)

            $index = $computerCheckedListBox.Items.IndexOf($item)
            if ($index -ge 0) {
                $computerCheckedListBox.SetItemChecked($index, $false)
            }

            $selectedCheckedListBox.Items.Remove($item)
            $newName = $script:originalToNewNameMap[$item]
            if ($newName) {
                $newNamesListBox.Items.Remove($newName)
            }

            foreach ($change in $script:changesList) {
                if ($change.ComputerNames -contains $item) {
                    $change.ComputerNames = $change.ComputerNames | Where-Object { $_ -ne $item }
                    if ($change.ComputerNames.Count -eq 0) {
                        $tempChangesToRemove += $change
                    }
                }
            }
        }

        foreach ($changeToRemove in $tempChangesToRemove) {
            $script:changesList.Remove($changeToRemove)
        }

        Set-FormState -IsEnabled $true -Form $form
        Write-Host "All devices removed from the list"
    })

$menuSelectAll = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuSelectAll.Text = "Select All"
$menuSelectAll.Add_Click({
    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        $selectedCheckedListBox.SetItemChecked($i, $true)
        $item = $selectedCheckedListBox.Items[$i]
        $script:checkedItems[$item] = $true
    }
    Write-Host "All items selected"
})

$menuUnselectAll = [System.Windows.Forms.ToolStripMenuItem]::new()
$menuUnselectAll.Text = "Unselect All"
$menuUnselectAll.Add_Click({
    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        $selectedCheckedListBox.SetItemChecked($i, $false)
        $item = $selectedCheckedListBox.Items[$i]
        $script:checkedItems.Remove($item)
    }
    Write-Host "All items unselected"
})

$contextMenu.Items.Add($menuSelectAll)
$contextMenu.Items.Add($menuUnselectAll)

$contextMenu.add_Opening({
    $allSelected = $true
    $allUnselected = $true

    for ($i = 0; $i -lt $selectedCheckedListBox.Items.Count; $i++) {
        if ($selectedCheckedListBox.GetItemChecked($i)) {
            $allUnselected = $false
        } else {
            $allSelected = $false
        }
    }

    $menuSelectAll.Enabled = -not $allSelected
    $menuUnselectAll.Enabled = -not $allUnselected

    $selectedItems = @($selectedCheckedListBox.CheckedItems)
    $itemsInBox = @($selectedCheckedListBox.Items)

    if ($itemsInBox.Count -gt 0) {
        $menuRemoveAll.Enabled = $true
    } else {
        $menuRemoveAll.Enabled = $false
    }

    if ($selectedItems.Count -gt 0) {
        $menuRemove.Enabled = $true
    } else {
        $menuRemove.Enabled = $false
    }
})

$selectedCheckedListBox.add_KeyDown({
        param ($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $menuFindAndReplace.PerformClick()
            $e.Handled = $true
        }
    })

$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemove) | Out-Null
$contextMenu.Items.Add([System.Windows.Forms.ToolStripItem]$menuRemoveAll) | Out-Null

$selectedCheckedListBox.ContextMenuStrip = $contextMenu
$form.Controls.Add($selectedCheckedListBox)


$form.Controls.Add($colorPanel3)
$form.Controls.Add($colorPanel)
$form.Controls.Add($colorPanel2)

$colorPanel3.BringToFront()
$colorPanel.BringToFront()
$colorPanel2.BringToFront()

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10,(355 + $formStartY))
$searchBox.Size = New-Object System.Drawing.Size(230, 20)
$searchBox.ForeColor = $defaultBoxForeColor
$searchBox.BackColor = $defaultBoxBackColor
$searchBox.Text = "Search"
$searchBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$searchBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$searchBox.add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $searchBox.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })

$searchBox.Add_Enter({
        if ($this.Text -eq "Search") {
            $this.Text = ''
            $this.ForeColor = [System.Drawing.Color]::Black
            $this.BackColor = [System.Drawing.Color]::White
        }
    })

$searchBox.Add_Leave({
        if ($this.Text -eq '') {
            $this.Text = "Search"
            $this.ForeColor = $defaultBoxForeColor
            $this.BackColor = $defaultBoxBackColor
        }
    })

$searchBox.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Set-FormState -IsEnabled $false -Form $form
            $e.SuppressKeyPress = $true
            $e.Handled = $true

            $computerCheckedListBox.Items.Clear()

            $searchTerm = $searchBox.Text
            $filteredList = $script:filteredComputers | Where-Object { $_.Name -like "*$searchTerm*" }
            foreach ($computer in $filteredList) {
                $isChecked = $false
                if ($script:checkedItems.ContainsKey($computer.Name)) {
                    $isChecked = $script:checkedItems[$computer.Name]
                }
                $computerCheckedListBox.Items.Add($computer.Name, $isChecked)
            }
            Set-FormState -IsEnabled $true -Form $form
            Write-Host ""
        }
    })

$form.Controls.Add($searchBox)

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
    $textBox.ForeColor = $defaultBoxForeColor
    $textBox.BackColor = $defaultBoxBackColor
    $textBox.Text = $defaultText
    $textBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $textBox.ReadOnly = $true
    $textBox.MaxLength = [Math]::Min(15, $maxLength)
    $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $textBox.Tag = $defaultText
    $textBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)

    $textBox.add_KeyDown({
            param($s, $e)
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
                $s.SelectAll()
                $e.SuppressKeyPress = $true
                $e.Handled = $true
            }
        })

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

    $textBox.Add_Leave({
            param($s, $e)
            $defaultText = $s.Tag
            if ($s.Text -eq '') {
                $s.ReadOnly = $true
                $s.Text = $defaultText
                $s.ForeColor = $defaultBoxForeColor
                $s.BackColor = $defaultBoxBackColor
            }
        })

    return $textBox
}

$form.add_MouseDown({
        param($s, $e)

        function SetReadOnlyIfNotFocused($textBox) {
            $textBox.ReadOnly = $true
            if ($textBox.Text -eq '') {
                $textBox.Text = "$($textBox.Tag)"
                $textBox.ForeColor = $defaultBoxForeColor
                $textBox.BackColor = $defaultBoxBackColor
                $textbox.Enabled = $false
                $textbox.Enabled = $true
            }
        }

        $searchBox.Enabled = $false
        $searchBox.Enabled = $true

        SetReadOnlyIfNotFocused $part0Input
        SetReadOnlyIfNotFocused $part1Input
        SetReadOnlyIfNotFocused $part2Input
        SetReadOnlyIfNotFocused $part3Input
    })

$textBoxSize = New-Object System.Drawing.Size(166, 20)
$gap = 35
$startX = 10

$script:part0DefaultText = "X-O-O-O"
$script:part1DefaultText = "O-X-O-O"
$script:part2DefaultText = "O-O-X-O"
$script:part3DefaultText = "O-O-O-X"

$part0Input = New-CustomTextBox -name "part0Input" -defaultText $script:part0DefaultText -x $startX -y (385 + $formStartY) -size $textBoxSize -maxLength 15
$form.Controls.Add($part0Input)

$part1Input = New-CustomTextBox -name "part1Input" -defaultText $script:part1DefaultText -x ($startX + $textBoxSize.Width + $gap) -y (385 + $formStartY) -size $textBoxSize -maxLength 15
$form.Controls.Add($part1Input)

$part2Input = New-CustomTextBox -name "part2Input" -defaultText $script:part2DefaultText -x ($startX + 2 * ($textBoxSize.Width + $gap)) -y (385 + $formStartY) -size $textBoxSize -maxLength 15
$form.Controls.Add($part2Input)

$part3Input = New-CustomTextBox -name "part3Input" -defaultText $script:part3DefaultText -x ($startX + 3 * ($textBoxSize.Width + $gap)) -y (385 + $formStartY) -size $textBoxSize -maxLength 15
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

$part0Input.Add_TextChanged({ UpdateCommitChangesButton })
$part1Input.Add_TextChanged({ UpdateCommitChangesButton })
$part2Input.Add_TextChanged({ UpdateCommitChangesButton })
$part3Input.Add_TextChanged({ UpdateCommitChangesButton })

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
$commitChangesButton = New-StyledButton -text "Commit Changes" -x 260 -y (355 + $formStartY) -width ($listBoxWidth + 2) -height 26 -enabled $false
$commitChangesButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$commitChangesButton.BackColor = $catPurple
$commitChangesButton.ForeColor = $defaultForeColor

$commitChangesButton.Add_Click({
        Set-FormState -IsEnabled $false -Form $form
        $script:selectedCtrlA = 1
        ProcessCommittedChanges
        UpdateAndSyncListBoxes
        $script:globalTopIndex = 0

        function ResetTextBox($textBox, $defaultText) {
            $textBox.Text = $defaultText
            $textBox.ForeColor = $defaultBoxForeColor
            $textBox.BackColor = $defaultBoxBackColor
            $textBox.ReadOnly = $true
        }

        ResetTextBox $part0Input $script:part0DefaultText
        ResetTextBox $part1Input $script:part1DefaultText
        ResetTextBox $part2Input $script:part2DefaultText
        ResetTextBox $part3Input $script:part3DefaultText

        foreach ($index in 0..($selectedCheckedListBox.Items.Count - 1)) {
            $selectedCheckedListBox.SetItemChecked($index, $false)
        }

        $commitChangesButton.Enabled = $false
        $ApplyRenameButton.Enabled = $true
        Set-FormState -IsEnabled $true -Form $form
    })

$form.Controls.Add($commitChangesButton)

$applyRenameButton = New-StyledButton -text "Apply Rename" -x 530 -y (355 + $formStartY) -width ($listBoxWidth + 2) -height 26 -enabled $false
$applyRenameButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$applyRenameButton.BackColor = $catRed
$applyRenameButton.ForeColor = $defaultForeColor
$applyRenameButton.Add_Click({
        $applyRenameButton.Enabled = $false

        if ($script:invalidNamesList.Count -gt 0) {
            Set-FormState -IsEnabled $false -Form $form
            RefreshInvalidNamesListBox
            $invalidRenameForm.ShowDialog() | Out-Null

            if ($global:formResult -eq "Yes") {
                Start-Process $renameGuideURL
            } elseif ($global:formResult -eq "Cancel") {
                Set-FormState -IsEnabled $true -Form $form
                $applyRenameButton.Enabled = $true
                return
            }
        }

        $userResponse = [System.Windows.Forms.MessageBox]::Show(("`nDo you want to proceed with renaming? `n`n"), "Apply Rename", [System.Windows.Forms.MessageBoxButtons]::YesNo)

        $successfulRenames = @()
        $failedRenames = @()
        $successfulRestarts = @()
        $failedRestarts = @()
        $loggedOnUsers = @()
        $loggedOnDevices = @()
        $totalTime = [System.TimeSpan]::Zero

        if ($userResponse -eq "Yes") {
            $logContent = ""

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

            Set-FormState -IsEnabled $false -Form $form
            Write-Host " "
            Write-Host "Starting rename operation..."
            Write-Host " "

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

                        if ($oldName -eq $newName) {
                            Write-Host "New name for $oldName is the same as the old one. Ignoring this device." -ForegroundColor Yellow
                            Write-Host " "
                            continue
                        }

                        $individualTime = [System.TimeSpan]::Zero

                        if ($online) {
                            $checkOfflineTime = Measure-Command {
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
                            $checkOfflineTime = Measure-Command {
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

                        Write-Host "Time taken to check if $oldName was online: $($checkOfflineTime.TotalSeconds) seconds" -ForegroundColor Blue
                        $individualTime = $individualTime.Add($checkOfflineTime)

                        if ($online) {
                            $testComp = Get-WmiObject Win32_ComputerSystem -ComputerName $oldName -Credential $cred
                        }

                        Write-Host "Checking if $oldName was renamed successfully..."

                        if ($online) {
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
                                    continue
                                }
                            }
                        }
                        else {
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
                                    continue
                                }
                            }
                        }

                        if ($renameResult.ReturnValue -eq 0) {
                            Write-Host "Computer $oldName successfully renamed to $newName." -ForegroundColor Green

                            Write-Host "Time taken to rename $oldName`: $($checkRenameTime.TotalSeconds) seconds" -ForegroundColor Blue

                            $individualTime = $individualTime.Add($checkRenameTime)
                            $successfulRenames += [PSCustomObject]@{OldName = $oldName; NewName = $newName }

                            if ($online) {
                                $checkLoginTime = Measure-Command {
                                    Write-Host "Checking if $oldname has a user logged on..."
                                    $loggedOnUser = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $oldName -Credential $cred | Select-Object -ExpandProperty UserName
                                }
                            }
                            else {
                                $checkLoginTime = Measure-Command {
                                    Write-Host "Checking if $oldName has a user logged on..."
                                    $loggedOnUser = Get-RandomOutcome -outcomes $loggedOnUserss
                                }

                                if ($loggedOnUser -eq "none") {
                                    $loggedOnUser = $null
                                }
                            }

                            if ($online) {
                                $checkRestartTime = Measure-Command {
                                    if (-not $loggedOnUser) {
                                        try {
                                            Write-Host "Computer $oldname ($newName) has no users logged on." -ForegroundColor Green

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

                                        Write-Host "Time taken to check $oldname ($newName) has logged on users`: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                        $individualTime = $individualTime.Add($checkLoginTime)

                                        $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                        $loggedOnUsers += "$oldName`: $loggedOnUser"

                                        $loggedOnDevices += [PSCustomObject]@{
                                            OldName  = $oldName
                                            NewName  = $newName
                                            UserName = $loggedOnUser
                                        }
                                    }
                                }
                            }
                            else {
                                $checkRestartTime = Measure-Command {
                                    if (-not $loggedOnUser) {
                                        try {
                                            Write-Host "Computer $oldName ($newName) has no users logged on." -ForegroundColor Green

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

                                        Write-Host "Time taken to check $oldName ($newName) for logged on users: $($checkLoginTime.TotalSeconds) seconds" -ForegroundColor Blue
                                        $individualTime = $individualTime.Add($checkLoginTime)

                                        $failedRestarts += [PSCustomObject]@{OldName = $oldName; NewName = $newName }
                                        $loggedOnUsers += "$oldName`: $loggedOnUser"

                                        $loggedOnDevices += [PSCustomObject]@{
                                            OldName  = $oldName
                                            NewName  = $newName
                                            UserName = $loggedOnUser
                                        }
                                    }
                                }
                            }
                            Write-Host "Time taken to send restart to $oldname`: $($checkRestartTime.TotalSeconds) seconds" -ForegroundColor Blue
                            $individualTime = $individualTime.Add($checkRestartTime)
                        }
                        else {
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

            $resultsFolderPath = Join-Path -Path $scriptDirectory -ChildPath "RESULTS"
            $logsFolderPath = Join-Path -Path $scriptDirectory -ChildPath "LOGS"

            if (-not (Test-Path -Path $resultsFolderPath)) {
                New-Item -Path $resultsFolderPath -ItemType Directory | Out-Null
            }

            if (-not (Test-Path -Path $logsFolderPath)) {
                New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
            }

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

            $csvData = @()

            $maxCount = $successfulRenames.Count
            if ($successfulRestarts.Count -gt $maxCount) { $maxCount = $successfulRestarts.Count }
            if ($failedRenames.Count -gt $maxCount) { $maxCount = $failedRenames.Count }
            if ($failedRestarts.Count -gt $maxCount) { $maxCount = $failedRestarts.Count }
            if ($loggedOnUsers.Count -gt $maxCount) { $maxCount = $loggedOnUsers.Count }

            $csvData = ""
            $headers = @(
                "Successful Renames,,Successful Restarts,,Failed Renames,,Failed Restarts,,Logged On Users,,,"
                "New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,New Names,Old Names,Logged On User"
            ) -join "`r`n"

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

            $csvOutput = "$headers`r`n$csvData"
            $dateTimeString = (Get-Date).ToString("yy-MM-dd_HH-mmtt")

            $csvFileName = "ADRenamer_Results_$dateTimeString"
            $csvFilePath = Join-Path -Path $resultsFolderPath -ChildPath "$csvFileName.csv"

            $csvOutput | Out-File -FilePath $csvFilePath -Encoding utf8

            Write-Host "RESULTS CSV file created at $csvFilePath" -ForegroundColor Yellow
            Write-Host " "

            $logFileName = "ADRenamer_Log_$dateTimeString"
            $logFilePath = Join-Path -Path $logsFolderPath -ChildPath "$logFileName.txt"
            $script:logContent | Out-File -FilePath $logFilePath -Encoding utf8

            Write-Host "LOGS TXT file created at $logFilePath" -ForegroundColor Yellow
            Write-Host " "

            $csvFileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
            $csvBase64Content = [Convert]::ToBase64String($csvFileContent)

            $logFileContent = [System.IO.File]::ReadAllBytes($logFilePath)
            $logBase64Content = [Convert]::ToBase64String($logFileContent)

            $flowUrl = "https://prod-166.westus.logic.azure.com:443/workflows/5e172f6d92d24c6a995023362c53472f/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7-7I6wW8ga9i3hSfzjP7-O_AFLNFmE-_cxCGt6g3f9A"

            $body = @{
                csvFileName    = "$csvFileName`-$username.csv"
                csvFileContent = $csvBase64Content
                logFileName    = "$logFileName`-$username.txt"
                logFileContent = $logBase64Content
            }

            $jsonBody = $body | ConvertTo-Json -Depth 3

            $headers = @{
                "Content-Type" = "application/json"
            }

            if ($online) {
                try {
                    Invoke-RestMethod -Uri $flowUrl -Method Post -Headers $headers -Body $jsonBody
                    Write-Host "Triggered Power Automate flow to upload the log files to SharePoint" -ForegroundColor Yellow
                    Write-Host " "
                }
                catch {
                    Write-Host "Failed to trigger Power Automate flow: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host " "
                }
            }
            else {
                Write-Host "Triggered Power Automate flow to upload the log files to SharePoint (OFFLINE - IGNORED)" -ForegroundColor Yellow
                Write-Host " "
            }

            if ($loggedOnUsers.Count -gt 0) {
                Write-Host "Logged on users:" -ForegroundColor Yellow
                $loggedOnUsers | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
                Show-EmailDrafts -loggedOnDevices $loggedOnDevices | Out-Null
            }
            else {
                Write-Host "No users were logged on to the renamed computers." -ForegroundColor Green
            }
            Write-Host " "

            foreach ($renameEntry in $script:validNamesList) {
                $oldName, $newName = $renameEntry -split ' -> '
                $script:checkedItems.Remove($oldName)
                $script:filteredComputers = $script:filteredComputers | Where-Object { $_.Name -ne $oldName }

                $index = $computerCheckedListBox.Items.IndexOf($oldName)
                if ($index -ge 0) {
                    $computerCheckedListBox.Items.RemoveAt($index)
                }
            }

            $selectedCheckedListBox.Items.Clear()
            $newNamesListBox.Items.Clear()
            $script:changesList.Clear()
            $script:checkedItems.Clear()
            $script:validNamesList = @()
            $script:invalidNamesList = @()
            $script:logContent = ""

            $searchTerm = $searchBox.Text
            $filteredList = $script:filteredComputers | Where-Object { $_.Name -like "*$searchTerm*" }

            $computerCheckedListBox.Items.Clear()
            foreach ($computer in $filteredList) {
                $isChecked = $false
                if ($script:checkedItems.ContainsKey($computer.Name)) {
                    $isChecked = $script:checkedItems[$computer.Name]
                }
                $computerCheckedListBox.Items.Add($computer.Name, $isChecked)
            }
            Set-FormState -IsEnabled $true -Form $form
            Write-Host " "
        } else{
            Set-FormState -IsEnabled $true -Form $form
            $applyRenameButton.Enabled = $true
        }
    })
$form.Controls.Add($applyRenameButton)

LoadAndFilterComputers -computerCheckedListBox $computerCheckedListBox | Out-Null

$form.ShowDialog()