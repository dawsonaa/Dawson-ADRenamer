# Define the shortcut name and location
$shortcutName = "Dawson's ADRenamer.lnk"
$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$shortcutLocation = "$scriptDirectory\$shortcutName"

# Define the target script name
$scriptName = "Dawson's ADRenamer.ps1"
$scriptPath = "$scriptDirectory\$scriptName"

# Check if the shortcut already exists
if (-Not (Test-Path -Path $shortcutLocation)) {
    # Create a WScript.Shell COM object
    $WScriptShell = New-Object -ComObject WScript.Shell

    # Create the shortcut
    $shortcut = $WScriptShell.CreateShortcut($shortcutLocation)
    $shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
    $shortcut.WorkingDirectory = $scriptDirectory
    $shortcut.Save()

    Write-Output "Shortcut created successfully at $shortcutLocation."
}
else {
    Write-Output "Shortcut already exists at $shortcutLocation."
}

# Prompt to press any key to close
Read-Host "Press any key to close"
