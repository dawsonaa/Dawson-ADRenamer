; Define basic information about the installer
Name "Dawson's ADRenamer"
OutFile "DawsonADRenamerInstaller.exe"
InstallDir "$PROFILE\Dawson's ADRenamer"  ; Set installation directory to current user's home
RequestExecutionLevel admin  ; Request admin rights

; Declare variables
Var PROFILE_PATH
Var LENGTH

; Pages to display
Page directory  ; Let the user choose the install directory
Page instfiles  ; Show the installation progress

; Default section for installation
Section "Install Dawson's ADRenamer"

    ; Set the output path to the installation directory
    SetOutPath $INSTDIR
    File "icon2.ico"  ; Ensure the icon is included in the files copied during installation

    ; Put your PowerShell script in the install directory
    File "Dawson's ADRenamer.ps1"

    File "..\LICENSE"
    File "..\README.md"

    File "settings.txt"

    ; Extract the profile-specific base path from $INSTDIR
    StrCpy $PROFILE_PATH $INSTDIR
    ; Find the length of "Dawson's ADRenamer" (adjust for your actual folder name length)
    StrLen $LENGTH "Dawson's ADRenamer"
    IntOp $LENGTH $LENGTH + 1 ; Account for the trailing slash/backslash
    ; Trim the application folder from $INSTDIR
    StrCpy $PROFILE_PATH $PROFILE_PATH -$LENGTH

    ; Create directories in the user's profile-specific Start Menu programs folder
    CreateDirectory "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer"

    ; Create application-specific subdirectories
    CreateDirectory "$INSTDIR\LOGS"
    CreateDirectory "$INSTDIR\RESULTS"

    ; Create Start Menu and Desktop shortcuts
    CreateShortCut "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer\Dawson's ADRenamer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` \
        "$INSTDIR\icon2.ico" 0
    CreateShortCut "$PROFILE_PATH\Desktop\Dawson's ADRenamer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` \
        "$INSTDIR\icon2.ico" 0

    ; Optional: Run a PowerShell command to unblock all files in the installation directory
    nsExec::Exec 'powershell -Command "Get-ChildItem -Path ''$INSTDIR\*'' -Recurse | Unblock-File"'

SectionEnd
