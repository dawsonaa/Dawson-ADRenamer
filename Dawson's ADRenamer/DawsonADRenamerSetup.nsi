; Define basic information about the installer
Name "Dawson's ADRenamer"
OutFile "DawsonADRenamerInstaller.exe"
InstallDir "$PROFILE\Dawson's ADRenamer"  ; Set installation directory to current user's home
RequestExecutionLevel admin  ; Request admin rights

; Pages to display
Page directory  ; Let the user choose the install directory
Page instfiles  ; Show the installation progress

; Default section for installation
Section "Install Dawson's ADRenamer"

    ; Set the output path to the installation directory
    SetOutPath $INSTDIR
	File "icon.ico"  ; Ensure the icon is included in the files copied during installation

    ; Put your PowerShell script in the install directory
    File "Dawson's ADRenamer.ps1"

    File "LICENSE"
    File "README.md"

CreateDirectory "$SMPROGRAMS\Dawson's ADRenamer"

CreateDirectory "$INSTDIR\LOGS"
CreateDirectory "$INSTDIR\RESULTS"

 CreateShortCut "$SMPROGRAMS\Dawson's ADRenamer\Dawson's ADRenamer.lnk" "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` "$INSTDIR\icon.ico" 0
    CreateShortCut "$DESKTOP\Dawson's ADRenamer.lnk" "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` "$INSTDIR\icon.ico" 0


    ; Optional: Run a PowerShell command to unblock all files in the installation directory
    nsExec::Exec 'powershell -Command "Get-ChildItem -Path ''$INSTDIR\*'' -Recurse | Unblock-File"'

SectionEnd
