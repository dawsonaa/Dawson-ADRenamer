Name "Dawson's ADRenamer"
OutFile "DawsonADRenamerInstaller.exe"
InstallDir "$PROFILE\Dawson's ADRenamer"
RequestExecutionLevel admin

Var PROFILE_PATH
Var LENGTH

Page directory
Page instfiles

Section "Install Dawson's ADRenamer"

    SetOutPath $INSTDIR

    File "..\Dawson's ADRenamer.ps1"
    File "..\icon.ico"
    File "..\README.md"
    File "..\settings.txt"

    StrCpy $PROFILE_PATH $INSTDIR
    StrLen $LENGTH "Dawson's ADRenamer"
    IntOp $LENGTH $LENGTH + 1 ; Account for the trailing slash/backslash
    StrCpy $PROFILE_PATH $PROFILE_PATH -$LENGTH

    CreateDirectory "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer"

    CreateDirectory "$INSTDIR\images"
    setOutPath "$INSTDIR\images"
    File /r "..\images\*.*"

    CreateDirectory "$INSTDIR\LOGS"
    CreateDirectory "$INSTDIR\RESULTS"

    CreateShortCut "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Dawson's ADRenamer\Dawson's ADRenamer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` \
        "$INSTDIR\icon.ico" 0
    CreateShortCut "$PROFILE_PATH\Desktop\Dawson's ADRenamer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\Dawson's ADRenamer.ps1"` \
        "$INSTDIR\icon.ico" 0

    nsExec::Exec 'powershell -Command "Get-ChildItem -Path ''$INSTDIR\*'' -Recurse | Unblock-File"'

SectionEnd
