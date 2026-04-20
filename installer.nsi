Unicode True
!include "MUI2.nsh"
!include "LogicLib.nsh"

!define APP_NAME "Offline CRM"
!define APP_VERSION "2"
!define APP_PUBLISHER "Siahtiri Lab"
!define LEGACY_APP_PUBLISHER "POWEREN"
!define APP_WEBSITE "https://github.com/siahtirilab/Offline-CRM"
!define APP_GITHUB "https://github.com/siahtirilab/Offline-CRM"
!define APP_EXE "OfflineCRM.exe"
!define OUTPUT_NAME "OfflineCRM_Setup_v2.exe"
!define APP_DATA_ROOT "$LOCALAPPDATA\OfflineCRM"
!define APP_DATA_DIR "$LOCALAPPDATA\OfflineCRM\data"

Name "${APP_NAME}"
OutFile "output\${OUTPUT_NAME}"
InstallDir "$PROGRAMFILES64\${APP_NAME}"
InstallDirRegKey HKLM "Software\${APP_PUBLISHER}\${APP_NAME}" "InstallDir"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
Icon "Icon.ico"
UninstallIcon "Icon.ico"

!define MUI_ABORTWARNING
!define MUI_ICON "Icon.ico"
!define MUI_UNICON "Icon.ico"
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_TEXT "Open the Offline CRM GitHub page"
!define MUI_FINISHPAGE_RUN_FUNCTION FinishOpenGitHub
!define MUI_FINISHPAGE_SHOWREADME
!define MUI_FINISHPAGE_SHOWREADME_TEXT "Create a desktop shortcut"
!define MUI_FINISHPAGE_SHOWREADME_FUNCTION FinishCreateDesktopShortcut

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Var PreviousInstallDir
Var PreviousUninstallCmd

Function .onInit
  ReadRegStr $PreviousUninstallCmd HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString"
  ReadRegStr $PreviousInstallDir HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "InstallLocation"
  ${If} $PreviousInstallDir == ""
    ReadRegStr $PreviousUninstallCmd HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString"
    ReadRegStr $PreviousInstallDir HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "InstallLocation"
  ${EndIf}
  ${If} $PreviousInstallDir == ""
    ReadRegStr $PreviousInstallDir HKLM "Software\${APP_PUBLISHER}\${APP_NAME}" "InstallDir"
  ${EndIf}
  ${If} $PreviousInstallDir == ""
    ReadRegStr $PreviousInstallDir HKCU "Software\${APP_PUBLISHER}\${APP_NAME}" "InstallDir"
  ${EndIf}
  ${If} $PreviousInstallDir == ""
    ReadRegStr $PreviousInstallDir HKLM "Software\${LEGACY_APP_PUBLISHER}\${APP_NAME}" "InstallDir"
  ${EndIf}
  ${If} $PreviousInstallDir == ""
    ReadRegStr $PreviousInstallDir HKCU "Software\${LEGACY_APP_PUBLISHER}\${APP_NAME}" "InstallDir"
  ${EndIf}

  ${If} $PreviousInstallDir != ""
    Call BackupLegacyData
    IfFileExists "$PreviousInstallDir\Uninstall.exe" 0 +2
      ExecWait '"$PreviousInstallDir\Uninstall.exe" /S _?=$PreviousInstallDir'
  ${EndIf}
FunctionEnd

Function BackupLegacyData
  IfFileExists "$PreviousInstallDir\data\*.*" 0 done
  CreateDirectory "${APP_DATA_ROOT}"
  nsExec::ExecToLog '"$SYSDIR\xcopy.exe" "$PreviousInstallDir\data" "${APP_DATA_DIR}\" /E /I /Y /Q'
done:
FunctionEnd

Section "Install"
  SetOutPath "$INSTDIR"
  RMDir /r "$INSTDIR"
  File /r "dist\OfflineCRM\*"

  WriteRegStr HKLM "Software\${APP_PUBLISHER}\${APP_NAME}" "InstallDir" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${APP_PUBLISHER}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "URLInfoAbout" "${APP_WEBSITE}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\GitHub - New Updates.lnk" "${APP_GITHUB}"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk"
  Delete "$SMPROGRAMS\${APP_NAME}\GitHub - New Updates.lnk"
  RMDir "$SMPROGRAMS\${APP_NAME}"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKLM "Software\${APP_PUBLISHER}\${APP_NAME}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
  DeleteRegKey HKCU "Software\${APP_PUBLISHER}\${APP_NAME}"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
SectionEnd

Function FinishCreateDesktopShortcut
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}"
FunctionEnd

Function FinishOpenGitHub
  ExecShell "open" "${APP_GITHUB}"
FunctionEnd
