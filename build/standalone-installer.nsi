Unicode true

!ifndef APP_VERSION
  !error "APP_VERSION is required"
!endif
!ifndef SOURCE_DIR
  !error "SOURCE_DIR is required"
!endif
!ifndef OUT_FILE
  !error "OUT_FILE is required"
!endif

!include "MUI2.nsh"
!include "FileFunc.nsh"

Var VerificationMode

Name "Exhibit Builder"
OutFile "${OUT_FILE}"
InstallDir "$LOCALAPPDATA\Programs\Exhibit Builder"
RequestExecutionLevel user
SetCompress off

!define MUI_ABORTWARNING
!define MUI_LICENSEPAGE_CHECKBOX
!define MUI_LICENSEPAGE_CHECKBOX_TEXT "I accept the terms of the Licence Agreement"
!define MUI_FINISHPAGE_RUN "$INSTDIR\Exhibit Builder.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Launch Exhibit Builder"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "${SOURCE_DIR}\EULA.txt"
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Function .onInit
  StrCpy $VerificationMode "0"
  ClearErrors
  ${GetOptions} "$CMDLINE" "/RELEASEVERIFY" $0
  IfErrors +2 0
  StrCpy $VerificationMode "1"
FunctionEnd

Section "Exhibit Builder" MainSection
  SetShellVarContext current
  SetOutPath "$INSTDIR"
  File /r "${SOURCE_DIR}\*.*"
  FileOpen $0 "$INSTDIR\.exhibit-builder-install-marker" w
  FileWrite $0 "Exhibit Builder ${APP_VERSION}"
  FileClose $0
  WriteUninstaller "$INSTDIR\Uninstall Exhibit Builder.exe"

  StrCmp $VerificationMode "1" verification_install normal_install

  verification_install:
    FileOpen $0 "$INSTDIR\.exhibit-builder-verification-install" w
    FileWrite $0 "Temporary release verification install"
    FileClose $0
    Goto install_complete

  normal_install:
  CreateDirectory "$SMPROGRAMS\Exhibit Builder"
  CreateShortcut "$SMPROGRAMS\Exhibit Builder\Exhibit Builder.lnk" "$INSTDIR\Exhibit Builder.exe"
  CreateShortcut "$SMPROGRAMS\Exhibit Builder\Uninstall Exhibit Builder.lnk" "$INSTDIR\Uninstall Exhibit Builder.exe"
  CreateShortcut "$DESKTOP\Exhibit Builder.lnk" "$INSTDIR\Exhibit Builder.exe"

  WriteRegStr HKCU "Software\Exhibit Builder" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "DisplayName" "Exhibit Builder"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "Publisher" "Exhibit Builder"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "UninstallString" '"$INSTDIR\Uninstall Exhibit Builder.exe"'
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder" "NoRepair" 1

  install_complete:
SectionEnd

Section "Uninstall"
  SetShellVarContext current
  IfFileExists "$INSTDIR\.exhibit-builder-install-marker" 0 unsafe_uninstall
  IfFileExists "$INSTDIR\.exhibit-builder-verification-install" verification_uninstall normal_uninstall

  verification_uninstall:
    RMDir /r "$INSTDIR"
    Goto uninstall_complete

  normal_uninstall:
  Delete "$DESKTOP\Exhibit Builder.lnk"
  RMDir /r "$SMPROGRAMS\Exhibit Builder"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\Exhibit Builder"
  DeleteRegKey HKCU "Software\Exhibit Builder"
  RMDir /r "$INSTDIR"
  Goto uninstall_complete

  unsafe_uninstall:
    MessageBox MB_ICONSTOP "The Exhibit Builder installation marker is missing. No application files were removed."
    Abort

  uninstall_complete:
SectionEnd
