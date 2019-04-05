;!include nsDialogs.nsh
;!include LogicLib.nsh


; example1.nsi
;
; This script is perhaps one of the simplest NSIs you can make. All of the
; optional settings are left to their default settings. The installer simply 
; prompts the user asking them where to install, and drops a copy of example1.nsi
; there. 

XPStyle on

;--------------------------------

; The name of the installer
Name "TestRail Test Cases to HTML"

; The file to write
OutFile "setup.exe"

; The default installation directory
InstallDir "C:\TestRail Test Cases to HTML"

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------


; Pages

Page directory
Page instfiles


;--------------------------------


; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File "TestRail Test Cases to HTML.exe"
  File "curl.exe"
  File "multimarkdown.exe"
  File "7z.exe"
  File *.dll

  CreateDirectory "$SMPROGRAMS\TestRail Test Cases to HTML"
  CreateShortCut "$SMPROGRAMS\TestRail Test Cases to HTML\TestRail Test Cases to HTML.lnk" "$INSTDIR\TestRail Test Cases to HTML.exe"

SectionEnd ; end the section
