; installer written for NSIS version 2.46
!include MUI2.nsh
!include Library.nsh

Name "Code Help - VB6 IDE Extensions"
Outfile "CodeHelpInstaller.exe"
InstallDir "$ProgramFiles\CodeHelp"
RequestExecutionLevel admin

Function RegisterDotNet
  Exch $R0
  Push $R1

  ReadRegStr $R1 HKEY_LOCAL_MACHINE "Software\Microsoft\.NETFramework" "InstallRoot"

  IfFileExists "$R1\v4.0.30319\regasm.exe" FileExists
    MessageBox MB_ICONSTOP|MB_OK "Microsoft .NET Framework 4.x was not detected!"
  Abort

  FileExists:
  ExecWait '"$R1\v4.0.30319\regasm.exe" "$R0" /tlb /codebase /silent'

  Pop $R1
  Pop $R0
FunctionEnd

Function UnregisterDotNet
  Exch $R0
  Push $R1

  ReadRegStr $R1 HKEY_LOCAL_MACHINE "Software\Microsoft\.NETFramework" "InstallRoot"

  IfFileExists "$R1\v4.0.30319\regasm.exe" FileExists
    MessageBox MB_ICONSTOP|MB_OK "Microsoft .NET Framework 4.x was not detected!"
  Abort

  FileExists:
  ExecWait '"$R1\v4.0.30319\regasm.exe" "$R0" /unregister /silent'

  Pop $R1
  Pop $R0
FunctionEnd

!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Install the non-optional components
Section
  ; create the necessary subdirectories
  SetOutPath $InstDir\Interfaces
  SetOutPath $InstDir\Plugins
  SetOutPath $InstDir

  ; setup the stuff the CHCore depends on
  !insertmacro InstallLib RegDllTlb NotShared NoReboot_Protected \
    Interfaces\chlib.tlb $InstDir\Interfaces\CHLib.tlb $SYSDIR
  !insertmacro InstallLib RegDllTlb NotShared NoReboot_Protected \
    Interfaces\WinApiForVb.tlb $InstDir\Interfaces\WinApiForVb.tlb $SYSDIR

  ; install CHCore
  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\CHGlobalLib.dll $InstDir\CHGlobalLib.dll $SYSDIR
  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\CHCore.dll $InstDir\CHCore.dll $SYSDIR

  ; install the tabs plugin
  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\Plugins\CHTabMDI2.dll $InstDir\Plugins\CHTabMDI2.dll $SYSDIR

  ; tell windows about the uninstaller
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\CodeHelp" \
                 "DisplayName" "Code Help - VB6 IDE Extensions"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\CodeHelp" \
                 "UninstallString" "$\"$INSTDIR\uninstall.exe$\""
  WriteUninstaller uninstall.exe
SectionEnd

; install the fullscreen plugin
Section "Fullscreen"
  SetOutPath $InstDir

  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\Plugins\CHFullScreen.dll $InstDir\Plugins\CHFullScreen.dll $SYSDIR
SectionEnd

; install the complexity plugin
Section "Code Complexity"
  SetOutPath "$InstDir\3rd Party"
  SetOutPath $InstDir

  ; register antlr and the CodeAnalysis library
  File "/oname=3rd Party\Antlr4.Runtime.dll" Plugins\CHCodeComplexity\Antlr4.Runtime.dll
  File "/oname=3rd Party\CodeAnalysis.dll" Plugins\CHCodeComplexity\CodeAnalysis.dll
  Push "$InstDir\3rd Party\CodeAnalysis.dll"
  Call RegisterDotNet

  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\Plugins\CHCodeComplexity.dll $InstDir\Plugins\CHCodeComplexity.dll $SYSDIR
SectionEnd

; install the comment plugin
Section "Comment/Uncomment"
  SetOutPath $InstDir

  ; !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
  ;   CHCore\bin\Plugins\CHFullScreen.dll Plugins\CHFullScreen.dll $SYSDIR
SectionEnd

; install the snippets plugin
Section "Snippets"
  SetOutPath $InstDir

  !insertmacro InstallLib RegDll NotShared NoReboot_Protected \
    CHCore\bin\Plugins\CHCoder.dll Plugins\CHCoder.dll $SYSDIR
  File "/oname=Plugins\code_templates.mdb" CHCore\bin\Plugins\code_templates.mdb
SectionEnd


; uninstall everything
Section "Uninstall"
  Delete $InstDir\uninstall.exe

  ; remove interfaces
  !insertmacro UninstallLib RegDllTlb NotShared NoReboot_Protected $InstDir\Interfaces\CHLib.tlb
  !insertmacro UninstallLib RegDllTlb NotShared NoReboot_Protected $InstDir\Interfaces\WinApiForVb.tlb
  Delete "$InstDir\Interfaces\CHLib.tlb"
  Delete "$InstDir\Interfaces\WinApiForVb.tlb"
  RmDir  "$InstDir\Interfaces"

  ; remove CHCore
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\CHCore.dll
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\CHGlobalLib.dll
  Delete "$InstDir\CHCore.dll"
  Delete "$InstDir\CHGlobalLib.dll"

  ; remove plugins
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\Plugins\CHTabMDI2.dll
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\Plugins\CHFullScreen.dll
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\Plugins\CHCodeComplexity.dll
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\Plugins\CHCoder.dll
  !insertmacro UninstallLib RegDll NotShared NoReboot_Protected $InstDir\Plugins\CHComment.dll

  Delete "$InstDir\Plugins\CHTabMDI2.dll"
  Delete "$InstDir\Plugins\CHFullScreen.dll"
  Delete "$InstDir\Plugins\CHCoder.dll"
  Delete "$InstDir\Plugins\CHComment.dll"
  Delete "$InstDir\Plugins\CHCodeComplexity.dll"
  Delete "$InstDir\Plugins\code_templates.mdb"
  RmDir  "$InstDir\Plugins"

  Delete "$InstDir\3rd Party\Antlr4.Runtime.dll"
  Delete "$InstDir\3rd Party\CodeAnalysis.dll"
  Delete "$InstDir\3rd Party\CodeAnalysis.tlb"
  RmDir  "$InstDir\3rd Party"
SectionEnd
