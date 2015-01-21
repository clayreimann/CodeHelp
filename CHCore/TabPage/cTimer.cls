VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cTimer - module-less, IDE safe, machine code timer thunk
'
'Paul_Caton@hotmail.com
'Copyright free, use and abuse as you see fit.
'
'v1.00 20030107 First cut..........................................................................
'v1.01 20031118 Allow control over callback gating
'               Use global memory for the machine code buffer
'               Reform the assembler...............................................................
'v1.02 20040118 Use EbMode for breakpoint/stop detection rather than callback gating
'               Further reform the assembler for greater speed and smaller size
'               Made InIDE public..................................................................
'v1.03 20040210 Added support for an Optional ID to indentify amongst multiple timers..............
'==================================================================================================
Option Explicit

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private nTimerID As Long                                    'Timer ID
Private nAddrTmr As Long                                    'The address of our timer thunk

'============================================
'Class creation/destruction
'============================================

'Build the timer thunk into allocated memory
Private Sub Class_Initialize()
Const PATCH_01 As Long = 3                                  'Code buffer offset to the location of the relative address to EbMode
Const PATCH_05 As Long = 52                                 'Code buffer offset to the location of the relative address to KillTimer
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk know if the IDE has stopped or is on a breakpoint
Const FUNC_KIL As String = "KillTimer"                      'KillTimer allows the cTimer machine code thunk to kill the timer itself if it detects via the EbMode function that the IDE has stopped
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const MOD_USER As String = "user32"                         'Location of the KillTimer function
  Dim i        As Long                                      'Loop index
  Dim nLen     As Long                                      'String length
  Dim sHex     As String                                    'Hex code representation of the machine code
  Dim sCode    As String                                    'ASCII string, holds the machine code before copying to allocated memory

  'Store the hex pair machine code representation in sHex
  sHex = "EB0EE8xxxxx01x83F802741E85C0741D68xxxxx02x8B4424142Dxxxxx03x50B8xxxxx04x508B00FF501CC210008B4C240C5150E8xxxxx05xEBF0"
  nLen = Len(sHex)                                          'Length of the hex pair string

  'Convert the code string from hex pairs to bytes and store in the ASCII string opcode buffer
  For i = 1 To nLen Step 2                                  'For each pair of hex characters
    sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))     'Convert a pair of hex characters to a byte and append to the ASCII string
  Next i                                                    'Next pair

  nLen = LenB(sCode)                                        'Get the machine code length
  nAddrTmr = GlobalAlloc(0, nLen)               'Allocate fixed memory for for the machine code buffer

  'Copy the code to allocated memory
  Call CopyMemory(ByVal nAddrTmr, ByVal StrPtr(sCode), nLen)

  If InIDE Then
    'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
    Call CopyMemory(ByVal nAddrTmr, &H9090, 2)

    i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll
    If i = 0 Then                                           'Found?
      i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
    End If

    Debug.Assert i                                          'Ensure the EbMode function was found
    Call PatchRel(PATCH_01, i)                              'Patch the relative address to the EbMode api function
  End If
  
  Call PatchRel(PATCH_05, AddrFunc(MOD_USER, FUNC_KIL))     'Patch the relative address to the KillTimer api function
End Sub

'Stop the timer and release the allocated memory
Private Sub Class_Terminate()
  Call Me.TmrStop                                           'Kill the timer
  Call GlobalFree(nAddrTmr)                     'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'Return whether we're running in the IDE. Public for general utility purposes
Public Function InIDE() As Boolean
  Debug.Assert SetTrue(InIDE)
End Function

'Start the timer
Public Function TmrStart(ByVal Owner As ITimer, _
                         ByVal nInterval As Long, _
                         Optional ByVal nID As Long = 0) As Boolean
Const PATCH_02 As Long = 17                                 'Timer ID
Const PATCH_03 As Long = 26                                 'Code buffer offset to the location of the timer start time
Const PATCH_04 As Long = 32                                 'Code buffer offset to the location of the owner object address

  If nTimerID = 0 Then
    Call PatchVal(PATCH_02, nID)                            'Set the Timer ID - Allows the programmer to distinguish amongst multiple timers
    Call PatchVal(PATCH_03, GetTickCount)       'Set the start time
    Call PatchVal(PATCH_04, ObjPtr(Owner))                  'Owner object address for iTimer_Proc

    'Create the timer
    nTimerID = SetTimer(0, 0, nInterval, nAddrTmr)
    TmrStart = (nTimerID <> 0)                              'nTimer ID is non-zero on success
  End If
    If TmrStart = False Then
        Debug.Print "Timer Failed"
    End If
  'Debug.Assert TmrStart                                     'Let programmer know if TmrStart failed
End Function

'Stop the timer
Public Function TmrStop() As Boolean
  If nTimerID <> 0 Then
    If KillTimer(0, nTimerID) <> 0 Then         'Kill the timer
      nTimerID = 0                                          'Indicate the timer is inactive
      TmrStop = True
    End If
  Else
    TmrStop = True                                          'Timer wasn't running, but that's no cause for complaint
  End If

  Debug.Assert TmrStop                                      'Let programmer know if TmrStop failed
End Function

'============================================
'Private interface
'============================================

'Return the address of the passed function in the passed dll
Private Function AddrFunc(ByVal sDLL As String, _
                          ByVal sProc As String) As Long
  AddrFunc = GetProcAddress(A_GetModuleHandle(sDLL), sProc)
  
  'You may want to comment out the following line if you're using vb5 else the EbMode
  'GetProcAddress will stop here everytime because we look in vba6.dll first
  Debug.Assert AddrFunc
End Function

'Patch the machine code buffer offset with the relative address to the target address
Private Sub PatchRel(ByVal nOffset As Long, _
                     ByVal nTargetAddr As Long)
  Call CopyMemory(ByVal (nAddrTmr + nOffset), nTargetAddr - nAddrTmr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer offset with the passed value
Private Sub PatchVal(ByVal nOffset As Long, _
                     ByVal nValue As Long)
  Call CopyMemory(ByVal (nAddrTmr + nOffset), nValue, 4)
End Sub

'Worker function for InIDE - will only be called whilst running in the IDE
Private Function SetTrue(ByRef bValue As Boolean) As Boolean
  bValue = True
  SetTrue = True
End Function
