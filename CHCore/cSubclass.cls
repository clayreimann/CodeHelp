VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'==================================================================================================
'cSubclass - module-less, IDE safe, machine code subclassing thunk
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
'
'==================================================================================================
Option Explicit
Private Declare Function IsBadReadPtr _
                Lib "kernel32.dll" (lp As Any, _
                                    ByVal ucb As Long) As Long

Private Const PATCH_05  As Long = 93                        'Table B (before) entry count
Private Const PATCH_09  As Long = 137                       'Table A (after) entry count

Private nMsgCntB        As Long                             'Before msg table entry count
Private nMsgCntA        As Long                             'After msg table entry count
Private aMsgTblB()      As Long                 'Before msg table array
Private aMsgTblA()      As Long                 'After msg table array
Private hWndSub         As Long                             'Handle of the window being subclassed
Private nAddrSubclass   As Long                             'The address of our WndProc
Private nAddrOriginal   As Long                             'The address of the existing WndProc

Public Enum enMsgWhen
    MSG_AFTER = 1&
    MSG_BEFORE = 2&
    MSG_BEFORE_AND_AFTER = 3&
    ALL_MESSAGES = &HFFFFFFFF
End Enum

'============================================
'Class creation/destruction
'============================================

'Build the subclass thunk into allocated memory
Private Sub Class_Initialize()
    Const PATCH_01 As Long = 18                                 'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_03 As Long = 78                                 'Relative address of SetWindowsLong
    Const PATCH_07 As Long = 121                                'Relative address of CallWindowProc
    Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"                 'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const FUNC_CWP As String = "CallWindowProcA"                'We use CallWindowProc to call the original WndProc
    Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
    Const MOD_USER As String = "user32"                         'Location of the SetWindowLong & CallWindowProc functions
    Dim i        As Long                                      'Loop index
    Dim nLen     As Long                                      'String lengths
    Dim sHex     As String                                    'Hex code string
    Dim sCode    As String                                    'Binary code string
  
    'Store the hex pair machine code representation in sHex
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE8xxxxx01x83F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168xxxxx02x6AFCFF7508E8xxxxx03xEBE031D24ABFxxxxx04xB9xxxxx05xE82D000000C3FF7514FF7510FF750CFF750868xxxxx06xE8xxxxx07x8945FCC331D2BFxxxxx08xB9xxxxx09xE801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B8xxxxx0Ax508B00FF501CC3"
    nLen = Len(sHex)                                          'Length of hex pair string
  
    'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
    For i = 1 To nLen Step 2                                  'For each pair of hex characters
        sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))     'Convert a pair of hex characters to a byte and append to the ASCII string
    Next i                                                    'Next pair
  
    nLen = LenB(sCode)                                        'Get the machine code length
    nAddrSubclass = GlobalAlloc(0, nLen)                      'Allocate fixed memory for machine code buffer

    'Copy the code to allocated memory
    Call CopyMemory(ByVal nAddrSubclass, ByVal StrPtr(sCode), nLen)

    If InIDE Then
        'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
        Call CopyMemory(ByVal nAddrSubclass + 15, &H9090, 2)
    
        i = AddrFunc(MOD_VBA6, FUNC_EBM)                        'Get the address of EbMode in vba6.dll

        If i = 0 Then                                           'Found?
            i = AddrFunc(MOD_VBA5, FUNC_EBM)                      'VB5 perhaps, try vba5.dll
        End If

        '    Debug.Assert i                                          'Ensure the EbMode function was found
        Call PatchRel(PATCH_01, i)                              'Patch the relative address to the EbMode api function
    End If
  
    Call PatchRel(PATCH_03, AddrFunc(MOD_USER, FUNC_SWL))     'Address of the SetWindowLong api function
    Call PatchVal(PATCH_05, 0)                                'Initial before table entry count
    Call PatchRel(PATCH_07, AddrFunc(MOD_USER, FUNC_CWP))     'Address of the CallWindowProc api function
    Call PatchVal(PATCH_09, 0)                                'Initial after table entry count
End Sub

'UnSubclass and release the allocated memory
Private Sub Class_Terminate()
    On Error Resume Next
    Call Me.UnSubclass                                        'UnSubclass if the Subclass thunk is active
    Call GlobalFree(nAddrSubclass)                            'Release the allocated memory
End Sub

'============================================
'Public interface
'============================================

'Add the message to the callback table
Public Sub AddMsg(ByVal uMsg As Long, _
                  ByVal When As enMsgWhen)

    If When And MSG_BEFORE Then                   'If Before
        'Add the message, pass the before table and before table message count variables ByRef
        Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, MSG_BEFORE)
    End If
  
    If When And MSG_AFTER Then                    'If After
        'Add the message, pass the after table and after table message count variables ByRef
        Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, MSG_AFTER)
    End If

End Sub

'Arbitarily call the original WndProc
Public Function CallOrigWndProc(ByVal uMsg As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long

    If hWndSub <> 0 Then
        CallOrigWndProc = A_CallWindowProc(nAddrOriginal, hWndSub, uMsg, wParam, lParam)  'Call the original WndProc
    End If

End Function

'Delete the message from the msg table
Public Sub DelMsg(ByVal uMsg As Long, _
                  ByVal When As enMsgWhen)
                  
    If When And MSG_BEFORE Then                   'If before
        'Delete the message, pass the Before table and before message count variables ByRef
        Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, MSG_BEFORE)
    End If
  
    If When And MSG_AFTER Then                    'If After
        'Delete the message, pass the After table and after message count variables ByRef
        Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, MSG_AFTER)
    End If

End Sub

'Return whether we're running in the IDE. Public for general utility purposes
Public Function InIDE() As Boolean
    #If IN_ADDIN = 0 Then
    InIDE = (GetModuleHandle("vba6") <> 0)
    #End If
End Function

'Set the window subclass
Public Function Subclass(ByVal hwnd As Long, _
                         ByVal Owner As ISubclass) As Boolean
    Const PATCH_02 As Long = 68                                 'Address of the previous WndProc
    Const PATCH_06 As Long = 116                                'Address of the previous WndProc
    Const PATCH_0A As Long = 186                                'Address of the owner object
  
    If hWndSub = 0 Then
        Debug.Assert IsWindow(hwnd)                             'Invalid window handle
        hWndSub = hwnd                                          'Store the window handle
    
        'Get the original window proc
        nAddrOriginal = A_GetWindowLong(hwnd, GWL_WNDPROC)
        Call PatchVal(PATCH_02, nAddrOriginal)                  'Original WndProc address for CallWindowProc, call the original WndProc
        Call PatchVal(PATCH_06, nAddrOriginal)                  'Original WndProc address for SetWindowLong, unsubclass on IDE stop
        Call PatchVal(PATCH_0A, ObjPtr(Owner))                  'Owner object address for iSubclass_Proc
    
        'Set our WndProc in place of the original
        nAddrOriginal = A_SetWindowLong(hwnd, GWL_WNDPROC, nAddrSubclass)

        If nAddrOriginal <> 0 Then
            Subclass = True                                       'Success
        End If
    End If
  
    Debug.Assert Subclass
End Function

'Stop subclassing the window
Public Function UnSubclass() As Boolean

    If hWndSub <> 0 Then
        Call PatchVal(PATCH_05, 0)                              'Patch the Table B entry count to ensure no further iSubclass_Proc callbacks
        Call PatchVal(PATCH_09, 0)                              'Patch the Table A entry count to ensure no further iSubclass_Proc callbacks
    
        'Restore the original WndProc
        ' try to unsubclass
        If IsWindow(hWndSub) Then
            If A_GetWindowLong(hWndSub, GWL_WNDPROC) = nAddrSubclass Then
                If Not IsBadReadPtr(nAddrOriginal, 4&) Then
                    Call A_SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
                End If
            End If
        End If

        'Call SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
    
        hWndSub = 0                                             'Indicate the subclasser is inactive
        nMsgCntB = 0                                            'Message before count equals zero
        nMsgCntA = 0                                            'Message after count equals zero
        UnSubclass = True                                       'Success
    End If
  
    '  Debug.Assert UnSubclass
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
    '  Debug.Assert AddrFunc
End Function

'Worker sub for AddMsg
Private Sub AddMsgSub(ByVal uMsg As Long, _
                      ByRef aMsgTbl() As Long, _
                      ByRef nMsgCnt As Long, _
                      ByVal When As enMsgWhen)
    Const PATCH_04 As Long = 88                                 'Table B (before) address
    Const PATCH_08 As Long = 132                                'Table A (after) address
    Dim nEntry   As Long
    Dim nOff1    As Long
    Dim nOff2    As Long
  
    If uMsg = ALL_MESSAGES Then                   'If ALL_MESSAGES
        nMsgCnt = -1                                            'Indicates that all messages shall callback
    Else                                                      'Else a specific message number

        For nEntry = 1 To nMsgCnt                               'For each existing entry. NB will skip if nMsgCnt = 0

            Select Case aMsgTbl(nEntry)                           'Select on the message number stored in this table entry

                Case -1                                               'This msg table slot is a deleted entry
                    aMsgTbl(nEntry) = uMsg                              'Re-use this entry
                    Exit Sub                                            'Bail

                Case uMsg                                             'The msg is already in the table!
                    Exit Sub                                            'Bail
            End Select

        Next nEntry                                             'Next entry

        'Make space for the new entry
        ReDim Preserve aMsgTbl(1 To nEntry)                     'Increase the size of the table. NB nEntry = nMsgCnt + 1
        nMsgCnt = nEntry                                        'Bump the entry count
        aMsgTbl(nEntry) = uMsg                                  'Store the message number in the table
    End If
  
    If When = MSG_BEFORE Then                     'If before
        nOff1 = PATCH_04                                        'Offset to the Before table address
        nOff2 = PATCH_05                                        'Offset to the Before table entry count
    Else                                                      'Else after
        nOff1 = PATCH_08                                        'Offset to the After table address
        nOff2 = PATCH_09                                        'Offset to the After table entry count
    End If

    'Patch the appropriate table entries
    Call PatchVal(nOff1, AddrMsgTbl(aMsgTbl))                 'Patch the appropriate table address. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a message number is added.
    Call PatchVal(nOff2, nMsgCnt)                             'Patch the appropriate table entry count
End Sub

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl(ByRef aMsgTbl() As Long) As Long
    On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
    AddrMsgTbl = VarPtr(aMsgTbl(1))                       'Get the address of the first element of the passed message table
    On Error GoTo 0                                         'Switch off error protection
End Function

'Worker sub for DelMsg
Private Sub DelMsgSub(ByVal uMsg As Long, _
                      ByRef aMsgTbl() As Long, _
                      ByRef nMsgCnt As Long, _
                      ByVal When As enMsgWhen)
    Dim nEntry As Long
  
    If uMsg = ALL_MESSAGES Then                   'If deleting all messages (specific or ALL_MESSAGES)
        nMsgCnt = 0                                             'Message count is now zero

        If When = MSG_BEFORE Then                   'If before
            nEntry = PATCH_05                                     'Patch the before table message count location
        Else                                                    'Else after
            nEntry = PATCH_09                                     'Patch the after table message count location
        End If

        Call PatchVal(nEntry, 0)                                'Patch the table message count
    Else                                                      'Else deleteting a specific message

        For nEntry = 1 To nMsgCnt                               'For each table entry

            If aMsgTbl(nEntry) = uMsg Then                        'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = -1                                'Mark the table slot as available
                Exit For                                            'Bail
            End If

        Next nEntry                                             'Next entry

    End If

End Sub

'Patch the machine code buffer offset with the relative address to the target address
Private Sub PatchRel(ByVal nOffset As Long, _
                     ByVal nTargetAddr As Long)
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
End Sub

'Patch the machine code buffer offset with the passed value
Private Sub PatchVal(ByVal nOffset As Long, _
                     ByVal nValue As Long)
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
End Sub

