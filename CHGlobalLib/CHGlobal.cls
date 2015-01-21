VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'***********************************
'CodeHelp Helper class
'***********************************

Private Const ccCFBitmap = 2
Private Const ccCFMetafile = 3
Private Const ccCFDIB = 8

Private Declare Function CreateIC Lib "gdi32" Alias "CreateICA" _
     (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
     ByVal lpOutput As String, ByVal lpInitData As String) As Long


Function GetCHCore(ByVal lPtr As Long) As ICHCore
  Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
  If lPtr = 0 Then Exit Function
  
  If IsBadReadPtr(lPtr, ByVal 4) Then Exit Function  ' better to be safe
  CopyMemory oTemp, lPtr, 4
  
  ' Assign to legal reference
  Set GetCHCore = oTemp
  
  ' Destroy the illegal reference
  CopyMemory oTemp, 0&, 4
End Function

Sub CopyIconToClipBoardAsBmp(oIcon As StdPicture, oBMP As StdPicture)
    Dim Rc As Long
    Dim hdc As Long
    Dim hdcMem As Long
    Dim hBmOld As Long
    
    hdc = CreateIC("DISPLAY", vbNullChar, vbNullChar, vbNullChar)
    hdcMem = CreateCompatibleDC(hdc)
    hBmOld = SelectObject(hdcMem, oBMP.Handle)
    Rc = DrawIconEx(hdcMem, 0, 0, oIcon.Handle, 16, 16, 0, 0, DI_NORMAL)
    SelectObject hdcMem, hBmOld
    DeleteDC hdc
    DeleteDC hdcMem
        
    Clipboard.Clear
    Clipboard.SetData oBMP, ccCFBitmap
    Clipboard.SetData oBMP, ccCFDIB

End Sub

Sub LogToNotePad(ByVal sMsg As String)
    'Help trace in compiled mode
    'The message will be appended to first notepad instance found
    
    Dim hNote As Long
    hNote = A_FindWindowEx(0, 0, "Notepad", "Untitled - Notepad")
    
    If hNote <> 0 Then
        hNote = A_FindWindowEx(hNote, 0, "Edit", vbNullString)
        A_SendMessageStr hNote, EM_REPLACESEL, 0, sMsg
    End If
End Sub
