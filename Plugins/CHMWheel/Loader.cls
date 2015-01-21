VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SCROLL_CLASS = "ScrollBar"
Private m_lPtr As Long
Private m_hMDIClient As Long
Private m_Enabled As Boolean

Private scrollLines As Long

Implements ICHPlugin

Private Property Let ICHPlugin_CHCore(ByVal RHS As Long)
    m_lPtr = RHS
End Property

Private Property Get ICHPlugin_CopyRight() As String
    ICHPlugin_CopyRight = "luthv@yahoo.com"
End Property

Private Property Get ICHPlugin_Description() As String
    ICHPlugin_Description = "Provide support for mousewheel in VB Code editor."
End Property

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
    m_Enabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
    ICHPlugin_Enabled = m_Enabled
End Property

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean

End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean
    ICHPlugin_HaveProperties = True
End Property

Private Property Get ICHPlugin_LongName() As String
    ICHPlugin_LongName = "CodeHelp MouseWheel Support"
End Property

Private Property Get ICHPlugin_Name() As String
    ICHPlugin_Name = App.Title
End Property

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, custom() As Variant)
    Dim Dsr As ICHCore
    Set Dsr = GetCHCore(m_lPtr)
    m_hMDIClient = Dsr.VBE.MainWindow.hWnd
    m_hMDIClient = A_FindWindowEx(m_hMDIClient, 0, "MDIClient", vbNullString)
    
    scrollLines = GetSetting("CodeHelp", ICHPlugin_Name, "ScrollLines", 3)
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, custom() As Variant)

End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, lreturn As Long, wParam As Long, lParam As Long)

End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bHandled As Boolean, lreturn As Long)

End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, placeSubclass As Boolean, before As Boolean)
    Select Case uMsg
        Case WM_MOUSEWHEEL
            If hWnd = m_hMDIClient Then OnMouseWheel wParam, lParam
    End Select
End Sub

Private Sub ICHPlugin_ShowHelp()

End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
    Dim dlg As frmProperties
    Dim newLines As Long
    
    On Error GoTo ERR_HANDLER
    
    Set dlg = New frmProperties
    dlg.txtNumber.Text = scrollLines
    dlg.Show vbModal
    
    newLines = CLng(dlg.txtNumber.Text)
    If newLines > 0 Then
        If newLines <> scrollLines Then
            scrollLines = newLines
            SaveSetting "CodeHelp", ICHPlugin_Name, "ScrollLines", CStr(newLines)
        End If
    End If
    
ERR_HANDLER:
    If Not dlg Is Nothing Then
        On Error Resume Next
        Unload dlg
    End If
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub OnMouseWheel(ByVal wParam As Long, ByVal lParam As Long)
    Dim delta As Long
    Dim lStep As Long
    Dim hActive As Long, hScroll As Long
    Dim lStyle As Long
    Dim i As Long
    

    hActive = A_SendMessage(m_hMDIClient, WM_MDIGETACTIVE, 0, ByVal 0)
    
    hScroll = GetWindow(hActive, GW_CHILD)
    
    Do While hScroll <> 0
        If GetWinText(hScroll, True) = SCROLL_CLASS Then
            lStyle = A_GetWindowLong(hScroll, GWL_STYLE)
            If (lStyle And SBS_VERT) = SBS_VERT Then
                If (lStyle And WS_VISIBLE) = WS_VISIBLE Then
                    Exit Do
                End If
            End If
        End If
        hScroll = GetWindow(hScroll, GW_HWNDNEXT)
    Loop
    
    If hScroll <> 0 Then
        
        delta = HiWord(wParam)
                
        If delta > 0 Then
            'lStart = delta * scrollLines
            'For i = -120 To delta Step 120
            For i = 1 To scrollLines
                A_SendMessage hActive, WM_VSCROLL, MakeDWord(SB_LINEUP, 0), ByVal hScroll
            Next
        Else
            'For i = delta - 120 To 0 Step 120
            For i = 1 To scrollLines
                A_SendMessage hActive, WM_VSCROLL, MakeDWord(SB_LINEDOWN, 0), ByVal hScroll
            Next
        End If
        
    End If
End Sub

Private Function GetWinText(hWnd As Long, Optional className As Boolean = False) As String
    'some static vars to speed up things, this func will be called many times
    Static sBuffer As String * 128& 'is it safe to use 128 bytes? should be enough..
    Static textLength As Long
  
    If className Then
        textLength = A_GetClassName(hWnd, sBuffer, 129&)
    Else
        textLength = A_GetWindowText(hWnd, sBuffer, 129&)
    End If
  
    If textLength > 0 Then
        GetWinText = Left$(sBuffer, textLength)
    End If

End Function

Private Function HiWord(lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
' by Karl E. Peterson, http://www.mvps.org/vb, 20001207
  ' High word is coerced to Long to allow it to
  ' overflow limits of multiplication which shifts
  ' it left.
  MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

