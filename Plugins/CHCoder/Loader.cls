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

'Plugin const
Private Const CH_LONGNAME As String = "CodeHelp Code Complete"
Private Const CH_DESCRIPTION As String = "Template based code completion"
Private Const CH_COPYRIGHT As String = "luthv@yahoo.com"

Private Const VBA_WIN As String = "VbaWindow"

Private m_CHCorePtr As Long
Private m_IsEnabled As Boolean
Private m_CodeWinFocused As Boolean

Private m_MenuItem As CommandBarControl
Private WithEvents m_oMenuEvents As VBIDE.CommandBarEvents
Attribute m_oMenuEvents.VB_VarHelpID = -1

Private m_TemplateDataSet As Recordset
Private m_MarkerDataSet As Recordset
Private m_sConnection As String

Implements ICHPlugin

Private Sub Class_Initialize()
    Call InitData
    
    Set m_TemplateDataSet = OpenData("SELECT * FROM Snippet ORDER BY key")
    Set m_MarkerDataSet = OpenData("SELECT * FROM Marker")
    
End Sub

Private Sub Class_Terminate()
    Call CloseData
End Sub

Private Sub InitData()
    m_sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\code_templates.mdb;"
End Sub

Private Function GetConnection() As Connection
    Set GetConnection = New Connection
    GetConnection.ConnectionString = m_sConnection
    Call GetConnection.Open
End Function

Private Function GetRecordSet() As Recordset
    Set GetRecordSet = New Recordset
    GetRecordSet.CursorLocation = adUseClient
    GetRecordSet.CursorType = adOpenStatic
    GetRecordSet.LOCKTYPE = adLockBatchOptimistic
End Function

Private Function OpenData(query As String) As Recordset
    Dim oRecords As Recordset
    Dim oConnection As Connection
    
    On Error GoTo EXIT_POINT
    Set oConnection = GetConnection
    Set oRecords = GetRecordSet

    Set oRecords.ActiveConnection = oConnection
    Call oRecords.Open(query)
    Set oRecords.ActiveConnection = Nothing

    Set OpenData = oRecords
    
EXIT_POINT:
    On Error Resume Next
    Set oRecords.ActiveConnection = Nothing
    
    Call oConnection.Close
    Set oConnection = Nothing
    
    If Err.Number <> 0 Then
        Call MsgBox("Error while opening code_templates.mdb." & vbCrLf & Err.Description & vbCrLf & _
            "Please make sure that code_templates.mdb file is placed in the same folder as CHCoder.dll", _
            vbInformation, "CodeHelp Coder Error")
        
    End If
    Err.Clear
End Function

Friend Sub SaveData()
    Dim oConnection As Connection
    
    Set oConnection = GetConnection
    Set m_TemplateDataSet.ActiveConnection = oConnection
    Call m_TemplateDataSet.UpdateBatch
    Set m_TemplateDataSet.ActiveConnection = Nothing
    
    
    Call oConnection.Close
    Set oConnection = Nothing
End Sub

Private Sub CloseData()
    Set m_TemplateDataSet = Nothing
    Set m_MarkerDataSet = Nothing
End Sub


Private Property Let ICHPlugin_CHCore(ByVal RHS As Long)
    'Save the Pointer for later use
    m_CHCorePtr = RHS
End Property

Private Property Get ICHPlugin_CopyRight() As String
    ICHPlugin_CopyRight = CH_COPYRIGHT
End Property

Private Property Get ICHPlugin_Description() As String
    ICHPlugin_Description = CH_DESCRIPTION
End Property

Private Property Let ICHPlugin_Enabled(ByVal Enabled As Boolean)
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
    m_IsEnabled = Enabled
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
    ICHPlugin_Enabled = m_IsEnabled
End Property

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean
'Enable/disable help button in plugin manager
End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean
'Enable/disable properties button in plugin manager
    ICHPlugin_HaveProperties = True
End Property

Private Property Get ICHPlugin_LongName() As String
    ICHPlugin_LongName = CH_LONGNAME
End Property

Private Property Get ICHPlugin_Name() As String
    ICHPlugin_Name = App.Title
End Property

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, custom() As Variant)
    'Sample use of the ICHCore pointer
    'It's advisable not to save the ICHCore object itself, always use the helper function to obtain the
    'ICHCore object from the pointer
    Dim hWnd As Long
    Dim oCHCore As ICHCore
    Dim sMakeItWork As String
    
    hWnd = GetFocus()
    
    If GetWinText(hWnd, True) = VBA_WIN Then
        If InStr(1, GetWinText(hWnd), "(Code)") Then
            m_CodeWinFocused = True
        End If
    End If
    
    Set oCHCore = CHHelper.GetCHCore(m_CHCorePtr)
    If m_MenuItem Is Nothing Then
        sMakeItWork = oCHCore.VBE.Name
        Set m_MenuItem = oCHCore.AddToCodeHelpMenu("Snippets", Nothing)
        Set m_oMenuEvents = oCHCore.VBE.Events.CommandBarEvents(m_MenuItem)
    End If
    
    Set oCHCore = Nothing
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, custom() As Variant)
    Call CloseData
End Sub

Private Function ShiftPressed() As Boolean
 ShiftPressed = (GetKeyState(vbKeyShift) And &H8000) <> 0
End Function

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, lReturn As Long, wParam As Long, lParam As Long)
    If m_CodeWinFocused Then
        If Not ShiftPressed Then Exit Sub  'ignore when the shift key is not down
        If wParam = vbKeySpace Then
            If CHHelper.IsKeyDownEvent(lParam) Then 'only process keydown
                If ProcessLine Then
                    lReturn = 1
                    bHandled = True
                End If
            End If
        End If
    End If
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, _
                                bHandled As Boolean, lReturn As Long)
    'subclassed message goes here
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                                    bHandled As Boolean, placeSubclass As Boolean, before As Boolean)
    'Hook msg goes here
    Select Case uMsg
        Case WM_SETFOCUS
            m_CodeWinFocused = False
            If GetWinText(hWnd, True) = VBA_WIN Then
                If InStr(1, GetWinText(hWnd), "(Code)") Then
                    m_CodeWinFocused = True
                End If
            End If
    End Select
End Sub

Private Sub ICHPlugin_ShowHelp()
'Show help instruction when user click on help button in plugin manager
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
    Dim frmProperties As frmProp
    Set frmProperties = New frmProp
    
    Call frmProperties.Initalize(m_TemplateDataSet, m_MarkerDataSet, Me)
    Call frmProperties.Show(vbModal)
    Call Unload(frmProperties)

    Set frmProperties = Nothing
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub m_oMenuEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call ICHPlugin_ShowPropertyDialog
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

Private Function ProcessLine() As Boolean
    Dim lStartLine As Long, lStartCol As Long
    Dim lEndLine As Long, lEndCol As Long
    Dim sLine As String, sMatch As String
    Dim sTemplate As String, sIndent As String
    Dim lMatchLength As Long
    Dim oActiveCodePane As CodePane

    Set oActiveCodePane = GetCHCore(m_CHCorePtr).VBE.ActiveCodePane
    
    If m_TemplateDataSet.RecordCount < 1 Then Exit Function
    
    Call oActiveCodePane.GetSelection(lStartLine, lStartCol, lEndLine, lEndCol)
    sLine = oActiveCodePane.CodeModule.Lines(lStartLine, 1)
    
    'parse the current line for the target
    Do While Mid$(sLine, lStartCol, 1) <> " " And lStartCol > 1
        lStartCol = lStartCol - 1
    Loop
    If Mid$(sLine, lStartCol, 1) = " " Then lStartCol = lStartCol + 1
    Do While Mid$(sLine, lEndCol, 1) <> " " _
            And Mid$(sLine, lEndCol, 1) <> vbCr _
            And lEndCol < Len(sLine)
        lEndCol = lEndCol + 1
    Loop
    
    sMatch = Mid$(sLine, lStartCol, lEndCol - lStartCol)
    lMatchLength = Len(sMatch)
    Call m_TemplateDataSet.MoveFirst
    Call m_TemplateDataSet.Find("Key='" & sMatch & "'")
    
    If m_TemplateDataSet.EOF = False Then
        sLine = Left$(sLine, Len(sLine) - lMatchLength)
        sTemplate = m_TemplateDataSet.Fields(1).Value
        
        'match indentation when the template is multi-line
        sIndent = vbCrLf & IndentationToMatchLine(sLine)
        sTemplate = Replace(sTemplate, vbCrLf, sIndent)

        ProcessLine = GetNewSelection(sLine, sTemplate, lStartLine, lStartCol, lEndLine, lEndCol)
        sLine = sLine & sTemplate
        Call oActiveCodePane.CodeModule.ReplaceLine(lStartLine, sLine)
        Call oActiveCodePane.SetSelection(lStartLine, lStartCol, lEndLine, lEndCol)
    End If
End Function

'if GetNewSelection returns true then we don't insert a space
Private Function GetNewSelection(sLine As String, sTemplate As String, lStartLine As Long, lStartCol As Long, _
                                                                       lEndLine As Long, lEndCol As Long) As Boolean
    Dim lLineLen        As Long
    Dim lSelStart       As Long
    Dim lSelEnd         As Long
    
    lSelStart = InStr(1, sTemplate, "<sel>")
    If lSelStart = 0 Then 'if no selection is specified then just do something default
        GetNewSelection = False
        lStartCol = InStr(1, sTemplate, vbCrLf)
        If lSelStart = 0 Then
            lStartCol = Len(sLine) + Len(sTemplate) + 1 'move to just after the template
        Else
            lStartCol = Len(sLine) + lSelStart 'move to the end of the first line inserted
        End If
        lEndCol = lStartCol
        lEndLine = lStartLine
        
    Else 'figure out where the selection is supposed to go
        GetNewSelection = True
        lSelEnd = InStr(1, sTemplate, "</>")
        lLineLen = InStr(1, sTemplate, vbCrLf)
        If lLineLen = 0 Then 'single line template
            lEndLine = lStartLine
            lStartCol = Len(sLine) + lSelStart
            lEndCol = Len(sLine) + lSelEnd + 3 '+3 to move past the selection delim
        Else
            
        End If
        
    End If
End Function

Private Function IndentationToMatchLine(sLine As String) As String
    Dim sIndent As String
    Dim lIndentLevel As Long
    
    lIndentLevel = 1
    Do While Mid$(sLine, lIndentLevel, 1) = " "
        lIndentLevel = lIndentLevel + 1
    Loop
    lIndentLevel = lIndentLevel - 1 'compensate for 1-based indicies
    
    sIndent = ""
    Do While lIndentLevel > 0
        sIndent = sIndent & " "
        lIndentLevel = lIndentLevel - 1
    Loop
    
    IndentationToMatchLine = sIndent
End Function

