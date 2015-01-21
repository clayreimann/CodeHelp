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

Private m_lPtr As Long
Private m_bEnabled As Boolean

Dim bMonitor As Boolean
Dim ActiveCodePane As CodePane

Implements ICHPlugin

Private Property Let ICHPlugin_CHCore(ByVal RHS As Long)
    'Save the Pointer for later use
    m_lPtr = RHS
End Property

Private Property Get ICHPlugin_CopyRight() As String
    ICHPlugin_CopyRight = CH_COPYRIGHT
End Property

Private Property Get ICHPlugin_Description() As String
    ICHPlugin_Description = CH_DESCRIPTION
End Property

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   m_bEnabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   ICHPlugin_Enabled = m_bEnabled
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

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, _
                                   custom() As Variant)
    
    'Sample use of the ICHCore pointer
    'It's advisable not to save the ICHCore object itself, always use the helper function to obtain the
    'ICHCore object from the pointer
    Dim hWnd As Long
    
    hWnd = GetFocus()
    
    OpenData
    If GetWinText(hWnd, True) = VBA_WIN Then
        
        If InStr(1, GetWinText(hWnd), "(Code)") Then
            bMonitor = True
        End If
        
    End If
    
    'Do your initializing stuff here
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    'Do your clean up here
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lreturn As Long, _
                                wParam As Long, _
                                lParam As Long)
    'do any keyboard related code here
    If bMonitor Then
        If wParam = vbKeySpace Then
            If (lParam And &H80000000) = 0 Then 'only process keydown
                ProcessLine
'                    lreturn = 1
'                    bHandled = True
            End If
        End If
    End If
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, _
                                ByVal uMsg As Long, _
                                wParam As Long, _
                                lParam As Long, _
                                bHandled As Boolean, _
                                lreturn As Long)
    'subclassed message goes here
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)
    'Hook msg goes here
    
    Select Case uMsg
        Case WM_SETFOCUS ', WM_CHILDACTIVATE
            
            bMonitor = False
            
            If GetWinText(hWnd, True) = VBA_WIN Then
                
                If InStr(1, GetWinText(hWnd), "(Code)") Then
                    bMonitor = True
                End If
                
            End If
            
    End Select
End Sub

Private Sub ICHPlugin_ShowHelp()
'Show help instruction when user click on help button in plugin manager
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
'show property dialog when user click on properties button in plugin manager
    Dim prop As frmProp
    Set prop = New frmProp
    prop.Show vbModal
    Set prop = Nothing
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

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
    Dim lStartLine As Long, lEndLine As Long, lStartCol As Long, lEndCol As Long
    Dim sCodes As String
    Dim sWords() As String
    Dim templateLenght As Long
    Dim sTemplate As String
    
    Set ActiveCodePane = GetCHCore(m_lPtr).VBE.ActiveCodePane
        
    If templateData.RecordCount > 0 Then
    
        With ActiveCodePane
            .GetSelection lStartLine, lStartCol, lEndLine, lEndCol
            sCodes = .CodeModule.Lines(lStartLine, 1)

            sWords = Split(sCodes, " ")
            sTemplate = Trim$(sWords(UBound(sWords)))
            templateLenght = Len(sTemplate)

            templateData.MoveFirst
            templateData.Find "Key='" & sTemplate & "'"
            
            If templateData.EOF = False Then
                sCodes = Left$(sCodes, Len(sCodes) - templateLenght) + _
                    templateData.Fields(1).Value
                .CodeModule.ReplaceLine lStartLine, sCodes
                templateLenght = Len(sCodes) + 1
                .SetSelection lStartLine, templateLenght, lStartLine, templateLenght
                
                ProcessLine = True
            End If
    
        End With
    End If
End Function
