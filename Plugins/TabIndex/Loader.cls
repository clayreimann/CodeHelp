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
Private Const CH_LONGNAME As String = "TabIndex Editor"
Private Const CH_DESCRIPTION As String = "Reorder form's controls tab index"
Private Const CH_COPYRIGHT As String = "luthv@yahoo.com"

Private Const HH_DISPLAY_TOPIC = &H0
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Private m_lPtr As Long
Private m_Enabled As Boolean

Implements ICHPlugin

Private WithEvents Trap             As MouseTrap
Attribute Trap.VB_VarHelpID = -1
Private m_TrapPlaced                As Boolean
Private m_MenuItem                  As CommandBarControl
Private WithEvents codeBarEvents    As VBIDE.CommandBarEvents
Attribute codeBarEvents.VB_VarHelpID = -1

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

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean
    ICHPlugin_HaveExtendedHelp = True
End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean

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
    
    Dim dsr As ICHCore
    
    Set dsr = GetCHCore(m_lPtr)
    
    'Do your initializing stuff here
    'create menu item if this is the first onconnect event
    If m_MenuItem Is Nothing Then
        Set m_MenuItem = dsr.AddToCodeHelpMenu("Tab Order", LoadResPicture(101, vbResIcon))
        
        Set codeBarEvents = dsr.VBE.Events.CommandBarEvents(m_MenuItem)
    End If
    m_MenuItem.Enabled = m_Enabled
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    'Do your clean up here
    If Not m_MenuItem Is Nothing Then
        m_MenuItem.Enabled = False
        Set m_MenuItem = Nothing
    End If
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lReturn As Long, _
                                wParam As Long, _
                                lParam As Long)
    'do any keyboard related code here
    If m_TrapPlaced Then
        If wParam = vbKeyEscape Then
            ShowHideTrap
        End If
    End If
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hwnd As Long, _
                                ByVal uMsg As Long, _
                                wParam As Long, _
                                lParam As Long, _
                                bHandled As Boolean, _
                                lReturn As Long)
    If m_TrapPlaced Then
        Trap.TrapProc True, bHandled, lReturn, hwnd, uMsg, wParam, lParam
    End If

End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hwnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)
    'Hook msg goes here
        
    If m_TrapPlaced Then
        If hwnd = Trap.hWndTrap Then
            Select Case uMsg
                Case WM_NCHITTEST, WM_LBUTTONDOWN, WM_MOUSEMOVE, WM_RBUTTONDOWN, _
                    WM_SYSCOMMAND, WM_SETCURSOR
                    
                    placeSubclass = True
            End Select
        
        ElseIf hwnd = Trap.hWndForm Then
            If uMsg = WM_DESTROY Then
                Trap.TrapProc True, False, 0, hwnd, uMsg, wParam, lParam
            End If
        
        ElseIf hwnd = Trap.hWndFormDesigner Then
            If uMsg = WM_DRAWITEM Then
                placeSubclass = True
                before = False
            End If
        End If
    End If
End Sub

Private Sub ICHPlugin_ShowHelp()
    HtmlHelp 0, App.Path & "\" & ICHPlugin_Name & ".chm", HH_DISPLAY_TOPIC, 0
    'MsgBox "Use left mouse button to assign tab index, first control clicked will have "
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()

End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
    m_Enabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
    ICHPlugin_Enabled = m_Enabled
End Property
'End of ICHPlugin Implementation************************************************


Private Sub Hide()
    
    On Error Resume Next
    If m_TrapPlaced Then
        'Unload Trap
        Set Trap = Nothing
        m_TrapPlaced = False
    End If
   
End Sub

Private Sub Show()
  
    Dim vbComp As VBComponent
    Dim dsr As ICHCore
    
    Set dsr = GetCHCore(m_lPtr)
    
    Set vbComp = dsr.VBE.SelectedVBComponent
    If Not vbComp Is Nothing Then
        If vbComp.HasOpenDesigner Then
            Select Case vbComp.Type
                
                Case vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl, vbext_ct_DocObject
                    
                    Set Trap = New MouseTrap
                    
                    Set Trap.FormDesigner = vbComp.Designer
                    Set Trap.ParentWindow = vbComp.DesignerWindow
                    
                    m_TrapPlaced = True
                    Trap.ShowTabOrder
            End Select
            
        End If
    End If
End Sub

Private Sub codeBarEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ShowHideTrap
End Sub

Private Sub Trap_Closed()
    ShowHideTrap
End Sub

Private Sub ShowHideTrap()
    Dim cbButton As CommandBarButton
    Set cbButton = m_MenuItem
    If m_TrapPlaced = False Then
        Show
        If m_TrapPlaced Then cbButton.State = msoButtonDown
    Else
        Hide
        cbButton.State = msoButtonUp
    End If
End Sub
