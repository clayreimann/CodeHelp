VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Plugin const
Private Const CH_LONGNAME As String = "CodeHelp Full Screen Code Editor"
Private Const CH_DESCRIPTION As String = "Provide Full Screen view capability for VB IDE"
Private Const CH_COPYRIGHT As String = "luthv@yahoo.com"

Private m_hiddenWindows As Collection
Private m_cmdBar As CommandBar
Private m_MenuItem As CommandBarControl

Private WithEvents cmdBarEvents As VBIDE.CommandBarEvents
Attribute cmdBarEvents.VB_VarHelpID = -1
Private WithEvents codeBarEvents As VBIDE.CommandBarEvents
Attribute codeBarEvents.VB_VarHelpID = -1

Private m_hWnd As Long
Private m_bInFullScreen As Boolean
Private m_bIsMaximized As Boolean
Private m_lPtr As Long
Private m_Enabled As Boolean

Implements ICHPlugin

Private Sub Invoke()
    On Error Resume Next
    Dim btnMenu As CommandBarButton
    
    Set btnMenu = m_MenuItem
    If m_bInFullScreen = False Then
        EnterFullScreen
        btnMenu.State = msoButtonDown
    Else
        ExitFullScreen
        btnMenu.State = msoButtonUp
    End If

    m_bInFullScreen = Not m_bInFullScreen
    
    'Redraw Frame
    SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Sub

Private Sub EnterFullScreen()
    
    If IsWindow(m_hWnd) = 0 Then Exit Sub
    
    Dim win As Window
    Dim vbeInst As VBE
    Dim dsr As ICHCore
    
    Set m_hiddenWindows = New Collection
    
    Set dsr = GetCHCore(m_lPtr)
    Set vbeInst = dsr.VBE
    
    For Each win In vbeInst.Windows

        Select Case win.Type

            Case vbext_wt_CodeWindow, vbext_wt_Designer
        
            Case Else

                If win.Visible Then
                    win.Visible = False
                    
                    m_hiddenWindows.Add win
                End If

        End Select

    Next

    Dim cmb As CommandBar

    For Each cmb In vbeInst.CommandBars

        If cmb.Name <> "Menu Bar" Then
            If cmb.Visible Then
                cmb.Visible = False
                m_hiddenWindows.Add cmb
            End If
        End If

    Next

    If m_cmdBar Is Nothing Then
        Dim picIcon As StdPicture
        Dim picBmp As StdPicture
        
        Dim cmdCtl As CommandBarButton
    
        Set picIcon = LoadResPicture(101, vbResIcon)
    
        Set m_cmdBar = vbeInst.CommandBars.Add("FullScreen")
        m_cmdBar.Position = msoBarFloating
        Set cmdCtl = m_cmdBar.Controls.Add(msoControlButton)
        cmdCtl.Caption = "Exit Fullscreen"
        cmdCtl.Style = msoButtonIconAndCaption
        
        On Error Resume Next 'Clipboard could be lock by other application
        Clipboard.Clear
        cmdCtl.CopyFace
        
        Set picBmp = Clipboard.GetData
        
        CopyIconToClipBoardAsBmp picIcon, picBmp
        cmdCtl.PasteFace
        Clipboard.Clear
        
        Set picIcon = Nothing
        Set picBmp = Nothing
        
        m_cmdBar.Protection = msoBarNoChangeDock
        m_cmdBar.Top = 0
        m_cmdBar.Left = (Screen.Width \ Screen.TwipsPerPixelX) - m_cmdBar.Width
        Set cmdBarEvents = vbeInst.Events.CommandBarEvents(m_cmdBar.Controls(1))
    
    End If

    m_cmdBar.Visible = True
    
    'save state
    m_bIsMaximized = IsZoomed(m_hWnd)

    If Not m_bIsMaximized Then ShowWindow m_hWnd, SW_MAXIMIZE
    A_SetWindowLong m_hWnd, GWL_STYLE, A_GetWindowLong(m_hWnd, GWL_STYLE) And Not WS_CAPTION
End Sub

Private Sub ExitFullScreen()
    
    Dim item As Variant

    For Each item In m_hiddenWindows
        item.Visible = True
    Next
    
    Set cmdBarEvents = Nothing
    m_cmdBar.Delete
    Set m_cmdBar = Nothing

    Set m_hiddenWindows = Nothing
    
    If IsWindow(m_hWnd) Then
        If m_bIsMaximized = False Then
            ShowWindow m_hWnd, SW_RESTORE
        End If

        A_SetWindowLong m_hWnd, GWL_STYLE, A_GetWindowLong(m_hWnd, GWL_STYLE) Or WS_CAPTION
    End If

End Sub

Private Sub Class_Terminate()

    If m_bInFullScreen Then
        Invoke
    End If
End Sub

Private Sub cmdBarEvents_Click(ByVal CommandBarControl As Object, _
                               handled As Boolean, _
                               CancelDefault As Boolean)
    Invoke
End Sub

Private Sub codeBarEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Invoke
End Sub

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
    Dim dsr As ICHCore
    Dim cmdBtn As CommandBarButton
    
    Set dsr = GetCHCore(m_lPtr)
    
    If Not dsr Is Nothing Then
        m_hWnd = dsr.VBE.MainWindow.hWnd
        
        'now it's possible for user to disable a plugin at runtime
        'so we have to check whether the plugin is connected for the first time
        If m_MenuItem Is Nothing Then
            Set m_MenuItem = dsr.AddToCodeHelpMenu("Full Screen", LoadResPicture(101, vbResIcon))
            Set cmdBtn = m_MenuItem
            cmdBtn.ShortcutText = "Shift+Enter"
            Set codeBarEvents = dsr.VBE.Events.CommandBarEvents(m_MenuItem)
        End If
        
        m_MenuItem.Enabled = m_Enabled
    
    End If
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    
    If Not m_MenuItem Is Nothing Then
        'can not delete menu item, bug in office command bar?
        'so we just disable it
        m_MenuItem.Enabled = False
    End If
    
    Class_Terminate
    
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lreturn As Long, _
                                wParam As Long, _
                                lParam As Long)
        
    Select Case wParam

        Case vbKeyReturn
            
            If (lParam And &HC0000000) = 0 Then
                If (GetKeyState(vbKeyShift) And &H8000) <> 0 Then
                    lreturn = 1
                    bHandled = True

                    Invoke
                    
                End If
            End If

    End Select
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bHandled As Boolean, lreturn As Long)

End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)

End Sub

Private Sub ICHPlugin_ShowHelp()

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
