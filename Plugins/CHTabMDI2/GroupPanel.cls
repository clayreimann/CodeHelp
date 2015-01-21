VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Plugin const
Private Const CH_LONGNAME As String = "Tabbed Tool Window"
Private Const CH_DESCRIPTION As String = "Group multiple tool windows into one tabbed tool window"
Private Const CH_COPYRIGHT As String = "luthv@yahoo.com"

Private Const TOOL_GUID As String = "{84AA41C0-3FE3-4ea0-8F6F-E70319277E9E}"
Private Const TAB_HEIGHT = 20

Private Const VBA_WIN As String = "VbaWindow"
Private Const IMMED As String = "Immediate"
Private Const WATCH As String = "Watches"
Private Const LOCALS As String = "Locals"
Private Const TASK_LIST As String = "Task List"
Private Const VBPALETTE As String = "VBFloatingPalette"

Private Type PanelStateStruct
    Caption As String
    ParenthWnd As Long
    VbahWnd As Long
    PreviouslyDocked As Long
End Type

Private m_lPtr          As Long
Private toolWindow      As Window
Private docWindow       As DockWindow
Private m_Enabled       As Boolean

Private WithEvents tabMgr       As TabManager
Attribute tabMgr.VB_VarHelpID = -1
Private m_hWnd                  As Long
Private m_DockedPanels()        As PanelStateStruct
Private m_PanelCount            As Long

Implements ICHPlugin

Private Sub Class_Terminate()
    Set toolWindow = Nothing
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

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
    m_Enabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
    ICHPlugin_Enabled = m_Enabled
End Property

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean
'Enable/disable help button in plugin manager
End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean
'Enable/disable properties button in plugin manager
End Property

Private Property Get ICHPlugin_LongName() As String
    ICHPlugin_LongName = CH_LONGNAME
End Property

Private Property Get ICHPlugin_Name() As String
    ICHPlugin_Name = "CHGroupPanel"
End Property

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, _
                                   custom() As Variant)
    
    'Sample use of the ICHCore pointer
    'It's advisable not to save the ICHCore object itself, always use the helper function to obtain the
    'ICHCore object from the pointer
        
    'Do your initializing stuff here
    
    Initialize GetCHCore(m_lPtr)
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    'Do your clean up here
    Dim tR As RECT
    Dim i As Long
    Dim dsr As ICHCore
    Dim vbaWin As Window
    
    m_Enabled = False
    If Not toolWindow Is Nothing Then
        If toolWindow.Visible Then
            
            toolWindow.Visible = False
            
            tabMgr.RemoveAll False
            
            Set dsr = GetCHCore(m_lPtr)
            
            If m_PanelCount > 0 Then
                For i = 0 To UBound(m_DockedPanels)
                    GetClientRect m_DockedPanels(i).ParenthWnd, tR
                    
                    SetParent m_DockedPanels(i).VbahWnd, m_DockedPanels(i).ParenthWnd
                    ShowWindow m_DockedPanels(i).ParenthWnd, SW_SHOW
                    
                    SetWindowPos m_DockedPanels(i).VbahWnd, 0, 0, 0, tR.Right, tR.Bottom, SWP_NOZORDER Or SWP_FRAMECHANGED
                    
                    If m_DockedPanels(i).PreviouslyDocked = 1 Then
                        For Each vbaWin In dsr.VBE.Windows
                            If vbaWin.Caption = m_DockedPanels(i).Caption Then
                                dsr.VBE.MainWindow.LinkedWindows.Add vbaWin
                                Exit For
                            End If
                        Next
                        
                    End If
                Next
            End If
            
            Erase m_DockedPanels
        End If
    End If
    m_PanelCount = 0
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lReturn As Long, _
                                wParam As Long, _
                                lParam As Long)
    'do any keyboard related code here
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, _
                                ByVal uMsg As Long, _
                                wParam As Long, _
                                lParam As Long, _
                                bHandled As Boolean, _
                                lReturn As Long)
    'subclassed message goes here
    On Error GoTo ERR_HANDLER
    
    Select Case uMsg
        Case WM_NCCALCSIZE
            OnNCCalcSize lParam
        
        Case WM_NCHITTEST
            lReturn = OnNCHittest(lParam)
            
        Case WM_ACTIVATE
            tabMgr.RequestRedraw
            
        Case WM_WINDOWPOSCHANGING
            'don't let VB reshow the old floating palette
'            Dim tWP As WINDOWPOS, bProcess As Boolean, i As Long
'            If m_PanelCount > 0 Then
'                For i = 0 To m_PanelCount - 1
'                    If m_DockedPanels(i).ParenthWnd = hWnd Then
'                        bProcess = True
'                        Exit For
'                    End If
'                Next
'
'                If bProcess Then
'                    CopyMemory tWP, ByVal lParam, Len(tWP)
'
'                    If (tWP.flags And SWP_SHOWWINDOW) = SWP_SHOWWINDOW Then
'                        tWP.flags = tWP.flags And Not SWP_SHOWWINDOW
'                        CopyMemory ByVal lParam, tWP, Len(tWP)
'                    End If
'                End If
'
'            End If
            
            OnWindowPosChanging hWnd, lParam
    End Select
    Exit Sub
ERR_HANDLER:
    tabMgr.LockUpdateRequest False, True
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)
    'Hook msg goes here
    Dim bProcess As Boolean, i As Long
    Dim item As TabItem
    
    Select Case uMsg
        Case WM_NCCALCSIZE, WM_NCHITTEST
            If hWnd <> m_hWnd Then Exit Sub
            'need to subclass to modify the return value
            placeSubclass = True
            before = False
            
        Case WM_NCPAINT
            If hWnd <> m_hWnd Then Exit Sub
            tabMgr.RequestRedraw
        
        Case WM_WINDOWPOSCHANGED
            OnWindowPosChanged hWnd, lParam
            
        Case WM_WINDOWPOSCHANGING

'            If m_PanelCount > 0 Then
'                For i = 0 To m_PanelCount - 1
'                    If m_DockedPanels(i).ParenthWnd = hWnd Then
'                        bProcess = True
'                        Exit For
'                    End If
'                Next
'
'                If bProcess Then
'                    placeSubclass = True
'                    before = True
'                End If
'            End If
            placeSubclass = True
            before = True
                    
        Case (WM_USER + 3139) 'Catch window activation from menu
            If wParam = 1 And m_PanelCount > 0 Then
                For i = 0 To m_PanelCount - 1
                    If m_DockedPanels(i).VbahWnd = hWnd Then
                        bProcess = True
                        Exit For
                    End If
                Next
                If bProcess Then
                    For Each item In tabMgr.Items
                        If CLng(item.Tag) = hWnd Then
                            item.Selected = True
                            tabMgr_ItemClick item
                            Exit For
                        End If
                    Next
                    'in case the CodeHelp panel is closed
                    toolWindow.Visible = True
                    SetFocusApi hWnd
                End If
            End If
        
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMOUSEMOVE, _
            WM_NCLBUTTONUP, WM_NCRBUTTONUP, WM_NCLBUTTONDBLCLK
            
            Dim tPt As POINTAPI
            Dim Button As MouseButtonConstants
            
            If hWnd <> m_hWnd Then Exit Sub
            tPt = GetMousePos(lParam)
            
            Select Case uMsg

                Case WM_NCLBUTTONDOWN
                    tabMgr.OnLMouseDown tPt.x, tPt.y
                
                Case WM_NCMOUSEMOVE
                    If GetAsyncKeyState(VK_LBUTTON) Then
                        Button = vbLeftButton
                    End If

                    tabMgr.OnMouseMove Button, tPt.x, tPt.y

                Case WM_NCLBUTTONUP
                    tabMgr.OnMouseUp vbLeftButton, tPt.x, tPt.y
                    
                Case WM_NCRBUTTONUP
                    tabMgr.OnMouseUp vbRightButton, tPt.x, tPt.y
                    
                Case WM_NCLBUTTONDBLCLK 'treat double click as one click
                    tabMgr.OnLMouseDown tPt.x, tPt.y
                    tabMgr.OnMouseUp vbLeftButton, tPt.x, tPt.y
                    
            End Select
    End Select
End Sub

Private Sub ICHPlugin_ShowHelp()
'Show help instruction when user click on more info button in plugin manager
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
'show property dialog when user click on properties button in plugin manager
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property


Private Sub Initialize(ByVal dsr As ICHCore)
    
    Set tabMgr = New TabManager
    
    Set dsr = GetCHCore(m_lPtr)
        
    DockVBAWindow dsr, WATCH
    DockVBAWindow dsr, LOCALS
    DockVBAWindow dsr, IMMED
    
    tabMgr.InsertItem TASK_LIST
    
    SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or _
        SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End Sub

Private Sub Terminate()
    Set tabMgr = Nothing
End Sub

Private Sub OnNCCalcSize(ByRef lParam As Long)
    Dim tNC As NCCALCSIZE_PARAMS
    
    'm_bShowTabStrip = tabMgr.IsAnyVisibleItem
    
    CopyMemory tNC, ByVal lParam, Len(tNC)
'    InflateRect tNC.rgrc(0), 1, 1
    
    'If m_bShowTabStrip Then
        tNC.rgrc(0).Top = tNC.rgrc(0).Top + TAB_HEIGHT
        LSet tNC.rgrc(1) = tNC.rgrc(0)
        MoveTabStrip
    'End If
    
    CopyMemory ByVal lParam, tNC, Len(tNC)
End Sub

Private Sub MoveTabStrip()
    Dim tR As RECT
    Dim lWidth As Long
    Dim i As Long
    
    GetWindowRect m_hWnd, tR
    lWidth = tR.Right - tR.Left

    tabMgr.Move 0&, 0&, lWidth, TAB_HEIGHT
    
    If m_PanelCount > 0 Then
        For i = 0 To m_PanelCount - 1
            
            MoveWindow m_DockedPanels(i).VbahWnd, 1, 2, lWidth, tR.Bottom - tR.Top - TAB_HEIGHT, 1
            
        Next
    End If
        
End Sub

Private Function OnNCHittest(ByVal lParam As Long) As Long
    Dim tPt As POINTAPI
    tPt.x = LoWord(lParam): tPt.y = HiWord(lParam)
    MapWindowPoints 0, m_hWnd, tPt, 1
    
    If tPt.y < 0 Then
        OnNCHittest = HTBORDER
    Else
        tabMgr.OnMouseLeave
    End If

End Function

Private Sub tabMgr_DrawFinished(ByVal hdc As Long)
    ReleaseDC m_hWnd, hdc
End Sub

Private Sub tabMgr_ItemChanging(ByVal toItem As TabItem, Cancel As Boolean)
    Dim item As TabItem
    Dim childHwnd As Long
    
    Set item = tabMgr.SelectedItem
    If Not item Is Nothing Then
        If item.Caption = TASK_LIST Then
            childHwnd = docWindow.RealhWnd
        Else
            childHwnd = item.Tag
        End If
        ShowWindow childHwnd, SW_HIDE
    End If
End Sub

Private Sub tabMgr_ItemClick(ByVal item As TabItem)
    If item.Caption = TASK_LIST Then
        ShowWindow docWindow.RealhWnd, SW_SHOW
    Else
        ShowWindow CLng(item.Tag), SW_SHOW
    End If
End Sub

Private Sub tabMgr_RequestRedraw(hdc As Long)
    hdc = GetWindowDC(m_hWnd)
    SetBkMode hdc, TRANSPARENT
End Sub

Private Sub OnWindowPosChanged(ByVal hWnd As Long, ByVal lParam As Long)
    Dim tWP As WINDOWPOS
    Dim winCaption As String
    Dim bProcess As Boolean
        
    If hWnd = m_hWnd Then
        CopyMemory tWP, ByVal lParam, Len(tWP)
        If (tWP.flags And SWP_NOSIZE) = 0 Then
            MoveTabStrip
        End If
    End If
End Sub

'Private Sub OnWindowPosChanged(ByVal hWnd As Long, ByVal lParam As Long)
'    Dim tWP As WINDOWPOS
'    Dim winCaption As String
'    Dim bProcess As Boolean
'
'    If hWnd = m_hWnd Then
'        bProcess = True
'    Else
'        If GetWinText(hWnd, True) = VBA_WIN Then
'            winCaption = GetWinText(hWnd)
'
'            Select Case winCaption
'
'                Case IMMED, WATCH, LOCALS
'
'                    bProcess = True
'
'            End Select
'        End If
'    End If
'
'    If bProcess Then
'        CopyMemory tWP, ByVal lParam, Len(tWP)
'        If (tWP.flags And SWP_NOSIZE) = 0 Then
'            MoveTabStrip
'        ElseIf (tWP.flags And SWP_SHOWWINDOW) = SWP_SHOWWINDOW Then
'
'            If m_hWnd <> hWnd Then
'                DockVBAWindow GetCHCore(m_lPtr), winCaption
'
'                SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or _
'                    SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
'            End If
'        End If
'    End If
'End Sub

Private Sub DockVBAWindow(ByVal dsr As ICHCore, ByVal winName As String) 'As Long
    Dim vbaWin As Window
    Dim item As TabItem
    Dim panelState As PanelStateStruct
    Dim bNeedReDocked As Boolean
    
    Set vbaWin = dsr.VBE.Windows(winName)
    
    panelState.VbahWnd = A_FindWindowEx(dsr.VBE.MainWindow.hWnd, 0, VBA_WIN, winName)
    
    If panelState.VbahWnd <> 0 Then
        For Each item In tabMgr.Items
            If CLng(item.Tag) = panelState.VbahWnd Then
                If IsChild(m_hWnd, panelState.VbahWnd) Then
                    Exit Sub
                Else
                'vb has redock our window to it's original parent
                    bNeedReDocked = True
                    Exit For
                End If
            End If
        Next
        
        'window is docked, undock it
        panelState.PreviouslyDocked = 1
        'temporary disable add in, cause undocking caused nested call to WM_WINDOWPOSCHANGED
        m_Enabled = False
        dsr.VBE.MainWindow.LinkedWindows.Remove vbaWin
        'reenable
        m_Enabled = True
    End If
    
    'now find the floating palette containing our window
    panelState.ParenthWnd = A_FindWindowEx(0, panelState.ParenthWnd, VBPALETTE, vbNullString)
    'see if it belong to current IDE instance
    Do While (panelState.ParenthWnd <> 0)
        If GetWindowThreadProcessId(panelState.ParenthWnd, 0) = App.ThreadID Then
        'get the real immediate window from the palette
            panelState.VbahWnd = A_FindWindowEx(panelState.ParenthWnd, 0, VBA_WIN, winName)
            If panelState.VbahWnd <> 0 Then Exit Do
        End If
        panelState.ParenthWnd = A_FindWindowEx(0, panelState.ParenthWnd, VBPALETTE, vbNullString)
    Loop
        
    If panelState.VbahWnd <> 0 Then
        bNeedReDocked = False
        For Each item In tabMgr.Items
            If CLng(item.Tag) = panelState.VbahWnd Then
                If IsChild(m_hWnd, panelState.VbahWnd) Then
                    Exit Sub
                Else
                'vb has redock our window to it's original parent
                    
                    bNeedReDocked = True
                    Exit For
                End If
            End If
        Next
        
        If toolWindow Is Nothing Then
            Set toolWindow = dsr.VBE.Windows.CreateToolWindow(dsr.AddInInst, "CHTabMDI.DockWindow", "CodeHelp Panel", TOOL_GUID, docWindow)
        End If
        toolWindow.Visible = True
        m_hWnd = docWindow.hWnd

        SetParent panelState.VbahWnd, m_hWnd
        panelState.Caption = winName
        
        ShowWindow panelState.ParenthWnd, SW_HIDE

        If Not bNeedReDocked Then
            Set item = tabMgr.InsertItem(winName, , , vbaWin.Visible)
            item.Tag = panelState.VbahWnd
            
            If item.Visible Then item.Selected = True
            
            ReDim Preserve m_DockedPanels(m_PanelCount)
            m_DockedPanels(m_PanelCount) = panelState
                    
            m_PanelCount = m_PanelCount + 1&
        Else
            
            item.Visible = True
            item.Selected = True
            ShowWindow panelState.VbahWnd, SW_SHOW
            SetFocusApi panelState.VbahWnd
        End If
    End If
    
End Sub

Private Function GetMousePos(ByVal lParam As Long) As POINTAPI
    GetMousePos.x = LoWord(lParam)
    GetMousePos.y = HiWord(lParam)
    
    'MapWindowPoints HWND_DESKTOP, m_hMDIClient, GetMousePos, 1
    MapWindowPoints 0, m_hWnd, GetMousePos, 1
    
    GetMousePos.x = GetMousePos.x + 1&
    GetMousePos.y = GetMousePos.y + tabMgr.Height
End Function

Private Sub OnWindowPosChanging(ByVal hWnd As Long, ByVal lParam As Long)
    'don't let VB reshow the old floating palette
    Dim tWP As WINDOWPOS, i As Long
    Dim winCaption As String
    
    If m_PanelCount > 0 Then
        For i = 0 To m_PanelCount - 1
            If m_DockedPanels(i).ParenthWnd = hWnd Then
                CopyMemory tWP, ByVal lParam, Len(tWP)
    
                If (tWP.flags And SWP_SHOWWINDOW) = SWP_SHOWWINDOW Then
                    tWP.flags = tWP.flags And Not SWP_SHOWWINDOW
                    CopyMemory ByVal lParam, tWP, Len(tWP)
                End If

                Exit Sub
            End If
        Next
    End If

    If GetWinText(hWnd, True) = VBA_WIN Then
        winCaption = GetWinText(hWnd)

        Select Case winCaption

            Case IMMED, WATCH, LOCALS
                CopyMemory tWP, ByVal lParam, Len(tWP)
                If (tWP.flags And SWP_SHOWWINDOW) = SWP_SHOWWINDOW Then
        
                    If m_hWnd <> hWnd Then
                        DockVBAWindow GetCHCore(m_lPtr), winCaption
        
                        SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or _
                            SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
                    End If
                End If

        End Select
    End If

End Sub
