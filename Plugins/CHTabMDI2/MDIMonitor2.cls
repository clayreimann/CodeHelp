VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDIMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const CH_LONGNAME As String = "CodeHelp Tabbed MDI Environtment"
Private Const CH_DESCRIPTION As String = "Provide tabbed workspace in VB IDE"
Private Const CH_COPYRIGHT As String = "luthv@yahoo.com"

Private Const TAB_HEIGHT = 22&

Private Const ID_CLOSE As Long = 8
Private Const ID_CLOSE_ALL As Long = 9
Private Const ID_CLOSE_BUT_ACTIVE As Long = 10

Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByRef lprc As Any) As Long

Private Const HH_DISPLAY_TOPIC = &H0
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Implements ICHPlugin

Private m_CHCorePtr As Long
Private m_hMDIClient As Long
Private m_hMenu As Long
Private m_hSubMenu As Long
Private m_hWndMain As Long
Private m_Enabled As Boolean

Private m_ShowTabStrip As Boolean
Private m_ShowCloseButton As CloseButtonPosition

Private WithEvents m_tabManager As TabManager
Attribute m_tabManager.VB_VarHelpID = -1

Dim m_FromTabBar As Boolean
Dim m_CtrlPressed As Boolean

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

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
    m_Enabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
    ICHPlugin_Enabled = m_Enabled
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

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, custom() As Variant)
    Dim CHCore As ICHCore
    
    Set CHCore = GetCHCore(m_CHCorePtr)
    m_hWndMain = CHCore.VBE.MainWindow.hWnd
    Set CHCore = Nothing
    
    Call StartMonitor(ConnectMode = ext_cm_AfterStartup)
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, custom() As Variant)
    Call EndMonitor
End Sub

' This is where we ctrl-switch tabs
Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, lReturn As Long, wParam As Long, lParam As Long)
    ' Call into VBWinApi to get the control key state
    CtrlDown = GetAsyncKeyState(vbKeyControl)
    
    If m_CtrlPressed Then
        Select Case wParam
            Case vbKey0 To vbKey9
                Call m_tabManager.ShortcutActivate(wParam)
        End Select
    End If
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                                    bHandled As Boolean, placeSubclass As Boolean, before As Boolean)
    On Error GoTo ERR_HANDLER
    
    Dim tPt As POINTAPI
    Dim Button As MouseButtonConstants
    
    Select Case uMsg
            
        Case WM_SHOWWINDOW
            Call OnShowWindow(hWnd, wParam)
            
        Case WM_SETTEXT
            Call OnTextChanged(hWnd, lParam)
            
        Case WM_PARENTNOTIFY
            Call OnParentNotify(hWnd, wParam, lParam)
            
        Case WM_NCPAINT
            If hWnd <> m_hMDIClient Then Exit Sub
            Call m_tabManager.RequestRedraw
            
        Case WM_NCCALCSIZE, WM_NCHITTEST
            If hWnd <> m_hMDIClient Then Exit Sub
            'need to subclass to modify the return value
            placeSubclass = True
            before = False
        
        Case WM_SIZE
            If hWnd <> m_hMDIClient Then Exit Sub
            Call MoveTabStrip
        
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMOUSEMOVE, _
            WM_NCLBUTTONUP, WM_NCRBUTTONUP, WM_NCLBUTTONDBLCLK
            
            If hWnd <> m_hMDIClient Then Exit Sub
            tPt = GetMousePos(lParam)
            
            Select Case uMsg

                Case WM_NCLBUTTONDOWN
                    Call m_tabManager.OnLMouseDown(tPt.x, tPt.y)
                
                Case WM_NCMOUSEMOVE
                    If GetAsyncKeyState(VK_LBUTTON) Then
                        Button = vbLeftButton
                    End If

                    Call m_tabManager.OnMouseMove(Button, tPt.x, tPt.y)

                Case WM_NCLBUTTONUP
                    Call m_tabManager.OnMouseUp(vbLeftButton, tPt.x, tPt.y)
                    
                Case WM_NCRBUTTONUP
                    Call m_tabManager.OnMouseUp(vbRightButton, tPt.x, tPt.y)
                    
                Case WM_NCLBUTTONDBLCLK 'treat double click as one click
                    Call m_tabManager.OnLMouseDown(tPt.x, tPt.y)
                    Call m_tabManager.OnMouseUp(vbLeftButton, tPt.x, tPt.y)
                    
            End Select
            
        Case WM_MOUSEMOVE
            Call m_tabManager.OnMouseLeave
        
        Case WM_CHILDACTIVATE
            If GetParent(hWnd) = m_hMDIClient Then
                FindItem(hWnd).Selected = True
            End If
            
        Case WM_ACTIVATE
            If hWnd = m_hWndMain Then
                If LowWord(wParam) = WA_ACTIVE Then
                    placeSubclass = True
                End If
            End If
    End Select
    
    Exit Sub
ERR_HANDLER:
    Call m_tabManager.LockUpdateRequest(False, True)
    'LogToNotePad Err.Description & " " & Erl
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, _
                                bHandled As Boolean, lReturn As Long)

    On Error GoTo ERR_HANDLER
    
    Select Case uMsg
        Case WM_NCCALCSIZE
            Call OnNCCalcSize(lParam)
        
        Case WM_NCHITTEST
            lReturn = OnNCHittest(lParam)
            
        Case WM_ACTIVATE
            Call m_tabManager.RequestRedraw
            
    End Select
    
    Exit Sub
ERR_HANDLER:
    Call m_tabManager.LockUpdateRequest(False, True)
    'LogToNotePad Err.Description
End Sub

Private Sub ICHPlugin_ShowHelp()
    HtmlHelp 0, App.Path & "\" & ICHPlugin_Name & ".chm", HH_DISPLAY_TOPIC, 0
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()

End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub StartMonitor(ByVal AfterStartUp As Boolean)
    Call EndMonitor
        
    m_hMDIClient = A_FindWindowEx(m_hWndMain, 0, "MDIClient", vbNullString)

    If m_hMDIClient = 0 Then Exit Sub
  
    Set m_tabManager = New TabManager
    
    If AfterStartUp Then
        Call RegisterExistingWindows
    End If
    
    Call SetWindowPos(m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
    'update client window in case child is not in maximized state
    Call RedrawWindow(m_hMDIClient, ByVal 0, ByVal 0, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ERASE)
    
    Call CreatePopup
    Call LoadSetting
End Sub

Private Sub EndMonitor()

    If m_hMDIClient = 0 Then Exit Sub
    If m_hWndMain = 0 Then Exit Sub
    
    On Error GoTo ERR_HANDLER
  
    If m_hMenu <> 0 Then
        Call DestroyMenu(m_hMenu)
        m_hMenu = 0
    End If
    
    Call SaveSetting("CodeHelp", ICHPlugin_Name, "CloseButton", m_ShowCloseButton)
    
    Call SetWindowPos(m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
    
ERR_HANDLER:
    Set m_tabManager = Nothing
    m_hWndMain = 0
    m_hMDIClient = 0
End Sub

Private Sub RegisterExistingWindows()
    Dim childWindow     As Long
    Dim oTabItem        As TabItem
    
    childWindow = GetWindow(m_hMDIClient, GW_CHILD)
    'disable update to speed up loading large project
    Call m_tabManager.LockUpdateRequest(True)

    Do While childWindow <> 0
        Call AddNewWindow(childWindow)
        childWindow = GetWindow(childWindow, GW_HWNDNEXT)
    Loop

    If m_tabManager.IsAnyVisibleItem Then
        childWindow = A_SendMessage(m_hMDIClient, WM_MDIGETACTIVE, 0, ByVal 0)
        
        ' I want to find a way to maximize the active window here :(
        
        If childWindow <> 0 Then
            Set oTabItem = FindItem(childWindow)

            If Not oTabItem Is Nothing Then oTabItem.Selected = True
        End If
    End If

    Call m_tabManager.LockUpdateRequest(False)
End Sub

Private Sub AddNewWindow(ByVal hWnd As Long)
    Dim oTabItem As TabItem
    
    Set oTabItem = m_tabManager.InsertItem(Caption:=GetCaption(GetWinText(hWnd, False)), Visible:=IsWindowVisible(hWnd))
    oTabItem.ChildhWnd = hWnd
End Sub

Private Function GetCaption(ByVal Text As String) As String

    If Len(Text) > 0 Then
        Text = Replace(Text, "(Form)", "(Design)")
        Text = Replace(Text, "(UserControl)", "(Design)")
        Text = Replace(Text, "(DataEnvironment)", "(Design)")
        If InStr(1, Text, "-") Then
            'Remove the project name prefix
            'Thanks to Zirro Tolerance
            Text = Trim$(Replace(Split(Text, "-")(1), "(Code)", ""))
        Else
            Text = Trim$(Replace(Text, "(Code)", ""))
        End If

        GetCaption = Text
    End If

End Function

Private Function FindItem(ByVal hWnd As Long) As TabItem
    Dim oTabItem As TabItem

    For Each oTabItem In m_tabManager.Items

        If CLng(oTabItem.ChildhWnd) = hWnd Then
            Set FindItem = oTabItem
            Exit For
        End If

    Next

End Function

Private Sub m_tabManager_DrawFinished(ByVal hdc As Long)
    Call ReleaseDC(m_hMDIClient, hdc)
End Sub

Private Sub m_tabManager_ItemClick(ByVal oTabItem As TabItem)

    If oTabItem.Index > 0 Then
        Call A_SendMessage(m_hMDIClient, WM_MDIACTIVATE, ByVal CLng(oTabItem.ChildhWnd), ByVal 0)
    End If

End Sub

Private Sub m_tabManager_ItemClosing(ByVal oTabItem As TabItem, ByVal FromButton As Boolean, Cancel As Boolean)

    If FromButton Then m_FromTabBar = True
    If m_FromTabBar Then
        Call A_SendMessage(CLng(oTabItem.ChildhWnd), WM_SYSCOMMAND, SC_CLOSE, ByVal 0)
    End If

    If FromButton Then m_FromTabBar = False
End Sub

Private Sub m_tabManager_MouseUp(ByVal Button As MouseButtonConstants, ByVal oTabItem As TabItem)
    Dim menuItemID As Long
    Dim itemhWnd As Long
    
    If Button = vbRightButton Then
        menuItemID = ShowPopup(oTabItem)

        If menuItemID = 0 Then Exit Sub
        
        Select Case menuItemID

            Case ID_CLOSE
                itemhWnd = CLng(oTabItem.ChildhWnd)
                Call A_SendMessage(itemhWnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0)
                
            Case ID_CLOSE_ALL, ID_CLOSE_BUT_ACTIVE
                m_FromTabBar = True

                If menuItemID = ID_CLOSE_ALL Then
                    m_tabManager.RemoveAll True
                Else
                    m_tabManager.RemoveAllButActive True
                End If

                m_FromTabBar = False
                Call SetWindowPos(m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
                            
            Case CloseButtonHidden, CloseButtonRightMost, CloseButtonOnActive
                m_ShowCloseButton = menuItemID
                m_tabManager.PaintManager.ShowCloseButton = menuItemID
                
                Call UpdateMenu(menuItemID)
                
        End Select
    
    End If

End Sub

Private Sub m_tabManager_RequestRedraw(hdc As Long)
    ' we only want to show the tab strip when a window is open
    If m_ShowTabStrip Then
        hdc = GetWindowDC(m_hMDIClient)
        Call SetBkMode(hdc, TRANSPARENT)
    End If

End Sub

Private Sub MoveTabStrip()
    Dim curRect As RECT
    Dim lWidth As Long
  
    Call GetWindowRect(m_hMDIClient, curRect)
    lWidth = curRect.Right - curRect.Left

    Call m_tabManager.Move(1&, 1&, lWidth - 1&, TAB_HEIGHT)
End Sub

Private Sub CreatePopup()
    
    m_hMenu = CreatePopupMenu()
    m_hSubMenu = CreatePopupMenu()
    Call A_AppendMenuStr(m_hMenu, MF_STRING, ID_CLOSE, "Close")
    Call A_AppendMenuStr(m_hMenu, MF_STRING, ID_CLOSE_ALL, "Close All")
    Call A_AppendMenuStr(m_hMenu, MF_STRING, ID_CLOSE_BUT_ACTIVE, "Close All But Active")
    Call A_AppendMenu(m_hMenu, MF_SEPARATOR, 0, ByVal 0)
    Call A_AppendMenuStr(m_hMenu, MF_POPUP, m_hSubMenu, "Close Button")
    
    Call A_AppendMenuStr(m_hSubMenu, MF_STRING, CloseButtonHidden, "Hidden")
    Call A_AppendMenuStr(m_hSubMenu, MF_STRING, CloseButtonRightMost, "Rightmost")
    Call A_AppendMenuStr(m_hSubMenu, MF_STRING, CloseButtonOnActive, "On Active Tab")
End Sub

Private Function ShowPopup(ByVal oTabItem As TabItem) As Long
    Dim tPt As POINTAPI
  
    If m_tabManager.Items.Count > 1 Then
        Call EnableMenuItem(m_hMenu, ID_CLOSE_ALL, MF_ENABLED)
        Call EnableMenuItem(m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_ENABLED)
    Else
        Call EnableMenuItem(m_hMenu, ID_CLOSE_ALL, MF_GRAYED)
        Call EnableMenuItem(m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_GRAYED)
    End If
    
    If Not oTabItem Is Nothing Then
        Call EnableMenuItem(m_hMenu, ID_CLOSE, MF_ENABLED)
    Else
        Call EnableMenuItem(m_hMenu, ID_CLOSE, MF_GRAYED)
    End If
    
    Call GetCursorPos(tPt)
    ShowPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD Or TPM_NONOTIFY, tPt.x, tPt.y, ByVal 0, m_hWndMain, ByVal 0)
    
End Function

Private Sub LoadSetting()
    On Error Resume Next
    m_ShowCloseButton = CLng(GetSetting("CodeHelp", ICHPlugin_Name, "CloseButton", CloseButtonOnActive))
    m_tabManager.PaintManager.ShowCloseButton = m_ShowCloseButton
    Call UpdateMenu(m_tabManager.PaintManager.ShowCloseButton)
End Sub

Private Sub UpdateMenu(ByVal menuID As Long)
    Dim idx As Long
    
    For idx = 0 To 2
        If menuID = 2 ^ idx Then
            Call CheckMenuItem(m_hSubMenu, menuID, MF_BYCOMMAND Or MF_CHECKED)
        Else
            Call CheckMenuItem(m_hSubMenu, menuID, MF_BYCOMMAND Or MF_UNCHECKED)
        End If
    Next

End Sub

'Message processing
Private Sub OnShowWindow(ByVal hWnd As Long, ByVal wParam As Long)
    If GetParent(hWnd) <> m_hMDIClient Then Exit Sub
    
    Dim oTabItem As TabItem
    
    Set oTabItem = FindItem(hWnd)
    
    If oTabItem Is Nothing Then Exit Sub
    
    oTabItem.Visible = (wParam = 1)
    
    Call RequestFrameChanges
        
    If wParam Then
        If IsZoomed(hWnd) = 0 Then
            Call RedrawWindow(m_hMDIClient, ByVal 0, ByVal 0, RDW_ERASENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN)
        End If
    End If
End Sub

Private Sub OnTextChanged(ByVal hWnd As Long, ByVal lParam As Long)
    Dim idx As Long
    Dim sText As String
    
    If GetParent(hWnd) <> m_hMDIClient Then Exit Sub
    
    sText = String$(128, vbNullChar)
    
    ' copy the stringPtr contents into a string of our own
    Call A_lstrcpyPtrStr(sText, ByVal lParam)
    
    ' figure out where it ends
    idx = InStr(1, sText, vbNullChar) - 1&
    If idx > 0 Then
        sText = Left$(sText, idx)
    End If
    
    FindItem(hWnd).Caption = GetCaption(sText)
End Sub

Private Sub OnParentNotify(ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long)
    If hWnd <> m_hMDIClient Then Exit Sub
    
    Select Case LowWord(wParam)
        Case WM_CREATE
            Call AddNewWindow(lParam)
            
        Case WM_DESTROY
            If m_FromTabBar Then Exit Sub
            Call m_tabManager.RemoveItem(FindItem(lParam), False)

    End Select
End Sub

Private Sub OnNCCalcSize(ByRef lParam As Long)
    Dim tNC As NCCALCSIZE_PARAMS
    
    m_ShowTabStrip = m_tabManager.IsAnyVisibleItem
    
    Call CopyMemory(tNC, ByVal lParam, Len(tNC))
    Call InflateRect(tNC.rgrc(0), 1, 1)
    
    If m_ShowTabStrip Then
        tNC.rgrc(0).Top = tNC.rgrc(0).Top + TAB_HEIGHT '- 1&
        LSet tNC.rgrc(1) = tNC.rgrc(0)
        Call MoveTabStrip
    End If
    
    CopyMemory ByVal lParam, tNC, Len(tNC)
    
End Sub

Private Function OnNCHittest(ByVal lParam As Long) As Long
    Dim tPt As POINTAPI
    tPt.x = LowWord(lParam)
    tPt.y = HiWord(lParam)
    
    Call MapWindowPoints(0, m_hMDIClient, tPt, 1)
    
    If tPt.y < 0 Then
        OnNCHittest = HTBORDER
    Else
        Call m_tabManager.OnMouseLeave
    End If

End Function

Private Sub RequestFrameChanges()
    Call SetWindowPos(m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Private Function GetMousePos(ByVal lParam As Long) As POINTAPI
    GetMousePos.x = LowWord(lParam)
    GetMousePos.y = HiWord(lParam)
    
    Call MapWindowPoints(0, m_hMDIClient, GetMousePos, 1)
    
    GetMousePos.x = GetMousePos.x + 1&
    GetMousePos.y = GetMousePos.y + m_tabManager.Height
End Function

Private Property Let CtrlDown(ByVal Value As Boolean)
    If m_CtrlPressed <> Value Then
        m_CtrlPressed = Value
        m_tabManager.ShowShortcut = Value
    End If
End Property
