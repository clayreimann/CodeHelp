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

Private m_lPtr As Long
Private m_hMDIClient As Long
Private m_hMenu As Long
Private hSubMenu As Long
Private m_hWndMain As Long
Private m_bEnabled As Boolean

Private m_bShowTabStrip As Boolean
Private m_eShowCloseButton As CloseButtonPosition

Private WithEvents tabMgr As TabManager
Attribute tabMgr.VB_VarHelpID = -1

Dim groupClose As Boolean
Dim bCtrlPressed As Boolean

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
    m_bEnabled = RHS
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
    ICHPlugin_Enabled = m_bEnabled
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
    Dim dsr As ICHCore
    
    Set dsr = GetCHCore(m_lPtr)
    m_hWndMain = dsr.VBE.MainWindow.hWnd
    
    StartMonitor (ConnectMode = ext_cm_AfterStartup)
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    EndMonitor
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lReturn As Long, _
                                wParam As Long, _
                                lParam As Long)
    
    CtrlDown = GetAsyncKeyState(vbKeyControl)
    
    If bCtrlPressed Then
        Select Case wParam
            Case vbKey0 To vbKey9
                tabMgr.ShortcutActivate wParam
        End Select
    End If
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)
    On Error GoTo ERR_HANDLER
    
    Dim tPt As POINTAPI
    Dim Button As MouseButtonConstants
    
    Select Case uMsg
            
        Case WM_SHOWWINDOW
            OnShowWindow hWnd, wParam
            
        Case WM_SETTEXT
            OnTextChanged hWnd, lParam
            
        Case WM_PARENTNOTIFY
            OnParentNotify hWnd, wParam, lParam
            
        Case WM_NCPAINT
            If hWnd <> m_hMDIClient Then Exit Sub
            tabMgr.RequestRedraw
            
        Case WM_NCCALCSIZE, WM_NCHITTEST
            If hWnd <> m_hMDIClient Then Exit Sub
            'need to subclass to modify the return value
            placeSubclass = True
            before = False
        
        Case WM_SIZE
            If hWnd <> m_hMDIClient Then Exit Sub
            MoveTabStrip
        
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMOUSEMOVE, _
            WM_NCLBUTTONUP, WM_NCRBUTTONUP, WM_NCLBUTTONDBLCLK
            
            If hWnd <> m_hMDIClient Then Exit Sub
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
            
        Case WM_MOUSEMOVE
            tabMgr.OnMouseLeave
        
'        Case WM_NCACTIVATE
'            If wParam = 1 Then
'                If GetParent(hWnd) = m_hMDIClient Then
'                    FindItem(hWnd).Selected = True
'                End If
'            End If
        
        Case WM_CHILDACTIVATE
            If GetParent(hWnd) = m_hMDIClient Then
                FindItem(hWnd).Selected = True
            End If
            
        Case WM_ACTIVATE
            If hWnd = m_hWndMain Then
                If LoWord(wParam) = WA_ACTIVE Then
                    placeSubclass = True
                End If
            End If
    End Select
    
    Exit Sub
ERR_HANDLER:
    tabMgr.LockUpdateRequest False, True
    'LogToNotePad Err.Description & " " & Erl
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, _
                                ByVal uMsg As Long, _
                                wParam As Long, _
                                lParam As Long, _
                                bHandled As Boolean, _
                                lReturn As Long)

    On Error GoTo ERR_HANDLER
    
    Select Case uMsg
        Case WM_NCCALCSIZE
            OnNCCalcSize lParam
        
        Case WM_NCHITTEST
            lReturn = OnNCHittest(lParam)
            
        Case WM_ACTIVATE
            tabMgr.RequestRedraw
    End Select
    Exit Sub
ERR_HANDLER:
    tabMgr.LockUpdateRequest False, True
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
    
    EndMonitor
        
    m_hMDIClient = A_FindWindowEx(m_hWndMain, 0, "MDIClient", vbNullString)

    If m_hMDIClient = 0 Then Exit Sub
  
    Set tabMgr = New TabManager
    
    If AfterStartUp Then
        RegisterExistingWindows
    End If
    
    SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    'update client window in case child is not in maximized state
    RedrawWindow m_hMDIClient, ByVal 0, ByVal 0, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ERASE
    
    CreatePopup
    LoadSetting
End Sub

Private Sub EndMonitor()

    If m_hMDIClient = 0 Then Exit Sub
    If m_hWndMain = 0 Then Exit Sub
    
    On Error GoTo ERR_HANDLER
  
    If m_hMenu <> 0 Then
        DestroyMenu m_hMenu
        m_hMenu = 0
    End If
    
    SaveSetting "CodeHelp", ICHPlugin_Name, "CloseButton", m_eShowCloseButton
    
    SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    
ERR_HANDLER:
    Set tabMgr = Nothing
    m_hWndMain = 0
    m_hMDIClient = 0
End Sub

Private Sub RegisterExistingWindows()
    Dim lastWin As Long
    Dim item As TabItem
    
    lastWin = GetWindow(m_hMDIClient, GW_CHILD)
    'disable update to speed up loading large project
    tabMgr.LockUpdateRequest True

    Do While lastWin <> 0
        AddNewWindow lastWin
        lastWin = GetWindow(lastWin, GW_HWNDNEXT)
    Loop

    If tabMgr.IsAnyVisibleItem Then
        lastWin = A_SendMessage(m_hMDIClient, WM_MDIGETACTIVE, 0, ByVal 0)

        If lastWin <> 0 Then
            Set item = FindItem(lastWin)

            If Not item Is Nothing Then item.Selected = True
        End If
    End If

    tabMgr.LockUpdateRequest False
End Sub

Private Sub AddNewWindow(ByVal hWnd As Long)
    Dim item As TabItem
    
    Set item = tabMgr.InsertItem(GetCaption(GetWinText(hWnd, False)), , , IsWindowVisible(hWnd))
    item.Tag = hWnd
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
    Dim item As TabItem

    For Each item In tabMgr.Items

        If CLng(item.Tag) = hWnd Then
            Set FindItem = item
            Exit For
        End If

    Next

End Function

Private Sub tabMgr_DrawFinished(ByVal hdc As Long)
    ReleaseDC m_hMDIClient, hdc
End Sub

Private Sub tabMgr_ItemClick(ByVal item As TabItem)

    If item.Index > 0 Then
        A_SendMessage m_hMDIClient, WM_MDIACTIVATE, ByVal CLng(item.Tag), ByVal 0
    End If

End Sub

Private Sub tabMgr_ItemClosing(ByVal item As TabItem, _
                               ByVal FromButton As Boolean, _
                               Cancel As Boolean)

    If FromButton Then groupClose = True
    If groupClose Then
        A_SendMessage CLng(item.Tag), WM_SYSCOMMAND, SC_CLOSE, ByVal 0
    End If

    If FromButton Then groupClose = False
End Sub

Private Sub tabMgr_MouseUp(ByVal Button As MouseButtonConstants, _
                           ByVal item As TabItem)
    Dim menuID As Long
    Dim itemhWnd As Long
    
    If Button = vbRightButton Then
        menuID = ShowPopup(item)

        If menuID = 0 Then Exit Sub
        
        Select Case menuID

            Case ID_CLOSE
                itemhWnd = CLng(item.Tag)
                A_SendMessage itemhWnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0
                
            Case ID_CLOSE_ALL, ID_CLOSE_BUT_ACTIVE
                groupClose = True

                If menuID = ID_CLOSE_ALL Then
                    tabMgr.RemoveAll True
                Else
                    tabMgr.RemoveAllButActive True
                End If

                groupClose = False
                SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
                            
            Case CloseButtonHidden, CloseButtonRightMost, CloseButtonOnActive
                m_eShowCloseButton = menuID
                tabMgr.PaintManager.ShowCloseButton = menuID
                
                UpdateMenu menuID
                
        End Select
    
    End If

End Sub

Private Sub tabMgr_RequestRedraw(hdc As Long)
       
    If m_bShowTabStrip Then
        hdc = GetWindowDC(m_hMDIClient)
        SetBkMode hdc, TRANSPARENT
    End If

End Sub

Private Sub MoveTabStrip()
    Dim tR As RECT
    Dim lWidth As Long
  
    GetWindowRect m_hMDIClient, tR
    lWidth = tR.Right - tR.Left

    tabMgr.Move 1&, 1&, lWidth - 1&, TAB_HEIGHT
End Sub

Private Sub CreatePopup()
    
    m_hMenu = CreatePopupMenu()
    hSubMenu = CreatePopupMenu()
    A_AppendMenuStr m_hMenu, MF_STRING, ID_CLOSE, "Close"
    A_AppendMenuStr m_hMenu, MF_STRING, ID_CLOSE_ALL, "Close All"
    A_AppendMenuStr m_hMenu, MF_STRING, ID_CLOSE_BUT_ACTIVE, "Close All But Active"
    A_AppendMenu m_hMenu, MF_SEPARATOR, 0, ByVal 0
    A_AppendMenuStr m_hMenu, MF_POPUP, hSubMenu, "Close Button"
    
    A_AppendMenuStr hSubMenu, MF_STRING, CloseButtonHidden, "Hidden"
    A_AppendMenuStr hSubMenu, MF_STRING, CloseButtonRightMost, "Rightmost"
    A_AppendMenuStr hSubMenu, MF_STRING, CloseButtonOnActive, "On Active Tab"
    
End Sub

Private Function ShowPopup(ByVal item As TabItem) As Long
    Dim tPt As POINTAPI
  
    If tabMgr.Items.Count > 1 Then
        EnableMenuItem m_hMenu, ID_CLOSE_ALL, MF_ENABLED
        EnableMenuItem m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_ENABLED
    Else
        EnableMenuItem m_hMenu, ID_CLOSE_ALL, MF_GRAYED
        EnableMenuItem m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_GRAYED
    End If
    
    If Not item Is Nothing Then
        EnableMenuItem m_hMenu, ID_CLOSE, MF_ENABLED
    Else
        EnableMenuItem m_hMenu, ID_CLOSE, MF_GRAYED
    End If
    
    GetCursorPos tPt
    ShowPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD Or TPM_NONOTIFY, tPt.x, tPt.y, ByVal 0, m_hWndMain, ByVal 0)
    
End Function

Private Sub LoadSetting()
'    Dim reg As cRegistry
'
'    Set reg = New cRegistry
    
    On Error Resume Next

'    With reg
'        .ClassKey = HKEY_CURRENT_USER
'        .SectionKey = "Software\Microsoft\Visual Basic\6.0\Addins\CHCore.Connect"
'        .ValueKey = "HideClose"
'        .ValueType = REG_DWORD
'        On Error Resume Next
'        m_eShowCloseButton = Abs(CLng(.Value))
'    End With
    m_eShowCloseButton = CLng(GetSetting("CodeHelp", ICHPlugin_Name, "CloseButton", CloseButtonOnActive))
    tabMgr.PaintManager.ShowCloseButton = m_eShowCloseButton
    UpdateMenu tabMgr.PaintManager.ShowCloseButton
'    Set reg = Nothing
End Sub

'Private Sub SaveSetting()
'    Dim reg As cRegistry
'    Set reg = New cRegistry
'
'    With reg
'        .ClassKey = HKEY_CURRENT_USER
'        .SectionKey = "Software\Microsoft\Visual Basic\6.0\Addins\CHCore.Connect"
'        .ValueKey = "HideClose"
'        .ValueType = REG_DWORD
'        .Value = CLng(m_eShowCloseButton)
'    End With
'
'    Set reg = Nothing
    
'End Sub

Private Sub UpdateMenu(ByVal menuID As Long)
    Dim i As Long
    Dim id As Long
    
    For i = 0 To 2
        id = 2 ^ i

        If id = menuID Then
            CheckMenuItem hSubMenu, id, MF_BYCOMMAND Or MF_CHECKED
        Else
            CheckMenuItem hSubMenu, id, MF_BYCOMMAND Or MF_UNCHECKED
        End If

    Next

End Sub

'Message processing
Private Sub OnShowWindow(ByVal hWnd As Long, ByVal wParam As Long)
    If GetParent(hWnd) <> m_hMDIClient Then Exit Sub
    
    Dim item As TabItem
    
    Set item = FindItem(hWnd)
    
    If item Is Nothing Then Exit Sub
    
    item.Visible = (wParam = 1)
    
    RequestFrameChanges
        
    If wParam Then
        If IsZoomed(hWnd) = 0 Then _
            RedrawWindow m_hMDIClient, ByVal 0, ByVal 0, _
                RDW_ERASENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
    
    End If
End Sub

Private Sub OnTextChanged(ByVal hWnd As Long, ByVal lParam As Long)
    If GetParent(hWnd) <> m_hMDIClient Then Exit Sub
    
    Dim sText As String, i As Long
    
    sText = String$(128, vbNullChar)
    A_lstrcpyPtrStr sText, ByVal lParam
    i = InStr(1, sText, vbNullChar) - 1&
    If i > 0 Then
        sText = Left$(sText, i)
    End If
    FindItem(hWnd).Caption = GetCaption(sText)
End Sub

Private Sub OnParentNotify(ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long)
    If hWnd <> m_hMDIClient Then Exit Sub
    
    Dim item As TabItem
    Select Case LoWord(wParam)
        Case WM_CREATE
            AddNewWindow lParam
            
        Case WM_DESTROY
            If groupClose Then Exit Sub
            Call tabMgr.RemoveItem(FindItem(lParam), False)

    End Select
End Sub

Private Sub OnNCCalcSize(ByRef lParam As Long)
    Dim tNC As NCCALCSIZE_PARAMS
    
    m_bShowTabStrip = tabMgr.IsAnyVisibleItem
    
    CopyMemory tNC, ByVal lParam, Len(tNC)
    InflateRect tNC.rgrc(0), 1, 1
    
    If m_bShowTabStrip Then
        tNC.rgrc(0).Top = tNC.rgrc(0).Top + TAB_HEIGHT '- 1&
        LSet tNC.rgrc(1) = tNC.rgrc(0)
        MoveTabStrip
    End If
    
    CopyMemory ByVal lParam, tNC, Len(tNC)
    
End Sub

Private Function OnNCHittest(ByVal lParam As Long) As Long
    Dim tPt As POINTAPI
    tPt.x = LoWord(lParam): tPt.y = HiWord(lParam)
    MapWindowPoints 0, m_hMDIClient, tPt, 1
    
    If tPt.y < 0 Then
        OnNCHittest = HTBORDER
    Else
        tabMgr.OnMouseLeave
    End If

End Function

Private Sub RequestFrameChanges()
    SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Function GetMousePos(ByVal lParam As Long) As POINTAPI
    GetMousePos.x = LoWord(lParam)
    GetMousePos.y = HiWord(lParam)
    
    'MapWindowPoints HWND_DESKTOP, m_hMDIClient, GetMousePos, 1
    MapWindowPoints 0, m_hMDIClient, GetMousePos, 1
    
    GetMousePos.x = GetMousePos.x + 1&
    GetMousePos.y = GetMousePos.y + tabMgr.Height
End Function

Private Property Let CtrlDown(ByVal Value As Boolean)
    If bCtrlPressed <> Value Then
        bCtrlPressed = Value
        tabMgr.ShowShortcut = Value
        tabMgr.RequestRedraw
    End If
End Property
