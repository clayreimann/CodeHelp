VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Items As SafeCollection
Private m_oFont As IFont
Private m_SelectedItem As TabItem
Private m_oPaintMgr As TabPaintManager

Event RequestRedraw(ByRef hdc As Long)
Event DrawFinished(ByVal hdc As Long)

Event ItemChanging(ByVal toItem As TabItem, ByRef Cancel As Boolean)
Event ItemClick(ByVal item As TabItem)
Event MouseUp(ByVal Button As MouseButtonConstants, ByVal item As TabItem)
Event ItemClosing(ByVal item As TabItem, ByVal FromButton As Boolean, ByRef Cancel As Boolean)

Private m_lFirstItem As Long
Private m_lLastItem As Long
Private m_lLeft As Long
Private m_lTop As Long
Private m_lheight As Long
Private m_lWidth As Long
Private m_lVisibleCount As Long
Private m_ShowShortcut As Boolean
Private m_ShortcutItems(vbKey1 To vbKey9) As TabItem
Private m_lUpdateLock As Long
Dim m_CloseFromButton As Boolean

Dim m_lLeftOffset As Long
Dim m_oClickedTab As TabItem
Dim m_oHoveredTab As TabItem
Dim m_ScrollTimer As cTimer

Implements ITimer

Private Sub Class_Initialize()
    Set m_Items = New SafeCollection
    Set m_oPaintMgr = New TabPaintManager
    Set m_oPaintMgr.Owner = Me
    m_lFirstItem = -1
    m_lLastItem = -1
    m_lheight = 22
    Set m_ScrollTimer = New cTimer
    
End Sub

Private Sub Class_Terminate()
    'clear shortcut cache
    Dim lIdx As Long
    For lIdx = vbKey1 To vbKey9
        Set m_ShortcutItems(lIdx) = Nothing
    Next

    Set m_SelectedItem = Nothing
    Set m_oClickedTab = Nothing
    Set m_Items = Nothing
    Set m_oFont = Nothing
    Set m_oPaintMgr = Nothing
    Set m_ScrollTimer = Nothing
End Sub

Private Sub ResetState()
    Set m_oClickedTab = Nothing
    Set m_SelectedItem = Nothing
    m_lLeftOffset = 0
    m_lFirstItem = -1
    m_lLastItem = -1
    m_lVisibleCount = 0
End Sub

Public Property Get PaintManager() As TabPaintManager
    Set PaintManager = m_oPaintMgr
End Property

Public Property Get Font() As IFont

    If m_oFont Is Nothing Then
        Set m_oFont = New StdFont
        m_oFont.Name = "Tahoma"
        m_oFont.Size = 8
    End If

    Set Font = m_oFont
End Property

Public Property Set Font(ByVal objFont As IFont)
    Set m_oFont = objFont
End Property

Function InsertItem(ByVal Caption As String, Optional ByVal Index As Long = -1, Optional Icon As IPictureDisp = Nothing, _
                    Optional Visible As Boolean = True) As TabItem
    
    Dim oNewTab As TabItem
    
    Set oNewTab = New TabItem
    
    Call m_Items.AddItem(oNewTab, "#" & CStr(m_Items.Count + 1))
    
    With oNewTab
        If Index = -1 Then
            .Index = m_Items.Count
        End If

        .Caption = Caption
        .Visible = Visible
        .Width = GetTabWidth(oNewTab)
        Set .Icon = Icon
        Set .Owner = Me
    End With
    
    If m_lFirstItem < 1 And Visible Then
        m_lFirstItem = 1
        m_lLastItem = 1
    End If

    Call RecalculateLayout
    
    Set InsertItem = oNewTab

    If Visible Then
        m_lVisibleCount = m_lVisibleCount + 1&
        Call RequestRedraw
    End If

End Function

Public Sub RemoveAll(Optional ByVal RequestConfirmation As Boolean = False)
    Dim lIdx As Long
    Call LockUpdateRequest(True)
    
    If RequestConfirmation Then
        For lIdx = 1 To m_Items.Count
            If RemoveItem(m_Items("#1")) = False Then Exit For
        Next
    Else
        Set m_Items = New SafeCollection
    End If
    
    Call ResetState
    Call LockUpdateRequest(False)
    Call RecalculateLayout(True)
End Sub

Public Sub RemoveAllButActive(Optional ByVal RequestConfirmation As Boolean = False)
    Dim oNewTab As TabItem
    Dim lIdx As Long
    
    Call LockUpdateRequest(True)

    If RequestConfirmation = False Then
        ' destroy the collection and re-add the active oNewTab to the collection
        Set oNewTab = m_SelectedItem
        RemoveAll False

        If Not oNewTab Is Nothing Then
            Set oNewTab = InsertItem(oNewTab.Caption, , oNewTab.Icon, True)
            oNewTab.Selected = True
        End If

        m_lVisibleCount = 1
    Else
        ' while the first item in the collection is not the selected item remove it
        ' then remove the second item till there is only one left
        For lIdx = 1 To m_Items.Count - 1&
            Set oNewTab = m_Items("#1")

            If oNewTab.Selected Then
                Set oNewTab = m_Items("#2")
            End If

            If RemoveItem(oNewTab) = False Then Exit For
        Next

        Call CaptionChanged(1)
    End If
    
    Call LockUpdateRequest(False)
    Call RecalculateLayout(True)
End Sub

Friend Sub RecalculateLayout(Optional Redraw As Boolean = False)
    Dim lIdx As Long
    Dim oNewTab As TabItem
    Dim tR As RECT
    Dim startCount As Long
    Dim lTabBarWidth As Long
    Dim lCount As Long
    
    If m_lUpdateLock > 0 Then Exit Sub
    
    'need to calculate to flag the navbar even if items are empty
    lTabBarWidth = m_oPaintMgr.GetMaxWidth(m_lWidth)
    lCount = m_Items.Count
    
    If lCount = 0 Then
        Call m_oPaintMgr.LayoutChanged
        Exit Sub
    End If
        
    tR.Right = lTabBarWidth
    startCount = m_lLastItem

    If Not m_SelectedItem Is Nothing Then
        If m_SelectedItem.Index > m_lFirstItem Then
            'we try to keep the selected item visible
            startCount = m_SelectedItem.Index
        End If
    End If

    If startCount > lCount Then startCount = lCount
    m_lFirstItem = 1

    For lIdx = startCount To 1 Step -1
        Set oNewTab = m_Items("#" & CStr(lIdx))

        If oNewTab.Visible Then
            tR.Left = tR.Right - oNewTab.Width

            'If lIdx = m_lFirstItem Then Exit For
            If tR.Left <= 0 Then
                m_lFirstItem = lIdx
                m_lLeftOffset = tR.Left
                
                Exit For
            End If

            tR.Right = tR.Left
        End If

    Next
    
    If m_lLeftOffset > 0 Then m_lLeftOffset = 0
    If tR.Left > 0 Then m_lLeftOffset = 0
    If m_lFirstItem = 1 And m_lLeftOffset = 0 Then
        'Calculate last visible oNewTab
        m_lLastItem = -2 'all shown
        tR.Left = m_lLeftOffset

        For lIdx = m_lFirstItem To lCount
            Set oNewTab = m_Items("#" & lIdx)

            If oNewTab.Visible Then
                tR.Right = tR.Left + oNewTab.Width

                If tR.Right >= lTabBarWidth Then
                    m_lLastItem = lIdx
                    Exit For
                End If

                tR.Left = tR.Right
            End If

        Next

        If m_lLastItem = -2 Then m_lLastItem = lCount
    End If

    Call m_oPaintMgr.LayoutChanged
    
    If Redraw Then Call RequestRedraw
End Sub

Private Function GetTabWidth(ByVal item As TabItem) As Long
    GetTabWidth = m_oPaintMgr.GetTabWidth(item)
End Function

Public Property Set SelectedItem(ByVal Value As TabItem)
    Dim oTabItem As TabItem
    
    'ensure only one item is active
    If m_SelectedItem Is Value Then
        If Not m_SelectedItem Is Nothing Then
            Call EnsureVisible(m_SelectedItem.Index)
        End If
    Else
        For Each oTabItem In m_Items
            Call oTabItem.SetActive(False)
        Next
        
        Call LockUpdateRequest(True)
        
        If Not m_SelectedItem Is Nothing Then
            m_SelectedItem.Selected = False
            m_SelectedItem.Width = GetTabWidth(m_SelectedItem)
        End If
                
        Set m_SelectedItem = Value
        m_SelectedItem.SetActive True
        m_SelectedItem.Width = GetTabWidth(m_SelectedItem)
        
        Call EnsureVisible(Value.Index)
        Call LockUpdateRequest(False, True)
        Call RequestRedraw
    End If

End Property

Public Property Get SelectedItem() As TabItem
    Set SelectedItem = m_SelectedItem
End Property

Public Property Get Items() As SafeCollection
    Set Items = m_Items
End Property

Public Property Let Width(ByVal Value As Long)
    If m_lWidth <> Value Then
        m_lWidth = Value
        Call RecalculateLayout(True)
    End If
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Sub Refresh(ByVal hdc As Long)
    If hdc = 0 Then Exit Sub
    
    Call DrawTabs(hdc)
    RaiseEvent DrawFinished(hdc)
End Sub

Friend Sub RequestRedraw()
    Dim hdc As Long

    If m_lUpdateLock = 0 Then
        RaiseEvent RequestRedraw(hdc)
        Call Refresh(hdc)
    End If

End Sub

Private Sub EnsureVisible(ByVal Index As Long)
    Dim lIdx As Long
    Dim item As TabItem
    Dim tR As RECT
    Dim maxWidth As Long
    
    If Index = m_lFirstItem Then
        If m_lLeftOffset <> 0 Then
            m_lLeftOffset = 0
            Call RequestRedraw
        End If

    ElseIf Index < m_lFirstItem Then
        m_lFirstItem = Index
        m_lLeftOffset = 0
        Call RequestRedraw
        
    Else
        maxWidth = m_oPaintMgr.GetMaxWidth(m_lWidth)

        If IsOffScreenToTheRight(Index, maxWidth) Then
            'Calculate new visible starting index from right to left
            tR.Right = maxWidth

            For lIdx = Index To 1 Step -1
                Set item = m_Items("#" & lIdx)

                If item.Visible Then
                    tR.Left = tR.Right - item.Width

                    If tR.Left <= 0 Then
                        m_lFirstItem = lIdx
                        m_lLeftOffset = tR.Left
                        Exit For
                    End If

                    tR.Right = tR.Left
                End If

            Next

            Call RequestRedraw
        End If
    End If

End Sub

Friend Function IsOffScreenToTheRight(ByVal Index As Long, ByVal maxWidth As Long) As Boolean
    Dim lIdx As Long
    Dim oTabItem As TabItem
    Dim lRightPos As Long
    
    lRightPos = m_lLeftOffset

    For lIdx = m_lFirstItem To Index
        Set oTabItem = m_Items("#" & CStr(lIdx))

        If oTabItem.Visible Then
            lRightPos = lRightPos + oTabItem.Width
        End If

    Next
    
    IsOffScreenToTheRight = ((lRightPos - maxWidth) > 2)
End Function

Friend Sub CaptionChanged(ByVal Index As Long)
    Dim item As TabItem
    
    Set item = m_Items("#" & Index)
    item.Width = GetTabWidth(item)

    If item.Visible Then
        If Index >= m_lFirstItem Then
            If IsOffScreenToTheRight(Index, m_oPaintMgr.GetMaxWidth(m_lWidth)) = False Then
                Call RecalculateLayout(True)
            End If
        End If
    End If

End Sub

Friend Sub VisibleChanged(ByVal Index As Long)
    If m_Items("#" & Index).Visible Then
        m_lVisibleCount = m_lVisibleCount + 1&
    Else
        m_lVisibleCount = m_lVisibleCount - 1&
    End If

    Call RecalculateLayout(True)
End Sub

Private Sub DrawTabs(ByVal hdc As Long)
    Dim lIdx As Long
    Dim item As TabItem
    Dim tR As RECT
    Dim maxSpace As Long
    Dim lCount As Long
    Dim oDC As MemoryDC
    Dim cnt As Long
    
    cnt = 1
    
    Set oDC = New MemoryDC
    Call oDC.CreateDC(m_lWidth, 22)
    Call SetBkMode(oDC.hdc, TRANSPARENT)
    
    'draw background
    tR.Left = 0
    tR.Top = 0
    tR.Right = m_lWidth
    tR.Bottom = 22
    
    'clear shortcut cache
    For lIdx = vbKey1 To vbKey9
        Set m_ShortcutItems(lIdx) = Nothing
    Next
    
    Call m_oPaintMgr.DrawItem(oDC, tR, Nothing, DrawBackground)
    
    tR.Left = m_lLeftOffset
    maxSpace = m_oPaintMgr.GetMaxWidth(m_lWidth)
    
    If m_lFirstItem > 0 Then
        lCount = VisibleTabCount

        For lIdx = m_lFirstItem To m_Items.Count
            Set item = m_Items("#" & CStr(lIdx))

            If item.Visible Then
                tR.Right = tR.Left + item.Width
                item.LeftMostItem = (cnt = 0)
                item.RightMostItem = (tR.Right >= maxSpace)
                
                                
                Call m_oPaintMgr.DrawItem(oDC, tR, item, DrawTabItem)
                
                If m_ShowShortcut And cnt < 10 Then
                    'we only support 0 to 9 shortcut number for the time being
                    Call m_oPaintMgr.DrawShortcut(oDC, tR, item, cnt)
                    
                    'save item for shortcut key activation
                    Set m_ShortcutItems(vbKey0 + cnt) = item
                End If
                
                cnt = cnt + 1&
                
                If item.RightMostItem Then
                    Exit For
                End If

                tR.Left = tR.Right
                
            End If
            
        Next

        m_lLastItem = lIdx  'mark last item for hittest boundary

        If m_lLastItem > lCount Then m_lLastItem = lCount
    End If

    Call BitBlt(hdc, m_lLeft, m_lTop, m_lWidth, m_lheight, oDC.hdc, 0, 0, vbSrcCopy)
    Set oDC = Nothing
End Sub

Friend Sub OnLMouseDown(ByVal x As Long, ByVal y As Long)
    Set m_oClickedTab = HitTest(x, y)

    If Not m_oClickedTab Is Nothing Then
        If m_oClickedTab.Index = 0 Then 'Close button, fire click on mouse up
            m_oClickedTab.Selected = True
            Call RequestRedraw
        Else
            Call OnItemClick(m_oClickedTab)
        End If
    End If

End Sub

Friend Sub OnMouseMove(ByVal Button As MouseButtonConstants, ByVal x As Long, ByVal y As Long)
    Dim oHoveredTab As TabItem
    
    Set oHoveredTab = HitTest(x, y)
    
    If Not m_oHoveredTab Is oHoveredTab Then
        If Not m_oHoveredTab Is Nothing Then m_oHoveredTab.Hovered = False
        
        Set m_oHoveredTab = oHoveredTab
        If Not m_oHoveredTab Is Nothing Then
            
            If Button = vbLeftButton Then 'Dragging
                If Not m_oClickedTab Is Nothing Then
                    If m_oClickedTab.Index > 0 And m_oHoveredTab.Index > 0 Then
                        Call SwapItem(m_oClickedTab, m_oHoveredTab)
                        Set m_oHoveredTab = m_oClickedTab
                    Else
                        m_oHoveredTab.Hovered = True
                    End If
                End If

            Else
                m_oHoveredTab.Hovered = True
            End If
        End If

        Call RequestRedraw
    End If

End Sub

Friend Sub OnMouseUp(ByVal Button As MouseButtonConstants, ByVal x As Long, ByVal y As Long)
    Dim upItem As TabItem
    Set upItem = HitTest(x, y)
    
    If Button = vbLeftButton Then
        If m_oClickedTab Is Nothing Then Exit Sub
        
        If m_oClickedTab.Index < 1 Then
            m_oClickedTab.Selected = False 'Reset flag
            Call RequestRedraw
        End If
        
        If upItem Is m_oClickedTab Then 'MouseDown and MouseUp in the same item
            
            If m_oClickedTab.Index > -1 Then 'nav button executed on mousedown
                RaiseEvent MouseUp(vbLeftButton, upItem)
                Call OnItemClick(m_oClickedTab)
            End If
            
        ElseIf m_oClickedTab Is m_SelectedItem Then

            If Not upItem Is Nothing Then
                If upItem.Index > 0 Then
                    Call SwapItem(m_oClickedTab, upItem)
                End If
            End If
        End If

    Else

        If Not upItem Is Nothing Then
            If upItem.Index < 1 Then Exit Sub
        End If

        RaiseEvent MouseUp(vbRightButton, upItem)
    End If

    Call m_ScrollTimer.TmrStop
    Set m_oClickedTab = Nothing
End Sub

Friend Sub OnMouseLeave()
    If Not m_oHoveredTab Is Nothing Then
        m_oHoveredTab.Hovered = False
        Set m_oHoveredTab = Nothing
        Call RequestRedraw
    End If

End Sub

Friend Sub ShortcutActivate(ByVal key As Long)
    Dim item As TabItem
    
    Set item = m_ShortcutItems(key)
    If Not item Is Nothing Then
        Call OnItemClick(item)
    End If
End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long) As TabItem
    Dim lIdx As Long, item As TabItem
    Dim tR As RECT
    Dim lMaxWidth As Long
    
    'try nav buttons first
    Set HitTest = m_oPaintMgr.ButtonHittest(x, y)

    If HitTest Is Nothing Then
        
        tR.Top = m_lTop
        tR.Bottom = m_lTop + m_lheight
        tR.Left = m_lLeftOffset + m_lLeft

        If m_lFirstItem > 0 Then
            lMaxWidth = m_oPaintMgr.GetMaxWidth(m_lWidth)

            For lIdx = m_lFirstItem To m_Items.Count
                Set item = m_Items("#" & lIdx)

                If item.Visible Then
                    tR.Right = tR.Left + item.Width

                    If PtInRect(tR, x, y) Then
                        Set HitTest = item
                        Exit For
                    End If

                    'iterate only in visible boundary
                    If tR.Right >= lMaxWidth Then Exit For
                    tR.Left = tR.Right
                End If

            Next

        End If

    Else

        If HitTest.Enabled = False Then Set HitTest = Nothing
    End If

End Function

Public Function RemoveItem(ByVal item As TabItem, Optional AutoActivateNext As Boolean = True) As Boolean
    'We're synchronizing the collection key with its index,
    'since item access by key is much faster than by its index
    Dim key As String, itemIdx As Long
    Dim lIdx As Long, lPos As Long
    Dim nextItem As TabItem
    Dim isActive As Boolean
    Dim Cancel As Boolean
    Dim lCount As Long
    Dim isVisible As Boolean
    
    If item Is Nothing Then Exit Function
    
    itemIdx = item.Index
    key = "#" & CStr(itemIdx)
    isActive = item Is m_SelectedItem
    
    If m_Items.Exists(key) Then
        
        RaiseEvent ItemClosing(item, m_CloseFromButton, Cancel)
        
        If Cancel Then Exit Function
        isVisible = item.Visible
        Call LockUpdateRequest(True)
        Call m_Items.Remove(key)
        RemoveItem = True
        lCount = m_Items.Count

        If lCount > 0 Then
            If itemIdx <> lCount + 1& Then
                'Shift all element down to fill the void & modify the key
                For lIdx = itemIdx + 1& To lCount + 1&
                    Set nextItem = m_Items("#" & CStr(lIdx))
                    Call m_Items.Remove("#" & CStr(nextItem.Index))
                    
                    lPos = lIdx - 1
                    nextItem.Index = lPos
                    Call m_Items.AddItem(nextItem, "#" & CStr(lPos), lPos)
                Next
                
                If AutoActivateNext And isActive Then
                    If lIdx > 1 Then
                        Set Me.SelectedItem = m_Items("#" & (itemIdx - 1&))
                    Else
                        Set Me.SelectedItem = m_Items(key)
                    End If
                End If
                
            ElseIf isActive Then 'Deleted item was the active item and the last item
                If AutoActivateNext Then Set Me.SelectedItem = m_Items("#" & lCount)
                
            ElseIf lCount = 1 Then
                For Each item In m_Items
                    If item.Visible Then
                        Call EnsureVisible(item.Index)
                        Exit For
                    End If
                Next
                
            End If
            
            If isVisible Then m_lVisibleCount = m_lVisibleCount - 1&
        Else
            Call ResetState
        End If
        
        Call LockUpdateRequest(False)
        Call RecalculateLayout(True)
    End If

End Function

Private Sub SwapItem(ByVal itemFrom As TabItem, ByVal itemTo As TabItem)
    Dim fromIdx As Long, toIdx As Long
    Dim item1 As TabItem, item2 As TabItem
    
    'Hold references
    Set item1 = itemFrom
    Set item2 = itemTo
    
    fromIdx = itemFrom.Index
    toIdx = itemTo.Index
    
    'swap index
    item1.Index = toIdx
    item2.Index = fromIdx
    
    'delete and re-add from collection
    Call m_Items.Remove("#" & CStr(fromIdx))
    Call m_Items.AddItem(itemTo, "#" & CStr(fromIdx), fromIdx)
    Call m_Items.Remove("#" & CStr(toIdx))
    Call m_Items.AddItem(itemFrom, "#" & CStr(toIdx), toIdx)
    
    Call RequestRedraw
End Sub

'prevent the painting and layout operation from executing
'use this when doing several updates at once, make sure calls are balanced
Friend Sub LockUpdateRequest(ByVal ShouldLock As Boolean, Optional ResetCount As Boolean = False)
    If ResetCount Then
        m_lUpdateLock = 0
        Exit Sub
    End If
    
    If ShouldLock Then
        m_lUpdateLock = m_lUpdateLock + 1&
    Else
        m_lUpdateLock = m_lUpdateLock - 1&

        If m_lUpdateLock < 0 Then m_lUpdateLock = 0
    End If

End Sub

Private Sub OnItemClick(ByVal item As TabItem)
    Dim Cancel As Boolean
    
    Select Case item.Index
        Case 0 ' close button
            If Not m_SelectedItem Is Nothing Then
                m_CloseFromButton = True
                Call RemoveItem(m_SelectedItem, False)
                m_CloseFromButton = False
            End If
        
        Case -1, -2 ' left/right
            m_oClickedTab.Selected = True
            Call m_ScrollTimer.TmrStart(Me, 50)
            
        Case Else
            RaiseEvent ItemChanging(item, Cancel)

            If Cancel = False Then
                Set SelectedItem = item
                RaiseEvent ItemClick(item)
            End If

    End Select

End Sub

Friend Property Get StartItem() As Long
    StartItem = m_lFirstItem
End Property

Friend Property Let StartItem(ByVal lStartItem As Long)
    If lStartItem < 1 Then lStartItem = 1
    If lStartItem <> m_lFirstItem Then
        m_lFirstItem = lStartItem
    End If
End Property

Friend Property Get LastItem() As Long
    LastItem = m_lLastItem
End Property

Friend Property Let LastItem(ByVal lLastItem As Long)
    Dim lCount As Long
    lCount = VisibleTabCount

    If lLastItem > lCount Then lLastItem = lCount
    If m_lLastItem <> lLastItem Then
        m_lLastItem = lLastItem
    End If

End Property

Friend Property Get LeftItemOffset() As Long
    LeftItemOffset = m_lLeftOffset
End Property

Friend Function GetTotalWidth() As Long
    Dim item As TabItem

    For Each item In m_Items
        If item.Visible Then
            GetTotalWidth = GetTotalWidth + item.Width
        End If
    Next
End Function

Friend Property Get VisibleTabCount() As Long
    Dim lIdx As Long, item As TabItem

    If m_Items Is Nothing Then Exit Property

    For Each item In m_Items
        If item.Visible Then
            lIdx = lIdx + 1&
        End If
    Next
    
    VisibleTabCount = lIdx
End Property

Public Sub Move(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    m_lLeft = lLeft
    m_lTop = lTop
    m_lWidth = lWidth
    m_lheight = lHeight
    
    Call RecalculateLayout(True)
End Sub

Public Property Get Left() As Long
    Left = m_lLeft
End Property

Public Property Let Left(ByVal lLeft As Long)
    If m_lLeft <> lLeft Then
        m_lLeft = lLeft
        Call RecalculateLayout(True)
    End If
End Property

Public Property Get Top() As Long
    Top = m_lTop
End Property

Public Property Let Top(ByVal lTop As Long)
    If m_lTop <> lTop Then
        m_lTop = lTop
        Call RecalculateLayout(True)
    End If
End Property

Public Property Get Height() As Long
    Height = m_lheight
End Property

Public Property Let Height(ByVal lHeight As Long)
    If m_lheight <> lHeight Then
        m_lheight = lHeight
        Call RecalculateLayout(True)
    End If
End Property


Public Property Get ShowShortcut() As Boolean
    ShowShortcut = m_ShowShortcut
End Property

Public Property Let ShowShortcut(ByVal bShowShortcut As Boolean)
    If m_ShowShortcut <> bShowShortcut Then
        m_ShowShortcut = bShowShortcut
        Call RequestRedraw
    End If
End Property

Private Sub Scroll(ByVal direction As Long)
    '-1 left, -2 Right
    Dim Cancel As Boolean
    Cancel = True

    If direction = -1 Then
        If m_lLeftOffset = 0 Then
            If m_lFirstItem > 1 Then
                m_lFirstItem = m_lFirstItem - 2&
                Cancel = False
            End If
        Else
            m_lLeftOffset = 0
            Cancel = False
        End If

        If Not Cancel Then
            Call m_oPaintMgr.LayoutChanged
            Call RequestRedraw
        End If
    
    Else
        LastItem = m_lLastItem + 2&
        Call EnsureVisible(m_lLastItem)
        Call m_oPaintMgr.LayoutChanged
    End If
    
    If Not m_oClickedTab.Enabled Then
        Call m_ScrollTimer.TmrStop
        Call RequestRedraw
    End If

End Sub

Private Sub ITimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
    If Not m_oClickedTab Is Nothing Then
        Call Scroll(m_oClickedTab.Index)
    Else
        Call m_ScrollTimer.TmrStop
    End If

End Sub

Public Function IsAnyVisibleItem() As Boolean
    Dim item As TabItem
    
    For Each item In m_Items
        If item.Visible Then
            IsAnyVisibleItem = True
            Exit For
        End If
    Next
End Function


