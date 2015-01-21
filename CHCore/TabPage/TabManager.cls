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

Private m_Items As ReadOnlyColl
Private m_objFont As IFont
Private m_selectedItem As TabItem
Private m_oPaintMgr As TabPaintManager

Event RequestRedraw(ByRef hdc As Long)
Event DrawFinished(ByVal hdc As Long)

Event ItemChanging(ByVal toItem As TabItem, ByRef Cancel As Boolean)
Event ItemClick(ByVal item As TabItem)
Event MouseUp(ByVal Button As MouseButtonConstants, ByVal item As TabItem)
Event ItemClosing(ByVal item As TabItem, ByVal FromButton As Boolean, ByRef Cancel As _
        Boolean)

Private m_lStartItem As Long
Private m_lLastItem As Long
Private m_lLeft As Long
Private m_lTop As Long
Private m_lheight As Long
Private m_lWidth As Long
Private m_lVisibleCount As Long
Private m_bShowShortcut As Boolean
Private m_ShortcutItems(vbKey0 To vbKey9) As TabItem

Dim lockUpdate As Long
Dim LeftOffset As Long
Dim downItem As TabItem
Dim itemHover As TabItem
Dim fromCloseButton As Boolean
Dim scrollTimer As cTimer

Implements ITimer

Private Sub Class_Initialize()
    Set m_Items = New ReadOnlyColl
    Set m_oPaintMgr = New TabPaintManager
    Set m_oPaintMgr.Owner = Me
    m_lStartItem = -1
    m_lLastItem = -1
    m_lheight = 22
    Set scrollTimer = New cTimer
    
End Sub

Private Sub Class_Terminate()
    'clear shortcut cache
    Dim i As Long
    For i = vbKey0 To vbKey9
        Set m_ShortcutItems(i) = Nothing
    Next

    Set m_selectedItem = Nothing
    Set downItem = Nothing
    Set m_Items = Nothing
    Set m_objFont = Nothing
    Set m_oPaintMgr = Nothing
    Set scrollTimer = Nothing
End Sub

Public Property Get PaintManager() As TabPaintManager
    Set PaintManager = m_oPaintMgr
End Property

Public Property Get Font() As IFont

    If m_objFont Is Nothing Then
        Set m_objFont = New StdFont
        m_objFont.Name = "Tahoma"
        m_objFont.Size = 8
    End If

    Set Font = m_objFont
End Property

Public Property Set Font(ByVal objFont As IFont)
    Set m_objFont = objFont
End Property

Function InsertItem(ByVal Caption As String, Optional ByVal Index As Long = -1, _
        Optional Icon As IPictureDisp = Nothing, Optional Visible As Boolean = True) As _
        TabItem
    
    Dim newItem As TabItem
    
    Set newItem = New TabItem
    
    m_Items.AddItem newItem, "#" & CStr(m_Items.Count + 1)
    
    With newItem

        If Index = -1 Then
            .Index = m_Items.Count
        End If

        .Caption = Caption
        .Visible = Visible
        .Width = GetTabWidth(newItem)
        Set .Icon = Icon
        Set .Owner = Me
    End With
    
    If m_lStartItem < 1 And Visible Then
        m_lStartItem = 1
        m_lLastItem = 1
    End If

    RecalculateLayout
    
    Set InsertItem = newItem

    If Visible Then
        m_lVisibleCount = m_lVisibleCount + 1&
        RequestRedraw
    End If

End Function

Public Sub RemoveAll(Optional ByVal RequestConfirmation As Boolean = False)
    Dim i As Long
    LockUpdateRequest True
    
    If RequestConfirmation Then

        For i = 1 To m_Items.Count

            If RemoveItem(m_Items("#1")) = False Then Exit For
        Next

    Else
        Set m_Items = New ReadOnlyColl
    End If
    
    ResetState
    LockUpdateRequest False
    RecalculateLayout
    RequestRedraw
End Sub

Public Sub RemoveAllButActive(Optional ByVal RequestConfirmation As Boolean = False)
    Dim item As TabItem
    Dim i As Long
    
    LockUpdateRequest True

    If RequestConfirmation = False Then
        'Fast way
        'destroy the collection and re-add the active item to the collection
        Set item = m_selectedItem
        RemoveAll False

        If Not item Is Nothing Then
            Set item = InsertItem(item.Caption, , item.Icon, True)
            item.Selected = True
        End If

        m_lVisibleCount = 1
    Else

        For i = 1 To m_Items.Count - 1&
            Set item = m_Items("#1")

            If item.Selected Then
                Set item = m_Items("#2")
            End If

            If RemoveItem(item) = False Then Exit For
        Next

        CaptionChanged 1
    End If
    
    LockUpdateRequest False
    RecalculateLayout
    RequestRedraw
End Sub

Friend Sub RecalculateLayout()
    
    Dim i As Long
    Dim item As TabItem
    Dim tR As RECT
    Dim startCount As Long
    Dim maxSpace As Long
    Dim lCount As Long
    
    If lockUpdate > 0 Then Exit Sub
    'need to calculate to flag the navbar even if items are empty
    maxSpace = m_oPaintMgr.GetMaxWidth(m_lWidth)
    lCount = m_Items.Count
    
    If lCount > 0 Then
        
        tR.Right = maxSpace
        startCount = m_lLastItem

        If Not m_selectedItem Is Nothing Then
            If m_selectedItem.Index > m_lStartItem Then
                'we try to keep selected item visible
                startCount = m_selectedItem.Index
                
            End If
        End If

        If startCount > lCount Then startCount = lCount
        m_lStartItem = 1

        For i = startCount To 1 Step -1
            Set item = m_Items("#" & i)

            If item.Visible Then
                tR.Left = tR.Right - item.Width

                'If i = StartItem Then Exit For
                If tR.Left <= 0 Then
                    m_lStartItem = i
                    LeftOffset = tR.Left
                    
                    Exit For
                End If

                tR.Right = tR.Left
            End If

        Next
        
        If LeftOffset > 0 Then LeftOffset = 0
        If tR.Left > 0 Then LeftOffset = 0
        If m_lStartItem = 1 And LeftOffset = 0 Then
            'Calculate last visible item
            m_lLastItem = -2 'all shown
            tR.Left = LeftOffset

            For i = m_lStartItem To lCount
                Set item = m_Items("#" & i)

                If item.Visible Then
                    tR.Right = tR.Left + item.Width

                    If tR.Right >= maxSpace Then
                        m_lLastItem = i
                        Exit For
                    End If

                    tR.Left = tR.Right
                End If

            Next
    
            If m_lLastItem = -2 Then m_lLastItem = lCount
        End If
        
    End If

    m_oPaintMgr.LayoutChanged
End Sub

Private Function GetTabWidth(ByVal item As TabItem) As Long
    GetTabWidth = m_oPaintMgr.GetTabWidth(item)
End Function

Public Property Set SelectedItem(ByVal Value As TabItem)
    Dim item As TabItem
    
    'ensure only one item is active
    
    If Not m_selectedItem Is Value Then
        For Each item In m_Items
            item.SetActive False
        Next
        
        LockUpdateRequest True
        
        If Not m_selectedItem Is Nothing Then
            m_selectedItem.Selected = False
            m_selectedItem.Width = GetTabWidth(m_selectedItem)
        End If
                
        Set m_selectedItem = Value
        m_selectedItem.SetActive True
        m_selectedItem.Width = GetTabWidth(m_selectedItem)
        
        EnsureVisible Value.Index
        
        LockUpdateRequest False, True

        RequestRedraw
    Else
        If Not m_selectedItem Is Nothing Then
            EnsureVisible m_selectedItem.Index
        End If
    End If

End Property

Public Property Get SelectedItem() As TabItem
    Set SelectedItem = m_selectedItem
End Property

Public Property Get Items() As ReadOnlyColl
    Set Items = m_Items
End Property

Public Property Let Width(ByVal Value As Long)

    If m_lWidth <> Value Then
        m_lWidth = Value
        RecalculateLayout
        RequestRedraw
    End If

End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Sub Refresh(ByVal hdc As Long)

    If hdc = 0 Then Exit Sub
    DrawTabs hdc
    RaiseEvent DrawFinished(hdc)
End Sub

Friend Sub RequestRedraw()
    Dim hdc As Long

    If lockUpdate = 0 Then
        RaiseEvent RequestRedraw(hdc)
        Refresh hdc
    End If

End Sub

Private Sub EnsureVisible(ByVal Index As Long)
    Dim i As Long, item As TabItem
    Dim tR As RECT
    Dim maxWidth As Long
    
    If Index = StartItem Then
        If LeftOffset <> 0 Then
            LeftOffset = 0
            RequestRedraw
        End If

    ElseIf Index < StartItem Then
        StartItem = Index
        LeftOffset = 0
        RequestRedraw
    Else
        maxWidth = m_oPaintMgr.GetMaxWidth(m_lWidth)

        If IsOffScreenToTheRight(Index, maxWidth) Then
            'Calculate new visible starting index from right to left
            tR.Right = maxWidth

            For i = Index To 1 Step -1
                Set item = m_Items("#" & i)

                If item.Visible Then
                    tR.Left = tR.Right - item.Width

                    If tR.Left <= 0 Then
                        StartItem = i
                        LeftOffset = tR.Left
                        Exit For
                    End If

                    tR.Right = tR.Left
                End If

            Next

            RequestRedraw
        End If
    End If

End Sub

Friend Function IsOffScreenToTheRight(ByVal Index As Long, _
                                      ByVal maxWidth As Long) As Boolean
    Dim i As Long
    Dim item As TabItem
    Dim lRightPos As Long
    
    lRightPos = LeftOffset

    For i = StartItem To Index
        Set item = m_Items("#" & i)

        If item.Visible Then
            lRightPos = lRightPos + item.Width
        End If

    Next
    
    IsOffScreenToTheRight = ((lRightPos - maxWidth) > 2)
End Function

Friend Sub CaptionChanged(ByVal Index As Long)
    Dim item As TabItem
    
    Set item = m_Items("#" & Index)
    item.Width = GetTabWidth(item)

    If item.Visible Then
        If Index >= StartItem Then
            If IsOffScreenToTheRight(Index, m_oPaintMgr.GetMaxWidth(m_lWidth)) = False _
                    Then
                RecalculateLayout
                RequestRedraw
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

    RecalculateLayout
    RequestRedraw
End Sub

Private Sub DrawTabs(ByVal hdc As Long)
    
    Dim i As Long
    Dim item As TabItem
    Dim tR As RECT
    Dim maxSpace As Long
    Dim lCount As Long
    Dim oDC As MemoryDC
    Dim j As Long
    
    Set oDC = New MemoryDC
    oDC.CreateDC m_lWidth, 22
    SetBkMode oDC.hdc, TRANSPARENT
    'draw background
    tR.Left = 0
    tR.Top = 0
    tR.Right = m_lWidth
    tR.Bottom = 22
    
    'clear shortcut cache
    For i = vbKey0 To vbKey9
        Set m_ShortcutItems(i) = Nothing
    Next
    
    m_oPaintMgr.DrawItem oDC, tR, Nothing, DrawBackGround
    
    tR.Left = LeftOffset
    maxSpace = m_oPaintMgr.GetMaxWidth(m_lWidth)
    
    If m_lStartItem > 0 Then
        lCount = VisibleTabCount

        For i = m_lStartItem To m_Items.Count
            Set item = m_Items("#" & i)

            If item.Visible Then
                tR.Right = tR.Left + item.Width
                item.LeftMostItem = (j = 0)
                item.RightMostItem = (tR.Right >= maxSpace)
                
                                
                m_oPaintMgr.DrawItem oDC, tR, item, DrawTabItem
                
                If m_bShowShortcut And j < 10 Then
                    'we only support 0 to 9 shortcut number for the time being
                    m_oPaintMgr.DrawShortcut oDC, tR, item, j
                    'save item for shortcut key activation
                    Set m_ShortcutItems(vbKey0 + j) = item
                End If
                
                j = j + 1&
                
                If item.RightMostItem Then
                    Exit For
                End If

                tR.Left = tR.Right
                
            End If
            
        Next

        m_lLastItem = i  'mark last item for hittest boundary

        If m_lLastItem > lCount Then m_lLastItem = lCount
    End If

    BitBlt hdc, m_lLeft, m_lTop, m_lWidth, m_lheight, oDC.hdc, 0, 0, vbSrcCopy
    Set oDC = Nothing
End Sub

Friend Sub OnLMouseDown(ByVal x As Long, _
                        ByVal y As Long)
    Set downItem = Hittest(x, y)

    If Not downItem Is Nothing Then
        If downItem.Index = 0 Then 'Close button, fire click on mouse up
            downItem.Selected = True
            RequestRedraw
        Else
            OnItemClick downItem
        End If
    End If

End Sub

Friend Sub OnMouseMove(ByVal Button As MouseButtonConstants, _
                       ByVal x As Long, _
                       ByVal y As Long)
    Dim lastHover As TabItem
    
    Set lastHover = itemHover
    Set itemHover = Hittest(x, y)
    
    If Not itemHover Is lastHover Then
        
        If Not lastHover Is Nothing Then lastHover.Hovered = False
        If Not itemHover Is Nothing Then
            
            If Button = vbLeftButton Then 'Dragging
                If Not downItem Is Nothing Then
                    If downItem.Index > 0 And itemHover.Index > 0 Then
                        SwapItem downItem, itemHover
                        Set itemHover = downItem
                    Else
                        itemHover.Hovered = True
                    End If
                End If

            Else
                itemHover.Hovered = True
            End If
        End If

        RequestRedraw
    End If

End Sub

Friend Sub OnMouseUp(ByVal Button As MouseButtonConstants, _
                     ByVal x As Long, _
                     ByVal y As Long)
    Dim upItem As TabItem
    Set upItem = Hittest(x, y)
    
    If Button = vbLeftButton Then
        If downItem Is Nothing Then Exit Sub
        
        If downItem.Index < 1 Then
            downItem.Selected = False 'Reset flag
            RequestRedraw
        End If
        
        If upItem Is downItem Then 'MouseDown and MouseUp in the same item
            
            If downItem.Index > -1 Then 'nav button executed on mousedown
                RaiseEvent MouseUp(vbLeftButton, upItem)
                OnItemClick downItem
            End If
            
        ElseIf downItem Is m_selectedItem Then

            If Not upItem Is Nothing Then
                If upItem.Index > 0 Then
                    SwapItem downItem, upItem
                End If
            End If
        End If

    Else

        If Not upItem Is Nothing Then
            If upItem.Index < 1 Then Exit Sub
        End If

        RaiseEvent MouseUp(vbRightButton, upItem)
    End If

    scrollTimer.TmrStop
    Set downItem = Nothing
End Sub

Friend Sub OnMouseLeave()

    If Not itemHover Is Nothing Then
        itemHover.Hovered = False
        Set itemHover = Nothing
        RequestRedraw
    End If

End Sub

Friend Sub ShortcutActivate(ByVal key As Long)
    Dim item As TabItem
    
    Set item = m_ShortcutItems(key)
    If Not item Is Nothing Then
        OnItemClick item
    End If
End Sub

Public Function Hittest(ByVal x As Long, _
                        ByVal y As Long) As TabItem
    Dim i As Long, item As TabItem
    Dim tR As RECT
    Dim lMaxWidth As Long
    
    'try nav buttons first
    Set Hittest = m_oPaintMgr.ButtonHittest(x, y)

    If Hittest Is Nothing Then
        
        tR.Top = m_lTop
        tR.Bottom = m_lTop + m_lheight
        tR.Left = LeftOffset + m_lLeft

        If m_lStartItem > 0 Then
            lMaxWidth = m_oPaintMgr.GetMaxWidth(m_lWidth)

            For i = StartItem To m_Items.Count
                Set item = m_Items("#" & i)

                If item.Visible Then
                    tR.Right = tR.Left + item.Width

                    If PtInRect(tR, x, y) Then
                        Set Hittest = item
                        Exit For
                    End If

                    'iterate only in visible boundary
                    If tR.Right >= lMaxWidth Then Exit For
                    tR.Left = tR.Right
                End If

            Next

        End If

    Else

        If Hittest.Enabled = False Then Set Hittest = Nothing
    End If

End Function

Public Function RemoveItem(ByVal item As TabItem, Optional AutoActivateNext As Boolean = True) As Boolean
    'We're synchronizing the collection key with its index,
    'since item access by key is much faster than by its index

    Dim key As String, idx As Long
    Dim i As Long, j As Long
    Dim nextItem As TabItem
    Dim isActive As Boolean
    Dim Cancel As Boolean
    Dim lCount As Long
    Dim isVisible As Boolean
    
    If item Is Nothing Then Exit Function
    
    idx = item.Index
    key = "#" & idx
    isActive = item Is m_selectedItem
    
    If m_Items.Exists(key) Then
        
        RaiseEvent ItemClosing(item, fromCloseButton, Cancel)
        
        If Cancel Then Exit Function
        isVisible = item.Visible
        LockUpdateRequest True
        m_Items.Remove key
        RemoveItem = True
        lCount = m_Items.Count

        If lCount > 0 Then
            If idx <> lCount + 1& Then

                'Shift all element down to fill the void & modify the key
                For i = idx + 1& To lCount + 1&
                    'Debug.Assert i <> m_Items.Count + 1
                    Set nextItem = m_Items("#" & i)
                    j = i - 1&
                    m_Items.Remove "#" & nextItem.Index
                    nextItem.Index = j

                    m_Items.AddItem nextItem, "#" & j, j
                Next
                
                If AutoActivateNext Then
                    If isActive Then 'Deleted item was the active item
                        If idx > 1 Then
                            Set SelectedItem = m_Items("#" & (idx - 1&))
                        Else
                            Set SelectedItem = m_Items(key)
                        End If
                    End If
                End If
            ElseIf isActive Then 'Deleted item was the active item and the last item
                If AutoActivateNext Then _
                    Set SelectedItem = m_Items("#" & lCount)
            ElseIf lCount = 1 Then

                For Each item In m_Items

                    If item.Visible Then
                        EnsureVisible item.Index
                        Exit For
                    End If

                Next
            
                '            ElseIf m_lLastItem > m_Items.Count Then
                '                m_lLastItem = m_Items.Count
            End If
            
            If isVisible Then m_lVisibleCount = m_lVisibleCount - 1&
        Else
            ResetState
        End If
        
        LockUpdateRequest False
        RecalculateLayout
        RequestRedraw
    End If

End Function

Private Sub SwapItem(ByVal itemFrom As TabItem, _
                     ByVal itemTo As TabItem)
    Dim fromIdx As Long
    Dim toIdx As Long
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
    m_Items.Remove "#" & fromIdx
    m_Items.AddItem itemTo, "#" & fromIdx, fromIdx
    m_Items.Remove "#" & toIdx
    m_Items.AddItem itemFrom, "#" & toIdx, toIdx
    
    RequestRedraw
End Sub

Friend Sub LockUpdateRequest(ByVal Value As Boolean, _
                             Optional reset As Boolean = False)

    'prevent the painting and layout operation from executing
    'use this when doing several updates at once
    'make sure to call it in pairs
    If reset = False Then
        If Value Then
            lockUpdate = lockUpdate + 1&
        Else
            lockUpdate = lockUpdate - 1&

            If lockUpdate < 0 Then lockUpdate = 0
        End If

    Else
        lockUpdate = 0
    End If

End Sub

Private Sub OnItemClick(ByVal item As TabItem)
    Dim Cancel As Boolean

    Select Case item.Index

        Case 0

            'This is from close button
            If Not m_selectedItem Is Nothing Then
                fromCloseButton = True
                RemoveItem m_selectedItem, False
                fromCloseButton = False
            End If

            Exit Sub
        
        Case -1, -2 'Left

            downItem.Selected = True
            scrollTimer.TmrStart Me, 50
            
        Case Else
            RaiseEvent ItemChanging(item, Cancel)

            If Cancel = False Then
                Set SelectedItem = item
                RaiseEvent ItemClick(item)
            End If

    End Select

End Sub

Friend Property Get StartItem() As Long
    StartItem = m_lStartItem
End Property

Friend Property Let StartItem(ByVal lStartItem As Long)

    If lStartItem < 1 Then lStartItem = 1
    If lStartItem <> m_lStartItem Then
        m_lStartItem = lStartItem
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
    LeftItemOffset = LeftOffset
End Property

Friend Function GetTotalWidth() As Long
    Dim item As TabItem

    For Each item In m_Items

        If item.Visible Then
            GetTotalWidth = GetTotalWidth + item.Width
        End If

    Next

End Function

Private Sub ResetState()
    Set downItem = Nothing
    Set m_selectedItem = Nothing
    LeftOffset = 0
    m_lStartItem = -1
    m_lLastItem = -1
    m_lVisibleCount = 0
End Sub

Friend Property Get VisibleTabCount() As Long
    Dim i As Long, item As TabItem

    If m_Items Is Nothing Then Exit Property

    For Each item In m_Items

        If item.Visible Then
            i = i + 1&
        End If

    Next
    
    VisibleTabCount = i
    'VisibleTabCount = m_lVisibleCount
End Property

Public Sub Move(ByVal lLeft As Long, _
                ByVal lTop As Long, _
                ByVal lWidth As Long, _
                ByVal lHeight As Long)
    
    m_lLeft = lLeft
    m_lTop = lTop
    m_lWidth = lWidth
    m_lheight = lHeight
    
    RecalculateLayout
    RequestRedraw
End Sub

Public Property Get Left() As Long
    Left = m_lLeft
End Property

Public Property Let Left(ByVal lLeft As Long)

    If m_lLeft <> lLeft Then
        m_lLeft = lLeft
        RecalculateLayout
        RequestRedraw
    End If

End Property

Public Property Get Top() As Long
    Top = m_lTop
End Property

Public Property Let Top(ByVal lTop As Long)

    If m_lTop <> lTop Then
        m_lTop = lTop
        RecalculateLayout
        RequestRedraw
    End If

End Property

Public Property Get Height() As Long
    Height = m_lheight
End Property

Public Property Let Height(ByVal lHeight As Long)

    If m_lheight <> lHeight Then
        m_lheight = lHeight
        RecalculateLayout
        RequestRedraw
    End If

End Property


Public Property Get ShowShortcut() As Boolean
    ShowShortcut = m_bShowShortcut
End Property

Public Property Let ShowShortcut(ByVal bShowShortcut As Boolean)
    m_bShowShortcut = bShowShortcut
End Property

Private Sub Scroll(ByVal direction As Long)
    '-1 left, -2 Right
    Dim Cancel As Boolean
    Cancel = True

    If direction = -1 Then
        If LeftOffset = 0 Then

            If m_lStartItem > 1 Then
                StartItem = m_lStartItem - 2&
                Cancel = False
            End If
            
        Else
            LeftOffset = 0
            Cancel = False
        End If

        If Cancel = False Then
            m_oPaintMgr.LayoutChanged
            RequestRedraw
        End If
    
    Else
        LastItem = m_lLastItem + 2&
        EnsureVisible m_lLastItem
        m_oPaintMgr.LayoutChanged
    End If
    
    If downItem.Enabled = False Then
        scrollTimer.TmrStop
        RequestRedraw
    End If

End Sub

Private Sub ITimer_Proc(ByVal lElapsedMS As Long, _
                        ByVal lTimerID As Long)

    If Not downItem Is Nothing Then
        Scroll downItem.Index
    Else
        scrollTimer.TmrStop
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
