VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabPaintManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LEFT_NAV = 1
Private Const RIGHT_NAV = 2

Public Enum CloseButtonPosition
    CloseButtonHidden = 1
    CloseButtonRightMost = 2
    CloseButtonOnActive = 4
End Enum

Private m_oTheme As ITabPainter
Private m_oOwner As TabManager
Private m_eShowCloseButton As CloseButtonPosition
Private m_Buttons(2) As TabItem 'Nav buttons and close button

Dim oDC As MemoryDC 'For string measurement
Dim bNavShown As Boolean

Private Sub Class_Initialize()
    Set oDC = New MemoryDC
    oDC.CreateDC 1, 1
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To 2
        Set m_Buttons(i) = Nothing
    Next
    Set m_oTheme = Nothing
    Set m_oOwner = Nothing
    Set oDC = Nothing
    
End Sub

'Public Property Get TabTheme() As ITabPainter
'    Set TabTheme = m_oTheme
'End Property

Public Property Set TabTheme(ByVal objTabTheme As ITabPainter)
    Set m_oTheme = objTabTheme
End Property

Public Property Let ShowCloseButton(ByVal eShowCloseButton As CloseButtonPosition)
    If m_eShowCloseButton <> eShowCloseButton Then
        m_eShowCloseButton = eShowCloseButton
        
        m_Buttons(0).Visible = (m_eShowCloseButton <> CloseButtonHidden)
        m_oOwner.RecalculateLayout
        
        If Not m_oOwner.SelectedItem Is Nothing Then
            m_oOwner.CaptionChanged m_oOwner.SelectedItem.Index
        End If
        m_oOwner.RequestRedraw

    End If
End Property

Public Property Get ShowCloseButton() As CloseButtonPosition
    ShowCloseButton = m_eShowCloseButton
End Property

Friend Function GetTabWidth(ByVal item As TabItem) As Long
    GetTabWidth = m_oTheme.CalculateTabWidth(item, oDC)
    If item.Selected And m_eShowCloseButton = CloseButtonOnActive Then
        GetTabWidth = GetTabWidth + 16&
    End If
End Function

Friend Function GetMaxWidth(ByVal lWidth As Long) As Long
    
    If m_eShowCloseButton = CloseButtonRightMost Then
        lWidth = lWidth - 16&
    End If
    bNavShown = m_oOwner.GetTotalWidth > lWidth
    If bNavShown Then
        lWidth = lWidth - 34&
    End If
    GetMaxWidth = lWidth
End Function

Friend Property Set Owner(objOwner As TabManager)
    Set m_oOwner = objOwner
    If m_oTheme Is Nothing Then
        Set TabTheme = New DefaultPainter
    End If
    Set m_Buttons(0) = New TabItem
    m_Buttons(0).Caption = "r"
    m_Buttons(0).Index = 0  'Close Button
    
    Set m_Buttons(1) = New TabItem
    m_Buttons(LEFT_NAV).Index = -1 'Left Nav Button
    m_Buttons(LEFT_NAV).Caption = "3"
    
    Set m_Buttons(2) = New TabItem
    m_Buttons(RIGHT_NAV).Index = -2 'Right Nav Button
    m_Buttons(RIGHT_NAV).Caption = "4"
End Property

Friend Sub DrawItem(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal itemType As TabItemType)
    Dim tR As RECT
    'pass a copy of the struct since we can not pass byval and the client may wish to modify the bounds
    LSet tR = bounds
    
    If Not item Is Nothing Then
        If item.RightMostItem Then
            
            ExcludeClipRect oDC.hdc, GetMaxWidth(oDC.Width), 0, oDC.Width, oDC.Height
            
        End If
    End If
    m_oTheme.DrawItem oDC, tR, item, itemType
    
    If itemType = DrawTabItem Then
        If m_eShowCloseButton = CloseButtonOnActive Then
            If item.Selected Then
                tR.Left = bounds.Right - 18&
                tR.Right = tR.Left + 14&
                tR.Top = bounds.Top + 5&
                tR.Bottom = bounds.Bottom - 3&
                m_Buttons(0).Width = tR.Left ' Borrow width property to store left coord for hittest
                m_oTheme.DrawItem oDC, tR, m_Buttons(0), DrawCloseButton
            End If
        End If
    
    ElseIf itemType = DrawBackGround Then
        If m_eShowCloseButton = CloseButtonRightMost Then
            tR.Left = bounds.Right - 16&
            tR.Right = tR.Left + 14&
            tR.Top = bounds.Top + 5&
            tR.Bottom = bounds.Bottom - 3&
            m_Buttons(0).Width = tR.Left ' Borrow width property to store left coord for hittest
            m_oTheme.DrawItem oDC, tR, m_Buttons(0), DrawCloseButton
        End If
        
        
        If m_Buttons(LEFT_NAV).Visible Then
            If m_eShowCloseButton = CloseButtonRightMost Then
                tR.Left = bounds.Right - 32&
            Else
                tR.Left = bounds.Right - 16&
            End If
            tR.Right = tR.Left + 14&
            tR.Top = bounds.Top + 5&
            tR.Bottom = bounds.Bottom - 3&
            m_Buttons(RIGHT_NAV).Width = tR.Left ' Borrow width property to store left coord for hittest
            
            m_oTheme.DrawItem oDC, tR, m_Buttons(RIGHT_NAV), DrawNavRightButton
            
            tR.Left = tR.Left - 16&
            tR.Right = tR.Left + 14&
            tR.Top = bounds.Top + 5&
            tR.Bottom = bounds.Bottom - 3&
            m_Buttons(LEFT_NAV).Width = tR.Left ' Borrow width property to store left coord for hittest
            m_oTheme.DrawItem oDC, tR, m_Buttons(LEFT_NAV), DrawNavLeftButton
        End If
    End If
End Sub

Friend Sub DrawShortcut(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal visibleIndex As Long)
    m_oTheme.DrawShortcut oDC, bounds, item, visibleIndex
End Sub

Friend Function ButtonHittest(ByVal x As Long, ByVal y As Long) As TabItem
    Dim i As Long, item As TabItem
    Dim tR As RECT
    'iterate only in visible boundary
    tR.Top = m_oOwner.Top + 3
    tR.Bottom = m_oOwner.Top + 19
    
    For i = 0 To 2
        Set item = m_Buttons(i)
        If item.Visible Then
            tR.Left = m_oOwner.Left + item.Width 'use width property as left coord
            tR.Right = tR.Left + 14&
            If PtInRect(tR, x, y) Then
                Set ButtonHittest = item
                Exit For
            End If
        End If
    Next
End Function

Friend Sub LayoutChanged()
    If bNavShown Then
        m_Buttons(RIGHT_NAV).Enabled = False
        
        'Check whether it's partially displayed
        If m_oOwner.LastItem > 1 Then
                If m_oOwner.IsOffScreenToTheRight(m_oOwner.LastItem, GetMaxWidth(m_oOwner.Width)) Then
                    m_Buttons(RIGHT_NAV).Enabled = True
                ElseIf m_oOwner.LastItem < m_oOwner.VisibleTabCount Then
                    m_Buttons(RIGHT_NAV).Enabled = True
                End If
    '        End If
        End If
        
        'Check whether it's partially displayed
    
        m_Buttons(LEFT_NAV).Enabled = (m_oOwner.StartItem > 1) Or (m_oOwner.LeftItemOffset <> 0)
    End If
    m_Buttons(LEFT_NAV).Visible = bNavShown
    m_Buttons(RIGHT_NAV).Visible = bNavShown
        
End Sub
