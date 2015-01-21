VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Tab item class for Tab Pages and buttons

Private m_sCaption As String
Private m_objIcon As IPictureDisp
Private m_bVisible As Boolean
Private m_tabMgr As TabManager
Private m_lIndex As Long
Private m_lWidth As Long
Private m_IsRightMostItem As Boolean
Private m_IsLeftMostItem As Boolean
Private m_lState As Long
Private m_hWnd As Variant

Private Sub Class_Terminate()
    Set m_objIcon = Nothing
    Set m_tabMgr = Nothing
    If IsObject(m_hWnd) Then
        Set m_hWnd = Nothing
    Else
        m_hWnd = vbEmpty
    End If
End Sub

Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Let Caption(ByVal sCaption As String)
    If m_sCaption <> sCaption Then
        m_sCaption = sCaption
        If Not m_tabMgr Is Nothing Then
            m_tabMgr.CaptionChanged m_lIndex
        End If
    End If
End Property

Public Property Get Icon() As IPictureDisp
    Set Icon = m_objIcon
End Property

Public Property Set Icon(objIcon As IPictureDisp)
    Set m_objIcon = objIcon
End Property

Public Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    If m_bVisible <> bVisible Then
        m_bVisible = bVisible
        If Not m_tabMgr Is Nothing Then
            m_tabMgr.VisibleChanged m_lIndex
        End If
    End If
End Property

Public Property Get Selected() As Boolean
    Selected = ((m_lState And ODS_SELECTED) = ODS_SELECTED)
End Property

Public Property Let Selected(ByVal bSelected As Boolean)
    If Me.Selected <> bSelected Then
        
        If m_bVisible Then
            If bSelected Then
                m_lState = (m_lState Or ODS_SELECTED)
            Else
                m_lState = (m_lState And Not ODS_SELECTED)
            End If
            
            If bSelected Then
                If Not m_tabMgr Is Nothing Then
                    Set m_tabMgr.SelectedItem = Me
                End If
            End If
        End If
    End If
End Property

Friend Sub SetActive(ByVal active As Boolean)
    If active Then
        m_lState = (m_lState Or ODS_SELECTED)
    Else
        m_lState = (m_lState And Not ODS_SELECTED)
    End If
End Sub

Public Property Get Index() As Long
    Index = m_lIndex
End Property

Public Property Let Index(ByVal lIndex As Long)
    m_lIndex = lIndex
End Property

Friend Property Set Owner(ByVal Value As TabManager)
    Set m_tabMgr = Value
End Property

Friend Property Get Width() As Long
    Width = m_lWidth
End Property

Friend Property Let Width(ByVal lWidth As Long)
    m_lWidth = lWidth
End Property

Friend Property Get RightMostItem() As Boolean
    RightMostItem = m_IsRightMostItem
End Property

Friend Property Let RightMostItem(ByVal bRightMostItem As Boolean)
    m_IsRightMostItem = bRightMostItem
End Property

Friend Property Get LeftMostItem() As Boolean
    LeftMostItem = m_IsLeftMostItem
End Property

Friend Property Let LeftMostItem(ByVal Value As Boolean)
    m_IsLeftMostItem = Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = ((m_lState And ODS_DISABLED) <> ODS_DISABLED)
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    If bEnabled Then
        m_lState = (m_lState And Not ODS_DISABLED)
    Else
        m_lState = (m_lState Or ODS_DISABLED)
    End If
End Property

Public Property Get Hovered() As Boolean
    Hovered = ((m_lState And ODS_HOTLIGHT) = ODS_HOTLIGHT)
End Property

Friend Property Let Hovered(ByVal bHovered As Boolean)
    If bHovered Then
        If Enabled Then
            m_lState = (m_lState Or ODS_HOTLIGHT)
        End If
    Else
        m_lState = (m_lState And Not ODS_HOTLIGHT)
    End If
End Property

Public Property Get Pushed() As Boolean
    Hovered = ((m_lState And ODS_CHECKED) = ODS_CHECKED)
End Property

Friend Property Let Pushed(ByVal Value As Boolean)
    If Value Then
        m_lState = (m_lState Or ODS_CHECKED)
    Else
        m_lState = (m_lState And Not ODS_CHECKED)
    End If
End Property

Public Property Get ChildhWnd() As Variant
    ChildhWnd = m_hWnd
End Property

Public Property Let ChildhWnd(ByVal hWnd As Variant)
    m_hWnd = hWnd
End Property

