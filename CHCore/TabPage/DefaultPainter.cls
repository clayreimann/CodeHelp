VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultPainter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TEXT_MARGIN As Long = 6&

Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" _
  (ByVal hdc As Long, _
   pVertex As Any, _
   ByVal dwNumVertex As Long, _
   pMesh As Any, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long

Private Declare Function OleTranslateColor2 Lib "oleaut32.dll" Alias "OleTranslateColor" _
  (ByVal Clr As Long, _
  ByVal hPal As Long, _
  ByVal lpcolorref As Long) As Long

Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long

Dim hBrBack As Long
Dim hBrBackSelected As Long
Dim hBrHover As Long
Dim hBrPushed As Long

Dim hPen As Long
Dim hActivePen As Long
Dim hHoverPen As Long

Dim myFont As StdFont
Dim shortcutFont As IFont
Dim lClrHover As Long

Dim lClrHoverDark As Long

Implements ITabPainter

Private Sub Class_Initialize()
    
    
    lClrHover = AlphaBlend(vbHighlight, vbWindowBackground, 75&)
    lClrHoverDark = AlphaBlend(vbHighlight, vbWindowBackground, 128&)
    
    hBrBack = GetSysColorBrush(COLOR_3DHILIGHT)
    hBrBackSelected = GetSysColorBrush(COLOR_3DFACE)
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DSHADOW))
    hActivePen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DHILIGHT))
    hHoverPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_HIGHLIGHT))
    
    Set myFont = New StdFont
    myFont.Name = "Arial"
    myFont.Size = 8
    
    Set shortcutFont = New StdFont
    shortcutFont.Name = "Arial"
    shortcutFont.Size = 10
    shortcutFont.Bold = True
    
    hBrHover = CreateSolidBrush(lClrHover)
    hBrPushed = CreateSolidBrush(lClrHoverDark)
    
End Sub

Private Sub Class_Terminate()
    DeleteObject hPen
    DeleteObject hActivePen
    DeleteObject hHoverPen
    DeleteObject hBrHover
    DeleteObject hBrPushed
    Set myFont = Nothing
    Set shortcutFont = Nothing
End Sub

Private Function ITabPainter_CalculateTabWidth(ByVal item As TabItem, ByVal oDC As MemoryDC) As Long
    Dim tS As SIZEL
    Dim textCount As Long
    
    textCount = Len(item.Caption)
    If textCount > 0 Then
        
        myFont.Bold = item.Selected
        myFont.Size = 8
        myFont.Name = "Arial"
        Set oDC.Font = myFont
        A_GetTextExtentPoint oDC.hdc, item.Caption, textCount, tS
        ITabPainter_CalculateTabWidth = tS.cx + 2 * TEXT_MARGIN
        If item.LeftMostItem Then
            'add small margin for first tabitem
            ITabPainter_CalculateTabWidth = ITabPainter_CalculateTabWidth + TEXT_MARGIN
        End If
    End If
End Function

Private Sub ITabPainter_DrawItem(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal itemType As TabItemType)
    Dim lClr3DFace As Long
    Dim lClr3DHilite As Long
    
    lClr3DFace = GetSysColor(COLOR_3DFACE)
    lClr3DHilite = GetSysColor(COLOR_3DHILIGHT)
    SetBkMode oDC.hdc, TRANSPARENT
    
    Select Case itemType
        Case DrawBackGround
            DrawGradientFill oDC.hdc, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, _
                 lClr3DFace, lClr3DHilite, True
            
        Case DrawTabItem
            oDC.Pen = hPen
            
            If item.LeftMostItem Then
                'draw small margin
                bounds.Left = bounds.Left + 4&
            End If
            
            MoveToEx oDC.hdc, bounds.Left, bounds.Bottom, ByVal 0
            bounds.Top = bounds.Top + 3&
            bounds.Right = bounds.Right - 2&
            
            If item.Selected Then
                'FillRect oDC.hDC, bounds, hBrBack
                DrawGradientFill oDC.hdc, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, _
                    lClr3DHilite, lClr3DFace, True
                SetPixel oDC.hdc, bounds.Left, bounds.Top, lClr3DFace
                
                oDC.Pen = 0
                oDC.Pen = hActivePen
                LineTo oDC.hdc, bounds.Left, bounds.Top + 2&
                LineTo oDC.hdc, bounds.Left + 2&, bounds.Top
                LineTo oDC.hdc, bounds.Right - 2&, bounds.Top
                oDC.Pen = hPen
            Else
                FillRect oDC.hdc, bounds, hBrBackSelected
                LineTo oDC.hdc, bounds.Left, bounds.Top + 2&
                LineTo oDC.hdc, bounds.Left + 2&, bounds.Top
                LineTo oDC.hdc, bounds.Right - 2&, bounds.Top
            End If
            
            LineTo oDC.hdc, bounds.Right, bounds.Top + 2&
            LineTo oDC.hdc, bounds.Right, bounds.Bottom - 1&
            
            If item.Selected = False Then
                oDC.Pen = 0
                oDC.Pen = hActivePen
                MoveToEx oDC.hdc, bounds.Left, bounds.Bottom - 1&, ByVal 0
                LineTo oDC.hdc, bounds.Right + 2&, bounds.Bottom - 1&
            End If
            
            'Draw Text
            
            If Len(item.Caption) > 0 Then
                myFont.Bold = item.Selected
                myFont.Name = "Arial"
                myFont.Size = 8
                Set oDC.Font = myFont
                
                bounds.Left = bounds.Left + TEXT_MARGIN
                
                SetTextColor oDC.hdc, GetSysColor(COLOR_BTNTEXT)
                OffsetRect bounds, -1, -1
                A_DrawText oDC.hdc, item.Caption, -1, bounds, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
            End If
            
        Case DrawCloseButton, DrawNavLeftButton, DrawNavRightButton
        
            Dim symFont As IFont, hOldFont As Long
            Set symFont = myFont
            symFont.Name = "Marlett"
            symFont.Bold = False
            If itemType <> DrawCloseButton Then
                symFont.Size = 12
            Else
                symFont.Size = 8
            End If
            
            oDC.Brush = 0
            oDC.Pen = 0
            
            If item.Enabled Then
                If item.Selected Then
                    If item.Hovered Then
                        SetTextColor oDC.hdc, GetSysColor(COLOR_BTNHILIGHT)
                    Else
                        SetTextColor oDC.hdc, GetSysColor(COLOR_BTNTEXT)
                    End If
                Else
                    SetTextColor oDC.hdc, GetSysColor(COLOR_BTNTEXT)
                End If
                
                If item.Selected Then
                    If item.Index <> 0 Then
                        oDC.Pen = hHoverPen
                        oDC.Brush = hBrPushed
                        Rectangle oDC.hdc, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom
                    
                    ElseIf item.Hovered Then
                        oDC.Pen = hHoverPen
                        oDC.Brush = hBrPushed
                        Rectangle oDC.hdc, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom
                        
                    End If
                    
                ElseIf item.Hovered Then
                    oDC.Pen = hHoverPen
                    oDC.Brush = hBrHover
                    Rectangle oDC.hdc, bounds.Left, bounds.Top, bounds.Right, bounds.Bottom
                End If
                
            Else
                SetTextColor oDC.hdc, GetSysColor(COLOR_GRAYTEXT)
            End If
            
            
            hOldFont = SelectObject(oDC.hdc, symFont.hFont)
            OffsetRect bounds, 1, 0
            A_DrawText oDC.hdc, item.Caption, -1, bounds, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
            SelectObject oDC.hdc, hOldFont
            
    End Select
    oDC.Brush = 0
    oDC.Pen = 0
End Sub

Private Sub DrawGradientFill(ByVal hdc As Long, _
  ByVal lLeft As Long, ByVal lTop As Long, _
  ByVal lRight As Long, ByVal lBottom As Long, _
  ByVal dwColour1 As Long, _
  ByVal dwColour2 As Long, _
  Optional ByVal Vertical As Boolean = False)
    
  Dim vert(0 To 1) As TRIVERTEX
  Dim grRc As GRADIENT_RECT
  
  'Colour at upper-left corner
  With vert(0)
    .x = lLeft
    .y = lTop
    .Red = LongToSignedShort((dwColour1 And &HFF&) * 256&)
    .Green = LongToSignedShort(((dwColour1 And &HFF00&) \ &H100&) * 256&)
    .Blue = LongToSignedShort(((dwColour1 And &HFF0000) \ &H10000) * 256&)
  End With
  
  'Colour at bottom-right corner
  With vert(1)
     .x = lRight
     .y = lBottom
     .Red = LongToSignedShort((dwColour2 And &HFF&) * 256&)
     .Green = LongToSignedShort(((dwColour2 And &HFF00&) \ &H100&) * 256&)
     .Blue = LongToSignedShort(((dwColour2 And &HFF0000) \ &H10000) * 256&)
  End With
  
  With grRc
     .LowerRight = 0&
     .UpperLeft = 1&
  End With
     
  Call GradientFill(hdc, vert(0), 2&, grRc, 1&, Abs(Vertical))
End Sub

Private Function LongToSignedShort(dwUnsigned As Long) As Integer
    
  'convert from long to signed short
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
    
End Function

Private Function AlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  Dim clrFore         As RGBQUAD
  Dim clrBack         As RGBQUAD
  Dim lDif As Long
  On Error Resume Next
  OleTranslateColor2 clrFirst, 0, VarPtr(clrFore)
  OleTranslateColor2 clrSecond, 0, VarPtr(clrBack)
  If lAlpha < 0& Then lAlpha = 0&
  If lAlpha > 255& Then lAlpha = 255&
  With clrFore
    lDif = CLng(.rgbRed) - CLng(clrBack.rgbRed)
    .rgbRed = (lAlpha * lDif) \ 255 + clrBack.rgbRed
    lDif = CLng(.rgbGreen) - CLng(clrBack.rgbGreen)
    .rgbGreen = (lAlpha * lDif) \ 255 + clrBack.rgbGreen
    lDif = CLng(.rgbBlue) - CLng(clrBack.rgbBlue)
    .rgbBlue = (lAlpha * lDif) \ 255 + clrBack.rgbBlue
  End With
  CopyMemory ByVal VarPtr(AlphaBlend), ByVal VarPtr(clrFore), 4
End Function

Private Sub ITabPainter_DrawShortcut(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal visibleIndex As Long)
    Dim hOldFont As Long
    
    
    hOldFont = SelectObject(oDC.hdc, shortcutFont.hFont)
    SetTextColor oDC.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
    SetBkMode oDC.hdc, OPAQUE
    SetBkColor oDC.hdc, GetSysColor(COLOR_HIGHLIGHT)
    A_DrawText oDC.hdc, CStr(visibleIndex), -1, bounds, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
    SelectObject oDC.hdc, hOldFont
    
End Sub

