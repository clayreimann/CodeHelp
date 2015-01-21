VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Modified class from vbaccelerator.com
' ---------------------------------------------------------------------------
' vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' =========================================================

Private Type PicBmp
    Size As Long
    lType As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function CreateDCAsNull _
                Lib "gdi32" _
                Alias "CreateDCA" (ByVal lpDriverName As String, _
                                   lpDeviceName As Any, _
                                   lpOutput As Any, _
                                   lpInitData As Any) As Long
Private Declare Function OleCreatePictureIndirect _
                Lib "olepro32.dll" (PicDesc As PicBmp, _
                                    RefIID As GUID, _
                                    ByVal fPictureOwnsHandle As Long, _
                                    IPic As IPicture) As Long
Private Declare Function GetObjectAPI _
                Lib "gdi32.dll" _
                Alias "GetObjectA" (ByVal hObject As Long, _
                                    ByVal nCount As Long, _
                                    ByRef lpObject As Any) As Long

Private m_hDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lheight As Long
Private m_hFont As Long
Private m_hOldFont As Long
Private m_MemoryBrush As Long
Private m_OriginalBrush As Long
Private m_MemoryPen As Long
Private m_OriginalPen As Long
Private m_Created As Boolean
Private m_lClrText As Long
Private m_bVertical As Boolean
Private m_IsBold As Boolean

Public Sub CreateFromPicture(sPic As IPicture)
    Dim tB As BITMAP
    Dim lhDCC As Long, lHDC As Long
    Dim lhBmpOld As Long
    GetObjectAPI sPic.Handle, Len(tB), tB
    lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    lHDC = CreateCompatibleDC(lhDCC)
    lhBmpOld = SelectObject(lHDC, sPic.Handle)
    BitBlt hDC, 0, 0, tB.bmWidth, tB.bmHeight, lHDC, 0, 0, vbSrcCopy
    SelectObject lHDC, lhBmpOld
    DeleteDC ByVal lHDC
    DeleteDC ByVal lhDCC
    m_Created = True
End Sub

Public Sub CreateDC(ByVal Width As Long, _
                    ByVal Height As Long)
    pDestroy
    pCreate Width, Height
    m_Created = True
End Sub

'Public Sub CreateFromDC(ByVal hOrgDC As Long, Width As Long, Height As Long)
'  Dim tm              As TEXTMETRIC
'  Dim sFaceName       As String * 80
'  Dim oFnt            As StdFont
'
'  pDestroy
'  m_hDC = CreateCompatibleDC(hOrgDC)
'  If hOrgDC <> 0 Then
'    m_hBmp = CreateCompatibleBitmap(hOrgDC, Width, Height)
'  Else
'    m_hBmp = CreateCompatibleBitmap(m_hDC, Width, Height)
'  End If
'  m_hBmpOld = SelectObject(m_hDC, m_hBmp)
'  If m_hBmpOld = 0 Then
'     pDestroy
'     Err.Raise vbObjectError + 512, "MEM_DC", "Can not create CompatibleDC"
'  Else
'    m_lWidth = Width
'    m_lheight = Height
'    If hOrgDC <> 0& Then
'      If hOrgDC <> 0& Then
'        GetTextMetrics hOrgDC, tm
'        GetTextFace hOrgDC, 79, sFaceName
'      Else
'        GetTextMetrics m_hDC, tm
'        GetTextFace m_hDC, 79, sFaceName
'      End If
'
'      Set oFnt = New StdFont
'      With oFnt
'          .Name = sFaceName 'StrConv(sFaceName, vbUnicode)
'          .Bold = (tm.tmWeight >= FW_NORMAL)
'          .Charset = tm.tmCharSet
'          .Italic = (tm.tmItalic <> 0)
'          .Strikethrough = (tm.tmStruckOut <> 0)
'          .Underline = (tm.tmUnderlined <> 0)
'          .Weight = tm.tmWeight
'          .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
'      End With
'      Set Font = oFnt
'      Set oFnt = Nothing
'    End If
'    m_Created = True
'  End If
'End Sub

Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With Pic
        .Size = Len(Pic) ' Length of structure
        .lType = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Property Get hBitmap() As Long

    If m_Created = False Then Exit Property
    hBitmap = m_hBmp
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Get Height() As Long
    Height = m_lheight
End Property

Public Property Get FontName() As String

    If m_Created = False Then Exit Property
    Dim lRetval As Long
    FontName = String$(80&, vbNullChar)
    lRetval = A_GetTextFace(m_hDC, 81&, FontName)
  
    FontName = Left$(FontName, lRetval)
  
End Property

Public Property Get fontSize() As Long

    If m_Created = False Then Exit Property
    Dim ttM As A_TEXTMETRIC
    A_GetTextMetrics m_hDC, ttM
    fontSize = (ttM.tmHeight - ttM.tmInternalLeading) * 72& / ttM.tmDigitizedAspectY
End Property

Public Property Get FontColor() As Long

    If m_Created = False Then Exit Property
    FontColor = m_lClrText
End Property

Public Property Let FontColor(ByVal Value As Long)

    If m_Created = False Then Exit Property
    If Value <> m_lClrText Then
        m_lClrText = Value
        SetTextColor m_hDC, m_lClrText
    End If

End Property

Property Get FontBold() As Boolean
    FontBold = m_IsBold
End Property

Public Property Get Brush() As Long
    Brush = m_MemoryBrush
End Property

Public Property Let Brush(ByVal lValue As Long)

    '--- state check
    If Not m_Created Then Exit Property
    If lValue = m_MemoryBrush Then Exit Property
  
    '--- set new brush (and save original)
    If m_OriginalBrush <> 0 Then
        SelectObject m_hDC, m_OriginalBrush
        m_OriginalBrush = 0&
    End If

    m_MemoryBrush = lValue

    If lValue <> 0& Then
        m_OriginalBrush = SelectObject(m_hDC, m_MemoryBrush)
    End If

End Property

'Purpose: Returns or sets the currently selected GDI pen (HPEN) for the device context of the <b>memoryDC</b> object.
Public Property Get Pen() As Long
    Pen = m_MemoryPen
End Property

Public Property Let Pen(ByVal lValue As Long)

    '--- state check
    If Not m_Created Then Exit Property
    If lValue = m_MemoryPen Then Exit Property
  
    '--- set new brush (and save original)
    If m_OriginalPen <> 0 Then
        SelectObject m_hDC, m_OriginalPen
        m_OriginalPen = 0&
    End If

    m_MemoryPen = lValue

    If lValue <> 0& Then
        m_OriginalPen = SelectObject(m_hDC, m_MemoryPen)
    End If

End Property

Public Property Get Font() As StdFont
    Dim tm              As A_TEXTMETRIC
  
    Dim oFont As StdFont
  
    '--- state check
    If Not m_Created Then
        Exit Property
    End If

    On Error Resume Next
  
    A_GetTextMetrics m_hDC, tm
  
    Set oFont = New StdFont

    With oFont
        .Name = FontName
        .Bold = (tm.tmWeight > FW_NORMAL)
        .Charset = tm.tmCharSet
        .Italic = (tm.tmItalic <> 0)
        .Strikethrough = (tm.tmStruckOut <> 0)
        .Underline = (tm.tmUnderlined <> 0)
        .Weight = tm.tmWeight
        .Size = (tm.tmHeight - tm.tmInternalLeading) * 72 / tm.tmDigitizedAspectY
    End With

    Set Font = oFont
    Set oFont = Nothing
End Property

Public Property Set Font(ByVal oValue As StdFont)
    Dim tFont           As A_LOGFONT

    '--- state check
    If Not m_Created Then Exit Property
  
    On Error Resume Next
    m_IsBold = oValue.Bold

    With tFont
        StrToBytes .lfFaceName, oValue.Name
        .lfCharSet = oValue.Charset
        .lfItalic = (-oValue.Italic)
        .lfStrikeOut = (-oValue.Strikethrough)
        .lfUnderline = (-oValue.Underline)
        .lfWeight = oValue.Weight
        .lfHeight = -MulDiv((oValue.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)

        If m_bVertical Then
            .lfEscapement = 900
        End If

    End With

    If m_hOldFont <> 0& Then
        SelectObject m_hDC, m_hOldFont
        m_hOldFont = 0&
    End If
  
    If m_hFont <> 0& Then
        DeleteObject m_hFont
        m_hFont = 0
    End If
  
    m_hFont = A_CreateFontIndirect(tFont)

    If m_hFont <> 0 Then m_hOldFont = SelectObject(m_hDC, m_hFont)
End Property

Public Property Let FontVertical(ByVal bVertical As Boolean)
    m_bVertical = bVertical
End Property

Private Sub pCreate(ByVal lW As Long, _
                    ByVal lH As Long)
    Dim lHDC As Long
    pDestroy
    lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_hDC = CreateCompatibleDC(lHDC)
    m_hBmp = CreateCompatibleBitmap(lHDC, lW, lH)
    m_hBmpOld = SelectObject(m_hDC, m_hBmp)

    If m_hBmpOld = 0 Then
        pDestroy
    Else
        m_lWidth = lW
        m_lheight = lH
    End If

    DeleteDC ByVal lHDC
    m_Created = True
End Sub

Private Sub pDestroy()
  
    If m_OriginalPen <> 0 Then
        Call SelectObject(ByVal m_hDC, ByVal m_OriginalPen)
        m_OriginalPen = 0
    End If

    If m_OriginalBrush <> 0 Then
        Call SelectObject(ByVal m_hDC, ByVal m_OriginalBrush)
        m_OriginalBrush = 0
    End If
  
    If m_MemoryPen <> 0 Then
        Call DeleteObject(ByVal m_MemoryPen)
        m_MemoryPen = 0
    End If

    If m_MemoryBrush <> 0 Then
        Call DeleteObject(ByVal m_MemoryBrush)
        m_MemoryBrush = 0
    End If
  
    If m_hBmpOld <> 0 Then
        SelectObject ByVal m_hDC, ByVal m_hBmpOld
        m_hBmpOld = 0
    End If
  
    If m_hOldFont <> 0& Then
        SelectObject ByVal m_hDC, ByVal m_hOldFont
        m_hOldFont = 0&
    End If
  
    If m_hFont <> 0 Then
        DeleteObject ByVal m_hFont
        m_hFont = 0
    End If
  
    If m_hBmp <> 0 Then
        DeleteObject ByVal m_hBmp
        m_hBmp = 0
        m_lWidth = 0
        m_lheight = 0
    End If
  
    If Not m_hDC = 0 Then
        DeleteDC ByVal m_hDC
        m_hDC = 0
    End If

    m_Created = False
End Sub

Private Sub Class_Terminate()
    pDestroy
End Sub

Private Sub StrToBytes(ab() As Byte, _
                       s As String)
    Dim cab As Long
    ' Copy to existing array, padding or truncating if necessary
    cab = UBound(ab) - LBound(ab) + 1

    If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
    CopyMemory ab(LBound(ab)), ByVal s, cab
End Sub

