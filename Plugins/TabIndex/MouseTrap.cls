VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SS_OWNERDRAW As Long = &HD&
Private Const TEXT_FORMAT As Long = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER

Private Declare Function LoadCursorLong Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private m_VBForm As VBForm
Private m_ParentWindow As Window

Private m_hMDIClient As Long
Private m_hDesigner As Long
Private m_hForm As Long

Private m_tR As RECT
'Private m_DC As MemoryDC
Private m_highestIdx As Integer
Private m_currentIdx As Integer
Private m_hWnd As Long
Private m_Font As IFont
'Private m_CtlOver As Object

Dim tOffset As POINTAPI
Dim hCursor As Long
Event Closed()


'Private Property Set HoveredControl(ByVal value As Object)
'
'    If Not (value Is m_CtlOver) Then
''        If Not m_CtlOver Is Nothing Then
''
''        End If
'        Set m_CtlOver = value
'        PostMessage m_hWnd, WM_PAINT, ByVal 0, ByVal 0
'        'RedrawWindow m_hWnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ERASE
'    End If
'
'End Property

Public Property Set FormDesigner(ByVal objFormDesigner As VBForm)
    Set m_VBForm = objFormDesigner
End Property

Public Property Set ParentWindow(ByVal objParentWindow As Window)
    Dim hChild As Long
    
    'Set m_DC = New MemoryDC
    Set m_ParentWindow = objParentWindow
    
    m_hMDIClient = A_FindWindowEx(m_ParentWindow.VBE.MainWindow.hwnd, 0, "MDIClient", vbNullString)
    m_hDesigner = A_FindWindowEx(m_hMDIClient, 0, "DesignerWindow", m_ParentWindow.Caption & vbNullChar)
    
    If m_hDesigner = 0 Then Err.Raise vbObjectError, "Create", "Can not find Designer Window"
    hChild = GetWindow(m_hDesigner, GW_CHILD)
    
    Do While hChild <> 0
        If Left$(GetWinText(hChild, True), 7) = "Thunder" Then
            m_hForm = hChild
            Exit Do
        End If
        hChild = GetWindow(hChild, GW_HWNDNEXT)
    Loop
    
    If m_hForm <> 0 Then
        CreateTrap
    End If
End Property

Public Property Get hWndTrap() As Long
    hWndTrap = m_hWnd
End Property

Public Property Get hWndFormDesigner() As Long
    hWndFormDesigner = m_hDesigner
End Property

Public Property Get hWndForm() As Long
    hWndForm = m_hForm
End Property

Private Sub CreateTrap()
    Dim tCS As CREATESTRUCT
    
    m_hWnd = A_CreateWindowEx(WS_EX_TRANSPARENT, "Static", "", WS_CHILD Or SS_OWNERDRAW, m_tR.Left, m_tR.Top, _
        m_tR.Right - m_tR.Left, m_tR.Bottom - m_tR.Top, m_hDesigner, 0, App.hInstance, tCS)
    
    If m_hWnd <> 0 Then
        Set m_Font = New StdFont
        m_Font.Name = "Arial"
        m_Font.Size = 8
        m_Font.Bold = True
        hCursor = LoadCursorLong(0, IDC_CROSS)
    End If
End Sub

Sub ShowTabOrder()
    
    SynchContents
    SetWindowPos m_hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOSENDCHANGING Or SWP_NOACTIVATE
    m_highestIdx = m_VBForm.VBControls.Count - 1&
    DoEvents 'Allow window to show before redraw
    RedrawWindow m_hWnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW
End Sub

Private Sub OnMouseDown(Button As MouseButtonConstants, ByVal x As Long, ByVal y As Long)
    Dim vbCtl As Object
    
    If Button = vbLeftButton Then
        Set vbCtl = FindControlOnPoint(x, y)
        If Not vbCtl Is Nothing Then
            vbCtl.TabIndex = m_currentIdx
            RedrawWindow m_hWnd, ByVal 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW
            m_currentIdx = m_currentIdx + 1
            If m_currentIdx > m_highestIdx Then m_currentIdx = 0
        End If
    Else
        m_currentIdx = 0
    End If
    
End Sub

Private Sub OnPaint()
    
    Dim parentForm As Object
        
    Dim tR As RECT
    Dim hDC As Long
    Dim hOldBr As Long, hOldPen As Long
    Dim hBr As Long, hPen As Long
    Dim hOldFont As Long
    Dim lPenColor As Long
    
    tOffset.x = 0
    tOffset.y = 0
    
    hDC = GetDC(m_hWnd)
    SetBkMode hDC, TRANSPARENT
    'BitBlt hdc, 0, 0, m_DC.Width, m_DC.Height, m_DC.hdc, 0, 0, vbSrcCopy
    
    'Draw Border
    lPenColor = GetSysColor(COLOR_HIGHLIGHT)
    hPen = CreatePen(PS_SOLID, 2, lPenColor)
    hOldPen = SelectObject(hDC, hPen)
    
    'Get the form coordinate
    GetClientRect m_hForm, tR
    MapWindowPoints m_hForm, m_hWnd, tOffset, 1
    OffsetRect tR, tOffset.x, tOffset.y
    
    'Draw box around the form, as indicator we're in tab order mode
    MoveToExNoRet hDC, tR.Left, tR.Top, ByVal 0
    LineTo hDC, tR.Right, tR.Top
    LineTo hDC, tR.Right, tR.Bottom
    LineTo hDC, tR.Left, tR.Bottom
    LineTo hDC, tR.Left, tR.Top - 1&
    
    SelectObject hDC, hOldPen
    DeleteObject hPen

    'Draw quick tip
    Dim tipFont As IFont
    
    Set tipFont = New StdFont
    tipFont.Name = "Arial"
    tipFont.Size = 8
    
    hBr = GetSysColorBrush(COLOR_INFOBK)
    hOldBr = SelectObject(hDC, hBr)
    
    hOldFont = SelectObject(hDC, tipFont.hFont)
    'draw tip in caption area
    OffsetRect tR, 0, -20
    
    Rectangle hDC, tR.Left - 1&, tR.Top - 1&, tR.Right + 1&, tR.Top + 18
    A_DrawText hDC, "  Reset click index to 0 with right click", -1, tR, DT_LEFT
    
    SelectObject hDC, hOldFont
    SelectObject hDC, hOldBr
    Set tipFont = Nothing
    
    OffsetRect tR, 0, 20
    
    'Draw tab indexes
    hOldFont = SelectObject(hDC, m_Font.hFont)
    hBr = GetSysColorBrush(COLOR_HIGHLIGHT)
    hOldBr = SelectObject(hDC, hBr)
    
    lPenColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
    hPen = CreatePen(PS_SOLID, 1, lPenColor)
    hOldPen = SelectObject(hDC, hPen)
    
    SetTextColor hDC, lPenColor
    
    If m_VBForm.ContainedVBControls.Count > 0 Then
        
        DrawContainedControls hDC, parentForm, Nothing, m_VBForm.ContainedVBControls
    
    End If
    'Clean up GDI resources
    SelectObject hDC, hOldFont
    SelectObject hDC, hOldBr
    SelectObject hDC, hOldPen
    DeleteObject hPen
    ReleaseDC m_hWnd, hDC
End Sub

Private Sub DrawContainedControls(ByVal hDC As Long, ByRef parentForm As Object, ByVal parentCtl As Object, ByVal container As ContainedVBControls)
    
    On Error Resume Next
    
    Dim vbCtl As VBControl, ctl As Object
    Dim tR As RECT
    Dim iIdx As Integer
    Dim tPt As POINTAPI
    
    For Each vbCtl In container
        Set ctl = vbCtl.ControlObject
        
        If parentForm Is Nothing Then 'set the root form object
            Set parentForm = ctl.Parent
        End If
        
        iIdx = ctl.TabIndex
        
'        tPt.x = ctl.Left
'        tPt.y = ctl.Top
'
'        If Not (parentCtl Is Nothing) Then
'            If parentForm.ScaleMode <> parentCtl.ScaleMode Then
'                tPt.x = parentForm.ScaleX(tPt.x, parentCtl.ScaleMode, parentForm.ScaleMode)
'                tPt.y = parentForm.ScaleY(tPt.y, parentCtl.ScaleMode, parentForm.ScaleMode)
'            End If
'
'            tPt.x = tPt.x + parentCtl.Left
'            tPt.y = tPt.y + parentCtl.Top
'        End If
'
'        tPt.x = parentForm.ScaleX(tPt.x, parentForm.ScaleMode, vbPixels)
'        tPt.y = parentForm.ScaleY(tPt.y, parentForm.ScaleMode, vbPixels)
        
        tR = GetChildRect(ctl, parentForm)
        tR.Left = tR.Left + tOffset.x
        tR.Top = tR.Top + tOffset.y
        tR.Right = tR.Left + 20&
        tR.Bottom = tR.Top + 20&
        
        Rectangle hDC, tR.Left, tR.Top, tR.Right, tR.Bottom
        
        A_DrawText hDC, CStr(iIdx), -1, tR, TEXT_FORMAT
        
        'recursively draw controls in container
        If vbCtl.ContainedVBControls.Count > 0 Then
            DrawContainedControls hDC, parentForm, ctl, vbCtl.ContainedVBControls
        End If
    Next

End Sub

Private Sub DestroyTrap()
    If m_hWnd <> 0 Then
        Set m_VBForm = Nothing
        Set m_ParentWindow = Nothing
                
        A_PostMessage m_hWnd, WM_CLOSE, ByVal 0, ByVal 0
        A_PostMessage m_hMDIClient, WM_PAINT, ByVal 0, ByVal 0
        A_PostMessage m_hDesigner, WM_PAINT, ByVal 0, ByVal 0
        Set m_Font = Nothing
        
    End If
End Sub

Private Sub SynchContents()
        
    GetWindowRect m_hForm, m_tR
    InflateRect m_tR, 8, 8
    'convert to client coordinate
    MapWindowPoints 0, m_hDesigner, m_tR, 2

    MoveWindow m_hWnd, m_tR.Left, m_tR.Top, m_tR.Right - m_tR.Left, m_tR.Bottom - m_tR.Top, 1
    
End Sub

Private Function GetWinText(hwnd As Long, Optional className As Boolean = False) As String
    'some static vars to speed up things, this func will be called many times
    Static sBuffer As String * 128&
    Static textLength As Long
  
    If className Then
        textLength = A_GetClassName(hwnd, sBuffer, 129&)
    Else
        textLength = A_GetWindowText(hwnd, sBuffer, 129&)
    End If
  
    If textLength > 0 Then
        GetWinText = Left$(sBuffer, textLength)
    End If

End Function

Private Function FindControlOnPoint(ByVal x As Long, ByVal y As Long) As Object

    Dim parentForm As Object
    
    Set FindControlOnPoint = HitTestContainer(x, y, parentForm, Nothing, m_VBForm.ContainedVBControls)
    
End Function

Private Function HitTestContainer(ByVal x As Long, ByVal y As Long, _
    ByRef parentForm As Object, ByVal parentCtl As Object, _
    ByVal container As ContainedVBControls) As Object
    
    Dim vbCtl As VBControl, ctl As Object
    Dim result As Object
    Dim tR As RECT
    
    For Each vbCtl In container
        
        Set ctl = vbCtl.ControlObject
        
        If parentForm Is Nothing Then
            Set parentForm = ctl.Parent
        End If
        
        If vbCtl.ContainedVBControls.Count > 0 Then
            Set result = HitTestContainer(x, y, parentForm, ctl, vbCtl.ContainedVBControls)
        End If
        
        If result Is Nothing Then
            tR = GetChildRect(ctl, parentForm)
            
            If PtInRect(tR, x - tOffset.x, y - tOffset.y) Then
                Set result = ctl
                Exit For
            End If
        Else
            Exit For
        End If
        
    Next
    
    Set HitTestContainer = result
End Function

Private Sub Class_Terminate()
    DestroyTrap
End Sub

Sub TrapProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    
    Select Case hwnd
        Case m_hForm
            RaiseEvent Closed
            
        Case m_hWnd
            Select Case uMsg
                Case WM_NCHITTEST
                    lReturn = HTCLIENT
                    bHandled = True
            
                Case WM_LBUTTONDOWN
                    OnMouseDown vbLeftButton, LoWord(lParam), HiWord(lParam)
                    
                Case WM_RBUTTONDOWN
                    OnMouseDown vbRightButton, LoWord(lParam), HiWord(lParam)
                    
                Case WM_SYSCOMMAND
                    If wParam = SC_CLOSE Then
                        bHandled = True
                        A_PostMessage m_hWnd, WM_CLOSE, ByVal 0, ByVal 0
                        A_PostMessage m_hForm, uMsg, ByVal wParam, ByVal lParam
                    End If
                    
                Case WM_SETCURSOR
                    SetCursor hCursor
                    bHandled = True
                    
'                Case WM_MOUSEMOVE
'                    Set HoveredControl = FindControlOnPoint(LoWord(lParam), HiWord(lParam))
                    
            End Select
            
            
        Case m_hDesigner
            OnPaint
    End Select
End Sub

Private Function LoWord(lDWord As Long) As Integer

    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If

End Function

Private Function HiWord(lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Private Function GetChildRect(ctl As Object, ByVal parentForm As Object) As RECT
    Dim tR As RECT
    Dim parentCtl As Object, parentOrg As Object
    Dim tCtlOffset As POINTAPI
    Dim lastParent As Object
    
    On Error Resume Next

    Set parentCtl = ctl.container
    Set parentOrg = parentCtl

    Do While Not (parentCtl Is Nothing)
        If Not lastParent Is Nothing Then
            tR.Left = tR.Left + GetScaledMeasure(tCtlOffset.x, lastParent, parentForm, True)
            tR.Top = tR.Top + GetScaledMeasure(tCtlOffset.y, lastParent, parentForm, False)
        End If
        
        If Not parentCtl Is parentForm Then
            tCtlOffset.x = parentCtl.Left
            tCtlOffset.y = parentCtl.Top
        End If
        Set lastParent = parentCtl
        Err.Clear
        Set parentCtl = parentCtl.container
        If Err.Number <> 0 Then
            Set parentCtl = Nothing
        End If
    Loop
        
    tR.Left = tR.Left + GetScaledMeasure(ctl.Left, parentOrg, parentForm, True)
    tR.Top = tR.Top + GetScaledMeasure(ctl.Top, parentOrg, parentForm, False)
    tR.Right = tR.Left + GetScaledMeasure(ctl.Width, parentOrg, parentForm, True)
    tR.Bottom = tR.Top + GetScaledMeasure(ctl.Height, parentOrg, parentForm, False)
    
    GetChildRect = tR
End Function


Private Function GetScaledMeasure(ByVal value As Long, ByVal container As Object, ByVal parentForm As Object, ByVal isX As Boolean) As Long
    On Error Resume Next
    
    If isX Then
        GetScaledMeasure = parentForm.ScaleX(value, container.ScaleMode, vbPixels)
        If Err.Number <> 0 Then
            GetScaledMeasure = parentForm.ScaleX(value, vbTwips, vbPixels)
        End If
    Else
        GetScaledMeasure = parentForm.ScaleY(value, container.ScaleMode, vbPixels)
        If Err.Number <> 0 Then
            GetScaledMeasure = parentForm.ScaleY(value, vbTwips, vbPixels)
        End If
    End If
End Function
