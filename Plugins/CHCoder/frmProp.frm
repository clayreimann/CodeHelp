VERSION 5.00
Begin VB.Form frmProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CodeHelp Coder Options"
   ClientHeight    =   6435
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtKey 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   135
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1650
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarker 
      Caption         =   "Insert Marker"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   4995
      Width           =   1215
   End
   Begin VB.ComboBox cboMarker 
      Height          =   330
      Left            =   2430
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5025
      Width           =   3615
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   1710
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   570
      Width           =   6945
   End
   Begin VB.ListBox lstKey 
      Height          =   4680
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Double Click to edit"
      Top             =   570
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7470
      TabIndex        =   9
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Marker:"
      Height          =   210
      Index           =   2
      Left            =   1710
      TabIndex        =   4
      Top             =   5025
      Width           =   540
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Snippet:"
      Height          =   210
      Index           =   1
      Left            =   1755
      TabIndex        =   2
      Top             =   300
      Width           =   585
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Shortcut:"
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   300
      Width           =   660
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLOSE_TAG As String = "</>"
Dim editMode As Boolean

Private m_Parent As CHCoder.Loader
Private m_Templates As Recordset
Private m_Markers As Recordset

Public Sub Initalize(templates As Recordset, markers As Recordset, parent As CHCoder.Loader)
    Set m_Templates = templates
    Set m_Markers = markers
    Set m_Parent = parent
    
    Call SetupUI
End Sub

Private Sub SetupUI()
    With m_Templates
        If .RecordCount > 0 Then
            Call .MoveFirst
            Do While Not .EOF
                Call lstKey.AddItem(.Fields("Key").Value)
                Call .MoveNext
            Loop
        End If
    End With
    
    With m_Markers
        If .RecordCount > 0 Then
            Call .MoveFirst
            Do While Not .EOF
                Call cboMarker.AddItem(.Fields(0).Value & "   " & .Fields(1).Value)
                Call .MoveNext
            Loop
        End If
    End With
End Sub


Private Sub cmdDelete_Click()
    Call m_Templates.MoveFirst
    Call m_Templates.Find(KeyForListItem)
    Call m_Templates.Delete(adAffectCurrent)
    Call m_Templates.MoveLast 'if we don't MoveLast then everything explodes awfully

    Call lstKey.RemoveItem(lstKey.ListIndex)
    txtCode.Text = ""
    
    Exit Sub
ErrHandler:
    Call MsgBox(Err.Description)
End Sub

Private Sub cmdMarker_Click()
    If cboMarker.ListIndex > -1 Then
        txtCode.SelText = Split(cboMarker.Text, "   ")(0) & CLOSE_TAG
    End If
End Sub

Private Sub cmdNew_Click()
    Dim idx As Long
    
    With txtKey
        lstKey.ListIndex = -1
        .Text = "New"

        idx = lstKey.TopIndex
        idx = lstKey.ListCount - idx
        
        Call .Move(lstKey.Left + 30, lstKey.Top + (idx * .Height) + 30, lstKey.Width - 60)

        .Visible = True
        .SelStart = 0
        .SelLength = 3
        Call .SetFocus
    End With
    
    editMode = False
End Sub

Private Sub cmdCancel_Click()
    Call Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ERR_HANDLER
    Call ValidateData
    Call m_Parent.SaveData
    Call Me.Hide
    
    Exit Sub
ERR_HANDLER:
    MsgBox Err.Description
End Sub

Private Sub lstKey_Click()
    If m_Templates.RecordCount > 0 Then
        Call m_Templates.MoveFirst
        Call m_Templates.Find(KeyForListItem)
        If Not m_Templates.EOF Then
            txtCode.Text = Trim$(m_Templates.Fields("Code").Value & " ")
        Else
            txtCode.Text = ""
        End If
    End If
End Sub

Private Sub lstKey_DblClick()
    Dim idx As Long
    idx = lstKey.TopIndex
    idx = lstKey.ListIndex - idx
    Call txtKey.Move(lstKey.Left + 30, lstKey.Top + (idx * txtKey.Height) + 30, lstKey.Width - 60)
    
    txtKey.Text = lstKey.Text
    txtKey.SelStart = 10
    txtKey.Visible = True
    Call txtKey.SetFocus
    
    editMode = True
End Sub

Private Sub lstKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call lstKey_DblClick
    End If
End Sub

Private Sub txtCode_Change()
    If m_Templates.EOF = False Then
        m_Templates.Fields(1).Value = txtCode.Text
    End If
End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If editMode = False Then
            Call txtCode.SetFocus
        Else
            m_Templates.Fields(0).Value = txtKey.Text
            txtKey.Visible = False
        End If
    End If
End Sub

Private Sub txtKey_LostFocus()
    txtKey.Visible = False
    Dim newKey As String
    
    newKey = LCase$(txtKey.Text)
    If Len(newKey) > 0 And newKey <> "new" Then
        If editMode = False Then
            Call lstKey.AddItem(newKey)
            Call m_Templates.AddNew
            m_Templates.Fields(0).Value = newKey
            lstKey.ListIndex = lstKey.ListCount - 1&
        Else
            lstKey.List(lstKey.ListIndex) = newKey
            m_Templates.Fields(0).Value = newKey
        End If
    End If
End Sub

Private Function KeyForListItem() As String
    KeyForListItem = "Key='" & lstKey.List(lstKey.ListIndex) & "'"
End Function

Private Sub ValidateData()
    With m_Templates
        If .RecordCount > 0 Then
            .Filter = adRecModified
            If .RecordCount > 0 Then
                Call .MoveFirst
                Do While Not .EOF
                    Call ParseCode(.Fields(1).Value)
                    Call .MoveNext
                Loop
            End If
            .Filter = adFilterNone
        End If
    End With
End Sub

Private Sub ParseCode(ByVal sCode As String)

End Sub
