VERSION 5.00
Begin VB.Form frmProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CodeHelp Coder Options"
   ClientHeight    =   6720
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
   ScaleHeight     =   6720
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKey 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   135
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2250
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   4815
      TabIndex        =   8
      Top             =   6300
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   8805
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8805
   End
   Begin VB.CommandButton cmdMarker 
      Caption         =   "Insert Marker"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5595
      Width           =   1215
   End
   Begin VB.ComboBox cboMarker 
      Height          =   330
      Left            =   2430
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5625
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
      TabIndex        =   4
      Top             =   1170
      Width           =   6945
   End
   Begin VB.ListBox lstKey 
      Height          =   4680
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Double Click to edit"
      Top             =   1170
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7470
      TabIndex        =   10
      Top             =   6300
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   6300
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Marker:"
      Height          =   210
      Index           =   2
      Left            =   1710
      TabIndex        =   5
      Top             =   5625
      Width           =   540
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Expanded Code:"
      Height          =   210
      Index           =   1
      Left            =   1755
      TabIndex        =   3
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Keys:"
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   900
      Width           =   420
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMarker_Click()
    If cboMarker.ListIndex > -1 Then
        txtCode.SelText = Split(cboMarker.Text, "   ")(0) & CLOSE_TAG
    End If
End Sub

Private Sub cmdNew_Click()
    Dim i As Long
    
    With txtKey
        lstKey.ListIndex = -1
        .Text = "New"

        i = lstKey.TopIndex
        i = lstKey.ListCount - i
        
        .Move lstKey.Left + 30, lstKey.Top + (i * .Height) + 30, lstKey.Width - 60

        .Visible = True
        .SelStart = 0
        .SelLength = 3
        .SetFocus
    End With
    editMode = False
    'templatedata.AddNew
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ERR_HANDLER
    ValidateData
    SaveData
    Unload Me
ERR_HANDLER:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    OpenData
    With templateData
        If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            lstKey.AddItem .Fields("Key").Value
            .MoveNext
        Loop
        End If
    End With
    
    With markerData
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboMarker.AddItem .Fields(0).Value & "   " & .Fields(1).Value
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub lstKey_Click()
    If templateData.RecordCount > 0 Then
        templateData.MoveFirst
        templateData.Find "Key='" & lstKey.List(lstKey.ListIndex) & "'"
        If Not templateData.EOF Then
            txtCode.Text = Trim$(templateData.Fields("Code").Value & " ")
        Else
            txtCode.Text = ""
        End If
    End If
End Sub

Private Sub lstKey_DblClick()
    Dim i As Long
    i = lstKey.TopIndex
    i = lstKey.ListIndex - i
    txtKey.Move lstKey.Left + 30, lstKey.Top + (i * txtKey.Height) + 30, lstKey.Width - 60
    txtKey.Text = lstKey.Text
    txtKey.SelStart = 10
    txtKey.Visible = True
    txtKey.SetFocus
    editMode = True
End Sub

Private Sub lstKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        lstKey_DblClick
    End If
End Sub

Private Sub txtCode_Change()
    If templateData.EOF = False Then
        templateData.Fields(1).Value = txtCode.Text
    End If
End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If editMode = False Then
            txtCode.SetFocus
        Else
            templateData.Fields(0).Value = txtKey.Text
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
            lstKey.AddItem newKey
            templateData.AddNew
            templateData.Fields(0).Value = newKey
            lstKey.ListIndex = lstKey.ListCount - 1&
        Else
            templateData.Fields(0).Value = newKey
        End If
    End If
End Sub

Private Sub ValidateData()
    With templateData
    If .RecordCount > 0 Then
        .Filter = adRecModified
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                ParseCode .Fields(1).Value
                .MoveNext
            Loop
        End If
        .Filter = adFilterNone
    End If
    End With
End Sub

Private Sub ParseCode(ByVal sCode As String)

End Sub
