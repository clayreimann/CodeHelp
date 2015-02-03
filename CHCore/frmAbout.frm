VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About CodeHelp Core"
   ClientHeight    =   2565
   ClientLeft      =   2130
   ClientTop       =   1605
   ClientWidth     =   5160
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   1935
      Width           =   1215
   End
   Begin VB.Label lblGithub 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GitHub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   2160
      MousePointer    =   4  'Icon
      TabIndex        =   5
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "frmAbout.frx":058A
      Top             =   0
      Width           =   1920
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © luthv@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2190
      TabIndex        =   3
      Top             =   1110
      Width           =   2205
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2190
      TabIndex        =   2
      Top             =   690
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CodeHelp Core IDE Framework"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2190
      TabIndex        =   1
      Top             =   270
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1845
      Width           =   1920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) _
        As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub lblGithub_Click()
    Call Clipboard.SetText("https://github.com/clayreimann/CodeHelp")
    Call MsgBox("URL copied to clipboard")
End Sub
