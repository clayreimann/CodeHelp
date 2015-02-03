VERSION 5.00
Begin VB.Form frmPlugins 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugins Manager"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   Icon            =   "frmPlugins.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "More Info..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6030
      TabIndex        =   13
      Top             =   4350
      Width           =   1155
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4815
      TabIndex        =   12
      Top             =   4350
      Width           =   1155
   End
   Begin VB.CheckBox chkLoad 
      Caption         =   "Enabled"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3300
      TabIndex        =   8
      Top             =   3870
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   7365
      TabIndex        =   10
      Top             =   0
      Width           =   7365
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmPlugins.frx":058A
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPlugins.frx":0E54
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   780
         TabIndex        =   11
         Top             =   210
         Width           =   6345
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   9
      Top             =   4350
      Width           =   1155
   End
   Begin VB.ListBox lstPlugin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   225
      TabIndex        =   1
      Top             =   1395
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Copyright:"
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
      Left            =   3300
      TabIndex        =   5
      Top             =   2175
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   3300
      TabIndex        =   3
      Top             =   1785
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Installed plugin(s):"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   1305
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description goes here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3300
      TabIndex        =   7
      Top             =   2565
      Width           =   3765
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
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
      Left            =   3300
      TabIndex        =   2
      Top             =   1365
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
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
      Left            =   4170
      TabIndex        =   4
      Top             =   1785
      Width           =   615
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
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
      Left            =   4170
      TabIndex        =   6
      Top             =   2175
      Width           =   2205
   End
End
Attribute VB_Name = "frmPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Plugins As Plugins
Private m_UpdateFromCode As Boolean

Property Set Plugins(ByVal value As Plugins)
    Dim oPlugin As ICHPlugin
    Set m_Plugins = value
    
    For Each oPlugin In m_Plugins
        Call lstPlugin.AddItem(oPlugin.Name)
    Next
    
    If lstPlugin.ListCount > 0 Then lstPlugin.ListIndex = 0
End Property

Private Sub chkLoad_Click()
    Dim oPlugin As ICHPlugin
    Dim Enabled As Boolean
    
    If m_UpdateFromCode Then Exit Sub
    
    On Error GoTo ERR_HANDLER
    Enabled = (chkLoad.value = vbChecked)
    Set oPlugin = m_Plugins(lstPlugin.ListIndex + 1&)
    oPlugin.Enabled = Enabled
    
    If Enabled Then
        oPlugin.CHCore = gPtr
        Call oPlugin.OnConnection(ext_cm_AfterStartup, customVar)
    Else
        Call oPlugin.OnDisconnect(ext_dm_UserClosed, customVar)
    End If
    
    Call SaveSetting("CodeHelp", oPlugin.Name, "Enabled", oPlugin.Enabled)
ERR_HANDLER:
End Sub

Private Sub cmdHelp_Click()
    Dim oPlugin As ICHPlugin
    On Error GoTo ERR_HANDLER
    
    Set oPlugin = m_Plugins(lstPlugin.ListIndex + 1&)
    Call oPlugin.ShowHelp

ERR_HANDLER:
End Sub

Private Sub cmdOK_Click()
    Call Me.Hide
End Sub

Private Sub cmdProperties_Click()
    Dim oPlugin As ICHPlugin
    On Error GoTo ERR_HANDLER
    
    Set oPlugin = m_Plugins(lstPlugin.ListIndex + 1&)
    Call oPlugin.ShowPropertyDialog
    
ERR_HANDLER:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Plugins = Nothing
End Sub

Private Sub lstPlugin_Click()
    Dim oPlugin As ICHPlugin
    
    Set oPlugin = m_Plugins(lstPlugin.ListIndex + 1&)
    lblName.Caption = oPlugin.LongName
    lblVersion.Caption = oPlugin.Version
    lblCopyright.Caption = oPlugin.CopyRight
    lblDesc.Caption = oPlugin.Description
    
    m_UpdateFromCode = True
    chkLoad.Enabled = True
    chkLoad.value = IIf(oPlugin.Enabled, vbChecked, vbUnchecked)
    m_UpdateFromCode = False
    
    cmdHelp.Enabled = oPlugin.HaveExtendedHelp
    cmdProperties.Enabled = oPlugin.HaveProperties
End Sub
