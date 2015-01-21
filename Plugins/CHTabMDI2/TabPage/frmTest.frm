VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Toggle Visible"
      Height          =   330
      Left            =   720
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   330
      Left            =   765
      TabIndex        =   7
      Top             =   1350
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Close Button:"
      Height          =   1365
      Left            =   2160
      TabIndex        =   3
      Top             =   900
      Width           =   1995
      Begin VB.OptionButton optCloseButton 
         Caption         =   "On Active Tab"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   1455
      End
      Begin VB.OptionButton optCloseButton 
         Caption         =   "Right Most"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   607
         Width           =   1230
      End
      Begin VB.OptionButton optCloseButton 
         Caption         =   "Hidden"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Active"
      Height          =   330
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   330
      Left            =   780
      TabIndex        =   1
      Top             =   915
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   150
      TabIndex        =   0
      Text            =   "15"
      Top             =   915
      Width           =   600
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCloseItem 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close &All"
      End
      Begin VB.Menu mnuCloseButActive 
         Caption         =   "Close All &But Active"
      End
      Begin VB.Menu mnuPopSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButtons 
         Caption         =   "Buttons"
         Begin VB.Menu mnuClosePop 
            Caption         =   "Close Button"
            Begin VB.Menu mnuClosePosition 
               Caption         =   "Hidden"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuClosePosition 
               Caption         =   "Rightmost"
               Index           =   2
            End
            Begin VB.Menu mnuClosePosition 
               Caption         =   "On Active Tab"
               Index           =   4
            End
         End
         Begin VB.Menu mnuNavButtons 
            Caption         =   "Navigation Buttons"
         End
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents tabMgr As TabManager
Attribute tabMgr.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
    tabMgr.InsertItem Text1.Text
End Sub

Private Sub cmdDelete_Click()
    tabMgr.RemoveItem tabMgr.SelectedItem
End Sub

Private Sub cmdVisible_Click()
    Dim i As Long
    
    i = CLng(Text1.Text)
    If tabMgr.Items.Exists("#" & i) Then
        tabMgr.Items(i).Visible = Not tabMgr.Items(i).Visible
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Long
    
    i = CLng(Text1.Text)
    If tabMgr.Items.Exists("#" & i) Then
        tabMgr.Items(i).Selected = True
    End If
End Sub


Private Sub Form_Load()
    Set tabMgr = New TabManager
    Dim i As Long
    Dim item As TabItem
    
    Form_Resize
    
    For i = 1 To 100
        Set item = tabMgr.InsertItem("Hello " & i)
    Next
    
    tabMgr.Top = 20
    'tabMgr.Items(1).Selected = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        'GetKeyState VB
        tabMgr.OnLMouseDown ScaleX(x, ScaleMode, vbPixels), ScaleY(y, ScaleMode, vbPixels)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    tabMgr.OnMouseMove Button, ScaleX(x, ScaleMode, vbPixels), ScaleY(y, ScaleMode, vbPixels)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    tabMgr.OnMouseUp Button, ScaleX(x, ScaleMode, vbPixels), ScaleY(y, ScaleMode, vbPixels)
    
End Sub

Private Sub Form_Paint()
    tabMgr.Refresh hdc
End Sub

Private Sub Form_Resize()
    tabMgr.Width = ScaleX(ScaleWidth, ScaleMode, vbPixels)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tabMgr = Nothing
End Sub

Private Sub mnuCloseAll_Click()
    
    tabMgr.RemoveAll True

End Sub

Private Sub mnuCloseButActive_Click()
    tabMgr.RemoveAllButActive
End Sub

Private Sub mnuCloseItem_Click()
    Dim item As TabItem
    
    Set item = tabMgr.Items("#" & mnuCloseItem.Tag)
    tabMgr.RemoveItem item
End Sub

Private Sub mnuClosePosition_Click(Index As Integer)
    optCloseButton_Click (Index)
End Sub

Private Sub optCloseButton_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To 4 Step 2
        optCloseButton(i).Value = (Index = i)
        mnuClosePosition(i).Checked = (Index = i)
    Next
    tabMgr.PaintManager.ShowCloseButton = Index
End Sub

Private Sub tabMgr_MouseUp(ByVal Button As MouseButtonConstants, ByVal item As TabItem)
    If Button = vbRightButton Then
        'item can be nothing if user click on empty space
        mnuCloseItem.Enabled = Not (item Is Nothing)
        If mnuCloseItem.Enabled Then mnuCloseItem.Tag = item.Index
        mnuCloseButActive.Enabled = (tabMgr.Items.Count > 1)
        mnuCloseAll.Enabled = mnuCloseButActive.Enabled
        PopupMenu mnuPop
    End If
End Sub

Private Sub tabMgr_RequestRedraw(hdc As Long)
    hdc = Me.hdc
End Sub
