VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICHCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'****************************************
'CodeHelp Core Parent Interface
'****************************************

Property Get VBE() As VBIDE.VBE

End Property

Function AddToCodeHelpMenu(ByVal Caption As String, _
    Optional ByVal iconBitmap As StdPicture = Nothing) As CommandBarControl

End Function
