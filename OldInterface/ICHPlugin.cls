VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICHPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'CodeHelp Plugin Interface
'****************************************

Property Get Name() As String

End Property

Property Get LongName() As String

End Property

Property Get Version() As String

End Property

Property Get Description() As String

End Property

Property Get CopyRight() As String

End Property

Sub OnConnection(ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, custom() As Variant)

End Sub

Sub OnDisconnect(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

End Sub

Sub OnWinProcHook(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    bHandled As Boolean, placeSubclass As Boolean, before As Boolean)
'Purpose :
'CWP Hook and MSG Hook notification, Called by CHCore
'Params:
'- bHandled, set to true to prevent other plugins from getting the notification
'- placeSublass, set to true to notify CHCore to place a subclass for the window for the particular msg
'               Use this if you want to make modification to the return value, wParam or lParam value
'- before, place the subclass before or after default winproc
End Sub

Sub OnWinProc(ByVal hWnd As Long, ByVal uMsg As Long, wParam As Long, lParam As Long, bHandled As Boolean, lreturn As Long)
'Purpose:
'To allow client plugin to modify the default window message processing
End Sub

Sub OnKeyHook(bHandled As Boolean, lreturn As Long, wParam As Long, lParam As Long)
'Purpose:
'Keyboard hook notification
End Sub

Property Let CHCore(ByVal value As Long)
'Purpose:
'Save the pointer to Parent AddIn Designer object
End Property

Property Get Enabled() As Boolean
'Allow enable/disabling the plugin at runtime
End Property

Property Let Enabled(ByVal value As Boolean)
'Allow enable/disabling the plugin at runtime
End Property

Property Get HaveProperties() As Boolean

End Property

Sub ShowPropertyDialog()

End Sub

Property Get HaveExtendedHelp() As Boolean

End Property

Sub ShowHelp()

End Sub

