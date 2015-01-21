VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Plugin const
Private Const CH_LONGNAME As String = "Your Plugin Title"
Private Const CH_DESCRIPTION As String = "Description of what your plugin do"
Private Const CH_COPYRIGHT As String = "Your copyright info"

Private m_lPtr As Long
private m_bEnabled as boolean

Implements ICHPlugin

Private Property Let ICHPlugin_CHCore(ByVal RHS As Long)
    'Save the Pointer for later use
    m_lPtr = RHS
End Property

Private Property Get ICHPlugin_CopyRight() As String
    ICHPlugin_CopyRight = CH_COPYRIGHT
End Property

Private Property Get ICHPlugin_Description() As String
    ICHPlugin_Description = CH_DESCRIPTION
End Property

Private Property Let ICHPlugin_Enabled(ByVal RHS As Boolean)
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   m_bEnabled = rhs
End Property

Private Property Get ICHPlugin_Enabled() As Boolean
'Enable/disable this plugin in plugin manager
'if disable the ondisconnect method will be called, and the plugin will be excluded from
'msg processing
   ICHPlugin_Enabled = m_bEnabled
End Property

Private Property Get ICHPlugin_HaveExtendedHelp() As Boolean
'Enable/disable help button in plugin manager
End Property

Private Property Get ICHPlugin_HaveProperties() As Boolean
'Enable/disable properties button in plugin manager
End Property

Private Property Get ICHPlugin_LongName() As String
    ICHPlugin_LongName = CH_LONGNAME
End Property

Private Property Get ICHPlugin_Name() As String
    ICHPlugin_Name = App.Title
End Property

Private Sub ICHPlugin_OnConnection(ByVal ConnectMode As CodeHelpDef.ext_ConnectMode, _
                                   custom() As Variant)
    
    'Sample use of the ICHCore pointer
    'It's advisable not to save the ICHCore object itself, always use the helper function to obtain the
    'ICHCore object from the pointer
    
    Dim dsr As ICHCore
    
    Set dsr = GetCHCore(m_lPtr)
    
    'Do your initializing stuff here
End Sub

Private Sub ICHPlugin_OnDisconnect(ByVal RemoveMode As CodeHelpDef.ext_DisconnectMode, _
                                   custom() As Variant)
    'Do your clean up here
End Sub

Private Sub ICHPlugin_OnKeyHook(bHandled As Boolean, _
                                lreturn As Long, _
                                wParam As Long, _
                                lParam As Long)
    'do any keyboard related code here
End Sub

Private Sub ICHPlugin_OnWinProc(ByVal hWnd As Long, _
                                ByVal uMsg As Long, _
                                wParam As Long, _
                                lParam As Long, _
                                bHandled As Boolean, _
                                lreturn As Long)
    'subclassed message goes here
End Sub

Private Sub ICHPlugin_OnWinProcHook(ByVal hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long, _
                                    bHandled As Boolean, _
                                    placeSubclass As Boolean, _
                                    before As Boolean)
    'Hook msg goes here
End Sub

Private Sub ICHPlugin_ShowHelp()
'Show help instruction when user click on help button in plugin manager
End Sub

Private Sub ICHPlugin_ShowPropertyDialog()
'show property dialog when user click on properties button in plugin manager
End Sub

Private Property Get ICHPlugin_Version() As String
    ICHPlugin_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

