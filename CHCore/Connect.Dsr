VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12765
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   18000
   _ExtentX        =   31750
   _ExtentY        =   22516
   _Version        =   393216
   Description     =   "CodeHelp Core IDE Extender Framework"
   DisplayName     =   "CodeHelp IDE Extender"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_CHCommandBarItem      As CommandBarControl
Private m_CHCoreMenuGroup       As CommandBarControl
Private WithEvents ctlAbout     As CommandBarEvents
Attribute ctlAbout.VB_VarHelpID = -1
Private WithEvents ctlPlugins   As CommandBarEvents
Attribute ctlPlugins.VB_VarHelpID = -1

Private m_VBE                   As VBIDE.VBE
Private m_AddInInst             As Object

Implements ICHCore

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Dim cmdNew As CommandBarControl
    Dim oPlugin As ICHPlugin
    
    'save the vb instance
    Set m_VBE = Application
    Set m_AddInInst = AddInInst
    
    'Use index for International Version of VB, thanks bicio!
    Dim menuBar As CommandBar
    
    On Error GoTo GeneralError
    Set menuBar = m_VBE.CommandBars(1)
    Set m_CHCommandBarItem = menuBar.Controls.Add(msoControlPopup, , , menuBar.Controls.Count - 1)
    m_CHCommandBarItem.Caption = "&CodeHelp"
    
    Set cmdNew = AddMenuItem("&Plugins Manager...", , False)
    Set ctlPlugins = m_VBE.Events.CommandBarEvents(cmdNew)
    
    Set cmdNew = AddMenuItem("&About...", , False)
    Set ctlAbout = m_VBE.Events.CommandBarEvents(cmdNew)
    
    Call LoadPlugins(ConnectMode, custom)
    'start low level message monitoring
    Set HookMon = New HookMonitor
    Call HookMon.StartMonitor
    
    'tell plugins we're ready
    For Each oPlugin In mCHCore.Plugins
        If oPlugin.Enabled Then
            On Error GoTo PluginConnectFailed
            Call oPlugin.OnConnection(ConnectMode, custom)
            GoTo NextPlugin
PluginConnectFailed:
            Call MsgBox("Error enabling oPlugin: " & oPlugin.Name & " in file " & Err.Source & " on line: " & Erl & vbCrLf _
                        & Err.Description, vbInformation, "Couldn't connect oPlugin")
NextPlugin:
        End If
    Next
    
    customVar = custom
    Exit Sub
    
GeneralError:
    Call MsgBox(Err.Description, vbInformation, "Error Encountered")
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    Call EndMonitor
    Call RemovePlugins(RemoveMode, custom)
    
    Set m_CHCoreMenuGroup = Nothing
    Set ctlAbout = Nothing
    Set ctlPlugins = Nothing
    Call m_CHCommandBarItem.Delete
    Set m_CHCommandBarItem = Nothing
    
    Set m_VBE = Nothing
End Sub

Private Sub ctlAbout_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call frmAbout.Show(vbModal)
End Sub

Private Sub ctlPlugins_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim f As frmPlugins
    
    Set f = New frmPlugins
    Set f.Plugins = mCHCore.Plugins
    
    Call f.Show(vbModal)
    Call Unload(f)
    
    Set f = Nothing
End Sub


Private Function AddMenuItem(ByVal Caption As String, Optional ByVal iconPic As stdole.Picture = Nothing, Optional aboveSeparator As Boolean = True) As CommandBarControl
    Dim dropDown As CommandBarPopup
    Dim newButton As CommandBarButton
    Dim iconBmp As StdPicture
    
    If m_CHCommandBarItem Is Nothing Then Exit Function
         
    Set dropDown = m_CHCommandBarItem
    If aboveSeparator Then 'add menu item above the menuseparator
        Set newButton = dropDown.Controls.Add(msoControlButton, , , m_CHCoreMenuGroup.Index)
    
    Else 'add menu item below the separator
        Set newButton = dropDown.Controls.Add(msoControlButton)
        If m_CHCoreMenuGroup Is Nothing Then
            Set m_CHCoreMenuGroup = newButton
            m_CHCoreMenuGroup.BeginGroup = True 'add separator
        End If
        
    End If
    
    newButton.Caption = Caption
            
    If Not iconPic Is Nothing Then
        On Error GoTo SKIP_FACE
        Call Clipboard.Clear
        
        Call newButton.CopyFace
        Set iconBmp = Clipboard.GetData
        Call CopyIconToClipBoardAsBmp(iconPic, iconBmp)
        Call newButton.PasteFace
        
        Call Clipboard.Clear
    End If
    
SKIP_FACE:
    Set AddMenuItem = newButton

End Function

Private Property Get ICHCore_AddInInst() As Object
    Set ICHCore_AddInInst = m_AddInInst
End Property

Private Function ICHCore_AddToCodeHelpMenu(ByVal Caption As String, Optional ByVal iconBitmap As Variant) As Object
    Set ICHCore_AddToCodeHelpMenu = AddMenuItem(Caption, iconBitmap, True)
End Function

Private Property Get ICHCore_VBE() As VBIDE.VBE
    Set ICHCore_VBE = m_VBE
End Property

Private Sub EndMonitor()
    Call HookMon.EndMonitor
    Set HookMon = Nothing
End Sub

Private Sub LoadPlugins(ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, custom() As Variant)
    Dim sPath As String
    Dim sFile As String
    
    Set mCHCore.Plugins = New Plugins
    sPath = App.Path & "\Plugins\"
    sFile = Dir(sPath & "*.dll")
    Do While Len(sFile) > 0
        sFile = sPath & sFile
        Call LoadPluginDLL(sFile, ConnectMode, custom)
        sFile = Dir()
    Loop
End Sub

Private Sub LoadPluginDLL(ByVal fileName As String, ByVal ConnectMode As ext_ConnectMode, custom() As Variant)
    'ICHPlugin Guid**************************************************************
    'This is defined in CHLib.tlb
    'All plugins must inplements this interface to be succesfully load by CHCore
    Const GUID_ID = "{0412CF22-0411-4255-9EE1-57354438E4EB}"
    '****************************************************************************
    Dim tliApp As TLIApplication
    Dim tliInfo As TypeLibInfo
    Dim ccI As CoClassInfo
    Dim inf As InterfaceInfo
    
    Dim oPlugin As ICHPlugin
    Dim className As String
    
    On Error Resume Next
    gPtr = ObjPtr(Me)
    
    Set tliApp = New TLIApplication
    Set tliInfo = tliApp.TypeLibInfoFromFile(fileName)
        
    For Each ccI In tliInfo.CoClasses
        For Each inf In ccI.Interfaces
            'more than one class in the dll can implement ICHPlugin
            If inf.Guid = GUID_ID Then
                'this class implements ICHPlugin
                className = tliInfo.Name & "." & ccI.Name
                Set oPlugin = CreateObject(className)
                
                If Not oPlugin Is Nothing Then
                    oPlugin.CHCore = gPtr
                    oPlugin.Enabled = CBool(GetSetting("CodeHelp", oPlugin.Name, "Enabled", True))
                    Call mCHCore.Plugins.Add(oPlugin)
                    
                    Set oPlugin = Nothing
                End If
                
                Exit For
            End If
        Next
    Next
    
End Sub

Private Sub RemovePlugins(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Dim oPlugin As ICHPlugin
    Dim oPluginList As Plugins
    Dim idx As Long
    
    Set oPluginList = mCHCore.Plugins
    
    For Each oPlugin In oPluginList
        Call oPlugin.OnDisconnect(RemoveMode, custom)
    Next
    
    'Delete oPlugin from collection
    For idx = 1 To oPluginList.Count
        Call oPluginList.Remove(1)
    Next
    
    Set mCHCore.Plugins = Nothing
End Sub

