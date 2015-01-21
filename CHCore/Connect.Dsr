VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8490
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13350
   _ExtentX        =   23548
   _ExtentY        =   14975
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

Private m_VBE                   As VBIDE.VBE
Private cbarCodeHelp            As CommandBarControl
Private WithEvents ctlAbout     As CommandBarEvents
Attribute ctlAbout.VB_VarHelpID = -1
Private WithEvents ctlPlugins   As CommandBarEvents
Attribute ctlPlugins.VB_VarHelpID = -1
Private coreGroup               As CommandBarControl
Private m_AddInInst             As Object
Implements ICHCore

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    Dim cmdNew As CommandBarControl
    Dim plugin As ICHPlugin
    
    'save the vb instance
    Set m_VBE = Application
    Set m_AddInInst = AddInInst
    
    'Use index for International Version of VB, thanks bicio!
    Dim menuBar As CommandBar
        
    Set menuBar = m_VBE.CommandBars(1)
    Set cbarCodeHelp = menuBar.Controls.Add(msoControlPopup, , , menuBar.Controls.Count - 1)
    
    cbarCodeHelp.Caption = "&CodeHelp"
    
    Set cmdNew = AddMenuItem("&Plugins Manager...", , False)
    Set ctlPlugins = m_VBE.Events.CommandBarEvents(cmdNew)
    
    Set cmdNew = AddMenuItem("&About...", , False)
    Set ctlAbout = m_VBE.Events.CommandBarEvents(cmdNew)
    
    LoadPlugins ConnectMode, custom
    'start low level message monitoring
    Set HookMon = New HookMonitor
    HookMon.StartMonitor
    
    'tell plugins we're ready
    For Each plugin In mCHCore.Plugins
        If plugin.enabled Then
            plugin.OnConnection ConnectMode, custom
        End If
    Next
    customVar = custom
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description, vbInformation, "Error Encountered"
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    EndMonitor
    RemovePlugins RemoveMode, custom
    
    Set coreGroup = Nothing
    Set ctlAbout = Nothing
    Set ctlPlugins = Nothing
    cbarCodeHelp.Delete
    Set cbarCodeHelp = Nothing
    
    Set m_VBE = Nothing
End Sub

Private Sub ctlAbout_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmAbout.Show vbModal
End Sub

Private Sub ctlPlugins_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim f As frmPlugins
    Set f = New frmPlugins
    Set f.Plugins = mCHCore.Plugins
    f.Show vbModal
    Unload f
    Set f = Nothing
End Sub



Private Function AddMenuItem(ByVal Caption As String, _
    Optional ByVal iconPic As stdole.Picture = Nothing, _
    Optional aboveSeparator As Boolean = True) As CommandBarControl
    
    If Not cbarCodeHelp Is Nothing Then
        Dim dropDown As CommandBarPopup
        Dim newButton As CommandBarButton
        Dim iconBmp As StdPicture
         
        Set dropDown = cbarCodeHelp
        
        If aboveSeparator Then
            'add menu item above the menuseparator
            Set newButton = dropDown.Controls.Add(msoControlButton, , , coreGroup.Index)
        
        Else
            
            'add menu item below the separator
            Set newButton = dropDown.Controls.Add(msoControlButton)
            If coreGroup Is Nothing Then
                'add separator
                Set coreGroup = newButton
                coreGroup.BeginGroup = True
            End If
            
        End If
        
        newButton.Caption = Caption
                
        If Not iconPic Is Nothing Then

On Error GoTo SKIP_FACE
            
            Clipboard.Clear
            newButton.CopyFace
            Set iconBmp = Clipboard.GetData
            
            CopyIconToClipBoardAsBmp iconPic, iconBmp

            newButton.PasteFace
            Clipboard.Clear
SKIP_FACE:
        End If
        
        Set AddMenuItem = newButton
    End If
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
    HookMon.EndMonitor
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
        LoadPluginDLL sFile, ConnectMode, custom
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
    
    Dim plugin As ICHPlugin
    Dim className As String
    
    On Error Resume Next
    gPtr = ObjPtr(Me)
    
    Set tliApp = New TLIApplication
    
    Set tliInfo = tliApp.TypeLibInfoFromFile(fileName)
        
    For Each ccI In tliInfo.CoClasses
        For Each inf In ccI.Interfaces
            If inf.Guid = GUID_ID Then
                'this class implements ICHPlugin
                className = tliInfo.Name & "." & ccI.Name
                Set plugin = CreateObject(className)
                
                If Not plugin Is Nothing Then
                    
                    plugin.CHCore = gPtr
                    plugin.enabled = (CLng(GetSetting("CodeHelp", plugin.Name, "Enabled", vbChecked)) = vbChecked)
                    
                    mCHCore.Plugins.Add plugin
                    Set plugin = Nothing
                End If
                
                Exit For
            End If
        Next
        'continue in case there are more than one class that implements ICHPlugin
    Next
    
End Sub

Private Sub RemovePlugins(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Dim plugin As ICHPlugin
    Dim pList As Plugins
    Dim i As Long
    
    Set pList = mCHCore.Plugins
    
    For Each plugin In pList
        plugin.OnDisconnect RemoveMode, custom
    Next
    
    'Delete plugin from collection
    For i = 1 To pList.Count
        pList.Remove 1
    Next
    
    Set mCHCore.Plugins = Nothing
End Sub

