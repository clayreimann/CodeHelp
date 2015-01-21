VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HookGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Private Type tagMSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private m_Hook As cHook
Private m_hookType As Long

Implements IHook

Public Sub StartHook(ByVal HookType As Long)
    Set m_Hook = New cHook
    m_Hook.Hook Me, HookType
    m_hookType = HookType
    
End Sub

Public Sub EndHook()

    If Not m_Hook Is Nothing Then
        m_Hook.UnHook
        Set m_Hook = Nothing
    End If
End Sub

Private Sub IHook_Proc(ByVal bBefore As Boolean, _
                       bHandled As Boolean, _
                       lReturn As Long, _
                       nCode As Long, _
                       wParam As Long, _
                       lParam As Long)

    If bBefore = False Then Exit Sub
    If nCode <> HC_ACTION Then Exit Sub
    
    Select Case m_hookType

        Case WH_KEYBOARD
            HookMon.KeyProc bHandled, lReturn, wParam, lParam
            
        Case WH_CALLWNDPROC
            Dim tCWP As CWPSTRUCT
            CopyMemory tCWP, ByVal lParam, Len(tCWP)
            HookMon.CWPMSGProc tCWP.hwnd, tCWP.message, tCWP.wParam, tCWP.lParam

        Case WH_GETMESSAGE
            Dim tMSG As tagMSG
            CopyMemory tMSG, ByVal lParam, Len(tMSG)
            
            If wParam = PM_REMOVE Then
                HookMon.CWPMSGProc tMSG.hwnd, tMSG.message, tMSG.wParam, tMSG.lParam
            End If
    End Select

End Sub
