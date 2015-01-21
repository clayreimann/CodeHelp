VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HookMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_KeyHook As HookGate
Private m_CWPHook As HookGate
Private m_MSGHook As HookGate

Private m_Subclass As cSubclass
Private m_SubClassPlugin As ICHPlugin

Implements ISubclass

Sub StartMonitor()
    
    Set m_KeyHook = New HookGate
    m_KeyHook.StartHook WH_KEYBOARD

    Set m_CWPHook = New HookGate
    m_CWPHook.StartHook WH_CALLWNDPROC

    Set m_MSGHook = New HookGate
    m_MSGHook.StartHook WH_GETMESSAGE
    
End Sub

Sub EndMonitor()
On Error Resume Next
    m_KeyHook.EndHook
    Set m_KeyHook = Nothing
    
    m_CWPHook.EndHook
    Set m_CWPHook = Nothing
    
    m_MSGHook.EndHook
    Set m_MSGHook = Nothing
    
End Sub

Sub KeyProc(bHandled As Boolean, lReturn As Long, wParam As Long, lParam As Long)
On Error Resume Next
    Dim plugin As ICHPlugin
    
    For Each plugin In mCHCore.Plugins
        
        If plugin.Enabled Then
            If bHandled = False Then
                plugin.OnKeyHook bHandled, lReturn, wParam, lParam
                
            End If
        End If
    Next
End Sub

Sub CWPMSGProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'On Error Resume Next
    Dim plugin As ICHPlugin
    Dim bHandled As Boolean, placeSubclass As Boolean, before As Boolean
    For Each plugin In mCHCore.Plugins
        
        If plugin.Enabled Then
            
            If bHandled = False Then
                placeSubclass = False
                before = False
                plugin.OnWinProcHook hwnd, uMsg, wParam, lParam, bHandled, placeSubclass, before
                
                If placeSubclass Then
                    Set m_SubClassPlugin = plugin
                    If before Then
                        SubclassThisWindowMsg hwnd, uMsg, MSG_BEFORE
                    Else
                        SubclassThisWindowMsg hwnd, uMsg, MSG_AFTER
                    End If
                    Exit For 'currently only support one plugin per subclassed msg
                End If
            End If
        End If
    Next
End Sub

Private Sub ISubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    If m_SubClassPlugin.Enabled Then
        m_SubClassPlugin.OnWinProc hwnd, uMsg, wParam, lParam, bHandled, lReturn
    End If
    'we're done so unsubclass right away
    'add reference lock for win98
    Set lockSubclass = m_Subclass
    m_Subclass.UnSubclass
    Set m_Subclass = Nothing
    Set m_SubClassPlugin = Nothing
End Sub

Private Sub SubclassThisWindowMsg(ByVal hwnd As Long, ByVal uMsg As Long, ByVal MsgWhen As enMsgWhen)
    Set m_Subclass = New cSubclass
    m_Subclass.AddMsg uMsg, MsgWhen
    m_Subclass.Subclass hwnd, Me
End Sub
