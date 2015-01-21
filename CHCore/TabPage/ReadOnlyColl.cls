VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Panel"
Attribute VB_Ext_KEY = "Member0" ,"Panel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_Col As Collection

Friend Sub AddItem(ByVal newObject As Object, ByVal key As String, Optional Index As Long = 0)
    If Index > 0 And Index <= m_Col.Count Then
        m_Col.Add newObject, key, Index
    Else
        m_Col.Add newObject, key
    End If
End Sub

Friend Sub Remove(ByVal key As String)
  If Exists(key) Then m_Col.Remove key
End Sub

Public Property Get item(vntIndexKey As Variant) As Object
Attribute item.VB_UserMemId = 0
  Set item = m_Col(vntIndexKey)
End Property

Public Property Get Count() As Long
  Count = m_Col.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = m_Col.[_NewEnum]
End Property

Function Exists(ByVal key As String) As Boolean
  Dim childItem As Object
  
  On Error GoTo ERR_HANDLER
  
  Set childItem = m_Col.item(key)
  Exists = True
  Set childItem = Nothing
  
ERR_HANDLER:
  
End Function

Private Sub Class_Initialize()
  Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_Col = Nothing
End Sub

