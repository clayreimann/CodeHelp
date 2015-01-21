VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITabPainter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TabPainter Interface

Public Enum TabItemType
    DrawBackGround
    DrawTabItem
    DrawCloseButton
    DrawNavLeftButton
    DrawNavRightButton
End Enum

Public Sub DrawItem(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal itemType As TabItemType)

End Sub

Public Function CalculateTabWidth(ByVal item As TabItem, ByVal oDC As MemoryDC) As Long

End Function

Public Sub DrawShortcut(ByVal oDC As MemoryDC, bounds As RECT, ByVal item As TabItem, ByVal visibleIndex As Long)

End Sub
