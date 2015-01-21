Attribute VB_Name = "mMain"
Option Explicit

Public templateData As Recordset
Public markerData As Recordset

Sub OpenData()
    
    Dim con As Connection
    Dim errNum As Long, errDesc As String
    
    On Error GoTo ERR_HANDLER
    
    If templateData Is Nothing Then
        Set con = New Connection
        With con
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\template.mdb;"
            .Open
        End With
        
        Set templateData = New Recordset
        With templateData
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LOCKTYPE = adLockBatchOptimistic
            Set .ActiveConnection = con
            .Open "SELECT * FROM Coder ORDER BY key"
            Set .ActiveConnection = Nothing
        End With
        
        Set markerData = New Recordset
        With markerData
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LOCKTYPE = adLockReadOnly
            Set .ActiveConnection = con
            .Open "SELECT * FROM Marker"
            
        End With
        
    End If

EXIT_POINT:
    On Error Resume Next
    Set templateData.ActiveConnection = Nothing
    Set markerData.ActiveConnection = Nothing
    con.Close
    Set con = Nothing
    Err.Clear
    If errNum <> 0 Then
        MsgBox "Error while opening template.mdb." & vbCrLf & _
        "Please make sure that template.mdb file is placed in the same folder as CHCoder.dll", vbInformation, "CodeHelp Coder Error"
    End If
    
ERR_HANDLER:
    errNum = Err.Number
    errDesc = Err.Description
    Resume EXIT_POINT
End Sub

Sub CloseData()
    Set templateData = Nothing
    Set markerData = Nothing
    
End Sub

Sub SaveData()
    Dim con As Connection
    Set con = New Connection
    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\template.mdb;"
        .Open
    End With

    With templateData
        Set .ActiveConnection = con
        .UpdateBatch
        Set .ActiveConnection = Nothing
    End With
    con.Close
    Set con = Nothing
End Sub
