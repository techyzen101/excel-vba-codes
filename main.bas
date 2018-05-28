Attribute VB_Name = "main"
Public Sub Test()
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim lo As ListObject
    
    CreateSheet "ws2", ws2
    GetSheet "ws", ws
    GetTable ws, "table", lo
    
    If Not ws Is Nothing Then
        ws.Cells(1, 1) = "Success"
    Else
        MsgBox "Failed"
    End If
    
    If Not ws2 Is Nothing Then
        ws2.Cells(2, 2) = "Success v2"
    Else
        MsgBox "Failed 2"
    End If
End Sub
