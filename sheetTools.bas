Attribute VB_Name = "sheetTools"
' Sheet Related
Public Sub CreateSheet(ByVal sheetName As String, ByRef sheetObject As Worksheet)
    GetSheet sheetName, sheetObject
    If sheetObject Is Nothing Then
        Set sheetObject = ThisWorkbook.Sheets.Add
        sheetObject.Name = sheetName
    End If
End Sub

Public Sub DeleteSheet(ByVal sheetObject As Worksheet)

End Sub

Public Sub GetSheet(ByVal sheetName As String, ByRef sheetObject As Worksheet)
    On Error Resume Next
    Set sheetObject = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

' Table Related
Public Sub GetTable(ByVal inSheet As Worksheet, ByVal tableName As String, ByRef tableObject As ListObject)
    On Error Resume Next
    Set tableObject = inSheet.ListObjects(tableName)
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
