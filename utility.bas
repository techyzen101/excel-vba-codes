Attribute VB_Name = "utility"

Public Function lastUsedRow() As Long
    lastUsedRow = Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(lastUsedRow, 1).MergeCells Then
        lastUsedRow = Cells(lastUsedRow, 1).MergeArea.Rows.Count
    End If
End Function

Public Function lastUsedColumn() As Long
    lastUsedColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    If Cells(1, lastUsedColumn).MergeCells Then
        lastUsedColumn = Cells(1, lastUsedColumn).MergeArea.Columns.Count
    End If
End Function
