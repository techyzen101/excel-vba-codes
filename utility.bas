Attribute VB_Name = "utility"
Sub Range_End_Method()
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    If Cells(1, lCol).MergeCells Then
        lCol = Cells(1, lCol).MergeArea.Columns.Count
    End If
    
    If Cells(lRow, 1).MergeCells Then
        lRow = Cells(lRow, 1).MergeArea.Rows.Count
    End If
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol
  
End Sub

