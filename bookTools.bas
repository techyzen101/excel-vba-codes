Attribute VB_Name = "bookTools"
Public Function IsOutsideCall() As Boolean
    IsOutsideCall = False
    
    If Not ActiveWorkbook Is ThisWorkbook Then
        IsOutsideCall = True
    End If
End Function
