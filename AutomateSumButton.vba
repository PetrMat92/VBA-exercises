Public Sub AutomateTotalSum()

    Dim LastRow As Long  ' variable to store the last row with data
    Dim TotalExpenseColumn As Long  ' variable to store the column number of the "Total Expense" column
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
    
        ' find the column number of the "Total Expense" column in the first row
        TotalExpenseColumn = Range("1:1").Find("Total Expense").Column
    
        ' find the last row with data in the "Total Expense" column
        LastRow = Cells(Rows.Count, TotalExpenseColumn).End(xlUp).Row
    
        ' enter the SUM formula in the next row of the "Total Expense" column
        Cells(LastRow + 1, TotalExpenseColumn).Formula = "=SUM(" & Cells(2, TotalExpenseColumn).Address(False, False) & ":" & Cells(LastRow, TotalExpenseColumn).Address(False, False) & ")"
        
    Next ws
End Sub

