Sub LoopYearlyReport()
    ' Declare worksheet variable and boolean variable for first loop iteration
    Dim ws As Worksheet
    Dim firsttime As Boolean
    firsttime = True
    
    ' Loop through all worksheets in the workbook
    For Each ws In Worksheets
        ' Select worksheet and skip the "YEARLY REPORT" worksheet
        Worksheets(ws.Name).Select
        If ws.Name <> "YEARLY REPORT" Then
        
            ' Call subroutines to insert, format headers and calculate total sum
            InsertHeaders
            FormatHeaders
            AutomateTotalSUM
            
            ' Select and copy data in the current worksheet
            Range("A2").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
        
            ' Select "YEARLY REPORT" worksheet and paste data
            Worksheets("YEARLY REPORT").Select
            Range("A30000").Select
            Selection.End(xlUp).Select
            If firsttime <> True Then
                ActiveCell.Offset(1, 0).Select
            Else
                firsttime = False
            End If
            
            ActiveSheet.Paste
        End If
    ' Move to the next worksheet in the loop
    Next ws
    
    ' Select "YEARLY REPORT" worksheet and call subroutines to insert, format headers and calculate total sum
    Worksheets("YEARLY REPORT").Select
    InsertHeaders
    FormatHeaders
    AutomateTotalSUM
End Sub

Public Sub AutomateTotalSUM()
    Dim lastCell As String
    Dim ws As Worksheet
    
    Range("F2").Select
        
    Selection.End(xlDown).Select
        
    lastCell = ActiveCell.Address(False, False)
        
    ActiveCell.Offset(1, 0).Select
        
    ActiveCell.Value = "=sum(F2:" & lastCell & ")"
    
End Sub

Sub InsertHeaders()
'
' InsertHeaders Macro
' Inserts a new row and add the list headers
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Division"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A2").Select
End Sub


Sub FormatHeaders()
'
' FormatHeaders Macro
' Formats list headers and list content
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Size = 12
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "Currency"
    Columns("B:F").Select
    Columns("B:F").EntireColumn.AutoFit
    Range("A2").Select
End Sub

