
Sub loopInsertFormatHeaders()
    Dim ws As Worksheet
    
      
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        If Range("A1").Value <> "Division" Then
            insertHeaders
            formatHeaders
        End If
    Next ws

End Sub

Sub insertHeaders()
'
' Insert_Headers Macro
' Inserts a new row and adds the list headers

    Rows("1:1").Select
    Calculate
    Selection.Borders(xlLeft).LineStyle = xlNone
    Selection.Borders(xlRight).LineStyle = xlNone
    Selection.Borders(xlTop).LineStyle = xlNone
    Selection.Borders(xlBottom).LineStyle = xlNone
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
    Range("C2").Select
End Sub
Sub formatHeaders()
'
' FormatHeaders Macro
' formats headers and lists content
'
    Range("A1:F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("C2:F15").Select
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.NumberFormat = _
        "_([$$-en-US]* #,##0.00_);_([$$-en-US]* (#,##0.00);_([$$-en-US]* ""-""??_);_(@_)"
    Columns("B:F").Select
    Columns("B:F").EntireColumn.AutoFit
    Range("A2").Select
End Sub
