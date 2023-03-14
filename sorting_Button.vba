  
Public Sub UserSortInput()
    'get the sort order from the user
    Dim sortOrder As Integer
    Dim promptMSG As String
    Dim tryAgain As Integer
    
    On Error GoTo errHandler
    
    
    promptMSG = "How would you like to sort the list?" & vbCrLf & _
    "l - Sort by Division" & vbCrLf & _
    "2 - Sort by Category" & vbCrLf & _
    "3 - Sort by Total"
    
    sortOrder = InputBox(promptMSG, "Sort Order")
    
    If sortOrder = 1 Then
        Division_sort
    ElseIf sortOrder = 2 Then
        Category_sort
    ElseIf sortOrder = 3 Then
        Total_sort
    Else
errHandler:
        tryAgain = MsgBox("Invalid input. Try again?", vbYesNo)
        If tryAgain = 6 Then
            UserSortInput
        
        End If
    End If  

End Sub

Public Sub Division_sort()
    ' sorts the list by Division column
    Columns("A:F").Sort key1:=Range("A2"), order1:=xlDescending, Header:=xlYes

End Sub
Public Sub Category_sort()
    ' sorts the list by Category column
    Columns("A:F").Sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes

End Sub
Public Sub Total_sort()
    ' sorts the list by Total column
    Columns("A:F").Sort key1:=Range("f2"), order1:=xlDescending, Header:=xlYes

End Sub

