Attribute VB_Name = "FindHighestSpendingModule"
Sub FindHighestSpending()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim maxAmount As Double
    Dim categories As String

    ' Set worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Expenses")

    ' Find the last row in column B (Amount Spent) to define the range
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set rng = ws.Range("K4:L" & lastRow) ' Data starts in Row 4

    ' Initialize maxAmount to a very small number
    maxAmount = -1

    ' Loop through each cell in the range
    For Each cell In rng.Columns(2).Cells ' Look at column K
        If cell.Value > maxAmount Then
            ' Update maxAmount and reset categories if a new max is found
            maxAmount = cell.Value
            categories = ws.Cells(cell.Row, "K").Value ' Category in column K
        ElseIf cell.Value = maxAmount Then
            ' Append category if it has the same max amount
            categories = categories & ", " & ws.Cells(cell.Row, "K").Value
        End If
    Next cell
    
    If maxAmount <> 0 Then
    ' Display result in a message box
        MsgBox "The highest amount spent in any category is $" & maxAmount & _
           ". Try to reduce spending in the following category/categories: " & categories, vbInformation, "Highest Spending"
    Else
        MsgBox "You have not entered any spending data. Please enter your information into the " & "Expenses" & " sheet to get advice."
    End If
        
End Sub


