Attribute VB_Name = "AnalyzeActualSpendingModule"
Sub AnalyzeActualSpending()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim category As String
    Dim overSpentCategory As String
    Dim closestToExpectedCategory As String
    Dim overspendAmount As Double
    Dim closestDifference As Double
    Dim difference As Double
    Dim resultMsg As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Expected Spending")
    
    ' Find the last row in the table (assumes no gaps in column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize variables
    overspendAmount = 0
    closestDifference = WorksheetFunction.Max(ws.Range("B2:B" & lastRow))
    overSpentCategory = ""
    closestToExpectedCategory = ""
    
    ' Loop through the table
    For i = 4 To lastRow
        ' Get the category, expected spending, and actual spending
        category = ws.Cells(i, "A").Value
        ExpectedSpending = ws.Cells(i, "B").Value
        actualSpending = ws.Cells(i, "C").Value
        
        ' Calculate the difference
        difference = actualSpending - ExpectedSpending
        
        ' Check for overspending
        If difference > 0 Then
            If difference > overspendAmount Then
                overspendAmount = difference
                overSpentCategory = category
            End If
        Else
            ' Check for the closest category to expected spending (when under budget)
            If Abs(difference) < closestDifference Then
                closestDifference = Abs(difference)
                closestToExpectedCategory = category
            End If
        End If
    Next i
    
    If overSpentCategory <> "" Then
        resultMsg = "You overspent by $" & Format(overspendAmount, "0.00") & " in the '" & overSpentCategory & "' category." & vbCrLf & _
                    "Consider cutting back on this category."
    Else
        resultMsg = "All categories are under budget!" & vbCrLf & _
                    "The category closest to the expected spending is '" & closestToExpectedCategory & "'. Try cutting back on this category."
    End If
    
    ' Display the message box
    MsgBox resultMsg, vbInformation, "Spending Analysis"
End Sub


