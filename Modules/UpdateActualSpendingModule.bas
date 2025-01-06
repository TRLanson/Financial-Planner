Attribute VB_Name = "UpdateActualSpendingModule"
Sub UpdateActualSpending()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceTable As Range
    Dim targetColumn As Range
    Dim rowIndex As Integer
    
    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Sheets("Expenses") ' Replace with the name of the source sheet
    Set wsTarget = ThisWorkbook.Sheets("Expected Spending") ' Replace with the name of the target sheet
    
    ' Set the range of the source and target columns
    Set sourceTable = wsSource.Range("L4:L11")
    Set targetColumn = wsTarget.Range("C4:C11")
    
    ' Loop through each row in the source table and copy the value to the target column
    For rowIndex = 1 To sourceTable.Rows.Count
        ' Update the target column with the value from the source table
        targetColumn.Cells(rowIndex, 1).Value = sourceTable.Cells(rowIndex, 1).Value
    Next rowIndex
End Sub



