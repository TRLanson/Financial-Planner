Attribute VB_Name = "LinkExpensesTableModule"
Sub CopyAndLinkTable()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceTable As ListObject
    Dim destTable As ListObject
    Dim destRange As Range
    Dim i As Long, j As Long
    
    ' Set source and destination worksheets
    Set wsSource = ThisWorkbook.Worksheets("Expenses")
    Set wsDest = ThisWorkbook.Worksheets("Income&Goals")
    
    ' Set the source table
    Set sourceTable = wsSource.ListObjects("SummaryTable")
    
    ' Clear any existing data in the destination range
    wsDest.Cells.Clear
    
    ' Set the destination range to start where the source table headers are copied
    Set destRange = wsDest.Range("O3")
    
    ' Copy the headers
    For j = 1 To sourceTable.HeaderRowRange.Columns.Count
        destRange.Cells(1, j).Value = sourceTable.HeaderRowRange.Cells(1, j).Value
    Next j
    
    ' Link the data using formulas
    For i = 1 To sourceTable.DataBodyRange.Rows.Count
        For j = 1 To sourceTable.DataBodyRange.Columns.Count
            destRange.Cells(i + 1, j).Formula = "=" & sourceTable.DataBodyRange.Cells(i, j).Address(External:=True)
        Next j
    Next i
    
    ' Format the destination range as a table
    Set destTable = wsDest.ListObjects.Add(xlSrcRange, destRange.CurrentRegion, , xlYes)
    destTable.Name = "SummaryTable2"
    

End Sub

