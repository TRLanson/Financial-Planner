Attribute VB_Name = "ClearExpensesModule"
Sub ClearTableContents()
    Dim tbl As ListObject
    
    Set tbl = ThisWorkbook.Sheets("Expenses").ListObjects("expenses_table")
    
    ' Clear all data within the table's data body range
    tbl.DataBodyRange.ClearContents
End Sub

