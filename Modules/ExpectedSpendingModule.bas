Attribute VB_Name = "ExpectedSpendingModule"
Sub ShowExpectedSpendingForm()
    ExpectedSpendingForm.Show
End Sub

Sub ClearExpectedSpendings()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colName As String
    Dim tblColumn As ListColumn
    
    Set ws = ThisWorkbook.Worksheets("Expected Spending")
    Set tbl = ws.ListObjects("ExpectedSpendingTable")
    
    ' Specify the column name to clear
    colName = "Expected Spending"
    Set tblColumn = tbl.ListColumns(colName)
    tblColumn.DataBodyRange.ClearContents
End Sub

